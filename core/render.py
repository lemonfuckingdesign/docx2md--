#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""YAML/Markdown 生成与章条/公文/无编号结构解析。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
"""

from pathlib import Path
import re
from config import RE_ARTICLE, RE_BARE_WEN, RE_CHAPTER, RE_INLINE_SPLIT, RE_QUESTION, RE_SIGN_DATE, TINY_SECTION_BODY_LEN, is_flat_mode
from .recognize import _label_heading_ok, _title_as_body, cn_to_int, int_to_cn, is_heading_like, level_marker
from .split import make_theme


def yaml_escape(s):
    return str(s).replace('"', "'")


def _part_lines(extra):
    """多篇文档时，补充 part / part_total 两个 YAML 字段。"""
    if not extra:
        return []
    d = dict(extra)
    if d.get("part_total", 1) > 1:
        return [("part", d["part"]), ("part_total", d["part_total"])]
    return []


def build_yaml(path: Path, meta, extra_lines):
    """生成 YAML front matter。extra_lines 为 [(key, value)] 形式，value 已序列化。"""
    title = meta["title"] or path.stem
    out = ["---"]
    out.append(f'title: "{yaml_escape(title)}"')
    out.append("doc_type: 法规制度")
    if meta.get("docno"):
        out.append(f'doc_number: "{meta["docno"]}"')
    if meta.get("issuer"):
        out.append(f'issuer: "{yaml_escape(meta["issuer"])}"')
    for key in ("date_original", "date_revised", "date_issued", "date_effective"):
        if meta.get(key):
            out.append(f"{key}: {meta[key]}")
    out.append("status: 现行有效")
    for k, v in extra_lines:
        if v not in (None, "", []):
            out.append(f"{k}: {v}")
    out.append(f'source_file: "{path.name}"')
    out.append("---")
    out.append("")
    return out


def doc_info_lines(meta):
    out = []
    if meta.get("docno"):
        out.append(f"- **发文字号**：{meta['docno']}")
    if meta.get("issuer"):
        out.append(f"- **制发机关**：{meta['issuer']}")
    if meta.get("related_doc"):
        out.append(f"- **解读文件**：{meta['related_doc']}")
    if meta.get("date_original"):
        out.append(f"- **制定时间**：{meta['date_original']}")
    if meta.get("date_revised"):
        out.append(f"- **修订时间**：{meta['date_revised']}")
    if meta.get("date_issued"):
        out.append(f"- **发布时间**：{meta['date_issued']}")
    return out


def index_line(labels, noun="问答", sep="、", suffix=""):
    """生成「内容索引」行：1.xxx、2.xxx、…"""
    if not labels:
        return ""
    items = sep.join(f"{i}.{lb}" for i, lb in enumerate(labels, 1))
    return f"**内容索引**：本篇共 {len(labels)} 组{noun}{suffix}，依次涵盖——{items}。"


def infer_keywords(title, related_doc, extra=()):
    """从标题与关联法规名推导关键词，供 Dify 元数据过滤使用。"""
    kws = []
    base = related_doc.strip("《》") if related_doc else ""
    if base:
        kws.append(base)
    if title:
        # 「XX负责人就印发《YY》答记者问」→ 取负责人主体与主题
        m = re.match(r"^(.{2,20}?)负责人", title)
        if m:
            kws.append(m.group(1))
        if "解读" in title:
            kws.append(base.replace("条例", "条例解读") if base else "政策解读")
        elif "答记者问" in title:
            kws.append("答记者问")
    for e in extra:
        if e:
            kws.append(e)
    # 去重保序
    seen, out = set(), []
    for k in kws:
        k = k.strip().strip("《》")
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out


def build_md_qa(path: Path, meta, intro, qas, extra=None):
    """问答体例输出：## 第N问 {主题} + **问：** + **答：**。"""
    title = meta["title"] or path.stem
    themes = [make_theme(qa["q"]) if qa["q"] else f"补充说明{i}"
              for i, qa in enumerate(qas, 1)]
    kws = infer_keywords(title, meta.get("related_doc", ""))
    lines = [
        ("structure_type", "问答体例"),
        ("qa_count", len(qas)),
        ("keywords", "[" + ", ".join(kws) + "]"),
    ]
    lines.extend(_part_lines(extra))
    out = build_yaml(path, meta, lines)
    out.append(f"# {title}")
    out.append("")
    out.extend(doc_info_lines(meta))
    if meta.get("issuer") or meta.get("date_issued") or meta.get("related_doc"):
        out.append("")
    for p in intro:
        if p != title:
            out.append(p)
            out.append("")
    idx = index_line(themes, "问答")
    if idx:
        out.append(idx)
        out.append("")
    out.append(f"> 说明：本文件由原始文档（{path.name}）自动结构化重排，问答文字未作改写；"
               "每组问答独立成节，便于知识库检索与引证。")
    out.append("")
    for i, qa in enumerate(qas, 1):
        out.append(f"## 第{i}问 {themes[i - 1]}")
        out.append("")
        if qa["q"]:
            q = qa["q"].strip().rstrip("。")
            if not RE_QUESTION.match(q):
                q = "问：" + q
            out.append(f"**{q}**" if q.endswith(("？", "?")) else f"**{q}？**")
            out.append("")
        for p in qa["a"]:
            # 答案段落在句末标点+序号处再切分，保证分段合理、不切断句子
            pieces = split_long(p)
            for j2, text in enumerate(pieces):
                if re.match(r"^答\s*[：:]", text) and j2 == 0:
                    text = re.sub(r"^答\s*[：:]\s*", "", text)
                    # 只有第一段加粗标注「答：」，后续段落去掉重复前缀
                    out.append(f"**答：**{text}")
                else:
                    out.append(text)
                out.append("")
    return "\n".join(out).rstrip() + "\n"


def md_header(path: Path, title, meta, extra_lines):
    """生成 YAML front matter + 标题区，两条解析路径共用。"""
    out = build_yaml(path, meta, [tuple(x.split(": ", 1)) if isinstance(x, str) else x
                                 for x in extra_lines])
    out.append(f"# {title}")
    out.append("")
    return out


def normalize_levels(items):
    """把实际出现的编号序列映射成连续的标题层级。

    公文里常见「一、」下面直接跟「（一）」，中间没有「1.」；
    如果按固定层级出标题会跳级（H2 直接到 H4）。这里按出现顺序重新压缩。
    """
    raw_levels = sorted({lv for it in items if (lv := it["level"]) > 0})
    mapping = {lv: i + 2 for i, lv in enumerate(raw_levels)}  # 从 H2 起
    for it in items:
        if it["level"] > 0:
            it["level"] = mapping[it["level"]]
    return items


def split_inline_items(text):
    """在段内「句末标点 + 序号标记」处切开，把挤在同一段里的多个编号条目拆开。

    切分点前必有句末标点（。！？；），因此不会切断任何一句话；
    只在序号（第X条/一、/（一）/1./①）前下刀，保证每个片段都以编号开头。
    """
    if not text:
        return [text]
    parts = [p.strip() for p in RE_INLINE_SPLIT.split(text) if p and p.strip()]
    # 去掉首片段只剩标点的退化情况
    parts = [p for p in parts if re.search(r"[\u4e00-\u9fa0-9A-Za-z]", p)]
    return parts or [text]


def split_long(text, limit=500):
    """超长段落按「完整句子」边界切分，绝不把一句话拆成两半。

    切分分两层：
      1. 先在段内「句末标点 + 序号标记」处切开（第X条/一、/（一）/1. 等），
         满足“在第几等序号处分段”的要求——即使整段未超长也会拆开；
      2. 每个片段若仍超长，再按句末标点（。！？）聚合切分。
    分号、逗号属于句内停顿，不作为切分点。
    单句本身超长时原样保留（宁可块长，也不裁断句子）。
    """
    parts = []
    for seg in split_inline_items(text):
        if len(seg) <= limit:
            parts.append(seg)
            continue
        # 仅在句末标点后切分，保证每个片段都是完整句子
        sents = [s for s in re.split(r"(?<=[。！？])", seg) if s]
        buf = ""
        for s in sents:
            if buf and len(buf) + len(s) > limit:
                parts.append(buf)
                buf = s
            else:
                buf += s
        if buf:
            parts.append(buf)
    return parts


def parse_gongwen(texts, blanks):
    """解析「非第X条」体例的公文：按 一、/（一）/1. 多级编号分层。

    返回 (head, items, warnings)：
      head  —— 文头（标题、文号、主送机关、正文前言等），原样保留
      items —— [{"level":int, "title":str, "paras":[str]}]，level 为最终 H 层级
    """
    warnings = []
    head = []
    items = []
    cur = None

    # 文头：第一个编号标题之前的全部内容
    start = 0
    for i, t in enumerate(texts):
        if level_marker(t)[0]:
            start = i
            break
    else:
        start = len(texts)
    head = texts[:start]
    body = texts[start:]
    body_blanks = blanks[start:]

    for idx, (t, has_blank) in enumerate(zip(body, body_blanks)):
        lv, _num = level_marker(t)
        if lv:
            # 「（一）加强组织领导。各级党委……」这类标题与正文同行的写法，
            # 若编号后第一句很短且紧跟句号，则拆成标题 + 正文两段
            tm = re.match(
                r"^([\s　]*[（(]?[一二三四五六七八九十百0-9①-⑳]+[)）]?[、．.]?[\s　]*)"
                r"([^。；！？]{2,20}[。:：])(.+)$", t)
            if tm and len(tm.group(2)) <= 22:
                cur = {"level": lv, "title": tm.group(1).strip() + tm.group(2).rstrip("。:："),
                       "paras": [], "blank": has_blank}
                body_rest = tm.group(3).strip()
            else:
                cur = {"level": lv, "title": t, "paras": [], "blank": has_blank}
                body_rest = ""
            items.append(cur)
            # 正文剩余部分继续走序号切分：段内挤进的下级条目（（一）/1. 等）
            # 会在句末标点处切开并升格为新的小节
            pending = split_long(body_rest) if body_rest else []
        else:
            pending = [t]
            # —— 以下判断仅针对独立成段的非编号段落 ——
            if cur is None:
                head.append(t)
                continue
            # 落款日期：单独记录，不进正文
            if RE_SIGN_DATE.match(t):
                cur.setdefault("signdates", []).append(t)
                continue
            # 落款机关：紧跟在落款日期之前、短且无标点的机构名，不并入正文
            if (len(t) <= 20 and not re.search(r"[。；！？，,、]$", t)
                    and any(RE_SIGN_DATE.match(x) for x in body[idx + 1: idx + 3])):
                cur["signoff"] = t
                continue
            # 无编号段落：判断是“标题的延续”还是“正文”
            # 仅当标题行短、不以句号结尾、且原文中该段与前一段未空行分隔时才合并
            if (not cur["paras"] and len(cur["title"]) <= 25 and is_heading_like(t)
                    and not has_blank and len(t) <= 20):
                cur["title"] += t
                continue
        # 正文里挤进多个编号条目（「……。（二）……。」）时，
        # 在序号处切开：序号片段升格为新小节标题，其余仍是正文。
        # split_long 只在句末标点+序号处下刀，不切断句子。
        for piece in pending:
            lv2, _n2 = level_marker(piece)
            if lv2:
                # 「（二）加强管理。各级……」：首句短且带句号时拆成标题+正文
                tm2 = re.match(
                    r"^([\s　]*[（(]?[一二三四五六七八九十百0-9①-⑳]+[)）]?[、．.]?[\s　]*)"
                    r"([^。；！？]{2,20}[。:：])(.+)$", piece)
                if tm2 and len(tm2.group(2)) <= 22:
                    cur = {"level": lv2, "title": tm2.group(1).strip() + tm2.group(2).rstrip("。:："),
                           "paras": [], "blank": False}
                    items.append(cur)
                    rest2 = tm2.group(3).strip()
                    # 标题后正文里若还嵌套下级条目，继续递进切分
                    for piece3 in split_long(rest2):
                        lv3, _n3 = level_marker(piece3)
                        if lv3:
                            cur = {"level": lv3, "title": piece3, "paras": [], "blank": False}
                            items.append(cur)
                        else:
                            cur["paras"].append(piece3)
                else:
                    cur = {"level": lv2, "title": piece.rstrip("。：:；;"),
                           "paras": [], "blank": False}
                    items.append(cur)
            else:
                cur["paras"].append(piece)

    # 只有标题没有正文的结构（多为纯目录）：若条目数很少且都是短行，
    # 说明该文件本身只有框架、没有正文，则按原编号逐条输出即可；
    # 否则（长文档却全无正文，多为目录页）也按条目保留，避免内容丢失。
    only_titles = items and all(not it["paras"] for it in items)
    no_content = not items
    if only_titles:
        short = all(len(it["title"]) <= 30 for it in items)
        if short and len(items) >= 2:
            warnings.append("文件仅有编号框架、无正文，按原编号逐条保留")
            items = normalize_levels(items)
            return head, items, warnings
        warnings.append("全文仅有编号条目无实质正文，改为按段落切分")
    if no_content or only_titles:
        items = []
        cur = None
        for t in texts:
            if cur is None or len("".join(cur["paras"])) >= 500:
                cur = {"level": 2, "title": "", "paras": [], "blank": True}
                items.append(cur)
            cur["paras"].append(t)
        # 文头只保留真正的版头信息（标题、文号、主送机关等），
        # 长段正文不属于文头，避免与下面的分块重复
        head = [h for h in head if len(h) <= 40 and not re.search(r"。", h)]

    items = normalize_levels(items)

    # 层级分布统计，写进备注便于核对
    dist = {}
    for it in items:
        dist[it["level"]] = dist.get(it["level"], 0) + 1
    if dist:
        warnings.append(
            "按编号分层：" + "、".join(f"H{k} {v}个" for k, v in sorted(dist.items())))

    return head, items, warnings


def parse_structure(texts, blanks=None):
    """切出 (章名, [条文...]) 结构。返回 chapters 列表与告警信息。"""
    chapters = []
    warnings = []
    blanks = blanks or [False] * len(texts)

    # 定位正文起点：第一条/第一章 出现的位置，之前视为文头
    start = 0
    for i, t in enumerate(texts):
        if RE_CHAPTER.match(t) or RE_ARTICLE.match(t):
            start = i
            break

    body = texts[start:]

    # 目录识别：仅在“同一章标题重复出现”且两次出现之间没有任何条文时才认定为目录
    # （法规正文里章标题不会重复；某些 docx 会在正文前贴一份目录）
    toc_end = 0
    first_seen = {}
    for i, t in enumerate(body):
        cm = RE_CHAPTER.match(t)
        if cm:
            name = (cm.group(1), cm.group(2).strip())
            if name in first_seen:
                toc_end = i
                break
            first_seen[name] = i

    if toc_end > 0:
        body = body[toc_end:]
        warnings.append(f"跳过开头 {toc_end} 行重复的目录条目")

    cur = None
    for t in body:
        cm = RE_CHAPTER.match(t)
        if cm:
            num = cn_to_int(cm.group(1))
            name = cm.group(2).strip()
            cur = {"num": num, "raw_num": cm.group(1), "name": name, "articles": []}
            chapters.append(cur)
            continue
        am = RE_ARTICLE.match(t)
        if am:
            if cur is None:
                # 没有章标题的公文：建一个虚拟章
                cur = {"num": 0, "raw_num": "0", "name": "", "articles": []}
                chapters.append(cur)
            # 条文剩余部分里挤进下一条（「……。第十二条 ……」）时，
            # 在序号处切开并各自升格为独立条文
            rest = am.group(2).strip()
            pieces = split_long(rest) if rest else []
            if pieces and RE_ARTICLE.match(pieces[0]):
                # 剩余部分整体以新条开头（极少数情况），按普通段落处理
                cur["articles"].append({"no": am.group(1), "paras": pieces})
            else:
                cur["articles"].append({"no": am.group(1), "paras": [pieces[0]] if pieces else []})
                for piece in pieces[1:]:
                    am2 = RE_ARTICLE.match(piece)
                    if am2:
                        cur["articles"].append({"no": am2.group(1), "paras": [am2.group(2)]})
                    else:
                        cur["articles"][-1]["paras"].append(piece)
            continue
        if cur and cur["articles"]:
            # 段内挤进多个「第X条」时，在序号处切开并各自升格为独立条文
            for piece in split_long(t):
                am2 = RE_ARTICLE.match(piece)
                if am2:
                    cur["articles"].append({"no": am2.group(1), "paras": [am2.group(2)]})
                elif cur["articles"]:
                    cur["articles"][-1]["paras"].append(piece)
                else:
                    cur["articles"].append({"no": "", "paras": [piece]})
        elif cur is not None:
            # 章标题之后、第一条之前的内容，作为导语
            cur.setdefault("preamble", []).append(t)

    if any(c["num"] == 0 for c in chapters):
        warnings.append("文件无章标题，条文归入「未分章」")

    return chapters, warnings


def _demote_empty_leaf_titles(items):
    """把「没有正文支撑的叶子标题」还原为正文条目。

    会议记录/方案类文档常见一连串序号标题彼此紧挨、中间无任何正文
    （如 (一)~(四) 四个阶段条目逐条排列）。它们挨得很近、分得很细，
    逐个生成分段标记只会把文档切碎。规则：
      * 标题下没有正文，且下一个标题的层级不比它深（同级或更浅）
        → 它是「叶子」，还原为正文一行，并入前一个条目；
      * 下一个标题层级更深 → 它是分组标题（如「三、活动内容与步骤」），
        保留，降级后的子条目自然成为它的正文；
      * 「第N问」问答标题始终保留。
    """
    if not items:
        return items
    kept = []
    for idx, it in enumerate(items):
        nxt = items[idx + 1] if idx + 1 < len(items) else None
        t = (it.get("title") or "").strip()
        is_qa = bool(re.match(r"^第[一二三四五六七八九十百\d]+问", t))
        empty_leaf = (t and not it.get("paras") and not is_qa
                      and (nxt is None or nxt.get("level", 2) <= it["level"]))
        if empty_leaf:
            if kept:
                kept[-1].setdefault("paras", []).append(t)
            else:
                it = {"level": it["level"], "title": "", "paras": [t]}
                kept.append(it)
            continue
        kept.append(it)
    return kept


def _merge_bare_wen_titles(items):
    """「1、问」这类只有问字、不带问题内容的标题：把问题并回标题。

    谈话记录体例常见「1、问」单独成行，问题正文在下一段（以「？」收尾）。
    处理：标题改为「N、问：<问题原文？>」，剩余文字（含答句）留在正文。
    若下一段提取不到问号句，则整个标题降为正文，避免孤零零一个问字。
    """
    for it in items:
        t = (it.get("title") or "").strip()
        m = RE_BARE_WEN.match(t)
        if not m or not it.get("paras"):
            continue
        first = it["paras"][0].strip()
        # 问号或句号收尾的首句都视作问题（有的文档问句以句号结尾）
        qm = re.match(r"^(.{2,60}?[？?。])(.*)$", first, re.S)
        if qm:
            it["title"] = f"{m.group(1).strip()}问：{qm.group(1).strip()}"
            rest = qm.group(2).strip()
            if rest:
                it["paras"][0] = rest
            else:
                it["paras"].pop(0)
        else:
            it["paras"].insert(0, t)
            it["title"] = ""
    return items


def _merge_tiny_leaf_sections(items):
    """把「标题+一句话正文」的极短叶子小节并回正文段落。

    Dify 等知识库按 Markdown 标题切分检索块，像
      #### 1、调查对象
      天津市青年京剧团业务办公楼、演员招待所工程
    这样标题下只有一句话的小节会各自成为一个几十字的碎块，
    检索时既缺上下文又浪费块数。规则：
      * 仅处理 level>=3 的叶子小节（下一个条目层级不比它深）；
      * 自身正文 < TINY_SECTION_BODY_LEN 字；
      * 处理方式：标题还原为正文，并与正文合并成一个自然段
        （「1、调查对象：天津市青年京剧团…」），
        使该内容归入上一级小节的检索块。
    「第N问」问答标题不受影响。
    """
    if not items:
        return items
    kept = []
    for idx, it in enumerate(items):
        nxt = items[idx + 1] if idx + 1 < len(items) else None
        title = (it.get("title") or "").strip()
        body_len = sum(len(x.strip()) for x in it.get("paras", []))
        is_qa = bool(re.match(r"^第[一二三四五六七八九十百\d]+问", title))
        tiny_leaf = (title and it.get("level", 2) >= 3 and not is_qa
                     and body_len < TINY_SECTION_BODY_LEN
                     and (nxt is None or nxt.get("level", 2) >= it["level"]))
        if tiny_leaf:
            body = "".join(x.strip() for x in it["paras"])
            if body and not title.endswith(("：", ":", "。", "；")):
                line = f"{title}：{body}" if body else title
            else:
                line = title + body
            if kept and kept[-1].get("title") is not None:
                kept[-1].setdefault("paras", []).append(line)
            else:
                it = {"level": it["level"], "title": "", "paras": [line]}
                kept.append(it)
            continue
        kept.append(it)
    return kept


def _unify_sibling_marking(items):
    """同一父节下的子条目标题标记保持一致。

    背景：同一父节（如「二、党支部委员会议」）下的编号条目，如果有的
    标 ###/####、有的还原为正文，层级观感混乱。规则：取某个标题节点的
    全部下级条目（直到遇到同级或更浅的标题为止）为一组；组内各条目
    「最终是否保留标记」的判定结果不一致时，整组统一降为正文。

    判定某条目最终会保留标记 = 标题不是句子式（_title_as_body 为 False）
    且不是「标题+一句话正文」的极短叶子小节（会被 _merge_tiny_leaf_sections
    合并回正文）。
    """
    n = len(items)
    i = 0
    while i < n:
        if not items[i].get("title"):
            i += 1
            continue
        lvl = items[i].get("level", 2)
        # 子块范围：层级比当前标题深、直到遇到同级或更浅标题为止
        j = i + 1
        while j < n and items[j].get("level", 2) > lvl:
            j += 1
        children = items[i + 1:j]
        if lvl >= 2 and len(children) >= 2:
            marked = []
            for k, it in enumerate(children):
                if not it.get("title"):
                    marked.append(False)
                    continue
                # 与 _merge_tiny_leaf_sections 的叶子判定保持一致：
                # 用全局下一项（而非子块切片内）判断是否叶子
                nxt = items[i + 2 + k] if i + 2 + k < n else None
                leaf = nxt is None or nxt.get("level", 2) >= it.get("level", 2)
                body_len = sum(len(p) for p in it.get("paras", []))
                tiny = leaf and body_len < TINY_SECTION_BODY_LEN
                marked.append(not _title_as_body(it["title"]) and not tiny)
            if any(marked) and not all(marked):
                for it in children:
                    if it.get("title"):
                        it["_force_body"] = True
        i = j
    return items


def build_md_gongwen(path: Path, meta, head, items, warnings, extra=None):
    """生成「非第X条」体例公文的 md。"""
    items = _unify_sibling_marking(items)
    items = _merge_tiny_leaf_sections(items)
    items = _demote_empty_leaf_titles(items)
    items = _merge_bare_wen_titles(items)
    title = meta["title"] or path.stem
    out = md_header(path, title, meta, [
        ("structure_type", "公文编号体例"),
        ("section_count", len(items)),
    ] + _part_lines(extra))
    out.extend(doc_info_lines(meta))
    numbered = any(it["title"] for it in items)
    if numbered:
        out.append(f"- **文件结构**：按原文编号分 {len(items)} 个小节")
    else:
        out.append(f"- **文件结构**：全文无编号，按内容长度切分为 {len(items)} 段")
    out.append("")

    # 目录（仅列最外层标题，避免目录本身过长）
    if items:
        top = min(it["level"] for it in items)
        tops = [it for it in items if it["level"] == top and it["title"]]
        if tops and len(items) > len(tops):
            labels = [re.sub(r"^[\s　]*[（(]?[一二三四五六七八九十百0-9①-⑳]+[)）]?[、．.]?\s*", "",
                             it["title"]).strip() for it in tops]
            # 有子节时说明总节数，避免索引只列顶层造成“共 N 组”与实际不符
            extra = len(items) - len(tops)
            suffix = f"（含下级分解共 {len(items)} 节）" if extra > 0 else ""
            line = index_line(labels, "小节", suffix=suffix)
            out.append(line)
            out.append("")

    if not numbered:
        # 无编号文本：把完整短首段作为主题标签补进索引，便于 Dify 概览。
        # 长段/半句截断的标签不进索引，避免索引行碎片化。
        labels = []
        for it in items:
            if it["paras"]:
                lb = re.split(r"[。；！？，,]", it["paras"][0])[0][:16]
                if lb and _label_heading_ok(lb) and it["paras"][0].strip() == lb.strip():
                    labels.append(lb)
        if labels:
            out.append(index_line(labels, "段落"))
            out.append("")

    if numbered:
        out.append(f"> 说明：本文件由原始文档（{path.name}）自动结构化重排，正文文字未作改写；"
                   "标题层级依据原文编号（一、/（一）/1. 等）生成，以便知识库检索与引证。")
    else:
        out.append(f"> 说明：本文件由原始文档（{path.name}）自动结构化重排，正文文字未作改写；"
                   "原文无编号结构，按段落与语义长度切分，以便知识库检索与引证。")
    out.append("")

    # 文头原样保留：直接接在标题下，不再套「文头信息」标题，避免与正文重复
    rest = [t for t in head if t and t != meta.get("title")]
    if rest:
        # 行间不留空行：整个文头内容块在 Dify 按空行切分时只占一个分段，
        # 避免「文头内容：」「文件标题」各自成为几字块的碎分段
        out.append("**文头内容**：")
        for line in rest:
            out.append(f"> {line}")
        out.append("")

    # 逐节输出；为下级小节补充「所属」上下文行，
    # 使 Dify 按标题分块后，每个 chunk 仍能自证其归属层级，便于检索与引证
    for it in items:
        level = min(max(it["level"], 2), 6)
        hashes = "#" * level
        if is_flat_mode() and level >= 3:
            # 扁平模式：下级标题还原为正文行，内容并入上级块
            if it["title"]:
                it = {"level": level, "title": "", "paras": [it["title"]] + it.get("paras", [])}
            elif not it.get("paras"):
                continue
        if (it["title"] and not _title_as_body(it["title"])
                and not it.get("_force_body")):
            out.append(f"{hashes} {it['title']}")
            out.append("")
        elif it["title"]:
            # 太短的/句子式的「标题」不做分段标记：整行还原为正文段落，
            # 避免每个一句话序号都变成 ###/#### 小节（如「1、讨论如何进一步
            # 优化村内环境。」这类会议记录条目）。
            out.append(it["title"])
            out.append("")
            for p in it["paras"]:
                out.append(p)
                out.append("")
            continue
        elif it["paras"]:
            # 无标题的兜底块：仅当首段本身就是一个完整短行（如独立标题行）时，
            # 才把它升格为小标题；否则（首段是长句、标签只是截断的半句话，
            # 如「对xx同志批评意见：给你提2条批」「封一座城」）直接作正文输出，
            # 避免把正文切碎成无意义的 ###/#### 标题。
            first = it["paras"][0]
            label = re.split(r"[。；！？，,]", first)[0][:16]
            # flat 模式下刚把标题降为正文段，不能再升格回去
            if (not is_flat_mode() and label and _label_heading_ok(label)
                    and first.strip() == label.strip()):
                out.append(f"{hashes} {label}")
                out.append("")
                # 若首段与标签完全相同（如纯目录文件），不再重复输出正文
                it = {"paras": it["paras"][1:]}
        for p in it["paras"]:
            out.append(p)
            out.append("")

    # 落款单独成块，不混在前一条正文末尾
    signoffs = [it for it in items if it.get("signoff") or it.get("signdates")]
    if signoffs:
        out.append("---")
        out.append("")
        out.append("**落款**：")
        out.append("")
        for it in signoffs:
            if it.get("signoff"):
                out.append(f"- 发文机关：{it['signoff']}")
            for d in it.get("signdates", []):
                out.append(f"- 成文日期：{d}")
        out.append("")

    return "\n".join(out).rstrip() + "\n", len(items), 0


def build_md(path: Path, meta, chapters, warnings, extra=None):
    title = meta["title"] or path.stem
    total_articles = sum(len(c["articles"]) for c in chapters)
    real_chapters = [c for c in chapters if c["num"] > 0]

    out = md_header(path, title, meta, [
        ("chapter_count", len(real_chapters)),
        ("article_count", total_articles),
    ] + _part_lines(extra))
    out.extend(doc_info_lines(meta))
    if real_chapters:
        out.append(f"- **文件结构**：共{int_to_cn(len(real_chapters)) if len(real_chapters) < 100 else len(real_chapters)}章{int_to_cn(total_articles) if total_articles < 100 else total_articles}条")
    else:
        out.append(f"- **文件结构**：不分章，共{int_to_cn(total_articles) if total_articles < 100 else total_articles}条")
    out.append("")

    if real_chapters:
        # 内容索引：一句话概述全部章节，替代原表格（表格在 Dify 里易被切碎）
        labels = [f"第{c['raw_num']}章 {c['name']}".rstrip() for c in real_chapters]
        out.append(index_line(labels, "章"))
        out.append("")
        out.append("| 章次 | 章名 | 条文范围 |")
        out.append("| --- | --- | --- |")
        for c in real_chapters:
            if c["articles"]:
                rng = f"第{c['articles'][0]['no']}条—第{c['articles'][-1]['no']}条"
            else:
                rng = "—"
            out.append(f"| 第{c['raw_num']}章 | {c['name']} | {rng} |")
        out.append("")

    out.append(f"> 说明：本文件由原始文档（{path.name}）自动结构化重排，条文文字未作改写；仅调整标题层级、补充元数据与章节归属，以便知识库检索与引证。")
    out.append("")

    for c in chapters:
        if c["num"] > 0:
            out.append("---")
            out.append("")
            out.append(f"## 第{c['raw_num']}章 {c['name']}".rstrip())
            out.append("")
            out.append(f"> 本部分出自《{title}》第{c['raw_num']}章 {c['name']}".rstrip() + "。")
            out.append("")
            for line in c.get("preamble", []):
                out.append(line)
                out.append("")
        for a in c["articles"]:
            if a.get("no"):
                out.append(f"### 第{a['no']}条")
                out.append("")
            for p in a["paras"]:
                if p.strip():
                    out.append(p.strip())
                    out.append("")

    return "\n".join(out).rstrip() + "\n", total_articles, len(real_chapters)


