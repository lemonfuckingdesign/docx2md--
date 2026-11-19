#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""聚合层：单篇转换总控、覆盖率质检、合集标题处理。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
"""

from collections import Counter
from pathlib import Path
import re
from config import CLEAN_FOR_DIFF, RE_ARTICLE, RE_CHAPTER
from .text import _merge_wrapped_paras, _split_glued_paras, demote_fake_headings, normalize_qa_prefix
from .render import build_md, build_md_gongwen, build_md_qa, parse_gongwen, parse_structure, split_long
from .recognize import detect_meta, extract_date_context, extract_related_doc, level_marker
from .split import parse_qa


def verify(orig_texts, md_text, baseline_texts=None):
    """核验原文内容是否完整落入 md。

    baseline_texts 为“内容基准”：问答/公文路径会先做标题降级去重，
    假标题的文字已在下一段完整保留，若仍按去重前的原文统计，
    重复计数会把覆盖率压低（出现假缺失）。因此传入去重后的序列。

    有条文编号时逐条比对；无编号时按字符多重集覆盖率校验（顺序无关，
    因为结构化会重排段落顺序）。
    返回 (原文单元数, 缺失列表, 覆盖率)。
    """
    orig = re.sub(r"\s+", "", "".join(baseline_texts if baseline_texts is not None else orig_texts))
    body = md_text.split("---", 2)[-1]
    # 只去掉 Markdown 结构标记本身，保留标题与正文文字参与比对
    body = re.sub(r"^#{1,6}\s*", "", body, flags=re.M)
    body = re.sub(r"^>\s*", "", body, flags=re.M)
    body = re.sub(r"^\|", "", body, flags=re.M)
    body = body.replace("|", " ")
    body = re.sub(r"^[-*]\s*", "", body, flags=re.M)
    flat = re.sub(r"\s+", "", body)

    arts = re.findall(
        r"第[一二三四五六七八九十百零〇]+条.*?(?=第[一二三四五六七八九十百零〇]+条|第[一二三四五六七八九十]+章|$)",
        orig)
    if arts:
        missing = []
        for a in arts:
            idx = a.index("条") + 1
            body_txt = a[idx:]
            if not body_txt:
                continue
            # 顺序严格匹配在条文含 Markdown 分隔符/多段换行时易误报：
            # 改为「条文正文的实义字符 98% 以上被 md 覆盖」判定。
            body_cnt = Counter(CLEAN_FOR_DIFF.sub("", body_txt))
            total_c = sum(body_cnt.values())
            if not total_c:
                continue
            out_cnt = Counter(CLEAN_FOR_DIFF.sub("", flat))
            hit = sum(min(n, out_cnt.get(ch, 0)) for ch, n in body_cnt.items())
            if hit / total_c < 0.98:
                missing.append(a[:idx])
        cover = 1.0 if not missing else 0.0
        return len(arts), missing, cover

    # 无条文编号：逐段落计算字符覆盖率，再按段落长度加权。
    #
    # 为什么不用全文字符多重集：重排后同一字符（如「请」「最」）在原文与
    # 输出中出现的“次数”会因标题合并/前缀归一而变化，多重集相减会产生
    # 少量“假缺失”，把 100% 压到 99%，掩盖真正的问题。
    # 逐段落判断则直接回答“这段话有没有进 md”，含义更清晰。
    if not orig:
        return 0, [], 1.0
    units = baseline_texts if baseline_texts is not None else orig_texts
    out_flat = CLEAN_FOR_DIFF.sub("", flat)
    out_cnt = Counter(out_flat)
    total_len, covered, bad = 0, 0, []
    for t in units:
        seg = CLEAN_FOR_DIFF.sub("", t)
        if not seg:
            continue
        total_len += len(seg)
        seg_cnt = Counter(seg)
        hit = sum(min(n, out_cnt.get(ch, 0)) for ch, n in seg_cnt.items())
        covered += hit
        if hit / len(seg) < 0.99:
            bad.append(t[:20])
    cover = covered / total_len if total_len else 1.0
    return len(units), bad, cover


def convert_one(texts, blanks, stem, src_name, part_no=0, part_total=1, title_override=""):
    """把一篇文章的段落序列转成 (md 文本, 体例, 条/节数, 章数, 告警, 覆盖率)。

    part_no/part_total 用于一个 docx 含多篇时标注「第X篇/共Y篇」；
    title_override 用于强制以指定文本（如文件名）作为文章标题。
    """
    # 先消除 docx 硬换行造成的断句（机器的发明与/使用导致了……）
    texts, blanks = _merge_wrapped_paras(list(texts), list(blanks))
    # 再拆分句读/编号/制度名粘连的段落（合集类 docx 常见整段粘死）
    texts, blanks = _split_glued_paras(list(texts), list(blanks))
    meta = detect_meta(texts)
    if title_override:
        meta["title"] = title_override
    # 多篇时：文件名不是文章标题，用识别出的标题；同名冲突由调用方处理
    fake_path = Path(src_name)  # 保留原始 .docx 文件名，供 YAML source_file 使用

    baseline = normalize_qa_prefix(demote_fake_headings(texts))
    intro, qas, _cands = parse_qa(texts, meta.get("title", ""), stem)
    has_answered = sum(1 for qa in qas if qa["q"] and qa["a"])

    extra = []
    if part_total > 1:
        extra = [("part", part_no), ("part_total", part_total)]

    if has_answered >= 1 and len(qas) >= 1:
        meta["related_doc"] = extract_related_doc(meta["title"], texts)
        meta["date_context"] = extract_date_context(texts, meta)
        md = build_md_qa(fake_path, meta, intro, qas, extra=extra)
        kind, n_art, n_ch, warns = "问答体例", len(qas), 0, []
    else:
        chapters, warns = parse_structure(texts, blanks)
        n_total = sum(len(c["articles"]) for c in chapters)
        start_i = len(texts)
        for i, t in enumerate(texts):
            if RE_CHAPTER.match(t) or RE_ARTICLE.match(t):
                start_i = i
                break
        numbered_head = any(level_marker(t)[0] for t in texts[:start_i])
        if (n_total >= 2 or any(c["num"] > 0 for c in chapters)) and not numbered_head:
            md, n_art, n_ch = build_md(fake_path, meta, chapters, warns, extra=extra)
            kind = "章条体例"
        else:
            head, items, gwarns = parse_gongwen(texts, blanks)
            if not items:
                head = texts
                items = []
                buf = []
                for t in texts:
                    for piece in split_long(t):
                        buf.append(piece)
                        if len("".join(buf)) >= 500:
                            items.append({"level": 2, "title": "", "paras": buf, "blank": True})
                            buf = []
                if buf:
                    items.append({"level": 2, "title": "", "paras": buf, "blank": True})
                gwarns = ["全文无编号，按段落长度兜底切分"]
                kind = "无编号文本"
            else:
                kind = "公文编号体例"
            md, n_art, n_ch = build_md_gongwen(fake_path, meta, head, items, gwarns, extra=extra)
            warns = gwarns

    _n, missing, cover = verify(texts, md, baseline)
    if missing:
        warns = list(warns) + ["缺失：" + "、".join(missing)]
    elif cover < 0.97:
        warns = list(warns) + [f"内容覆盖率 {cover:.1%}，请抽查"]
    return md, kind, n_art, n_ch, warns, cover


def demote_headings(md_text, levels=1):
    """把 md 里的标题整体降 level 级（篇内 ## → ###），避免与外层篇标题冲突。"""
    out = []
    for line in md_text.split("\n"):
        m = re.match(r"^(#{1,6})\s", line)
        if m and len(m.group(1)) + levels <= 6:
            out.append("#" * (len(m.group(1)) + levels) + line[len(m.group(1)):])
        else:
            out.append(line)
    return "\n".join(out)


def strip_md_head(md_text):
    """剥离单篇 md 的 YAML front matter、一级标题与「说明」行，供合并进合集。"""
    lines = md_text.split("\n")
    i = 0
    if lines and lines[0].strip() == "---":
        i = 1
        while i < len(lines) and lines[i].strip() != "---":
            i += 1
        i += 1
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and lines[i].startswith("# "):
        i += 1
    # 合集场景下剔除每篇重复出现的「说明」行与元信息行
    lines = [ln for ln in lines[i:]
             if not ln.startswith("> 说明：")
             and not ln.startswith(("- **发文字号**", "- **制发机关**",
                                    "- **解读文件**", "- **制定时间**",
                                    "- **修订时间**", "- **发布时间**"))
             and ln.strip() != "**文头内容**："]
    while lines and not lines[0].strip():
        lines = lines[1:]
    return "\n".join(lines).strip()


