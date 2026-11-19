#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""多篇切分与问答解析。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
"""

import re
from config import RE_ANSWER, RE_ANSWER_ALT, RE_ARTICLE, RE_CHAPTER, RE_DATE_LINE, RE_PIECE_MARK, RE_PIECE_MARK_BRACKET, RE_PIECE_MARK_TAIL, RE_PIECE_MARK_TAIL_CN, RE_QUESTION, RE_QUESTION_ALT, STOP_PREFIX
from .text import _strip_qa_prefix, demote_fake_headings, split_head_toc
from .recognize import is_heading_like, level_marker


def match_piece_mark(t):
    """匹配篇标记行，返回 (篇号, 篇标题)；非标记行返回 None。

    优先行首式（「篇1 标题」），其次行尾式（「标题 篇1」/「标题 第一篇」）、
    括号式（「文章名【第一篇】」）。
    """
    t = t.strip()
    m = RE_PIECE_MARK.match(t)
    if m:
        return m.group(1), m.group(2).strip()
    m = RE_PIECE_MARK_TAIL.match(t)
    if m:
        return m.group(2), m.group(1).strip()
    m = RE_PIECE_MARK_TAIL_CN.match(t)
    if m:
        return m.group(2), m.group(1).strip()
    m = RE_PIECE_MARK_BRACKET.match(t)
    if m:
        num = m.group(2) or m.group(3)
        if num:
            return num, m.group(1).strip()
    return None


def _title_piece_count(stem):
    """从文档标题解析篇数提示：「……【最新4篇】」→ 4；无提示返回 0。"""
    m = re.search(r"([一二三四五六七八九十]{1,3}|\d{1,3})\s*篇", stem)
    if not m:
        return 0
    v = m.group(1)
    if v.isdigit():
        return int(v)
    d = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
         "六": 6, "七": 7, "八": 8, "九": 9}
    if "十" in v:
        a, _, b = v.partition("十")
        return (d.get(a, 1) if a else 1) * 10 + (d.get(b, 0) if b else 0)
    return d.get(v, 0)


def _lcs_len(a, b):
    """最长公共子串长度（衡量标题行与汇编标题的相似度）。"""
    import difflib
    m = difflib.SequenceMatcher(None, a, b, autojunk=False) \
        .find_longest_match(0, len(a), 0, len(b))
    return m.size


def _strip_count_suffix(t):
    """去掉篇标题尾部携带的汇编篇数括注（「……【最新4篇】」→「……」）。"""
    return re.sub(r"\s*[（(【\[][^）)】\]]*\d\s*篇[^）)】\]]*[）)】\]]\s*$", "", t).strip()


def _piece_title_candidates(texts, base, expected):
    """收集「看起来像文章标题、且与汇编标题高度相似」的独立标题行。"""
    cands = []
    for i, raw in enumerate(texts):
        t = raw.strip()
        if not t or match_piece_mark(t):
            continue
        if RE_CHAPTER.match(t) or RE_ARTICLE.match(t) or RE_DATE_LINE.match(t):
            continue
        if level_marker(t)[0]:
            continue
        if not is_heading_like(t) or not (6 <= len(t) <= 40):
            continue
        score = _lcs_len(t, base)
        if score >= 4:
            cands.append((i, score))
    # 近重复标题去重：汇编标题与第1篇标题几乎相同（LCS≥较短者-1）时
    # 只保留靠后的那个（汇编标题行自然落入文头）
    dedup = []
    for idx, sc in cands:
        ti = _strip_count_suffix(texts[idx])
        dup = False
        for jdx, _sc2 in cands:
            if jdx == idx or jdx < idx:
                continue
            tj = _strip_count_suffix(texts[jdx])
            if _lcs_len(ti, tj) >= min(len(ti), len(tj)) - 1:
                dup = True
                break
        if not dup:
            dedup.append((idx, sc))
    cands = dedup
    # 超出标题标注的篇数时，只保留相似度最高的前 expected 个
    if expected and len(cands) > expected:
        cands = sorted(cands, key=lambda c: (-c[1], c[0]))[:expected]
    cands.sort()
    return cands


def split_articles(texts, blanks, expected=0, base=""):
    """把段落序列切成多篇合集，返回 ((文头,文头bl), [(篇标题, seg, bl), ...], [篇号,...])。

    篇界来源（按优先级）：
      1. 显式「篇N」标记行（RE_PIECE_MARK 各式）；标记行本身不进正文；
      2. 无标记/标记不足时：文档标题标注了「N篇」→ 用「与汇编标题高度相似的
         独立标题行」补切（用户约定：文档标题里提到几篇，里面一般就有几篇文章）；
      3. 都没有 → 整体作为一篇。
    第一个篇界之前的内容（汇编标题、编辑说明等）归入合集文头，不算任何一篇
    （用户约定：只有原文明确标注「篇几」的才标注第几篇）。
    num 为原文显式标注的篇号（如「篇2」→ 2），无标注的篇为 None。
    """
    if not texts:
        return [], [("", texts, blanks)], []
    marks = [i for i, t in enumerate(texts) if match_piece_mark(t)]

    bounds = []   # [(idx, title, num_or_None, is_candidate)]
    if marks:
        for m in marks:
            _num, title = match_piece_mark(texts[m])
            try:
                num = int(_num)
            except (TypeError, ValueError):
                num = None
            bounds.append((m, title, num, False))
        # 标记数量不足标题标注的篇数 → 用相似标题行补切
        if expected >= 2 and base and len(marks) < expected:
            cand = _piece_title_candidates(texts, base, expected)
            mark_set = set(marks)
            for ci, _sc in cand:
                # 紧邻显式标记（前 3 行内或后 2 行内）的标题行不是新篇界：
                # 篇界由显式标记承担，标题行是篇内重复的题名或汇编标题
                if any(abs(ci - mk) <= 3 for mk in mark_set):
                    continue
                bounds.append((ci, _strip_count_suffix(texts[ci]), None, True))
        bounds.sort(key=lambda b: b[0])
        # 同一位置去重（显式标记优先）
        seen, dedup = set(), []
        for b in bounds:
            if b[0] in seen:
                continue
            seen.add(b[0])
            dedup.append(b)
        bounds = dedup
        # 清理：某边界切出的篇过短（<30 字）说明切多了——
        # 优先删候选边界，其次删后一个边界
        changed = True
        while changed and len(bounds) >= 2:
            changed = False
            for bi in range(len(bounds)):
                s = bounds[bi][0]
                e = bounds[bi + 1][0] if bi + 1 < len(bounds) else len(texts)
                seg = texts[s + 1:e]
                if sum(len(x) for x in seg) >= 30 or any(
                        ("。" in x or "！" in x or "？" in x) for x in seg):
                    continue
                drop = None
                if bounds[bi][3]:
                    drop = bi
                elif bounds[bi + 1][3]:
                    drop = bi + 1
                elif bounds[bi][0] > 0 and bi > 0:
                    drop = bi
                if drop is not None:
                    del bounds[drop]
                    changed = True
                    break
    elif expected >= 2 and base:
        cand = _piece_title_candidates(texts, base, expected)
        if len(cand) >= 2:
            bounds = [(ci, _strip_count_suffix(texts[ci]), None, True)
                      for ci, _sc in cand]
            # 清理：某边界切出的篇过短（<30 字）说明切多了，删掉该边界
            changed = True
            while changed and len(bounds) >= 2:
                changed = False
                for bi in range(len(bounds)):
                    s = bounds[bi][0]
                    e = bounds[bi + 1][0] if bi + 1 < len(bounds) else len(texts)
                    seg = texts[s + 1:e]
                    # 篇过短且不含句读（纯标题行残留）才视为切多了
                    if sum(len(x) for x in seg) >= 30 or any(
                            ("。" in x or "！" in x or "？" in x) for x in seg):
                        continue
                    del bounds[bi]
                    changed = True
                    break

    if not bounds:
        return [], [("", texts, blanks)], []

    # 篇界之前的内容（汇编标题、说明等）作为整个合集的文头单独返回
    first = bounds[0][0]
    preface = [texts[k] for k in range(first)]
    preface_bl = [blanks[k] for k in range(first)]
    parts = []
    nums = []
    for bi, (s, title, num, _c) in enumerate(bounds):
        e = bounds[bi + 1][0] if bi + 1 < len(bounds) else len(texts)
        seg = [texts[k] for k in range(s + 1, e)]
        bl = [blanks[k] for k in range(s + 1, e)]
        parts.append((title, seg, bl))
        nums.append(num)
    return (preface, preface_bl), parts, nums


def make_theme(q, limit=26):
    """从问句生成小节主题标签。

    处理顺序：去「问：」前缀 → 去「请问/请您…」套话 → 按第一个问号截断
    → 超长时按标点截断 → 若结果只剩「最后/另外」等空泛词，改用实义短语。
    """
    def strip_court(t):
        for p in STOP_PREFIX:
            if t.startswith(p):
                return t[len(p):].strip()
        return t

    base = _strip_qa_prefix(q).strip()
    t = re.split(r"[？?]", base)[0].strip()
    t = strip_court(t).rstrip("。、，；：").strip()

    if len(t) > limit:
        cut = t[:limit + 1]
        for sp in ["，", "、", "；", "："]:
            if sp in cut:
                cut = cut[:cut.rindex(sp)]
        t = cut.rstrip("。、，；：").strip()

    # 截断后只剩空泛词：回退到问句的实义短语
    if t in ("最后", "另外", "还有", "此外", "补充", "同时", "", "问"):
        parts = [x.strip() for x in re.split(r"[，,、]", base) if x.strip()]
        cand = ""
        for x in parts:
            x = strip_court(re.split(r"[？?]", x)[0]).strip()
            if len(x) > len(cand):
                cand = x
        if len(cand) > len(t):
            t = cand[:limit].rstrip("。、，；：").strip()
    return t or "问答"


def restore_question(q, candidates):
    """被截断的问题用候选完整问句补全（前缀匹配）。

    docx 里有时把问题行在 20 字左右截断（当作标题排版），完整问句只出现在
    文头目录或下一段中。这里用前缀匹配把它还原成完整问句。
    """
    q = q.strip()
    if q.endswith(("？", "?")):
        return q
    key = _strip_qa_prefix(q)
    if len(key) < 6:
        return q
    best = None
    head4 = key[:4]
    for cand in candidates:
        ck = _strip_qa_prefix(cand)
        if ck == key or len(ck) <= len(key):
            continue
        # 宽松前缀匹配：候选以截断串开头，或两者前 4 字一致且截断串是其子串
        if ck.startswith(key[: max(6, len(key) - 6)]) or (len(head4) >= 4 and ck.startswith(head4)):
            if best is None or len(ck) < len(best):  # 取最短的更长问句，避免误配
                best = ck
    if best:
        return "问：" + best
    # 未能补全：保持原样，但补齐「问：」前缀，保证体例统一
    return q if RE_QUESTION.match(q) else "问：" + q


def parse_qa(texts, doc_title="", file_stem=""):
    """解析问答体例。返回 (intro_paras, qas, candidates)。

    qas: [{"q": str, "a": [str,...]}, ...]
    candidates: 全文出现的所有「问：…」文本（含目录块），用于问题补全

    doc_title / file_stem 用于识别「首行其实是文档标题」的情况：
    有些 docx 的一级标题本身写成一句问题（如「问：最后，请您谈谈…？」），
    若不加区分会被当成第 1 问，把后续真实问答全部并入它的答案里。
    """
    body, toc_qs = split_head_toc(texts)
    paras = demote_fake_headings(body)
    candidates = [_strip_qa_prefix(t) for t in paras
                  if RE_QUESTION.match(t) or RE_QUESTION_ALT.match(t)]
    candidates = [c for c in candidates if c] + [q for q in toc_qs if q]

    # 首段「问：…」是文档标题还是真问题？用结构信号判断：
    #   真实问答的首段问句后面应当紧跟「答：」或至少没有长段导语；
    #   若其后先出现「标题样式段落」（与文档标题同名、或极短的居中标题行），
    #   说明它是被排版成标题的整句标题，记为导语、不占第1问。
    def _is_titleish(txt):
        t2 = txt.strip()
        if doc_title and t2 == doc_title:
            return True
        if file_stem and t2 == file_stem:
            return True
        return False

    if paras and (RE_QUESTION.match(paras[0]) or RE_QUESTION_ALT.match(paras[0])):
        nxt = paras[1] if len(paras) > 1 else ""
        # 下一段不是答案、且下一段是标题样式/另一段问句前的标题行 → 首段为标题
        if not (RE_ANSWER.match(nxt) or RE_ANSWER_ALT.match(nxt)):
            if _is_titleish(nxt) or (nxt and len(nxt) <= 60 and nxt == doc_title):
                paras = paras[1:]

    intro, qas, cur = [], [], None
    for idx, t in enumerate(paras):
        mq = RE_QUESTION.match(t) or RE_QUESTION_ALT.match(t)
        ma = RE_ANSWER.match(t) or RE_ANSWER_ALT.match(t)
        if mq:
            qtext = mq.group(1).strip()
            cur = {"q": ("问：" + qtext) if qtext else t, "a": []}
            qas.append(cur)
        elif ma:
            if cur is None:
                # 没有前置问题的答案：仍独立成节，不丢内容
                cur = {"q": "", "a": []}
                qas.append(cur)
            cur["a"].append(t)
        elif cur is None:
            intro.append(t)
        else:
            cur["a"].append(t)

    for qa in qas:
        if qa["q"]:
            qa["q"] = restore_question(qa["q"], candidates)
    # 丢弃既无问题又无答案的空节
    qas = [qa for qa in qas if qa["q"] or qa["a"]]

    # 同一问题重复出现（目录/回目残留）时合并为一节，答案按序拼接
    merged, seen = [], {}
    for qa in qas:
        key = re.sub(r"[\s。？?]", "", _strip_qa_prefix(qa["q"])) if qa["q"] else ""
        if key and key in seen:
            seen[key]["a"].extend(qa["a"])
            continue
        if key:
            seen[key] = qa
        merged.append(qa)

    # 清理：既无答案、问句又短到不成句的残缺节（多为标题被截断的排版残留），
    # 其文字通常已在相邻问句里，直接丢弃；有答案的一律保留。
    merged = [qa for qa in merged
              if qa["a"] or len(_strip_qa_prefix(qa["q"])) >= 12]
    return intro, merged, candidates


