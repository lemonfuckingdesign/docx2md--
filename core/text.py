#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx 读取与文本清洗（段落修复、文头目录剥离、粘连拆分）。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
"""

from docx import Document
from pathlib import Path
import re
from config import RE_GLUED_TITLE, RE_HYPHEN_CJK, RE_ITEM_AFTER_PUNCT, RE_NUM_RUN, RE_NUM_TITLE, RE_PIECE_INLINE, RE_QUESTION, RE_QUESTION_ALT, RE_TITLE_BEFORE_NUM, RE_TITLE_TITLE, RE_TRAILING_TITLE, RE_WRAP_TERMINAL
from .recognize import is_heading_like


def read_paragraphs(path: Path):
    """读取 docx 段落文本，尽量保留空段落作为分段线索。

    返回 (texts, blank_flags)：texts 为非空段落，blank_flags 记录每段之前是否有空行。
    """
    doc = Document(str(path))
    texts, blanks = [], []
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            if texts:
                blanks[-1] = True  # 标记上一段之后出现了空行
            continue
        texts.append(t)
        blanks.append(False)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    t = p.text.strip()
                    if t:
                        texts.append(t)
                        blanks.append(True)
    return texts, blanks


def demote_fake_headings(texts):
    """降级「假标题」：某段文本几乎完整地重复出现在紧邻的下一段里。

    docx 里常见排版是「一行短标题 + 一段正文」，若那行短文本其实是正文的
    开头（被自动排版误当成标题层级），直接当作标题输出就会：
      * 标题被 Dify 当作 chunk 边界，正文首段被割裂；
      * 校验时该行在标题与正文各出现一次，统计口径混乱。

    这里返回去重后的段落序列：假标题被丢弃（其文字已在下一段中完整保留），
    真标题原样保留。

    判定规则（任一成立即为假标题）：
      1) 下一段以该行开头（严格前缀），且该行长度 ≥ 6；
      2) 该行去掉「问：/答：」前缀后，仍以 ≥85% 的字符相似度出现在下一段中。
    """
    out = []
    i = 0
    n = len(texts)
    while i < n:
        cur = texts[i]
        nxt = texts[i + 1] if i + 1 < n else None
        if nxt is not None and _is_dup_heading(cur, nxt):
            # 丢弃假标题，只保留正文段
            i += 1
            continue
        out.append(cur)
        i += 1
    return out


def _strip_qa_prefix(t):
    """去掉行首的「问：/答：」前缀，用于比较标题与正文。"""
    return re.sub(r"^(?:问|答|记者|提问人|主持人|回答|答复|回应)\s*[：:]\s*", "", t).strip()


def _is_dup_heading(cur, nxt):
    """判断 cur 是否为 nxt 的“假标题”（同一句话被重复成标题+正文）。"""
    if not cur or not nxt:
        return False
    if cur == nxt:
        return True
    # 只考虑短行（标题候选），长段落不可能是标题
    if len(cur) > 60:
        return False
    body = _strip_qa_prefix(cur).rstrip("。：:；;，,")
    if len(body) < 6:
        return False
    # 规则 1：严格前缀
    nxt_clean = _strip_qa_prefix(nxt)
    if nxt_clean.startswith(body):
        return True
    # 规则 2：该行内容有 85% 以上出现在下一段前 1.5 倍长度内（容忍标点/空格差异）
    window = nxt_clean[: int(len(body) * 1.5) + 10]
    if not window:
        return False
    hit = sum(1 for ch in body if ch in window)
    return hit / len(body) >= 0.85


def split_head_toc(texts):
    """剥离「文头内容 / 目录 / 内容提要」之类的目录块。

    有些 docx（或本工具早期产物）在正文前贴一份「文头内容：」列表，
    里面是「> 问：……」形式的目录条目。这些条目如果混入正文解析，
    会被当成真实问答，导致问答错位、重复。

    返回 (正文段落, 目录问题候选列表)。
    """
    toc_markers = ("文头内容", "内容目录", "目录：", "篇目")

    # 第一轮：先用显式标记（**文头内容**：）定位目录块
    body, toc_qs = [], []
    in_toc = False
    for t in texts:
        s = t.strip()
        bare = s.lstrip(">*-# ").strip()
        if re.match(r"^\*{0,2}(文头内容|内容目录|篇目)\*{0,2}\s*[：:]?$", bare):
            in_toc = True
            continue
        if in_toc:
            if s.startswith(">") or RE_QUESTION.match(bare) or RE_QUESTION_ALT.match(bare):
                if RE_QUESTION.match(bare) or RE_QUESTION_ALT.match(bare):
                    toc_qs.append(_strip_qa_prefix(bare))
                continue
            if not s:
                continue
            in_toc = False
        if any(m in bare for m in toc_markers) and len(bare) < 20:
            continue
        body.append(t)

    # 第二轮：识别「无标记的 > 引用块」，块内问句 ≥2 个即判为目录
    # （某些 docx 直接把旧 md 的目录粘进正文，没有「文头内容」字样）
    cleaned, i, n = [], 0, len(body)
    while i < n:
        if body[i].strip().startswith(">"):
            j = i
            while j < n and body[j].strip().startswith(">"):
                j += 1
            block = body[i:j]
            q_in_block = [b for b in block
                          if RE_QUESTION.match(b.lstrip(">*-# ").strip())
                          or RE_QUESTION_ALT.match(b.lstrip(">*-# ").strip())]
            if len(q_in_block) >= 2:
                for b in q_in_block:
                    toc_qs.append(_strip_qa_prefix(b.lstrip(">*-# ").strip()))
                i = j
                continue
            cleaned.extend(block)
            i = j
            continue
        cleaned.append(body[i])
        i += 1
    return cleaned, toc_qs


def split_head_toc_aligned(texts, blanks):
    """split_head_toc 的保序号版本：同时维护 blanks，保证与段落一一对应。

    做法：先跑 split_head_toc 得到保留后的正文序列，再按顺序把原 blanks
    映射到被保留的段落上（用文本相等 + 顺序游标匹配，容忍重复文本）。
    """
    kept, _toc_qs = split_head_toc(list(texts))
    texts = list(texts)
    if len(kept) == len(texts):
        return kept, list(blanks)
    out_b, cursor = [], 0
    for t in kept:
        j = cursor
        while j < len(texts) and texts[j] != t:
            j += 1
        if j < len(texts):
            out_b.append(blanks[j])
            cursor = j + 1
        else:
            out_b.append(False)
    return kept, out_b


def _merge_wrapped_paras(texts, blanks):
    """消除 docx 硬换行造成的断句。

    同一段落在 docx 里被换行拆成两段时，前段往往没有句末标点
    （如「……例如：机器的发明与」+「使用导致了第一次工业革命。」）。
    规则：前一段 ≥15 字、未以句末标点收尾、且两段之间无空行分隔
    → 判定为换行断句，合并为一段。
    短行（说话人名、引导语等 <15 字）不参与合并，避免误粘。
    """
    i = 0
    while i < len(texts) - 1:
        cur = texts[i].strip()
        nxt = texts[i + 1].strip()
        if (cur and nxt and len(cur) >= 15
                and not cur.endswith(RE_WRAP_TERMINAL)
                and not is_heading_like(cur)
                and not blanks[i]):
            texts[i] = cur + nxt
            texts.pop(i + 1)
            blanks.pop(i + 1)
            continue
        i += 1
    return texts, blanks


def _split_double_title(p):
    """拆开「××制度××制度」式双标题粘连；后面紧跟连接词的不拆（防误伤正文）。"""
    def _repl(m):
        nxt = p[m.end():m.end() + 1]
        if nxt in ("的", "和", "与", "及", "或", "、", "，", "。"):
            return m.group(0)
        return m.group(0) + "\n"
    return RE_TITLE_TITLE.sub(_repl, p)


def _split_glued_paras(texts, blanks):
    """拆分粘连段落：合集类 docx 常把小标题、多条编号条目粘进同一自然段。

    处理（按序）：
      * 修复「共-产-党员」式拆字；
      * 句末标点后的「（1）（2）」条目各自成行；
      * 「……。结对共建帮扶制度1、……」粘连的制度名拆为独立行；
      * 「一、支部大会制度1、会议时间」这类「制度名+编号」拆开；
      * 段中粘连的「×× [篇N]」篇标记拆为独立行（供合集切分识别）；
      * 行首起、≥3 个「N、」连续编号的长段拆为逐条成行。
    拆出的新行之间不计空行（blanks 用 False 填充，保持与原段一致的段间关系）。
    """
    out_t, out_b = [], []
    for idx, raw in enumerate(texts):
        t = raw.strip()
        if t:
            t = RE_HYPHEN_CJK.sub("", t)
        if not t:
            out_t.append(raw)
            out_b.append(blanks[idx])
            continue
        parts = RE_ITEM_AFTER_PUNCT.sub("\n", t).split("\n")
        frags = []
        for p in parts:
            p = _split_double_title(p)
            p = RE_GLUED_TITLE.sub(lambda m: "\n" + m.group(1) + "\n", p)
            p = RE_TITLE_BEFORE_NUM.sub(r"\1\n", p)
            p = RE_NUM_TITLE.sub(lambda m: "\n" + m.group(1) + "\n", p)
            # 段中出现的「×× [篇N]」篇标记：前后拆开成独立行
            if not RE_PIECE_INLINE.fullmatch(p.strip()):
                p = RE_PIECE_INLINE.sub(lambda m: "\n" + m.group(1) + "\n", p)
            frags.extend(x for x in p.split("\n") if x.strip())
        res = []
        for p in frags:
            # 句尾粘连的制度名标题：连环拆（「…内容。结对共建帮扶制度村企共建制度」）
            tails = []
            while True:
                m = RE_TRAILING_TITLE.match(p)
                if not m:
                    break
                tails.insert(0, m.group(2))
                p = m.group(1).strip()
            if tails:
                res.append(p)
                res.extend(tails)
                continue
            marks = list(RE_NUM_RUN.finditer(p))
            if len(marks) >= 3:
                idxs = [m.start() for m in marks]
                if idxs[0] > 0:
                    res.append(p[:idxs[0]].strip())
                res.extend(p[s:(idxs[k + 1] if k + 1 < len(idxs) else len(p))].strip()
                           for k, s in enumerate(idxs))
            else:
                res.append(p)
        if len(res) == 1:
            out_t.append(res[0])
            out_b.append(blanks[idx])
        else:
            out_t.extend(res)
            out_b.extend([False] * (len(res) - 1) + [blanks[idx]])
    return out_t, out_b


def normalize_qa_prefix(texts):
    """剥掉每段行首的「问：/答：」标签，供校验使用。

    输出 md 里问答标签统一为「**问：…**」「**答：**」，标签文字只出现一次；
    而原文每问每答都可能带标签。若按原文统计，标签会被算作“缺失”，
    造成覆盖率虚低（典型的假缺失）。校验时统一剔除标签再比对。
    """
    return [re.sub(r"^(问|答|记者|提问人|主持人|回答|答复|回应)\s*[：:]\s*", "", t.strip())
            for t in texts]


