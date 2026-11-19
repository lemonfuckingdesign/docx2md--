#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中文数字转换、结构识别、元数据提取。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
"""

import re
from config import CN_DIGITS, CN_YEAR_DIGIT, LABEL_HEADING_MIN_LEN, RE_ARTICLE, RE_CHAPTER, RE_DATE_LINE, RE_DOCNO, RE_L1, RE_L2, RE_L3, RE_QUESTION, RE_QUESTION_ALT, RE_REDHEAD, RE_SIGN_DATE, _TITLE_AS_BODY_MAX_LEN


def cn_to_int(s: str) -> int:
    """把『一 / 十 / 十二 / 二十三』等中文数字转成 int。"""
    s = s.strip()
    if not s:
        return 0
    if s.isdigit():
        return int(s)
    if s == "十":
        return 10
    if s.startswith("十"):
        return 10 + CN_DIGITS.get(s[1], 0)
    if "十" in s:
        left, _, right = s.partition("十")
        return CN_DIGITS.get(left, 0) * 10 + (CN_DIGITS.get(right, 0) if right else 0)
    total = 0
    for ch in s:
        total = total * 10 + CN_DIGITS.get(ch, 0)
    return total


def int_to_cn(n: int) -> str:
    """1-99 转中文章节号，保证章号与原文章号风格一致。"""
    d = "零一二三四五六七八九"
    if n < 10:
        return d[n]
    if n == 10:
        return "十"
    if n < 20:
        return "十" + d[n % 10]
    return d[n // 10] + "十" + (d[n % 10] if n % 10 else "")


def is_heading_like(t):
    """判断一行是否像标题：短、无句号、不是套话起始。"""
    if len(t) > 40:
        return False
    if re.search(r"[。；，,]$", t):
        return False
    if re.match(r"^(为|根据|按照|现将|特此|经研究|为深入|为进一步)", t):
        return False
    return True


def _title_as_body(title):
    """判断编号「标题」是否应还原为正文段落而不是小节标题。"""
    t = title.strip()
    if not t:
        return False
    # 问答体例的题干标题（### 第N问 …… / N、问：……）始终保留
    if re.match(r"^第[一二三四五六七八九十百\d]+问", t):
        return False
    if "问：" in t or "问:" in t:
        return False
    if "。" in t or t.endswith(("；", ";")):
        return True
    # 阿拉伯数字括号条目（(1)…／（1)…）是答话/论述里的列举句，
    # 不是章节：不做分段标记（汉字序号（一）(一) 不受影响）
    if re.match(r"^[（(]\d+[)）]", t):
        return True
    if len(t) > _TITLE_AS_BODY_MAX_LEN:
        return True
    return False


def _label_heading_ok(label):
    """兜底标签能否作为小标题输出。

    太短的标签（如「加油」「转眼间」「第二」「这次疫情里」等半句话）
    不应生成分级标题，否则正文会被切得过碎。
    """
    label = label.strip()
    if len(label) < LABEL_HEADING_MIN_LEN:
        return False
    return is_heading_like(label)


def level_marker(t):
    """识别通用公文的层级标记。返回 (层级, 编号) 或 (0, None)。

    只在行首出现编号时才认为是标题行，避免把正文中的「1.」误判。
    """
    if RE_CHAPTER.match(t) or RE_ARTICLE.match(t):
        return 0, None
    m = RE_L2.match(t)
    if m:
        return 2, m.group(1)
    m = RE_L3.match(t)
    if m:
        return 3, (m.group(1) or m.group(2) or m.group(3))
    m = RE_L1.match(t)
    if m:
        return 1, m.group(1)
    return 0, None


def parse_sign_date(t):
    """把落款日期（阿拉伯 / 中文数字）解析成 YYYY-MM-DD，失败返回空串。"""
    if not t:
        return ""
    s = re.sub(r"\s+", "", t)
    m = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$", s)
    if m:
        y, mo, d = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    m = re.match(r"^([〇零一二三四五六七八九]{4})年([一二三四五六七八九十]{1,2})月"
                 r"([一二三四五六七八九十]{1,3})日$", s)
    if m:
        y = "".join(CN_YEAR_DIGIT.get(c, "") for c in m.group(1))
        mo, d = cn_to_int(m.group(2)), cn_to_int(m.group(3))
        if len(y) == 4 and 1 <= mo <= 12 and 1 <= d <= 31:
            return f"{y}-{mo:02d}-{d:02d}"
    m = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$", s.replace("〇", "0"))
    return ""


def detect_meta(texts):
    """从开头段落里提取元数据。"""
    meta = {}
    title = None
    # 取“第一条/第一章/第一个编号标题/第一个问：”之前的所有段落作为文头区
    head = []
    for t in texts:
        if (RE_CHAPTER.match(t) or RE_ARTICLE.match(t) or level_marker(t)[0]
                or RE_QUESTION.match(t) or RE_QUESTION_ALT.match(t)):
            break
        head.append(t)
    # 文头区可能包含多个“标题”行（有的 docx 首行是「标题（日期）」），
    # 取最长的那个作为正式标题，其余作日期行
    date_lines = [t for t in head if RE_DATE_LINE.match(t)]
    title_cands = [t for t in head if t not in date_lines]
    if title_cands:
        # 优先挑「关于……的通知/意见/办法/规定」这类标准公文标题
        official = [t for t in title_cands if re.search(
            r"^(关于|印发|转发).{2,}的(通知|意见|办法|规定|细则|方案|决定|批复|通报|报告|纪要|规则|制度|规范|指引|要点|计划|意见)$",
            re.sub(r"[（(][^（()）]*[）)]$", "", t))]
        if official:
            title = max(official, key=len)
        else:
            heads = [t for t in title_cands if is_heading_like(t)]
            title = max(heads, key=len) if heads else max(title_cands, key=len)
        title = re.sub(r"[（(][^（()）]*\d{4}年[^（()）]*[）)]\s*$", "", title).strip()
        title = re.sub(r"[（(][^（()）]*[）)]\s*$", "", title).strip() if len(title) > 30 else title
    meta["title"] = title or ""
    meta["head"] = head
    meta["date_lines"] = date_lines

    # 发文字号（保留发文机关简称前缀，如「X宣发〔2026〕15号」）
    for t in texts[:12]:
        m = RE_DOCNO.search(t)
        if m:
            meta["docno"] = re.sub(r"\s+", "", m.group(0))
            break

    # 红头文件的发文机关标志，作为 issuer 的兜底
    redhead = ""
    for t in texts[:5]:
        m = RE_REDHEAD.match(t)
        if m:
            redhead = t
            break
    meta["redhead"] = redhead

    joined = " ".join(date_lines)
    for key, pat in (
        ("date_original", r"(\d{4})年(\d{1,2})月(\d{1,2})日.{0,20}?制定"),
        ("date_revised", r"(\d{4})年(\d{1,2})月(\d{1,2})日.{0,20}?修订"),
        ("date_issued", r"(\d{4})年(\d{1,2})月(\d{1,2})日(?:发布|印发)"),
        ("date_effective", r"自(\d{4})年(\d{1,2})月(\d{1,2})日起施行"),
    ):
        m = re.search(pat, joined)
        if m:
            y, mo, d = m.groups()
            meta.setdefault(key, f"{y}-{int(mo):02d}-{int(d):02d}")

    # 落款日期（公文末尾单独一行的 2025年3月1日）
    if "date_issued" not in meta:
        for t in reversed(texts[-6:]):
            if RE_SIGN_DATE.match(t):
                d = parse_sign_date(t)
                if d:
                    meta["date_issued"] = d
                break

    # 制发机关：去掉日期与「制定/修订/发布」等动词后剩下的机构名
    issuer = ""
    for line in date_lines:
        body = re.sub(r"\d{4}年\d{1,2}月\d{1,2}日", "", line.strip("（）() "))
        body = re.sub(r"(制定|修订|发布|印发|通过)", "", body)
        cands = [x.strip(" 、，,　") for x in body.split() if x.strip(" 、，,　")]
        cands = [x for x in cands if 2 <= len(x) <= 40]
        if cands:
            issuer = max(cands, key=len)
            break
    # 兜底：从正文首段里抽「XX负责人就…」「XX办公室负责人」这类主体名
    if not issuer:
        for t in texts[:12]:
            m = re.search(r"([\u4e00-\u9fa5]{2,20}?(?:部|委|办|局|院|署|中心|小组办公室))"
                          r"(?:负责人|有关负责人|新闻发言人)", t)
            if m:
                issuer = m.group(1)
                break
    # 兜底：落款机关（正文末尾、日期之前那一行），再兜底用红头标志
    if not issuer:
        for i, t in enumerate(texts[-6:]):
            if RE_SIGN_DATE.match(t):
                j = len(texts) - 6 + i
                if j > 0 and 2 <= len(texts[j - 1]) <= 20 and not t == texts[j - 1]:
                    issuer = texts[j - 1]
                break
    if not issuer and redhead:
        issuer = redhead
    meta["issuer"] = issuer
    return meta


def extract_related_doc(title, texts):
    """抽取被解读的法规名，如「《中国共产党农村工作条例》」。"""
    m = re.search(r"《[^》]{2,60}》", title or "")
    if m:
        return m.group(0)
    m = re.search(r"《[^》]{2,60}》", " ".join(texts[:6]))
    return m.group(0) if m else ""


def extract_date_context(texts, meta):
    """推断政策解读的时间语境（YYYY-MM）。"""
    if meta.get("date_issued"):
        return meta["date_issued"][:7]
    blob = " ".join(texts[:10])
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*\d{1,2}\s*日", blob)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.search(r"(\d{4})\s*年\s*(\d{1,2})\s*月", blob)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    return ""


