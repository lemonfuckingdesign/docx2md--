#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""全局常量与可变状态（FLAT_MODE 访问器）。不依赖任何本地模块。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
"""

import re

# ---- 常量（逐字搬运）----
CN_DIGITS = {c: i + 1 for i, c in enumerate("一二三四五六七八九")}


RE_CHAPTER = re.compile(r"^第([一二三四五六七八九十百零〇\d]+)章[\s　]*(.*)$")


RE_ARTICLE = re.compile(r"^第([一二三四五六七八九十百零〇\d]+)条[\s　]*(.*)$")


RE_DATE_LINE = re.compile(r"^[（(].*?(制定|修订|发布|通过|施行).*?[）)]$")


RE_CHAPTER_INDEX = re.compile(r"^(目\s*录|目\s*次)$")


RE_L1 = re.compile(r"^([一二三四五六七八九十百]+)[、．.][\s　]*(.*)$")


RE_L2 = re.compile(r"^[（(]([一二三四五六七八九十百]+)[)）][、．.]?[\s　]*(.*)$")


RE_L3 = re.compile(
    r"^(?:([0-9]{1,2})[、．.](?![0-9])|[（(]([0-9]{1,2})[)）]|([①-⑳]))[\s　]*(.*)$")


RE_DOCNO = re.compile(r"([\u4e00-\u9fa5A-Za-z]{1,12})?[〔\[［]\s*(\d{4})\s*[〕\]］]\s*(\d+)\s*号")


RE_SIGN_DATE = re.compile(
    r"^(?:\d{4}|[〇零一二三四五六七八九]{4})\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日$"
    r"|^(?:\d{4}|[〇零一二三四五六七八九]{4})\s*年\s*[一二三四五六七八九十]{1,2}\s*月"
    r"\s*[一二三四五六七八九十]{1,3}\s*日$"
)


RE_ADDRESSEE = re.compile(r"^[^。；，,]{1,40}[：:]$")


RE_REDHEAD = re.compile(r"^([\u4e00-\u9fa5（）()]{2,20})(文件|党组|党委)$")


RE_QUESTION = re.compile(r"^问\s*[：:]\s*(.*)$")


RE_ANSWER = re.compile(r"^答\s*[：:]\s*(.*)$")


RE_QUESTION_ALT = re.compile(r"^(?:记者|提问人|主持人)\s*[：:]\s*(.*)$")


RE_ANSWER_ALT = re.compile(r"^(?:回答|答复|回应)\s*[：:]\s*(.*)$")


LABEL_HEADING_MIN_LEN = 6


_TITLE_AS_BODY_MAX_LEN = 40


CN_YEAR_DIGIT = {"〇": "0", "零": "0", "一": "1", "二": "2", "三": "3", "四": "4",
                 "五": "5", "六": "6", "七": "7", "八": "8", "九": "9"}


RE_PIECE_MARK = re.compile(
    r"^第?\s*篇\s*([一二三四五六七八九十百零〇\d]+)\s*[：:、．.]?\s*(.*)$"
)


RE_PIECE_MARK_TAIL = re.compile(
    r"^([^。！？，,]{2,40}?)\s*第?\s*篇\s*([一二三四五六七八九十百零〇\d]+)\s*[：:、．.]?\s*$"
)


RE_PIECE_MARK_TAIL_CN = re.compile(
    r"^([^。！？，,]{2,40}?)\s*第\s*([一二三四五六七八九十百零〇\d]+)\s*篇\s*[：:、．.]?\s*$"
)


RE_PIECE_MARK_BRACKET = re.compile(
    r"^([^。！？，,]{2,40}?)\s*[【\[（(]\s*(?:第\s*([一二三四五六七八九十百零〇\d]+)\s*篇"
    r"|篇\s*([一二三四五六七八九十百零〇\d]+))\s*[】\]）)]\s*$"
)


STOP_PREFIX = ["请问，", "请问：", "请问", "请您介绍一下", "请您谈谈", "请您", "请介绍一下",
               "请谈谈", "请你", "请介绍", "请"]


RE_INLINE_SPLIT = re.compile(
    r"(?<=[。！？；?!；])\s*(?="
    r"第[一二三四五六七八九十百零〇\d]+[条章]"   # 第X条 / 第X章
    r"|[一二三四五六七八九十百]+[、．.]"          # 一、 二、
    r"|[（(][一二三四五六七八九十百\d]+[)）][、．.]?"  # （一）（2）
    r"|\d{1,2}[、．.](?!\d)"                     # 1. 2、
    r"|[①-⑳]"                                    # ①
    r"|[一二三四五六七八九十]+是"                 # 一是 二是（答问常用列举）
    r")"
)


RE_BARE_WEN = re.compile(
    r"^([\s　]*[一二三四五六七八九十百\d]+[、.．]?\s*)问\s*[:：]?\s*$")


RE_WRAP_TERMINAL = ("。", "！", "？", "；", "：", "…", "”", "）", ")", "】")


RE_HYPHEN_CJK = re.compile(r"(?<=[\u4e00-\u9fa5])-+(?=[\u4e00-\u9fa5])")


RE_ITEM_AFTER_PUNCT = re.compile(
    r"(?<=[。；;！？])\s*(?=[（(][一二三四五六七八九十百\d]{1,3}[)）]"
    r"|\d{1,2}[、.．](?!\d))")


RE_GLUED_TITLE = re.compile(
    r"(?<=[。；;！？：])\s*([\u4e00-\u9fa5]{2,14}制度)\s*(?=[（(]?[\d一二三四五六七八九十]{1,2}[、.．:：])")


RE_TITLE_BEFORE_NUM = re.compile(
    r"([\u4e00-\u9fa5]{2,14}制度)(?=\d{1,2}[、.．])")


RE_PIECE_INLINE = re.compile(
    r"([\u4e00-\u9fa5]{2,20}\s*[【\[]\s*篇\s*\d+\s*[】\]])")


RE_NUM_TITLE = re.compile(
    r"([一二三四五六七八九十]{1,3}、[\u4e00-\u9fa5]{2,14}制度)(?=\s*\d{1,2}[、.．])")


RE_NUM_RUN = re.compile(r"(?<![\d（(])\d{1,2}、")


RE_TRAILING_TITLE = re.compile(
    r"^(.+[。；;！？])([\u4e00-\u9fa5]{2,14}制度)$")


RE_TITLE_TITLE = re.compile(r"[\u4e00-\u9fa5]{2,14}制度(?=[\u4e00-\u9fa5]{2,14}制度)")


TINY_SECTION_BODY_LEN = 40


CLEAN_FOR_DIFF = re.compile(
    r"[#>*`|：:（）()【】\[\]、。？！?!,.；;\u201c\u201d\u2018\u2019'\"\s\-—\d]")


# ---- FLAT_MODE 访问器（原为模块级全局，改为函数访问以避免 from-import 值拷贝）----
_FLAT_MODE = False


def set_flat_mode(enabled: bool) -> None:
    """设置 Dify 扁平模式（### 及以下标题还原为正文）。"""
    global _FLAT_MODE
    _FLAT_MODE = bool(enabled)


def is_flat_mode() -> bool:
    """读取当前是否为 Dify 扁平模式。"""
    return _FLAT_MODE
