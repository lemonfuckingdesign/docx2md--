#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx2md —— 批量把法规/制度/解读类 .docx 转成适合 Dify 知识库入库的 Markdown。

用法：
    python docx2md.py <输入文件夹> [输出文件夹] [flat]
    python docx2md.py <输入文件夹> -o <输出文件夹> --flat -v --log-file run.log

    不指定输出文件夹时，默认写到 <输入文件夹>/dify_md/
    旧用法第三个参数传 flat 可启用扁平模式（等同 --flat）

支持的四种体例（自动识别，优先级从上到下）：
  1. 问答体例     —— 「问：…／答：…」答记者问、访谈、政策问答
  2. 章条体例     —— 「第X章／第X条」法规
  3. 公文编号体例 —— 「一、／（一）／1.」多级编号
  4. 无编号文本   —— 按段落与语义长度切分兜底
  ※ 一个 docx 含多篇时：以「篇N」行为界切分，合集输出，不再拆成多个文件

特性：
  * 一个 docx → 一个同名 .md（以文件名为总标题），多篇也合并在同一个文档里
  * 落款日期通吃两种写法：2025年3月1日 / 二〇二五年三月一日
  * 标题降级去重、问题还原、内容覆盖率校验，正文文字零改写
  * YAML front matter 写入 title / doc_type / issuer 等字段，便于 Dify 元数据筛选
  * 处理完输出「处理清单.csv」，标注体例、问答数、内容覆盖率

依赖：pip install -r requirements.txt（仅需 python-docx，Python >= 3.8）
完整文档：https://github.com/lemonfuckingdesign/docx2md
"""
from cli import main


if __name__ == "__main__":
    main()
