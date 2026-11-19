#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""命令行编排入口。

由原单文件 docx2md.py 机械拆分而来；函数体逐字搬运，未改动解析算法。
命令行参数用 argparse 管理，同时兼容旧的三位置参数用法。
"""

import argparse
import logging
import sys
from pathlib import Path

from engine import convert

__version__ = "1.0.0"

logger = logging.getLogger("docx2md")


def build_parser():
    """构建命令行参数解析器（兼容旧用法 + 提供新选项）。"""
    p = argparse.ArgumentParser(
        prog="docx2md",
        description="批量把法规/制度/解读类 .docx 转成适合 Dify 知识库入库的 Markdown。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python docx2md.py <输入文件夹>\n"
            "  python docx2md.py <输入文件夹> <输出文件夹>\n"
            "  python docx2md.py <输入文件夹> -o <输出文件夹> --flat -v\n"
            "\n兼容旧用法：python docx2md.py <输入文件夹> [输出文件夹] [flat]\n"
        ),
    )
    p.add_argument("src", help="输入文件夹（递归扫描 *.docx）")
    p.add_argument("dst", nargs="?", default=None,
                   help="输出文件夹（默认 <输入文件夹>/dify_md）")
    p.add_argument("legacy_flat", nargs="?", default="",
                   help="兼容旧用法：传 flat 启用扁平模式（忽略大小写），其他值忽略")
    p.add_argument("-o", "--output", dest="dst_opt", default=None,
                   help="输出文件夹（与位置参数同时给出时以本项为准）")
    p.add_argument("--flat", action="store_true",
                   help="Dify 扁平模式：### 及以下标题还原为正文")
    p.add_argument("-v", "--verbose", action="store_true",
                   help="输出 DEBUG 级日志")
    p.add_argument("--log-file", default=None,
                   help="除控制台外，额外把日志写入指定文件（UTF-8）")
    p.add_argument("--version", action="version", version="docx2md " + __version__)
    return p


def setup_logging(verbose=False, log_file=None):
    """初始化日志：控制台走 stderr，可选另写 UTF-8 文件。"""
    level = logging.DEBUG if verbose else logging.INFO
    lg = logging.getLogger("docx2md")
    lg.setLevel(level)
    lg.handlers.clear()
    fmt = logging.Formatter("%(levelname)s %(message)s")
    sh = logging.StreamHandler(sys.stderr)
    sh.setFormatter(fmt)
    lg.addHandler(sh)
    if log_file:
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        lg.addHandler(fh)
    return lg


def main(argv=None):
    args = build_parser().parse_args(argv)
    setup_logging(args.verbose, args.log_file)
    # Windows 控制台中文编码兜底
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8")
            except (ValueError, OSError):
                pass

    src = Path(args.src).expanduser().resolve()
    if not src.is_dir():
        sys.exit(f"输入路径不是文件夹：{src}")
    dst_raw = args.dst_opt or args.dst
    dst = Path(dst_raw).expanduser().resolve() if dst_raw else src / "dify_md"
    flat = args.flat or args.legacy_flat.lower() == "flat"

    result = convert(src, dst, flat=flat, log=lambda m: logger.info(m))

    if not result["ok"]:
        sys.exit(result["message"])


