#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""可编程转换引擎。

把命令行 main() 里的核心编排逻辑抽成函数，供图形界面与命令行共用，
保证两种入口行为完全一致。函数体逐字沿用原逻辑，仅把 sys.exit/print
改为「返回结果 + 进度回调」，不改动任何解析算法。
"""

import csv
import logging
import re
from pathlib import Path

from config import set_flat_mode
from core.text import _merge_wrapped_paras, _split_glued_paras, read_paragraphs, split_head_toc_aligned
from core.split import _strip_count_suffix, _title_piece_count, split_articles
from core.render import build_yaml
from core.convert import convert_one, demote_headings, strip_md_head
from core.recognize import detect_meta, is_heading_like


def convert(src, dst=None, flat=False, log=None, progress=None):
    """批量转换一个文件夹下的所有 docx。

    参数:
        src: 输入文件夹（str 或 Path）
        dst: 输出文件夹（str 或 Path）；None 时默认 <src>/dify_md
        flat: 是否启用 Dify 扁平模式
        log: 可选回调 log(message: str)，用于界面显示日志
        progress: 可选回调 progress(done: int, total: int, filename: str)

    返回:
        dict: {
            "ok": bool,               # 是否整体成功
            "message": str,           # 结果说明
            "src": str, "dst": str,
            "total": int, "succeeded": int, "failed": int,
            "rows": list,             # 清单行
            "report": str,            # 处理清单.csv 路径
            "error": str,             # 致命错误（若有）
        }
    """
    def _emit(msg):
        if log:
            log(str(msg))

    result = {
        "ok": False, "message": "", "src": "", "dst": "",
        "total": 0, "succeeded": 0, "failed": 0,
        "rows": [], "report": "", "error": "",
    }

    src = Path(src).expanduser().resolve()
    result["src"] = str(src)
    if not src.is_dir():
        result["error"] = f"输入路径不是文件夹：{src}"
        result["message"] = result["error"]
        return result

    dst_raw = dst
    dst = Path(dst_raw).expanduser().resolve() if dst_raw else src / "dify_md"
    result["dst"] = str(dst)

    # 扁平模式：每次转换先复位，避免同一进程内多次调用时状态残留
    set_flat_mode(bool(flat))
    try:
        dst.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        result["error"] = f"无法创建输出文件夹：{exc}"
        result["message"] = result["error"]
        return result

    files = sorted([p for p in src.rglob("*.docx") if not p.name.startswith("~$")])
    if not files:
        result["error"] = f"在 {src} 下没有找到 .docx 文件"
        result["message"] = result["error"]
        return result

    rows = []
    total = len(files)
    result["total"] = total
    _emit(f"发现 {total} 个 docx，输出目录：{dst}\n")

    for idx, p in enumerate(files, 1):
        if progress:
            try:
                progress(idx - 1, total, p.name)
            except Exception:
                pass
        try:
            texts, blanks = read_paragraphs(p)
            texts, blanks = split_head_toc_aligned(texts, blanks)
            texts, blanks = _merge_wrapped_paras(list(texts), list(blanks))
            texts, blanks = _split_glued_paras(list(texts), list(blanks))

            doc_title = p.stem
            expected = _title_piece_count(doc_title)
            base = re.sub(r"^\d+[_\-、.．]\s*", "", doc_title)
            base = re.sub(r"\s*[（(【\[][^）)】\]]*[）)】\]]\s*$", "", base).strip()
            if not expected:
                for t0 in texts[:3]:
                    t0s = t0.strip()
                    if 0 < len(t0s) <= 40 and is_heading_like(t0s) and _title_piece_count(t0s):
                        expected = _title_piece_count(t0s)
                        base = _strip_count_suffix(t0s)
                        break

            preface, parts, piece_nums = split_articles(
                texts, blanks, expected=expected, base=base)
            total_parts = len(parts)
            out_path = dst / (p.stem + ".md")

            if total_parts == 1:
                _t, seg, bl = parts[0]
                md, kind, n_art, n_ch, warns, cover = convert_one(
                    seg, bl, p.stem, p.name, title_override=doc_title)
                warns_all = [(kind, warns)]
                covers = [cover]
                kinds = [kind]
                n_arts, n_chs = [n_art], [n_ch]
            else:
                body = []
                kinds, n_arts, n_chs, warns_all, covers = [], [], [], [], []
                titles_unique = len({pt for pt, _s, _b in parts}) > 1
                used = {n for n in piece_nums if n}
                seq_pool = [k for k in range(1, total_parts + len(used) + 1) if k not in used]
                seq_i = 0
                labels = []
                for pi, (ptitle, _s, _b) in enumerate(parts):
                    num = piece_nums[pi] if pi < len(piece_nums) else None
                    if not num:
                        num = seq_pool[seq_i]
                        seq_i += 1
                    labels.append(f"第{num}篇"
                                  + (f"：{ptitle}" if titles_unique and ptitle else ""))
                for pi, (ptitle, seg, bl) in enumerate(parts, 1):
                    md_i, kind, n_art, n_ch, warns, cover = convert_one(
                        seg, bl, p.stem, p.name,
                        part_no=pi, part_total=total_parts, title_override=doc_title)
                    kinds.append(kind)
                    n_arts.append(n_art)
                    n_chs.append(n_ch)
                    warns_all.append((kind, warns))
                    covers.append(cover)
                    body.append(f"## {labels[pi - 1]}")
                    body.append("")
                    body.append(demote_headings(strip_md_head(md_i)))
                    body.append("")
                meta = detect_meta([t for _t, seg, _b in parts for t in seg])
                meta["title"] = doc_title
                merged_yaml = build_yaml(
                    Path(p.name), meta,
                    [("structure_type", "多篇合集"), ("part_count", total_parts)])
                merged = list(merged_yaml)
                merged.append(f"# {doc_title}")
                merged.append("")
                merged.append(f"- **文件构成**：本文件为合集，共 {total_parts} 篇——"
                              + "、".join(labels))
                merged.append("")
                merged.append("> 说明：本文件由原始文档（%s）自动结构化重排；"
                              "原文以「篇N」分隔的多篇文章合并收录于此，"
                              "各篇内部标题已整体降一级。" % p.name)
                merged.append("")
                pre_texts = preface[0] if preface else []
                if pre_texts:
                    merged.append("**文头内容**：")
                    merged.append("")
                    for pf in pre_texts:
                        merged.append(f"> {pf}")
                        merged.append("")
                merged.extend(body)
                md = "\n".join(merged).rstrip() + "\n"
                n_art, n_ch = sum(n_arts), sum(n_chs)

            out_path.write_text(md, encoding="utf-8")
            cover = min(covers) if covers else 1.0
            if any(any(w.startswith("缺失") for w in ws) for _k, ws in warns_all):
                status = "条文缺失"
            elif cover < 0.97:
                status = "内容不全"
            else:
                status = "正常"
            note_parts = []
            for ki, (kind, ws) in enumerate(warns_all, 1):
                ws2 = [w for w in ws if not w.startswith(("缺失", "内容覆盖"))]
                if len(warns_all) > 1 and ws2:
                    note_parts.append(f"篇{ki}[{kind}]：" + "；".join(ws2))
                elif ws2:
                    note_parts.append("；".join(ws2))
            note = "；".join(note_parts)
            rows.append([p.name,
                         kinds[0] if total_parts == 1 else f"合集（{total_parts}篇）",
                         status, n_ch, n_art, f"{cover:.0%}", note or "-", str(out_path)])
            if total_parts == 1:
                unit = "组问答" if kinds[0] == "问答体例" else ("章" if kinds[0] == "章条体例" else "节")
                _emit(f"  [{status}] {p.name} → {kinds[0]}，{n_arts[0]}{unit}（覆盖 {cover:.0%}）")
            else:
                detail = "；".join(f"篇{i+1}「{(parts[i][0] or '未命名')[:14]}」{kinds[i]}"
                                   for i in range(total_parts))
                _emit(f"  [{status}] {p.name} → 合集 {total_parts} 篇（覆盖 {cover:.0%}）：{detail}")
        except Exception as e:  # noqa: BLE001
            _emit(f"  [失败] {p.name}：{e}")
            rows.append([p.name, "-", "失败", 0, 0, "-", f"{type(e).__name__}: {e}", "-"])

    if progress:
        try:
            progress(total, total, "")
        except Exception:
            pass

    report = dst / "处理清单.csv"
    try:
        with open(report, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["源文件", "体例", "状态", "章数", "节数/问答数", "内容覆盖率", "备注", "输出文件"])
            w.writerows(rows)
    except OSError as exc:
        result["error"] = f"写入处理清单失败（请检查输出目录权限或文件是否被占用）：{exc}"
        result["message"] = result["error"]
        return result

    succeeded = len([r for r in rows if r[2] == "正常"])
    failed = len([r for r in rows if r[2] == "失败"])
    result["ok"] = True
    result["succeeded"] = succeeded
    result["failed"] = failed
    result["rows"] = rows
    result["report"] = str(report)
    result["message"] = f"完成：{succeeded}/{len(rows)} 正常"
    _emit(f"\n{result['message']}")
    _emit(f"清单：{report}")
    return result
