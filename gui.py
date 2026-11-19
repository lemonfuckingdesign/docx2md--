#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx2md 图形界面版 —— 傻瓜式操作，无需命令行、无需安装 Python。

用法（普通用户）：双击 exe 即可。开发者调试：python gui.py
"""

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

# 让打包后的程序能正确 import 同目录模块（PyInstaller 下由 spec 处理，这里兜底）
sys.path.insert(0, str(Path(__file__).resolve().parent))

APP_TITLE = "docx2md 文档转换助手"
APP_VERSION = "1.0.0"
CONTACT_EMAIL = "dhly2004@gmail.com"


def resource_path(rel):
    """兼容源码运行与 PyInstaller 单文件打包的资源路径。"""
    base = getattr(sys, "_MEIPASS", None)  # PyInstaller 解压目录
    if base:
        return Path(base) / rel
    return Path(__file__).resolve().parent / rel


# ---------- 配色（简洁浅色卡片风）----------
BG = "#F4F6FB"          # 浅灰蓝页面背景
CARD = "#FFFFFF"        # 白色卡片
PRIMARY = "#4A6CF7"     # 主蓝
PRIMARY_DARK = "#3556D9"
TEXT = "#1F2937"
SUBTEXT = "#8B93A7"
LINE = "#EDF0F6"
SOFT_BLUE = "#EEF3FF"
WARN = "#B45309"
LOG_BG = "#F7F8FC"      # 浅色日志区


def _round_rect(canvas, x1, y1, x2, y2, r=16, **kw):
    """在 Canvas 上画圆角矩形（用多边形平滑模拟）。"""
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kw)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self.msg_queue = queue.Queue()
        self.running = False
        self.src_var = tk.StringVar()
        self.custom_out_var = tk.BooleanVar(value=False)
        self.out_var = tk.StringVar()
        self.flat_var = tk.BooleanVar(value=False)
        self.last_result = None

        self._build_ui()
        self._center(920, 830)
        self.root.after(100, self._poll_queue)

    # ---------- 界面搭建 ----------
    def _build_ui(self):
        cv = tk.Canvas(self.root, bg=BG, highlightthickness=0, bd=0)
        cv.pack(fill="both", expand=True)
        self.cv = cv

        # 顶部标题
        cv.create_text(40, 40, text="docx2md 转换中心", anchor="w",
                       font=("Microsoft YaHei", 23, "bold"), fill=TEXT)
        cv.create_text(42, 76, text="把 Word 文档批量转换成知识库可用的 Markdown —— 选好文件夹，点一下即可",
                       anchor="w", font=("Microsoft YaHei", 10), fill=SUBTEXT)
        cv.create_text(880, 46, text=f"v{APP_VERSION}", anchor="e",
                       font=("Microsoft YaHei", 9), fill=SUBTEXT)

        # ---- 主转换卡片（左，淡蓝渐变感）----
        self._card(36, 104, 588, 452, fill="#EAF1FF", outline="")
        cv.create_text(64, 132, text="文档转换", anchor="w",
                       font=("Microsoft YaHei", 16, "bold"), fill=TEXT)
        cv.create_text(64, 162, text="选择文件夹，自动扫描里面的 Word 文档",
                       anchor="w", font=("Microsoft YaHei", 10), fill=SUBTEXT)

        # 步骤 1：选择文件夹
        cv.create_text(64, 200, text="01   选择装着 Word 文档的文件夹", anchor="w",
                       font=("Microsoft YaHei", 11, "bold"), fill=PRIMARY)
        self.src_entry = tk.Entry(cv, textvariable=self.src_var, font=("Microsoft YaHei", 10),
                                  relief="flat", bd=0, bg="white", fg=TEXT,
                                  highlightthickness=1, highlightbackground="#D7E2F7",
                                  highlightcolor=PRIMARY)
        cv.create_window(64, 222, window=self.src_entry, anchor="nw", width=390, height=40)
        self._button(cv, 464, 222, 132, 40, "选择文件夹", self._choose_src,
                     bg=PRIMARY, fg="white", radius=10, font_size=10)

        # 步骤 2：输出位置
        cv.create_text(64, 286, text="02   结果保存到哪里", anchor="w",
                       font=("Microsoft YaHei", 11, "bold"), fill=PRIMARY)
        ttk.Checkbutton(cv, text="自己指定保存位置（不勾选则自动放到文档文件夹里的 dify_md）",
                        variable=self.custom_out_var, command=self._toggle_out).place(
                            x=58, y=306)
        self.out_entry = tk.Entry(cv, textvariable=self.out_var, font=("Microsoft YaHei", 10),
                                  relief="flat", bd=0, bg="#F3F5FA", fg=TEXT,
                                  highlightthickness=1, highlightbackground="#E3E8F2",
                                  highlightcolor=PRIMARY, state="disabled")
        cv.create_window(64, 336, window=self.out_entry, anchor="nw", width=390, height=40)
        self.out_btn = self._button(cv, 464, 336, 132, 40, "选择位置", self._choose_out,
                                    bg="#D7E0F5", fg="#9AA6BD", radius=10, font_size=10,
                                    enabled=False)

        # 选项
        ttk.Checkbutton(cv, text="扁平模式（推荐用于 Dify：把小标题变成正文，避免切出只有标题的碎块）",
                        variable=self.flat_var).place(x=58, y=392)
        cv.create_text(64, 424, text="转换只调整标题结构与格式，绝不改动正文里的任何一个字。",
                       anchor="w", font=("Microsoft YaHei", 9), fill=SUBTEXT)

        # 操作按钮
        self.start_btn = self._button(cv, 64, 452, 168, 50, "开始转换", self._start,
                                      bg=PRIMARY, fg="white", radius=12, font_size=14, bold=True)
        self.open_btn = self._button(cv, 246, 452, 168, 50, "打开结果文件夹", self._open_result,
                                     bg="#D7E0F5", fg="#9AA6BD", radius=12, font_size=10,
                                     enabled=False)

        # ---- 右侧打赏卡片（白底）----
        self._card(640, 104, 244, 452, fill=CARD, outline=LINE)
        cv.create_text(662, 134, text="支持作者", anchor="w",
                       font=("Microsoft YaHei", 15, "bold"), fill=TEXT)
        cv.create_text(662, 162, text="觉得好用？扫码打赏一杯咖啡", anchor="w",
                       font=("Microsoft YaHei", 9), fill=SUBTEXT)

        # 完整二维码（不裁切），居中显示
        self._qr_card_img = None
        try:
            self._qr_card_img = tk.PhotoImage(file=str(resource_path("assets/donate_qr_card.png")))
            qr_label = tk.Label(cv, image=self._qr_card_img, bg=CARD, cursor="hand2",
                                highlightthickness=1, highlightbackground=LINE)
            cv.create_window(762, 182, window=qr_label, anchor="n")
            qr_label.bind("<Button-1>", self._show_qr)
        except Exception:
            cv.create_text(762, 260, text="[二维码图片缺失]", anchor="center",
                           font=("Microsoft YaHei", 10), fill=SUBTEXT)
        cv.create_text(762, 434, text="微信扫一扫 · 点击二维码可放大", anchor="center",
                       font=("Microsoft YaHei", 9), fill=SUBTEXT)

        # ---- 底部：进度 + 日志卡片 ----
        self._card(36, 572, 848, 180, fill=CARD, outline=LINE)
        cv.create_text(64, 600, text="转换进度", anchor="w",
                       font=("Microsoft YaHei", 13, "bold"), fill=TEXT)
        self.progress = ttk.Progressbar(cv, mode="determinate", maximum=100)
        cv.create_window(64, 624, window=self.progress, anchor="nw", width=792, height=16)

        log_frame = tk.Frame(cv, bg=LOG_BG, highlightthickness=1, highlightbackground=LINE)
        cv.create_window(64, 652, window=log_frame, anchor="nw", width=792, height=82)
        self.log_text = tk.Text(log_frame, bg=LOG_BG, fg=TEXT,
                                font=("Consolas", 10), relief="flat", wrap="word",
                                padx=10, pady=6, height=4, bd=0)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.tag_config("ok", foreground="#1A7F45")
        self.log_text.tag_config("err", foreground="#D64545")
        self.log_text.tag_config("info", foreground="#4A6CF7")
        self._log("欢迎使用！选好文件夹后点「开始转换」即可。", "info")

        # 底部报错邮箱（可点击复制）—— 放在进度卡片下方的留白处
        cv.create_text(40, 774, text="遇到问题 / 有建议？邮箱：", anchor="w",
                       font=("Microsoft YaHei", 9), fill=SUBTEXT)
        mail = cv.create_text(172, 774, text=CONTACT_EMAIL, anchor="w",
                              font=("Microsoft YaHei", 9, "underline"), fill=PRIMARY)
        cv.tag_bind(mail, "<Button-1>", self._copy_email)
        cv.tag_bind(mail, "<Enter>", lambda e: cv.configure(cursor="hand2"))
        cv.tag_bind(mail, "<Leave>", lambda e: cv.configure(cursor=""))

    def _card(self, x, y, w, h, fill=CARD, outline=LINE):
        return _round_rect(self.cv, x, y, x + w, y + h, r=18, fill=fill,
                           outline=outline if outline else "", width=1)

    def _button(self, parent, x, y, w, h, text, command, bg, fg="white",
                radius=10, font_size=10, bold=False, enabled=True):
        """用 Canvas 画圆角按钮，返回一个带 .configure(state=...) 的代理对象。"""
        frame = tk.Frame(parent, bg=parent.cget("bg") if isinstance(parent, tk.Canvas) else BG,
                         highlightthickness=0, bd=0)
        cvs = tk.Canvas(frame, width=w, height=h, bg=frame.cget("bg"),
                        highlightthickness=0, bd=0)
        cvs.pack()
        btn = _RoundButton(cvs, w, h, text, command, bg, fg, radius, font_size, bold, enabled)
        parent.create_window(x, y, window=frame, anchor="nw")
        return btn

    def _copy_email(self, event=None):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(CONTACT_EMAIL)
            self._log(f"[提示] 已复制邮箱：{CONTACT_EMAIL}", "ok")
        except Exception:
            pass

    def _show_qr(self, event=None):
        win = tk.Toplevel(self.root)
        win.title("打赏支持")
        win.resizable(False, False)
        win.configure(bg="white")
        try:
            img = tk.PhotoImage(file=str(resource_path("assets/donate_qr.png")))
            lbl = tk.Label(win, image=img, bg="white")
            lbl.image = img  # 保留引用
            lbl.pack(padx=16, pady=(16, 6))
        except Exception:
            tk.Label(win, text="[二维码图片缺失]", bg="white").pack(padx=20, pady=20)
        tk.Label(win, text="感谢你的支持！微信扫一扫即可打赏", bg="white", fg=TEXT,
                 font=("Microsoft YaHei", 10, "bold")).pack(pady=(0, 14))
        win.transient(self.root)
        win.grab_set()

    def _toggle_out(self):
        if self.custom_out_var.get():
            self.out_entry.configure(state="normal", bg="white",
                                     highlightbackground="#D7E2F7")
            self.out_btn.configure(state="normal")
            self.out_btn.set_colors("#4A6CF7", "white")
        else:
            self.out_entry.configure(state="disabled", bg="#F3F5FA",
                                     highlightbackground="#E3E8F2")
            self.out_btn.configure(state="disabled")
            self.out_btn.set_colors("#D7E0F5", "#9AA6BD")

    # ---------- 交互 ----------
    def _choose_src(self):
        d = filedialog.askdirectory(title="选择装着 Word 文档的文件夹")
        if d:
            self.src_var.set(d)

    def _choose_out(self):
        d = filedialog.askdirectory(title="选择结果保存位置")
        if d:
            self.out_var.set(d)

    def _center(self, w, h):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _log(self, msg, tag=None):
        self.log_text.insert("end", str(msg) + "\n", tag)
        self.log_text.see("end")

    def _open_result(self):
        if not self.last_result:
            return
        path = Path(self.last_result["dst"])
        if path.exists():
            try:
                os.startfile(str(path))  # Windows
            except AttributeError:
                subprocess.run(["open", str(path)])  # macOS
            except Exception:
                subprocess.run(["xdg-open", str(path)])  # Linux

    # ---------- 转换线程 ----------
    def _start(self):
        if self.running:
            return
        src = self.src_var.get().strip()
        if not src:
            messagebox.showwarning("还没选择文件夹", "请先选择装着 Word 文档的文件夹。")
            return
        if not Path(src).is_dir():
            messagebox.showerror("文件夹不存在", f"找不到这个文件夹：\n{src}\n\n有问题欢迎反馈：{CONTACT_EMAIL}")
            return

        dst = self.out_var.get().strip() if self.custom_out_var.get() else None
        flat = self.flat_var.get()

        self.running = True
        self.last_result = None
        self.start_btn.configure(state="disabled")
        self.start_btn.set_text("转换中…")
        self.open_btn.configure(state="disabled")
        self.progress["value"] = 0
        self.log_text.delete("1.0", "end")
        self._log("开始转换，请稍候……", "info")

        t = threading.Thread(target=self._worker, args=(src, dst, flat), daemon=True)
        t.start()

    def _worker(self, src, dst, flat):
        try:
            from engine import convert

            def emit(msg):
                self.msg_queue.put(("log", str(msg)))

            def prog(done, total, name):
                self.msg_queue.put(("progress", (done, total, name)))

            result = convert(src, dst, flat=flat, log=emit, progress=prog)
            self.msg_queue.put(("done", result))
        except Exception as e:  # noqa: BLE001
            self.msg_queue.put(("fatal", f"{type(e).__name__}: {e}"))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    self._log(payload)
                elif kind == "progress":
                    done, total, name = payload
                    if total:
                        self.progress["value"] = done / total * 100
                    if name:
                        self._log(f"[{done}/{total}] {name}", "info")
                elif kind == "done":
                    self._on_done(payload)
                elif kind == "fatal":
                    self._on_fatal(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    def _on_done(self, result):
        self.running = False
        self.last_result = result
        self.start_btn.configure(state="normal")
        self.start_btn.set_text("开始转换")
        self.progress["value"] = 100 if result["ok"] else self.progress["value"]

        if not result["ok"]:
            self._log("[错误] " + result["message"], "err")
            messagebox.showerror("转换失败", result["message"] + f"\n\n如反复失败，请截图发到：{CONTACT_EMAIL}")
            return

        self._log(f"\n[完成] {result['message']}", "ok")
        self._log(f"[完成] 结果清单：{result['report']}", "ok")
        self.open_btn.configure(state="normal")
        self.open_btn.set_colors("#4A6CF7", "white")
        warn = ""
        if result["failed"]:
            warn = f"\n其中 {result['failed']} 个文件转换失败（多为文件损坏或加密），详见清单。"
        messagebox.showinfo(
            "转换完成",
            f"处理了 {result['total']} 个文档：\n\n"
            f"• 成功：{result['succeeded']} 个\n"
            f"• 失败：{result['failed']} 个{warn}\n\n"
            f"结果已保存到：\n{result['dst']}\n\n"
            f"可以点「打开结果文件夹」查看。"
        )

    def _on_fatal(self, msg):
        self.running = False
        self.start_btn.configure(state="normal")
        self.start_btn.set_text("开始转换")
        self._log("[错误] 发生错误：" + msg, "err")
        messagebox.showerror("发生错误", f"转换过程出错：\n{msg}\n\n请截图发到：{CONTACT_EMAIL}，我们会尽快处理。")


class _RoundButton:
    """圆角按钮代理：在 Canvas 上绘制，提供 configure/state/text 接口。"""

    def __init__(self, canvas, w, h, text, command, bg, fg, radius, font_size, bold, enabled):
        self.cvs = canvas
        self.w, self.h = w, h
        self.text = text
        self.command = command
        self.bg, self.fg = bg, fg
        self.radius = radius
        self.font = ("Microsoft YaHei", font_size, "bold" if bold else "normal")
        self.state = "normal" if enabled else "disabled"
        self.item = None
        self.txt = None
        self._draw()

    def _draw(self):
        if self.item is not None:
            self.cvs.delete(self.item)
            self.cvs.delete(self.txt)
        disp_bg = self.bg if self.state == "normal" else "#E3E8F2"
        disp_fg = self.fg if self.state == "normal" else "#AAB3C4"
        self.item = _round_rect(self.cvs, 1, 1, self.w - 1, self.h - 1,
                                self.radius, fill=disp_bg, outline="")
        self.txt = self.cvs.create_text(self.w / 2, self.h / 2, text=self.text,
                                        font=self.font, fill=disp_fg)
        if self.state == "normal":
            self.cvs.bind("<Button-1>", lambda e: self.command())
            self.cvs.configure(cursor="hand2")
        else:
            self.cvs.unbind("<Button-1>")
            self.cvs.configure(cursor="")

    def configure(self, **kw):
        if "state" in kw:
            self.state = kw["state"]
        self._draw()

    def set_text(self, text):
        self.text = text
        self._draw()

    def set_colors(self, bg, fg):
        self.bg, self.fg = bg, fg
        self._draw()


def _selftest(argv):
    """隐藏自测入口：exe --selftest <输入> <输出> <结果文件>。"""
    import json
    src, dst, out_file = argv[2], argv[3], argv[4]
    from engine import convert
    result = convert(src, dst, flat=False)
    Path(out_file).write_text(
        json.dumps({"ok": result["ok"], "total": result["total"],
                    "succeeded": result["succeeded"], "message": result["message"]},
                   ensure_ascii=False),
        encoding="utf-8",
    )


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        _selftest(sys.argv)
        return
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
