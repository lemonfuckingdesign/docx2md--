# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置：把 GUI 版打成单个 exe。"""
from PyInstaller.utils.hooks import collect_submodules

# 收集我们自己的包，确保 core 子模块全部打进去
hidden = collect_submodules("core") + ["engine", "cli", "config", "gui", "docx2md"]

a = Analysis(
    ["gui.py"],
    pathex=["."],
    binaries=[],
    datas=[
        ("assets/donate_qr.png", "assets"),
        ("assets/donate_qr_small.png", "assets"),
        ("assets/donate_qr_card.png", "assets"),
    ],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["matplotlib", "numpy", "PIL", "tkinter.test", "unittest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="docx2md-文档转换助手",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 不显示黑色命令行窗口（GUI 程序）
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
