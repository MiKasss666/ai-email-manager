"""gui/theme.py — 视觉常量与 ttk 样式。

界面里所有颜色、字号都从这里取，避免每个文件各写一套导致风格漂移。
"""
from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import ttk
from typing import Optional

from core import config

# ── 字体（优先中文显示效果好的雅黑）────────────────────────────────
FAMILY = "Microsoft YaHei UI"
FONT = (FAMILY, 10)
FONT_BOLD = (FAMILY, 10, "bold")
FONT_TITLE = (FAMILY, 15, "bold")
FONT_SECTION = (FAMILY, 10, "bold")
FONT_MONO = ("Consolas", 9)
FONT_BUTTON = (FAMILY, 11, "bold")

# ── 颜色 ──────────────────────────────────────────────────────────
BG = "#f4f5f7"          # 窗口底色
CARD = "#ffffff"        # 卡片底色
BORDER = "#d8dbe0"
TEXT = "#1f2328"
MUTED = "#8b9198"
PRIMARY = "#2f6fd0"     # 主操作（分类）
ACCENT = "#1f8a4c"      # 次操作（总结）
DANGER = "#c0392b"
LOG_BG = "#1e1f22"      # 日志区（深色终端风）
LOG_FG = "#d8dbe0"

PAD = 12


def apply(root: tk.Misc) -> None:
    """应用 ttk 样式。"""
    style = ttk.Style(root)
    for candidate in ("vista", "winnative", "clam"):
        try:
            if candidate in style.theme_names():
                style.theme_use(candidate)
                break
        except Exception:
            continue

    style.configure("TLabel", font=FONT)
    style.configure("TSection.TLabel", font=FONT_SECTION)
    style.configure("TMuted.TLabel", font=FONT, foreground=MUTED)
    style.configure("TButton", font=FONT, padding=(10, 5))
    style.configure("TCheckbutton", font=FONT)
    style.configure("TRadiobutton", font=FONT)
    style.configure("TEntry", padding=3)
    style.configure("TCombobox", padding=2)


def card(master: tk.Misc, **kwargs) -> tk.Frame:
    """一个带边框的白色卡片容器。"""
    options = dict(bg=CARD, highlightbackground=BORDER, highlightthickness=1, bd=0)
    options.update(kwargs)
    return tk.Frame(master, **options)


def action_button(master: tk.Misc, text: str, command, color: str = PRIMARY) -> tk.Button:
    """主操作按钮（用 tk.Button 而不是 ttk.Button，因为 ttk 无法跨主题改背景色）。"""
    return tk.Button(master, text=text, command=command, font=FONT_BUTTON,
                     bg=color, fg="#ffffff", activebackground=color, activeforeground="#ffffff",
                     relief="flat", bd=0, padx=18, pady=9, cursor="hand2")


def plain_button(master: tk.Misc, text: str, command) -> tk.Button:
    """次要按钮（白底描边）。"""
    return tk.Button(master, text=text, command=command, font=FONT,
                     bg=CARD, fg=TEXT, activebackground=BG, activeforeground=TEXT,
                     relief="solid", bd=1, padx=12, pady=5, cursor="hand2",
                     highlightbackground=BORDER)


# ── 窗口图标 ──────────────────────────────────────────────────────
ICON_NAME = "EmailManager.ico"


def icon_path() -> Optional[str]:
    """独立图标文件的路径；找不到返回 None。

    打包成单文件 exe 时本来就没有独立的 .ico（图标被嵌进 exe 资源里），
    这时返回 None 是正常的，交给 icon_source() 去 exe 自身取。
    """
    candidates = []
    bundle = getattr(sys, "_MEIPASS", None)
    if bundle:
        candidates.append(os.path.join(bundle, ICON_NAME))
    candidates.append(os.path.join(config.app_dir(), ICON_NAME))
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def icon_source() -> Optional[str]:
    """给 iconbitmap 用的图标来源。

    优先用旁边的 .ico 文件（源码运行 / onedir 打包都能找到）；
    单文件 exe 里没有独立文件，就直接从 **exe 自身的内嵌图标资源**取 ——
    Tk 在 Windows 上支持从 .exe/.dll 抽图标（已实测可行）。
    这样单文件 exe 真正自包含：不需要在它旁边放 .ico。
    """
    path = icon_path()
    if path:
        return path
    if getattr(sys, "frozen", False):
        exe = os.path.abspath(sys.executable)
        if os.path.exists(exe):
            return exe
    return None


def apply_icon(window: tk.Misc) -> None:
    """给这个窗口设置图标，并把它设为后续新建 Toplevel 的默认图标。

    取不到图标时静默跳过 —— 图标缺失不该让程序打不开。
    （任务栏图标还需要 app.py 里的 AppUserModelID 配合，见那里的说明。）
    """
    source = icon_source()
    if not source:
        return
    try:
        window.iconbitmap(source)              # 本窗口自己的图标
    except Exception:
        pass
    try:
        window.iconbitmap(default=source)      # 之后新建的 Toplevel 也用它
    except Exception:
        pass
