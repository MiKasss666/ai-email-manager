"""gui/widgets.py — 两个设置对话框共用的表单小部件。"""
from __future__ import annotations

import tkinter as tk

from . import theme


def center_on_parent(window: tk.Misc, parent: tk.Misc) -> None:
    """把窗口摆到父窗口中间；父窗口还没显示时摆到屏幕中间。"""
    window.update_idletasks()
    width = window.winfo_width() or window.winfo_reqwidth()
    height = window.winfo_height() or window.winfo_reqheight()

    try:
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        pw, ph = parent.winfo_width(), parent.winfo_height()
    except Exception:
        px = py = pw = ph = 0

    if pw > 1 and ph > 1:
        x = px + (pw - width) // 2
        y = py + (ph - height) // 3
    else:
        x = (window.winfo_screenwidth() - width) // 2
        y = (window.winfo_screenheight() - height) // 3
    window.geometry(f"+{max(x, 0)}+{max(y, 0)}")


def make_modal(window: tk.Misc) -> None:
    """把窗口尽量做成模态。

    某些环境会失败——窗口还没映射（如隐藏窗口做自测）、或别处已持有 grab——
    这时忽略即可，不该因为模态失败就打不开设置窗口。
    """
    try:
        window.grab_set()
    except Exception:
        pass
    try:
        window.focus_force()
    except Exception:
        pass


def form_row(parent: tk.Misc, row: int, label: str, widget: tk.Misc, label_width: int = 15) -> tk.Misc:
    """一行表单：左标签右控件；控件所在列会横向拉伸。"""
    tk.Label(parent, text=label, bg=theme.CARD, fg=theme.TEXT, font=theme.FONT,
             anchor="e", width=label_width).grid(row=row, column=0, sticky="e", padx=(0, 10), pady=4)
    widget.grid(row=row, column=1, sticky="ew", pady=4)
    return widget


def section(parent: tk.Misc, row: int, title: str, span: int = 2) -> tk.Frame:
    """分组标题：文字 + 右侧一条细横线。"""
    frame = tk.Frame(parent, bg=theme.CARD)
    frame.grid(row=row, column=0, columnspan=span, sticky="ew", pady=(14, 4))
    tk.Label(frame, text=title, bg=theme.CARD, fg=theme.PRIMARY,
             font=theme.FONT_SECTION).pack(side="left")
    tk.Frame(frame, bg=theme.BORDER, height=1).pack(side="left", fill="x", expand=True, padx=(10, 0))
    return frame


def hint(parent: tk.Misc, row: int, text: str, span: int = 2, wraplength: int = 430) -> tk.Label:
    """灰色小字提示。"""
    label = tk.Label(parent, text=text, bg=theme.CARD, fg=theme.MUTED, font=theme.FONT,
                     justify="left", wraplength=wraplength, anchor="w")
    label.grid(row=row, column=0, columnspan=span, sticky="w", pady=(0, 4))
    return label
