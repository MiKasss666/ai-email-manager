"""gui/user_info_dialog.py — User info 编辑表单（对应草图图 2）。

改完点 Save 直接写回 config.json。

实现上走「读出现有配置 → 覆盖表单字段 → from_dict 归一 → save」这条路，
好处有两个：
  · AI 后端等其它字段原样保留，不会因为改个学校名就把模型配置弄丢；
  · 归一逻辑与程序运行时完全同一套，不会出现「界面填 0、程序里变成 1」两套规则。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Optional

from core import config

from . import theme, widgets

# 表单字段顺序与草图一致
_IDENTITY_FIELDS = (
    ("email", "email"),
    ("name", "name"),
    ("student_id", "student id"),
    ("school", "school"),
    ("faculty", "faculty"),
    ("courses", "courses"),
)


class UserInfoDialog(tk.Toplevel):
    """User info 表单窗口。"""

    def __init__(self, master: tk.Misc, cfg: config.Config,
                 on_saved: Optional[Callable[[config.Config], None]] = None):
        super().__init__(master)
        self._cfg = cfg
        self._on_saved = on_saved

        self.title("User info")
        self.configure(bg=theme.CARD)
        self.resizable(False, False)
        self.transient(master)
        theme.apply_icon(self)

        self._vars = {
            "email": tk.StringVar(),
            "name": tk.StringVar(),
            "student_id": tk.StringVar(),
            "school": tk.StringVar(),
            "faculty": tk.StringVar(),
            "courses": tk.StringVar(),
            "language": tk.StringVar(),
            "summary_recipient": tk.StringVar(),
            "summary_days": tk.StringVar(),
        }

        self._build()
        self._refresh_from_cfg()

        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        widgets.center_on_parent(self, master)
        widgets.make_modal(self)

    # ── 界面 ──────────────────────────────────────────────────
    def _build(self) -> None:
        # 注意：控件构造时的 padx/pady 只能是单个数值，元组写法只在 pack()/grid() 里合法
        body = tk.Frame(self, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)

        row = 0
        widgets.section(body, row, "Identity"); row += 1
        for key, label in _IDENTITY_FIELDS:
            widgets.form_row(body, row, label,
                             ttk.Entry(body, textvariable=self._vars[key], width=44))
            row += 1
        widgets.hint(body, row, "courses is comma separated (e.g. COMP1021, MATH1013). "
                                "Filling in student id noticeably improves detection of "
                                "personal mail addressed to you.")
        row += 1

        # 草图里的 "text" 段 = 输出语言 + 总结相关文案设置（language 就放在这一组里，不重复出现）
        widgets.section(body, row, "text (output & wording)"); row += 1
        widgets.form_row(body, row, "language",
                         ttk.Combobox(body, textvariable=self._vars["language"],
                                      values=("zh", "en"), state="readonly", width=8))
        row += 1
        widgets.form_row(body, row, "summary recipient",
                         ttk.Entry(body, textvariable=self._vars["summary_recipient"], width=44))
        row += 1
        widgets.form_row(body, row, "summary days",
                         ttk.Spinbox(body, from_=1, to=365, textvariable=self._vars["summary_days"],
                                     width=8))
        row += 1
        widgets.hint(body, row, "summary recipient left empty means the report goes to your own "
                                "email. summary days is how far back the report looks. "
                                "language controls the wording of the report email, not this window.",
                     wraplength=460)

        footer = tk.Frame(self, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        footer.pack(fill="x")
        self._status = tk.Label(footer, text="", bg=theme.CARD, fg=theme.MUTED,
                                font=theme.FONT, anchor="w")
        self._status.pack(side="left")
        ttk.Button(footer, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(footer, text="Save", command=self._save).pack(side="right", padx=(0, 8))

    # ── 数据 ──────────────────────────────────────────────────
    def _refresh_from_cfg(self) -> None:
        cfg = self._cfg
        self._vars["email"].set(cfg.email)
        self._vars["name"].set(cfg.name)
        self._vars["student_id"].set(cfg.student_id)
        self._vars["school"].set(cfg.school)
        self._vars["faculty"].set(cfg.faculty)
        self._vars["courses"].set(", ".join(cfg.courses))
        self._vars["language"].set(cfg.language)
        self._vars["summary_recipient"].set(cfg.summary_recipient)
        self._vars["summary_days"].set(str(cfg.summary_days))

    def _save(self) -> None:
        raw = config.to_dict(self._cfg)                       # 保住 ai 段等其它配置
        raw.update({key: var.get() for key, var in self._vars.items()})
        new_cfg = config.from_dict(raw)                       # 复用程序同一套归一逻辑

        try:
            config.save(new_cfg)
        except Exception as ex:
            messagebox.showerror("Could not save", f"{type(ex).__name__}: {ex}", parent=self)
            return

        self._cfg = new_cfg
        self._refresh_from_cfg()                              # 回填归一后的真实值

        if self._on_saved is not None:
            self._on_saved(new_cfg)

        problems = new_cfg.identity_problems()
        if problems:
            self._status.configure(text="Saved, but problems remain", fg=theme.DANGER)
            messagebox.showwarning("Saved, but these need attention",
                                   "\n".join(f"- {p}" for p in problems), parent=self)
        else:
            self._status.configure(text="Saved to config.json", fg=theme.ACCENT)
