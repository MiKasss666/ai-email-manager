"""gui/ai_settings_dialog.py — AI 后端设置面板（标题：AI setting）。

左边列出 config.json 里配置的所有后端，右边编辑选中那个的细节。
「Test connection」会真的发一个最小请求（消耗极少量额度），用来确认
base_url / model / API key 填对了 —— 这是给「用户自填模型和 API」准备的最重要的一块。

编辑的是配置的深拷贝，只有点 Save 才写回磁盘；因此 Cancel 一定是干净的。
"""
from __future__ import annotations

import copy
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from typing import Callable, Dict, Optional

from core import ai, config

from . import theme, widgets
from .task import TaskRunner

# 这些字段只有非命令行后端才有意义
_API_ONLY_FIELDS = ("base_url", "model", "api_key", "api_key_env",
                    "api_key_required", "temperature", "max_tokens", "json_mode")


class AiSettingsDialog(tk.Toplevel):
    """AI 后端设置窗口。"""

    def __init__(self, master: tk.Misc, cfg: config.Config,
                 on_saved: Optional[Callable[[config.Config], None]] = None):
        super().__init__(master)
        self._cfg = cfg
        self._on_saved = on_saved

        self._profiles: Dict[str, config.AiProfile] = copy.deepcopy(cfg.ai_profiles)
        self._active = cfg.ai_active
        self._names = list(self._profiles)
        self._current: Optional[str] = self._names[0] if self._names else None
        self._loading = False
        self._runner = TaskRunner(self)

        self.title("AI setting")
        self.configure(bg=theme.CARD)
        self.resizable(False, False)
        self.transient(master)
        theme.apply_icon(self)

        self._vars = {key: tk.StringVar() for key in
                      ("name", "type", "base_url", "model", "api_key", "api_key_env",
                       "temperature", "max_tokens", "timeout", "cli_command")}
        self._var_key_required = tk.BooleanVar(value=True)
        self._var_json_mode = tk.BooleanVar(value=True)
        self._fields: Dict[str, tk.Misc] = {}

        self._build()
        self._refresh_list()
        if self._current:
            self._load_profile(self._current)

        self.bind("<Escape>", lambda _event: self._close())
        self.protocol("WM_DELETE_WINDOW", self._close)
        widgets.center_on_parent(self, master)
        widgets.make_modal(self)

    # ── 界面 ──────────────────────────────────────────────────
    def _build(self) -> None:
        outer = tk.Frame(self, bg=theme.CARD)
        outer.pack(fill="both", expand=True)

        self._build_list(outer)

        tk.Frame(outer, bg=theme.BORDER, width=1).pack(side="left", fill="y", pady=theme.PAD)

        form = tk.Frame(outer, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        form.pack(side="left", fill="both", expand=True)
        form.columnconfigure(1, weight=1)

        row = 0
        widgets.section(form, row, "Backend"); row += 1
        widgets.form_row(form, row, "name",
                         ttk.Entry(form, textvariable=self._vars["name"], width=36, state="readonly"))
        row += 1
        type_box = ttk.Combobox(form, textvariable=self._vars["type"],
                                values=config.PROVIDER_TYPES, state="readonly", width=14)
        type_box.bind("<<ComboboxSelected>>", lambda _e: self._apply_type_state())
        widgets.form_row(form, row, "type", type_box); row += 1

        widgets.section(form, row, "Endpoint"); row += 1
        self._fields["base_url"] = widgets.form_row(
            form, row, "base_url", ttk.Entry(form, textvariable=self._vars["base_url"], width=36)); row += 1
        self._fields["model"] = widgets.form_row(
            form, row, "model", ttk.Entry(form, textvariable=self._vars["model"], width=36)); row += 1
        self._fields["api_key"] = widgets.form_row(
            form, row, "api key", ttk.Entry(form, textvariable=self._vars["api_key"], width=36)); row += 1
        self._fields["api_key_env"] = widgets.form_row(
            form, row, "api key env var",
            ttk.Entry(form, textvariable=self._vars["api_key_env"], width=36)); row += 1
        self._fields["api_key_required"] = widgets.form_row(
            form, row, "",
            ttk.Checkbutton(form, text="requires an API key (uncheck for local Ollama / LM Studio)",
                            variable=self._var_key_required)); row += 1
        widgets.hint(form, row, "Prefer the env var route: the key never lands in config.json, so the "
                                "file is safe to share. If both are filled, the env var wins.",
                     wraplength=460)
        row += 1

        widgets.section(form, row, "Generation"); row += 1
        self._fields["temperature"] = widgets.form_row(
            form, row, "temperature", ttk.Entry(form, textvariable=self._vars["temperature"], width=8)); row += 1
        self._fields["max_tokens"] = widgets.form_row(
            form, row, "max tokens",
            ttk.Spinbox(form, from_=0, to=200000, textvariable=self._vars["max_tokens"], width=10)); row += 1
        self._fields["timeout"] = widgets.form_row(
            form, row, "timeout (s)",
            ttk.Spinbox(form, from_=1, to=3600, textvariable=self._vars["timeout"], width=10)); row += 1
        self._fields["json_mode"] = widgets.form_row(
            form, row, "",
            ttk.Checkbutton(form, text="request JSON mode (falls back automatically if unsupported)",
                            variable=self._var_json_mode)); row += 1

        widgets.section(form, row, "CLI backend only"); row += 1
        self._fields["cli_command"] = widgets.form_row(
            form, row, "cli_command",
            ttk.Entry(form, textvariable=self._vars["cli_command"], width=36)); row += 1
        widgets.hint(form, row, "Leave empty to auto-detect the Claude Code CLI. Any command that reads "
                                "stdin works, e.g. ollama run qwen2.5:7b.\n"
                                "extra_headers (special headers for Azure / OpenRouter) can only be "
                                "edited in config.json.",
                     wraplength=460)
        row += 1

        actions = tk.Frame(form, bg=theme.CARD)
        actions.grid(row=row, column=0, columnspan=2, sticky="w", pady=(14, 0))
        ttk.Button(actions, text="Set as active", command=self._set_active).pack(side="left")
        ttk.Button(actions, text="Test connection", command=self._test).pack(side="left", padx=(8, 0))

        self._build_footer()

    def _build_list(self, outer: tk.Frame) -> None:
        left = tk.Frame(outer, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        left.pack(side="left", fill="y")

        tk.Label(left, text="Backends", bg=theme.CARD, fg=theme.TEXT,
                 font=theme.FONT_SECTION).pack(anchor="w")
        self._list = tk.Listbox(left, width=22, height=17, font=theme.FONT,
                                exportselection=False, activestyle="none",
                                relief="solid", bd=1, highlightthickness=0)
        self._list.pack(fill="y", expand=True, pady=(4, 6))
        self._list.bind("<<ListboxSelect>>", self._on_select)

        buttons = tk.Frame(left, bg=theme.CARD)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Add", width=7, command=self._new_profile).pack(side="left")
        ttk.Button(buttons, text="Remove", width=8, command=self._delete_profile).pack(side="left", padx=(6, 0))

        self._active_label = tk.Label(left, text="", bg=theme.CARD, fg=theme.PRIMARY,
                                      font=theme.FONT, wraplength=170, justify="left", anchor="w")
        self._active_label.pack(anchor="w", pady=(10, 0))
        tk.Label(left, text="* = the active backend", bg=theme.CARD, fg=theme.MUTED,
                 font=theme.FONT, anchor="w").pack(anchor="w", pady=(4, 0))

    def _build_footer(self) -> None:
        footer = tk.Frame(self, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        footer.pack(fill="x", side="bottom")
        self._status = tk.Label(footer, text="", bg=theme.CARD, fg=theme.MUTED,
                                font=theme.FONT, anchor="w", wraplength=440, justify="left")
        self._status.pack(side="left")
        ttk.Button(footer, text="Cancel", command=self._close).pack(side="right")
        ttk.Button(footer, text="Save", command=self._save).pack(side="right", padx=(0, 8))

    # ── 列表 ──────────────────────────────────────────────────
    def _refresh_list(self) -> None:
        self._names = list(self._profiles)
        self._list.delete(0, "end")
        for name in self._names:
            self._list.insert("end", ("* " if name == self._active else "   ") + name)
        if self._current in self._names:
            index = self._names.index(self._current)
            self._list.selection_clear(0, "end")
            self._list.selection_set(index)
            self._list.see(index)
        self._refresh_active_label()

    def _refresh_active_label(self) -> None:
        self._active_label.configure(text=f"Active: {self._active}")

    def _on_select(self, _event=None) -> None:
        if self._loading:
            return
        selection = self._list.curselection()
        if not selection:
            return
        name = self._names[selection[0]]
        if name == self._current:
            return
        self._collect_current()
        self._current = name
        self._load_profile(name)

    # ── 表单 ↔ 数据 ───────────────────────────────────────────
    def _load_profile(self, name: str) -> None:
        profile = self._profiles[name]
        self._loading = True
        try:
            self._vars["name"].set(name)
            self._vars["type"].set(profile.type)
            self._vars["base_url"].set(profile.base_url)
            self._vars["model"].set(profile.model)
            self._vars["api_key"].set(profile.api_key)
            self._vars["api_key_env"].set(profile.api_key_env)
            self._var_key_required.set(profile.api_key_required)
            self._vars["temperature"].set(str(profile.temperature))
            self._vars["max_tokens"].set(str(profile.max_tokens))
            self._vars["timeout"].set(str(profile.timeout))
            self._var_json_mode.set(profile.json_mode)
            self._vars["cli_command"].set(profile.cli_command)
        finally:
            self._loading = False
        self._apply_type_state()
        self._status.configure(text="")

    def _collect_current(self) -> None:
        """把表单写回内存中的 profile（保存前 / 切换前调用）。"""
        if self._current is None or self._current not in self._profiles:
            return
        previous = self._profiles[self._current]
        raw = {
            "type": self._vars["type"].get(),
            "base_url": self._vars["base_url"].get(),
            "model": self._vars["model"].get(),
            "api_key": self._vars["api_key"].get(),
            "api_key_env": self._vars["api_key_env"].get(),
            "api_key_required": self._var_key_required.get(),
            "temperature": self._vars["temperature"].get(),
            "max_tokens": self._vars["max_tokens"].get(),
            "timeout": self._vars["timeout"].get(),
            "json_mode": self._var_json_mode.get(),
            "extra_headers": dict(previous.extra_headers),   # 界面不编辑，原样保留
            "cli_command": self._vars["cli_command"].get(),
        }
        self._profiles[self._current] = config.ai_profile_from_dict(self._current, raw)

    def _apply_type_state(self) -> None:
        """type=cli 时只留命令和超时，其余字段置灰，避免用户填了却不生效。"""
        is_cli = self._vars["type"].get() == "cli"
        for key in _API_ONLY_FIELDS:
            widget = self._fields.get(key)
            if widget is not None:
                self._set_state(widget, "disabled" if is_cli else "normal")
        cli = self._fields.get("cli_command")
        if cli is not None:
            self._set_state(cli, "normal" if is_cli else "disabled")

    @staticmethod
    def _set_state(widget: tk.Misc, state: str) -> None:
        try:
            widget.configure(state=state)
        except Exception:
            pass

    # ── 操作 ──────────────────────────────────────────────────
    def _new_profile(self) -> None:
        self._collect_current()
        name = simpledialog.askstring("Add backend",
                                      "Name for this backend (e.g. deepseek, ollama):", parent=self)
        if not name or not name.strip():
            return
        name = name.strip()
        if name in self._profiles:
            messagebox.showwarning("Name already used", f'"{name}" already exists.', parent=self)
            return
        self._profiles[name] = config.ai_profile_from_dict(name, {"type": "openai"})
        self._current = name
        self._refresh_list()
        self._load_profile(name)

    def _delete_profile(self) -> None:
        if self._current is None:
            return
        name = self._current
        if len(self._profiles) <= 1:
            messagebox.showwarning("Cannot remove", "At least one backend must remain.", parent=self)
            return
        if not messagebox.askyesno("Remove backend", f'Remove "{name}"?', parent=self):
            return
        del self._profiles[name]
        if self._active == name:
            self._active = next(iter(self._profiles))
        self._current = next(iter(self._profiles))
        self._refresh_list()
        self._load_profile(self._current)

    def _set_active(self) -> None:
        if self._current is None:
            return
        self._collect_current()
        self._active = self._current
        self._refresh_list()
        self._status.configure(text=f'"{self._active}" is now the active backend '
                                    f"(click Save to write config.json)", fg=theme.TEXT)

    def _test(self) -> None:
        if self._current is None:
            return
        self._collect_current()
        profile = self._profiles[self._current]
        problems = profile.problems()
        if problems:
            messagebox.showwarning("Incomplete configuration",
                                   "\n".join(f"- {p}" for p in problems), parent=self)
            return
        self._status.configure(text=f'Testing "{profile.name}" ...', fg=theme.MUTED)
        started = self._runner.run(lambda: ai.probe(profile),
                                   on_done=self._on_test_ok, on_error=self._on_test_fail)
        if not started:
            self._status.configure(text="A task is already running, please wait", fg=theme.DANGER)

    def _on_test_ok(self, reply) -> None:
        self._status.configure(text=f"Connection OK, model replied: {str(reply)[:40]}", fg=theme.ACCENT)

    def _on_test_fail(self, ex: BaseException) -> None:
        self._status.configure(text=f"Test failed: {ex}", fg=theme.DANGER)

    def _save(self) -> None:
        self._collect_current()
        self._cfg.ai_profiles = self._profiles
        self._cfg.ai_active = self._active
        try:
            config.save(self._cfg)
        except Exception as ex:
            messagebox.showerror("Could not save", f"{type(ex).__name__}: {ex}", parent=self)
            return

        if self._on_saved is not None:
            self._on_saved(self._cfg)

        problems = self._cfg.ai_problems()
        if problems:
            self._status.configure(text="Saved, but the active backend is not usable yet", fg=theme.DANGER)
            messagebox.showwarning("Saved, but the active backend has problems",
                                   "\n".join(f"- {p}" for p in problems), parent=self)
        else:
            self._status.configure(text=f"Saved; active backend: {self._active}", fg=theme.ACCENT)

    def _close(self) -> None:
        self._runner.stop()
        self.destroy()
