"""gui/main_window.py — 主窗口（对应草图图 1）。

    [User info]          打开 User info 表单，点进去改用户信息
    [AI setting]         打开 AI 后端设置（模型 / API Key / 测试连通）
    [Categorize]         后台调用 classifier.classify_inbox()
    [Sum]                后台调用 summarizer.summarize_inbox()
    [Auto-watch]         配置填好后自动开始监听收件箱，见 gui/watcher.py
    Activity log         内核的 log= 回调直接接到日志区

界面层不含任何业务判断：分级/总结的规则全在 core 里，这里只负责按钮、参数和展示。
界面文案统一英文（内核里会显示的字符串也都是英文），代码注释保持中文。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any, Optional

from core import ai, classifier, config, outlook, summarizer

from . import theme, watcher as watch_mod
from .ai_settings_dialog import AiSettingsDialog
from .log_view import LogView, default_log_path
from .task import TaskRunner
from .user_info_dialog import UserInfoDialog

DEFAULT_DAYS = 7
WATCHDOG_MS = 1000          # 状态栏每秒刷新一次
STUCK_HINT_SECONDS = 90     # 超过这个时长还没结果，就提示用户去看 Outlook


class MainWindow:
    """主窗口装配器（不是 tk.Frame 子类：它管理的是整个 root）。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.cfg = config.ensure()
        self.runner = TaskRunner(root)

        root.title("AI Email Manager")
        root.configure(bg=theme.BG)
        root.minsize(820, 660)
        theme.apply(root)
        theme.apply_icon(root)      # 同时把之后创建的 Toplevel 也设成同一图标

        self._build()

        # 监听控制器要用到日志区，所以放在 _build 之后创建
        self.watcher = watch_mod.WatchController(
            root, self.runner, lambda: self.cfg, self.log.write, self._on_watch_notify)
        self._interval_var.set(str(self.watcher.interval))

        self._refresh_header()
        self.log.write("Ready.")
        self.log.write(f"Config file: {config.config_path()}")
        self.log.write(f"Log file: {self.log.log_file() or '(disabled)'}")
        level = outlook.integrity_level()
        self.log.write(f"[env] security context: {level} integrity")
        self.log.write(f"[env] launched by: {outlook.parent_process()}")
        if level in ("Low", "Untrusted"):
            self.log.write("[env] WARNING: Outlook COM is blocked at this integrity level - "
                           "launch Email Manager normally (not from a sandboxed shell/agent)")
        self._report_problems()
        self._refresh_status()

        root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._hinted_stuck = False
        self._watchdog_id = root.after(WATCHDOG_MS, self._watchdog)
        # 配置齐全就自动开始监听；没填完就先不开，并说明缺什么
        self._maybe_start_watching("startup")

    # ── 界面 ──────────────────────────────────────────────────
    def _build(self) -> None:
        self._build_header()
        self._build_actions()

        log_card = theme.card(self.root)
        log_card.pack(fill="both", expand=True, padx=theme.PAD, pady=(theme.PAD, 0))
        self.log = LogView(log_card, log_path=default_log_path())
        self.log.pack(fill="both", expand=True)

        self._status = tk.Label(self.root, text="", bg=theme.BG, fg=theme.MUTED,
                                font=theme.FONT, anchor="w")
        self._status.pack(fill="x", padx=theme.PAD, pady=(6, theme.PAD))

    def _build_header(self) -> None:
        card = theme.card(self.root)
        card.pack(fill="x", padx=theme.PAD, pady=(theme.PAD, 0))
        inner = tk.Frame(card, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        inner.pack(fill="x")

        tk.Label(inner, text="AI Email Manager", bg=theme.CARD, fg=theme.TEXT,
                 font=theme.FONT_TITLE).pack(anchor="w")
        self._header_info = tk.Label(inner, text="", bg=theme.CARD, fg=theme.MUTED,
                                     font=theme.FONT, justify="left", anchor="w",
                                     wraplength=760)
        self._header_info.pack(anchor="w", pady=(4, 10))

        buttons = tk.Frame(inner, bg=theme.CARD)
        buttons.pack(fill="x")
        self._btn_user = ttk.Button(buttons, text="User info", command=self._open_user_info)
        self._btn_user.pack(side="left")
        self._btn_ai = ttk.Button(buttons, text="AI setting", command=self._open_ai_settings)
        self._btn_ai.pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Reload config", command=self._reload_config).pack(side="right")

    def _build_actions(self) -> None:
        card = theme.card(self.root)
        card.pack(fill="x", padx=theme.PAD, pady=(theme.PAD, 0))

        row = tk.Frame(card, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        row.pack(fill="x")
        self._btn_categorize = theme.action_button(row, "Categorize",
                                                   self._run_categorize, theme.PRIMARY)
        self._btn_categorize.pack(side="left")
        self._btn_sum = theme.action_button(row, "Sum", self._run_sum, theme.ACCENT)
        self._btn_sum.pack(side="left", padx=(10, 0))

        options = tk.Frame(card, bg=theme.CARD, padx=theme.PAD)
        options.pack(fill="x", pady=(0, 8))

        tk.Label(options, text="Scope: last", bg=theme.CARD, fg=theme.TEXT,
                 font=theme.FONT).pack(side="left")
        self._days = tk.StringVar(value=str(DEFAULT_DAYS))
        ttk.Spinbox(options, from_=1, to=365, textvariable=self._days, width=5).pack(side="left", padx=6)
        tk.Label(options, text="days", bg=theme.CARD, fg=theme.TEXT, font=theme.FONT).pack(side="left")

        self._dry_run = tk.BooleanVar(value=False)
        ttk.Checkbutton(options, text="Dry run (print only, no write-back)",
                        variable=self._dry_run).pack(side="left", padx=(18, 0))
        self._force = tk.BooleanVar(value=False)
        ttk.Checkbutton(options, text="Include already classified",
                        variable=self._force).pack(side="left", padx=(18, 0))

        self._build_watch_row(card)

    def _build_watch_row(self, card: tk.Frame) -> None:
        """自动监听那一行：开关 + 轮询间隔。

        监听开关在任务运行期间**不禁用** —— 用户任何时候都该能把监听关掉。
        """
        row = tk.Frame(card, bg=theme.CARD, padx=theme.PAD, pady=theme.PAD)
        row.pack(fill="x")

        self._watch_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row, text="Auto-watch new mail", variable=self._watch_var,
                        command=self._on_watch_toggle).pack(side="left")

        tk.Label(row, text="poll every", bg=theme.CARD, fg=theme.TEXT,
                 font=theme.FONT).pack(side="left", padx=(18, 4))
        self._interval_var = tk.StringVar(value=str(watch_mod.DEFAULT_INTERVAL))
        spin = ttk.Spinbox(row, from_=watch_mod.MIN_INTERVAL, to=watch_mod.MAX_INTERVAL,
                           textvariable=self._interval_var, width=6)
        spin.pack(side="left")
        spin.bind("<Return>", self._apply_interval)
        spin.bind("<FocusOut>", self._apply_interval)
        tk.Label(row, text="seconds", bg=theme.CARD, fg=theme.TEXT,
                 font=theme.FONT).pack(side="left", padx=(4, 0))

        # 卡死逃生口：Outlook 弹着对话框时 COM 会一直挂着，界面不能就此锁死
        ttk.Button(row, text="Force reset", command=self._force_reset).pack(side="right")

    # ── 状态展示 ──────────────────────────────────────────────
    def _refresh_header(self) -> None:
        cfg = self.cfg
        self._header_info.configure(
            text=f"Email: {cfg.email or '(not set)'}    |    Name: {cfg.name or '(not set)'}\n"
                 f"AI backend: {cfg.ai_active} — {ai.describe(cfg.ai)}")

    def _refresh_status(self) -> None:
        try:
            if self.runner.busy:
                text = f"Working ... {self.runner.elapsed:.0f}s"
            elif self.watcher.running:
                text = f"Watching the inbox every {self.watcher.interval}s"
            else:
                text = "Idle — watching is off"
            self._status.configure(text=text)
        except Exception:
            pass

    def _watchdog(self) -> None:
        """每秒刷一次状态栏，并把「卡住」和「在跑」区分开。

        Outlook 弹着对话框（选配置文件、安全提示）时，COM 调用会一直挂着 ——
        既不返回也不报错。没有这个提示，用户只会看到界面不动却不知道原因。
        """
        try:
            if self.runner.busy:
                elapsed = self.runner.elapsed
                self._status.configure(text=f"Working ... {elapsed:.0f}s")
                if elapsed > STUCK_HINT_SECONDS and not self._hinted_stuck:
                    self._hinted_stuck = True
                    self.log.write(
                        f"[warn] Outlook has not answered for {elapsed:.0f}s. It is most likely "
                        f"showing a dialog (profile picker or security prompt) or still starting. "
                        f"Bring Outlook to the front and dismiss it. If the task really is stuck, "
                        f"click Force reset.")
            else:
                self._hinted_stuck = False
                self._refresh_status()
            self._sync_buttons()      # 每秒自愈一次按钮状态
        except Exception:
            pass
        try:
            self._watchdog_id = self.root.after(WATCHDOG_MS, self._watchdog)
        except Exception:
            self._watchdog_id = None      # 窗口已销毁

    def _force_reset(self) -> None:
        """卡死逃生口：只解除界面封锁，卡住的线程杀不掉。"""
        if not self.runner.busy:
            self.log.write("[ui] nothing to reset — no task is running")
            return
        elapsed = self.runner.elapsed
        if not messagebox.askyesno(
                "Force reset",
                f"A task has been running for {elapsed:.0f}s without finishing.\n\n"
                f"This only unlocks the interface; the stuck operation itself cannot be "
                f"cancelled. A stuck task is usually Outlook waiting on a dialog.\n\n"
                f"Reset now?",
                parent=self.root):
            return
        self.runner.force_reset()
        self._refresh_busy()
        self.log.write(f"[ui] force-reset after {elapsed:.0f}s — the stuck task was abandoned")
        if self.watcher.running:
            # 被放弃的任务永远不会回调，监听链要手动接回去，否则它会静默死掉
            self.watcher.resume_schedule()

    def _report_problems(self) -> None:
        problems = self.cfg.problems()
        for problem in problems:
            self.log.write(f"[config] needs attention: {problem}")
        if not problems:
            self.log.write("[config] OK")

    def _sync_buttons(self) -> None:
        """控件可用性只由 runner.busy 推导，并且每秒重新同步一次。

        这样即使某处漏了一次「解除禁用」，也会在 1 秒内自愈 ——「按钮点不动」
        是用户唯一无法自救的故障，不能只指望回调成对出现。
        监听开关（Auto-watch）永远不禁用：任何时候都要能关掉监听。
        """
        state = "disabled" if self.runner.busy else "normal"
        for widget in (self._btn_categorize, self._btn_sum, self._btn_user, self._btn_ai):
            try:
                if str(widget.cget("state")) != state:
                    widget.configure(state=state)
            except Exception:
                pass

    def _refresh_busy(self) -> None:
        self._sync_buttons()
        self._refresh_status()

    # ── 监听 ──────────────────────────────────────────────────
    def _log_problems(self, prefix: str, problems) -> None:
        self.log.write(f"{prefix} configuration is incomplete:")
        for problem in problems:
            self.log.write(f"{prefix}   - {problem}")

    def _maybe_start_watching(self, reason: str) -> None:
        """配置齐全才自动开监听。

        没填完就开监听只会每 30 秒刷一屏错误，然后连续失败 3 次自动暂停 ——
        不如干脆不开，并把缺什么直接写在日志里。
        """
        if self.watcher.running:
            return
        problems = self.cfg.problems()
        if problems:
            self.log.write(f"[watch] not started ({reason}) - finish the setup first, "
                           f"then it turns on by itself:")
            for problem in problems:
                self.log.write(f"[watch]   - {problem}")
            self._watch_var.set(False)
            self._refresh_status()
            return
        self.watcher.start()

    def _on_watch_toggle(self) -> None:
        if self._watch_var.get():
            problems = self.cfg.problems()
            if problems:
                self._watch_var.set(False)
                self._log_problems("[watch] not started -", problems)
                messagebox.showwarning("Configuration problems",
                                       "\n".join(f"- {p}" for p in problems), parent=self.root)
                self._refresh_status()
                return
            self.watcher.start()
        else:
            self.watcher.stop()
        self._refresh_status()

    def _on_watch_notify(self, kind: str, value: Any) -> None:
        """监听控制器的回调，一定在主线程执行。"""
        if kind == "running":
            self._watch_var.set(bool(value))
        elif kind == "busy":
            self._refresh_busy()
        self._refresh_status()

    def _apply_interval(self, _event: Optional[tk.Event] = None) -> None:
        self.watcher.set_interval(self._interval_var.get())
        self._interval_var.set(str(self.watcher.interval))   # 回填被夹取后的值
        self._refresh_status()

    # ── 设置窗口 ──────────────────────────────────────────────
    def _open_user_info(self) -> None:
        UserInfoDialog(self.root, self.cfg, on_saved=self._on_settings_saved)

    def _open_ai_settings(self) -> None:
        AiSettingsDialog(self.root, self.cfg, on_saved=self._on_settings_saved)

    def _on_settings_saved(self, new_cfg: config.Config) -> None:
        self.cfg = new_cfg
        self._refresh_header()
        self.log.write(f"[config] saved to {config.config_path()}")
        for problem in self.cfg.problems():
            self.log.write(f"[config] needs attention: {problem}")
        # 刚把配置填全 → 不用重启，立刻开始监听
        self._maybe_start_watching("setup complete")

    def _reload_config(self) -> None:
        self.cfg = config.ensure()
        self._refresh_header()
        self.log.write("[config] reloaded config.json")
        self._report_problems()

    # ── 手动任务 ──────────────────────────────────────────────
    def _days_int(self) -> int:
        try:
            value = int(self._days.get())
        except (TypeError, ValueError):
            value = DEFAULT_DAYS
        value = max(1, min(365, value))
        self._days.set(str(value))
        return value

    def _blocked_by_config(self) -> bool:
        problems = self.cfg.problems()
        if not problems:
            return False
        messagebox.showwarning("Configuration problems",
                               "\n".join(f"- {p}" for p in problems), parent=self.root)
        return True

    def _run_categorize(self) -> None:
        if self.runner.busy or self._blocked_by_config():
            return

        cfg = self.cfg
        days = self._days_int()
        dry_run = self._dry_run.get()
        force = self._force.get()

        # 这是唯一会改动邮箱的操作，写之前必须确认一次
        if not dry_run:
            if not messagebox.askyesno(
                    "Confirm classification",
                    f"Read the last {days} day(s) of your inbox, classify them with AI "
                    f"and write the S/A/B/C categories back to Outlook.\n\n"
                    f"This modifies categories and flags on your real mail. Continue?",
                    parent=self.root):
                return

        self.log.write("")
        self.log.write(f"=== Categorize started (last {days} day(s)"
                       + (", dry run - no write-back" if dry_run else "")
                       + (", including already classified" if force else "") + ") ===")
        started = self.runner.run(
            lambda: classifier.classify_inbox(cfg, days=days, dry_run=dry_run,
                                              force=force, log=self.log.write),
            on_done=self._on_categorize_done,
            on_error=self._on_task_error,
        )
        if started:
            self._refresh_busy()

    def _on_categorize_done(self, stats: Any) -> None:
        self._refresh_busy()
        self.log.write(f"Done: read {stats['read']}, S={stats['S']} A={stats['A']} "
                       f"B={stats['B']} C={stats['C']}, ok {stats['ok']}, failed {stats['failed']}")

    def _run_sum(self) -> None:
        if self.runner.busy or self._blocked_by_config():
            return

        cfg = self.cfg
        self.log.write("")
        self.log.write("=== Sum started (all levels S+A+B+C) ===")
        started = self.runner.run(
            lambda: summarizer.summarize_inbox(cfg, "sabc", log=self.log.write),
            on_done=self._on_sum_done,
            on_error=self._on_task_error,
        )
        if started:
            self._refresh_busy()

    def _on_sum_done(self, count: Any) -> None:
        self._refresh_busy()
        self.log.write(f"Done: {count} mail(s) summarised" if count
                       else "Done: nothing to summarise, no report sent")

    def _on_task_error(self, ex: BaseException) -> None:
        self._refresh_busy()
        self.log.write(f"[error] {type(ex).__name__}: {ex}")
        messagebox.showerror("Something went wrong",
                             f"{type(ex).__name__}: {ex}", parent=self.root)

    # ── 关闭 ──────────────────────────────────────────────────
    def _on_close(self) -> None:
        if self.runner.busy:
            if not messagebox.askyesno("Task in progress",
                                       "A task is still running. Quit anyway?", parent=self.root):
                return
        if self._watchdog_id is not None:
            try:
                self.root.after_cancel(self._watchdog_id)
            except Exception:
                pass
            self._watchdog_id = None
        self.watcher.stop()
        self.runner.stop()
        try:
            self.root.destroy()
        except Exception:
            pass
