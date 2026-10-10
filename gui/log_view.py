"""gui/log_view.py — 「运行过程」输出区 + 落盘日志。

线程安全的关键：Tkinter 的 widget 只能在主线程碰。
所以后台线程只往队列里投递文本，主线程定时取出来写进 widget。
内核里的 log= 回调直接传 self.write 即可（它是 print 风格的可调用对象）。

同时把同样的内容**追加写入日志文件**（默认 app.log，在 config.json 旁边）——
出问题时可以直接把文件发出来，不用靠截图。文件里的每行带 HH:MM:SS 时间戳。
"""
from __future__ import annotations

import os
import queue
import time
import tkinter as tk
from tkinter import scrolledtext
from typing import Any, Optional

from core import config

from . import theme

MAX_LINES = 2000            # widget 里最多保留多少行，避免长时间运行把内存吃光
DRAIN_MS = 120              # 主线程取日志的间隔
MAX_LOG_BYTES = 1024 * 1024  # 日志文件超过 1 MB 就在下次启动时清空，避免无限增长
DEFAULT_LOG_NAME = "app.log"


def default_log_path() -> str:
    """默认日志文件路径：和应用配置放在一起（打包后就是 exe 所在目录）。"""
    return os.path.join(config.app_dir(), DEFAULT_LOG_NAME)


class LogView(tk.Frame):
    """带标题栏、「复制」「清空」按钮，并会把内容落盘的日志区。"""

    def __init__(self, master: tk.Misc, log_path: Optional[str] = None, **kwargs):
        super().__init__(master, bg=theme.CARD, **kwargs)

        self._log_path = log_path
        self._queue: "queue.Queue[str]" = queue.Queue()
        self._after_id: Optional[str] = None
        self._alive = True

        self._prepare_log_file()

        header = tk.Frame(self, bg=theme.CARD)
        header.pack(fill="x", padx=theme.PAD, pady=(theme.PAD, 0))
        tk.Label(header, text="Activity log", bg=theme.CARD, fg=theme.TEXT,
                 font=theme.FONT_SECTION).pack(side="left")
        for text, command in (("Clear", self.clear), ("Copy", self.copy_to_clipboard)):
            tk.Button(header, text=text, command=command, relief="flat",
                      bg=theme.CARD, fg=theme.MUTED, activebackground=theme.CARD,
                      font=theme.FONT, cursor="hand2", bd=0).pack(side="right", padx=(8, 0))

        self._text = scrolledtext.ScrolledText(
            self, height=16, wrap="word", font=theme.FONT_MONO,
            bg=theme.LOG_BG, fg=theme.LOG_FG, insertbackground=theme.LOG_FG,
            relief="flat", bd=0, padx=10, pady=8,
        )
        self._text.pack(fill="both", expand=True, padx=theme.PAD, pady=(6, theme.PAD))
        self._text.configure(state="disabled")

        self.bind("<Destroy>", self._on_destroy)
        self._schedule()

    # ── 对外（任何线程都可调用）────────────────────────────────
    def write(self, *args: Any) -> None:
        """print 风格写入。内核的 log= 回调直接用它。"""
        self._queue.put(" ".join(str(a) for a in args) + "\n")

    def clear(self) -> None:
        self._text.configure(state="normal")
        self._text.delete("1.0", "end")
        self._text.configure(state="disabled")

    def contents(self) -> str:
        """当前日志全文（自测 / 复制用）。"""
        return self._text.get("1.0", "end")

    def log_file(self) -> Optional[str]:
        """日志文件路径（没启用落盘时返回 None）。"""
        return self._log_path

    def copy_to_clipboard(self) -> None:
        """把整段日志放进剪贴板 —— 出问题时直接粘给我就行，不用截图。"""
        try:
            self.clipboard_clear()
            self.clipboard_append(self.contents())
        except Exception:
            pass

    # ── 内部 ──────────────────────────────────────────────────
    def _prepare_log_file(self) -> None:
        """启动时处理日志文件：太大就清空，并写一条分隔线。"""
        if not self._log_path:
            return
        try:
            if os.path.exists(self._log_path) and os.path.getsize(self._log_path) > MAX_LOG_BYTES:
                os.remove(self._log_path)
            with open(self._log_path, "a", encoding="utf-8") as handle:
                handle.write(f"\n===== session started {time.strftime('%Y-%m-%d %H:%M:%S')} "
                             f"=====\n")
        except Exception:
            self._log_path = None      # 写不了就静默放弃落盘，不影响界面

    def _append_log_file(self, text: str) -> None:
        if not self._log_path or not text:
            return
        stamp = time.strftime("%H:%M:%S")
        stamped = "".join(f"[{stamp}] {line}" if line.strip() else line
                          for line in text.splitlines(True))
        try:
            with open(self._log_path, "a", encoding="utf-8") as handle:
                handle.write(stamped)
        except Exception:
            pass

    def _schedule(self) -> None:
        try:
            self._after_id = self.after(DRAIN_MS, self._drain)
        except Exception:
            self._alive = False

    def _drain(self) -> None:
        if not self._alive:
            return

        chunks = []
        try:
            while True:
                chunks.append(self._queue.get_nowait())
        except queue.Empty:
            pass

        if chunks:
            text = "".join(chunks)
            self._text.configure(state="normal")
            self._text.insert("end", text)
            self._trim()
            self._text.see("end")
            self._text.configure(state="disabled")
            self._append_log_file(text)

        self._schedule()

    def _trim(self) -> None:
        """行数超过上限时，从头部删掉多余的行。"""
        last_line = int(self._text.index("end-1c").split(".")[0])
        if last_line > MAX_LINES:
            self._text.delete("1.0", f"{last_line - MAX_LINES + 1}.0")

    def _on_destroy(self, event: tk.Event) -> None:
        if event.widget is not self:      # 子控件销毁也会冒泡到这里，忽略
            return
        self._alive = False
        if self._after_id is not None:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
