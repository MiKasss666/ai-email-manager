"""gui/task.py — 后台任务执行器。

界面上的「分类」「总结」要跑几十秒到几分钟（AI 调用 + Outlook COM），
必须放到后台线程，否则窗口会一直假死。

两条线程规则（踩过就知道疼）：
  1. Tkinter 的 widget 只能在主线程碰 —— 所以结果不直接回调，
     而是丢进 queue，由主线程的 after 轮询取出来执行。
  2. 后台线程要用 Outlook COM 必须先 CoInitialize()，否则报
     「尚未调用 CoInitialize」；任务结束后要 CoUninitialize() 配对释放。
"""
from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from typing import Any, Callable, Optional, Tuple

try:
    import pythoncom
except ImportError:      # 没装 pywin32 时界面仍能打开，只是真正跑任务才会报错
    pythoncom = None

POLL_MS = 100


def _com_initialize() -> None:
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass


def _com_uninitialize() -> None:
    if pythoncom is not None:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class TaskRunner:
    """一次只允许跑一个任务，避免用户连点按钮把邮箱写乱。"""

    def __init__(self, widget: tk.Misc):
        self._widget = widget
        self._results: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self._busy = False
        self._started_at: Optional[float] = None
        self._after_id: Optional[str] = None
        self._on_done: Optional[Callable[[Any], None]] = None
        self._on_error: Optional[Callable[[BaseException], None]] = None
        self._poll()

    # ── 状态 ──────────────────────────────────────────────────
    @property
    def busy(self) -> bool:
        return self._busy

    @property
    def elapsed(self) -> float:
        """当前任务已经跑了多少秒；空闲时为 0。

        用来把「界面卡住了」和「任务真的在跑」区分开 —— Outlook 弹着对话框时
        COM 调用会一直挂着（既不返回也不报错），用户需要看得见时间在走。
        """
        if not self._busy or self._started_at is None:
            return 0.0
        return time.time() - self._started_at

    def force_reset(self) -> None:
        """把一个卡死的任务标记为结束，让界面重新可用。

        卡住的线程杀不掉（通常正卡在 Outlook 的某个模态对话框上），
        这里只是解除界面封锁 —— 所以只应在用户确认「确实卡住了」之后调用。
        残留线程是 daemon，程序退出时会被一起收掉。
        """
        self._busy = False
        self._started_at = None

    # ── 对外 ──────────────────────────────────────────────────
    def run(self, work: Callable[[], Any], *, on_done: Optional[Callable[[Any], None]] = None,
            on_error: Optional[Callable[[BaseException], None]] = None) -> bool:
        """在后台线程执行 work()。

        返回 False 表示已有任务在跑，本次被拒绝。
        on_done / on_error 保证在主线程被调用，可以安全操作界面。
        """
        if self._busy:
            return False

        self._busy = True
        self._started_at = time.time()
        self._on_done = on_done
        self._on_error = on_error

        def worker() -> None:
            _com_initialize()
            try:
                result = work()
            except BaseException as ex:   # 后台异常必须兜住，否则界面会永远卡在 busy
                self._results.put(("error", ex))
            else:
                self._results.put(("done", result))
            finally:
                _com_uninitialize()

        threading.Thread(target=worker, daemon=True, name="email-manager-task").start()
        return True

    def stop(self) -> None:
        """窗口关闭时调用，停掉轮询。"""
        if self._after_id is not None:
            try:
                self._widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    # ── 主线程轮询 ────────────────────────────────────────────
    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self._results.get_nowait()
                self._busy = False
                self._started_at = None
                callback = self._on_done if kind == "done" else self._on_error
                if callback is not None:
                    try:
                        callback(payload)
                    except Exception as ex:
                        print(f"[界面] 回调出错：{type(ex).__name__}: {ex}")
        except queue.Empty:
            pass

        try:
            self._after_id = self._widget.after(POLL_MS, self._poll)
        except Exception:
            self._after_id = None      # 窗口已销毁
