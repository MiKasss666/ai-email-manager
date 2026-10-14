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


class RunToken:
    """一次任务的「身份证 + 取消信号」。

    为什么需要：Force reset 之后，被放弃的线程**杀不掉**（通常正卡在 Outlook 的
    模态对话框上），它会继续往下跑。没有令牌的时候，它的结果会：

      · 被当成当前任务的结果，打到**新任务的回调**上（Categorize 的 dict 落到 Sum 的
        回调里会直接抛 KeyError）；
      · 把界面状态清零（`busy=False`），于是用户还能再点出第三个任务 ——
        三个线程同时给同一个邮箱写分类，标签互相覆盖。

    有了令牌：结果按「代次」认领，过期的直接丢弃；`cancel` 事件让旧任务在下一个
    安全点自己停下来（比如下一批 AI 调用之前、下一封写回之前）。
    """

    __slots__ = ("generation", "cancel")

    def __init__(self, generation: int):
        self.generation = generation
        self.cancel = threading.Event()

    @property
    def cancelled(self) -> bool:
        return self.cancel.is_set()

    def should_stop(self) -> bool:
        """给长任务用：该不该现在收手。"""
        return self.cancel.is_set()


class TaskRunner:
    """一次只允许跑一个任务，避免用户连点按钮把邮箱写乱。"""

    def __init__(self, widget: tk.Misc,
                 on_callback_error: Optional[Callable[[str], None]] = None):
        self._widget = widget
        self._results: "queue.Queue[Tuple[int, str, Any]]" = queue.Queue()
        self._busy = False
        self._started_at: Optional[float] = None
        self._after_id: Optional[str] = None
        self._on_done: Optional[Callable[[Any], None]] = None
        self._on_error: Optional[Callable[[BaseException], None]] = None
        self._generation = 0                 # 每启动一个任务 +1
        self._token: Optional[RunToken] = None
        self._local = threading.local()      # 工作线程各自的令牌
        # 回调自己抛异常时往哪儿报（默认没地方报 —— 打包成 --windowed 后 print 是空操作）
        self.on_callback_error = on_callback_error
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

    @property
    def token(self) -> Optional[RunToken]:
        """**当前工作线程**的任务令牌（主线程里为 None）。

        用线程本地而不是实例属性，是因为被放弃的任务可能还在跑 —— 它必须拿到
        自己那一份令牌，而不是新任务刚建的那一份。
        """
        return getattr(self._local, "token", None)

    def should_stop(self) -> bool:
        """当前工作线程的任务是否已被放弃（没有令牌时恒为 False）。

        直接当回调用：`classify(..., should_stop=runner.should_stop)`。
        """
        token = self.token
        return bool(token is not None and token.cancelled)

    def force_reset(self) -> None:
        """把一个卡死的任务标记为结束，让界面重新可用。

        卡住的线程杀不掉（通常正卡在 Outlook 的某个模态对话框上），这里做三件事：

          1. `cancel.set()`：通知它**在下一个安全点停下来** —— 不再继续写邮箱；
          2. 代次 +1：它后面才回来的结果会被丢弃，不会再打到新任务的回调上；
          3. 解除界面封锁。

        残留线程仍然是 daemon，程序退出时被一起收掉。
        """
        if self._token is not None:
            self._token.cancel.set()
        self._generation += 1
        self._token = None
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

        self._generation += 1
        token = RunToken(self._generation)
        self._token = token
        self._busy = True
        self._started_at = time.time()
        self._on_done = on_done
        self._on_error = on_error

        def worker() -> None:
            # 令牌放线程本地：被放弃的任务后知后觉回来时，读到的仍是**它自己**那份
            self._local.token = token
            _com_initialize()
            try:
                result = work()
            except BaseException as ex:   # 后台异常必须兜住，否则界面会永远卡在 busy
                self._results.put((token.generation, "error", ex))
            else:
                self._results.put((token.generation, "done", result))
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
                generation, kind, payload = self._results.get_nowait()
                if generation != self._generation:
                    # 被放弃的旧任务迟到的结果：丢弃。绝不能让它清掉 busy，
                    # 也不能让它打到当前任务的回调上（那会把界面状态搅乱）。
                    continue
                self._busy = False
                self._started_at = None
                callback = self._on_done if kind == "done" else self._on_error
                if callback is not None:
                    try:
                        callback(payload)
                    except Exception as ex:
                        self._report_callback_error(ex)
        except queue.Empty:
            pass

        try:
            self._after_id = self._widget.after(POLL_MS, self._poll)
        except Exception:
            self._after_id = None      # 窗口已销毁

    def _report_callback_error(self, ex: BaseException) -> None:
        """回调自己抛异常：写进日志区，别用 print。

        `--windowed` 打包后 `sys.stdout is None`，print 是静默空操作 —— 异常会
        彻底消失（日志区、app.log、弹窗里都没有）。而 watcher 的 `_on_done`
        把「排下一轮」放在最后，一旦回调中途抛异常，**自动监听就静默死掉**，
        而日志正是官方唯一给出的排查渠道。
        """
        message = f"[ui] callback error: {type(ex).__name__}: {ex}"
        if self.on_callback_error is None:
            return
        try:
            self.on_callback_error(message)
        except Exception:
            pass
