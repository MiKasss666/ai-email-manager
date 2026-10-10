"""app.py — 图形界面入口。

用法：
    py -3 app.py          打开界面
    双击 app.py（若已关联）也可以

界面只是一层外壳：User info 写 config.json，Categorize / Sum 在后台线程调用
core.classifier / core.summarizer，日志实时打到「运行过程」。

命令行入口仍然可用，见 run.py。
"""
from __future__ import annotations

import sys


def _setup_console() -> None:
    """双击运行时没有控制台，内核里的 print 会走 config.log 兜住；这里只处理有控制台的情况。"""
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _enable_dpi_awareness() -> None:
    """Windows 高分屏下让文字不发虚（Tk 会据此自动按 DPI 缩放字体）。"""
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def _set_app_user_model_id(app_id: str = "AIEmailManager.1") -> None:
    """让 Windows 任务栏按我们自己的图标显示，而不是 python.exe 的图标。

    必须在创建任何窗口之前调用。右键「固定到任务栏」也靠它区分进程。
    """
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def main() -> int:
    _setup_console()
    _enable_dpi_awareness()
    _set_app_user_model_id()

    import tkinter as tk
    from tkinter import messagebox

    from gui.main_window import MainWindow

    try:
        root = tk.Tk()
    except Exception as ex:
        messagebox.showerror("Could not start the interface",
                             f"Creating the window failed: {ex}\n\n"
                             f"If you are in a session without a desktop, use the command line "
                             f"instead: py -3 run.py check")
        return 1

    MainWindow(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
