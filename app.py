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

# 单实例互斥体：会话级名字（不加 Global\ 前缀），不同用户各跑各的互不影响
_SINGLE_INSTANCE_NAME = "AIEmailManager.single_instance"
_ERROR_ALREADY_EXISTS = 183
_mutex_handle = None       # 进程存活期间必须一直持有，不能释放


def _another_instance_running() -> bool:
    """已经有一个实例在跑吗？

    两个实例会各自开一份监听：同一封新邮件被两边分别送 AI（双倍花费）、
    分别写 Categories（互相覆盖），还会同时读写 watch_state.json 与 config.json。

    判断不了（非 Windows / API 不可用）就放行 —— 绝不因为"检测失败"把程序挡在门外。
    """
    global _mutex_handle
    try:
        import ctypes
        # use_last_error=True：ctypes 会把这次调用的 LastError 存进它自己的槽，
        # 之后用 get_last_error() 取才可靠（直接调 kernel32.GetLastError() 中间可能被覆盖）。
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p)
        kernel32.CloseHandle.restype = ctypes.c_bool
        kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)

        handle = kernel32.CreateMutexW(None, False, _SINGLE_INSTANCE_NAME)
        if not handle:                  # 建不出来就别拦（命名空间/权限问题）
            return False
        if ctypes.get_last_error() == _ERROR_ALREADY_EXISTS:
            kernel32.CloseHandle(handle)
            return True
        _mutex_handle = handle          # 持有到进程结束；释放就等于放弃保护
        return False
    except Exception:
        return False


def _alert(title: str, message: str) -> None:
    """不依赖 Tk 的提示框（Tk 起不来的时候也得能把话说出来）。"""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(None, message, title, 0x40)   # MB_ICONINFORMATION
    except Exception:
        pass


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

    if _another_instance_running():
        _alert("AI Email Manager",
               "AI Email Manager is already running.\n\n"
               "Two copies would both watch the inbox and overwrite each other's "
               "categories, so this one will exit now.\n\n"
               "Switch to the window that is already open (check the taskbar).")
        return 1

    import tkinter as tk

    from gui.main_window import MainWindow

    try:
        root = tk.Tk()
    except Exception as ex:
        # 这里**不能**用 tkinter 的 messagebox：它自己也要一个可用的 Tk，
        # 会再抛一次异常，用户在无桌面会话里就什么都看不到（连指引都没有）。
        _alert("Could not start the interface",
               f"Creating the window failed: {ex}\n\n"
               f"If you are in a session without a desktop, use the command line "
               f"instead: py -3 run.py check")
        return 1

    MainWindow(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
