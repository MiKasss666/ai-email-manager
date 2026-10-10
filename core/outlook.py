"""core/outlook.py — Outlook 经典版（桌面客户端）COM 交互层。

只提供「原子操作」：连接、分类维护、读信、写回、发信。
业务判断（该判几级、总结些什么）不在这里 —— 见 classifier.py / summarizer.py。

依赖：Outlook 经典版桌面客户端（OUTLOOK.EXE）+ pywin32。
      新版 Outlook / 网页版没有可用的 MAPI COM 接口，本模块不适用。
"""
from __future__ import annotations

import os
import re
import subprocess
import time
from datetime import datetime, timedelta
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import pythoncom  # noqa: F401  pywin32 运行时依赖，显式导入便于打包时被发现
import win32com.client

from .config import LEVELS

# ── Outlook 枚举常量 ──────────────────────────────────────────────
OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 0
OL_FORMAT_HTML = 2
OL_FLAG_NONE = 0
OL_FLAG_MARKED = 2

# OlCategoryColor 枚举值
CATEGORY_COLORS = {"S": 2, "A": 9, "B": 8, "C": 13}  # 橙 / 紫 / 蓝 / 灰
LEGACY_CATEGORY_NAMES = ("S 紧急", "A 重要", "B 普通", "C 可忽略")

SUM_NAME = "Sum"
SUM_COLOR = 4  # 黄

# 我们自己生成的总结邮件，主题里一定含这个词 → 扫收件箱时要跳过，别把自己的报告再分一次级
REPORT_SUBJECT_MARK = "邮件总结"        # 中文标记（历史版本的报告用的就是它）
REPORT_SUBJECT_MARK_EN = "Email summary"   # 英文标记（language = en 时用）
# 两个都认：报告主题跟随 language，但用户随时可能切换语言，
# 历史报告（另一种语言的主题）也必须继续被认出来 ——
# 否则程序会把自己发的报告当成普通邮件去分级。
REPORT_SUBJECT_MARKS = (REPORT_SUBJECT_MARK, REPORT_SUBJECT_MARK_EN)
# 用户发给自己的总结指令：sum / sum s / sum sa / sum sabc
# 属于「收件箱约定」，统一在这里解析（summarizer 复用本题的正则）
SUM_COMMAND_RE = re.compile(r"^sum(?:\s+([sabc]+))?$", re.IGNORECASE)


# ── 连接 ──────────────────────────────────────────────────────────
OL_WINDOW_MINIMIZED = 1           # OlWindowState.olMinimized
OUTLOOK_MANUAL_HINT = ("Outlook is not running. Open Outlook classic first and wait until "
                       "your inbox shows up, then try again.")
# 启动 Outlook 的退避：只约束「主动启动」，**绝不约束「附加到已在运行的实例」**
# （早期版本把退避挡在 Dispatch 前面，导致用户手动开了 Outlook 也连不上）
LAUNCH_BACKOFF_SECONDS = 300
OUTLOOK_START_TIMEOUT = 30        # 启动 OUTLOOK.EXE 后等它出现的上限（秒）
# 说明：这里**故意没有**"启动失败后的冷却"之类的守卫 —— 见 connect() 的文档字符串。
# 重试节奏交给调用方（监听那边对"Outlook 暂时用不了"单独退避）。

# 上次尝试启动 Outlook 的时刻 / 上次打印退避提示的时刻
_last_launch_attempt = 0.0
_backoff_logged_at = 0.0
# "例行附加成功"这类心跳日志已经打过哪些 key（见 _notice）
_seen_notices: set = set()


class OutlookNotRunning(RuntimeError):
    """Outlook 当前不可用（没开、正在启动还没就绪、或启动失败）。

    单独做一种异常类型，是为了让上层能区分「Outlook 暂时用不了」和「真的出错了」：
    前者是正常情况，不该被当成失败去累计、更不该触发自动暂停。
    """


def _process_running(exe_name: str) -> bool:
    """系统里有没有某个进程在跑（用 Windows API 枚举，不启动子进程）。

    不用 tasklist 是因为要捕获它的输出（管道），而管道在某些受限环境下会被拦；
    直接用 CreateToolhelp32Snapshot 更快也更稳。
    """
    if os.name != "nt":
        return False
    target = (exe_name or "").strip().lower()
    if not target:
        return False
    try:
        import ctypes
        from ctypes import wintypes

        TH32CS_SNAPPROCESS = 0x00000002
        INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [("dwSize", wintypes.DWORD),
                        ("cntUsage", wintypes.DWORD),
                        ("th32ProcessID", wintypes.DWORD),
                        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
                        ("th32ModuleID", wintypes.DWORD),
                        ("cntThreads", wintypes.DWORD),
                        ("th32ParentProcessID", wintypes.DWORD),
                        ("pcPriClassBase", ctypes.c_long),
                        ("dwFlags", wintypes.DWORD),
                        ("szExeFile", ctypes.c_char * 260)]

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        # 必须显式声明参数与返回类型：ctypes 默认按 32 位 int 处理，
        # 而 CreateToolhelp32Snapshot 返回的是 64 位 HANDLE，不声明会被截断
        # （截断后第一个进程都读不出来，函数会静默返回 False）。
        kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        kernel32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
        kernel32.Process32First.restype = wintypes.BOOL
        kernel32.Process32First.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32)]
        kernel32.Process32Next.restype = wintypes.BOOL
        kernel32.Process32Next.argtypes = [wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32)]
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

        snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if not snapshot or snapshot == INVALID_HANDLE_VALUE:
            return False
        try:
            entry = PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
            if not kernel32.Process32First(snapshot, ctypes.byref(entry)):
                return False
            while True:
                name = entry.szExeFile.decode("ascii", "ignore").strip().lower()
                if name == target:
                    return True
                if not kernel32.Process32Next(snapshot, ctypes.byref(entry)):
                    break
        finally:
            kernel32.CloseHandle(snapshot)
    except Exception:
        return False
    return False


def parent_process() -> str:
    """父进程的映像名和 PID —— 用来判断"是谁把本程序拉起来的"。

    排障价值：被代理 / 沙箱界面拉起时，父进程会是那个代理（例如
    "DeepSeek Harness.exe"），而不是 explorer.exe；而父进程是 Low 完整性的话，
    本进程也就是 Low，于是连不上 Outlook。这一行能直接终结"到底是谁启动的"的猜测。
    """
    pid = os.getppid()
    name = "unknown"
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes

            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.OpenProcess.restype = wintypes.HANDLE
            kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
            kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
            kernel32.QueryFullProcessImageNameW.argtypes = [
                wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR,
                ctypes.POINTER(wintypes.DWORD)]
            kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

            handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
            if handle:
                try:
                    buf = ctypes.create_unicode_buffer(1024)
                    size = wintypes.DWORD(len(buf))
                    if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                        name = os.path.basename(buf.value) or name
                finally:
                    kernel32.CloseHandle(handle)
        except Exception:
            pass
    return f"{name} (pid {pid})"


def _outlook_process_running() -> bool:
    return _process_running("outlook.exe")


def integrity_level() -> str:
    """当前进程的完整性级别：Untrusted / Low / Medium / High / System / unknown。

    为什么需要它：**Low 完整性的进程无法通过 COM 启动（甚至无法连接）Medium 完整性的
    Outlook** —— 表现就是 `0x80080005 CO_E_SERVER_EXEC_FAILURE`
    （com_error: 'Server execution failed'）外加 Outlook 自己弹出的
    "Cannot start Microsoft Outlook."。这个信息直接决定排障方向，
    所以启动时就写进日志，别让人对着一个 COM 错误码猜。
    """
    if os.name != "nt":
        return "unknown"
    try:
        import ctypes
        from ctypes import wintypes

        TOKEN_QUERY = 0x0008
        TOKEN_INTEGRITY_LEVEL = 25

        class SID_AND_ATTRIBUTES(ctypes.Structure):
            _fields_ = [("Sid", ctypes.c_void_p), ("Attributes", wintypes.DWORD)]

        class TOKEN_MANDATORY_LABEL(ctypes.Structure):
            _fields_ = [("Label", SID_AND_ATTRIBUTES)]

        advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        advapi32.OpenProcessToken.restype = wintypes.BOOL
        advapi32.OpenProcessToken.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                              ctypes.POINTER(wintypes.HANDLE)]
        advapi32.GetTokenInformation.restype = wintypes.BOOL
        advapi32.GetTokenInformation.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p,
                                                 wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]
        kernel32.GetCurrentProcess.restype = wintypes.HANDLE
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

        token = wintypes.HANDLE()
        if not advapi32.OpenProcessToken(kernel32.GetCurrentProcess(), TOKEN_QUERY,
                                        ctypes.byref(token)):
            return "unknown"
        try:
            size = wintypes.DWORD()
            advapi32.GetTokenInformation(token, TOKEN_INTEGRITY_LEVEL, None, 0, ctypes.byref(size))
            if not size.value:
                return "unknown"
            buf = ctypes.create_string_buffer(size.value)
            if not advapi32.GetTokenInformation(token, TOKEN_INTEGRITY_LEVEL, buf, size.value,
                                                ctypes.byref(size)):
                return "unknown"
            label = ctypes.cast(buf, ctypes.POINTER(TOKEN_MANDATORY_LABEL)).contents
            raw = ctypes.cast(label.Label.Sid, ctypes.POINTER(ctypes.c_ubyte))
            sub_count = raw[1]                       # SID: 版本(1) 子授权数(1) 标识(6) 子授权(4n)
            offset = 8 + 4 * (sub_count - 1)
            rid = int.from_bytes(bytes(raw[offset:offset + 4]), "little")
            return {0x0000: "Untrusted", 0x1000: "Low", 0x2000: "Medium",
                    0x3000: "High", 0x4000: "System"}.get(rid, f"RID 0x{rid:x}")
        finally:
            kernel32.CloseHandle(token)
    except Exception:
        return "unknown"


def _active_outlook():
    """附加到已经在运行的 Outlook；没有则返回 None（不会启动它）。"""
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return None


def _mapi_namespace(app, *, log: Callable[..., None] = print):
    """取 MAPI namespace。Outlook 刚启动时 profile 可能还在加载，给它一点时间。"""
    last: Optional[Exception] = None
    for _ in range(15):
        try:
            return app.GetNamespace("MAPI")
        except Exception as ex:
            last = ex
            time.sleep(1)
    raise OutlookNotRunning(
        f"Outlook is up but its MAPI profile is not available yet ({last}). "
        "If this keeps happening, check Control Panel > Mail (Microsoft Outlook) > Profiles.")


def _minimize_outlook(app) -> None:
    """尽力把 Outlook 主窗口最小化到任务栏（后台运行）。

    刚启动时主窗口可能还没建好，所以轮询几次；拿不到就算了 ——
    最小化失败不影响读邮件，不该因此报错。
    """
    for _ in range(10):
        try:
            explorer = app.ActiveExplorer()
        except Exception:
            explorer = None
        if explorer is not None:
            try:
                explorer.WindowState = OL_WINDOW_MINIMIZED
            except Exception:
                pass
            return
        time.sleep(0.5)


def _notice(log: Callable[..., None], key: str, message: str) -> bool:
    """打一条"例行"连接日志；同一个 key 每次进程只打一次。

    为什么需要：监听每 30 秒跑一轮，每轮都连一次 Outlook。如果每轮都打
    "attaching / attached"，活动日志会被这两行刷屏 —— 界面日志区只保留 2000 行，
    真正有用的行（分类了多少封、补了标签、出了报告）会被挤掉，app.log 也每天涨约 260 KB。

    状态**变化**的消息（连不上、要启动 Outlook、启动好了）不走这里，照常每次都打 ——
    所以"能不能连上"这件事的变化不会被静音，被静音的只是"又一次例行附加"。
    """
    if key in _seen_notices:
        return False
    _seen_notices.add(key)
    log(message)
    return True


def connect(*, auto_launch: bool = True,
            log: Callable[..., None] = print) -> Tuple[Any, Any]:
    """连接 Outlook；没开的话**自动启动它**（并尽力最小化到后台）。返回 (app, MAPI namespace)。

    做法就是旧版 Donald 的 `connect_outlook()`：

        app = win32com.client.Dispatch("Outlook.Application")

    Outlook 是**单实例**应用：已经在跑时 Dispatch 返回的就是现有实例（不会新开一个），
    没在跑时它才会启动。所以**一个调用同时覆盖「附加」和「启动」两种情况**。

    这里有一条踩过两次的教训：**不要在 Dispatch 前面加任何"守卫"**。

      · 只用 GetActiveObject 不行 —— 它只查 ROT，Outlook 明明在跑也可能附加不上，
        于是程序以为自己连不上，反复重试、永远读不到邮件；
      · 加"启动失败后的冷却"也不行 —— 冷却会把**附加**也一起挡掉：
        用户手动打开 Outlook 之后，程序照样抛
        "could not be started 63s ago; not retrying for another 57s"。

    所以现在的顺序就是：看一眼有没有 OUTLOOK.EXE 进程（只为了日志措辞和"要不要最小化"），
    然后直接 Dispatch。真的连不上就抛 OutlookNotRunning，由调用方决定重试节奏。

    auto_launch=False 时退化为「只附加」，供只读自检使用 —— 自检不该顺手启动你的 Outlook。
    """
    if not auto_launch:
        app = _active_outlook()
        if app is None:
            raise OutlookNotRunning(OUTLOOK_MANUAL_HINT)
        return app, _mapi_namespace(app, log=log)

    # 只为两件事看进程：日志措辞，以及"只有我们自己启动的才去最小化"
    # （否则会把你正在用的 Outlook 窗口拽到任务栏）
    was_running = _outlook_process_running()
    if was_running:
        _notice(log, "attach", "[outlook] attaching to the running Outlook")
    else:
        # 没看到 OUTLOOK.EXE 进程。可能是真没开，也可能是枚举受限 —— 所以下面
        # **只对"要不要主动拉起"做退避，绝不阻止后面的 Dispatch 附加**。
        if _launch_allowed():
            _notice(log, "launch", "[outlook] Outlook is not running - starting it")
            _launch_outlook(log)
        else:
            _notice_backoff(log)

    app = _dispatch()

    if was_running or _outlook_process_running():
        _notice(log, "attached", "[outlook] attached")
    else:
        # 极少数情况：Dispatch 成功但进程还是看不到（枚举受限），照样能用，不最小化
        _notice(log, "attached-blind", "[outlook] connected")
    if not was_running:
        _minimize_outlook(app)          # 只最小化我们自己拉起来的那个
    return app, _mapi_namespace(app, log=log)


def outlook_exe_path() -> Optional[str]:
    """定位 OUTLOOK.EXE（先查注册表 App Paths，再退回常见安装路径）。

    用途：主动拉起 Outlook 时**正常启动这个可执行文件**，而不是靠 COM 激活 —— 见
    _launch_outlook() 里的说明。
    """
    if os.name != "nt":
        return None
    candidates: List[str] = []
    try:
        import winreg
        for sub in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe"):
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub) as key:
                    value = winreg.QueryValue(key, None)
                    if value:
                        candidates.append(value)
            except OSError:
                continue
    except Exception:
        pass
    candidates += [r"C:\Program Files\Microsoft Office\Root\Office16\OUTLOOK.EXE",
                   r"C:\Program Files (x86)\Microsoft Office\Root\Office16\OUTLOOK.EXE"]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


def _launch_allowed() -> bool:
    """现在允许再尝试「启动」Outlook 吗？（退避只约束启动，不约束附加）"""
    return (time.time() - _last_launch_attempt) >= LAUNCH_BACKOFF_SECONDS


def _notice_backoff(log: Callable[..., None]) -> None:
    """退避期内不打日志刷屏，但每次进入退避窗口提示一次。"""
    global _backoff_logged_at
    if time.time() - _backoff_logged_at < LAUNCH_BACKOFF_SECONDS:
        return
    _backoff_logged_at = time.time()
    log(f"[outlook] Outlook is not running and the last start attempt was less than "
        f"{LAUNCH_BACKOFF_SECONDS}s ago - not starting it again yet "
        f"(open Outlook yourself and this will reconnect by itself)")


def _launch_outlook(log: Callable[..., None]) -> None:
    """主动把 Outlook 拉起来。

    **优先「正常启动 OUTLOOK.EXE」，而不是靠 COM 激活。** 这是实测踩出来的：
    `Dispatch` 拉起来的往往是一个**自动化实例** —— 本轮任务结束、COM 引用释放后它就退出，
    于是下一轮又发现"没在运行"→ 再拉一次，形成**每 30 秒拉起一次 Outlook 的死循环**
    （实测一小时 100 次，每次都要重新加载一遍加载项，用户看到的是 Outlook 不断弹出，
    同时白白吃掉内存和 CPU）。

    直接启动 OUTLOOK.EXE 得到的是**完整的交互式 Outlook**，它会一直留着，下一轮直接附加即可。
    """
    global _last_launch_attempt
    _last_launch_attempt = time.time()

    exe = outlook_exe_path()
    if not exe:
        log("[outlook] could not locate OUTLOOK.EXE - falling back to COM activation")
        return
    try:
        subprocess.Popen([exe], close_fds=True)
    except Exception as ex:
        log(f"[outlook] could not start OUTLOOK.EXE ({type(ex).__name__}: {ex}) - "
            f"falling back to COM activation")
        return

    log(f"[outlook] started {exe} - waiting for it to come up")
    deadline = time.time() + OUTLOOK_START_TIMEOUT
    while time.time() < deadline:
        if _outlook_process_running():
            log("[outlook] OUTLOOK.EXE is up (it stays open now, no more repeated launches)")
            return
        time.sleep(1)
    log(f"[outlook] OUTLOOK.EXE did not appear within {OUTLOOK_START_TIMEOUT}s")


def _dispatch():
    """Dispatch("Outlook.Application")；失败时抛 OutlookNotRunning（带完整性提示）。"""
    try:
        return win32com.client.Dispatch("Outlook.Application")
    except Exception as ex:
        level = integrity_level()
        hint = ""
        if level in ("Low", "Untrusted"):
            hint = (f" This program is running at {level} integrity, and a {level}-integrity "
                    "process cannot start or reach Outlook through COM. Launch Email Manager "
                    "normally (double-click it in Explorer, or from the Start menu) instead of "
                    "from a sandboxed shell or agent.")
        raise OutlookNotRunning(
            f"Could not connect to Outlook: {type(ex).__name__}: {ex}.{hint}") from ex


def newest_message(ns):
    """收件箱里最新的那封邮件；空收件箱返回 None。

    只取第一封就返回，所以很快 —— watcher 每轮用它判断「有没有新邮件到了」。
    """
    items = ns.GetDefaultFolder(OL_FOLDER_INBOX).Items
    items.Sort("[ReceivedTime]", True)      # 新 → 旧
    for msg in items:
        return msg
    return None


# ── 收件箱约定 ────────────────────────────────────────────────────
def parse_sum_command(subject: str) -> Optional[str]:
    """主题是总结指令时返回规范化的级别串（如 'sum sa' → 'sa'），否则 None。"""
    m = SUM_COMMAND_RE.match((subject or "").strip())
    if not m:
        return None
    letters = (m.group(1) or "sabc").lower()
    return "".join(c for c in "sabc" if c in letters)


def is_report_subject(subject: str) -> bool:
    """主题是否是我们自己生成的总结报告（中英标记都认，见 REPORT_SUBJECT_MARKS）。"""
    text = subject or ""
    return any(mark in text for mark in REPORT_SUBJECT_MARKS)


def is_report(msg) -> bool:
    """是否是我们自己生成的总结邮件。"""
    return is_report_subject(msg.Subject or "")


def category_names(msg) -> List[str]:
    """邮件当前的分类名列表（已去空白、去空项）。"""
    return [c.strip() for c in (msg.Categories or "").split(",") if c.strip()]


def has_our_category(msg) -> bool:
    """是否已打过 S/A/B/C 或 Sum —— 都是本程序写的，默认不该再被覆盖。"""
    return any(c in LEVELS or c == SUM_NAME for c in category_names(msg))


def pending_sum_commands(ns, last_seen: Optional[str], *, days: Optional[int] = None,
                         max_items: int = 5) -> List[Any]:
    """「还没处理过」的 sum 指令邮件（新 → 旧，最多 max_items 封）。

    为什么需要（和 untagged_reports 同一类洞）：`plan_tick` 只看**最新那一封**。
    sum 指令到达后如果 30 秒内又来了别的邮件，指令就不再是"最新"，classify 路径
    又会把指令邮件当控制邮件跳过 —— 于是这条指令永远轮不到执行，
    基线一推进就彻底丢了（用户发了指令却什么都没发生）。

    判据：主题是合法指令 + **还没打 Sum 标签**（处理时会打上，所以它同时是"已处理"的标志）。
    只从最新往回走到 last_seen（"上次已经看过的那封"）就停，所以不遍历整个收件箱。
    """
    items = ns.GetDefaultFolder(OL_FOLDER_INBOX).Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=days) if days else None
    out: List[Any] = []
    for msg in items:
        if last_seen and msg.EntryID == last_seen:
            break                            # 到上次看过的位置了，后面都是旧的
        received = _naive(msg.ReceivedTime)
        if cutoff is not None and received is not None and received < cutoff:
            break
        if parse_sum_command(msg.Subject) is None:
            continue
        if SUM_NAME in category_names(msg):  # 打过 Sum = 已经处理过
            continue
        out.append(msg)
        if len(out) >= max_items:
            break
    return out


# 按主题标记直接筛报告（DASL 语法；Jet 的 "[Subject] Like ..." Outlook 不接受）
def report_filter(mark: str) -> str:
    """按一个主题标记生成筛选条件。"""
    return '@SQL="urn:schemas:httpmail:subject" LIKE \'%' + mark + '%\''


REPORT_FILTER = report_filter(REPORT_SUBJECT_MARK)   # 中文标记的条件（供参考/自测）


def untagged_reports(ns, *, days: Optional[int] = 30) -> List[Any]:
    """收件箱里「是总结报告、但还没打 Sum 标签」的邮件（新 → 旧）。

    为什么需要单独找一遍：`plan_tick` 只看**最新那一封**，而 classify 路径又把报告
    当控制邮件跳过 —— 于是只要报告到达之后又有别的邮件更新，它就永远轮不到打标签
    （实际发生过：10 封报告里有 1 封缺 Sum）。

    中英两个标记各查一次再合并（而不是拼一个 OR 条件）：两次查询各 0.03 秒，
    比赌 DASL 的 OR 语法可靠。合并时按 EntryID 去重。
    实测 829 封的收件箱两次查询合计约 0.05 秒，不需要遍历整个收件箱。
    """
    items = ns.GetDefaultFolder(OL_FOLDER_INBOX).Items

    candidates: List[Any] = []
    seen: set = set()
    for mark in REPORT_SUBJECT_MARKS:
        try:
            found = items.Restrict(report_filter(mark))
        except Exception:
            continue                      # 查询语法不被支持时跳过这一种标记
        for msg in found:
            try:
                key = msg.EntryID
            except Exception:
                key = None
            if key is not None:
                if key in seen:
                    continue
                seen.add(key)
            candidates.append(msg)

    cutoff = datetime.now() - timedelta(days=days) if days else None
    out: List[Any] = []
    for msg in candidates:
        received = _naive(msg.ReceivedTime)
        if cutoff is not None and received is not None and received < cutoff:
            continue
        if not is_report(msg):              # 主题里带标记但不是我们的报告格式
            continue
        if SUM_NAME in category_names(msg):
            continue
        out.append(msg)
    return out


# ── 分类维护 ──────────────────────────────────────────────────────
def ensure_categories(ns) -> List[str]:
    """确保 S/A/B/C 四个分类存在且颜色正确，并清理旧命名。

    返回本次改动说明（空列表 = 一切本来就正确），由调用方决定要不要打印。
    """
    cats = ns.Categories
    changes: List[str] = []

    for name in LEGACY_CATEGORY_NAMES:
        try:
            cats.Item(name).Delete()
            changes.append(f"removed legacy category {name}")
        except Exception:
            pass  # 本来就没有，属正常情况

    for level in LEVELS:
        color = CATEGORY_COLORS[level]
        try:
            cat = cats.Item(level)
        except Exception:
            cat = None
        if cat is None:
            try:
                cats.Add(level, color)
                changes.append(f"created category {level} (color={color})")
            except Exception as ex:
                changes.append(f"could not create category {level}: {ex}")
        elif cat.Color != color:
            try:
                cat.Color = color
                changes.append(f"fixed category color {level} -> {color}")
            except Exception as ex:
                changes.append(f"could not fix category color {level}: {ex}")
    return changes


def ensure_sum_category(ns) -> bool:
    """确保黄色 Sum 分类存在且颜色正确。返回是否发生了改动。"""
    cats = ns.Categories
    try:
        cat = cats.Item(SUM_NAME)
    except Exception:
        cat = None
    if cat is None:
        try:
            cats.Add(SUM_NAME, SUM_COLOR)
            return True
        except Exception:
            return False
    if cat.Color != SUM_COLOR:
        try:
            cat.Color = SUM_COLOR
            return True
        except Exception:
            pass
    return False


def existing_categories(ns) -> List[str]:
    """列出 Outlook 里现有的分类名（只读，供自检使用）。"""
    cats = ns.Categories
    out = []
    for i in range(1, cats.Count + 1):
        try:
            out.append(cats.Item(i).Name)
        except Exception:
            pass
    return out


def our_category_colors(ns) -> Dict[str, int]:
    """返回我们已经存在的分类（S/A/B/C/Sum）的实际颜色值。

    只读，供自检用：用来发现「分类被 Outlook 自动加进去、但颜色是空的」这种情况
    —— 全新用户如果先被监听器打了标签，就会出现这种没颜色的分类。
    下一次 ensure_categories() 会把颜色修正回来。
    """
    cats = ns.Categories
    out: Dict[str, int] = {}
    for name in list(LEVELS) + [SUM_NAME]:
        try:
            out[name] = int(cats.Item(name).Color)
        except Exception:
            continue
    return out


def expected_category_colors() -> Dict[str, int]:
    """S/A/B/C/Sum 各自应有的颜色值。"""
    colors = dict(CATEGORY_COLORS)
    colors[SUM_NAME] = SUM_COLOR
    return colors


# ── 读信 ──────────────────────────────────────────────────────────
def _naive(dt):
    """COM 返回的 datetime 可能带 tzinfo，统一去掉，方便与本地时间比较。"""
    if dt is not None and getattr(dt, "tzinfo", None) is not None:
        return dt.replace(tzinfo=None)
    return dt


def message_info(msg) -> Dict[str, Any]:
    """把一个 MailItem 转成内部统一使用的邮件字典。"""
    return {
        "entryid": msg.EntryID,
        "msg": msg,
        "subject": msg.Subject or "(no subject)",
        "sender_name": msg.SenderName or "",
        "sender": msg.SenderEmailAddress or "",
        "body": (msg.Body or "").strip(),
        "received": _naive(msg.ReceivedTime),
    }


def collect_new_mail(ns, last_seen: Optional[str], *, days: Optional[int] = None,
                     max_items: int = 100) -> List[Dict[str, Any]]:
    """收件箱里比 last_seen 更新的邮件（新 → 旧）。

    走到 last_seen 就停，所以只会处理「真正新到的那些」，不会去啃整个历史收件箱。
    旧实现用 read_inbox(days=None) 会遍历全部邮件并逐封读正文，大邮箱下要几分钟，
    表现为「检测到新邮件却什么都没发生，按钮还一直是灰的」。

    跳过：我们自己的报告、sum 指令邮件、已经打过 S/A/B/C/Sum 的邮件。
    days / max_items 是兜底：万一 last_seen 那封已被删除或移走，也能收住范围。
    """
    items = ns.GetDefaultFolder(OL_FOLDER_INBOX).Items
    items.Sort("[ReceivedTime]", True)
    cutoff = datetime.now() - timedelta(days=days) if days is not None else None

    out: List[Dict[str, Any]] = []
    for msg in items:
        try:
            entry_id = msg.EntryID
        except Exception:
            continue
        if last_seen and entry_id == last_seen:
            break

        received = _naive(msg.ReceivedTime)
        if cutoff is not None and received is not None and received < cutoff:
            break
        if is_report(msg) or parse_sum_command(msg.Subject) is not None:
            continue
        if has_our_category(msg):
            continue

        out.append(message_info(msg))
        if len(out) >= max_items:
            break
    return out


def read_inbox(ns, *, limit: Optional[int] = None, days: Optional[int] = None,
               skip_classified: bool = True) -> List[Dict[str, Any]]:
    """读收件箱待分级邮件（新 → 旧）。

    limit            最多读多少封（None = 不限）
    days             只读最近 N 天（利用时间降序，遇到更旧的就直接停）
    skip_classified  跳过已打 S/A/B/C/Sum 的邮件（要重新分级时传 False）

    始终跳过：我们自己生成的总结邮件、sum 指令邮件。
    """
    items = ns.GetDefaultFolder(OL_FOLDER_INBOX).Items
    items.Sort("[ReceivedTime]", True)  # 大集合下标索引不可靠，一律用迭代

    cutoff = datetime.now() - timedelta(days=days) if days is not None else None
    emails: List[Dict[str, Any]] = []

    for msg in items:
        received = _naive(msg.ReceivedTime)
        if cutoff is not None and received is not None and received < cutoff:
            break  # 已按时间降序，后面只会更旧
        if is_report(msg) or parse_sum_command(msg.Subject) is not None:
            continue
        if skip_classified and has_our_category(msg):
            continue

        emails.append(message_info(msg))
        if limit is not None and len(emails) >= limit:
            break
    return emails


def read_recent(ns, *, days: int, letters: str = "sabc") -> List[Dict[str, Any]]:
    """读最近 N 天内、已打 S/A/B/C 分类且级别属于 letters 的邮件。

    返回按 S→A→B→C 排序（组内保持新 → 旧）。
    """
    items = ns.GetDefaultFolder(OL_FOLDER_INBOX).Items
    items.Sort("[ReceivedTime]", True)

    cutoff = datetime.now() - timedelta(days=days)
    out: List[Dict[str, Any]] = []

    for msg in items:
        received = _naive(msg.ReceivedTime)
        if received is not None and received < cutoff:
            break
        level = next((c for c in category_names(msg) if c in LEVELS), None)
        if level is None or level.lower() not in letters:
            continue
        out.append({
            "entryid": msg.EntryID,
            "level": level,
            "subject": msg.Subject or "(no subject)",
            "sender": msg.SenderName or msg.SenderEmailAddress or "",
            "body": (msg.Body or "").strip(),
            "received": received,
        })

    out.sort(key=lambda e: LEVELS.index(e["level"]))  # 稳定排序 → 组内保持时间降序
    return out


# ── 写回 / 发信 ───────────────────────────────────────────────────
def write_categories(msg, names: Sequence[str], flag: Optional[int] = None) -> None:
    """把分类写回邮件；flag 不为 None 时同时设置后续标志（红旗）。

    Save() 必须调用，否则分类不会落盘。
    """
    msg.Categories = ", ".join(names)
    if flag is not None:
        msg.FlagStatus = flag
    msg.Save()


def send_html(app, to: str, subject: str, html: str, categories: str = "") -> None:
    """发送 HTML 邮件（总结报告用）。"""
    mail = app.CreateItem(OL_MAIL_ITEM)
    mail.To = to
    mail.Subject = subject
    if categories:
        mail.Categories = categories
    mail.BodyFormat = OL_FORMAT_HTML
    mail.HTMLBody = html
    mail.Send()
