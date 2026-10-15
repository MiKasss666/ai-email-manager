"""guitest.py — 界面冒烟测试：不弹对话框、不联网、不碰 Outlook、不改 config.json。

用法（在 Email_manager_beta 目录下）：
    py -3 guitest.py

Tkinter 只能在有桌面会话的环境里创建窗口，所以这里把主窗口 withdraw 隐藏起来跑，
正常情况下你看不到任何窗口（构造 Toplevel 的瞬间可能极短暂地闪一下，属 Tk 的限制）。

覆盖界面里最容易出错的三块：
  1. 主窗口能否装配出来
  2. 日志「后台线程 → 队列 → widget」这条路通不通
  3. 后台任务器：线程执行、回调必须回主线程、busy 语义、异常兜底
  4. 两个设置窗口的「表单 ↔ 配置」往返，以及「没点保存就不落盘」

注意：本文件会读到 gui 内部的私有字段（_vars / _list / _fields），
这是冒烟测试的取舍——只验证装配与数据通路，不重复验证业务逻辑（那在 selftest.py）。
"""
from __future__ import annotations

import threading
import time
import traceback
import tkinter as tk

from core import config


def pump(root: tk.Tk, seconds: float = 0.4) -> None:
    """跑一段事件循环，让 after 回调有机会执行（不用 mainloop，避免卡住）。"""
    end = time.time() + seconds
    while time.time() < end:
        try:
            root.update()
        except tk.TclError:
            return
        time.sleep(0.02)


def widget_texts(widget: tk.Misc) -> list:
    """递归收集所有子控件的 text 文案（用来验证界面是否英文化）。"""
    texts = []
    try:
        value = widget.cget("text")
        if value:
            texts.append(str(value))
    except Exception:
        pass
    for child in widget.winfo_children():
        texts.extend(widget_texts(child))
    return texts


def has_chinese(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def _is_valid_ico(path) -> bool:
    """ICO 文件头是 00 00 01 00，紧跟的第 5-6 字节是内含图像数量。

    Tk 在 Windows 上不支持查询窗口当前图标（wm iconbitmap 查询返回空串），
    所以这里退一步验证「文件确实是个能用的多分辨率图标」——
    这正是分发时忘了带 .ico 会坏掉的地方。
    """
    if not path:
        return False
    try:
        with open(path, "rb") as handle:
            header = handle.read(6)
    except OSError:
        return False
    if len(header) < 6 or header[:4] != b"\x00\x00\x01\x00":
        return False
    return int.from_bytes(header[4:6], "little") > 0


def main() -> int:
    ok = 0

    def check(name: str, cond) -> None:
        nonlocal ok
        assert cond, f"FAIL: {name}"
        ok += 1
        print(f"  ok  {name}")

    # ── 1. 主窗口 ─────────────────────────────────────────────
    print("[1] 主窗口装配")
    root = tk.Tk()
    root.withdraw()

    from gui.main_window import MainWindow
    app = MainWindow(root)
    auto_started = app.watcher.running     # 默认不该自动开监听
    root.update_idletasks()

    check("主窗口装配成功", app.cfg is not None)
    check("日志区创建成功", app.log is not None)
    check("任务器初始空闲", not app.runner.busy)
    check("配置对象与磁盘一致", app.cfg.email == config.load().email)

    print("[1b] 界面文案是英文")
    check("窗口标题", root.title() == "AI Email Manager")
    check("按钮 AI setting", app._btn_ai["text"] == "AI setting")
    check("按钮 Categorize / Sum",
          app._btn_categorize["text"] == "Categorize" and app._btn_sum["text"] == "Sum")
    check("按钮 User info", app._btn_user["text"] == "User info")
    check("日志区标题", "Activity log" in widget_texts(app.log))
    check("没有残留的中文按钮文案",
          not any(has_chinese(t) for t in widget_texts(app.log)))

    print("[1c] 自动监听：配置不全会拒绝启动")
    check("配置不全时不自动开监听", auto_started is False)
    check("界面开关默认未勾选", bool(app._watch_var.get()) is False)
    check("未监听时不排轮询", app.watcher._after_id is None)

    import gui.main_window as mw_mod
    _orig_cfg = app.cfg                      # 后面几节还要用原配置，测完还原
    _orig_warn = mw_mod.messagebox.showwarning
    mw_mod.messagebox.showwarning = lambda *a, **k: None
    try:
        app._watch_var.set(True)             # 模拟用户硬要勾上
        app._on_watch_toggle()
    finally:
        mw_mod.messagebox.showwarning = _orig_warn
    check("配置不全时手动勾选被拒绝", app.watcher.running is False)
    check("被拒绝后勾选自动回弹", bool(app._watch_var.get()) is False)

    # 补全配置后（有 email + 能用的 AI 后端）应当自动开
    app.cfg = config.from_dict({"email": "me@example.edu",
                                "ai": {"active": "c", "profiles": {"c": {"type": "cli"}}}})
    check("补齐后的配置无阻塞问题", app.cfg.problems() == [])
    app._maybe_start_watching("test")
    check("配置齐全后自动开始监听", app.watcher.running is True)
    check("界面开关同步为勾选", bool(app._watch_var.get()) is True)
    app.watcher.stop()
    check("可以手动关掉", app.watcher.running is False)
    check("监听停止后不排下一轮", app.watcher._after_id is None)
    app.cfg = _orig_cfg                      # 还原，避免影响后面几节

    print("[1d] 窗口图标")
    from gui import theme as theme_mod
    _icon = theme_mod.icon_path()
    check("图标文件随包存在", _icon is not None and _icon.endswith("EmailManager.ico"))
    check("图标是合法 ICO 且含多分辨率", _is_valid_ico(_icon))
    check("图标来源可用（源码/onedir 用 .ico，单文件 exe 会用 exe 自身）",
          theme_mod.icon_source() is not None)
    theme_mod.apply_icon(root)          # 可安全重复调用
    check("apply_icon 不抛异常", True)

    # ── 2. 日志通路 ───────────────────────────────────────────
    print("[2] 日志通路（后台线程 → 队列 → widget）")
    app.log.clear()

    def background_write() -> None:
        for i in range(3):
            app.log.write(f"后台第 {i} 行")

    thread = threading.Thread(target=background_write, daemon=True)
    thread.start()
    thread.join()
    pump(root, 0.6)

    text = app.log.contents()
    check("后台线程写的日志出现在 widget 里", "后台第 2 行" in text)
    check("print 风格多参数被拼接", " " in text)

    print("[2b] 日志落盘 + 复制按钮")
    _logfile = app.log.log_file()
    check("启用了日志文件", bool(_logfile) and _logfile.endswith("app.log"))
    pump(root, 0.5)                       # 让 drain 把队列里的内容写进文件
    _on_disk = ""
    try:
        with open(_logfile, "r", encoding="utf-8") as handle:
            _on_disk = handle.read()
    except OSError:
        pass
    check("日志文件真的写进去了", "后台第 0 行" in _on_disk)
    check("文件里每行带时间戳", "[0" in _on_disk or "[1" in _on_disk or "[2" in _on_disk)
    app.log.copy_to_clipboard()           # 不应抛异常
    check("复制到剪贴板不报错", True)

    # ── 3. 后台任务器 ─────────────────────────────────────────
    print("[3] 后台任务器")
    got = {}

    started = app.runner.run(lambda: 6 * 7,
                             on_done=lambda r: got.setdefault("done", r),
                             on_error=lambda e: got.setdefault("error", e))
    check("任务已启动", started is True)
    check("启动后立即处于忙碌", app.runner.busy is True)
    check("忙碌时拒绝第二个任务", app.runner.run(lambda: None) is False)
    pump(root, 1.0)
    check("on_done 收到返回值", got.get("done") == 42)
    check("任务结束后回到空闲", app.runner.busy is False)

    def boom():
        raise ValueError("故意炸一下")

    got.pop("error", None)
    app.runner.run(boom, on_error=lambda e: got.setdefault("error", e))
    pump(root, 1.0)
    check("异常交给 on_error", isinstance(got.get("error"), ValueError))
    check("异常后仍回到空闲（不会永久卡死）", app.runner.busy is False)

    threads = {}
    app.runner.run(lambda: "x", on_done=lambda _r: threads.setdefault("name", threading.current_thread().name))
    pump(root, 1.0)
    check("回调在主线程执行（能安全操作界面）", threads.get("name") == "MainThread")

    print("[3b] 卡死逃生口（模拟 Outlook 不响应）")
    check("状态栏看门狗已启动", app._watchdog_id is not None)

    hang = threading.Event()
    app.runner.run(lambda: hang.wait(30))       # 一个永远不结束的任务
    pump(root, 0.6)
    check("卡住时仍是 busy", app.runner.busy is True)
    check("elapsed 能反映已等待时长", app.runner.elapsed > 0.4)

    import gui.main_window as main_window_mod
    original_askyesno = main_window_mod.messagebox.askyesno
    main_window_mod.messagebox.askyesno = lambda *a, **k: True
    app.watcher.start()                          # 监听开着，验证断链会被接回
    try:
        app._force_reset()
    finally:
        main_window_mod.messagebox.askyesno = original_askyesno
    check("force reset 后界面解锁", app.runner.busy is False)
    check("force reset 后 elapsed 归零", app.runner.elapsed == 0.0)
    check("监听轮询链被接回（否则会静默死掉）", app.watcher._after_id is not None)
    app.watcher.stop()
    hang.set()                                   # 放掉卡住的线程
    pump(root, 0.6)                              # 它的结果会被代次检查丢掉（见下）

    # 被放弃的任务的结果必须被丢弃：既不能打到新任务的回调上，也不能搅乱界面状态。
    # 旧实现在这里会把僵尸的结果交给当前任务的 on_done，并把 busy 清零 ——
    # 于是用户能再点出第三个任务，三个线程同时给同一个邮箱写分类。
    print("[3b-2] 被放弃的任务的结果不会回灌")
    stale_calls = []
    zombie = threading.Event()
    app.runner.run(lambda: zombie.wait(30))
    pump(root, 0.3)
    app.runner.force_reset()                     # 放弃它（线程还在后台挂着）
    check("放弃后界面解锁", app.runner.busy is False)
    app.runner.run(lambda: "fresh", on_done=lambda r: stale_calls.append(r))
    pump(root, 0.6)
    check("新任务正常收到自己的结果", stale_calls == ["fresh"])
    zombie.set()                                 # 僵尸现在才结束 —— 结果必须被丢掉
    pump(root, 0.8)
    check("僵尸的结果没有回灌到新任务的回调（也不会清掉 busy）",
          stale_calls == ["fresh"] and app.runner.busy is False)
    check("主线程读不到任务令牌（令牌是各工作线程自己的）", app.runner.token is None)

    # 卡住提示必须按「现在在等什么」说话：卡在 AI 和卡在 Outlook 的处理方式完全不同。
    # 实测踩过：等 DeepSeek 回话等了 8 分钟，界面却一直说 "Outlook has not answered"。
    print("[3b-3] 卡住提示按阶段说话（不再一律怪 Outlook）")
    import gui.main_window as _mw
    _saved_hint = _mw.STUCK_HINT_SECONDS
    _mw.STUCK_HINT_SECONDS = 0                    # 让提示立刻触发
    app.log.clear()
    hold = threading.Event()

    def _stall_in_ai():
        app.runner.set_phase("waiting for the AI (deepseek)")
        hold.wait(15)

    try:
        app.runner.run(_stall_in_ai)
        pump(root, 1.6)                           # 看门狗每秒一次，要跨过它的 tick
        _text = app.log.contents()
        check("状态栏显示当前阶段", app.runner.phase.startswith("waiting for the AI"))
        check("提示指向 AI 服务商，而不是让用户去折腾 Outlook",
              "AI provider has not answered" in _text and "Outlook has not answered" not in _text)
    finally:
        _mw.STUCK_HINT_SECONDS = _saved_hint
        app.runner.force_reset()
        hold.set()
        pump(root, 0.5)

    print("[3c] 按钮状态每秒自愈")
    # 直接改 runner 的内部状态，模拟「某处漏了一次解除禁用」这种状态漂移
    app.runner._busy = True
    app.runner._started_at = time.time()
    pump(root, 1.4)                              # 让真实的看门狗跑一轮
    check("busy 时 Sum 被禁用", str(app._btn_sum.cget("state")) == "disabled")
    check("busy 时 Categorize 被禁用", str(app._btn_categorize.cget("state")) == "disabled")

    app.runner._busy = False                     # 任务其实早就结束了
    app.runner._started_at = None
    pump(root, 1.4)
    check("看门狗把 Sum 恢复可点", str(app._btn_sum.cget("state")) == "normal")
    check("看门狗把 Categorize 恢复可点", str(app._btn_categorize.cget("state")) == "normal")
    check("看门狗把 AI setting 恢复可点", str(app._btn_ai.cget("state")) == "normal")

    print("[3d] Outlook 用不了时不当作失败")
    from core.outlook import OutlookNotRunning
    app.watcher.start(delay=60)          # 推迟首轮，避免测试真去连 Outlook
    pump(root, 0.3)
    app.watcher._errors = 0
    app.log.clear()
    app.watcher._on_error(OutlookNotRunning("outlook closed"))
    check("Outlook 用不了不累计失败次数", app.watcher._errors == 0)
    check("仍然保持监听（不会自动暂停）", app.watcher.running is True)
    check("轮询链继续排下去", app.watcher._after_id is not None)
    pump(root, 0.4)
    check("日志里给出了原因", "Outlook is not available yet" in app.log.contents())
    check("并说明会继续重试", "still watching; retrying every" in app.log.contents())
    app.log.clear()
    app.watcher._on_error(OutlookNotRunning("outlook closed"))
    pump(root, 0.4)
    check("同一种情况只提示一次（不刷屏）", "Outlook is not running" not in app.log.contents())
    app.watcher._on_error(ValueError("a real error"))
    check("真正的错误照常累计", app.watcher._errors == 1)
    app.watcher.stop()
    app.watcher._errors = 0

    # ── 4. User info 表单 ─────────────────────────────────────
    print("[4] User info 表单")
    from gui.user_info_dialog import UserInfoDialog
    dialog = UserInfoDialog(root, app.cfg)
    dialog.withdraw()
    pump(root, 0.3)

    check("email 回填正确", dialog._vars["email"].get() == app.cfg.email)
    check("courses 以逗号回显", dialog._vars["courses"].get() == ", ".join(app.cfg.courses))
    check("text 组里的 language 回填正确", dialog._vars["language"].get() == app.cfg.language)
    check("summary days 回填正确", dialog._vars["summary_days"].get() == str(app.cfg.summary_days))
    dialog.destroy()

    # ── 5. AI 设置面板 ────────────────────────────────────────
    print("[5] AI 设置面板")
    from gui.ai_settings_dialog import AiSettingsDialog
    ai_dialog = AiSettingsDialog(root, app.cfg)
    ai_dialog.withdraw()
    pump(root, 0.3)

    names = list(app.cfg.ai_profiles)
    check("列表列出全部后端", ai_dialog._list.size() == len(names))
    check("默认选中第一个后端", ai_dialog._current == names[0])
    check("表单回填 type", ai_dialog._vars["type"].get() == app.cfg.ai_profiles[names[0]].type)

    if app.cfg.ai_profiles[names[0]].type == "cli":
        check("cli 后端把 base_url 置灰", str(ai_dialog._fields["base_url"]["state"]) == "disabled")
        check("cli 后端启用 cli_command", str(ai_dialog._fields["cli_command"]["state"]) == "normal")

    ai_dialog._vars["model"].set("自测模型")
    ai_dialog._collect_current()
    check("表单改动被收进内存副本", ai_dialog._profiles[names[0]].model == "自测模型")
    check("没点保存就不落盘（取消是干净的）",
          config.load().ai_profiles[names[0]].model != "自测模型")

    if len(names) > 1:
        ai_dialog._list.selection_clear(0, "end")
        ai_dialog._list.selection_set(1)
        ai_dialog._on_select()
        pump(root, 0.2)
        check("切换后端后表单跟着变", ai_dialog._current == names[1])
        check("切换前把改动收好了（不丢数据）",
              ai_dialog._profiles[names[0]].model == "自测模型")

    ai_dialog._close()
    pump(root, 0.2)

    # ── 收尾 ──────────────────────────────────────────────────
    app.runner.stop()
    root.destroy()
    print(f"\n全部通过：{ok} 项")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        print("\n界面冒烟测试失败（见上面的 traceback）")
        raise SystemExit(1)
