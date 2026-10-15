"""core/mail.py — 邮件后端选择（Graph 或 Outlook COM）。

上层（classifier / watcher / summarizer / 界面 / run.py）只 import 这一个模块，
用的名字和以前 `core.outlook` 完全一样，所以**换后端不需要改上层代码**：

    from core import mail as outlook
    app, ns = outlook.connect(log=...)
    mails = outlook.read_inbox(ns, days=7, ...)
    outlook.write_categories(msg, ["S"], flag)

选哪个由 `config.json` 的 `mail_source` 决定（默认 `graph`）：

    "mail_source": "graph"      → core/graph.py   （Microsoft Graph，不需要 Outlook）
    "mail_source": "outlook"    → core/outlook.py （Outlook 经典版 + COM，旧行为）

临时覆盖：环境变量 `EM_MAIL_SOURCE=outlook`（自检/排障用，不必改配置）。
"""
from __future__ import annotations

import os
from typing import Any, List

# 端口：函数
_PORT_FUNCS = (
    "connect", "newest_message", "message_info", "read_inbox", "collect_new_mail",
    "read_recent", "pending_sum_commands", "untagged_reports", "write_categories",
    "send_html", "ensure_categories", "ensure_sum_category", "existing_categories",
    "our_category_colors", "expected_category_colors", "category_names", "has_our_category",
    "parse_sum_command", "is_report", "is_report_subject", "sender_address",
    "integrity_level", "parent_process", "outlook_exe_path",
)
# 端口：常量 / 约定（两个后端必须一致，缺的从 outlook 退回）
_PORT_CONSTS = (
    "SUM_NAME", "SUM_COLOR", "CATEGORY_COLORS", "LEVELS", "OUR_CATEGORY_NAMES",
    "LEGACY_CATEGORY_NAMES", "REPORT_SUBJECT_MARK", "REPORT_SUBJECT_MARK_EN",
    "REPORT_SUBJECT_MARKS", "SUM_COMMAND_RE", "OL_FOLDER_INBOX", "OL_FLAG_NONE",
    "OL_FLAG_MARKED", "OL_MAIL_ITEM", "OL_FORMAT_HTML", "OutlookNotRunning",
)


def _choice() -> str:
    """决定用哪个后端：环境变量 > config.json > 默认 graph。"""
    from . import config as config_mod
    name = (os.environ.get("EM_MAIL_SOURCE") or "").strip().lower()
    if name in config_mod.MAIL_SOURCES:
        return name
    try:
        name = (config_mod.load().mail_source or "").strip().lower()
    except Exception:
        name = ""
    return name if name in config_mod.MAIL_SOURCES else config_mod.DEFAULT_MAIL_SOURCE


BACKEND = _choice()

if BACKEND == "outlook":
    from . import outlook as _impl
else:
    from . import graph as _impl

# 常量/纯逻辑的共同出处：无论选哪个后端，都从 outlook 取一份兜底（它不 import COM 也能用）
from . import outlook as _base                                  # noqa: E402

for _name in _PORT_FUNCS + _PORT_CONSTS:
    _value = getattr(_impl, _name, None)
    if _value is None:
        _value = getattr(_base, _name, None)
    if _value is not None:
        globals()[_name] = _value
del _name, _value


def describe() -> str:
    """给日志/自检用的一句话说明当前后端（日志统一用英文）。"""
    if BACKEND == "outlook":
        return "outlook (Outlook Classic + COM)"
    try:
        from . import config as config_mod
        gcfg = config_mod.load().graph
        who = f"client_id={gcfg.client_id[:8]}..." if gcfg.client_id else "Microsoft public client"
        return f"graph (Microsoft Graph, tenant={gcfg.tenant}, {who})"
    except Exception:
        return "graph (Microsoft Graph)"


def logged_in() -> bool:
    """当前后端是否已经"登录"（Graph 查凭据缓存；COM 恒为 True，因为它不需要登录）。"""
    if BACKEND != "graph":
        return True
    try:
        from . import graph as graph_mod
        return graph_mod.logged_in()
    except Exception:
        return False


def backend_module() -> Any:
    """返回真正的实现模块（自检/测试用）。"""
    return _impl


__all__: List[str] = ["BACKEND", "describe", "logged_in", "backend_module",
                      *_PORT_FUNCS, *_PORT_CONSTS]
