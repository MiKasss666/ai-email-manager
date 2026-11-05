"""core/mail.py — 邮件后端门面（Microsoft Graph，唯一实现）。

历史版本这里在 Graph 与 Outlook COM 两个后端之间切换；v3.0.0 起只保留 Graph。
本模块保留原来的端口名，上层（classifier / summarizer / gui / run.py）继续
`from core import mail as outlook` 即可，不必关心实现细节。

纯逻辑与约定（分类名、报告/指令识别、邮件字典形状）在 core/mailmeta.py。
"""
from __future__ import annotations

from typing import Any

from . import graph as _graph
from .graph import (  # noqa: F401
    collect_new_mail,
    connect,
    logged_in,
    newest_message,
    pending_sum_commands,
    read_inbox,
    read_recent,
    send_html,
    untagged_reports,
    write_categories,
)
from .mailmeta import (  # noqa: F401
    FLAG_MARKED,
    FLAG_NONE,
    LEGACY_CATEGORY_NAMES,
    OUR_CATEGORY_NAMES,
    REPORT_SUBJECT_MARK,
    REPORT_SUBJECT_MARK_EN,
    REPORT_SUBJECT_MARKS,
    SUM_COMMAND_RE,
    SUM_NAME,
    MailboxUnavailable,
    category_names,
    has_our_category,
    is_report,
    is_report_subject,
    merge_categories,
    message_info,
    note_once,
    parse_sum_command,
)

BACKEND = "graph"


def describe() -> str:
    """给日志/自检用的一句话说明当前后端（日志统一用英文）。"""
    try:
        from . import config as config_mod
        gcfg = config_mod.load().graph
        who = f"client_id={gcfg.client_id[:8]}..." if gcfg.client_id else "Microsoft public client"
        return f"graph (Microsoft Graph, tenant={gcfg.tenant}, {who})"
    except Exception:
        return "graph (Microsoft Graph)"


def backend_module() -> Any:
    """返回真正的实现模块（自检/测试用）。"""
    return _graph


__all__ = [
    "BACKEND", "describe", "backend_module",
    "connect", "newest_message", "read_inbox", "collect_new_mail", "read_recent",
    "pending_sum_commands", "untagged_reports", "write_categories", "send_html", "logged_in",
    "SUM_NAME", "OUR_CATEGORY_NAMES", "LEGACY_CATEGORY_NAMES", "REPORT_SUBJECT_MARK",
    "REPORT_SUBJECT_MARK_EN", "REPORT_SUBJECT_MARKS", "SUM_COMMAND_RE",
    "FLAG_NONE", "FLAG_MARKED", "MailboxUnavailable",
    "category_names", "has_our_category", "is_report", "is_report_subject",
    "merge_categories", "message_info", "note_once", "parse_sum_command",
]
