"""core/mailmeta.py — 邮件领域的纯逻辑与约定。

本模块**不联网、不碰 UI、不依赖第三方库**，是下列约定的唯一真源：

  · 邮件字典的形状（message_info）
  · 本程序写进邮件的分类名（OUR_CATEGORY_NAMES / SUM_NAME / 旧版命名）
  · 写回分类的合并规则（merge_categories：只替换自己的标签）
  · 报告邮件与 sum 指令的识别
  · 「邮箱暂时不可用」这一异常类型（上层据此退避，而不是计失败）

graph.py（后端实现）、classifier / summarizer（业务）、gui 都从这里取，
避免同一套约定在两个后端里各写一份、慢慢漂移。
"""
from __future__ import annotations

import re
from typing import Any, Callable, Dict, List, Optional, Sequence

from .config import LEVELS

# 本程序自己的分类名。写回时**只替换这些**：用户自己定义的、以及服务器规则打的
# 分类一律原样保留（见 merge_categories）。
SUM_NAME = "Sum"
LEGACY_CATEGORY_NAMES = ("S 紧急", "A 重要", "B 普通", "C 可忽略")
OUR_CATEGORY_NAMES = frozenset(list(LEVELS) + [SUM_NAME] + list(LEGACY_CATEGORY_NAMES))

# 我们自己生成的总结邮件，主题里一定含这个词 —— 扫收件箱时要跳过，别把报告再分一次级。
REPORT_SUBJECT_MARK = "邮件总结"            # 中文标记（language = zh）
REPORT_SUBJECT_MARK_EN = "Email summary"    # 英文标记（language = en）
# 两个都认：主题跟随 language，但用户随时可能切换语言，历史报告也必须继续被认出来。
REPORT_SUBJECT_MARKS = (REPORT_SUBJECT_MARK, REPORT_SUBJECT_MARK_EN)

# 用户发给自己的总结指令：sum / sum s / sum sa / sum sabc
SUM_COMMAND_RE = re.compile(r"^sum(?:\s+([sabc]+))?$", re.IGNORECASE)

# write_categories 的 flag 语义（Graph 的 flag.flagStatus）
FLAG_NONE = 0
FLAG_MARKED = 2


class MailboxUnavailable(RuntimeError):
    """邮箱当前不可用（没网、被限流、服务端 5xx、登录态暂时拿不到）。

    单独做一种异常类型，是为了让上层能区分「邮箱暂时用不了」和「真的出错了」：
    前者是正常情况，不该被当成失败去累计、更不该触发自动暂停。
    """


# ── 日志节流 ─────────────────────────────────────────────────────
_seen_notices: set = set()


def note_once(log: Callable[..., None], key: str, message: str) -> bool:
    """打一条「例行」日志；同一个 key 每次进程只打一次。

    监听每 30 秒跑一轮、每轮都连一次邮箱；如果每轮都打 "connected as ..."，
    日志区会被刷屏，真正有用的行（分类了多少封、补了标签、出了报告）会被挤掉。

    状态**变化**的消息（连不上、恢复）不走这里，照常每次都打 —— 被静音的只是
    「又一次例行连接」。
    """
    if key in _seen_notices:
        return False
    _seen_notices.add(key)
    log(message)
    return True


_notice = note_once        # 兼容旧名字


# ── 收件箱约定（纯函数）──────────────────────────────────────────
def parse_sum_command(subject: str) -> Optional[str]:
    """主题是总结指令时返回规范化的级别串（如 'sum sa' → 'sa'），否则 None。"""
    match = SUM_COMMAND_RE.match((subject or "").strip())
    if not match:
        return None
    letters = (match.group(1) or "sabc").lower()
    return "".join(char for char in "sabc" if char in letters)


def is_report_subject(subject: str) -> bool:
    """主题是否是我们自己生成的总结报告（中英标记都认）。"""
    text = subject or ""
    return any(mark in text for mark in REPORT_SUBJECT_MARKS)


def is_report(msg) -> bool:
    """是否是我们自己生成的总结邮件。"""
    return is_report_subject(getattr(msg, "Subject", "") or "")


def category_names(msg) -> List[str]:
    """邮件当前的分类名列表（已去空白、去空项）。"""
    raw = getattr(msg, "Categories", "") or ""
    return [name.strip() for name in raw.split(",") if name.strip()]


def has_our_category(msg) -> bool:
    """是否已打过 S/A/B/C 或 Sum —— 都是本程序写的，默认不该再被覆盖。"""
    return any(name in LEVELS or name == SUM_NAME for name in category_names(msg))


def merge_categories(existing: Sequence[str], names: Sequence[str]) -> List[str]:
    """把新标签合并进已有分类：只剔掉本系统自己的名字，其余原样保留并追加新名字。

    这是「不误删用户分类」这条承诺的唯一实现点，抽成纯函数便于单测：
    旧实现是整串覆盖，用户自己打的标签会在分级那一刻被抹掉且无法撤销。
    """
    kept = [str(name).strip() for name in existing
            if str(name).strip() and str(name).strip() not in OUR_CATEGORY_NAMES]
    return kept + [str(name) for name in names]


def message_info(msg) -> Dict[str, Any]:
    """把一封邮件对象（GraphMessage 或任何同属性对象）转成内部统一字典。

    **逐字段兜错**：读不出的字段留空，尽量保住这封邮件的其余信息
    （有主题就还能判级），一封坏邮件不该让整轮挂掉。

    注意：`Body` 可能尚未加载（Graph 后端按需下载正文，见 graph.GraphSession.fill_body）。
    """
    def _read(name, default=None):
        try:
            value = getattr(msg, name)
        except Exception:
            return default
        return default if value is None else value

    return {
        "entryid": _read("EntryID", "") or "",
        "msg": msg,
        "subject": _read("Subject", "") or "(no subject)",
        "sender_name": _read("SenderName", ""),
        "sender": _read("SenderEmailAddress", ""),
        "body": str(_read("Body", "") or "").strip(),
        "received": _read("ReceivedTime"),
        "web_link": str(_read("WebLink", "") or "").strip(),
    }
