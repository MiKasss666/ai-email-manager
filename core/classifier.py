"""core/classifier.py —【核心 1】邮件分级。

职责：邮件列表 → 每封的级别（S/A/B/C）+ 截止日期，并可写回 Outlook 分类。

分级规则不写死在代码里：规则文本由 config 现场拼装（build_spec），
确定性的「我参与的课程 → S」也由 config 的课程列表生成。
换学校 / 换学院 / 换课程，只需要改 config.json。

决策顺序（与 spec.md 一致）：
    第 0 步  我参与的课程                        → S（确定性正则，优先于 AI）
    第 1 步  来源（学院 / 宿舍 / 其它课程 / 个人） → A
    第 2 步  A 中 urgent_days 天内 deadline 或需交材料 → 升 S
    第 3 步  个人发展 / 学习类                    → B
    第 4 步  娱乐活动 / 其余                      → C
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from . import ai, outlook
from .config import LEVELS, Config

BODY_CHARS = 400  # 送进模型的正文截断长度（判级只需要依据，不需要全文）

_LEVEL_KEYS = ("级别", "level")
_DEADLINE_KEYS = ("截止日期", "deadline", "due")


# ── 结果类型 ──────────────────────────────────────────────────────
@dataclass
class Verdict:
    """一封邮件的分级结果。"""

    level: str = "C"
    deadline: Optional[date] = None
    urgent: bool = False          # 仅 level == "S" 时才有意义（决定是否打红旗）
    raw: Optional[dict] = None    # 模型原始输出，排查问题用


# ── 规则文本（提示词的唯一真源）────────────────────────────────────
def build_spec(cfg: Config) -> str:
    """用配置拼出分级规则文本。

    config 里每个影响判级的字段都必须出现在这里，否则它就是「永远不起作用的开关」。
    """
    na = "(not set)" if cfg.is_english else "(未填写)"
    courses = "、".join(cfg.courses) or na
    sources = "、".join(cfg.faculty_terms) or na
    dorms = "、".join(cfg.dorm_keywords) or na
    return f"""分级规则（S/A/B/C）：
- 含这些我参与的课程 → 直接 S：{courses}
- A（来源，满足任一）：发件人来自「{sources}」或 {cfg.school or na}；宿舍/住宿相关邮件（关键词：{dorms}）；其它课程编号（正则 {cfg.course_code_regex}）；称呼含「{cfg.name or na}」或正文含我的学号 {cfg.student_id or na}
- A 内升 S：{cfg.urgent_days} 天内 deadline，或需交资料/文件
- B（非 A）：个人发展(申研/计划申请/教授科研)、学习类
- C：娱乐/活动类（未报名）、其余兜底
- 已报名活动 → A 及以上"""


def _email_block(index: int, email: Dict[str, Any]) -> str:
    body = (email.get("body") or "")[:BODY_CHARS] or "(无正文)"
    return f"[{index}] {email.get('subject')}\n发件人: {email.get('sender')}\n正文: {body}"


def build_prompt(emails: Sequence[Dict[str, Any]], cfg: Config) -> str:
    """组装整批邮件的分级提示词。"""
    blocks = "\n\n".join(_email_block(i, e) for i, e in enumerate(emails, 1))
    return f"""按规则直接给每封邮件分级（S/A/B/C），不要写摘要或解释。

{build_spec(cfg)}

处理下面 {len(emails)} 封邮件，每封输出一个 JSON 对象，字段：
- "级别": "S"|"A"|"B"|"C"
- "截止日期": YYYY-MM-DD，无则 null

只输出 JSON 数组（顺序与邮件一致，数量必须正好 {len(emails)} 个），不要多余文字。

{blocks}
"""


# ── 模型输出 → Verdict ────────────────────────────────────────────
def _pick(raw: dict, keys: Sequence[str]) -> Any:
    for key in keys:
        if key in raw and raw[key] not in (None, ""):
            return raw[key]
    return None


def normalize_level(value) -> str:
    """把模型输出规范成 S/A/B/C；任何非法值一律回落 C。"""
    level = str(value or "").strip().upper()[:1]
    return level if level in LEVELS else "C"


def parse_deadline(value) -> Optional[date]:
    """解析 YYYY-MM-DD（也容忍 ISO 时间串）；解析不出来返回 None。"""
    text = str(value or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _is_urgent(level: str, deadline: Optional[date], cfg: Config) -> bool:
    """S 级且 deadline 落在 [今天, 今天 + urgent_days] 内 → 加急。"""
    if level != "S" or deadline is None:
        return False
    return 0 <= (deadline - date.today()).days <= cfg.urgent_days


def to_verdicts(payload: Any, expected: int, cfg: Config) -> Optional[List[Verdict]]:
    """把模型返回的 JSON 变成 Verdict 列表。

    数量与 expected 不一致时返回 None：宁可整批跳过，也绝不能错位贴标签
    （旧版直接 zip，模型少数一条就会把后面所有邮件的级别整体错位）。
    """
    if not isinstance(payload, list) or len(payload) != expected:
        return None

    out: List[Verdict] = []
    for item in payload:
        if not isinstance(item, dict):
            return None
        level = normalize_level(_pick(item, _LEVEL_KEYS))
        deadline = parse_deadline(_pick(item, _DEADLINE_KEYS))
        out.append(Verdict(level=level, deadline=deadline,
                           urgent=_is_urgent(level, deadline, cfg), raw=item))
    return out


# ── 确定性规则：我参与的课程 → S ──────────────────────────────────
def enrolled_regex(cfg: Config):
    """「我参与的课程」正则（大小写不敏感，两侧不能紧邻字母数字）；无课程返回 None。"""
    if not cfg.courses:
        return None
    pattern = "|".join(re.escape(c) for c in cfg.courses)
    return re.compile(r"(?<![A-Za-z0-9])(" + pattern + r")(?![A-Za-z0-9])", re.IGNORECASE)


def apply_enrolled_override(emails: Sequence[Dict[str, Any]], verdicts: Sequence[Verdict],
                            cfg: Config) -> int:
    """命中我参与的课程 → 强制 S（优先于 AI）。原地修改，返回改写的封数。"""
    rx = enrolled_regex(cfg)
    if rx is None:
        return 0

    changed = 0
    for email, verdict in zip(emails, verdicts):
        haystack = f"{email.get('subject') or ''}\n{email.get('body') or ''}"
        if rx.search(haystack) and verdict.level != "S":
            verdict.level = "S"
            verdict.urgent = _is_urgent("S", verdict.deadline, cfg)
            changed += 1
    return changed


# ── 分批调用模型 ──────────────────────────────────────────────────
def classify(emails: Sequence[Dict[str, Any]], cfg: Config,
             log: Callable[..., None] = print) -> List[Optional[Verdict]]:
    """分批调用模型，返回与 emails 同序的结果（失败的位置为 None）。

    单批失败（调用失败，或返回条数对不上）只影响这一批：这些邮件保持 None，
    既不写回也不猜 —— 宁可漏贴标签，不可错贴标签。
    """
    results: List[Optional[Verdict]] = [None] * len(emails)
    size = cfg.batch_size

    for start in range(0, len(emails), size):
        batch = emails[start:start + size]
        end = start + len(batch)
        log(f"[batch {start // size + 1}] mails {start + 1}-{end} ...")

        payload = ai.ask_json(build_prompt(batch, cfg), cfg.ai, log=log)
        verdicts = to_verdicts(payload, len(batch), cfg) if payload is not None else None
        if verdicts is None:
            log(f"  batch failed, {len(batch)} mail(s) left unlabelled")
            continue

        changed = apply_enrolled_override(batch, verdicts, cfg)
        if changed:
            log(f"  deterministic rule (enrolled course) forced {changed} mail(s) to S")

        results[start:end] = verdicts

    return results


# ── 写回 Outlook ──────────────────────────────────────────────────
def apply_verdicts(emails: Sequence[Dict[str, Any]], verdicts: Sequence[Optional[Verdict]],
                   log: Callable[..., None] = print) -> Tuple[int, int]:
    """把分级写回 Outlook：分类 + 仅「S 且加急」打红旗。返回 (成功, 失败)。"""
    ok = failed = 0

    for email, verdict in zip(emails, verdicts):
        if verdict is None:
            failed += 1
            continue

        flag = None
        if verdict.level == "S":
            flag = outlook.OL_FLAG_MARKED if verdict.urgent else outlook.OL_FLAG_NONE

        subject = str(email.get("subject") or "")[:45]
        try:
            outlook.write_categories(email["msg"], [verdict.level], flag)
        except Exception as ex:
            failed += 1
            log(f"  [write failed] {verdict.level} {subject}: {ex}")
        else:
            ok += 1
            log(f"  [{verdict.level}]{' 🔴' if verdict.urgent else ''} {subject}")

    return ok, failed


def counts(verdicts: Sequence[Optional[Verdict]]) -> Dict[str, int]:
    """统计各档封数。"""
    out = {level: 0 for level in LEVELS}
    for verdict in verdicts or []:
        if verdict is not None:
            out[verdict.level] += 1
    return out


# ── 端到端入口 ────────────────────────────────────────────────────
def classify_inbox(cfg: Config, *, limit: Optional[int] = None, days: Optional[int] = None,
                   dry_run: bool = False, force: bool = False,
                   log: Callable[..., None] = print) -> Dict[str, int]:
    """连接 → 读收件箱 → 分级 →（可选）写回。返回统计字典。

    dry_run=True 时只打印分级结果，绝不改邮件。
    force=True   时连已打过标签的邮件也重新分级。
    """
    _app, ns = outlook.connect(log=log)

    for change in outlook.ensure_categories(ns):
        log(f"  {change}")

    emails = outlook.read_inbox(ns, limit=limit, days=days, skip_classified=not force)
    log(f"read {len(emails)} mail(s) to classify")
    if not emails:
        return {"read": 0, "ok": 0, "failed": 0, **{level: 0 for level in LEVELS}}

    verdicts = classify(emails, cfg, log=log)

    if dry_run:
        for email, verdict in zip(emails, verdicts):
            label = verdict.level if verdict else "?"
            mark = " 🔴" if verdict and verdict.urgent else ""
            log(f"  [{label}]{mark} {str(email.get('subject'))[:45]}")
        ok = sum(1 for v in verdicts if v is not None)
        failed = len(verdicts) - ok
    else:
        ok, failed = apply_verdicts(emails, verdicts, log=log)

    stats = {"read": len(emails), "ok": ok, "failed": failed}
    stats.update(counts(verdicts))
    return stats
