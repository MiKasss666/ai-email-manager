"""core/classifier.py —【核心 1】邮件分级。

职责：邮件列表 → 每封的级别（S/A/B/C）+ 截止日期，并可写回 Outlook 分类。

分级规则不写死在代码里：规则文本由 config 现场拼装（build_spec），
确定性的「我参与的课程 → S」也由 config 的课程列表生成。
换学校 / 换学院 / 换课程，只需要改 config.json。

决策顺序（与 spec.md 一致）：
    第 -1 步 学校邮件判定：AI 判「这封是不是学校邮件」；发件人域名命中 school_domains 的
             由代码直接认定（确定性，优先于 AI）。不是学校邮件 → 不贴任何标签
    第 0 步  我参与的课程                        → S（确定性正则，优先于 AI）
    第 1 步  来源（学院 / 宿舍 / 其它课程 / 个人） → A
    第 2 步  A 中 urgent_days 天内 deadline 或需交材料 → 升 S
    第 3 步  个人发展 / 学习类                    → B
    第 4 步  娱乐活动 / 其余                      → C

第 -1 步为什么必须让 AI 参与：只看发件人域名会漏 —— 学校事务经常由外部平台代发
（Canvas 通知、问卷、报名/活动系统…），域名不是学校的，内容却是学校事务。
反过来，域名命中的一定是学校邮件，所以那一类不问 AI，由代码直接认定。
判为「不是学校邮件」的邮件不写回 Outlook（保持原样，不贴 S/A/B/C）。
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
_SCHOOL_KEYS = ("学校", "school", "is_school", "属于学校")

# parse_school 认得的字面量（模型可能给 bool、0/1，也可能给中英文的是/否）
_SCHOOL_TRUE = {"true", "1", "yes", "y", "是", "对", "学校", "school"}
_SCHOOL_FALSE = {"false", "0", "no", "n", "否", "不是", "非学校", "not_school"}
# 模型偶尔整句回答，按前缀兜底：先判否定，再判肯定
_SCHOOL_FALSE_PREFIXES = ("不是", "不属于", "不算", "否", "非学校", "not", "no")
_SCHOOL_TRUE_PREFIXES = ("是", "属于", "算", "yes")


# ── 结果类型 ──────────────────────────────────────────────────────
@dataclass
class Verdict:
    """一封邮件的分级结果。"""

    level: str = "C"
    deadline: Optional[date] = None
    urgent: bool = False          # 仅 level == "S" 时才有意义（决定是否打红旗）
    is_school: bool = True        # 是不是学校邮件；False → 不写回、不贴标签
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
    domains = "、".join(cfg.school_domains) or na
    return f"""分级规则（先判来源属性，再判 S/A/B/C）：

第一步 · 这封是不是「学校邮件」？
- 算：学校/学院/教务/课程/宿舍/图书馆/校方服务发出的事务邮件 —— 考试、选课、缴费、报名、
  成绩、讲座与校方活动，以及学校采购的外部平台代发（Canvas 通知、问卷、报名系统…）
- 不算：私人邮件、商家推广、与学校无关的外部通知（电商、账单、社交平台、新闻订阅…）
- 发件人域名本身就是学校的（{domains}）→ 一律算学校邮件
- 不是学校邮件 → "学校": false，级别字段随便填；是 → "学校": true，再按第二步判级

第二步 · 学校邮件的 S/A/B/C 判级：
- 含这些我参与的课程 → 直接 S：{courses}
- A（来源，满足任一）：发件人来自「{sources}」；宿舍/住宿相关邮件（关键词：{dorms}）；其它课程编号（正则 {cfg.course_code_regex}）；称呼含「{cfg.name or na}」或正文含我的学号 {cfg.student_id or na}
- A 内升 S：{cfg.urgent_days} 天内 deadline，或需交资料/文件
- B（非 A）：个人发展(申研/计划申请/教授科研)、学习类
- C：娱乐/活动类（未报名）、其余兜底
- 已报名活动 → A 及以上"""


def _email_block(index: int, email: Dict[str, Any]) -> str:
    body = (email.get("body") or "")[:BODY_CHARS] or "(无正文)"
    return f"[{index}] {email.get('subject')}\n发件人: {email.get('sender')}\n正文: {body}"


def build_prompt(emails: Sequence[Dict[str, Any]], cfg: Config) -> str:
    """组装整批邮件的分级提示词。

    注意输出契约是**一个 JSON 对象**（里面 `items` 才是数组），不是裸数组：
    `ai.profiles.<name>.json_mode` 打开时请求的是 `response_format={"type":"json_object"}`，
    而它要求顶层必须是对象。旧提示词要裸数组，两者矛盾 —— 实测模型会先在正文里
    "纠结该听谁的"（日志里留下过整段内心戏），白烧输出额度、拖长到超时。
    """
    blocks = "\n\n".join(_email_block(i, e) for i, e in enumerate(emails, 1))
    return f"""按规则逐封处理，不要写摘要或解释。

{build_spec(cfg)}

处理下面 {len(emails)} 封邮件，输出**一个** JSON 对象：
- "items"：数组，每封邮件一个对象，顺序与邮件一致，数量必须正好 {len(emails)} 个
- 每个元素三个字段：
  · "学校"：true 或 false —— 这封是不是学校邮件（依据见「第一步」）
  · "级别"："S"|"A"|"B"|"C" —— 不是学校邮件时随便填
  · "截止日期"：YYYY-MM-DD，无则 null

形如：{{"items": [{{"学校": true, "级别": "S", "截止日期": "2026-03-05"}}]}}

只输出这一个 JSON 对象，不要多余文字。

{blocks}
"""


# ── 模型输出 → Verdict ────────────────────────────────────────────
def _pick(raw: dict, keys: Sequence[str]) -> Any:
    for key in keys:
        if key in raw and raw[key] not in (None, ""):
            return raw[key]
    return None


# 对象包裹时，数组可能挂在这些键名下面
_ITEM_KEYS = ("items", "results", "data", "mails", "邮件", "结果")


def _as_item_list(payload: Any) -> Optional[list]:
    """把模型输出规范成「每个元素是一封邮件的数组」；形状不对返回 None。

    兼容两种形状（不是每个后端都支持 response_format=json_object，降级重试后
    模型可能又给裸数组）：

      · 裸数组                    [ {...}, {...} ]
      · 对象包一层（现在的契约）   {"items": [ {...}, {...} ]}

    只认「列表」这一件事，元素类型仍由 to_verdicts 逐个校验。
    """
    if isinstance(payload, list):
        return payload
    if not isinstance(payload, dict):
        return None
    for key in _ITEM_KEYS:
        value = payload.get(key)
        if isinstance(value, list):
            return value
    # 键名不认识时：如果对象里**只有一个**列表值，就认它（别因为换个名字就整批作废）
    lists = [v for v in payload.values() if isinstance(v, list)]
    return lists[0] if len(lists) == 1 else None


def normalize_level(value) -> str:
    """把模型输出规范成 S/A/B/C；任何非法值一律回落 C。"""
    level = str(value or "").strip().upper()[:1]
    return level if level in LEVELS else "C"


def parse_school(value) -> Optional[bool]:
    """解析模型输出的「学校」字段；认不出来返回 None。

    模型可能给 bool、0/1，也可能给"是"/"不是学校邮件"这类字面量。
    返回 None 时调用方按「是学校邮件」保守处理 —— 漏判学校邮件（永远不贴标签）
    比多判一封（多贴个标签）严重得多。
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    text = str(value or "").strip().lower()
    if not text:
        return None
    if text in _SCHOOL_TRUE:
        return True
    if text in _SCHOOL_FALSE:
        return False
    # 短语兜底：先看否定，再看肯定（"不是学校邮件"同时以"不是"和"是"开头）
    if text.startswith(_SCHOOL_FALSE_PREFIXES):
        return False
    if text.startswith(_SCHOOL_TRUE_PREFIXES):
        return True
    return None


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

    形状兼容两种（见 _as_item_list）：裸数组，或对象里的 "items" 数组。
    """
    items = _as_item_list(payload)
    if items is None or len(items) != expected:
        return None

    out: List[Verdict] = []
    for item in items:
        if not isinstance(item, dict):
            return None
        level = normalize_level(_pick(item, _LEVEL_KEYS))
        deadline = parse_deadline(_pick(item, _DEADLINE_KEYS))
        school = parse_school(_pick(item, _SCHOOL_KEYS))
        out.append(Verdict(level=level, deadline=deadline,
                           urgent=_is_urgent(level, deadline, cfg),
                           is_school=True if school is None else school,
                           raw=item))
    return out


# ── 确定性规则：发件人域名属于学校 → 一定是学校邮件 ───────────────
def apply_school_override(emails: Sequence[Dict[str, Any]], verdicts: Sequence[Verdict],
                          cfg: Config) -> int:
    """发件人域名命中 school_domains → 强制 is_school=True（确定性证据优先于 AI）。

    只做「捞回来」这一个方向：AI 把学校邮件误判成非学校邮件时兜住它（那会导致
    该邮件永远不贴标签）。反方向不覆盖 —— AI 说"是学校邮件"就照它的判断走，
    因为「域名不是学校、内容却是学校事务」正是这一版要解决的场景。
    原地修改，返回捞回的封数。
    """
    changed = 0
    for email, verdict in zip(emails, verdicts):
        if not verdict.is_school and cfg.is_school_address(email.get("sender")):
            verdict.is_school = True
            changed += 1
    return changed


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
             log: Callable[..., None] = print,
             should_stop: Optional[Callable[[], bool]] = None) -> List[Optional[Verdict]]:
    """分批调用模型，返回与 emails 同序的结果（失败的位置为 None）。

    单批失败（调用失败，或返回条数对不上）只影响这一批：这些邮件保持 None，
    既不写回也不猜 —— 宁可漏贴标签，不可错贴标签。

    should_stop：任务被放弃（Force reset）时返回 True。只在**批与批之间**检查 ——
    已经发出去的那次 AI 调用收不回来，但不必再开下一批、白花额度。
    """
    results: List[Optional[Verdict]] = [None] * len(emails)
    size = cfg.batch_size

    for start in range(0, len(emails), size):
        if should_stop is not None and should_stop():
            log("  cancelled — not starting the next batch")
            break
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

        rescued = apply_school_override(batch, verdicts, cfg)
        if rescued:
            log(f"  deterministic rule (school domain) kept {rescued} mail(s) in scope")

        results[start:end] = verdicts

    return results


# ── 写回 Outlook ──────────────────────────────────────────────────
def apply_verdicts(emails: Sequence[Dict[str, Any]], verdicts: Sequence[Optional[Verdict]],
                   log: Callable[..., None] = print,
                   should_stop: Optional[Callable[[], bool]] = None) -> Tuple[int, int, int]:
    """把分级写回 Outlook：分类 + 仅「S 且加急」打红旗。返回 (成功, 失败, 跳过)。

    跳过 = 被判为「不是学校邮件」的：不写回、不贴标签，但**也不算失败** ——
    算失败会让监听以为这批没处理完，反复重试。

    should_stop：任务被放弃时返回 True → 逐封检查、立刻停止写邮箱。
    这是 Force reset 之后最要紧的一件事：被放弃的线程可能拿着指向**已经退出的
    Outlook 实例**的 COM 指针，继续写下去就是一片 RPC 报错（实际发生过）。
    没写到的那些不算失败，调用方据此不推进基线，下一轮重来。
    """
    ok = failed = skipped = 0

    for email, verdict in zip(emails, verdicts):
        if should_stop is not None and should_stop():
            log("  cancelled — stopping before writing the rest")
            break
        if verdict is None:
            failed += 1
            continue
        if not verdict.is_school:
            skipped += 1
            log(f"  [skip] {str(email.get('subject') or '')[:45]}")
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

    return ok, failed, skipped


def counts(verdicts: Sequence[Optional[Verdict]]) -> Dict[str, int]:
    """统计各档封数；被判为「不是学校邮件」的不计入（它们没被打标签）。"""
    out = {level: 0 for level in LEVELS}
    for verdict in verdicts or []:
        if verdict is not None and verdict.is_school:
            out[verdict.level] += 1
    return out


def not_school_entryids(emails: Sequence[Dict[str, Any]],
                        verdicts: Sequence[Optional[Verdict]]) -> List[str]:
    """被判为「不是学校邮件」的 EntryID 列表。

    给调用方（watcher）记录用：这些邮件不会被打标签，所以下一轮的
    read_inbox(skip_classified=True) 认不出它们 —— 不记名单就会每次启动
    把同一批垃圾重新问一遍 AI。
    """
    return [str(email.get("entryid")) for email, verdict in zip(emails, verdicts)
            if verdict is not None and not verdict.is_school and email.get("entryid")]


# ── 端到端入口 ────────────────────────────────────────────────────
def classify_inbox(cfg: Config, *, limit: Optional[int] = None, days: Optional[int] = None,
                   dry_run: bool = False, force: bool = False,
                   log: Callable[..., None] = print,
                   should_stop: Optional[Callable[[], bool]] = None) -> Dict[str, int]:
    """连接 → 读收件箱 → 分级 →（可选）写回。返回统计字典。

    dry_run=True 时只打印判定结果，绝不改邮件。
    force=True   时连已打过标签的邮件也重新判定（非学校邮件仍然不写回）。
    """
    _app, ns = outlook.connect(log=log)

    for change in outlook.ensure_categories(ns):
        log(f"  {change}")

    emails = outlook.read_inbox(ns, limit=limit, days=days, skip_classified=not force, log=log)
    log(f"read {len(emails)} mail(s) to classify")
    if not emails:
        return {"read": 0, "ok": 0, "failed": 0, "skipped": 0,
                **{level: 0 for level in LEVELS}}

    verdicts = classify(emails, cfg, log=log, should_stop=should_stop)

    if dry_run:
        for email, verdict in zip(emails, verdicts):
            if verdict is None:
                label, mark = "?", ""
            elif not verdict.is_school:
                label, mark = "skip", ""
            else:
                label, mark = verdict.level, (" 🔴" if verdict.urgent else "")
            log(f"  [{label}]{mark} {str(email.get('subject'))[:45]}")
        ok = sum(1 for v in verdicts if v is not None and v.is_school)
        skipped = sum(1 for v in verdicts if v is not None and not v.is_school)
        failed = len(verdicts) - ok - skipped
    else:
        ok, failed, skipped = apply_verdicts(emails, verdicts, log=log,
                                             should_stop=should_stop)

    stats = {"read": len(emails), "ok": ok, "failed": failed, "skipped": skipped}
    stats.update(counts(verdicts))
    return stats
