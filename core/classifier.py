"""core/classifier.py —【核心 1】邮件分级。

职责：邮件列表 → 每封的级别（S/A/B/C）+ 截止日期，并可写回 Outlook 分类。

分级规则不写死在代码里：规则文本由 config 现场拼装（build_spec），
确定性的「我参与的课程 → S」也由 config 的课程列表生成。
换学校 / 换学院 / 换课程，只需要改 config.json。

决策顺序：
    第 -1 步 学校邮件判定：AI 判「这封是不是学校邮件」；发件人域名命中 school_domains 的
              由代码直接认定（确定性，优先于 AI）。不是学校邮件 → 不贴任何标签
    第 0 步  我参与的课程                        → 至少 A（确定性正则，优先于 AI；
                                                    S 留给"要动手 + 有截止"）
    第 1 步  我学院/系发来的                     → 至少 A（确定性：显示名/地址/签名匹配，只升不降）
    第 2 步  需我做事 + urgent_days 天内截止      → S（**不限来源** —— 校级单位的报名截止同样算）
    第 3 步  来源（学院 / 宿舍 / 其它课程 / 点名到我） → A
    第 4 步  需要我报名/申请/提交                 → B
    第 5 步  通知型（简报/新闻/讲座通知/过期）     → C

第 -1 步为什么必须让 AI 参与：只看发件人域名会漏 —— 学校事务经常由外部平台代发
（Canvas 通知、问卷、报名/活动系统…），域名不是学校的，内容却是学校事务。
反过来，域名命中的一定是学校邮件，所以那一类不问 AI，由代码直接认定。
判为「不是学校邮件」的邮件不写回 Outlook（保持原样，不贴 S/A/B/C）。

第 1 步为什么必须由代码兜底：学院的真实身份常常**不在发件地址里** ——
`advise@ust.hk` 的显示名才是 `SCI/HOME, School of Science`，而正文签名里的学院名
又落在开头 400 字符之外。模型两处都看不到，于是同一个发件人会时 A 时 B 时 C
（实测：同一发件人的 6 封邮件出现过 4 个不同级别）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from . import ai, mail as outlook
from .config import LEVELS, Config

BODY_CHARS = 400        # 送进模型的正文**开头**截断长度
BODY_TAIL_CHARS = 300   # 另外附上正文**末尾**这么多字符 —— 签名（学院名、单位名）通常在那里

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
def faculty_terms(cfg: Config) -> List[str]:
    """我所在学院/系的识别词（拆词与去重规则在 config.faculty_terms，这里只转发）。"""
    return list(cfg.faculty_terms)


def build_spec(cfg: Config) -> str:
    """用配置拼出分级规则文本。

    config 里每个影响判级的字段都必须出现在这里，否则它就是「永远不起作用的开关」。
    反过来，字段为空时**不要**把条件写进去 —— 否则规则里会留下
    「称呼含「(未填写)」」这种永远不可能成立的废话，白白干扰模型。
    """
    na = "(not set)" if cfg.is_english else "(未填写)"
    courses = "、".join(cfg.courses) or na
    faculty = "、".join(faculty_terms(cfg)) or na
    dorms = "、".join(cfg.dorm_keywords) or na
    domains = "、".join(cfg.school_domains) or na

    # 只在真的配置了的时候才写进规则；没配置的字段绝不出现，免得留下
    # 「称呼含「(未填写)」」这种永远不可能成立的条件去干扰模型。
    terms = faculty_terms(cfg)
    if terms:
        a_clauses = [f"来自我所在的学院/系/课程组 —— 发件人显示名、发件地址或正文签名里出现"
                     f"「{'、'.join(terms)}」（学院组织的讲座、活动、招聘会也算 A）"]
    else:
        a_clauses = []
    if cfg.dorm_keywords:
        a_clauses.append(f"宿舍/住宿相关（关键词：{dorms}）")
    if cfg.course_code_regex:
        a_clauses.append(f"其它课程编号（正则 {cfg.course_code_regex}）")
    who = []
    if cfg.name:
        who.append(f"称呼含「{cfg.name}」")
    if cfg.student_id:
        who.append(f"正文含我的学号 {cfg.student_id}")
    if who:
        a_clauses.append("点名到我的：" + " 或 ".join(who))
    a_clauses.append("与我的学业直接相关的官方事务：选课、考试、成绩、报名、缴费、"
                     "表格/材料、毕业、交换项目、图书馆到期提醒")
    a_list = "\n".join(f"  · {c}" for c in a_clauses)
    s_line = (f"- 含这些我参与的课程 → 直接 S：{courses}" if cfg.courses
              else "- （没有配置课程清单，不按课程名判 S；S 由下面的 deadline / 需交材料决定）")

    return f"""分级规则（先判来源属性，再判 S/A/B/C）：

第一步 · 这封是不是「学校邮件」？
- 算：学校/学院/教务/课程/宿舍/图书馆/校方服务发出的事务邮件 —— 考试、选课、缴费、报名、
  成绩、讲座与校方活动，以及学校采购的外部平台代发（Canvas 通知、问卷、报名系统…）
- 不算：私人邮件、商家推广、与学校无关的外部通知（电商、账单、社交平台、新闻订阅…）
- 发件人域名本身就是学校的（{domains}）→ 一律算学校邮件
- 不是学校邮件 → "学校": false，级别字段随便填；是 → "学校": true，再按第二步判级

第二步 · 学校邮件的 S/A/B/C 判级（**先判 S，再判 A/B/C**）：
{s_line}
- S（第二条件，不看发件人是谁）：**需要我做某件事**且有 {cfg.urgent_days} 天内的截止日期 ——
  交作业、交表格/材料、报名截止、缴费截止、考试时间。来自学院是 S，来自校级单位也是 S。
  注意两点：
  · **活动举办日期不算截止**。「某讲座/招聘会/欢迎会 10 月 20 日举行」不是 S，
    只有"我必须在那之前做点什么"才是；
  · **已经过期的**（报名/活动日期已过、也不再需要我补做什么）→ 不是 S，按 A/B/C 判，
    通常落到 C。
- A（满足任一）：
{a_list}
- B（非 A 也非 S）：**需要我报名/申请/提交**的东西 —— 招聘会与 Career Center 的 Co-op/实习项目、
  工作坊、语言或心理课程、创业比赛、交流与夏令营报名、科研机会、申研与计划申请
- C：通知型内容，我不需要做任何事 —— HKUST Daily Event Alert、Press Release 新闻稿、
  讲座通知（如 IAS Distinguished Lecture）、图书馆开放时间、社团活动宣传、比赛围观、
  商家推广、已过期的通知、其余兜底
- 已报名的活动 → A（要交材料或有截止才升 S）

判断 B 还是 C 只看一件事：**要不要我报名/申请/提交**。要 → B；只是告诉我有个活动 → C。"""


def _email_block(index: int, email: Dict[str, Any]) -> str:
    """一封邮件送进模型的样子。

    两处细节都是踩出来的：
      · **发件人显示名必须带上** —— 学院/单位的真实身份常常只在这里
        （例：`SCI/HOME, School of Science <advise@ust.hk>`，光看地址看不出是学院）；
      · **正文末尾也要给一段** —— 签名（学院名、单位名、系别）几乎总在末尾，
        只给开头的 400 字符时模型看不到，于是同一个发件人会时 A 时 C。
    """
    body = (email.get("body") or "").strip()
    head = body[:BODY_CHARS] or "(无正文)"
    tail = ""
    if len(body) > BODY_CHARS:
        # 不重复：短一点的正文补上"头部没覆盖到的部分"，长正文只取最后一段
        tail = (body[BODY_CHARS:] if len(body) <= BODY_CHARS + BODY_TAIL_CHARS
                else body[-BODY_TAIL_CHARS:])

    name = str(email.get("sender_name") or "").strip()
    addr = str(email.get("sender") or "").strip()
    who = f"{name} <{addr}>" if name and addr else (name or addr or "(未知发件人)")

    block = f"[{index}] {email.get('subject')}\n发件人: {who}\n正文: {head}"
    if tail:
        block += f"\n（正文末尾节选，签名通常在这里）: {tail}"
    return block


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
    """命中我参与的课程 → **至少 A**（优先于 AI，只升不降）。原地修改，返回改写的封数。

    早先是"直接强制 S"，实测太钝：Canvas / StashTag 那些「lecture 12 has started!」
    「Assignment Graded」之类**不需要我动手**的通知也被顶成 S，S 里三分之一都是这类，
    真正紧急的反而淹没了。
    现在这条规则只保证「课程邮件不会被当成可忽略」；S 仍然由"要不要我做 + 有没有截止"
    来判（AI 说 S 就保持 S，不降级）。
    """
    rx = enrolled_regex(cfg)
    if rx is None:
        return 0

    changed = 0
    for email, verdict in zip(emails, verdicts):
        haystack = f"{email.get('subject') or ''}\n{email.get('body') or ''}"
        if rx.search(haystack) and LEVELS.index(verdict.level) > LEVELS.index("A"):
            verdict.level = "A"
            verdict.urgent = False
            changed += 1
    return changed


# ── 确定性规则：我学院的邮件 → 至少 A ─────────────────────────────
def apply_faculty_override(emails: Sequence[Dict[str, Any]], verdicts: Sequence[Verdict],
                           cfg: Config) -> int:
    """学院/系发来的邮件 → 至少 A（优先于 AI，只升不降）。原地修改，返回改写封数。

    为什么必须由代码兜底：学院的真实身份常常**不在发件地址里** ——
    `advise@ust.hk` 的显示名才是 `SCI/HOME, School of Science`，
    而正文签名里的学院名又落在开头 400 字符之外。模型两处都看不到，
    于是同一个发件人会时 A 时 B 时 C（实测：6 封里出现过 4 个不同级别）。
    """
    terms = [t.lower() for t in faculty_terms(cfg)]
    if not terms:
        return 0

    changed = 0
    for email, verdict in zip(emails, verdicts):
        if not verdict.is_school:            # 非学校邮件本来就不写回，不动它
            continue
        if LEVELS.index(verdict.level) <= LEVELS.index("A"):
            continue                          # 已经是 S/A，不降级
        haystack = "\n".join(str(email.get(k) or "")
                             for k in ("sender_name", "sender", "subject", "body")).lower()
        if any(term in haystack for term in terms):
            verdict.level = "A"
            verdict.urgent = False
            changed += 1
    return changed


# ── 分批调用模型 ──────────────────────────────────────────────────
SPLIT_MIN = 2        # 拆到这个大小还不成功就不再拆（单封失败没有更小的可拆）
SPLIT_MAX_DEPTH = 3  # 最多拆三层（40 → 20 → 10 → 5）


def _ask_batch(batch: Sequence[Dict[str, Any]], cfg: Config,
               log: Callable[..., None] = print,
               should_stop: Optional[Callable[[], bool]] = None,
               depth: int = 0) -> List[Optional[Verdict]]:
    """一批邮件 → 与 batch 同序的结果（失败的位置为 None）。

    整批失败时**对半拆开重试**。踩过的坑：一批 40 封时模型会因为输出额度被撑爆
    （它有时先写一段推理再吐 JSON）而整批作废 —— 40 封邮件一起白跑。
    拆成两半后即使一半仍然失败，另一半也能正常贴标签。
    """
    if should_stop is not None and should_stop():
        return [None] * len(batch)

    payload = ai.ask_json(build_prompt(batch, cfg), cfg.ai, log=log)
    verdicts = to_verdicts(payload, len(batch), cfg) if payload is not None else None
    if verdicts is not None:
        return list(verdicts)

    if len(batch) <= SPLIT_MIN or depth >= SPLIT_MAX_DEPTH:
        return [None] * len(batch)

    half = len(batch) // 2
    log(f"  batch of {len(batch)} failed — retrying as {half} + {len(batch) - half}")
    return (_ask_batch(batch[:half], cfg, log, should_stop, depth + 1)
            + _ask_batch(batch[half:], cfg, log, should_stop, depth + 1))


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

        got = _ask_batch(batch, cfg, log, should_stop)
        ok = [(email, verdict) for email, verdict in zip(batch, got) if verdict is not None]
        if not ok:
            log(f"  batch failed, {len(batch)} mail(s) left unlabelled")
            continue
        if len(ok) < len(batch):
            log(f"  {len(batch) - len(ok)} mail(s) in this batch could not be classified")

        batch_emails = [email for email, _ in ok]
        batch_verdicts = [verdict for _, verdict in ok]

        changed = apply_enrolled_override(batch_emails, batch_verdicts, cfg)
        if changed:
            log(f"  deterministic rule (enrolled course) raised {changed} mail(s) to A")

        faculty = apply_faculty_override(batch_emails, batch_verdicts, cfg)
        if faculty:
            log(f"  deterministic rule (my faculty) raised {faculty} mail(s) to A")

        rescued = apply_school_override(batch_emails, batch_verdicts, cfg)
        if rescued:
            log(f"  deterministic rule (school domain) kept {rescued} mail(s) in scope")

        results[start:end] = got

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
                   should_stop: Optional[Callable[[], bool]] = None,
                   phase: Optional[Callable[[str], None]] = None) -> Dict[str, int]:
    """连接 → 读收件箱 → 分级 →（可选）写回。返回统计字典。

    dry_run=True 时只打印判定结果，绝不改邮件。
    force=True   时连已打过标签的邮件也重新判定（非学校邮件仍然不写回）。
    phase        当前阶段的回调 —— 界面据此把「卡在 AI」和「卡在 Outlook」分开报，
                 不再笼统地让用户去折腾 Outlook。
    """
    def _phase(text: str) -> None:
        if phase is not None:
            phase(text)

    _phase("connecting to Outlook")
    _app, ns = outlook.connect(log=log)

    for change in outlook.ensure_categories(ns):
        log(f"  {change}")

    emails = outlook.read_inbox(ns, limit=limit, days=days, skip_classified=not force, log=log)
    log(f"read {len(emails)} mail(s) to classify")
    if not emails:
        return {"read": 0, "ok": 0, "failed": 0, "skipped": 0,
                **{level: 0 for level in LEVELS}}

    _phase(f"waiting for the AI ({cfg.ai_active})")
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
        _phase("writing categories to Outlook")
        ok, failed, skipped = apply_verdicts(emails, verdicts, log=log,
                                             should_stop=should_stop)

    stats = {"read": len(emails), "ok": ok, "failed": failed, "skipped": skipped}
    stats.update(counts(verdicts))
    return stats
