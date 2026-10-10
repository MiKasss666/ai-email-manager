"""core/summarizer.py —【核心 2】邮件总结。

职责：把最近 N 天里已分级的邮件，按级别整理成一份 HTML 报告发到自己邮箱。
每封邮件配一句内容要点，报告里的主题是可点开的 outlook: 链接（点开直接跳回原邮件）。
报告邮件本身打黄色 Sum 分类，下次扫收件箱时会被自动跳过。
"""
from __future__ import annotations

from typing import Any, Callable, Dict, List, Sequence

from . import ai, outlook
from .config import Config

BODY_CHARS = 800  # 生成要点时的正文截断长度（要点要写出截止日期/课程代号等细节，给足上下文）

LEVEL_COLORS = {"S": "#e67e22", "A": "#8e44ad", "B": "#2980b9", "C": "#7f8c8d"}
FONT_STACK = "'Segoe UI','Microsoft YaHei',sans-serif"
LEVEL_NAMES = {"s": "S", "a": "A", "b": "B", "c": "C"}


# ── 级别串与文案 ──────────────────────────────────────────────────
def parse_letters(value: str) -> str:
    """把 'sabc' / 'SA' / 's,a' 之类规范成去重且按 S→A→B→C 排列的字母串；空则 sabc。"""
    letters = "".join(c for c in str(value or "").lower() if c in "sabc")
    return "".join(c for c in "sabc" if c in letters) or "sabc"


def label_for(letters: str, english: bool = False) -> str:
    """'sa' → 'S+A 级'；'sabc' → '全部'。"""
    if letters == "sabc":
        return "all levels" if english else "全部"
    names = [LEVEL_NAMES[c] for c in letters if c in LEVEL_NAMES]
    return "+".join(names) + (" levels" if english else " 级")


def _esc(text) -> str:
    """HTML 转义，避免邮件主题里的尖括号破坏报告结构。"""
    return (str(text or "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# ── 一句话要点 ────────────────────────────────────────────────────
def build_points_prompt(emails: Sequence[Dict[str, Any]], cfg: Config) -> str:
    """组装「每封一句要点」的提示词；语言跟随 cfg.language。"""
    en = cfg.is_english
    blocks = []
    for i, email in enumerate(emails, 1):
        body = (email.get("body") or "")[:BODY_CHARS] or ("(no body)" if en else "(无正文)")
        if en:
            blocks.append(f"[{i}] Subject: {email.get('subject')}\n"
                          f"From: {email.get('sender')}\nBody: {body}")
        else:
            blocks.append(f"[{i}] 主题:{email.get('subject')}\n"
                          f"发件人:{email.get('sender')}\n正文:{body}")

    count = len(emails)
    listing = chr(10).join(blocks)
    if en:
        return (f"For each email below, write ONE short line saying what it is and what I need to do.\n"
                f"Rules:\n"
                f"- If the email has a deadline, date, course code, amount, venue or required "
                f"document, include it in the line.\n"
                f"- Never mention the level (S/A/B/C) or the sender.\n"
                f"- No filler like \"this email\"; get straight to the point.\n"
                f"- One line per email, at most about 40 words.\n"
                f"Output only a JSON array of strings, one per email in the same order, "
                f"exactly {count} items.\n\n{listing}\n")
    return (f"给下面每封邮件写一句中文要点（一句话说清：这封邮件是什么、要我做什么）。\n"
            f"要求：\n"
            f"- 邮件里若有明确的截止时间/日期、课程代号、金额、地点、需要提交的材料，写进要点；\n"
            f"- 不要写级别（S/A/B/C）和发件人；\n"
            f"- 不要写「这封邮件」「通知你」这类废话，直接说重点；\n"
            f"- 每封只写一句，控制在 40 字左右。\n"
            f"只输出 JSON 数组，元素按邮件顺序，每个是一个字符串（要点），数量必须正好 {count} 个。\n\n"
            f"{listing}\n")


def _as_string_list(payload: Any) -> Optional[List[str]]:
    """把模型返回的东西规范成「字符串列表」；形状不对就返回 None。

    实测见过的三种形状都要吃下：
      · ["要点1","要点2"]                                        —— 正常
      · {"summaries": ["要点1","要点2"]}                          —— 单键包装
      · {"type":"json_object","content":"[\\"要点1\\", ...]"}     —— 把数组当字符串塞进包装对象
        （deepseek-flash 会偶发这种；不兼容的话每次都要靠重试救，白白多花一次调用）

    只接受**全是字符串**的数组：万一模型在解释性文字里夹了杂乱数组，
    也不会被当成要点混进报告。
    """
    candidates = list(payload.values()) if isinstance(payload, dict) else [payload]
    for candidate in candidates:
        if isinstance(candidate, str):
            candidate = ai.parse_json(candidate) or ai.salvage_array(candidate)
        if (isinstance(candidate, list) and candidate
                and all(isinstance(item, str) for item in candidate)):
            return candidate
    return None


def _points_batch(batch: Sequence[Dict[str, Any]], cfg: Config,
                  log: Callable[..., None]) -> List[str]:
    """给一批邮件要要点，返回长度恰好等于 len(batch) 的列表（拿不到的填空串）。

    为什么不是"失败就整批留空"（旧实现）：整批要求"数量正好 N 个"，而响应可能被
    max_tokens 截断 —— 于是 99 封里最后 19 封会全部没有要点（实际发生过）。
    现在的策略从便宜到贵：

      1. 要一次；形状和数量都对 → 直接用（绝大多数情况）
      2. 只拿到前 k 个（被截断）→ 先收下，只对剩下的 k+1..N 再要
      3. 整批什么都没拿到 → 劈成两半递归重试，最差退化到单封，只影响那一封
    """
    if not batch:
        return []

    payload = ai.ask_json_lenient(build_points_prompt(batch, cfg), cfg.ai, log=log)
    items = _as_string_list(payload)

    if items is not None and len(items) >= len(batch):
        return items[:len(batch)]

    got = len(items) if items else 0
    if got:
        log(f"  only {got}/{len(batch)} one-liner(s) came back, asking for the rest")
    if len(batch) == 1:
        log(f"  no one-liner for \"{str(batch[0].get('subject'))[:44]}\", left blank")
        return [""]

    if not got:
        log(f"  batch of {len(batch)} produced nothing, splitting and retrying")

    rest = list(batch[got:])
    mid = len(rest) // 2
    return (list(items or [])[:got]
            + _points_batch(rest[:mid], cfg, log)
            + _points_batch(rest[mid:], cfg, log))


def summarize_points(emails: Sequence[Dict[str, Any]], cfg: Config,
                     log: Callable[..., None] = print) -> List[str]:
    """给每封邮件生成一句要点，返回与 emails 同序的要点列表（缺的填空串）。"""
    points: List[str] = []
    size = max(1, cfg.batch_size)

    for start in range(0, len(emails), size):
        batch = list(emails[start:start + size])
        points.extend(_points_batch(batch, cfg, log))

    missing = sum(1 for p in points if not p)
    if missing:
        log(f"  {missing} mail(s) got no one-liner (the rest of the report is unaffected)")
    return points


# ── HTML 报告 ─────────────────────────────────────────────────────
def build_html(emails: Sequence[Dict[str, Any]], points: Sequence[str], cfg: Config,
               letters: str) -> str:
    """按级别分组生成 HTML：每封一块（级别+发件人 / 主题链接 / 要点 / 时间）。"""
    en = cfg.is_english
    title = f"📬 Email summary（{label_for(letters, en)}）" if en \
        else f"📬 邮件总结（{label_for(letters, en)}）"
    scope = f"last {cfg.summary_days} days" if en else f"最近 {cfg.summary_days} 天"
    total = f"{len(emails)} messages" if en else f"共 {len(emails)} 封"

    html = [
        f"<html><body style=\"font-family:{FONT_STACK};font-size:14px;color:#222;line-height:1.6;\">",
        f"<h2 style='margin:0 0 4px 0;'>{title} · {scope}</h2>",
        f"<p style='margin:0 0 8px 0;color:#888;'>{total}</p>",
    ]

    current = None
    for email, point in zip(emails, points):
        if email["level"] != current:
            if current is not None:
                html.append("</div>")
            current = email["level"]
            html.append(f"<h3 style='margin:22px 0 10px 0;color:{LEVEL_COLORS.get(current, '#000')};'>"
                        f"【{current}】</h3>")
            html.append("<div>")

        lines = [f"<div>[{current}] <b>{_esc(email.get('sender'))}</b></div>"]
        lines.append(f"<div><a href='outlook:{email.get('entryid')}'>{_esc(email.get('subject'))}</a></div>")
        if point:
            lines.append(f"<div>{_esc(point)}</div>")
        received = email.get("received")
        if received:
            stamp = received.strftime("%b %d %H:%M") if en else received.strftime("%m月%d日 %H:%M")
            lines.append(f"<div style='color:#999;font-size:12px;'>{stamp}</div>")

        html.append("<div style='margin:0 0 16px 0;padding-bottom:10px;"
                    "border-bottom:1px solid #e6e6e6;'>" + "".join(lines) + "</div>")

    if current is not None:
        html.append("</div>")
    html.append("</body></html>")
    return "\n".join(html)


def build_subject(letters: str, cfg: Config) -> str:
    """报告邮件主题。标记跟随 language —— 因为 outlook.is_report_subject()
    中英两个标记都认，所以切换语言之后历史报告仍然能被识别、不会被误分级。
    """
    label = label_for(letters, cfg.is_english)
    if cfg.is_english:
        return f"📬 {outlook.REPORT_SUBJECT_MARK_EN} ({label}) · last {cfg.summary_days} days"
    return f"📬 {outlook.REPORT_SUBJECT_MARK}（{label}）· 最近{cfg.summary_days}天"


# ── 发送 ──────────────────────────────────────────────────────────
def send_report(app, ns, cfg: Config, emails: Sequence[Dict[str, Any]],
                points: Sequence[str], letters: str) -> str:
    """生成并发出报告邮件，返回主题。"""
    subject = build_subject(letters, cfg)
    outlook.ensure_sum_category(ns)
    outlook.send_html(app, cfg.recipient, subject, build_html(emails, points, cfg, letters),
                      categories=outlook.SUM_NAME)
    return subject


# ── 端到端入口 ────────────────────────────────────────────────────
def summarize_inbox(cfg: Config, letters: str = "sabc", *,
                    log: Callable[..., None] = print) -> int:
    """连接 → 读最近 N 天已分级邮件 → 生成要点 → 发报告给自己。

    返回整理的邮件数；0 表示没有匹配邮件，不会发送任何东西。
    """
    letters = parse_letters(letters)
    app, ns = outlook.connect(log=log)

    log(f"scanning the last {cfg.summary_days} day(s), levels: {letters.upper()} ...")
    emails = outlook.read_recent(ns, days=cfg.summary_days, letters=letters)
    log(f"matched {len(emails)} mail(s)")
    if not emails:
        log("nothing to summarise, no report sent.")
        return 0

    log("writing one line per mail ...")
    points = summarize_points(emails, cfg, log=log)

    subject = send_report(app, ns, cfg, emails, points, letters)
    log(f"sent to {cfg.recipient}: {subject}")
    return len(emails)
