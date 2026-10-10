"""gui/watcher.py — 自动监听收件箱。

**配置齐全时才自动开始**（缺 email / AI 后端不可用就先不开，见
main_window._maybe_start_watching）：每 N 秒看一次收件箱里最新的那封，
是新邮件就分级；是 `sum` 指令就出总结报告。

四个设计要点：
  · **复用 TaskRunner** —— 监听任务和手动点的 Categorize / Sum 走同一个执行器，
    天然串行，永远不会出现两个任务同时写邮箱。
  · **只处理「比上次记住的那封更新」的邮件**（collect_new_mail，上限 100 封 / 往回 7 天），
    不去遍历整个历史收件箱 —— 逐封读正文在 COM 上非常慢，大邮箱下要几分钟。
  · **状态存 watch_state.json**（上次记住的邮件 + 轮询间隔），所以关掉程序期间到的
    新邮件，下次启动会被补上。
  · **Outlook 没开不算失败**：只提示一次、继续轮询，用户一打开 Outlook 就自动接上；
    真正的错误连续 3 次才自动暂停。

判断与推进逻辑抽成了纯函数 plan_tick() / advance_policy()，可以脱离 Outlook 单测。
"""
from __future__ import annotations

import json
import os
import tkinter as tk
from typing import Any, Callable, Dict, Optional, Tuple

from core import classifier, config, outlook, summarizer

STATE_NAME = "watch_state.json"
DEFAULT_INTERVAL = 30      # 秒
MIN_INTERVAL = 5
MAX_INTERVAL = 3600
MAX_CONSECUTIVE_ERRORS = 3
FIRST_TICK_DELAY = 1       # 启动后 1 秒跑第一轮，让界面先显示出来
BUSY_RETRY_DELAY = 2       # 有任务在跑时，过 2 秒再看

MAX_MAIL_PER_ROUND = 100   # 一轮最多处理多少封新邮件（处理不完的下一轮接着来）
CATCHUP_DAYS = 7           # 最多往回补多少天（兜住「上次记住的那封被删了/移走了」）
MAX_FAILED_ROUNDS = 3      # 同一批连续失败几轮后放弃推进，避免无限重试烧额度
OUTLOOK_RETRY_DELAY = 120  # Outlook 暂时用不了时的退避，别每 30 秒去撞一次
REPORT_TAG_DAYS = 30       # 往前找多少天内的"漏了 Sum 标签"的总结报告


# ── 状态文件 ──────────────────────────────────────────────────────
def state_path() -> str:
    return os.path.join(config.app_dir(), STATE_NAME)


def load_state() -> Dict[str, Any]:
    try:
        with open(state_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_state(state: Dict[str, Any]) -> None:
    try:
        with open(state_path(), "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def clamp_interval(value) -> int:
    try:
        seconds = int(value)
    except (TypeError, ValueError):
        return DEFAULT_INTERVAL
    return max(MIN_INTERVAL, min(MAX_INTERVAL, seconds))


# ── 状态机（纯函数，可单测）───────────────────────────────────────
def plan_tick(last_seen: Optional[str],
              newest: Optional[Dict[str, Any]]) -> Tuple[str, Optional[str]]:
    """决定这一轮该做什么。返回 (action, 新的 last_seen)。

    action:
      idle        没有新邮件
      baseline    第一次运行：只记基线，不碰收件箱里的存量邮件
      summarize   新邮件是 sum 指令
      tag-report  新邮件是我们自己的总结邮件（只打 Sum 标签，不再分类）
      classify    其它新邮件 → 走分级

    注意：classify 分支的 next_seen 只是「最新那封」，
    实际推进到哪儿由调用方按「本轮真正处理完的最旧那封」决定（见 _work）。
    """
    if newest is None:                       # 收件箱是空的
        return "idle", last_seen

    entry_id = newest.get("entryid")
    if last_seen is None:
        return "baseline", entry_id
    if not entry_id or entry_id == last_seen:
        return "idle", last_seen

    subject = newest.get("subject") or ""
    if outlook.parse_sum_command(subject) is not None:
        return "summarize", entry_id
    if outlook.is_report_subject(subject):
        return "tag-report", entry_id
    return "classify", entry_id


def advance_policy(pending, failed: int, failed_rounds: int, newest_entry_id: Optional[str],
                   max_failed_rounds: int = MAX_FAILED_ROUNDS
                   ) -> Tuple[Optional[str], int, bool]:
    """决定这一轮之后基线推进到哪儿。返回 (新基线, 新的连续失败轮数, 是否已放弃)。

      · 有失败且还没到放弃阈值 → 基线**不动**（None）：下一轮重试，
        已经贴上标签的会被 collect_new_mail 跳过，所以只重试失败的那些；
      · 放弃了 → 推进到「本轮处理过的最旧那封」，并记一条日志由调用方打出；
      · 全部成功 → 同样推进到最旧那封：一轮没处理完的（超过 MAX_MAIL_PER_ROUND）
        下一轮会自然接着来，不会漏；
      · 没有待处理邮件（新到的全是控制邮件）→ 推进到最新那封。

    抽成纯函数是因为这是最容易写错、又最难在真实邮箱上复现的一段逻辑。
    """
    if not pending:
        return newest_entry_id, failed_rounds, False
    oldest = pending[-1].get("entryid")
    if failed:
        rounds = failed_rounds + 1
        if rounds < max_failed_rounds:
            return None, rounds, False
        return oldest, 0, True
    return oldest, 0, False


def retry_delay_for(ex: BaseException) -> Optional[int]:
    """这次失败之后隔多久再试；返回 None 表示用默认间隔。

    Outlook 暂时用不了（没开/启动失败）时退避久一点 —— 目的是**别每 30 秒就去撞一次**，
    把 Outlook 自己的错误框刷屏。注意这个退避只影响"多久试一次"，
    绝不能做成"在某个时间内不许连"：那会把「用户已经手动打开 Outlook」也一起挡掉。
    """
    return OUTLOOK_RETRY_DELAY if isinstance(ex, outlook.OutlookNotRunning) else None


# ── 监听控制器 ────────────────────────────────────────────────────
class WatchController:
    """自动监听。所有界面回调都在主线程执行（靠 TaskRunner 的队列机制）。"""

    def __init__(self, widget: tk.Misc, runner, get_cfg: Callable[[], config.Config],
                 log: Callable[..., None], notify: Callable[[str, Any], None]):
        self._widget = widget
        self._runner = runner
        self._get_cfg = get_cfg
        self._log = log
        self._notify = notify

        state = load_state()
        self._last_seen: Optional[str] = state.get("last_entryid")
        self._interval = clamp_interval(state.get("interval", DEFAULT_INTERVAL))
        self._running = False
        self._after_id: Optional[str] = None
        self._errors = 0
        self._failed_rounds = 0
        self._outlook_warned = False
        self._catchup_pending = False       # 本次会话是否还要做「启动补分类」

    # ── 状态 ──────────────────────────────────────────────────
    @property
    def running(self) -> bool:
        return self._running

    @property
    def interval(self) -> int:
        return self._interval

    # ── 开关 ──────────────────────────────────────────────────
    def start(self, delay: int = FIRST_TICK_DELAY) -> None:
        if self._running:
            return
        self._running = True
        self._errors = 0
        self._outlook_warned = False
        self._catchup_pending = True        # 每次启动补分类一次（受 catchup_days 控制）
        self._notify("running", True)
        self._log(f"[watch] watching the inbox every {self._interval}s "
                  f"(new mail gets classified automatically)")
        days = self._get_cfg().catchup_days
        if days > 0:
            self._log(f"[watch] startup catch-up: unclassified mail from the last "
                      f"{days} day(s) will be classified")
        else:
            self._log("[watch] startup catch-up disabled (catchup_days = 0)")
        self._schedule(delay)

    def stop(self) -> None:
        was_running = self._running
        self._running = False
        self._cancel()
        self._notify("busy", False)
        self._notify("running", False)
        if was_running:
            self._log("[watch] stopped")

    def set_interval(self, seconds) -> None:
        self._interval = clamp_interval(seconds)
        self._persist()
        if self._running:
            self._schedule()

    def resume_schedule(self) -> None:
        """把一个被放弃的任务留下的断链接回去。

        正常路径下每轮任务结束都会通过 on_done/on_error 排下一轮；
        但如果任务卡死被 Force reset 放弃，就永远不会有回调 —— 监听会静默死掉。
        """
        if self._running:
            self._schedule(BUSY_RETRY_DELAY)

    # ── 调度 ──────────────────────────────────────────────────
    def _schedule(self, delay: Optional[int] = None) -> None:
        self._cancel()
        if not self._running:
            return
        seconds = self._interval if delay is None else delay
        try:
            self._after_id = self._widget.after(seconds * 1000, self._tick)
        except Exception:
            self._running = False

    def _cancel(self) -> None:
        if self._after_id is not None:
            try:
                self._widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _persist(self) -> None:
        save_state({"last_entryid": self._last_seen, "interval": self._interval})

    # ── 主线程：发起一轮检查 ──────────────────────────────────
    def _tick(self) -> None:
        if not self._running:
            return
        if self._runner.busy:
            self._schedule(BUSY_RETRY_DELAY)
            return

        self._notify("busy", True)
        started = self._runner.run(self._work, on_done=self._on_done, on_error=self._on_error)
        if not started:
            self._notify("busy", False)
            self._schedule(BUSY_RETRY_DELAY)

    # ── 后台线程：真正干活 ────────────────────────────────────
    def _work(self) -> Dict[str, Any]:
        cfg = self._get_cfg()
        _app, ns = outlook.connect(log=self._log)

        # 先给"还没打 Sum 标签的总结报告"补标签。
        # 不能只靠下面 plan_tick 的 tag-report 分支 —— 它检查的是"最新那一封"，
        # 只要报告之后又到了一封别的邮件，报告就永远轮不到（实际发生过）。
        self._tag_reports(ns)

        newest = outlook.newest_message(ns)
        info = outlook.message_info(newest) if newest is not None else None
        previous = self._last_seen
        action, next_seen = plan_tick(previous, info)

        if action == "idle":
            # 没有新邮件不等于没事干：启动补分类还没做完的话，这一轮就做它
            if self._catchup_pending and cfg.catchup_days > 0:
                return self._catchup(ns, cfg)
            return {"action": action}

        if action == "baseline":
            self._commit(next_seen)          # 第一次运行：只记基线（存量交给补分类）
            if self._catchup_pending and cfg.catchup_days > 0:
                return self._catchup(ns, cfg)
            return {"action": action}

        if action == "tag-report":
            outlook.ensure_sum_category(ns)
            outlook.write_categories(newest, [outlook.SUM_NAME])
            self._commit(next_seen)
            return {"action": action}

        if action == "summarize":
            # 已经处理过的指令邮件（打过 Sum 就是"处理完成"的标志）：只推进基线，
            # 不要再出一份报告 —— 否则"补扫指令"那套逻辑会和这里互相触发。
            if outlook.SUM_NAME in outlook.category_names(newest):
                self._commit(next_seen)
                return {"action": "skip-control"}

            letters = outlook.parse_sum_command(newest.Subject) or "sabc"
            outlook.ensure_sum_category(ns)
            outlook.write_categories(newest, [outlook.SUM_NAME])
            self._log(f'[watch] "sum" command received (levels: {letters.upper()})')
            count = summarizer.summarize_inbox(cfg, letters, log=self._log)
            extra = self._run_commands(ns, cfg)      # 同一批里还有别的指令就一起办
            self._commit(next_seen)
            return {"action": action, "letters": letters, "count": count,
                    "extra": len(extra)}

        # classify：只处理「比上次记住的那封更新」的邮件。
        # 关键：不要用 classify_inbox(days=None) —— 那会遍历整个收件箱并逐封读正文，
        # 大邮箱下要几分钟，期间什么都没分类、按钮还一直灰着。
        # 存量未分类邮件由 _catchup() 负责（见下），不混进这条路径。

        # 全新用户的邮箱里没有 S/A/B/C 分类，必须先确保它们存在且颜色正确：
        # 否则 Outlook 会把 Categories 里出现的陌生名字自动加成「无颜色」分类，
        # 四档颜色编码就全丢了。
        for change in outlook.ensure_categories(ns):
            self._log(f"[watch] {change}")

        pending = outlook.collect_new_mail(ns, previous, days=CATCHUP_DAYS,
                                           max_items=MAX_MAIL_PER_ROUND)
        # 埋在这批新邮件下面的 sum 指令：顺手办掉，别让它们被基线推过去丢掉
        commands = self._run_commands(ns, cfg)
        if not pending:
            self._commit(next_seen)          # 新到的都是控制邮件，没有要分级的东西
            return {"action": "idle", "commands": len(commands)}

        self._log(f"[watch] {len(pending)} new mail(s) to classify")
        verdicts = classifier.classify(pending, cfg, log=self._log)
        ok, failed = classifier.apply_verdicts(pending, verdicts, log=self._log)
        stats = {"read": len(pending), "ok": ok, "failed": failed}
        stats.update(classifier.counts(verdicts))

        baseline, self._failed_rounds, gave_up = advance_policy(
            pending, failed, self._failed_rounds, next_seen)
        if baseline is None:
            self._log(f"[watch] {failed} mail(s) failed — retrying next round "
                      f"({self._failed_rounds}/{MAX_FAILED_ROUNDS})")
            return {"action": action, "stats": stats, "retry": True}
        if gave_up:
            self._log(f"[watch] giving up on {failed} mail(s) after {MAX_FAILED_ROUNDS} rounds; "
                      f"run Categorize manually if you still need them")
        self._commit(baseline)
        return {"action": action, "stats": stats}

    def _tag_reports(self, ns) -> int:
        """把「是总结报告但还没打 Sum 标签」的邮件补上标签，返回补了几封。

        每一轮都跑（Restrict 查询约 0.03 秒），所以历史遗留的漏标签也会被补上。
        已打过标签的会被跳过，重复执行无害。
        """
        try:
            reports = outlook.untagged_reports(ns, days=REPORT_TAG_DAYS)
        except Exception as ex:
            self._log(f"[watch] could not scan for untagged reports: {type(ex).__name__}: {ex}")
            return 0
        if not reports:
            return 0

        outlook.ensure_sum_category(ns)
        tagged = 0
        for msg in reports:
            try:
                outlook.write_categories(msg, [outlook.SUM_NAME])
                tagged += 1
            except Exception as ex:
                self._log(f"[watch] could not tag a report: {type(ex).__name__}: {ex}")
        if tagged:
            self._log(f"[watch] tagged {tagged} summary report(s) with Sum")
        return tagged

    def _run_commands(self, ns, cfg) -> List[Any]:
        """执行「还没处理过」的 sum 指令，返回 [(letters, count), ...]。

        为什么需要：`plan_tick` 只看最新那一封。指令邮件到达后 30 秒内若又来了别的邮件，
        它就不再是"最新"，而 classify 路径又会把指令邮件当控制邮件跳过 ——
        于是这条指令永远轮不到执行，基线一推进就彻底丢了。

        处理顺序：**先打 Sum 标签再出报告**。标签就是"已处理"的标志，
        所以即使总结抛异常，也不会被反复重跑（异常会走 _on_error 的错误计数）。
        """
        try:
            commands = outlook.pending_sum_commands(ns, self._last_seen, days=CATCHUP_DAYS)
        except Exception as ex:
            self._log(f"[watch] could not scan for sum commands: {type(ex).__name__}: {ex}")
            return []

        done: List[Any] = []
        for msg in commands:
            letters = outlook.parse_sum_command(msg.Subject) or "sabc"
            self._log(f'[watch] "sum" command found (levels: {letters.upper()})')
            outlook.ensure_sum_category(ns)
            outlook.write_categories(msg, [outlook.SUM_NAME])
            count = summarizer.summarize_inbox(cfg, letters, log=self._log)
            done.append((letters, count))
        return done

    def _catchup(self, ns, cfg) -> Dict[str, Any]:
        """启动补分类：把最近 catchup_days 天里**还没打过标签**的邮件处理掉。

        为什么需要（用户实际反馈）：监听只处理「比上次记住的那封更新」的邮件，
        所以收件箱里早就堆着、从来没分类过的邮件永远不会被碰 —— 表现就是
        「打开了程序，但那些未分类邮件没人管」。

        范围由 cfg.catchup_days 控制（0 = 关闭）。一轮最多 MAX_MAIL_PER_ROUND 封，
        处理满一轮就继续下一轮，直到范围内没有未分类邮件为止。
        `read_inbox(skip_classified=True)` 天然会跳过已打标签的、我们自己的总结邮件
        和 sum 指令邮件，所以重复触发是安全的。
        """
        for change in outlook.ensure_categories(ns):
            self._log(f"[watch] {change}")

        pending = outlook.read_inbox(ns, days=cfg.catchup_days,
                                     skip_classified=True, limit=MAX_MAIL_PER_ROUND)
        if not pending:
            self._catchup_pending = False
            return {"action": "catchup", "stats": {"read": 0, "ok": 0, "failed": 0}, "more": False}

        self._log(f"[watch] catch-up: {len(pending)} unclassified mail(s) in the last "
                  f"{cfg.catchup_days} day(s)")
        verdicts = classifier.classify(pending, cfg, log=self._log)
        ok, failed = classifier.apply_verdicts(pending, verdicts, log=self._log)
        stats = {"read": len(pending), "ok": ok, "failed": failed}
        stats.update(classifier.counts(verdicts))

        # 凑满一批说明范围内可能还有，下一轮接着来；没满就是清完了
        self._catchup_pending = len(pending) >= MAX_MAIL_PER_ROUND
        if failed:
            self._log(f"[watch] catch-up: {failed} mail(s) could not be labelled")
        return {"action": "catchup", "stats": stats, "more": self._catchup_pending}

    def _commit(self, entry_id: Optional[str]) -> None:
        """推进基线并落盘。

        只在「这一轮真的处理完了」之后调用 —— 旧实现在处理前就推进，
        一旦分级失败这封就永远不会被重试了。
        """
        if entry_id and entry_id != self._last_seen:
            self._last_seen = entry_id
            self._persist()

    # ── 主线程：结果处理 ──────────────────────────────────────
    def _on_done(self, result: Dict[str, Any]) -> None:
        self._errors = 0
        if self._outlook_warned:          # 之前提示过连不上，现在通了
            self._log("[watch] Outlook is available now")
        self._outlook_warned = False      # 下次断了再提示一次
        action = result.get("action")

        if action == "baseline":
            self._log("[watch] baseline recorded — existing mail is left untouched")
        elif action == "classify":
            stats = result["stats"]
            if result.get("retry"):
                self._log(f"[watch] {stats['failed']} mail(s) still unlabelled, will retry shortly")
            else:
                self._log(f"[watch] classified {stats['ok']} mail(s): "
                          f"S={stats['S']} A={stats['A']} B={stats['B']} C={stats['C']}")
        elif action == "catchup":
            stats = result["stats"]
            if not stats["read"]:
                self._log("[watch] catch-up done: no unclassified mail left in scope")
            else:
                self._log(f"[watch] catch-up classified {stats['ok']} mail(s): "
                          f"S={stats['S']} A={stats['A']} B={stats['B']} C={stats['C']}")
                if result.get("more"):
                    self._log("[watch] catch-up: more unclassified mail in range, continuing")
        elif action == "summarize":
            self._log(f"[watch] summary sent ({result.get('count', 0)} mail(s))")
            if result.get("extra"):
                self._log(f"[watch] {result['extra']} more sum command(s) handled in the same pass")
        elif action == "skip-control":
            pass                          # 已处理过的控制邮件：只推进基线，不重复出报告
        elif action == "idle" and result.get("commands"):
            self._log(f"[watch] handled {result['commands']} sum command(s) found in this batch")
        elif action == "tag-report":
            self._log("[watch] tagged our own report email with Sum")

        self._notify("busy", False)
        # 刚处理过邮件就紧接着再确认一轮（可能还有别的没分类）；否则等下一个间隔
        soon = action in ("classify", "summarize") or (
            action == "catchup" and result.get("more"))
        self._schedule(BUSY_RETRY_DELAY if soon else None)

    def _on_error(self, ex: BaseException) -> None:
        self._notify("busy", False)

        # Outlook 用不了（没开、正在启动、或启动失败）：不算失败、不累计、不暂停。
        # 只提示一次，然后继续轮询 —— connect() 每轮都会再试，通了就自动接上。
        if isinstance(ex, outlook.OutlookNotRunning):
            delay = retry_delay_for(ex)
            if not self._outlook_warned:
                self._outlook_warned = True
                self._log(f"[watch] Outlook is not available yet - {ex}")
                self._log(f"[watch] still watching; retrying every {delay}s")
            self._schedule(delay)
            return

        self._errors += 1
        self._log(f"[watch] error ({self._errors}/{MAX_CONSECUTIVE_ERRORS}): "
                  f"{type(ex).__name__}: {ex}")
        if self._errors >= MAX_CONSECUTIVE_ERRORS:
            self._log("[watch] too many failures in a row — auto-paused. "
                      "Check the log above, then turn watching back on.")
            self.stop()
            return
        self._schedule()
