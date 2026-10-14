"""run.py — 命令行入口。

用法（在 Email_manager_beta 目录下执行）：

    py -3 run.py check [--ai]
        自检：配置 / AI 后端 / Outlook 连接。只读，不改任何邮件。
        --ai  额外发一个最小请求，实测 AI 后端是否真的通（会消耗极少量额度）

    py -3 run.py providers
        列出 config.json 里配置的所有 AI 后端，并标明当前生效的是哪个。

    py -3 run.py classify [--limit N] [--days N] [--dry-run] [--force]
        读收件箱 → AI 分级 → 写回 S/A/B/C 分类（S 且加急的额外打红旗）。

    py -3 run.py summarize [letters]
        把最近 N 天已分级的邮件整理成报告发给自己（N 见 config.json 的 summary_days）。

界面入口见 app.py（Tkinter 界面，配置填好后会自动开始监听收件箱）。

退出码：0 成功 ／ 1 配置或环境有问题 ／ 2 有部分邮件处理失败。
"""
from __future__ import annotations

import argparse
import sys
from typing import List, Optional

from core import ai, classifier, config, outlook, summarizer


def _setup_console() -> None:
    """让输出在 Windows 控制台正常显示（隐藏控制台时 stdout 可能是 None）。"""
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def cmd_providers(cfg: config.Config) -> int:
    """列出所有 AI 后端配置。不联网。"""
    print(f"Config file: {config.config_path()}")
    print(f"Active: ai.active = \"{cfg.ai_active}\"\n")
    for name, profile in cfg.ai_profiles.items():
        mark = "*" if name == cfg.ai_active else " "
        print(f" {mark} {name}")
        print(f"     {ai.describe(profile)}")
        for problem in profile.problems():
            print(f"     ! {problem}")
    if cfg.ai.missing:
        print(f"\n! ai.active points at \"{cfg.ai_active}\", which is not in the list above.")
        return 1
    return 0


def cmd_check(cfg: config.Config, test_ai: bool = False) -> int:
    """只读自检：不读信、不发信、不改分类。"""
    print(f"Config file: {config.config_path()}")

    problems = cfg.problems()
    code = 1 if problems else 0
    if problems:
        print("[config] problems found:")
        for problem in problems:
            print(f"  - {problem}")
    else:
        print("[config] OK")

    print(f"  email {cfg.email or '(empty)'} / name {cfg.name or '(empty)'} / "
          f"student id {cfg.student_id or '(empty)'}")
    print(f"  faculty {cfg.faculty or '(empty)'} / school {cfg.school or '(empty)'}")
    print(f"  school domains (decided as school mail without asking the AI): "
          f"{', '.join(cfg.school_domains) or '(empty)'}")
    print(f"  enrolled courses: {', '.join(cfg.courses) or '(empty)'}")
    print(f"  summary recipient {cfg.recipient or '(empty)'} / covering the last "
          f"{cfg.summary_days} day(s)")
    print(f"  batch size {cfg.batch_size} / urgent threshold {cfg.urgent_days} day(s)")
    print(f"  startup catch-up: {'last %d day(s)' % cfg.catchup_days if cfg.catchup_days else 'disabled'}")

    # ── AI backend ────────────────────────────────────────────
    print(f"[AI] active backend: \"{cfg.ai_active}\"")
    print(f"     {ai.describe(cfg.ai)}")
    others = [name for name in cfg.ai_profiles if name != cfg.ai_active]
    if others:
        print(f"     switchable: {', '.join(others)} "
              f"(change ai.active in config.json, or run: py -3 run.py providers)")

    if test_ai:
        if problems:
            print("[AI] configuration has problems, skipping the live test.")
        else:
            print("[AI] sending a minimal test request ...")
            try:
                reply = ai.probe(cfg.ai, log=print)
                print(f"[AI] connection OK, model replied: {reply[:60]!r}")
            except ai.AiError as ex:
                print(f"[AI] connection test failed: {ex}")
                code = 1

    # ── Outlook ───────────────────────────────────────────────
    try:
        _app, ns = outlook.connect(auto_launch=False)   # 只读自检不该顺手启动 Outlook
        inbox = ns.GetDefaultFolder(outlook.OL_FOLDER_INBOX)
        print(f"[Outlook] connected; inbox \"{inbox.Name}\" holds {inbox.Items.Count} mail(s)")
        names = outlook.existing_categories(ns)
        missing = [level for level in config.LEVELS if level not in names]
        if outlook.SUM_NAME not in names:
            missing.append(outlook.SUM_NAME)
        print(f"[Outlook] existing categories: {', '.join(names) or '(none)'}")
        print(f"[Outlook] missing categories: {', '.join(missing) or '(none, S/A/B/C/Sum all present)'}")
        if missing:
            print("          (classify creates missing categories automatically)")

        colors = outlook.our_category_colors(ns)
        if colors:
            expected = outlook.expected_category_colors()
            wrong = {name: value for name, value in colors.items()
                     if expected.get(name) != value}
            if wrong:
                print(f"[Outlook] colour mismatch: {wrong} "
                      f"(expected {expected}) - fixed automatically on the next classify")
            else:
                print("[Outlook] category colours: all correct")
    except Exception as ex:
        print(f"[Outlook] connection failed: {ex}")
        print("          Make sure Outlook classic (desktop) is installed and has been "
              "started at least once; the new Outlook does not expose COM.")
        code = 1

    return code


def main(argv: Optional[List[str]] = None) -> int:
    _setup_console()

    parser = argparse.ArgumentParser(
        prog="run.py",
        description="AI email manager for Outlook classic (core, command line)",
        epilog="Examples: py -3 run.py check --ai | py -3 run.py classify --days 7 --dry-run | "
               "py -3 run.py summarize sa",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_check = sub.add_parser("check", help="self-check: config / AI backend / Outlook (read-only)")
    p_check.add_argument("--ai", action="store_true",
                         help="also send a minimal request to verify the AI backend really works "
                              "(costs a tiny amount of credit)")

    sub.add_parser("providers", help="list all configured AI backends and which one is active")

    p_classify = sub.add_parser("classify", help="read the inbox, classify with AI, write S/A/B/C back")
    p_classify.add_argument("--limit", type=int, default=None, help="process at most N mail(s)")
    p_classify.add_argument("--days", type=int, default=None, help="only process the last N day(s)")
    p_classify.add_argument("--dry-run", action="store_true", help="print results only, no write-back")
    p_classify.add_argument("--force", action="store_true", help="also reclassify already labelled mail")

    p_summarize = sub.add_parser("summarize", help="build a report of recently classified mail and send it to yourself")
    p_summarize.add_argument("letters", nargs="?", default="sabc",
                             help="levels to include, e.g. s / sa / sabc (default: sabc)")

    args = parser.parse_args(argv)
    cfg = config.ensure()

    if args.cmd == "providers":
        return cmd_providers(cfg)
    if args.cmd == "check":
        return cmd_check(cfg, test_ai=args.ai)

    problems = cfg.problems()
    if problems:
        print("[config] problems found, aborting:")
        for problem in problems:
            print(f"  - {problem}")
        print(f"Edit {config.config_path()} and retry (py -3 run.py check re-checks).")
        return 1

    if args.cmd == "classify":
        if args.days is not None:
            scope = f"the last {args.days} day(s)"
        elif args.limit is not None:
            scope = f"at most {args.limit} mail(s)"
        else:
            scope = "all unclassified mail"
        print(f"Connecting to Outlook, reading {scope}"
              + (" (dry run, no write-back)" if args.dry_run else "") + " ...")
        print(f"AI backend: {ai.describe(cfg.ai)}")

        stats = classifier.classify_inbox(cfg, limit=args.limit, days=args.days,
                                          dry_run=args.dry_run, force=args.force, log=print)
        print(f"\nDone: read {stats['read']}; "
              f"S={stats['S']} A={stats['A']} B={stats['B']} C={stats['C']}; "
              f"ok {stats['ok']}, failed {stats['failed']}, "
              f"skipped {stats.get('skipped', 0)} (not school mail)")
        return 2 if stats["failed"] else 0

    count = summarizer.summarize_inbox(cfg, args.letters, log=print)
    print(f"\nDone: summarised {count} mail(s)" if count
          else "\nDone: nothing to summarise")
    return 0


if __name__ == "__main__":
    sys.exit(main())
