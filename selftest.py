"""selftest.py — 纯逻辑自测：不连接 Outlook、不调用 AI、不消耗任何额度。

用法（在 Email_manager_beta 目录下）：
    py -3 selftest.py

覆盖 config / ai.parse_json / classifier / summarizer 里所有不依赖 COM 的逻辑。
旧版没有任何测试，规则改错只能等线上邮件暴露；这里是这套内核的第一道防线。
"""
import os
from datetime import date, datetime, timedelta

from core import ai, classifier, config, outlook, summarizer

ok = 0


def check(name, cond):
    global ok
    assert cond, f"FAIL: {name}"
    ok += 1
    print(f"  ok  {name}")


print("[1] config")
cfg = config.from_dict({
    "email": "  a@b.edu ",
    "courses": "COMP1021, MATH1013；LANG1002、COMP1021",
    "summary_days": "9999",
    "batch_size": 0,
    "urgent_days": "abc",
    "language": "EN",
    "faculty": "School of Engineering",
})
check("email 去空白", cfg.email == "a@b.edu")
check("courses 多种分隔符 + 去重保序", cfg.courses == ["COMP1021", "MATH1013", "LANG1002"])
check("summary_days 被夹到上限 365", cfg.summary_days == 365)
check("batch_size 非法 0 → 下限 1", cfg.batch_size == 1)
check("urgent_days 非数字 → 默认 30", cfg.urgent_days == 30)
check("language EN → en / is_english", cfg.language == "en" and cfg.is_english)
check("recipient 回落本人邮箱", config.from_dict({"email": "x@y.z"}).recipient == "x@y.z")
check("problems：缺 email 报错", bool(config.from_dict({}).problems()))
check("problems：完整合法配置无报错",
      not config.from_dict({"email": "a@b.edu",
                            "ai": {"active": "c", "profiles": {"c": {"type": "cli"}}}}).problems())
check("problems：坏正则被抓出",
      any("regex" in p for p in config.from_dict({"email": "a@b", "course_code_regex": "([A-"}).problems()))
check("to_dict/from_dict 往返", config.from_dict(config.to_dict(cfg)).courses == cfg.courses)
check("缺 school 键时补港科大默认值",
      config.from_dict({}).school == "香港科技大学 (HKUST)")
check("显式清空 school 时尊重用户（不回落默认）",
      config.from_dict({"school": ""}).school == "")
check("缺 dorm_keywords 时补港科大默认值",
      "UGHousing" in config.from_dict({}).dorm_keywords)
check("显式清空 dorm_keywords 时尊重用户",
      config.from_dict({"dorm_keywords": []}).dorm_keywords == [])
check("缺 language 键时用默认值 en", config.from_dict({}).language == "en")
check("显式 language=zh 生效", config.from_dict({"language": "zh"}).language == "zh")
check("无法识别的 language 回落默认值", config.from_dict({"language": "fr"}).language == "en")
check("缺 ai.active 时用默认值 deepseek", config.from_dict({}).ai_active == "deepseek")
check("DeepSeek 模板用官方模型名 deepseek-flash",
      config.from_dict({}).ai_profiles["deepseek"].model == "deepseek-flash")
check("DeepSeek 模板用官方 base_url（不带 /v1）",
      config.from_dict({}).ai_profiles["deepseek"].base_url == "https://api.deepseek.com")
check("拼出的请求地址与官方 curl 示例一致",
      ai._endpoint(config.from_dict({}).ai_profiles["deepseek"].base_url, "/chat/completions")
      == "https://api.deepseek.com/chat/completions")

print("[2] ai.parse_json")
check("剥 ```json 围栏", ai.parse_json('```json\n[{"级别":"S"}]\n```') == [{"级别": "S"}])
check("容忍前后废话", ai.parse_json('好的，结果如下：\n[1,2]\n以上。') == [1, 2])
check("纯 JSON 对象", ai.parse_json('{"a": 1}') == {"a": 1})
check("垃圾输入 → None", ai.parse_json("完全不是 JSON") is None)

print("[3] classifier")
spec = classifier.build_spec(cfg)
check("spec 含课程", "COMP1021" in spec)
check("spec 含学院", "School of Engineering" in spec)
check("spec 含宿舍关键词（港科大默认值）", "UGHousing" in spec)
check("spec 含学校（港科大默认值）", "香港科技大学" in spec)
check("spec 含加急天数", "30 天内 deadline" in spec)

check("normalize_level 合法", [classifier.normalize_level(x) for x in ("s", "A", " b ", "C")] == list("SABC"))
check("normalize_level 非法 → C", classifier.normalize_level("X") == "C"
      and classifier.normalize_level(None) == "C")
check("parse_deadline 正常", classifier.parse_deadline("2026-03-05") == date(2026, 3, 5))
check("parse_deadline ISO 串", classifier.parse_deadline("2026-03-05T10:00:00Z") == date(2026, 3, 5))
check("parse_deadline 空/非法 → None",
      classifier.parse_deadline("") is None and classifier.parse_deadline("下周一") is None)

soon = (date.today() + timedelta(days=3)).isoformat()
far = (date.today() + timedelta(days=90)).isoformat()
v = classifier.to_verdicts([{"级别": "S", "截止日期": soon},
                            {"级别": "a", "截止日期": None},
                            {"级别": "B", "截止日期": far},
                            {"级别": "??", "截止日期": "bad"}], 4, cfg)
check("4 条正常解析", v is not None and len(v) == 4)
check("A→S 加急判定", v[0].level == "S" and v[0].urgent)
check("小写 a 归一为 A 且不加急", v[1].level == "A" and not v[1].urgent)
check("B 级即使有远处 deadline 也不加急", v[2].level == "B" and not v[2].urgent)
check("非法级别 → C", v[3].level == "C")

check("条数对不上 → None（防错位）", classifier.to_verdicts([{"级别": "S"}], 3, cfg) is None)
check("非数组 → None", classifier.to_verdicts({"级别": "S"}, 1, cfg) is None)
check("元素非 dict → None", classifier.to_verdicts(["S"], 1, cfg) is None)

emails = [
    {"subject": "COMP1021 Homework 3", "body": "due soon"},
    {"subject": "Welcome", "body": "nothing special"},
]
verdicts = classifier.to_verdicts([{"级别": "C", "截止日期": None},
                                   {"级别": "B", "截止日期": None}], 2, cfg)
check("确定性规则改写封数 = 1", classifier.apply_enrolled_override(emails, verdicts, cfg) == 1)
check("课程邮件被强制 S", verdicts[0].level == "S")
check("无关邮件不受影响", verdicts[1].level == "B")
check("counts 统计", classifier.counts(verdicts) == {"S": 1, "A": 0, "B": 1, "C": 0})
check("counts 容忍 None", classifier.counts([None, verdicts[0]])["S"] == 1)
check("无课程时不误伤", classifier.apply_enrolled_override(
    emails, verdicts, config.from_dict({"email": "a@b.c"})) == 0)

print("[4] summarizer")
check("parse_letters 归一排序", summarizer.parse_letters("BA") == "ab"
      and summarizer.parse_letters("s,a") == "sa")
check("parse_letters 空 → sabc", summarizer.parse_letters("") == "sabc"
      and summarizer.parse_letters("zzz") == "sabc")
check("label_for 全部", summarizer.label_for("sabc") == "全部")
check("label_for 组合", summarizer.label_for("sa") == "S+A 级")
check("label_for 英文", summarizer.label_for("sa", True) == "S+A levels")

samples = [
    {"entryid": "AAA111", "level": "S", "subject": "作业 <重要> & 紧急",
     "sender": "Prof. A", "body": "b", "received": datetime(2026, 3, 1, 9, 30)},
    {"entryid": "BBB222", "level": "C", "subject": "Newsletter",
     "sender": "Shop", "body": "b", "received": datetime(2026, 3, 2, 8, 0)},
]
html = summarizer.build_html(samples, ["要点一", ""], cfg, "sc")
check("HTML 含可点开链接", "outlook:AAA111" in html and "outlook:BBB222" in html)
check("HTML 转义主题尖括号", "&lt;重要&gt;" in html and "&amp;" in html)
check("HTML 分级标题", "【S】" in html and "【C】" in html)
check("HTML 块数正确（2 封各一个 div 块）", html.count("border-bottom:1px solid #e6e6e6") == 2)
check("空要点不产生空行", html.count("要点一") == 1)
check("div 标签配平", html.count("<div") == html.count("</div>"))
check("HTML 为空列表时也不崩",
      summarizer.build_html([], [], config.from_dict({"email": "a@b.c"}), "sabc").startswith("<html>"))
check("主题含识别标记（跟随 language，见 [11] 节的中英双标记测试）",
      outlook.is_report_subject(summarizer.build_subject("sa", cfg)))

print("[5] ai 后端")
check("endpoint 正常拼接",
      ai._endpoint("https://a.com/v1", "/chat/completions") == "https://a.com/v1/chat/completions")
check("endpoint 容忍多余斜杠",
      ai._endpoint("https://a.com/v1/", "/chat/completions") == "https://a.com/v1/chat/completions")
check("endpoint 已写全路径不重复",
      ai._endpoint("https://a.com/v1/chat/completions", "/chat/completions")
      == "https://a.com/v1/chat/completions")
check("endpoint 带查询串原样使用（Azure 式）",
      ai._endpoint("https://a.com/x/chat/completions?api-version=1", "/chat/completions")
      == "https://a.com/x/chat/completions?api-version=1")
check("endpoint 空 base_url 回落路径", ai._endpoint("", "/v1/messages") == "/v1/messages")

check("anthropic base 原样保留", ai._anthropic_base("https://api.anthropic.com") == "https://api.anthropic.com")
check("anthropic base 容忍用户多写的 /v1",
      ai._anthropic_base("https://api.anthropic.com/v1") == "https://api.anthropic.com")
check("anthropic base 保留 DeepSeek 的 /anthropic 前缀",
      ai._anthropic_base("https://api.deepseek.com/anthropic") == "https://api.deepseek.com/anthropic")
check("DeepSeek 走 Anthropic 格式时地址正确",
      ai._endpoint(ai._anthropic_base("https://api.deepseek.com/anthropic"), "/v1/messages")
      == "https://api.deepseek.com/anthropic/v1/messages")
check("anthropic base 为空时回落官方地址",
      ai._anthropic_base("") == "https://api.anthropic.com")

check("content 为字符串", ai._flatten_content("hi") == "hi")
check("content 为分块列表", ai._flatten_content([{"text": "a"}, {"text": "b"}]) == "ab")
check("content 为其他类型 → 空串", ai._flatten_content(None) == "")
check("错误信息优先取 error.message",
      ai._error_snippet('{"error":{"message":"bad key"}}') == "bad key")
check("错误信息非 JSON 时截断原文", ai._error_snippet("x" * 500) == "x" * 300)
check("错误信息空输入", ai._error_snippet("") == "")

_prof = config.from_dict({"ai": {"active": "p", "profiles": {
    "p": {"type": "openai", "base_url": "u", "model": "m", "api_key": "k"}}}}).ai
check("2xx 不抛错", ai._raise_for_status(200, "", _prof, "url") is None)
for status, expect_transient, keyword in ((401, False, "API key"), (403, False, "API key"),
                                          (404, False, "base_url"), (429, True, "rate limited"),
                                          (400, False, "request rejected"), (500, True, "server error")):
    try:
        ai._raise_for_status(status, '{"error":{"message":"detail"}}', _prof, "url")
        check(f"HTTP {status} 应当抛错", False)
    except ai.AiError as ex:
        check(f"HTTP {status} → {keyword}", keyword in str(ex) and ex.transient is expect_transient)

os.environ["EMAIL_MANAGER_SELFTEST_KEY"] = "from-env"
_prof2 = config.from_dict({"ai": {"active": "p", "profiles": {"p": {
    "type": "openai", "base_url": "u", "model": "m",
    "api_key": "plain", "api_key_env": "EMAIL_MANAGER_SELFTEST_KEY"}}}}).ai
check("环境变量优先于明文 api_key", _prof2.key() == "from-env")
del os.environ["EMAIL_MANAGER_SELFTEST_KEY"]
check("环境变量缺失时回落到明文", _prof2.key() == "plain")
check("Key 来源标注正确", _prof2.key_source() == "config.json api_key")

print("[6] AI 配置校验")
check("api_key_required=false 时不再报 Key 缺失",
      not any("API key" in p for p in config.from_dict({"ai": {"active": "local", "profiles": {"local": {
          "type": "openai", "base_url": "http://x/v1", "model": "m",
          "api_key_required": False}}}}).ai_problems()))
check("api_key_required=true（默认）时仍报 Key 缺失",
      any("API key" in p for p in config.from_dict({"ai": {"active": "local", "profiles": {"local": {
          "type": "openai", "base_url": "http://x/v1", "model": "m"}}}}).ai_problems()))
check("需要 Key 但没填时确实报错",
      any("API key" in p for p in config.from_dict({"ai": {"active": "x", "profiles": {"x": {
          "type": "openai", "base_url": "http://x/v1", "model": "m"}}}}).problems()))
check("缺 base_url / model 会被抓出",
      len(config.from_dict({"ai": {"active": "x", "profiles": {"x": {
          "type": "openai", "api_key": "k"}}}}).ai_problems()) == 2)
check("ai.active 指向不存在的后端会被抓出",
      any("does not exist" in p for p in config.from_dict({"ai": {"active": "no-such", "profiles": {
          "x": {"type": "cli"}}}}).ai_problems()))
check("cli 后端不需要 Key / url / model",
      not config.from_dict({"ai": {"active": "c", "profiles": {
          "c": {"type": "cli"}}}}).ai_problems())
check("不认识的 type 会被抓出",
      any("unknown type" in p for p in config.from_dict({"ai": {"active": "x", "profiles": {
          "x": {"type": "gemini"}}}}).ai_problems()))
check("identity_problems 不掺 AI 的问题",
      all("AI backend" not in p for p in config.from_dict({}).identity_problems()))

print("[7] User info 编辑往返（界面保存走的就是这条路）")
_edited = config.from_dict({**config.to_dict(cfg), "school": "另一所大学",
                            "courses": "COMP1021, MATH1013", "summary_days": "3"})
check("改 User info 不丢 AI 后端列表", list(_edited.ai_profiles) == list(cfg.ai_profiles))
check("改 User info 不丢 ai.active", _edited.ai_active == cfg.ai_active)
check("改 User info 后课程被重新解析", _edited.courses == ["COMP1021", "MATH1013"])
check("改 User info 后天数被重新归一", _edited.summary_days == 3)
check("往返不改变身份字段", _edited.email == cfg.email and _edited.faculty == cfg.faculty)

print("[8] 自动监听的状态机（plan_tick / advance_policy，纯函数）")
from gui.watcher import (DEFAULT_INTERVAL, MAX_INTERVAL, MIN_INTERVAL,  # noqa: E402
                         OUTLOOK_RETRY_DELAY, advance_policy, clamp_interval, plan_tick,
                         retry_delay_for)

_seen = "ID-A"
check("收件箱空了 → idle", plan_tick(_seen, None) == ("idle", _seen))
check("首次运行（无基线）→ baseline",
      plan_tick(None, {"entryid": "ID-A", "subject": "Hi"}) == ("baseline", "ID-A"))
check("还是同一封 → idle", plan_tick(_seen, {"entryid": "ID-A", "subject": "Hi"}) == ("idle", _seen))
check("来了新邮件 → classify",
      plan_tick(_seen, {"entryid": "ID-B", "subject": "Hello"}) == ("classify", "ID-B"))
check("新邮件是 sum 指令 → summarize",
      plan_tick(_seen, {"entryid": "ID-B", "subject": "sum sa"}) == ("summarize", "ID-B"))
check("新邮件是自己的报告 → tag-report",
      plan_tick(_seen, {"entryid": "ID-B", "subject": "📬 邮件总结（全部）· 最近7天"})
      == ("tag-report", "ID-B"))
check("sum 指令优先于报告判定",
      plan_tick(_seen, {"entryid": "ID-B", "subject": "sum"})[0] == "summarize")
check("entryid 缺失时不误判为新邮件",
      plan_tick(_seen, {"entryid": None, "subject": "x"}) == ("idle", _seen))
check("基线只在首次运行时建立（第二次不会重记）",
      plan_tick("ID-A", {"entryid": "ID-B", "subject": "x"})[0] != "baseline")

check("间隔夹取下限", clamp_interval(0) == MIN_INTERVAL)
check("间隔夹取上限", clamp_interval(99999) == MAX_INTERVAL)
check("间隔非数字 → 默认值", clamp_interval("abc") == DEFAULT_INTERVAL)
check("间隔正常值原样返回", clamp_interval(45) == 45)

# ── 失败后的重试节奏（退避，而不是"一段时间内禁止连接"）────────
check("Outlook 暂时用不了 → 退避到 120 秒",
      retry_delay_for(outlook.OutlookNotRunning("nope")) == OUTLOOK_RETRY_DELAY)
check("其它错误 → 用默认间隔（None）", retry_delay_for(ValueError("boom")) is None)
check("退避只影响节奏，不做成'禁止连接'",
      not hasattr(outlook, "_last_launch_attempt"))

# ── 启动补分类的范围设置 ──────────────────────────────────────
check("catchup_days 默认 7 天", config.from_dict({}).catchup_days == 7)
check("catchup_days=0 是合法的（0 = 关闭补分类）",
      config.from_dict({"catchup_days": 0}).catchup_days == 0)
check("catchup_days 支持自定义", config.from_dict({"catchup_days": 30}).catchup_days == 30)
check("catchup_days 负数被夹到 0", config.from_dict({"catchup_days": -5}).catchup_days == 0)
check("catchup_days 超大值被夹到 365", config.from_dict({"catchup_days": 9999}).catchup_days == 365)
check("catchup_days 会写回 to_dict（保存配置不丢）",
      config.to_dict(config.from_dict({"catchup_days": 14}))["catchup_days"] == 14)

# ── 基线推进策略（监听里最容易写错的一段）─────────────────────
_pending = [{"entryid": "ID-3"}, {"entryid": "ID-2"}, {"entryid": "ID-1"}]   # 新 → 旧
check("全部成功 → 推进到最旧那封（一轮没处理完的下一轮接着来）",
      advance_policy(_pending, 0, 0, "ID-3") == ("ID-1", 0, False))
check("没有待处理邮件 → 推进到最新那封",
      advance_policy([], 0, 0, "ID-3") == ("ID-3", 0, False))
check("首轮有失败 → 不推进基线（None）并记一次失败",
      advance_policy(_pending, 2, 0, "ID-3") == (None, 1, False))
check("第二轮仍失败 → 继续不推进",
      advance_policy(_pending, 2, 1, "ID-3") == (None, 2, False))
check("连续失败到阈值 → 放弃并推进，避免无限重试烧额度",
      advance_policy(_pending, 2, 2, "ID-3") == ("ID-1", 0, True))
check("成功后失败计数清零",
      advance_policy(_pending, 0, 2, "ID-3")[1] == 0)
check("部分成功也算有失败（未贴标签的会重试）",
      advance_policy(_pending, 1, 0, "ID-3")[0] is None)

print("[9] Outlook 进程检测（启动前的防撞车判断）")
import sys  # noqa: E402
from core.outlook import _process_running  # noqa: E402

check("OutlookNotRunning 是 RuntimeError 子类", issubclass(outlook.OutlookNotRunning, RuntimeError))
# 用「自己所在的进程」做阳性用例：一定能枚举到，且不依赖环境里跑着哪些别的程序
check("进程检测：能枚举到自己 → True",
      _process_running(os.path.basename(sys.executable)) is True)
check("进程检测：不存在的进程 → False", _process_running("no_such_process_xyz.exe") is False)
check("进程检测：大小写不敏感",
      _process_running(os.path.basename(sys.executable).upper()) is True)
check("进程检测：空名字 → False", _process_running("") is False)

check("完整性级别检测可用（Windows 上不该是 unknown）", outlook.integrity_level() != "unknown")
check("完整性级别是已知取值之一",
      outlook.integrity_level() in ("Untrusted", "Low", "Medium", "High", "System"))
check("父进程信息可读（形如 name (pid N)）",
      outlook.parent_process().endswith(")") and "pid " in outlook.parent_process())

print("[10] connect() 不许把「附加」也挡掉（针对实际故障的回归测试）")
# 这一节来自一次真实故障：我在 Dispatch 前面加了"启动失败后 120 秒冷却"，
# 结果用户手动打开 Outlook 之后照样报
#   "could not be started 63s ago; not retrying for another 57s"
# —— 冷却把「附加到已在运行的实例」也一起挡掉了。
# 下面用桩替换 Dispatch，确保**每一次连接尝试都会真的去 Dispatch**。
_saved_dispatch = outlook.win32com.client.Dispatch
_silent = lambda *a, **k: None


class _FakeApp:
    def GetNamespace(self, name):
        return {"ns": name}


try:
    outlook.win32com.client.Dispatch = lambda progid: _FakeApp()
    _app, _ns = outlook.connect(log=_silent)
    check("Dispatch 能拿到实例时 connect 成功", _ns == {"ns": "MAPI"})

    _calls = []

    def _boom(progid):
        _calls.append(progid)
        raise RuntimeError("simulated failure")

    outlook.win32com.client.Dispatch = _boom
    _raised = 0
    for _ in range(3):
        try:
            outlook.connect(log=_silent)
        except outlook.OutlookNotRunning as ex:
            _raised += 1
            check("失败时抛的是 OutlookNotRunning 且说明原因", "Could not connect" in str(ex))
    check("连续 3 次失败都抛 OutlookNotRunning", _raised == 3)
    check("连续 3 次都真的尝试了 Dispatch（没有被冷却/宽限期挡住）", len(_calls) == 3)
    check("尝试的是 Outlook.Application", set(_calls) == {"Outlook.Application"})
finally:
    outlook.win32com.client.Dispatch = _saved_dispatch

print("[11] 漏打 Sum 标签的总结报告能被找出来（假 COM 对象，不连 Outlook）")
# 回归来源：plan_tick 只看"最新那一封"，所以只要报告之后又来了一封别的邮件，
# classify 路径会把报告当控制邮件跳过 —— 它就永远轮不到打标签（实际发生过）。
from datetime import datetime as _dt  # noqa: E402


class _FakeMsg:
    def __init__(self, subject, categories, days_ago=0):
        self.Subject = subject
        self.Categories = categories
        self.ReceivedTime = _dt.now() - timedelta(days=days_ago)
        self.EntryID = "E-%d-%s" % (days_ago, subject[:10])


class _FakeItems:
    def __init__(self, msgs):
        self._msgs = msgs

    def Restrict(self, _filter):
        return list(self._msgs)          # 假实现：返回全部，让代码里的二次过滤起作用


class _FakeNS:
    def __init__(self, msgs):
        self._msgs = msgs

    def GetDefaultFolder(self, _fid):
        msgs = self._msgs

        class _Folder:
            Items = _FakeItems(msgs)
        return _Folder()


_msgs = [
    _FakeMsg("📬 邮件总结（全部）· 最近7天", "Sum", 0),        # 已打标签 → 不返回
    _FakeMsg("📬 邮件总结（全部）· 最近7天", "", 1),            # 漏标签 → 返回
    _FakeMsg("📬 邮件总结（全部）· 最近7天", "", 40),           # 超出 30 天 → 不返回
    _FakeMsg("转发：邮件总结 是什么", "", 2),                   # 主题含标记 → 按现有定义算报告
    _FakeMsg("[Press Release] HKUST Convenes Global Leaders", "", 1),   # 普通邮件 → 不返回
]
_got = outlook.untagged_reports(_FakeNS(_msgs), days=30)
check("只返回漏标签的报告（已打标签的跳过）", _msgs[0] not in _got)
check("窗口内漏标签的报告被找到", _msgs[1] in _got)
check("超出天数窗口的不返回", _msgs[2] not in _got)
check("普通邮件不返回", _msgs[4] not in _got)
check("返回顺序与传入顺序一致", _got == [_msgs[1], _msgs[3]])
check("days=None 时不限时间窗口", _msgs[2] in outlook.untagged_reports(_FakeNS(_msgs), days=None))
check("查询不被支持时返回空列表而不是崩掉",
      outlook.untagged_reports(type("NS", (), {"GetDefaultFolder": lambda s, f: type(
          "F", (), {"Items": type("I", (), {"Restrict": lambda s, x: (_ for _ in ()).throw(
              RuntimeError("nope"))})()})()})(), days=30) == [])
check("筛选条件用的是 DASL 主题属性",
      "httpmail:subject" in outlook.REPORT_FILTER
      and outlook.REPORT_SUBJECT_MARK in outlook.REPORT_FILTER)

# 中英标记都要认：报告主题跟随 language，但切换语言后历史报告仍须被识别
check("识别报告：中文主题 → True", outlook.is_report_subject("📬 邮件总结（全部）· 最近7天"))
check("识别报告：英文主题 → True",
      outlook.is_report_subject("📬 Email summary (all levels) · last 7 days"))
check("识别报告：普通邮件 → False", outlook.is_report_subject("HKUST Daily Event Alert") is False)
check("识别报告：空主题 → False", outlook.is_report_subject("") is False)
check("英文筛选条件也用 DASL 且带英文标记",
      outlook.report_filter(outlook.REPORT_SUBJECT_MARK_EN) ==
      '@SQL="urn:schemas:httpmail:subject" LIKE \'%Email summary%\'')
check("两种标记合并查询时按 EntryID 去重（假 Restrict 返回全部，正好考验去重）",
      len(_got) == len({m.EntryID for m in _got}))
_en_msgs = [_FakeMsg("📬 Email summary (all levels) · last 7 days", "", 1)]
check("英文主题的报告同样能被找出来（切语言后不会被误分级）",
      [m.EntryID for m in outlook.untagged_reports(_FakeNS(_en_msgs), days=30)]
      == [_en_msgs[0].EntryID])

check("主题跟随语言：zh → 含中文标记",
      "邮件总结" in summarizer.build_subject("sabc", config.from_dict({"language": "zh"})))
_en_subject = summarizer.build_subject("sabc", config.from_dict({"language": "en"}))
check("主题跟随语言：en → 含英文标记", "Email summary" in _en_subject)
check("主题跟随语言：en → 不再出现中文", "邮件总结" not in _en_subject)
check("主题里的级别标签也跟着语言", "(all levels)" in _en_subject)

print("[12] 埋在别的邮件下面的 sum 指令也能被找到（同一类洞的修复）")


class _OrderedItems:
    """假 Items：支持 Sort 和迭代（pending_sum_commands 走的是遍历路径）。"""

    def __init__(self, msgs):
        self._msgs = msgs

    def Sort(self, *_args):
        pass

    def __iter__(self):
        return iter(self._msgs)


class _CmdMsg:
    def __init__(self, subject, categories, entryid, days_ago=0, hours_ago=0):
        self.Subject = subject
        self.Categories = categories
        self.EntryID = entryid
        self.ReceivedTime = _dt.now() - timedelta(days=days_ago, hours=hours_ago)


class _CmdNS:
    def __init__(self, msgs):
        self._msgs = msgs

    def GetDefaultFolder(self, _fid):
        msgs = self._msgs

        class _Folder:
            Items = _OrderedItems(msgs)
        return _Folder()


# 关键场景：指令**不是**最新那一封（下面第 2 封才是指令），这正是原来会丢指令的情况
_cmds = [
    _CmdMsg("[Press Release] HKUST Convenes Global Leaders", "", "ID-1", hours_ago=0),
    _CmdMsg("sum", "Sum", "ID-2", hours_ago=1),          # 已处理 → 不该再返回
    _CmdMsg("sum sa", "", "ID-3", hours_ago=2),          # 未处理 → 应返回
    _CmdMsg("Your monthly summary is ready", "", "ID-4", hours_ago=3),   # 非指令 → 不返回
    _CmdMsg("Summary of AWS billing", "", "ID-5", hours_ago=4),          # 非指令 → 不返回
    _CmdMsg("sum", "", "ID-6", hours_ago=5),             # 未处理 → 应返回
    _CmdMsg("OLDER-THAN-BASELINE", "", "ID-7", hours_ago=6),
]
_found = outlook.pending_sum_commands(_CmdNS(_cmds), "ID-7", days=7)
check("埋在普通邮件下面的指令被找到", [m.EntryID for m in _found] == ["ID-3", "ID-6"])
check("已打过 Sum 标签的指令不重复执行", all(m.EntryID != "ID-2" for m in _found))
check("主题含 summary 的普通邮件不会被当指令", all(m.EntryID not in ("ID-4", "ID-5") for m in _found))
check("走到基线（上次看过的位置）就停，不看更旧的",
      all(m.EntryID != "ID-7" for m in _found))
check("max_items 上限生效",
      len(outlook.pending_sum_commands(_CmdNS(_cmds), "ID-7", days=7, max_items=1)) == 1)
check("超出天数窗口的指令不返回",
      outlook.pending_sum_commands(_CmdNS([_CmdMsg("sum", "", "ID-9", days_ago=30)]),
                                   None, days=7) == [])
check("没有基线时（首次运行）也能扫到窗口内的指令",
      len(outlook.pending_sum_commands(_CmdNS([_CmdMsg("sum", "", "ID-9", hours_ago=1)]),
                                       None, days=7)) == 1)

print("[13] 总结要点：被截断/失败不再整批留空")
import json as _json  # noqa: E402
import re as _re  # noqa: E402

check("salvage：完整数组原样解出", ai.salvage_array('["a","b"]') == ["a", "b"])
check("salvage：截断的数组只取完整元素", ai.salvage_array('["a","b","c') == ["a", "b"])
check("salvage：一个完整元素都没有 → None", ai.salvage_array('["没写完') is None)
check("salvage：容忍围栏与前后废话",
      ai.salvage_array('说明\n```json\n["a","b"]\n```') == ["a", "b"])
check("salvage：对象元素也能解", ai.salvage_array('[{"x":1},{"y":2}') == [{"x": 1}, {"y": 2}])
check("salvage：只缺结尾括号时全部元素都收下（比预想更宽）",
      ai.salvage_array('["a","b"') == ["a", "b"])
check("salvage：最后一个元素写到一半 → 丢掉它",
      ai.salvage_array('[{"x":1},{"y":') == [{"x": 1}])

check("截断判断：缺右括号 → True", ai.looks_truncated('["a","b"') is True)
check("截断判断：对象缺右括号 → True", ai.looks_truncated('{"a": 1') is True)
check("截断判断：完整的数组 → False", ai.looks_truncated('["a","b"]') is False)
check("截断判断：完整的对象 → False", ai.looks_truncated('{"a": 1}') is False)
check("截断判断：纯文字说明 → False（不是截断，是模型跑偏）",
      ai.looks_truncated("我不太确定该怎么总结这些邮件。") is False)

print("[14] 连接日志的心跳抑制（监听每 30 秒一轮，不能刷屏）")
outlook._seen_notices.clear()
_notice_calls: list = []
_outlook_log = lambda m: _notice_calls.append(m)   # noqa: E731

check("第一次的连接日志会打出来",
      outlook._notice(_outlook_log, "attach", "attaching") is True)
check("同一个 key 再来就被抑制（这就是刷屏的来源）",
      outlook._notice(_outlook_log, "attach", "attaching") is False)
check("抑制不影响别的 key", outlook._notice(_outlook_log, "attached", "attached") is True)
check("换个状态（如改成要启动 Outlook）仍然会打",
      outlook._notice(_outlook_log, "launch", "trying to start") is True)
check("实际只打了 3 条，而不是 4 条", _notice_calls == ["attaching", "attached", "trying to start"])
outlook._seen_notices.clear()
check("清空后可以重新打（模拟新一次会话）",
      outlook._notice(_outlook_log, "attach", "attaching") is True)

_cfg = config.from_dict({"email": "a@b.c", "language": "zh", "batch_size": 40,
                         "ai": {"active": "x", "profiles": {"x": {"type": "cli"}}}})
_logs: list = []


def _log(*args):
    _logs.append(" ".join(str(a) for a in args))


def _fake_ask(good, mode="truncate"):
    """good: {主题: 要点}；不在里面的主题模拟"这一项生成不出来"。

    mode="truncate" 模拟 max_tokens 截断（前面几项还能救回来）；
    mode="hard"     模拟整批失败（模型只回了一句解释性文字）。
    """
    calls: list = []

    def fake(prompt, ai_cfg, *, log=print, retries=3):
        subjects = [s.strip() for s in _re.findall(r"主题:(.*)", prompt)]
        calls.append(subjects)
        if mode == "hard" and any(s not in good for s in subjects):
            return "我不太确定该怎么总结这些邮件。"
        parts = []
        for subject in subjects:
            if subject not in good:
                return "[" + ",".join(parts) + ("," if parts else "") + '"未写完'
            parts.append(_json.dumps(good[subject], ensure_ascii=False))
        return "[" + ",".join(parts) + "]"

    return fake, calls


def _emails(subjects):
    return [{"subject": s, "sender": "someone@ust.hk", "body": "body"} for s in subjects]


_saved_ask = ai.ask
try:
    fake, calls = _fake_ask({"m1": "要点1", "m2": "要点2", "m3": "要点3"})
    ai.ask = fake
    _pts = summarizer.summarize_points(_emails(["m1", "m2", "m3"]), _cfg, log=_log)
    check("正常情况：要点数量与顺序正确", _pts == ["要点1", "要点2", "要点3"])
    check("正常情况：只问了一次模型", len(calls) == 1)

    _logs.clear()
    fake, calls = _fake_ask({"m1": "要点1", "m2": "要点2"})     # m3 生成不出来
    ai.ask = fake
    _pts = summarizer.summarize_points(_emails(["m1", "m2", "m3"]), _cfg, log=_log)
    check("截断时：已拿到的两个要点被保留（旧的实现会全丢）", _pts[:2] == ["要点1", "要点2"])
    check("截断时：只有拿不到的那一封填空", _pts == ["要点1", "要点2", ""])
    check("截断时：日志说明只拿回两个", any("only 2/3" in m for m in _logs))
    check("截断时：日志点名了是哪一封", any("m3" in m for m in _logs))

    _logs.clear()
    fake, calls = _fake_ask({"m1": "要点1", "m3": "要点3", "m4": "要点4"}, mode="hard")
    ai.ask = fake
    _pts = summarizer.summarize_points(_emails(["m1", "m2", "m3", "m4"]), _cfg, log=_log)
    check("整批失败：顺序不错位", _pts == ["要点1", "", "要点3", "要点4"])
    check("整批失败：确实劈半重试了", any("splitting" in m for m in _logs))
    check("整批失败：只影响那一封", sum(1 for p in _pts if p) == 3)
    check("整批失败：劈半也救不回来的那封被点名", any("m2" in m for m in _logs))

    _logs.clear()
    ai.ask = lambda prompt, ai_cfg, **kw: _json.dumps({"summaries": ["A", "B"]})
    check("模型多包一层 {\"summaries\": [...]} 也能解析",
          summarizer.summarize_points(_emails(["m1", "m2"]), _cfg, log=_log) == ["A", "B"])

    _logs.clear()
    # 实测 deepseek-flash 会偶发这种：数组被当成字符串，塞进 {"type","content"} 包装里
    ai.ask = lambda prompt, ai_cfg, **kw: _json.dumps(
        {"type": "json_object", "content": _json.dumps(["甲", "乙"], ensure_ascii=False)})
    check("数组被塞进 {\"type\",\"content\"} 字符串包装里也能解析",
          summarizer.summarize_points(_emails(["m1", "m2"]), _cfg, log=_log) == ["甲", "乙"])

    _logs.clear()
    ai.ask = lambda prompt, ai_cfg, **kw: "我不太确定该怎么总结这些邮件。"
    check("完全拿不到时：全部填空但不抛异常",
          summarizer.summarize_points(_emails(["m1", "m2"]), _cfg, log=_log) == ["", ""])

    _logs.clear()
    ai.ask = lambda prompt, ai_cfg, **kw: _json.dumps([123, 456])
    check("返回的不是字符串数组时：当作不可用（不把数字塞进报告）",
          summarizer.summarize_points(_emails(["m1", "m2"]), _cfg, log=_log) == ["", ""])
finally:
    ai.ask = _saved_ask

print(f"\n全部通过：{ok} 项")
print("\n--- build_spec 实际产出（配置缺字段时的样子）---")
print(classifier.build_spec(config.from_dict({"email": "a@b.c"})))
