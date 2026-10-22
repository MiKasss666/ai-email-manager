"""selftest.py — 纯逻辑自测：不连接 Outlook、不调用 AI、不消耗任何额度。

用法（在 Email_manager_beta 目录下）：
    py -3 selftest.py

覆盖 config / ai.parse_json / classifier / summarizer 里所有不依赖 COM 的逻辑。
旧版没有任何测试，规则改错只能等线上邮件暴露；这里是这套内核的第一道防线。
"""
import os
from datetime import date, datetime, timedelta

# 自测固定走 COM 后端：这一套断言验证的是「纯逻辑 + COM 那条路的行为」，
# 必须确定性，不能受 config.json 里 mail_source 的影响。Graph 后端的断言见文末 [15]。
os.environ["EM_MAIL_SOURCE"] = "outlook"

from core import ai, classifier, config, outlook, summarizer  # noqa: E402

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
check("urgent_days 非数字 → 默认 14", cfg.urgent_days == 14)
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
check("缺 school_domains 键时补默认值 ust.hk",
      config.from_dict({}).school_domains == ["ust.hk"])
check("显式清空 school_domains 回落默认值（否则整箱邮件都不分级）",
      config.from_dict({"school_domains": []}).school_domains == ["ust.hk"])
check("school_domains 支持逗号分隔字符串",
      config.from_dict({"school_domains": "ust.hk, hku.hk"}).school_domains == ["ust.hk", "hku.hk"])
check("to_dict/from_dict 往返保留 school_domains",
      config.from_dict(config.to_dict(cfg)).school_domains == cfg.school_domains)
check("缺 language 键时用默认值 en", config.from_dict({}).language == "en")
check("显式 language=zh 生效", config.from_dict({"language": "zh"}).language == "zh")
check("无法识别的 language 回落默认值", config.from_dict({"language": "fr"}).language == "en")

print("[1b] config.is_school_address（分级前置门槛）")
school = config.from_dict({"school_domains": ["ust.hk"]})
check("学生邮箱 @connect.ust.hk 命中", school.is_school_address("student@connect.ust.hk"))
check("教职员邮箱 @ust.hk 命中", school.is_school_address("office@ust.hk"))
check("大小写不敏感", school.is_school_address("STUDENT@CONNECT.UST.HK"))
check("带显示名的地址命中", school.is_school_address("Alex Doe <student@connect.ust.hk>"))
check("校外邮箱不命中", not school.is_school_address("news@instructure.com"))
check("相似后缀不误判（ust.hk.evil.com 不是校内）",
      not school.is_school_address("a@ust.hk.evil.com"))
check("Exchange 内部地址（没有 @）不命中", not school.is_school_address("/O=EXCHANGELABS/CN=abc"))
check("空地址不命中", not school.is_school_address(""))
check("多域名配置生效",
      config.from_dict({"school_domains": ["ust.hk", "hku.hk"]}).is_school_address("x@hku.hk"))
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
check("对象里含数组时不会被内层数组劫持（整段优先）",
      ai.parse_json('{"a": [1, 2], "b": 3}') == {"a": [1, 2], "b": 3})
check("JSONL：一行一个对象也能解析成数组",
      ai.parse_json('{"级别": "A"}\n{"级别": "S"}\n{"级别": "C"}')
      == [{"级别": "A"}, {"级别": "S"}, {"级别": "C"}])
check("JSONL 最后一行写坏了：前面完整的仍能拿到",
      ai.parse_json('{"级别": "A"}\n{"级别": "S"}\n{"级别')
      == [{"级别": "A"}, {"级别": "S"}])

print("[2b] AI 调用有墙钟上限（socket 超时兜不住的那种挂起）")
import time as _t  # noqa: E402

_saved_callers = dict(ai._CALLERS)
_slow = config.AiProfile(name="slow", type="openai", base_url="https://example.invalid",
                         model="m", api_key="k", timeout=1)
try:
    ai._CALLERS["openai"] = lambda prompt, cfg: _t.sleep(30)
    _t0 = _t.time()
    _timed_out = None
    try:
        ai.ask("hi", _slow, log=lambda *a: None, retries=1)
    except ai.AiError as ex:
        _timed_out = ex
    _took = _t.time() - _t0
    check("挂起的调用会在 timeout 秒后被放弃（不再无限等）",
          _timed_out is not None and _took < 10)
    check("这种超时算瞬时故障（会重试，而不是判死）",
          _timed_out is not None and _timed_out.transient)
    check("错误信息说清是对方没回话、不是 Outlook 的问题",
          _timed_out is not None and "did not answer within" in str(_timed_out))

    class _Boom(RuntimeError):
        pass

    def _throw(prompt, cfg):
        raise _Boom("boom")

    ai._CALLERS["openai"] = _throw
    _other = None
    try:
        ai.ask("hi", _slow, log=lambda *a: None, retries=1)
    except BaseException as ex:
        _other = ex
    check("调用里抛的其它异常原样带回（不被兜成超时）", isinstance(_other, _Boom))

    ai._CALLERS["openai"] = lambda prompt, cfg: "OK"
    check("正常返回照样能用", ai.ask("hi", _slow, log=lambda *a: None, retries=1) == "OK")
finally:
    ai._CALLERS.clear()
    ai._CALLERS.update(_saved_callers)

print("[3] classifier")
spec = classifier.build_spec(cfg)
check("spec 含课程", "COMP1021" in spec)
check("spec 含学院", "School of Engineering" in spec)
check("spec 含宿舍关键词（港科大默认值）", "UGHousing" in spec)
check("spec 含加急天数（且出现在「截止」语境里，不是死开关）",
      f"{cfg.urgent_days} 天内" in spec and "截止" in spec)
check("spec 明确「活动举办日期不算截止」", "活动举办日期不算截止" in spec)
check("spec 明确过期通知的归属", "已过期" in spec)
check("spec 含「学校邮件」判定步骤", "学校邮件" in spec)
check("spec 用学校域名做确定性判据", cfg.school_domains[0] in spec)
check("spec 不再用学校名当判级条件", "香港科技大学" not in spec)

print("[3b] 学校邮件判定：AI 判断 + 域名确定性兜底")
check("parse_school：bool", classifier.parse_school(True) is True
      and classifier.parse_school(False) is False)
check("parse_school：0/1", classifier.parse_school(1) is True
      and classifier.parse_school(0) is False)
check("parse_school：中文是/否", classifier.parse_school("是") is True
      and classifier.parse_school("不是") is False)
check("parse_school：否定短语", classifier.parse_school("不是学校邮件") is False)
check("parse_school：整句否定（否，这是商家推广）", classifier.parse_school("否，这是商家推广") is False)
check("parse_school：整句肯定（是，属于学校邮件）", classifier.parse_school("是，属于学校邮件") is True)
check("parse_school：不属于 → False", classifier.parse_school("不属于学校邮件") is False)
check("parse_school：肯定短语", classifier.parse_school("是学校邮件") is True)
check("parse_school：认不出来 → None", classifier.parse_school("maybe") is None
      and classifier.parse_school("") is None)

_school_cfg = config.from_dict({"school_domains": ["ust.hk"]})
_mixed = classifier.to_verdicts([
    {"学校": True, "级别": "A", "截止日期": None},
    {"学校": False, "级别": "S", "截止日期": None},
    {"学校": True, "级别": "B", "截止日期": None},
    {"级别": "C", "截止日期": None},          # 模型漏写「学校」
], 4, _school_cfg)
check("to_verdicts 解析「学校」字段",
      [v.is_school for v in _mixed] == [True, False, True, True])
check("漏写「学校」时保守按学校邮件处理（不漏分级）", _mixed[3].is_school)
check("counts 不计入非学校邮件", classifier.counts(_mixed) == {"S": 0, "A": 1, "B": 1, "C": 1})

_rescue_emails = [
    {"subject": "Campus notice", "sender": "office@ust.hk", "entryid": "R-1"},
    {"subject": "Shop promo", "sender": "promo@shop.com", "entryid": "R-2"},
]
_rescue = classifier.to_verdicts([{"学校": False, "级别": "C"},
                                  {"学校": False, "级别": "C"}], 2, _school_cfg)
check("AI 把学校邮件误判为非学校 → 域名兜底捞回来",
      classifier.apply_school_override(_rescue_emails, _rescue, _school_cfg) == 1
      and _rescue[0].is_school and not _rescue[1].is_school)
check("not_school_entryids 只列被判为非学校邮件的（供 watcher 记名单）",
      classifier.not_school_entryids(_rescue_emails, _rescue) == ["R-2"])

print("[3c] apply_verdicts：非学校邮件一个字节都不写")


class _WriteMsg:
    """假 MailItem：只记录分类、标志位和 Save() 次数。"""

    def __init__(self):
        self.Categories = ""
        self.FlagStatus = None
        self.saves = 0

    def Save(self):
        self.saves += 1


_m_school, _m_other = _WriteMsg(), _WriteMsg()
_w_emails = [{"subject": "School stuff", "msg": _m_school},
             {"subject": "Shop promo", "msg": _m_other}]
_w_verdicts = classifier.to_verdicts([{"学校": True, "级别": "S", "截止日期": None},
                                      {"学校": False, "级别": "A", "截止日期": None}],
                                     2, _school_cfg)
check("apply_verdicts 返回 (成功, 失败, 跳过)",
      classifier.apply_verdicts(_w_emails, _w_verdicts, log=lambda *a: None) == (1, 0, 1))
check("学校邮件被写入分类", _m_school.Categories == "S" and _m_school.saves == 1)
check("非学校邮件完全没被动过", _m_other.Categories == "" and _m_other.saves == 0)
check("拿不到判定算失败（触发重试），不算跳过",
      classifier.apply_verdicts([{"subject": "x", "msg": _WriteMsg()}], [None],
                                log=lambda *a: None) == (0, 1, 0))

print("[3d] write_categories：只换自己的标签，用户自己的分类一律不动")


class _CatMsg:
    """假 MailItem：带一个已有的 Categories 字符串。"""

    def __init__(self, categories):
        self.Categories = categories
        self.FlagStatus = None
        self.saves = 0

    def Save(self):
        self.saves += 1


_mc = _CatMsg("跟进, 重要")
outlook.write_categories(_mc, ["A"])
check("用户自己的分类被保留（并追加我们的级别）", _mc.Categories == "跟进, 重要, A")
check("写回后确实 Save 了", _mc.saves == 1)

_mc = _CatMsg("S, 跟进")
outlook.write_categories(_mc, ["B"])
check("我们自己的旧级别被替换，不叠加", _mc.Categories == "跟进, B")

_mc = _CatMsg("A, B, C")
outlook.write_categories(_mc, ["S"])
check("多档叠加的脏数据被收敛成一个级别", _mc.Categories == "S")

_mc = _CatMsg("S 紧急, 学期")
outlook.write_categories(_mc, ["C"])
check("历史命名（S 紧急）被顺手清掉", _mc.Categories == "学期, C")

_mc = _CatMsg("Sum, 归档")
outlook.write_categories(_mc, ["A"])
check("报告标签 Sum 也属于我们自己，会被换掉", _mc.Categories == "归档, A")

_mc = _CatMsg("")
outlook.write_categories(_mc, ["Sum"])
check("原本没有任何分类时只写我们的", _mc.Categories == "Sum")

_mc = _CatMsg(" 跟进 ,  ")
outlook.write_categories(_mc, ["A"])
check("已有分类里的空白项被清掉", _mc.Categories == "跟进, A")

check("本系统标签集合 = S/A/B/C/Sum + 历史命名",
      set(outlook.OUR_CATEGORY_NAMES) == {"S", "A", "B", "C", "Sum",
                                          "S 紧急", "A 重要", "B 普通", "C 可忽略"})

print("[3e] 一封读不出正文的邮件不能拖垮整轮（逐封兜错）")


class _PoisonMsg:
    """假邮件：Body 会抛异常（IRM/加密邮件、正在被移动的邮件就是这样）。"""

    def __init__(self, subject):
        self.Subject = subject
        self.SenderName = "Someone"
        self.SenderEmailAddress = "a@ust.hk"
        self.EntryID = "P-" + subject
        self.Categories = ""
        self.ReceivedTime = datetime.now()

    @property
    def Body(self):
        raise RuntimeError("Body unavailable (IRM/encrypted or item moved)")


class _PoisonItems:
    def __init__(self, msgs):
        self._msgs = msgs

    def Sort(self, *_args):
        pass

    def __iter__(self):
        return iter(self._msgs)


class _PoisonNS:
    def __init__(self, msgs):
        self._msgs = msgs

    def GetDefaultFolder(self, _fid):
        msgs = self._msgs

        class _Folder:
            Items = _PoisonItems(msgs)
        return _Folder()


_poison_msgs = [_PoisonMsg("good-1"), _PoisonMsg("poison"), _PoisonMsg("good-2")]
_poison_ok = True
_poison_got = []
try:
    _poison_got = outlook.read_inbox(_PoisonNS(_poison_msgs), days=7, log=lambda *a: None)
except Exception as _ex:                      # 旧行为：整轮直接抛异常
    _poison_ok = False
    print(f"      read_inbox raised: {type(_ex).__name__}: {_ex}")
check("坏邮件不再让整轮抛异常", _poison_ok)
check("三封邮件都在（坏的那封留空正文，不丢）",
      [m["subject"] for m in _poison_got] == ["good-1", "poison", "good-2"])
check("坏邮件其余字段仍可用（靠主题也能判级）",
      _poison_got[1]["body"] == "" and _poison_got[1]["entryid"] == "P-poison")

print("[3f] config.save 的临时文件名带 pid（两个实例不会互相写坏配置）")
import tempfile  # noqa: E402
import core.config as config_mod  # noqa: E402

_tmp_dir = tempfile.mkdtemp()
_tmp_cfg_path = os.path.join(_tmp_dir, "config.json")
_seen_sources = []
_orig_replace = config_mod.os.replace
try:
    def _spy_replace(src, dst):
        _seen_sources.append(src)
        return _orig_replace(src, dst)

    config_mod.os.replace = _spy_replace
    config.save(config.from_dict({"email": "a@b.c"}), _tmp_cfg_path)
finally:
    config_mod.os.replace = _orig_replace
check("配置写进去了", os.path.exists(_tmp_cfg_path))
check("临时文件名里带自己的 pid",
      bool(_seen_sources) and f".{os.getpid()}." in os.path.basename(_seen_sources[0]))
check("没留下临时文件", not [f for f in os.listdir(_tmp_dir) if f.endswith(".tmp")])
os.remove(_tmp_cfg_path)
os.rmdir(_tmp_dir)

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
check("对象包裹（契约里的 items）也能解析",
      (classifier.to_verdicts({"items": [{"级别": "S"}]}, 1, cfg) or [None])[0].level == "S")
check("键名不认识、但只有一个数组值时也认",
      (classifier.to_verdicts({"结果集": [{"级别": "B"}]}, 1, cfg) or [None])[0].level == "B")
check("对象里有多个数组 → 不猜，返回 None",
      classifier.to_verdicts({"a": [{"级别": "S"}], "b": [{"级别": "A"}]}, 1, cfg) is None)
check("提示词要的是对象包裹（与 json_object 模式一致）",
      '"items"' in classifier.build_prompt([{"subject": "x", "sender": "a@ust.hk", "body": "b"}], cfg))

emails = [
    {"subject": "COMP1021 Homework 3", "body": "due soon"},
    {"subject": "Welcome", "body": "nothing special"},
]
verdicts = classifier.to_verdicts([{"级别": "C", "截止日期": None},
                                   {"级别": "B", "截止日期": None}], 2, cfg)
check("确定性规则改写封数 = 1", classifier.apply_enrolled_override(emails, verdicts, cfg) == 1)
check("课程邮件被抬到至少 A（不再一律顶成 S）", verdicts[0].level == "A")
check("无关邮件不受影响", verdicts[1].level == "B")
check("counts 统计", classifier.counts(verdicts) == {"S": 0, "A": 1, "B": 1, "C": 0})
check("counts 容忍 None", classifier.counts([None, verdicts[0]])["A"] == 1)
check("无课程时不误伤", classifier.apply_enrolled_override(
    emails, verdicts, config.from_dict({"email": "a@b.c"})) == 0)
_s_verdict = classifier.to_verdicts([{"级别": "S", "截止日期": "2026-10-20"}], 1, cfg)
classifier.apply_enrolled_override(emails[:1], _s_verdict, cfg)
check("AI 已经判 S 的课程邮件不被降级", _s_verdict[0].level == "S")

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
     "sender": "Prof. A", "body": "b", "received": datetime(2026, 3, 1, 9, 30),
     "web_link": "https://outlook.office.com/mail/deeplink/read/AAA111"},
    {"entryid": "BBB222", "level": "C", "subject": "Newsletter",
     "sender": "Shop", "body": "b", "received": datetime(2026, 3, 2, 8, 0)},
]
html = summarizer.build_html(samples, ["要点一", ""], cfg, "sc")
check("有网页版链接时标题可点击",
      "https://outlook.office.com/mail/deeplink/read/AAA111" in html)
check("不再生成点不动的 outlook: 假链接", "outlook:" not in html)
check("拿不到链接时标题退化成纯文本（不生成空 href）",
      "<a href=''>" not in html and "Newsletter" in html)
check("HTML 转义主题尖括号", "&lt;重要&gt;" in html and "&amp;" in html)
check("HTML 分级标题", "【S】" in html and "【C】" in html)
check("HTML 块数正确（2 封各一个 div 块）", html.count("border-bottom:1px solid #e6e6e6") == 2)
check("空要点不产生空行", html.count("要点一") == 1)
check("div 标签配平", html.count("<div") == html.count("</div>"))
check("HTML 为空列表时也不崩",
      summarizer.build_html([], [], config.from_dict({"email": "a@b.c"}), "sabc").startswith("<html>"))
check("主题含识别标记（跟随 language，见 [11] 节的中英双标记测试）",
      outlook.is_report_subject(summarizer.build_subject("sa", cfg)))

# 界面上的 Scope（last N days）对 Sum 也生效：build_subject / build_html 都要能收这个天数
check("报告主题可用界面 Scope 的天数覆盖",
      "last 3 days" in summarizer.build_subject("sabc", config.from_dict({"language": "en"}), 3)
      and "最近3天" in summarizer.build_subject("sabc", config.from_dict({"language": "zh"}), 3))
check("报告标题同样跟随界面 Scope",
      "last 3 days" in summarizer.build_html(samples, ["", ""],
                                             config.from_dict({"language": "en"}), "sc", 3))
check("不传天数时仍用配置的 summary_days",
      "last 7 days" in summarizer.build_subject(
          "sabc", config.from_dict({"language": "en", "summary_days": 7})))

_saved_connect = summarizer.outlook.connect
_saved_recent = summarizer.outlook.read_recent
_seen = {}
try:
    summarizer.outlook.connect = lambda log=None: (object(), object())
    summarizer.outlook.read_recent = \
        lambda ns, *, days, letters: (_seen.update(days=days, letters=letters), [])[1]
    _count = summarizer.summarize_inbox(
        config.from_dict({"email": "a@b.c", "summary_days": 7}),
        "sa", days=2, log=lambda *a: None)
    check("手动 Sum 时按界面 Scope 取数（days=2，而不是配置里的 7）",
          _seen.get("days") == 2 and _count == 0)
    check("取数级别仍然按参数走", _seen.get("letters") == "sa")
finally:
    summarizer.outlook.connect = _saved_connect
    summarizer.outlook.read_recent = _saved_recent

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
# 注意：这条断言在 2026-10 改过。原来的写法是"不许存在 _last_launch_attempt"，
# 目的是防止把退避挡在 Dispatch 前面（那样用户手动开了 Outlook 也连不上）。
# 现在退避确实存在，但只用于「启动」这一条路径 —— 真正的防回归断言在 [15] 节：
# 它检查 connect() 里 _dispatch() 出现在退避判断**之后**，即附加永不被挡。

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
_saved_launch = outlook._launch_outlook
_saved_minimize = outlook._minimize_outlook
_saved_process_running = outlook._outlook_process_running
_saved_launch_attempt = outlook._last_launch_attempt
_silent = lambda *a, **k: None
_launched = []
import time as _time  # noqa: E402  桩里复刻退避计时用


class _FakeApp:
    def GetNamespace(self, name):
        return {"ns": name}


try:
    # 自测**绝不能启动用户的 Outlook**：把「启动」和「最小化」换成桩，
    # 并把「有没有 OUTLOOK.EXE 进程」固定成否 —— 这样代码走的仍是启动分支，
    # 但真正干活的 subprocess.Popen 不会执行（曾经真的把用户的 Outlook 拉起来过）。
    def _fake_launch(log):
        # 复刻真实 _launch_outlook 的关键副作用：记下本次启动尝试的时间（退避靠它）
        _launched.append("launch")
        outlook._last_launch_attempt = _time.time()

    outlook._launch_outlook = _fake_launch
    outlook._minimize_outlook = lambda app: None
    outlook._outlook_process_running = lambda: False
    outlook._last_launch_attempt = 0.0

    outlook.win32com.client.Dispatch = lambda progid: _FakeApp()
    _app, _ns = outlook.connect(log=_silent)
    check("Dispatch 能拿到实例时 connect 成功", _ns == {"ns": "MAPI"})
    check("允许启动时走了一次「启动」分支（已被桩住，不会真的启动 Outlook）",
          len(_launched) == 1)

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
    check("退避期内不再重复尝试启动 Outlook", len(_launched) == 1)
finally:
    outlook.win32com.client.Dispatch = _saved_dispatch
    outlook._launch_outlook = _saved_launch
    outlook._minimize_outlook = _saved_minimize
    outlook._outlook_process_running = _saved_process_running
    outlook._last_launch_attempt = _saved_launch_attempt

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

print("[15] 启动 Outlook 的退避（只约束「启动」，绝不约束「附加」）")
import time as _time  # noqa: E402

check("能找到 OUTLOOK.EXE（本机装了 Office）", bool(outlook.outlook_exe_path()))
if outlook.outlook_exe_path():
    check("定位到的路径确实存在", os.path.exists(outlook.outlook_exe_path()))
    check("路径以 OUTLOOK.EXE 结尾", outlook.outlook_exe_path().upper().endswith("OUTLOOK.EXE"))

_saved_launch = outlook._last_launch_attempt
try:
    outlook._last_launch_attempt = 0.0
    check("从未尝试过 → 允许启动", outlook._launch_allowed() is True)
    outlook._last_launch_attempt = _time.time()
    check("刚尝试过 → 退避期内不再启动", outlook._launch_allowed() is False)
    outlook._last_launch_attempt = _time.time() - (outlook.LAUNCH_BACKOFF_SECONDS + 1)
    check("超过退避时长 → 重新允许启动", outlook._launch_allowed() is True)
finally:
    outlook._last_launch_attempt = _saved_launch

# 这一条是防回归的关键：退避**不能**出现在 Dispatch 之前
import inspect as _inspect  # noqa: E402
_src = _inspect.getsource(outlook.connect)
check("connect() 里 Dispatch 出现在退避判断之后（退避不会挡掉附加）",
      _src.find("_dispatch()") > _src.find("_launch_allowed()"))
check("退避提示里有明确的人工处理指引",
      "open Outlook yourself" in _inspect.getsource(outlook._notice_backoff))

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

print("[14] 采集层过滤器：已判定为非学校邮件的 EntryID 不再重复送 AI")


class _SchoolMsg:
    """假邮件。被过滤掉的邮件，其 Body 属性会抛异常 —— 只要过滤器在 message_info()
    之前生效，就绝不会碰到它；碰到了就是回归。"""

    def __init__(self, subject, sender, entryid, days_ago=0):
        self.Subject = subject
        self.SenderName = "Someone"
        self.SenderEmailAddress = sender
        self.EntryID = entryid
        self.Categories = ""
        self.ReceivedTime = _dt.now() - timedelta(days=days_ago)
        # 独立按发件人域名判断，不抄代码的结论
        self._school = str(sender).lower().endswith("ust.hk")

    @property
    def Body(self):
        if not self._school:
            raise RuntimeError("被过滤掉的邮件不该读正文")
        return "（校方邮件正文）"


class _SchoolNS:
    def __init__(self, msgs):
        self._msgs = msgs

    def GetDefaultFolder(self, _fid):
        msgs = self._msgs

        class _Folder:
            Items = _OrderedItems(msgs)
        return _Folder()


_smsgs = [
    _SchoolMsg("Assignment Graded: Homework 2, MATH1023", "ta@connect.ust.hk", "S-1", days_ago=0),
    _SchoolMsg("Weekly deals from a shop", "promo@shop.com", "N-1", days_ago=1),
    _SchoolMsg("HKUST Daily Event Alert", "events@ust.hk", "S-2", days_ago=1),
    _SchoolMsg("Your AWS billing statement", "billing@amazon.com", "N-2", days_ago=2),
    _SchoolMsg("sum", "me@connect.ust.hk", "S-3", days_ago=3),        # 指令邮件 → 照旧不进分级
    _SchoolMsg("📬 邮件总结（全部）· 最近7天", "me@connect.ust.hk", "S-4", days_ago=1),  # 报告 → 照旧跳过
]
_sns = _SchoolNS(_smsgs)
# watcher 会把「判定为非学校邮件」的 EntryID 记进 watch_state.json，这里模拟同一份名单
_ignore = {"N-1", "N-2"}


def _keep(msg):
    return msg.EntryID not in _ignore


_slog = []

_sread = outlook.read_inbox(_sns, days=7, skip_classified=True,
                            keep=_keep, log=_slog.append)
check("read_inbox：名单里的邮件不再采集", [m["entryid"] for m in _sread] == ["S-1", "S-2"])
check("read_inbox：被过滤的邮件记了条数",
      any("skipped 2 mail(s)" in m for m in _slog))
check("read_inbox：sum 指令与总结报告照旧不进分级",
      all(m["entryid"] not in ("S-3", "S-4") for m in _sread))

_slog.clear()
_scollect = outlook.collect_new_mail(_sns, "S-3", days=7, max_items=100,
                                     keep=_keep, log=_slog.append)
check("collect_new_mail：走到基线为止，且跳过名单里的邮件",
      [m["entryid"] for m in _scollect] == ["S-1", "S-2"])
check("collect_new_mail：被过滤的邮件同样记条数",
      any("skipped 2 mail(s)" in m for m in _slog))

check("limit 只数通过过滤器的邮件",
      [m["entryid"] for m in outlook.read_inbox(_sns, limit=1, days=7,
                                                keep=_keep, log=lambda *a: None)]
      == ["S-1"])

print("[16] 邮件后端选择（graph / outlook）与 Graph 后端的纯逻辑")
import core.graph as graph_mod                                 # noqa: E402
import core.mail as mail_mod                                   # noqa: E402

check("自测固定走 COM 后端（避免联网/登录）", mail_mod.BACKEND == "outlook")
check("两个后端共用同一个「连不上」异常类（否则 watcher 的 isinstance 会失效）",
      graph_mod.OutlookNotRunning is outlook.OutlookNotRunning)
check("mail 暴露的端口名齐全",
      all(hasattr(mail_mod, name) for name in
          ("connect", "read_inbox", "collect_new_mail", "write_categories", "send_html",
           "newest_message", "message_info", "read_recent", "pending_sum_commands",
           "untagged_reports", "ensure_categories", "ensure_sum_category",
           "existing_categories", "our_category_colors", "category_names",
           "has_our_category", "parse_sum_command", "is_report", "integrity_level")))
check("mail.describe() 说得清当前后端", "outlook" in mail_mod.describe())

check("mail_source 默认是 graph", config.from_dict({}).mail_source == "graph")
check("mail_source 认得出 outlook",
      config.from_dict({"mail_source": "outlook"}).mail_source == "outlook")
check("mail_source 写错时回落默认",
      config.from_dict({"mail_source": "pop3"}).mail_source == "graph")
check("graph 段有默认值",
      config.from_dict({}).graph.tenant == "common"
      and config.from_dict({}).graph.open_browser is True)
check("graph 段可配置",
      config.from_dict({"graph": {"tenant": "hkust.edu.hk",
                                  "client_id": "abc"}}).graph.tenant == "hkust.edu.hk")
check("to_dict/from_dict 往返保留 graph 设置",
      config.from_dict(config.to_dict(
          config.from_dict({"graph": {"client_id": "xyz"}}))).graph.client_id == "xyz")

_raw = {
    "id": "AAMkAGI2", "subject": "Homework 2", "receivedDateTime": "2026-10-15T02:04:33Z",
    "categories": ["S"], "isRead": False,
    "from": {"emailAddress": {"name": "Canvas", "address": "notifications@instructure.com"}},
    "body": {"contentType": "html", "content": "<p>Due <b>soon</b></p><br><div>MATH1023</div>"},
    "bodyPreview": "Due soon MATH1023",
}
_info = graph_mod.normalize(_raw)
check("Graph 邮件归一：主题 / 发件人 / id / 分类",
      _info["subject"] == "Homework 2"
      and _info["sender"] == "notifications@instructure.com"
      and _info["entryid"] == "AAMkAGI2" and _info["categories"] == ["S"])
check("Graph 邮件归一：HTML 压成纯文本", "soon" in _info["body"] and "<" not in _info["body"])
check("Graph 邮件归一：时间是本地 naive",
      _info["received"] is not None and _info["received"].tzinfo is None)
check("Graph 邮件归一：带出网页版链接 web_link",
      graph_mod.normalize(dict(_raw, webLink="https://outlook.office.com/x"))["web_link"]
      == "https://outlook.office.com/x")
check("没有 webLink 时 web_link 是空串（报告会退化成纯文本）",
      _info["web_link"] == "")

_str_body = dict(_raw)
_str_body["body"] = "plain text body"
_str_body["from"] = "someone@ust.hk"
_info2 = graph_mod.normalize(_str_body)
check("body 是字符串时不炸（真实邮箱里遇到过）", _info2["body"] == "plain text body")
check("from 是字符串时也能取到地址", _info2["sender"] == "someone@ust.hk")
check("几乎空的邮件也不炸", graph_mod.normalize({})["subject"] == "(no subject)")

_gmsg = graph_mod.GraphMessage(_raw, client=None)
check("GraphMessage 扮得像 COM MailItem（Subject/Categories/EntryID）",
      _gmsg.Subject == "Homework 2" and _gmsg.Categories == "S"
      and _gmsg.EntryID == "AAMkAGI2")
check("core.outlook 的纯逻辑能直接用在 GraphMessage 上",
      outlook.message_info(_gmsg)["sender"] == "notifications@instructure.com"
      and outlook.category_names(_gmsg) == ["S"]
      and outlook.has_our_category(_gmsg) is True
      and outlook.is_report(_gmsg) is False)
check("logged_in 只查本地凭据（不联网）", isinstance(graph_mod.logged_in(), bool))
check("token 缓存默认放在 config.json 旁边",
      os.path.basename(graph_mod.token_path(config.from_dict({}).graph)) == "graph_token.json")

# ── 回归：二次归一化把邮件 id 弄丢，导致所有写回 PATCH /me/messages/ → HTTP 405 ──
# 实测踩过：GraphSession 拿到的是已归一化的字典，再喂给 GraphMessage 时 id 变空串，
# 个人版日志里刷 "could not tag a report: HTTP 405 ..."。单测当时用的是原始条目，所以没抓到。
check("normalize 幂等：二次归一化不丢 id",
      graph_mod.normalize(graph_mod.normalize(_raw))["entryid"] == "AAMkAGI2")


class _FakeClient:
    """只实现 GraphSession/写回需要的那几个方法。"""

    def __init__(self):
        self.calls = []

    def inbox_raw(self, top=10, days=None, select=""):
        return [_raw]                      # 必须是**原始**条目

    def set_categories(self, message_id, names, flag=None):
        self.calls.append((message_id, list(names), flag))


_fclient = _FakeClient()
_ns = graph_mod.GraphSession(_fclient)
_sess_msgs = _ns.recent()
check("GraphSession 保住邮件 id（用原始条目构造）",
      bool(_sess_msgs) and _sess_msgs[0].id == "AAMkAGI2"
      and _sess_msgs[0].EntryID == "AAMkAGI2")
check("GraphSession 保住收件时间", _sess_msgs[0].ReceivedTime is not None)
check("message_info 给出的 entryid 非空",
      bool(outlook.message_info(_sess_msgs[0]).get("entryid")))
graph_mod.write_categories(_sess_msgs[0], ["Sum"])
check("写回用的是真实 id（不是空串）",
      _fclient.calls == [("AAMkAGI2", ["Sum"], None)])
try:
    graph_mod.GraphMail("token").set_categories("", ["Sum"])
    _empty_guard = False
except graph_mod.GraphError:
    _empty_guard = True
check("空 id 会被拦下（不发注定 405 的请求）", _empty_guard)

print("[17] Graph 登录交互：设备码进日志 + 自动打开浏览器（离线桩，不联网）")
_saved_post = graph_mod._post_form
_saved_open = graph_mod.webbrowser.open
_saved_clip = graph_mod._copy_to_clipboard
_logs: list = []
_opened: list = []
_clipped: list = []


def _fake_post(url, fields):
    if url.endswith("/devicecode"):
        return 200, {"device_code": "DC", "user_code": "ABCD-1234",
                     "verification_uri": "https://microsoft.com/devicelogin",
                     "interval": 1, "expires_in": 30}
    return 200, {"access_token": "AT", "refresh_token": "RT"}


try:
    graph_mod._post_form = _fake_post
    graph_mod.webbrowser.open = lambda url: _opened.append(url)
    graph_mod._copy_to_clipboard = lambda text: (_clipped.append(text), True)[1]
    _tok = graph_mod._device_login(config.from_dict({}).graph, _logs.append, "Mail.Read")
    _joined = "\n".join(_logs)
    check("设备码出现在日志里（用户可复制）", "ABCD-1234" in _joined)
    check("登录网址也在日志里", "microsoft.com/devicelogin" in _joined)
    check("自动打开了浏览器", len(_opened) == 1 and _opened[0].startswith("https://"))
    check("浏览器链接里带着设备码（免手输）", "ABCD-1234" in _opened[0])
    check("设备码被放进剪贴板（可直接 Ctrl+V）", _clipped == ["ABCD-1234"])
    check("登录成功后拿到 access_token", _tok.get("access_token") == "AT")

    _opened.clear()
    _logs.clear()
    graph_mod._device_login(
        config.from_dict({"graph": {"open_browser": False}}).graph, _logs.append, "Mail.Read")
    check("open_browser=false 时不弹浏览器（但日志里照样给码）",
          _opened == [] and "ABCD-1234" in "\n".join(_logs))

    check("有 verification_uri_complete 时优先用它",
          graph_mod._sign_in_url({"verification_uri": "https://a/device",
                                  "verification_uri_complete": "https://a/device?otc=ZZZ"},
                                 "CODE") == ("https://a/device?otc=ZZZ", True))
    check("没有该字段时把设备码附在 user_code 参数上",
          graph_mod._sign_in_url({"verification_uri": "https://a/device"}, "AB-12")
          == ("https://a/device?user_code=AB-12", True))
    check("连设备码都没有时只给基础网址",
          graph_mod._sign_in_url({"verification_uri": "https://a/device"}, "")
          == ("https://a/device", False))
finally:
    graph_mod._post_form = _saved_post
    graph_mod.webbrowser.open = _saved_open
    graph_mod._copy_to_clipboard = _saved_clip

print("[18] 分级规则：学院邮件必须稳定为 A（不再时 A 时 C）")
_terms_cfg = config.from_dict({"faculty": "School of Science / 理学院",
                               "faculty_keywords": ["SCI", "ab"]})
check("faculty_terms 按分隔符拆开、丢掉过短的词",
      _terms_cfg.faculty_terms == ["School of Science", "理学院", "SCI"])

_fac_cfg = config.from_dict({"faculty": "School of Science"})
_fac_mail = {"subject": "Recruitment Talk", "sender": "advise@ust.hk",
             "sender_name": "SCI/HOME, School of Science", "body": "short body"}
_other_mail = {"subject": "HKUST Daily Event Alert", "sender": "calendar@ust.hk",
               "sender_name": "HKUST Event Calendar", "body": "events"}

_v = [classifier.Verdict(level="C", is_school=True),
      classifier.Verdict(level="B", is_school=True)]
_n = classifier.apply_faculty_override([_fac_mail, _other_mail], _v, _fac_cfg)
check("学院发来的邮件从 C 提升为 A", _v[0].level == "A" and _n == 1)
check("非学院的校级邮件不受影响", _v[1].level == "B")

_v = [classifier.Verdict(level="S", is_school=True)]
classifier.apply_faculty_override([_fac_mail], _v, _fac_cfg)
check("已经是 S 的学院邮件不被降级", _v[0].level == "S")

_v = [classifier.Verdict(level="C", is_school=False)]
classifier.apply_faculty_override([_fac_mail], _v, _fac_cfg)
check("非学校邮件不参与该规则（保持不写回）",
      _v[0].level == "C" and _v[0].is_school is False)

_by_signature = [{"subject": "x", "sender": "noreply@example.com", "sender_name": "Office",
                  "body": "Best regards, School of Science, HKUST"}]
_v = [classifier.Verdict(level="C", is_school=True)]
classifier.apply_faculty_override(_by_signature, _v, _fac_cfg)
check("显示名看不出、但签名里的学院名同样能命中", _v[0].level == "A")

_v = [classifier.Verdict(level="C", is_school=True)]
check("没填学院时该规则不生效",
      classifier.apply_faculty_override([_fac_mail], _v, config.from_dict({})) == 0
      and _v[0].level == "C")

_block = classifier._email_block(1, {"subject": "Hi", "sender": "advise@ust.hk",
                                     "sender_name": "SCI/HOME, School of Science",
                                     "body": "H" * 500 + " SIGNATURE School of Science"})
check("提示词里带上发件人显示名（学院身份藏在这里）",
      "SCI/HOME, School of Science" in _block)
check("长正文会附上末尾节选（签名在那里）", "SIGNATURE" in _block)

_spec_empty = classifier.build_spec(config.from_dict({"faculty": "", "name": "",
                                                      "student_id": "", "courses": []}))
check("空配置不会把规则写成「(未填写)」这种废话", "(未填写)" not in _spec_empty)
check("没填学院时规则里不出现学院条件", "我所在的学院" not in _spec_empty)
_spec_full = classifier.build_spec(config.from_dict({"faculty": "School of Science"}))
check("填了学院时规则里出现学院条件", "School of Science" in _spec_full)

print("[19] 连接日志不刷屏 + 批次失败自动拆分")
_saved_get_token = graph_mod.get_token
_saved_settings = graph_mod._settings
_saved_me = graph_mod.GraphMail.me
_conn_logs: list = []
try:
    graph_mod.get_token = lambda *a, **kw: "TOKEN"
    graph_mod._settings = lambda: config.from_dict({}).graph
    graph_mod.GraphMail.me = lambda self: {"displayName": "Tester",
                                           "userPrincipalName": "t@example.com"}
    for _ in range(3):
        graph_mod.connect(auto_launch=False, log=_conn_logs.append)
    check("连续三次连接只打一行 connected as（监听每 30 秒一轮，不能刷屏）",
          len([line for line in _conn_logs if "connected as" in line]) == 1)
finally:
    graph_mod.get_token = _saved_get_token
    graph_mod._settings = _saved_settings
    graph_mod.GraphMail.me = _saved_me

_saved_ask = ai.ask_json
_batch_sizes: list = []
_split_logs: list = []


def _flaky_ask(prompt, ai_cfg, log=print, **kwargs):
    """4 封及以上就整批失败 —— 模拟"输出额度被撑爆"那种整批作废。

    期望的拆分阶梯（batch_size=8）：8 失败 → 4+4 失败 → 2+2+2+2 成功。
    """
    size = prompt.count("\n发件人: ")
    _batch_sizes.append(size)
    if size >= 4:
        return None
    return {"items": [{"学校": True, "级别": "C", "截止日期": None}] * size}


try:
    ai.ask_json = _flaky_ask
    _eight = [{"subject": f"m{i}", "sender": "a@ust.hk", "body": "x"} for i in range(8)]
    _got = classifier.classify(_eight, config.from_dict({"courses": [], "batch_size": 8}),
                               log=_split_logs.append)
    check("整批失败后对半拆分重试",
          _batch_sizes[:1] == [8] and _batch_sizes.count(4) == 2 and _batch_sizes.count(2) == 4)
    check("拆开后每一封最终都拿到了结果", all(v is not None for v in _got))
    check("日志里说明了在拆分", any("retrying as" in line for line in _split_logs))
finally:
    ai.ask_json = _saved_ask

print("[20] 分类主列表建不出来时：只提示一次，判级与写回照常")
_saved_blocked = graph_mod._category_create_blocked


class _Cat403Client:
    """模拟「只有 Mail.ReadWrite、没有 MailboxSettings.ReadWrite」的登录。"""

    def master_categories(self):
        return [{"displayName": "Red category"}]

    def ensure_master_category(self, name):
        raise graph_mod.GraphError("HTTP 403 ErrorAccessDenied: Access is denied.", 403)


class _CatHaveClient:
    def __init__(self):
        self.attempts = []

    def master_categories(self):
        return [{"displayName": n} for n in ("S", "A", "B", "C", "Sum")]

    def ensure_master_category(self, name):
        self.attempts.append(name)
        return False


try:
    graph_mod._category_create_blocked = False
    _c403 = _Cat403Client()
    _ns403 = graph_mod.GraphSession(_c403)
    _first = graph_mod.ensure_categories(_ns403)
    check("缺权限时只返回一条聚合提示（不再是 4 行报错）",
          len(_first) == 1 and "cannot create" in _first[0])
    check("提示里说明这是正常现象（该权限多数租户要管理员批准）",
          "administrator approval" in _first[0] and "Approval required" in _first[0])
    check("提示里给出最省事的做法（自己在 Outlook 里建 5 个分类）",
          "All Categories" in _first[0] and "Sum" in _first[0])
    check("提示里仍然提到所需权限名，方便查证",
          "MailboxSettings.ReadWrite" in _first[0])
    check("之后不再反复尝试（不会每 30 秒刷屏）",
          graph_mod.ensure_categories(_ns403) == [])
    graph_mod._category_create_blocked = False
    check("Sum 分类建不出来时不抛异常、只返回 False",
          graph_mod.ensure_sum_category(_ns403) is False)
    check("Sum 分类失败后也进入「不再重试」状态", graph_mod._category_create_blocked is True)

    graph_mod._category_create_blocked = False
    _have = _CatHaveClient()
    check("四个分类都在时不尝试创建",
          graph_mod.ensure_categories(graph_mod.GraphSession(_have)) == []
          and _have.attempts == [])
finally:
    graph_mod._category_create_blocked = _saved_blocked

print(f"\n全部通过：{ok} 项")
print("\n--- build_spec 实际产出（配置缺字段时的样子）---")
print(classifier.build_spec(config.from_dict({"email": "a@b.c"})))
