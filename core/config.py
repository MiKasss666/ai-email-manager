"""core/config.py — 配置层。

单一职责：把 config.json 变成一个经过校验、可直接使用的 Config 对象。
本模块不认识 Outlook，也不认识具体某个 AI 服务（只知道它们的配置长什么样）。

四条原则（针对旧版的问题）：
  · 只保留被代码真正使用的字段 —— 不做「以后可能用」的预留开关。
  · 缺失字段补默认值，非法值回落到安全区间，所以上层无需写防御性判断。
  · 不做全局缓存 —— 配置由调用方读取一次后逐层传入，避免缓存失效类 bug。
  · 纯标准库。
"""
from __future__ import annotations

import json
import os
import re
import sys
from dataclasses import dataclass, field
from typing import Dict, List, Optional

CONFIG_NAME = "config.json"

# 四个级别（业务概念）。颜色属于 Outlook，见 outlook.CATEGORY_COLORS。
LEVELS = ("S", "A", "B", "C")

# AI 后端类型（ai.py 的调用分发必须覆盖这里列出的每一种）
#   openai    OpenAI 兼容 /chat/completions —— DeepSeek、Kimi、通义、硅基流动、OpenRouter、Ollama、vLLM…
#   anthropic Anthropic 原生 /v1/messages
#   cli       本地命令行（提示词写进 stdin，输出当答案）—— claude、ollama run、任意 CLI
PROVIDER_TYPES = ("openai", "anthropic", "cli")

# 出厂默认值。注意：这些常量必须同时被 Config 数据类和 from_dict() 使用 ——
# 之前 language / ai_active 的默认值在两处各写了一份，改了数据类却被 from_dict
# 里的字面量架空，实际行为没变。默认值只允许有一个出处。
DEFAULT_LANGUAGE = "en"
DEFAULT_AI_ACTIVE = "deepseek"

DEFAULT_COURSE_CODE_RE = r"[A-Z]{2,4}\d{3,4}[A-Z]?"

# 学校特化的出厂默认值：本程序面向香港科技大学学生开发，所以默认填好港科大的信息，
# 其他学校的用户改 config.json 即可（这些都是「学校属性」，不是某个人的隐私）。
# 注意：真正的个人信息（邮箱、姓名、学号、你自己的选课列表）一律不写进代码，
# 由用户首次运行时自己填写。
DEFAULT_SCHOOL = "香港科技大学 (HKUST)"
DEFAULT_DORM_KEYWORDS = [
    "UG Hall", "UGHall", "UGHousing", "Hall", "宿舍", "宿生",
    "Student Housing", "Residence", "Dorm",
]


def app_dir() -> str:
    """应用根目录：打包后取 exe 所在目录，否则取 core 的上一级目录。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def config_path() -> str:
    return os.path.join(app_dir(), CONFIG_NAME)


def log(*args) -> None:
    """输出日志；隐藏控制台（pythonw）时 stdout 为 None，这里兜住。"""
    if sys.stdout is None:
        return
    try:
        print(*args)
    except Exception:
        pass


# ── 类型归一 ──────────────────────────────────────────────────────
def _as_int(value, default: int, low: int, high: int) -> int:
    try:
        v = int(value)
    except (TypeError, ValueError):
        return default
    return max(low, min(high, v))


def _as_float(value, default: float, low: float, high: float) -> float:
    try:
        v = float(value)
    except (TypeError, ValueError):
        return default
    return max(low, min(high, v))


def _as_bool(value, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in ("1", "true", "yes", "on", "y", "是", "开")


def _as_str(value, default: str = "") -> str:
    if value is None:
        return default
    s = str(value).strip()
    return s or default


def _as_list(value) -> List[str]:
    """列表或「逗号/分号/顿号/换行/竖线」分隔的字符串 → 去空白去重且保序的列表。"""
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        items = [str(v) for v in value]
    else:
        s = str(value)
        for sep in (";", "；", "、", "\n", "|"):
            s = s.replace(sep, ",")
        items = s.split(",")
    out, seen = [], set()
    for it in items:
        v = it.strip()
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def _normalize_language(value) -> str:
    """把语言值规范成 zh / en；识别不出来时回落到默认值。"""
    text = _as_str(value, DEFAULT_LANGUAGE).lower()
    if text.startswith("en"):
        return "en"
    if text.startswith("zh"):
        return "zh"
    return DEFAULT_LANGUAGE


def _as_headers(value) -> Dict[str, str]:
    """自定义请求头：只接受「键值都是字符串」的字典。"""
    if not isinstance(value, dict):
        return {}
    return {str(k): str(v) for k, v in value.items() if str(k).strip()}


# ── AI 后端配置 ───────────────────────────────────────────────────
@dataclass
class AiProfile:
    """一个 AI 后端。用户可以配多个，用 ai.active 选一个生效。"""

    name: str = ""
    type: str = "openai"          # openai / anthropic / cli
    base_url: str = ""            # 例：https://api.deepseek.com（官方 OpenAI 兼容地址，不带 /v1）
    model: str = ""               # 例：deepseek-flash
    api_key: str = ""             # 直接写在配置里（不推荐公开分享的机器这么用）
    api_key_env: str = ""         # 推荐：从环境变量读，密钥不落盘
    api_key_required: bool = True  # 本地服务（Ollama / LM Studio 等）不需要 Key，设为 false
    temperature: float = 0.0      # 判级/摘要都要稳定，默认 0
    # 输出上限。默认给得很宽（实测 DeepSeek 接受到 32768）：推理模型会把「思考」也计入
    # 这个预算，太小的话正文（JSON 数组）会被写到一半掐断 —— 表现是"整批要点变空"。
    # 完整性优先于省钱，所以这里不设紧。
    max_tokens: int = 16384
    timeout: int = 300            # 秒
    json_mode: bool = True        # 请求 response_format=json_object；不支持会自动降级重试
    extra_headers: Dict[str, str] = field(default_factory=dict)  # Azure / OpenRouter 等特殊头
    cli_command: str = ""         # type=cli 时用；留空则自动找 claude
    missing: bool = False         # 内部：ai.active 指向了不存在的 profile

    # ── 密钥 ──────────────────────────────────────────────────
    def key(self) -> str:
        """API Key：优先环境变量（推荐，密钥不落盘），其次配置里的明文。"""
        if self.api_key_env:
            from_env = (os.environ.get(self.api_key_env) or "").strip()
            if from_env:
                return from_env
        return (self.api_key or "").strip()

    def key_source(self) -> str:
        """给自检打印用：说明当前 Key 是从哪儿取到的。"""
        if self.api_key_env and (os.environ.get(self.api_key_env) or "").strip():
            return f"env var {self.api_key_env}"
        if (self.api_key or "").strip():
            return "config.json api_key"
        return "not set"

    def problems(self) -> List[str]:
        label = f'AI backend "{self.name}"'
        errs: List[str] = []

        if self.type not in PROVIDER_TYPES:
            return [f"{label} has an unknown type {self.type!r} "
                    f"(supported: {' / '.join(PROVIDER_TYPES)})"]
        if self.timeout <= 0:
            errs.append(f"{label}: timeout must be greater than 0")
        if self.type == "cli":
            return errs  # 命令行后端不需要 url / model / key

        if not self.base_url:
            errs.append(f"{label} has no base_url")
        if not self.model:
            errs.append(f"{label} has no model")
        if not self.key() and self.api_key_required:
            hint = (f'set the environment variable {self.api_key_env}'
                    if self.api_key_env else "fill in api_key or api_key_env")
            errs.append(f"{label} has no usable API key ({hint})")
        return errs


def _default_ai_profiles() -> Dict[str, AiProfile]:
    """出厂模板：常见后端各给一份，改 ai.active 选一个、填上 Key 就能用。

    注意：这里的 model 名只是示例，服务商的模型名会变，请以官方文档为准。
    """
    return {
        "claude-cli": AiProfile(name="claude-cli", type="cli"),
        # DeepSeek 官方给出的 OpenAI 兼容地址是 https://api.deepseek.com（不带 /v1），
        # 模型名用 deepseek-flash（旧的 deepseek-chat 已不在官方模型列表里）。
        # 官方还提供 Anthropic 格式：base_url 填 https://api.deepseek.com/anthropic
        "deepseek": AiProfile(name="deepseek", type="openai",
                              base_url="https://api.deepseek.com",
                              model="deepseek-flash", api_key_env="DEEPSEEK_API_KEY"),
        "openai": AiProfile(name="openai", type="openai",
                            base_url="https://api.openai.com/v1",
                            model="gpt-4o-mini", api_key_env="OPENAI_API_KEY"),
        "anthropic": AiProfile(name="anthropic", type="anthropic",
                               base_url="https://api.anthropic.com",
                               model="claude-sonnet-4-5", api_key_env="ANTHROPIC_API_KEY"),
        "ollama": AiProfile(name="ollama", type="openai",
                            base_url="http://localhost:11434/v1", model="qwen2.5:7b",
                            api_key_required=False),
    }


def ai_profile_from_dict(name: str, raw) -> AiProfile:
    raw = raw if isinstance(raw, dict) else {}
    return AiProfile(
        name=name,
        type=_as_str(raw.get("type"), "openai").lower(),
        base_url=_as_str(raw.get("base_url")),
        model=_as_str(raw.get("model")),
        api_key=_as_str(raw.get("api_key")),
        api_key_env=_as_str(raw.get("api_key_env")),
        api_key_required=_as_bool(raw.get("api_key_required"), True),
        temperature=_as_float(raw.get("temperature"), 0.0, 0.0, 2.0),
        max_tokens=_as_int(raw.get("max_tokens"), 16384, 0, 200000),
        timeout=_as_int(raw.get("timeout"), 300, 1, 3600),
        json_mode=_as_bool(raw.get("json_mode"), True),
        extra_headers=_as_headers(raw.get("extra_headers")),
        cli_command=_as_str(raw.get("cli_command")),
    )


def ai_profile_to_dict(profile: AiProfile) -> dict:
    return {
        "type": profile.type,
        "base_url": profile.base_url,
        "model": profile.model,
        "api_key": profile.api_key,
        "api_key_env": profile.api_key_env,
        "api_key_required": profile.api_key_required,
        "temperature": profile.temperature,
        "max_tokens": profile.max_tokens,
        "timeout": profile.timeout,
        "json_mode": profile.json_mode,
        "extra_headers": dict(profile.extra_headers),
        "cli_command": profile.cli_command,
    }


# ── 主配置 ────────────────────────────────────────────────────────
@dataclass
class Config:
    """用户配置。字段与 config.json 一一对应。"""

    # 身份：用于判定「这封邮件与我有关」
    email: str = ""
    name: str = ""
    student_id: str = ""
    school: str = DEFAULT_SCHOOL    # 预填港科大；其他学校改这里
    faculty: str = ""               # 各人学院不同，留空由用户填
    courses: List[str] = field(default_factory=list)

    # 输出与总结
    language: str = DEFAULT_LANGUAGE     # zh / en —— 只影响总结报告的文案（界面固定英文）
    summary_recipient: str = ""          # 留空则发给 email 自己
    summary_days: int = 7                # 总结涵盖最近多少天

    # 分级行为
    batch_size: int = 40                 # 每次送给模型的邮件数
    urgent_days: int = 30                # deadline 在多少天内算「加急」（A 升 S）
    catchup_days: int = 7                # 启动时补分类最近 N 天的未分类邮件（0 = 关闭）
    dorm_keywords: List[str] = field(default_factory=lambda: list(DEFAULT_DORM_KEYWORDS))
    course_code_regex: str = DEFAULT_COURSE_CODE_RE
    faculty_keywords: List[str] = field(default_factory=list)

    # AI 后端
    ai_active: str = DEFAULT_AI_ACTIVE
    ai_profiles: Dict[str, AiProfile] = field(default_factory=_default_ai_profiles)

    # ── 派生属性 ──────────────────────────────────────────────
    @property
    def is_english(self) -> bool:
        return self.language == "en"

    @property
    def recipient(self) -> str:
        """总结邮件收件人：未单独配置时回落到本人邮箱。"""
        return self.summary_recipient or self.email

    @property
    def faculty_terms(self) -> List[str]:
        """「来源 = 学院」的判定词：faculty 本身 + 额外关键词。"""
        return ([self.faculty] if self.faculty else []) + list(self.faculty_keywords)

    @property
    def ai(self) -> AiProfile:
        """当前生效的 AI 后端。

        ai.active 写错时返回一个带 missing 标记的占位对象，让 problems() 和
        ai.ask() 都能给出「名字写错了、现有的是哪些」这种明确提示，
        而不是在这里抛异常把 check 也一起带崩。
        """
        profile = self.ai_profiles.get(self.ai_active)
        if profile is not None:
            return profile
        return AiProfile(name=self.ai_active, missing=True)

    # ── 校验 ──────────────────────────────────────────────────
    def identity_problems(self) -> List[str]:
        """身份 / 输出 / 判级相关的问题（User info 表单负责这部分）。"""
        errs: List[str] = []
        if "@" not in self.email:
            errs.append("email is empty or not a valid address (must contain @)")
        if self.summary_recipient and "@" not in self.summary_recipient:
            errs.append("summary_recipient is not a valid address (must contain @)")
        for label, value in (("summary_days", self.summary_days),
                             ("batch_size", self.batch_size),
                             ("urgent_days", self.urgent_days)):
            if value <= 0:
                errs.append(f"{label} must be greater than 0 (currently {value})")
        try:
            re.compile(self.course_code_regex)
        except re.error as ex:
            errs.append(f"course_code_regex is not a valid regex: {ex}")
        return errs

    def ai_problems(self) -> List[str]:
        """AI 后端相关的问题（AI 设置面板负责这部分）。"""
        if self.ai.missing:
            return [f'ai.active = "{self.ai_active}" does not exist in ai.profiles '
                    f"(available: {', '.join(self.ai_profiles) or 'none'})"]
        return self.ai.problems()

    def problems(self) -> List[str]:
        """全部阻塞性问题；空列表表示可以正常运行。"""
        return self.identity_problems() + self.ai_problems()


# ── 字典 ↔ Config ─────────────────────────────────────────────────
def from_dict(raw: dict) -> Config:
    raw = raw if isinstance(raw, dict) else {}
    # 这两个字段要区分「键不存在」（补学校默认值）和「显式留空」（尊重用户的清空操作）
    dorm = _as_list(raw.get("dorm_keywords")) if "dorm_keywords" in raw else list(DEFAULT_DORM_KEYWORDS)
    school = _as_str(raw.get("school")) if "school" in raw else DEFAULT_SCHOOL

    ai_raw = raw.get("ai") if isinstance(raw.get("ai"), dict) else {}
    profiles_raw = ai_raw.get("profiles") if isinstance(ai_raw.get("profiles"), dict) else None
    if profiles_raw:
        profiles = {str(name): ai_profile_from_dict(str(name), value)
                    for name, value in profiles_raw.items()}
    else:
        profiles = _default_ai_profiles()

    return Config(
        email=_as_str(raw.get("email")),
        name=_as_str(raw.get("name")),
        student_id=_as_str(raw.get("student_id")),
        school=school,
        faculty=_as_str(raw.get("faculty")),
        courses=_as_list(raw.get("courses")),
        language=_normalize_language(raw.get("language")),
        summary_recipient=_as_str(raw.get("summary_recipient")),
        summary_days=_as_int(raw.get("summary_days"), 7, 1, 365),
        batch_size=_as_int(raw.get("batch_size"), 40, 1, 200),
        urgent_days=_as_int(raw.get("urgent_days"), 30, 1, 365),
        # 0 是合法值（= 关闭启动补分类），所以下限是 0
        catchup_days=_as_int(raw.get("catchup_days"), 7, 0, 365),
        dorm_keywords=dorm,
        course_code_regex=_as_str(raw.get("course_code_regex"), DEFAULT_COURSE_CODE_RE),
        faculty_keywords=_as_list(raw.get("faculty_keywords")),
        ai_active=_as_str(ai_raw.get("active"), DEFAULT_AI_ACTIVE),
        ai_profiles=profiles,
    )


def to_dict(cfg: Config) -> dict:
    return {
        "email": cfg.email,
        "name": cfg.name,
        "student_id": cfg.student_id,
        "school": cfg.school,
        "faculty": cfg.faculty,
        "courses": list(cfg.courses),
        "language": cfg.language,
        "summary_recipient": cfg.summary_recipient,
        "summary_days": cfg.summary_days,
        "batch_size": cfg.batch_size,
        "urgent_days": cfg.urgent_days,
        "catchup_days": cfg.catchup_days,
        "dorm_keywords": list(cfg.dorm_keywords),
        "course_code_regex": cfg.course_code_regex,
        "faculty_keywords": list(cfg.faculty_keywords),
        "ai": {
            "active": cfg.ai_active,
            "profiles": {name: ai_profile_to_dict(profile)
                         for name, profile in cfg.ai_profiles.items()},
        },
    }


# ── 读写 ──────────────────────────────────────────────────────────
def load(path: Optional[str] = None) -> Config:
    """读配置。文件不存在或损坏时回落到默认值并说明原因，不抛异常。"""
    path = path or config_path()
    if not os.path.exists(path):
        return Config()
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            return from_dict(json.load(f))
    except Exception as ex:
        log(f"[config] could not parse {path}, using defaults: {ex}")
        return Config()


def save(cfg: Config, path: Optional[str] = None) -> str:
    """原子写入（先写 .tmp 再替换），避免中途失败留下半个文件。"""
    path = path or config_path()
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(to_dict(cfg), f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    return path


def ensure(path: Optional[str] = None) -> Config:
    """启动时调用：保证 config.json 存在，并返回读到的配置。"""
    path = path or config_path()
    exists = os.path.exists(path)
    cfg = load(path)
    if not exists:
        save(cfg, path)
        log(f"[config] created default config: {path}")
    return cfg
