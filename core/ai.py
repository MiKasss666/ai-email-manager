"""core/ai.py — AI 传输层（后端无关，零第三方依赖）。

唯一职责：把一段提示词发给「用户自己配置的模型」，拿回文本或 JSON。

支持三类后端（在 config.json 的 ai.profiles 里选，ai.active 指定生效的那个）：

  openai     任何 OpenAI 兼容的 /chat/completions
             —— DeepSeek、Kimi、通义、智谱、硅基流动、OpenRouter、Groq、Together、
                Azure OpenAI、以及本地 Ollama / LM Studio / vLLM …
  anthropic  Anthropic 原生 /v1/messages
  cli        本地命令行：提示词写进 stdin，stdout 当答案
             —— claude、`ollama run qwen2.5`、任何吃 stdin 的 CLI

只用标准库 urllib，所以不装任何 SDK、打包进 PyInstaller 也不会缺模块；
换服务商只是改配置，不需要改代码。

上层唯一的约定：ask() 输入提示词、输出文本，瞬时故障自己重试。
"""
from __future__ import annotations

import json
import os
import re
import shlex
import shutil
import subprocess
import threading
import time
import urllib.error
import urllib.request
from typing import Any, Callable, Dict, Optional, Tuple

from .config import AiProfile

# Claude Code CLI 的常见安装位置（npm 全局安装时 bin 下才是真 exe）
_CLI_CANDIDATES = (
    os.path.expandvars(r"%APPDATA%\npm\node_modules\@anthropic-ai\claude-code\bin\claude.exe"),
    os.path.expandvars(r"%APPDATA%\npm\claude.cmd"),
)

PROVIDER_LABELS = {
    "openai": "OpenAI-compatible API",
    "anthropic": "Anthropic Messages API",
    "cli": "local CLI",
}

ANTHROPIC_VERSION = "2023-06-01"
_MAX_ERROR_CHARS = 300


class AiError(RuntimeError):
    """AI 调用失败。transient=True 表示属于瞬时故障，重试可能成功。"""

    def __init__(self, message: str, transient: bool = False):
        super().__init__(message)
        self.transient = transient


# ── 对外：描述当前后端（自检用）──────────────────────────────────
def find_cli() -> str:
    """定位 claude 可执行文件；找不到时回落到 PATH 上的名字。"""
    for path in _CLI_CANDIDATES:
        if path and os.path.exists(path):
            return path
    return shutil.which("claude") or "claude"


def describe(ai_cfg: AiProfile) -> str:
    """把人话描述给 check 打印：用的什么、连哪儿、Key 从哪儿来。"""
    if ai_cfg.missing:
        return f'no backend profile named "{ai_cfg.name}"'
    kind = PROVIDER_LABELS.get(ai_cfg.type, ai_cfg.type)
    if ai_cfg.type == "cli":
        command = ai_cfg.cli_command or f"{find_cli()} -p --output-format text (auto-detected)"
        return f"{kind}: {command}"
    if ai_cfg.key():
        key_desc = f"key from {ai_cfg.key_source()} (ready)"
    elif ai_cfg.api_key_required:
        key_desc = f"key from {ai_cfg.key_source()} (missing)"
    else:
        key_desc = "no API key needed"
    return (f"{kind}: {ai_cfg.base_url or '(no base_url)'} / "
            f"model={ai_cfg.model or '(not set)'} / {key_desc}")


# ── 对外：主入口 ──────────────────────────────────────────────────
def _call_with_deadline(work: Callable[[], Any], seconds: int, what: str) -> Any:
    """跑 work()，最多等 `seconds` 秒；超时抛瞬时 AiError。

    为什么需要：`urllib` 的 timeout 只约束**单次 socket 操作**。实测遇到过一次
    「连接建好了、请求发出去了、服务器一直不回」——socket 超时始终没触发，程序
    卡在那个 `response.read()` 里 8 分钟以上（界面一直 Working …，日志里连一行
    `transient failure` 都没有）。这里用墙钟时间兜底，保证**任何**一次调用都有上限。

    超时后残留的取数线程是 daemon：socket 最终失败时它自己会结束，不会拖着进程。
    """
    box: Dict[str, Any] = {}

    def run() -> None:
        try:
            box["value"] = work()
        except BaseException as ex:          # 原样带回给调用方判断
            box["error"] = ex

    worker = threading.Thread(target=run, daemon=True, name="ai-call")
    worker.start()
    worker.join(seconds)
    if worker.is_alive():
        raise AiError(
            f"{what} did not answer within {seconds}s "
            f"(the connection was accepted but nothing came back — this is a problem on the "
            f"provider's side, not with Outlook)", transient=True)
    if "error" in box:
        raise box["error"]
    return box["value"]


def ask(prompt: str, ai_cfg: AiProfile, *, log: Callable[..., None] = print,
        retries: int = 3) -> str:
    """把提示词发给模型，返回原始输出文本。

    瞬时故障（限流 / 5xx / 超时 / 网络中断）按 5s、10s… 退避重试；
    重试用尽或属于配置类错误（401、404、model 名不对）时抛 AiError。

    每次尝试都有 `ai_cfg.timeout` 秒的**墙钟上限**（见 _call_with_deadline）——
    不能只靠 socket 超时，那种"连接挂着不回"的情况它兜不住。
    """
    if ai_cfg.missing:
        raise AiError(f'ai.active = "{ai_cfg.name}" does not exist. '
                      f"Check ai.profiles in config.json, or run: py -3 run.py check")
    call = _CALLERS.get(ai_cfg.type)
    if call is None:
        raise AiError(f"unknown AI backend type {ai_cfg.type!r} "
                      f"(supported: {' / '.join(sorted(_CALLERS))})")

    last: Optional[AiError] = None
    for attempt in range(1, retries + 1):
        try:
            return _call_with_deadline(lambda: call(prompt, ai_cfg), ai_cfg.timeout,
                                       f'AI backend "{ai_cfg.name}"')
        except AiError as ex:
            last = ex
            if not ex.transient or attempt == retries:
                break
            wait = 5 * attempt
            log(f"  [AI] transient failure, retrying in {wait}s "
                f"({attempt}/{retries - 1}): {ex}")
            time.sleep(wait)

    raise last if last is not None else AiError("AI call failed (unknown reason)")


def looks_truncated(text: str) -> bool:
    """粗判"写到一半被截断"：JSON 结构符号不配平。

    推理模型把「思考」也计入 max_tokens，思考一长正文就被掐断 —— 表现是数组少一个 `]`。
    有这个判断，日志就能直接说清是"输出额度用尽"而不是笼统的"不是合法 JSON"。
    """
    if not text:
        return False
    cleaned = re.sub(r"```(?:json)?", "", text)
    return (cleaned.count("[") > cleaned.count("]")
            or cleaned.count("{") > cleaned.count("}"))


def ask_json(prompt: str, ai_cfg: AiProfile, *, log: Callable[..., None] = print,
             retries: int = 2) -> Optional[Any]:
    """要求模型返回 JSON。解析失败就重试（模型偶发跑偏），最终失败返回 None。"""
    for attempt in range(1, retries + 1):
        try:
            text = ask(prompt, ai_cfg, log=log, retries=2)
        except AiError as ex:
            log(f"  [warn] {ex}")
            return None

        data = parse_json(text)
        if data is not None:
            return data

        hint = (" (the model ran out of output tokens - raise ai.profiles.<name>.max_tokens)"
                if looks_truncated(text) else "")
        log(f"  [warn] attempt {attempt}/{retries} did not return valid JSON{hint}, "
            f"first 500 chars:")
        log((text or "")[:500])
    return None


def salvage_array(text: str) -> Optional[List[Any]]:
    """把（可能被截断的）JSON 数组里**已经完整的元素**逐个解出来。

    截断几乎总是切在最后一个元素中间，前面的都是好的 —— 与其把一整批邮件留空，
    不如先用上已经拿到的部分，剩下的由调用方补（见 summarizer._points_batch）。

    用 raw_decode 从头逐元素推进，遇到解不动的（就是那个残缺元素）就停。
    """
    if not text:
        return None
    cleaned = re.sub(r"```(?:json)?", "", text)
    start = cleaned.find("[")
    if start == -1:
        return None

    decoder = json.JSONDecoder()
    out: List[Any] = []
    index = start + 1
    while index < len(cleaned):
        while index < len(cleaned) and cleaned[index] in " \t\r\n,":
            index += 1
        if index >= len(cleaned):
            break
        try:
            value, end = decoder.raw_decode(cleaned, index)
        except ValueError:
            break                      # 剩下的正是被截断的那个元素
        out.append(value)
        index = end
    return out or None


def ask_json_lenient(prompt: str, ai_cfg: AiProfile, *,
                     log: Callable[..., None] = print,
                     retries: int = 2) -> Optional[Any]:
    """要 JSON；解析成功就**原样返回**（形状由调用方判断），数组被截断时返回已完整的部分。

    和 ask_json 的分工：
      · ask_json        只有整体合法才返回 —— 判级那种「顺序必须严格对齐」的场景必须用它，
                        数量不对就该整批作废，绝不能错位。
      · ask_json_lenient 每个元素互相独立、可以逐个补 —— 一句话要点就是这种，
                        被截断时先收下前面完整的，剩下的再要一次，不必整批留空。

    注意这里**不检查形状**：模型可能把数组包一层（{"summaries": [...]} 之类），
    该由调用方（summarizer._as_string_list）去拆 —— 在这里过滤会把能用的结果丢掉。
    """
    for attempt in range(1, retries + 1):
        try:
            text = ask(prompt, ai_cfg, log=log, retries=2)
        except AiError as ex:
            log(f"  [warn] {ex}")
            return None

        data = parse_json(text)
        if data is not None:
            return data

        partial = salvage_array(text)
        if partial:
            log(f"  [warn] response looks truncated - recovered {len(partial)} complete item(s)")
            return partial

        hint = (" (the model ran out of output tokens - raise ai.profiles.<name>.max_tokens)"
                if looks_truncated(text) else "")
        log(f"  [warn] attempt {attempt}/{retries} did not return valid JSON{hint}, "
            f"first 500 chars:")
        log((text or "")[:500])
    return None


def probe(ai_cfg: AiProfile, *, log: Callable[..., None] = print) -> str:
    """发一个最小请求，验证后端是否真的可用（会消耗极少量额度）。"""
    return ask("Reply with exactly OK and nothing else.", ai_cfg, log=log, retries=1).strip()


def parse_json(text: str) -> Optional[Any]:
    """从模型输出里抠出 JSON。

    三种写法都要吃下：
      1. **整段就是合法 JSON** —— 必须最先试。否则下面「找第一个 [ 到最后一个 ]」的
         启发式会把 `{"a": [1, 2]}` 里那个内层数组当成整个结果，语义就漂了；
      2. 前后有解释文字或 ```json 围栏 → 按括号切片抠出来（老路径，保持不变）；
      3. **JSONL**：一行一个 JSON 对象。实测 deepseek-flash 在 json_object 模式下
         会把数组拆成一行一个对象写出来 —— 每行都合法、整体却不是合法 JSON，
         旧实现整批作废，白丢一次调用（日志里真的出现过）。
    """
    if not text:
        return None
    cleaned = re.sub(r"```(?:json)?", "", text).strip()
    if not cleaned:
        return None

    try:
        return json.loads(cleaned)
    except Exception:
        pass

    for opener, closer in (("[", "]"), ("{", "}")):
        start, end = cleaned.find(opener), cleaned.rfind(closer)
        if start == -1 or end <= start:
            continue
        try:
            return json.loads(cleaned[start:end + 1])
        except Exception:
            continue

    values = _decode_stream(cleaned)
    if values:
        return values if len(values) > 1 else values[0]
    return None


def _decode_stream(text: str) -> Optional[list]:
    """从第一个 `{` 或 `[` 开始，逐个解码连续出现的 JSON 值（JSONL 那种）。"""
    starts = [i for i in (text.find("{"), text.find("[")) if i != -1]
    if not starts:
        return None
    decoder = json.JSONDecoder()
    out = []
    index = min(starts)
    while index < len(text):
        while index < len(text) and text[index] in " \t\r\n,":
            index += 1
        if index >= len(text):
            break
        try:
            value, index = decoder.raw_decode(text, index)
        except ValueError:
            break                      # 剩下的正是写坏了的那一段
        out.append(value)
    return out or None


# ── HTTP 底座 ─────────────────────────────────────────────────────
def _endpoint(base_url: str, path: str) -> str:
    """拼接接口地址：容忍多余的斜杠、已写全的路径、以及带查询串的完整 URL。"""
    base = (base_url or "").strip().rstrip("/")
    if not base:
        return path
    if "?" in base or base.endswith(path):
        return base  # 用户给的是完整 URL（如 Azure 带 ?api-version=…），原样使用
    return base + path


def _anthropic_base(base_url: str) -> str:
    """Anthropic 的 base_url 只到域名或前缀，路径由调用方拼 /v1/messages。

    这里额外容忍两种常见写法：
      · 用户按 OpenAI 习惯写成 https://api.anthropic.com/v1 → 去掉末尾的 /v1；
      · DeepSeek 的 Anthropic 兼容入口是 https://api.deepseek.com/anthropic → 原样保留。
    """
    base = (base_url or "https://api.anthropic.com").strip().rstrip("/")
    return re.sub(r"/v1/?$", "", base)


def _post_json(url: str, payload: dict, headers: Dict[str, str], timeout: int) -> Tuple[int, str]:
    """POST JSON，返回 (状态码, 响应文本)。4xx/5xx 不抛异常，交给调用方判断。"""
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status, response.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as ex:      # 注意：HTTPError 是 URLError 的子类，必须先捕
        body = ""
        try:
            body = ex.read().decode("utf-8", "replace")
        except Exception:
            pass
        return ex.code, body
    except urllib.error.URLError as ex:
        raise AiError(f"network error or timeout: {ex.reason}", transient=True)
    except (TimeoutError, OSError) as ex:
        raise AiError(f"{type(ex).__name__}: {ex}", transient=True)


def _error_snippet(body: str) -> str:
    """从错误响应里抠出最有用的那句话。"""
    if not body:
        return ""
    try:
        data = json.loads(body)
    except Exception:
        return body.strip()[:_MAX_ERROR_CHARS]
    if isinstance(data, dict):
        error = data.get("error")
        if isinstance(error, dict) and error.get("message"):
            return str(error["message"])[:_MAX_ERROR_CHARS]
        if isinstance(error, str):
            return error[:_MAX_ERROR_CHARS]
        for key in ("message", "detail", "msg"):
            if data.get(key):
                return str(data[key])[:_MAX_ERROR_CHARS]
    return body.strip()[:_MAX_ERROR_CHARS]


def _raise_for_status(status: int, body: str, ai_cfg: AiProfile, url: str) -> None:
    """把 HTTP 状态码翻译成用户能自己动手解决的提示。"""
    if 200 <= status < 300:
        return

    if status in (401, 403):
        hint = f"API key rejected (currently from {ai_cfg.key_source()})"
    elif status == 404:
        hint = f"wrong base_url or model: {url} / model={ai_cfg.model!r}"
    elif status == 429:
        hint = "rate limited (retry later, or lower batch_size)"
    elif status == 400:
        hint = "request rejected (common causes: unknown model, unsupported parameter)"
    elif status >= 500:
        hint = "server error"
    else:
        hint = ""

    snippet = _error_snippet(body)
    message = f"HTTP {status}" + (f" - {hint}" if hint else "")
    if snippet:
        message += f"; response: {snippet}"
    raise AiError(message, transient=(status == 429 or status >= 500))


# ── 各后端实现 ────────────────────────────────────────────────────
def _flatten_content(content: Any) -> str:
    """模型返回的 content 可能是字符串，也可能是分块列表，统一成字符串。"""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if isinstance(block, str):
                parts.append(block)
            elif isinstance(block, dict) and isinstance(block.get("text"), str):
                parts.append(block["text"])
        return "".join(parts)
    return ""


def _call_openai(prompt: str, ai_cfg: AiProfile) -> str:
    url = _endpoint(ai_cfg.base_url, "/chat/completions")
    payload: Dict[str, Any] = {
        "model": ai_cfg.model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": ai_cfg.temperature,
    }
    if ai_cfg.max_tokens > 0:
        payload["max_tokens"] = ai_cfg.max_tokens
    if ai_cfg.json_mode:
        payload["response_format"] = {"type": "json_object"}

    headers = {"Content-Type": "application/json"}
    if ai_cfg.key():
        headers["Authorization"] = f"Bearer {ai_cfg.key()}"
    headers.update(ai_cfg.extra_headers)

    status, body = _post_json(url, payload, headers, ai_cfg.timeout)

    # 不少「兼容」实现并不支持 response_format，遇到 400 就去掉它再试一次
    if status == 400 and ai_cfg.json_mode:
        payload.pop("response_format", None)
        status, body = _post_json(url, payload, headers, ai_cfg.timeout)

    _raise_for_status(status, body, ai_cfg, url)

    try:
        data = json.loads(body)
    except Exception:
        raise AiError(f"response is not valid JSON: {body[:_MAX_ERROR_CHARS]}")
    choices = data.get("choices") or []
    if not choices:
        raise AiError(f"response has no choices: {body[:_MAX_ERROR_CHARS]}")
    message = choices[0].get("message") or {}

    text = _flatten_content(message.get("content"))
    if not text.strip():  # 推理模型有时把正文放在 reasoning_content
        text = _flatten_content(message.get("reasoning_content"))
    if not text.strip():
        raise AiError(f"model returned empty content: {body[:_MAX_ERROR_CHARS]}")
    return text


def _call_anthropic(prompt: str, ai_cfg: AiProfile) -> str:
    url = _endpoint(_anthropic_base(ai_cfg.base_url), "/v1/messages")

    payload = {
        "model": ai_cfg.model,
        "max_tokens": ai_cfg.max_tokens or 4096,   # Anthropic 必填
        "temperature": min(max(ai_cfg.temperature, 0.0), 1.0),
        "messages": [{"role": "user", "content": prompt}],
    }
    headers = {
        "Content-Type": "application/json",
        "x-api-key": ai_cfg.key(),
        "anthropic-version": ANTHROPIC_VERSION,
    }
    headers.update(ai_cfg.extra_headers)

    status, body = _post_json(url, payload, headers, ai_cfg.timeout)
    _raise_for_status(status, body, ai_cfg, url)

    try:
        data = json.loads(body)
    except Exception:
        raise AiError(f"response is not valid JSON: {body[:_MAX_ERROR_CHARS]}")
    text = _flatten_content(data.get("content"))
    if not text.strip():
        raise AiError(f"model returned empty content: {body[:_MAX_ERROR_CHARS]}")
    return text


def _call_cli(prompt: str, ai_cfg: AiProfile) -> str:
    if ai_cfg.cli_command.strip():
        # posix=False：Windows 路径里的反斜杠不被当转义符；再手动剥掉外层引号
        argv = [token.strip("\"'") for token in shlex.split(ai_cfg.cli_command, posix=False)]
    else:
        argv = [find_cli(), "-p", "--output-format", "text"]
    if not argv or not argv[0]:
        raise AiError("cli_command is empty and no working CLI backend was found")

    # 打包成 --windowed 的 exe 后没有控制台，命令行子进程会闪一个黑框，这里压掉
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0

    try:
        proc = subprocess.run(argv, input=prompt, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", timeout=ai_cfg.timeout,
                              creationflags=creationflags)
    except FileNotFoundError:
        raise AiError(f"command not found: {argv[0]!r} "
                      f"(set the full path in ai.profiles.<name>.cli_command)")
    except Exception as ex:
        raise AiError(f"{type(ex).__name__}: {ex}", transient=True)

    if proc.returncode != 0 or not (proc.stdout or "").strip():
        stderr = (proc.stderr or "").strip()
        blob = f"{proc.stdout or ''}\n{stderr}".lower()
        transient = any(k in blob for k in ("503", "overload", "429", "rate limit", "timed out", "timeout"))
        message = f"command exited with code {proc.returncode}" + (f"; stderr: {stderr[:300]}" if stderr else "")
        raise AiError(message, transient=transient)
    return proc.stdout


_CALLERS = {
    "openai": _call_openai,
    "anthropic": _call_anthropic,
    "cli": _call_cli,
}
