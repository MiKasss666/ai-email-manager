"""core/graph.py — 唯一的邮件后端：Microsoft Graph。

不依赖 Outlook、不需要 COM、只使用 Python 标准库（urllib / json / webbrowser）。

能力：
  · 登录：OAuth2 设备码。token 缓存里同时保存 **access_token 与过期时间**，
    未过期就静默复用（不再每轮都去刷新 refresh token）；refresh token 失效才重登。
    网络错误与「凭据失效」严格区分 —— 断网不会把你推进重新登录流程。
  · 读信：轮询只拉邮件头；**正文按需单封下载**（只有真正要判级/总结的邮件才付这份成本）。
  · 限流：429 / 5xx 读 Retry-After 退避重试，并归类为「邮箱暂时不可用」，不计失败。
  · 写回：只改邮件自己的 categories/flag，**不碰 Outlook 的分类主列表**
    （不创建、不删除、不读取 masterCategories —— 那需要额外权限且常被租户拦）。
"""
from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from email.utils import parsedate_to_datetime
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from . import config as config_mod
from .config import LEVELS
from .mailmeta import (
    FLAG_MARKED,
    FLAG_NONE,
    SUM_NAME,
    MailboxUnavailable,
    category_names,
    has_our_category,
    is_report_subject,
    merge_categories,
    message_info,
    note_once,
    parse_sum_command,
)

GRAPH = "https://graph.microsoft.com/v1.0"
LOGIN = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0"
DEVICE_GRANT = "urn:ietf:params:oauth:grant-type:device_code"
TIMEOUT = 30

# 登录缓存里 access_token 的提前刷新余量（秒）：避免"刚好在请求途过期"。
TOKEN_MARGIN = 300

# 微软自家公开客户端（无密钥 public client）。只在用户没填 client_id 时兜底，**仅适合自己用**；
# 要分发给别人请在 entra.microsoft.com 注册自己的应用，并把 id 填进 config.json -> graph.client_id。
FALLBACK_CLIENTS: Tuple[Tuple[str, str], ...] = (
    ("14d82eec-204b-4c2f-b7e8-296a70dab67e", "Microsoft Graph Command Line Tools"),
    ("04b07795-8ddb-461a-bbee-02f9e1bf7b46", "Azure CLI"),
    ("1950a258-227b-4e31-a9cf-717495945fc2", "Azure PowerShell"),
)

# 邮件头：轮询/列表/报告识别都只需要这些，不含 body（体积差一个数量级）。
SELECT_HEADERS = ("id,subject,from,sender,receivedDateTime,categories,isRead,hasAttachments,"
                  "bodyPreview,conversationId,webLink")
SELECT_BODY = "body,bodyPreview"

RECENT_MAX = 200          # 一次最多取多少封（收件箱扫描窗口）
RECENT_TTL = 10           # 同一轮监听里多个查询共用一次请求

RETRY_STATUS = (429, 500, 502, 503, 504)
MAX_ATTEMPTS = 3


class GraphError(RuntimeError):
    """Graph 调用失败。status 是 HTTP 状态码（0 = 网络层失败）。

    transient=True 表示「过一会儿重试可能就好了」（限流 / 5xx / 网络中断），
    上层据此退避，而不是把它当成一次失败去累计。
    """

    def __init__(self, message: str, status: int = 0):
        super().__init__(message)
        self.status = status

    @property
    def transient(self) -> bool:
        return self.status == 0 or self.status in RETRY_STATUS


# ── 设置与 token 缓存 ─────────────────────────────────────────────
def _settings():
    """取 config.json 里的 graph 段（读不到就用默认值）。"""
    try:
        return config_mod.load().graph
    except Exception:
        return config_mod.GraphConfig()


def token_path(gcfg=None) -> str:
    gcfg = gcfg or _settings()
    name = (gcfg.token_cache or "graph_token.json").strip() or "graph_token.json"
    if os.path.isabs(name):
        return name
    return os.path.join(config_mod.app_dir(), name)


def load_cache(gcfg=None) -> Dict[str, Any]:
    try:
        with open(token_path(gcfg), "r", encoding="utf-8") as handle:
            data = json.load(handle)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_cache(data: Dict[str, Any], gcfg=None) -> None:
    """原子写（先 .tmp 再 replace）：崩溃时不会留下半个凭据文件。"""
    path = token_path(gcfg)
    try:
        config_mod.write_json_atomic(path, data)
        try:
            os.chmod(path, 0o600)      # Windows 上基本无效，但 Unix 下有意义
        except Exception:
            pass
    except Exception:
        pass


def cached_token(cache: Dict[str, Any], now: Optional[float] = None,
                 scopes: Optional[str] = None) -> Optional[str]:
    """缓存里未过期、且权限与本次请求一致的 access_token；否则 None。

    **必须比对 scopes**：只读自检（run.py check）只申请 Mail.Read，若把那个
    token 留在缓存里给写回用，PATCH 会 403。纯函数（now 可注入），便于单测。
    """
    token = str(cache.get("access_token") or "").strip()
    if not token:
        return None
    if scopes is not None and str(cache.get("scopes") or "") != scopes:
        return None
    try:
        expires_at = float(cache.get("expires_at") or 0)
    except (TypeError, ValueError):
        return None
    if not expires_at:
        return None
    return token if (now or time.time()) < expires_at - TOKEN_MARGIN else None


def logged_in() -> bool:
    """本地是否已有可用的登录凭据（只查缓存，不发请求）。"""
    cache = load_cache()
    return bool(cache.get("refresh_token") or cached_token(cache))


# ── HTTP 底座 ─────────────────────────────────────────────────────
def _decode(raw: bytes) -> Any:
    text = raw.decode("utf-8", "replace")
    try:
        return json.loads(text)
    except Exception:
        return {"raw": text[:600]}


def _http(method: str, url: str, *, data: Optional[bytes] = None,
          headers: Optional[Dict[str, str]] = None) -> Tuple[int, Any, Dict[str, str]]:
    request = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            return response.status, _decode(response.read()), dict(response.headers)
    except urllib.error.HTTPError as ex:
        try:
            payload = _decode(ex.read())
        except Exception:
            payload = {"raw": ""}
        return ex.code, payload, dict(ex.headers or {})
    except Exception as ex:                     # DNS / TLS / 网络中断
        return 0, {"raw": f"{type(ex).__name__}: {ex}"}, {}


def _post_form(url: str, fields: Dict[str, str]) -> Tuple[int, Any]:
    body = urllib.parse.urlencode(fields).encode()
    status, payload, _headers = _http("POST", url, data=body,
                                      headers={"Content-Type": "application/x-www-form-urlencoded"})
    return status, payload


def _explain(status: int, payload: Any) -> str:
    if isinstance(payload, dict):
        err = payload.get("error")
        if isinstance(err, dict):
            return f"HTTP {status} {err.get('code') or ''}: {str(err.get('message') or '')[:400]}"
        if payload.get("raw"):
            return f"HTTP {status}: {payload['raw']}"[:400]
    return f"HTTP {status}: {str(payload)[:300]}"


def _retry_delay(status: int, headers: Dict[str, str], attempt: int) -> float:
    """这次失败后等多久；优先用服务端给的 Retry-After。"""
    value = (headers or {}).get("Retry-After") or (headers or {}).get("retry-after") or ""
    text = str(value).strip()
    try:
        seconds = float(text)                  # 常见形式：秒数
        if 0 <= seconds <= 120:
            return seconds
    except (TypeError, ValueError):
        pass
    try:                                       # RFC 7231 的 HTTP-date 形式
        when = parsedate_to_datetime(text)
        if when is not None:
            seconds = (when - datetime.now(timezone.utc)).total_seconds()
            if 0 <= seconds <= 120:
                return float(seconds)
    except (TypeError, ValueError):
        pass
    return float(min(30, 2 ** attempt))        # 1s、2s …（封顶 30s）


def graph_call(method: str, path_or_url: str, token: str, *,
               body: Optional[dict] = None) -> Any:
    """调 Graph。GET 遇到限流/5xx/网络错误会按 Retry-After 退避重试；
    写操作**不自动重试**（sendMail 这类不是幂等的），失败抛带 transient 标记的 GraphError。

    非 2xx 一律抛 GraphError —— 调用方可以用 ex.transient 决定「退避重试」
    还是「当成真错误」。
    """
    url = path_or_url if path_or_url.startswith("http") else GRAPH + path_or_url
    headers = {
        "Authorization": "Bearer " + token,
        # 邮件 id 保持稳定（移动文件夹后不变）—— 监听基线依赖它
        "Prefer": 'IdType="ImmutableId"',
        "Accept": "application/json",
    }
    data = None
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"

    attempts = MAX_ATTEMPTS if method.upper() == "GET" else 1
    last_status, last_payload = 0, None
    for attempt in range(attempts):
        status, payload, response_headers = _http(method, url, data=data, headers=headers)
        if 200 <= status < 300:
            return payload
        last_status, last_payload = status, payload
        retryable = status == 0 or status in RETRY_STATUS
        if not retryable or attempt == attempts - 1:
            break
        time.sleep(_retry_delay(status, response_headers, attempt))
    raise GraphError(_explain(last_status, last_payload), last_status)


# ── 登录 ──────────────────────────────────────────────────────────
def _copy_to_clipboard(text: str) -> bool:
    """尽力把设备码放进剪贴板（用 clip.exe，不碰 Tk 的剪贴板）。"""
    try:
        import subprocess
        proc = subprocess.run("clip", input=text.encode("utf-8"), shell=True,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
        return proc.returncode == 0
    except Exception:
        return False


def _sign_in_url(device: Dict[str, Any], code: str) -> Tuple[str, bool]:
    """给浏览器用的登录链接。返回 (url, 是否已带上设备码)。"""
    base = device.get("verification_uri") or "https://login.microsoft.com/device"
    complete = device.get("verification_uri_complete")
    if complete:
        return str(complete), True
    if not code:
        return base, False
    joiner = "&" if "?" in base else "?"
    return f"{base}{joiner}user_code={urllib.parse.quote(code)}", True


def _device_login(gcfg, log: Callable[..., None], scopes: str) -> Dict[str, Any]:
    """设备码登录：把设备码打进日志 + 自动打开浏览器。"""
    tenant = (gcfg.tenant or "common").strip() or "common"

    candidates: List[Tuple[str, str]] = []
    if (gcfg.client_id or "").strip():
        candidates.append((gcfg.client_id.strip(), "config.json 里的 client_id"))
    candidates.extend(FALLBACK_CLIENTS)

    device = None
    client_id = label = ""
    for cid, name in candidates:
        status, payload = _post_form(LOGIN.format(tenant=tenant) + "/devicecode",
                                     {"client_id": cid, "scope": scopes})
        if status == 200 and isinstance(payload, dict) and payload.get("device_code"):
            client_id, label, device = cid, name, payload
            break
        log(f"[graph] client {name} is not usable: {_explain(status, payload)[:150]}")
    if device is None:
        raise MailboxUnavailable(
            "Could not obtain a device code: this tenant may block third-party apps or the "
            "device-code flow.\n"
            "        Register an app at https://entra.microsoft.com and put its client id "
            "into config.json -> graph.client_id")

    code = device.get("user_code") or ""
    sign_in, carries_code = _sign_in_url(device, code)
    log("")
    log("[graph] ================ Microsoft sign-in required ================")
    log(f"[graph] Device code: {code}")        # 单独一行，方便在日志区复制
    log(f"[graph] Open this page: {sign_in}")
    log(f"[graph] Permissions requested: {scopes}")
    log(f"[graph] Signing in with app: {label}")
    log("[graph] ===========================================================")
    if _copy_to_clipboard(code):
        log("[graph] Device code copied to the clipboard - just press Ctrl+V on that page")
    if getattr(gcfg, "open_browser", True):
        try:
            webbrowser.open(sign_in)
            hint = " (the link already carries the device code)" if carries_code else ""
            log(f"[graph] Opened your browser{hint}; if it did not appear, "
                f"open the page above yourself")
        except Exception as ex:
            log(f"[graph] Could not open a browser ({type(ex).__name__}) - "
                f"please open the page above yourself")
    log("[graph] This window continues by itself once you finish signing in ...")
    log("")

    interval = int(device.get("interval") or 5)
    deadline = time.time() + int(device.get("expires_in") or 900)
    started = last_notice = time.time()
    network_errors = 0
    while time.time() < deadline:
        time.sleep(interval)
        now = time.time()
        if now - last_notice >= 20:            # 等用户操作时给心跳，别让界面看起来像卡死
            last_notice = now
            log(f"[graph] still waiting for you to finish signing in ... "
                f"{int(now - started)}s (device code {code}, "
                f"expires in {int(deadline - now)}s)")
        status, payload = _post_form(LOGIN.format(tenant=tenant) + "/token", {
            "grant_type": DEVICE_GRANT, "client_id": client_id,
            "device_code": device["device_code"]})
        if status == 200 and isinstance(payload, dict) and payload.get("access_token"):
            payload["_client_id"] = client_id
            log(f"[graph] signed in ({label})")
            return payload
        if status == 0:
            network_errors += 1
            if network_errors >= 3:
                raise MailboxUnavailable(
                    f"Cannot reach the Microsoft sign-in service: {_explain(status, payload)}")
            continue
        network_errors = 0
        error = (payload or {}).get("error")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        raise MailboxUnavailable(
            f"Sign-in failed: {error} - "
            f"{str((payload or {}).get('error_description'))[:300]}\n"
            f"        If it says admin approval is required, this tenant does not allow "
            f"users to consent; register your own Entra app or contact your IT department.")
    raise MailboxUnavailable("The device code expired after 15 minutes - start the app again "
                             "to get a new one")


# 这些 OAuth 错误码表示「凭据真的不能用了」，需要重新登录；
# 其它情况（尤其是网络错误）不该把用户推进登录流程。
_DEAD_CREDENTIAL_ERRORS = {"invalid_grant", "interaction_required", "invalid_request",
                           "bad_token", "expired_token", "unauthorized_client"}


def _refresh(gcfg, cache: Dict[str, Any], log: Callable[..., None],
             scopes: str) -> Optional[Dict[str, Any]]:
    """用 refresh token 换新 token。

    返回新 token（含旋转后的 refresh_token）；**凭据失效**返回 None（调用方去重新登录）；
    **网络/服务端故障**抛 MailboxUnavailable（绝不能因此让用户重新登录）。
    """
    refresh_token, client_id = cache.get("refresh_token"), cache.get("client_id")
    if not refresh_token or not client_id:
        return None
    tenant = (gcfg.tenant or "common").strip() or "common"
    status, payload = _post_form(LOGIN.format(tenant=tenant) + "/token", {
        "grant_type": "refresh_token", "client_id": client_id,
        "refresh_token": refresh_token, "scope": scopes})
    if status == 200 and isinstance(payload, dict) and payload.get("access_token"):
        payload["_client_id"] = client_id
        return payload
    if status == 0 or status >= 500:
        raise MailboxUnavailable(
            f"Cannot reach the Microsoft sign-in service to renew the credential: "
            f"{_explain(status, payload)[:200]}")
    error = str((payload or {}).get("error") or "")
    if status == 400 and error not in _DEAD_CREDENTIAL_ERRORS:
        raise MailboxUnavailable(
            f"The credential could not be renewed ({_explain(status, payload)[:200]}); "
            f"will retry later")
    log(f"[graph] the cached credential has expired ({_explain(status, payload)[:120]}); "
        f"signing in again")
    return None


def _store_token(token: Dict[str, Any], scopes: str, gcfg) -> None:
    """把 token 落盘（原子写）：access_token + 过期时间 + 旋转后的 refresh_token。"""
    try:
        expires_in = int(token.get("expires_in") or 3600)
    except (TypeError, ValueError):
        expires_in = 3600
    cache = load_cache(gcfg)
    save_cache({
        "access_token": token.get("access_token"),
        "expires_at": time.time() + max(60, expires_in),
        "refresh_token": token.get("refresh_token") or cache.get("refresh_token"),
        "client_id": token.get("_client_id") or cache.get("client_id"),
        "scopes": scopes,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }, gcfg)


def get_token(gcfg=None, log: Callable[..., None] = print, *, write: bool = True) -> str:
    """拿 access_token：**缓存优先**，过期才刷新；refresh 失效才走设备码登录。"""
    gcfg = gcfg or _settings()
    scopes = (gcfg.scopes if write else gcfg.probe_scopes) or gcfg.scopes
    cache = load_cache(gcfg)

    token = cached_token(cache, scopes=scopes)
    if token:
        return token

    refreshed = _refresh(gcfg, cache, log, scopes)      # 可能抛 MailboxUnavailable
    if refreshed is None:
        refreshed = _device_login(gcfg, log, scopes)
    _store_token(refreshed, scopes, gcfg)
    return refreshed["access_token"]


# ── 邮件对象 ─────────────────────────────────────────────────────
def _strip_html(html: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?</\1>", " ", html or "")
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(p|div|tr|li|h[1-6])>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    for entity, char in (("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"),
                         ("&gt;", ">"), ("&quot;", '"'), ("&#39;", "'")):
        text = text.replace(entity, char)
    text = re.sub(r"[ \t\u00a0]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _address_of(value: Any) -> Tuple[str, str]:
    if isinstance(value, dict):
        inner = value.get("emailAddress")
        if isinstance(inner, dict):
            return str(inner.get("name") or ""), str(inner.get("address") or "")
        return str(value.get("name") or ""), str(value.get("address") or "")
    if isinstance(value, str):
        return "", value
    return "", ""


def _body_text(raw: Any) -> str:
    """Graph 的 body 字段 → 纯文本（HTML 会被剥成文本）。"""
    if isinstance(raw, dict):
        content = str(raw.get("content") or "")
        is_html = (raw.get("contentType") or "").lower() == "html"
    elif isinstance(raw, str):
        content, is_html = raw, True
    else:
        return ""
    return _strip_html(content) if is_html else content.strip()


def _received_at(value: Optional[str]):
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00")).astimezone() \
            .replace(tzinfo=None)
    except Exception:
        return None


def normalize(message: Dict[str, Any]) -> Dict[str, Any]:
    """Graph 的 message → 内部统一字典（字段名与 mailmeta.message_info 对齐）。

    **幂等**：已经是归一化过的字典就直接返回（少了这道保护，二次归一化会把 id
    变成空字符串，写回时 PATCH 会打到 /me/messages/ 上）。正文允许为空 ——
    Graph 后端按需下载正文，见 GraphSession.fill_body()。
    """
    if isinstance(message, dict) and "entryid" in message and "id" not in message:
        return message

    name, address = _address_of(message.get("from"))
    if not address:
        alt_name, alt_address = _address_of(message.get("sender"))
        name, address = name or alt_name, address or alt_address

    body = _body_text(message.get("body")) or str(message.get("bodyPreview") or "")
    categories = message.get("categories")
    return {
        "entryid": message.get("id") or "",
        "msg": message,
        "subject": message.get("subject") or "(no subject)",
        "sender_name": name,
        "sender": address,
        "body": body,
        "received": _received_at(message.get("receivedDateTime")),
        "categories": list(categories) if isinstance(categories, list) else [],
        "is_read": bool(message.get("isRead")),
        "web_link": str(message.get("webLink") or ""),
    }


class GraphMessage:
    """把一封 Graph 邮件包装成「像邮件对象」的东西。

    mailmeta 里的纯逻辑（message_info / category_names / is_report / parse_sum_command）
    都按属性取用，所以这一层只要把属性摆好。
    """

    def __init__(self, raw: Dict[str, Any], client: "GraphMail"):
        info = normalize(raw)
        self._client = client
        self._raw = raw
        self.id = info["entryid"]
        self.EntryID = info["entryid"]
        self.Subject = info["subject"]
        self.SenderName = info["sender_name"]
        self.SenderEmailAddress = info["sender"]
        self.Body = info["body"]
        self.body_loaded = bool(info["body"])
        self.ReceivedTime = info["received"]
        self.Categories = ", ".join(info["categories"])
        self.UnRead = not info["is_read"]
        self.FlagStatus = (raw.get("flag") or {}).get("flagStatus")
        self.WebLink = info["web_link"]
        self.level_hint = next((c for c in info["categories"] if c in LEVELS), None)

    def __repr__(self) -> str:
        return f"<GraphMessage {self.ReceivedTime} {self.Subject[:40]!r}>"


class GraphMail:
    """一个登录后的邮箱。方法都很薄：一次调用对应一个 Graph 端点。"""

    def __init__(self, token: str):
        self.token = token

    def me(self) -> Dict[str, Any]:
        return graph_call("GET", "/me?$select=displayName,userPrincipalName,mail,id", self.token)

    def inbox_raw(self, top: int = 10, *, days: Optional[int] = None,
                  select: str = SELECT_HEADERS) -> List[Dict[str, Any]]:
        """收件箱最新若干封的原始 Graph 条目（新 → 旧）。

        默认只取邮件头（不含正文）：轮询每 30 秒跑一次，正文留给真正要判级的邮件。
        """
        params = {"$top": str(int(top)), "$select": select,
                  "$orderby": "receivedDateTime desc"}
        if days:
            since = (datetime.now(timezone.utc) - timedelta(days=int(days))) \
                .strftime("%Y-%m-%dT%H:%M:%SZ")
            params["$filter"] = f"receivedDateTime ge {since}"
        # 必须 urlencode：`$orderby` 里有空格，直接拼进 URL 会被 urllib 拒绝
        payload = graph_call("GET", "/me/mailFolders/inbox/messages?"
                             + urllib.parse.urlencode(params), self.token)
        return list(payload.get("value") or [])

    def fetch_body(self, message_id: str) -> str:
        """按需下载一封邮件的正文（纯文本化后返回）；失败返回空串。"""
        if not message_id:
            return ""
        payload = graph_call(
            "GET", "/me/messages/" + urllib.parse.quote(message_id) + "?$select=" + SELECT_BODY,
            self.token)
        return _body_text((payload or {}).get("body")) or str((payload or {}).get("bodyPreview") or "")

    def inbox_folder(self) -> Dict[str, Any]:
        return graph_call(
            "GET", "/me/mailFolders/inbox?$select=displayName,totalItemCount,unreadItemCount",
            self.token)

    def set_categories(self, message_id: str, names: Sequence[str],
                       flag: Optional[int] = None) -> None:
        # 空 id 会打到 /me/messages/ 上 —— Graph 对它的 PATCH 回 405，报错还很难懂。
        if not (message_id or "").strip():
            raise GraphError("no message id to write to (the mail object lost its id)")
        body: Dict[str, Any] = {"categories": list(names)}
        if flag is not None:
            body["flag"] = {"flagStatus": "flagged" if flag == FLAG_MARKED else "notFlagged"}
        graph_call("PATCH", "/me/messages/" + urllib.parse.quote(message_id),
                   self.token, body=body)

    def send_mail(self, to: str, subject: str, html: str) -> None:
        graph_call("POST", "/me/sendMail", self.token, body={
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html},
                "toRecipients": [{"emailAddress": {"address": to}}],
            },
            "saveToSentItems": True,
        })


# ── 会话（一轮任务内复用查询结果）────────────────────────────────
class GraphSession:
    """一次连接。带 10 秒缓存：一轮监听里会连着问好几次，合并成一次请求。"""

    def __init__(self, client: GraphMail):
        self.client = client
        self._recent: List[GraphMessage] = []
        self._recent_at = 0.0
        self._recent_days: Optional[int] = None     # 缓存窗口；None = 服务器未按时间过滤

    def recent(self, *, days: Optional[int] = None, force: bool = False) -> List[GraphMessage]:
        """收件箱最新若干封（新 → 旧），带 10 秒缓存。

        缓存**按窗口区分**：一轮里先查 30 天再查 7 天时，不能把 30 天的结果当成
        7 天用（会让 catch-up 处理超出配置范围的邮件）；反过来，先查 7 天再要全量
        时也不能拿窄窗口冒充。窗口变宽就重取，变窄就在本地按时间过滤。
        """
        now = time.time()
        stale = force or (now - self._recent_at > RECENT_TTL)
        if not stale:
            if days is None:
                stale = self._recent_days is not None        # 缓存是窄窗口，这次要全量
            elif self._recent_days is not None and days > self._recent_days:
                stale = True                                 # 缓存窗口不够宽
        if stale:
            raw = self.client.inbox_raw(top=RECENT_MAX, days=days)
            self._recent = [GraphMessage(item, self.client) for item in raw]
            self._recent_at = now
            self._recent_days = days
        if days is None:
            return self._recent
        cutoff = datetime.now() - timedelta(days=days)
        return [msg for msg in self._recent
                if msg.ReceivedTime is None or msg.ReceivedTime >= cutoff]

    def fill_body(self, msg: GraphMessage) -> None:
        """按需把一封邮件的正文下载进来（幂等；失败只留摘要，不影响判级流程）。

        这是「轮询只拉邮件头」的配套：真正要被 AI 看的邮件才付正文的下载成本。
        已加载过就直接返回；下载失败时退回 bodyPreview，并标记完成避免反复重试。
        """
        if getattr(msg, "body_loaded", False):
            return
        try:
            text = self.client.fetch_body(msg.EntryID)
            if text:
                msg.Body = text
        except GraphError as ex:
            if ex.transient:
                # 邮箱暂时不可用：让整轮退避重来，别用残缺信息去判级
                raise MailboxUnavailable(f"could not load a message body: {ex}") from ex
        except Exception:
            pass
        msg.body_loaded = True

    # 判定"要不要处理"——与收件箱约定一致
    @staticmethod
    def _is_report(msg: GraphMessage) -> bool:
        return is_report_subject(msg.Subject or "")

    @staticmethod
    def _is_command(msg: GraphMessage) -> bool:
        return parse_sum_command(msg.Subject or "") is not None

    @staticmethod
    def _is_classified(msg: GraphMessage) -> bool:
        return has_our_category(msg)


# ── 端口实现 ──────────────────────────────────────────────────────
def connect(*, read_only: bool = False,
            log: Callable[..., None] = print) -> Tuple[Any, Any]:
    """连接邮箱，返回 (client, session)。

    · 缓存里的 access_token 未过期 → 直接复用，不发登录请求；
    · 需要重新登录时，日志里给出设备码 + 自动打开浏览器（设备码 15 分钟内有效）；
    · read_only=True（自检）只申请只读权限，免得为了看一眼配置就要写授权。

    网络/限流类失败抛 MailboxUnavailable（上层退避重试），不是「配置错误」。
    """
    gcfg = _settings()
    read_only = read_only or os.environ.get("EM_GRAPH_READ_ONLY") == "1"
    if read_only:
        log("[graph] read-only mode (Mail.Read only - nothing is written back)")
    token = get_token(gcfg, log, write=not read_only)
    client = GraphMail(token)
    try:
        me = client.me()
    except GraphError as ex:
        raise MailboxUnavailable(f"Graph could not read your account: {ex}") from ex
    # 例行连接日志每个进程只打一次：监听每 30 秒连一次，不然活动日志会被刷屏。
    note_once(log, "graph-connected",
              f"[graph] connected as {me.get('displayName')} "
              f"<{me.get('userPrincipalName') or me.get('mail')}>")
    return client, GraphSession(client)


def newest_message(ns: GraphSession) -> Optional[GraphMessage]:
    items = ns.recent()
    return items[0] if items else None


def _collect(ns: GraphSession, *, last_seen: Optional[str], days: Optional[int],
             max_items: int, skip_classified: bool,
             keep: Optional[Callable[[Any], bool]],
             log: Callable[..., None],
             with_body: bool) -> List[Dict[str, Any]]:
    """从最新往回扫，收集要处理的邮件（内部统一实现，行为由参数区分）。"""
    out: List[Dict[str, Any]] = []
    skipped = 0
    for msg in ns.recent(days=days):
        try:
            if last_seen and msg.EntryID == last_seen:
                break
            if GraphSession._is_report(msg) or GraphSession._is_command(msg):
                continue
            if skip_classified and GraphSession._is_classified(msg):
                continue
            if keep is not None and not keep(msg):
                skipped += 1
                continue
            if with_body:
                ns.fill_body(msg)             # 只有要送 AI 的邮件才下载正文
            out.append(message_info(msg))
        except MailboxUnavailable:
            raise                             # 邮箱不可用：整轮退避，不做半截判断
        except Exception as ex:               # 单封读不出来不该拖垮整轮
            log(f"  [skip] could not read a mail: {type(ex).__name__}: {ex}")
            continue
        if len(out) >= max_items:
            break
    if skipped:
        log(f"  skipped {skipped} mail(s) already ruled out of scope")
    return out


def read_inbox(ns: GraphSession, *, limit: Optional[int] = None, days: Optional[int] = None,
               skip_classified: bool = True,
               keep: Optional[Callable[[Any], bool]] = None,
               log: Callable[..., None] = print) -> List[Dict[str, Any]]:
    """读收件箱待分级邮件（新 → 旧），正文按需下载。"""
    return _collect(ns, last_seen=None, days=days,
                    max_items=limit if limit is not None else RECENT_MAX,
                    skip_classified=skip_classified, keep=keep, log=log, with_body=True)


def collect_new_mail(ns: GraphSession, last_seen: Optional[str], *, days: Optional[int] = None,
                     max_items: int = 100,
                     keep: Optional[Callable[[Any], bool]] = None,
                     log: Callable[..., None] = print) -> List[Dict[str, Any]]:
    """比 last_seen 更新的邮件（新 → 旧）；走到 last_seen 就停。

    正文按需下载：只有真正要被判级的邮件才付这份成本。
    """
    return _collect(ns, last_seen=last_seen, days=days, max_items=max_items,
                    skip_classified=True, keep=keep, log=log, with_body=True)


def read_recent(ns: GraphSession, *, days: int, letters: str = "sabc") -> List[Dict[str, Any]]:
    """读最近 N 天内已打 S/A/B/C 且级别在 letters 里的邮件（总结报告用）。

    正文按需下载（要点需要正文），返回按 S→A→B→C 排序（组内新 → 旧）。
    """
    wanted = "".join(char for char in "sabc" if char in (letters or "sabc").lower())
    rows: List[Dict[str, Any]] = []
    for msg in ns.recent(days=days):
        level = msg.level_hint
        if level is None or level.lower() not in wanted:
            continue
        ns.fill_body(msg)
        rows.append({
            "entryid": msg.EntryID, "level": level,
            "subject": msg.Subject or "(no subject)",
            "sender": msg.SenderName or msg.SenderEmailAddress or "",
            "body": msg.Body or "", "received": msg.ReceivedTime,
            "web_link": msg.WebLink or "",
        })
    rows.sort(key=lambda item: LEVELS.index(item["level"]))
    return rows


def pending_sum_commands(ns: GraphSession, last_seen: Optional[str], *,
                         days: Optional[int] = None,
                         max_items: int = 5) -> List[GraphMessage]:
    """「还没处理过」的 sum 指令邮件（新 → 旧，最多 max_items 封）。

    `plan_tick` 只看最新那一封；指令到达后 30 秒内又来了别的邮件，指令就不再是
    "最新"，于是永远轮不到执行。这里从最新往回扫到 last_seen 为止。
    判据：主题是合法指令 + 还没打 Sum 标签（标签即"已处理"标志）。
    """
    out: List[GraphMessage] = []
    for msg in ns.recent(days=days):
        if last_seen and msg.EntryID == last_seen:
            break
        if not GraphSession._is_command(msg):
            continue
        if SUM_NAME in category_names(msg):
            continue
        out.append(msg)
        if len(out) >= max_items:
            break
    return out


def untagged_reports(ns: GraphSession, *, days: Optional[int] = 30) -> List[GraphMessage]:
    """收件箱里「是总结报告、但还没打 Sum 标签」的邮件（新 → 旧）。

    单独找一遍：`plan_tick` 只看最新那一封，而 classify 路径又把报告当控制邮件跳过，
    于是报告之后只要来了别的邮件，它就永远轮不到打标签。
    """
    cutoff = datetime.now() - timedelta(days=days) if days else None
    out: List[GraphMessage] = []
    for msg in ns.recent(days=days):
        if not GraphSession._is_report(msg):
            continue
        if cutoff is not None and msg.ReceivedTime is not None and msg.ReceivedTime < cutoff:
            continue
        if SUM_NAME in category_names(msg):
            continue
        out.append(msg)
    return out


def write_categories(msg, names: Sequence[str], flag: Optional[int] = None) -> None:
    """写回分类：只替换本系统自己的标签，用户自己的分类原样保留。

    只改**这一封邮件**的 categories，不碰 Outlook 的分类主列表。
    """
    existing = category_names(msg)
    merged = merge_categories(existing, names)

    flag_value: Optional[int] = None
    if flag is not None:
        flag_value = FLAG_MARKED if flag == FLAG_MARKED else FLAG_NONE

    message_id = getattr(msg, "id", None) or getattr(msg, "EntryID", "")
    client = getattr(msg, "_client", None)
    if client is None:
        raise RuntimeError("write_categories 需要 GraphMessage（带 _client）")
    client.set_categories(message_id, merged, flag_value)
    msg.Categories = ", ".join(merged)          # 同一轮里的后续判断看到最新值
    msg.level_hint = next((name for name in merged if name in LEVELS), None)


def send_html(app, to: str, subject: str, html: str, categories: str = "") -> None:
    """发送 HTML 邮件（总结报告用）。

    categories 参数保留在签名里只为兼容调用方（报告靠主题标记识别，不依赖分类）。
    """
    app.send_mail(to, subject, html)
