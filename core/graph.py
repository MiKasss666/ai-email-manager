"""core/graph.py — Microsoft Graph 邮件后端（不依赖 Outlook，也不需要 COM）。

和 `core/outlook.py` 提供**同一套端口**：

    connect / newest_message / read_inbox / collect_new_mail / read_recent /
    pending_sum_commands / untagged_reports / write_categories / send_html /
    ensure_categories / ensure_sum_category / existing_categories / our_category_colors

`core/mail.py` 按 `config.json` 的 `mail_source` 在其中之一做选择，所以
classifier / watcher / summarizer / 界面都不知道邮件是从哪来的。

登录（OAuth2 设备码）——按需求实现成：
  · 启动先看本地 token 缓存：能刷新就**静默复用**，平时完全无感；
  · 需要重新登录时，把**设备码打进日志**（界面日志区右上角有 Copy，可直接复制），
    并**自动打开浏览器**到登录页；用户输一次即可，之后长期免登录（refresh token 滚动续期）。

只用 Python 标准库（urllib / json / webbrowser），不引入 msal 之类的依赖。
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from datetime import datetime, timedelta, timezone
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

from . import config as config_mod
from .config import LEVELS

# 纯逻辑（不碰 COM）直接从 COM 后端复用 —— 分类约定、报告识别、sum 指令解析必须是
# 同一份，否则换个后端行为就漂了。
from .outlook import (                                            # noqa: F401
    LEGACY_CATEGORY_NAMES,
    OL_FLAG_MARKED,
    OL_FLAG_NONE,
    OL_FOLDER_INBOX,
    OUR_CATEGORY_NAMES,
    REPORT_SUBJECT_MARKS,
    SUM_COMMAND_RE,
    SUM_NAME,
    OutlookNotRunning,
    _notice,
    category_names,
    has_our_category,
    integrity_level,
    is_report,
    is_report_subject,
    message_info,
    parent_process,
    parse_sum_command,
    sender_address,
)

GRAPH = "https://graph.microsoft.com/v1.0"
LOGIN = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0"
DEVICE_GRANT = "urn:ietf:params:oauth:grant-type:device_code"
TIMEOUT = 30

# 微软自家公开客户端（无密钥的 public client）。只在用户没填 client_id 时兜底 ——
# 正式分发给别人用，请到 entra.microsoft.com 注册自己的应用并把 id 填进 config.json。
FALLBACK_CLIENTS: Tuple[Tuple[str, str], ...] = (
    ("14d82eec-204b-4c2f-b7e8-296a70dab67e", "Microsoft Graph Command Line Tools"),
    ("04b07795-8ddb-461a-bbee-02f9e1bf7b46", "Azure CLI"),
    ("1950a258-227b-4e31-a9cf-717495945fc2", "Azure PowerShell"),
)

SELECT = ("id,subject,from,sender,receivedDateTime,categories,isRead,hasAttachments,"
          "bodyPreview,body,conversationId,webLink")
RECENT_MAX = 200          # 一次最多取多少封（收件箱扫描窗口）
RECENT_TTL = 10           # 同一轮监听里多个查询共用一次请求


class GraphError(RuntimeError):
    """Graph 调用失败。status 是 HTTP 状态码（0 = 网络层失败）。"""

    def __init__(self, message: str, status: int = 0):
        super().__init__(message)
        self.status = status


# OutlookNotRunning 由 core/outlook.py 统一定义并在这里复用 —— 「邮箱暂时连不上」
# 只能有**一个**异常类，否则 watcher 里 `isinstance(ex, outlook.OutlookNotRunning)`
# 会随后端不同而认不出来（实测踩过）。


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
    path = token_path(gcfg)
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
        os.chmod(path, 0o600)
    except Exception:
        pass


def logged_in() -> bool:
    """本地是否已有可用的登录凭据（只查缓存，不发请求）。"""
    return bool(load_cache().get("refresh_token"))


# ── HTTP 底座 ─────────────────────────────────────────────────────
def _decode(raw: bytes) -> Any:
    text = raw.decode("utf-8", "replace")
    try:
        return json.loads(text)
    except Exception:
        return {"raw": text[:600]}


def _http(method: str, url: str, *, data: Optional[bytes] = None,
          headers: Optional[Dict[str, str]] = None) -> Tuple[int, Any]:
    request = urllib.request.Request(url, data=data, headers=headers or {}, method=method)
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            return response.status, _decode(response.read())
    except urllib.error.HTTPError as ex:
        return ex.code, _decode(ex.read())
    except Exception as ex:                     # DNS / TLS / 网络中断
        return 0, {"raw": f"{type(ex).__name__}: {ex}"}


def _post_form(url: str, fields: Dict[str, str]) -> Tuple[int, Any]:
    body = urllib.parse.urlencode(fields).encode()
    return _http("POST", url, data=body,
                 headers={"Content-Type": "application/x-www-form-urlencoded"})


def _explain(status: int, payload: Any) -> str:
    if isinstance(payload, dict):
        err = payload.get("error")
        if isinstance(err, dict):
            return f"HTTP {status} {err.get('code') or ''}: {str(err.get('message') or '')[:400]}"
        if payload.get("raw"):
            return f"HTTP {status}: {payload['raw']}"[:400]
    return f"HTTP {status}: {str(payload)[:300]}"


def graph_call(method: str, path_or_url: str, token: str, *,
               body: Optional[dict] = None) -> Any:
    url = path_or_url if path_or_url.startswith("http") else GRAPH + path_or_url
    headers = {
        "Authorization": "Bearer " + token,
        # 邮件 id 保持稳定（移动文件夹后不变）—— watcher 的基线依赖它
        "Prefer": 'IdType="ImmutableId"',
        "Accept": "application/json",
    }
    data = None
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    status, payload = _http(method, url, data=data, headers=headers)
    if status == 0 or status >= 400:
        raise GraphError(_explain(status, payload), status)
    return payload


# ── 登录 ──────────────────────────────────────────────────────────
def _copy_to_clipboard(text: str) -> bool:
    """尽力把设备码放进剪贴板，用户在网页里直接 Ctrl+V 即可。

    用 Windows 自带的 clip.exe，而不是 tkinter 的剪贴板：登录跑在后台线程里，
    从非主线程碰 Tk 是不安全的。设备码只含字母数字，编码无坑。
    """
    try:
        proc = subprocess.run("clip", input=text.encode("utf-8"), shell=True,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                              timeout=5)
        return proc.returncode == 0
    except Exception:
        return False


def _sign_in_url(device: Dict[str, Any], code: str) -> Tuple[str, bool]:
    """给浏览器用的登录链接。返回 (url, 是否已带上设备码)。

    · 优先用 RFC 8628 的 `verification_uri_complete`（点开即登录；微软目前不给这个字段）；
    · 退而求其次，把设备码作为 `user_code` 参数附上 —— 实测该参数在跳转
      （login.microsoft.com/device → login.microsoftonline.com/.../deviceauth）后会被保留，
      页面能拿到；万一将来失效，也只是退化成"手动输入"，不会出错。
    · 注意：`?code=` 会返回空白页，不能用。
    """
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
        raise GraphError(
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
    while time.time() < deadline:
        time.sleep(interval)
        now = time.time()
        # 心跳：这一步是在等**用户**操作，可能等几分钟。不吭声的话界面看起来像卡死，
        # 实测就有人等到 132 秒然后点了 Force reset。
        if now - last_notice >= 20:
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
        error = (payload or {}).get("error")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        raise GraphError(
            f"Sign-in failed: {error} - "
            f"{str((payload or {}).get('error_description'))[:300]}\n"
            f"        If it says admin approval is required, this tenant does not allow "
            f"users to consent; use the legacy backend instead "
            f"(config.json -> mail_source: \"outlook\")")
    raise GraphError("The device code expired after 15 minutes - start the app again "
                     "to get a new one")


def _refresh(gcfg, cache: Dict[str, Any], log: Callable[..., None],
             scopes: str) -> Optional[Dict[str, Any]]:
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
    log(f"[graph] the cached credential has expired ({_explain(status, payload)[:120]}); "
        f"signing in again")
    return None


def get_token(gcfg=None, log: Callable[..., None] = print, *, write: bool = True) -> str:
    """拿 access_token：缓存优先；没有或失效就走设备码登录。"""
    gcfg = gcfg or _settings()
    scopes = (gcfg.scopes if write else gcfg.probe_scopes) or gcfg.scopes
    token = _refresh(gcfg, load_cache(gcfg), log, scopes)
    if token is None:
        token = _device_login(gcfg, log, scopes)
    save_cache({"client_id": token.get("_client_id"),
                "refresh_token": token.get("refresh_token"),
                "scopes": scopes,
                "saved_at": datetime.now().isoformat(timespec="seconds")}, gcfg)
    return token["access_token"]


# ── 邮件字典（形状与 COM 后端一致）────────────────────────────────
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


def normalize(message: Dict[str, Any]) -> Dict[str, Any]:
    """Graph 的 message → 内部统一字典（字段名与 COM 版一致）。

    真实邮箱形状并不统一（`body` 可能是对象也可能是字符串、`from` 同理），
    这里按"能拿到多少算多少"处理：一封形状奇怪的邮件不该让整轮挂掉。

    **幂等**：已经是归一化过的字典就直接返回。少了这道保护会出事 ——
    归一化后的字段叫 `entryid` 而不是 `id`，再归一化一次 id 就变成空字符串，
    写回时 PATCH 会打到 `/me/messages/`（HTTP 405）。（实测踩过：所有写回全废。）
    """
    if isinstance(message, dict) and "entryid" in message and "id" not in message:
        return message

    name, address = _address_of(message.get("from"))
    if not address:
        alt_name, alt_address = _address_of(message.get("sender"))
        name, address = name or alt_name, address or alt_address

    body = message.get("body")
    if isinstance(body, dict):
        content = str(body.get("content") or "")
        is_html = (body.get("contentType") or "").lower() == "html"
    elif isinstance(body, str):
        content, is_html = body, True
    else:
        content, is_html = "", False
    if is_html:
        content = _strip_html(content)
    if not content:
        content = str(message.get("bodyPreview") or "")

    received = message.get("receivedDateTime")
    when = None
    if received:
        try:
            when = datetime.fromisoformat(received.replace("Z", "+00:00")).astimezone() \
                .replace(tzinfo=None)
        except Exception:
            when = None
    categories = message.get("categories")
    return {
        "entryid": message.get("id") or "",
        "msg": message,
        "subject": message.get("subject") or "(no subject)",
        "sender_name": name,
        "sender": address,
        "body": content,
        "received": when,
        "categories": list(categories) if isinstance(categories, list) else [],
        "is_read": bool(message.get("isRead")),
        # 微软给的「在 Outlook 网页版里打开这封邮件」链接，总结报告的标题指向它。
        # 经典版 Outlook 没有能深链到某封邮件的 URL 协议（`outlook:` 只认 inbox/calendar
        # 这类固定命令，带 EntryID 会被直接忽略），所以点标题打开的是网页版。
        "web_link": str(message.get("webLink") or ""),
    }


class GraphMessage:
    """把一封 Graph 邮件包装成"像 COM MailItem"的对象。

    这样 core.outlook 里那些只按属性取用的纯逻辑（message_info / category_names /
    is_report / parse_sum_command）可以原样复用。
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
        self.ReceivedTime = info["received"]
        self.Categories = ", ".join(info["categories"])
        self.UnRead = not info["is_read"]
        self.FlagStatus = (raw.get("flag") or {}).get("flagStatus")
        self.web_link = info["web_link"]
        self.level_hint = next((c for c in info["categories"] if c in LEVELS), None)

    def __repr__(self) -> str:
        return f"<GraphMessage {self.ReceivedTime} {self.Subject[:40]!r}>"


class GraphMail:
    """一个登录后的邮箱。方法都很薄，和 COM 版逐条对得上。"""

    def __init__(self, token: str):
        self.token = token

    def me(self) -> Dict[str, Any]:
        return graph_call("GET", "/me?$select=displayName,userPrincipalName,mail,id", self.token)

    def inbox_raw(self, top: int = 10, *, days: Optional[int] = None,
                  select: str = SELECT) -> List[Dict[str, Any]]:
        """收件箱最新若干封的**原始** Graph 条目（新 → 旧）。

        内部读邮件一律用这个 + GraphMessage —— GraphMessage 需要原始条目里的 `id`
        和 `receivedDateTime`；喂它归一化后的字典会把这些字段弄丢。
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

    def inbox(self, top: int = 10, *, days: Optional[int] = None,
              select: str = SELECT) -> List[Dict[str, Any]]:
        """归一化后的收件箱列表（命令行/诊断脚本用）。"""
        return [normalize(item)
                for item in self.inbox_raw(top=top, days=days, select=select)]

    def categories_of(self, message_id: str) -> List[str]:
        payload = graph_call("GET", "/me/messages/"
                             + urllib.parse.quote(message_id) + "?$select=categories", self.token)
        return list(payload.get("categories") or [])

    def master_categories(self) -> List[Dict[str, Any]]:
        payload = graph_call("GET", "/me/outlook/masterCategories", self.token) or {}
        return list(payload.get("value") or [])

    def inbox_folder(self) -> Dict[str, Any]:
        return graph_call(
            "GET", "/me/mailFolders/inbox?$select=displayName,totalItemCount,unreadItemCount",
            self.token)

    def set_categories(self, message_id: str, names: Sequence[str],
                       flag: Optional[int] = None) -> None:
        # 空 id 会打到 /me/messages/ 上 —— Graph 对它的 PATCH 回 405，报错还很难懂。
        # 宁可在这里明确失败，也不要发出一个注定被拒的请求。
        if not (message_id or "").strip():
            raise GraphError("no message id to write to (the mail object lost its id)")
        body: Dict[str, Any] = {"categories": list(names)}
        if flag is not None:
            body["flag"] = {"flagStatus": "flagged" if flag else "notFlagged"}
        graph_call("PATCH", "/me/messages/" + urllib.parse.quote(message_id),
                   self.token, body=body)

    def ensure_master_category(self, name: str) -> bool:
        if name in {row.get("displayName") for row in self.master_categories()}:
            return False
        graph_call("POST", "/me/outlook/masterCategories", self.token,
                   body={"displayName": name})
        return True

    def send_mail(self, to: str, subject: str, html: str) -> None:
        graph_call("POST", "/me/sendMail", self.token, body={
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": html},
                "toRecipients": [{"emailAddress": {"address": to}}],
            },
            "saveToSentItems": True,
        })


# ── 会话（相当于 COM 的 ns）───────────────────────────────────────
class GraphSession:
    """一次连接。带 10 秒缓存：一轮监听里会连着问好几次，合并成一次请求。"""

    def __init__(self, client: GraphMail):
        self.client = client
        self._recent: List[GraphMessage] = []
        self._recent_at = 0.0

    def recent(self, *, days: Optional[int] = None, force: bool = False) -> List[GraphMessage]:
        now = time.time()
        if force or now - self._recent_at > RECENT_TTL:
            # 必须用原始条目：GraphMessage 要靠里面的 id / receivedDateTime
            raw = self.client.inbox_raw(top=RECENT_MAX, days=days)
            self._recent = [GraphMessage(item, self.client) for item in raw]
            self._recent_at = now
        return self._recent

    def GetDefaultFolder(self, folder_id: int = OL_FOLDER_INBOX):
        """只为 `run.py check` 的打印服务：返回一个"像 COM 文件夹"的小对象。"""
        info = self.client.inbox_folder()

        class _Items:
            Count = info.get("totalItemCount") or 0

        class _Folder:
            Name = info.get("displayName") or "Inbox"
            Items = _Items()

        return _Folder()

    # 判定"要不要处理"——与 COM 版规则一致
    @staticmethod
    def _is_report(msg: GraphMessage) -> bool:
        return is_report_subject(msg.Subject or "")

    @staticmethod
    def _is_command(msg: GraphMessage) -> bool:
        return SUM_COMMAND_RE.match((msg.Subject or "").strip()) is not None

    @staticmethod
    def _is_classified(msg: GraphMessage) -> bool:
        names = [c.strip() for c in (msg.Categories or "").split(",") if c.strip()]
        return any(c in LEVELS or c == SUM_NAME for c in names)


# ── 端口实现 ──────────────────────────────────────────────────────
def connect(*, auto_launch: bool = True,
            log: Callable[..., None] = print) -> Tuple[Any, Any]:
    """连接邮箱，返回 (app, ns)。

    · 已经有凭据 → 静默复用，不打扰用户；
    · 没有凭据 → 日志里给出设备码 + 自动打开浏览器，等用户登录（设备码 15 分钟内有效）。
    auto_launch=False（只读自检）时只申请只读权限，免得为了看一眼配置就要写授权。
    """
    gcfg = _settings()
    read_only = (not auto_launch) or os.environ.get("EM_GRAPH_READ_ONLY") == "1"
    if read_only:
        log("[graph] read-only mode (Mail.Read only - nothing is written back)")
    try:
        token = get_token(gcfg, log, write=not read_only)
    except GraphError as ex:
        raise OutlookNotRunning(f"Graph connection failed: {ex}") from ex
    client = GraphMail(token)
    try:
        me = client.me()
    except GraphError as ex:
        raise OutlookNotRunning(f"Graph could not read your account: {ex}") from ex
    # 例行连接日志每个进程只打一次：监听每 30 秒连一次，不然活动日志会被
    # "connected as ..." 刷屏（实测截图上连着刷了 7 行），有用的事件被挤掉。
    # 复用 COM 后端那套抑制机制，日志语义保持一致。
    _notice(log, "graph-connected",
            f"[graph] connected as {me.get('displayName')} "
            f"<{me.get('userPrincipalName') or me.get('mail')}>")
    return client, GraphSession(client)


def newest_message(ns: GraphSession) -> Optional[GraphMessage]:
    items = ns.recent()
    return items[0] if items else None


def read_inbox(ns: GraphSession, *, limit: Optional[int] = None, days: Optional[int] = None,
               skip_classified: bool = True,
               keep: Optional[Callable[[Any], bool]] = None,
               log: Callable[..., None] = print) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    skipped = 0
    for msg in ns.recent(days=days):
        if GraphSession._is_report(msg) or GraphSession._is_command(msg):
            continue
        if skip_classified and GraphSession._is_classified(msg):
            continue
        if keep is not None and not keep(msg):
            skipped += 1
            continue
        out.append(message_info(msg))
        if limit is not None and len(out) >= limit:
            break
    if skipped:
        log(f"  skipped {skipped} mail(s) already ruled out of scope")
    return out


def collect_new_mail(ns: GraphSession, last_seen: Optional[str], *, days: Optional[int] = None,
                     max_items: int = 100,
                     keep: Optional[Callable[[Any], bool]] = None,
                     log: Callable[..., None] = print) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    skipped = 0
    for msg in ns.recent(days=days):
        if last_seen and msg.EntryID == last_seen:
            break
        if GraphSession._is_report(msg) or GraphSession._is_command(msg):
            continue
        if GraphSession._is_classified(msg):
            continue
        if keep is not None and not keep(msg):
            skipped += 1
            continue
        out.append(message_info(msg))
        if len(out) >= max_items:
            break
    if skipped:
        log(f"  skipped {skipped} mail(s) already ruled out of scope")
    return out


def read_recent(ns: GraphSession, *, days: int, letters: str = "sabc") -> List[Dict[str, Any]]:
    wanted = "".join(c for c in "sabc" if c in (letters or "sabc").lower())
    rows: List[Dict[str, Any]] = []
    for msg in ns.recent(days=days):
        level = msg.level_hint
        if level is None or level.lower() not in wanted:
            continue
        rows.append({
            "entryid": msg.EntryID, "level": level,
            "subject": msg.Subject or "(no subject)",
            "sender": msg.SenderName or msg.SenderEmailAddress or "",
            "body": msg.Body or "", "received": msg.ReceivedTime,
            "web_link": msg.web_link,
        })
    rows.sort(key=lambda item: LEVELS.index(item["level"]))
    return rows


def pending_sum_commands(ns: GraphSession, last_seen: Optional[str], *, days: Optional[int] = None,
                         max_items: int = 5) -> List[GraphMessage]:
    out: List[GraphMessage] = []
    for msg in ns.recent(days=days):
        if last_seen and msg.EntryID == last_seen:
            break
        if not GraphSession._is_command(msg):
            continue
        if SUM_NAME in [c.strip() for c in (msg.Categories or "").split(",")]:
            continue
        out.append(msg)
        if len(out) >= max_items:
            break
    return out


def untagged_reports(ns: GraphSession, *, days: Optional[int] = 30) -> List[GraphMessage]:
    out: List[GraphMessage] = []
    for msg in ns.recent(days=days):
        if not GraphSession._is_report(msg):
            continue
        if SUM_NAME in [c.strip() for c in (msg.Categories or "").split(",")]:
            continue
        out.append(msg)
    return out


def write_categories(msg, names: Sequence[str], flag: Optional[int] = None) -> None:
    """写回分类：只替换本系统自己的标签，用户自己的分类原样保留（与 COM 版同语义）。"""
    existing = [c.strip() for c in (getattr(msg, "Categories", "") or "").split(",") if c.strip()]
    merged = [c for c in existing if c not in OUR_CATEGORY_NAMES] + list(names)

    flag_value: Optional[int] = None
    if flag is not None:
        flag_value = OL_FLAG_MARKED if flag == OL_FLAG_MARKED else OL_FLAG_NONE

    message_id = getattr(msg, "id", None) or getattr(msg, "EntryID", "")
    client = getattr(msg, "_client", None)
    if client is None:
        raise RuntimeError("write_categories 需要 GraphMessage（带 _client）")
    client.set_categories(message_id, merged, flag_value)
    msg.Categories = ", ".join(merged)          # 同一轮里的后续判断看到最新值


def send_html(app, to: str, subject: str, html: str, categories: str = "") -> None:
    app.send_mail(to, subject, html)


# ── 分类维护 ──────────────────────────────────────────────────────
def existing_categories(ns: GraphSession) -> List[str]:
    return [row.get("displayName") or "" for row in ns.client.master_categories()]


def our_category_colors(ns: GraphSession) -> Dict[str, int]:
    """Graph 的颜色是 preset 名，与 Outlook 的数字枚举不是一回事 → 这项不检查。

    与其编一个可能错的映射（自检会据此报"颜色不对、下次分类自动修复"，那是误导），
    不如老实返回空字典。名字是否存在仍由 existing_categories 负责。
    """
    return {}


# 创建「分类主列表」需要这个权限；邮件读写（Mail.ReadWrite）不够。
# 它是一个**可选**权限：有些租户不允许用户自行同意，所以默认不申请 ——
# 没它照样能把 S/A/B/C 贴到邮件上，只是 Outlook 的分类列表里看不到这几项。
CATEGORY_SCOPE = "MailboxSettings.ReadWrite"

# 一旦确认"这个登录建不了分类"，本次运行就不再反复尝试 / 反复刷日志。
_category_create_blocked = False


def ensure_categories(ns: GraphSession) -> List[str]:
    """确保 S/A/B/C 存在（缺了就建）。返回改动说明，供调用方打日志。

    创建失败（缺 MailboxSettings.ReadWrite 权限 → HTTP 403）时**只报一次**，
    并说明怎么开启；判级与写回不受影响（Graph 允许把主列表里不存在的分类名
    直接写到邮件上，实测通过）。
    """
    global _category_create_blocked
    try:
        have = set(existing_categories(ns))
    except GraphError as ex:
        return [f"could not read categories: {ex}"]

    missing = [level for level in LEVELS if level not in have]
    if not missing or _category_create_blocked:
        return []

    try:
        for level in missing:
            ns.client.ensure_master_category(level)
        return ["created categories: " + ", ".join(missing)]
    except GraphError as ex:
        _category_create_blocked = True
        return [f"cannot create the S/A/B/C categories ({ex}). Levels are still written to "
                f"mail, but Outlook's category list stays empty. To let the app create them, "
                f"add {CATEGORY_SCOPE} to graph.scopes in config.json and sign in again."]


def ensure_sum_category(ns: GraphSession) -> bool:
    global _category_create_blocked
    if _category_create_blocked:
        return False
    try:
        if SUM_NAME in set(existing_categories(ns)):
            return False
        ns.client.ensure_master_category(SUM_NAME)
        return True
    except GraphError:
        _category_create_blocked = True
        return False


def outlook_exe_path() -> Optional[str]:
    """Graph 后端不需要 Outlook；保留这个名字只是让 run.py 的打印逻辑不炸。"""
    return None
