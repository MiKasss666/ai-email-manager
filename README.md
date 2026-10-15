# Email Manager · AI inbox assistant for your Microsoft mailbox

**Part 1 · English below ｜ Part 2 · 中文在下方（中文全文）**

---

# Part 1 · English

## ⬇️ Just want to use it? Download here

> ### 👉 **[Download the packaged build — EmailManager.zip](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest)** 👈
>
> Unzip → double-click `EmailManager.exe`. **No Python required.**

> ⚠️ **Don't use the green `Code` → `Download ZIP` button**
> That downloads the **source code** (29 files, ~435 KB, **no exe**, nothing happens when you double-click); it is only for developers who want to modify the code.
> Regular users should use the **Download** link above, or grab `EmailManager.zip` (13 MB) from [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest).

**Built for students at HKUST (Hong Kong University of Science and Technology)** — an AI assistant that sorts your mailbox into S / A / B / C and emails you a summary report on demand. It talks to **Microsoft Graph** over HTTPS, so **desktop Outlook is not required**. Course codes, school and dorm keywords ship pre-filled for HKUST; other schools only need a few config changes.

Four sections: **Prerequisites → Configuration → Classification logic → Usage**, plus **Troubleshooting** at the end.

---

## What's new in v2.0.0

- **The mail backend moved from Outlook COM to Microsoft Graph.** Desktop (Classic) Outlook is **no longer required at all** — the app works with Outlook not installed, not running and not signed in. Reads (inbox, bodies, categories, read state) and writes (the S/A/B/C categories, the red flag, the summary email) all go to the mailbox through the Graph REST API over HTTPS.
- **The only new user step: a one-time Microsoft sign-in on first launch.** No stored credential → the app prints the device code in the **Activity log**, opens your default browser with the code already in the URL, and puts the code on the clipboard too. After that it is silent: the credential is cached in `graph_token.json` and renewed automatically.
- **New config keys**: `mail_source` (`graph` default, or `outlook` for the legacy backend) and the `graph` section (`tenant`, `client_id`, `scopes`, `probe_scopes`, `open_browser`, `token_cache`). See [section 2](#2-configuration).
- **`pywin32` is now optional.** The Graph path uses only the Python standard library; `pywin32` is needed only by the legacy Outlook backend.
- **Sign-in asks for `Mail.ReadWrite`** (plus `offline_access openid profile User.Read`); the read-only self-check asks only for `Mail.Read`.

The legacy Outlook Classic backend is still available as a fallback — see [Optional: legacy Outlook backend](#optional-legacy-outlook-backend).

---

## 1. Prerequisites

### Get the program

| Option | What to do | For |
|---|---|---|
| **Download the build** | Grab `EmailManager.zip` from [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest), unzip, run `EmailManager.exe` | No Python needed |
| **Build it yourself** | Clone → install Python → double-click `build.bat` | You want to modify the code |
| **Run from source** | Clone → install Python → double-click `start.bat` | Quick iteration |

> The repository does **not** contain the build output (1025 binary files / 29 MB), so a fresh clone has no `EmailManager\` folder — either download it from Releases or run `build.bat` once (1–3 minutes).

### Hard requirements

| Item | Requirement |
|---|---|
| OS | Windows 10 / 11 |
| **Mail account** | A Microsoft / Microsoft 365 account you can sign in to. **Outlook does not have to be installed, running or signed in.** |
| **Sign-in** | **One-time Microsoft sign-in on first launch** (OAuth2 device-code flow). You can sign out at any time by deleting the cached credential — see [Troubleshooting](#troubleshooting). |
| **How to launch** | **Launch it as a normal user**: double-click it in File Explorer, or start it from the Start menu. |

### Mail backend

| Backend | What it uses | Outlook needed? |
|---|---|---|
| **graph** (default) | Microsoft Graph REST API over HTTPS, Python standard library only | **No** — not installed, not running, not signed in |
| outlook (legacy) | Outlook Classic driven through COM (`pywin32`) | Yes |

Switch with `mail_source` in `config.json`; the Graph section is described in [section 2](#2-configuration).

### How to run (pick one)

| Option | Entry point | Requirements |
|---|---|---|
| **Packaged build** | `EmailManager\EmailManager.exe` | Nothing to install. SmartScreen blocks it once on first run: click `More info` → `Run anyway`. |
| **From source** | `start.bat` | Python 3.9+ (`py -3 -m pip install -r requirements.txt` installs the optional `pywin32`, which only the legacy Outlook backend needs). **No SmartScreen prompt.** |

### Optional: legacy Outlook backend

A legacy backend that drives **Outlook Classic (desktop)** through COM is still shipped, for anyone who wants the old behaviour (no Microsoft sign-in at all).

| Step | What to do |
|---|---|
| 1 | Install the optional dependency: `py -3 -m pip install pywin32` |
| 2 | Make sure Outlook Classic (desktop) is installed, signed in and able to send/receive mail — Outlook on the web and the new Outlook are **not** supported on this path |
| 3 | Set `"mail_source": "outlook"` in `config.json` and restart the app |

> ⚠️ This path talks to a local COM server, so do not launch the app from a sandboxed / restricted shell there: a **Low**-integrity process cannot reach a Medium-integrity Outlook, and the log then shows `com_error: (-2146959355, 'Server execution failed')`. `check_context.bat` prints the integrity level of the current window, the exe's label and whether Outlook is running.

### AI backends (one must be usable)

| Backend | Type | Endpoint / Model | API key | Tested |
|---|---|---|---|---|
| **DeepSeek** | OpenAI compatible | `https://api.deepseek.com` / `deepseek-flash` | yes | ✅ **fully tested** |
| OpenAI | OpenAI | `https://api.openai.com/v1` / `gpt-4o-mini` | yes | ⚠️ **not tested** |
| Anthropic | Anthropic | `https://api.anthropic.com` / `claude-sonnet-4-5` | yes | ⚠️ **not tested** |
| Claude CLI | local CLI | Any machine with Claude Code CLI installed | no | ⚠️ **not tested** |
| Ollama | OpenAI compatible | `http://localhost:11434/v1` / `qwen2.5:7b` | no (local) | ⚠️ **not tested** |

Model names change over time — always check your provider's docs.

### ⚠️ Known limitations

- **Only the DeepSeek backend has been fully tested.** Classification, summarisation and real mailbox read/write were all verified against the official DeepSeek API with `deepseek-flash`.
  The other four backends are covered by **unit tests only** (URL building, request shape, response parsing, error translation) — **no real API call has ever been made to them**, so whether they work and how well they classify is currently **unknown**.
  To try one: `py -3 run.py check --ai` (sends one tiny request to confirm the backend answers).

- **The classification prompt is currently Chinese-only.** The rules sent to the model are written in Chinese (the GUI and docs are English; this one piece is not). Classification was verified on mixed Chinese/English mail, but whether it is optimal for an all-English mailbox is **not yet verified**.

- **The legacy Outlook backend is Windows + Outlook Classic only.** Outlook on the web, the new Outlook, macOS and Linux are not supported on that path (it drives Outlook through COM). The Graph backend is not tied to desktop Outlook, but the app itself is still a Windows/Tkinter program.

- **A network connection is required.** The Graph backend reaches your mailbox over HTTPS; the legacy Outlook backend talked to a local Outlook instead and could work offline.

- **How much the app may do is decided by your Microsoft 365 tenant.** If the organisation requires administrator approval for new apps (or blocks them), sign-in will be refused — ask your IT department, or fall back to the legacy Outlook backend.

---

## 2. Configuration

### Where the config lives

| Mode | Location |
|---|---|
| From source | `Email_manager_beta\config.json` |
| Packaged build | `EmailManager\config.json` (next to the exe) |

**Created automatically on first run** — no need to make it yourself. Two ways to fill it in:

- **In the GUI (recommended)**: `User info` for your identity, `AI setting` for the backend and key
- Edit `config.json` directly (reopen the app to pick it up)

`mail_source` and the `graph` section have no GUI control yet, so **edit `config.json` for those**. The credential cache `graph_token.json` (see below) is written next to it after the first sign-in; both are personal data and are already in `.gitignore`.

### Identity fields

| Field | Purpose | Default |
|---|---|---|
| `email` | Your address. **Required** — tasks are refused while it is empty. | empty |
| `name` | Your name. Mail addressing you by name → level A. | empty |
| `student_id` | Student ID. Appearing in the body → level A. | empty |
| `school` | School name. **Display only**; the school-mail test uses `school_domains` below. | `香港科技大学 (HKUST)` |
| `school_domains` | **Deterministic school-mail test**: a sender domain matching one of these is school mail without asking the AI. Not matching does **not** mean it isn't — school business is often relayed by external platforms, and the AI judges those from the content. | `ust.hk` |
| `faculty` | Faculty / department, comma-separated. Match → level A. | empty |
| `courses` | This term's course codes, comma-separated. **Match → level S directly** (deterministic, overrides the AI). | empty |
| `faculty_keywords` | Extra faculty keywords. | empty |
| `dorm_keywords` | Dorm / housing keywords. | HKUST defaults |
| `course_code_regex` | Regex for course codes, used to spot mail from *other* courses → level A. | `[A-Z]{2,4}\d{3,4}[A-Z]?` |

### Behaviour fields

| Field | Purpose | Default |
|---|---|---|
| `mail_source` | Which mail backend: `graph` (Microsoft Graph, no Outlook needed) or `outlook` (legacy Outlook Classic + COM). No GUI control — edit the file and reopen the app. | `graph` |
| `graph.tenant` | Which Microsoft tenant sign-in is restricted to. `common` accepts any account; set your school's domain (e.g. `hkust.edu.hk`) to refuse accounts from outside it. | `common` |
| `graph.client_id` | Microsoft Entra app id. Empty = use Microsoft's own public client, which is fine for personal use; register your own app at [entra.microsoft.com](https://entra.microsoft.com) to distribute to others. | empty |
| `graph.scopes` | Permissions requested when signing in to read **and write** mail. | `offline_access openid profile User.Read Mail.ReadWrite` |
| `graph.probe_scopes` | Read-only permissions used by the self-check. | `offline_access openid profile User.Read Mail.Read` |
| `graph.open_browser` | Open the default browser automatically when a sign-in is needed. `false` = no window opens, the code still appears in the log. | `true` |
| `graph.token_cache` | Where the sign-in credential is cached, relative to `config.json`. Delete it to force a fresh sign-in. | `graph_token.json` |
| `language` | `zh` / `en`. **Affects only the summary report and its subject** (the GUI stays English). | `en` |
| `summary_recipient` | Where the summary goes; empty = send to yourself. | empty |
| `summary_days` | How many days the summary covers. | `7` |
| `batch_size` | Emails per AI request. | `40` |
| `urgent_days` | A deadline within N days counts as urgent (A → S). | `30` |
| `catchup_days` | On startup, classify unlabelled mail from the last N days; **`0` disables**. | `7` |
| `ai.active` | Which backend is active. | `deepseek` |
| `ai.profiles.<name>.base_url` | API endpoint. | see above |
| `ai.profiles.<name>.model` | Model name. | see above |
| `ai.profiles.<name>.api_key` | API key. | empty |
| `ai.profiles.<name>.max_tokens` | Max output tokens. **Too small leaves summary bullets blank.** | `16384` |
| `ai.profiles.<name>.json_mode` | Requests `response_format=json_object`; falls back automatically if unsupported. | `true` |

### Two sources for the API key

The `api_key` in the config wins; when it is empty the environment variable is used:

| Backend | Env var |
|---|---|
| deepseek | `DEEPSEEK_API_KEY` |
| openai | `OPENAI_API_KEY` |
| anthropic | `ANTHROPIC_API_KEY` |

Leave `api_key` empty to use the env var, so the key never lands in a file.

### Signing in to the mailbox

With the default `graph` backend the app signs in to Microsoft **once**:

1. When no stored credential exists it prints the **device code on its own line in the Activity log** (the log panel has a `Copy` button).
2. It **opens your default browser** automatically. The URL it opens already carries the code (`...?user_code=XXXX`), so on the sign-in page you usually just confirm.
3. It also puts the device code **on the clipboard**, so pressing `Ctrl+V` on the sign-in page is enough.
4. It waits for you (up to about **15 minutes**) and **continues by itself** once the sign-in finishes — no button to press.

Afterwards the credential is cached in `graph_token.json` next to `config.json` and reused silently (refresh token, roughly 90 days, rolling renewal), so normal use never asks you to sign in again. Delete that file to sign out — the next launch asks you to sign in again.

The permissions requested are **`Mail.ReadWrite`** (needed to write the S/A/B/C categories back to the mailbox), plus `offline_access openid profile User.Read`; the read-only self-check asks only for **`Mail.Read`**. If `graph.open_browser` is `false` the browser is not opened — the code still appears in the log for you to enter by hand.

---

## 3. How classification works

### Four levels + one marker

| Level | Meaning | Colour |
|---|---|---|
| **S** | Act now: homework, exams, documents to submit, your enrolled courses | orange |
| **A** | Important and from a trusted source: faculty, dorm, teachers, other courses | purple |
| **B** | Worth a look: personal growth, talks, internships, study resources | blue |
| **C** | Everything else: event notices, promotions, fallback | grey |
| **Sum** | The app's own summary reports and `sum` command mail | yellow |

When writing a category back, **only the app's own labels are replaced** (`S`/`A`/`B`/`C`/`Sum`): categories you defined yourself, or those applied by server rules, are **left untouched**; re-judging replaces the level instead of stacking. The app never deletes entries from your Outlook category list.

### Decision order (first match wins)

```
step-1 AI decides whether the mail is school mail; a sender domain in `school_domains` is
       decided by code without asking. Not school mail → no label at all.
step0  sender/subject/body matches one of your `courses` → S   ★ deterministic rule, overrides the AI
step1  trusted source (faculty / dorm / other course code / your name / student ID) → A
step2  step1 matched and a deadline in the body is within `urgent_days` → A upgraded to S
step3  personal growth / study related → B
step4  everything else → C
```

- **Why step-1 asks the AI**: a domain-only rule misses school business that is **sent on behalf of the school by external platforms** (Canvas notifications, surveys, sign-up systems…) — the domain is not the school's, but the content is. So the AI judges it from the content; only domain matches are decided by code (deterministic evidence outranks the AI).
- **Non-school mail is never written to**: no label, no flag, and not counted as a failure (so no retry loop). Such mails are remembered in `watch_state.json` and are not sent to the AI again.
- **Only step0 and the domain match are deterministic**: local checks, no AI involved.
- **step-1 and step1–step4 are done by the AI**: the rules go into the prompt and the model judges each mail.
- Each mail sent to the AI contains only: **sender, subject, first 400 characters of the body**. ⚠️ Because step-1 reads the content, **every not-yet-judged inbox mail is sent to the AI** — if that matters to you, point the AI backend at a local model such as `ollama`.
- `batch_size` mails go per request; if the model returns a different number of items the **whole batch is discarded** — better to miss than to mislabel.

### Summary flow

1. Read mail labelled `S/A/B/C` from the last `summary_days` days
2. Group by `S → A → B → C`
3. Generate one bullet per mail (first 800 characters of the body, so bullets can carry **deadlines, course codes, weightings, company names**)
4. Render an HTML report → send it to `summary_recipient` (yourself by default)
5. The subject carries a marker (`邮件总结` / `Email summary`, following `language`) so the app recognises its own report, tags it `Sum`, and **never classifies it**

**Bullets never come back blank in bulk**: the model is asked for a JSON array of *exactly* N items; when the response is cut off by the output limit, a naive implementation throws the whole batch away. Three layers prevent that now: salvage the complete items and re-ask only for the rest → split a failed batch in half and retry → worst case only a single mail is affected, and the log names it.

### Auto-watch

| When | Behaviour |
|---|---|
| Startup | **Only starts when the config is complete** (missing `email` or an unusable AI backend keeps it off, and the log lists exactly what is missing) |
| After saving settings | Turns on **the moment** the config becomes complete — no restart |
| First tick | Classifies **existing unlabelled mail** within `catchup_days` and records the baseline |
| Every N seconds (30 by default) | Only mail newer than the last-seen message is classified |
| New mail is a `sum` command | Tag `Sum` → generate and send the summary |
| Our own reports | Each tick scans for reports missing the `Sum` tag and fixes them (including historical ones) |
| Mailbox temporarily unreachable | **Not treated as a failure**: one notice, then retry with a 120-second backoff and reconnect automatically (this covers both a Graph connection problem and the legacy backend finding Outlook unavailable) |

---

## 4. Usage

### GUI controls

| Control | What it does |
|---|---|
| **User info** | Your identity (open, edit, `Save`) |
| **AI setting** | Pick a backend, fill in `base_url` / `model` / API key |
| **Categorize** | Classify the inbox now |
| **Sum** | Generate and send the summary report now |
| `Scope: last N days` | How far back the two buttons above look |
| `Dry run (print only, no write-back)` | **Print only, no write-back.** Strongly recommended on your first run |
| `Include already classified` | Re-judge already-labelled mail. The level is re-judged, but **your own category labels are preserved**; AI usage multiplies |
| `Auto-watch new mail` | Auto-watch toggle (refused and reverted while the config is incomplete) |
| `poll every [N] seconds` | Poll interval (5–3600) |
| **Activity log** | Runtime log. `Copy` grabs the whole thing, `Clear` wipes it |
| `Force reset` | Unlocks the UI when a task hangs. **Does not kill the stuck thread**, but makes it stop at the next safe point (it stops writing to the mailbox), and its late result is discarded |

The log is also written to `app.log` (next to `config.json`) with `HH:MM:SS` timestamps. Repeated connection heartbeats are logged once, so what remains is mostly **things that actually happened**.

The app is **single-instance**: a second launch tells you it is already running and exits. Two copies would each run their own watcher and write categories to the same mailbox at the same time.

### First run in five steps

1. Open the app → **User info** → fill in `email` (plus `student_id`, `faculty`, `courses` if you can) → **Save**
2. The **one-time Microsoft sign-in** may appear on first use (with the Graph backend): follow it in the browser once — see [Signing in to the mailbox](#signing-in-to-the-mailbox)
3. **AI setting** → pick a backend → enter the API key → save
4. The watcher **starts by itself** once the config is complete (the log shows `watching the inbox every 30s`)
5. To preview first: tick **Dry run** → click **Categorize**; untick it once the levels look right

### Command line

```powershell
py -3 run.py check [--ai]         # self-check: config / mail access / category colours (--ai sends one tiny request)
py -3 run.py providers            # list backends and their status
py -3 run.py classify [--limit N] [--days N] [--dry-run] [--force]
py -3 run.py summarize [letters]  # build and send the summary; letters like s / sa / sabc
```

Exit codes: `0` success, `1` config or connection problem, `2` partial failure.

### Files

| File / folder | Description |
|---|---|
| `EmailManager\` | Build output (**not in the repo**): download from Releases or run `build.bat`. Ship the whole folder. |
| `start.bat` | Run from source (needs Python locally, plus `pywin32` only for the legacy backend; no SmartScreen prompt) |
| `build.bat` | Rebuild the distributable folder; the last step clears Low integrity labels (which matter for the legacy COM backend) and prints the result |
| `check_context.bat` | Diagnostic for the legacy COM backend: prints this window's integrity level, the exe's label and whether Outlook is running |
| `run.py` | Command-line entry point |
| `core/` `gui/` | Source (layered: `config ← (ai, mail) ← (classifier, summarizer) ← gui`; `core/mail.py` picks the mail backend, implemented by `core/graph.py` or `core/outlook.py`) |
| `selftest.py` `guitest.py` | Tests: `py -3 selftest.py` (pure logic, no mailbox/AI); `py -3 guitest.py` (GUI smoke test) |
| `config.json` / `graph_token.json` / `watch_state.json` / `app.log` | **Created at runtime, contains your personal data — never redistribute** (already in `.gitignore`). `graph_token.json` holds the cached Microsoft sign-in credential. |

---

## Troubleshooting

| Symptom | What it means / what to do |
|---|---|
| The sign-in page says the app **needs administrator approval** | Your school's Microsoft 365 tenant does not let users consent to new apps. Ask your IT department, or set `"mail_source": "outlook"` in `config.json` and use the legacy Outlook backend. |
| The **device code expired** | The code is valid for about **15 minutes**. Restart the app to get a fresh code, then finish the sign-in without waiting. |
| Sign-in worked but **mail cannot be read** | Check your network connection: the Graph backend needs internet access, unlike the old local COM path. |
| You want to sign in **again / as another account** | Delete `graph_token.json` (next to `config.json`) to force a fresh sign-in. |

---

## License

[MIT](LICENSE) © 2026 MiKasss666

The app sends **email content** (sender, subject, body excerpts) to the AI provider *you* configure — check that provider's privacy policy yourself. Your API key is stored in plain text in the local `config.json`; never commit it to any repository.

---

## Credit & Citation

**Designed by pkfok@hkust** ｜ Hong Kong University of Science and Technology (HKUST)

| Item | Detail |
|---|---|
| Author | **pkfok@hkust** (HKUST) |
| First published | **2026-10-10** — the public, timestamped commit history of this repository is the record of authorship |
| License | [MIT](LICENSE) © 2026 MiKasss666 |

### Terms of use

This project is released under the **MIT license**: you are free to use, modify and redistribute it.
**The MIT license requires that the copyright notice and this permission notice be included in all
copies or substantial portions of the Software.**

In other words: **stripping the credit, the LICENSE file or this notice and using the code anyway is
already outside the terms of the licence.**

### How to cite

> This project uses Email Manager by pkfok@hkust (HKUST),
> available at https://github.com/MiKasss666/hkust-ai-email-manager

A machine-readable [CITATION.cff](CITATION.cff) is included, so GitHub shows a
**"Cite this repository"** button in the sidebar.

---
---

# Part 2 · 中文

# Email Manager · AI 邮件管理助手

## ⬇️ 只是想用？点这里下载

> ### 👉 **[下载打包版 EmailManager.zip](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest)** 👈
>
> 解压 → 双击 `EmailManager.exe` → 完事。**不需要安装 Python。**

> ⚠️ **别点右上角绿色的 `Code` → `Download ZIP`**
> 那个下载的是**源码**（29 个文件 / 约 435 KB，里面**没有 exe**，双击没反应），只有想改代码的开发者才需要。
> 普通用户请用上面那个 **Download** 链接，或到 [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest) 页面下载 `EmailManager.zip`（13 MB）。

**为香港科技大学（HKUST）学生开发** —— 用 AI 把收件箱自动分成 S / A / B / C 四级，并按需生成一份总结报告发给自己。程序通过 HTTPS 直接对接 **Microsoft Graph**，所以**不需要安装桌面版 Outlook**。课程代号、学院、宿舍关键词等默认值都按 HKUST 预置，换成别的学校只需在配置里改几个字段。

本文四节：**前置要求 → 环境配置 → 运行分类逻辑 → 使用方法**，最后另加一节**故障排查**。

---

## 版本更新 · v2.0.0

- **邮件后端从 Outlook COM 换成 Microsoft Graph。** **不再需要**桌面版（经典版）Outlook——没装、没开、没登录都能用。读取（收件箱、正文、分类、已读状态）与写入（写回 S/A/B/C 分类、红旗、发送总结邮件）全部通过 HTTPS 调用 Graph REST API 完成。
- **用户唯一新增的步骤：首次启动时登录一次 Microsoft 账号。** 本地没有凭据时，程序会把**设备码打到 Activity log**、自动打开默认浏览器（网址里已带上设备码），并把设备码放进剪贴板。之后就无感了：凭据缓存在 `graph_token.json`，会自动续期。
- **新增配置项**：`mail_source`（默认 `graph`，可设为 `outlook` 走旧后端）与 `graph` 段（`tenant`、`client_id`、`scopes`、`probe_scopes`、`open_browser`、`token_cache`）。详见 [第 2 节](#2-环境配置)。
- **`pywin32` 变成可选依赖。** Graph 后端只用 Python 标准库；只有旧的 Outlook 后端才需要 `pywin32`。
- **登录时申请的权限是 `Mail.ReadWrite`**（另外还有 `offline_access openid profile User.Read`）；只读自检只申请 `Mail.Read`。

旧的 Outlook 经典版后端仍保留为回退方案：见[可选：旧版 Outlook 后端](#可选旧版-outlook-后端)。

---

## 1. 前置要求

### 先拿到程序

| 方式 | 怎么做 | 适合 |
|---|---|---|
| **下载打包版** | 到 [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest) 下载 `EmailManager.zip`，解压后运行里面的 `EmailManager.exe` | 不想装 Python 的人 |
| **自己打包** | 克隆仓库 → 装好 Python → 双击 `build.bat` | 想改代码的人 |
| **直接跑源码** | 克隆仓库 → 装好 Python → 双击 `start.bat` | 改完立刻试 |

> 仓库里**不含**打包产物（1025 个文件 / 29 MB 的二进制），所以克隆下来没有 `EmailManager\` —— 要么从 Releases 下载，要么自己跑一次 `build.bat`（约 1–3 分钟）。

### 必须满足

| 项 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11 |
| **邮箱账号** | 一个你能登录的 Microsoft / Microsoft 365 账号。**Outlook 不需要安装、不需要运行、也不需要登录。** |
| **登录** | **首次启动时登录一次 Microsoft 账号**（OAuth2 设备码流程）。随时可以退出登录：删掉缓存的凭据即可，见[故障排查](#故障排查)。 |
| **启动方式** | **必须以普通用户身份启动**：在文件资源管理器里双击，或从开始菜单启动。 |

### 邮件后端

| 后端 | 用什么 | 需要 Outlook 吗 |
|---|---|---|
| **graph**（默认） | 通过 HTTPS 调用 Microsoft Graph REST API，只用 Python 标准库 | **不需要**——没装、没开、没登录都行 |
| outlook（旧版） | 通过 COM 驱动 Outlook 经典版（需要 `pywin32`） | 需要 |

在 `config.json` 里用 `mail_source` 切换；`graph` 段各字段见[第 2 节](#2-环境配置)。

### 运行方式（二选一）

| 方式 | 入口 | 需要装什么 |
|---|---|---|
| **打包版** | `EmailManager\EmailManager.exe` | 什么都不用装。首次运行会被 SmartScreen 拦一下：点 `More info` → `Run anyway`。 |
| **源码运行** | `start.bat` | Python 3.9+（`py -3 -m pip install -r requirements.txt` 会装上可选的 `pywin32`，只有旧版 Outlook 后端才用得到）。**没有 SmartScreen 提示**。 |

### 可选：旧版 Outlook 后端

程序仍保留一个通过 COM 驱动 **Outlook 经典版（桌面版）** 的旧后端，想用回旧行为（完全不需要登录 Microsoft）的人可以用它。

| 步骤 | 怎么做 |
|---|---|
| 1 | 装可选依赖：`py -3 -m pip install pywin32` |
| 2 | 确保 Outlook 经典版（桌面版）已安装、已登录且能正常收发邮件——这条路径**不支持**网页版 Outlook 和新版 Outlook |
| 3 | 在 `config.json` 里设 `"mail_source": "outlook"`，然后重启程序 |

> ⚠️ 这条路径要连本机的 COM 服务，所以不要从被沙箱 / 受限的 shell 里启动程序：**Low** 完整性进程连不上 Medium 完整性的 Outlook，日志里会出现 `com_error: (-2146959355, 'Server execution failed')`。`check_context.bat` 会打印当前窗口的完整性级别、exe 的标签以及 Outlook 是否在运行。

### AI 后端（必须有一个可用）

| 后端 | 类型 | 地址 / 模型 | 需要 Key | 实测状态 |
|---|---|---|---|---|
| **DeepSeek** | OpenAI 兼容 | `https://api.deepseek.com` / `deepseek-flash` | 是 | ✅ **已完整实测** |
| OpenAI | OpenAI | `https://api.openai.com/v1` / `gpt-4o-mini` | 是 | ⚠️ **未实测** |
| Anthropic | Anthropic | `https://api.anthropic.com` / `claude-sonnet-4-5` | 是 | ⚠️ **未实测** |
| Claude CLI | 本地命令行 | 本机装了 Claude Code CLI 即可 | 否 | ⚠️ **未实测** |
| Ollama | OpenAI 兼容 | `http://localhost:11434/v1` / `qwen2.5:7b` | 否（本地） | ⚠️ **未实测** |

模型名会随服务商更新，以官方文档为准。

### ⚠️ 目前的限制

- **只用 DeepSeek 做过完整实测。** 分类、总结、真实邮箱读写都跑通过（DeepSeek 官方 API + `deepseek-flash`）。
  其余四个后端**只做过代码层面的单元测试**（URL 拼接、请求构造、响应解析、错误码翻译），**没有对真实服务发起过调用**，所以它们能不能用、判得准不准，目前**未知**。
  想试的话先跑一次自检：`py -3 run.py check --ai`（会发一个极小请求，确认后端真的可用）。

- **分类提示词目前固定为中文。** 送给模型的规则文本是中文的（界面和文档都是英文，只有这段提示词不是）。中英文邮件都实测能正常判级，但纯英文环境下是否最优**尚未验证**。

- **仅支持 Windows。** 程序本身是 Windows 上的 Python/Tkinter 程序；旧版 Outlook 后端另外还要求 Outlook 经典版（它通过 COM 操作 Outlook，网页版、新版 Outlook、macOS、Linux 在这条路径上都不支持）。默认的 Graph 后端不依赖桌面版 Outlook。

- **必须能联网。** Graph 后端通过 HTTPS 访问你的邮箱；旧版 Outlook 后端连的是本机 Outlook，离线也能工作。

- **程序能做什么，取决于你所在组织的 Microsoft 365 租户设置。** 如果学校要求第三方应用必须由管理员批准（或直接禁用），登录会被拒绝——可以找 IT 处理，或改回旧版 Outlook 后端。

---

## 2. 环境配置

### 配置文件在哪

| 运行方式 | 位置 |
|---|---|
| 源码运行 | `Email_manager_beta\config.json` |
| 打包运行 | `EmailManager\config.json`（exe 同目录） |

**首次运行自动生成**，不用手动创建。两种填写方式：

- **界面里填（推荐）**：`User info` 填身份，`AI setting` 选后端填 Key
- 直接编辑 `config.json`（保存后重新打开界面即可生效）

`mail_source` 与 `graph` 段目前**没有界面入口**，只能**直接编辑 `config.json`**。首次登录后程序会在同目录生成凭据缓存 `graph_token.json`（见下文）；两者都是个人信息，且都已写进 `.gitignore`。

### 身份字段

| 字段 | 作用 | 默认值 |
|---|---|---|
| `email` | 你的邮箱。**必填**——空着程序会拒绝执行任务。 | 空 |
| `name` | 姓名。邮件里称呼你 → 判 A 级。 | 空 |
| `student_id` | 学号。正文里出现学号 → 判 A 级。 | 空 |
| `school` | 学校名。**仅用于显示/署名**；判"是不是学校邮件"改用下面的 `school_domains`。 | `香港科技大学 (HKUST)` |
| `school_domains` | **「是学校邮件」的确定性判据**：发件人域名命中这里（逗号分隔）→ 直接认定为学校邮件，不再问 AI。没命中**不代表不是**——学校事务常由外部平台代发，那一类由 AI 按内容判断。 | `ust.hk` |
| `faculty` | 学院 / 系名，可逗号分隔多个。命中 → 判 A 级。 | 空 |
| `courses` | 本学期课程代号，逗号分隔。**命中 → 直接判 S 级**（确定性规则，覆盖 AI）。 | 空 |
| `faculty_keywords` | 额外的学院关键词。 | 空 |
| `dorm_keywords` | 宿舍 / 宿位相关关键词。 | HKUST 常用词 |
| `course_code_regex` | 课程代号正则，用于识别"别的课程"的邮件 → 判 A 级。 | `[A-Z]{2,4}\d{3,4}[A-Z]?` |

### 行为字段

| 字段 | 作用 | 默认值 |
|---|---|---|
| `mail_source` | 用哪个邮件后端：`graph`（Microsoft Graph，不需要 Outlook）或 `outlook`（旧的 Outlook 经典版 + COM）。没有界面入口——改完文件重开程序生效。 | `graph` |
| `graph.tenant` | 登录限定在哪个 Microsoft 租户。`common` 接受任何账号；填本校域名（如 `hkust.edu.hk`）可拒绝校外账号登录。 | `common` |
| `graph.client_id` | Microsoft Entra 应用 id。留空 = 用微软自家公开客户端，自己用没问题；要分发给别人请到 [entra.microsoft.com](https://entra.microsoft.com) 注册自己的应用。 | 空 |
| `graph.scopes` | 登录时申请、用于读**并写**邮件的权限。 | `offline_access openid profile User.Read Mail.ReadWrite` |
| `graph.probe_scopes` | 自检用的只读权限。 | `offline_access openid profile User.Read Mail.Read` |
| `graph.open_browser` | 需要登录时自动打开默认浏览器。`false` = 不弹浏览器，设备码仍然会打进日志。 | `true` |
| `graph.token_cache` | 登录凭据缓存的文件名（相对 `config.json` 所在目录）。删掉它 = 强制重新登录。 | `graph_token.json` |
| `language` | `zh` / `en`。**只影响总结报告和邮件主题的语言**（界面固定英文）。 | `en` |
| `summary_recipient` | 总结报告发给谁；留空 = 发给自己。 | 空 |
| `summary_days` | 总结覆盖最近多少天。 | `7` |
| `batch_size` | 每次送给模型的邮件数。 | `40` |
| `urgent_days` | deadline 在多少天内算"加急"（A 升 S）。 | `30` |
| `catchup_days` | 启动时补分类最近多少天的未分类邮件；**`0` = 关闭**。 | `7` |
| `ai.active` | 当前使用的后端名。 | `deepseek` |
| `ai.profiles.<名>.base_url` | 接口地址。 | 见上表 |
| `ai.profiles.<名>.model` | 模型名。 | 见上表 |
| `ai.profiles.<名>.api_key` | API Key。 | 空 |
| `ai.profiles.<名>.max_tokens` | 单次输出上限。**调小会导致总结要点空白**。 | `16384` |
| `ai.profiles.<名>.json_mode` | 请求 `response_format=json_object`；后端不支持会自动降级。 | `true` |

### API Key 的两种来源

配置里的 `api_key` **优先**；留空时读环境变量：

| 后端 | 环境变量 |
|---|---|
| deepseek | `DEEPSEEK_API_KEY` |
| openai | `OPENAI_API_KEY` |
| anthropic | `ANTHROPIC_API_KEY` |

想用环境变量方式，把配置里的 `api_key` 留空即可（这样 Key 不落在文件里）。

### 登录邮箱

用默认的 `graph` 后端时，程序只需要登录**一次**：

1. 本地没有凭据时，程序会把**设备码单独一行打进 Activity log**（日志区右上角有 `Copy` 按钮）。
2. 它会**自动打开默认浏览器**，打开的网址里已经带着设备码（`...?user_code=XXXX`），所以在登录页上通常只要确认一下。
3. 它还会把设备码**放进剪贴板**，在登录页直接 `Ctrl+V` 就行。
4. 之后它会等你（最长约 **15 分钟**），登录一完成就**自己继续**，不需要再点任何按钮。

登录成功后凭据缓存在与 `config.json` 同目录的 `graph_token.json` 里，之后静默复用（refresh token，大约 90 天，滚动续期），日常使用不会再要求登录。删掉这个文件即等于退出登录——下次启动会重新要求登录一次。

登录申请的权限是 **`Mail.ReadWrite`**（写回 S/A/B/C 分类必需），另外还有 `offline_access openid profile User.Read`；只读自检只申请 **`Mail.Read`**。若把 `graph.open_browser` 设为 `false`，浏览器不会自动弹出——设备码仍然会打进日志，手动输入即可。

---

## 3. 运行分类逻辑

### 四个级别 + 一个标记

| 级别 | 含义 | 颜色 |
|---|---|---|
| **S** | 必须马上处理：作业 / 考试 / 要交的材料 / 你选的课程 | 橙色 |
| **A** | 重要，且来源可信：学院、宿舍、老师、其它课程 | 紫色 |
| **B** | 值得一看：个人成长、讲座、实习、学习资源 | 蓝色 |
| **C** | 其余：活动通知、推广、兜底 | 灰色 |
| **Sum** | 程序自己的总结报告与 `sum` 指令邮件 | 黄色 |

写回分类时**只替换本程序自己的标签**（`S`/`A`/`B`/`C`/`Sum`）：你自己定义的分类、以及服务器规则打上的分类会**原样保留**，重判也只换级别、不会层层叠加。程序从不删除你 Outlook 里的分类主列表条目。

### 判定顺序（前一步命中即停止）

```
step-1 先问 AI：这封是不是学校邮件？发件人域名命中 school_domains 的由代码直接认定
       └─ 不是学校邮件 → 完全不贴标签（Outlook 里保持原样）
step0  邮件命中你填的 courses（课程代号）        → S   ★ 确定性规则，直接覆盖 AI 判断
step1  来源可信（faculty / 宿舍 / 别的课程代号 / 你的姓名 / 学号） → A
step2  step1 命中，且正文里的 deadline 在 urgent_days 天内 → A 升 S
step3  属于个人成长 / 学习相关                  → B
step4  其余                                      → C
```

- **step-1 为什么要问 AI**：只看发件人域名会漏——学校事务经常由**外部平台代发**（Canvas 通知、问卷、报名系统…），域名不是学校的，内容却是学校事务。所以"是不是学校邮件"也交给 AI 按内容判断；只有域名命中的那一类由代码直接认定（确定性证据优先于 AI）。
- **不是学校邮件的邮件一个字节都不写**：不贴 S/A/B/C、不动分类、不动红旗，也不算失败（不会触发重试）。这些邮件会被记进 `watch_state.json`，之后不再重复送 AI。
- **只有 step0 和"域名命中"是确定性规则**：本地判断，不经过 AI，也就不会判错。
- **step-1 / step1–step4 由 AI 完成**：程序把这套规则写进提示词，由模型逐封判断。
- 送给 AI 的每封邮件只包含：**发件人、主题、正文前 400 字符**。⚠️ 因为 step-1 要看内容，**收件箱里的邮件（只要还没被判定过）都会送进 AI**——介意隐私的话请把 AI 后端换成 `ollama` 等本地模型，或在 `config.json` 里不要配云端 Key。
- 一次送 `batch_size` 封；模型返回的条数与请求**不一致时整批作废**——宁可漏判，也不允许错位贴错标签。

### 总结流程

1. 读最近 `summary_days` 天内已打 `S/A/B/C` 的邮件
2. 按 `S → A → B → C` 分组
3. 每封生成一句要点（正文取前 800 字符，让要点能带上**截止日期、课程代号、占比、公司名**等具体信息）
4. 合成 HTML 报告 → 发到 `summary_recipient`（默认自己）
5. 报告主题带标记（`邮件总结` / `Email summary`，跟随 `language`），程序据此认出自己的报告并打 `Sum` 标签，**不会被自己分级**

**要点不会整批空白**：要求模型返回"数量正好 N 个"的 JSON 数组，一旦响应被输出上限截断，旧做法会把整批丢掉。现在分三层兜底——捞回已完整的部分并只补剩下的 → 整批失败就劈半重试 → 最差只影响单封，日志会点名是哪一封。

### 自动监听

| 时机 | 行为 |
|---|---|
| 启动 | **配置齐全才开**（缺 `email` 或 AI 后端不可用就不开，日志逐条说明缺什么） |
| 保存配置后 | 配置补全的**那一刻自动开启**，不用重启 |
| 启动第一轮 | 按 `catchup_days` **补分类收件箱里存量的未分类邮件**，同时记录基线 |
| 每 N 秒（默认 30） | 只处理"比上次记住的那封更新"的邮件 → 分级 |
| 新邮件是 `sum` 指令 | 打 `Sum` 标签 → 生成总结报告 |
| 我们自己的报告 | 每轮扫描"是报告但没打 `Sum` 标签"的并补上（历史遗漏也会补） |
| 邮箱暂时连不上 | **不当作失败**：只提示一次、继续按 120 秒退避重试，一恢复就自动接上（Graph 连不上、旧后端连不上 Outlook，都走这条） |

---

## 4. 使用方法

### 界面控件

| 控件 | 作用 |
|---|---|
| **User info** | 填个人信息（点开编辑 → `Save` 保存） |
| **AI setting** | 选 AI 后端、填 `base_url` / `model` / API Key |
| **Categorize** | 手动给收件箱分级 |
| **Sum** | 手动生成并发送总结报告 |
| `Scope: last N days` | 上面两个按钮的作用范围 |
| `Dry run (print only, no write-back)` | **只打印不写回**。第一次用强烈建议勾上 |
| `Include already classified` | 连已打过标签的也重新判一遍。级别会被重判，**但你自己定义的分类标签会保留**；AI 调用量成倍增加 |
| `Auto-watch new mail` | 自动监听开关（配置不齐时勾选会被拒绝并回弹） |
| `poll every [N] seconds` | 轮询间隔（5–3600） |
| **Activity log** | 运行日志。右上角 `Copy` 一键复制全文，`Clear` 清屏 |
| `Force reset` | 任务卡死时解锁界面。**不会杀死卡住的线程**，但会让它在下一个安全点停下来（不再继续写邮箱），它在后台的结果也会被丢弃 |

日志同时写入 `app.log`（与 `config.json` 同目录），每行带 `HH:MM:SS` 时间戳。同一种连接心跳只记一次，所以日志里保留下来的基本都是**真正发生的事**。

程序是**单实例**的：第二次启动会提示"已经在运行"并直接退出。两个实例会各自开一份监听、同时给同一个邮箱写分类（标签互相覆盖），所以必须挡住。

### 首次使用五步

1. 打开程序 → **User info** 填 `email`（建议把 `student_id`、`faculty`、`courses` 一起填上）→ **Save**
2. 用 Graph 后端时，首次使用可能会弹出**一次性的 Microsoft 登录**：跟着浏览器完成一次即可，见[登录邮箱](#登录邮箱)
3. **AI setting** 选后端 → 填 API Key → 保存
4. 配置齐全后**监听会自动开启**（日志出现 `watching the inbox every 30s`）
5. 想先看效果：勾 **Dry run** → 点 **Categorize**，确认判级合理后取消 Dry run 再正式跑

### 命令行

```powershell
py -3 run.py check [--ai]         # 自检：配置 / 邮箱访问 / 分类颜色（--ai 会发一次极小请求）
py -3 run.py providers            # 列出所有 AI 后端及可用状态
py -3 run.py classify [--limit N] [--days N] [--dry-run] [--force]
py -3 run.py summarize [letters]  # 生成并发送总结报告；letters 如 s / sa / sabc
```

退出码：`0` 成功，`1` 配置或连接有问题，`2` 任务部分失败。

### 文件说明

| 文件 / 目录 | 说明 |
|---|---|
| `EmailManager\` | 打包产物（**不在仓库里**）：从 Releases 下载，或自己跑 `build.bat` 生成。分发时带上整个文件夹。 |
| `start.bat` | 从源码启动（需本机有 Python；`pywin32` 只有旧版后端才需要，无 SmartScreen 提示） |
| `build.bat` | 重新打包成可分发文件夹；最后一步会自动清除产物上的 Low 完整性标签（对旧的 COM 后端才有影响）并打印结果 |
| `check_context.bat` | 旧 COM 后端的诊断脚本：打印当前窗口的完整性级别、exe 的标签、Outlook 是否在运行 |
| `run.py` | 命令行入口 |
| `core/` `gui/` | 源码（分层：`config ← (ai, mail) ← (classifier, summarizer) ← gui`；邮件后端由 `core/mail.py` 选择，实现在 `core/graph.py` 或 `core/outlook.py`） |
| `selftest.py` `guitest.py` | 测试：`py -3 selftest.py` 纯逻辑自测（不连邮箱、不调 AI）；`py -3 guitest.py` 界面冒烟 |
| `config.json` / `graph_token.json` / `watch_state.json` / `app.log` | **运行时生成，含你的个人信息，不要一起分发**（已在 `.gitignore` 里排除）。`graph_token.json` 里是登录 Microsoft 的凭据缓存。 |

---

## 故障排查

| 现象 | 原因 / 怎么办 |
|---|---|
| 登录页提示**需要管理员批准** | 你所在学校的 Microsoft 365 租户不允许用户自行同意新应用。找 IT 处理，或在 `config.json` 里设 `"mail_source": "outlook"` 改用旧版 Outlook 后端。 |
| **设备码过期** | 设备码有效期约 **15 分钟**。重启程序拿一个新码，然后尽快完成登录。 |
| 登录成功了但**读不到邮件** | 检查网络：Graph 后端需要联网，这一点和旧的本地 COM 路径不同。 |
| 想**换账号 / 重新登录** | 删掉 `graph_token.json`（与 `config.json` 同目录）即可强制重新登录。 |

---

## 许可

[MIT](LICENSE) © 2026 MiKasss666

程序会把**邮件内容**（发件人、主题、正文片段）发送给你自己配置的 AI 服务商，请自行确认该服务商的隐私条款；API Key 明文保存在本机 `config.json` 中，不要把它提交到任何仓库。

---

## 声明与引用

**Designed by pkfok@hkust** ｜ 香港科技大学（HKUST）

| 项 | 内容 |
|---|---|
| 作者 | **pkfok@hkust**（HKUST） |
| 首次发布 | **2026-10-10** —— 本仓库的**公开提交记录**（带时间戳）即为著作权证明 |
| 许可 | [MIT](LICENSE) © 2026 MiKasss666 |

### 使用条款

本项目以 **MIT 许可**发布，你可以自由使用、修改、分发。**但 MIT 许可有一项硬性要求：
所有副本都必须保留版权声明与许可声明**（见 [LICENSE](LICENSE) 第一段）。

也就是说 —— **把署名、LICENSE 或这段声明删掉再拿去用，本身就已经超出了授权范围。**

### 引用格式

> 本项目使用了 Email Manager（作者 pkfok@hkust，HKUST），
> 仓库地址：https://github.com/MiKasss666/hkust-ai-email-manager

仓库里还有一份机器可读的 [CITATION.cff](CITATION.cff) —— GitHub 会据此在侧边栏显示
「**Cite this repository**」按钮，点一下就能复制标准引用格式。
