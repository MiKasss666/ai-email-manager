# Email Manager · AI 邮件管理助手（Outlook 经典版）

**中文** ｜ English below each section

> ## ⬇️ 只是想用？点这里下载
> ## ⬇️ Just want to use it? Download here
>
> ### 👉 **[下载打包版 Download EmailManager.zip](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest)** 👈
>
> 解压 → 双击 `EmailManager.exe` → 完事。**不需要安装 Python。**
> Unzip → double-click `EmailManager.exe`. **No Python required.**

> ⚠️ **别点右上角绿色的 `Code` → `Download ZIP`**
> 那个下载的是**源码**（约 360 KB，里面**没有 exe**），只有想改代码的开发者才需要。
> 普通用户请用上面那个 **Download** 链接，或到 [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest) 页面下载 `EmailManager.zip`（13 MB）。
>
> ⚠️ **Don't use the green `Code` → `Download ZIP` button**
> That downloads the **source code** (~360 KB, **no exe**); it is only for developers who want to modify the code.
> Regular users should use the **Download** link above, or grab `EmailManager.zip` (13 MB) from [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest).

**为香港科技大学（HKUST）学生开发** —— 用 AI 把收件箱自动分成 S / A / B / C 四级，并按需生成一份总结报告发给自己。课程代号、学院、宿舍关键词等默认值都按 HKUST 预置，换成别的学校只需在配置里改几个字段。

**Built for students at HKUST (Hong Kong University of Science and Technology)** — an AI assistant that sorts your Outlook inbox into S / A / B / C and emails you a summary report on demand. Course codes, school and dorm keywords ship pre-filled for HKUST; other schools only need a few config changes.

本文四节：**前置要求 → 环境配置 → 运行分类逻辑 → 使用方法**。
Four sections: **Prerequisites → Configuration → Classification logic → Usage**.

---

## 1. 前置要求 / Prerequisites

### 先拿到程序 / Get the program

| 方式 / Option | 怎么做 / What to do | 适合 / For |
|---|---|---|
| **下载打包版**<br>Download the build | 到 [Releases](https://github.com/MiKasss666/hkust-ai-email-manager/releases/latest) 下载 `EmailManager.zip`，解压后运行里面的 `EmailManager.exe`<br>Grab `EmailManager.zip` from Releases, unzip, run `EmailManager.exe` | 不想装 Python 的人<br>No Python needed |
| **自己打包**<br>Build it yourself | 克隆仓库 → 装好 Python + `pywin32` → 双击 `build.bat`<br>Clone → install Python + `pywin32` → double-click `build.bat` | 想改代码的人<br>You want to modify the code |
| **直接跑源码**<br>Run from source | 克隆仓库 → 装好 Python + `pywin32` → 双击 `start.bat`<br>Clone → install Python + `pywin32` → double-click `start.bat` | 改完立刻试<br>Quick iteration |

> 仓库里**不含**打包产物（1025 个文件 / 29 MB 的二进制），所以克隆下来没有 `EmailManager\` —— 要么从 Releases 下载，要么自己跑一次 `build.bat`（约 1–3 分钟）。
>
> The repository does **not** contain the build output (1025 binary files / 29 MB), so a fresh clone has no `EmailManager\` folder — either download it from Releases or run `build.bat` once (1–3 minutes).

### 必须满足 / Hard requirements

| 项 / Item | 要求 / Requirement |
|---|---|
| 操作系统<br>OS | Windows 10 / 11 |
| **Outlook** | **Outlook 经典版（桌面版）**，已登录、能正常收发邮件。程序通过 COM 接口直接操作它——网页版、新版 Outlook 不支持。<br>**Outlook Classic (desktop)**, signed in and able to send/receive mail. The app drives it through COM — Outlook on the web and the new Outlook are **not** supported. |
| **启动方式**<br>How to launch | **必须以普通用户身份启动**：在文件资源管理器里双击，或从开始菜单启动。<br>**Launch it as a normal user**: double-click it in File Explorer, or start it from the Start menu. |

> ⚠️ **不要从被沙箱 / 代理 / 受限 shell 里启动本程序。**
> 那样启动的进程会带上 **Low 完整性级别**，而 Windows 不允许 Low 完整性进程启动或连接 Medium 完整性的 Outlook。症状是 Outlook 自己弹出 "Cannot start Microsoft Outlook."，程序日志报 `com_error: (-2146959355, 'Server execution failed')`。程序启动时会把当前级别写进日志（`[env] security context: ...`），是 `Medium` 才正常；拿不准就跑一次 `check_context.bat`。
>
> ⚠️ **Do not launch this app from a sandboxed / proxied / restricted shell.**
> Such a process gets a **Low integrity level**, and Windows does not allow a Low-integrity process to start or reach a Medium-integrity Outlook. The symptom: Outlook itself pops up "Cannot start Microsoft Outlook." and the log shows `com_error: (-2146959355, 'Server execution failed')`. The app writes its level to the log on startup (`[env] security context: ...`) — it must say `Medium`. Not sure? Run `check_context.bat`.

### 运行方式 / How to run（二选一 / pick one）

| 方式 / Option | 入口 / Entry point | 需要装什么 / Requirements |
|---|---|---|
| **打包版**<br>Packaged build | `EmailManager\EmailManager.exe` | 什么都不用装。首次运行会被 SmartScreen 拦一下：点 `More info` → `Run anyway`。<br>Nothing to install. SmartScreen blocks it once on first run: click `More info` → `Run anyway`. |
| **源码运行**<br>From source | `start.bat` | Python 3.9+ 与 pywin32（`py -3 -m pip install -r requirements.txt`）。**没有 SmartScreen 提示**。<br>Python 3.9+ and pywin32 (`py -3 -m pip install -r requirements.txt`). **No SmartScreen prompt.** |

### AI 后端 / AI backends（必须有一个可用 / one must be usable）

| 后端 / Backend | 类型 / Type | 地址 / 模型<br>Endpoint / Model | 需要 Key<br>API key |
|---|---|---|---|
| DeepSeek | OpenAI 兼容 / compatible | `https://api.deepseek.com` / `deepseek-flash` | 是 / yes |
| OpenAI | OpenAI | `https://api.openai.com/v1` / `gpt-4o-mini` | 是 / yes |
| Anthropic | Anthropic | `https://api.anthropic.com` / `claude-sonnet-4-5` | 是 / yes |
| Claude CLI | 本地命令行 / local CLI | 本机装了 Claude Code CLI 即可<br>Any machine with Claude Code CLI installed | 否 / no |
| Ollama | OpenAI 兼容 / compatible | `http://localhost:11434/v1` / `qwen2.5:7b` | 否（本地）/ no (local) |

模型名会随服务商更新，以官方文档为准。
Model names change over time — always check your provider's docs.

---

## 2. 环境配置 / Configuration

### 配置文件在哪 / Where the config lives

| 运行方式 / Mode | 位置 / Location |
|---|---|
| 源码运行<br>From source | `Email_manager_beta\config.json` |
| 打包运行<br>Packaged build | `EmailManager\config.json`（exe 同目录 / next to the exe） |

**首次运行自动生成**，不用手动创建。两种填写方式：
**Created automatically on first run** — no need to make it yourself. Two ways to fill it in:

- **界面里填（推荐）**：`User info` 填身份，`AI setting` 选后端填 Key
  **In the GUI (recommended)**: `User info` for your identity, `AI setting` for the backend and key
- 直接编辑 `config.json`（保存后重新打开界面即可生效）
  Edit `config.json` directly (reopen the app to pick it up)

### 身份字段 / Identity fields

| 字段 / Field | 作用 / Purpose | 默认值 / Default |
|---|---|---|
| `email` | 你的邮箱。**必填**——空着程序会拒绝执行任务。<br>Your address. **Required** — tasks are refused while it is empty. | 空 / empty |
| `name` | 姓名。邮件里称呼你 → 判 A 级。<br>Your name. Mail addressing you by name → level A. | 空 / empty |
| `student_id` | 学号。正文里出现学号 → 判 A 级。<br>Student ID. Appearing in the body → level A. | 空 / empty |
| `school` | 学校名。发件人/正文命中 → 判 A 级。<br>School name. Matching sender/body → level A. | `香港科技大学 (HKUST)` |
| `faculty` | 学院 / 系名，可逗号分隔多个。命中 → 判 A 级。<br>Faculty / department, comma-separated. Match → level A. | 空 / empty |
| `courses` | 本学期课程代号，逗号分隔。**命中 → 直接判 S 级**（确定性规则，覆盖 AI）。<br>This term's course codes, comma-separated. **Match → level S directly** (deterministic, overrides the AI). | 空 / empty |
| `faculty_keywords` | 额外的学院关键词。<br>Extra faculty keywords. | 空 / empty |
| `dorm_keywords` | 宿舍 / 宿位相关关键词。<br>Dorm / housing keywords. | HKUST 常用词 / HKUST defaults |
| `course_code_regex` | 课程代号正则，用于识别"别的课程"的邮件 → 判 A 级。<br>Regex for course codes, used to spot mail from *other* courses → level A. | `[A-Z]{2,4}\d{3,4}[A-Z]?` |

### 行为字段 / Behaviour fields

| 字段 / Field | 作用 / Purpose | 默认值 / Default |
|---|---|---|
| `language` | `zh` / `en`。**只影响总结报告和邮件主题的语言**（界面固定英文）。<br>`zh` / `en`. **Affects only the summary report and its subject** (the GUI stays English). | `en` |
| `summary_recipient` | 总结报告发给谁；留空 = 发给自己。<br>Where the summary goes; empty = send to yourself. | 空 / empty |
| `summary_days` | 总结覆盖最近多少天。<br>How many days the summary covers. | `7` |
| `batch_size` | 每次送给模型的邮件数。<br>Emails per AI request. | `40` |
| `urgent_days` | deadline 在多少天内算"加急"（A 升 S）。<br>A deadline within N days counts as urgent (A → S). | `30` |
| `catchup_days` | 启动时补分类最近多少天的未分类邮件；**`0` = 关闭**。<br>On startup, classify unlabelled mail from the last N days; **`0` disables**. | `7` |
| `ai.active` | 当前使用的后端名。<br>Which backend is active. | `deepseek` |
| `ai.profiles.<名>.base_url` | 接口地址。<br>API endpoint. | 见上表 / see above |
| `ai.profiles.<名>.model` | 模型名。<br>Model name. | 见上表 / see above |
| `ai.profiles.<名>.api_key` | API Key。<br>API key. | 空 / empty |
| `ai.profiles.<名>.max_tokens` | 单次输出上限。**调小会导致总结要点空白**。<br>Max output tokens. **Too small leaves summary bullets blank.** | `16384` |
| `ai.profiles.<名>.json_mode` | 请求 `response_format=json_object`；后端不支持会自动降级。<br>Requests `response_format=json_object`; falls back automatically if unsupported. | `true` |

### API Key 的两种来源 / Two sources for the API key

配置里的 `api_key` **优先**；留空时读环境变量：
The `api_key` in the config wins; when it is empty the environment variable is used:

| 后端 / Backend | 环境变量 / Env var |
|---|---|
| deepseek | `DEEPSEEK_API_KEY` |
| openai | `OPENAI_API_KEY` |
| anthropic | `ANTHROPIC_API_KEY` |

想用环境变量方式，把配置里的 `api_key` 留空即可（这样 Key 不落在文件里）。
Leave `api_key` empty to use the env var, so the key never lands in a file.

---

## 3. 运行分类逻辑 / How classification works

### 四个级别 + 一个标记 / Four levels + one marker

| 级别 / Level | 含义 / Meaning | 颜色 / Colour |
|---|---|---|
| **S** | 必须马上处理：作业 / 考试 / 要交的材料 / 你选的课程<br>Act now: homework, exams, documents to submit, your enrolled courses | 橙色 / orange |
| **A** | 重要，且来源可信：学校、学院、宿舍、老师、其它课程<br>Important and from a trusted source: school, faculty, dorm, teachers, other courses | 紫色 / purple |
| **B** | 值得一看：个人成长、讲座、实习、学习资源<br>Worth a look: personal growth, talks, internships, study resources | 蓝色 / blue |
| **C** | 其余：活动通知、推广、兜底<br>Everything else: event notices, promotions, fallback | 灰色 / grey |
| **Sum** | 程序自己的总结报告与 `sum` 指令邮件<br>The app's own summary reports and `sum` command mail | 黄色 / yellow |

### 判定顺序 / Decision order（前一步命中即停止 / first match wins）

```
step0  邮件命中你填的 courses（课程代号）        → S   ★ 确定性规则，直接覆盖 AI 判断
step1  来源可信（faculty / school / 宿舍 / 别的课程代号 / 你的姓名 / 学号） → A
step2  step1 命中，且正文里的 deadline 在 urgent_days 天内 → A 升 S
step3  属于个人成长 / 学习相关                  → B
step4  其余                                      → C

step0  sender/subject/body matches one of your `courses` → S   ★ deterministic rule, overrides the AI
step1  trusted source (faculty / school / dorm / other course code / your name / student ID) → A
step2  step1 matched and a deadline in the body is within `urgent_days` → A upgraded to S
step3  personal growth / study related → B
step4  everything else → C
```

- **只有 step0 是确定性规则**：本地正则匹配，不经过 AI，也就不会判错。
  **Only step0 is deterministic**: local regex matching, no AI involved, so it cannot be misjudged.
- **step1–step4 由 AI 完成**：程序把这套规则写进提示词，由模型逐封判断。
  **step1–step4 are done by the AI**: the rules go into the prompt and the model judges each mail.
- 送给 AI 的每封邮件只包含：**发件人、主题、正文前 400 字符**。
  Each mail sent to the AI contains only: **sender, subject, first 400 characters of the body**.
- 一次送 `batch_size` 封；模型返回的条数与请求**不一致时整批作废**——宁可漏判，也不允许错位贴错标签。
  `batch_size` mails go per request; if the model returns a different number of items the **whole batch is discarded** — better to miss than to mislabel.

### 总结流程 / Summary flow

1. 读最近 `summary_days` 天内已打 `S/A/B/C` 的邮件
   Read mail labelled `S/A/B/C` from the last `summary_days` days
2. 按 `S → A → B → C` 分组
   Group by `S → A → B → C`
3. 每封生成一句要点（正文取前 800 字符，让要点能带上**截止日期、课程代号、占比、公司名**等具体信息）
   Generate one bullet per mail (first 800 characters of the body, so bullets can carry **deadlines, course codes, weightings, company names**)
4. 合成 HTML 报告 → 发到 `summary_recipient`（默认自己）
   Render an HTML report → send it to `summary_recipient` (yourself by default)
5. 报告主题带标记（`邮件总结` / `Email summary`，跟随 `language`），程序据此认出自己的报告并打 `Sum` 标签，**不会被自己分级**
   The subject carries a marker (`邮件总结` / `Email summary`, following `language`) so the app recognises its own report, tags it `Sum`, and **never classifies it**

**要点不会整批空白 / Bullets never come back blank in bulk**：要求模型返回"数量正好 N 个"的 JSON 数组，一旦响应被输出上限截断，旧做法会把整批丢掉。现在分三层兜底——捞回已完整的部分并只补剩下的 → 整批失败就劈半重试 → 最差只影响单封，日志会点名是哪一封。

The model is asked for a JSON array of *exactly* N items; when the response is cut off by the output limit, a naive implementation throws the whole batch away. Three layers prevent that now: salvage the complete items and re-ask only for the rest → split a failed batch in half and retry → worst case only a single mail is affected, and the log names it.

### 自动监听 / Auto-watch

| 时机 / When | 行为 / Behaviour |
|---|---|
| 启动<br>Startup | **配置齐全才开**（缺 `email` 或 AI 后端不可用就不开，日志逐条说明缺什么）<br>**Only starts when the config is complete** (missing `email` or an unusable AI backend keeps it off, and the log lists exactly what is missing) |
| 保存配置后<br>After saving settings | 配置补全的**那一刻自动开启**，不用重启<br>Turns on **the moment** the config becomes complete — no restart |
| 启动第一轮<br>First tick | 按 `catchup_days` **补分类收件箱里存量的未分类邮件**，同时记录基线<br>Classifies **existing unlabelled mail** within `catchup_days` and records the baseline |
| 每 N 秒（默认 30）<br>Every N seconds (30 by default) | 只处理"比上次记住的那封更新"的邮件 → 分级<br>Only mail newer than the last-seen message is classified |
| 新邮件是 `sum` 指令<br>New mail is a `sum` command | 打 `Sum` 标签 → 生成总结报告<br>Tag `Sum` → generate and send the summary |
| 我们自己的报告<br>Our own reports | 每轮扫描"是报告但没打 `Sum` 标签"的并补上（历史遗漏也会补）<br>Each tick scans for reports missing the `Sum` tag and fixes them (including historical ones) |
| Outlook 不可用<br>Outlook unavailable | **不当作失败**：只提示一次、继续按 120 秒退避重试，一恢复就自动接上<br>**Not treated as a failure**: one notice, then retry with a 120-second backoff and reconnect automatically |

---

## 4. 使用方法 / Usage

### 界面控件 / GUI controls

| 控件 / Control | 作用 / What it does |
|---|---|
| **User info** | 填个人信息（点开编辑 → `Save` 保存）<br>Your identity (open, edit, `Save`) |
| **AI setting** | 选 AI 后端、填 `base_url` / `model` / API Key<br>Pick a backend, fill in `base_url` / `model` / API key |
| **Categorize** | 手动给收件箱分级<br>Classify the inbox now |
| **Sum** | 手动生成并发送总结报告<br>Generate and send the summary report now |
| `Scope: last N days` | 上面两个按钮的作用范围<br>How far back the two buttons above look |
| `Dry run (print only, no write-back)` | **只打印不写回**。第一次用强烈建议勾上<br>**Print only, no write-back.** Strongly recommended on your first run |
| `Include already classified` | 连已打过标签的也重新判一遍。**会覆盖你手动改过的分类**，且 AI 调用量成倍增加<br>Re-judge already-labelled mail. **Overwrites manual edits** and multiplies AI usage |
| `Auto-watch new mail` | 自动监听开关（配置不齐时勾选会被拒绝并回弹）<br>Auto-watch toggle (refused and reverted while the config is incomplete) |
| `poll every [N] seconds` | 轮询间隔（5–3600）<br>Poll interval (5–3600) |
| **Activity log** | 运行日志。右上角 `Copy` 一键复制全文，`Clear` 清屏<br>Runtime log. `Copy` grabs the whole thing, `Clear` wipes it |
| `Force reset` | 任务卡死时解锁界面（**不会杀死卡住的线程**，只解除界面封锁）<br>Unlocks the UI when a task hangs (**does not kill the stuck thread**) |

日志同时写入 `app.log`（与 `config.json` 同目录），每行带 `HH:MM:SS` 时间戳。同一种连接心跳只记一次，所以日志里保留下来的基本都是**真正发生的事**。

The log is also written to `app.log` (next to `config.json`) with `HH:MM:SS` timestamps. Repeated connection heartbeats are logged once, so what remains is mostly **things that actually happened**.

### 首次使用四步 / First-run in four steps

1. 打开程序 → **User info** 填 `email`（建议把 `student_id`、`faculty`、`courses` 一起填上）→ **Save**
   Open the app → **User info** → fill in `email` (plus `student_id`, `faculty`, `courses` if you can) → **Save**
2. **AI setting** 选后端 → 填 API Key → 保存
   **AI setting** → pick a backend → enter the API key → save
3. 配置齐全后**监听会自动开启**（日志出现 `watching the inbox every 30s`）
   The watcher **starts by itself** once the config is complete (the log shows `watching the inbox every 30s`)
4. 想先看效果：勾 **Dry run** → 点 **Categorize**，确认判级合理后取消 Dry run 再正式跑
   To preview first: tick **Dry run** → click **Categorize**; untick it once the levels look right

### 命令行 / Command line

```powershell
py -3 run.py check [--ai]         # 自检：配置 / Outlook 连接 / 分类颜色（--ai 会发一次极小请求）
                                  # self-check: config / Outlook / category colours (--ai sends one tiny request)
py -3 run.py providers            # 列出所有 AI 后端及可用状态 / list backends and their status
py -3 run.py classify [--limit N] [--days N] [--dry-run] [--force]
py -3 run.py summarize [letters]  # 生成并发送总结报告；letters 如 s / sa / sabc
                                  # build and send the summary; letters like s / sa / sabc
```

退出码 / Exit codes：`0` 成功 / success，`1` 配置或连接有问题 / config or connection problem，`2` 任务部分失败 / partial failure。

### 文件说明 / Files

| 文件 / 目录<br>File / folder | 说明 / Description |
|---|---|
| `EmailManager\` | 打包产物（**不在仓库里**）：从 Releases 下载，或自己跑 `build.bat` 生成。分发时带上整个文件夹。<br>Build output (**not in the repo**): download from Releases or run `build.bat`. Ship the whole folder. |
| `start.bat` | 从源码启动（需本机有 Python + pywin32，无 SmartScreen 提示）<br>Run from source (needs Python + pywin32 locally; no SmartScreen prompt) |
| `build.bat` | 重新打包成可分发文件夹；最后一步会自动清除产物上的 Low 完整性标签并打印结果<br>Rebuild the distributable folder; the last step clears Low integrity labels and prints the result |
| `check_context.bat` | 诊断脚本：打印当前窗口的完整性级别、exe 的标签、Outlook 状态<br>Diagnostic: prints this window's integrity level, the exe's label and whether Outlook is running |
| `run.py` | 命令行入口 / command-line entry point |
| `core/` `gui/` | 源码（分层：`config ← (ai, outlook) ← (classifier, summarizer) ← gui`）<br>Source (layered: `config ← (ai, outlook) ← (classifier, summarizer) ← gui`) |
| `selftest.py` `guitest.py` | 测试：`py -3 selftest.py` 纯逻辑自测（不连 Outlook、不调 AI）；`py -3 guitest.py` 界面冒烟<br>Tests: `py -3 selftest.py` (pure logic, no Outlook/AI); `py -3 guitest.py` (GUI smoke test) |
| `config.json` / `watch_state.json` / `app.log` | **运行时生成，含你的个人信息，不要一起分发**（已在 `.gitignore` 里排除）<br>**Created at runtime, contains your personal data — never redistribute** (already in `.gitignore`) |

---

## 许可 / License

[MIT](LICENSE) © 2026 MiKasss666

程序会把**邮件内容**（发件人、主题、正文片段）发送给你自己配置的 AI 服务商，请自行确认该服务商的隐私条款；API Key 明文保存在本机 `config.json` 中，不要把它提交到任何仓库。

The app sends **email content** (sender, subject, body excerpts) to the AI provider *you* configure — check that provider's privacy policy yourself. Your API key is stored in plain text in the local `config.json`; never commit it to any repository.
