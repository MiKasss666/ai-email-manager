# Email Manager — AI 邮件管理助手（Outlook 经典版）

**为香港科技大学（HKUST）学生开发**：用 AI 自动把收件箱分成 S / A / B / C 四级，
并按需生成一份总结报告发给自己。课程代号、学院、宿舍关键词等默认值都按 HKUST 预置，
**换成别的学校也只需在配置里改几个字段**。

本文只讲四件事：**前置要求 → 环境配置 → 运行分类逻辑 → 使用方法**。

---

## 1. 前置要求

### 先拿到程序

| 方式 | 怎么做 | 适合 |
|---|---|---|
| **下载打包版** | 到本仓库的 **Releases** 页面下载 `EmailManager.zip`，解压后运行里面的 `EmailManager.exe` | 不想装 Python 的人 |
| **自己打包** | 克隆仓库 → 装好 Python + `pywin32` → 双击 `build.bat` | 想改代码 / 自己控制的人 |
| **直接跑源码** | 克隆仓库 → 装好 Python + `pywin32` → 双击 `start.bat` | 改完立刻试 |

> 仓库里**不含**打包产物（1025 个文件 / 29 MB 的二进制），所以克隆下来是没有 `EmailManager\` 的 ——
> 要么从 Releases 下载，要么自己跑一次 `build.bat`（约 1–3 分钟）。

### 必须满足

| 项 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11 |
| **Outlook** | **Outlook 经典版（桌面版）**，已登录、能正常收发邮件。程序通过 COM 接口直接操作它——网页版、新版 Outlook 不支持 |
| **启动方式** | **必须以普通用户身份启动**：在文件资源管理器里双击，或从开始菜单启动 |

> ⚠️ **不要从被沙箱 / 代理 / 受限 shell 里启动本程序。**
> 那样启动的进程会带上 **Low 完整性级别**，而 Windows 不允许 Low 完整性进程
> 启动或连接 Medium 完整性的 Outlook。症状是 Outlook 自己弹出
> "Cannot start Microsoft Outlook."，程序日志报
> `com_error: (-2146959355, 'Server execution failed')`。
> 程序启动时会把当前级别写进日志（`[env] security context: ...`），
> 是 `Medium` 才正常；拿不准就跑一次 `check_context.bat`。

### 运行方式（二选一）

| 方式 | 入口 | 需要装什么 |
|---|---|---|
| **打包版** | `EmailManager\EmailManager.exe` | 什么都不用装。首次运行会被 SmartScreen 拦一下：点 `More info` → `Run anyway` |
| **源码运行** | `start.bat` | Python 3.9+ 与 pywin32（`py -3 -m pip install -r requirements.txt`）。**没有 SmartScreen 提示** |

### AI 后端（必须有一个可用）

| 后端 | 类型 | 地址 / 模型 | 需要 API Key |
|---|---|---|---|
| DeepSeek | OpenAI 兼容 | `https://api.deepseek.com` / `deepseek-flash` | 是 |
| OpenAI | OpenAI | `https://api.openai.com/v1` / `gpt-4o-mini` | 是 |
| Anthropic | Anthropic | `https://api.anthropic.com` / `claude-sonnet-4-5` | 是 |
| Claude CLI | 本地命令行 | 本机装了 Claude Code CLI 即可 | 否 |
| Ollama | OpenAI 兼容 | `http://localhost:11434/v1` / `qwen2.5:7b` | 否（本地） |

模型名会随服务商更新，以官方文档为准。

---

## 2. 环境配置

### 配置文件在哪

| 运行方式 | 位置 |
|---|---|
| 源码运行 | `Email_manager_beta\config.json` |
| 打包运行 | `EmailManager\config.json`（exe 同目录） |

**首次运行自动生成**，不用手动创建。两种填写方式：

- **界面里填（推荐）**：`User info` 填身份，`AI setting` 选后端填 Key
- 直接编辑 `config.json`（保存后界面上重新打开即可生效）

### 身份字段

| 字段 | 作用 | 默认值 |
|---|---|---|
| `email` | 你的邮箱。**必填**——空着程序会拒绝执行任务 | 空 |
| `name` | 姓名。邮件里称呼你 → 判 A 级 | 空 |
| `student_id` | 学号。正文里出现学号 → 判 A 级 | 空 |
| `school` | 学校名。发件人/正文命中 → 判 A 级 | `香港科技大学 (HKUST)` |
| `faculty` | 学院 / 系名，可逗号分隔多个。命中 → 判 A 级 | 空 |
| `courses` | 本学期课程代号，逗号分隔。**命中 → 直接判 S 级**（确定性规则，覆盖 AI） | 空 |
| `faculty_keywords` | 额外的学院关键词（`faculty` 之外的补充） | 空 |
| `dorm_keywords` | 宿舍 / 宿位相关关键词 | HKUST 常用词 |
| `course_code_regex` | 课程代号正则，用于识别"别的课程"的邮件 → 判 A 级 | `[A-Z]{2,4}\d{3,4}[A-Z]?` |

### 行为字段

| 字段 | 作用 | 默认值 |
|---|---|---|
| `language` | `zh` / `en`。**只影响总结报告和邮件主题的语言**（界面固定英文） | `en` |
| `summary_recipient` | 总结报告发给谁；留空 = 发给自己 | 空 |
| `summary_days` | 总结覆盖最近多少天 | `7` |
| `batch_size` | 每次送给模型的邮件数 | `40` |
| `urgent_days` | deadline 在多少天内算"加急"（A 升 S） | `30` |
| `catchup_days` | 启动时补分类最近多少天的未分类邮件；**`0` = 关闭** | `7` |
| `ai.active` | 当前使用的后端名 | `deepseek` |
| `ai.profiles.<名>.base_url` | 接口地址 | 见上表 |
| `ai.profiles.<名>.model` | 模型名 | 见上表 |
| `ai.profiles.<名>.api_key` | API Key | 空 |
| `ai.profiles.<名>.max_tokens` | 单次输出上限。**调小会导致总结要点空白** | `16384` |
| `ai.profiles.<名>.json_mode` | 请求 `response_format=json_object`；后端不支持会自动降级 | `true` |

### API Key 的两种来源

配置里的 `api_key` **优先**；留空时读环境变量：

| 后端 | 环境变量 |
|---|---|
| deepseek | `DEEPSEEK_API_KEY` |
| openai | `OPENAI_API_KEY` |
| anthropic | `ANTHROPIC_API_KEY` |

想用环境变量方式，把配置里的 `api_key` 留空即可（这样 Key 不落在文件里）。

---

## 3. 运行分类逻辑

### 四个级别 + 一个标记

| 级别 | 含义 | 颜色 |
|---|---|---|
| **S** | 必须马上处理：作业 / 考试 / 要交的材料 / 你选的课程 | 橙色 |
| **A** | 重要，且来源可信：学校、学院、宿舍、老师、其它课程 | 紫色 |
| **B** | 值得一看：个人成长、讲座、实习、学习资源 | 蓝色 |
| **C** | 其余：活动通知、推广、兜底 | 灰色 |
| **Sum** | 程序自己的总结报告与 `sum` 指令邮件 | 黄色 |

### 判定顺序（前一步命中即停止）

```
step0  邮件命中你填的 courses（课程代号）        → S   ★ 确定性规则，直接覆盖 AI 判断
step1  来源可信（faculty / school / 宿舍 / 别的课程代号 / 你的姓名 / 学号） → A
step2  step1 命中，且正文里的 deadline 在 urgent_days 天内 → A 升 S
step3  属于个人成长 / 学习相关                  → B
step4  其余                                      → C
```

- **只有 step0 是确定性规则**：本地正则匹配，不经过 AI，也就不会判错。
- **step1–step4 由 AI 完成**：程序把这套规则写进提示词，由模型逐封判断。
- 送给 AI 的每封邮件只包含：**发件人、主题、正文前 400 字符**。
- 一次送 `batch_size` 封；模型返回的条数与请求**不一致时整批作废**——宁可漏判，也不允许错位贴错标签。

### 总结（Sum）流程

1. 读最近 `summary_days` 天内已打 `S/A/B/C` 的邮件
2. 按 `S → A → B → C` 分组
3. 每封生成一句要点（正文取前 800 字符，让要点能带上**截止日期、课程代号、占比、公司名**等具体信息）
4. 合成 HTML 报告 → 发到 `summary_recipient`（默认自己）
5. 报告主题带标记（`邮件总结` / `Email summary`，跟随 `language`），程序据此认出自己的报告并打 `Sum` 标签，**不会被自己分级**

**要点不会整批空白**（这一点做过专门加固）：要求模型返回"数量正好 N 个"的 JSON 数组，一旦响应被输出上限截断，旧做法会把整批丢掉。现在分三层兜底——捞回已完整的部分并只补剩下的 → 整批失败就劈半重试 → 最差只影响单封，日志会点名是哪一封。

### 自动监听（watch）

| 时机 | 行为 |
|---|---|
| 启动 | **配置齐全才开**（缺 `email` 或 AI 后端不可用就不开，日志逐条说明缺什么） |
| 保存配置后 | 配置补全的**那一刻自动开启**，不用重启 |
| 启动第一轮 | 按 `catchup_days` **补分类收件箱里存量的未分类邮件**，同时记录基线 |
| 每 N 秒（默认 30） | 只处理"比上次记住的那封更新"的邮件 → 分级 |
| 新邮件是 `sum` 指令 | 打 `Sum` 标签 → 生成总结报告 |
| 我们自己的报告 | 每轮扫描"是报告但没打 `Sum` 标签"的并补上（历史遗漏也会补） |
| Outlook 不可用 | **不当作失败**：只提示一次、继续按 120 秒退避重试，一恢复就自动接上 |

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
| `Include already classified` | 连已打过标签的也重新判一遍。**会覆盖你手动改过的分类**，且 AI 调用量成倍增加 |
| `Auto-watch new mail` | 自动监听开关（配置不齐时勾选会被拒绝并回弹） |
| `poll every [N] seconds` | 轮询间隔（5–3600） |
| **Activity log** | 运行日志。右上角 `Copy` 一键复制全文，`Clear` 清屏 |
| `Force reset` | 任务卡死时解锁界面（**不会杀死卡住的线程**，只解除界面封锁） |

日志同时写入 `app.log`（与 `config.json` 同目录），每行带 `HH:MM:SS` 时间戳。
同一种连接心跳只记一次，所以日志里保留下来的基本都是**真正发生的事**。

### 首次使用四步

1. 打开程序 → **User info** 填 `email`（建议把 `student_id`、`faculty`、`courses` 一起填上）→ **Save**
2. **AI setting** 选后端 → 填 API Key → 保存
3. 配置齐全后**监听会自动开启**（日志出现 `watching the inbox every 30s`）
4. 想先看效果：勾 **Dry run** → 点 **Categorize**，确认判级合理后取消 Dry run 再正式跑

### 命令行

```powershell
py -3 run.py check [--ai]         # 自检：配置 / Outlook 连接 / 分类颜色（--ai 会发一次极小请求）
py -3 run.py providers            # 列出所有 AI 后端及可用状态
py -3 run.py classify [--limit N] [--days N] [--dry-run] [--force]
py -3 run.py summarize [letters]  # 生成并发送总结报告；letters 如 s / sa / sabc
```

退出码：`0` 成功，`1` 配置或连接有问题，`2` 任务部分失败。

### 文件说明

| 文件 / 目录 | 说明 |
|---|---|
| `EmailManager\` | 打包产物（**不在仓库里**）：从 Releases 下载，或自己跑 `build.bat` 生成。分发时带上整个文件夹 |
| `start.bat` | 从源码启动（需本机有 Python + pywin32，无 SmartScreen 提示） |
| `build.bat` | 重新打包成可分发文件夹；最后一步会自动清除产物上的 Low 完整性标签并打印结果 |
| `check_context.bat` | 诊断脚本：打印当前窗口的完整性级别、exe 的标签、Outlook 状态 |
| `run.py` | 命令行入口 |
| `core/` `gui/` | 源码（分层：`config ← (ai, outlook) ← (classifier, summarizer) ← gui`） |
| `selftest.py` `guitest.py` | 测试：`py -3 selftest.py` 纯逻辑自测（不连 Outlook、不调 AI）；`py -3 guitest.py` 界面冒烟 |
| `config.json` / `watch_state.json` / `app.log` | **运行时生成，含你的个人信息，不要一起分发**（已在 `.gitignore` 里排除） |

---

## 许可

[MIT](LICENSE) © 2026 MiKasss666

程序会把**邮件内容**（发件人、主题、正文片段）发送给你自己配置的 AI 服务商，
请自行确认该服务商的隐私条款；API Key 明文保存在本机 `config.json` 中，不要把它提交到任何仓库。
