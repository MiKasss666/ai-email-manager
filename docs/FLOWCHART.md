# Email Manager v3.0.0 · 程序完整流程图

> 依据 **v3.0.0** 源码（commit `45ed398`，Graph-only 后端）绘制。
> 源文件就是本 Markdown：GitHub / VS Code / Typora 可直接渲染 Mermaid。
> 本地可视化版：双击同目录的 `flowchart.html`（离线可用），由 `build-flowchart-html.py` 生成。

## 目录

1. [总体架构](#1-总体架构)
2. [启动与初始化](#2-启动与初始化)
3. [登录与凭据缓存](#3-登录与凭据缓存)
4. [自动监听：调度与一轮工作](#4-自动监听调度与一轮工作)
5. [启动补分类](#5-启动补分类)
6. [分类流水线](#6-分类流水线)
7. [AI 调用层](#7-ai-调用层)
8. [总结报告](#8-总结报告)
9. [监听状态机与基线推进](#9-监听状态机与基线推进)
10. [失败与退避](#10-失败与退避)
11. [数据与文件](#11-数据与文件)
12. [构建 / 运行 / 发布入口](#12-构建--运行--发布入口)
13. [附录 A：关键常量](#附录-a关键常量)
14. [附录 B：失败语义速查](#附录-b失败语义速查)

---

## 1. 总体架构

```mermaid
flowchart LR
  subgraph ENTRY["入口"]
    APP["app.py<br/>GUI 入口 + 单实例互斥"]
    RUN["run.py<br/>CLI: check / providers / classify / summarize"]
  end

  subgraph GUI["gui/ 界面层（只做线程调度与展示）"]
    MW["main_window.py<br/>主窗口 + watchdog + Force reset"]
    WATCH["watcher.py<br/>自动监听调度"]
    TASK["task.py<br/>TaskRunner：单任务串行 + 代次令牌"]
    DLG["ai_settings_dialog.py / user_info_dialog.py"]
    LOG["log_view.py<br/>界面日志 + app.log"]
  end

  subgraph CORE["core/ 业务层"]
    CFG["config.py<br/>config.json 加载/校验/原子写"]
    META["mailmeta.py<br/>分类名·报告与指令识别·merge_categories·message_info"]
    MAIL["mail.py<br/>邮件后端门面"]
    GRAPH["graph.py<br/>Microsoft Graph 唯一实现<br/>设备码登录·token 缓存·限流退避·按需取正文"]
    CLS["classifier.py<br/>S/A/B/C 判定 + 写回"]
    SUM["summarizer.py<br/>一句话要点 + HTML 报告"]
    AI["ai.py<br/>OpenAI 兼容 / Anthropic / CLI"]
    VER["version.py<br/>3.0.0 单一真源"]
  end

  subgraph EXT["外部服务"]
    MS["Microsoft Graph API<br/>HTTPS"]
    LLM["AI 服务商<br/>DeepSeek / OpenAI / Anthropic / CLI / Ollama"]
  end

  subgraph FILES["本地运行时文件（.gitignore，不进分发）"]
    F1["config.json"]
    F2["graph_token.json"]
    F3["watch_state.json"]
    F4["app.log"]
  end

  APP --> MW
  RUN --> CLS
  RUN --> SUM
  RUN --> MAIL
  MW --> TASK
  MW --> WATCH
  MW --> DLG
  MW --> LOG
  MW --> CFG
  MW --> VER
  WATCH --> CLS
  WATCH --> SUM
  WATCH --> MAIL
  CLS --> AI
  SUM --> AI
  CLS --> META
  SUM --> META
  MAIL --> GRAPH
  GRAPH --> META
  GRAPH --> MS
  AI --> LLM
  CFG --> F1
  GRAPH --> F2
  WATCH --> F3
  LOG --> F4
```

---

## 2. 启动与初始化

```mermaid
flowchart TD
  A["双击 EmailManager.exe 或 start.bat"] --> B["app.py：控制台编码 / DPI 感知 / AppUserModelID"]
  B --> C{"CreateMutex 单实例检查"}
  C -->|"已存在"| D["Win32 弹窗：已在运行 → exit 1"]
  C -->|"不存在，或检测失败放行"| E["tk.Tk() 创建主窗口"]
  E -->|"失败（无桌面会话）"| F["Win32 弹窗：请改用 py -3 run.py check → exit 1"]
  E --> G["MainWindow(root)"]
  G --> H["config.ensure()<br/>缺失则原子写默认 config.json"]
  H --> I["theme + 控件 + LogView + TaskRunner + WatchController"]
  I --> J["打印 [env] mail backend / [config] 检查结果"]
  J --> K["watchdog 每 1 秒刷新状态栏与按钮"]
  K --> L{"config.problems() 为空？"}
  L -->|"否"| M["日志逐条列出缺什么；Auto-watch 保持关闭"]
  L -->|"是"| N["WatchController.start()：1 秒后跑首轮"]
  N --> O["root.mainloop()"]
  M --> O
```

---

## 3. 登录与凭据缓存

```mermaid
flowchart TD
  S["connect(read_only?)"] --> T["get_token()：scopes = Mail.ReadWrite（默认）或 Mail.Read（自检）"]
  T --> U["读取 graph_token.json"]
  U --> V{"cached_token：未过期 且 scopes 与本次一致？"}
  V -->|"是"| W["直接复用 access_token<br/>（不发任何登录请求）"]
  V -->|"否"| X{"缓存里有 refresh_token + client_id？"}
  X -->|"否"| L0
  X -->|"是"| Y["POST /token<br/>grant_type=refresh_token"]
  Y -->|"200"| Z["_store_token：原子写<br/>access_token / expires_at / 轮换后的 refresh_token"]
  Y -->|"网络错误 或 5xx"| E1["MailboxUnavailable<br/>上层退避；绝不逼用户重新登录"]
  Y -->|"invalid_grant 等凭据失效"| L0["设备码登录"]
  L0 --> L1["候选 client_id：config 里的 → 微软公开客户端 ×3"]
  L1 --> L2["POST /devicecode 取 user_code"]
  L2 --> L3["日志单独一行打印设备码<br/>复制到剪贴板 + 自动打开浏览器"]
  L3 --> L4["每 interval 秒轮询 /token"]
  L4 -->|"authorization_pending"| L4
  L4 -->|"slow_down"| L5["interval += 5 秒"] --> L4
  L4 -->|"成功"| Z
  L4 -->|"15 分钟超时 / 管理员未批准"| E2["MailboxUnavailable<br/>提示重启或联系 IT"]
  Z --> W
  W --> X2["GraphSession(client)"]
  X2 --> X3["GET /me 验证账号<br/>失败 → MailboxUnavailable"]
  X3 --> X4["返回 (client, session)<br/>（连接日志每进程只打一次）"]
```

---

## 4. 自动监听：调度与一轮工作

```mermaid
flowchart TD
  T0["after(interval) 触发 _tick<br/>（默认 30 秒，范围 5–3600）"] --> T1{"watcher.running？"}
  T1 -->|"否"| TZ["结束，不再排队"]
  T1 -->|"是"| T2{"TaskRunner 正忙？"}
  T2 -->|"是"| T3["2 秒后重试本 tick"] --> T0
  T2 -->|"否"| T4["后台线程执行 _work"]

  T4 --> W1["connect() 建立 GraphSession"]
  W1 --> W2["_tag_reports：扫 30 天内缺 Sum 标签的自家报告 → 补标签"]
  W2 --> W3["newest_message + plan_tick(last_seen)"]
  W3 --> W4{"action"}

  W4 -->|"idle"| I1{"启动补分类未完成<br/>且 catchup_days > 0？"}
  I1 -->|"是"| CATCH["进入 _catchup（见第 5 节）"]
  I1 -->|"否"| DONE["本轮结束"]

  W4 -->|"baseline（首次运行）"| B1["只记录基线，不碰存量邮件"] --> I1

  W4 -->|"tag-report"| R1["给这封报告写 Sum 标签"] --> R2["推进基线"]

  W4 -->|"summarize（sum 指令）"| S1{"已经打过 Sum？"}
  S1 -->|"是"| S2["只推进基线（防重复出报告）"]
  S1 -->|"否"| S3["写 Sum 标签 → summarizer.summarize_inbox"]
  S3 --> S4["_run_commands：顺手处理同批遗漏的指令"] --> S5["推进基线"]

  W4 -->|"classify（有新邮件）"| K1["collect_new_mail：比 last_seen 新、最多 100 封<br/>跳过已分类 / 自家报告 / sum 指令 / ignored 名单"]
  K1 --> K2["_run_commands：扫出埋在下面的 sum 指令并执行"]
  K2 --> K3{"有待分类邮件？"}
  K3 -->|"否"| K4["推进基线 → 本轮结束"]
  K3 -->|"是"| K5["classifier.classify：分批 AI 判定"]
  K5 --> K6["apply_verdicts：写回分类与红旗"]
  K6 --> K7["remember_ignored：记住被判为非学校邮件的 EntryID"]
  K7 --> K8["advance_policy 决定基线"]
  K8 -->|"有失败且 < 3 轮"| K9["基线不动；下轮只重试失败的"]
  K8 -->|"连续 3 轮失败"| K10["放弃并推进，日志点名失败封数"]
  K8 -->|"全部成功"| K11["推进到本轮处理过的最旧一封"]

  DONE --> D1["on_done 回调（主线程）：清零错误计数、写日志、排下一轮"]
  K9 --> D1
  K10 --> D1
  K11 --> D1
  R2 --> D1
  S2 --> D1
  S5 --> D1
  CATCH --> D1
  D1 --> D2{"本轮处理过邮件 / 或补分类还没完？"}
  D2 -->|"是"| D3["2 秒后立刻再来一轮"]
  D2 -->|"否"| D4["等下一个轮询间隔"]
```

---

## 5. 启动补分类

```mermaid
flowchart TD
  A["_catchup（每次启动最多补一次，受 catchup_days 控制）"] --> B["read_inbox(catchup_days, limit=100)<br/>跳过已分类、自家报告、sum 指令、ignored"]
  B --> C{"有待分类邮件？"}
  C -->|"否"| D["标记补分类完成"]
  C -->|"是"| E["classifier.classify 分批判定"]
  E --> F["apply_verdicts 写回"]
  F --> G["remember_ignored"]
  G --> H{"本轮凑满 100 封？"}
  H -->|"是"| I["下一轮继续补<br/>（不推进基线，不打断新邮件处理逻辑）"]
  H -->|"否"| D
  E -->|"有失败"| J["日志说明失败封数；本轮结束，下轮继续"]
```

---

## 6. 分类流水线

```mermaid
flowchart TD
  A["入口：GUI Categorize / CLI classify / 监听 classify / 启动补分类"] --> B["connect() → GraphSession"]
  B --> C["读邮件头：最新 200 封（RECENT_MAX）<br/>跳过报告 / sum 指令 / 已分类 / ignored"]
  C --> D["对每封候选：fill_body 按需下载正文<br/>（瞬时故障→整轮退避；其它失败→留 bodyPreview）"]
  D --> E["按 batch_size = 40 分批"]
  E --> F["build_prompt：规则文本 + 每封<br/>显示名/地址/主题/正文前 400 + 末尾 300"]
  F --> G["ai.ask_json → parse_json"]
  G --> H{"返回条数正好等于本批？"}
  H -->|"否"| I["整批作废 → 对半拆分重试<br/>最多 3 层：40→20→10→5"]
  I --> G
  H -->|"是"| J["to_verdicts：级别归一 S/A/B/C、解析 deadline<br/>is_school 缺失时保守当作学校邮件"]
  J --> K["确定性规则（只升不降）：<br/>1 命中 courses → 至少 A<br/>2 命中 faculty → 至少 A<br/>3 发件域名命中 school_domains → 捞回为学校邮件"]
  K --> L{"dry_run？"}
  L -->|"是"| M["只打印 [S]/[A]/[B]/[C]，不写邮箱"]
  L -->|"否"| N["apply_verdicts：<br/>非学校邮件跳过（不写、不算失败）<br/>S 且加急 → 打红旗<br/>merge_categories 只替换自家标签 → PATCH"]
  N --> O["统计：read / ok / failed / skipped + S/A/B/C 计数"]
```

---

## 7. AI 调用层

```mermaid
flowchart TD
  A["ai.ask(prompt, profile, retries=3)"] --> B{"ai.active 存在？"}
  B -->|"否"| E1["AiError：名字写错 + 现有列表"]
  B -->|"是"| C["按 type 选实现：openai / anthropic / cli"]
  C --> D["_call_with_deadline：工作线程 + 墙钟超时<br/>（profile.timeout，默认 300s）"]
  D --> F{"HTTP 状态"}
  F -->|"200–299"| G["解析响应文本"]
  F -->|"400 且启用 json_mode"| H["去掉 response_format 重发一次"] --> F
  F -->|"429 / 5xx / 网络错误"| I["AiError(transient=True)"]
  F -->|"401 / 403 / 404"| J["AiError(transient=False)<br/>提示 Key / base_url / 模型名"]
  I --> K{"还有重试机会？"}
  K -->|"是"| L["退避 5s、10s… 后重试"] --> D
  K -->|"否"| E2["抛出 AiError（上层记录为该批失败）"]
  G --> M["ask_json：parse_json 容错<br/>裸 JSON / 围栏切片 / JSONL 流"]
  M --> N{"解析成功？"}
  N -->|"否"| O["重试 2 次；被截断时提示调大 max_tokens"]
  N -->|"是"| P["返回数据给调用方"]
  O --> Q["总结用宽松解析：ask_json_lenient<br/>截断时 salvage_array 收下已完整元素"]
  subgraph CLI["CLI 后端"]
    C1["subprocess.run(argv, input=prompt)"] --> C2["非 0 退出 → 判断 stderr 是否限流/超时"]
  end
  C --> CLI
```

---

## 8. 总结报告

```mermaid
flowchart TD
  A["触发：GUI Sum / 监听发现 sum 指令邮件 / CLI summarize"] --> B["parse_letters：sabc 顺序、默认 sabc"]
  B --> C["summarize_inbox：connect → read_recent(days=Scope, letters)"]
  C --> D["遍历最新 200 封：level_hint 在 letters 内 → fill_body 取正文"]
  D --> E{"有匹配邮件？"}
  E -->|"否"| F["不发送报告，日志说明"]
  E -->|"是"| G["按 batch_size = 40 分批生成要点"]
  G --> H["ask_json_lenient → 形状归一<br/>数组 / 单键对象 / 字符串包 JSON"]
  H --> I{"拿到数量 ≥ 本批？"}
  I -->|"是"| J["截断取前 n 个"]
  I -->|"部分"| K["收下已有的，只对剩余再要一次"]
  I -->|"无"| L["对半拆分递归；最小到单封 → 该封要点留空"]
  K --> M
  L --> M
  J --> M["build_html：按 S→A→B→C 分组<br/>标题用 Graph webLink（没有则纯文本）<br/>时间本地化、正文要点 HTML 转义"]
  M --> N["build_subject：语言标记（邮件总结 / Email summary）<br/>+ 级别 + 最近 N 天"]
  N --> O["send_html → POST /me/sendMail（saveToSentItems）"]
  O --> P["下一轮监听识别出这封报告并打 Sum 标签<br/>（is_report_subject 中英标记都认）"]
```

### 6.1 一次 GUI 分类的时序（补充）

```mermaid
sequenceDiagram
  autonumber
  actor U as 用户
  participant MW as 主窗口
  participant TR as TaskRunner
  participant CL as classifier
  participant AI as ai
  participant G as graph
  participant M as Microsoft Graph
  U->>MW: 点 Categorize（非 dry-run 时先弹确认框）
  MW->>TR: run(classify_inbox)
  TR->>TR: 后台线程 + RunToken（代次令牌）
  TR->>CL: classify_inbox(cfg, days, ...)
  CL->>G: connect()
  G->>M: 读/刷新 token（按需）
  G-->>CL: (client, session)
  CL->>G: read_inbox(days)
  G->>M: GET 收件箱邮件头
  loop 每封待判级邮件
    G->>M: GET /me/messages/{id}?$select=body
  end
  CL->>AI: ask_json(批次提示词)
  AI-->>CL: JSON（含超时/重试）
  CL->>CL: to_verdicts + 确定性规则
  CL->>G: write_categories(mail, level, flag)
  G->>M: PATCH /me/messages/{id}
  CL-->>TR: stats
  TR-->>MW: on_done（主线程回调）
  MW-->>U: 日志 + 状态栏
```

---

## 9. 监听状态机与基线推进

```mermaid
stateDiagram-v2
  direction LR
  [*] --> Startup
  Startup --> Baseline: 首次运行
  Startup --> Idle: 收件箱为空
  Baseline --> Watching: 记录基线
  Watching --> Idle: 无新邮件
  Watching --> Classify: 普通新邮件
  Watching --> Summarize: sum 指令
  Watching --> TagReport: 自家报告缺 Sum
  Classify --> Watching: 全部成功
  Classify --> Classify: 失败 <3 轮
  Classify --> Watching: 连续 3 轮失败
  Summarize --> Watching: 报告已发
  TagReport --> Watching: 补标签
  Idle --> Watching: 本轮结束
  note right of Classify
    Force reset 时立即停止
    写邮箱且不推进基线
  end note
```

---

## 10. 失败与退避

```mermaid
flowchart TD
  A["发生异常"] --> B{"属于哪一类"}
  B -->|"MailboxUnavailable<br/>网络 / 429 / 5xx / 登录服务不可用"| C["_on_error：不计失败次数<br/>只提示一次；120 秒后再试"]
  B -->|"其它异常（配置/写回/代码错误）"| D["_errors += 1；日志 (n/3)"]
  D --> E{"n ≥ 3？"}
  E -->|"是"| F["自动暂停监听<br/>（用户需手动重新打开 Auto-watch）"]
  E -->|"否"| G["按当前轮询间隔重试"]
  C --> H["恢复后自动接上，日志说明邮箱已可达"]

  I["单封写回失败"] --> J["failed += 1；基线不推进（未达 3 轮）"]
  J --> K["下一轮：已贴标签的会被跳过，只重试失败的"]

  L["AI 批次失败（条数不符 / 调用失败）"] --> M["整批作废，对半拆分重试"]
  M --> N{"最小到单封仍失败？"}
  N -->|"是"| O["该封不写回、不猜级别；只影响这一封"]
```

---

## 11. 数据与文件

```mermaid
flowchart LR
  CFG["config.ensure / save<br/>原子写"] --> C1
  GRAPH["graph.get_token<br/>原子写"] --> C2
  WATCH["watcher._persist<br/>原子写"] --> C3
  LOGV["log_view"] --> C4

  subgraph LOCAL["本地运行时文件（.gitignore，绝不进 zip）"]
    C1["config.json<br/>身份 / 行为 / graph / AI 配置"]
    C2["graph_token.json<br/>access_token + expires_at + refresh_token"]
    C3["watch_state.json<br/>last_entryid / interval / ignored 名单"]
    C4["app.log<br/>启动时超过 1 MB 清空"]
  end

  C1 --> R1["read_only 自检：只用 Mail.Read<br/>不会覆盖写回用 token（scopes 校验）"]
  C2 --> R2["过期前 5 分钟自动刷新"]
  C3 --> R3["关掉程序期间的新邮件<br/>下次启动会补上"]
```

---

## 12. 构建 / 运行 / 发布入口

```mermaid
flowchart LR
  A["start.bat"] --> B["py -3 app.py（源码运行）"]
  C["build.bat"] --> D["PyInstaller --onedir --windowed"]
  D --> E["dist/EmailManager → EmailManager/"]
  E --> F["压缩 EmailManager.zip（9.7 MB / 985 项）"]
  F --> G["GitHub Release v3.0.0"]
  B --> H["GUI 主窗口"]
  G --> I["用户解压 → 双击 EmailManager.exe"]
  I --> H
  J["run.py"] --> K["check / providers / classify / summarize（CLI）"]
  J --> H

  subgraph HELP["命令行"]
    K1["check：配置 / AI / 邮箱连通（只读）"]
    K2["providers：列出 AI 后端与状态"]
    K3["classify --days --dry-run --force"]
    K4["summarize sa / sabc"]
  end
  K --> HELP
```

---

## 附录 A：关键常量

| 常量 | 值 | 作用 |
|---|---|---|
| `DEFAULT_INTERVAL` | 30 s | 监听轮询间隔（可调 5–3600） |
| `MAX_MAIL_PER_ROUND` | 100 | 一轮最多处理的邮件数 |
| `CATCHUP_DAYS` | 7 天 | 启动补分类 / 新邮件回溯窗口 |
| `MAX_FAILED_ROUNDS` | 3 | 同一批连续失败几轮后放弃推进 |
| `MAX_CONSECUTIVE_ERRORS` | 3 | 连续错误多少次自动暂停监听 |
| `MAILBOX_RETRY_DELAY` | 120 s | 邮箱不可用时的退避 |
| `REPORT_TAG_DAYS` | 30 天 | 补 Sum 标签的回溯窗口 |
| `RECENT_MAX` | 200 | 一次读取的邮件头上限 |
| `RECENT_TTL` | 10 s | 同一轮内查询结果复用的时长 |
| `TOKEN_MARGIN` | 300 s | access_token 提前刷新余量 |
| `MAX_ATTEMPTS` | 3 | Graph GET 限流/5xx 的最大尝试次数 |
| `batch_size` | 40 | 每次送 AI 的邮件数（config 可改） |
| `BODY_CHARS` / `BODY_TAIL_CHARS` | 400 / 300 | 判级时送给模型的正文头/尾字符数 |
| `BODY_CHARS`（总结） | 800 | 生成一句话要点时的正文长度 |
| `urgent_days` | 14 天 | S 级「截止在 N 天内」阈值 |
| `summary_days` | 7 天 | 总结报告默认回溯窗口 |
| `SPLIT_MIN` / `SPLIT_MAX_DEPTH` | 2 / 3 | 批次失败拆分的最小批量 / 最大层数 |

## 附录 B：失败语义速查

| 场景 | 行为 |
|---|---|
| 模型返回条数与请求不符 | **整批作废**，对半拆分重试；绝不错位贴标签 |
| 非学校邮件 | 不写回、不计失败；EntryID 记入 `watch_state.json` 的 ignored |
| 单封写回失败 | 计失败；基线不推进（< 3 轮），下轮只重试失败的 |
| 连续 3 轮失败 | 放弃并推进基线，日志点名，避免无限重试烧 AI 额度 |
| 监听基线推进 | 只在「本轮真的处理完」后推进；Force reset 中断则不推进 |
| 邮箱不可用（网络/429/5xx） | 不计失败、不暂停；120 秒退避后自动重连 |
| 断网 | 不触发重新登录（只有 refresh token 真的失效才走设备码） |
| 写回范围 | 只替换程序自己的标签（S/A/B/C/Sum）；用户分类与服务器规则分类保留 |
| 分类主列表 | **从不创建、删除、读取**；S/A/B/C/Sum 直接写在邮件上 |
| 报告邮件 | 靠主题标记识别并打 Sum，自己不会被再分类 |
