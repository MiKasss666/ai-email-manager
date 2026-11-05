# Changelog · 版本历史（详细提交记录见 git log 与 Releases 页面）

## v3.0.0 — 2026-11-05

### Breaking changes 破坏性变更

- **移除旧版 Outlook 经典版（COM）后端，只保留 Microsoft Graph。**
  不再需要 pywin32、不再需要桌面版 Outlook、不再有「完整性级别 / Server execution failed」这类问题。
  `config.json` 里的 `mail_source` 字段被移除（旧值会被忽略并在下次保存时丢弃）。
- **不再创建或删除 Outlook 分类主列表条目。**
  S/A/B/C/Sum 仍然直接写在邮件本身上；但程序不再调用 `/me/outlook/masterCategories`，
  也不再需要 `MailboxSettings.ReadWrite` 权限。若想在 Outlook 分类列表里看到这 5 个分类，
  手动新建一次即可。
- **删除 `check_context.bat` 与 `EmailManager.spec`**（前者是 COM 诊断脚本，后者是死掉的第二套构建定义）。

### Added 新增

- `core/mailmeta.py`：邮件约定与纯逻辑的唯一真源（分类名、报告/指令识别、写回合并规则、邮件字典形状）。
- `core/version.py`：版本号唯一真源（界面 / `run.py` / CITATION 从这里对齐）。
- `CHANGELOG.md`：版本历史。
- `graph_token.json` 现在缓存 **access_token + 过期时间**；未过期直接复用，不再每轮刷新 refresh token。
- AI 设置里的 API Key 默认以圆点遮蔽，可勾选临时显示。

### Changed 行为改进

- **轮询只拉邮件头**（主题/发件人/时间/分类等）；只有真正要判级或写总结的邮件才单独下载正文。
- **429 / 5xx 按 `Retry-After` 退避重试**，并归类为「邮箱暂时不可用」：不计失败、不会因此暂停监听。
- 断网不再被误判成「登录失败」而弹出设备码登录；只有 refresh token 真的失效才要求重新登录。
- `watch_state.json` 与 `graph_token.json` 改为原子写，崩溃不会留下半个文件。
- 界面与日志不再提及 Outlook / COM；窗口标题显示版本号。

### Fixed 修复

- 旧 COM 后端里「Exchange 内部发件人拿不到 SMTP 地址」导致学校域名规则失效的问题（随后端移除而消失）。
- 旧 COM 后端五处收件箱遍历没有逐字段兜错、一封坏邮件可拖垮整轮的问题（随后端移除而消失）。
- AI 设置窗口保存失败时会提前改写内存中的共享配置（保存失败后状态与磁盘不一致）。
