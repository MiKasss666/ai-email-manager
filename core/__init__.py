"""core — Email Manager 的内核。

两支核心能力：
    classifier   邮件分级：邮件 → S/A/B/C + 截止日期，并可写回邮件分类
    summarizer   邮件总结：最近 N 天已分级邮件 → 一句话要点 + HTML 报告发给自己

公共支撑：
    config       配置（读写 config.json，纯标准库）
    ai           AI 传输层（后端无关，换服务商只改配置）
    mail         邮件后端门面（当前唯一实现：Microsoft Graph）
    mailmeta     邮件约定与纯逻辑（分类名 / 报告与指令识别 / 写回合并规则）
    version      版本号唯一真源

分层：config ← (ai, graph/mail) ← (classifier, summarizer) ← gui / run.py
这里不做 eager import，导入 core.config 不会连带建立网络连接或创建窗口。
"""
