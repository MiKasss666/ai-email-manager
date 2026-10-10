"""core — Outlook 经典版 AI 邮件管理助手的内核。

两支核心能力：
    classifier   邮件分级：邮件 → S/A/B/C + 截止日期，并可写回 Outlook 分类
    summarizer   邮件总结：最近 N 天已分级邮件 → 一句话要点 + HTML 报告发给自己

公共支撑：
    config       配置（读写 config.json，纯标准库）
    ai           AI 传输层（claude CLI，换后端只改这一个文件）
    outlook      Outlook COM 原子操作（连接 / 分类 / 读信 / 写回 / 发信）

分层：config ← (ai, outlook) ← (classifier, summarizer) ← run.py
这里不做 eager import，因此 import core.config 不会连带加载 win32com。
"""
