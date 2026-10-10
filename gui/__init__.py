"""gui — Tkinter 界面层。

只是一层「外壳」：把 config.py / classifier.py / summarizer.py 的能力包成
按钮和表单，业务逻辑一行都不在这里。

    theme.py              视觉常量与 ttk 样式
    log_view.py           「运行过程」日志区（线程安全）
    task.py               后台任务执行器（队列回主线程 + COM 初始化）
    main_window.py        主窗口：User info / Categorize / Sum / AI 设置
    user_info_dialog.py   User info 编辑表单
    ai_settings_dialog.py AI 后端设置（模型 / API Key / 测试连通）
"""
