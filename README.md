
    Selenium: 用于Web自动化。

    安装命令:

    shell

pip install selenium

python-docx: 用于创建和修改Word文档。

安装命令:

shell

pip install python-docx

PyAutoGUI: 用于GUI操作，如屏幕截图等。

安装命令:

shell

    pip install pyautogui

ChromeDriver，将ChromeDriver的路径设置为实际的路径
（chrome打开页面可以使用百度搜索引擎打开）

在Windows上：

可以使用任务计划程序来设置定时任务。

    打开“任务计划程序”。
    点击“创建基本任务...”。
    设置触发器为每天，然后设置重复任务的间隔为1小时，并在“高级设置”中限定重复任务时间为9:00 AM到6:00 PM。
    在“操作”步骤中，选择“启动程序”，并浏览到您的Python解释器的路径，通常是python.exe，在参数中添加您的脚本路径
（建议每个时间段进行测试，防止出现问题）
