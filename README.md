不打断思路的情况，通过快捷键快速记录文本、图片到excel作为知识库。 

**痛点：** 
- 日常工作生活中会与不同的桌面程序打交道完成复杂任务，记录的信息是分散的！收藏的文章在浏览器中，代码在编辑器中。
- 之前做过的复杂任务，当我们过一段时间再次重复的时候，大脑没有多少印象了，又要重头开始，效率低下。


 **解决方案：**
- 好记性不如烂笔头，用户通过快捷键记录重要的文字和图片信息到excel知识库。
- 按天存储。
- 粗略的用户ToDo List管理
- excel表格天然支持层级显示，这样看起来层次清楚。
- excel作为存储和展示：tab（按天管理搜集的素材）；表格提供了层级结构；支持手机、平板跨平台共享，天然的（一个人开发，我就不需要再做轮子）。
    - ctrl+\：打开excel
    - ctrl+q: 关闭excel
- 全局的快捷键帮助我们快速查看、截取文字和图片
    - ctrl+C两次: 保存剪切板的信息
    - ctrl+shift+enter: 一键保存所选中文字+截屏+标题+链接（chrome）
    - PrintScreen/alt+PrintScreen两次：保存图片
    - ctrl+enter: 如果信息已经放在剪切板，保存到剪切板
    - alt+/: 查看最近一次保存的信息
    - alt+left/right: 跳转层级
    - alt+up/down: 往上下一行（增量式添加）
- 气泡提醒保存进度，不打断用户思考。提醒用户保存了哪些文字和图片
- TODO: 任务和记录结合
- TODO: 结合向量数据库模糊检索匹配的文字和图片，实现快速查找定位。
- TODO: 问题的解决过程通过训练本地LLM，不存在泄漏公司个人信息。
- TODO: 通过录屏来，ocr来解决主动记录所遗漏的重要信息。


# shouyu
Quickly record the content (text & image) of clipboard to MS/WPS Excel file by using hot keys. Suitable for users whose record habits and currently only support Windows users.


# Cases
- When users are studying a complex problem, they often need to take notes quickly without being disturbed, but all note-taking tools on the market need to switch to another interface to paste and copy, which causes the user's thinking to be interrupted. shouyu provides a shortcut to save, using the bubble pop-up box does not disturb the user's thinking.
- New tab records are generated every day in a tree hierarchy to make the timeline clear and easy to retrieve.


# Features
- Please refer to [kb.ini](kb.ini) to set/change excel path and shortcuts.
- <img src="resources/screenshort/ui.png" alt="excel UI" title="Excel UI">
- <img src="resources/screenshort/bubble_msg_box.png" alt="Bubble message box" title="Bubble message box">
- <img src="resources/screenshort/img_bubble_msg_box.png" alt="Bubble message box for image" title="Bubble message box for image">
- <img src="resources/screenshort/tray.png" alt="Tray" title="Tray">
