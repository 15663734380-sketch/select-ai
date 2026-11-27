# 划词问答

一款 Windows 划词问答小工具：在任意软件里选中文字即可向 AI 提问；支持加载 PDF / Word 等文档追问、DeepSeek 联网搜索与自定义提示词，最小化后自动暂停划词；内置可拖拽缩放的鲸鱼娘桌宠，问答历史自动保存并支持导出。

## 功能

- 划词即问，AI 边生成边显示
- 加载 PDF / Word / TXT / Markdown 作为文档上下文
- 在回答里继续划词追问
- 历史问答自动保存，可导出 Markdown / Word
- Ctrl+Alt+Q 暂停 / 恢复
- 提示词（决定回答风格）

## 运行

安装依赖后执行 `pythonw select_ai.py`，首次使用在「设置」里填写 API Key。

## 其他
https://platform.deepseek.com/usage；  在这个平台上买API key然后用Flash模型便宜还好用，推荐一下

为了防止内存不断叠加增大所做的轻量化处理，历史问答仅保留近 100 轮对话，当有超过 100 轮的新一轮问答时，最早的一轮问答会被删除掉
