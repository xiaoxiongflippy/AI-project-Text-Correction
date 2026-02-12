# FormatPilot（LLM 文本排版优化器）

一个可产品化交付的本地桌面软件，用于处理“大模型回答复制后格式混乱”的问题。

## 你可以做什么

- 清理 Markdown 符号（标题、加粗、链接、代码标记）
- 智能消除重复噪声（如“主题:主题:主题”叠词、重复标点）
- 统一中英文标点与空白
- 自动优化中英文与数字间距（如“支持Python3.11”→“支持 Python 3.11”）
- 合并被错误断开的句子和段落
- 段首自动空两格（中文文档更美观）
- 可选移除 Emoji
- 桌面界面一键优化 + 复制结果
- 命令行批量处理文本文件
- 产品化工作台界面（状态栏、快捷键、导出专区）
- 配置自动记忆（上次参数、窗口大小、打开目录）
- 现代化 UI（侧边栏 + 卡片布局 + 深浅色主题）
- 发布预览视图（编辑 / 发布一键切换）
- 质量评分与交付状态提示


## 推荐环境（重要）

为了避免你遇到的 `lxml/pillow` 编译报错，推荐固定使用 **Python 3.11**。

### 手动初始化（Windows）

```bash
py -3.11 -m venv .venv311
.\.venv311\Scripts\python.exe -m pip install -U pip setuptools wheel
.\.venv311\Scripts\python.exe -m pip install -r requirements-py311.txt
```

## 运行方式

### 1) 桌面软件（推荐）

```bash
python run.py
```

桌面版额外依赖：

```bash
pip install customtkinter
```

桌面版快捷键：
- `Ctrl+Enter`：一键优化
- `Ctrl+Shift+C`：复制结果
- `Ctrl+S`：保存 TXT

### 2) 命令行模式

```bash
python run.py cli --input raw.txt --output clean.txt
```

导出为 Word / PDF：

```bash
python run.py cli --input raw.txt --export-word result.docx --export-pdf result.pdf
```

也可以直接传文本：

```bash
python run.py cli --text "### hello **world**"
```

## 命令行参数

- `--text`：直接传入文本
- `--input`：输入文件路径
- `--output`：输出文件路径
- `--keep-markdown`：保留 Markdown 标记
- `--keep-lines`：保留原始换行
- `--remove-emoji`：移除 Emoji
- `--no-indent`：关闭段首空两格
- `--export-word`：导出为 Word（`.docx`）
- `--export-pdf`：导出为 PDF（`.pdf`）

## 导出依赖

- Word 导出：`pip install python-docx`
- PDF 导出：`pip install reportlab`

建议一次安装完整依赖（若你不用一键脚本）：

```bash
pip install -r requirements-py311.txt
```

## 文件说明

- `run.py`：统一入口（GUI / CLI）
- `requirements-py311.txt`：Python 3.11 稳定依赖锁定
- `app_gui.py`：产品化桌面界面（工作台）
- `app_config.py`：本地配置持久化
- `text_cleanup.py`：文本清洗核心逻辑
- `export_utils.py`：Word/PDF 导出器（保留段落与列表格式）

## 产品定位

- 本项目聚焦“AI 回复内容格式优化”，不做模型改写或文案生成。
