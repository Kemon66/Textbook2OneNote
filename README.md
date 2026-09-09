---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: 24a9a0152c7e8926064293f264998d15_0a8a3ba9ab9711f18874525400287e28
    ReservedCode1: O1Z5rHMXe7viBu1UWETjrVIQpp4X+WUydS26zMo0Hv98b+OqU95L0/LM3D16BChhYJx/SLLjgzwQoDVHF3HwXs1k8HZKBbvlma2yfxdxr53j5klUzh2J7VsjqfwwmmM7NHew0fHFITiDPRBZcqJwfo/VFdAtBxq8NvA1O5idgtqQV3V1la4AXhOZTAI=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: 24a9a0152c7e8926064293f264998d15_0a8a3ba9ab9711f18874525400287e28
    ReservedCode2: O1Z5rHMXe7viBu1UWETjrVIQpp4X+WUydS26zMo0Hv98b+OqU95L0/LM3D16BChhYJx/SLLjgzwQoDVHF3HwXs1k8HZKBbvlma2yfxdxr53j5klUzh2J7VsjqfwwmmM7NHew0fHFITiDPRBZcqJwfo/VFdAtBxq8NvA1O5idgtqQV3V1la4AXhOZTAI=
---

# Textbook2OneNote

教材自动导入 OneNote 工具：识别 PDF / Word 教材目录层级，生成可直接导入 OneNote 的 .one 笔记本（OfficeIMO 离线引擎，免装 .NET 运行时）。

## 使用方式

### 方式一：双击 Textbook2OneNote.exe（推荐，免安装环境）
1. 双击 `Textbook2OneNote.exe` 打开界面；
2. 将教材文件（PDF / Word）**直接拖入界面中央的拖放区**，或点击拖放区选择文件；
3. 输出目录默认 OneDrive 桌面（便于移动端同步），可自行修改；
4. 点「开始导入」，完成后在输出目录得到 `书名` 笔记本文件夹（含 `.one` 分区与 `.onetoc2` 结构）。

### 方式二：命令行走 importer.py（需 Python 3.9+ 与依赖）
```
pip install pymupdf python-docx
python importer.py 教材.pdf
python importer.py 教材.pdf --out D:\输出目录
```
无参数运行会弹出文件选择框。

## 文件说明
| 文件 | 作用 |
|------|------|
| `Textbook2OneNote.exe` | 免环境 GUI 入口（已打包） |
| `gui.py` | GUI 源码 |
| `importer.py` | 核心逻辑：目录解析 / 页码偏移 / 渲染 / 调 notegen |
| `notegen.exe` | .one 生成引擎（OfficeIMO 离线构建，自包含 70MB） |
| `run.bat` | 本机装有 Python 时双击启动 GUI |

## 输出结构
```
输出目录/
└── 书名/
    ├── Open Notebook.onetoc2
    └── PART xx .../
        ├── Open Notebook.onetoc2
        └── Chapter xx ....one
```
在 OneNote 中「打开备份」选择该文件夹即可载入。

## 已知限制
- PDF 目录依赖「Brief Contents / Contents」页识别，无目录或纯扫描件无法使用；
- Word 需使用「标题 1/2/3」样式组织层级。
*（内容由AI生成，仅供参考）*
