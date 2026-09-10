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
    ├── PART xx .../           ← 目录覆盖的正文章节（PART/CHAPTER/小节）
    └── PART xx 其他内容/       ← 自动识别被目录删减的页并分区分组
        ├── 目录 Contents        （目录页）
        ├── 前言 Preface         （前言/序/致谢）
        ├── 附录 Appendix
        ├── 尾注 Endnotes / 术语表 Glossary
        └── 索引 Index / 其他内容 Other
```
在 OneNote 中「打开备份」选择该文件夹即可载入。

## 删减内容整理（v0.1 新增）
未被目录层级覆盖的 PDF 页不会丢弃，程序会自动：
1. **书首区**（封面/版权/前言/目录等最早页）按语义归为「前言 Preface / 目录 Contents / 其他内容 Other」；
2. **书尾区**（附录/尾注/术语表/索引/后记等末尾页）通过**页眉里程碑**切段：检测到 `Endnotes / Glossary / Index / References / 尾注 / 索引 / 附录` 等独立成行的标题词时自动切出新分区；
3. 同类的多个不相连片段合并为一个分区（分区内分 multiple 页片段展示）；
4. 全部归入新增分区组「其他内容」并追加到笔记本末尾。

> 说明：正文中间因目录页码偏移而未被覆盖的散页不计入删减分区（会打印日志提示数量），避免把正文误当删减内容。

## 扫描版 / 图片型 PDF 处理（v2.0.1 新增）
对每页一张图、无文字层（或仅残留极少量 OCR 碎片）的扫描版教材，程序不再笼统报「未识别到有效目录层级」，而是：
1. 自动检测文本层：以内嵌大纲（Outline / get_toc）为强信号，并**全文多页采样**统计「有实质文本页占比 / 单页文本长度分布」；
2. 判定为扫描版时给出明确诊断：`检测到扫描版/图片型 PDF（无文字层）… 请先对教材进行 OCR 处理，或更换带文字层的版本后重试`，并附带文本层统计（采样页数 / 有文本页占比 / 总字符 / 单页文本中位数）；
3. 含文字层但前 40 页未定位到目录页（Contents / 目录）的教材，单独提示缺少目录页，不再与扫描版混淆。

## 已知限制
- 目录识别基于语义（PART/CHAPTER/UNIT、`N > Title`、`N Title`、`N-N 小节`、全大写主节），不依赖字号与排版本，兼容绝大多数有文本层的教材目录；
- PDF 依赖「Brief Contents / Contents（含中文“目录”）」页定位；扫描版/图片型（无文字层）需先 OCR 或更换带文字层版本后才能识别目录；
- 带文字层但缺少目录页的教材无法自动定位章节层级，需提供包含目录页的版本；
- Word 需使用「标题 1/2/3」样式组织层级。
*（内容由AI生成，仅供参考）*
