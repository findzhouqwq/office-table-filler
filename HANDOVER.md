# 交接文档：表格批量填写工具（Windows GUI）

> 面向接手维护的人。先看「快速开始」，再看「架构」和「扩展点」。

## 1. 这是什么

一个 Windows 桌面工具，用图形界面把各种来源的数据批量填进表格 / 文档。
主要适配 **Microsoft Office**（Excel / Word / PowerPoint），同时适配 **WPS Office** 与 **OnlyOffice**。

三种填表方式：

| 页面 | 功能 | 典型场景 |
| --- | --- | --- |
| ① 图片识别填表 | 批量图片 OCR + 表格线检测，整表或按字段填入 Excel | 发票、证件、报关单、表格截图批量录入 |
| ② 文字填表 | 粘贴文本 / CSV / TXT / JSON / XLSX → Excel | 从聊天记录、报告里抽字段入表 |
| ③ 跨文件填表 | 多源文件 → 多目标文件（Excel / Word / PPT） | 用 Excel 名单批量生成 Word 入职表、PPT 人员页 |

## 2. 快速开始

### 2.1 环境

- Windows 10/11
- Python 3.10+（推荐 3.11 / 3.12，已在 3.12 验证）
- 联网安装依赖（首次会下载 RapidOCR 的中文模型）

### 2.2 安装

```bat
cd office-table-filler
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

或直接双击 `run.bat`（自动建虚拟环境 + 装依赖）。

### 2.3 环境自检

```bat
python main.py --check
```

输出表格引擎、OCR 引擎和各模块是否可用。

### 2.4 跑演示

```bat
python make_samples.py
python demo.py
```

`make_samples.py` 生成样例文件（含一张带表格线的合成图片）；`demo.py` 依次演示三种填表方式，结果写到 `demo_output/`。

### 2.5 跑测试

```bat
python -m unittest discover -s tests -v
```

共 34 个用例：工具函数、列映射、目标解析、表格引擎、三种填表、GUI 冒烟、真实 OCR 端到端。
OCR 集成用例在缺少 `rapidocr-onnxruntime` 时自动跳过。

## 3. 架构

```
office-table-filler/
├── main.py                 入口：--check / 启动 GUI
├── make_samples.py         生成演示样例
├── demo.py                 三种填表方式的端到端演示
├── run.bat / build.bat     一键运行 / 打包
├── requirements.txt
├── app/
│   ├── core/               引擎层（与 UI 无关，可单独调用）
│   │   ├── spreadsheet.py  表格引擎：openpyxl 直接读写 + Excel/WPS COM
│   │   ├── data_io.py      读取 xlsx/csv/txt/json/docx/pptx
│   │   ├── ocr.py          OCR 引擎抽象：RapidOCR / Tesseract / Windows OCR
│   │   ├── table_detect.py OpenCV 表格线检测 + 词聚类
│   │   ├── mapping.py      字段提取规则 + 列匹配
│   │   ├── targets.py      跨文件目标位置解析（紧凑字符串）
│   │   ├── filler.py       三种填表方式的编排（核心）
│   │   ├── models.py       数据结构
│   │   ├── errors.py       用户可见异常
│   │   └── utils.py        文本/数字/日期工具
│   └── ui/                 PySide6 界面
│       ├── main_window.py
│       ├── image_tab.py    图片识别填表
│       ├── text_tab.py     文字填表
│       ├── cross_tab.py    跨文件填表
│       ├── mapping_widgets.py  规则/映射表格控件
│       └── common.py       文件对话框、后台线程
└── tests/
```

### 3.1 Office 适配原理（两层引擎）

1. **直接读写引擎（默认，始终可用）**
   用 openpyxl / python-docx / python-pptx 直接读写 OOXML 文件。
   Office、WPS、OnlyOffice 都用同一套 `.xlsx/.docx/.pptx` 格式，所以**三家通吃**，
   不需要 Office 在运行。

2. **COM 引擎（可选，自动检测）**
   检测到已安装 Microsoft Excel 或 WPS 表格时，通过 Windows COM 驱动真实程序。
   适合保留图表/公式/复杂格式，或需要“开着 Excel 实时刷新”的场景。
   优先 COM，失败回退直接读写。

> 建议默认用直接读写；只有文件含图表/透视表或要实时刷新时才依赖 COM。

### 3.2 关键设计约定

- **坐标一致性**：OCR 词框与表格线检测必须在同一张（预处理后）图上，否则词会串列。
  `OcrEngine.recognize_with_image()` 返回识别结果 + 预处理图，网格检测复用同一张图。
  （曾经的坑：在原始图上找线、在放大图上取词，导致词全部错位。）

- **填表模式**：
  - `grid`：每张图一个完整表格，检测行列后整表写入。
  - `fields`：每张图提取字段，按列映射写一行。

- **跨文件映射**：一条映射 = `源字段 → 目标类型 + 位置参数`。
  位置参数示例：
  - Excel 单元格：`Sheet1!B2`（可加 `步进=1` 让多条记录逐行下移）
  - Word 单元格：`table=0,row=1,col=2`
  - Word 占位符：`{{姓名}}`（正文里写 `{{姓名}}`）
  - PPT 单元格：`slide=0,table=0,row=1,col=1`

  按目标文件扩展名过滤映射（`_KIND_EXTENSIONS`），互不兼容的目标自动跳过。

- **批量语义**：
  - `batch_by_row=True`：每个源记录生成一份目标文件（`template_001.docx` 等）。
  - `batch_by_row=False`：所有记录写进同一个文件。

## 4. 常见问题

| 现象 | 原因 / 处理 |
| --- | --- |
| 图片识别填进去的内容串列 | 图片有透视变形，或预处理放大导致词框与线不在同一坐标系。先用规则字段模式，或把图片裁正。 |
| 识别不出内容 | 确认 OCR 引擎已安装（`--check`）。默认 RapidOCR，中文效果好。 |
| 照片表格识别差 | 表格线检测对规整扫描/截图效果好；照片建议先截图裁正。 |
| 图表/透视表丢失 | 直接读写引擎不保留复杂对象；改用 COM 引擎或装 Office/WPS。 |
| pip 装依赖很慢 | 国内网络可加镜像：`pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple` |

## 5. 已知限制

- 图片为照片/透视变形时，网格检测退化为按词聚类，建议先裁正。
- 直接读写引擎保留单元格值与样式，但复杂对象（图表、透视表）可能丢失。
- Word/PPT 表格行追加依赖 python-docx / python-pptx；复杂模板请在目标程序中复核。
- `.xls`（旧版二进制）不支持，请另存为 `.xlsx`。
- COM 引擎需要已安装 Office 或 WPS，且未在本机验证（本机未装 Office）。

## 6. 扩展点（建议的下一步）

1. **透视校正**：在 `table_detect.detect_grid` 前加一步四点透视变换，提升照片表格准确率。
2. **模板保存/加载**：把映射规则导出为 JSON，便于复用（当前 `MappingDoc` 已定义结构但未接 UI）。
3. **更多 OCR**：接入 PaddleOCR / 百度/腾讯云 OCR，替换 `ocr.py` 中的引擎即可。
4. **Excel 公式/样式保护**：直接读写时对公式单元格做保护标记，避免误覆盖。
5. **目标单元格可视化选择**：跨文件填表目前靠手填位置参数，可做点击选单元格。
6. **OnlyOffice 专用适配器**：OnlyOffice 无 COM，可加其文档服务器 API 适配（当前走 OOXML 已可用）。
7. **打包**：`build.bat` 用 PyInstaller 打包为单 exe；注意 OCR 模型要用 `--collect-all rapidocr_onnxruntime`。

## 7. 推送到 GitHub

本机未安装 git，也不保存 GitHub 凭据。请自行安装 git 后执行：

```bat
cd D:\work\office-table-filler
git init
git add .
git commit -m "Initial commit: Windows Office/WPS/OnlyOffice table batch filler (GUI)"
git branch -M main
git remote add origin https://github.com/<你的用户名>/office-table-filler.git
git push -u origin main
```

仓库地址模板：`https://github.com/<你的用户名>/office-table-filler.git`

`.gitignore` 已排除 `.venv/`、`demo_output/`、`samples/`、日志与临时文件，仓库保持干净。

## 8. 目录地址

- 项目根目录：`D:\work\office-table-filler`
- 虚拟环境：`D:\work\office-table-filler\.venv`
- 样例数据：`D:\work\office-table-filler\samples`
- 演示输出：`D:\work\office-table-filler\demo_output`
