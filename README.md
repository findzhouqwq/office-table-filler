# 表格批量填写工具（Windows GUI）

一个 Windows 桌面工具，用图形界面把各种来源的数据批量填进表格 / 文档。
主要适配 **Microsoft Office**（Excel / Word / PowerPoint），同时适配 **WPS Office** 与 **OnlyOffice**。

## 三种填表方式

| 页面 | 功能 | 典型场景 |
| --- | --- | --- |
| ① 图片识别填表 | 批量选择图片，OCR 识别后填进 Excel；支持整张表格的**网格识别**，也支持**按字段提取** | 身份证、发票、报关单、表格截图批量录入 |
| ② 文字填表 | 粘贴文本 / 导入 CSV、TXT、JSON、XLSX，按列映射填进 Excel | 从聊天记录、报告、文本里抽字段入表 |
| ③ 跨文件填表 | 从一个或多个源文件读数据，按字段映射写入多个目标文件（Excel / Word / PPT） | 用 Excel 名单批量生成 Word 入职表、PPT 人员页 |

## 日常办公增强

| 能力 | 说明 |
| --- | --- |
| 模板保存 / 复用 | 每个页面顶部都有模板栏：保存、应用、删除、导入、导出 JSON 模板，重复任务一键还原 |
| 生成前预览 / 校验 | “预览 / 校验”按钮先展示将要写入的行、目标列和问题清单；有错误时禁止执行 |
| 自动备份 / 安全输出 | 覆盖已有文件前自动生成 `文件名.bak-时间戳` 备份；写入走临时文件 + 原子替换，失败不会留下半成品 |
| PDF / 扫描件 | 图片页可直接选择 PDF，自动按页渲染后 OCR；文字 / 跨文件页支持有文字层的 PDF |

模板默认保存在用户配置目录（Windows：`%APPDATA%\TableFiller\templates`，
macOS：`~/Library/Application Support/TableFiller/templates`，
Linux：`~/.config/TableFiller/templates`），可用环境变量 `TABLE_FILLER_HOME` 覆盖。

## 对 Office 套件的适配

工具的写入引擎有两层：

1. **直接读写引擎（默认，始终可用）** —— 用 openpyxl / python-docx / python-pptx 直接读写 OOXML 文件。
   Office、WPS、OnlyOffice 都使用同一套 `.xlsx / .docx / .pptx` 格式，因此**三者通吃**，
   不需要 Office 正在运行。
2. **COM 引擎（可选，自动检测）** —— 当检测到已安装 Microsoft Excel 或 WPS 表格时，
   自动通过 Windows COM 驱动真实程序填写。适合需要保留复杂格式 / 公式 / 图表的场景。

> 建议：默认的“直接读写”引擎覆盖绝大多数需求；当文件里有图表、透视表等复杂对象、
> 或需要“开着 Excel 实时刷新”时，安装 Office / WPS 后会自动启用 COM 引擎。
> 工具会优先用 COM 引擎，失败时回退到直接读写。

## 图形化安装程序（推荐）

`installer/` 目录提供与常见 Windows 软件一致的**向导式图形安装程序**，
全程无命令框（黑色控制台窗口）：

- `installer/TableFiller-Setup.exe`：编译好的安装程序，双击即可（需先按下面的步骤生成）。
- `installer/Install.vbs`：未编译时的等价入口，双击同样以无窗口方式启动安装向导。
- `installer/build-setup.bat`：在 Windows 上把安装脚本编译成单文件 EXE。

安装向导包含「欢迎 → 选择安装位置 → 附加任务 → 正在安装 → 完成」五个页面：

| 功能 | 说明 |
| --- | --- |
| 自定义安装位置 | 默认 `%LOCALAPPDATA%\Programs\TableFiller`，可浏览选择其他目录；安装在当前用户下，无需管理员权限 |
| 附加任务 | 可勾选创建**桌面快捷方式**、**开始菜单快捷方式**、**固定到任务栏** |
| 自动下载程序本体 | 依次尝试 `ghfast.top` 镜像 → `gh-proxy.com` 镜像 → GitHub 官方，任一成功即继续 |
| 自动准备 Python | 若未检测到 Python 3.10–3.13，自动下载 Python 3.12.8：华为云镜像 → 阿里云镜像 → Python 官方，按当前用户静默安装 |
| 自动安装依赖 | 创建独立 `.venv`，pip 依次尝试清华 PyPI 镜像 → 阿里云 PyPI 镜像 → PyPI 官方 |
| 卸载 | 「设置 → 应用」中卸载，或运行安装目录下的 `uninstall.ps1`，会清理快捷方式、任务栏固定项与注册表信息 |

> 安装过程需要联网下载程序与依赖，首次安装依赖体积较大，请耐心等待进度条完成。

### 生成安装程序 EXE

在 Windows 上安装 PowerShell 5.1+ 后，双击或在命令行运行：

```bat
installer\build-setup.bat
```

脚本会自动安装 `ps2exe` 模块并生成 `installer\TableFiller-Setup.exe`
（无控制台、带图标、版本 1.1.0.0）。分发该 EXE 即可；若希望窗口标题也显示图标，
可将 `installer\TableFiller.ico` 与 EXE 放在同一目录。

## 安装

需要 Python 3.10+（推荐 3.11 / 3.12）。

```bat
cd office-table-filler
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

也可以直接双击 `run.bat`（首次会自动创建虚拟环境并安装依赖）。

OCR 说明：默认引擎是 **RapidOCR（中文 + 英文）**，安装时自动下载 ONNX 模型，无需额外配置。
如已安装 Tesseract，也可在界面“识别引擎”中切换。

## 使用方法

### ① 图片识别填表

1. 选择目标 Excel 文件（表头行用于列映射）。
2. 添加图片（可多选）。
3. 选择方式：
   - **表格网格识别**：图片中画有表格线时，自动检测行列并整表写入。
   - **字段提取**：为每个字段配置提取规则（标签 / 关键词 / 正则 / 行号），每张图片填一行。
4. 选择写入位置（如 `A1`）、是否追加、输出文件，点击“开始批量填表”。

### ② 文字填表

1. 选择目标 Excel，读取目标表头。
2. 添加来源文件（CSV / TXT / JSON / XLSX）或直接粘贴文本。
3. 建立“目标列 ↔ 源字段”映射，可点“按同名表头自动生成映射”。
4. 点击“开始填表”。

### ③ 跨文件填表

1. 添加源文件（可多个），载入预览。
2. 添加目标文件（可多个，Excel / Word / PPT）。
3. 添加映射行：`源字段 → 目标类型 → 位置参数`。
   - Excel 单元格：`Sheet1!B2`（可用 `步进=1` 让多条记录逐行下移）
   - Word 单元格：`table=0,row=1,col=2`
   - Word 占位符：`{{姓名}}`（正文里写 `{{姓名}}` 即可）
   - PPT 单元格：`slide=0,table=0,row=1,col=1`
4. 选择输出目录，勾选“按源数据逐行生成”，点击“开始跨文件填表”。

## 打包为 exe

```bat
pyinstaller --noconfirm --windowed --name TableFiller --add-data "app;app" main.py
```

或直接运行 `build.bat`。

> 打包后 OCR 模型需要一并带上。`build.bat` 中已包含
> `--add-data ".venv\Lib\site-packages\rapidocr_onnxruntime;rapidocr_onnxruntime"`，
> 并设置 `--collect-all rapidocr_onnxruntime`。

## 项目结构

```
office-table-filler/
├── main.py                 入口
├── make_samples.py         生成演示样例
├── requirements.txt
├── run.bat / build.bat
├── installer/
│   ├── TableFillerSetup.ps1  图形化安装向导（PowerShell + WinForms）
│   ├── Install.vbs           无命令框启动安装向导
│   ├── build-setup.bat       编译 TableFiller-Setup.exe
│   └── TableFiller.ico       安装程序图标
├── app/
│   ├── core/
│   │   ├── spreadsheet.py   表格引擎（openpyxl / Excel COM）
│   │   ├── data_io.py       读取 xlsx / csv / txt / json / docx / pptx / pdf
│   │   ├── ocr.py           OCR 引擎抽象（RapidOCR / Tesseract / Windows OCR）
│   │   ├── table_detect.py  OpenCV 表格线检测与词聚类
│   │   ├── mapping.py       字段提取与列匹配
│   │   ├── targets.py       跨文件目标位置解析
│   │   ├── filler.py        三种填表方式的编排
│   │   ├── validation.py    执行前校验
│   │   ├── preview.py       执行前预览
│   │   ├── templates.py     模板保存 / 复用
│   │   ├── safeio.py        原子写入与自动备份
│   │   ├── pdf_io.py        PDF 文本提取与逐页渲染
│   │   ├── models.py        数据结构
│   │   └── utils.py         文本 / 数字 / 日期工具
│   └── ui/
│       ├── main_window.py
│       ├── image_tab.py     图片识别填表
│       ├── text_tab.py      文字填表
│       ├── cross_tab.py     跨文件填表
│       ├── template_bar.py  模板工具栏
│       ├── preview_dialog.py 预览 / 校验对话框
│       ├── mapping_widgets.py
│       └── common.py
└── tests/                  单元测试
```

## 运行测试

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python make_samples.py
python -m unittest discover -s tests -v
```

## 已知限制

- 图片是**照片 / 有透视变形**时，表格线检测会退化为按词聚类，建议先截成规整图片。
- 直接读写引擎会保留单元格值与样式，但复杂对象（图表、透视表）请改用 COM 引擎。
- Word / PPT 的表格行追加依赖 `python-docx` / `python-pptx` 能力，复杂模板建议在目标程序中检查。
