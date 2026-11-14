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

## 界面与主题

- 统一主题层 `app/ui/theme.py`：集中管理配色与 Qt 样式表，主色靛蓝、浅灰背景、白色圆角卡片。
- 三个功能页均为「卡片 + 标签页」布局：配置项收进卡片，预览 / 规则 / 映射用标签页切换，
  表格获得更多纵向空间；主操作按钮统一为靛蓝实心主按钮。
- 执行报告为可折叠卡片，默认收起，运行结束自动展开。
- 应用图标在运行时绘制，无需外部图片资源；窗口默认 1360×800，适配 1366×768 及以上屏幕。

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
| 自动下载程序本体 | 从 GitHub Release 获取 `app-<版本>.zip` + `manifest-<版本>.json` + 签名，镜像依次 `ghfast.top` → `gh-proxy.com` → GitHub 官方 |
| 完整性校验 | 清单用内置 RSA 公钥验签，`app-<版本>.zip` 再校验 SHA256 与大小；任一不符自动切换下一源 |
| 自动准备 Python | 若未检测到 Python 3.10–3.13，自动下载 Python 3.12.8：华为云 → 阿里云 → Python 官方，并校验 Authenticode 签名 |
| VC++ 运行库 | 自动检测 Microsoft Visual C++ 2015–2022 x64 运行库，缺失时静默安装微软官方 `vc_redist.x64.exe`（校验微软签名；该组件可能触发一次 UAC 授权） |
| 自动安装依赖 | 创建独立 `.venv`，pip 依次尝试清华 PyPI → 阿里云 PyPI → PyPI 官方 |
| 卸载 | 「设置 → 应用」中卸载，或运行安装目录下的 `uninstall.ps1`；会清理快捷方式、任务栏固定项与注册表 |
| 卸载时的模板 | 卸载确认框可选择是否删除 `%APPDATA%\TableFiller`（用户模板与设置），默认选择“否”即保留 |
| 升级 | 覆盖安装前自动把旧版程序文件备份到 `%LOCALAPPDATA%\Programs\TableFiller-backup-<时间戳>`，用户模板不受影响 |

> 安装过程需要联网下载程序与依赖，首次安装依赖体积较大，请耐心等待进度条完成。

### 安全与完整性

镜像只负责搬运字节，可信性由签名兜底：

1. 安装程序内置发布公钥，只有用对应私钥签名的 `manifest-<版本>.json` 才会被接受；
2. 清单通过后，再校验 `app-<版本>.zip` 的 **SHA256 + 文件大小**；
3. 校验失败会自动切换到下一个镜像；全部失败则中止安装，不会运行未经验证的程序；
4. Python 安装包与 VC++ 运行库额外做 **Authenticode** 校验（分别要求 Python Software Foundation / Microsoft 签名）。

发布签名文件由 `installer/sign-release.ps1` 生成（需要私钥，私钥不进入仓库）：

```powershell
$env:TABLEFILLER_SIGNING_KEY = 'D:\keys\tablefiller-signing.xml'
pwsh -File installer\sign-release.ps1 -Version 1.2.1
```

生成 `dist\app-1.2.1.zip`、`dist\manifest-1.2.1.json`、`dist\manifest-1.2.1.sig` 后，作为 GitHub Release 资产上传；
安装程序从 `releases/latest/download/` 获取并校验。资产名带版本号，
可避免 CDN / 镜像缓存把旧版本的清单或安装包混进新版本。

### 代码签名与 SmartScreen

未签名的 `TableFiller-Setup.exe` 和 PyInstaller 产物会触发
“Windows 已保护你的电脑”以及部分杀软拦截。建议用代码签名证书签名：

```bat
set SIGN_PFX=D:\certs\my.pfx
set SIGN_PASSWORD=******
installer\build-setup.bat
```

`build-setup.bat` 检测到 `SIGN_PFX` 或 `SIGN_THUMBPRINT` 时会自动调用
`installer\sign-setup.bat`，使用 `signtool` 做 SHA256 签名并打时间戳。
签名后仍需一定下载量才能建立 SmartScreen 信誉；EV 证书可立即获得信誉。

### 生成安装程序 EXE

在 Windows 上安装 PowerShell 5.1+ 后，双击或在命令行运行：

```bat
installer\build-setup.bat
```

脚本会自动安装 `ps2exe` 模块并生成 `installer\TableFiller-Setup.exe`
（无控制台、带图标、版本 1.2.1.0）。分发该 EXE 即可；若希望窗口标题也显示图标，
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
│   ├── sign-setup.bat        signtool 代码签名
│   ├── sign-release.ps1      生成签名发布清单（app.zip / manifest）
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
│       ├── theme.py         主题：配色 / 样式表 / 卡片组件
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
