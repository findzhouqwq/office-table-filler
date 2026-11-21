# AGENTS.md

面向在本仓库工作的编码助手。先读本文件，再读 `HANDOVER.md` 的「架构」与「关键设计约定」。

## 1. 项目概览

- Windows 桌面工具：PySide6 GUI + Python 3.10+，把图片 / 文本 / 多个源文件的数据批量填入 Excel、Word、PPT。
- 三种填表方式：图片 OCR 填表、文字填表、跨文件填表；另有右侧「填表助手」面板（含跨文件绿色框选复制）。
- 表格写入有两层引擎：openpyxl / python-docx / python-pptx 直接读写（默认），以及 Windows COM（Excel / WPS）。
- 目标格式为 OOXML（`.xlsx/.docx/.pptx`），因此 Office、WPS、OnlyOffice 均可读写。

## 2. 常用命令

```bash
# 安装依赖
pip install -r requirements.txt

# 启动 GUI / 环境自检
python main.py
python main.py --check

# 跑测试（Windows）
python -m unittest discover -s tests -v

# 跑测试（Linux / CI，无显示器）
QT_QPA_PLATFORM=offscreen python -m unittest discover -s tests
```

Linux 上跑 GUI 测试需要先安装：

```bash
pip install PySide6
apt-get install -y libegl1 libgl1 libxkbcommon0 libdbus-1-3 libfontconfig1 libxcb-cursor0
```

- 打包：`build.bat`；安装程序相关在 `installer/`。
- 演示：`python make_samples.py && python demo.py`，输出到 `demo_output/`。

## 3. 目录与分层

```
main.py                 入口（--check / GUI）
app/core/               引擎层，禁止依赖 app/ui
app/ui/                 PySide6 界面
tests/                  unittest 用例
installer/              安装与发布脚本
```

分层规则：

- `app/core` 可以单独在无 GUI 环境调用；**不要**在 core 里 import UI。
- UI 只通过 core 的公开函数/类交互；耗时任务走 `app/ui/common.py` 的 `Worker`。
- core 的用户可见错误统一用 `app/core/errors.py` 的 `AppError` 子类；UI 用 `show_error()` 展示。
- 引擎的公开能力包括 `read_block()`、`write_block()`、`write_rows()`、`worksheet()`；
  不要跨模块调用 `engine._ws()` 等私有成员，也不要导入其他模块的 `_private` 函数。
- 预览逻辑（`preview.py`）只使用 filler 的公开规划函数；新增共享逻辑时先公开命名，再复用。

## 4. 关键约定与已知坑

- **OCR 坐标一致性**：词框与表格线检测必须使用同一张预处理图（见 `OcrEngine.recognize_with_image()`）。
- **COM 仅 Windows**：`screen_copy.py` / `ExcelComEngine` 依赖 `pywin32`；调用 `CoInitialize()` 后必须在
  `CoUninitialize()` 前释放所有 COM 代理。非 Windows 上必须优雅报错，不能 import 崩溃。
- **屏幕框选**：需要正在运行的可见 Excel / WPS；WPS 不支持 `RangeFromPoint` 时提示改用 Excel。
- **安全输出**：所有写文件走 `safeio.save_with_backup()`（临时文件 + 原子替换 + 时间戳备份）。
  受保护目录（受控文件夹访问）依赖 `probe_cache()`；探测必须用内建 `open()`，**不要用 `tempfile.mkstemp()`**。
- **模板目录**：默认用户配置目录，可用环境变量 `TABLE_FILLER_HOME` 覆盖（测试依赖此变量）。
- **Excel 限制**：1,048,576 行 × 16,384 列；单元格文本上限 32,767 字符；以 `=` 开头的值按文本写入。
- **低端机性能**：
  - 自绘控件（框选覆盖层、侧边栏）只重绘脏区，缓存 `QColor/QPen/QFont`；
  - 大面积读写用 `read_block()` / `write_block()` / `write_rows()`，不要逐单元格解析坐标；
  - 批量填充 `QTableWidget` 时先 `setUpdatesEnabled(False)`，结束再恢复；
  - 预览只保留前 200 行；复制缓冲有块数/单元格上限，避免占满内存。
- **线程**：不要从工作线程直接操作 Qt 控件；结果通过信号回到主线程。

## 5. 测试

- 使用标准库 `unittest`，入口 `python -m unittest discover -s tests`。
- 当前基线：**122 个通过、5 个跳过**（OCR / Windows 相关用例在缺依赖时跳过）。
- 新增 core 行为时补单元测试；GUI 交互用 `QT_QPA_PLATFORM=offscreen` 验证。
- 不要为了通过测试而删除跳过逻辑或放宽断言。

## 6. 版本与发布

- 每个版本在 `CHANGELOG.md` 顶部**单独新增**小节；不要把新功能追加进旧版本小节。
- 升版本时同步这些位置：
  - `app/__init__.py` 的 `__version__`
  - `installer/TableFillerSetup.ps1` 的 `$script:Version`
  - `installer/build-setup.bat` 的 `-version`
  - `installer/sign-release.ps1` 的默认版本与示例
  - `README.md` / `HANDOVER.md` 中的发布示例
- 发布资产由 `installer/sign-release.ps1` 生成，私钥不进入仓库。

## 7. 协作边界

- 除非用户明确要求，不要执行 `git commit` / `git push` / 打标签或覆盖已有标签。
- 不要在仓库、脚本或文档中写入 GitHub PAT、密码等凭据；用户粘贴的令牌只用于当次明确授权的操作。
- 修改保持最小范围：不顺手重构、不新增无关抽象、不生成多余文档。
- 代码标识符用英文，界面文案用中文；保持现有风格与类型标注习惯。
