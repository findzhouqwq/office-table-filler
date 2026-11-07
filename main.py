"""表格批量填写工具 — 主入口。

用法::

    python main.py            # 启动图形界面
    python main.py --check    # 仅检查环境依赖
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import traceback


def setup_logging() -> str:
    log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "table_filler.log")
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    try:
        handlers.insert(0, logging.FileHandler(log_path, encoding="utf-8"))
    except OSError:
        # Installed under Program Files / a read-only folder: fall back to temp.
        import tempfile

        log_path = os.path.join(tempfile.gettempdir(), "table_filler.log")
        try:
            handlers.insert(0, logging.FileHandler(log_path, encoding="utf-8"))
        except OSError:
            pass
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=handlers,
    )
    return log_path


def check_environment() -> int:
    """Print which optional components are available."""
    from app.core import ocr as ocr_module
    from app.core.spreadsheet import available_engines
    print("== 表格批量填写工具 环境检查 ==")
    engines = available_engines()
    print(f"表格引擎：直接读写={'可用' if engines['openpyxl'] else '缺失'}，"
          f"COM(Excel/WPS)={'可用' if engines['com'] else '缺失（将使用直接读写引擎）'}")
    ocr = ocr_module.available_engines()
    print(f"OCR 引擎：{', '.join(ocr) if ocr else '无（请安装 rapidocr-onnxruntime）'}")
    for module in ("openpyxl", "docx", "pptx", "cv2", "PIL", "pandas", "win32com"):
        try:
            __import__(module)
            print(f"  · {module}: OK")
        except Exception as exc:
            print(f"  · {module}: 缺失（{exc.__class__.__name__}）")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="表格批量填写工具")
    parser.add_argument("--check", action="store_true", help="只检查环境依赖")
    args = parser.parse_args()

    setup_logging()
    if args.check:
        return check_environment()

    try:
        from PySide6.QtWidgets import QApplication
        from PySide6.QtCore import Qt
        from PySide6.QtGui import QFont
    except ImportError:
        print("未安装 PySide6，请先执行：pip install -r requirements.txt", file=sys.stderr)
        return 2

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    app = QApplication(sys.argv)
    app.setApplicationName("表格批量填写工具")
    app.setFont(QFont("Microsoft YaHei UI", 10))

    from app.ui.main_window import MainWindow
    window = MainWindow()
    window.show()
    logging.info("GUI started")
    return app.exec()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:  # noqa: BLE001
        traceback.print_exc()
        raise SystemExit(1)
