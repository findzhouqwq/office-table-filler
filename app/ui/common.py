"""UI helpers shared by the tabs."""

from __future__ import annotations

import os
import traceback

from PySide6.QtCore import QSettings, QThread, Signal
from PySide6.QtWidgets import (QFileDialog, QMessageBox, QWidget,
                               QApplication)


def app_settings() -> QSettings:
    """Application-wide settings (last folders, window geometry)."""
    return QSettings("TableFiller", "TableFiller")


def _last_dir(key: str) -> str:
    value = app_settings().value(f"paths/{key}", "")
    return value if isinstance(value, str) and os.path.isdir(value) else ""


def _remember_dir(path: str, key: str) -> None:
    if path and os.path.isdir(path):
        app_settings().setValue(f"paths/{key}", path)


def pick_files(parent: QWidget, title: str, extensions: list[str],
               multi: bool = True, key: str = "files") -> list[str]:
    filters = ";;".join(extensions)
    start = _last_dir(key)
    if multi:
        files, _ = QFileDialog.getOpenFileNames(parent, title, start, filters)
        result = list(files)
    else:
        path, _ = QFileDialog.getOpenFileName(parent, title, start, filters)
        result = [path] if path else []
    if result:
        _remember_dir(os.path.dirname(result[0]), key)
    return result


def pick_file(parent: QWidget, title: str, extensions: list[str],
              key: str = "files") -> str:
    files = pick_files(parent, title, extensions, multi=False, key=key)
    return files[0] if files else ""


def pick_dir(parent: QWidget, title: str = "选择目录", key: str = "dir") -> str:
    path = QFileDialog.getExistingDirectory(parent, title, _last_dir(key))
    if path:
        _remember_dir(path, key)
    return path


def save_file(parent: QWidget, title: str, default_name: str = "output.xlsx",
              extensions: list[str] | None = None, key: str = "save") -> str:
    extensions = extensions or ["Excel 表格 (*.xlsx *.xlsm)", "所有文件 (*.*)"]
    start = os.path.join(_last_dir(key), default_name) if _last_dir(key) else default_name
    path, _ = QFileDialog.getSaveFileName(parent, title, start, ";;".join(extensions))
    if path:
        _remember_dir(os.path.dirname(path), key)
    return path


def show_error(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.critical(parent, title, message)


def show_info(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.information(parent, title, message)


def show_warning(parent: QWidget, title: str, message: str) -> None:
    QMessageBox.warning(parent, title, message)


def confirm(parent: QWidget, title: str, message: str) -> bool:
    return QMessageBox.question(parent, title, message,
                                QMessageBox.Yes | QMessageBox.No,
                                QMessageBox.No) == QMessageBox.Yes


class Worker(QThread):
    """Run a callable off the UI thread and stream its results back."""

    finished_ok = Signal(object)
    failed = Signal(str)
    progress = Signal(str)

    def __init__(self, func, *args, parent: QWidget | None = None):
        super().__init__(parent)
        self._func = func
        self._args = args
        self._progress = None
        try:
            import inspect

            params = inspect.signature(func).parameters
            self._progress = "progress" in params or any(
                p.kind == inspect.Parameter.VAR_KEYWORD for p in params.values()
            )
        except (TypeError, ValueError):
            self._progress = False

    def run(self) -> None:  # noqa: D401
        try:
            if self._progress:
                result = self._func(*self._args, progress=self.progress.emit)
            else:
                result = self._func(*self._args)
        except Exception as exc:  # noqa: BLE001 - surface to the UI
            self.failed.emit(f"{exc}\n\n{traceback.format_exc()}")
            return
        self.finished_ok.emit(result)


def format_report(report) -> str:
    lines = [f"【{report.title}】完成：写入 {report.written} 处"
             f"{'，失败 ' + str(report.failed) if report.failed else ''}"]
    if report.output_path:
        lines.append(f"输出：{report.output_path}")
    for warning in report.warnings:
        lines.append(f"  · {warning}")
    return "\n".join(lines)
