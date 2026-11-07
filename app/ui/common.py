"""UI helpers shared by the tabs."""

from __future__ import annotations

import os
import traceback

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (QFileDialog, QMessageBox, QWidget,
                               QApplication)


def pick_files(parent: QWidget, title: str, extensions: list[str],
               multi: bool = True) -> list[str]:
    filters = ";;".join(extensions)
    if multi:
        files, _ = QFileDialog.getOpenFileNames(parent, title, "", filters)
        return list(files)
    path, _ = QFileDialog.getOpenFileName(parent, title, "", filters)
    return [path] if path else []


def pick_file(parent: QWidget, title: str, extensions: list[str]) -> str:
    return pick_files(parent, title, extensions, multi=False)[0]


def pick_dir(parent: QWidget, title: str = "选择目录") -> str:
    return QFileDialog.getExistingDirectory(parent, title)


def save_file(parent: QWidget, title: str, default_name: str = "output.xlsx",
              extensions: list[str] | None = None) -> str:
    extensions = extensions or ["Excel 表格 (*.xlsx *.xlsm)", "所有文件 (*.*)"]
    path, _ = QFileDialog.getSaveFileName(parent, title, default_name, ";;".join(extensions))
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
