"""Main window: the three feature tabs plus a status bar."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QIcon
from PySide6.QtWidgets import (QLabel, QMainWindow, QMessageBox, QTabWidget,
                               QVBoxLayout, QWidget)

from .. import __version__
from .common import confirm
from .cross_tab import CrossTab
from .image_tab import ImageTab
from .text_tab import TextTab


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"表格批量填写工具 v{__version__}")
        self.resize(1180, 780)
        self._build()

    def _build(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        self.tabs = QTabWidget()
        self.tabs.addTab(ImageTab(), "① 图片识别填表")
        self.tabs.addTab(TextTab(), "② 文字填表")
        self.tabs.addTab(CrossTab(), "③ 跨文件填表")
        root.addWidget(self.tabs, 1)
        self.setCentralWidget(central)
        self.statusBar().showMessage("就绪。")

        self._build_menu()
        self._build_toolbar()

    def _build_menu(self) -> None:
        file_menu = self.menuBar().addMenu("文件(&F)")
        open_action = QAction("打开目标表格…", self)
        open_action.triggered.connect(self._open_target)
        file_menu.addAction(open_action)
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        help_menu = self.menuBar().addMenu("帮助(&H)")
        about_action = QAction("关于", self)
        about_action.triggered.connect(self._about)
        help_menu.addAction(about_action)

    def _build_toolbar(self) -> None:
        from PySide6.QtWidgets import QToolBar
        toolbar = QToolBar("主工具栏")
        toolbar.setMovable(False)
        open_action = QAction("打开目标表格", self)
        open_action.triggered.connect(self._open_target)
        toolbar.addAction(open_action)
        self.addToolBar(toolbar)

    def _open_target(self) -> None:
        tab = self.tabs.currentWidget()
        if isinstance(tab, ImageTab):
            tab._pick_target()
        elif isinstance(tab, TextTab):
            tab._pick_target()

    def _about(self) -> None:
        QMessageBox.about(
            self, "关于",
            f"<b>表格批量填写工具</b> v{__version__}<br><br>"
            "支持三种填表方式：<br>"
            "① 图片识别批量填表（OCR + 表格线检测）<br>"
            "② 文字 / CSV / JSON 填表<br>"
            "③ 跨文件填表（源文件 → 多个目标文件）<br><br>"
            "适配 Microsoft Office（Excel / Word / PowerPoint）、WPS、OnlyOffice。",
        )

    def _has_running_workers(self) -> bool:
        for index in range(self.tabs.count()):
            worker = getattr(self.tabs.widget(index), "worker", None)
            if worker is not None and worker.isRunning():
                return True
        return False

    def closeEvent(self, event) -> None:  # noqa: N802 - Qt naming
        if self._has_running_workers() and not confirm(
            self, "确认退出", "有任务正在运行，确定要退出吗？"
        ):
            event.ignore()
            return
        for index in range(self.tabs.count()):
            widget = self.tabs.widget(index)
            closer = getattr(widget, "on_before_close", None)
            if callable(closer):
                closer()
        event.accept()
