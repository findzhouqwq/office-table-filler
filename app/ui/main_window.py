"""Main window: a modern header, the three feature tabs and a status bar."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (QFrame, QHBoxLayout, QLabel, QMainWindow,
                               QMessageBox, QSizePolicy, QTabWidget,
                               QVBoxLayout, QWidget)

from .. import __version__
from .common import app_settings, confirm
from .cross_tab import CrossTab
from .image_tab import ImageTab
from .text_tab import TextTab
from .theme import ACCENT, make_app_icon, ghost_button


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"表格批量填写工具 v{__version__}")
        self.setWindowIcon(make_app_icon())
        self.resize(1360, 800)
        self.setMinimumSize(1100, 680)
        self._build()
        self._restore_state()

    def _restore_state(self) -> None:
        geometry = app_settings().value("window/geometry")
        if geometry:
            self.restoreGeometry(geometry)
        state = app_settings().value("window/state")
        if state:
            self.restoreState(state)

    def _build(self) -> None:
        central = QWidget()
        central.setObjectName("AppRoot")
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_header())

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(22, 14, 22, 12)
        content_layout.setSpacing(10)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.addTab(ImageTab(), "① 图片识别填表")
        self.tabs.addTab(TextTab(), "② 文字填表")
        self.tabs.addTab(CrossTab(), "③ 跨文件填表")
        content_layout.addWidget(self.tabs, 1)
        root.addWidget(content, 1)

        self.setCentralWidget(central)

        status = self.statusBar()
        status.showMessage("就绪。")
        status.setSizeGripEnabled(False)

        self._build_menu()

    def _build_header(self) -> QWidget:
        header = QFrame()
        header.setObjectName("Header")
        header.setFixedHeight(68)
        layout = QHBoxLayout(header)
        layout.setContentsMargins(22, 12, 22, 12)
        layout.setSpacing(12)

        logo = QLabel()
        logo.setPixmap(make_app_icon(72).pixmap(38, 38))
        logo.setFixedSize(38, 38)
        layout.addWidget(logo)

        titles = QVBoxLayout()
        titles.setContentsMargins(0, 0, 0, 0)
        titles.setSpacing(1)
        title = QLabel("表格批量填写工具")
        title.setObjectName("AppTitle")
        subtitle = QLabel("Office · WPS · OnlyOffice 批量填表")
        subtitle.setObjectName("AppSubtitle")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        layout.addLayout(titles)

        layout.addStretch(1)

        version = QLabel(f"v{__version__}")
        version.setObjectName("Pill")
        layout.addWidget(version)

        open_button = ghost_button("打开目标表格")
        open_button.clicked.connect(self._open_target)
        layout.addWidget(open_button)
        return header

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
        app_settings().setValue("window/geometry", self.saveGeometry())
        app_settings().setValue("window/state", self.saveState())
        event.accept()