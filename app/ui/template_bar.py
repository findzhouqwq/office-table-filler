"""Reusable toolbar for saving and loading fill templates."""

from __future__ import annotations

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (QComboBox, QFileDialog, QHBoxLayout,
                               QInputDialog, QLabel, QMessageBox, QPushButton,
                               QWidget)

from ..core import templates
from ..core.errors import AppError
from .common import save_file, show_error, show_info


class TemplateBar(QWidget):
    """A combo box plus save/load/delete/import/export buttons.

    The owning tab receives :attr:`save_requested` and :attr:`load_requested`
    and is responsible for reading/writing its own settings.
    """

    save_requested = Signal(str)
    load_requested = Signal(str)

    def __init__(self, kind: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.kind = kind
        self.setObjectName("TemplateBar")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 6, 14, 6)
        layout.setSpacing(8)

        title = QLabel("模板")
        title.setObjectName("CardTitle")
        layout.addWidget(title)

        self.combo = QComboBox()
        self.combo.setMinimumWidth(220)
        layout.addWidget(self.combo, 1)

        layout.addWidget(QPushButton("保存…", self, clicked=self._save))
        apply_btn = QPushButton("应用", self)
        apply_btn.setProperty("role", "ghost")
        apply_btn.clicked.connect(self._load)
        layout.addWidget(apply_btn)
        delete_btn = QPushButton("删除", self)
        delete_btn.setProperty("role", "danger")
        delete_btn.clicked.connect(self._delete)
        layout.addWidget(delete_btn)
        layout.addWidget(QPushButton("导入…", self, clicked=self._import))
        layout.addWidget(QPushButton("导出…", self, clicked=self._export))

        self.setToolTip(f"模板目录：{templates.templates_root()}")
        self.refresh()

    # -- helpers ------------------------------------------------------------
    def refresh(self) -> None:
        current = self.combo.currentText()
        self.combo.blockSignals(True)
        self.combo.clear()
        for template in templates.list_templates(self.kind):
            self.combo.addItem(template.name)
        self.combo.blockSignals(False)
        if current:
            self.select(current)

    def current_name(self) -> str:
        return self.combo.currentText().strip()

    def select(self, name: str) -> None:
        index = self.combo.findText(name)
        if index >= 0:
            self.combo.setCurrentIndex(index)

    # -- actions ------------------------------------------------------------
    def _save(self) -> None:
        default = self.current_name() or "新模板"
        name, ok = QInputDialog.getText(self, "保存模板", "模板名称：", text=default)
        if not ok or not name.strip():
            return
        try:
            templates.validate_name(name)
        except AppError as exc:
            show_error(self, "模板名称无效", str(exc))
            return
        self.save_requested.emit(name.strip())

    def _load(self) -> None:
        name = self.current_name()
        if not name:
            QMessageBox.information(self, "提示", "还没有可用的模板。")
            return
        self.load_requested.emit(name)

    def _delete(self) -> None:
        name = self.current_name()
        if not name:
            return
        answer = QMessageBox.question(
            self, "删除模板", f"确定删除模板 “{name}” 吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if answer != QMessageBox.Yes:
            return
        try:
            templates.delete_template(name, self.kind)
        except AppError as exc:
            show_error(self, "删除失败", str(exc))
            return
        self.refresh()

    def _import(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "导入模板", "", "模板文件 (*.json);;所有文件 (*.*)")
        if not path:
            return
        try:
            template = templates.import_template(path, self.kind)
        except AppError as exc:
            show_error(self, "导入失败", str(exc))
            return
        self.refresh()
        self.select(template.name)
        show_info(self, "导入成功", f"已导入模板：{template.name}")

    def _export(self) -> None:
        name = self.current_name()
        if not name:
            QMessageBox.information(self, "提示", "请先选择一个模板。")
            return
        path = save_file(self, "导出模板", f"{name}.json",
                         ["模板文件 (*.json)", "所有文件 (*.*)"], key="template")
        if not path:
            return
        try:
            saved = templates.export_template(name, self.kind, path)
        except AppError as exc:
            show_error(self, "导出失败", str(exc))
            return
        show_info(self, "导出成功", f"模板已导出到：\n{saved}")
