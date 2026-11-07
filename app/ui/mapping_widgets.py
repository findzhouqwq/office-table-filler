"""Editable tables for extraction rules and column plans."""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QHeaderView,
                               QTableWidget, QTableWidgetItem, QWidget)

from ..core.models import ExtractionRule, FieldMapping


class _BaseTable(QTableWidget):
    changed = Signal()

    def _emit(self) -> None:
        self.changed.emit()


class RuleTable(_BaseTable):
    """Edit :class:`ExtractionRule` objects."""

    HEADERS = ["字段名", "提取方式", "规则内容", "默认值", "转换", "必填"]
    TYPES = ["label", "keyword", "regex", "line", "word", "grid"]
    TRANSFORMS = ["text", "raw", "upper", "lower", "title", "number", "strip"]

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(0, len(self.HEADERS), parent)
        self.setHorizontalHeaderLabels(self.HEADERS)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        header = self.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.verticalHeader().setVisible(False)
        self._emit()

    def load(self, rules: list[ExtractionRule]) -> None:
        self.setRowCount(len(rules))
        for row, rule in enumerate(rules):
            self._text(row, 0, rule.field)
            combo = QComboBox()
            combo.addItems(self.TYPES)
            combo.setCurrentText(rule.pattern_type)
            combo.currentTextChanged.connect(lambda _: self._emit())
            self.setCellWidget(row, 1, combo)
            self._text(row, 2, rule.pattern)
            self._text(row, 3, rule.default)
            tf = QComboBox()
            tf.addItems(self.TRANSFORMS)
            tf.setCurrentText(rule.transform if rule.transform in self.TRANSFORMS else "text")
            tf.currentTextChanged.connect(lambda _: self._emit())
            self.setCellWidget(row, 4, tf)
            self._text(row, 5, "是" if rule.required else "")
        self._emit()

    def to_rules(self) -> list[ExtractionRule]:
        rules: list[ExtractionRule] = []
        for row in range(self.rowCount()):
            field = self._value(row, 0).strip()
            if not field:
                continue
            combo = self.cellWidget(row, 1)
            tf = self.cellWidget(row, 4)
            rules.append(ExtractionRule(
                field=field,
                pattern_type=combo.currentText() if combo else "label",
                pattern=self._value(row, 2),
                default=self._value(row, 3),
                transform=tf.currentText() if tf else "text",
                required=self._value(row, 5) == "是",
            ))
        return rules

    def _text(self, row: int, col: int, value: str) -> None:
        item = QTableWidgetItem(value)
        self.setItem(row, col, item)

    def _value(self, row: int, col: int) -> str:
        item = self.item(row, col)
        return item.text() if item else ""

    def add_row(self) -> None:
        self.insertRow(self.rowCount())
        combo = QComboBox()
        combo.addItems(self.TYPES)
        self.setCellWidget(self.rowCount() - 1, 1, combo)
        tf = QComboBox()
        tf.addItems(self.TRANSFORMS)
        self.setCellWidget(self.rowCount() - 1, 4, tf)
        self._emit()

    def remove_row(self, row: int) -> None:
        if 0 <= row < self.rowCount():
            self.removeRow(row)
            self._emit()


class PlanTable(_BaseTable):
    """Edit :class:`FieldMapping` objects."""

    HEADERS = ["目标列 / 表头", "源字段", "转换", "默认值"]
    TRANSFORMS = ["text", "raw", "upper", "lower", "title", "number", "strip"]

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(0, len(self.HEADERS), parent)
        self.setHorizontalHeaderLabels(self.HEADERS)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        header = self.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.verticalHeader().setVisible(False)
        self._emit()

    def load(self, plan: list[FieldMapping]) -> None:
        self.setRowCount(len(plan))
        for row, mapping in enumerate(plan):
            self._text(row, 0, mapping.target)
            self._text(row, 1, mapping.source)
            tf = QComboBox()
            tf.addItems(self.TRANSFORMS)
            tf.setCurrentText(mapping.transform if mapping.transform in self.TRANSFORMS else "text")
            tf.currentTextChanged.connect(lambda _: self._emit())
            self.setCellWidget(row, 2, tf)
            self._text(row, 3, mapping.default)
        self._emit()

    def to_plan(self) -> list[FieldMapping]:
        plan: list[FieldMapping] = []
        for row in range(self.rowCount()):
            target = self._value(row, 0).strip()
            source = self._value(row, 1).strip()
            if not source:
                continue
            tf = self.cellWidget(row, 2)
            plan.append(FieldMapping(
                target=target or source,
                source=source,
                transform=tf.currentText() if tf else "text",
                default=self._value(row, 3),
            ))
        return plan

    def _text(self, row: int, col: int, value: str) -> None:
        self.setItem(row, col, QTableWidgetItem(value))

    def _value(self, row: int, col: int) -> str:
        item = self.item(row, col)
        return item.text() if item else ""

    def add_row(self) -> None:
        self.insertRow(self.rowCount())
        tf = QComboBox()
        tf.addItems(self.TRANSFORMS)
        self.setCellWidget(self.rowCount() - 1, 2, tf)
        self._emit()

    def remove_row(self, row: int) -> None:
        if 0 <= row < self.rowCount():
            self.removeRow(row)
            self._emit()


class TargetTable(_BaseTable):
    """Edit cross-file mappings (source field -> target reference)."""

    HEADERS = ["源字段", "目标类型", "目标文件", "位置参数", "实际位置", "转换", "默认值"]
    KINDS = ["xlsx_cell", "xlsx_rows", "docx_table_cell", "docx_table_rows",
             "pptx_table_cell", "docx_placeholder", "pptx_placeholder"]
    TRANSFORMS = ["text", "raw", "upper", "lower", "title", "number", "strip"]
    SPEC_COLUMN = 3
    DESCRIBE_COLUMN = 4

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(0, len(self.HEADERS), parent)
        self.setHorizontalHeaderLabels(self.HEADERS)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
        header = self.horizontalHeader()
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.verticalHeader().setVisible(False)
        self.itemChanged.connect(self._on_changed)
        self._emit()

    def load(self, entries) -> None:
        self.blockSignals(True)
        self.setRowCount(len(entries))
        for row, entry in enumerate(entries):
            self._text(row, 0, entry.source)
            combo = QComboBox()
            combo.addItems(self.KINDS)
            combo.setCurrentText(entry.target.kind)
            combo.currentTextChanged.connect(lambda _: self._emit())
            self.setCellWidget(row, 1, combo)
            self._text(row, 2, entry.target.file)
            spec = _spec_for(entry.target)
            self._text(row, self.SPEC_COLUMN, spec)
            self._refresh_describe(row)
            tf = QComboBox()
            tf.addItems(self.TRANSFORMS)
            tf.setCurrentText(entry.transform if entry.transform in self.TRANSFORMS else "text")
            tf.currentTextChanged.connect(lambda _: self._emit())
            self.setCellWidget(row, 5, tf)
            self._text(row, 6, entry.default)
        self.blockSignals(False)
        self._emit()

    def to_entries(self, default_file: str = "") -> list:
        from ..core.models import MappingEntry
        from ..core.targets import parse_target_spec
        entries = []
        for row in range(self.rowCount()):
            source = self._value(row, 0).strip()
            if not source:
                continue
            combo = self.cellWidget(row, 1)
            kind = combo.currentText() if combo else "xlsx_cell"
            spec = self._value(row, self.SPEC_COLUMN)
            tf = self.cellWidget(row, 5)
            file_name = self._value(row, 2).strip()  # empty = any compatible target
            entries.append(MappingEntry(
                source=source,
                target=parse_target_spec(kind, spec, file_name),
                transform=tf.currentText() if tf else "text",
                default=self._value(row, 6),
            ))
        return entries

    def _refresh_describe(self, row: int) -> None:
        from ..core.targets import describe_target
        combo = self.cellWidget(row, 1)
        kind = combo.currentText() if combo else "xlsx_cell"
        spec = self._value(row, self.SPEC_COLUMN)
        desc = describe_target(kind, spec)
        item = self.item(row, self.DESCRIBE_COLUMN)
        if item:
            item.setText(desc)

    def _on_changed(self, item: QTableWidgetItem) -> None:
        if item.column() == self.SPEC_COLUMN:
            self._refresh_describe(item.row())
        self._emit()

    def _text(self, row: int, col: int, value: str) -> None:
        item = QTableWidgetItem(value)
        if col == self.DESCRIBE_COLUMN:
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        self.setItem(row, col, item)

    def _value(self, row: int, col: int) -> str:
        item = self.item(row, col)
        return item.text() if item else ""

    def add_row(self) -> None:
        self.insertRow(self.rowCount())
        combo = QComboBox()
        combo.addItems(self.KINDS)
        combo.currentTextChanged.connect(lambda _: self._emit())
        self.setCellWidget(self.rowCount() - 1, 1, combo)
        tf = QComboBox()
        tf.addItems(self.TRANSFORMS)
        tf.currentTextChanged.connect(lambda _: self._emit())
        self.setCellWidget(self.rowCount() - 1, 5, tf)
        self._refresh_describe(self.rowCount() - 1)
        self._emit()

    def remove_row(self, row: int) -> None:
        if 0 <= row < self.rowCount():
            self.removeRow(row)
            self._emit()


def _spec_for(ref) -> str:
    """Render a TargetRef back into the compact spec text."""
    kind = ref.kind
    if kind == "xlsx_cell":
        coord = (ref.sheet + "!" if ref.sheet else "") + (ref.cell or "A1")
        if ref.row >= 0:
            coord += f" 步进={ref.row}"
        return coord
    if kind == "xlsx_rows":
        return (ref.sheet + "!" if ref.sheet else "") + (ref.cell or "A1")
    if kind == "docx_table_cell":
        return f"table={ref.table},row={ref.row},col={ref.col}"
    if kind == "docx_table_rows":
        return f"table={ref.table},row_start={ref.row_start}"
    if kind == "pptx_table_cell":
        return f"slide={ref.slide},table={ref.table},row={ref.row},col={ref.col}"
    if kind in ("docx_placeholder", "pptx_placeholder"):
        return "{{" + (ref.placeholder or "字段") + "}}"
    return ""
