"""Tab: fill a spreadsheet from structured text / CSV / JSON."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QPushButton, QSpinBox,
                               QSplitter, QTableWidget, QTableWidgetItem,
                               QTextEdit, QVBoxLayout, QWidget)

from ..core import templates
from ..core.data_io import load_table
from ..core.errors import AppError
from ..core.filler import TextFillRequest, fill_from_text
from ..core.mapping import auto_map_headers
from ..core.models import FieldMapping
from ..core.preview import build_text_preview
from ..core.utils import clean_text
from .common import pick_files, save_file, show_error, Worker, format_report
from .mapping_widgets import PlanTable
from .preview_dialog import show_preview
from .template_bar import TemplateBar

SOURCE_EXTS = ["数据 / PDF 文件 (*.csv *.txt *.tsv *.json *.xlsx *.xlsm *.pdf)",
               "所有文件 (*.*)"]
TABLE_EXTS = ["Excel 表格 (*.xlsx *.xlsm *.xltx *.xltm)", "所有文件 (*.*)"]


class TextTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.records: list[dict] = []
        self.headers: list[str] = []
        self.worker: Worker | None = None
        self._build()

    def _build(self) -> None:
        root = QVBoxLayout(self)

        self.template_bar = TemplateBar("text")
        self.template_bar.save_requested.connect(self._save_template)
        self.template_bar.load_requested.connect(self._load_template)
        root.addWidget(self.template_bar)

        # ---- target ----
        target = QGroupBox("目标表格")
        tgrid = QVBoxLayout(target)
        row1 = QHBoxLayout()
        self.target_edit = QLineEdit()
        self.target_edit.setReadOnly(True)
        row1.addWidget(self.target_edit, 1)
        row1.addWidget(QPushButton("选择…", self, clicked=lambda: self._pick_target()))
        tgrid.addLayout(row1)
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("工作表"))
        self.sheet_combo = QComboBox()
        row2.addWidget(self.sheet_combo, 1)
        row2.addWidget(QLabel("表头行"))
        self.header_spin = QSpinBox()
        self.header_spin.setRange(1, 200)
        self.header_spin.setValue(1)
        row2.addWidget(self.header_spin)
        tgrid.addLayout(row2)
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("输出文件"))
        self.output_edit = QLineEdit()
        row3.addWidget(self.output_edit, 1)
        row3.addWidget(QPushButton("另存为…", self, clicked=lambda: self._pick_output()))
        tgrid.addLayout(row3)
        row4 = QHBoxLayout()
        self.append_box = QCheckBox("追加到已有内容之后")
        self.append_box.setChecked(True)
        self.clear_box = QCheckBox("写入前清空目标工作表")
        row4.addWidget(self.append_box)
        row4.addWidget(self.clear_box)
        row4.addStretch(1)
        tgrid.addLayout(row4)
        root.addWidget(target)

        splitter = QSplitter(Qt.Horizontal)

        # ---- left: source ----
        left = QWidget()
        lgrid = QVBoxLayout(left)
        lgrid.addWidget(QLabel("来源文件"))
        self.source_list = QListWidget()
        lgrid.addWidget(self.source_list, 1)
        lbtns = QHBoxLayout()
        lbtns.addWidget(QPushButton("添加文件…", self, clicked=lambda: self._add_sources()))
        lbtns.addWidget(QPushButton("移除", self, clicked=lambda: self._remove_source()))
        lbtns.addWidget(QPushButton("清空", self, clicked=lambda: self.source_list.clear()))
        lgrid.addLayout(lbtns)
        lgrid.addWidget(QLabel("数据预览"))
        self.preview_table = QTableWidget(0, 0)
        lgrid.addWidget(self.preview_table, 1)
        lgrid.addWidget(QLabel("或直接粘贴文本（每段一条记录，可用空行分隔）"))
        self.text_edit = QTextEdit()
        lgrid.addWidget(self.text_edit, 1)
        splitter.addWidget(left)

        # ---- right: mapping ----
        right = QWidget()
        rgrid = QVBoxLayout(right)
        rgrid.addWidget(QLabel("目标列映射（目标列 ↔ 源字段）"))
        self.plan_table = PlanTable()
        rgrid.addWidget(self.plan_table, 1)
        pbtns = QHBoxLayout()
        pbtns.addWidget(QPushButton("添加映射", self, clicked=lambda: self.plan_table.add_row()))
        pbtns.addWidget(QPushButton("删除映射", self, clicked=lambda: self.plan_table.remove_row(
            self.plan_table.currentRow())))
        pbtns.addStretch(1)
        rgrid.addLayout(pbtns)
        rgrid.addWidget(QLabel("目标表头（用于自动匹配）"))
        self.target_header_edit = QTextEdit()
        self.target_header_edit.setMaximumHeight(70)
        rgrid.addWidget(self.target_header_edit)
        rgrid.addWidget(QPushButton("按同名表头自动生成映射", self,
                                    clicked=lambda: self._auto_map()))
        splitter.addWidget(right)
        splitter.setSizes([460, 380])
        root.addWidget(splitter, 1)

        run = QHBoxLayout()
        self.validate_btn = QPushButton("预览 / 校验")
        self.validate_btn.clicked.connect(self._preview_run)
        run.addWidget(self.validate_btn)
        self.run_btn = QPushButton("开始填表")
        self.run_btn.setStyleSheet("font-weight: bold; padding: 6px;")
        self.run_btn.clicked.connect(self._run)
        run.addWidget(self.run_btn)
        run.addStretch(1)
        root.addLayout(run)
        self.status_label = QLabel("就绪。")
        root.addWidget(self.status_label)
        self.report_edit = QTextEdit()
        self.report_edit.setReadOnly(True)
        root.addWidget(QLabel("执行报告"))
        root.addWidget(self.report_edit, 1)

    # -- helpers ------------------------------------------------------------
    def _pick_target(self) -> None:
        files = pick_files(self, "选择目标表格", TABLE_EXTS, multi=False)
        if not files:
            return
        self.target_edit.setText(files[0])
        self.output_edit.setText(files[0])
        self._load_target(files[0])

    def _load_target(self, path: str) -> None:
        from ..core.spreadsheet import OpenpyxlEngine
        eng = OpenpyxlEngine()
        try:
            eng.open(path)
            names = eng.sheets()
            self.sheet_combo.clear()
            self.sheet_combo.addItems(names)
            sheet = self.sheet_combo.currentText()
            _, max_col = eng.dims(sheet)
            header_row = max(1, self.header_spin.value())
            from openpyxl.utils import get_column_letter
            headers = []
            for col in range(1, max_col + 1):
                value = clean_text(eng.cell_value(sheet, f"{get_column_letter(col)}{header_row}"))
                if value:
                    headers.append(value)
            self.target_header_edit.setPlainText("、".join(headers))
        except Exception as exc:
            show_error(self, "无法读取表格", str(exc))
        finally:
            eng.close()

    def _pick_output(self) -> None:
        path = save_file(self, "另存为")
        if path:
            self.output_edit.setText(path)

    def _sources(self) -> list[str]:
        return [self.source_list.item(i).text()
                for i in range(self.source_list.count())]

    def _add_sources(self) -> None:
        files = pick_files(self, "选择来源数据", SOURCE_EXTS, multi=True)
        for path in files:
            self.source_list.addItem(QListWidgetItem(path))
        self._load_sources()

    def _remove_source(self) -> None:
        for item in self.source_list.selectedItems():
            self.source_list.takeItem(self.source_list.row(item))
        self._load_sources()

    def _load_sources(self) -> None:
        paths = self._sources()
        if not paths:
            return
        try:
            table = load_table(paths[0])
            self.headers = table.headers
            self.records = table.as_records()
        except Exception as exc:
            show_error(self, "读取来源失败", str(exc))
            return
        self._render_preview()
        self.status_label.setText(f"已载入 {len(self.records)} 行 / {len(self.headers)} 列。")

    def _render_preview(self) -> None:
        self.preview_table.clear()
        if not self.headers:
            return
        self.preview_table.setRowCount(min(len(self.records), 200))
        self.preview_table.setColumnCount(len(self.headers))
        self.preview_table.horizontalHeader().setVisible(True)
        self.preview_table.setHorizontalHeaderLabels(self.headers)
        for r, record in enumerate(self.records[:200]):
            for c, head in enumerate(self.headers):
                self.preview_table.setItem(r, c, QTableWidgetItem(str(record.get(head, ""))))

    def _auto_map(self) -> None:
        target_headers = [h for h in clean_text(self.target_header_edit.toPlainText()).split("、") if h]
        if not target_headers:
            show_error(self, "缺少目标表头", "请先读取目标表格。")
            return
        pairs = auto_map_headers(self.headers, target_headers)
        plan = [FieldMapping(target=target_headers[t], source=self.headers[s])
                for s, t in pairs]
        if not plan:
            plan = [FieldMapping(target=h, source=h) for h in target_headers]
        self.plan_table.load(plan)

    def _build_request(self) -> TextFillRequest:
        text = self.text_edit.toPlainText()
        records = list(self.records)
        if not records and text.strip():
            records = []
        return TextFillRequest(
            text=text,
            records=records,
            target_path=self.target_edit.text().strip(),
            output_path=self.output_edit.text().strip(),
            sheet=self.sheet_combo.currentText(),
            header_row=max(1, self.header_spin.value()),
            append=self.append_box.isChecked(),
            clear_sheet=self.clear_box.isChecked(),
            column_plan=self.plan_table.to_plan(),
        )

    def _set_busy(self, busy: bool) -> None:
        self.run_btn.setEnabled(not busy)
        self.validate_btn.setEnabled(not busy)

    def _run(self) -> None:
        request = self._build_request()
        if not request.target_path:
            show_error(self, "缺少目标", "请选择目标表格。")
            return
        if not request.records and not request.text.strip():
            show_error(self, "缺少数据", "请添加来源文件或粘贴文本。")
            return
        self._set_busy(True)
        self.status_label.setText("正在填表…")
        self.worker = Worker(lambda: fill_from_text(request), parent=self)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _preview_run(self) -> None:
        request = self._build_request()
        if not request.target_path:
            show_error(self, "缺少目标", "请选择目标表格。")
            return
        if not request.records and not request.text.strip():
            show_error(self, "缺少数据", "请添加来源文件或粘贴文本。")
            return
        result = build_text_preview(request)
        self.status_label.setText(
            f"预览完成：{result.total_rows} 行，"
            f"{len(result.issues.errors)} 个错误 / {len(result.issues.warnings)} 个提示。")
        if show_preview(self, result):
            self._run()

    def _on_done(self, report) -> None:
        self.status_label.setText(f"完成：写入 {report.written} 处。")
        self.report_edit.setPlainText(format_report(report))

    def _on_fail(self, message: str) -> None:
        self.status_label.setText("失败。")
        show_error(self, "填表失败", message)

    # -- templates ---------------------------------------------------------
    def _template_data(self) -> dict:
        return {
            "sources": self._sources(),
            "target_path": self.target_edit.text().strip(),
            "output_path": self.output_edit.text().strip(),
            "sheet": self.sheet_combo.currentText(),
            "header_row": self.header_spin.value(),
            "append": self.append_box.isChecked(),
            "clear_sheet": self.clear_box.isChecked(),
            "column_plan": [plan.to_dict() for plan in self.plan_table.to_plan()],
        }

    def _apply_template_data(self, data: dict) -> None:
        self.source_list.clear()
        for path in data.get("sources") or []:
            self.source_list.addItem(QListWidgetItem(path))
        target = data.get("target_path") or ""
        if target:
            self.target_edit.setText(target)
            self._load_target(target)
        self.output_edit.setText(data.get("output_path") or "")
        sheet = data.get("sheet") or ""
        if sheet:
            index = self.sheet_combo.findText(sheet)
            if index >= 0:
                self.sheet_combo.setCurrentIndex(index)
        self.header_spin.setValue(int(data.get("header_row", 1) or 1))
        self.append_box.setChecked(bool(data.get("append", True)))
        self.clear_box.setChecked(bool(data.get("clear_sheet", False)))
        self.plan_table.load(
            [FieldMapping.from_dict(item) for item in data.get("column_plan", [])])
        if self._sources():
            self._load_sources()

    def _save_template(self, name: str) -> None:
        try:
            templates.save_template(name, "text", self._template_data())
        except AppError as exc:
            show_error(self, "保存模板失败", str(exc))
            return
        self.template_bar.refresh()
        self.template_bar.select(name)
        self.status_label.setText(f"模板已保存：{name}")

    def _load_template(self, name: str) -> None:
        try:
            template = templates.load_template(name, "text")
        except AppError as exc:
            show_error(self, "读取模板失败", str(exc))
            return
        self._apply_template_data(template.data)
        self.status_label.setText(f"已应用模板：{name}")

    def on_before_close(self) -> bool:
        if self.worker is not None and self.worker.isRunning():
            self.worker.wait(3000)
        return True
