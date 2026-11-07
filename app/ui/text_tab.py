"""Tab: fill a spreadsheet from structured text / CSV / JSON."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QPushButton, QSpinBox,
                               QSplitter, QTableWidget, QTableWidgetItem,
                               QTextEdit, QVBoxLayout, QWidget)

from ..core.data_io import load_table
from ..core.filler import TextFillRequest, fill_from_text
from ..core.mapping import auto_map_headers
from ..core.models import FieldMapping
from ..core.utils import clean_text
from .common import pick_files, save_file, show_error, Worker, format_report
from .mapping_widgets import PlanTable

SOURCE_EXTS = ["数据文件 (*.csv *.txt *.tsv *.json *.xlsx *.xlsm)", "所有文件 (*.*)"]
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
        from ..core.spreadsheet import OpenpyxlEngine
        eng = OpenpyxlEngine()
        try:
            eng.open(files[0])
            names = eng.sheets()
            self.sheet_combo.clear()
            self.sheet_combo.addItems(names)
            sheet = self.sheet_combo.currentText()
            max_row, max_col = eng.dims(sheet)
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

    def _run(self) -> None:
        request = self._build_request()
        if not request.target_path:
            show_error(self, "缺少目标", "请选择目标表格。")
            return
        if not request.records and not request.text.strip():
            show_error(self, "缺少数据", "请添加来源文件或粘贴文本。")
            return
        self.run_btn.setEnabled(False)
        self.status_label.setText("正在填表…")
        self.worker = Worker(lambda: fill_from_text(request), parent=self)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self.run_btn.setEnabled(True))
        self.worker.start()

    def _on_done(self, report) -> None:
        self.status_label.setText(f"完成：写入 {report.written} 处。")
        self.report_edit.setPlainText(format_report(report))

    def _on_fail(self, message: str) -> None:
        self.status_label.setText("失败。")
        show_error(self, "填表失败", message)

    def on_before_close(self) -> bool:
        if self.worker is not None and self.worker.isRunning():
            self.worker.wait(3000)
        return True
