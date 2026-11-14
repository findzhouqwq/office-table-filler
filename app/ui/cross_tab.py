"""Tab: read several source files and fill several target files."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QGroupBox, QHBoxLayout, QLabel,
                               QLineEdit, QListWidget, QListWidgetItem,
                               QPushButton, QSpinBox, QSplitter, QTabWidget,
                               QTableWidget, QTableWidgetItem, QTextEdit,
                               QVBoxLayout, QWidget)

from ..core import templates
from ..core.data_io import load_docx, load_pptx, load_table
from ..core.errors import AppError
from ..core.filler import CrossFillRequest, fill_cross_files
from ..core.models import MappingEntry
from ..core.preview import build_cross_preview
from ..core.utils import clean_text
from .common import pick_files, pick_dir, show_error, Worker, format_report
from .mapping_widgets import TargetTable
from .preview_dialog import show_preview
from .template_bar import TemplateBar
from .theme import (Card, CollapsibleCard, ghost_button, hint_label,
                    primary_button)

SOURCE_EXTS = ["数据 / PDF 文件 (*.csv *.txt *.tsv *.json *.xlsx *.xlsm *.pdf)",
               "所有文件 (*.*)"]
TARGET_EXTS = ["目标文件 (*.xlsx *.xlsm *.docx *.pptx)", "所有文件 (*.*)"]


class CrossTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.source_paths: list[str] = []
        self.target_paths: list[str] = []
        self.headers: list[str] = []
        self.records: list[dict] = []
        self.worker: Worker | None = None
        self._build()

    def _build(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        self.template_bar = TemplateBar("cross")
        self.template_bar.save_requested.connect(self._save_template)
        self.template_bar.load_requested.connect(self._load_template)
        root.addWidget(self.template_bar)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        # ---- left: source files + preview ----
        left = QWidget()
        lgrid = QVBoxLayout(left)
        lgrid.setContentsMargins(0, 0, 0, 0)
        lgrid.setSpacing(0)

        src = Card("源数据", "读取数据，可多选")
        self.source_tabs = QTabWidget()
        self.source_tabs.setObjectName("DetailTabs")
        self.source_tabs.setDocumentMode(True)

        files_page = QWidget()
        fgrid = QVBoxLayout(files_page)
        fgrid.setContentsMargins(0, 6, 0, 0)
        self.source_list = QListWidget()
        self.source_list.setAlternatingRowColors(True)
        fgrid.addWidget(self.source_list, 1)
        sbtns = QHBoxLayout()
        sbtns.addWidget(QPushButton("添加…", self, clicked=lambda: self._add_sources()))
        sbtns.addWidget(QPushButton("移除", self, clicked=lambda: self._remove(self.source_list)))
        sbtns.addWidget(QPushButton("清空", self, clicked=lambda: self.source_list.clear()))
        sbtns.addStretch(1)
        fgrid.addLayout(sbtns)
        sopt = QHBoxLayout()
        sopt.addWidget(QLabel("工作表"))
        self.src_sheet_edit = QLineEdit()
        sopt.addWidget(self.src_sheet_edit, 1)
        sopt.addWidget(QLabel("表头行"))
        self.src_header_spin = QSpinBox()
        self.src_header_spin.setRange(1, 200)
        self.src_header_spin.setValue(1)
        sopt.addWidget(self.src_header_spin)
        fgrid.addLayout(sopt)
        load_btn = ghost_button("载入并预览")
        load_btn.clicked.connect(self._load_sources)
        fgrid.addWidget(load_btn)
        self.source_tabs.addTab(files_page, "源文件")

        preview_page = QWidget()
        vgrid = QVBoxLayout(preview_page)
        vgrid.setContentsMargins(0, 6, 0, 0)
        self.preview_table = QTableWidget(0, 0)
        self.preview_table.setAlternatingRowColors(True)
        vgrid.addWidget(self.preview_table, 1)
        self.source_tabs.addTab(preview_page, "数据预览")

        src.body.addWidget(self.source_tabs, 1)
        lgrid.addWidget(src, 1)
        splitter.addWidget(left)

        # ---- right: targets + output ----
        right = QWidget()
        rgrid = QVBoxLayout(right)
        rgrid.setContentsMargins(0, 0, 0, 0)
        rgrid.setSpacing(0)

        tgt = Card("目标与输出", "写入结果，可多选")
        self.target_tabs = QTabWidget()
        self.target_tabs.setObjectName("DetailTabs")
        self.target_tabs.setDocumentMode(True)

        target_page = QWidget()
        tg = QVBoxLayout(target_page)
        tg.setContentsMargins(0, 6, 0, 0)
        self.target_list = QListWidget()
        self.target_list.setAlternatingRowColors(True)
        tg.addWidget(self.target_list, 1)
        tbtns = QHBoxLayout()
        tbtns.addWidget(QPushButton("添加…", self, clicked=lambda: self._add_targets()))
        tbtns.addWidget(QPushButton("移除", self, clicked=lambda: self._remove(self.target_list)))
        tbtns.addWidget(QPushButton("清空", self, clicked=lambda: self.target_list.clear()))
        tbtns.addStretch(1)
        tg.addLayout(tbtns)
        inspect_btn = ghost_button("查看目标结构")
        inspect_btn.clicked.connect(self._inspect_targets)
        tg.addWidget(inspect_btn)
        self.target_tabs.addTab(target_page, "目标文件")

        out_page = QWidget()
        og = QVBoxLayout(out_page)
        og.setContentsMargins(0, 6, 0, 0)
        row = QHBoxLayout()
        row.addWidget(QLabel("输出目录"))
        self.outdir_edit = QLineEdit()
        self.outdir_edit.setPlaceholderText("默认目标文件所在目录")
        row.addWidget(self.outdir_edit, 1)
        row.addWidget(QPushButton("选择…", self, clicked=lambda: self._pick_dir()))
        og.addLayout(row)
        self.batch_box = QCheckBox("按源数据逐行生成（多条记录 → 多份文件）")
        self.batch_box.setChecked(True)
        og.addWidget(self.batch_box)
        og.addStretch(1)
        self.target_tabs.addTab(out_page, "输出设置")

        tgt.body.addWidget(self.target_tabs, 1)
        rgrid.addWidget(tgt, 1)
        splitter.addWidget(right)
        splitter.setSizes([560, 460])
        root.addWidget(splitter, 2)

        map_card = Card("字段映射", "源字段 → 目标位置")
        self.mapping_table = TargetTable()
        map_card.body.addWidget(self.mapping_table, 1)
        mbtns = QHBoxLayout()
        mbtns.addWidget(QPushButton("添加映射", self, clicked=lambda: self.mapping_table.add_row()))
        mbtns.addWidget(QPushButton("删除映射", self, clicked=lambda: self.mapping_table.remove_row(
            self.mapping_table.currentRow())))
        mbtns.addStretch(1)
        map_card.body.addLayout(mbtns)
        root.addWidget(map_card, 2)

        run = QHBoxLayout()
        run.setSpacing(10)
        self.validate_btn = QPushButton("预览 / 校验")
        self.validate_btn.clicked.connect(self._preview_run)
        run.addWidget(self.validate_btn)
        self.run_btn = primary_button("开始跨文件填表")
        self.run_btn.clicked.connect(self._run)
        run.addWidget(self.run_btn)
        run.addStretch(1)
        self.status_label = QLabel("就绪。")
        self.status_label.setObjectName("StatusText")
        run.addWidget(self.status_label)
        root.addLayout(run)

        self.report_card = CollapsibleCard("结构信息 / 执行报告", expanded=False)
        self.info_edit = QTextEdit()
        self.info_edit.setReadOnly(True)
        self.info_edit.setMinimumHeight(90)
        self.report_card.body.addWidget(self.info_edit)
        root.addWidget(self.report_card)

    # -- helpers ------------------------------------------------------------
    def _remove(self, widget: QListWidget) -> None:
        for item in widget.selectedItems():
            widget.takeItem(widget.row(item))

    def _add_sources(self) -> None:
        files = pick_files(self, "选择源文件", SOURCE_EXTS, multi=True)
        for path in files:
            self.source_list.addItem(QListWidgetItem(path))

    def _add_targets(self) -> None:
        files = pick_files(self, "选择目标文件", TARGET_EXTS, multi=True)
        for path in files:
            self.target_list.addItem(QListWidgetItem(path))

    def _pick_dir(self) -> None:
        path = pick_dir(self, "选择输出目录")
        if path:
            self.outdir_edit.setText(path)

    def _load_sources(self) -> None:
        paths = [self.source_list.item(i).text() for i in range(self.source_list.count())]
        if not paths:
            show_error(self, "缺少源", "请先添加源文件。")
            return
        try:
            table = load_table(paths[0], sheet=self.src_sheet_edit.text().strip(),
                               header_row=max(1, self.src_header_spin.value()))
            self.headers = table.headers
            self.records = table.as_records()
        except Exception as exc:
            show_error(self, "读取源失败", str(exc))
            return
        self.preview_table.clear()
        if self.headers:
            self.preview_table.setRowCount(min(len(self.records), 200))
            self.preview_table.setColumnCount(len(self.headers))
            self.preview_table.setHorizontalHeaderLabels(self.headers)
            for r, record in enumerate(self.records[:200]):
                for c, head in enumerate(self.headers):
                    self.preview_table.setItem(r, c, QTableWidgetItem(str(record.get(head, ""))))
        self.status_label.setText(f"源数据：{len(self.records)} 行 / {len(self.headers)} 列。")

    def _inspect_targets(self) -> None:
        paths = [self.target_list.item(i).text() for i in range(self.target_list.count())]
        lines: list[str] = []
        for path in paths:
            lower = path.lower()
            try:
                if lower.endswith((".docx", ".docm")):
                    model = load_docx(path)
                    lines.append(f"{path}\n  Word 表格数：{model.table_count}")
                    for i, table in enumerate(model.tables[:10]):
                        lines.append(f"    表{i}: {len(table)}行 × {len(table[0]) if table else 0}列")
                elif lower.endswith((".pptx", ".pptm")):
                    model = load_pptx(path)
                    lines.append(f"{path}\n  PPT 页数：{len(model.slides)}")
                    for slide in model.slides:
                        info = f"    第{slide.index + 1}页：表格{len(slide.tables)}个"
                        if slide.tables:
                            info += f"（{len(slide.tables[0])}行×{len(slide.tables[0][0])}列）"
                        lines.append(info)
                else:
                    from ..core.spreadsheet import OpenpyxlEngine
                    eng = OpenpyxlEngine()
                    eng.open(path)
                    names = eng.sheets()
                    lines.append(f"{path}\n  工作表：{', '.join(names)}")
                    eng.close()
            except Exception as exc:
                lines.append(f"{path}\n  读取失败：{exc}")
        self.info_edit.setPlainText("\n".join(lines))
        self.report_card.set_expanded(True)

    def _build_request(self) -> CrossFillRequest | None:
        sources = [self.source_list.item(i).text() for i in range(self.source_list.count())]
        targets = [self.target_list.item(i).text() for i in range(self.target_list.count())]
        if not sources:
            show_error(self, "缺少源", "请添加源文件。")
            return None
        if not targets:
            show_error(self, "缺少目标", "请添加目标文件。")
            return None
        entries = self.mapping_table.to_entries(targets[0])
        if not entries:
            show_error(self, "缺少映射", "请添加至少一条字段映射。")
            return None
        return CrossFillRequest(
            source_paths=sources,
            source_sheet=self.src_sheet_edit.text().strip(),
            source_header_row=max(1, self.src_header_spin.value()),
            target_paths=targets,
            mappings=entries,
            output_dir=self.outdir_edit.text().strip(),
            batch_by_row=self.batch_box.isChecked(),
        )

    def _set_busy(self, busy: bool) -> None:
        self.run_btn.setEnabled(not busy)
        self.validate_btn.setEnabled(not busy)

    def _run(self) -> None:
        request = self._build_request()
        if request is None:
            return
        self._set_busy(True)
        self.status_label.setText("正在跨文件填表…")
        self.worker = Worker(lambda: fill_cross_files(request), parent=self)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _preview_run(self) -> None:
        request = self._build_request()
        if request is None:
            return
        result = build_cross_preview(request)
        self.status_label.setText(
            f"预览完成：{result.total_rows} 行源数据，"
            f"{len(result.issues.errors)} 个错误 / {len(result.issues.warnings)} 个提示。")
        if show_preview(self, result):
            self._run()

    def _on_done(self, report) -> None:
        self.status_label.setText(f"完成：写入 {report.written} 处。")
        self.info_edit.setPlainText(format_report(report))
        self.report_card.set_expanded(True)

    def _on_fail(self, message: str) -> None:
        self.status_label.setText("失败。")
        show_error(self, "跨文件填表失败", message)

    # -- templates ---------------------------------------------------------
    def _template_data(self) -> dict:
        return {
            "sources": [self.source_list.item(i).text()
                        for i in range(self.source_list.count())],
            "targets": [self.target_list.item(i).text()
                        for i in range(self.target_list.count())],
            "source_sheet": self.src_sheet_edit.text().strip(),
            "source_header_row": self.src_header_spin.value(),
            "output_dir": self.outdir_edit.text().strip(),
            "batch_by_row": self.batch_box.isChecked(),
            "mappings": [entry.to_dict() for entry in
                         self.mapping_table.to_entries("")],
        }

    def _apply_template_data(self, data: dict) -> None:
        self.source_list.clear()
        for path in data.get("sources") or []:
            self.source_list.addItem(QListWidgetItem(path))
        self.target_list.clear()
        for path in data.get("targets") or []:
            self.target_list.addItem(QListWidgetItem(path))
        self.src_sheet_edit.setText(data.get("source_sheet") or "")
        self.src_header_spin.setValue(int(data.get("source_header_row", 1) or 1))
        self.outdir_edit.setText(data.get("output_dir") or "")
        self.batch_box.setChecked(bool(data.get("batch_by_row", True)))
        self.mapping_table.load(
            [MappingEntry.from_dict(item) for item in data.get("mappings", [])])
        if self.source_list.count():
            self._load_sources()

    def _save_template(self, name: str) -> None:
        try:
            templates.save_template(name, "cross", self._template_data())
        except AppError as exc:
            show_error(self, "保存模板失败", str(exc))
            return
        self.template_bar.refresh()
        self.template_bar.select(name)
        self.status_label.setText(f"模板已保存：{name}")

    def _load_template(self, name: str) -> None:
        try:
            template = templates.load_template(name, "cross")
        except AppError as exc:
            show_error(self, "读取模板失败", str(exc))
            return
        self._apply_template_data(template.data)
        self.status_label.setText(f"已应用模板：{name}")

    def on_before_close(self) -> bool:
        if self.worker is not None and self.worker.isRunning():
            self.worker.wait(3000)
        return True
