"""Tab: read several source files and fill several target files."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QGroupBox, QHBoxLayout, QLabel,
                               QLineEdit, QListWidget, QListWidgetItem,
                               QPushButton, QSpinBox, QSplitter, QTableWidget,
                               QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget)

from ..core.data_io import load_docx, load_pptx, load_table
from ..core.filler import CrossFillRequest, fill_cross_files
from ..core.utils import clean_text
from .common import pick_files, pick_dir, show_error, Worker, format_report
from .mapping_widgets import TargetTable

SOURCE_EXTS = ["数据文件 (*.csv *.txt *.tsv *.json *.xlsx *.xlsm)", "所有文件 (*.*)"]
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

        splitter = QSplitter(Qt.Horizontal)

        # ---- left: sources ----
        left = QWidget()
        lgrid = QVBoxLayout(left)
        src = QGroupBox("源文件（读取数据）")
        sgrid = QVBoxLayout(src)
        srow = QHBoxLayout()
        self.source_list = QListWidget()
        srow.addWidget(self.source_list, 1)
        sgrid.addLayout(srow)
        sbtns = QHBoxLayout()
        sbtns.addWidget(QPushButton("添加…", self, clicked=lambda: self._add_sources()))
        sbtns.addWidget(QPushButton("移除", self, clicked=lambda: self._remove(self.source_list)))
        sbtns.addWidget(QPushButton("清空", self, clicked=lambda: self.source_list.clear()))
        sgrid.addLayout(sbtns)
        sopt = QHBoxLayout()
        sopt.addWidget(QLabel("工作表"))
        self.src_sheet_edit = QLineEdit()
        sopt.addWidget(self.src_sheet_edit, 1)
        sopt.addWidget(QLabel("表头行"))
        self.src_header_spin = QSpinBox()
        self.src_header_spin.setRange(1, 200)
        self.src_header_spin.setValue(1)
        sopt.addWidget(self.src_header_spin)
        sgrid.addLayout(sopt)
        sgrid.addWidget(QPushButton("载入并预览", self, clicked=lambda: self._load_sources()))
        lgrid.addWidget(src)
        lgrid.addWidget(QLabel("源数据预览"))
        self.preview_table = QTableWidget(0, 0)
        lgrid.addWidget(self.preview_table, 1)
        splitter.addWidget(left)

        # ---- right: targets ----
        right = QWidget()
        rgrid = QVBoxLayout(right)
        tgt = QGroupBox("目标文件（写入结果）")
        tgrid = QVBoxLayout(tgt)
        trow = QHBoxLayout()
        self.target_list = QListWidget()
        trow.addWidget(self.target_list, 1)
        tgrid.addLayout(trow)
        tbtns = QHBoxLayout()
        tbtns.addWidget(QPushButton("添加…", self, clicked=lambda: self._add_targets()))
        tbtns.addWidget(QPushButton("移除", self, clicked=lambda: self._remove(self.target_list)))
        tbtns.addWidget(QPushButton("清空", self, clicked=lambda: self.target_list.clear()))
        tgrid.addLayout(tbtns)
        tgrid.addWidget(QPushButton("查看目标结构", self, clicked=lambda: self._inspect_targets()))
        rgrid.addWidget(tgt)
        rgrid.addWidget(QLabel("输出目录"))
        row = QHBoxLayout()
        self.outdir_edit = QLineEdit()
        row.addWidget(self.outdir_edit, 1)
        row.addWidget(QPushButton("选择…", self, clicked=lambda: self._pick_dir()))
        rgrid.addLayout(row)
        row2 = QHBoxLayout()
        self.batch_box = QCheckBox("按源数据逐行生成（多条记录→多份文件）")
        self.batch_box.setChecked(True)
        row2.addWidget(self.batch_box)
        row2.addStretch(1)
        rgrid.addLayout(row2)
        splitter.addWidget(right)
        splitter.setSizes([480, 420])
        root.addWidget(splitter, 1)

        root.addWidget(QLabel("字段映射（源字段 → 目标位置）"))
        self.mapping_table = TargetTable()
        root.addWidget(self.mapping_table, 1)
        mbtns = QHBoxLayout()
        mbtns.addWidget(QPushButton("添加映射", self, clicked=lambda: self.mapping_table.add_row()))
        mbtns.addWidget(QPushButton("删除映射", self, clicked=lambda: self.mapping_table.remove_row(
            self.mapping_table.currentRow())))
        mbtns.addStretch(1)
        root.addLayout(mbtns)

        run = QHBoxLayout()
        self.run_btn = QPushButton("开始跨文件填表")
        self.run_btn.setStyleSheet("font-weight: bold; padding: 6px;")
        self.run_btn.clicked.connect(self._run)
        run.addWidget(self.run_btn)
        run.addStretch(1)
        root.addLayout(run)
        self.status_label = QLabel("就绪。")
        root.addWidget(self.status_label)
        self.info_edit = QTextEdit()
        self.info_edit.setReadOnly(True)
        root.addWidget(QLabel("结构信息 / 报告"))
        root.addWidget(self.info_edit, 1)

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

    def _run(self) -> None:
        sources = [self.source_list.item(i).text() for i in range(self.source_list.count())]
        targets = [self.target_list.item(i).text() for i in range(self.target_list.count())]
        if not sources:
            show_error(self, "缺少源", "请添加源文件。")
            return
        if not targets:
            show_error(self, "缺少目标", "请添加目标文件。")
            return
        entries = self.mapping_table.to_entries(targets[0])
        if not entries:
            show_error(self, "缺少映射", "请添加至少一条字段映射。")
            return
        request = CrossFillRequest(
            source_paths=sources,
            source_sheet=self.src_sheet_edit.text().strip(),
            source_header_row=max(1, self.src_header_spin.value()),
            target_paths=targets,
            mappings=entries,
            output_dir=self.outdir_edit.text().strip(),
            batch_by_row=self.batch_box.isChecked(),
        )
        self.run_btn.setEnabled(False)
        self.status_label.setText("正在跨文件填表…")
        self.worker = Worker(lambda: fill_cross_files(request), parent=self)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self.run_btn.setEnabled(True))
        self.worker.start()

    def _on_done(self, report) -> None:
        self.status_label.setText(f"完成：写入 {report.written} 处。")
        self.info_edit.setPlainText(format_report(report))

    def _on_fail(self, message: str) -> None:
        self.status_label.setText("失败。")
        show_error(self, "跨文件填表失败", message)

    def on_before_close(self) -> bool:
        if self.worker is not None and self.worker.isRunning():
            self.worker.wait(3000)
        return True
