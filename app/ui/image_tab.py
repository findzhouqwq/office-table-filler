"""Tab: recognise images and batch-fill a spreadsheet."""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox,
                               QDoubleSpinBox, QFileDialog, QGroupBox,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QListWidget, QListWidgetItem, QMessageBox,
                               QPushButton, QSpinBox, QSplitter, QTableWidget,
                               QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget)

from ..core import ocr as ocr_module
from ..core.filler import ImageFillRequest, preview_image
from ..core.utils import clean_text
from .common import pick_files, pick_file, save_file, show_error, Worker, format_report
from .mapping_widgets import PlanTable, RuleTable

IMAGE_EXTS = ["图片文件 (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp)", "所有文件 (*.*)"]
TABLE_EXTS = ["Excel 表格 (*.xlsx *.xlsm *.xltx *.xltm)", "所有文件 (*.*)"]


class ImageTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.images: list[str] = []
        self.worker: Worker | None = None
        self._build()

    # -- UI -----------------------------------------------------------------
    def _build(self) -> None:
        root = QVBoxLayout(self)

        # ---- target group ----
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
        self.header_spin.setRange(0, 200)
        self.header_spin.setValue(1)
        row2.addWidget(self.header_spin)
        tgrid.addLayout(row2)
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("写入位置"))
        self.anchor_edit = QLineEdit("A1")
        row3.addWidget(self.anchor_edit, 1)
        self.append_box = QCheckBox("追加到已有内容之后")
        self.append_box.setChecked(True)
        row3.addWidget(self.append_box)
        tgrid.addLayout(row3)
        row4 = QHBoxLayout()
        row4.addWidget(QLabel("输出文件"))
        self.output_edit = QLineEdit()
        row4.addWidget(self.output_edit, 1)
        row4.addWidget(QPushButton("另存为…", self, clicked=lambda: self._pick_output()))
        tgrid.addLayout(row4)
        self.clear_box = QCheckBox("写入前清空目标工作表")
        tgrid.addWidget(self.clear_box)
        root.addWidget(target)

        # ---- ocr group ----
        ocr = QGroupBox("识别引擎")
        ogrid = QHBoxLayout(ocr)
        ogrid.addWidget(QLabel("OCR 引擎"))
        self.engine_combo = QComboBox()
        self.engine_combo.addItems(["自动", "rapidocr", "tesseract", "windows"])
        ogrid.addWidget(self.engine_combo, 1)
        self.detect_refresh_btn = QPushButton("刷新引擎")
        self.detect_refresh_btn.clicked.connect(self.refresh_engines)
        ogrid.addWidget(self.detect_refresh_btn)
        root.addWidget(ocr)

        # ---- mode group ----
        mode = QGroupBox("填写方式")
        mgrid = QHBoxLayout(mode)
        mgrid.addWidget(QLabel("方式"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["表格网格识别（整张图片是一个表格）",
                                  "字段提取（每张图片填一行）"])
        self.mode_combo.currentIndexChanged.connect(self._toggle_mode)
        mgrid.addWidget(self.mode_combo, 1)
        root.addWidget(mode)

        splitter = QSplitter(Qt.Vertical)

        # ---- left: images + preview ----
        left = QWidget()
        lgrid = QVBoxLayout(left)
        lgrid.addWidget(QLabel("图片列表"))
        self.image_list = QListWidget()
        lgrid.addWidget(self.image_list, 1)
        btns = QHBoxLayout()
        self.add_btn = QPushButton("添加图片…")
        self.add_btn.clicked.connect(self._add_images)
        self.remove_btn = QPushButton("移除")
        self.remove_btn.clicked.connect(lambda: self._remove_selected())
        self.clear_btn = QPushButton("清空")
        self.clear_btn.clicked.connect(self.image_list.clear)
        self.preview_btn = QPushButton("预览识别")
        self.preview_btn.clicked.connect(self._preview)
        btns.addWidget(self.add_btn)
        btns.addWidget(self.remove_btn)
        btns.addWidget(self.clear_btn)
        btns.addWidget(self.preview_btn)
        lgrid.addLayout(btns)
        self.preview_table = QTableWidget(0, 0)
        self.preview_table.setVerticalHeaderLabels(["识别结果"])
        self.preview_table.verticalHeader().setVisible(False)
        lgrid.addWidget(QLabel("识别预览（网格）"))
        lgrid.addWidget(self.preview_table, 1)
        splitter.addWidget(left)

        # ---- right: rules / plan ----
        right = QWidget()
        rgrid = QVBoxLayout(right)
        self.rule_title = QLabel("字段提取规则（每张图片生成一条记录）")
        rgrid.addWidget(self.rule_title)
        self.rule_table = RuleTable()
        rgrid.addWidget(self.rule_table, 2)
        rule_btns = QHBoxLayout()
        self.rule_add = QPushButton("添加规则")
        self.rule_add.clicked.connect(self.rule_table.add_row)
        self.rule_del = QPushButton("删除规则")
        self.rule_del.clicked.connect(lambda: self.rule_table.remove_row(
            self.rule_table.currentRow()))
        rule_btns.addWidget(self.rule_add)
        rule_btns.addWidget(self.rule_del)
        rule_btns.addStretch(1)
        rgrid.addLayout(rule_btns)
        rgrid.addWidget(QLabel("目标列映射（目标列 ↔ 提取字段）"))
        self.plan_table = PlanTable()
        rgrid.addWidget(self.plan_table, 2)
        plan_btns = QHBoxLayout()
        self.plan_add = QPushButton("添加映射")
        self.plan_add.clicked.connect(self.plan_table.add_row)
        self.plan_del = QPushButton("删除映射")
        self.plan_del.clicked.connect(lambda: self.plan_table.remove_row(
            self.plan_table.currentRow()))
        plan_btns.addWidget(self.plan_add)
        plan_btns.addWidget(self.plan_del)
        plan_btns.addStretch(1)
        rgrid.addLayout(plan_btns)
        splitter.addWidget(right)
        splitter.setSizes([420, 320])
        root.addWidget(splitter, 1)

        # ---- run / report ----
        run = QHBoxLayout()
        self.run_btn = QPushButton("开始批量填表")
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

        self._toggle_mode()
        self.refresh_engines()

    # -- helpers ------------------------------------------------------------
    def refresh_engines(self) -> None:
        available = ocr_module.available_engines()
        self.engine_combo.clear()
        self.engine_combo.addItems(["自动"])
        self.engine_combo.addItems(available if available else ["（无可用引擎）"])
        if not available:
            self.status_label.setText("未检测到 OCR 引擎，请在“识别引擎”中安装 rapidocr。")

    def _toggle_mode(self) -> None:
        grid_mode = self.mode_combo.currentIndex() == 0
        self.rule_title.setEnabled(not grid_mode)
        self.rule_table.setEnabled(not grid_mode)
        self.rule_add.setEnabled(not grid_mode)
        self.rule_del.setEnabled(not grid_mode)
        self.plan_table.setEnabled(not grid_mode)
        self.plan_add.setEnabled(not grid_mode)
        self.plan_del.setEnabled(not grid_mode)

    def _pick_target(self) -> None:
        files = pick_files(self, "选择目标表格", TABLE_EXTS, multi=False)
        if files:
            self.target_edit.setText(files[0])
            self.output_edit.setText(files[0])
            self._load_sheets(files[0])

    def _pick_output(self) -> None:
        path = save_file(self, "另存为")
        if path:
            self.output_edit.setText(path)

    def _load_sheets(self, path: str) -> None:
        from ..core.spreadsheet import OpenpyxlEngine
        try:
            eng = OpenpyxlEngine()
            eng.open(path)
            names = eng.sheets()
            eng.close()
        except Exception as exc:
            show_error(self, "无法读取表格", str(exc))
            return
        self.sheet_combo.clear()
        self.sheet_combo.addItems(names)

    def _add_images(self) -> None:
        files = pick_files(self, "选择图片", IMAGE_EXTS, multi=True)
        for path in files:
            item = QListWidgetItem(path)
            self.image_list.addItem(item)

    def _remove_selected(self) -> None:
        for item in self.image_list.selectedItems():
            self.image_list.takeItem(self.image_list.row(item))

    def _images(self) -> list[str]:
        return [self.image_list.item(i).text()
                for i in range(self.image_list.count())]

    # -- preview ------------------------------------------------------------
    def _preview(self) -> None:
        items = self.image_list.selectedItems() or [
            self.image_list.item(0)] if self.image_list.count() else []
        if not items:
            QMessageBox.information(self, "提示", "请先添加图片。")
            return
        path = items[0].text()
        self.status_label.setText(f"正在识别 {path} …")
        engine_pref = self.engine_combo.currentText()
        if engine_pref == "自动":
            engine_pref = "auto"
        self.worker = Worker(preview_image, path, None, engine_pref, parent=self)
        self.worker.finished_ok.connect(self._on_preview)
        self.worker.failed.connect(self._on_fail)
        self.worker.start()

    def _on_preview(self, data: dict) -> None:
        self.status_label.setText(
            f"识别完成：{data['engine']}，{len(data['words'])} 个词，"
            f"{data['rows']}行 × {data['cols']}列，耗时 {data['elapsed']:.2f}s")
        self._render_preview(data)

    def _render_preview(self, data: dict) -> None:
        matrix = data.get("matrix") or []
        self.preview_table.clear()
        if not matrix:
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            return
        self.preview_table.setRowCount(len(matrix))
        self.preview_table.setColumnCount(len(matrix[0]))
        for r, row in enumerate(matrix):
            for c, value in enumerate(row):
                self.preview_table.setItem(r, c, QTableWidgetItem(clean_text(value)))

    def _on_fail(self, message: str) -> None:
        self.status_label.setText("识别失败。")
        show_error(self, "识别失败", message)

    # -- run ---------------------------------------------------------------
    def _run(self) -> None:
        images = self._images()
        target = self.target_edit.text()
        if not images:
            show_error(self, "缺少图片", "请先添加至少一张图片。")
            return
        if not target or not target.strip():
            show_error(self, "缺少目标", "请选择目标表格。")
            return

        from ..core.models import ExtractionRule, FieldMapping
        rules = [] if self.mode_combo.currentIndex() == 0 else self.rule_table.to_rules()
        plan = [] if self.mode_combo.currentIndex() == 0 else self.plan_table.to_plan()
        if self.mode_combo.currentIndex() == 1 and not plan:
            plan = [FieldMapping(target="A", source="内容")]

        request = ImageFillRequest(
            images=images,
            target_path=target,
            output_path=self.output_edit.text().strip(),
            sheet=self.sheet_combo.currentText(),
            mode="grid" if self.mode_combo.currentIndex() == 0 else "fields",
            anchor=self.anchor_edit.text().strip() or "A1",
            append=self.append_box.isChecked(),
            clear_sheet=self.clear_box.isChecked(),
            header_row=max(1, self.header_spin.value()),
            ocr_engine=self.engine_combo.currentText(),
            extraction_rules=rules,
            column_plan=plan,
        )
        engine_pref = self.engine_combo.currentText()
        self.run_btn.setEnabled(False)
        self.status_label.setText("正在批量识别并填表…")

        def job(progress=lambda *_: None):
            engine = ocr_module.get_ocr_engine(engine_pref)
            from ..core.filler import fill_from_images
            return fill_from_images(request, ocr_engine=engine)

        self.worker = Worker(job, parent=self)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self.run_btn.setEnabled(True))
        self.worker.start()

    def _on_done(self, report) -> None:
        self.status_label.setText(f"完成：写入 {report.written} 处。")
        lines = [format_report(report), ""]
        for record in report.records[:80]:
            lines.append(f"{record.target}  ←  {record.source}  =  {record.value}")
        self.report_edit.setPlainText("\n".join(lines))

    def on_before_close(self) -> bool:
        if self.worker is not None and self.worker.isRunning():
            self.worker.wait(3000)
        return True
