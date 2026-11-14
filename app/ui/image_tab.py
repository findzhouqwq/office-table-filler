"""Tab: recognise images and batch-fill a spreadsheet."""

from __future__ import annotations

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox,
                               QDoubleSpinBox, QFileDialog, QGroupBox,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit,
                               QListWidget, QListWidgetItem, QMessageBox,
                               QPushButton, QSpinBox, QSplitter, QTabWidget,
                               QTableWidget, QTableWidgetItem, QTextEdit,
                               QVBoxLayout, QWidget)

from ..core import ocr as ocr_module
from ..core import templates
from ..core.errors import AppError
from ..core.filler import ImageFillRequest, preview_image
from ..core.models import ExtractionRule, FieldMapping
from ..core.preview import build_image_preview
from ..core.utils import clean_text
from .common import (format_report, pick_files, save_file, show_error, Worker)
from .mapping_widgets import PlanTable, RuleTable
from .preview_dialog import show_preview
from .template_bar import TemplateBar
from .theme import (Card, CollapsibleCard, ghost_button, hint_label,
                    primary_button)

IMAGE_EXTS = ["图片 / PDF (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf)",
              "所有文件 (*.*)"]
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
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(12)

        self.template_bar = TemplateBar("image")
        self.template_bar.save_requested.connect(self._save_template)
        self.template_bar.load_requested.connect(self._load_template)
        root.addWidget(self.template_bar)

        # ---- target + recognition settings (side by side) ----
        top = QHBoxLayout()
        top.setSpacing(12)

        target = Card("目标表格", "选择要写入的 Excel 文件")
        tgrid = target.body
        row1 = QHBoxLayout()
        self.target_edit = QLineEdit()
        self.target_edit.setReadOnly(True)
        self.target_edit.setPlaceholderText("尚未选择目标表格")
        row1.addWidget(self.target_edit, 1)
        row1.addWidget(QPushButton("选择…", self, clicked=lambda: self._pick_target()))
        tgrid.addLayout(row1)
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("工作表"))
        self.sheet_combo = QComboBox()
        row2.addWidget(self.sheet_combo, 2)
        row2.addWidget(QLabel("表头行"))
        self.header_spin = QSpinBox()
        self.header_spin.setRange(0, 200)
        self.header_spin.setValue(1)
        row2.addWidget(self.header_spin)
        row2.addWidget(QLabel("写入位置"))
        self.anchor_edit = QLineEdit("A1")
        self.anchor_edit.setFixedWidth(72)
        row2.addWidget(self.anchor_edit)
        tgrid.addLayout(row2)
        row3 = QHBoxLayout()
        row3.addWidget(QLabel("输出文件"))
        self.output_edit = QLineEdit()
        self.output_edit.setPlaceholderText("默认覆盖目标表格（自动备份）")
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
        top.addWidget(target, 3)

        settings = Card("识别设置")
        sgrid = settings.body
        orow = QHBoxLayout()
        orow.addWidget(QLabel("OCR 引擎"))
        self.engine_combo = QComboBox()
        self.engine_combo.addItems(["自动", "rapidocr", "tesseract", "windows"])
        orow.addWidget(self.engine_combo, 1)
        self.detect_refresh_btn = QPushButton("刷新")
        self.detect_refresh_btn.clicked.connect(self.refresh_engines)
        orow.addWidget(self.detect_refresh_btn)
        sgrid.addLayout(orow)
        mrow = QHBoxLayout()
        mrow.addWidget(QLabel("填写方式"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["表格网格识别（整张图片是一个表格）",
                                  "字段提取（每张图片填一行）"])
        self.mode_combo.currentIndexChanged.connect(self._toggle_mode)
        mrow.addWidget(self.mode_combo, 1)
        sgrid.addLayout(mrow)
        sgrid.addWidget(hint_label("网格识别依赖图片中的表格线；照片建议先裁正。"))
        sgrid.addStretch(1)
        top.addWidget(settings, 2)
        root.addLayout(top)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        # ---- left: images + preview ----
        left = QWidget()
        lgrid = QVBoxLayout(left)
        lgrid.setContentsMargins(0, 0, 0, 0)
        lgrid.setSpacing(12)

        images_card = Card("图片 / PDF 列表")
        images_card.setMinimumHeight(110)
        igrid = images_card.body
        self.image_list = QListWidget()
        self.image_list.setAlternatingRowColors(True)
        igrid.addWidget(self.image_list, 1)
        btns = QHBoxLayout()
        self.add_btn = QPushButton("添加…")
        self.add_btn.clicked.connect(self._add_images)
        self.remove_btn = QPushButton("移除")
        self.remove_btn.clicked.connect(lambda: self._remove_selected())
        self.clear_btn = QPushButton("清空")
        self.clear_btn.clicked.connect(self.image_list.clear)
        self.preview_btn = ghost_button("预览识别")
        self.preview_btn.clicked.connect(self._preview)
        btns.addWidget(self.add_btn)
        btns.addWidget(self.remove_btn)
        btns.addWidget(self.clear_btn)
        btns.addStretch(1)
        btns.addWidget(self.preview_btn)
        igrid.addLayout(btns)
        lgrid.addWidget(images_card, 1)
        splitter.addWidget(left)

        # ---- right: detail tabs (preview / rules / plan) ----
        right = QWidget()
        rgrid = QVBoxLayout(right)
        rgrid.setContentsMargins(0, 0, 0, 0)
        rgrid.setSpacing(0)

        detail_card = Card("识别详情")
        self.detail_tabs = QTabWidget()
        self.detail_tabs.setObjectName("DetailTabs")
        self.detail_tabs.setDocumentMode(True)

        preview_page = QWidget()
        pg = QVBoxLayout(preview_page)
        pg.setContentsMargins(0, 6, 0, 0)
        self.preview_table = QTableWidget(0, 0)
        self.preview_table.verticalHeader().setVisible(False)
        self.preview_table.setAlternatingRowColors(True)
        pg.addWidget(self.preview_table, 1)
        self.detail_tabs.addTab(preview_page, "识别预览")

        rules_page = QWidget()
        rl = QVBoxLayout(rules_page)
        rl.setContentsMargins(0, 6, 0, 0)
        self.rule_title = hint_label("每张图片生成一条记录")
        rl.addWidget(self.rule_title)
        self.rule_table = RuleTable()
        rl.addWidget(self.rule_table, 1)
        rule_btns = QHBoxLayout()
        self.rule_add = QPushButton("添加规则")
        self.rule_add.clicked.connect(self.rule_table.add_row)
        self.rule_del = QPushButton("删除规则")
        self.rule_del.clicked.connect(lambda: self.rule_table.remove_row(
            self.rule_table.currentRow()))
        rule_btns.addWidget(self.rule_add)
        rule_btns.addWidget(self.rule_del)
        rule_btns.addStretch(1)
        rl.addLayout(rule_btns)
        self.detail_tabs.addTab(rules_page, "字段规则")

        plan_page = QWidget()
        pl = QVBoxLayout(plan_page)
        pl.setContentsMargins(0, 6, 0, 0)
        self.plan_table = PlanTable()
        pl.addWidget(self.plan_table, 1)
        plan_btns = QHBoxLayout()
        self.plan_add = QPushButton("添加映射")
        self.plan_add.clicked.connect(self.plan_table.add_row)
        self.plan_del = QPushButton("删除映射")
        self.plan_del.clicked.connect(lambda: self.plan_table.remove_row(
            self.plan_table.currentRow()))
        plan_btns.addWidget(self.plan_add)
        plan_btns.addWidget(self.plan_del)
        plan_btns.addStretch(1)
        pl.addLayout(plan_btns)
        self.detail_tabs.addTab(plan_page, "目标列映射")

        detail_card.body.addWidget(self.detail_tabs, 1)
        rgrid.addWidget(detail_card, 1)
        splitter.addWidget(right)
        splitter.setSizes([520, 640])
        root.addWidget(splitter, 4)

        # ---- run bar ----
        run = QHBoxLayout()
        run.setSpacing(10)
        self.validate_btn = QPushButton("预览 / 校验")
        self.validate_btn.clicked.connect(self._preview_run)
        run.addWidget(self.validate_btn)
        self.run_btn = primary_button("开始批量填表")
        self.run_btn.clicked.connect(self._run)
        run.addWidget(self.run_btn)
        run.addStretch(1)
        self.status_label = QLabel("就绪。")
        self.status_label.setObjectName("StatusText")
        run.addWidget(self.status_label)
        root.addLayout(run)

        self.report_card = CollapsibleCard("执行报告", expanded=False)
        self.report_edit = QTextEdit()
        self.report_edit.setReadOnly(True)
        self.report_edit.setMinimumHeight(90)
        self.report_card.body.addWidget(self.report_edit)
        root.addWidget(self.report_card)

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
        if getattr(self, "detail_tabs", None) is not None:
            self.detail_tabs.setTabEnabled(1, not grid_mode)
            self.detail_tabs.setTabEnabled(2, not grid_mode)
            if grid_mode and self.detail_tabs.currentIndex() in (1, 2):
                self.detail_tabs.setCurrentIndex(0)
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
        eng = OpenpyxlEngine()
        try:
            eng.open(path)
            names = eng.sheets()
        except Exception as exc:
            show_error(self, "无法读取表格", str(exc))
            return
        finally:
            eng.close()
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
    def _build_request(self) -> ImageFillRequest | None:
        images = self._images()
        target = self.target_edit.text().strip()
        if not images:
            show_error(self, "缺少图片", "请先添加至少一张图片或 PDF。")
            return None
        if not target:
            show_error(self, "缺少目标", "请选择目标表格。")
            return None

        rules = [] if self.mode_combo.currentIndex() == 0 else self.rule_table.to_rules()
        plan = [] if self.mode_combo.currentIndex() == 0 else self.plan_table.to_plan()
        if self.mode_combo.currentIndex() == 1 and not plan:
            plan = [FieldMapping(target="A", source="内容")]
        return ImageFillRequest(
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

    def _set_busy(self, busy: bool) -> None:
        self.run_btn.setEnabled(not busy)
        self.validate_btn.setEnabled(not busy)

    def _run(self) -> None:
        request = self._build_request()
        if request is None:
            return
        engine_pref = request.ocr_engine
        self._set_busy(True)
        self.status_label.setText("正在批量识别并填表…")

        def job(progress=lambda *_: None):
            engine = ocr_module.get_ocr_engine(engine_pref)
            from ..core.filler import fill_from_images
            return fill_from_images(request, ocr_engine=engine)

        self.worker = Worker(job, parent=self)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _preview_run(self) -> None:
        request = self._build_request()
        if request is None:
            return
        engine_pref = request.ocr_engine
        self._set_busy(True)
        self.status_label.setText("正在生成预览（OCR）…")

        def job(progress=lambda *_: None):
            engine = ocr_module.get_ocr_engine(engine_pref)
            return build_image_preview(request, ocr_engine=engine)

        self.worker = Worker(job, parent=self)
        self.worker.finished_ok.connect(self._on_preview_result)
        self.worker.failed.connect(self._on_fail)
        self.worker.finished.connect(lambda: self._set_busy(False))
        self.worker.start()

    def _on_preview_result(self, result) -> None:
        self.status_label.setText(
            f"预览完成：预计写入 {result.total_rows} 行，"
            f"{len(result.issues.errors)} 个错误 / {len(result.issues.warnings)} 个提示。")
        if show_preview(self, result):
            self._run()

    def _on_done(self, report) -> None:
        self.status_label.setText(f"完成：写入 {report.written} 处。")
        lines = [format_report(report), ""]
        for record in report.records[:80]:
            lines.append(f"{record.target}  ←  {record.source}  =  {record.value}")
        self.report_edit.setPlainText("\n".join(lines))
        self.report_card.set_expanded(True)

    # -- templates ---------------------------------------------------------
    def _template_data(self) -> dict:
        return {
            "images": self._images(),
            "target_path": self.target_edit.text().strip(),
            "output_path": self.output_edit.text().strip(),
            "sheet": self.sheet_combo.currentText(),
            "header_row": self.header_spin.value(),
            "anchor": self.anchor_edit.text().strip(),
            "append": self.append_box.isChecked(),
            "clear_sheet": self.clear_box.isChecked(),
            "mode": self.mode_combo.currentIndex(),
            "ocr_engine": self.engine_combo.currentText(),
            "extraction_rules": [rule.to_dict() for rule in self.rule_table.to_rules()],
            "column_plan": [plan.to_dict() for plan in self.plan_table.to_plan()],
        }

    def _apply_template_data(self, data: dict) -> None:
        self.image_list.clear()
        for path in data.get("images") or []:
            self.image_list.addItem(QListWidgetItem(path))
        target = data.get("target_path") or ""
        if target:
            self.target_edit.setText(target)
            self._load_sheets(target)
        self.output_edit.setText(data.get("output_path") or "")
        sheet = data.get("sheet") or ""
        if sheet:
            index = self.sheet_combo.findText(sheet)
            if index >= 0:
                self.sheet_combo.setCurrentIndex(index)
        self.header_spin.setValue(int(data.get("header_row", 1) or 1))
        self.anchor_edit.setText(data.get("anchor") or "A1")
        self.append_box.setChecked(bool(data.get("append", True)))
        self.clear_box.setChecked(bool(data.get("clear_sheet", False)))
        self.mode_combo.setCurrentIndex(int(data.get("mode", 0) or 0))
        engine = data.get("ocr_engine") or "自动"
        index = self.engine_combo.findText(engine)
        if index >= 0:
            self.engine_combo.setCurrentIndex(index)
        self.rule_table.load(
            [ExtractionRule.from_dict(item) for item in data.get("extraction_rules", [])])
        self.plan_table.load(
            [FieldMapping.from_dict(item) for item in data.get("column_plan", [])])
        self._toggle_mode()

    def _save_template(self, name: str) -> None:
        try:
            templates.save_template(name, "image", self._template_data())
        except AppError as exc:
            show_error(self, "保存模板失败", str(exc))
            return
        self.template_bar.refresh()
        self.template_bar.select(name)
        self.status_label.setText(f"模板已保存：{name}")

    def _load_template(self, name: str) -> None:
        try:
            template = templates.load_template(name, "image")
        except AppError as exc:
            show_error(self, "读取模板失败", str(exc))
            return
        self._apply_template_data(template.data)
        self.status_label.setText(f"已应用模板：{name}")

    def on_before_close(self) -> bool:
        if self.worker is not None and self.worker.isRunning():
            self.worker.wait(3000)
        return True
