"""Side panel: a frameless assistant docked to the right edge of the screen.

The panel is a small companion window that stays on top of Excel / WPS while
the user fills a form by hand:

* it is attached to the right edge of the current screen;
* it can collapse into a slim gradient rail and expand again;
* it holds a list of ``field name + value`` rows, each value can be copied
  individually, or the whole row can be copied as Tab separated text so it can
  be pasted straight into a spreadsheet row;
* fields can be generated from the headers of the table currently selected in
  the main window.

Everything (fields, pinned state, vertical position) is stored in QSettings.
"""

from __future__ import annotations

import json
import os
from typing import Callable

from PySide6.QtCore import (QEasingCurve, QPointF, QPropertyAnimation, QRect,
                            QRectF, Qt, QTimer, Signal)
from PySide6.QtGui import (QColor, QFont, QGuiApplication, QLinearGradient,
                           QPainter, QPainterPath)
from PySide6.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel,
                               QLayout, QLineEdit, QPushButton, QScrollArea,
                               QSizePolicy, QToolButton, QVBoxLayout, QWidget)

from .common import app_settings, confirm
from .theme import ACCENT, primary_button, set_role

DEFAULT_FIELDS = ["姓名", "联系电话", "部门"]


# ---------------------------------------------------------------------------
# Small pieces
# ---------------------------------------------------------------------------

class _FieldRow(QWidget):
    """One editable ``field name + value + copy`` row."""

    copied = Signal(str)
    changed = Signal()
    removed = Signal(object)

    def __init__(self, label: str = "", value: str = "",
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("FieldRow")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.name_edit = QLineEdit(label, self)
        self.name_edit.setObjectName("FieldName")
        self.name_edit.setPlaceholderText("字段名")
        self.name_edit.setFixedWidth(84)
        self.name_edit.setToolTip("字段名，可双击修改")

        self.value_edit = QLineEdit(value, self)
        self.value_edit.setObjectName("FieldValue")
        self.value_edit.setPlaceholderText("填写内容，回车复制")
        self.value_edit.setClearButtonEnabled(True)

        self.copy_btn = QToolButton(self)
        self.copy_btn.setObjectName("RowCopyButton")
        self.copy_btn.setText("复制")
        self.copy_btn.setCursor(Qt.PointingHandCursor)
        self.copy_btn.setToolTip("复制该字段的值")

        self.del_btn = QToolButton(self)
        self.del_btn.setObjectName("RowDeleteButton")
        self.del_btn.setText("✕")
        self.del_btn.setCursor(Qt.PointingHandCursor)
        self.del_btn.setToolTip("删除该字段")

        layout.addWidget(self.name_edit)
        layout.addWidget(self.value_edit, 1)
        layout.addWidget(self.copy_btn)
        layout.addWidget(self.del_btn)

        # Keep long values readable from the start instead of scrolled to the end.
        self.name_edit.setCursorPosition(0)
        self.value_edit.setCursorPosition(0)

        self.name_edit.textChanged.connect(self.changed)
        self.value_edit.textChanged.connect(self.changed)
        self.value_edit.returnPressed.connect(self._copy)
        self.copy_btn.clicked.connect(self._copy)
        self.del_btn.clicked.connect(lambda: self.removed.emit(self))

    def _copy(self) -> None:
        text = self.value_edit.text().strip()
        if not text:
            return
        QApplication.clipboard().setText(text)
        self.copied.emit(text)

    def data(self) -> tuple[str, str]:
        return self.name_edit.text().strip(), self.value_edit.text()


class _Rail(QWidget):
    """The slim gradient tab shown while the panel is collapsed."""

    clicked = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip("展开填表助手")
        self.setAttribute(Qt.WA_Hover, True)
        self._hover = False

    def enterEvent(self, event) -> None:  # noqa: N802 - Qt naming
        self._hover = True
        self.update()
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: N802
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)

    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        rect = QRectF(self.rect()).adjusted(0.5, 0.5, -0.5, -0.5)
        gradient = QLinearGradient(rect.topLeft(), rect.bottomLeft())
        if self._hover:
            gradient.setColorAt(0.0, QColor("#4338CA"))
            gradient.setColorAt(1.0, QColor("#6D28D9"))
        else:
            gradient.setColorAt(0.0, QColor(ACCENT))
            gradient.setColorAt(1.0, QColor("#7C3AED"))
        path = QPainterPath()
        path.addRoundedRect(rect, 15, 15)
        painter.fillPath(path, gradient)

        painter.setPen(QColor(255, 255, 255, 235))
        font = QFont(self.font())
        font.setPointSize(9)
        font.setBold(True)
        painter.setFont(font)
        painter.save()
        painter.translate(self.width() / 2.0, self.height() / 2.0)
        painter.rotate(90)
        painter.drawText(QRectF(-60, -10, 120, 20), Qt.AlignCenter, "填表助手")
        painter.restore()

        # Chevron pill at the bottom, a hint that the rail expands.
        painter.setPen(QColor(255, 255, 255, 200))
        chevron = QFont(self.font())
        chevron.setPointSize(11)
        chevron.setBold(True)
        painter.setFont(chevron)
        painter.drawText(QRectF(0, self.height() - 34, self.width(), 22),
                         Qt.AlignCenter, "‹")
        painter.end()


class _DragHeader(QFrame):
    """Gradient header that can be dragged vertically."""

    drag_moved = Signal(int)
    drag_finished = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("PanelHeader")
        self.setFixedHeight(70)
        self.setCursor(Qt.OpenHandCursor)
        self._press: QPointF | None = None

    def mousePressEvent(self, event) -> None:  # noqa: N802
        if event.button() == Qt.LeftButton:
            self._press = event.globalPosition()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self._press is not None:
            pos = event.globalPosition()
            delta = pos.y() - self._press.y()
            if delta:
                self.drag_moved.emit(int(delta))
                self._press = pos
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if self._press is not None:
            self._press = None
            self.setCursor(Qt.OpenHandCursor)
            self.drag_finished.emit()
            event.accept()
            return
        super().mouseReleaseEvent(event)


# ---------------------------------------------------------------------------
# Panel
# ---------------------------------------------------------------------------

class SidePanel(QWidget):
    """Frameless, always-on-top assistant attached to the screen's right edge."""

    PANEL_WIDTH = 348
    RAIL_WIDTH = 38
    SHADOW_MARGIN = 14  # transparent left + right margins used by the shadow

    closed = Signal()
    collapsed_changed = Signal(bool)

    def __init__(self, anchor: QWidget | None = None) -> None:
        super().__init__(anchor)
        self._anchor = anchor
        self._collapsed = False
        self._top: int | None = self._load_top()
        self._always_on_top = bool(
            app_settings().value("side_panel/pinned", True, type=bool))
        self.headers_provider: Callable[[], list[str]] | None = None
        self.target_provider: Callable[[], dict | None] | None = None

        flags = Qt.Tool | Qt.FramelessWindowHint
        if self._always_on_top:
            flags |= Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setWindowTitle("填表助手")

        self._build()

        self._anim = QPropertyAnimation(self, b"geometry", self)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)
        self._anim.finished.connect(self._on_anim_finished)
        self._anim_done: Callable[[], None] | None = None

        self._status_timer = QTimer(self)
        self._status_timer.setSingleShot(True)
        self._status_timer.timeout.connect(self._reset_status)

        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.timeout.connect(self._save_fields)

        self._load_fields()

    # -- construction ------------------------------------------------------
    def _build(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 12, 2, 12)
        outer.setSpacing(0)
        # The window shrinks below its content width while collapsing, so the
        # layout must not impose a minimum width on the frameless window.
        outer.setSizeConstraint(QLayout.SetNoConstraint)

        self._card = QFrame()
        self._card.setObjectName("SidePanelCard")
        outer.addWidget(self._card)

        card_layout = QHBoxLayout(self._card)
        card_layout.setContentsMargins(0, 0, 0, 0)
        card_layout.setSpacing(0)

        self._rail = _Rail(self._card)
        self._rail.setFixedWidth(self.RAIL_WIDTH)
        self._rail.clicked.connect(lambda: self.set_collapsed(False))
        self._rail.setVisible(False)
        card_layout.addWidget(self._rail)

        self._content = QWidget(self._card)
        card_layout.addWidget(self._content, 1)
        content_layout = QVBoxLayout(self._content)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self._header = self._build_header()
        content_layout.addWidget(self._header)

        self._body = self._build_body()
        content_layout.addWidget(self._body, 1)

        content_layout.addWidget(self._build_footer())

    def _build_header(self) -> QWidget:
        header = _DragHeader()
        header.drag_moved.connect(self._on_drag)
        header.drag_finished.connect(self._save_top)

        layout = QHBoxLayout(header)
        layout.setContentsMargins(14, 12, 10, 12)
        layout.setSpacing(8)

        logo = QLabel("表", header)
        logo.setObjectName("PanelLogo")
        logo.setFixedSize(32, 32)
        logo.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo)

        titles = QVBoxLayout()
        titles.setContentsMargins(0, 0, 0, 0)
        titles.setSpacing(1)
        title = QLabel("填表助手", header)
        title.setObjectName("PanelTitle")
        subtitle = QLabel("侧边栏 · 识别 / 复制 / 直填", header)
        subtitle.setObjectName("PanelSubtitle")
        titles.addWidget(title)
        titles.addWidget(subtitle)
        layout.addLayout(titles)
        layout.addStretch(1)

        self._pin_btn = QToolButton(header)
        self._pin_btn.setObjectName("PanelHeaderButton")
        self._pin_btn.setCheckable(True)
        self._pin_btn.setChecked(self._always_on_top)
        self._pin_btn.setText("已置顶" if self._always_on_top else "置顶")
        self._pin_btn.setCursor(Qt.PointingHandCursor)
        self._pin_btn.setToolTip("切换窗口置顶")
        self._pin_btn.toggled.connect(self._on_pin_toggled)
        layout.addWidget(self._pin_btn)

        collapse = QToolButton(header)
        collapse.setObjectName("PanelHeaderButton")
        collapse.setText("›")
        collapse.setCursor(Qt.PointingHandCursor)
        collapse.setToolTip("收起为屏幕右缘窄条（Esc）")
        collapse.clicked.connect(lambda: self.set_collapsed(True))
        layout.addWidget(collapse)

        close = QToolButton(header)
        close.setObjectName("PanelHeaderButton")
        close.setText("✕")
        close.setCursor(Qt.PointingHandCursor)
        close.setToolTip("关闭侧边栏")
        close.clicked.connect(self.hide_panel)
        layout.addWidget(close)
        return header

    def _build_body(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setObjectName("PanelScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.viewport().setAutoFillBackground(False)

        body = QWidget()
        body.setObjectName("PanelBody")
        layout = QVBoxLayout(body)
        layout.setContentsMargins(14, 12, 14, 14)
        layout.setSpacing(10)

        card = QFrame()
        card.setObjectName("PanelCard")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(8)

        head = QHBoxLayout()
        head.setContentsMargins(0, 0, 0, 0)
        head.setSpacing(6)
        title = QLabel("字段速填", card)
        title.setObjectName("PanelCardTitle")
        head.addWidget(title)
        head.addStretch(1)
        saved = QLabel("自动保存", card)
        saved.setObjectName("PanelHint")
        head.addWidget(saved)
        card_layout.addLayout(head)

        self.detect_btn = QToolButton(card)
        self.detect_btn.setObjectName("PanelDetectButton")
        self.detect_btn.setText("⚡ 一键识别剪贴板")
        self.detect_btn.setToolTip("读取剪贴板文字，自动分类为字段并匹配目标表头")
        self.detect_btn.setCursor(Qt.PointingHandCursor)
        self.detect_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.detect_btn.clicked.connect(self.detect_clipboard)
        card_layout.addWidget(self.detect_btn)

        self._empty_hint = QLabel(
            "还没有字段。点「⚡ 一键识别剪贴板」，或读取主界面目标表格的表头。", card)
        self._empty_hint.setObjectName("PanelHint")
        self._empty_hint.setWordWrap(True)
        card_layout.addWidget(self._empty_hint)

        self._rows_layout = QVBoxLayout()
        self._rows_layout.setContentsMargins(0, 0, 0, 0)
        self._rows_layout.setSpacing(6)
        card_layout.addLayout(self._rows_layout)

        actions = QHBoxLayout()
        actions.setContentsMargins(0, 2, 0, 0)
        actions.setSpacing(2)
        add_btn = self._flat_button("＋ 添加字段", "在末尾添加一个字段")
        add_btn.clicked.connect(lambda: self.add_field(focus=True))
        read_btn = self._flat_button("读取表头", "读取主界面当前目标表格的表头")
        read_btn.clicked.connect(self._read_headers)
        clear_btn = self._flat_button("清空", "删除全部字段", danger=True)
        clear_btn.clicked.connect(self._clear_fields)
        actions.addWidget(add_btn)
        actions.addWidget(read_btn)
        actions.addStretch(1)
        actions.addWidget(clear_btn)
        card_layout.addLayout(actions)

        card_layout.addStretch(1)
        hint = QLabel("按列顺序排列；「复制整行」用 Tab 分隔，可直接粘贴到表格。", card)
        hint.setObjectName("PanelHint")
        hint.setWordWrap(True)
        card_layout.addWidget(hint)

        layout.addWidget(card, 1)

        tips = QFrame()
        tips.setObjectName("PanelCard")
        tips_layout = QVBoxLayout(tips)
        tips_layout.setContentsMargins(12, 10, 12, 10)
        tips_layout.setSpacing(5)
        tips_title = QLabel("使用提示", tips)
        tips_title.setObjectName("PanelCardTitle")
        tips_layout.addWidget(tips_title)
        for text in (
            "复制一段包含姓名、电话等内容的文字，点「⚡ 一键识别剪贴板」。",
            "识别结果会按目标表头对齐；点「直接填入目标表格」追加新行。",
            "也可以逐项「复制」或「复制整行」，手动粘贴到 Excel / WPS。",
        ):
            tip = QLabel(f"· {text}", tips)
            tip.setObjectName("PanelHint")
            tip.setWordWrap(True)
            tips_layout.addWidget(tip)
        layout.addWidget(tips)
        scroll.setWidget(body)
        return scroll

    def _build_footer(self) -> QWidget:
        footer = QFrame()
        footer.setObjectName("PanelFooter")
        layout = QVBoxLayout(footer)
        layout.setContentsMargins(14, 10, 14, 12)
        layout.setSpacing(6)

        self.status_label = QLabel("就绪。识别剪贴板或读取表头后，可直接填入目标表格。", footer)
        self.status_label.setObjectName("PanelStatus")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        self.copy_all_btn = primary_button("复制整行（Tab 分隔）", footer)
        self.copy_all_btn.clicked.connect(self.copy_row)
        layout.addWidget(self.copy_all_btn)

        self.fill_btn = QPushButton("直接填入目标表格", footer)
        self.fill_btn.setCursor(Qt.PointingHandCursor)
        self.fill_btn.setToolTip("把当前字段作为新行追加到主界面选中的表格（自动备份）")
        set_role(self.fill_btn, "outline")
        self.fill_btn.clicked.connect(self.fill_target)
        layout.addWidget(self.fill_btn)
        return footer

    @staticmethod
    def _flat_button(text: str, tooltip: str, danger: bool = False) -> QToolButton:
        button = QToolButton()
        button.setObjectName("PanelGhostDanger" if danger else "PanelGhostButton")
        button.setText(text)
        button.setToolTip(tooltip)
        button.setCursor(Qt.PointingHandCursor)
        return button

    # -- fields ------------------------------------------------------------
    def add_field(self, label: str = "", value: str = "",
                  focus: bool = False, save: bool = True) -> _FieldRow:
        row = _FieldRow(label, value)
        row.changed.connect(self._schedule_save)
        row.copied.connect(self._on_field_copied)
        row.removed.connect(self._remove_row)
        self._rows.append(row)
        self._rows_layout.addWidget(row)
        self._empty_hint.setVisible(False)
        if save:
            self._schedule_save()
        if focus:
            row.value_edit.setFocus()
        return row

    def fields(self) -> list[tuple[str, str]]:
        return [row.data() for row in self._rows]

    def set_fields_from_headers(self, headers: list[str]) -> None:
        """Replace the field names, keeping values of rows with the same name."""
        existing = {label: value for label, value in self.fields()}
        self._clear_rows()
        for header in headers:
            header = str(header).strip()
            if header:
                self.add_field(header, existing.get(header, ""), save=False)
        self._empty_hint.setVisible(not self._rows)
        self._schedule_save()

    def copy_row(self) -> None:
        values = [value.strip() for _, value in self.fields() if value.strip()]
        if not values:
            self._set_status("还没有可复制的内容，请先填写字段值。")
            return
        QApplication.clipboard().setText("\t".join(values))
        self._set_status(f"已复制整行（{len(values)} 个字段），可直接粘贴到表格行。")

    def _read_headers(self) -> None:
        headers = self._headers()
        if not headers:
            self._set_status("未找到目标表格，请先在主界面选择目标表格。")
            return
        self.set_fields_from_headers(headers)
        self._set_status(f"已读取 {len(headers)} 个表头字段，可逐项填写。")

    def _headers(self) -> list[str]:
        if self.headers_provider is None:
            return []
        try:
            return [str(item).strip()
                    for item in self.headers_provider() if str(item).strip()]
        except Exception:  # noqa: BLE001 - a helper must never crash the panel
            return []

    # -- clipboard detection ----------------------------------------------
    def detect_clipboard(self) -> None:
        """Read the clipboard, classify it and align it with the target headers."""
        from ..core.field_detect import detect_fields, map_to_headers

        text = QApplication.clipboard().text()
        if not text or not text.strip():
            self._set_status("剪贴板是空的，先复制一段文字再试。")
            return
        detected = detect_fields(text)
        if not detected:
            self._set_status("没有识别到可分类的内容，可手动填写或读取表头。")
            return

        headers = self._headers()
        if not headers:
            self._clear_rows()
            for field in detected:
                self.add_field(field.label, field.value, save=False)
            self._empty_hint.setVisible(False)
            self._schedule_save()
            self._set_status(
                f"已识别 {len(detected)} 个字段（未找到目标表头，已按识别结果显示）。")
            return

        matched, unmatched = map_to_headers(detected, headers)
        existing = {label: value for label, value in self.fields() if value.strip()}
        self._clear_rows()
        used_labels: set[str] = set()
        for header in headers:
            header = str(header).strip()
            if not header or header in used_labels:
                continue
            used_labels.add(header)
            value = matched.get(header, existing.get(header, ""))
            self.add_field(header, value, save=False)
        for field in unmatched:
            if field.label in used_labels:
                continue
            used_labels.add(field.label)
            self.add_field(field.label, field.value, save=False)
        self._empty_hint.setVisible(not self._rows)
        self._schedule_save()
        self._set_status(
            f"已识别 {len(detected)} 个字段，命中 {len(matched)}/{len(headers)} 个表头；"
            "点「直接填入目标表格」追加新行。")

    def fill_target(self) -> None:
        """Append the current fields as a new row to the selected workbook."""
        from ..core.errors import FillError
        from ..core.field_detect import append_fields

        info = None
        if self.target_provider is not None:
            try:
                info = self.target_provider()
            except Exception:  # noqa: BLE001 - a helper must never crash the panel
                info = None
        path = str((info or {}).get("path", "") or "").strip()
        if not path:
            self._set_status("请先在主界面选择目标表格，再点「直接填入」。")
            return
        values = {label: value for label, value in self.fields() if value.strip()}
        if not values:
            self._set_status("没有可填入的内容，请先识别剪贴板或填写字段值。")
            return

        sheet = str(info.get("sheet", "") or "")
        header_row = max(1, int(info.get("header_row", 1) or 1))
        output = str(info.get("output_path", "") or "").strip()
        target = output or path
        if not confirm(
            self, "直接填入表格",
            f"将把 {len(values)} 个字段作为新行追加到：\n{target}\n"
            f"工作表：{sheet or '当前工作表'}（表头第 {header_row} 行）\n\n"
            "写入前会自动备份原文件。继续吗？",
        ):
            return
        try:
            report = append_fields(path, values, sheet=sheet,
                                   header_row=header_row, output_path=output)
        except FillError as exc:
            self._set_status(str(exc))
            return
        except Exception as exc:  # noqa: BLE001 - surface IO errors in the panel
            self._set_status(f"写入失败：{exc}")
            return
        self._set_status(
            f"已写入第 {report.row} 行（{report.matched} 个字段）→ "
            f"{os.path.basename(report.path)}")

    def _clear_fields(self) -> None:
        self._clear_rows()
        self._empty_hint.setVisible(True)
        self._schedule_save()
        self._set_status("已清空字段。")

    def _clear_rows(self) -> None:
        while self._rows:
            row = self._rows.pop()
            self._rows_layout.removeWidget(row)
            row.setParent(None)
            row.deleteLater()

    def _remove_row(self, row: _FieldRow) -> None:
        if row in self._rows:
            self._rows.remove(row)
            self._rows_layout.removeWidget(row)
            row.setParent(None)
            row.deleteLater()
        self._empty_hint.setVisible(not self._rows)
        self._schedule_save()

    def _on_field_copied(self, text: str) -> None:
        shown = text if len(text) <= 24 else text[:24] + "…"
        self._set_status(f"已复制「{shown}」到剪贴板。")

    def _schedule_save(self) -> None:
        self._save_timer.start(400)

    def _load_fields(self) -> None:
        self._rows: list[_FieldRow] = []
        raw = app_settings().value("side_panel/fields", "")
        items: list[list[str]] | None = None
        if raw:
            try:
                parsed = json.loads(raw)
                if isinstance(parsed, list):
                    items = [[str(pair[0]), str(pair[1])]
                             for pair in parsed
                             if isinstance(pair, (list, tuple)) and len(pair) >= 2]
            except (TypeError, ValueError):
                items = None
        if items is None:
            items = [[name, ""] for name in DEFAULT_FIELDS]
        for label, value in items:
            self.add_field(label, value, save=False)
        self._empty_hint.setVisible(not self._rows)

    def _save_fields(self) -> None:
        app_settings().setValue("side_panel/fields",
                                json.dumps(self.fields(), ensure_ascii=False))

    # -- status ------------------------------------------------------------
    def _set_status(self, text: str) -> None:
        self.status_label.setText(text)
        self._status_timer.start(5000)

    def _reset_status(self) -> None:
        self.status_label.setText("就绪。识别剪贴板或读取表头后，可直接填入目标表格。")

    # -- geometry ----------------------------------------------------------
    def _available_geometry(self) -> QRect:
        screen = None
        if self._anchor is not None:
            try:
                screen = self._anchor.screen()
            except AttributeError:
                screen = None
        screen = screen or self.screen() or QGuiApplication.primaryScreen()
        return screen.availableGeometry()

    def _rect_for_width(self, width: int, top: int | None = None) -> QRect:
        available = self._available_geometry()
        height = max(460, min(int(available.height() * 0.86), 960))
        height = min(height, max(320, available.height() - 24))
        if top is None:
            top = self._top if self._top is not None else (
                available.y() + (available.height() - height) // 2)
        top = max(available.y(),
                  min(top, available.y() + available.height() - height))
        x = available.x() + available.width() - width
        return QRect(x, top, width, height)

    def _on_drag(self, delta: int) -> None:
        new_top = self.y() + delta
        self._top = new_top
        self.setGeometry(self._rect_for_width(self.width(), top=new_top))

    def _load_top(self) -> int | None:
        value = app_settings().value("side_panel/top", "")
        try:
            return int(value) if value not in ("", None) else None
        except (TypeError, ValueError):
            return None

    def _save_top(self) -> None:
        if self._top is not None:
            app_settings().setValue("side_panel/top", int(self._top))

    # -- show / hide / collapse -------------------------------------------
    def show_panel(self, animated: bool = True) -> None:
        if self._collapsed:
            self._collapsed = False
            self._content.setVisible(True)
            self._rail.setVisible(False)
            self.collapsed_changed.emit(False)

        end = self._rect_for_width(self.PANEL_WIDTH)
        if self.isVisible():
            if animated:
                self._animate_to(end)
            else:
                self.setGeometry(end)
            self.raise_()
            self.activateWindow()
            return

        if animated:
            start = QRect(end.x() + 48, end.y(), end.width(), end.height())
            self.setGeometry(start)
            self.show()
            self.raise_()
            self._animate_to(end, duration=260)
        else:
            self.setGeometry(end)
            self.show()
            self.raise_()
        self.activateWindow()

    def hide_panel(self) -> None:
        self._anim.stop()
        self._anim_done = None
        self._save_timer.stop()
        self._save_fields()
        self.hide()
        self.closed.emit()

    def toggle_panel(self) -> None:
        if self.isVisible():
            self.hide_panel()
        else:
            self.show_panel()

    def set_collapsed(self, collapsed: bool, animated: bool = True) -> None:
        if self._collapsed == collapsed:
            return
        self._collapsed = collapsed
        width = (self.RAIL_WIDTH + self.SHADOW_MARGIN if collapsed
                 else self.PANEL_WIDTH)
        target = self._rect_for_width(width)

        if not self.isVisible() or not animated:
            self._anim.stop()
            self.setGeometry(target)
            self._content.setVisible(not collapsed)
            self._rail.setVisible(collapsed)
        elif collapsed:
            self._animate_to(target, duration=230, on_done=self._finish_collapse)
        else:
            self._content.setVisible(True)
            self._rail.setVisible(False)
            self._animate_to(target, duration=250)
        self.collapsed_changed.emit(collapsed)

    def is_collapsed(self) -> bool:
        return self._collapsed

    def _finish_collapse(self) -> None:
        if not self._collapsed:
            return
        self._content.setVisible(False)
        self._rail.setVisible(True)

    def _animate_to(self, target: QRect, duration: int = 230,
                    on_done: Callable[[], None] | None = None) -> None:
        self._anim.stop()
        self._anim_done = on_done
        self._anim.setDuration(duration)
        self._anim.setStartValue(self.geometry())
        self._anim.setEndValue(target)
        self._anim.start()

    def _on_anim_finished(self) -> None:
        callback = self._anim_done
        self._anim_done = None
        if callback is not None:
            callback()

    # -- window behaviour --------------------------------------------------
    def _on_pin_toggled(self, checked: bool) -> None:
        self._always_on_top = checked
        self._pin_btn.setText("已置顶" if checked else "置顶")
        app_settings().setValue("side_panel/pinned", checked)
        visible = self.isVisible()
        geometry = self.geometry()
        self.setWindowFlag(Qt.WindowStaysOnTopHint, checked)
        self.setGeometry(geometry)
        if visible:
            self.show()
            self.raise_()

    def paintEvent(self, event) -> None:  # noqa: N802
        """Draw the soft shadow behind the rounded card.

        A ``QGraphicsDropShadowEffect`` distorts the child layout when the
        frameless translucent window is grabbed/rendered, so the shadow is
        painted manually instead.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(Qt.NoPen)
        base = QRectF(self._card.geometry())
        for spread, alpha in ((8, 10), (5, 14), (2, 20)):
            rect = base.adjusted(-spread, -spread + 3, spread, spread + 3)
            path = QPainterPath()
            path.addRoundedRect(rect, 16 + spread, 16 + spread)
            painter.fillPath(path, QColor(15, 23, 42, alpha))
        painter.end()

    def keyPressEvent(self, event) -> None:  # noqa: N802
        if event.key() == Qt.Key_Escape and not self._collapsed:
            self.set_collapsed(True)
            event.accept()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event) -> None:  # noqa: N802
        event.ignore()
        self.hide_panel()
