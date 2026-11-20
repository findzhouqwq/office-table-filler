"""Full-screen translucent overlay used to select spreadsheet blocks.

The user long-presses the right mouse button on top of Excel / WPS, drags a
green rubber band and releases to add one block.  Several blocks can be
selected in one go; ``Enter`` reads them through COM and emits
:attr:`SelectionOverlay.finished`.

The overlay itself never touches COM: it stores both the Qt (logical) rectangle
for painting and the physical screen rectangle that ``RangeFromPoint`` needs.
"""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import QPoint, QRect, QRectF, Qt, QTimer, Signal
from PySide6.QtGui import (QColor, QFont, QGuiApplication, QPainter, QPen)
from PySide6.QtWidgets import QApplication, QWidget

from ..core.screen_copy import CopyBlock, ScreenCopyError, read_screen_rect

GREEN = "#22C55E"
GREEN_DARK = "#16A34A"

HOLD_MS = 380
MIN_DRAG = 6


def _physical_cursor() -> tuple[int, int] | None:
    """Current cursor position in physical screen pixels (Windows)."""
    try:
        import win32api  # type: ignore

        x, y = win32api.GetCursorPos()
        return int(x), int(y)
    except Exception:
        pass
    try:
        import ctypes

        class _Point(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        point = _Point()
        if ctypes.windll.user32.GetCursorPos(ctypes.byref(point)):
            return int(point.x), int(point.y)
    except Exception:
        pass
    return None


@dataclass
class _Region:
    """One finished selection: logical rectangle + physical screen rectangle."""

    widget_rect: QRect
    screen_rect: tuple[int, int, int, int]


class SelectionOverlay(QWidget):
    """Translucent, always-on-top layer with a green rubber band selector."""

    finished = Signal(list, list)  # (list[CopyBlock], list[str] errors)
    cancelled = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        flags = (Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint
                 | Qt.NoDropShadowWindowHint)
        self.setWindowFlags(flags)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setCursor(Qt.CrossCursor)
        self.setWindowTitle("跨表框选")

        # Injected so tests can replace the COM reader.
        self.reader = read_screen_rect

        self._regions: list[_Region] = []
        self._press_widget: QPoint | None = None
        self._press_screen: tuple[int, int] | None = None
        self._current: QRect | None = None
        self._dragging = False
        self._done = False
        self._hint = ""

        self._hold_timer = QTimer(self)
        self._hold_timer.setSingleShot(True)
        self._hold_timer.setInterval(HOLD_MS)
        self._hold_timer.timeout.connect(self._start_drag)

        # Cached paint resources: rebuilding QColor/QPen/QFont for every mouse
        # move is measurable on low-end machines with software rendering.
        self._veil = QColor(2, 6, 23, 36)
        self._region_fill = QColor(34, 197, 94, 58)
        self._green = QColor(GREEN)
        self._green_dark = QColor(GREEN_DARK)
        self._white = QColor("#FFFFFF")
        self._hud_bg = QColor(15, 23, 42, 225)
        self._hud_text = QColor(226, 232, 240, 205)
        self._hud_accent = QColor("#86EFAC")
        self._pen_solid = QPen(self._green, 2)
        self._pen_dashed = QPen(self._green, 2)
        self._pen_dashed.setStyle(Qt.DashLine)
        self._badge_font = QFont(self.font())
        self._badge_font.setPointSize(9)
        self._badge_font.setBold(True)
        self._title_font = QFont(self.font())
        self._title_font.setPointSize(12)
        self._title_font.setBold(True)
        self._hud_font = QFont(self.font())
        self._hud_font.setPointSize(10)

    # -- lifecycle ---------------------------------------------------------
    def start(self) -> None:
        """Show the overlay over the whole virtual desktop."""
        self._regions.clear()
        self._press_widget = None
        self._press_screen = None
        self._current = None
        self._dragging = False
        self._done = False
        self._hint = ""
        self.setGeometry(self._virtual_geometry())
        self.show()
        self.raise_()
        self.activateWindow()
        self.setFocus()
        self.update()

    def _virtual_geometry(self) -> QRect:
        rect = QRect()
        for screen in QGuiApplication.screens():
            rect = rect.united(screen.geometry())
        if not rect.isValid():
            primary = QGuiApplication.primaryScreen()
            rect = primary.geometry() if primary is not None else QRect(0, 0, 1920, 1080)
        return rect

    def _hud_rect(self) -> QRect:
        width = min(660, max(360, self.width() - 48))
        return QRect(int((self.width() - width) / 2) - 2, 22, width + 4, 100)

    def _update_hud(self) -> None:
        self.update(self._hud_rect())

    @staticmethod
    def _inflate(rect: QRect) -> QRect:
        return rect.adjusted(-3, -3, 3, 3)

    # -- physical coordinates ---------------------------------------------
    def _to_physical(self, point: QPoint) -> tuple[int, int]:
        physical = _physical_cursor()
        if physical is not None:
            return physical
        # Fallback for environments without GetCursorPos (tests, non-Windows).
        ratio = self.devicePixelRatioF() or 1.0
        global_point = self.mapToGlobal(point)
        return (round(global_point.x() * ratio), round(global_point.y() * ratio))

    # -- mouse -------------------------------------------------------------
    def mousePressEvent(self, event) -> None:  # noqa: N802 - Qt naming
        if event.button() != Qt.RightButton:
            event.ignore()
            return
        self._press_widget = event.position().toPoint()
        self._press_screen = self._to_physical(self._press_widget)
        self._dragging = False
        self._current = None
        self._hint = "按住右键稍候，拖动即可框选"
        self._hold_timer.start()
        self._update_hud()
        event.accept()

    def mouseMoveEvent(self, event) -> None:  # noqa: N802
        if self._press_widget is None:
            super().mouseMoveEvent(event)
            return
        pos = event.position().toPoint()
        if not self._dragging:
            if (pos - self._press_widget).manhattanLength() < MIN_DRAG:
                event.accept()
                return
            # The user pressed and dragged right away: do not force them to
            # wait for the hold timer.
            self._start_drag()
        old = self._current
        self._current = QRect(self._press_widget, pos).normalized()
        # Repaint only the old + new rubber band instead of the whole screen;
        # this keeps dragging smooth on 4K / low-end machines.
        dirty = self._inflate(self._current)
        if old is not None:
            dirty = dirty.united(self._inflate(old))
        self.update(dirty)
        event.accept()

    def mouseReleaseEvent(self, event) -> None:  # noqa: N802
        if event.button() != Qt.RightButton:
            super().mouseReleaseEvent(event)
            return
        self._hold_timer.stop()
        dirty = self._hud_rect()
        if self._current is not None:
            dirty = dirty.united(self._inflate(self._current))
        if self._dragging and self._press_widget is not None:
            pos = event.position().toPoint()
            rect = QRect(self._press_widget, pos).normalized()
            dirty = dirty.united(self._inflate(rect))
            right = self._to_physical(pos)
            if (rect.width() >= MIN_DRAG and rect.height() >= MIN_DRAG
                    and self._press_screen is not None):
                x1, y1 = self._press_screen
                x2, y2 = right
                screen_rect = (min(x1, x2), min(y1, y2),
                               max(x1, x2), max(y1, y2))
                self._regions.append(_Region(rect, screen_rect))
                self._hint = "已选好一块，可继续长按右键框选；Enter 完成"
            else:
                self._hint = "选区太小了，请拖出至少一个单元格"
        else:
            self._hint = "长按右键约 0.4 秒后再拖动，松开完成框选"
        self._dragging = False
        self._current = None
        self._press_widget = None
        self._press_screen = None
        self.update(dirty)
        event.accept()

    def _start_drag(self) -> None:
        self._hold_timer.stop()
        if self._press_widget is None:
            return
        self._dragging = True
        self._current = QRect(self._press_widget, self._press_widget)
        self._hint = "拖动到选区右下角，松开完成这一块"
        self.update(self._hud_rect().united(self._inflate(self._current)))

    # -- keyboard ----------------------------------------------------------
    def keyPressEvent(self, event) -> None:  # noqa: N802
        key = event.key()
        modifiers = event.modifiers()
        if key in (Qt.Key_Return, Qt.Key_Enter):
            self.finish()
            event.accept()
            return
        if key == Qt.Key_Escape:
            self.cancel()
            event.accept()
            return
        if key == Qt.Key_Backspace or (
                key == Qt.Key_Z and modifiers & Qt.ControlModifier):
            dirty = self._hud_rect()
            if self._regions:
                removed = self._regions.pop()
                dirty = dirty.united(self._inflate(removed.widget_rect))
                self._hint = f"已撤销一块，还剩 {len(self._regions)} 块"
            elif self._dragging:
                if self._current is not None:
                    dirty = dirty.united(self._inflate(self._current))
                self._dragging = False
                self._current = None
                self._press_widget = None
                self._hint = "已取消当前框选"
            self.update(dirty)
            event.accept()
            return
        if key == Qt.Key_Delete:
            if self._regions:
                dirty = self._hud_rect()
                for region in self._regions:
                    dirty = dirty.united(self._inflate(region.widget_rect))
                self._regions.clear()
                self._hint = "已清空全部选区"
                self.update(dirty)
            event.accept()
            return
        super().keyPressEvent(event)

    # -- reading / closing -------------------------------------------------
    def finish(self) -> None:
        """Read every selected region and emit :attr:`finished`."""
        if not self._regions:
            self._hint = "请先长按右键拖动，框选至少一块单元格区域"
            self._update_hud()
            return

        blocks: list[CopyBlock] = []
        errors: list[str] = []
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for region in self._regions:
                try:
                    block = self.reader(*region.screen_rect)
                except ScreenCopyError as exc:
                    errors.append(str(exc))
                    continue
                except Exception as exc:  # noqa: BLE001 - Windows COM errors
                    errors.append(f"读取选区失败：{exc}")
                    continue
                if block.is_empty:
                    errors.append("有一块选区是空的，已跳过")
                else:
                    blocks.append(block)
        finally:
            QApplication.restoreOverrideCursor()

        if not blocks:
            self._hint = errors[0] if errors else "选区内没有读到数据，请重新框选"
            self._update_hud()
            return
        self._done = True
        self.hide()
        self.finished.emit(blocks, errors)

    def cancel(self) -> None:
        self._hold_timer.stop()
        self._done = True
        self.hide()
        self.cancelled.emit()

    def closeEvent(self, event) -> None:  # noqa: N802
        self._hold_timer.stop()
        if not self._done:
            self._done = True
            self.cancelled.emit()
        event.accept()

    # -- painting ----------------------------------------------------------
    def paintEvent(self, event) -> None:  # noqa: N802
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        dirty = event.rect()
        # Only touch the exposed area: a full-screen fill on a 4K desktop costs
        # several milliseconds per mouse move with software rendering.
        painter.fillRect(dirty, self._veil)

        for index, region in enumerate(self._regions, start=1):
            if region.widget_rect.intersects(dirty):
                self._paint_region(painter, region.widget_rect, index)
        if self._current is not None and self._current.intersects(dirty):
            self._paint_region(painter, self._current, None, dashed=True)
        if self._hud_rect().intersects(dirty):
            self._paint_hud(painter)
        painter.end()

    def _paint_region(self, painter: QPainter, rect: QRect,
                      index: int | None, dashed: bool = False) -> None:
        if rect.width() < 2 or rect.height() < 2:
            return
        painter.fillRect(rect, self._region_fill)
        painter.setPen(self._pen_dashed if dashed else self._pen_solid)
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(rect.adjusted(1, 1, -1, -1))

        if index is not None:
            badge = QRectF(rect.left() + 6, rect.top() + 6, 28, 21)
            painter.setPen(Qt.NoPen)
            painter.setBrush(self._green_dark)
            painter.drawRoundedRect(badge, 6, 6)
            painter.setFont(self._badge_font)
            painter.setPen(self._white)
            painter.drawText(badge, Qt.AlignCenter, str(index))

    def _paint_hud(self, painter: QPainter) -> None:
        box = QRectF(self._hud_rect().adjusted(2, 2, -2, -2))
        painter.setPen(Qt.NoPen)
        painter.setBrush(self._hud_bg)
        painter.drawRoundedRect(box, 14, 14)

        # Small green indicator dot.
        painter.setBrush(self._green)
        painter.drawEllipse(QRectF(box.left() + 16, box.top() + 17, 9, 9))

        painter.setFont(self._title_font)
        painter.setPen(self._white)
        painter.drawText(
            QRectF(box.left() + 32, box.top() + 8, box.width() - 44, 26),
            Qt.AlignLeft | Qt.AlignVCenter,
            "跨文件复制 · 长按右键拖动框选单元格")

        painter.setFont(self._hud_font)
        painter.setPen(self._hud_text)
        painter.drawText(
            QRectF(box.left() + 17, box.top() + 36, box.width() - 34, 20),
            Qt.AlignLeft | Qt.AlignVCenter,
            "Enter 完成 · Esc 取消 · Backspace 撤销 · Delete 清空")
        painter.setPen(self._hud_accent)
        count = f"已选 {len(self._regions)} 块"
        if self._hint:
            count += f" · {self._hint}"
        painter.drawText(
            QRectF(box.left() + 17, box.top() + 60, box.width() - 34, 24),
            Qt.AlignLeft | Qt.AlignVCenter, count)
