"""Modern visual theme for the application.

Centralises the colour palette, the Qt stylesheet and a few small reusable
widgets so every page looks consistent.  Call :func:`apply_theme` once right
after creating the ``QApplication``.
"""

from __future__ import annotations

from PySide6.QtCore import QPointF, QRectF, Qt
from PySide6.QtGui import (QColor, QFont, QIcon, QLinearGradient, QPainter,
                           QPainterPath, QPixmap)
from PySide6.QtWidgets import (QApplication, QFrame, QHBoxLayout, QLabel,
                               QPushButton, QSizePolicy, QToolButton,
                               QVBoxLayout, QWidget)

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------

ACCENT = "#4F46E5"
ACCENT_HOVER = "#4338CA"
ACCENT_PRESSED = "#3730A3"
ACCENT_SOFT = "#EEF2FF"

BG = "#F4F6FB"
SURFACE = "#FFFFFF"
SURFACE_ALT = "#F8FAFC"
BORDER = "#E3E8F0"
BORDER_STRONG = "#CBD5E1"
TEXT = "#1F2937"
TEXT_STRONG = "#111827"
MUTED = "#6B7280"
SUCCESS = "#10B981"
WARNING = "#F59E0B"
DANGER = "#EF4444"

RADIUS = 10


def build_stylesheet() -> str:
    """Return the application-wide Qt stylesheet."""
    return f"""
    * {{
        font-family: "Microsoft YaHei UI", "Segoe UI", "PingFang SC", "Noto Sans CJK SC", sans-serif;
        font-size: 13px;
        color: {TEXT};
        outline: none;
    }}

    QMainWindow, #AppRoot, QDialog {{
        background: {BG};
    }}

    /* ---------- top header ---------- */
    QFrame#Header {{
        background: {SURFACE};
        border: none;
        border-bottom: 1px solid {BORDER};
    }}
    QLabel#AppTitle {{
        font-size: 17px;
        font-weight: 700;
        color: {TEXT_STRONG};
    }}
    QLabel#AppSubtitle {{
        font-size: 12px;
        color: {MUTED};
    }}
    QLabel#Pill {{
        background: {ACCENT_SOFT};
        color: {ACCENT};
        border-radius: 10px;
        padding: 3px 10px;
        font-size: 12px;
        font-weight: 600;
    }}

    /* ---------- tabs ---------- */
    QTabWidget::pane {{
        border: none;
        background: transparent;
        top: 8px;
    }}
    QTabWidget::tab-bar {{
        alignment: left;
    }}
    QTabBar {{
        background: transparent;
    }}
    QTabBar::tab {{
        background: transparent;
        color: {MUTED};
        border: 1px solid transparent;
        border-radius: 9px;
        padding: 9px 18px;
        margin-right: 6px;
        margin-bottom: 2px;
        font-weight: 600;
    }}
    QTabBar::tab:hover {{
        background: #E9EDF7;
        color: {TEXT};
    }}
    QTabBar::tab:selected {{
        background: {SURFACE};
        color: {ACCENT};
        border: 1px solid {BORDER};
    }}

    /* ---------- cards ---------- */
    QFrame#Card {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 12px;
    }}
    QFrame#TemplateBar {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 12px;
    }}
    QLabel#CardTitle {{
        font-size: 13px;
        font-weight: 700;
        color: {TEXT_STRONG};
    }}
    QLabel#SectionHint {{
        font-size: 12px;
        color: {MUTED};
    }}
    QLabel#StatusText {{
        color: {MUTED};
        font-size: 12px;
    }}
    QToolButton#CardToggle {{
        border: none;
        background: transparent;
        color: {TEXT_STRONG};
        font-weight: 700;
        font-size: 13px;
        padding: 2px 2px;
        text-align: left;
    }}
    QToolButton#CardToggle:hover {{
        color: {ACCENT};
    }}

    /* nested detail tabs inside a card */
    QTabWidget#DetailTabs::pane {{
        border: none;
        background: transparent;
        top: 6px;
    }}
    QTabWidget#DetailTabs QTabBar::tab {{
        background: transparent;
        color: {MUTED};
        border: 1px solid transparent;
        border-radius: 7px;
        padding: 6px 14px;
        margin-right: 4px;
        font-weight: 600;
    }}
    QTabWidget#DetailTabs QTabBar::tab:hover {{
        background: #F1F3FA;
        color: {TEXT};
    }}
    QTabWidget#DetailTabs QTabBar::tab:selected {{
        background: {ACCENT_SOFT};
        color: {ACCENT};
        border: 1px solid #E0E7FF;
    }}
    QTabWidget#DetailTabs QTabBar::tab:disabled {{
        color: #C3C8D2;
        background: transparent;
    }}

    /* ---------- buttons ---------- */
    QPushButton {{
        background: {SURFACE};
        color: {TEXT};
        border: 1px solid {BORDER_STRONG};
        border-radius: 8px;
        padding: 7px 14px;
        min-height: 16px;
    }}
    QPushButton:hover {{
        background: #F5F7FF;
        border-color: #A5B4FC;
    }}
    QPushButton:pressed {{
        background: {ACCENT_SOFT};
    }}
    QPushButton:disabled {{
        color: #A0A7B4;
        background: #F3F5F9;
        border-color: #E8ECF3;
    }}
    QPushButton[role="primary"] {{
        background: {ACCENT};
        color: #FFFFFF;
        border: 1px solid {ACCENT};
        font-weight: 700;
        padding: 9px 22px;
    }}
    QPushButton[role="primary"]:hover {{
        background: {ACCENT_HOVER};
        border-color: {ACCENT_HOVER};
    }}
    QPushButton[role="primary"]:pressed {{
        background: {ACCENT_PRESSED};
    }}
    QPushButton[role="primary"]:disabled {{
        background: #C7D2FE;
        border-color: #C7D2FE;
        color: #FFFFFF;
    }}
    QPushButton[role="ghost"] {{
        background: transparent;
        border: 1px solid transparent;
        color: {ACCENT};
        font-weight: 600;
        padding: 6px 10px;
    }}
    QPushButton[role="ghost"]:hover {{
        background: {ACCENT_SOFT};
    }}
    QPushButton[role="danger"] {{
        background: transparent;
        border: 1px solid transparent;
        color: {DANGER};
        padding: 6px 10px;
    }}
    QPushButton[role="danger"]:hover {{
        background: #FEF2F2;
        border-color: #FECACA;
    }}

    /* ---------- inputs ---------- */
    QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {{
        background: {SURFACE};
        border: 1px solid {BORDER_STRONG};
        border-radius: 8px;
        padding: 6px 10px;
        min-height: 18px;
        selection-background-color: #C7D2FE;
        selection-color: {TEXT_STRONG};
    }}
    QLineEdit:hover, QComboBox:hover, QSpinBox:hover, QDoubleSpinBox:hover {{
        border-color: #A5B4FC;
    }}
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {{
        border: 1px solid {ACCENT};
    }}
    QLineEdit:read-only {{
        background: {SURFACE_ALT};
        color: {MUTED};
    }}
    QComboBox::drop-down {{
        border: none;
        width: 24px;
        subcontrol-origin: padding;
        subcontrol-position: center right;
    }}
    QComboBox QAbstractItemView {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 8px;
        padding: 4px;
        selection-background-color: {ACCENT_SOFT};
        selection-color: {ACCENT};
        outline: none;
    }}
    QSpinBox::up-button, QSpinBox::down-button,
    QDoubleSpinBox::up-button, QDoubleSpinBox::down-button {{
        width: 16px;
        border: none;
        background: transparent;
    }}

    QTextEdit, QPlainTextEdit {{
        background: {SURFACE};
        border: 1px solid {BORDER_STRONG};
        border-radius: 10px;
        padding: 8px;
        selection-background-color: #C7D2FE;
    }}
    QTextEdit:focus, QPlainTextEdit:focus {{
        border: 1px solid {ACCENT};
    }}

    /* ---------- lists & tables ---------- */
    QListWidget, QTableWidget, QTreeWidget {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 10px;
        gridline-color: #EEF1F6;
        alternate-background-color: {SURFACE_ALT};
        selection-background-color: {ACCENT_SOFT};
        selection-color: {TEXT_STRONG};
    }}
    QListWidget::item {{
        padding: 7px 8px;
        border-radius: 6px;
    }}
    QListWidget::item:hover {{
        background: {SURFACE_ALT};
    }}
    QListWidget::item:selected {{
        background: {ACCENT_SOFT};
        color: {ACCENT};
    }}
    QTableWidget::item {{
        padding: 5px 8px;
    }}
    QHeaderView::section {{
        background: #F7F9FC;
        color: #4B5563;
        border: none;
        border-bottom: 1px solid {BORDER};
        border-right: 1px solid #EEF1F6;
        padding: 8px 8px;
        font-weight: 700;
    }}
    QTableCornerButton::section {{
        background: #F7F9FC;
        border: none;
    }}

    /* ---------- misc ---------- */
    QCheckBox, QRadioButton {{
        spacing: 8px;
        color: {TEXT};
    }}
    QCheckBox::indicator, QRadioButton::indicator {{
        width: 16px;
        height: 16px;
    }}
    QGroupBox {{
        border: 1px solid {BORDER};
        border-radius: 12px;
        margin-top: 16px;
        padding: 14px 12px 12px 12px;
        background: {SURFACE};
        font-weight: 700;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 14px;
        padding: 0 6px;
        color: {TEXT_STRONG};
    }}
    QSplitter::handle {{
        background: transparent;
    }}
    QSplitter::handle:horizontal {{
        width: 10px;
    }}
    QSplitter::handle:vertical {{
        height: 10px;
    }}
    QSplitter::handle:hover {{
        background: #E3E8F0;
        border-radius: 3px;
    }}

    QScrollBar:vertical {{
        background: transparent;
        width: 10px;
        margin: 2px;
    }}
    QScrollBar::handle:vertical {{
        background: #CBD5E1;
        border-radius: 4px;
        min-height: 28px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: #A5B4FC;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}
    QScrollBar:horizontal {{
        background: transparent;
        height: 10px;
        margin: 2px;
    }}
    QScrollBar::handle:horizontal {{
        background: #CBD5E1;
        border-radius: 4px;
        min-width: 28px;
    }}
    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
        width: 0;
    }}

    QStatusBar {{
        background: {SURFACE};
        border-top: 1px solid {BORDER};
        color: {MUTED};
        font-size: 12px;
    }}
    QStatusBar::item {{
        border: none;
    }}
    QMenuBar {{
        background: {SURFACE};
        border-bottom: 1px solid {BORDER};
    }}
    QMenuBar::item {{
        padding: 6px 10px;
        border-radius: 6px;
        background: transparent;
    }}
    QMenuBar::item:selected {{
        background: {ACCENT_SOFT};
        color: {ACCENT};
    }}
    QMenu {{
        background: {SURFACE};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 6px;
    }}
    QMenu::item {{
        padding: 7px 22px 7px 14px;
        border-radius: 6px;
    }}
    QMenu::item:selected {{
        background: {ACCENT_SOFT};
        color: {ACCENT};
    }}
    QToolTip {{
        background: {TEXT_STRONG};
        color: #FFFFFF;
        border: none;
        border-radius: 6px;
        padding: 6px 9px;
    }}
    """


def apply_theme(app: QApplication) -> None:
    """Apply the Fusion style, palette and stylesheet to *app*."""
    app.setStyle("Fusion")
    app.setStyleSheet(build_stylesheet())
    try:
        app.setWindowIcon(make_app_icon())
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Reusable widgets
# ---------------------------------------------------------------------------

class Card(QFrame):
    """A rounded white panel with an optional title row.

    ``card.body`` is the vertical layout to add content to.
    """

    def __init__(self, title: str = "", hint: str = "",
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(14, 12, 14, 12)
        self._layout.setSpacing(10)

        if title:
            head = QHBoxLayout()
            head.setContentsMargins(0, 0, 0, 0)
            head.setSpacing(8)
            bar = QFrame()
            bar.setFixedSize(4, 15)
            bar.setStyleSheet(
                f"background: {ACCENT}; border-radius: 2px; border: none;")
            label = QLabel(title)
            label.setObjectName("CardTitle")
            head.addWidget(bar)
            head.addWidget(label)
            if hint:
                hint_label = QLabel(hint)
                hint_label.setObjectName("SectionHint")
                head.addSpacing(6)
                head.addWidget(hint_label)
            head.addStretch(1)
            self._layout.addLayout(head)

        self.body = self._layout


def section_title(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("CardTitle")
    return label


def hint_label(text: str) -> QLabel:
    label = QLabel(text)
    label.setObjectName("SectionHint")
    return label


def set_role(button: QPushButton, role: str) -> QPushButton:
    """Set a dynamic style role: primary / ghost / danger / (empty)."""
    button.setProperty("role", role)
    button.style().unpolish(button)
    button.style().polish(button)
    return button


def primary_button(text: str, parent: QWidget | None = None) -> QPushButton:
    button = QPushButton(text, parent)
    button.setCursor(Qt.PointingHandCursor)
    return set_role(button, "primary")


def ghost_button(text: str, parent: QWidget | None = None) -> QPushButton:
    button = QPushButton(text, parent)
    button.setCursor(Qt.PointingHandCursor)
    return set_role(button, "ghost")


def danger_button(text: str, parent: QWidget | None = None) -> QPushButton:
    button = QPushButton(text, parent)
    button.setCursor(Qt.PointingHandCursor)
    return set_role(button, "danger")


def divider() -> QFrame:
    line = QFrame()
    line.setFrameShape(QFrame.HLine)
    line.setStyleSheet(f"background: {BORDER}; border: none; max-height: 1px;")
    return line


class CollapsibleCard(QFrame):
    """A card whose body can be collapsed to a single header row.

    Collapsed by default, which keeps secondary panels such as the run report
    from stealing vertical space.
    """

    def __init__(self, title: str, expanded: bool = False,
                 parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("Card")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 8, 14, 10)
        outer.setSpacing(6)

        self._button = QToolButton()
        self._button.setObjectName("CardToggle")
        self._button.setText(title)
        self._button.setCheckable(True)
        self._button.setChecked(expanded)
        self._button.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self._button.setArrowType(Qt.DownArrow if expanded else Qt.RightArrow)
        self._button.setCursor(Qt.PointingHandCursor)
        self._button.clicked.connect(self._on_clicked)
        outer.addWidget(self._button)

        self._content = QWidget()
        self.body = QVBoxLayout(self._content)
        self.body.setContentsMargins(0, 0, 0, 0)
        self.body.setSpacing(8)
        outer.addWidget(self._content)

        self._expanded = expanded
        self._content.setVisible(expanded)

    def _on_clicked(self) -> None:
        self.set_expanded(self._button.isChecked())

    def set_expanded(self, expanded: bool) -> None:
        self._expanded = expanded
        self._button.setChecked(expanded)
        self._button.setArrowType(Qt.DownArrow if expanded else Qt.RightArrow)
        self._content.setVisible(expanded)

    def is_expanded(self) -> bool:
        return self._expanded


# ---------------------------------------------------------------------------
# Application icon (drawn at runtime, no asset files needed)
# ---------------------------------------------------------------------------

def make_app_icon(size: int = 64) -> QIcon:
    """Draw a simple rounded-square icon with a grid glyph."""
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)

    gradient = QLinearGradient(QPointF(0, 0), QPointF(size, size))
    gradient.setColorAt(0.0, QColor("#6366F1"))
    gradient.setColorAt(1.0, QColor("#4338CA"))
    path = QPainterPath()
    path.addRoundedRect(QRectF(2, 2, size - 4, size - 4), size * 0.24, size * 0.24)
    painter.fillPath(path, gradient)

    # grid glyph
    painter.setPen(Qt.NoPen)
    painter.setBrush(QColor(255, 255, 255, 235))
    margin = size * 0.24
    gap = size * 0.06
    cell = (size - 2 * margin - gap) / 2.0
    for row in range(2):
        for col in range(2):
            x = margin + col * (cell + gap)
            y = margin + row * (cell + gap)
            painter.drawRoundedRect(QRectF(x, y, cell, cell), cell * 0.22, cell * 0.22)
    painter.end()
    return QIcon(pixmap)
