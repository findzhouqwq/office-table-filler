"""Dialog that shows a fill preview and blocks execution on validation errors."""

from __future__ import annotations

import html

from PySide6.QtWidgets import (QAbstractItemView, QDialog, QDialogButtonBox,
                               QLabel, QTableWidget, QTableWidgetItem,
                               QTextEdit, QVBoxLayout, QWidget)

from ..core.preview import PreviewResult


class PreviewDialog(QDialog):
    def __init__(self, result: PreviewResult, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.result = result
        self.setWindowTitle(result.title or "预览 / 校验")
        self.resize(920, 640)
        self._build()

    def _build(self) -> None:
        layout = QVBoxLayout(self)

        summary = QLabel(self.result.summary or "执行预览")
        summary.setWordWrap(True)
        layout.addWidget(summary)

        if self.result.details:
            details = QLabel("\n".join(self.result.details))
            details.setWordWrap(True)
            layout.addWidget(details)

        issues = QTextEdit()
        issues.setReadOnly(True)
        issues.setMaximumHeight(130)
        issues.setHtml(self._issues_html())
        layout.addWidget(issues)

        table = QTableWidget(0, 0)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        if self.result.headers:
            table.setColumnCount(len(self.result.headers))
            table.setHorizontalHeaderLabels(self.result.headers)
        table.setRowCount(len(self.result.rows))
        for r, row in enumerate(self.result.rows):
            for c, value in enumerate(row):
                table.setItem(r, c, QTableWidgetItem(str(value)))
        if self.result.truncated:
            table.setToolTip(f"仅显示前 {len(self.result.rows)} 行")
        layout.addWidget(table, 1)

        buttons = QDialogButtonBox()
        self.run_button = buttons.addButton("开始执行", QDialogButtonBox.AcceptRole)
        buttons.addButton("取消", QDialogButtonBox.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        if not self.result.ok:
            self.run_button.setEnabled(False)
            self.run_button.setToolTip("存在校验错误，请先修正后再执行。")

    def _issues_html(self) -> str:
        issues = self.result.issues.issues
        if not issues:
            return '<div style="color:#2e7d32">校验通过，没有发现问题。</div>'
        parts: list[str] = []
        for issue in issues:
            color = "#b00020" if issue.level == "error" else "#8a6d00"
            parts.append(f'<div style="color:{color}">{html.escape(str(issue))}</div>')
        return "".join(parts)


def show_preview(parent: QWidget, result: PreviewResult) -> bool:
    """Show *result* and return True when the user chose to continue."""
    dialog = PreviewDialog(result, parent)
    return dialog.exec() == QDialog.Accepted
