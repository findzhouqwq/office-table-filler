"""Headless smoke tests for the PySide6 UI.

Run with::

    QT_QPA_PLATFORM=offscreen python -m unittest tests.test_gui -v
"""

import os
import sys
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication  # noqa: E402


class GuiSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication(sys.argv)

    def test_mapping_widgets_construct(self):
        from app.ui.mapping_widgets import PlanTable, RuleTable, TargetTable
        from app.core.models import ExtractionRule, FieldMapping, MappingEntry, TargetRef

        rules = RuleTable()
        rules.load([ExtractionRule(field="姓名", pattern_type="label", pattern="姓名")])
        self.assertEqual(len(rules.to_rules()), 1)

        plan = PlanTable()
        plan.load([FieldMapping(target="A", source="姓名")])
        self.assertEqual(plan.to_plan()[0].target, "A")

        targets = TargetTable()
        targets.load([MappingEntry(source="姓名",
                                    target=TargetRef(kind="xlsx_cell", file="t.xlsx",
                                                      sheet="S", cell="A1", row=1))])
        self.assertEqual(len(targets.to_entries("t.xlsx")), 1)
        # close the C-level objects we created
        for w in (rules, plan, targets):
            w.close()

    def test_main_window_constructs(self):
        from app.ui.main_window import MainWindow
        win = MainWindow()
        self.assertEqual(win.tabs.count(), 3)
        win.close()

    def test_tabs_construct(self):
        from app.ui.image_tab import ImageTab
        from app.ui.text_tab import TextTab
        from app.ui.cross_tab import CrossTab
        for tab in (ImageTab(), TextTab(), CrossTab()):
            tab.close()

    def test_side_panel_behaviour(self):
        from app.ui.side_panel import SidePanel

        panel = SidePanel()
        panel.set_fields_from_headers(["姓名", "电话"])
        self.assertEqual([label for label, _ in panel.fields()], ["姓名", "电话"])

        panel._rows[0].value_edit.setText("张三")
        panel._rows[1].value_edit.setText("13800000000")
        panel.copy_row()
        self.assertEqual(QApplication.clipboard().text(), "张三\t13800000000")

        panel.headers_provider = lambda: ["姓名", "部门"]
        panel._read_headers()
        self.assertEqual([label for label, _ in panel.fields()], ["姓名", "部门"])
        self.assertEqual(panel.fields()[0][1], "张三")

        panel.show_panel(animated=False)
        self.assertTrue(panel.isVisible())
        panel.set_collapsed(True, animated=False)
        self.assertTrue(panel.is_collapsed())
        panel.set_collapsed(False, animated=False)
        self.assertFalse(panel.is_collapsed())
        panel.hide_panel()
        self.assertFalse(panel.isVisible())

    def test_side_panel_clipboard_detection(self):
        from app.ui.side_panel import SidePanel

        panel = SidePanel()
        panel.headers_provider = lambda: ["姓名", "联系电话", "电子邮箱", "部门"]
        QApplication.clipboard().setText(
            "姓名：李四\n手机：13900001111\n邮箱：lisi@example.com")
        panel.detect_clipboard()
        self.assertEqual([label for label, _ in panel.fields()],
                         ["姓名", "联系电话", "电子邮箱", "部门"])
        self.assertEqual(panel.fields()[0][1], "李四")
        self.assertEqual(panel.fields()[1][1], "13900001111")
        self.assertEqual(panel.fields()[2][1], "lisi@example.com")
        panel.hide_panel()

    def test_side_panel_copies_empty_columns(self):
        from app.ui.side_panel import SidePanel

        panel = SidePanel()
        panel.set_fields_from_headers(["姓名", "电话", "邮箱", "部门"])
        panel._rows[0].value_edit.setText("张三")
        panel._rows[2].value_edit.setText("a@b.com")
        panel._rows[3].value_edit.setText("市场部")
        panel.copy_row()
        self.assertEqual(QApplication.clipboard().text(),
                         "张三\t\ta@b.com\t市场部")
        panel.hide_panel()

    def test_side_panel_positional_row(self):
        from app.ui.side_panel import SidePanel

        panel = SidePanel()
        panel.headers_provider = lambda: ["姓名", "联系电话", "部门"]
        QApplication.clipboard().setText("张三\t13800138000\t市场部")
        panel.detect_clipboard()
        self.assertEqual(panel.fields(), [
            ("姓名", "张三"),
            ("联系电话", "13800138000"),
            ("部门", "市场部"),
        ])
        panel.hide_panel()

    def test_side_panel_keeps_manual_values_without_headers(self):
        from app.ui.side_panel import SidePanel

        panel = SidePanel()
        panel.headers_provider = lambda: []
        panel._clear_fields()
        panel.add_field("姓名", "手动填写")
        QApplication.clipboard().setText("姓名：李四\n手机：13900001111")
        panel.detect_clipboard()
        values = dict(panel.fields())
        self.assertEqual(values["姓名"], "手动填写")
        self.assertEqual(values["手机"], "13900001111")
        panel.hide_panel()

    def test_side_panel_fills_target_table(self):
        import tempfile

        from openpyxl import Workbook, load_workbook

        import app.ui.side_panel as side_panel_module
        from app.ui.main_window import MainWindow

        tmp = tempfile.mkdtemp(prefix="panel_")
        path = os.path.join(tmp, "target.xlsx")
        workbook = Workbook()
        workbook.active.title = "数据"
        workbook.active.append(["姓名", "联系电话", "电子邮箱", "部门"])
        workbook.save(path)

        win = MainWindow()
        win.tabs.setCurrentIndex(1)  # TextTab owns target_edit / output_edit
        tab = win.tabs.widget(1)
        tab.target_edit.setText(path)
        tab.sheet_combo.addItem("数据")
        tab.header_spin.setValue(1)
        tab.output_edit.setText(path)

        panel = win._ensure_side_panel()
        QApplication.clipboard().setText(
            "姓名：李四\n手机：13900001111\n邮箱：lisi@example.com\n部门：市场部")
        panel.detect_clipboard()

        original = side_panel_module.confirm
        side_panel_module.confirm = lambda *args, **kwargs: True
        try:
            panel.fill_target()
        finally:
            side_panel_module.confirm = original

        sheet = load_workbook(path)["数据"]
        self.assertEqual(
            [sheet.cell(row=2, column=col).value for col in range(1, 5)],
            ["李四", "13900001111", "lisi@example.com", "市场部"])
        panel.hide_panel()
        win.close()

    def test_selection_overlay_collects_blocks(self):
        from PySide6.QtCore import QEvent, QPointF, Qt
        from PySide6.QtGui import QKeyEvent, QMouseEvent

        from app.core.screen_copy import CopyBlock
        from app.ui.screen_select import SelectionOverlay

        overlay = SelectionOverlay()
        overlay.reader = lambda left, top, right, bottom: CopyBlock(
            (("张三", "13900001111"),), "src.xlsx", "数据", "A1:B1")

        def mouse(kind, x, y, button, buttons):
            event = QMouseEvent(kind, QPointF(x, y), QPointF(x, y),
                                button, buttons, Qt.NoModifier)
            QApplication.sendEvent(overlay, event)

        mouse(QEvent.MouseButtonPress, 30, 30, Qt.RightButton, Qt.RightButton)
        overlay._start_drag()
        mouse(QEvent.MouseMove, 160, 120, Qt.NoButton, Qt.RightButton)
        mouse(QEvent.MouseButtonRelease, 160, 120, Qt.RightButton, Qt.NoButton)
        self.assertEqual(len(overlay._regions), 1)

        received = []
        overlay.finished.connect(
            lambda blocks, errors: received.append((blocks, errors)))
        QApplication.sendEvent(overlay, QKeyEvent(
            QEvent.KeyPress, Qt.Key_Return, Qt.NoModifier))
        self.assertEqual(len(received), 1)
        self.assertEqual(received[0][0][0].rows, (("张三", "13900001111"),))
        self.assertEqual(received[0][1], [])

    def test_selection_overlay_escape_cancels(self):
        from PySide6.QtCore import QEvent, Qt
        from PySide6.QtGui import QKeyEvent

        from app.ui.screen_select import SelectionOverlay

        overlay = SelectionOverlay()
        cancelled = []
        overlay.cancelled.connect(lambda: cancelled.append(True))
        QApplication.sendEvent(overlay, QKeyEvent(
            QEvent.KeyPress, Qt.Key_Escape, Qt.NoModifier))
        self.assertEqual(cancelled, [True])

    def test_side_panel_pastes_copied_blocks(self):
        import tempfile

        from openpyxl import Workbook, load_workbook

        import app.ui.side_panel as side_panel_module
        from app.core.screen_copy import CopyBlock
        from app.ui.side_panel import SidePanel

        tmp = tempfile.mkdtemp(prefix="blocks_")
        path = os.path.join(tmp, "target.xlsx")
        workbook = Workbook()
        workbook.active.title = "数据"
        workbook.active.append(["姓名", "联系电话", "部门"])
        workbook.active.append(["旧人", "13800000000", "旧部门"])
        workbook.save(path)

        panel = SidePanel()
        panel.target_provider = lambda: {
            "path": path, "sheet": "数据", "header_row": 1, "output_path": ""}
        panel._copy_blocks = [
            CopyBlock((("张三", "13900001111", "市场部"),)),
            CopyBlock((("李四", "13700002222", "技术部"),)),
        ]
        panel._update_copy_buttons()
        self.assertTrue(panel.paste_blocks_btn.isEnabled())

        original = side_panel_module.confirm
        side_panel_module.confirm = lambda *args, **kwargs: True
        try:
            panel.paste_blocks()
        finally:
            side_panel_module.confirm = original

        sheet = load_workbook(path)["数据"]
        self.assertEqual(
            [sheet.cell(row=3, column=col).value for col in range(1, 4)],
            ["张三", "13900001111", "市场部"])
        self.assertEqual(sheet["A4"].value, "李四")
        self.assertIn("已粘贴 2 块", panel.status_label.text())
        # The buffer survives so the same blocks can go to another file.
        self.assertEqual(len(panel._copy_blocks), 2)
        panel.hide_panel()

    def test_collapsed_rail_has_screen_copy_shortcut(self):
        from PySide6.QtCore import Qt
        from PySide6.QtTest import QTest

        from app.ui.side_panel import SidePanel

        panel = SidePanel()
        panel.show_panel(animated=False)
        panel.set_collapsed(True, animated=False)
        calls = []
        panel.start_screen_copy = lambda: calls.append(True)

        self.assertTrue(panel._rail.isVisible())
        rect = panel._rail.copy_hit_rect()
        QTest.mouseClick(panel._rail, Qt.LeftButton, pos=rect.center())

        self.assertEqual(calls, [True])
        self.assertFalse(panel.is_collapsed())
        panel.hide_panel()

    def test_main_window_toggles_side_panel(self):
        from app.ui.main_window import MainWindow

        win = MainWindow()
        win._toggle_side_panel()
        self.assertIsNotNone(win.side_panel)
        self.assertTrue(win.side_panel.isVisible())
        self.assertTrue(win._side_action.isChecked())
        win._toggle_side_panel()
        self.assertFalse(win.side_panel.isVisible())
        self.assertFalse(win._side_action.isChecked())
        win.close()


if __name__ == "__main__":
    unittest.main()
