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
