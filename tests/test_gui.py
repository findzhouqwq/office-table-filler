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


if __name__ == "__main__":
    unittest.main()
