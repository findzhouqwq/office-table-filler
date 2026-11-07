import os
import tempfile
import unittest

from app.core.spreadsheet import OpenpyxlEngine


class SpreadsheetTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="tf_")
        self.path = os.path.join(self.tmp, "book.xlsx")

    def test_create_and_write(self):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.title = "数据"
        wb.active.append(["姓名", "工号"])
        wb.active.append(["张三", "E001"])
        wb.save(self.path)

        eng = OpenpyxlEngine()
        eng.open(self.path)
        self.assertEqual(eng.sheets(), ["数据"])
        self.assertEqual(eng.dims("数据"), (2, 2))
        self.assertEqual(eng.cell_value("数据", "A1"), "姓名")
        self.assertEqual(eng.cell_value("数据", "B2"), "E001")
        self.assertTrue(eng.cell_exists("数据", "A1"))
        self.assertFalse(eng.cell_exists("数据", "Z9"))

        eng.write_value("数据", "C1", "部门")
        eng.write_value("数据", "C2", "市场部")
        self.assertEqual(eng.cell_value("数据", "C1"), "部门")
        self.assertEqual(eng.cell_value("数据", "A2"), "张三")
        self.assertEqual(eng.cell_value("数据", "C2"), "市场部")
        self.assertEqual(eng.header_map("数据", 1), {"姓名": "A", "工号": "B", "部门": "C"})
        eng.save(self.path)
        eng.close()

        # reopen to confirm persistence
        eng = OpenpyxlEngine()
        eng.open(self.path)
        self.assertEqual(eng.cell_value("数据", "C1"), "部门")
        eng.close()

    def test_append_rows(self):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.append(["a", "b"])
        wb.active.append([1, 2])
        wb.save(self.path)
        eng = OpenpyxlEngine()
        eng.open(self.path)
        written = eng.append_rows("Sheet", [[3, 4], [5, 6]])
        self.assertEqual(written, 2)
        eng.save(self.path)
        self.assertEqual(eng.cell_value("Sheet", "A4"), 5)
        eng.close()

    def test_missing_file(self):
        from app.core.errors import EngineError
        eng = OpenpyxlEngine()
        with self.assertRaises(EngineError):
            eng.open(os.path.join(self.tmp, "nope.xlsx"))


if __name__ == "__main__":
    unittest.main()
