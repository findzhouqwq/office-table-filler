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

    def test_read_block_and_bulk_write(self):
        from openpyxl import Workbook, load_workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "数据"
        ws.append(["姓名", "工号"])
        ws.append(["张三", "E001"])
        ws.append([None, None])
        ws.append(["李四", "E002"])
        wb.save(self.path)

        eng = OpenpyxlEngine()
        eng.open(self.path)
        self.assertEqual(eng.read_block("数据", 1, 1, 4, 2), [
            ["姓名", "工号"], ["张三", "E001"], [None, None], ["李四", "E002"]])
        self.assertEqual(eng.read_block("数据", 3, 1, 4, 2), [
            [None, None], ["李四", "E002"]])
        self.assertEqual(eng.read_block("数据", 5, 1, 4, 2), [])
        self.assertEqual(eng.write_rows("数据", 6, [["王五", "E003"], ["赵六", "E004"]]), 2)
        eng.save(self.path)
        eng.close()
        self.assertEqual(load_workbook(self.path)["数据"]["A7"].value, "赵六")

    def test_load_spreadsheet_skips_blank_rows(self):
        from openpyxl import Workbook

        from app.core.data_io import load_spreadsheet

        wb = Workbook()
        ws = wb.active
        ws.title = "数据"
        ws.append(["姓名", "工号"])
        ws.append(["张三", "E001"])
        ws.append([None, None])
        ws.append(["李四", "E002"])
        wb.save(self.path)

        table = load_spreadsheet(self.path, header_row=1)
        self.assertEqual(table.headers, ["姓名", "工号"])
        self.assertEqual(table.rows, [["张三", "E001"], ["李四", "E002"]])

    def test_missing_file(self):
        from app.core.errors import EngineError
        eng = OpenpyxlEngine()
        with self.assertRaises(EngineError):
            eng.open(os.path.join(self.tmp, "nope.xlsx"))

    def test_formula_text_stays_text(self):
        """Values from OCR/imported data must not become live formulas."""
        from openpyxl import Workbook, load_workbook
        wb = Workbook()
        wb.active.append(["值"])
        wb.save(self.path)
        eng = OpenpyxlEngine()
        eng.open(self.path)
        eng.write_value("Sheet", "A2", '=HYPERLINK("http://example.com")')
        eng.save(self.path)
        eng.close()
        cell = load_workbook(self.path)["Sheet"]["A2"]
        self.assertEqual(cell.data_type, "s")
        self.assertTrue(cell.value.startswith("="))

    def test_invalid_coordinate_rejected(self):
        from openpyxl import Workbook
        from app.core.errors import FillError
        wb = Workbook()
        wb.save(self.path)
        eng = OpenpyxlEngine()
        eng.open(self.path)
        with self.assertRaises(FillError):
            eng.write_value("Sheet", "ZZZ1", "x")
        eng.close()


if __name__ == "__main__":
    unittest.main()
