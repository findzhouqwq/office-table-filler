import unittest

from app.core.targets import parse_target_spec, describe_target


class TargetSpecTests(unittest.TestCase):
    def test_xlsx_cell_with_sheet(self):
        ref = parse_target_spec("xlsx_cell", "Sheet1!B2 步进=1", "t.xlsx")
        self.assertEqual(ref.sheet, "Sheet1")
        self.assertEqual(ref.cell, "B2")
        self.assertEqual(ref.row, 1)
        self.assertEqual(ref.file, "t.xlsx")

    def test_xlsx_cell_without_sheet(self):
        ref = parse_target_spec("xlsx_cell", "C3")
        self.assertEqual(ref.cell, "C3")
        self.assertEqual(ref.row, -1)

    def test_docx_table_cell(self):
        ref = parse_target_spec("docx_table_cell", "table=0,row=1,col=2")
        self.assertEqual((ref.table, ref.row, ref.col), (0, 1, 2))

    def test_docx_table_rows(self):
        ref = parse_target_spec("docx_table_rows", "table=1,row_start=2")
        self.assertEqual((ref.table, ref.row_start), (1, 2))

    def test_pptx_table_cell(self):
        ref = parse_target_spec("pptx_table_cell", "slide=2,table=0,row=1,col=1")
        self.assertEqual((ref.slide, ref.table, ref.row, ref.col), (2, 0, 1, 1))

    def test_placeholder(self):
        ref = parse_target_spec("docx_placeholder", "{{姓名}}")
        self.assertEqual(ref.placeholder, "姓名")

    def test_describe(self):
        self.assertIn("Word 表格", describe_target("docx_table_cell", "table=0,row=1,col=2"))
        self.assertIn("占位符", describe_target("docx_placeholder", "{{姓名}}"))

    def test_invalid_placeholder(self):
        from app.core.errors import MappingError
        with self.assertRaises(MappingError):
            parse_target_spec("docx_placeholder", "没有花括号")


if __name__ == "__main__":
    unittest.main()
