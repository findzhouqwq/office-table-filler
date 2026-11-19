"""Tests for copying spreadsheet blocks from screen selections."""

import os
import sys
import tempfile
import unittest

from app.core.errors import FillError
from app.core.screen_copy import (CopyBlock, ScreenCopyError, append_blocks,
                                  normalize_range_values, read_screen_rect,
                                  screen_copy_available, trim_block)


class BlockShapeTests(unittest.TestCase):
    def test_trims_trailing_blank_rows_and_columns(self):
        rows = [["姓名", "电话", None], ["张三", "138", None], [None, None, None]]
        self.assertEqual(trim_block(rows), (("姓名", "电话"), ("张三", "138")))

    def test_block_drops_all_blank_selection(self):
        block = CopyBlock(((None, ""), ("  ", None)))
        self.assertTrue(block.is_empty)
        self.assertEqual(block.height, 0)

    def test_dimensions_and_source(self):
        block = CopyBlock(((1, 2, 3), (4, 5, 6)),
                          r"C:\data\源表.xlsx", "数据", "A1:C2")
        self.assertEqual(block.height, 2)
        self.assertEqual(block.width, 3)
        self.assertEqual(block.source, "源表.xlsx · 数据 · A1:C2")

    def test_normalises_com_value_shapes(self):
        self.assertEqual(normalize_range_values("张三"), (("张三",),))
        self.assertEqual(normalize_range_values((1, 2, 3)), ((1, 2, 3),))
        self.assertEqual(normalize_range_values(((1, 2), (3, 4))),
                         ((1, 2), (3, 4)))
        self.assertEqual(normalize_range_values(None), ())
        # Ragged input is padded so every row has the same width.
        self.assertEqual(normalize_range_values(((1,), (2, 3))),
                         ((1, None), (2, 3)))


class AppendBlocksTests(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook

        self.tmp = tempfile.mkdtemp(prefix="sc_")
        self.path = os.path.join(self.tmp, "target.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名", "联系电话", "部门"])
        sheet.append(["旧人", "13800000000", "旧部门"])
        workbook.save(self.path)

    def test_appends_multiple_blocks_below_existing_data(self):
        from openpyxl import load_workbook

        blocks = [
            CopyBlock((("张三", "13900001111", "市场部"),)),
            CopyBlock((("李四", "13700002222", "技术部"),
                       ("王五", "13600003333", "财务部"))),
        ]
        report = append_blocks(self.path, blocks, sheet="数据")

        self.assertEqual((report.start_row, report.end_row), (3, 5))
        self.assertEqual(report.blocks, 2)
        self.assertEqual(report.rows, 3)
        sheet = load_workbook(self.path)["数据"]
        self.assertEqual(
            [sheet.cell(row=3, column=col).value for col in range(1, 4)],
            ["张三", "13900001111", "市场部"])
        self.assertEqual(sheet["A5"].value, "王五")
        self.assertTrue(any(name.startswith("target.bak-")
                            for name in os.listdir(self.tmp)))

    def test_skips_empty_blocks_and_writes_output_file(self):
        from openpyxl import load_workbook

        output = os.path.join(self.tmp, "output.xlsx")
        blocks = [CopyBlock(()), CopyBlock((("赵六",),))]
        report = append_blocks(self.path, blocks, sheet="数据",
                               output_path=output)

        self.assertEqual(report.blocks, 1)
        self.assertEqual(report.path, os.path.abspath(output))
        self.assertEqual(load_workbook(output)["数据"]["A3"].value, "赵六")
        # The original file is untouched when an output path is used.
        self.assertIsNone(load_workbook(self.path)["数据"]["A3"].value)

    def test_formatted_empty_rows_are_ignored(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "formatted.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名"])
        sheet.cell(row=100, column=1).number_format = "@"
        workbook.save(path)

        report = append_blocks(path, [CopyBlock((("张三",),))], sheet="数据")
        self.assertEqual(report.start_row, 2)
        self.assertEqual(load_workbook(path)["数据"]["A2"].value, "张三")

    def test_merged_rows_are_skipped(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "merged.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名", "电话"])
        sheet.merge_cells("A1:B2")
        workbook.save(path)

        report = append_blocks(path, [CopyBlock((("张三", "138"),))],
                               sheet="数据")
        self.assertEqual(report.start_row, 3)
        self.assertEqual(load_workbook(path)["数据"]["A3"].value, "张三")

    def test_control_characters_are_removed_and_formulas_stay_text(self):
        from openpyxl import load_workbook

        append_blocks(self.path, [CopyBlock((("张\x00三", "=1+1"),))],
                      sheet="数据")
        sheet = load_workbook(self.path)["数据"]
        self.assertEqual(sheet["A3"].value, "张三")
        self.assertEqual(sheet["B3"].value, "=1+1")
        self.assertEqual(sheet["B3"].data_type, "s")

    def test_too_long_cell_content_raises(self):
        with self.assertRaises(FillError):
            append_blocks(self.path, [CopyBlock((("x" * 32768,),))],
                          sheet="数据")

    def test_empty_selection_raises(self):
        with self.assertRaises(FillError):
            append_blocks(self.path, [CopyBlock(())], sheet="数据")

    def test_empty_sheet_starts_at_the_first_row(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "blank.xlsx")
        workbook = Workbook()
        workbook.active.title = "空白"
        workbook.save(path)

        report = append_blocks(path, [CopyBlock((("张三", "138"),))],
                               sheet="空白")
        self.assertEqual(report.start_row, 1)
        sheet = load_workbook(path)["空白"]
        self.assertEqual(sheet["A1"].value, "张三")
        self.assertEqual(sheet["B1"].value, "138")

    def test_missing_target_path_raises(self):
        with self.assertRaises(FillError):
            append_blocks("", [CopyBlock((("张三",),))])


class ScreenReadTests(unittest.TestCase):
    def test_screen_copy_is_reported_unavailable_off_windows(self):
        if sys.platform.startswith("win"):
            self.skipTest("Windows has pywin32 / Excel available")
        available, reason = screen_copy_available()
        self.assertFalse(available)
        self.assertIn("Windows", reason)

    def test_read_screen_rect_fails_gracefully_off_windows(self):
        if sys.platform.startswith("win"):
            self.skipTest("Windows may have a live Excel instance")
        with self.assertRaises(ScreenCopyError):
            read_screen_rect(10, 10, 200, 120)


if __name__ == "__main__":
    unittest.main()
