"""Tests for clipboard field detection, header mapping and quick fill."""

import os
import tempfile
import unittest

from app.core.field_detect import (append_fields, detect_fields, map_to_headers,
                                   map_values_to_headers, split_positional_row)


class DetectFieldsTests(unittest.TestCase):
    def _by_kind(self, text: str) -> dict[str, str]:
        return {field.kind: field.value for field in detect_fields(text)}

    def test_labeled_contact_card(self):
        text = "\n".join([
            "姓名：张三",
            "手机：13800138000",
            "身份证号：110101199003071234",
            "邮箱：zhangsan@example.com",
            "公司：北京字节跳动科技有限公司",
            "部门：市场部",
            "职位：经理",
            "出生日期：1990-03-07",
            "工资：12000元",
            "备注：请尽快联系",
        ])
        found = self._by_kind(text)
        self.assertEqual(found["name"], "张三")
        self.assertEqual(found["phone"], "13800138000")
        self.assertEqual(found["id_card"], "110101199003071234")
        self.assertEqual(found["email"], "zhangsan@example.com")
        self.assertEqual(found["company"], "北京字节跳动科技有限公司")
        self.assertEqual(found["department"], "市场部")
        self.assertEqual(found["position"], "经理")
        self.assertEqual(found["date"], "1990-03-07")
        self.assertEqual(found["amount"], "12000元")
        self.assertEqual(found["remark"], "请尽快联系")

    def test_values_on_the_line_below_the_label(self):
        found = self._by_kind("姓名：\n张三\n手机：\n13800138000")
        self.assertEqual(found["name"], "张三")
        self.assertEqual(found["phone"], "13800138000")

    def test_sentence_lead_ins_are_trimmed(self):
        found = self._by_kind("我的地址是北京市朝阳区建国路88号")
        self.assertEqual(found["address"], "北京市朝阳区建国路88号")
        found = self._by_kind("他在北京字节跳动科技有限公司上班")
        self.assertEqual(found["company"], "北京字节跳动科技有限公司")
        found = self._by_kind("联系地址为广东省深圳市南山区科技园1号")
        self.assertEqual(found["address"], "广东省深圳市南山区科技园1号")

    def test_company_name_is_not_trimmed_by_its_own_prefix(self):
        found = self._by_kind("在线科技有限公司")
        self.assertEqual(found["company"], "在线科技有限公司")
        found = self._by_kind("工作细胞有限公司")
        self.assertEqual(found["company"], "工作细胞有限公司")

    def test_single_tab_separated_row(self):
        self.assertEqual(
            split_positional_row("张三\t13800138000\t市场部"),
            ["张三", "13800138000", "市场部"])
        # A two-column label/value dump is handled by the label scanner.
        self.assertIsNone(split_positional_row("姓名\t张三\t手机\t13800138000"))
        self.assertIsNone(split_positional_row("电话\t13800138000"))
        self.assertIsNone(split_positional_row("张三\t13800138000\n李四\t13900001111"))

    def test_several_fields_on_one_line(self):
        found = self._by_kind("姓名：张三 性别：男 年龄：28")
        self.assertEqual(found["name"], "张三")
        self.assertEqual(found["gender"], "男")
        self.assertEqual(found["age"], "28")

    def test_unlabeled_values_are_classified(self):
        found = self._by_kind(
            "联系 13900001111，邮箱 lisi@example.com，"
            "地址 北京市朝阳区建国路88号")
        self.assertEqual(found["phone"], "13900001111")
        self.assertEqual(found["email"], "lisi@example.com")
        self.assertEqual(found["address"], "北京市朝阳区建国路88号")

    def test_known_and_unknown_labels(self):
        found = self._by_kind("工号：A123\n项目名称：星河计划\n入职日期：2024-05-06")
        self.assertEqual(found["number"], "A123")
        self.assertEqual(found["custom"], "星河计划")
        self.assertEqual(found["date"], "2024-05-06")

    def test_duplicate_values_are_reported_once(self):
        found = detect_fields("13800138000\n13800138000")
        self.assertEqual(len(found), 1)

    def test_large_input_is_capped(self):
        found = detect_fields("13800138000\n" * 5000)
        self.assertTrue(found)
        self.assertIn("phone", {field.kind for field in found})


class MappingTests(unittest.TestCase):
    def test_alias_mapping_onto_headers(self):
        detected = detect_fields("手机：13900001111\n邮箱：a@b.com")
        matched, unmatched = map_to_headers(
            detected, ["联系电话", "电子邮箱", "部门"])
        self.assertEqual(matched, {"联系电话": "13900001111",
                                   "电子邮箱": "a@b.com"})
        self.assertEqual(unmatched, [])

    def test_custom_label_containment(self):
        matched, _ = map_values_to_headers(
            {"工号": "A123"}, ["员工工号", "姓名"])
        self.assertEqual(matched, {"员工工号": "A123"})

    def test_unmatched_fields_are_returned(self):
        detected = detect_fields("职位：经理")
        matched, unmatched = map_to_headers(detected, ["姓名", "电话"])
        self.assertEqual(matched, {})
        self.assertEqual([field.value for field in unmatched], ["经理"])

    def test_prefix_does_not_cross_field_meanings(self):
        # “单位电话” is a phone column, not a company column.
        matched, _ = map_to_headers(
            detect_fields("公司：北京某科技有限公司"), ["单位电话", "姓名"])
        self.assertEqual(matched, {})
        matched, _ = map_to_headers(
            detect_fields("手机：13800138000"), ["单位电话"])
        self.assertEqual(matched, {"单位电话": "13800138000"})
        # “联系人电话” is a phone column, not a name column.
        matched, _ = map_to_headers(detect_fields("姓名：张三"), ["联系人电话"])
        self.assertEqual(matched, {})

    def test_header_decorations_and_numbered_columns(self):
        matched, _ = map_to_headers(
            detect_fields("手机：13800138000"), ["手机号码（必填）"])
        self.assertEqual(matched, {"手机号码（必填）": "13800138000"})
        matched, _ = map_to_headers(
            detect_fields("手机：13800138000"), ["电话1"])
        self.assertEqual(matched, {"电话1": "13800138000"})
        matched, _ = map_to_headers(
            detect_fields("金额：12000元"), ["金额（元）"])
        self.assertEqual(matched, {"金额（元）": "12000元"})
        matched, _ = map_to_headers(
            detect_fields("身份证：110101199003071234"), ["身份证号码（18位）"])
        self.assertEqual(matched, {"身份证号码（18位）": "110101199003071234"})


class AppendFieldsTests(unittest.TestCase):
    def setUp(self):
        from openpyxl import Workbook

        self.tmp = tempfile.mkdtemp(prefix="fd_")
        self.path = os.path.join(self.tmp, "book.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名", "联系电话", "电子邮箱", "部门"])
        sheet.append(["旧人", "13800000000", "old@example.com", "旧部门"])
        workbook.save(self.path)

    def test_append_mapped_row_and_backup(self):
        from openpyxl import load_workbook

        report = append_fields(
            self.path,
            {"姓名": "张三", "手机": "13900001111",
             "邮箱": "zhangsan@example.com", "职位": "经理", "部门": "市场部"},
            sheet="数据", header_row=1)

        self.assertEqual(report.row, 3)
        self.assertEqual(report.matched, 4)
        self.assertEqual(report.skipped, ("职位",))

        sheet = load_workbook(self.path)["数据"]
        self.assertEqual(
            [sheet.cell(row=3, column=col).value for col in range(1, 5)],
            ["张三", "13900001111", "zhangsan@example.com", "市场部"])
        self.assertTrue(any(name.startswith("book.bak-")
                            for name in os.listdir(self.tmp)))

    def test_missing_headers_raise(self):
        from openpyxl import Workbook

        from app.core.errors import FillError

        path = os.path.join(self.tmp, "empty.xlsx")
        workbook = Workbook()
        workbook.active.title = "空白"
        workbook.save(path)
        with self.assertRaises(FillError):
            append_fields(path, {"姓名": "张三"}, sheet="空白", header_row=1)

    def test_duplicate_header_names_use_separate_columns(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "duplicate.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["电话", "电话"])
        workbook.save(path)

        report = append_fields(
            path, {"手机": "13800138000", "联系电话": "13900139000"},
            sheet="数据", header_row=1)
        self.assertEqual(report.matched, 2)
        sheet = load_workbook(path)["数据"]
        self.assertEqual(sheet["A2"].value, "13800138000")
        self.assertEqual(sheet["B2"].value, "13900139000")

    def test_formatted_empty_rows_are_ignored(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "formatted.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名", "联系电话"])
        # Formatting far below the table makes max_row large but has no data.
        sheet.cell(row=100, column=1).number_format = "@"
        workbook.save(path)

        report = append_fields(path, {"姓名": "张三"}, sheet="数据", header_row=1)
        self.assertEqual(report.row, 2)
        sheet = load_workbook(path)["数据"]
        self.assertEqual(sheet["A2"].value, "张三")
        self.assertIsNone(sheet["A100"].value)

    def test_control_characters_are_removed(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "control.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名"])
        workbook.save(path)

        append_fields(path, {"姓名": "张\x00三"}, sheet="数据", header_row=1)
        sheet = load_workbook(path)["数据"]
        self.assertEqual(sheet["A2"].value, "张三")

    def test_writes_to_separate_output_file(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "target.xlsx")
        output = os.path.join(self.tmp, "output.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["姓名"])
        workbook.save(path)

        report = append_fields(path, {"姓名": "张三"}, sheet="数据",
                               header_row=1, output_path=output)
        self.assertEqual(os.path.abspath(report.path), os.path.abspath(output))
        self.assertIsNone(load_workbook(path)["数据"]["A2"].value)
        self.assertEqual(load_workbook(output)["数据"]["A2"].value, "张三")

    def test_formula_like_values_are_stored_as_text(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "formula.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet.append(["备注"])
        workbook.save(path)

        append_fields(path, {"备注": "=HYPERLINK(\"http://x\")"},
                      sheet="数据", header_row=1)
        sheet = load_workbook(path)["数据"]
        self.assertEqual(sheet["A2"].data_type, "s")
        self.assertEqual(sheet["A2"].value, "=HYPERLINK(\"http://x\")")

    def test_merged_header_rows_are_skipped(self):
        from openpyxl import Workbook, load_workbook

        path = os.path.join(self.tmp, "merged.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "数据"
        sheet["A1"] = "姓名"
        sheet.merge_cells("A1:A2")
        workbook.save(path)

        report = append_fields(path, {"姓名": "张三"}, sheet="数据", header_row=1)
        self.assertEqual(report.row, 3)
        sheet = load_workbook(path)["数据"]
        self.assertEqual(sheet["A3"].value, "张三")


if __name__ == "__main__":
    unittest.main()
