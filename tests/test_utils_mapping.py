import unittest

from app.core.utils import (apply_transform, clean_text, detect_delimiter,
                            norm_header, parse_number, parse_date, to_number)
from app.core.mapping import (auto_map_headers, extract_fields,
                              parse_labeled_lines, resolve_column)
from app.core.models import ExtractionRule


class UtilsTests(unittest.TestCase):
    def test_clean_text_strips_and_collapses(self):
        self.assertEqual(clean_text("  张三 \t  "), "张三")
        self.assertEqual(clean_text("a\r\nb"), "a\nb")

    def test_norm_header_handles_punctuation(self):
        self.assertEqual(norm_header(" 姓名："), norm_header("姓名"))
        self.assertEqual(norm_header("A-B_C"), norm_header("abc"))

    def test_parse_number_variants(self):
        self.assertEqual(parse_number("1,234.5"), (1234.5, "1,234.5"))
        self.assertEqual(parse_number("１２３"), (123.0, "123"))
        self.assertIsNone(parse_number("无")[0])
        self.assertEqual(parse_number("50%"), (0.5, "50%"))

    def test_to_number_returns_int(self):
        self.assertEqual(to_number("12"), 12)
        self.assertEqual(to_number("12.5"), 12.5)
        self.assertEqual(to_number("abc", "N/A"), "N/A")

    def test_transforms(self):
        self.assertEqual(apply_transform("  hello ", "upper"), "HELLO")
        self.assertEqual(apply_transform("12.0", "number"), "12")
        self.assertEqual(apply_transform("x", "unknown"), "x")

    def test_detect_delimiter(self):
        self.assertEqual(detect_delimiter("a\tb\tc"), "\t")
        self.assertEqual(detect_delimiter("a,b,c"), ",")

    def test_parse_date(self):
        self.assertEqual(parse_date("2022年7月15日"), "2022-07-15")
        self.assertEqual(parse_date("2022/7/15"), "2022-07-15")


class MappingTests(unittest.TestCase):
    def test_extract_label(self):
        rules = [ExtractionRule(field="姓名", pattern_type="label", pattern="姓名")]
        result = extract_fields("姓名：李四\n工号：E010", rules)
        self.assertEqual(result["姓名"], "李四")

    def test_extract_regex(self):
        rules = [ExtractionRule(field="工号", pattern_type="regex", pattern=r"(\d{4})")]
        result = extract_fields("编号 1234 部门", rules)
        self.assertEqual(result["工号"], "1234")

    def test_extract_line(self):
        rules = [ExtractionRule(field="内容", pattern_type="line", position=1)]
        self.assertEqual(extract_fields("第一行\n第二行", rules)["内容"], "第二行")

    def test_parse_labeled_lines(self):
        data = parse_labeled_lines("姓名：张三\n年龄：30\n无标签行")
        self.assertEqual(data["姓名"], "张三")
        self.assertEqual(data["年龄"], "30")
        self.assertNotIn("无标签行", data)

    def test_auto_map_headers(self):
        pairs = auto_map_headers(["姓名", "工号", "部门"], ["部门", "姓名", "备注"])
        self.assertIn((0, 1), pairs)
        self.assertIn((2, 0), pairs)

    def test_resolve_column_by_letter_and_header(self):
        header_map = {"姓名": "A", "工号": "B"}
        self.assertEqual(resolve_column("C", header_map), "C")
        self.assertEqual(resolve_column("工号", header_map), "B")


if __name__ == "__main__":
    unittest.main()
