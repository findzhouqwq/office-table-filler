import os
import shutil
import tempfile
import unittest

from app.core import validation
from app.core.models import FieldMapping, MappingEntry, TargetRef


class ValidationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="tf_val_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _path(self, name):
        return os.path.join(self.tmp, name)

    def test_cell_and_anchor(self):
        self.assertTrue(validation.validate_cell("B2"))
        self.assertTrue(validation.validate_anchor("$AA$10"))
        self.assertFalse(validation.validate_cell("A0"))
        self.assertFalse(validation.validate_cell("A1:B2"))
        self.assertFalse(validation.validate_anchor("XFE1"))

    def test_plan_resolves_header_and_detects_duplicates(self):
        header_map = {"姓名": "A", "工号": "B"}
        plan = [
            FieldMapping(target="姓名", source="name"),
            FieldMapping(target="工号", source="id"),
        ]
        result = validation.validate_field_plan(plan, header_map)
        self.assertTrue(result.ok)

        duplicate = plan + [FieldMapping(target="A", source="other")]
        result = validation.validate_field_plan(duplicate, header_map)
        self.assertFalse(result.ok)
        self.assertTrue(any("同时使用" in issue.message for issue in result.errors))

    def test_unknown_header_is_error(self):
        result = validation.validate_field_plan(
            [FieldMapping(target="不存在", source="x")], {"姓名": "A"})
        self.assertFalse(result.ok)

    def test_records_required_and_duplicates(self):
        result = validation.validate_records(
            [{"姓名": ""}, {"姓名": "张三"}, {"姓名": "张三"}],
            required_fields=["姓名"])
        self.assertTrue(any("必填" in issue.message for issue in result.errors))
        self.assertTrue(any("完全相同" in issue.message for issue in result.warnings))

    def test_empty_records_is_error(self):
        self.assertFalse(validation.validate_records([]).ok)

    def test_mapping_entry_checks(self):
        ok = MappingEntry(
            source="姓名",
            target=TargetRef(kind="docx_table_cell", table=0, row=1, col=2))
        self.assertTrue(validation.validate_mapping_entry(ok, ["姓名"]).ok)

        bad = MappingEntry(
            source="姓名",
            target=TargetRef(kind="docx_table_cell", table=-1, row=1, col=2))
        self.assertFalse(validation.validate_mapping_entry(bad, ["姓名"]).ok)

        bad_xlsx = MappingEntry(source="姓名", target=TargetRef(kind="xlsx_cell", cell=""))
        self.assertFalse(validation.validate_mapping_entry(bad_xlsx, ["姓名"]).ok)

    def test_cross_mapping_warns_on_incompatible_target(self):
        entries = [MappingEntry(source="姓名",
                                target=TargetRef(kind="xlsx_cell", cell="A1"))]
        result = validation.validate_cross_mappings(entries, ["姓名"], ["某个.docx"])
        self.assertTrue(any("没有适用的映射" in issue.message for issue in result.warnings))

    def test_output_same_as_source_is_warned(self):
        path = self._path("book.xlsx")
        open(path, "w").close()
        result = validation.validate_output(path, path)
        self.assertTrue(result.ok)
        self.assertTrue(result.warnings)


if __name__ == "__main__":
    unittest.main()
