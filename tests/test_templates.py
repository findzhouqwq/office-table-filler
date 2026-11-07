import os
import shutil
import tempfile
import unittest

from app.core import templates
from app.core.errors import DataError


class TemplateTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="tf_tpl_")
        self._old_home = os.environ.get("TABLE_FILLER_HOME")
        os.environ["TABLE_FILLER_HOME"] = self.tmp

    def tearDown(self):
        if self._old_home is None:
            os.environ.pop("TABLE_FILLER_HOME", None)
        else:
            os.environ["TABLE_FILLER_HOME"] = self._old_home
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_root_follows_override(self):
        self.assertTrue(templates.templates_root().startswith(self.tmp))

    def test_save_list_load(self):
        templates.save_template("月度台账", "text", {"target_path": "a.xlsx"})
        templates.save_template("发票模板", "image", {"mode": 1})
        listed = templates.list_templates("text")
        self.assertEqual([t.name for t in listed], ["月度台账"])
        loaded = templates.load_template("月度台账", "text")
        self.assertEqual(loaded.data["target_path"], "a.xlsx")
        self.assertEqual(loaded.kind, "text")

    def test_update_preserves_created(self):
        first = templates.save_template("模板", "image", {"mode": 0})
        second = templates.save_template("模板", "image", {"mode": 1})
        self.assertEqual(first.created, second.created)
        self.assertEqual(second.data["mode"], 1)

    def test_invalid_names_rejected(self):
        for name in ("", "  ", "..", "a/b", "a\\b", "a:b"):
            with self.assertRaises(DataError):
                templates.save_template(name, "text", {})

    def test_unknown_kind_rejected(self):
        with self.assertRaises(DataError):
            templates.save_template("x", "unknown", {})

    def test_export_and_import(self):
        templates.save_template("共享模板", "cross", {"mappings": []})
        dest = os.path.join(self.tmp, "exported.json")
        templates.export_template("共享模板", "cross", dest)
        self.assertTrue(os.path.isfile(dest))
        imported = templates.import_template(dest, kind="cross", name="导入模板")
        self.assertEqual(imported.name, "导入模板")
        loaded = templates.load_template("导入模板", "cross")
        self.assertEqual(loaded.data, {"mappings": []})

    def test_delete(self):
        templates.save_template("临时", "text", {})
        self.assertTrue(templates.delete_template("临时", "text"))
        self.assertFalse(templates.delete_template("临时", "text"))
        with self.assertRaises(DataError):
            templates.load_template("临时", "text")


if __name__ == "__main__":
    unittest.main()
