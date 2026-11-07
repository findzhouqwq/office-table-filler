import json
import os
import shutil
import tempfile
import unittest

from app.core.safeio import (atomic_write, backup_file, save_with_backup,
                             write_json)


class SafeIoTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="tf_safeio_")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _path(self, name):
        return os.path.join(self.tmp, name)

    def _write(self, path, text):
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)

    def test_backup_creates_copy(self):
        path = self._path("report.xlsx")
        self._write(path, "original")
        backup = backup_file(path)
        self.assertTrue(os.path.isfile(backup))
        self.assertNotEqual(os.path.abspath(backup), os.path.abspath(path))
        with open(backup, encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "original")

    def test_backup_missing_file_returns_empty(self):
        self.assertEqual(backup_file(self._path("missing.xlsx")), "")

    def test_atomic_write_replaces_content(self):
        path = self._path("data.txt")
        atomic_write(path, lambda tmp: self._write(tmp, "new"))
        with open(path, encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "new")
        leftovers = [n for n in os.listdir(self.tmp) if n.startswith(".tablefiller-")]
        self.assertEqual(leftovers, [])

    def test_failed_write_keeps_original(self):
        path = self._path("data.txt")
        self._write(path, "original")

        def failing(tmp):
            self._write(tmp, "partial")
            raise RuntimeError("boom")

        with self.assertRaises(RuntimeError):
            atomic_write(path, failing)
        with open(path, encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "original")
        leftovers = [n for n in os.listdir(self.tmp) if n.startswith(".tablefiller-")]
        self.assertEqual(leftovers, [])

    def test_save_with_backup_returns_paths(self):
        path = self._path("book.xlsx")
        self._write(path, "v1")
        target, backup = save_with_backup(
            path, lambda tmp: self._write(tmp, "v2"))
        self.assertEqual(os.path.abspath(target), os.path.abspath(path))
        self.assertTrue(os.path.isfile(backup))
        with open(path, encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "v2")

    def test_write_json_roundtrip(self):
        path = self._path("tpl.json")
        write_json(path, {"姓名": "张三", "n": 1})
        with open(path, encoding="utf-8") as fh:
            self.assertEqual(json.load(fh), {"姓名": "张三", "n": 1})


if __name__ == "__main__":
    unittest.main()
