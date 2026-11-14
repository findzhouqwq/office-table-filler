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


class BackupFallbackTests(unittest.TestCase):
    """A protected target folder must not abort the save."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="tf_fallback_")
        self.backup_root = os.path.join(self.tmp, "backups")

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _write(self, path, text):
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)

    def test_backup_falls_back_when_target_folder_fails(self):
        src = os.path.join(self.tmp, "book.xlsx")
        self._write(src, "original")

        import app.core.safeio as safeio

        real_copy = safeio.shutil.copy2

        def flaky_copy(from_path, to_path, *args, **kwargs):
            # Refuse anything written next to the source, emulate CFA.
            if os.path.dirname(os.path.abspath(to_path)) == os.path.dirname(src):
                raise PermissionError(13, "Permission denied", to_path)
            return real_copy(from_path, to_path, *args, **kwargs)

        original_default = safeio.default_backup_dir
        safeio.shutil.copy2 = flaky_copy
        safeio.default_backup_dir = lambda: self.backup_root
        try:
            backup = safeio.backup_file(src)
        finally:
            safeio.shutil.copy2 = real_copy
            safeio.default_backup_dir = original_default

        self.assertTrue(backup.startswith(self.backup_root))
        self.assertTrue(os.path.isfile(backup))

    def test_save_with_backup_survives_backup_failure(self):
        path = os.path.join(self.tmp, "book.xlsx")
        self._write(path, "v1")
        warnings: list[str] = []

        import app.core.safeio as safeio

        real_backup = safeio.backup_file
        safeio.backup_file = lambda *a, **k: (_ for _ in ()).throw(
            PermissionError(13, "Permission denied"))
        try:
            target, backup = safeio.save_with_backup(
                path, lambda tmp: self._write(tmp, "v2"), warnings=warnings)
        finally:
            safeio.backup_file = real_backup

        self.assertEqual(backup, "")
        self.assertEqual(len(warnings), 1)
        self.assertIn("备份", warnings[0])
        with open(path, encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "v2")

    def test_atomic_write_falls_back_off_folder(self):
        target = os.path.join(self.tmp, "guarded", "out.txt")
        os.makedirs(os.path.dirname(target))

        import app.core.safeio as safeio

        real_probe = safeio._probe_cache
        safeio._probe_cache = lambda directory: False
        try:
            safeio.atomic_write(target, lambda tmp: self._write(tmp, "payload"))
        finally:
            safeio._probe_cache = real_probe

        with open(target, encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "payload")

    def test_write_probe_uses_fast_open(self):
        """The probe must not use tempfile.mkstemp (it can hang when blocked)."""
        from app.core.safeio import _write_probe

        self.assertTrue(_write_probe(self.tmp))
        missing_drive = os.path.join(self.tmp, "no-such-dir", "deeper")
        self.assertTrue(_write_probe(missing_drive))
        self.assertFalse(_write_probe(""))


if __name__ == "__main__":
    unittest.main()
