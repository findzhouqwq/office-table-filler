"""Atomic writes and timestamped backups.

Every write goes to a temporary file in the destination directory and is then
moved into place with :func:`os.replace`, so a failed save never leaves a
half-written workbook behind.  When the destination already exists it is
copied to a timestamped ``.bak-*`` file first.
"""

from __future__ import annotations

import datetime as _dt
import json
import os
import shutil
import tempfile
from typing import Callable

from .errors import FillError


def _timestamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def backup_file(path: str, backup_dir: str = "") -> str:
    """Copy *path* to a timestamped backup and return the backup path.

    Returns an empty string when *path* does not exist (nothing to back up).
    """
    src = os.path.abspath(path)
    if not os.path.isfile(src):
        return ""
    directory = os.path.abspath(backup_dir) if backup_dir else os.path.dirname(src)
    os.makedirs(directory, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(src))
    candidate = os.path.join(directory, f"{stem}.bak-{_timestamp()}{ext}")
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{stem}.bak-{_timestamp()}-{counter}{ext}")
        counter += 1
    shutil.copy2(src, candidate)
    return candidate


def atomic_write(target: str, writer: Callable[[str], None]) -> str:
    """Write through *writer* into a temp file, then atomically replace *target*.

    *writer* receives a temporary path that keeps the target's extension, which
    matters for libraries such as openpyxl that infer the format from it.
    """
    target = os.path.abspath(target)
    directory = os.path.dirname(target)
    if directory:
        os.makedirs(directory, exist_ok=True)
    suffix = os.path.splitext(target)[1] or ".tmp"
    fd, tmp = tempfile.mkstemp(prefix=".tablefiller-", suffix=suffix, dir=directory or None)
    os.close(fd)
    try:
        writer(tmp)
        os.replace(tmp, target)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    return target


def save_with_backup(target: str, writer: Callable[[str], None],
                     backup: bool = True, backup_dir: str = "") -> tuple[str, str]:
    """Atomically save *target*, backing up an existing file first.

    Returns ``(target, backup_path)`` where ``backup_path`` is empty when no
    backup was needed.
    """
    target = os.path.abspath(target)
    backup_path = ""
    if backup and os.path.isfile(target):
        try:
            backup_path = backup_file(target, backup_dir)
        except OSError as exc:
            raise FillError(f"无法创建备份文件：{exc}") from exc
    try:
        atomic_write(target, writer)
    except OSError as exc:
        raise FillError(f"保存文件失败：{exc}") from exc
    return target, backup_path


def write_json(path: str, data: object) -> str:
    """Atomically write *data* as UTF-8 JSON."""
    def _write(tmp: str) -> None:
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
            fh.write("\n")

    return atomic_write(path, _write)
