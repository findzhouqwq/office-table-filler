"""Atomic writes and timestamped backups.

Every write goes to a temporary file in the destination directory and is then
moved into place with :func:`os.replace`, so a failed save never leaves a
half-written workbook behind.  When the destination already exists it is
copied to a timestamped ``.bak-*`` file first.
"""

from __future__ import annotations

import datetime as _dt
import json
import logging
import os
import shutil
import tempfile
from typing import Callable

from .errors import FillError

# Resolved once so repeated saves do not re-probe the filesystem.
_BACKUP_DIR_CACHE: dict[str, str] = {}
# directory -> bool, so a protected folder is probed at most once per run.
_PROBE_CACHE: dict[str, bool] = {}


def _probe_cache(directory: str) -> bool:
    """Cached :func:`_write_probe` — never probe the same folder twice."""
    key = os.path.abspath(directory)
    if key not in _PROBE_CACHE:
        _PROBE_CACHE[key] = _write_probe(key)
    return _PROBE_CACHE[key]


def _timestamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def _backup_candidates(directory: str, stem: str, ext: str, stamp: str) -> str:
    candidate = os.path.join(directory, f"{stem}.bak-{stamp}{ext}")
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{stem}.bak-{stamp}-{counter}{ext}")
        counter += 1
    return candidate


def _copy_to_backup(src: str, directory: str, stem: str, ext: str,
                    stamp: str) -> str:
    os.makedirs(directory, exist_ok=True)
    candidate = _backup_candidates(directory, stem, ext, stamp)
    shutil.copy2(src, candidate)
    return candidate


def backup_file(path: str, backup_dir: str = "") -> str:
    """Copy *path* to a timestamped backup and return the backup path.

    Returns an empty string when *path* does not exist (nothing to back up).

    Writes next to the original first.  If that location is not writable
    (Controlled Folder Access, read-only folders, security software), it
    falls back to a per-user backup folder picked by :func:`default_backup_dir`.
    Raises :class:`OSError` only when every location failed.
    """
    src = os.path.abspath(path)
    if not os.path.isfile(src):
        return ""
    stem, ext = os.path.splitext(os.path.basename(src))
    stamp = _timestamp()

    directories: list[str] = []
    if backup_dir:
        directories.append(os.path.abspath(backup_dir))
    directories.append(os.path.dirname(src))
    fallback = default_backup_dir()
    if fallback not in directories:
        directories.append(fallback)

    last_error: OSError | None = None
    for directory in directories:
        try:
            return _copy_to_backup(src, directory, stem, ext, stamp)
        except OSError as exc:
            last_error = exc
            continue
    raise last_error or OSError("无法创建备份文件")


def default_backup_dir() -> str:
    """Return a writable folder used when the target folder is protected.

    Candidates are tried in order of preference; the first one that accepts a
    new file wins.  ``LOCALAPPDATA`` is preferred for normal desktop use, and
    the system temp folder is the last resort for locked-down machines where
    even app-data is protected (Controlled Folder Access, managed devices,
    sandboxed runtimes).
    """
    cached = _BACKUP_DIR_CACHE.get("path")
    if cached and os.path.isdir(cached):
        return cached

    candidates: list[str] = []
    local = os.environ.get("LOCALAPPDATA")
    if local:
        candidates.append(os.path.join(local, "TableFiller", "backups"))
    roaming = os.environ.get("APPDATA")
    if roaming:
        candidates.append(os.path.join(roaming, "TableFiller", "backups"))
    candidates.append(os.path.join(os.path.expanduser("~"), "TableFiller", "backups"))
    candidates.append(os.path.join(tempfile.gettempdir(), "TableFiller", "backups"))

    for candidate in candidates:
        if _write_probe(candidate):
            _BACKUP_DIR_CACHE["path"] = candidate
            return candidate
    # Nothing writable: return the last candidate so the caller reports a real
    # path in the error message instead of an empty string.
    return candidates[-1]


def _write_probe(directory: str, name: str = ".tablefiller-probe") -> bool:
    """Return True when a new file can be created in *directory*.

    Deliberately uses the builtin ``open`` rather than ``tempfile.mkstemp``:
    security software that blocks a protected folder can make ``mkstemp``
    hang instead of failing, while a plain create returns ``PermissionError``
    immediately.  Failures of any kind are treated as "not writable".
    """
    if not directory:
        return False
    probe = os.path.join(directory, f"{name}-{os.getpid()}.tmp")
    try:
        os.makedirs(directory, exist_ok=True)
        with open(probe, "wb") as fh:
            fh.write(b"")
        return True
    except (OSError, ValueError):
        return False
    finally:
        try:
            if os.path.exists(probe):
                os.remove(probe)
        except OSError:
            pass


def atomic_write(target: str, writer: Callable[[str], None]) -> str:
    """Write through *writer* into a temp file, then atomically replace *target*.

    *writer* receives a temporary path that keeps the target's extension, which
    matters for libraries such as openpyxl that infer the format from it.

    The temp file is normally created next to *target* so ``os.replace`` stays
    atomic.  When that folder denies new files (Controlled Folder Access,
    read-only location), the temp file is built in the system temp folder and
    copied over, so the write still succeeds instead of aborting.
    """
    target = os.path.abspath(target)
    directory = os.path.dirname(target)
    suffix = os.path.splitext(target)[1] or ".tmp"

    if directory and _probe_cache(directory):
        fd, tmp = tempfile.mkstemp(prefix=".tablefiller-", suffix=suffix, dir=directory)
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

    # Fall back: build elsewhere, then copy over the existing file in place.
    fd, tmp = tempfile.mkstemp(prefix=".tablefiller-", suffix=suffix)
    os.close(fd)
    try:
        writer(tmp)
        with open(tmp, "rb") as src_fh, open(target, "wb") as dst_fh:
            shutil.copyfileobj(src_fh, dst_fh)
    except BaseException:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    try:
        os.remove(tmp)
    except OSError:
        pass
    return target


def save_with_backup(target: str, writer: Callable[[str], None],
                     backup: bool = True, backup_dir: str = "",
                     warnings: list[str] | None = None) -> tuple[str, str]:
    """Atomically save *target*, backing up an existing file first.

    Returns ``(target, backup_path)`` where ``backup_path`` is empty when no
    backup was needed or none could be made.

    A failed backup is **not fatal**: the save still proceeds and a note is
    appended to *warnings* (or logged) so the caller can surface it.
    """
    target = os.path.abspath(target)
    backup_path = ""
    if backup and os.path.isfile(target):
        try:
            backup_path = backup_file(target, backup_dir)
        except OSError as exc:
            message = (
                f"无法创建备份文件（{exc.strerror or exc}），已跳过备份继续保存。"
                f"原文件：{target}"
            )
            if warnings is not None:
                warnings.append(message)
            else:
                logging.getLogger(__name__).warning(message)
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
