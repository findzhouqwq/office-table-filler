"""Spreadsheet engines.

Two engines share the same interface:

* :class:`OpenpyxlEngine` — edits ``.xlsx`` files directly with openpyxl.  This
  works for **Microsoft Office, WPS Office and OnlyOffice** files because all
  three read/write the same OOXML format.  Nothing has to be installed or
  running, and it is the default.
* :class:`ExcelComEngine` — drives a live Excel / WPS ET instance through
  Windows COM.  Preserves charts and everything else, but needs Office or WPS
  installed.  Selected automatically when available and preferred.
"""

from __future__ import annotations

import os
import re
from typing import Iterable, Sequence

from .errors import EngineError, FillError
from .safeio import backup_file, save_with_backup
from .utils import clean_text, norm_header

try:  # pragma: no cover - import guard
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter, column_index_from_string
    from openpyxl.worksheet.worksheet import Worksheet

    _HAS_OPENPYXL = True
except Exception:  # pragma: no cover
    _HAS_OPENPYXL = False


# Excel / WPS ET ProgIDs, best first.
_EXCEL_PROGIDS = ("Excel.Application", "ET.Application", "ket.Application", "wps.Application")


def excel_progid() -> str | None:
    """Return the first registered Excel/WPS spreadsheet ProgID, or ``None``.

    The Windows registry is probed instead of dispatching COM: creating an
    application instance only to test whether it exists would leak an orphan
    Excel process on every call (this function is called at startup and for
    every engine lookup).
    """
    try:
        import winreg
    except ImportError:  # not running on Windows
        return None
    for progid in _EXCEL_PROGIDS:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{progid}\CLSID"):
                return progid
        except OSError:
            continue
    return None


def available_engines() -> dict[str, bool]:
    return {
        "openpyxl": _HAS_OPENPYXL,
        "com": excel_progid() is not None,
    }


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------

class SpreadsheetEngine:
    """Common spreadsheet operations used by the fillers."""

    name = "base"

    # -- lifecycle ---------------------------------------------------------
    def open(self, path: str, keep_vba: bool = False) -> None:
        raise NotImplementedError

    def close(self) -> None:
        raise NotImplementedError

    def save(self, path: str | None = None, backup: bool = True) -> str:
        raise NotImplementedError

    # -- inspection --------------------------------------------------------
    def sheets(self) -> list[str]:
        raise NotImplementedError

    def active(self) -> str:
        raise NotImplementedError

    def set_active(self, name: str) -> None:
        raise NotImplementedError

    def dims(self, sheet: str = "") -> tuple[int, int]:
        raise NotImplementedError

    def used_range(self, sheet: str = "") -> tuple[int, int, int, int]:
        """Return ``(min_row, min_col, max_row, max_col)`` of non-empty cells."""
        raise NotImplementedError

    def cell_value(self, sheet: str, coord: str) -> object:
        raise NotImplementedError

    def cell_exists(self, sheet: str, coord: str) -> bool:
        raise NotImplementedError

    def header_map(self, sheet: str, row: int = 1) -> dict[str, str]:
        raise NotImplementedError

    # -- writing -----------------------------------------------------------
    def write_value(self, sheet: str, coord: str, value: object) -> None:
        raise NotImplementedError

    def write_row(self, sheet: str, row: int, values: Sequence[object]) -> None:
        raise NotImplementedError

    def write_rows(self, sheet: str, start_row: int, rows: Iterable[Sequence[object]]) -> int:
        raise NotImplementedError

    def append_rows(self, sheet: str, rows: Iterable[Sequence[object]]) -> int:
        raise NotImplementedError

    # -- helpers -----------------------------------------------------------
    def sheet_name(self, sheet: str) -> str:
        if not sheet:
            return self.active()
        names = self.sheets()
        if sheet in names:
            return sheet
        lowered = {n.lower(): n for n in names}
        if sheet.lower() in lowered:
            return lowered[sheet.lower()]
        raise FillError(f"找不到工作表：{sheet}（可用：{', '.join(names)}）")

    def resolve_coord(self, ref: str) -> tuple[str, str]:
        """Split ``Sheet1!A1`` into ``(sheet, coord)``."""
        ref = (ref or "").strip()
        if "!" in ref:
            sheet, coord = ref.split("!", 1)
            return sheet.strip().strip("'"), coord.strip()
        return "", ref


# ---------------------------------------------------------------------------
# openpyxl implementation
# ---------------------------------------------------------------------------

_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_MAX_COLUMN = 16384
_MAX_ROW = 1048576


def _validate_coord(coord: str) -> None:
    """Reject coordinates outside the range Excel can represent."""
    match = _COORD_RE.match(coord or "")
    if not match:
        raise FillError(f"非法单元格地址：{coord}")
    if column_index_from_string(match.group(1).upper()) > _MAX_COLUMN:
        raise FillError(f"单元格地址超出 Excel 列范围：{coord}")
    if int(match.group(2)) > _MAX_ROW:
        raise FillError(f"单元格地址超出 Excel 行范围：{coord}")


def _set_cell_value(cell, value: object) -> None:
    """Assign *value* without letting untrusted text become a live formula.

    OCR output and imported data may start with ``=`` (e.g.
    ``=HYPERLINK(...)``).  openpyxl treats such strings as formulas, which is a
    spreadsheet-injection risk, so force them to be plain text.
    """
    cell.value = value
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"


def _set_com_value(cell, value: object) -> None:
    """COM counterpart of :func:`_set_cell_value`."""
    if isinstance(value, str) and value.startswith("="):
        cell.NumberFormat = "@"
    cell.Value = value


class OpenpyxlEngine(SpreadsheetEngine):
    name = "openpyxl"

    def __init__(self) -> None:
        if not _HAS_OPENPYXL:
            raise EngineError("未安装 openpyxl，无法使用直接读写引擎。")
        self.path = ""
        self._wb = None
        self._active = ""

    # -- lifecycle ---------------------------------------------------------
    def open(self, path: str, keep_vba: bool = False) -> None:
        if not os.path.exists(path):
            raise EngineError(f"文件不存在：{path}")
        self.path = os.path.abspath(path)
        self._wb = load_workbook(
            self.path,
            data_only=False,
            keep_vba=keep_vba or self.path.lower().endswith(".xlsm"),
        )
        self._active = self._wb.active.title if self._wb.active else ""

    def close(self) -> None:
        if self._wb is not None:
            self._wb.close()
            self._wb = None

    def save(self, path: str | None = None, backup: bool = True) -> str:
        if self._wb is None:
            raise FillError("工作簿尚未打开。")
        target = os.path.abspath(path or self.path)
        ensure_parent(target)
        save_with_backup(target, lambda tmp: self._wb.save(tmp), backup=backup)
        self.path = target
        return target

    # -- inspection --------------------------------------------------------
    def sheets(self) -> list[str]:
        return list(self._wb.sheetnames)

    def active(self) -> str:
        return self._active

    def set_active(self, name: str) -> None:
        real = self.sheet_name(name)
        self._wb.active = self._wb.index(self._wb[real])
        self._active = real

    def _ws(self, sheet: str) -> Worksheet:
        return self._wb[self.sheet_name(sheet)]

    def dims(self, sheet: str = "") -> tuple[int, int]:
        ws = self._ws(sheet)
        return ws.max_row, ws.max_column

    def used_range(self, sheet: str = "") -> tuple[int, int, int, int]:
        ws = self._ws(sheet)
        return ws.min_row, ws.min_column, ws.max_row, ws.max_column

    def cell_value(self, sheet: str, coord: str) -> object:
        ws = self._ws(sheet)
        _validate_coord(coord)
        return ws[coord].value

    def cell_exists(self, sheet: str, coord: str) -> bool:
        ws = self._ws(sheet)
        if not _COORD_RE.match(coord or ""):
            return False
        value = ws[coord].value
        return value is not None and str(value).strip() != ""

    def header_map(self, sheet: str, row: int = 1) -> dict[str, str]:
        ws = self._ws(sheet)
        mapping: dict[str, str] = {}
        used: dict[str, int] = {}
        for col in range(1, ws.max_column + 1):
            value = ws.cell(row=row, column=col).value
            if value is None or str(value).strip() == "":
                continue
            key = norm_header(value)
            if key in used:
                used[key] += 1
                mapping[key] = get_column_letter(col)
            else:
                used[key] = 1
                mapping[key] = get_column_letter(col)
        return mapping

    # -- writing -----------------------------------------------------------
    def write_value(self, sheet: str, coord: str, value: object) -> None:
        ws = self._ws(sheet)
        _validate_coord(coord)
        _set_cell_value(ws[coord], value)

    def write_row(self, sheet: str, row: int, values: Sequence[object]) -> None:
        ws = self._ws(sheet)
        for offset, value in enumerate(values, start=1):
            _set_cell_value(ws.cell(row=row, column=offset), value)

    def write_rows(self, sheet: str, start_row: int, rows: Iterable[Sequence[object]]) -> int:
        ws = self._ws(sheet)
        written = 0
        for i, row_values in enumerate(rows):
            self.write_row(sheet, start_row + i, row_values)
            written += 1
        return written

    def append_rows(self, sheet: str, rows: Iterable[Sequence[object]]) -> int:
        ws = self._ws(sheet)
        start = ws.max_row + 1
        # max_row is unreliable for empty sheets; normalise to 1 when blank.
        if ws.max_row <= 1 and all(
            ws.cell(row=1, column=c).value in (None, "") for c in range(1, ws.max_column + 1)
        ):
            start = 1
        return self.write_rows(sheet, start, rows)


# ---------------------------------------------------------------------------
# COM implementation (live Excel / WPS)
# ---------------------------------------------------------------------------

class ExcelComEngine(SpreadsheetEngine):
    name = "com"

    def __init__(self, progid: str | None = None) -> None:
        self.progid = progid or excel_progid()
        if not self.progid:
            raise EngineError("未检测到已安装的 Excel / WPS 表格程序，无法使用 COM 引擎。")
        self.path = ""
        self._app = None
        self._wb = None
        self._visible = False
        self._com_initialized = False

    def open(self, path: str, keep_vba: bool = False) -> None:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore

        if not os.path.exists(path):
            raise EngineError(f"文件不存在：{path}")
        pythoncom.CoInitialize()
        self._com_initialized = True
        try:
            # DispatchEx creates a private instance, so Quit() will never close
            # a workbook the user has open in their own Excel/WPS window.
            self._app = win32com.client.DispatchEx(self.progid)
            try:
                self._app.Visible = False
                self._app.DisplayAlerts = False
            except Exception:
                pass
            self.path = os.path.abspath(path)
            self._wb = self._app.Workbooks.Open(self.path)
            self._visible = False
        except Exception:
            self.close()
            raise

    def close(self) -> None:
        try:
            if self._wb is not None:
                try:
                    self._wb.Close(SaveChanges=False)
                except Exception:
                    pass
                self._wb = None
            if self._app is not None:
                try:
                    self._app.Quit()
                except Exception:
                    pass
                self._app = None
        finally:
            if self._com_initialized:
                try:
                    import pythoncom  # type: ignore

                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self._com_initialized = False

    def save(self, path: str | None = None, backup: bool = True) -> str:
        if self._wb is None:
            raise FillError("工作簿尚未打开。")
        target = os.path.abspath(path or self.path)
        ensure_parent(target)
        if backup and os.path.isfile(target):
            try:
                backup_file(target)
            except OSError as exc:
                raise FillError(f"无法创建备份文件：{exc}") from exc
        self._wb.SaveAs(target)
        self.path = target
        return target

    def sheets(self) -> list[str]:
        return [s.Name for s in self._wb.Worksheets]

    def active(self) -> str:
        return self._wb.ActiveSheet.Name

    def set_active(self, name: str) -> None:
        self._wb.Worksheets(self.sheet_name(name)).Activate()

    def _ws(self, sheet: str):
        return self._wb.Worksheets(self.sheet_name(sheet))

    def dims(self, sheet: str = "") -> tuple[int, int]:
        ws = self._ws(sheet)
        rng = ws.UsedRange
        return int(rng.Rows.Count), int(rng.Columns.Count)

    def used_range(self, sheet: str = "") -> tuple[int, int, int, int]:
        ws = self._ws(sheet)
        rng = ws.UsedRange
        return int(rng.Row), int(rng.Column), int(rng.Row + rng.Rows.Count - 1), int(rng.Column + rng.Columns.Count - 1)

    def cell_value(self, sheet: str, coord: str) -> object:
        return self._ws(sheet).Range(coord).Value

    def cell_exists(self, sheet: str, coord: str) -> bool:
        value = self._ws(sheet).Range(coord).Value
        return value is not None and str(value).strip() != ""

    def header_map(self, sheet: str, row: int = 1) -> dict[str, str]:
        ws = self._ws(sheet)
        mapping: dict[str, str] = {}
        for col in range(1, self.dims(sheet)[1] + 1):
            value = ws.Cells(row, col).Value
            if value is None or str(value).strip() == "":
                continue
            mapping.setdefault(norm_header(value), col)
        from openpyxl.utils import get_column_letter  # noqa: WPS433

        return {k: get_column_letter(v) for k, v in mapping.items()}

    def write_value(self, sheet: str, coord: str, value: object) -> None:
        _validate_coord(coord)
        _set_com_value(self._ws(sheet).Range(coord), value)

    def write_row(self, sheet: str, row: int, values: Sequence[object]) -> None:
        ws = self._ws(sheet)
        for offset, value in enumerate(values, start=1):
            _set_com_value(ws.Cells(row, offset), value)

    def write_rows(self, sheet: str, start_row: int, rows: Iterable[Sequence[object]]) -> int:
        written = 0
        for i, row_values in enumerate(rows):
            self.write_row(sheet, start_row + i, row_values)
            written += 1
        return written

    def append_rows(self, sheet: str, rows: Iterable[Sequence[object]]) -> int:
        ws = self._ws(sheet)
        start = ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count, 1).Row
        return self.write_rows(sheet, start, rows)


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def get_spreadsheet_engine(prefer: str = "auto") -> SpreadsheetEngine:
    """Create an engine.

    ``prefer`` is ``auto`` (COM when available, else openpyxl), ``direct`` or
    ``com``.
    """
    if prefer in ("auto", "com"):
        progid = excel_progid()
        if progid:
            try:
                return ExcelComEngine(progid)
            except Exception as exc:
                if prefer == "com":
                    raise EngineError(f"COM 引擎启动失败：{exc}") from exc
    try:
        return OpenpyxlEngine()
    except Exception as exc:
        if prefer == "direct":
            raise
        raise EngineError(f"直接读写引擎不可用：{exc}") from exc


def ensure_parent(path: str) -> None:
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent, exist_ok=True)
