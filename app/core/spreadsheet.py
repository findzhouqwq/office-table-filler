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
    """Return the first registered Excel/WPS spreadsheet ProgID, or ``None``."""
    for progid in _EXCEL_PROGIDS:
        try:
            import win32com.client  # noqa: F401
            from win32com.client import gencache  # noqa: F401

            import pythoncom

            pythoncom.CoInitialize()
            win32com.client.Dispatch(progid)
            pythoncom.CoUninitialize()
            return progid
        except Exception:
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

    def save(self, path: str | None = None) -> str:
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

_COORD_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d{1,7}$")


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

    def save(self, path: str | None = None) -> str:
        if self._wb is None:
            raise FillError("工作簿尚未打开。")
        target = os.path.abspath(path or self.path)
        ensure_parent(target)
        self._wb.save(target)
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
        if not _COORD_RE.match(coord or ""):
            raise FillError(f"非法单元格地址：{coord}")
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
        if not _COORD_RE.match(coord or ""):
            raise FillError(f"非法单元格地址：{coord}")
        ws[coord] = value

    def write_row(self, sheet: str, row: int, values: Sequence[object]) -> None:
        ws = self._ws(sheet)
        for offset, value in enumerate(values, start=1):
            ws.cell(row=row, column=offset, value=value)

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

    def open(self, path: str, keep_vba: bool = False) -> None:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore

        pythoncom.CoInitialize()
        self._app = win32com.client.Dispatch(self.progid)
        try:
            self._app.Visible = False
            self._app.DisplayAlerts = False
        except Exception:
            pass
        self.path = os.path.abspath(path)
        self._wb = self._app.Workbooks.Open(self.path)
        self._visible = False

    def close(self) -> None:
        try:
            if self._wb is not None:
                self._wb.Close(SaveChanges=False)
        finally:
            self._wb = None

    def save(self, path: str | None = None) -> str:
        target = os.path.abspath(path or self.path)
        ensure_parent(target)
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
        self._ws(sheet).Range(coord).Value = value

    def write_row(self, sheet: str, row: int, values: Sequence[object]) -> None:
        ws = self._ws(sheet)
        for offset, value in enumerate(values, start=1):
            ws.Cells(row, offset).Value = value

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
