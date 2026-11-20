"""Copy rectangular cell blocks from a live Excel / WPS window.

The side panel turns every rectangle the user drags over a spreadsheet into a
:class:`CopyBlock` (via COM ``RangeFromPoint``) and later appends all blocks to
the end of the target workbook with :func:`append_blocks`.

Only the *reading* half needs Windows and a running Excel / WPS; writing uses
openpyxl and works everywhere, which keeps the core testable on Linux.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from typing import Sequence

from .errors import AppError, FillError

#: A selection larger than this is almost certainly a stray drag over a whole
#: column; refuse it instead of transferring millions of cells over COM.  The
#: limit also keeps memory and the openpyxl write time reasonable on low-end
#: office PCs (50k cells is already a large clipboard).
MAX_BLOCK_CELLS = 50_000

_MAX_CELL_TEXT = 32767
_MAX_EXCEL_ROW = 1048576
_MAX_EXCEL_COLUMN = 16384

# xlsx cannot store most C0 control characters; COM values coming from copied
# text occasionally contain them.
_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

#: Excel / WPS ET ProgIDs, best first (same order as the spreadsheet engine).
_PROGIDS = ("Excel.Application", "ET.Application", "ket.Application",
            "wps.Application")


class ScreenCopyError(AppError):
    """A screen selection could not be read from the spreadsheet application."""


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

def _is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def _cell_at(row: Sequence[object], index: int) -> object:
    return row[index] if index < len(row) else None


def trim_block(rows: Sequence[Sequence[object]]) -> tuple[tuple[object, ...], ...]:
    """Drop trailing blank rows / columns from a rectangular block.

    Trailing blanks are a side effect of dragging the selection slightly too
    far; removing them keeps the pasted block flush with the target table.
    """
    grid = [list(row) for row in rows]
    while grid and all(_is_blank(value) for value in grid[-1]):
        grid.pop()
    if not grid:
        return ()
    width = max(len(row) for row in grid)
    while width and all(_is_blank(_cell_at(row, width - 1)) for row in grid):
        width -= 1
    if not width:
        return ()
    return tuple(tuple(row[:width]) for row in grid)


def normalize_range_values(raw: object) -> tuple[tuple[object, ...], ...]:
    """Normalise a COM ``Range.Value`` result into rows of cells.

    COM returns a scalar for a single cell, a flat tuple for one row and a
    tuple of tuples for several rows; all three shapes end up rectangular here.
    """
    if raw is None:
        return ()
    if isinstance(raw, (tuple, list)):
        if not raw:
            return ()
        if all(isinstance(item, (tuple, list)) for item in raw):
            grid = [list(item) for item in raw]
        else:
            grid = [list(raw)]
    else:
        grid = [[raw]]
    width = max(len(row) for row in grid)
    return tuple(tuple(row) + (None,) * (width - len(row)) for row in grid)


@dataclass(frozen=True)
class CopyBlock:
    """One rectangular selection copied out of a spreadsheet window."""

    rows: tuple[tuple[object, ...], ...] = ()
    source_path: str = ""
    sheet: str = ""
    address: str = ""
    source_app: str = ""

    def __post_init__(self) -> None:
        object.__setattr__(self, "rows", trim_block(self.rows))
        for name in ("source_path", "sheet", "address", "source_app"):
            value = getattr(self, name)
            if value is None:
                object.__setattr__(self, name, "")

    @property
    def height(self) -> int:
        return len(self.rows)

    @property
    def width(self) -> int:
        return max((len(row) for row in self.rows), default=0)

    @property
    def is_empty(self) -> bool:
        return not self.rows

    @property
    def source(self) -> str:
        """Short ``file · sheet · A1:C5`` description for status messages."""
        name = re.split(r"[\\/]", self.source_path)[-1] if self.source_path else ""
        parts = [part for part in (name, self.sheet, self.address) if part]
        return " · ".join(parts)


@dataclass(frozen=True)
class BlockPasteReport:
    """What :func:`append_blocks` did."""

    path: str
    sheet: str
    start_row: int
    end_row: int
    blocks: int
    rows: int
    warning: str = ""


# ---------------------------------------------------------------------------
# Availability
# ---------------------------------------------------------------------------

def screen_copy_available() -> tuple[bool, str]:
    """Return ``(available, reason)`` for screen selection on this machine."""
    if not sys.platform.startswith("win"):
        return False, "屏幕框选仅支持 Windows。"
    try:
        import win32com.client  # noqa: F401 - availability probe
    except Exception:
        return False, "屏幕框选需要安装 pywin32，请先运行 pip install pywin32。"
    from .spreadsheet import excel_progid

    if excel_progid() is None:
        return False, "未检测到已安装的 Excel / WPS 表格。"
    return True, ""


# ---------------------------------------------------------------------------
# Reading the live spreadsheet window
# ---------------------------------------------------------------------------

def _com_initialize():
    """Initialise COM for the current (GUI) thread."""
    try:
        import pythoncom  # type: ignore

        pythoncom.CoInitialize()
        return pythoncom
    except Exception:
        return None


def _active_application():
    """Return the visible, running Excel / WPS application object."""
    try:
        import win32com.client
    except Exception as exc:  # pragma: no cover - Windows only
        raise ScreenCopyError(
            "屏幕框选需要安装 pywin32，请先运行 pip install pywin32。") from exc

    for progid in _PROGIDS:
        try:
            app = win32com.client.GetActiveObject(progid)
        except Exception:
            continue
        try:
            if not bool(app.Visible):
                continue
            if not int(app.Workbooks.Count):
                continue
        except Exception:
            # Property access failed: still try to use it, RangeFromPoint will
            # produce a clearer error if the object is unusable.
            pass
        return app
    raise ScreenCopyError(
        "没有找到正在运行的 Excel / WPS 表格，请先打开源表格再框选。")


def _column_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def read_screen_rect(left: int, top: int, right: int, bottom: int) -> CopyBlock:
    """Read the cells under a screen rectangle from the active spreadsheet.

    Coordinates are physical screen pixels, exactly what
    ``Window.RangeFromPoint`` expects (Qt mouse events are logical pixels on
    high-DPI screens, so the overlay passes ``GetCursorPos`` values).
    """
    left, right = sorted((int(left), int(right)))
    top, bottom = sorted((int(top), int(bottom)))
    com = _com_initialize()
    try:
        # Keep all COM proxies inside this call: the frame is gone before
        # CoUninitialize runs, so no interface pointer outlives the apartment.
        return _read_screen_rect(left, top, right, bottom)
    finally:
        if com is not None:
            try:
                com.CoUninitialize()
            except Exception:
                pass


def _read_screen_rect(left: int, top: int, right: int,
                      bottom: int) -> CopyBlock:
    app = _active_application()
    try:
        window = app.ActiveWindow
    except Exception as exc:  # pragma: no cover - Windows only
        raise ScreenCopyError("当前表格窗口不可用，请先激活要复制的表格。") from exc

    try:
        start = window.RangeFromPoint(left, top)
        end = window.RangeFromPoint(right, bottom)
    except Exception as exc:  # pragma: no cover - WPS may not support it
        raise ScreenCopyError(
            "当前程序不支持按屏幕位置读取单元格；WPS 请改用"
            " Microsoft Excel 后再框选。") from exc
    if start is None or end is None:
        raise ScreenCopyError(
            "框选区域没有落在单元格上，请选中表格内的单元格区域。")

    try:
        sheet = start.Worksheet
        sheet_name = str(sheet.Name)
        try:
            source_path = str(sheet.Parent.FullName or "")
        except Exception:
            source_path = ""
        try:
            source_app = str(app.Name or "")
        except Exception:
            source_app = ""
    except Exception as exc:  # pragma: no cover - Windows only
        raise ScreenCopyError("无法读取选区所在的工作表。") from exc

    row1, col1 = int(start.Row), int(start.Column)
    row2, col2 = int(end.Row), int(end.Column)
    row1, row2 = sorted((row1, row2))
    col1, col2 = sorted((col1, col2))
    rows = row2 - row1 + 1
    columns = col2 - col1 + 1
    if rows * columns > MAX_BLOCK_CELLS:
        raise ScreenCopyError(
            f"选区包含 {rows} 行 × {columns} 列，超过 "
            f"{MAX_BLOCK_CELLS} 个单元格，请分几次框选。")

    address = f"{_column_letter(col1)}{row1}:{_column_letter(col2)}{row2}"
    try:
        cell_range = sheet.Range(sheet.Cells(row1, col1),
                                 sheet.Cells(row2, col2))
        raw = cell_range.Value
    except Exception as exc:  # pragma: no cover - Windows only
        raise ScreenCopyError(f"读取选区 {address} 失败：{exc}") from exc

    values = normalize_range_values(raw)
    return CopyBlock(values, source_path, sheet_name, address, source_app)


# ---------------------------------------------------------------------------
# Appending blocks to the target workbook
# ---------------------------------------------------------------------------

def _clean_cell(value: object, address: str) -> object:
    """Make a copied value safe for openpyxl / xlsx."""
    if value is None:
        return None
    if isinstance(value, str):
        text = _CONTROL_RE.sub("", value)
        if len(text) > _MAX_CELL_TEXT:
            raise FillError(
                f"选区 {address} 中有内容超过 Excel 单元格上限"
                f"（{_MAX_CELL_TEXT} 字符），请缩小选区。")
        return text
    if isinstance(value, float) and value != value:  # NaN
        return None
    return value


def _has_any_value(ws) -> bool:
    """True when the worksheet contains at least one non-blank value."""
    cells = getattr(ws, "_cells", None)
    if cells is not None:
        for cell in cells.values():
            value = cell.value
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            return True
        return False
    for row in ws.iter_rows():
        for cell in row:
            value = cell.value
            if value is not None and (not isinstance(value, str) or value.strip()):
                return True
    return False


def append_blocks(path: str, blocks: Sequence[CopyBlock], sheet: str = "",
                   header_row: int = 1, output_path: str = "") -> BlockPasteReport:
    """Append *blocks* below the existing data of the target workbook.

    Blocks keep the order they were selected in and are stacked vertically;
    the target's existing rows and formatting are left untouched.  The file is
    saved atomically with a timestamped backup.
    """
    from .errors import EngineError
    from .field_detect import avoid_merged_rows, last_data_row
    from .spreadsheet import OpenpyxlEngine

    path = (path or "").strip()
    if not path:
        raise FillError("请先在主界面选择目标表格。")
    usable = [block for block in blocks if not block.is_empty]
    if not usable:
        raise FillError("复制的选区里没有数据，请重新框选。")
    header_row = max(1, int(header_row or 1))

    engine = OpenpyxlEngine()
    try:
        try:
            engine.open(path)
        except EngineError as exc:
            raise FillError(str(exc)) from exc
        except Exception as exc:  # noqa: BLE001 - corrupt / unsupported workbook
            raise FillError(f"无法打开目标表格：{exc}") from exc

        sheet_name = engine.sheet_name(sheet)
        _, min_col, max_row, max_col = engine.used_range(sheet_name)
        if header_row > max_row:
            raise FillError(
                f"工作表「{sheet_name}」第 {header_row} 行没有表头。")

        ws = engine._ws(sheet_name)
        row = last_data_row(ws, header_row, min_col, max_col) + 1
        if row == header_row + 1 and not _has_any_value(ws):
            # A completely empty sheet has no header to protect: start at A1.
            row = 1
        row = avoid_merged_rows(ws, row, min_col, max_col)
        start_row = row

        widest = max(block.width for block in usable)
        if min_col + widest - 1 > _MAX_EXCEL_COLUMN:
            raise FillError("复制的数据列数超过 Excel 上限，请缩小框选范围。")
        # Validate every block before touching the workbook so a bad cell
        # cannot leave a half-written in-memory sheet behind.
        prepared: list[tuple[tuple[object, ...], ...]] = []
        next_row = row
        for block in usable:
            # Accumulate so several blocks cannot each pass a check that only
            # looks at the first block's start row.
            if next_row + block.height - 1 > _MAX_EXCEL_ROW:
                raise FillError("目标工作表已写满，无法再追加数据。")
            prepared.append(tuple(
                tuple(_clean_cell(value, block.address) for value in values)
                for values in block.rows))
            next_row += block.height

        for rows in prepared:
            engine.write_block(sheet_name, row, min_col, rows)
            row += len(rows)

        warnings: list[str] = []
        try:
            target = engine.save(output_path or path, backup=True,
                                 warnings=warnings)
        except FillError as exc:
            raise FillError(
                f"{exc}\n如果目标文件正在 Excel / WPS 中打开，请先关闭它，"
                "或在主界面把输出路径设为另一个文件。") from exc
        return BlockPasteReport(
            target, sheet_name, start_row, row - 1, len(usable),
            row - start_row, warnings[0] if warnings else "")
    except OSError as exc:
        raise FillError(
            "无法写入目标表格：目标文件可能正被 Excel / WPS 占用，"
            "请先关闭它，或在主界面把输出路径设为另一个文件。"
            f"（{exc}）") from exc
    finally:
        engine.close()
