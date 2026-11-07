"""Plain data structures used to describe fill plans and results."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


# ---------------------------------------------------------------------------
# Cross-file target kinds and the file extensions they support
# ---------------------------------------------------------------------------

KIND_EXTENSIONS: dict[str, tuple[str, ...]] = {
    "xlsx_cell": (".xlsx", ".xlsm", ".xltx", ".xltm"),
    "xlsx_rows": (".xlsx", ".xlsm", ".xltx", ".xltm"),
    "docx_table_cell": (".docx", ".docm"),
    "docx_table_rows": (".docx", ".docm"),
    "docx_placeholder": (".docx", ".docm"),
    "pptx_table_cell": (".pptx", ".pptm"),
    "pptx_placeholder": (".pptx", ".pptm"),
}


# ---------------------------------------------------------------------------
# Mapping rules
# ---------------------------------------------------------------------------

@dataclass
class ExtractionRule:
    """How to pull one value out of free text / an OCR result.

    ``pattern_type`` selects the strategy:

    * ``label``    — text that follows a label, e.g. ``姓名：张三``
    * ``keyword``  — text on the line that contains a keyword
    * ``regex``    — the first regex capture group
    * ``line``     — the Nth non-empty line (``position``)
    * ``word``     — the Nth OCR word (``position``)
    * ``grid``     — filled by :mod:`app.core.table_detect` (no rule needed)
    """

    field: str
    pattern_type: str = "label"
    pattern: str = ""
    position: int = 0
    default: str = ""
    transform: str = "text"
    required: bool = False

    def to_dict(self) -> dict:
        return {
            "field": self.field,
            "pattern_type": self.pattern_type,
            "pattern": self.pattern,
            "position": self.position,
            "default": self.default,
            "transform": self.transform,
            "required": self.required,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "ExtractionRule":
        return cls(
            field=str(data.get("field", "")),
            pattern_type=str(data.get("pattern_type", "label")),
            pattern=str(data.get("pattern", "")),
            position=int(data.get("position", 0) or 0),
            default=str(data.get("default", "")),
            transform=str(data.get("transform", "text")),
            required=bool(data.get("required", False)),
        )


@dataclass
class FieldMapping:
    """Maps a source field onto one destination cell / column."""

    target: str
    source: str
    transform: str = "text"
    default: str = ""

    def to_dict(self) -> dict:
        return {
            "target": self.target,
            "source": self.source,
            "transform": self.transform,
            "default": self.default,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "FieldMapping":
        return cls(
            target=str(data.get("target", "")),
            source=str(data.get("source", "")),
            transform=str(data.get("transform", "text")),
            default=str(data.get("default", "")),
        )


# ---------------------------------------------------------------------------
# Fill options
# ---------------------------------------------------------------------------

@dataclass
class FillOptions:
    """Options that control how a fill run behaves."""

    sheet: str = ""                    # spreadsheet sheet name (empty = active)
    header_row: int = 1               # 1-based row holding headers (for column mapping)
    anchor: str = "A1"                # where grid/append writes start
    mode: str = "append"              # append | overwrite | fill_cells
    header_to_column: bool = True     # map by header name rather than literal column
    keep_formulas: bool = False        # preserve formulas when reading back
    write_blank_for_missing: bool = True
    verbose: bool = False


# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------

@dataclass
class FillRecord:
    """One written value, kept for the report view."""

    target: str
    value: str
    source: str = ""
    ok: bool = True
    message: str = ""
    file: str = ""


@dataclass
class FillReport:
    """Summary of a fill operation."""

    title: str = ""
    records: list[FillRecord] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    output_path: str = ""

    @property
    def written(self) -> int:
        return sum(1 for r in self.records if r.ok)

    @property
    def failed(self) -> int:
        return sum(1 for r in self.records if not r.ok)

    def add(self, record: FillRecord) -> None:
        self.records.append(record)

    def warn(self, message: str) -> None:
        self.warnings.append(message)

    def to_rows(self) -> list[list[str]]:
        return [
            [r.file, r.target, r.source, r.value, "OK" if r.ok else "失败", r.message]
            for r in self.records
        ]


# ---------------------------------------------------------------------------
# Target references (cross-file fill)
# ---------------------------------------------------------------------------

@dataclass
class TargetRef:
    """A single write destination.

    ``kind`` is one of:

    * ``xlsx_cell``        — ``sheet`` + ``cell`` (e.g. ``Sheet1!A1``)
    * ``xlsx_rows``        — append rows to a sheet (grid source)
    * ``docx_table_cell``  — ``table`` + ``row`` + ``col``
    * ``docx_table_rows``  — fill a docx table from rows (``row_start``)
    * ``pptx_table_cell``  — ``slide`` + ``table`` + ``row`` + ``col``
    * ``docx_placeholder`` — replace ``{{name}}`` in paragraphs
    * ``pptx_placeholder`` — replace ``{{name}}`` in shapes
    """

    kind: str
    file: str = ""
    sheet: str = ""
    cell: str = ""
    table: int = -1
    row: int = -1
    col: int = -1
    slide: int = -1
    row_start: int = 0
    placeholder: str = ""
    label: str = ""

    def describe(self) -> str:
        if self.kind == "xlsx_cell":
            return f"{self.sheet or '当前表'}!{self.cell}" if self.sheet else self.cell
        if self.kind == "xlsx_rows":
            return f"{self.sheet or '当前表'}!{self.cell or 'A1'}(追加行)"
        if self.kind == "docx_table_cell":
            return f"Word 表格{self.table + 1} 第{self.row + 1}行第{self.col + 1}列"
        if self.kind == "docx_table_rows":
            return f"Word 表格{self.table + 1} 从第{self.row_start + 1}行开始填"
        if self.kind == "pptx_table_cell":
            return f"PPT 第{self.slide + 1}页 表格{self.table + 1} {self.row + 1}x{self.col + 1}"
        if self.kind in ("docx_placeholder", "pptx_placeholder"):
            return f"{self.file or ''} 占位符 {self.placeholder}"
        return self.label or self.kind

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "file": self.file,
            "sheet": self.sheet,
            "cell": self.cell,
            "table": self.table,
            "row": self.row,
            "col": self.col,
            "slide": self.slide,
            "row_start": self.row_start,
            "placeholder": self.placeholder,
            "label": self.label,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "TargetRef":
        return cls(
            kind=str(data.get("kind", "")),
            file=str(data.get("file", "")),
            sheet=str(data.get("sheet", "")),
            cell=str(data.get("cell", "")),
            table=int(data.get("table", -1) or -1),
            row=int(data.get("row", -1) if data.get("row") is not None else -1),
            col=int(data.get("col", -1) if data.get("col") is not None else -1),
            slide=int(data.get("slide", -1) if data.get("slide") is not None else -1),
            row_start=int(data.get("row_start", 0) or 0),
            placeholder=str(data.get("placeholder", "")),
            label=str(data.get("label", "")),
        )


@dataclass
class MappingEntry:
    """One row of a cross-file mapping: source field -> target ref(s)."""

    source: str
    target: TargetRef
    transform: str = "text"
    default: str = ""

    def to_dict(self) -> dict:
        return {
            "source": self.source,
            "target": self.target.to_dict(),
            "transform": self.transform,
            "default": self.default,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MappingEntry":
        return cls(
            source=str(data.get("source", "")),
            target=TargetRef.from_dict(data.get("target") or {}),
            transform=str(data.get("transform", "text")),
            default=str(data.get("default", "")),
        )


@dataclass
class MappingDoc:
    """A serialisable mapping between source fields and targets."""

    entries: list[MappingEntry] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "entries": [e.to_dict() for e in self.entries],
            "meta": self.meta,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "MappingDoc":
        return cls(
            entries=[MappingEntry.from_dict(e) for e in data.get("entries", [])],
            meta=dict(data.get("meta") or {}),
        )
