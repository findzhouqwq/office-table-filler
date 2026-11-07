"""Read source material (xlsx / csv / txt / json / docx / pptx) into models.

Everything here is read-only; writing is handled by the engines.
"""

from __future__ import annotations

import csv
import json
import os
from dataclasses import dataclass, field

from .errors import DataError
from .spreadsheet import OpenpyxlEngine, get_spreadsheet_engine
from .utils import clean_text, detect_delimiter, read_text_file, split_lines


# ---------------------------------------------------------------------------
# Structured table model
# ---------------------------------------------------------------------------

@dataclass
class TableData:
    headers: list[str] = field(default_factory=list)
    rows: list[list[object]] = field(default_factory=list)
    source: str = ""

    def as_records(self) -> list[dict]:
        records: list[dict] = []
        for row in self.rows:
            record = {}
            for i, head in enumerate(self.headers):
                record[head] = row[i] if i < len(row) else ""
            records.append(record)
        return records

    def __len__(self) -> int:
        return len(self.rows)


# ---------------------------------------------------------------------------
# Delimited text
# ---------------------------------------------------------------------------

def load_delimited(path: str, encoding: str | None = None, delimiter: str | None = None,
                   header_row: int = 1) -> TableData:
    text = read_text_file(path, encoding)
    if not text.strip():
        raise DataError(f"文件为空：{path}")
    delim = delimiter or detect_delimiter(text)
    lines = text.splitlines()
    # parse with csv so quoted fields survive
    rows = list(csv.reader(lines, delimiter=delim))
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    if not rows:
        raise DataError(f"文件没有可用的数据行：{path}")
    if header_row <= 1:
        headers = [clean_text(c) or f"列{i + 1}" for i, c in enumerate(rows[0])]
        data = rows[1:]
    else:
        headers = [clean_text(c) or f"列{i + 1}" for i, c in enumerate(rows[header_row - 1])]
        data = rows[header_row:]
    return TableData(headers=headers, rows=[list(r) for r in data], source=path)


def load_json_table(path: str) -> TableData:
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
    except json.JSONDecodeError as exc:
        raise DataError(f"JSON 解析失败：{exc}") from exc
    if isinstance(data, dict):
        # {"headers": [...], "rows": [[...]]} or {"name": "value", ...}
        if "rows" in data:
            headers = [str(h) for h in data.get("headers", [])]
            rows = [[r if not isinstance(r, dict) else r.get(h, "") for h in headers] for r in data["rows"]]
            return TableData(headers or [f"列{i + 1}" for i in range(len(rows[0]) if rows else 0)], rows, path)
        data = [data]
    if not isinstance(data, list) or not data:
        raise DataError("JSON 顶层应为数组或带 rows 的对象。")
    if all(isinstance(item, dict) for item in data):
        headers: list[str] = []
        for item in data:
            for key in item:
                if key not in headers:
                    headers.append(str(key))
        rows = [[item.get(h, "") for h in headers] for item in data]
        return TableData(headers, rows, path)
    rows = [item if isinstance(item, list) else [item] for item in data]
    width = max(len(r) for r in rows)
    return TableData([f"列{i + 1}" for i in range(width)], rows, path)


# ---------------------------------------------------------------------------
# Spreadsheets
# ---------------------------------------------------------------------------

def load_spreadsheet(path: str, sheet: str = "", header_row: int = 1,
                     engine: str = "auto") -> TableData:
    eng = get_spreadsheet_engine(engine)
    try:
        eng.open(path)
        names = eng.sheets()
        target = sheet or names[0]
        if target not in names:
            raise DataError(f"找不到工作表：{target}")
        max_row, max_col = eng.dims(target)
        if header_row > 1:
            headers = [str(eng.cell_value(target, _col_letter(c)) or "") for c in range(1, max_col + 1)]
            headers = [clean_text(h) or f"列{i + 1}" for i, h in enumerate(headers)]
            start = header_row + 1
        else:
            headers = []
            for c in range(1, max_col + 1):
                value = eng.cell_value(target, _col_letter(c))
                headers.append(clean_text(value) or f"列{i + 1}")
            start = 2
        rows: list[list[object]] = []
        for r in range(start, max_row + 1):
            row = [eng.cell_value(target, f"{_col_letter(c)}{r}") for c in range(1, max_col + 1)]
            if all(v is None or str(v).strip() == "" for v in row):
                continue
            rows.append(row)
        return TableData(headers, rows, path)
    finally:
        eng.close()


# ---------------------------------------------------------------------------
# Unified loader
# ---------------------------------------------------------------------------

def load_table(path: str, sheet: str = "", header_row: int = 1,
               encoding: str | None = None) -> TableData:
    if not os.path.exists(path):
        raise DataError(f"文件不存在：{path}")
    lower = path.lower()
    if lower.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
        return load_spreadsheet(path, sheet=sheet, header_row=header_row)
    if lower.endswith(".csv"):
        return load_delimited(path, encoding=encoding, header_row=header_row)
    if lower.endswith((".tsv", ".txt")):
        return load_delimited(path, encoding=encoding, header_row=header_row)
    if lower.endswith(".json"):
        return load_json_table(path)
    raise DataError(f"不支持的源文件类型：{os.path.basename(path)}")


# ---------------------------------------------------------------------------
# Word / PowerPoint models (for cross-file fill targets)
# ---------------------------------------------------------------------------

@dataclass
class DocxModel:
    path: str
    paragraphs: list[str] = field(default_factory=list)
    tables: list[list[list[str]]] = field(default_factory=list)

    @property
    def table_count(self) -> int:
        return len(self.tables)


@dataclass
class SlideModel:
    index: int
    shapes: list[tuple[str, str]] = field(default_factory=list)
    tables: list[list[list[str]]] = field(default_factory=list)


@dataclass
class PptxModel:
    path: str
    slides: list[SlideModel] = field(default_factory=list)


def load_docx(path: str) -> DocxModel:
    try:
        import docx  # type: ignore
    except Exception as exc:
        raise DataError("未安装 python-docx，无法读取 Word 文件。") from exc
    document = docx.Document(path)
    model = DocxModel(path=os.path.abspath(path))
    model.paragraphs = [p.text for p in document.paragraphs]
    for table in document.tables:
        rows: list[list[str]] = []
        for row in table.rows:
            rows.append([cell.text for cell in row.cells])
        model.tables.append(rows)
    return model


def load_pptx(path: str) -> PptxModel:
    try:
        from pptx import Presentation  # type: ignore
    except Exception as exc:
        raise DataError("未安装 python-pptx，无法读取 PPT 文件。") from exc
    prs = Presentation(path)
    model = PptxModel(path=os.path.abspath(path))
    for i, slide in enumerate(prs.slides):
        slide_model = SlideModel(index=i)
        for shape in slide.shapes:
            name = shape.name or f"形状{i}"
            if shape.has_text_frame:
                text = "\n".join(p.text for p in shape.text_frame.paragraphs)
                slide_model.shapes.append((name, text))
            if getattr(shape, "has_table", False) and shape.has_table:
                tbl = shape.table
                rows = [[cell.text for cell in row.cells] for row in tbl.rows]
                slide_model.tables.append(rows)
        model.slides.append(slide_model)
    return model


# ---------------------------------------------------------------------------
# Column letter helper (avoid importing openpyxl at module import time)
# ---------------------------------------------------------------------------

_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _col_letter(index: int) -> str:
    index = int(index)
    if index <= 0:
        raise DataError(f"列号必须为正数：{index}")
    result = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = _LETTERS[rem] + result
    return result
