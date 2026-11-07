"""High-level fill operations for the three features.

* :func:`fill_from_images`  — image OCR batch fill into a spreadsheet
* :func:`fill_from_text`    — text / records fill into a spreadsheet or table
* :func:`fill_cross_files`  — read one or more sources and write one or more targets
"""

from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field

from .errors import DataError, FillError
from .models import (ExtractionRule, FillReport, FillRecord, FieldMapping,
                     FillOptions, MappingEntry, TargetRef)
from .ocr import OcrEngine, get_ocr_engine
from .spreadsheet import OpenpyxlEngine, get_spreadsheet_engine
from .table_detect import build_grid_from_ocr, cluster_words_to_grid
from .utils import apply_transform, clean_text
from .mapping import extract_fields, parse_labeled_lines, resolve_column
from .data_io import TableData, load_table, load_docx, load_pptx


# ---------------------------------------------------------------------------
# Requests
# ---------------------------------------------------------------------------

@dataclass
class ImageFillRequest:
    images: list[str] = field(default_factory=list)
    target_path: str = ""
    output_path: str = ""
    sheet: str = ""
    mode: str = "grid"                 # grid | fields
    anchor: str = "A1"
    append: bool = True
    clear_sheet: bool = False
    header_row: int = 1
    ocr_engine: str = "auto"
    extraction_rules: list[ExtractionRule] = field(default_factory=list)
    column_plan: list[FieldMapping] = field(default_factory=list)
    max_cols: int = 30


@dataclass
class TextFillRequest:
    text: str = ""
    records: list[dict] = field(default_factory=list)
    target_path: str = ""
    output_path: str = ""
    sheet: str = ""
    header_row: int = 1
    append: bool = True
    clear_sheet: bool = False
    column_plan: list[FieldMapping] = field(default_factory=list)
    extraction_rules: list[ExtractionRule] = field(default_factory=list)
    delimiter: str = ""


@dataclass
class CrossFillRequest:
    source_paths: list[str] = field(default_factory=list)
    source_sheet: str = ""
    source_header_row: int = 1
    target_paths: list[str] = field(default_factory=list)
    mappings: list[MappingEntry] = field(default_factory=list)
    output_dir: str = ""
    batch_by_row: bool = True
    overwrite: bool = False
    sheet_for_target: str = ""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _col_letter(index: int) -> str:
    index = int(index)
    result = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[rem] + result
    return result


def _split_anchor(anchor: str) -> tuple[int, int]:
    match = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$", anchor or "A1")
    if not match:
        raise FillError(f"非法起始单元格：{anchor}")
    col = 0
    for ch in match.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    return int(match.group(2)), col


def _output_path(output: str, target: str, suffix: str = "") -> str:
    if output:
        return output
    base, ext = os.path.splitext(target)
    return base + (suffix or "") + ext


def _copy_to_output(src: str, output: str) -> str:
    if not output or os.path.abspath(output) == os.path.abspath(src):
        return src
    directory = os.path.dirname(os.path.abspath(output))
    if directory and not os.path.isdir(directory):
        os.makedirs(directory, exist_ok=True)
    shutil.copy2(src, output)
    return output


def _batch_output(output_dir: str, base: str, index: int, total: int) -> str:
    stem, ext = os.path.splitext(base)
    if total > 1:
        return os.path.join(output_dir, f"{stem}_{index + 1:03d}{ext}")
    return os.path.join(output_dir, base)


def _letter_to_index(letter: str) -> int:
    index = 0
    for ch in letter.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def _row_values_from_plan(plan: list[FieldMapping], fields: dict,
                          header_map: dict[str, str], width: int) -> list[object]:
    """Build a row where ``plan[i]`` writes the matching source field."""
    slots: list[tuple[int, FieldMapping]] = []
    if plan:
        for mapping in plan:
            try:
                col_letter = resolve_column(mapping.target, header_map)
            except Exception:
                col_letter = mapping.target.upper() if mapping.target else "A"
            slots.append((_letter_to_index(col_letter), mapping))
    elif fields:
        # No plan given: lay the fields out in order, one per column.
        slots = [(i + 1, FieldMapping(target="", source=key)) for i, key in enumerate(fields)]
    width = max(width, max([s for s, _ in slots], default=0))
    row: list[object] = ["" for _ in range(width)]
    for index, mapping in slots:
        value = str(fields.get(mapping.source, "") or "")
        if value == "" and mapping.default:
            value = mapping.default
        if index - 1 < width:
            row[index - 1] = apply_transform(value, mapping.transform)
    return row


# ---------------------------------------------------------------------------
# Feature 1: image OCR batch fill
# ---------------------------------------------------------------------------

def fill_from_images(request: ImageFillRequest,
                     ocr_engine: OcrEngine | None = None) -> FillReport:
    report = FillReport(title="图片识别批量填表")
    if not request.images:
        raise DataError("请至少选择一张图片。")
    if not request.target_path or not os.path.exists(request.target_path):
        raise DataError("请选择有效的目标表格文件。")

    engine_ocr = ocr_engine or get_ocr_engine(request.ocr_engine)
    engine = OpenpyxlEngine()
    engine.open(request.target_path)
    try:
        sheet = request.sheet or engine.active()
        engine.set_active(sheet)
        ws = engine._ws(sheet)
        if request.clear_sheet:
            ws.delete_rows(1, ws.max_row)
            ws.delete_cols(1, ws.max_column)

        anchor_row, anchor_col = _split_anchor(request.anchor)
        if request.append and anchor_row == 1 and anchor_col == 1:
            max_row, _ = engine.dims(sheet)
            anchor_row = max_row + 1 if max_row > 1 else 1
        header_map = engine.header_map(sheet, request.header_row)
        width = max(len(header_map), anchor_col, 1)

        cursor_row = anchor_row
        for image_path in request.images:
            result = engine_ocr.recognize(image_path)
            if request.mode == "fields":
                fields = extract_fields(result.text, request.extraction_rules)
                if not fields and not request.extraction_rules:
                    fields = parse_labeled_lines(result.text)
                row = _row_values_from_plan(request.column_plan, fields, header_map, width)
                engine.write_row(sheet, cursor_row, row)
                report.add(FillRecord(
                    target=f"{sheet}!{_col_letter(1)}{cursor_row}",
                    value=str(row), source=os.path.basename(image_path), file=image_path,
                ))
                cursor_row += 1
            else:
                result, prepared = engine_ocr.recognize_with_image(image_path)
                words = result.words_as_tuples()
                grid = build_grid_from_ocr(prepared, words)
                if not grid.used_lines and not grid.cells:
                    grid = cluster_words_to_grid(words)
                matrix = grid.matrix()
                if not matrix:
                    report.warn(f"{os.path.basename(image_path)}：未识别到表格内容。")
                    continue
                for row_index, row in enumerate(matrix):
                    trimmed = [clean_text(v) for v in row]
                    if not any(trimmed):
                        continue
                    engine.write_row(sheet, cursor_row, trimmed)
                    report.add(FillRecord(
                        target=f"{sheet}!{_col_letter(1)}{cursor_row}",
                        value=" | ".join(trimmed), source=os.path.basename(image_path),
                        file=image_path,
                    ))
                    cursor_row += 1
    finally:
        output = request.output_path or request.target_path
        engine.save(output)
        report.output_path = output
        engine.close()
    report.warn(f"共处理 {len(request.images)} 张图片，写入 {report.written} 处。")
    return report


# ---------------------------------------------------------------------------
# Feature 2: text fill
# ---------------------------------------------------------------------------

def _text_to_records(request: TextFillRequest) -> list[dict]:
    if request.records:
        return list(request.records)
    if not request.text.strip():
        raise DataError("请输入要填写的文本或选择数据文件。")
    blocks = re.split(r"\n\s*\n|^\s*[-=—]{3,}\s*$", request.text, flags=re.M)
    blocks = [b for b in blocks if b.strip()]
    records: list[dict] = []
    for block in blocks:
        if request.extraction_rules:
            records.append(extract_fields(block, request.extraction_rules))
        else:
            records.append(parse_labeled_lines(block) or {"内容": clean_text(block)})
    return records


def fill_from_text(request: TextFillRequest, engine: OpenpyxlEngine | None = None) -> FillReport:
    report = FillReport(title="文字填表")
    records = _text_to_records(request)
    if not records:
        raise DataError("没有可填写的数据。")
    if not request.target_path or not os.path.exists(request.target_path):
        raise DataError("请选择有效的目标表格文件。")

    eng = engine or OpenpyxlEngine()
    eng.open(request.target_path)
    try:
        sheet = request.sheet or eng.active()
        eng.set_active(sheet)
        ws = eng._ws(sheet)
        if request.clear_sheet:
            ws.delete_rows(1, ws.max_row)
            ws.delete_cols(1, ws.max_column)
        header_map = eng.header_map(sheet, request.header_row)
        width = max(len(header_map), 1)
        if request.append:
            start_row = ws.max_row + 1 if ws.max_row > 1 else 1
        else:
            start_row = 1
        plan = request.column_plan or [
            FieldMapping(target=h, source=h) for h in header_map.values()
        ]
        for offset, record in enumerate(records):
            row = _row_values_from_plan(plan, record, header_map, max(width, len(plan)))
            eng.write_row(sheet, start_row + offset, row)
            report.add(FillRecord(
                target=f"{sheet}!{_col_letter(1)}{start_row + offset}",
                value=" | ".join(str(v) for v in row),
                source=f"记录{offset + 1}", file=request.target_path,
            ))
    finally:
        output = request.output_path or request.target_path
        eng.save(output)
        report.output_path = output
        eng.close()
    return report


# ---------------------------------------------------------------------------
# Feature 3: cross-file fill
# ---------------------------------------------------------------------------

def _merge_sources(paths: list[str], sheet: str, header_row: int) -> TableData:
    combined = TableData(source=";".join(paths))
    for path in paths:
        table = load_table(path, sheet=sheet, header_row=header_row)
        if not combined.headers:
            combined.headers = list(table.headers)
            combined.rows = list(table.rows)
            continue
        # union headers by name
        for i, head in enumerate(table.headers):
            if head not in combined.headers:
                combined.headers.append(head)
                for row in combined.rows:
                    row.append("")
        for row in table.rows:
            new_row = ["" for _ in combined.headers]
            for i, head in enumerate(table.headers):
                if head in combined.headers:
                    new_row[combined.headers.index(head)] = row[i] if i < len(row) else ""
            combined.rows.append(new_row)
    return combined


def _set_docx_cell(cell, text: str) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    if paragraph.runs:
        paragraph.runs[0].text = text
    else:
        paragraph.add_run(text)


def _replace_in_runs(container, mapping: dict[str, str]) -> int:
    count = 0
    for paragraph in container.paragraphs:
        tokens = [t for k in mapping for t in ("{{" + k + "}}", "{" + k + "}")]
        if not any(t in paragraph.text for t in tokens):
            continue
        full = paragraph.text
        new = full
        for key, value in mapping.items():
            for token in ("{{" + key + "}}", "{" + key + "}"):
                new = new.replace(token, str(value))
        if new == full:
            continue
        runs = paragraph.runs
        if runs:
            runs[0].text = new
            for run in runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(new)
        count += 1
    return count


def _apply_docx(path: str, output: str, entries: list[MappingEntry],
                record_index: int, record: dict, report: FillReport) -> None:
    try:
        import docx  # type: ignore
    except Exception as exc:
        raise DataError("未安装 python-docx，无法写入 Word 文件。") from exc
    document = docx.Document(path)
    for entry in entries:
        ref = entry.target
        value = apply_transform(record.get(entry.source, entry.default), entry.transform)
        if ref.kind == "docx_table_cell":
            if ref.table < 0 or ref.table >= len(document.tables):
                raise FillError(f"Word 表格序号超出范围：{ref.table}")
            table = document.tables[ref.table]
            if ref.row < 0 or ref.row >= len(table.rows):
                raise FillError(f"Word 表格{ref.table + 1} 行号超出范围：{ref.row}")
            if ref.col < 0 or ref.col >= len(table.rows[ref.row].cells):
                raise FillError(f"Word 表格{ref.table + 1} 列号超出范围：{ref.col}")
            _set_docx_cell(table.rows[ref.row].cells[ref.col], value)
            report.add(FillRecord(target=ref.describe(), value=value,
                                  source=entry.source, file=output))
        elif ref.kind == "docx_table_rows":
            if ref.table < 0 or ref.table >= len(document.tables):
                raise FillError(f"Word 表格序号超出范围：{ref.table}")
            table = document.tables[ref.table]
            rows_data = record.get("_rows") or []
            if not isinstance(rows_data, list) or not rows_data:
                continue
            header = [cell.text for cell in table.rows[ref.row_start].cells]
            for r_offset, data_row in enumerate(rows_data):
                target_row = ref.row_start + 1 + r_offset
                if target_row >= len(table.rows):
                    table.add_row()
                cells = table.rows[target_row].cells
                for c_index, head in enumerate(header):
                    if c_index >= len(cells):
                        break
                    _set_docx_cell(cells[c_index], str(data_row.get(head, "")) if isinstance(data_row, dict) else (data_row[c_index] if c_index < len(data_row) else ""))
                report.add(FillRecord(target=ref.describe(), value="行" + str(r_offset + 1),
                                      source=entry.source, file=output))
        elif ref.kind == "docx_placeholder":
            mapping = {ref.placeholder: value}
            count = _replace_in_runs(document, mapping)
            report.add(FillRecord(target=ref.describe(), value=value,
                                  source=entry.source, file=output,
                                  message=f"替换 {count} 处" if count else "未找到占位符"))
    document.save(output)


def _apply_pptx(path: str, output: str, entries: list[MappingEntry],
                record_index: int, record: dict, report: FillReport) -> None:
    try:
        from pptx import Presentation  # type: ignore
    except Exception as exc:
        raise DataError("未安装 python-pptx，无法写入 PPT 文件。") from exc
    prs = Presentation(path)
    for entry in entries:
        ref = entry.target
        value = apply_transform(record.get(entry.source, entry.default), entry.transform)
        if ref.kind == "pptx_table_cell":
            slide = prs.slides[ref.slide]
            shape = _find_table_shape(slide, ref.table)
            if shape is None:
                raise FillError(f"PPT 第{ref.slide + 1}页没有表格{ref.table + 1}")
            table = shape.table
            if ref.row < 0 or ref.row >= len(table.rows) or ref.col < 0 or ref.col >= len(table.columns):
                raise FillError(f"PPT 表格单元格超出范围：{ref.row},{ref.col}")
            table.cell(ref.row, ref.col).text = value
            report.add(FillRecord(target=ref.describe(), value=value,
                                  source=entry.source, file=output))
        elif ref.kind == "pptx_placeholder":
            mapping = {ref.placeholder: value}
            replaced = 0
            for slide in prs.slides:
                for shape in slide.shapes:
                    if not getattr(shape, "has_text_frame", False):
                        continue
                    replaced += _replace_in_runs(shape.text_frame, mapping)
            report.add(FillRecord(target=ref.describe(), value=value,
                                  source=entry.source, file=output,
                                  message=f"替换 {replaced} 处" if replaced else "未找到占位符"))
    prs.save(output)


def _find_table_shape(slide, table_index: int):
    tables = [s for s in slide.shapes if getattr(s, "has_table", False) and s.has_table]
    if 0 <= table_index < len(tables):
        return tables[table_index]
    return None


_KIND_EXTENSIONS = {
    "xlsx_cell": (".xlsx", ".xlsm", ".xltx", ".xltm"),
    "xlsx_rows": (".xlsx", ".xlsm", ".xltx", ".xltm"),
    "docx_table_cell": (".docx", ".docm"),
    "docx_table_rows": (".docx", ".docm"),
    "docx_placeholder": (".docx", ".docm"),
    "pptx_table_cell": (".pptx", ".pptm"),
    "pptx_placeholder": (".pptx", ".pptm"),
}


def _entries_for_target(entries: list[MappingEntry], target: str) -> list[MappingEntry]:
    """Keep only the mappings that apply to *target*.

    A mapping applies when its target kind matches the file extension and its
    ``file`` is either empty (any compatible target) or matches *target* by
    full path or basename.
    """
    base = os.path.basename(target)
    lower = target.lower()
    matched: list[MappingEntry] = []
    for entry in entries:
        kind_exts = _KIND_EXTENSIONS.get(entry.target.kind)
        if kind_exts and not lower.endswith(kind_exts):
            continue
        ref_file = entry.target.file
        if not ref_file:
            matched.append(entry)
            continue
        if os.path.abspath(ref_file) == os.path.abspath(target) or os.path.basename(ref_file) == base:
            matched.append(entry)
    return matched


def _apply_xlsx(path: str, output: str, entries: list[MappingEntry],
                records: list[dict], report: FillReport,
                sheet_name: str = "") -> None:
    eng = get_spreadsheet_engine("direct")
    eng.open(path)
    try:
        sheet = sheet_name or eng.active()
        for index, record in enumerate(records):
            for entry in entries:
                ref = entry.target
                value = apply_transform(record.get(entry.source, entry.default), entry.transform)
                target_sheet = ref.sheet or sheet
                if ref.kind == "xlsx_cell":
                    coord = ref.cell or "A1"
                    advance = ref.row if ref.row >= 0 else 1
                    if index > 0:
                        coord = _advance_coord(coord, advance * index)
                    eng.write_value(target_sheet, coord, value)
                    report.add(FillRecord(target=f"{target_sheet}!{coord}", value=value,
                                          source=entry.source, file=output))
                elif ref.kind == "xlsx_rows":
                    rows = record.get("_rows")
                    if not rows:
                        report.warn("没有可写入的行数据（_rows 为空）。")
                        continue
                    start_row, _ = _split_anchor(ref.cell or "A1")
                    written = eng.write_rows(target_sheet, start_row, rows)
                    report.add(FillRecord(target=ref.describe(), value=f"写入{written}行",
                                          source=entry.source, file=output))
                else:
                    raise FillError(f"不支持的目标类型：{ref.kind}")
    finally:
        eng.save(output)
        eng.close()


def _advance_coord(coord: str, steps: int) -> str:
    if steps <= 0:
        return coord
    match = re.match(r"^(\$?[A-Za-z]{1,3})(\$?\d{1,7})$", coord)
    if not match:
        return coord
    prefix, row_part = match.group(1), match.group(2)
    row = int(re.sub(r"[^0-9]", "", row_part))
    return f"{prefix}{row + steps}"


def fill_cross_files(request: CrossFillRequest) -> FillReport:
    report = FillReport(title="跨文件填表")
    if not request.source_paths:
        raise DataError("请选择至少一个源文件。")
    if not request.target_paths:
        raise DataError("请选择至少一个目标文件。")
    if not request.mappings:
        raise DataError("请至少添加一条字段映射。")

    sources = _merge_sources(request.source_paths, request.source_sheet, request.source_header_row)
    records = sources.as_records()
    if not records:
        raise DataError("源文件没有数据行。")
    # Make the raw row set available to "fill rows" targets (xlsx_rows / table_rows).
    raw_rows = [list(row) for row in sources.rows]
    for record in records:
        record["_rows"] = raw_rows
    report.warn(f"源数据：{len(records)} 行 / {len(sources.headers)} 列。")

    output_dir = request.output_dir or os.path.dirname(request.target_paths[0]) or "."
    os.makedirs(output_dir, exist_ok=True)

    for target in request.target_paths:
        base = os.path.basename(target)
        lower = target.lower()
        entries = _entries_for_target(request.mappings, target)
        if not entries:
            report.warn(f"跳过 {base}：没有适用于该目标的映射。")
            continue
        if lower.endswith((".xlsx", ".xlsm", ".xltx", ".xltm")):
            out = os.path.join(output_dir, base)
            _copy_to_output(target, out)
            _apply_xlsx(target, out, entries, records, report,
                        sheet_name=request.sheet_for_target)
        elif lower.endswith((".docx", ".docm")):
            used = records if request.batch_by_row else [records[0]]
            for index, record in enumerate(used):
                out = _batch_output(output_dir, base, index, len(used))
                _copy_to_output(target, out)
                _apply_docx(target, out, entries, index, record, report)
        elif lower.endswith((".pptx", ".pptm")):
            used = records if request.batch_by_row else [records[0]]
            for index, record in enumerate(used):
                out = _batch_output(output_dir, base, index, len(used))
                _copy_to_output(target, out)
                _apply_pptx(target, out, entries, index, record, report)
        else:
            report.warn(f"跳过不支持的目标文件：{base}")

    report.output_path = output_dir
    report.warn(f"已生成 {len(request.target_paths)} 个目标文件于 {output_dir}")
    return report


# ---------------------------------------------------------------------------
# Preview helpers (used by the GUI)
# ---------------------------------------------------------------------------

def preview_image(image_path: str, ocr_engine: OcrEngine | None = None,
                  engine_pref: str = "auto") -> dict:
    engine = ocr_engine or get_ocr_engine(engine_pref)
    result, prepared = engine.recognize_with_image(image_path)
    words = result.words_as_tuples()
    grid = build_grid_from_ocr(prepared, words)
    if not grid.used_lines and not grid.cells:
        grid = cluster_words_to_grid(words)
    return {
        "path": image_path,
        "text": result.text,
        "words": [(w.text, w.x, w.y, w.w, w.h) for w in result.words],
        "engine": result.engine,
        "elapsed": result.elapsed,
        "rows": grid.rows,
        "cols": grid.cols,
        "used_lines": grid.used_lines,
        "matrix": grid.matrix(),
    }
