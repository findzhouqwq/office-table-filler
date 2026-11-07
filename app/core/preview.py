"""Build a preview of what a fill run would write, without touching disk.

The GUI calls one of the ``build_*_preview`` functions, shows the rows and
validation issues, and only then lets the user start the real run.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field

from . import pdf_io
from .errors import AppError
from .filler import (_default_plan, _entries_for_target, _merge_sources,
                     _row_values_from_plan, _text_to_records, preview_image)
from .mapping import extract_fields, parse_labeled_lines
from .ocr import OcrEngine, get_ocr_engine
from .utils import clean_text, norm_header
from .validation import (ValidationResult, validate_cross_request,
                         validate_field_plan, validate_image_request,
                         validate_text_request)

MAX_PREVIEW_ROWS = 200
MAX_PREVIEW_IMAGES = 5
MAX_PREVIEW_PDF_PAGES = 3


@dataclass
class PreviewResult:
    title: str = ""
    headers: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    issues: ValidationResult = field(default_factory=ValidationResult)
    summary: str = ""
    details: list[str] = field(default_factory=list)
    output_path: str = ""
    total_rows: int = 0
    truncated: bool = False

    @property
    def ok(self) -> bool:
        return self.issues.ok

    def to_text(self) -> str:
        lines = [self.summary] if self.summary else []
        lines.append(f"待写入行数：{self.total_rows}")
        if self.details:
            lines.extend(self.details)
        lines.append("")
        lines.append(self.issues.to_text())
        return "\n".join(lines)


def _target_info(path: str, sheet: str, header_row: int) -> tuple[list[str], dict[str, str]]:
    """Return display headers and a normalised header -> column-letter map."""
    from openpyxl.utils import get_column_letter

    from .spreadsheet import OpenpyxlEngine

    engine = OpenpyxlEngine()
    engine.open(path)
    try:
        names = engine.sheets()
        target = sheet or names[0]
        if target not in names:
            lowered = {name.lower(): name for name in names}
            if target.lower() not in lowered:
                raise AppError(f"找不到工作表：{target}")
            target = lowered[target.lower()]
        _, max_col = engine.dims(target)
        display: list[str] = []
        header_map: dict[str, str] = {}
        for col in range(1, max_col + 1):
            value = clean_text(engine.cell_value(target, f"{get_column_letter(col)}{header_row}"))
            if not value:
                continue
            display.append(value)
            header_map.setdefault(norm_header(value), get_column_letter(col))
        return display, header_map
    finally:
        engine.close()


def _make_headers(display: list[str], width: int) -> list[str]:
    headers = list(display[:width])
    for index in range(len(headers) + 1, width + 1):
        headers.append(f"列{index}")
    return headers


def _string_rows(rows: list[list[object]]) -> list[list[str]]:
    return [[clean_text(value) for value in row] for row in rows]


def _trim(result: PreviewResult) -> PreviewResult:
    result.total_rows = len(result.rows)
    if len(result.rows) > MAX_PREVIEW_ROWS:
        result.rows = result.rows[:MAX_PREVIEW_ROWS]
        result.truncated = True
        result.details.append(f"仅显示前 {MAX_PREVIEW_ROWS} 行，执行时会处理全部 {result.total_rows} 行。")
    return result


# ---------------------------------------------------------------------------
# Feature 1: image / PDF OCR
# ---------------------------------------------------------------------------

def build_image_preview(request, ocr_engine: OcrEngine | None = None) -> PreviewResult:
    result = PreviewResult(title="图片识别填表 · 预览")
    result.issues.extend(validate_image_request(request))
    if not result.issues.ok:
        return result

    try:
        display, header_map = _target_info(request.target_path, request.sheet,
                                           max(1, request.header_row))
    except AppError as exc:
        result.issues.error(str(exc))
        return result

    engine_pref = request.ocr_engine or "auto"
    try:
        engine = ocr_engine or get_ocr_engine(engine_pref)
    except AppError as exc:
        result.issues.error(str(exc))
        return result

    selected = list(request.images[:MAX_PREVIEW_IMAGES])
    temp_dir = ""
    try:
        paths, temp_dir = pdf_io.expand_image_inputs(
            selected, max_pdf_pages=MAX_PREVIEW_PDF_PAGES)
        if not paths:
            result.issues.warning("没有可预览的图片或 PDF 页面。")
            return result
        cursor = 1
        width = max(len(display), 1)
        for path in paths:
            name = os.path.basename(path)
            if request.mode == "fields":
                ocr_result = engine.recognize(path)
                fields = extract_fields(ocr_result.text, request.extraction_rules)
                if not fields and not request.extraction_rules:
                    fields = parse_labeled_lines(ocr_result.text)
                row = _row_values_from_plan(request.column_plan, fields, header_map, width)
                result.rows.append([clean_text(v) for v in row])
                result.details.append(f"第{cursor}行 ← {name}")
            else:
                data = preview_image(path, ocr_engine=engine, engine_pref=engine_pref)
                matrix = data.get("matrix") or []
                if not matrix:
                    result.issues.warning(f"{name}：未识别到表格内容。")
                    continue
                offset = 0
                anchor = (request.anchor or "A1").strip().upper()
                from .filler import _split_anchor
                _, anchor_col = _split_anchor(anchor)
                offset = anchor_col - 1
                for row in matrix:
                    trimmed = [clean_text(v) for v in row]
                    if not any(trimmed):
                        continue
                    values = [""] * offset + trimmed if offset else trimmed
                    result.rows.append([clean_text(v) for v in values])
                result.details.append(f"{name}：识别 {len(matrix)} 行")
            cursor += 1
    except AppError as exc:
        result.issues.error(str(exc))
    finally:
        pdf_io.cleanup_temp_dir(temp_dir)

    width = max((len(row) for row in result.rows), default=0)
    width = max(width, len(display), 1)
    result.headers = _make_headers(display, width)
    result.rows = [row + [""] * (width - len(row)) for row in result.rows]
    result.summary = (
        f"共 {len(request.images)} 个输入；预览前 {len(selected)} 个"
        f"（PDF 最多 {MAX_PREVIEW_PDF_PAGES} 页），预计写入 {len(result.rows)} 行。")
    if len(request.images) > MAX_PREVIEW_IMAGES:
        result.details.append(
            f"为加快预览，只识别了前 {MAX_PREVIEW_IMAGES} 个输入；执行时会处理全部。")
    result.output_path = request.output_path or request.target_path
    return _trim(result)


# ---------------------------------------------------------------------------
# Feature 2: text
# ---------------------------------------------------------------------------

def build_text_preview(request) -> PreviewResult:
    result = PreviewResult(title="文字填表 · 预览")
    try:
        records = _text_to_records(request)
    except AppError as exc:
        result.issues.error(str(exc))
        return result

    try:
        display, header_map = _target_info(request.target_path, request.sheet,
                                           max(1, request.header_row))
    except AppError as exc:
        display, header_map = [], {}

    result.issues.extend(validate_text_request(request, records, header_map))
    plan = request.column_plan or _default_plan(records, header_map)
    if plan:
        result.issues.extend(validate_field_plan(plan, header_map))

    width = max(len(display), len(plan), 1)
    for record in records:
        try:
            row = _row_values_from_plan(plan, record, header_map, width)
        except AppError as exc:
            result.issues.error(str(exc))
            break
        result.rows.append([clean_text(v) for v in row])
    result.headers = _make_headers(display, width)
    result.rows = [row + [""] * (width - len(row)) for row in result.rows]
    result.summary = f"源数据 {len(records)} 行，目标 {len(plan)} 列，将追加/写入到 "
    result.summary += os.path.basename(request.target_path or "（未选择）")
    result.output_path = request.output_path or request.target_path
    return _trim(result)


# ---------------------------------------------------------------------------
# Feature 3: cross-file
# ---------------------------------------------------------------------------

def build_cross_preview(request) -> PreviewResult:
    result = PreviewResult(title="跨文件填表 · 预览")
    try:
        sources = _merge_sources(request.source_paths, request.source_sheet,
                                 max(1, request.source_header_row))
    except AppError as exc:
        result.issues.error(str(exc))
        return result

    records = sources.as_records()
    result.issues.extend(validate_cross_request(request, sources.headers, records))
    result.headers = list(sources.headers)
    result.rows = _string_rows([list(row) for row in sources.rows])
    result.total_rows = len(records)

    details: list[str] = []
    for target in request.target_paths:
        entries = _entries_for_target(request.mappings, target)
        if not entries:
            continue
        lower = target.lower()
        if lower.endswith((".docx", ".docm", ".pptx", ".pptm")):
            copies = len(records) if request.batch_by_row else 1
            details.append(
                f"{os.path.basename(target)}：{len(entries)} 条映射，生成 {copies} 份。")
        else:
            details.append(
                f"{os.path.basename(target)}：{len(entries)} 条映射 × {len(records)} 行。")
    result.details = details
    result.summary = (f"源 {len(records)} 行 / {len(sources.headers)} 列，"
                      f"目标 {len(request.target_paths)} 个文件。")
    result.output_path = request.output_dir
    return _trim(result)
