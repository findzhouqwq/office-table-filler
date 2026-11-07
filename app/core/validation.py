"""Pre-flight validation for all three fill modes.

Validation never touches the target file for writing: it only reports the
problems (errors block the run, warnings are informational) so the user can
fix a mapping before anything is written.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from .errors import MappingError
from .mapping import resolve_column
from .models import KIND_EXTENSIONS, FieldMapping, MappingEntry
from .utils import norm_header

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_MAX_COL = 16384
_MAX_ROW = 1048576

_TABLE_EXTS = (".xlsx", ".xlsm", ".xltx", ".xltm")


@dataclass
class ValidationIssue:
    level: str          # error | warning
    message: str
    location: str = ""
    record: int = -1

    def __str__(self) -> str:
        prefix = "错误" if self.level == "error" else "提示"
        where = f"[{self.location}] " if self.location else ""
        return f"{prefix}：{where}{self.message}"


@dataclass
class ValidationResult:
    issues: list[ValidationIssue] = field(default_factory=list)

    def error(self, message: str, location: str = "", record: int = -1) -> None:
        self.issues.append(ValidationIssue("error", message, location, record))

    def warning(self, message: str, location: str = "", record: int = -1) -> None:
        self.issues.append(ValidationIssue("warning", message, location, record))

    def extend(self, other: "ValidationResult") -> None:
        self.issues.extend(other.issues)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.level == "error"]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.level == "warning"]

    @property
    def ok(self) -> bool:
        return not self.errors

    def to_text(self) -> str:
        if not self.issues:
            return "校验通过，没有发现问题。"
        return "\n".join(str(issue) for issue in self.issues)


def _col_index(letter: str) -> int:
    index = 0
    for ch in letter.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def validate_cell(coord: str) -> bool:
    match = _CELL_RE.match((coord or "").strip())
    if not match:
        return False
    return (_col_index(match.group(1)) <= _MAX_COL
            and 1 <= int(match.group(2)) <= _MAX_ROW)


def validate_anchor(anchor: str) -> bool:
    return validate_cell((anchor or "A1").strip())


def validate_output(output: str, source: str) -> ValidationResult:
    result = ValidationResult()
    out = (output or source or "").strip()
    if not out:
        result.error("请选择输出文件。")
        return result
    if source and os.path.abspath(out) == os.path.abspath(source):
        result.warning("输出文件与目标文件相同：将在覆盖前自动备份原文件。")
    directory = os.path.dirname(os.path.abspath(out))
    if directory:
        try:
            os.makedirs(directory, exist_ok=True)
        except OSError as exc:
            result.error(f"无法创建输出目录：{exc}")
            return result
        if not os.access(directory, os.W_OK):
            result.error(f"输出目录不可写：{directory}")
    return result


def validate_target_file(path: str) -> ValidationResult:
    result = ValidationResult()
    if not path:
        result.error("请选择目标表格文件。")
    elif not os.path.isfile(path):
        result.error(f"目标文件不存在：{path}")
    elif not path.lower().endswith(_TABLE_EXTS):
        result.error(f"目标文件不是 Excel 表格：{os.path.basename(path)}")
    return result


def validate_field_plan(plan: list[FieldMapping],
                        header_map: dict[str, str]) -> ValidationResult:
    """Check that every mapping resolves and no two share a destination."""
    result = ValidationResult()
    if not plan:
        result.warning("未配置目标列映射，将按源字段顺序写入。")
        return result
    used: dict[str, int] = {}
    for index, mapping in enumerate(plan, start=1):
        source = (mapping.source or "").strip()
        target = (mapping.target or "").strip()
        location = f"映射第{index}条"
        if not source:
            result.error("缺少源字段。", location)
        if not target:
            result.error("缺少目标列 / 表头。", location)
            continue
        try:
            letter = resolve_column(target, header_map)
        except MappingError as exc:
            result.error(f"目标列无法解析：{exc}", location)
            continue
        if not validate_cell(f"{letter}1"):
            result.error(f"目标列超出 Excel 范围：{letter}", location)
            continue
        if letter in used:
            result.error(
                f"目标列 {letter} 被第{used[letter]}条和第{index}条映射同时使用。",
                location)
        else:
            used[letter] = index
    return result


def validate_records(records: list[dict], plan: list[FieldMapping] | None = None,
                     required_fields: list[str] | None = None) -> ValidationResult:
    """Validate source records: emptiness, required fields and duplicates."""
    result = ValidationResult()
    if not records:
        result.error("没有可填写的数据行。")
        return result
    required = [f for f in (required_fields or []) if f]
    if not required and plan:
        required = []
    empty_rows = 0
    signatures: dict[tuple, int] = {}
    for index, record in enumerate(records, start=1):
        values = [record.get(key, "") for key in record.keys()]
        for name in required:
            if str(record.get(name, "") or "").strip() == "":
                result.error(f"必填字段 “{name}” 为空。", f"第{index}行")
        if all(str(v).strip() == "" for v in values):
            empty_rows += 1
            continue
        signature = tuple(str(record.get(k, "")) for k in sorted(record))
        if signature in signatures:
            result.warning(
                f"第{index}行与第{signatures[signature]}行内容完全相同。",
                f"第{index}行")
        else:
            signatures[signature] = index
    if empty_rows:
        result.warning(f"有 {empty_rows} 行内容为空，执行时会跳过。")
    return result


def _kind_matches_target(kind: str, target: str) -> bool:
    extensions = KIND_EXTENSIONS.get(kind)
    if not extensions:
        return False
    return target.lower().endswith(extensions)


def validate_mapping_entry(entry: MappingEntry, source_headers: list[str]) -> ValidationResult:
    result = ValidationResult()
    location = entry.source or "映射"
    if not (entry.source or "").strip():
        result.error("映射缺少源字段。", location)
    elif source_headers and norm_header(entry.source) not in {norm_header(h) for h in source_headers}:
        result.warning(f"源字段 “{entry.source}” 不在源数据表头中，将写入默认值。", location)
    ref = entry.target
    if not ref.kind:
        result.error("映射缺少目标类型。", location)
        return result
    if ref.kind in ("xlsx_cell", "xlsx_rows"):
        if not validate_cell(ref.cell or ""):
            result.error(f"Excel 位置无效：{ref.cell or '（空）'}", location)
    elif ref.kind == "docx_table_cell":
        if ref.table < 0 or ref.row < 0 or ref.col < 0:
            result.error("Word 单元格位置不完整（需要 table/row/col）。", location)
    elif ref.kind == "docx_table_rows":
        if ref.table < 0 or ref.row_start < 0:
            result.error("Word 行区域位置不完整（需要 table/row_start）。", location)
    elif ref.kind == "pptx_table_cell":
        if min(ref.slide, ref.table, ref.row, ref.col) < 0:
            result.error("PPT 单元格位置不完整（需要 slide/table/row/col）。", location)
    elif ref.kind in ("docx_placeholder", "pptx_placeholder"):
        if not (ref.placeholder or "").strip():
            result.error("占位符名称不能为空。", location)
    else:
        result.error(f"未知的目标类型：{ref.kind}", location)
    return result


def validate_cross_mappings(mappings: list[MappingEntry], source_headers: list[str],
                            targets: list[str]) -> ValidationResult:
    result = ValidationResult()
    if not mappings:
        result.error("请至少添加一条字段映射。")
        return result
    for target in targets:
        compatible = [m for m in mappings if _kind_matches_target(m.target.kind, target)]
        if not compatible:
            result.warning(f"{os.path.basename(target)} 没有适用的映射，执行时会跳过。")
    for entry in mappings:
        result.extend(validate_mapping_entry(entry, source_headers))
    return result


def validate_image_request(request) -> ValidationResult:
    """Validate an ``ImageFillRequest`` before OCR starts."""
    result = ValidationResult()
    if not request.images:
        result.error("请至少选择一张图片或 PDF。")
    else:
        for path in request.images:
            if not os.path.isfile(path):
                result.error(f"图片不存在：{path}")
    result.extend(validate_target_file(request.target_path))
    if request.mode not in ("grid", "fields"):
        result.error(f"未知的填写方式：{request.mode}")
    if not validate_anchor(request.anchor):
        result.error(f"写入位置无效：{request.anchor}")
    if request.header_row < 1:
        result.error("表头行必须从 1 开始。")
    if request.mode == "fields" and not request.column_plan and not request.extraction_rules:
        result.warning("未配置提取规则和列映射，将按识别到的字段顺序写入。")
    if request.output_path:
        result.extend(validate_output(request.output_path, request.target_path))
    return result


def validate_text_request(request, records: list[dict],
                          header_map: dict[str, str] | None = None) -> ValidationResult:
    result = ValidationResult()
    result.extend(validate_target_file(request.target_path))
    required = [r.field for r in request.extraction_rules if r.required]
    result.extend(validate_records(records, required_fields=required))
    if request.column_plan:
        result.extend(validate_field_plan(request.column_plan, header_map or {}))
    if request.output_path:
        result.extend(validate_output(request.output_path, request.target_path))
    return result


def validate_cross_request(request, source_headers: list[str],
                           records: list[dict]) -> ValidationResult:
    result = ValidationResult()
    if not request.source_paths:
        result.error("请选择至少一个源文件。")
    else:
        for path in request.source_paths:
            if not os.path.isfile(path):
                result.error(f"源文件不存在：{path}")
    if not request.target_paths:
        result.error("请选择至少一个目标文件。")
    else:
        for path in request.target_paths:
            if not os.path.isfile(path):
                result.error(f"目标文件不存在：{path}")
    result.extend(validate_records(records))
    result.extend(validate_cross_mappings(request.mappings, source_headers, request.target_paths))
    if request.output_dir:
        try:
            os.makedirs(request.output_dir, exist_ok=True)
        except OSError as exc:
            result.error(f"无法创建输出目录：{exc}")
    return result
