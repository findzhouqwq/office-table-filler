"""Parse compact destination specs into :class:`TargetRef` objects.

Examples of the ``位置参数`` text shown in the GUI:

* ``xlsx_cell``         ``Sheet1!B2``  or ``B2``  (optionally ``步进=1``)
* ``xlsx_rows``         ``Sheet1!A3``
* ``docx_table_cell``   ``table=0,row=1,col=2``
* ``docx_table_rows``   ``table=0,row_start=1``
* ``pptx_table_cell``   ``slide=0,table=0,row=1,col=2``
* ``docx_placeholder``  ``{{姓名}}``
* ``pptx_placeholder``  ``{{姓名}}``
"""

from __future__ import annotations

import re

from .errors import MappingError
from .models import TargetRef


def _split_cell(spec: str) -> tuple[str, str]:
    spec = (spec or "").strip()
    if not spec:
        raise MappingError("请填写位置参数。")
    if "!" in spec:
        sheet, cell = spec.split("!", 1)
        return sheet.strip().strip("'"), cell.strip()
    return "", spec


def parse_target_spec(kind: str, spec: str, default_file: str = "") -> TargetRef:
    """Turn a user-supplied spec string into a :class:`TargetRef`."""
    ref = TargetRef(kind=kind, file=default_file)
    spec = (spec or "").strip()

    if kind == "xlsx_cell":
        step_match = re.search(r"(?:步进|step)\s*=\s*(\d+)", spec, re.I)
        spec = re.sub(r"\s*(?:步进|step)\s*=\s*\d+", "", spec, flags=re.I).strip()
        ref.sheet, ref.cell = _split_cell(spec)
        ref.row = int(step_match.group(1)) if step_match else -1
    elif kind == "xlsx_rows":
        spec = re.sub(r"\s*(?:步进|step)\s*=\s*\d+", "", spec, flags=re.I).strip()
        ref.sheet, ref.cell = _split_cell(spec)
    elif kind == "docx_table_cell":
        ref.table, ref.row, ref.col = _ints(spec, ("table", "row", "col"))
    elif kind == "docx_table_rows":
        ref.table = _int(spec, "table")
        ref.row_start = _int(spec, "row_start")
    elif kind == "pptx_table_cell":
        ref.slide = _int(spec, "slide")
        ref.table = _int(spec, "table")
        ref.row = _int(spec, "row")
        ref.col = _int(spec, "col")
    elif kind == "docx_placeholder":
        ref.placeholder = _placeholder(spec)
    elif kind == "pptx_placeholder":
        ref.placeholder = _placeholder(spec)
    else:
        raise MappingError(f"未知的目标类型：{kind}")
    return ref


def _placeholder(spec: str) -> str:
    spec = (spec or "").strip()
    if not spec:
        raise MappingError("请填写占位符，例如 {{姓名}}。")
    match = re.search(r"\{\{\s*([^}]+?)\s*\}\}", spec)
    if not match:
        match = re.search(r"\{\s*([^}]+?)\s*\}", spec)
    if not match:
        raise MappingError(f"占位符格式应为 {{{{字段名}}}}：{spec}")
    return match.group(1).strip()


def _option(spec: str, name: str) -> str:
    match = re.search(rf"{name}\s*=\s*(\d+)", spec)
    return match.group(1) if match else ""


def _int(spec: str, name: str) -> int:
    match = re.search(rf"{name}\s*=\s*(\d+)", spec)
    if not match:
        raise MappingError(f"位置参数缺少 {name}=数字。")
    return int(match.group(1))


def _ints(spec: str, names: tuple[str, str, str]) -> tuple[int, int, int]:
    return tuple(_int(spec, name) for name in names)  # type: ignore[return-value]


def describe_target(kind: str, spec: str, default_file: str = "") -> str:
    try:
        ref = parse_target_spec(kind, spec, default_file)
    except MappingError as exc:
        return f"无效：{exc}"
    return ref.describe()
