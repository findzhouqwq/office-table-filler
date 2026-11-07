"""Field extraction from free text and column matching.

Used by the *image OCR* and *text fill* features to turn unstructured text
into a ``{field: value}`` dictionary.
"""

from __future__ import annotations

import re

from .errors import MappingError
from .models import ExtractionRule
from .utils import apply_transform, clean_text, norm_header, split_lines


def extract_fields(text: str, rules: list[ExtractionRule]) -> dict[str, str]:
    """Apply extraction rules to *text* and return a field dictionary."""
    result: dict[str, str] = {}
    lines = split_lines(text)
    for rule in rules:
        if rule.pattern_type == "grid":
            continue
        try:
            value = _extract_one(rule, text, lines)
        except Exception as exc:
            raise MappingError(f"字段 “{rule.field}” 提取失败：{exc}") from exc
        if value is None or value == "":
            if rule.required and not rule.default:
                raise MappingError(f"必填字段 “{rule.field}” 未提取到内容。")
            value = rule.default
        result[rule.field] = apply_transform(value, rule.transform) if value != "" else rule.default
    return result


def _extract_one(rule: ExtractionRule, text: str, lines: list[str]) -> str:
    kind = rule.pattern_type
    if kind == "line":
        idx = rule.position
        if 0 <= idx < len(lines):
            return lines[idx]
        return rule.default

    if kind == "word":
        # handled by caller via OCR words; fall back to words in text
        words = re.findall(r"[\w\u4e00-\u9fff]+", text)
        idx = rule.position
        if 0 <= idx < len(words):
            return words[idx]
        return rule.default

    if kind == "regex":
        if not rule.pattern:
            return rule.default
        match = re.search(rule.pattern, text, re.S)
        if not match:
            return rule.default
        if match.groups():
            return match.group(1)
        return match.group(0)

    if kind == "label":
        label = rule.pattern or rule.field
        pattern = re.escape(label) + r"\s*[:：]?\s*([^\n\r]*)"
        match = re.search(pattern, text, re.S)
        if match:
            return match.group(1).strip()
        # try a line that starts with the label
        for line in lines:
            if re.match(rf"\s*{re.escape(label)}\s*[:：]?\s*", line):
                return re.sub(rf"^\s*{re.escape(label)}\s*[:：]?\s*", "", line).strip()
        return rule.default

    if kind == "keyword":
        keyword = rule.pattern or rule.field
        for line in lines:
            if keyword in line:
                after = line.split(keyword, 1)[1]
                return after.lstrip(" :\u3000:：").strip()
        return rule.default

    raise MappingError(f"未知的提取方式：{kind}")


def parse_labeled_lines(text: str) -> dict[str, str]:
    """Parse ``label: value`` / ``label：value`` text into a dictionary."""
    result: dict[str, str] = {}
    pattern = re.compile(r"^\s*([^:：]{1,20})\s*[:：]\s*(.+?)\s*$")
    for line in split_lines(text):
        match = pattern.match(line)
        if match:
            key = clean_text(match.group(1))
            if key:
                result.setdefault(key, clean_text(match.group(2)))
    return result


def auto_map_headers(source_headers: list[str], target_headers: list[str]) -> list[tuple[int, int]]:
    """Return ``(source_index, target_index)`` pairs with matching headers."""
    source_norm = {norm_header(h): i for i, h in enumerate(source_headers) if h}
    target_norm = {norm_header(h): i for i, h in enumerate(target_headers) if h}
    pairs: list[tuple[int, int]] = []
    for key, src_idx in source_norm.items():
        if key in target_norm:
            pairs.append((src_idx, target_norm[key]))
    return pairs


def resolve_column(spec: str, header_map: dict[str, str]) -> str:
    """Resolve a column spec (letter like ``C`` or a header name) to a letter."""
    spec = (spec or "").strip()
    if not spec:
        raise MappingError("列映射目标为空。")
    if re.fullmatch(r"[A-Za-z]{1,3}", spec):
        return spec.upper()
    key = norm_header(spec)
    if key in header_map:
        return header_map[key]
    # try suffix matching (header may contain extra words)
    for header, letter in header_map.items():
        if key and (key in header or header in key):
            return letter
    raise MappingError(f"找不到列 “{spec}” 对应的表头。")


def expression_value(fields: dict, expr: str, default: str = "") -> str:
    """Resolve ``fields['name']`` / ``name`` / ``row.name`` expressions."""
    expr = (expr or "").strip()
    if not expr:
        return default
    match = re.match(r"^(?:fields?\[|\{)?\s*['\"]?([A-Za-z0-9_\u4e00-\u9fff.]+)['\"]?\s*\]?\}?$", expr)
    if match:
        name = match.group(1)
        if name in fields:
            return str(fields[name])
        if "." in name:
            row, field = name.split(".", 1)
            if row in fields and isinstance(fields[row], dict) and field in fields[row]:
                return str(fields[row][field])
    return default
