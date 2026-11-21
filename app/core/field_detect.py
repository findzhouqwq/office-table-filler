"""Fast, dependency-free detection of fields in clipboard text.

The side panel uses this to turn arbitrary copied text (a chat message, a
contact card, an e-mail signature, a ``label: value`` dump, or simply a few
values on separate lines) into classified ``label + value`` fields, then maps
them onto the headers of the target spreadsheet.

Design goals:

* **broad** – covers the fields that show up in Chinese forms most often;
* **fast** – a handful of pre-compiled regular expressions over at most
  :data:`MAX_TEXT` characters, no third party libraries and no ML model;
* **safe** – every match is only reported once and strongly typed patterns
  (phone, ID card, e-mail…) take precedence over generic numbers.

Nothing here imports openpyxl at module level so detection stays cheap.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, Sequence

from .utils import clean_text, norm_header

#: Detection only looks at the first 20k characters – far beyond what a form
#: clipboard ever contains, but a hard bound keeps the regular expressions
#: predictable even when someone copies a whole document.
MAX_TEXT = 20000

#: ``kind -> (canonical label, aliases…)``.  Aliases double as header names:
#: a detected phone number matches a column called “联系电话” or “手机号码”.
_FIELD_ALIASES: dict[str, tuple[str, ...]] = {
    "name": ("姓名", "名字", "联系人", "客户姓名", "客户名称", "员工姓名",
             "收件人", "申请人", "填报人", "负责人", "真实姓名"),
    "phone": ("手机", "手机号", "手机号码", "移动电话", "联系电话", "电话",
              "联系方式", "电话号码"),
    "telephone": ("座机", "固话", "固定电话", "办公电话", "公司电话"),
    "id_card": ("身份证", "身份证号", "身份证号码", "证件号", "证件号码"),
    "email": ("邮箱", "邮箱地址", "电子邮箱", "电子邮件", "邮件", "email",
              "e-mail"),
    "address": ("地址", "住址", "通讯地址", "家庭地址", "联系地址",
                "收货地址", "工作地址", "详细地址"),
    "company": ("公司", "公司名称", "单位", "单位名称", "企业名称",
                "工作单位", "所在单位", "公司全称"),
    "department": ("部门", "所属部门", "科室", "事业部", "所在部门"),
    "position": ("职位", "岗位", "职务", "头衔", "职称"),
    "date": ("日期", "出生日期", "生日", "入职日期", "签订日期",
             "报名日期", "填表日期"),
    "time": ("时间", "时刻", "具体时间"),
    "amount": ("金额", "费用", "价格", "工资", "薪资", "收入", "合计",
               "总额", "报销金额", "付款金额"),
    "number": ("数字", "数量", "数值", "序号", "工号", "学号"),
    "percent": ("百分比", "比例", "费率", "折扣", "税率"),
    "bank_card": ("银行卡", "银行卡号", "卡号", "银行账号", "账号", "开户账号"),
    "zip": ("邮编", "邮政编码"),
    "plate": ("车牌", "车牌号", "车牌号码"),
    "url": ("网址", "网站", "主页", "链接", "个人主页"),
    "gender": ("性别",),
    "age": ("年龄",),
    "education": ("学历", "学位", "最高学历"),
    "nation": ("民族",),
    "marital": ("婚姻", "婚姻状况", "婚否"),
    "political": ("政治面貌",),
    "credit_code": ("统一社会信用代码", "信用代码", "税号", "纳税人识别号"),
    "passport": ("护照", "护照号", "护照号码"),
    "ip": ("ip", "ip地址"),
    "mac": ("mac", "mac地址", "物理地址"),
    "qq": ("qq", "qq号", "qq号码"),
    "wechat": ("微信", "微信号", "wechat"),
    "order_no": ("订单号", "订单编号", "单号", "编号", "流水号"),
    "remark": ("备注", "说明", "描述", "备注信息"),
}

_CANONICAL: dict[str, str] = {kind: aliases[0] for kind, aliases in _FIELD_ALIASES.items()}

_ALIAS_TO_KIND: dict[str, str] = {}
for _kind, _aliases in _FIELD_ALIASES.items():
    for _alias in _aliases:
        _ALIAS_TO_KIND.setdefault(norm_header(_alias), _kind)


@dataclass(frozen=True)
class DetectedField:
    """One classified value extracted from text."""

    kind: str
    label: str
    value: str
    confidence: float = 0.8


# ---------------------------------------------------------------------------
# Pattern tables
# ---------------------------------------------------------------------------

# Longest aliases first so “联系电话” wins over “电话” and “公司电话” over
# “公司” when both start at the same position.
_LABEL_ALTERNATION = "|".join(
    re.escape(alias)
    for alias in sorted({alias for aliases in _FIELD_ALIASES.values()
                         for alias in aliases}, key=len, reverse=True)
)
_LABEL_COLON_RE = re.compile(
    rf"(?:{_LABEL_ALTERNATION})[ \t\u3000]*(?:[:：][ \t\u3000]*|[ \t\u3000]+)")

# Unknown labels such as “工号: A123” are still useful – keep them as custom
# fields.  Lines already consumed by a known label are blanked out first.
_GENERIC_LABEL_RE = re.compile(
    r"^[ \t\u3000]*([\u4e00-\u9fffA-Za-z][\u4e00-\u9fffA-Za-z0-9_\-（）() ]{0,15})"
    r"[ \t\u3000]*[:：][ \t\u3000]*(.+?)[ \t\u3000]*$",
    re.M,
)

_SKIP_KEYS = {"http", "https", "ftp", "mailto", "file"}

# Ordered by specificity: a match claims its character span so a generic
# number never duplicates a phone number, ID card or amount.
_PATTERNS: tuple[tuple[str, re.Pattern[str], float], ...] = (
    ("email", re.compile(
        r"(?<![A-Za-z0-9._%+\-])[A-Za-z0-9._%+\-]{1,64}@"
        r"[A-Za-z0-9.\-]{1,190}\.[A-Za-z]{2,10}"), 0.97),
    ("url", re.compile(r"(?:https?://|www\.)[^\s，。；;、,|）)】\]]+", re.I), 0.95),
    ("id_card", re.compile(
        r"(?<!\d)[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])"
        r"(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?!\d)"), 0.97),
    ("credit_code", re.compile(
        r"(?<![0-9A-Za-z])[0-9A-HJ-NPQRTUWXY]{2}\d{6}"
        r"[0-9A-HJ-NPQRTUWXY]{10}(?![0-9A-Za-z])"), 0.95),
    ("phone", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)"), 0.97),
    ("telephone", re.compile(
        r"(?<!\d)0\d{2,3}[-\s]?\d{7,8}(?:-\d{1,6})?(?!\d)"), 0.9),
    ("bank_card", re.compile(r"(?<!\d)(?:\d{4}[ -]?){3}\d{4,7}(?!\d)"), 0.88),
    ("plate", re.compile(
        r"[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领]"
        r"[A-HJ-NP-Z][A-HJ-NP-Z0-9]{4,5}[A-HJ-NP-Z0-9挂学警港澳]"), 0.93),
    ("ip", re.compile(
        r"(?<![\d.])(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}"
        r"(?:25[0-5]|2[0-4]\d|1?\d?\d)(?![\d.])"), 0.9),
    ("mac", re.compile(
        r"(?<![0-9A-Fa-f])(?:[0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2}(?![0-9A-Fa-f])"),
     0.9),
    ("date", re.compile(
        r"(?<!\d)(?:19|20)\d{2}\s*[-/.年]\s*\d{1,2}"
        r"(?:\s*[-/.月]\s*\d{1,2}\s*日?)?(?!\d)"), 0.9),
    ("time", re.compile(
        r"(?<!\d)(?:[01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d)?(?!\d)"), 0.85),
    ("percent", re.compile(
        r"(?<![\d.])\d{1,12}(?:\.\d{1,6})?\s*[%％](?!\d)"), 0.9),
    ("amount", re.compile(
        r"(?<![\d,.])"
        r"(?:[¥￥]\s?\d[\d,]{0,15}(?:\.\d{1,2})?"
        r"|\d[\d,]{0,15}(?:\.\d{1,2})?\s*(?:元|万元|人民币))"
        r"(?!\d)"), 0.88),
    ("company", re.compile(
        r"(?<![\u4e00-\u9fffA-Za-z0-9])[\u4e00-\u9fffA-Za-z0-9（）()]{2,40}"
        r"(?:股份有限公司|有限责任公司|有限公司|集团|公司|事务所|工作室)"), 0.68),
    ("address", re.compile(
        r"(?:[\u4e00-\u9fff]{2,20}(?:省|自治区|特别行政区))?"
        r"[\u4e00-\u9fff]{2,20}(?:市|自治州|地区)"
        r"[\u4e00-\u9fff0-9A-Za-z]{0,20}(?:区|县|镇|街道)"
        r"[\u4e00-\u9fff0-9A-Za-z\-#号栋室楼]{2,40}"), 0.7),
    ("age", re.compile(r"(?<!\d)\d{1,3}\s*岁"), 0.9),
    ("political", re.compile(
        r"中共党员|中共预备党员|共青团员|民主党派|无党派人士|群众|党员|团员"), 0.85),
    ("marital", re.compile(r"未婚|已婚|离异|丧偶|再婚"), 0.85),
    ("education", re.compile(
        r"博士研究生|硕士研究生|博士|硕士|研究生|本科|大专|专科|高中|中专|初中|小学"),
     0.82),
    ("gender", re.compile(r"(?<![\u4e00-\u9fff])(?:男|女)(?![\u4e00-\u9fff])"), 0.7),
    ("zip", re.compile(r"(?<![\d.])[1-9]\d{5}(?![\d.])"), 0.55),
    ("number", re.compile(
        r"(?<![\d.])\d{1,15}(?:\.\d{1,6})?(?![\d.])"), 0.35),
)

_NON_NAME = {
    "公司", "单位", "部门", "市场部", "财务部", "技术部", "人事部", "销售部",
    "未婚", "已婚", "离异", "男", "女", "是", "否",
}

# Cheap token checks that let long prose skip the two greedy patterns entirely.
_PREFILTERS: dict[str, tuple[str, ...]] = {
    "company": ("公司", "集团", "事务所", "工作室"),
    "address": ("省", "市", "区", "县", "镇", "街道", "路", "街", "巷", "号",
                "村", "小区", "大厦", "广场", "楼", "室"),
}

# Greedy values often swallow the words in front of them when the text is a
# sentence ("我的地址是北京市…", "他在某某公司上班").  Strip those lead-ins so
# the field value starts at the actual company / address.
_LEAD_IN: dict[str, re.Pattern[str]] = {
    "address": re.compile(
        r"^(?:(?:我的|他的|她的|我|他|她|客户的|联系|家庭|工作|收货|"
        r"通讯|详细|现)?(?:地址|住址)(?:是|为|在)?|住在|位于|在)"),
    "company": re.compile(
        r"^(?:(?:我的|他的|她的|我|他|她|客户的|本人|我们|他们|对方)"
        r"(?:在|就职于|工作于|供职于|就职|工作|来自|属于)"
        r"|(?:就职于|工作于|供职于|来自|属于))"),
}


def _clean_value(kind: str, raw: str) -> str:
    value = clean_text(raw)
    pattern = _LEAD_IN.get(kind)
    if pattern is not None:
        value = pattern.sub("", value, count=1).strip()
    return value


# ---------------------------------------------------------------------------
# Detection
# ---------------------------------------------------------------------------

def detect_fields(text: str) -> list[DetectedField]:
    """Return classified fields found in *text*, strongest matches first."""
    text = clean_text(text or "")[:MAX_TEXT]
    if not text:
        return []

    used = bytearray(len(text))
    found = list(_scan_labeled(text, used))
    found.extend(_scan_generic_labels(text, used))
    found.extend(_scan_patterns(text, used))

    if not found:
        fallback = _fallback_name(text)
        if fallback is not None:
            return [fallback]

    result: list[DetectedField] = []
    seen: set[tuple[str, str]] = set()
    for field in found:
        key = (field.kind, norm_header(field.value))
        if not field.value or key in seen:
            continue
        seen.add(key)
        result.append(field)
    return result


def _scan_labeled(text: str, used: bytearray) -> Iterable[DetectedField]:
    """Find known ``label: value`` pairs, including several per line."""
    matches = list(_LABEL_COLON_RE.finditer(text))
    for index, match in enumerate(matches):
        start = match.end()
        # The value normally ends at the next label or at a line break.  Keep
        # the untrimmed boundary first: vertical forms put the value on the
        # following line (``姓名：`` / ``张三``).
        raw_end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        end = raw_end
        for stop_char in ("\n", "\r", ";", "；", "|"):
            position = text.find(stop_char, start, end)
            if position != -1:
                end = position
        value = clean_text(text[start:end]).strip(",;|、")
        if not value:
            offset = start
            for line in text[start:raw_end].splitlines(keepends=True):
                if line.strip():
                    value = clean_text(line).strip(",;|、")
                    end = offset + len(line.rstrip("\r\n"))
                    break
                offset += len(line)
        if not value:
            continue
        label = clean_text(re.sub(r"\s*[:：]\s*$", "", match.group(0)))
        kind = _ALIAS_TO_KIND.get(norm_header(label))
        if not kind:
            continue
        _mark(used, match.start(), end)
        yield DetectedField(kind, _CANONICAL[kind], value, 0.99)


def _scan_generic_labels(text: str, used: bytearray) -> Iterable[DetectedField]:
    """Pick up unknown labels like ``工号: A123`` from the remaining text."""
    masked = "".join(
        char if not used[i] or char in "\r\n" else " "
        for i, char in enumerate(text)
    )
    for match in _GENERIC_LABEL_RE.finditer(masked):
        label = clean_text(match.group(1))
        value = clean_text(match.group(2))
        key = norm_header(label)
        if not value or not key or len(key) > 16:
            continue
        if key in _ALIAS_TO_KIND or key in _SKIP_KEYS:
            continue
        if value.startswith("//"):
            continue
        yield DetectedField("custom", label, value, 0.85)


def _scan_patterns(text: str, used: bytearray) -> Iterable[DetectedField]:
    for kind, pattern, confidence in _PATTERNS:
        tokens = _PREFILTERS.get(kind)
        if tokens and not any(token in text for token in tokens):
            continue
        for match in pattern.finditer(text):
            start, end = match.span()
            if start == end or 1 in used[start:end]:
                continue
            value = _clean_value(kind, match.group(0))
            if not value:
                continue
            _mark(used, start, end)
            yield DetectedField(kind, _CANONICAL[kind], value, confidence)


def _mark(used: bytearray, start: int, end: int) -> None:
    used[start:end] = b"\x01" * (end - start)


def _fallback_name(text: str) -> DetectedField | None:
    """Treat a single short Chinese line as a name when nothing else matched."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) != 1:
        return None
    candidate = lines[0]
    if not 2 <= len(candidate) <= 4:
        return None
    if not all("\u4e00" <= char <= "\u9fff" for char in candidate):
        return None
    if candidate in _NON_NAME:
        return None
    return DetectedField("name", "姓名", candidate, 0.45)


def classify_value(value: str) -> str:
    """Return the most likely kind of a single value (``""`` when unknown)."""
    fields = detect_fields(str(value))
    return fields[0].kind if fields else ""


_LABEL_VALUE_RE = re.compile(
    r"(?:姓名|名字|手机|电话|邮箱|公司|单位|部门|职位|地址|身份证)\s*[:：]")
# XML (and therefore xlsx) cannot store most C0 control characters; pasted
# clipboard data may contain them, which would make openpyxl fail on save.
_CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
_MAX_CELL_TEXT = 32767


def split_positional_row(text: str) -> list[str] | None:
    """Return the cells of a single tab-separated clipboard row.

    Copying one row out of Excel / WPS yields one line of tab separated values
    with no field names.  Such a row is used positionally by the side panel
    (column *n* → header *n*).  ``None`` is returned for multi-line text,
    ``label: value`` text, and two-column label/value dumps.
    """
    lines = [line for line in str(text or "").replace("\r\n", "\n")
             .replace("\r", "\n").split("\n") if line.strip()]
    if len(lines) != 1 or "\t" not in lines[0]:
        return None
    line = lines[0]
    if "：" in line or _LABEL_VALUE_RE.search(line):
        return None
    cells = [clean_text(cell) for cell in line.split("\t")]
    while cells and not cells[-1]:
        cells.pop()
    if len(cells) < 2 or not any(cells):
        return None
    # A two-column dump copied from a table (label / value / label / value)
    # is not a data row – let the normal label scanner handle it.
    pairs = sum(1 for index in range(0, len(cells) - 1, 2)
                if norm_header(cells[index]) in _ALIAS_TO_KIND)
    if pairs >= max(1 if len(cells) == 2 else 2, len(cells) // 4):
        return None
    return cells


def _fields_from_values(values: dict[str, str]) -> list[DetectedField]:
    detected: list[DetectedField] = []
    for label, value in values.items():
        label = str(label).strip()
        if value is None or not label:
            continue
        value = _CONTROL_RE.sub("", str(value)).strip()
        if not value:
            continue
        # The label is the strongest hint; only then inspect the value itself.
        kind = _ALIAS_TO_KIND.get(norm_header(label)) or classify_value(value) or "custom"
        detected.append(DetectedField(kind, label, value, 0.9))
    return detected


def map_values_to_headers(values: dict[str, str],
                          headers: Sequence[str]) -> tuple[dict[str, str], list[DetectedField]]:
    """Like :func:`map_to_headers` but for manually maintained label/value rows."""
    return map_to_headers(_fields_from_values(values), headers)


# Header decorations that may follow a field name without changing its meaning
# ("手机号码（必填）" → "手机号码必填").  A trailing word that is another kind's
# alias ("联系人电话") is *not* neutral, so a name never lands in a phone column.
_NEUTRAL_SUFFIXES = {
    "必填", "选填", "可选", "常用", "备用", "号码", "编号", "信息", "名称",
    "全称", "明细", "备注", "说明", "元", "万元", "人民币", "rmb", "岁",
    "个", "人", "名", "列", "cm", "厘米", "kg", "公斤", "米", "m", "年月日",
}
_NEUTRAL_SUFFIX_RE = re.compile(r"^\d+(?:位|个|人|名|列)$")


def map_to_headers(detected: Sequence[DetectedField],
                   headers: Sequence[str]) -> tuple[dict[str, str], list[DetectedField]]:
    """Map detected fields onto *headers*.

    Returns ``({header: value}, unmatched)``.  Exact labels win, then known
    aliases of the same kind, then suffix containment (``工号`` → ``员工工号``).
    A header is only filled once; high-confidence fields are placed first.
    """
    cleaned: list[tuple[str, str]] = []
    for header in headers:
        if header is None:
            continue
        text = str(header).strip()
        if text:
            cleaned.append((text, norm_header(text)))
    matches, unmatched = _match_all(detected, cleaned)
    filled = {cleaned[index][0]: value for index, value in matches}
    ordered = {header: filled[header] for header, _ in cleaned if header in filled}
    return ordered, unmatched


def _match_all(detected: Sequence[DetectedField],
               headers: list[tuple[str, str]]) -> tuple[list[tuple[int, str]],
                                                        list[DetectedField]]:
    """Return ``[(header_index, value)], unmatched`` (duplicate safe)."""
    used: set[int] = set()
    matches: list[tuple[int, str]] = []
    unmatched: list[DetectedField] = []
    for field in sorted(detected, key=lambda item: -item.confidence):
        index = _match_index(field, headers, used)
        if index is None:
            unmatched.append(field)
        else:
            used.add(index)
            matches.append((index, field.value))
    return matches, unmatched


def _match_index(field: DetectedField, headers: list[tuple[str, str]],
                 used: set[int]) -> int | None:
    label = norm_header(field.label)
    aliases = {norm_header(alias)
               for alias in _FIELD_ALIASES.get(field.kind, ())}
    aliases.add(label)
    aliases.discard("")

    for index, (_, key) in enumerate(headers):
        if index not in used and key and key == label:
            return index

    for index, (_, key) in enumerate(headers):
        if index not in used and key and key in aliases:
            return index

    best_index: int | None = None
    best_score = 0
    for index, (_, key) in enumerate(headers):
        if index in used or len(key) < 2:
            continue
        for alias in aliases:
            if len(alias) < 2:
                continue
            score = 0
            if key.endswith(alias):
                # Suffix matches are the common Chinese header pattern:
                # 员工工号、联系电话、家庭地址.
                score = 1000 + len(alias)
            elif key.startswith(alias):
                remainder = key[len(alias):]
                # Prefix matches are only safe when the rest is a neutral
                # decoration ("必填") or a column number ("电话2").
                if (remainder.isdigit() or remainder in _NEUTRAL_SUFFIXES
                        or _NEUTRAL_SUFFIX_RE.match(remainder)):
                    score = 100 + len(alias)
            if score > best_score:
                best_index, best_score = index, score
    return best_index


# ---------------------------------------------------------------------------
# Writing a row into the target workbook
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class FillReport:
    """What :func:`append_fields` did."""

    path: str
    sheet: str
    row: int
    matched: int
    skipped: tuple[str, ...] = ()
    warning: str = ""


_MAX_EXCEL_ROW = 1048576


def last_data_row(ws, header_row: int, min_col: int, max_col: int) -> int:
    """Last row that actually contains a value inside the table columns.

    ``Worksheet.max_row`` also counts empty cells that only carry formatting,
    which would append the new row far below the visible table.  Scanning the
    cell store (instead of every row up to ``max_row``) keeps this cheap even
    for sheets with a formatted empty area.
    """
    last = header_row
    cells = getattr(ws, "_cells", None)
    if cells is not None:
        for (row, col), cell in cells.items():
            if row <= header_row or col < min_col or col > max_col:
                continue
            value = cell.value
            if value is None:
                continue
            if isinstance(value, str) and not value.strip():
                continue
            if row > last:
                last = row
        return last
    for row in range(ws.max_row, header_row, -1):
        if any(ws.cell(row=row, column=col).value not in (None, "")
               for col in range(min_col, max_col + 1)):
            return row
    return header_row


def avoid_merged_rows(ws, row: int, min_col: int, max_col: int) -> int:
    """Move *row* below merged ranges that would make it unwritable.

    Multi-row merged headers (A1:A2) and merged footer blocks are common; a new
    row inside such a range would otherwise hit a read-only ``MergedCell``.
    """
    changed = True
    while changed:
        changed = False
        for merged in ws.merged_cells.ranges:
            if (merged.min_row <= row <= merged.max_row
                    and merged.min_col <= max_col and merged.max_col >= min_col):
                row = merged.max_row + 1
                changed = True
    return row


def append_fields(path: str, values: dict[str, str], sheet: str = "",
                  header_row: int = 1, output_path: str = "") -> FillReport:
    """Append one row built from *values* to the table in *path*.

    Values are matched against the real headers of the sheet, so a field called
    “手机” lands in a “联系电话” column.  Duplicate header names are written to
    separate, still unused columns.  The workbook is saved atomically with a
    backup next to it.
    """
    from openpyxl.utils import get_column_letter

    from .errors import EngineError, FillError
    from .spreadsheet import OpenpyxlEngine

    path = (path or "").strip()
    if not path:
        raise FillError("请先在主界面选择目标表格。")
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

        columns: list[tuple[str, str, str]] = []
        for col in range(1, max_col + 1):
            letter = get_column_letter(col)
            header = clean_text(engine.cell_value(
                sheet_name, f"{letter}{header_row}"))
            if header:
                columns.append((header, norm_header(header), letter))
        if not columns:
            raise FillError(
                f"工作表「{sheet_name}」第 {header_row} 行没有表头。")

        matches, unmatched = _match_all(
            _fields_from_values(values),
            [(header, key) for header, key, _ in columns],
        )
        if not matches:
            raise FillError(
                "没有字段能与目标表头匹配，请先「读取表头」或检查字段名。")

        ws = engine.worksheet(sheet_name)
        row = last_data_row(ws, header_row, min_col, max_col) + 1
        row = avoid_merged_rows(ws, row, min_col, max_col)
        if row > _MAX_EXCEL_ROW:
            raise FillError("目标工作表已写满，无法再追加新行。")

        for index, value in matches:
            if len(value) > _MAX_CELL_TEXT:
                raise FillError(
                    f"字段「{columns[index][0]}」的内容超过 Excel 单元格上限"
                    f"（{_MAX_CELL_TEXT} 字符）。")
            engine.write_value(sheet_name, f"{columns[index][2]}{row}", value)

        warnings: list[str] = []
        target = engine.save(output_path or path, backup=True, warnings=warnings)
        return FillReport(
            target, sheet_name, row, len(matches),
            tuple(field.label for field in unmatched),
            warnings[0] if warnings else "",
        )
    except OSError as exc:
        raise FillError(
            f"无法写入表格（可能正被 Excel / WPS 占用）：{exc}") from exc
    finally:
        engine.close()
