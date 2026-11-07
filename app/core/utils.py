"""Small, dependency-light helpers shared across the core layer."""

from __future__ import annotations

import datetime as _dt
import os
import re
import unicodedata
from typing import Iterable, Sequence

from .errors import DataError

# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

_WS_RE = re.compile(r"[ \t\u3000]+")


def clean_text(value: object) -> str:
    """Collapse whitespace/newlines and strip unicode control noise."""
    if value is None:
        return ""
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _WS_RE.sub(" ", text)
    text = "\n".join(line.strip() for line in text.split("\n"))
    return text.strip()


def flatten_text(value: object) -> str:
    """Like :func:`clean_text` but with no line breaks at all."""
    return clean_text(value).replace("\n", " ")


def norm_header(value: object) -> str:
    """Normalise a header so ``" 姓名 "`` and ``"姓名"`` compare equal."""
    text = flatten_text(value)
    text = re.sub(r"[\s:：_\-—/\\()（）\[\]【】]+", "", text)
    return text.lower()


# ---------------------------------------------------------------------------
# Number helpers
# ---------------------------------------------------------------------------

_NUM_RE = re.compile(r"-?\+?\d[\d,]*(?:\.\d+)?%?")


def parse_number(value: object) -> tuple[float | None, str]:
    """Return ``(number, raw_text)`` for the first number found in *value*.

    Full-width digits and Chinese units are normalised. ``None`` is returned
    when no number is present.
    """
    if value is None:
        return None, ""
    if isinstance(value, bool):
        return None, ""
    if isinstance(value, (int, float)):
        return float(value), str(value)
    text = unicodedata.normalize("NFKC", str(value))
    m = _NUM_RE.search(text)
    if not m:
        return None, text
    raw = m.group(0)
    try:
        number = float(raw.replace(",", "").rstrip("%"))
    except ValueError:
        return None, text
    if raw.endswith("%"):
        number /= 100.0
    return number, raw


def to_number(value: object, default: object = None):
    """Coerce *value* to int/float when it contains a number, else ``default``."""
    number, _ = parse_number(value)
    if number is None:
        return default
    if float(number).is_integer() and abs(number) < 1e15:
        return int(number)
    return number


def looks_numeric(value: object) -> bool:
    return parse_number(value)[0] is not None


# ---------------------------------------------------------------------------
# Transform helpers (used by mapping rules)
# ---------------------------------------------------------------------------

_TRANSFORMS = {
    "text": lambda v: flatten_text(v),
    "raw": lambda v: v if v is None else str(v),
    "upper": lambda v: flatten_text(v).upper(),
    "lower": lambda v: flatten_text(v).lower(),
    "title": lambda v: flatten_text(v).title(),
    "number": lambda v: to_number(v, ""),
    "percent": lambda v: to_number(v, ""),
    "strip": lambda v: clean_text(v),
}


def apply_transform(value: object, transform: str = "text") -> str:
    """Apply a named transform to a value.

    Unknown transforms fall back to plain text so a typo never crashes a fill.
    """
    fn = _TRANSFORMS.get(transform)
    if fn is None:
        return flatten_text(value)
    return str(fn(value))


# ---------------------------------------------------------------------------
# File / encoding helpers
# ---------------------------------------------------------------------------

_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "big5", "cp1252", "latin-1")


def read_text_file(path: str, encoding: str | None = None) -> str:
    """Read a text file, tolerating the encodings common on Chinese Windows."""
    if encoding:
        try:
            with open(path, "r", encoding=encoding) as fh:
                return fh.read()
        except UnicodeDecodeError as exc:  # pragma: no cover - user provided encoding
            raise DataError(f"无法用 {encoding} 读取文件：{exc}") from exc
    for enc in _ENCODINGS:
        try:
            with open(path, "r", encoding=enc) as fh:
                return fh.read()
        except UnicodeDecodeError:
            continue
        except OSError as exc:
            raise DataError(f"无法读取文件：{exc}") from exc
    raise DataError(f"无法识别文件编码：{path}")


def detect_delimiter(sample: str) -> str:
    """Guess the delimiter of a delimited text sample."""
    first_line = sample.splitlines()[0] if sample.splitlines() else ""
    candidates = {
        "\t": first_line.count("\t"),
        ";": first_line.count(";"),
        "|": first_line.count("|"),
        ",": first_line.count(","),
    }
    best = max(candidates, key=candidates.get)
    return best if candidates[best] > 0 else ","


def safe_filename(name: str, fallback: str = "output") -> str:
    bad = re.compile(r'[\\/:*?"<>|\r\n\t]+')
    cleaned = bad.sub("_", name).strip().strip(".")
    return cleaned or fallback


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def split_lines(text: str) -> list[str]:
    return [ln for ln in (line.strip() for line in text.splitlines()) if ln]


def chunked(items: Sequence, size: int) -> Iterable[list]:
    for i in range(0, len(items), size):
        yield list(items[i : i + size])


# ---------------------------------------------------------------------------
# Date helpers
# ---------------------------------------------------------------------------

_DATE_PATTERNS = (
    (re.compile(r"(\d{4})\s*[-/年]\s*(\d{1,2})\s*[-/月]\s*(\d{1,2})"), "%Y-%m-%d"),
    (re.compile(r"(\d{4})\s*[-/年]\s*(\d{1,2})\s*月?"), "%Y-%m"),
)


def parse_date(value: object) -> str:
    """Best-effort date extraction producing ``YYYY-MM-DD`` (or ``YYYY-MM``)."""
    if isinstance(value, (_dt.date, _dt.datetime)):
        return value.strftime("%Y-%m-%d")
    text = flatten_text(value)
    for pattern, fmt in _DATE_PATTERNS:
        m = pattern.search(text)
        if m:
            try:
                return _dt.datetime.strptime("-".join(m.groups()), fmt).strftime(fmt)
            except ValueError:
                continue
    return text
