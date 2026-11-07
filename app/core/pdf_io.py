"""PDF helpers: text extraction and page rendering for OCR.

Text-based PDFs are read directly with PyMuPDF.  Scanned PDFs have no text
layer, so their pages are rendered to temporary images and fed to the normal
OCR pipeline.  PyMuPDF is optional: every public function raises a readable
:class:`DataError` when it is missing.
"""

from __future__ import annotations

import os
import shutil
import tempfile
from dataclasses import dataclass

from .errors import DataError

try:  # pragma: no cover - import guard
    import fitz  # PyMuPDF

    _HAS_FITZ = True
except Exception:  # pragma: no cover
    fitz = None  # type: ignore
    _HAS_FITZ = False


def available() -> bool:
    return _HAS_FITZ


def _require() -> None:
    if not _HAS_FITZ:
        raise DataError("未安装 PyMuPDF，无法处理 PDF。请执行：pip install PyMuPDF")


def is_pdf(path: str) -> bool:
    return os.path.splitext(path or "")[1].lower() == ".pdf"


@dataclass
class PdfInfo:
    path: str
    pages: int
    text_chars: int

    @property
    def has_text(self) -> bool:
        return self.text_chars > 0


def inspect_pdf(path: str) -> PdfInfo:
    """Return page count and whether the PDF has an extractable text layer."""
    _require()
    if not os.path.isfile(path):
        raise DataError(f"文件不存在：{path}")
    try:
        document = fitz.open(path)
    except Exception as exc:
        raise DataError(f"无法打开 PDF：{exc}") from exc
    try:
        pages = document.page_count
        text = "".join(page.get_text("text") for page in document)
    finally:
        document.close()
    return PdfInfo(path=path, pages=pages, text_chars=len(text.strip()))


def extract_pdf_text(path: str, max_pages: int = 0) -> str:
    """Extract the text layer of a PDF, page by page."""
    _require()
    if not os.path.isfile(path):
        raise DataError(f"文件不存在：{path}")
    try:
        document = fitz.open(path)
    except Exception as exc:
        raise DataError(f"无法打开 PDF：{exc}") from exc
    try:
        limit = document.page_count if max_pages <= 0 else min(max_pages, document.page_count)
        parts = [document[i].get_text("text") for i in range(limit)]
    finally:
        document.close()
    return "\n".join(parts)


def render_pdf_pages(path: str, out_dir: str = "", dpi: int = 200,
                     max_pages: int = 0) -> list[str]:
    """Render PDF pages to PNG files and return their paths."""
    _require()
    if not os.path.isfile(path):
        raise DataError(f"文件不存在：{path}")
    directory = out_dir or tempfile.mkdtemp(prefix="tablefiller-pdf-")
    os.makedirs(directory, exist_ok=True)
    try:
        document = fitz.open(path)
    except Exception as exc:
        raise DataError(f"无法打开 PDF：{exc}") from exc
    zoom = max(dpi, 72) / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    rendered: list[str] = []
    try:
        limit = document.page_count if max_pages <= 0 else min(max_pages, document.page_count)
        for index in range(limit):
            pixmap = document[index].get_pixmap(matrix=matrix, alpha=False)
            out = os.path.join(directory, f"page_{index + 1:03d}.png")
            pixmap.save(out)
            rendered.append(out)
    finally:
        document.close()
    return rendered


def expand_image_inputs(paths: list[str], max_pdf_pages: int = 0) -> tuple[list[str], str]:
    """Expand PDFs into page images, returning ``(paths, temp_dir)``.

    The caller owns *temp_dir* and should remove it when finished.  It is an
    empty string when no PDF was expanded.
    """
    temp_dir = ""
    expanded: list[str] = []
    for path in paths:
        if not is_pdf(path):
            expanded.append(path)
            continue
        _require()
        if not temp_dir:
            temp_dir = tempfile.mkdtemp(prefix="tablefiller-pdf-")
        # One sub-directory per PDF keeps page file names from colliding.
        sub_dir = tempfile.mkdtemp(dir=temp_dir)
        expanded.extend(render_pdf_pages(path, out_dir=sub_dir, max_pages=max_pdf_pages))
    return expanded, temp_dir


def cleanup_temp_dir(path: str) -> None:
    if path and os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
