"""Locate table grids inside images and turn OCR words into a matrix.

The detector is deliberately forgiving: it works on scanned documents and
screenshots that contain ruled tables, and falls back to clustering words into
rows/columns when no ruling is found.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

try:  # pragma: no cover - import guard
    import cv2

    _HAS_CV2 = True
except Exception:  # pragma: no cover
    _HAS_CV2 = False

from .errors import OcrError


@dataclass
class CellBox:
    row: int
    col: int
    x: int
    y: int
    w: int
    h: int
    text: str = ""

    @property
    def cx(self) -> int:
        return self.x + self.w // 2

    @property
    def cy(self) -> int:
        return self.y + self.h // 2


@dataclass
class GridResult:
    rows: int = 0
    cols: int = 0
    cells: list[CellBox] = field(default_factory=list)
    confidence: float = 0.0
    used_lines: bool = False

    def matrix(self) -> list[list[str]]:
        if self.rows <= 0 or self.cols <= 0:
            return []
        grid = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        for cell in self.cells:
            if 0 <= cell.row < self.rows and 0 <= cell.col < self.cols:
                grid[cell.row][cell.col] = cell.text
        return grid

    def to_rows(self) -> list[list[str]]:
        return self.matrix()


def _imread(path: str) -> "np.ndarray":
    if not _HAS_CV2:
        raise OcrError("未安装 opencv-python，无法进行表格线检测。")
    data = np.fromfile(path, dtype=np.uint8)
    if data.size == 0:
        raise OcrError(f"无法读取图片：{path}")
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise OcrError(f"图片解码失败：{path}")
    return img


def _to_gray(img: "np.ndarray") -> "np.ndarray":
    if img.ndim == 2:
        return img
    return cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)


def _find_peaks(profile: np.ndarray, ratio: float = 0.25,
                min_gap: int = 8) -> list[int]:
    """Return centres of contiguous runs above ``max(profile)*ratio``."""
    if profile.size == 0:
        return []
    threshold = max(float(profile.max()) * ratio, 1.0)
    peaks: list[int] = []
    run_start: int | None = None
    for i, value in enumerate(profile):
        if value >= threshold:
            if run_start is None:
                run_start = i
        else:
            if run_start is not None:
                peaks.append((run_start + i - 1) // 2)
                run_start = None
    if run_start is not None:
        peaks.append((run_start + len(profile) - 1) // 2)
    # drop near-duplicates
    filtered: list[int] = []
    for peak in peaks:
        if not filtered or peak - filtered[-1] >= min_gap:
            filtered.append(peak)
    return filtered


def _deskew(gray: "np.ndarray") -> "np.ndarray":
    try:
        coords = np.column_stack(np.where(gray < 128))
        if coords.size == 0:
            return gray
        angle = cv2.minAreaRect(coords)[-1]
        angle = angle - 90 if angle > 45 else angle
        if abs(angle) < 0.1 or abs(angle) > 25:
            return gray
        (h, w) = gray.shape[:2]
        matrix = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
        return cv2.warpAffine(gray, matrix, (w, h), flags=cv2.INTER_CUBIC,
                              borderMode=cv2.BORDER_REPLICATE)
    except Exception:
        return gray


def detect_grid(path_or_img, min_line_ratio: float = 0.18,
                expect_min_rows: int = 2, expect_min_cols: int = 2) -> GridResult:
    """Detect a ruled grid in an image file or BGR/gray ndarray."""
    img = path_or_img if isinstance(path_or_img, str) else path_or_img
    if isinstance(img, str):
        img = _imread(img)
    gray = _to_gray(img)
    gray = _deskew(gray)

    h, w = gray.shape[:2]
    # binarise (lines end up white/255)
    if gray.max() - gray.min() < 10:
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    else:
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    # Open with long kernels to strip text, then dilate back to restore lines.
    h_open = cv2.getStructuringElement(cv2.MORPH_RECT, (max(25, w // 10), 1))
    v_open = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(25, h // 10)))
    h_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_open)
    h_mask = cv2.dilate(h_mask, h_open, iterations=1)
    v_mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_open)
    v_mask = cv2.dilate(v_mask, v_open, iterations=1)

    h_profile = h_mask.sum(axis=1) / 255.0
    v_profile = v_mask.sum(axis=0) / 255.0

    row_lines = _find_peaks(h_profile, ratio=min_line_ratio, min_gap=max(6, h // 120))
    col_lines = _find_peaks(v_profile, ratio=min_line_ratio, min_gap=max(6, w // 120))

    if len(row_lines) >= expect_min_rows and len(col_lines) >= expect_min_cols:
        row_lines = _expand_to_edges(row_lines, h, min_gap=max(6, h // 120))
        col_lines = _expand_to_edges(col_lines, w, min_gap=max(6, w // 120))
        confidence = 0.75 + 0.25 * min(1.0, (len(row_lines) * len(col_lines)) / 36.0)
        cells = _cells_from_lines(row_lines, col_lines, img.shape[1], img.shape[0])
        return GridResult(rows=len(row_lines) - 1, cols=len(col_lines) - 1,
                          cells=cells, confidence=min(confidence, 1.0), used_lines=True)

    # Fallback: cluster words later; return an empty grid with metadata.
    return GridResult(rows=0, cols=0, confidence=0.0, used_lines=False)


def _expand_to_edges(lines: list[int], size: int, min_gap: int) -> list[int]:
    result = list(lines)
    if not result:
        return [0, size]
    if result[0] > min_gap:
        result.insert(0, 0)
    if size - result[-1] > min_gap:
        result.append(size)
    return result


def _cells_from_lines(row_lines: list[int], col_lines: list[int], width: int, height: int) -> list[CellBox]:
    cells: list[CellBox] = []
    for r in range(len(row_lines) - 1):
        y0, y1 = row_lines[r], row_lines[r + 1]
        for c in range(len(col_lines) - 1):
            x0, x1 = col_lines[c], col_lines[c + 1]
            cells.append(CellBox(row=r, col=c, x=x0, y=y0, w=x1 - x0, h=y1 - y0))
    return cells


def assign_words_to_grid(grid: GridResult, words: list[tuple[str, int, int, int, int]]) -> GridResult:
    """Attach OCR words (``text, x, y, w, h``) to the nearest containing cell."""
    for cell in grid.cells:
        cell.text = ""
    for text, x, y, w, h in words:
        cx, cy = x + w / 2.0, y + h / 2.0
        best = None
        best_area = 0.0
        for cell in grid.cells:
            if cell.x - 2 <= cx <= cell.x + cell.w + 2 and cell.y - 2 <= cy <= cell.y + cell.h + 2:
                area = cell.w * cell.h
                if area > best_area:
                    best = cell
                    best_area = area
        if best is not None:
            best.text = (best.text + " " + text).strip()
        elif grid.used_lines and grid.cells:
            # nearest by centre
            best = min(grid.cells, key=lambda c: (c.x + c.w / 2 - cx) ** 2 + (c.y + c.h / 2 - cy) ** 2)
            best.text = (best.text + " " + text).strip()
    return grid


def cluster_words_to_grid(words: list[tuple[str, int, int, int, int]],
                          max_rows: int = 200, max_cols: int = 40) -> GridResult:
    """Cluster words into rows by y and columns by x (used when no ruling)."""
    if not words:
        return GridResult()
    sorted_words = sorted(words, key=lambda wd: (wd[2], wd[1]))
    # rows by y overlap
    rows_y: list[list[int]] = []
    for text, x, y, w, h in sorted_words:
        placed = False
        for band in rows_y:
            top = min(wd[2] for wd in band)
            bottom = max(wd[2] + wd[4] for wd in band)
            if y <= bottom and y + h >= top:
                band.append((text, x, y, w, h))
                placed = True
                break
        if not placed:
            rows_y.append([(text, x, y, w, h)])
    if len(rows_y) > max_rows:
        rows_y = rows_y[:max_rows]

    # columns by x clustering across all rows
    all_x = sorted({wd[1] for wd in words})
    col_bands: list[list[int]] = []
    for x in all_x:
        placed = False
        for band in col_bands:
            if abs(x - band[-1]) <= max(12, (max(all_x) - min(all_x)) * 0.02):
                band.append(x)
                placed = True
                break
        if not placed:
            col_bands.append([x])
    if len(col_bands) > max_cols:
        col_bands = col_bands[:max_cols]

    cells: list[CellBox] = []
    for r, band in enumerate(rows_y):
        for text, x, y, w, h in band:
            col = _nearest_col(x, col_bands)
            cells.append(CellBox(row=r, col=col, x=x, y=y, w=w, h=h, text=text))
    rows = len(rows_y)
    cols = len(col_bands)
    grid = GridResult(rows=rows, cols=cols, cells=cells, confidence=0.4, used_lines=False)
    return assign_words_to_grid(grid, words)


def _nearest_col(x: int, bands: list[list[int]]) -> int:
    centres = [sum(b) / len(b) for b in bands]
    return int(min(range(len(centres)), key=lambda i: abs(centres[i] - x)))


def build_grid_from_ocr(prepared_image, words: list[tuple[str, int, int, int, int]],
                        force_cluster: bool = False) -> GridResult:
    """Detect a grid on the same image the words came from, then attach them.

    *prepared_image* must be the preprocessed array that produced *words* so
    that box coordinates and grid cells share one coordinate system.  Falls back
    to word clustering when no ruling is found.
    """
    grid = detect_grid(prepared_image) if not force_cluster else GridResult()
    if grid.used_lines and grid.rows and grid.cols:
        return assign_words_to_grid(grid, words)
    if not grid.cells:
        return cluster_words_to_grid(words)
    return grid
