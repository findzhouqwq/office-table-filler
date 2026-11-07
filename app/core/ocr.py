"""OCR engine abstraction.

The default engine is RapidOCR (ONNX models bundled by the package), which
handles Chinese and English well with no external download.  Tesseract and the
Windows built-in OCR are optional alternatives.
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass, field

import numpy as np

from .errors import OcrError

try:  # pragma: no cover - import guard
    import cv2

    _HAS_CV2 = True
except Exception:  # pragma: no cover
    _HAS_CV2 = False


@dataclass
class OcrWord:
    text: str
    x: int
    y: int
    w: int
    h: int
    confidence: float = 0.0


@dataclass
class OcrResult:
    text: str
    words: list[OcrWord] = field(default_factory=list)
    engine: str = ""
    elapsed: float = 0.0
    width: int = 0
    height: int = 0

    def words_as_tuples(self) -> list[tuple[str, int, int, int, int]]:
        return [(w.text, w.x, w.y, w.w, w.h) for w in self.words]


# ---------------------------------------------------------------------------
# Base engine
# ---------------------------------------------------------------------------

class OcrEngine:
    name = "base"

    def available(self) -> bool:  # pragma: no cover - default
        return True

    # -- image preparation ---------------------------------------------------
    def prepare(self, image_path: str) -> "np.ndarray":
        """Load an image and return the preprocessed grayscale array."""
        img = self._imread(image_path)
        return preprocess_for_ocr(img)

    def recognize(self, source: str | "np.ndarray") -> OcrResult:
        """Recognise text from a file path or a pre-loaded image array."""
        if isinstance(source, str):
            img = self.prepare(source)
            result = self._recognize_array(img)
            result.width, result.height = img.shape[1], img.shape[0]
            return result
        img = preprocess_for_ocr(source)
        result = self._recognize_array(img)
        result.width, result.height = img.shape[1], img.shape[0]
        return result

    def _recognize_array(self, img: "np.ndarray") -> OcrResult:
        raise NotImplementedError

    def recognize_with_image(self, image_path: str) -> tuple[OcrResult, "np.ndarray"]:
        """Like :meth:`recognise` but also returns the preprocessed image.

        This guarantees the OCR word boxes and any grid detection share one
        coordinate system.
        """
        img = self.prepare(image_path)
        result = self._recognize_array(img)
        result.width, result.height = img.shape[1], img.shape[0]
        return result, img

    @staticmethod
    def _imread(path: str) -> "np.ndarray":
        if not _HAS_CV2:
            raise OcrError("未安装 opencv-python，无法读取图片。")
        data = np.fromfile(path, dtype=np.uint8)
        if data.size == 0:
            raise OcrError(f"无法读取图片：{path}")
        img = cv2.imdecode(data, cv2.IMREAD_COLOR)
        if img is None:
            raise OcrError(f"图片解码失败：{path}")
        return img


def preprocess_for_ocr(img: "np.ndarray", max_side: int = 2400) -> "np.ndarray":
    """Grayscale + contrast stretch + upscale small images."""
    if img is None:
        return img
    if img.ndim == 3:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    else:
        gray = img
    h, w = gray.shape[:2]
    scale = max_side / max(h, w)
    if scale > 1.0:
        gray = cv2.resize(gray, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
    # CLAHE for uneven lighting
    clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
    return clahe.apply(gray)


# ---------------------------------------------------------------------------
# RapidOCR
# ---------------------------------------------------------------------------

class RapidOcrEngine(OcrEngine):
    name = "rapidocr"

    def __init__(self) -> None:
        try:
            from rapidocr_onnxruntime import RapidOCR  # type: ignore
        except Exception as exc:
            raise OcrError("未安装 rapidocr-onnxruntime，无法使用内置中文识别引擎。") from exc
        self._engine = RapidOCR()

    def _recognize_array(self, img: "np.ndarray") -> OcrResult:
        started = time.time()
        result, _ = self._engine(img)
        words: list[OcrWord] = []
        texts: list[str] = []
        if result:
            for item in result:
                box, text, score = item[0], item[1], float(item[2])
                xs = [p[0] for p in box]
                ys = [p[1] for p in box]
                x0, y0, x1, y1 = int(min(xs)), int(min(ys)), int(max(xs)), int(max(ys))
                words.append(OcrWord(text=text, x=x0, y=y0, w=x1 - x0, h=y1 - y0, confidence=score))
                texts.append(text)
        text = "\n".join(t for t in texts if t)
        return OcrResult(text=text, words=words, engine=self.name,
                         elapsed=time.time() - started)


# ---------------------------------------------------------------------------
# Tesseract
# ---------------------------------------------------------------------------

class TesseractEngine(OcrEngine):
    name = "tesseract"

    def __init__(self, lang: str = "chi_sim+eng") -> None:
        self.lang = lang
        try:
            import pytesseract  # type: ignore
        except Exception as exc:
            raise OcrError("未安装 pytesseract，无法使用 Tesseract 引擎。") from exc
        self._pytesseract = pytesseract
        if not self._tesseract_ready():
            raise OcrError(
                "未找到 Tesseract 可执行文件。请安装后把它加入 PATH，"
                "或在界面中改用 RapidOCR 引擎。"
            )

    def _tesseract_ready(self) -> bool:
        try:
            self._pytesseract.get_tesseract_version()
            return True
        except Exception:
            return False

    def _recognize_array(self, img: "np.ndarray") -> OcrResult:
        started = time.time()
        data = self._pytesseract.image_to_data(
            img, lang=self.lang, output_type=self._pytesseract.Output.DICT
        )
        words: list[OcrWord] = []
        for i, raw in enumerate(data.get("text", [])):
            text = (raw or "").strip()
            if not text or int(data["level"][i]) != 5:
                continue  # only word level
            words.append(OcrWord(
                text=text,
                x=int(data["left"][i]), y=int(data["top"][i]),
                w=int(data["width"][i]), h=int(data["height"][i]),
                confidence=float(data["conf"][i] or 0) / 100.0,
            ))
        full_text = self._pytesseract.image_to_string(img, lang=self.lang)
        return OcrResult(text=full_text, words=words, engine=self.name,
                         elapsed=time.time() - started)


# ---------------------------------------------------------------------------
# Windows built-in OCR
# ---------------------------------------------------------------------------

class WindowsOcrEngine(OcrEngine):
    name = "windows"

    def __init__(self, language: str = "zh-Hans-CN") -> None:
        self.language = language
        try:
            from winsdk.windows.media.ocr import OcrEngine as WinOcrEngine  # type: ignore
        except Exception as exc:
            raise OcrError("未安装 winsdk，无法使用 Windows 内置 OCR。") from exc
        try:
            self._engine = WinOcrEngine.try_create_from_language(self.language)
        except Exception as exc:
            raise OcrError(f"Windows OCR 初始化失败（语言包可能未安装）：{exc}") from exc

    def recognize(self, image_path: str) -> OcrResult:
        from winsdk.windows.graphics.imaging import (  # type: ignore
            BitmapDecoder, ImagePixelFormat, BitmapFrame,
        )
        from winsdk.windows.storage.streams import FileRandomAccessStream, StorageFile  # type: ignore

        started = time.time()
        stream = FileRandomAccessStream.open_async(StorageFile.get_file_from_path_async(image_path).get_result())
        frame = BitmapFrame.decode_async(stream, BitmapDecoder.png_decoder).get_result()
        result = self._engine.recognize_async(frame).get_result()
        words: list[OcrWord] = []
        texts: list[str] = []
        for line in result.lines:
            texts.append(line.text)
            for word in line.words:
                box = word.bounding_box
                words.append(OcrWord(text=word.text, x=int(box.x), y=int(box.y),
                                     w=int(box.width), h=int(box.height),
                                     confidence=float(word.confidence or 0)))
        return OcrResult(text="\n".join(texts), words=words, engine=self.name,
                         elapsed=time.time() - started)

    def recognize_with_image(self, image_path: str) -> tuple[OcrResult, "np.ndarray"]:
        """Scale Windows OCR boxes into the preprocessed coordinate space."""
        result = self.recognize(image_path)
        img_color = self._imread(image_path)
        prepared = preprocess_for_ocr(img_color)
        scale_x = prepared.shape[1] / img_color.shape[1]
        scale_y = prepared.shape[0] / img_color.shape[0]
        for word in result.words:
            word.x = int(word.x * scale_x)
            word.y = int(word.y * scale_y)
            word.w = int(word.w * scale_x)
            word.h = int(word.h * scale_y)
        result.width, result.height = prepared.shape[1], prepared.shape[0]
        return result, prepared


# ---------------------------------------------------------------------------
# Manual / test engine
# ---------------------------------------------------------------------------

class ManualOcrEngine(OcrEngine):
    name = "manual"

    def __init__(self, words: list[tuple[str, int, int, int, int]], text: str = "") -> None:
        self._words = [OcrWord(t, x, y, w, h) for t, x, y, w, h in words]
        self._text = text or "\n".join(w.text for w in self._words)

    def recognize(self, image_path: str) -> OcrResult:
        return OcrResult(text=self._text, words=self._words, engine=self.name)

    def recognize_with_image(self, image_path: str) -> tuple[OcrResult, "np.ndarray"]:
        # Word boxes were authored in the original image coordinates, so skip
        # preprocessing here and return the original array for grid detection.
        img = self._imread(image_path)
        result = OcrResult(text=self._text, words=self._words, engine=self.name)
        result.width, result.height = img.shape[1], img.shape[0]
        return result, img


# ---------------------------------------------------------------------------
# Factory
# ---------------------------------------------------------------------------

def available_engines() -> list[str]:
    """List usable OCR engines without loading any model.

    Constructing :class:`RapidOcrEngine` loads the ONNX models (tens of MB),
    so probing by import/executable avoids a slow startup for a check that
    only feeds a combo box.
    """
    import importlib.util
    import shutil

    def module_exists(name: str) -> bool:
        try:
            return importlib.util.find_spec(name) is not None
        except (ImportError, ValueError):
            return False

    engines: list[str] = []
    if module_exists("rapidocr_onnxruntime"):
        engines.append("rapidocr")
    if module_exists("pytesseract") and shutil.which("tesseract"):
        engines.append("tesseract")
    if module_exists("winsdk"):
        engines.append("windows")
    return engines


def get_ocr_engine(prefer: str = "auto") -> OcrEngine:
    """Return a ready OCR engine, preferring the requested one."""
    order = {
        "auto": ["rapidocr", "tesseract", "windows"],
        "rapidocr": ["rapidocr"],
        "tesseract": ["tesseract"],
        "windows": ["windows"],
        "none": [],
    }.get(prefer, ["rapidocr", "tesseract", "windows"])
    last_error: Exception | None = None
    for name in order:
        try:
            if name == "rapidocr":
                return RapidOcrEngine()
            if name == "tesseract":
                return TesseractEngine()
            if name == "windows":
                return WindowsOcrEngine()
        except Exception as exc:
            last_error = exc
            continue
    if last_error:
        raise OcrError(
            f"没有可用的 OCR 引擎（{last_error}）。"
            "请在界面中安装 rapidocr-onnxruntime，或安装 Tesseract。"
        )
    raise OcrError("没有可用的 OCR 引擎。")
