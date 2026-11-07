"""User-facing exceptions.

Everything that can be caused by the user's input is wrapped in an
``AppError`` subclass so the GUI can show a readable message instead of a
traceback.
"""

from __future__ import annotations


class AppError(Exception):
    """Base class for all expected, user-facing failures."""


class DataError(AppError):
    """A source file could not be read or has an unexpected shape."""


class MappingError(AppError):
    """A mapping/extraction rule is malformed or references unknown fields."""


class FillError(AppError):
    """A target could not be written."""


class EngineError(AppError):
    """A spreadsheet/document engine is missing or unusable."""


class OcrError(AppError):
    """An OCR engine is missing or failed on an image."""
