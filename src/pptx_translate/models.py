from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple


@dataclass
class TranslatableUnit:
    """
    Represents a single text node to be translated.
    """

    id: str
    location: str
    source_text: str
    translated_text: Optional[str] = None
    context: Optional[str] = None


@dataclass
class OcrImageInput:
    """
    Represents an image to be sent to OCR.
    """

    slide_index: int
    shape_index: int
    image_name: str
    image_bytes: bytes
    width_px: int
    height_px: int


@dataclass
class OcrTextRegion:
    """
    OCR result region associated with an image.
    """

    slide_index: int
    shape_index: int
    image_name: str
    bbox: Tuple[int, int, int, int]  # left, top, width, height in pixels (image coords)
    source_text: str
    translated_text: Optional[str] = None
    unit_id: Optional[str] = None
