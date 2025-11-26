from __future__ import annotations

import logging
import io
from typing import Iterable, List, Optional

from PIL import Image

try:
    import pytesseract
except ImportError:  # pragma: no cover - optional dependency
    pytesseract = None

from pptx_translate.models import OcrImageInput, OcrTextRegion
from pptx_translate.ocr.base import OcrBackend


class PytesseractOcrBackend(OcrBackend):
    """
    OCR backend using pytesseract.
    """

    def __init__(self) -> None:
        if pytesseract is None:
            raise ImportError("pytesseract is required for OCR; install with `pip install pytesseract pillow` or extras.")
        self.logger = logging.getLogger(__name__)

    def recognize(
        self,
        images: Iterable[OcrImageInput],
        config: Optional[dict] = None,
    ) -> List[OcrTextRegion]:
        regions: List[OcrTextRegion] = []
        tesseract_config = config.get("tesseract_config") if config else None
        lang = config.get("lang") if config else None
        for img in images:
            pil_image = Image.open(io.BytesIO(img.image_bytes))
            data = pytesseract.image_to_data(
                pil_image,
                config=tesseract_config,
                lang=lang,
                output_type=pytesseract.Output.DICT,
            )
            for i, text in enumerate(data["text"]):
                if not text or text.strip() == "":
                    continue
                left = int(data["left"][i])
                top = int(data["top"][i])
                width = int(data["width"][i])
                height = int(data["height"][i])
                regions.append(
                    OcrTextRegion(
                        slide_index=img.slide_index,
                        shape_index=img.shape_index,
                        image_name=img.image_name,
                        bbox=(left, top, width, height),
                        source_text=text,
                    )
                )
        return regions
