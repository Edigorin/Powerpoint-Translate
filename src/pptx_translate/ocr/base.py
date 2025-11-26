from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Iterable, List, Optional

from pptx_translate.models import OcrImageInput, OcrTextRegion


class OcrBackend(ABC):
    """
    Interface for OCR backends.
    """

    @abstractmethod
    def recognize(
        self,
        images: Iterable[OcrImageInput],
        config: Optional[dict] = None,
    ) -> List[OcrTextRegion]:
        """
        Perform OCR on images and return text regions.
        """
        raise NotImplementedError
