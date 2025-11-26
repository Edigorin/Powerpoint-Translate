from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Iterable, List, Optional, Sequence

from pptx_translate.models import TranslatableUnit


class TranslationBackend(ABC):
    """
    Interface for translation backends.
    """

    @abstractmethod
    def translate(
        self,
        units: Iterable[TranslatableUnit],
        source_lang: Optional[str],
        target_lang: str,
        max_batch_chars: int = 4000,
        glossary: Optional[Sequence[dict]] = None,
        context: Optional[str] = None,
        max_concurrent_requests: int = 1,
    ) -> List[TranslatableUnit]:
        """
        Translate a list of units and return updated units.
        """
        raise NotImplementedError
