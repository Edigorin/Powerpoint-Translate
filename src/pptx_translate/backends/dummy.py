from __future__ import annotations

from typing import Iterable, List, Optional, Sequence

from pptx_translate.backends.base import TranslationBackend
from pptx_translate.models import TranslatableUnit


class DummyBackend(TranslationBackend):
    """
    Development backend that prefixes each string with the target language code.
    """

    def translate(
        self,
        units: Iterable[TranslatableUnit],
        source_lang: Optional[str],
        target_lang: str,
        max_batch_chars: int = 4000,
        glossary: Optional[Sequence[dict]] = None,
        context: Optional[str] = None,
    ) -> List[TranslatableUnit]:
        updated: List[TranslatableUnit] = []
        for unit in units:
            translated = f"[{target_lang}] {unit.source_text}"
            updated.append(
                TranslatableUnit(
                    id=unit.id,
                    location=unit.location,
                    source_text=unit.source_text,
                    translated_text=translated,
                    context=unit.context,
                )
            )
        return updated
