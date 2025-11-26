from __future__ import annotations

import json
import logging
from typing import Dict, Iterable, List, Optional

from openai import OpenAI

from pptx_translate.backends.base import TranslationBackend
from pptx_translate.models import TranslatableUnit


class OpenAIBackend(TranslationBackend):
    """
    OpenAI Chat Completions translation backend.

    Expects an API key via config or OPENAI_API_KEY env var.
    """

    def __init__(
        self,
        model: str = "gpt-4o-mini",
        api_key: Optional[str] = None,
        base_url: Optional[str] = None,
        temperature: float = 0.0,
        system_prompt: Optional[str] = None,
    ) -> None:
        if api_key:
            self.client = OpenAI(api_key=api_key, base_url=base_url)
        else:
            self.client = OpenAI(base_url=base_url)
        self.model = model
        self.temperature = temperature
        self.system_prompt = system_prompt or "You are a translation engine. Return only translations, preserving placeholders and numbering. Do not add explanations."
        self.logger = logging.getLogger(__name__)

    def translate(
        self,
        units: Iterable[TranslatableUnit],
        source_lang: Optional[str],
        target_lang: str,
        max_batch_chars: int = 4000,
    ) -> List[TranslatableUnit]:
        translated: List[TranslatableUnit] = []
        for batch in self._batch_units(units, max_batch_chars):
            translations = self._translate_batch(batch, source_lang, target_lang)
            for unit in batch:
                text = translations.get(unit.id)
                if text is None:
                    self.logger.warning("Missing translation for id %s; falling back to source text", unit.id)
                    text = unit.source_text
                translated.append(
                    TranslatableUnit(
                        id=unit.id,
                        location=unit.location,
                        source_text=unit.source_text,
                        translated_text=text,
                        context=unit.context,
                    )
                )
        return translated

    def _translate_batch(
        self,
        batch: List[TranslatableUnit],
        source_lang: Optional[str],
        target_lang: str,
    ) -> Dict[str, str]:
        items = [{"id": u.id, "text": u.source_text} for u in batch]
        user_content = (
            f"Translate each item from {source_lang or 'auto-detect'} to {target_lang}. "
            'Return JSON: {"translations": [{"id": "...", "text": "<translated>"} ...]} '
            "Do not drop or reorder items. Preserve placeholders and numbering. "
            "Only respond with valid JSON and nothing else.\n\n"
            f"Items: {json.dumps(items, ensure_ascii=False)}"
        )
        response = self.client.chat.completions.create(
            model=self.model,
            temperature=self.temperature,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": user_content},
            ],
        )
        content = response.choices[0].message.content
        data: Dict[str, List[str]] = json.loads(content)
        translations_list = data.get("translations")
        if not isinstance(translations_list, list):
            raise RuntimeError("OpenAI response missing 'translations' list")

        mapping: Dict[str, str] = {}
        for item in translations_list:
            if not isinstance(item, dict):
                continue
            item_id = item.get("id")
            text = item.get("text")
            if item_id is None or text is None:
                continue
            mapping[str(item_id)] = str(text)
        return mapping

    def _batch_units(self, units: Iterable[TranslatableUnit], max_batch_chars: int) -> List[List[TranslatableUnit]]:
        batches: List[List[TranslatableUnit]] = []
        current: List[TranslatableUnit] = []
        current_size = 0
        for unit in units:
            size = len(unit.source_text)
            if current and current_size + size > max_batch_chars:
                batches.append(current)
                current = [unit]
                current_size = size
            else:
                current.append(unit)
                current_size += size
        if current:
            batches.append(current)
        return batches
