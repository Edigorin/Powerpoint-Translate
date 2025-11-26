from __future__ import annotations

import json
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Iterable, List, Optional, Sequence

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
        glossary: Optional[Sequence[dict]] = None,
        context: Optional[str] = None,
        max_concurrent_requests: int = 1,
    ) -> List[TranslatableUnit]:
        translated: List[TranslatableUnit] = []
        batches = list(self._batch_units(units, max_batch_chars))
        if not batches:
            return translated

        def process_batch(idx: int, batch: List[TranslatableUnit]) -> Dict[str, str]:
            try:
                return self._translate_batch(batch, source_lang, target_lang, glossary, context)
            except Exception as exc:
                msg = str(exc).lower()
                if "context length" in msg or "maximum" in msg:
                    smaller = max(500, max_batch_chars // 2)
                    self.logger.warning("Batch %s failed due to size; retrying with smaller batches (%s chars)", idx, smaller)
                    retry_translations: Dict[str, str] = {}
                    for smaller_batch in self._batch_units(batch, smaller):
                        partial = self._translate_batch(smaller_batch, source_lang, target_lang, glossary, context)
                        retry_translations.update(partial)
                    return retry_translations
                raise

        if max_concurrent_requests > 1 and len(batches) > 1:
            with ThreadPoolExecutor(max_workers=max_concurrent_requests) as executor:
                futures = {executor.submit(process_batch, idx, batch): idx for idx, batch in enumerate(batches)}
                results: Dict[int, Dict[str, str]] = {}
                for future in as_completed(futures):
                    idx = futures[future]
                    results[idx] = future.result()
            ordered = [results[i] for i in sorted(results.keys())]
        else:
            ordered = [process_batch(idx, batch) for idx, batch in enumerate(batches)]

        for batch, translations in zip(batches, ordered):
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
        glossary: Optional[Sequence[dict]],
        context: Optional[str],
    ) -> Dict[str, str]:
        items = [{"id": u.id, "text": u.source_text} for u in batch]
        glossary_text = self._format_glossary(glossary) if glossary else ""
        context_text = f"Context: {context}\n" if context else ""
        user_content = (
            f"Translate each item from {source_lang or 'auto-detect'} to {target_lang}. "
            'Return JSON: {"translations": [{"id": "...", "text": "<translated>"} ...]} '
            "Do not drop or reorder items. Preserve placeholders and numbering. "
            "Only respond with valid JSON and nothing else.\n"
            f"{context_text}"
            f"{glossary_text}"
            "\n"
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

    def _format_glossary(self, glossary: Sequence[dict]) -> str:
        pairs = []
        for entry in glossary:
            src = entry.get("source")
            tgt = entry.get("target")
            if not src or not tgt:
                continue
            pairs.append(f"'{src}' -> '{tgt}'")
        if not pairs:
            return ""
        return "Glossary (must use these translations): " + "; ".join(pairs) + "\n"

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
