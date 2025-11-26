from __future__ import annotations

import logging
import os
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET

from pptx_translate.backends import TranslationBackend
from pptx_translate.models import TranslatableUnit

NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}


@dataclass
class DocumentPart:
    """
    Holds a parsed XML part and its translatable text nodes.
    """

    path: Path
    tree: ET.ElementTree
    nodes: List[Tuple[ET.Element, TranslatableUnit]]


class PptxTranslator:
    """
    Translate PowerPoint files while preserving layout/formatting.
    """

    def __init__(
        self,
        backend: TranslationBackend,
        include_notes: bool = True,
        include_masters: bool = True,
        max_batch_chars: int = 4000,
        dry_run: bool = False,
        dedupe_text: bool = True,
    ) -> None:
        self.backend = backend
        self.include_notes = include_notes
        self.include_masters = include_masters
        self.max_batch_chars = max_batch_chars
        self.dry_run = dry_run
        self.dedupe_text = dedupe_text
        self.logger = logging.getLogger(__name__)

    def translate_file(
        self,
        input_path: Path,
        output_path: Path,
        source_lang: str | None,
        target_lang: str,
        glossary: list[dict] | None = None,
        context: str | None = None,
    ) -> List[TranslatableUnit]:
        """
        Main entrypoint: translate `input_path` into `output_path`.
        Returns list of translated units.
        """
        input_path = input_path.resolve()
        output_path = output_path.resolve()
        self._validate_input(input_path)

        with TemporaryDirectory() as tmp_dir:
            temp_root = Path(tmp_dir)
            self._unpack_pptx(input_path, temp_root)
            parts = self._load_parts(temp_root)
            units = [unit for part in parts for _, unit in part.nodes]

            self.logger.info("Extracted %d text units", len(units))

            translated_units = self._translate_units(
                units,
                source_lang=source_lang,
                target_lang=target_lang,
                glossary=glossary,
                context=context,
            )
            translated_map: Dict[str, TranslatableUnit] = {u.id: u for u in translated_units}

            if not self.dry_run:
                self._inject_translations(parts, translated_map)
                self._repack_pptx(temp_root, output_path)
                self.logger.info("Wrote translated file to %s", output_path)
            else:
                self.logger.info("Dry run mode: no output file written")

            return translated_units

    def _validate_input(self, input_path: Path) -> None:
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")
        if input_path.suffix.lower() != ".pptx":
            raise ValueError("Input file must be a .pptx")

    def _unpack_pptx(self, input_path: Path, temp_root: Path) -> None:
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(temp_root)
        self.logger.debug("Extracted pptx to %s", temp_root)

    def _load_parts(self, temp_root: Path) -> List[DocumentPart]:
        paths = self._discover_xml_parts(temp_root)
        parts: List[DocumentPart] = []
        counter = 0

        for path in paths:
            tree = ET.parse(path)
            root = tree.getroot()
            nodes: List[Tuple[ET.Element, TranslatableUnit]] = []
            for idx, elem in enumerate(root.findall(".//a:t", namespaces=NAMESPACES)):
                text = elem.text if elem.text is not None else ""
                if text == "":
                    continue
                counter += 1
                location = f"{path.relative_to(temp_root)}::a:t[{idx}]"
                unit_id = f"t{counter}"
                unit = TranslatableUnit(
                    id=unit_id,
                    location=str(location),
                    source_text=text,
                    context=None,
                )
                nodes.append((elem, unit))
            parts.append(DocumentPart(path=path, tree=tree, nodes=nodes))

        return parts

    def _translate_units(
        self,
        units: List[TranslatableUnit],
        source_lang: str | None,
        target_lang: str,
        glossary: list[dict] | None,
        context: str | None,
    ) -> List[TranslatableUnit]:
        if not self.dedupe_text:
            return self.backend.translate(
                units,
                source_lang=source_lang,
                target_lang=target_lang,
                max_batch_chars=self.max_batch_chars,
                glossary=glossary,
                context=context,
            )

        text_to_units: Dict[str, List[TranslatableUnit]] = {}
        unique_units: List[TranslatableUnit] = []
        for unit in units:
            key = unit.source_text
            if key not in text_to_units:
                text_to_units[key] = [unit]
                unique_units.append(unit)
            else:
                text_to_units[key].append(unit)

        self.logger.info("Deduped %d texts down to %d unique entries", len(units), len(unique_units))

        translated_unique = self.backend.translate(
            unique_units,
            source_lang=source_lang,
            target_lang=target_lang,
            max_batch_chars=self.max_batch_chars,
            glossary=glossary,
            context=context,
        )
        by_text: Dict[str, str] = {u.source_text: (u.translated_text or u.source_text) for u in translated_unique}

        translated_all: List[TranslatableUnit] = []
        for unit in units:
            translated_text = by_text.get(unit.source_text, unit.source_text)
            translated_all.append(
                TranslatableUnit(
                    id=unit.id,
                    location=unit.location,
                    source_text=unit.source_text,
                    translated_text=translated_text,
                    context=unit.context,
                )
            )
        return translated_all

    def _inject_translations(
        self,
        parts: List[DocumentPart],
        translated_map: Dict[str, TranslatableUnit],
    ) -> None:
        for part in parts:
            for elem, unit in part.nodes:
                translated = translated_map.get(unit.id)
                if translated and translated.translated_text is not None:
                    elem.text = translated.translated_text
            part.tree.write(part.path, xml_declaration=True, encoding="utf-8", method="xml")

    def _repack_pptx(self, temp_root: Path, output_path: Path) -> None:
        if output_path.exists():
            output_path.unlink()
        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for folder, _, files in os.walk(temp_root):
                for filename in files:
                    abs_path = Path(folder) / filename
                    arcname = abs_path.relative_to(temp_root)
                    # Ensure POSIX separators inside the zip
                    zf.write(abs_path, arcname=str(arcname).replace(os.sep, "/"))

    def _discover_xml_parts(self, temp_root: Path) -> List[Path]:
        """
        Collect relevant XML parts that may contain text.
        """
        candidates: List[Path] = []
        slides_dir = temp_root / "ppt" / "slides"
        notes_dir = temp_root / "ppt" / "notesSlides"
        masters_dir = temp_root / "ppt" / "slideMasters"
        layouts_dir = temp_root / "ppt" / "slideLayouts"

        candidates.extend(sorted(slides_dir.glob("slide*.xml")))
        if self.include_notes:
            candidates.extend(sorted(notes_dir.glob("notesSlide*.xml")))
        if self.include_masters:
            candidates.extend(sorted(masters_dir.glob("slideMaster*.xml")))
            candidates.extend(sorted(layouts_dir.glob("slideLayout*.xml")))

        existing = [path for path in candidates if path.exists()]
        return existing


def sanitize_output_path(input_path: Path, user_output: Path | None, target_lang: str) -> Path:
    """
    Produce a default output path if the user didn't supply one.
    """
    if user_output:
        return user_output
    suffix = f".{target_lang}.pptx"
    return input_path.with_name(f"{input_path.stem}{suffix}")
