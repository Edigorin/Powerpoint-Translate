from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
from typing import Optional

from pptx_translate.backends import DummyBackend, OpenAIBackend, TranslationBackend
from pptx_translate.translator import PptxTranslator, sanitize_output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="pptx-translate",
        description="Translate PowerPoint .pptx files while preserving layout/formatting.",
    )
    parser.add_argument("input", type=Path, help="Path to input .pptx file")
    parser.add_argument("-o", "--output", type=Path, help="Path to output .pptx file")
    parser.add_argument("--source-lang", type=str, help="Source language code (optional if backend auto-detects)")
    parser.add_argument("--target-lang", type=str, required=True, help="Target language code")
    parser.add_argument("--backend", type=str, default="dummy", help="Translation backend id (default: dummy)")
    parser.add_argument(
        "--backend-config",
        type=Path,
        help="Path to backend config file (JSON). Used for future real backends.",
    )
    parser.add_argument("--include-notes", dest="include_notes", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--include-masters", dest="include_masters", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--dry-run", action="store_true", help="Extract and translate text but do not write output file")
    parser.add_argument(
        "--max-batch-chars",
        type=int,
        default=4000,
        help="Maximum characters per translation batch (backend-specific).",
    )
    parser.add_argument(
        "--log-level",
        type=str,
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging verbosity.",
    )
    return parser.parse_args()


def load_backend(name: str, config_path: Optional[Path] = None) -> TranslationBackend:
    config = {}
    if config_path:
        with config_path.open("r", encoding="utf-8") as f:
            config = json.load(f)
    normalized = name.lower()
    if normalized == "dummy":
        return DummyBackend()
    if normalized == "openai":
        return OpenAIBackend(**config)
    raise ValueError(f"Unknown backend: {name}")


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level))

    backend = load_backend(args.backend, args.backend_config)

    output_path = sanitize_output_path(args.input, args.output, args.target_lang)

    translator = PptxTranslator(
        backend=backend,
        include_notes=bool(args.include_notes),
        include_masters=bool(args.include_masters),
        max_batch_chars=args.max_batch_chars,
        dry_run=args.dry_run,
    )

    translated_units = translator.translate_file(
        input_path=args.input,
        output_path=output_path,
        source_lang=args.source_lang,
        target_lang=args.target_lang,
    )

    if args.dry_run:
        preview = [{"id": u.id, "location": u.location, "source": u.source_text, "translated": u.translated_text} for u in translated_units]
        print(json.dumps(preview, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
