from __future__ import annotations

import argparse
import csv
import json
import logging
from pathlib import Path
from typing import List, Optional

from pptx_translate.backends import DummyBackend, OpenAIBackend, TranslationBackend
from pptx_translate.ocr import OcrBackend, PytesseractOcrBackend
from pptx_translate.translator import PptxTranslator, generate_run_id, sanitize_output_path


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
    parser.add_argument(
        "--glossary",
        type=Path,
        help='Path to glossary file (JSON array of {"source","target"} or CSV with source,target columns).',
    )
    parser.add_argument(
        "--context",
        type=str,
        help="Short context string to guide translation (e.g., domain, product).",
    )
    parser.add_argument(
        "--context-file",
        type=Path,
        help="Path to a text file containing context instructions.",
    )
    parser.add_argument(
        "--dedupe-text/--no-dedupe-text",
        dest="dedupe_text",
        default=True,
        action=argparse.BooleanOptionalAction,
        help="Deduplicate identical source strings to reduce calls (default: on).",
    )
    parser.add_argument("--run-id", type=str, help="Run identifier to embed and append to output filename.")
    parser.add_argument("--no-run-id", action="store_true", help="Do not append run_id to output filename.")
    parser.add_argument(
        "--profile",
        type=str,
        default="balanced",
        choices=["fast", "balanced", "quality"],
        help="Performance/quality profile.",
    )
    parser.add_argument(
        "--max-concurrent-requests",
        type=int,
        default=1,
        help="Maximum concurrent translation requests (for backends that support it).",
    )
    parser.add_argument(
        "--translate-images",
        action="store_true",
        help="Enable OCR + translation for text inside images (requires OCR backend).",
    )
    parser.add_argument(
        "--image-ocr-backend",
        type=str,
        default="pytesseract",
        help="OCR backend to use (e.g., pytesseract).",
    )
    parser.add_argument(
        "--image-ocr-config",
        type=Path,
        help="OCR backend config file (JSON).",
    )
    parser.add_argument(
        "--qa-report",
        type=Path,
        help="Write a QA report to this path (JSON by default).",
    )
    parser.add_argument(
        "--qa-report-format",
        type=str,
        default="json",
        choices=["json", "markdown"],
        help="QA report format.",
    )
    parser.add_argument(
        "--qa-threshold-length-ratio",
        type=float,
        default=1.6,
        help="Flag translations whose length exceeds this ratio vs source.",
    )
    parser.add_argument(
        "--generate-glossary",
        type=Path,
        help="Generate a glossary suggestion file and exit (no translation performed).",
    )
    parser.add_argument(
        "--deck-profile-out",
        type=Path,
        help="Write derived deck context/profile to this file (optional).",
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


def load_ocr_backend(name: str) -> OcrBackend:
    normalized = name.lower()
    if normalized in ("pytesseract", "tesseract"):
        return PytesseractOcrBackend()
    raise ValueError(f"Unknown OCR backend: {name}")


def load_glossary(path: Path) -> List[dict]:
    if not path.exists():
        raise FileNotFoundError(f"Glossary file not found: {path}")
    suffix = path.suffix.lower()
    if suffix == ".csv":
        entries: List[dict] = []
        with path.open("r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                src = row.get("source")
                tgt = row.get("target")
                if src and tgt:
                    entries.append({"source": src, "target": tgt})
        return entries
    # default JSON
    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        entries = []
        for item in data:
            if isinstance(item, dict) and "source" in item and "target" in item:
                entries.append({"source": item["source"], "target": item["target"]})
        return entries
    raise ValueError("Glossary file must be a JSON list of {source,target} or CSV with source,target columns")


def apply_profile_defaults(args: argparse.Namespace) -> None:
    """
    Adjust defaults based on the chosen profile.
    """
    if args.profile == "fast":
        args.max_batch_chars = max(args.max_batch_chars, 6000)
        args.max_concurrent_requests = max(args.max_concurrent_requests, 4)
        args.include_masters = False
        args.include_notes = False
    elif args.profile == "quality":
        args.max_batch_chars = min(args.max_batch_chars, 2500)
        args.max_concurrent_requests = 1


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=getattr(logging, args.log_level))

    apply_profile_defaults(args)

    backend = load_backend(args.backend, args.backend_config)

    run_id = args.run_id or generate_run_id()
    output_path = sanitize_output_path(
        args.input,
        args.output,
        args.target_lang,
        run_id=run_id,
        no_run_id=args.no_run_id,
    )

    glossary = load_glossary(args.glossary) if args.glossary else None
    context = None
    if args.context_file:
        context = args.context_file.read_text(encoding="utf-8")
    elif args.context:
        context = args.context

    ocr_backend: Optional[OcrBackend] = None
    if args.translate_images:
        ocr_backend = load_ocr_backend(args.image_ocr_backend)

    ocr_config = {}
    if args.image_ocr_config:
        with args.image_ocr_config.open("r", encoding="utf-8") as f:
            ocr_config = json.load(f)

    translator = PptxTranslator(
        backend=backend,
        include_notes=bool(args.include_notes),
        include_masters=bool(args.include_masters),
        max_batch_chars=args.max_batch_chars,
        dry_run=args.dry_run,
        dedupe_text=bool(args.dedupe_text),
        translate_images=bool(args.translate_images),
        ocr_backend=ocr_backend,
        ocr_config=ocr_config,
        max_concurrent_requests=args.max_concurrent_requests,
        profile=args.profile,
        qa_report_path=args.qa_report,
        qa_report_format=args.qa_report_format,
        qa_threshold_length_ratio=args.qa_threshold_length_ratio,
    )

    translated_units = translator.translate_file(
        input_path=args.input,
        output_path=output_path,
        source_lang=args.source_lang,
        target_lang=args.target_lang,
        glossary=glossary,
        context=context,
        run_id=run_id,
        generate_glossary_path=args.generate_glossary,
        deck_profile_path=args.deck_profile_out,
    )

    if args.dry_run:
        preview = [
            {"id": u.id, "location": u.location, "source": u.source_text, "translated": u.translated_text}
            for u in translated_units
        ]
        print(json.dumps(preview, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
