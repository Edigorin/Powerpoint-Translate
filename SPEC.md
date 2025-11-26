# PowerPoint Translator – Functional & Technical Spec

## 1. Problem & Constraints

We need a tool that can translate the textual content of PowerPoint `.pptx` files while keeping:
- Slide layouts exactly the same.
- Positions, sizes, and alignment of all shapes and text boxes.
- Fonts, font sizes, colors, and bullet levels.
- Tables, charts, SmartArt (where possible), and speaker notes.

The tool must not:
- Rebuild slides from scratch.
- Change slide master/layout structure.
- Modify non-text content (images, shapes, animations) except to update text labels.

## 2. Scope of Translation

Translate:
- Slide titles and body text.
- Text in shapes, text boxes, and placeholders.
- Bullet lists and numbered lists.
- Table cell text.
- Chart labels and data labels where stored as normal text.
- Speaker notes.

Optional / later:
- SmartArt text.
- Embedded objects (e.g., Word/Excel).

Do not translate:
- Filenames, embedded binaries, or image data.
- Internal IDs, XML attributes, or style names.

## 3. User Stories

- As a user, I can run `pptx-translate input.pptx --source en --target de` and get `input.de.pptx` with the same layout but German text.
- As a user, I can choose the translation backend (e.g. DeepL, Azure, OpenAI, local model).
- As a user, I can run in “dry run” mode and see an extract of source/target text pairs without writing a new file.
- As a user, I can exclude certain slides, layouts, or patterns (e.g. `--exclude-notes`, `--exclude-title`).

## 4. CLI Interface (initial proposal)

Command: `pptx-translate`

Required arguments:
- `input`: path to input `.pptx` file.

Options:
- `-o, --output`: path for output `.pptx` (default: `<input>.<target>.pptx`).
- `--source-lang`: source language code (e.g. `en`); optional if backend can auto-detect.
- `--target-lang`: target language code (e.g. `de`).
- `--backend`: translation backend ID (`deepl`, `azure`, `openai`, `dummy`).
- `--backend-config`: path to a config file or JSON with backend-specific settings and API keys.
- `--include-notes` / `--exclude-notes`: whether to translate speaker notes (default: include).
- `--include-masters` / `--exclude-masters`: whether to translate slide masters/layouts.
- `--dry-run`: do not write output; print or export extracted text and translations.
- `--max-batch-chars`: max characters per translation request (backend-specific).

Exit codes:
- `0`: success.
- `1`: general failure.
- `2`: invalid input / arguments.
- `3`: translation backend error.

## 5. Architecture

Language: Python (3.10+).

High-level modules:
- `pptx_translate.cli`: argument parsing and orchestration.
- `pptx_translate.io`: file handling (open, unzip, zip, temp dirs).
- `pptx_translate.ooxml`: Open XML parsing/writing for `.pptx`.
- `pptx_translate.extract`: extract translatable units from the OOXML.
- `pptx_translate.translate`: translation orchestration and batching.
- `pptx_translate.backends`: pluggable backend interfaces and implementations.
- `pptx_translate.logging`: structured logging helpers.

Key requirement: Whenever possible, operate at the level of `<a:t>` text nodes in the underlying OOXML so that geometry and styling elements (e.g. `<a:rPr>`, `<p:spPr>`) remain untouched.

## 6. Data Model

### 6.1 Translatable Unit

`TranslatableUnit`:
- `id`: stable identifier (e.g. `slide_3_shape_5_paragraph_2_run_1`).
- `location`: structured path (slide index, notes/master, element type).
- `source_text`: original text from the OOXML.
- `translated_text`: populated after translation.
- `context`: optional context string, hints (e.g., slide title vs bullet).

### 6.2 Backend Interface

`TranslationBackend` interface:
- `translate(self, units: list[TranslatableUnit], source_lang: str | None, target_lang: str) -> list[TranslatableUnit]`

Responsibilities:
- Group units into batches (respecting `max_batch_chars`).
- Call external API or local model.
- Fill `translated_text` while preserving ordering.

At least one “dummy” backend will be implemented for development/testing (e.g., reversing strings or prefixing with `[XX]`).

## 7. PPTX / OOXML Handling

### 7.1 File Processing

- Treat `.pptx` as a ZIP archive.
- Use `zipfile` to:
  - Extract to a temp directory.
  - Read XML from:
    - `ppt/slides/slide*.xml`
    - `ppt/notesSlides/notesSlide*.xml`
    - `ppt/slideMasters/slideMaster*.xml`
    - `ppt/slideLayouts/slideLayout*.xml`
  - Write updated XML back and re-package into a `.pptx`.

### 7.2 XML Parsing Strategy

- Use `lxml` or `xml.etree.ElementTree`.
- Search for text elements typically under namespaces:
  - DrawingML text: `<a:t>` nodes.
- Do not:
  - Change or reorder elements.
  - Modify attributes like `x`, `y`, `cx`, `cy`, or style attributes.

### 7.3 Preserving Formatting

- Keep the existing `<a:r>` (run) and `<a:rPr>` (run properties) structures.
- Only replace the text content of `<a:t>` with `translated_text`.
- If the translation changes length or line breaks, avoid changing structure unless absolutely necessary:
  - By default, keep one `<a:t>` node per original node.
  - If backend returns multiple lines for a single unit, insert line breaks (e.g. `&#10;`) rather than new XML elements.

### 7.4 IDs and Relationships

- Do not change:
  - `r:id` relationships.
  - `p:cNvPr`, `p:nvSpPr`, or similar IDs.
  - Slide numbers, master/layout references.

## 8. Translation Backends

Initial backends (spec level; implementation later):

- `DummyBackend`:
  - For each unit, sets `translated_text = f\"[{target_lang}] {source_text}\"`.
  - No external calls.

- `OpenAIBackend`:
  - Uses OpenAI Chat Completions (e.g., GPT-4o) to translate batches of text.
  - Respects a max tokens/characters limit.
  - Accepts config JSON for `model`, `temperature`, `base_url`, and `api_key`.
  - Optional glossaries / style hints via system prompts.

Backend configuration:
- Read from:
  - CLI `--backend-config` (JSON/YAML file) or
  - Environment variables (e.g. `OPENAI_API_KEY`).

## 9. Error Handling & Logging

- Validate input file existence and extension (`.pptx`).
- Fail fast if backend configuration is invalid or API key is missing.
- If a single translation call fails:
  - Retry with exponential backoff (configurable).
  - On repeated failure, abort and return a non-zero exit code.
- Log:
  - Which slides/parts were processed.
  - Number of units translated.
  - Backend calls and failures (without leaking secrets).

## 10. Testing Strategy

- Unit tests:
  - XML parsing and extraction on small sample slides.
  - Round-tripping: extract units, apply dummy translations, reinsert, and compare non-text XML unchanged.
- Integration tests:
  - End-to-end run on sample PPTX with various layouts.
  - Verify:
    - No change in slide/master/notes counts.
    - No change in shape counts per slide.
    - Only `<a:t>` inner text differs.

## 11. Future Extensions

- GUI wrapper for drag-and-drop translation.
- Batch processing of folders.
- Glossary and terminology management.
- Per-shape or per-placeholder rules (e.g. never translate certain placeholders).
- Post-translation QA report (detected overflows, truncated text, etc.).

## 12. Next Steps (Performance, Context, Images)

### 12.1 Performance & Latency

- Reduce number of API calls by tuning batching:
  - Experiment with larger `max_batch_chars` defaults for OpenAI-backed translation while staying within model limits.
  - Optionally expose a CLI flag (e.g. `--max-batch-chars`) tuned to “fast but safe” presets.
- Add optional concurrency:
  - Spec an internal scheduler that can send multiple batches in parallel (subject to rate limits).
  - Future CLI flag: `--max-concurrent-requests` to control parallelism.
- De-duplicate repeated text:
  - Detect identical `source_text` across the deck (e.g. repeated bullet templates, footers).
  - Translate each unique string once, then reuse cached translations when writing back.
- Allow “fast mode” profiles:
  - Preset configurations that skip masters/notes or use a cheaper/faster model for bulk translation.

### 12.2 Context-Aware Translation & Terminology

- Deck-level context analysis:
  - First pass over all slides to build a “deck profile” (title, agenda, frequent terms, section headers).
  - Use this profile as part of the system prompt for all translation batches.
- Glossary and terminology enforcement:
  - Accept user-provided glossary files via CLI (e.g. `--glossary glossary.json` or CSV with `source,target` columns).
  - Pass glossary entries into the backend prompt with instructions like “must use these translations for the following terms”.
  - Optionally add a pre-check that flags glossary terms not found in the deck.
- Two-step context workflow (optional):
  - Step 1: Summarize the deck and extract candidate domain terms via LLM.
  - Step 2: Present or export a suggested glossary for user review.
  - Step 3: Run the final translation pass using the approved glossary and deck summary as context.
- Per-section context:
  - Group units by slide or section and include slide/section titles in the prompt so technical terms are translated consistently within a topic.

### 12.3 Translating Text in Images (OCR)

- Image text detection:
  - Identify picture shapes/images via OOXML (e.g. `<p:pic>` and related `blip` references).
  - Run an OCR backend (e.g. Tesseract, Azure Vision, or another API) on each image to extract text.
  - Represent each OCR result as a `TranslatableUnit` with a location referencing the image ID/file.
- Translation and reinsertion strategies:
  - Strategy A (non-destructive overlay):
    - Create new text boxes over or near the original image to display the translated text.
    - Preserve the original image and slide layout; keep overlay positioning as close as possible.
  - Strategy B (image replacement, later):
    - Render a new image with translated text baked in (requires image rendering pipeline, e.g. Pillow or external service).
    - Replace the original image in the PPTX while keeping geometry (size/position) unchanged.
- Controls and scope:
  - Add CLI toggle such as `--translate-images` (default off due to performance and API cost).
  - Allow selecting OCR backend and configuration via `--image-ocr-backend` and `--image-ocr-config`.
  - Document that image text translation can significantly increase processing time compared to plain text-only translation.
