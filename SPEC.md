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
- `--glossary`: path to glossary file (JSON/CSV with `source,target`).
- `--context` / `--context-file`: free-text domain/context hint for translation.
- `--dedupe-text` / `--no-dedupe-text`: control deduplication of identical source strings before translation.

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
  - Supports glossary enforcement and free-text context to bias terminology via prompt.

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

## 12. Advanced Features (Versioning, Performance, Context, Images, QA)

### 12.1 Run ID & Versioning

Goals:
- Avoid overwriting previous translated decks.
- Make each translation run traceable via a unique run identifier, both in the filename and inside the PPTX metadata.

Run identifier:
- Each invocation of the translator has a `run_id` string.
- If the user does not pass `--run-id`, the tool auto-generates one using UTC timestamp plus a short random suffix, e.g. `20241126-213045-3f7a`.
- The user can override with `--run-id <string>` (must be filesystem-safe: alphanumeric plus `-` / `_`).
- `--no-run-id` disables run-id in the output filename (but a run id is still generated for internal metadata).

Output filename semantics:
- Default output (no `--output`):
  - Current default is `<input>.<target>.pptx`. With run id, it becomes `<stem>.<target>.<run_id>.pptx`, e.g. `Deck.fr.20241126-213045-3f7a.pptx`.
- Explicit `--output` provided:
  - If `--run-id` or auto-run-id is in effect, the tool appends `.<run_id>` before `.pptx` unless `--no-run-id` is set, e.g. `output.fr.pptx` -> `output.fr.20241126-213045-3f7a.pptx`.
  - If `--no-run-id` is set, the tool writes exactly to `--output` (may overwrite if the file exists).
- Overwrite behavior:
  - By default, if the final output filename already exists and `--no-run-id` is set, the tool fails with a clear error message.
  - A future `--overwrite` flag may allow forced overwrite; without it, run-id is the primary mechanism to keep versions.

Embedded run metadata:
- The tool writes run metadata into `docProps/custom.xml` as custom properties, for example:
  - `pptx_translate_run_id`
  - `pptx_translate_source_lang`
  - `pptx_translate_target_lang`
  - `pptx_translate_backend`
  - `pptx_translate_model`
  - `pptx_translate_timestamp_utc`
  - `pptx_translate_profile` (fast/balanced/quality)
- The QA report and logs reference the same `run_id` so users can match files, logs, and reports.

CLI changes (summary):
- `--run-id <string>`: explicitly set run id used in filenames and metadata.
- `--no-run-id`: do not modify the output filename (but still generate run id internally).

### 12.2 Performance & Latency

Batching strategy:
- Reduce number of API calls by tuning batching:
  - Use `--max-batch-chars` as a soft upper bound and infer a token-aware limit per model (e.g., percentage of context window).
  - Auto-split large batches on backend errors (e.g. context-length or rate-limit) and retry with smaller batch sizes.

Concurrency:
- Introduce an internal scheduler to send multiple translation batches in parallel, subject to backend rate limits.
- CLI flag `--max-concurrent-requests` (default: 1) controls the maximum number of in-flight translation requests per run.
- Guarantees:
  - Translation units are grouped into batches, and batches are processed concurrently up to the limit.
  - Order of translations is preserved when reconstructing the final list of `TranslatableUnit`s.
  - Concurrency is only used for network-based backends; dummy/local backends may ignore it or run synchronously.

Fast mode profiles:
- Add `--profile` flag with values like:
  - `fast`: larger batches, higher `--max-concurrent-requests`, dedupe enabled, cheaper/faster model (e.g., `gpt-4o-mini`), and by default `--no-include-masters` and `--no-include-notes` unless overridden.
  - `balanced` (default): current behavior with moderate batch sizes, dedupe on, masters/notes included, standard model.
  - `quality`: smaller batches, possibly slower but higher-quality model (e.g., full `gpt-4o`), dedupe still on, masters/notes included, and no aggressive concurrency.
- Profiles set reasonable defaults for:
  - `max_batch_chars`
  - `max_concurrent_requests`
  - `backend`/`model` (where appropriate)
  - `include_notes` / `include_masters`
  - `dedupe_text`

Smarter auto-tuning of `max_batch_chars`:
- The translator keeps track of:
  - Observed error responses indicating too-large payloads.
  - Average response times per batch.
- Behavior:
  - On context-length or 413-like errors, reduce effective batch size for subsequent requests by a configurable factor (e.g., 50%). Persist for the rest of the run.
  - On consistently fast responses with small batches, optionally increase batch size up to the soft upper bound to reduce total calls.
  - These adjustments are logged, including the final effective batch size chosen.

### 12.3 Context-Aware Translation & Terminology

Deck-level context analysis:
- First pass over all slides to build a "deck profile" before translation:
  - Extract: deck title, agenda/TOC slide(s), section headers, repeated phrases, top keywords, and any existing glossary slides.
  - Optionally call an LLM once with a subset of representative slide text to summarize:
    - Domain (e.g., cybersecurity, finance, medical).
    - Target audience.
    - Register (formal/informal).
  - Store this as a structured `DeckProfile` object (in memory) and as an optional JSON file if requested via CLI.
- The `DeckProfile` is converted into a compact textual context string and appended to the system or user prompt for all translation batches.

User-provided context (already available):
- CLI flags `--context` / `--context-file` remain and are merged with the auto-generated `DeckProfile`:
  - Final context = user context + derived deck profile summary.
  - User context has precedence in case of conflicting instructions.

Glossary and terminology enforcement (two-step workflow):
- Step 1: Glossary suggestion run:
  - New CLI mode `--generate-glossary <output.csv>` (or `.json`), which:
    - Analyzes the deck texts (titles, headers, repeated terms).
    - Calls an LLM to propose candidate term pairs (source/target) with optional notes.
    - Writes a file with columns such as `source`, `preferred_target`, `notes`, `frequency`.
  - In this mode, no translation is performed; the tool exits after generating the glossary.
- Step 2: User review:
  - User edits the generated glossary to keep/remove/edit entries.
- Step 3: Translation run with enforced glossary:
  - User supplies the approved glossary with `--glossary` (CSV/JSON).
  - Backend prompt includes explicit instructions to respect these mappings.

Per-section context:
- The tool groups slides into sections as defined in the PPTX (if available) or infers sections based on title slides.
- For each batch of units belonging to the same slide/section:
  - The section title and slide title are included in the batch-specific context (e.g., "Section: Incident Response; Slide: Containment Step 1").
  - This per-section context is merged with the global `DeckProfile` and user context for that batch.

Modes of operation:
- Default (single-pass): build `DeckProfile` automatically and use it immediately; no user interaction required beyond optional `--glossary`.
- Two-step workflow:
  - First run with `--generate-glossary` (and optionally `--write-deck-profile profile.json`).
  - Second run with `--glossary` and optional `--context-file profile.json` for fully controlled terminology.

### 12.4 Translating Text in Images (OCR)

Overview:
- Enable translation of text that appears inside images (screenshots, diagrams, scanned documents) by running OCR and overlaying translated text inside the slide.

OCR backends:
- Define an `OcrBackend` interface similar to `TranslationBackend`, with at least:
  - `recognize(self, images: list[ImageInput], config: dict | None) -> list[OcrResult]`
  - `ImageInput` contains: image bytes, source filename/path, slide index, and shape identifier.
  - `OcrResult` contains: recognized text, confidence, bounding boxes (relative to image coordinates), language hints.
- Initial OCR backend implementations (design-level):
  - `tesseract`: local OCR via Tesseract/pytesseract (requires local installation).
  - `azure-vision`: remote OCR using Azure Cognitive Services.
  - Future: an OpenAI vision-based backend.

Image detection and OCR pipeline:
- Identify picture shapes in each slide via OOXML (e.g. `<p:pic>` with `a:blip` references to `ppt/media/image*.png`).
- For each image:
  - Extract the image binary and its placement/size (bounding box on the slide).
  - Send images to the configured `OcrBackend` (optionally batched) to get recognized text + bounding boxes.
  - For each recognized text region, create a `TranslatableUnit` with:
    - `location`: slide index + image id + region index.
    - `source_text`: OCR text.
    - `context`: marked as `"image_text"` plus nearby slide text/section info.

Translating and reinserting image text:
- Translation:
  - Image-derived `TranslatableUnit`s flow through the same translation pipeline as other units (including glossary and context).
- Reinsertion strategy (initial implementation: non-destructive overlay):
  - For each `OcrResult` region, create a new text box (`<p:sp>` shape) on the slide:
    - Positioned over (or near) the original bounding box scaled to slide coordinates.
    - With a semi-transparent background or contrasting color to ensure readability.
    - With font size scaled roughly to match the size of the original text region.
  - Keep the original image unchanged behind the overlay.
- Future strategy (image replacement):
  - Render a new image with translated text baked in (using Pillow or an external rendering service).
  - Replace the original `ppt/media/image*.png` while preserving the shape’s geometry on the slide.

Controls and CLI flags:
- `--translate-images` (bool, default: off):
  - When on, image OCR and translation pipeline are enabled.
  - When off, images are left untouched.
- `--image-ocr-backend`: select OCR backend (`tesseract`, `azure-vision`, etc.).
- `--image-ocr-config`: JSON/YAML config for OCR backend (API keys, endpoints, language hints).
- `--image-overlay-style` (future): preset styles for overlays (background color, opacity, font family).

Performance considerations:
- Image OCR can significantly increase processing time and API cost.
- When `--translate-images` is enabled, the tool:
  - Logs the number of images processed and OCR timings.
  - Optionally supports batching images and limiting max concurrent OCR requests (aligned with `--max-concurrent-requests` or a separate OCR concurrency limit).

### 12.5 Post-Translation QA Report

Goals:
- Help users quickly identify slides where translated text may be truncated, overflowing, or visually problematic.
- Provide a machine-readable report and an optional human-readable summary.

Scope of checks:
- Text overflow/truncation heuristics:
  - Compare length of translated text vs source text (e.g., flag if length ratio exceeds a configurable threshold such as 1.6×).
  - Analyze text frames versus their bounding boxes (width/height) using approximate character-per-line/lines-per-box heuristics.
  - Use PPTX layout properties (e.g., auto-fit vs fixed size) where available to infer likely overflow.
- Layout consistency:
  - Detect if bullet indentation levels changed in an unexpected way (e.g., significantly longer bullets causing wrap to additional lines).
  - Optionally detect if a slide has a much higher word count than others in the same section (potential readability issue).
- Image text overlays:
  - When `--translate-images` is enabled, check if overlay text boxes extend beyond image bounds by more than a small tolerance.

Report contents:
- QA report is associated with a `run_id` and includes:
  - `run_id`, source/target language, backend, model, timestamp.
  - Per-slide entries with:
    - Slide index and title (if any).
    - For each flagged shape/text box:
      - Shape identifier (slide-local index or OOXML id).
      - Original text and translated text (or truncated preview).
      - Reason(s) for flag (e.g., "length_ratio>1.8", "estimated_overflow", "overlay_outside_image").
- Optional summary section:
  - Number of slides scanned, number of slides with issues.
  - Top categories of issues.

Output formats and CLI:
- `--qa-report <path>`:
  - When provided, the tool writes a machine-readable QA report (JSON by default) to the given path.
  - If not provided, QA still runs but only severe issues may be logged to stdout/stderr.
- `--qa-report-format`:
  - `json` (default): structured data for tools/automation.
  - `markdown`: a readable summary suitable for manual review.
- `--qa-threshold-length-ratio` and `--qa-threshold-overflow-score` (future):
  - Allow users to tune how aggressive the overflow/truncation detector is.
