# PowerPoint Translator (layout-preserving)

Tool to translate PowerPoint `.pptx` files while preserving slide layout, formatting, positioning, and alignment as closely as possible.

## Goals

- Input: one or more `.pptx` files.
- Output: translated `.pptx` files with all shapes, images, animations, and layouts preserved.
- Preserve: positions, sizes, fonts, colors, bullet levels, text boxes, tables, charts, and notes.
- Pluggable translation backend (e.g. DeepL, Azure, OpenAI, custom).

## High-level Design

- CLI command (e.g. `pptx-translate`) that:
  - Reads a source `.pptx`.
  - Unzips the file and walks the underlying Open XML parts.
  - Extracts only actual text (`<a:t>` nodes in slides, notes, masters, layouts).
  - Batches text for translation via a backend adapter interface.
  - Writes translated text back into the XML in-place without changing any shape geometry or formatting attributes.
  - Re-zips the package into a new `.pptx`.

See `SPEC.md` for detailed requirements and architecture.

## Quickstart

```bash
# Install locally (editable)
pip install -e .

# Translate with the dummy backend (prefixes text with target lang code)
pptx-translate path/to/input.pptx --target-lang de -o path/to/output.pptx

# Dry run: print extracted + translated pairs, no file write
pptx-translate path/to/input.pptx --target-lang fr --dry-run
```

Key flags:
- `--source-lang`: optional if backend auto-detects.
- `--include-notes/--no-include-notes`: toggle notes translation (default on).
- `--include-masters/--no-include-masters`: toggle master/layout translation (default on).
- `--backend`: translation backend id (`dummy`, `openai`).
- `--backend-config`: backend-specific config (JSON).

### Step-by-step: run with OpenAI

1. Install Python 3.10+.
2. (Recommended) create and activate a virtual env:
   ```bash
   python -m venv .venv
   # Windows PowerShell
   .\.venv\Scripts\Activate.ps1
   ```
3. Install the tool (and dev extras if you want tests):
   ```bash
   pip install -e .
   # or with tests/deps
   pip install -e .[dev]
   ```
4. Provide your OpenAI API key (choose one):
   - One-time for the current shell (PowerShell):
     ```powershell
     $env:OPENAI_API_KEY="sk-..."
     ```
   - Or put it into a config file (JSON) so you don’t set env vars:
     ```json
     {
       "api_key": "sk-...",
       "model": "gpt-4o-mini",
       "temperature": 0.0
     }
     ```
     Save as `openai.json`.
5. Run the translator:
   ```bash
   # Using env var, default output: input.<lang>.pptx
   pptx-translate input.pptx --target-lang es --backend openai

   # Or using the config file
   pptx-translate input.pptx --target-lang es --backend openai --backend-config openai.json
   ```
6. Optional toggles:
   - Skip notes: `--no-include-notes`
   - Skip masters/layouts: `--no-include-masters`
   - Dry run (see translations, no file write): `--dry-run`

### Dev / tests

```bash
pip install -e .[dev]
pytest
```
