# tiwater-pdf

A Python command-line utility for inspecting PDF documents and extracting tabular data, heavily utilized in analytical reporting workflows (e.g., HPLC reports).

## Installation

This tool requires Python 3.11+. We recommend installing it using modern package managers like `uv` or `pipx` to avoid global environment conflicts:

```bash
# Recommend approach using uv:
uv tool install tiwater-pdf

# Or using pipx:
pipx install tiwater-pdf

# Fallback (may require --break-system-packages on newer OS):
pip install tiwater-pdf
```

## Commands Reference

The CLI provides four major functionalities:

### 1. Find a Specific Table
Searches the document for a table matching a specific heading or name and attempts to extract it.

```bash
tiwater-pdf find-table <report.pdf> "<table_name>" [--auto-span] [--json]
```
*   `--auto-span`: Enables heuristics to span tables that break across multiple pages.
*   `--json`: Outputs the table data entirely in machine-readable JSON format.

### 2. Extract All Tables
Extracts all tables detected within the PDF or from specific pages.

```bash
tiwater-pdf extract-tables <report.pdf> [--pages 1,3,4] [--auto-span] [--json]
```

### 3. Inspect PDF
Provides a high-level inspection of the PDF's structural layout and tables to determine its format.

```bash
tiwater-pdf inspect <report.pdf>
```

### 4. OCR Scanned PDFs With a Vision LLM
Extracts text from scanned or image-only PDFs using an OpenAI-compatible vision model.

```bash
tiwater-pdf ocr <scan.pdf> [--pages 1,2] [--json]
```

Configuration is read from explicit flags first, then environment variables:

- `--api-key`, `SUPEN_LLM_TOKEN`, `SUPEN_LLM_API_KEY`, `TIWATER_LLM_API_KEY`, `OPENAI_API_KEY`, or `OPENROUTER_API_KEY`
- `--base-url`, `SUPEN_LLM_GATEWAY_URL`, `SUPEN_LLM_BASE_URL`, `TIWATER_LLM_BASE_URL`, or `OPENAI_BASE_URL`
- `--llm-model`, `TIWATER_LLM_OCR_MODEL`, `TIWATER_LLM_VISION_MODEL`, or the built-in `gpt-4o-mini` OCR default

When only `OPENROUTER_API_KEY` is present, the default base URL is `https://openrouter.ai/api/v1`.
When running under Supen, `SUPEN_LLM_GATEWAY_URL` should point at the gateway's OpenAI-compatible route, for example `http://127.0.0.1:2755/api/llm/v1`.
