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

### 3. Extract Table Details
Extracts detected PDF tables with visual evidence for format validation: cell bounding boxes, null detected cells, text spans with font/size/color, and table-region line/rectangle drawings. PDF table cells are visual detections, so `null` cells may indicate merged or suppressed visual cells but are not authoritative DOCX-style rowspan/colspan.

```bash
tiwater-pdf extract-table-details <report.pdf> [--pages 1,3,4] [--json]
```

### 4. Inspect PDF
Provides a high-level inspection of the PDF's structural layout and tables to determine its format.

```bash
tiwater-pdf inspect <report.pdf>
```

### 5. OCR Scanned PDFs With a Vision LLM
Extracts text from scanned or image-only PDFs using an OpenAI-compatible vision model.

```bash
tiwater-pdf ocr <scan.pdf> [--pages 1,2] [--json]
```

Multiple PDFs can be OCRed concurrently with bounded parallelism. Batch mode
writes one JSON result and one status file per input plus a `manifest.json`:

```bash
tiwater-pdf ocr scan-1.pdf scan-2.pdf scan-3.pdf \
  --provider llm \
  --output-dir outputs/ocr-llm \
  --max-parallel 3 \
  --max-page-parallel 12 \
  --json
```

JSON OCR output preserves each model-returned markdown table and also exposes
deterministic `table_rows` evidence. Every row has a stable
`page-<n>-table-<n>-row-<n>` id, page/table/row coordinates, header flag, and
normalized `cells[]`; blank continuation cells remain blank. `table_rows` is
available both on each page and as a flattened top-level array so downstream
inventory validation can prove row coverage without reparsing markdown or
matching known business text.
Interior Markdown columns that contain no evidence in any row are removed
during normalization. This prevents a vision model from changing downstream
cell indexes by splitting one visual merged cell into multiple empty columns;
leading and trailing blank columns, and any column containing evidence in at
least one row, remain intact.
Top-level `table_logical_rows` additionally joins an unambiguous suffix-only
row at the start of the next page to its owning row at the previous page end.
It retains every contributing physical `source_row_ids` value, so downstream
field binding can use complete cross-page cells without guessing ownership.
The same output exposes `table_cell_lines[]`: every non-empty normalized line
inside every cell has a stable id derived from its row/cell/line coordinates.
This keeps source evidence stable when an OCR model represents repeated visual
rows either as separate markdown rows or as `<br>`-separated values in one
cell.

Configuration is read from explicit flags first, then environment variables:

- `--api-key`, `SUPEN_LLM_TOKEN`, `SUPEN_LLM_API_KEY`, `TIWATER_LLM_API_KEY`, `OPENAI_API_KEY`, or `OPENROUTER_API_KEY`
- `--base-url`, `SUPEN_LLM_GATEWAY_URL`, `SUPEN_LLM_BASE_URL`, `TIWATER_LLM_BASE_URL`, or `OPENAI_BASE_URL`
- `--llm-model`, `TIWATER_LLM_OCR_MODEL`, `TIWATER_LLM_VISION_MODEL`, or the built-in `qwen3.7-plus` OCR default
- `--max-tokens` or `TIWATER_PDF_OCR_MAX_TOKENS` to cap per-page OCR output
- `--max-parallel` or `TIWATER_PDF_OCR_MAX_PARALLEL` to cap concurrent PDFs in batch mode
- `--max-page-parallel` or `TIWATER_PDF_OCR_MAX_PAGE_PARALLEL` to cap concurrent pages within each PDF; page results are sorted before cross-page table normalization
- `--enable-thinking auto|true|false` or `TIWATER_LLM_ENABLE_THINKING` for vendor thinking mode

When only `OPENROUTER_API_KEY` is present, the default base URL is `https://openrouter.ai/api/v1`.
When running under Supen, `SUPEN_LLM_GATEWAY_URL` should point at the gateway's OpenAI-compatible route, for example `http://127.0.0.1:2755/api/llm/v1`.
In `auto` mode, bare Alibaba Model Studio Qwen3.5/Qwen3.6/Qwen3.7 model ids such as `qwen3.7-plus` disable thinking for OCR calls, which avoids unnecessary latency on extraction tasks. Provider-prefixed model ids such as `qwen/qwen3.7-plus` are left unchanged.
Each vision page request retries a bounded three times for transient gateway
timeouts, throttling, server errors, and the gateway's intermittent invalid-URL
response; successful page JSON records `request_attempts` for auditability.
