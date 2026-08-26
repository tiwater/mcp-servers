# pdf MCP server

Shared stdio MCP server for deterministic PDF inspection/table extraction and fixed-model OCR.

## Tools

- `pdf_inspect`
- `pdf_extract_tables`
- `pdf_find_table`
- `pdf_extract_table_details`
- `pdf_ocr` (fixed `qwen3.8-max`; per-invocation Supen Gateway credential only)

## Run

```bash
node servers/pdf/index.mjs
```

The server prefers the published `tiwater-pdf` command and falls back to `python3 -m tiwater_pdf.cli` from this repo.
