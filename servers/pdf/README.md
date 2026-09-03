# Tiwater PDF MCP

Published, Agent-facing stdio MCP server for PDF inspection, table extraction,
and OCR. It is a separate package and tool surface from `@tiwater/office-mcp`.

## Tools

- `pdf_inspect`
- `pdf_extract_tables`
- `pdf_find_table`
- `pdf_extract_table_details`
- `pdf_ocr` (pinned Aliyun `qwen3.8-max`; per-invocation Supen credential)

Every tool is read-only and records the exact input PDF identity. Read results
can be retained in immutable JSON artifacts. `pdf_inspect` always requires that
durable artifact and separately returns a bounded document identity without
traversing page content.

## Install and run

```bash
npm install --global @tiwater/pdf-mcp
tiwater-pdf-mcp
```

The server invokes the independently published `tiwater-pdf` executable from
`PATH`. It does not fall back to a repository checkout. OCR credentials are
provided to that child for one invocation by the runtime environment; the MCP
does not read or persist provider credentials.
