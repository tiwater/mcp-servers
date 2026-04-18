# tiwater/mcp-servers

Shared MCP server implementations and shared runtime packages for Tiwater agent skills.

## Layout

- `packages/docx-cli` — shared DOCX runtime
- `packages/xlsx-cli` — shared XLSX runtime
- `packages/pptx-cli` — shared PPTX runtime
- `packages/pdf-cli` — shared PDF runtime
- `servers/office` — shared Office MCP server for DOCX, XLSX, and PPTX operations
- `servers/pdf` — shared PDF MCP server for inspection and table extraction

## Run locally

```bash
node servers/office/index.mjs
node servers/pdf/index.mjs
```

The servers prefer published commands when available:

- `tiwater-docx`
- `tiwater-xlsx`
- `tiwater-pptx`
- `tiwater-pdf`

When those are not installed, they fall back to the local runtime sources in `packages/`.

## Boundary

This repository owns shared executable capabilities. Agent Skills live in `tiwater/skills`. Domain workflows live in downstream repos such as `tiwater/lucid`.
