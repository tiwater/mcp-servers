# tiwater/mcp-servers

Shared MCP server implementations and shared runtime packages for Tiwater agent skills.

## Layout

- `packages/docx-cli` — shared DOCX runtime
- `packages/xlsx-cli` — shared XLSX runtime
- `packages/pdf-cli` — shared PDF runtime
- `servers/office` — shared Office MCP server workspace
- `servers/pdf` — shared PDF MCP server workspace

## Boundary

This repository owns shared executable capabilities. Agent Skills live in `tiwater/skills`. Domain workflows live in downstream repos such as `tiwater/lucid`.
