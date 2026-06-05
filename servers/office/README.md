# office MCP server

Shared stdio MCP server for Office document workflows.

## Tools

- `docx_inspect`
- `docx_compare`
- `docx_validate_template_transform`
- `docx_strip_direct_formatting`
- `docx_replace_style_ids`
- `docx_export_json`
- `docx_fill_template`
- `docx_edit`
- `xlsx_inspect`
- `xlsx_export_json`
- `xlsx_fill_template`
- `xlsx_edit`
- `xlsx_validate`
- `pptx_inspect`
- `pptx_inspect_detail`
- `pptx_export_json`
- `pptx_fill_template`
- `pptx_apply_format_edits`

## Run

```bash
node servers/office/index.mjs
```

The server prefers published `tiwater-docx`, `tiwater-xlsx`, and `tiwater-pptx` commands.
It falls back to `dotnet run --project ...` for docx/xlsx/pptx.
