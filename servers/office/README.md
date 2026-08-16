# office MCP server

Shared stdio MCP server for Office document workflows.

## Tools

- `docx_inspect`
- `docx_list_migration_choices`
- `docx_migrate_template`
- `docx_verify_migration`
- `docx_compare`
- `docx_validate_template_transform`
- `docx_export_json`
- `xlsx_inspect`
- `xlsx_export_json`
- `xlsx_validate`
- `pptx_inspect`
- `pptx_export_json`

## Run

Install `@tiwater/office-mcp` together with the runtime versions required by
the consumer, then run `tiwater-office-mcp` as a stdio MCP server.

The server invokes published `tiwater-docx`, `tiwater-xlsx`, and
`tiwater-pptx` commands from `PATH`. It does not require a source checkout or
fall back to local projects.

The official MCP SDK derives the schemas advertised to clients and validates
tool arguments and structured results before they cross the protocol boundary.
Large observations and exports are written to a caller-selected new JSON
artifact. MCP returns only the artifact path, hash, and byte count.
