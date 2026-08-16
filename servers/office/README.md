# office MCP server

Shared stdio MCP server for Office document workflows.

## Tools

- `docx_inspect`
- `docx_list_migration_choices`
- `docx_query_migration_choices`
- `docx_migrate_template`
- `docx_verify_migration`
- `docx_compare`
- `docx_export_json`
- `office_render_pdf`
- `xlsx_inspect`
- `xlsx_export_json`
- `xlsx_validate`
- `pptx_inspect`
- `pptx_export_json`

## Run

Install `@tiwater/office-mcp` together with the runtime versions required by
the consumer, then run `tiwater-office-mcp` as a stdio MCP server.

The server invokes published `tiwater-docx`, `tiwater-xlsx`,
`tiwater-pptx`, and `tiwater-convert` commands from `PATH`. It does not require
a source checkout or fall back to local projects.

The official MCP SDK derives the schemas advertised to clients and validates
tool arguments and structured results before they cross the protocol boundary.
Large observations and exports are written to a caller-selected new JSON
artifact. MCP returns only the artifact path, hash, and byte count.
Template-migration choice artifacts are opaque evidence. List the choices once,
then query the same current source and baseline to page sources, request targets
for a business action, or inspect cleanup targets. Query results expose short
catalog-bound references. Submit one complete batch to migrate, then verify the
output independently from the same inputs and batch. The tool does not choose
the business mapping.
