# office MCP server

This workspace is reserved for the shared Office MCP server surface.

The underlying shared runtimes live in:

- `../../packages/docx-cli`
- `../../packages/xlsx-cli`

Lucid and other downstream skills should depend on the Office MCP capability rather than embedding DOCX/XLSX executables locally.
