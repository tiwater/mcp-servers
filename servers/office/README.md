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

## Template migration

Template migration separates business choice from document mechanics:

1. `docx_list_migration_choices` records the complete current source and target
   catalog in an opaque run-local artifact.
2. `docx_query_migration_choices` pages source items and returns bounded,
   document-compatible alternatives for one source item.
3. `docx_migrate_template` accepts one complete batch and derives the plan,
   edits, and readback receipt.
4. `docx_verify_migration` independently verifies the output from the same
   source, baseline, and batch.

The scenario supplies the business meaning. The query tool exposes three
orthogonal target actions:

- `place-content` moves current content into a target content position.
- `keep-template-label` keeps the target label and structure while migrating a
  uniquely identified current field value.
- `select-template-option` marks a target option represented by the current
  source fact.

Choose the action first, then query targets with that action filter. Returned
`alternativeRef` values bind the action and target together. Source exclusion
and genuine local review are target-free terminal choices. The caller never
supplies document text, selectors, coordinates, plans, or edit operations.

Template-migration choice artifacts are opaque evidence. List the choices once,
then query the same current source and baseline to page sources, request targets
for a source, or inspect cleanup targets. Target queries return complete
provider-compatible action-and-target alternatives under short catalog-bound
references. Submit selected alternatives and target-free terminal choices as
one batch, then verify the output independently from the same inputs and batch.
The tool does not choose the business mapping.

## Version 0.10 migration

Version 0.10 replaces the 0.9 template-migration identity form. Targeted
choices now use one `alternativeRef` returned by
`docx_query_migration_choices`; terminal choices use `sourceRef` plus
`exclude-source` or `review-source`. The server rejects the old combination of
raw source id, action, and raw target id so an action cannot be paired with a
target from a different alternative. Other Office tools keep their existing
inputs and outputs.
