# office MCP server

Published generic Office capabilities for DOCX, XLS/XLSX, and PPTX.

The server owns technical observation, native conversion, fixed-action edits,
package validation, and native WPS rendering. It owns no scenario meaning,
template-migration workflow, customer mapping, or Lucid lifecycle.

## Capability families

- DOCX: inspect document/tables, export, compare, validate OpenXML and font
  policy, fill placeholders, transform styles, and batch one fixed edit action.
- XLS/XLSX: convert legacy XLS with ET, inspect/export/fill, validate, and batch
  one fixed workbook edit action.
- PPTX: inspect/export/fill, bind selected masters/layouts, apply text formatting,
  set exact top-level object geometry, replace existing picture media, and
  validate OpenXML.
- Office: render DOC/DOCX/XLS/XLSX/PPT/PPTX to PDF with the corresponding native
  WPS backend.

Each mutation tool fixes its provider operation type. Callers submit only the
coordinates and values for that action, so they cannot provide an arbitrary
operation discriminator or a multi-action plan language. A call may batch
multiple changes only when every change has the same action kind.

Structural worksheet row deletion is exposed as `xlsx_delete_rows`. Each change
contains only `sheet`, `startRow`, and `count`; unsupported dependent workbook
structures fail atomically and are reported by the provider receipt.

The bounded DOCX object actions are:

- `docx_insert_body_range`: copies inclusive direct-body indexes from a bound
  source DOCX before a target direct-body boundary. Source documents are hashed
  into the receipt. Whole sections preserve supported style, numbering, media,
  hyperlink, header, and footer relationships; partial or unsafe sections fail.
- `docx_replace_drawing_image`: replaces one body drawing's embedded image while
  preserving its drawing geometry. Image inputs are hashed into the receipt.
- `docx_insert_body_image`: inserts one inline body drawing with explicit EMU
  dimensions. Image inputs are hashed into the receipt.
- `docx_set_table_row_repeat_as_header`: sets or unsets native repeat-as-header
  state on uniquely addressed direct body, header, or footer table rows. The
  whole same-action batch is validated before mutation; nested, missing,
  ambiguous, duplicate, or invalid targets fail closed.

`docx_inspect_tables` preserves the v1 body `Tables` view and additively exposes
header/footer topology in `StoryTables`. Header/footer tables carry part
coordinates plus section/reference bindings, and only supported direct-story
tables expose mutation addresses.

The catalog is intentionally open to new generic document capabilities, but a
scenario, template, customer, issue, work item, or model difference does not
justify a new tool. Add a tool only for a stable technical responsibility that
cannot be composed from the existing public capabilities; merge or remove
overlapping capabilities.

## Run

Install `@tiwater/office-mcp` with the exact published CLI versions required by
the consumer, then run:

```bash
tiwater-office-mcp
```

The server uses published `tiwater-docx`, `tiwater-xlsx`,
`tiwater-pptx`, and `tiwater-convert` commands from `PATH`.
