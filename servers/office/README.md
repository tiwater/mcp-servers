# office MCP server

Published generic Office capabilities for DOCX, XLS/XLSX, and PPTX.

The server owns technical observation, native conversion, fixed-action edits,
package validation, and native WPS rendering. Callers own document selections,
values, business mappings, workflow decisions, and delivery status.

Every inspect/export result binds its observation artifact to the exact source
file path, SHA-256, and byte count used by that invocation.
Full native object reads are also written as artifacts instead of being inlined.

Every filesystem argument is identified in the published MCP input schema by
`x-tiwater-file-role`. `read` marks an existing provider input; `write` marks a
new provider artifact. Strings without this metadata are document values or
OpenXML-internal identifiers, not filesystem arguments.

## Capability families

- DOCX: inspect document/tables, export, compare, validate OpenXML, apply or validate font and
  table-of-contents style policies, refresh document fields, replace object text, transform styles,
  and batch one fixed edit action.
- XLS/XLSX: convert legacy XLS with ET, inspect/export, validate, and batch
  one fixed workbook edit action.
- PPTX: inspect/export, bind selected masters/layouts, apply text formatting,
  set exact top-level object geometry, replace existing picture media, and
  validate OpenXML.
- Office: render DOC/DOCX/XLS/XLSX/PPT/PPTX to PDF with the corresponding native
  WPS backend.

Each mutation tool fixes its provider operation type. Callers submit only the
coordinates and values for that action, so they cannot provide an arbitrary
operation discriminator or a multi-action plan language. A call may batch
multiple changes only when every change has the same action kind.
When a completed mutation receipt reports `summary.pass=false`, the MCP result
also sets the standard `isError` field; consumers do not need a private failure
interpretation.

Structural worksheet row deletion is exposed as `xlsx_delete_rows`. Each change
contains only `sheet`, `startRow`, and `count`; unsupported dependent workbook
structures fail atomically and are reported by the provider receipt.

DOCX observation and mutation use the OpenXML part URI and native object path
directly. Content edits keep unaffected addresses usable. After a structural
edit, callers re-list only the changed parent when they need its new children.

Table-row copying is a distinct bulk table responsibility: the caller selects
current source and target tables, bounded row regions, excluded source rows,
and source/target columns. The provider expands the target region, preserves
target presentation, carries selected source paragraph content, and reproduces
the selected source row order and merge topology. It does not infer business
identity, language, row scope, or column meaning.

The catalog is intentionally open to new generic document capabilities, but a
specific workflow, template, customer, issue, input document, or model
difference does not justify a new tool. Add a tool only for a stable technical
responsibility that cannot be composed from the existing public capabilities;
merge or remove overlapping capabilities.

## Run

Install `@tiwater/office-mcp` with the exact published CLI versions required by
the consumer, then run:

```bash
tiwater-office-mcp
```

The server uses published `tiwater-docx`, `tiwater-xlsx`,
`tiwater-pptx`, and `tiwater-convert` commands from `PATH`.
