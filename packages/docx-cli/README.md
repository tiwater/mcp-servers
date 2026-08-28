# tiwater-docx

`tiwater-docx` provides technical DOCX observation, fixed Open XML mutation,
normalization, comparison, and package validation. It does not own business
mappings, template migration, workflow decisions, or delivery status.

The Agent-facing surface is the published Office MCP. The CLI exposes the same
provider behavior for diagnostics and package integration; it is not a second
workflow protocol.

## Observation

```bash
tiwater-docx inspect input.docx --json
tiwater-docx inspect-tables input.docx --json
tiwater-docx export-json input.docx output.json
tiwater-docx compare old.docx new.docx --json
```

Inspection reports current package, story, paragraph, run, table, row, cell,
field, drawing, font, flow, and formatting facts. Complete JSON can be written
to an artifact path by the Office MCP.

## Fixed technical mutation

```bash
tiwater-docx edit input.docx operations.json output.docx
tiwater-docx normalize-openxml input.docx output.docx
tiwater-docx strip-direct-formatting input.docx output.docx
tiwater-docx replace-style-ids input.docx output.docx style-map.json
```

`edit` accepts only the currently published fixed technical actions. The
provider validates document coordinates and Open XML constraints; callers own
the selected objects and values. Structural mutation requires fresh
observation before later calls.

## Validation

```bash
tiwater-docx validate-openxml input.docx
tiwater-docx validate-font-policy input.docx policy.json
tiwater-docx validate-toc-style-policy input.docx false 2
```

These commands prove technical package or requested property conditions only.
They do not determine whether document content is correct for a business task.

## Discovery

```bash
tiwater-docx --list-tools
tiwater-docx --describe-tool
tiwater-docx <command> --help
```

The provider tool list contains technical commands only. The Office MCP adapter
must expose the same provider-owned operations without adding business fields.
