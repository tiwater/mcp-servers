# tiwater-docx

`tiwater-docx` provides technical DOCX observation, native Open XML mutation,
normalization, comparison, and package validation. It does not own business
mappings, template migration, workflow decisions, or delivery status.

The Agent-facing surface is the published Office MCP. The CLI exposes the same
provider behavior for diagnostics and package integration; it is not a second
workflow protocol.

## Observation

```bash
tiwater-docx inspect input.docx --json
tiwater-docx inspect-tables input.docx --json
tiwater-docx docx_list_objects request.json
tiwater-docx docx_find_literal request.json
tiwater-docx docx_read_object request.json
tiwater-docx export-json input.docx output.json
tiwater-docx compare old.docx new.docx --json
```

Inspection reports current package, story, paragraph, run, table, row, cell,
field, drawing, font, flow, and formatting facts. Complete JSON can be written
to an artifact path by the Office MCP.

## Native object mutation

```bash
tiwater-docx docx_copy_content request.json
tiwater-docx docx_copy_object request.json
tiwater-docx docx_delete_object request.json
tiwater-docx docx_merge_cells request.json
tiwater-docx docx_split_cells request.json
tiwater-docx normalize-openxml input.docx output.docx
tiwater-docx strip-direct-formatting input.docx output.docx
tiwater-docx replace-style-ids input.docx output.docx style-map.json
```

Each `docx_*` mutation command consumes the matching provider-owned request
contract from `contracts/mcp-input/`. The provider validates revision-bound
native object references and Open XML constraints; callers own the selected
objects and values. Structural mutation requires fresh observation before later
calls.

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
