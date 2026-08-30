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
tiwater-docx docx_list_objects request.json
tiwater-docx docx_read_object request.json
tiwater-docx export-json input.docx output.json
tiwater-docx compare old.docx new.docx --json
```

`docx_find_literal` remains a CLI diagnostic for a human who already knows the
exact document and scope. It is not published to Agents because text occurrence
does not establish document identity or source selection.

Inspection reports current package, story, paragraph, run, table, row, cell,
field, drawing, font, flow, and formatting facts. Complete JSON can be written
to an artifact path by the Office MCP.

## Native object mutation

```bash
tiwater-docx docx_copy_content request.json
tiwater-docx docx_set_text request.json
tiwater-docx docx_replace_table_rows request.json
tiwater-docx docx_copy_object request.json
tiwater-docx docx_delete_object request.json
tiwater-docx docx_merge_cells request.json
tiwater-docx docx_split_cells request.json
tiwater-docx docx_apply_font_policy request.json
tiwater-docx docx_apply_toc_style_policy request.json
tiwater-docx normalize-openxml input.docx output.docx
tiwater-docx strip-direct-formatting input.docx output.docx
tiwater-docx replace-style-ids input.docx output.docx style-map.json
```

Each `docx_*` mutation command consumes the matching provider-owned request
contract from `contracts/mcp-input/`. The provider resolves the supplied OpenXML
part and native path and enforces only executable OpenXML constraints; callers
own the selected objects and values. After structural mutation, re-list the
changed parent when its child paths are needed again.

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
