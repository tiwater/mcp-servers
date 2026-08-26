# office MCP server

Shared stdio MCP server for Office document workflows.

## Tools

The public surface is organized by stable Office object actions. Mutation tools
accept a homogeneous `changes` batch; callers never select an operation `type`.
Scenario, customer, template, issue, work-item, and model differences do not add
tools. A new tool requires a new generic Office action that cannot be composed
from the existing actions.

DOCX observation and general transforms:

- `docx_inspect`, `docx_inspect_tables`, `docx_compare`
- `docx_validate_template_transform`, `docx_validate`, `docx_validate_font_policy`
- `docx_export_json`, `docx_fill_template`
- `docx_strip_direct_formatting`, `docx_replace_style_ids`

DOCX content and document-part actions:

- `docx_set_anchored_text`, `docx_set_paragraph_text`, `docx_set_paragraph_run_text`
- `docx_replace_body_text`, `docx_delete_body_paragraph`, `docx_delete_body_range`
- `docx_start_section`, `docx_collapse_trailing_empty_section`
- `docx_set_header_paragraph_text`, `docx_set_header_run_text`, `docx_replace_header_text`
- `docx_set_footer_paragraph_text`, `docx_set_footer_run_text`
- `docx_set_header_table_cell_text`, `docx_set_header_table_cell_run_text`
- `docx_set_footer_table_cell_text`, `docx_set_footer_table_cell_run_text`
- `docx_delete_comments`
- `docx_mark_fields_dirty`, `docx_sanitize_fields`, `docx_freeze_fields`
- `docx_apply_font_policy`

DOCX table actions:

- `docx_set_table_cell_text`, `docx_set_table_cell_run_text`, `docx_set_table_cell_rich_text`
- `docx_insert_table_rows`, `docx_delete_table_rows`, `docx_replace_table_rows`, `docx_insert_table_columns`
- `docx_set_table_width`, `docx_set_table_cell_alignment`, `docx_set_table_cell_no_wrap`, `docx_set_table_cell_font_size`
- `docx_set_table_row_height`, `docx_set_table_row_cant_split`, `docx_set_table_row_keep_next`
- `docx_merge_table_cells`, `docx_unmerge_table_row_cells`, `docx_unmerge_table_column_cells`

XLSX actions:

- `xlsx_inspect`, `xlsx_export_json`, `xlsx_fill_template`, `xlsx_validate`
- `xlsx_set_cell_value`, `xlsx_set_range_values`
- `xlsx_insert_rows`, `xlsx_copy_row`, `xlsx_expand_section_rows`
- `xlsx_set_print_area`, `xlsx_set_page_setup`, `xlsx_set_column_width`

PPTX actions:

- `pptx_inspect`, `pptx_export_json`, `pptx_fill_template`
- `pptx_set_text_format`, `pptx_apply_template`, `pptx_validate`

## Run

```bash
node servers/office/index.mjs
```

The server prefers published `tiwater-docx`, `tiwater-xlsx`, and `tiwater-pptx` commands.
It falls back to `dotnet run --project ...` for docx/xlsx/pptx.
