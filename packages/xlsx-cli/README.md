# tiwater-xlsx

`tiwater-xlsx` provides technical XLSX observation, fixed Open XML mutation,
region inventory, export, and package validation. It does not own business
mappings, workflow decisions, or delivery status.

The Agent-facing surface is the published Office MCP. The CLI exposes the same
provider requests for diagnostics and package integration; it is not a second
workflow protocol.

## Observation

```bash
tiwater-xlsx inspect input.xlsx --json
tiwater-xlsx export-json input.xlsx output.json
tiwater-xlsx inventory-regions input.xlsx output.json --schema v2
```

Observation reports current sheets, cells, values, formulas, styles, merges,
dimensions, and print settings. It does not infer headers, business fields, or
record identities. Inspection accepts current `.xls` and `.xlsx` sources;
mutation accepts `.xlsx` only.

## Fixed technical mutation

```bash
tiwater-xlsx xlsx_set_cell_value request.json
tiwater-xlsx xlsx_set_range_values request.json
tiwater-xlsx xlsx_insert_rows request.json
tiwater-xlsx xlsx_set_page_setup request.json
```

Each `xlsx_*` mutation command consumes the matching provider-owned request
contract from `contracts/mcp-input/`. Requests contain no operation
discriminator. One call batches only the named action. The provider preserves
unselected workbook content and fails without publishing an output when a
requested structural change is technically unsafe.

## Validation

```bash
tiwater-xlsx validate input.xlsx
```

Validation proves package integrity and technical postconditions only. It does
not decide whether workbook content is correct for a business task.

## Discovery

```bash
tiwater-xlsx --list-tools
tiwater-xlsx <command> --help
```

The provider tool list contains technical commands only. The Office MCP adapter
must expose the same provider-owned requests without adding business fields.
