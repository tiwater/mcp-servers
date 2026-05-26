# tiwater-xlsx

A .NET 9 globally installed command-line tool for inspecting, editing, and filling `.xlsx` workbooks.

## Installation

Install the tool from the NuGet global registry using the modern .NET CLI:

```bash
dotnet tool install -g tiwater.xlsx.cli
```

## Placeholder Syntax

In your target Excel template (`.xlsx`):
*   **Single Cells**: Should be formatted exactly as `{{placeholder_key}}` (e.g., `{{controlledNumber}}`). The entire cell's content must just be the placeholder text if it's meant to be replaced entirely.
*   **Data Grids/Tables**: Should be anchored with `{{table:placeholder_key}}`. The CLI will auto-fill a 2D array downwards and to the right starting directly from that anchored cell.

## Usage

### 1. Inspect a Workbook
Outputs sheet-level metrics, placeholders, used ranges, formula counts, and merged regions. This is the canonical low-level read surface for both placeholder templates and fixed-layout workbooks.

```bash
tiwater-xlsx inspect <template.xlsx> [--json]
```
*   `--json` returns structured output suitable for parsers.

### 2. Export Workbook JSON
Exports workbook content for template analysis. Each sheet includes row arrays for compatibility plus address-level cell records so callers can locate fixed-layout sections without relying on hardcoded row and column positions. Cell records include `reference`, `row`, `column`, `value`, `formattedValue`, and `formula` when present.

```bash
tiwater-xlsx export-json <input.xlsx> [<output.json>]
```

### 3. Fill a Template
Injects the defined JSON payload directly into an active Excel sheet, replacing matched placeholders and rendering the final result document.

```bash
tiwater-xlsx fill-template <template.xlsx> <data.json> <output.xlsx>
```

#### Expected JSON Model

The structured shape of `<data.json>` expected by `fill-template` must look like the following:

```json
{
  "cellValues": {
    "controlledNumber": "260359",
    "calculationResult": "0.98",
    "placeholder_name": "example_value"
  },
  "tableData": {
    "peakAreas": [
      ["Peak1", "Area1", "RT1"],
      ["Peak2", "Area2", "RT2"]
    ]
  }
}
```


### 4. Apply Explicit Edit Operations
Applies a batch of explicit fixed-layout workbook edits. Supported operation types are:
- `setCellValue` (optional `bold: true|false`)
- `setRangeValues`
- `insertRows`

`insertRows` copies row styles and formulas from `sourceRow`, inserts `count` rows before `targetRow`, shifts following rows downward, and translates relative formula references into the inserted rows. Non-formula cell values are blank by default so callers can fill the inserted rows with `setRangeValues`; set `copyValues: true` to duplicate values too.

By default, edit operations use `valueType: "auto"` semantics. Numeric-looking
values are written as numeric Excel cells unless the target cell is formatted as
text; other values are written as strings. The target cell's existing style and
number format are preserved. Set `valueType` to `"text"` or `"number"` on an
operation when a caller needs explicit behavior.

```bash
tiwater-xlsx edit <input.xlsx> <operations.json> <output.xlsx>
```

Example operations file:

```json
{
  "operations": [
    { "type": "insertRows", "sheet": "Sheet1", "sourceRow": 2, "targetRow": 3, "count": 2 },
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "D2", "value": "260359-01" },
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "E7", "value": "浅于黄色0.5号标准比色液", "bold": false },
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "E2", "value": "10.2" },
    { "type": "setRangeValues", "sheet": "Sheet1", "startCell": "F2", "values": [["233988", "383789"], ["252353", "341366"]], "valueType": "number" }
  ]
}
```

Scenario-specific fixed-layout planning workflows now live in Lucid skills and scripts. This CLI remains the generic workbook runtime for inspection, export, template filling, and explicit edit application.
