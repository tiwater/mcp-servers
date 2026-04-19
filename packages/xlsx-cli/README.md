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
Outputs sheet-level metrics, placeholders, used ranges, formula counts, merged regions, and note rows. This is the canonical low-level read surface for both placeholder templates and fixed-layout workbooks.

```bash
tiwater-xlsx inspect <template.xlsx> [--json]
```
*   `--json` returns structured output suitable for parsers.

### 2. Fill a Template
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


### 3. Apply Explicit Edit Operations
Applies a batch of explicit fixed-layout workbook edits. Supported operation types are:
- `setCellValue`
- `setRangeValues`

```bash
tiwater-xlsx edit <input.xlsx> <operations.json> <output.xlsx>
```

Example operations file:

```json
{
  "operations": [
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "D2", "value": "260359-01" },
    { "type": "setRangeValues", "sheet": "Sheet1", "startCell": "E2", "values": [["233988", "383789"], ["252353", "341366"]] }
  ]
}
```

Scenario-specific fixed-layout planning workflows now live in Lucid skills and scripts. This CLI remains the generic workbook runtime for inspection, export, template filling, and explicit edit application.
