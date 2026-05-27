# tiwater-xlsx

A .NET 9 globally installed command-line tool for inspecting, editing, validating, and filling `.xlsx` workbooks.

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
- `setCellValue` with required `sheet`, `cell`, and `value`; optional `valueType` and `bold`
- `setRangeValues` with required `sheet`, `startCell`, and `values`; optional `valueType`
- `insertRows` with required `sheet`, `startRow`, and `count`
- `copyRow` with required `sheet`, `sourceRow`, and `targetRow`; optional `translateFormulas`
- `expandSectionRows` with required `sheet`, `anchorText`, `exampleRows`, and `targetRows`; optional `preserveStyle`, `preserveFormulas`, and `preserveMergedRanges`

By default, edit operations use `valueType: "auto"` semantics. Numeric-looking
values are written as numeric Excel cells unless the target cell is formatted as
text; other values are written as strings. The target cell's existing style and
number format are preserved. Set `valueType` to `"text"` or `"number"` on an
operation when a caller needs explicit behavior.

Formula adjustment for `insertRows` and `copyRow` is intentionally conservative.
It supports A1-style cell references, including local references and sheet-qualified
references. Whole-row references, 3D references, structured table references, and
external workbook references are not guaranteed to be adjusted correctly.
`expandSectionRows` finds the first visible text cell exactly matching
`anchorText`, treats the following `exampleRows` as the template section, inserts
rows until the section reaches `targetRows`, and copies example rows cyclically
into generated rows. Styles, translated formulas, and merged-range movement are
preserved by default. Shrinking existing sections is reported as a warning and
does not delete rows.

```bash
tiwater-xlsx edit <input.xlsx> <operations.json> <output.xlsx>
```

Example operations file:

```json
{
  "operations": [
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "D2", "value": "260359-01" },
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "E7", "value": "浅于黄色0.5号标准比色液", "bold": false },
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "E2", "value": "10.2" },
    { "type": "setRangeValues", "sheet": "Sheet1", "startCell": "F2", "values": [["233988", "383789"], ["252353", "341366"]], "valueType": "number" },
    { "type": "insertRows", "sheet": "RP", "startRow": 8, "count": 2 },
    { "type": "copyRow", "sheet": "RP", "sourceRow": 12, "targetRow": 14, "translateFormulas": true },
    { "type": "expandSectionRows", "sheet": "RP", "anchorText": "impurity peak area", "exampleRows": 2, "targetRows": 4, "preserveStyle": true, "preserveFormulas": true, "preserveMergedRanges": true }
  ]
}
```

### 4. Validate a Workbook Package
Validates an `.xlsx` workbook as an Open XML spreadsheet package and returns JSON validation evidence. The command exits `0` when the workbook is valid and `1` when validation errors are found or the file is not a valid XLSX package.

```bash
tiwater-xlsx validate <input.xlsx>
```

Scenario-specific fixed-layout planning workflows now live in Lucid skills and scripts. This CLI remains the generic workbook runtime for inspection, export, template filling, explicit edit application, and package validation.
