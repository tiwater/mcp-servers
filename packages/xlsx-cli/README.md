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
*   `--json` returns the complete technical workbook readback.
*   Every addressed cell exposes style identity, number-format id/code, font/fill/border ids, horizontal/vertical alignment, and wrap state. Rich text cells also expose per-run text and formatting.

Inspection accepts current `.xls` and `.xlsx` sources; editing remains `.xlsx` only.

`tiwater-xlsx inspect <input.xlsx> --json` includes the versioned `tiwater.xlsx.evidence/v1`
readback envelope. It includes typed raw cells, style identity and alignment,
number formats, formula/shared-formula metadata, merges, dimensions, sheet view,
print/page settings and workbook date system for independent baseline comparison.
Each physical cell also exposes its formatted display value, effective inherited
style identifiers, and normalized number-format evidence (source, normalized code,
semantic kind, and date classification). Numeric date/time cells include an ISO-8601
value computed from the workbook's declared 1900 or 1904 date system. These are
technical readback facts; the tool does not infer business semantics.

### 2. Inventory Non-empty Regions

Publishes a deterministic, versioned inventory of non-empty row bands. A completely
empty row terminates one region and the next non-empty row starts another. Every
material cell retains its address, row/column coordinates, raw value, formatted
display value, and formula. The output is bound to the input SHA-256 and does not
infer headers, record identities, methods, samples, or other business semantics.

```bash
tiwater-xlsx inventory-regions <input.xlsx> [<output.json>] [--schema v1|v2]
```

The default output contract remains `tiwater.xlsx.region-inventory/v1`; its JSON
Schema is packaged as `contracts/tiwater.xlsx-region-inventory-v1.schema.json`.
Callers that need typed cell-value evidence must opt in with `--schema v2`.
`tiwater.xlsx.region-inventory/v2` preserves every v1 cell field unchanged and
adds a required `normalizedValue` derived from the workbook's declared date
system and effective number format. The provider reports technical spreadsheet
facts only: it does not infer business date meaning, locale intent, headers,
record identities, methods, or samples. The v2 schema is packaged as
`contracts/tiwater.xlsx-region-inventory-v2.schema.json`.

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
- `setCellValue` with required `sheet`, `cell`, and `value`; optional `valueType`, `bold`, `shrinkToFit`, and `wrapText`
- `setCellNumberFormat` with required target `sheet` and existing target `cell`; provide exactly one of an explicit `numberFormat` code or a same-workbook `sourceSheet` / `sourceCell` peer whose observed number format should be copied. Only the target cell's number-format component changes; its value and all other style components are preserved.
- `setPrintArea` with required `sheet` and A1-style `range`
- `setPageSetup` with required `sheet` and at least one of `fitToPagesWide`, `fitToPagesTall`, `orientation` (`portrait` or `landscape`), `paperSize` (`letter`, `legal`, `a3`, or `a4`), paired `repeatRowsStart` / `repeatRowsEnd`, or paired `repeatColsStart` / `repeatColsEnd`; repeated rows and columns use one-based indices and are persisted together as the sheet-local standard Excel print-title definition
- `setRowPageBreaks` with required `sheet` and a strictly increasing `breakBeforeRows` list; replaces the sheet's manual horizontal page breaks so every listed row begins a new printed page
- `setColumnWidth` with required `sheet`, bounded A1-style `column`, and Excel-compatible `width`
- `setRichTextCellValue` with required `sheet`, `cell`, `value`, and `bold`; writes one explicit rich-text run so value and all-run bold state are one operation
- `setRangeValues` with required `sheet`, `startCell`, and `values`; optional `valueType`
- `insertRows` with required `sheet`, `startRow`, and `count`; optional `expandAdjacentVerticalMergedRanges` extends vertical merged ranges that end immediately before the insertion point
- `copyRow` with required `sheet`, `sourceRow`, and `targetRow`; optional `translateFormulas` and `preserveHorizontalMergedRanges`
- `expandSectionRows` with required `sheet`, `anchorText`, `exampleRows`, and `targetRows`; optional `preserveStyle`, `preserveFormulas`, and `preserveMergedRanges`

By default, edit operations use `valueType: "auto"` semantics. Numeric-looking
values are written as numeric Excel cells unless the target cell is formatted as
text; other values are written as strings. The target cell's existing style and
number format are preserved. Set `valueType` to `"text"` or `"number"` on an
operation when a caller needs explicit behavior. `valueType: "date"` accepts an
ISO date/time and writes its Excel serial while preserving the target number format.

Formula adjustment for `insertRows` and `copyRow` is intentionally conservative.
It supports A1-style cell references, including local references and sheet-qualified
references. Whole-row references, 3D references, structured table references, and
external workbook references are not guaranteed to be adjusted correctly.
When requested, `copyRow` duplicates only horizontal merged ranges wholly contained
in the source row. It rejects a target that intersects a different existing merge.
Vertical category merges are expanded by `insertRows` only when
`expandAdjacentVerticalMergedRanges` is explicitly true.
`expandSectionRows` finds the first visible text cell exactly matching
`anchorText`, treats the following `exampleRows` as the template section, inserts
rows until the section reaches `targetRows`, and copies example rows cyclically
into generated rows. Styles, translated formulas, and merged-range movement are
preserved by default. A print area that contains the example section, including
one ending at its last example row, expands with the generated rows. Shrinking
existing sections is reported as a warning and
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
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "F2", "value": "a value that must remain visible", "shrinkToFit": true },
    { "type": "setCellValue", "sheet": "Sheet1", "cell": "G2", "value": "a long value that may use multiple lines", "wrapText": true },
    { "type": "setCellNumberFormat", "sheet": "Sheet1", "cell": "H2", "numberFormat": "yyyy-mm-dd" },
    { "type": "setCellNumberFormat", "sheet": "Sheet1", "cell": "I2", "sourceSheet": "Sheet1", "sourceCell": "H2" },
    { "type": "setPrintArea", "sheet": "Sheet1", "range": "A1:G12" },
    { "type": "setPageSetup", "sheet": "Sheet1", "fitToPagesWide": 1, "orientation": "landscape", "paperSize": "a3", "repeatRowsStart": 1, "repeatRowsEnd": 2, "repeatColsStart": 1, "repeatColsEnd": 2 },
    { "type": "setRowPageBreaks", "sheet": "Sheet1", "breakBeforeRows": [27, 40] },
    { "type": "setColumnWidth", "sheet": "Sheet1", "column": "G", "width": 60 },
    { "type": "setRangeValues", "sheet": "Sheet1", "startCell": "F2", "values": [["233988", "383789"], ["252353", "341366"]], "valueType": "number" },
    { "type": "insertRows", "sheet": "RP", "startRow": 8, "count": 2 },
    { "type": "copyRow", "sheet": "RP", "sourceRow": 12, "targetRow": 14, "translateFormulas": true },
    { "type": "expandSectionRows", "sheet": "RP", "anchorText": "impurity peak area", "exampleRows": 2, "targetRows": 4, "preserveStyle": true, "preserveFormulas": true, "preserveMergedRanges": true }
  ]
}
```

### 5. Validate a Workbook Package
Validates an `.xlsx` workbook as an Open XML spreadsheet package and returns JSON validation evidence. The command exits `0` when the workbook is valid and `1` when validation errors are found or the file is not a valid XLSX package.

```bash
tiwater-xlsx validate <input.xlsx>
```

The CLI is a generic workbook runtime for inspection, export, template filling, explicit edit application, and package validation.
`export-json` also includes each cell's complete `style` evidence and `richTextRuns`, so downstream planners and independent validators can prove formats without parsing package XML directly.
