# tiwater-docx

A .NET 9 globally installed command-line tool for inspecting, comparing, and transforming Word (`.docx`) documents.

## Installation

Install the tool from the NuGet global registry using the .NET CLI:

```bash
dotnet tool install -g tiwater.docx.cli
```

## Usage

The CLI provides several commands for document processing, structural inspection, and templating. Appending `--json` to querying commands outputs the data in a machine-readable JSON structure.

### 1. Inspect a Document
Outputs a unified structural report of a Word document, including paragraph styles, headings, placeholders, comments, annotation anchors, table previews, fields, drawings, and formatting metrics.

```bash
tiwater-docx inspect <input.docx> [--json]
```

### 1a. Inspect Table Details
Exports body table rows, cells, grid spans, vertical merges, paragraph alignment, run font, color, underline, and text-fill details. Use this for template-fidelity validation where row/cell merge structure and run-level formatting matter.

```bash
tiwater-docx inspect-tables <input.docx> [--json]
```

### 2. Compare Two Documents
Compares a baseline and an updated document. Reports on differences in package structure, overall metrics, and paragraph style usage changes.

```bash
tiwater-docx compare <old.docx> <new.docx> [--json]
```

### 3. Validate Template Transformation
Validates compatibility between a source template and a target template. Ensures that body field slots match and reports any structural discrepancies.

```bash
tiwater-docx validate-template-transform <source-template.docx> <target-template.docx> [--json]
```

### 4. Strip Direct Formatting
Removes direct formatting from paragraphs and runs. Useful for enforcing strict style adherence instead of manual styling.

```bash
tiwater-docx strip-direct-formatting <input.docx> <output.docx>
```

### 5. Replace Style IDs
Replaces internal Style IDs within a document based on a provided JSON mapping structure.

```bash
tiwater-docx replace-style-ids <input.docx> <output.docx> <style-map.json>
```

### 6. Export Body JSON
Exports body paragraphs and tables as structured JSON, including `paragraphIndex` on paragraph nodes and `tableIndex` on table nodes.

```bash
tiwater-docx export-json <input.docx> [<output.json>]
```

### 7. Fill Placeholder Template
Fills a classic placeholder-based template using JSON data.

```bash
tiwater-docx fill-template <template.docx> <data.json> <output.docx>
```

### 8. Normalize OpenXML
Canonicalizes known WordprocessingML namespace prefixes and orders common run/table property children so Word does not need to repair generated packages.

```bash
tiwater-docx normalize-openxml <input.docx> <output.docx>
```

### 9. Apply Explicit Edit Operations
Applies a batch of explicit edits to a DOCX. Supported operation types are:
- `replaceAnchoredText`
- `replaceParagraphText`
- `replaceBodyText`
- `replaceTableCellText`
- `replaceTableCellRichText`
- `replaceTable`
- `insertTableRows`
- `replaceTableRows`
- `setTableWidth`
- `setTableCellAlignment`
- `deleteComment`
- `deleteComments`
- `sanitizeFields`
- `freezeFields`
- `markFieldsDirty`

`replaceTableCellText` accepts optional `alignment` (`left`, `center`, `right`, `both`).
`replaceTableCellRichText` accepts `richText` segments with `text`, optional `color`, `underline`, and `bold`.
`replaceTable` row cell objects may use the same `richText` segments instead of plain `text`.
`insertTableRows` inserts `rows` before `rowIndex`; `templateRowIndex` controls which existing row supplies row/cell/run styling.
`replaceTableRows` replaces inclusive `startRowIndex`..`endRowIndex` with `rows`, preserving the surrounding table and using `templateRowIndex` for row/cell/run styling.
`setTableWidth` accepts `width` and `widthType` (`pct`, `dxa`, `auto`, `nil`).
`sanitizeFields` removes update-field prompts and dirty field markers from the package.
`freezeFields` converts visible field results into ordinary content so converters cannot recalculate cross-references or sequence numbers.

```bash
tiwater-docx edit <input.docx> <operations.json> <output.docx>
```

Example operations file:

```json
{
  "operations": [
    { "type": "replaceAnchoredText", "commentId": "12", "text": "Final narrative" },
    { "type": "replaceBodyText", "findText": "HSPXXX", "text": "HSP-PTMs" },
    { "type": "replaceTableCellText", "tableIndex": 2, "rowIndex": 0, "cellIndex": 3, "text": "2026-04-15" },
    {
      "type": "replaceTableCellRichText",
      "tableIndex": 2,
      "rowIndex": 1,
      "cellIndex": 2,
      "richText": [
        { "text": "QV" },
        { "text": "Q", "color": "FF0000", "underline": true },
        { "text": "LVQSGAEVK" }
      ]
    },
    {
      "type": "replaceTable",
      "tableIndex": 0,
      "rows": [
        [
          { "text": "检测项目", "bold": true },
          { "text": "时间点", "gridSpan": 2, "bold": true }
        ],
        [
          { "text": "颜色" },
          {
            "richText": [
              { "text": "1" },
              { "text": "月", "color": "FF0000", "underline": true }
            ]
          },
          { "text": "3月" }
        ]
      ]
    },
    { "type": "deleteComment", "commentId": "12" },
    { "type": "setTableWidth", "tableIndex": 0, "width": "5000", "widthType": "pct" },
    { "type": "setTableCellAlignment", "tableIndex": 1, "rowIndex": 2, "cellIndex": 3, "alignment": "center" },
    { "type": "sanitizeFields" },
    { "type": "freezeFields" },
    { "type": "markFieldsDirty" }
  ]
}
```

Scenario-specific planning and resolution workflows now live in Lucid skills and scripts. This CLI remains the generic document runtime for inspection, export, fill, comparison, and explicit edit application.
