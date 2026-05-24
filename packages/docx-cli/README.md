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

### 8. Apply Explicit Edit Operations
Applies a batch of explicit edits to a DOCX. Supported operation types are:
- `replaceAnchoredText`
- `replaceParagraphText`
- `replaceTableCellText`
- `replaceTable`
- `setTableWidth`
- `setTableCellAlignment`
- `deleteComment`
- `deleteComments`
- `sanitizeFields`
- `markFieldsDirty`

`replaceTableCellText` accepts optional `alignment` (`left`, `center`, `right`, `both`).
`setTableWidth` accepts `width` and `widthType` (`pct`, `dxa`, `auto`, `nil`).
`sanitizeFields` removes update-field prompts and dirty field markers from the package.

```bash
tiwater-docx edit <input.docx> <operations.json> <output.docx>
```

Example operations file:

```json
{
  "operations": [
    { "type": "replaceAnchoredText", "commentId": "12", "text": "Final narrative" },
    { "type": "replaceTableCellText", "tableIndex": 2, "rowIndex": 0, "cellIndex": 3, "text": "2026-04-15" },
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
          { "text": "1月" },
          { "text": "3月" }
        ]
      ]
    },
    { "type": "deleteComment", "commentId": "12" },
    { "type": "setTableWidth", "tableIndex": 0, "width": "5000", "widthType": "pct" },
    { "type": "setTableCellAlignment", "tableIndex": 1, "rowIndex": 2, "cellIndex": 3, "alignment": "center" },
    { "type": "sanitizeFields" },
    { "type": "markFieldsDirty" }
  ]
}
```

Scenario-specific planning and resolution workflows now live in Lucid skills and scripts. This CLI remains the generic document runtime for inspection, export, fill, comparison, and explicit edit application.
