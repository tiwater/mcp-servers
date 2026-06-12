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
- `startSectionBeforeParagraph`
- `replaceAllHeaderParagraphText`
- `replaceHeaderParagraphText`
- `replaceHeaderText`
- `replaceTableCellText`
- `replaceTableCellRichText`
- `replaceTable`
- `insertTableRows`
- `deleteTableRows`
- `replaceTableRows`
- `insertTableColumns`
- `setTableWidth`
- `setTableCellAlignment`
- `setTableCellNoWrap`
- `setTableCellFontSize`
- `setTableRowHeight`
- `mergeTableCells`
- `unmergeTableRowHorizontalCells`
- `unmergeTableColumnVerticalCells`
- `deleteComment`
- `deleteComments`
- `sanitizeFields`
- `freezeFields`
- `markFieldsDirty`

`replaceTableCellText` accepts optional `alignment` (`left`, `center`, `right`, `both`).
`replaceHeaderText` accepts `findText` and `text`, replacing matching text inside headers without overwriting other header content.
`replaceHeaderParagraphText` accepts `headerIndex`, `paragraphIndex`, and `text`.
`replaceAllHeaderParagraphText` accepts `paragraphIndex` and `text`, replacing that paragraph in every header part where it exists.
`startSectionBeforeParagraph` accepts `findText` and `orientation` (`landscape` or `portrait`); it inserts a section break before the matching direct body paragraph and applies the requested orientation to the following section.
`replaceTableCellRichText` accepts `richText` segments with `text`, optional `color`, `underline`, `bold`, and `fontName`.
When the target cell is empty, the generated runs inherit font-related formatting from the nearest table run so blank template cells do not fall back to Office default font size; emphasis such as bold/italic is not inherited from fallback runs.
`replaceTable` row cell objects may use the same `richText` segments instead of plain `text`.
`insertTableRows` inserts `rows` before `rowIndex`; `templateRowIndex` controls which existing row supplies row/cell/run styling.
`deleteTableRows` deletes inclusive `startRowIndex`..`endRowIndex`, preserving the surrounding table.
`replaceTableRows` replaces inclusive `startRowIndex`..`endRowIndex` with `rows`, preserving the surrounding table and using `templateRowIndex` for row/cell/run styling. When the replaced range contains multiple row shapes, replacement rows are matched to a template row with the same `gridSpan` pattern when possible, so mixed merged/unmerged rows keep their cell widths and paragraph properties.
`insertTableColumns` inserts empty columns before a visual grid `columnIndex`; `columnCount` defaults to `1`, and `templateColumnIndex` controls which existing grid column/cell supplies width and cell styling. If the insertion point falls inside an existing horizontally merged cell, that cell's `gridSpan` is expanded instead of creating a new physical cell in that row.
`setTableWidth` accepts `width` and `widthType` (`pct`, `dxa`, `auto`, `nil`) and preserves the template table layout (`fixed`, `autofit`, or absent) instead of changing it.
`setTableCellNoWrap` accepts optional `noWrap`; `true` or omitted writes Word `w:noWrap`, and `false` removes it.
`setTableCellFontSize` accepts `fontSize` as OpenXML half-points (`18`) or points (`9pt`).
`setTableRowHeight` accepts `height` in twips and optional `heightRule` (`atLeast`, `exact`, `auto`).
`mergeTableCells` merges a horizontal cell range when `rowIndex/startCellIndex/endCellIndex` are provided, or a vertical row range when `cellIndex/startRowIndex/endRowIndex` are provided.
`unmergeTableRowHorizontalCells` splits one horizontally merged visible cell in `tableIndex/rowIndex/cellIndex` back into its grid columns, preserving the original text in the first cell and inserting empty styled cells for the remaining columns.
`unmergeTableColumnVerticalCells` removes vertical merge markers in `tableIndex/cellIndex/startRowIndex/endRowIndex` and fills continuation cells from the latest visible content.
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
    { "type": "replaceHeaderText", "findText": "XX（客户项目代号）（与报告中HSPTEST对应）", "text": "HSPTEST" },
    { "type": "replaceTableCellText", "tableIndex": 2, "rowIndex": 0, "cellIndex": 3, "text": "2026-04-15" },
    {
      "type": "replaceTableCellRichText",
      "tableIndex": 2,
      "rowIndex": 1,
      "cellIndex": 2,
      "richText": [
        { "text": "QV" },
        { "text": "Q", "color": "FF0000", "underline": true, "fontName": "Times New Roman" },
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
    { "type": "insertTableColumns", "tableIndex": 0, "columnIndex": 6, "columnCount": 2, "templateColumnIndex": 5 },
    { "type": "deleteComment", "commentId": "12" },
    { "type": "setTableWidth", "tableIndex": 0, "width": "5000", "widthType": "pct" },
    { "type": "setTableCellAlignment", "tableIndex": 1, "rowIndex": 2, "cellIndex": 3, "alignment": "center" },
    { "type": "setTableCellNoWrap", "tableIndex": 1, "rowIndex": 2, "cellIndex": 3 },
    { "type": "setTableCellFontSize", "tableIndex": 1, "rowIndex": 2, "cellIndex": 3, "fontSize": "9pt" },
    { "type": "setTableRowHeight", "tableIndex": 1, "rowIndex": 2, "height": "240", "heightRule": "exact" },
    { "type": "sanitizeFields" },
    { "type": "freezeFields" },
    { "type": "markFieldsDirty" }
  ]
}
```

Scenario-specific planning and resolution workflows now live in Lucid skills and scripts. This CLI remains the generic document runtime for inspection, export, fill, comparison, and explicit edit application.
