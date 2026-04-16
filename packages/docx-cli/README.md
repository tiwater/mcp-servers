# tiwater-docx

A .NET 9 globally installed command-line tool for inspecting, planning, comparing, and transforming Word (`.docx`) documents.

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

### 2. Plan ANA03 Comment-Driven Edits
Builds a reviewable plan from an annotated DOCX. Use this when comments describe intent and you want explicit proposed edits before mutating the document.

```bash
tiwater-docx plan <input.docx> <plan-data.json>
```

The planner is ANA03-focused and classifies each comment into one of:

- `replace_anchor_text`
- `fill_table_block`
- `generate_paragraph`
- `manual_only`

The output includes target scope, candidate targets, required sources, confidence, reasoning, and proposed `docx_edit` operations. The intended workflow is:

```text
inspect -> plan -> edit -> inspect
```

### 3. Compare Two Documents
Compares a baseline and an updated document. Reports on differences in package structure, overall metrics, and paragraph style usage changes.

```bash
tiwater-docx compare <old.docx> <new.docx> [--json]
```

### 4. Validate Template Transformation
Validates compatibility between a source template and a target template. Ensures that body field slots match and reports any structural discrepancies.

```bash
tiwater-docx validate-template-transform <source-template.docx> <target-template.docx> [--json]
```

### 5. Strip Direct Formatting
Removes direct formatting from paragraphs and runs. Useful for enforcing strict style adherence instead of manual styling.

```bash
tiwater-docx strip-direct-formatting <input.docx> <output.docx>
```

### 6. Replace Style IDs
Replaces internal Style IDs within a document based on a provided JSON mapping structure.

```bash
tiwater-docx replace-style-ids <input.docx> <output.docx> <style-map.json>
```

### 7. Export Body JSON
Exports body paragraphs and tables as structured JSON.

```bash
tiwater-docx export-json <input.docx> [<output.json>]
```

### 8. Fill Placeholder Template
Fills a classic placeholder-based template using JSON data.

```bash
tiwater-docx fill-template <template.docx> <data.json> <output.docx>
```

### 9. Apply Explicit Edit Operations
Applies a batch of explicit edits to a DOCX. Supported operation types are:
- `replaceAnchoredText`
- `replaceParagraphText`
- `replaceTableCellText`
- `deleteComment`
- `deleteComments`
- `markFieldsDirty`

```bash
tiwater-docx edit <input.docx> <operations.json> <output.docx>
```

Example operations file:

```json
{
  "operations": [
    { "type": "replaceAnchoredText", "commentId": "12", "text": "Final narrative" },
    { "type": "replaceTableCellText", "tableIndex": 2, "rowIndex": 0, "cellIndex": 3, "text": "2026-04-15" },
    { "type": "deleteComment", "commentId": "12" },
    { "type": "markFieldsDirty" }
  ]
}
```

### 10. Resolve ANA03 Comment-Driven Edits
Turns a `stability-report` plan plus exported source JSON files into explicit `docx_edit` operations for the first ANA03 slice.

```bash
tiwater-docx resolve <stability-report-plan.json> <resolve-data.json>
```

The resolve data file points at exported JSON files:

```json
{
  "scenario": "stability-report",
  "stabilityDataPath": "./stability-data.json",
  "qualityStandardCnPath": "./quality-standard-cn.json",
  "reportPath": "./report-export.json",
  "protocolPath": "./protocol-export.json"
}
```

The resolver currently covers:
- comment `0` partial-span replacement
- comment `18` narrative paragraph replacement
- table `9` rows for color, clarity, pH, and protein concentration

### Plan Input

The `plan` command accepts a small JSON file:

```json
{
  "scenario": "stability-report",
  "sourceHints": ["summary sheet", "supporting notes"]
}
```
