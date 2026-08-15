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

`--json` includes the document report, table details, document flow, and font inspection in one result.

### 1a. Inspect Table Details
Exports a versioned `tiwater.docx.inspect-tables/v1` envelope with tool version and extraction-view identity. Tables are traversed depth-first from the body through arbitrarily nested table cells; every table carries a containment path, declared `Width`/`WidthType`, and nested tables carry their parent-cell runtime address. Cell `Text` and `Paragraphs` contain only direct cell paragraphs and exclude nested-table descendants. Rows expose normalized grid omissions/extents, and cells expose mutation address, grid range/span, vertical merge, paragraph alignment, run font, color, underline, and text-fill details.

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

### 3a. Analyze Cross-Template Migration

Exports hash-attested source and baseline object inventories plus unresolved
structural/content/style differences. Table-cell objects include canonical
`Topology` (`ContainerObjectId`, `Row`, `Column`). It does not infer a business
mapping or modify either document. A content object also includes `Selector`
when the existing semantic-selector fields identify it uniquely within that
same current inventory. The selector contains no object id or coordinate and
can be copied into a semantic candidate after the caller chooses the business
mapping. Objects that cannot be identified uniquely have no `Selector`.
When a mapped table cell contains multiple source paragraphs, the deterministic
copy operation retains their visible boundaries as line breaks instead of
joining them. Independent readback observes the source and output paragraphs
again and rejects flattened or otherwise changed visible line structure.

```bash
tiwater-docx analyze-template-migration <source.docx> <baseline.docx> [--json]
```

### 3aa. Derive an Exact-Text Mapping Candidate

Produces a plan for content that is unique within both source and baseline for
the same object kind. Repeated table-cell content is also mapped when the full
normalized cell topology identifies exactly one source table and exactly one
baseline table, with a one-to-one row/column correspondence. The provider uses
its own current inventories for that proof; callers do not supply object ids or
coordinates. Media content is mapped when its content hash is unique in both
inventories; drawings that reference mapped media are covered by that mapping.
Missing or repeated media hashes remain in `Unresolved`. Repeated content
remains in `Unresolved` when either table is ambiguous or the semantic
topology differs. Other absent or ambiguous content also remains there. The
candidate `Plan` marks these objects as `unresolved`; it never represents them
as the terminal `review-required` disposition.

```bash
tiwater-docx derive-template-migration-exact-text-plan <source.docx> <baseline.docx>
```

This is a candidate producer: `Pass` means the typed candidate was derived
successfully. A non-empty `Unresolved` list is valid current evidence for
semantic resolution, not an upstream failure and not an executable migration
plan. Passing that incomplete plan to the operation builder fails with
`template-migration-semantic-resolution-required`; it does not create a
customer-review terminal.

`template-migration-exact-text-match-missing` means only that this mechanical
comparison found no identical text in the selected baseline.
`template-migration-exact-text-match-non-unique` means only that identical text
did not identify one reciprocal pair. Neither reason decides whether a semantic
target exists or whether its business meaning is ambiguous.

Each unresolved entry repeats the compact current source observation and any
exact baseline options. The same receipt also lists unclaimed non-empty
baseline paragraphs and table cells plus their uniquely selectable child runs.
These observations contain no object ids or coordinates and make the plan
itself sufficient for semantic selection;
the full technical analysis remains an audit artifact rather than an Agent
join surface.

### 3ab. Derive an Anchor-Gap Mapping Candidate

Produces a deterministic candidate for unmatched current objects when their
nearest mapped objects before and after them identify the same unique gap in
the selected baseline. It preserves current object order and leaves any
missing, reversed, or non-unique gap unresolved for a semantic candidate.

```bash
tiwater-docx derive-template-migration-anchor-gap-plan <source.docx> <baseline.docx>
```

This command uses the same candidate status contract: unresolved gaps remain
in `Unresolved` while the successfully derived candidate remains `Pass` true.
When a reciprocal anchor gap identifies one baseline object mechanically, its
compact current observation is included without approving the mapping.

### 3b. Build Cross-Template Operations

Compiles a hash-bound, validated semantic mapping into deterministic edit
operations. It rejects missing/duplicate source content, duplicate targets,
hash drift, type mismatches, unsupported targets, and unresolved mappings.
The builder does not infer any source-to-target mapping. A mapping can target
an attested `run` as well as a paragraph or table cell; run operations preserve
the target template's surrounding labels and formatting while replacing only
the mapped run's text. Object ids are accepted only from the current
hash-attested inventories, never as caller-supplied document coordinates.
For a mixed label/value parent, a semantic candidate may use `retain-target`
for the attested target parent only together with at least one mapped child
run. Readback verifies that every untargeted target run remains unchanged.
For an explicitly selected label-only parent, `retain-target-label` records
that target label retention without inferring semantic equivalence or accepting
coordinates; it emits no edit and readback verifies every target run unchanged.

A versioned semantic candidate may project a typed value between uniquely
selected current paragraph or table-cell parents even when labels, run splits,
or parent kinds differ. The candidate declares only semantic identity, value
kind (`text`, `token`, `date`, `identifier`, or `version`), and an extraction
contract (`after-first-delimiter`, `unique-delimited-run-group`, or
`unique-delimited-value`), or the explicit `whole-parent` contract for a parent
that consists only of the typed value. Unicode identifiers need no delimiter;
date values must be real calendar dates rather than regex-shaped strings.
Resolution binds current source/baseline hashes and
object ids; the operation builder derives the value and affected target runs
from those inventories. Empty, ambiguous, duplicate, or ill-typed values fail
closed. Independent readback derives the expected value and target replacements
again and verifies that sibling fields and target formatting remain unchanged.

A semantic candidate may also declare one or more source body ranges to append.
Each range is bounded by unique current paragraph/table selectors (a table may
be selected by a unique descendant header). The runtime resolves those selectors
to the current inventories, copies only plain body paragraph/table blocks after
the baseline body, and independently verifies both copied blocks and every
pre-existing baseline object. It rejects duplicate selectors, overlapping
ranges, drawings/revisions/content controls, missing styles, and any structural
drift outside the declared append. This is explicit source preservation when a
target template has no compatible section; it does not infer a target location,
style conversion, or semantic equivalence.

For source-only paragraph ranges that belong between two adjacent target body
anchors, semantic candidate v3 may declare an anchored body insertion. Both
source range endpoints and both target anchors must resolve uniquely in the
current hash-attested inventories, and the target anchors must be adjacent and
ordered. The supported `target-after-context` style policy keeps the source
paragraph/run content and order while applying the following target paragraph's
paragraph style. Tables, drawings, revisions, content controls, ambiguous
anchors, overlapping ranges, and non-adjacent anchors fail closed. Independent
readback rebuilds the source/baseline/output inventories, translates shifted
target object identities, and verifies inserted content/order, contextual style,
and the relative structure of every pre-existing target body object.

Semantic candidate v4 can also bind a current, non-empty source membership
object to a unique target label run whose immediately preceding run is a
drawing-backed choice glyph. The builder emits only a `selected` state change:
it replaces that glyph's image relationship with the provider-owned checked
symbol while preserving the drawing, label text, paragraph, and table shape.
Duplicate members or labels, missing/ambiguous labels, non-drawing glyphs, and
unknown selections fail closed. Independent readback recomputes the selected
label set from image hashes, verifies every bound label is unchanged, and
rejects both missing and additional selections.

Semantic candidate v5 may select baseline-owned placeholder or default content
for clearing by a unique current baseline selector. The selector is resolved to
the hash-attested baseline inventory; callers cannot supply object ids or table
coordinates. `cell` clears one paragraph/table cell and `row` clears every cell
in the selected table row. Missing, ambiguous, duplicate, unsupported, or
copy-conflicting selections fail closed. Composite clear/projection/insertion/
choice plans use plan v7, and independent readback verifies both the declared
empty targets and every untouched baseline run's text and formatting hashes.

Semantic candidate v6 can select an inventory object whose normalized `Text`
is empty with `textState: "empty"`. The predicate is mutually exclusive with
`text`, `sha256`, and `descendantText`, requires an explicit scope plus at least one current
semantic context (`parentText`, `previousText`, or `nextText`), and still must
resolve to exactly one object in the fresh hash-attested inventory. It does not
accept object ids, table coordinates, or an unbound empty string. Each consumer
continues to enforce its existing mapping, range, projection, choice, or clear
semantics after selection.

For repeated legacy container labels with no business content, an
`out-of-scope` mapping may explicitly use `cardinality: all`; this terminates
every current hash-bound semantic match. `all` is forbidden for copy, retain,
projection, insertion, and choice behavior, which continue to require unique
selectors.

#### Resolve a semantic candidate

```bash
tiwater-docx resolve-template-migration-semantic-candidate <source.docx> <baseline.docx> <candidate.json>
```

The candidate is a closed JSON object. Version 5 requires `schema` and a
`mappings` array (which may be empty when another branch supplies content).
Its optional top-level branches are `bodyAppends`, `valueProjections`,
`bodyInsertions`, `choiceSelections`, and `baselineClears`.

Every selector requires `kind` and exactly one primary predicate: `text`,
`sha256`, or `descendantText`. It may further narrow the match with `scope`,
`parentText`, `previousText`, `nextText`, `sameRowText`, or
`sameColumnText`. Row/column context applies only to `table-cell`. Version 6
also permits `textState: "empty"` instead of the three v5 primary predicates,
with explicit scope and parent/previous/next context. Selectors never accept
object ids, indexes, or coordinates.

A mapping requires `source` and `disposition`. It also requires `baseline`
unless the disposition is `out-of-scope`. Existing dispositions are
`copy-text`, `copy-media`, `retain-target`, `retain-target-label`, and
`out-of-scope`. Optional `cardinality` is `one`, or `all` only for an
`out-of-scope` mapping. Choice entries require `sourceMember` and a
`baselineLabel` selector whose kind is `run`. Baseline clears require
`baseline` and mode `cell` or `row`.

Candidate source selectors address only items reported in `Unresolved` by the
current automatic plan. `Plan.Mappings` are already complete and must not be
repeated. Baseline-only cleanup may select current `UnclaimedBaseline` items.

This complete but minimal v5 example shows independent existing branches.
Placeholder strings stand for unique observations from the current source or
baseline; they are not fixed document values:

```json
{
  "schema": "tiwater.docx.template-migration-semantic-candidate/v5",
  "mappings": [
    {
      "source": {
        "kind": "paragraph",
        "scope": "body",
        "text": "<current source text>",
        "previousText": "<current source context>"
      },
      "baseline": {
        "kind": "paragraph",
        "scope": "body",
        "text": "<current baseline text>",
        "previousText": "<current baseline context>"
      },
      "disposition": "copy-text"
    },
    {
      "source": {
        "kind": "paragraph",
        "scope": "header",
        "text": "<current excluded source text>"
      },
      "disposition": "out-of-scope"
    }
  ],
  "choiceSelections": [
    {
      "sourceMember": {
        "kind": "table-cell",
        "scope": "body",
        "text": "<current selected member>"
      },
      "baselineLabel": {
        "kind": "run",
        "scope": "body",
        "text": "<current baseline choice label>"
      }
    }
  ],
  "baselineClears": [
    {
      "baseline": {
        "kind": "table-cell",
        "scope": "body",
        "text": "<current baseline placeholder>"
      },
      "mode": "cell"
    }
  ]
}
```

Other existing branches use these fields:

- `bodyAppends`: `sourceStart`, `sourceEnd`.
- `bodyInsertions`: `sourceStart`, `sourceEnd`, `baselineBefore`,
  `baselineAfter`, and `stylePolicy: "target-after-context"`.
- `valueProjections`: `sourceParent`, `baselineParent`, `semantic`, `valueKind`,
  and `extraction`. Existing value kinds are `text`, `token`, `date`,
  `identifier`, and `version`; existing extraction modes are
  `after-first-delimiter`, `unique-delimited-run-group`,
  `unique-delimited-value`, and `whole-parent`.

Unknown fields and invalid branch combinations fail closed. A non-zero resolve
exit still writes its typed unresolved result to stdout; only a passing result
with no unresolved mappings may proceed to operation building.

If a completed semantic attempt has already resolved every determinate item
and leaves only genuine local business ambiguity, close those remaining items
in a separate candidate containing only `review-required` source selectors:

```bash
tiwater-docx close-template-migration-reviews <source.docx> <baseline.docx> <resolution.json> <review-candidate.json>
tiwater-docx preview-template-migration <source.docx> <baseline.docx> <closed-review.json> <output.docx>
```

The close command rejects targets, other candidate branches, sources that are
not unresolved in that resolution, and incomplete review closure. Its output
is non-pass and remains review-required; preview independently reads back only
the verified subset and does not make it delivery-eligible.

```bash
tiwater-docx build-template-migration-operations <source.docx> <baseline.docx> <plan.json>
```

### 3c. Apply and Independently Read Back a Migration

Applies only a passing operation build to the baseline, then independently
re-inventories source, baseline, and output. It checks every copied value,
baseline structure/style preservation, and OpenXML validity. A failed builder
does not create an output.

```bash
tiwater-docx apply-template-migration <source.docx> <baseline.docx> <plan.json> <output.docx>
```

### 3d. Independently Validate a Migration Output

Rebuilds source, baseline, plan admission, and output evidence in a fresh
invocation. It does not accept an apply result or trust its embedded
`Readback.Pass`. The versioned verdict binds all four file hashes and fails on
an incomplete plan, content/media mismatch, baseline structure drift, body
append drift, or newly introduced OpenXML errors.

```bash
tiwater-docx validate-template-migration-output <source.docx> <baseline.docx> <plan.json> <output.docx>
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
- `deleteBodyParagraph`
- `deleteBodyDrawingBeforeParagraph`
- `deleteBodyRange`
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
- `applyDocumentFontPolicy`
- `setTableRowHeight`
- `setTableRowCantSplit`
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
`deleteBodyParagraph` removes exactly one body descendant paragraph selected by `findText`. `matchMode` is `exact` by default and may be `startsWith`; optional `paragraphStyle` binds the selector to the current paragraph style id. Missing or ambiguous matches fail the operation.
`deleteBodyDrawingBeforeParagraph` removes the direct body paragraph immediately before a uniquely selected direct body paragraph. The removed paragraph must contain exactly one drawing and no text. `findText`, `matchMode`, and optional `paragraphStyle` select the retained anchor paragraph; missing or ambiguous anchors, non-paragraph boundaries, text-bearing drawing paragraphs, and zero or multiple drawings fail the operation.
`deleteBodyRange` removes direct body elements beginning with the uniquely selected `findText` paragraph and ending immediately before the uniquely selected following `endFindText` paragraph. Use `deleteToBodyEnd: true` instead of `endFindText` for a final body range; the document-level final section properties are preserved. A final range may set `removePrecedingPageBreak: true` to remove the single explicit page break that separated the deleted range from the preceding retained content; missing or ambiguous boundary breaks fail the operation. `matchMode` and `endMatchMode` accept `exact` or `startsWith`; optional `paragraphStyle` and `endParagraphStyle` bind each selector to a paragraph style id. All missing, ambiguous, reversed, or unsafe ranges fail the operation.
`startSectionBeforeParagraph` accepts `findText` and `orientation` (`landscape` or `portrait`); it inserts a section break before the matching direct body paragraph and applies the requested orientation to the following section.
`replaceTableCellRichText` accepts `richText` segments with `text`, optional `color`, `underline`, `bold`, and `fontName`.
An explicit `bold: false` writes an off override for both Latin and complex-script bold so paragraph- or style-level bold is not inherited.
When the target cell is empty, the generated runs inherit font-related formatting from the nearest table run so blank template cells do not fall back to Office default font size; emphasis such as bold/italic is not inherited from fallback runs. Ordinary text written into a blank cell explicitly uses baseline vertical alignment, so paragraph-mark residue cannot turn new content into superscript or subscript. Existing text-bearing cells continue to preserve their declared vertical alignment.
`replaceTable` row cell objects may use the same `richText` segments instead of plain `text`.
`insertTableRows` inserts `rows` before `rowIndex`; `templateRowIndex` controls which existing row supplies row/cell/run styling.
`deleteTableRows` deletes inclusive `startRowIndex`..`endRowIndex`, preserving the surrounding table.
`replaceTableRows` replaces inclusive `startRowIndex`..`endRowIndex` with `rows`, preserving the surrounding table and using `templateRowIndex` for row/cell/run styling. When the replaced range contains multiple row shapes, replacement rows are matched to a template row with the same `gridSpan` pattern when possible, so mixed merged/unmerged rows keep their cell widths and paragraph properties.
`insertTableColumns` inserts empty columns before a visual grid `columnIndex`; `columnCount` defaults to `1`, and `templateColumnIndex` controls which existing grid column/cell supplies width and cell styling. If the insertion point falls inside an existing horizontally merged cell, that cell's `gridSpan` is expanded instead of creating a new physical cell in that row.
`setTableWidth` accepts `width` and `widthType` (`pct`, `dxa`, `auto`, `nil`) and preserves the template table layout (`fixed`, `autofit`, or absent) instead of changing it.
`setTableCellNoWrap` accepts optional `noWrap`; `true` or omitted writes Word `w:noWrap`, and `false` removes it.
`setTableCellFontSize` accepts `fontSize` as OpenXML half-points (`18`) or points (`9pt`).
`applyDocumentFontPolicy` accepts a closed `fontPolicy` with independent `body` and `table`
rules. Each rule declares `eastAsia`, `latin`, and `size` (half-points or `pt`) without document
coordinates. `size` may instead be `preserve`; that mode changes only the four `w:rFonts`
channels and leaves both existing size channels untouched. It applies the rule to every
text-bearing main-document run in the corresponding scope.
`validate-font-policy <input.docx> <policy.json>` independently reads back every text-bearing run
and exits non-zero when any font channel differs, or when a non-preserving rule's size differs.
Preserved-size equality is established by comparing the pre-edit and output inspection evidence.
`setTableRowHeight` accepts `height` in twips and optional `heightRule` (`atLeast`, `exact`, `auto`).
`setTableRowCantSplit` accepts `cantSplit: true|false` and controls the Word table-row `w:cantSplit` property. `inspect-tables` reports the row property as `CantSplit`.
`mergeTableCells` merges a horizontal cell range when `rowIndex/startCellIndex/endCellIndex` are provided, or a vertical row range when `startRowIndex/endRowIndex` and exactly one of `cellIndex` or logical `gridColumn` are provided. Prefer `gridColumn` when rows may have different horizontal spans.
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

The CLI is a generic document runtime for inspection, export, fill, comparison, and explicit edit application.
