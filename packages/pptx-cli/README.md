# tiwater-pptx

OpenXML-based command-line utility for PPTX inspection, text export, and placeholder filling.

## Usage

```bash
tiwater-pptx inspect <input.pptx> --json
tiwater-pptx inspect <input.pptx> --json --detail
tiwater-pptx export-json <input.pptx> [output.json]
tiwater-pptx fill-template <template.pptx> <data.json> <output.pptx>
tiwater-pptx apply-format-edits <input.pptx> <plan.json> <output.pptx>
tiwater-pptx apply-template <input.pptx> <template.pptx> <plan.json> <output.pptx>
tiwater-pptx map-render-findings <inspect.json> <render-manifest.json> <findings.json> <map.json>
tiwater-pptx validate-render-finding-map <inspect.json> <render-manifest.json> <findings.json> <map.json> <verdict.json>
```

For local development fallback:

```bash
dotnet run --project packages/pptx-cli/pptx.csproj -- inspect <input.pptx> --json
```

## Fill Data

`fill-template` accepts either a flat JSON object or `{ "textValues": { ... } }`.
Placeholders are matched as exact inline tokens like `{{title}}`.

## Detailed Inspect

`inspect --json --detail` emits slide size, slide paths, shape ids/names/kinds,
master/layout/theme paths, placeholder roles, native z-order, media path/hash
bindings, shape transforms, paragraph alignment, and direct run formatting. Each shape
also reports `placeholderPresent` and the nullable `placeholderIndex` directly from its
OpenXML `p:ph` element; `placeholderType` remains the existing nullable type field.
Font size is
reported in points. Shape coordinates remain in EMU so callers can compare
native PPTX positions without lossy conversion. The detail report includes
`artifactSha256`, computed directly from the inspected PPTX bytes.

The machine-readable shape evidence contract is
[`contracts/tiwater.pptx-inspect-shape-v1.schema.json`](contracts/tiwater.pptx-inspect-shape-v1.schema.json).
`placeholderPresent` is false only when the shape has no `p:ph` element; text that
looks like a placeholder and a missing `type` attribute do not change that result.

## Format Edit Plan

`apply-format-edits` copies the input PPTX to the output path, then applies only
the targeted run-format operations listed in the plan. Operations are addressed
by slide number, shape id, and run index from `inspect --detail`.

```json
{
  "operations": [
    {
      "slideNumber": 1,
      "shapeId": 2,
      "runIndex": 0,
      "fontFamily": "微软雅黑",
      "fontSize": 16,
      "color": "287341",
      "bold": true,
      "paragraphAlignment": "center"
    }
  ]
}
```

Supported `paragraphAlignment` values are `left`, `center`, `right`,
`justified`, and `distributed`. Missing targets are reported in `issues`; they
are not silently ignored.

Rendered finding mapping is evidence-only: it identifies the current slide,
layout, or master object associated with a hash-bound WPS raster finding but
does not infer an edit. See [RENDER_FINDING_CONTRACT.md](RENDER_FINDING_CONTRACT.md).

## Template Application Plan

`apply-template` imports one inspected template master and assigns every current
slide to an explicitly selected inspected template layout. Before reassignment,
it freezes the effective geometry and text formatting inherited by current-slide
placeholders so that the target master cannot reinterpret existing content. It
then removes superseded masters after complete assignment and copies the approved
template slide size.

`systemPlaceholderPolicy` defaults to `preserve`. Setting it to
`target-template` removes source date/footer/header/slide-number placeholders so
that those system elements come only from the selected target master/layout. A
slide may explicitly materialize visible, inspected source-layout shapes through
`sourceLayoutShapeIdsToPreserve`; every imported shape is reported with both its
source and output identity. The operation never guesses which layout content is
business content.

The result receipt reports `frozenPlaceholderCount` as the number of non-system
source-slide placeholder elements processed, including placeholders identified
only by an `idx` without an explicit `type`. System placeholders removed under
`target-template` are excluded from that count and are listed in
`removedSystemPlaceholders`; a `preserve` run leaves that list empty.

Optional content fitting is scoped, never slide-wide. When `contentBounds` is
present, `contentShapeIds` must explicitly identify the current-slide shapes to
fit. Supplying either field without the other is rejected before that slide is
mutated.

```json
{
  "targetMasterPath": "ppt/slideMasters/slideMaster1.xml",
  "systemPlaceholderPolicy": "target-template",
  "slides": [
    {
      "slideNumber": 1,
      "targetLayoutPath": "ppt/slideLayouts/slideLayout1.xml",
      "contentBounds": { "x": 608400, "y": 1490400, "cx": 10969200, "cy": 4759200 },
      "contentShapeIds": [2, 3],
      "sourceLayoutShapeIdsToPreserve": []
    }
  ]
}
```
