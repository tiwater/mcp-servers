# tiwater-pptx

OpenXML-based command-line utility for PPTX inspection, text export, and placeholder filling.

## Usage

```bash
tiwater-pptx inspect <input.pptx> --json
tiwater-pptx inspect <input.pptx> --json --detail
tiwater-pptx export-json <input.pptx> [output.json]
tiwater-pptx fill-template <template.pptx> <data.json> <output.pptx>
tiwater-pptx apply-format-edits <input.pptx> <plan.json> <output.pptx>
tiwater-pptx inspect-evidence --request request.json --output evidence.json
tiwater-pptx validate-inspect-evidence --request request.json --evidence evidence.json --output verdict.json
tiwater-pptx apply-template <input.pptx> <template.pptx> <plan.json> <output.pptx>
```

The evidence commands reuse detailed PPTX inspection. Validation re-opens the presentation and independently recomputes the evidence.

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
bindings, shape transforms, paragraph alignment, and direct run formatting. Font size is
reported in points. Shape coordinates remain in EMU so callers can compare
native PPTX positions without lossy conversion.

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

## Template Application Plan

`apply-template` imports one inspected template master and assigns every current
slide to an explicitly selected inspected template layout. It preserves current
slide shapes and content, removes superseded masters after complete assignment,
and copies the approved template slide size.

```json
{
  "targetMasterPath": "ppt/slideMasters/slideMaster1.xml",
  "slides": [
    { "slideNumber": 1, "targetLayoutPath": "ppt/slideLayouts/slideLayout1.xml" }
  ]
}
```
