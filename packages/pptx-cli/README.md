# tiwater-pptx

`tiwater-pptx` provides technical PPTX observation, fixed Open XML mutation,
export, and package validation. It does not own business mappings, workflow
decisions, or delivery status.

The Agent-facing surface is the published Office MCP. The CLI exposes the same
provider requests for diagnostics and package integration; it is not a second
workflow protocol.

## Observation

```bash
tiwater-pptx inspect input.pptx --json
tiwater-pptx export-json input.pptx output.json
```

Observation reports current slides, masters, layouts, shapes, transforms,
paragraphs, runs, pictures, and placeholders. It does not infer slide meaning,
preferred layouts, or business fields.

Paragraph observation includes effective bullet kind/character, numbering type
and start, left margin and hanging indent. Content/object placeholders (including
the omitted default object type) inherit the master body text style. Direct
paragraph values override inherited values; a direct no-bullet value suppresses
an inherited bullet.

## Fixed technical mutation

```bash
tiwater-pptx pptx_apply_template request.json
tiwater-pptx pptx_apply_format request.json
tiwater-pptx pptx_set_shape_geometry request.json
tiwater-pptx pptx_replace_picture_image request.json
```

Each mutation command consumes the matching provider-owned request contract
from `contracts/mcp-input/`. Requests contain no operation discriminator. The
provider preserves unselected slide content and publishes no output when a
requested technical action cannot be completed.

Template application may explicitly select ordinary source text shapes as
`sourceSystemShapes` within a slide assignment. The caller supplies each shape
identity and its date, footer, header, or slide-number role. This requires the
`target-template` system placeholder policy. The provider removes only those
selected ordinary text shapes; it does not infer roles from text or appearance.
Pictures, tables, native placeholders, and shapes also selected for content
fitting are rejected. Omitting this selection preserves the existing behavior.

Template application accepts `preserveSourceAppearance: true` to retain the source
slide theme and color mapping while changing layout geometry. Ordinary source
text keeps its existing direct styling; inherited placeholder styles are
materialized before changing layouts. Omission retains the previous behavior.

## Validation

```bash
tiwater-pptx validate input.pptx
```

Validation proves package integrity and technical postconditions only. It does
not decide whether presentation content or appearance is correct for a business
task.

## Discovery

```bash
tiwater-pptx --list-tools
tiwater-pptx <command> --help
```

The provider tool list contains technical commands only. The Office MCP adapter
must expose the same provider-owned requests without adding business fields.

Detailed inspection reports effective placeholder geometry: a slide overrides each offset/extent it declares, then inherits missing components from its uniquely matched layout and master. Slide/layout identity uses the placeholder index (default zero); layout/master identity uses placeholder type. Ambiguous or missing inheritance remains unavailable, never guessed from object names.
