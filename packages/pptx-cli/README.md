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
