# tiwater-pptx

Minimal command-line utility for PPTX inspection, text export, and placeholder filling.

## Usage

```bash
python3 packages/pptx-cli/cli.py inspect <input.pptx> --json
python3 packages/pptx-cli/cli.py export-json <input.pptx> [output.json]
python3 packages/pptx-cli/cli.py fill-template <template.pptx> <data.json> <output.pptx>
```

## Fill Data

`fill-template` accepts either a flat JSON object or `{ "textValues": { ... } }`.
Placeholders are matched as exact inline tokens like `{{title}}`.
