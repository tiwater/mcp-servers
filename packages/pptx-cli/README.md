# tiwater-pptx

OpenXML-based command-line utility for PPTX inspection, text export, and placeholder filling.

## Usage

```bash
tiwater-pptx inspect <input.pptx> --json
tiwater-pptx export-json <input.pptx> [output.json]
tiwater-pptx fill-template <template.pptx> <data.json> <output.pptx>
```

For local development fallback:

```bash
dotnet run --project packages/pptx-cli/pptx.csproj -- inspect <input.pptx> --json
```

## Fill Data

`fill-template` accepts either a flat JSON object or `{ "textValues": { ... } }`.
Placeholders are matched as exact inline tokens like `{{title}}`.
