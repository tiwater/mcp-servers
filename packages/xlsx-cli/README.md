# tiwater-xlsx

A .NET 9 globally installed command-line tool for dynamically inspecting and filling `.xlsx` templates.

## Installation

Install the tool from the NuGet global registry using the modern .NET CLI:

```bash
dotnet tool install -g tiwater.xlsx.cli
```

## Placeholder Syntax

In your target Excel template (`.xlsx`):
*   **Single Cells**: Should be formatted exactly as `{{placeholder_key}}` (e.g., `{{controlledNumber}}`). The entire cell's content must just be the placeholder text if it's meant to be replaced entirely.
*   **Data Grids/Tables**: Should be anchored with `{{table:placeholder_key}}`. The CLI will auto-fill a 2D array downwards and to the right starting directly from that anchored cell.

## Usage

### 1. Inspect a Template
Outputs all discovered data placeholders present inside an Excel template, which helps agentic workflows know what shape of data is required before filling it.

```bash
tiwater-xlsx inspect <template.xlsx> [--json]
```
*   `--json` returns structured output suitable for parsers.

### 2. Fill a Template
Injects the defined JSON payload directly into an active Excel sheet, replacing matched placeholders and rendering the final result document.

```bash
tiwater-xlsx fill-template <template.xlsx> <data.json> <output.xlsx>
```

#### Expected JSON Model

The structured shape of `<data.json>` expected by `fill-template` must look like the following:

```json
{
  "cellValues": {
    "controlledNumber": "260359",
    "calculationResult": "0.98",
    "placeholder_name": "example_value"
  },
  "tableData": {
    "peakAreas": [
      ["Peak1", "Area1", "RT1"],
      ["Peak2", "Area2", "RT2"]
    ]
  }
}
```
