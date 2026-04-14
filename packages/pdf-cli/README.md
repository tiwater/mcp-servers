# tiwater-pdf

A Python command-line utility for inspecting PDF documents and extracting tabular data, heavily utilized in analytical reporting workflows (e.g., HPLC reports).

## Installation

This tool requires Python 3.11+. We recommend installing it using modern package managers like `uv` or `pipx` to avoid global environment conflicts:

```bash
# Recommend approach using uv:
uv tool install tiwater-pdf

# Or using pipx:
pipx install tiwater-pdf

# Fallback (may require --break-system-packages on newer OS):
pip install tiwater-pdf
```

## Commands Reference

The CLI provides three major functionalities:

### 1. Find a Specific Table
Searches the document for a table matching a specific heading or name and attempts to extract it.

```bash
tiwater-pdf find-table <report.pdf> "<table_name>" [--auto-span] [--json]
```
*   `--auto-span`: Enables heuristics to span tables that break across multiple pages.
*   `--json`: Outputs the table data entirely in machine-readable JSON format.

### 2. Extract All Tables
Extracts all tables detected within the PDF or from specific pages.

```bash
tiwater-pdf extract-tables <report.pdf> [--pages 1,3,4] [--auto-span] [--json]
```

### 3. Inspect PDF
Provides a high-level inspection of the PDF's structural layout and tables to determine its format.

```bash
tiwater-pdf inspect <report.pdf>
```
