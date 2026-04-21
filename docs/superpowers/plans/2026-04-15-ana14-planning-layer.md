# ANA14 Planning Layer

> **2026-04-21 update:** This historical planning note is superseded. ANA14 scenario planning now lives in Lucid skill-local scripts, not in `mcp-servers`, and `Lucid skill-local planner` is no longer an active server/API surface.

## Goal

Add a planning layer for the ANA14 workflow so the agent can reliably:

1. inspect a fixed-layout workbook and extracted PDF tables
2. choose the correct source tables and rows
3. map them into the workbook's target regions
4. compute derived values for the third section
5. emit explicit `xlsx_edit` operations instead of improvising cell writes

## Scope

This plan covers:

- extending `xlsx_inspect` so it is useful as planning input
- adding a stable planning schema for ANA14-style workflows
- building sample-driven tests against synthetic workbooks/PDF-derived data
- verifying the planner logic against the real ANA14 sample assets

This plan does not yet cover:

- generic multi-scenario spreadsheet planning
- direct LLM orchestration changes in Supen
- full autonomous execution of ANA14 from Lucid queue input

## Current State

- `xlsx_fill_template` is suitable only for placeholder-based workbooks.
- `xlsx_edit` now supports explicit fixed-layout edits.
- ANA14 sample workbook has zero placeholders and contains embedded sheet instructions.
- PDF extraction returns multiple candidate tables, including the peak-area tables we need.
- The remaining hard problem is selecting the right extracted tables, mapping them to workbook regions, and computing derived values in a repeatable way.

## Target Design

### 1. Canonical Inspect Output

Evolve `xlsx_inspect` so it is the single low-level read surface for planning. The output should include enough structure to answer:

- workbook sheets and dimensions
- non-empty cell regions
- merged regions
- formulas vs literal cells
- likely instruction rows or note rows
- headings / section labels
- style-informed boundaries where possible

The goal is not to expose every spreadsheet detail. The goal is to make fixed-layout planning practical without adding a second inspection tool.

### 2. ANA14 Planning Schema

Add a planning representation in `xlsx-cli` for ANA14-like workflows. The planner should consume:

- workbook inspection output
- extracted PDF tables
- scenario-specific assumptions

And produce:

- chosen source tables for 280nm and 360nm
- row matches by sample / identifier
- target workbook ranges for each section
- derived values for the computed section
- explicit `xlsx_edit` operations
- warnings / unresolved ambiguities

Suggested shape:

```json
{
  "scenario": "experimental-record-attachment",
  "sheet": "附件1",
  "sections": [
    {
      "name": "280nm",
      "sourceTableRef": "pdf280:table-3",
      "targetRange": "B5:G12",
      "edits": []
    }
  ],
  "warnings": [],
  "confidence": "high"
}
```

### 3. Public Planner Entry Point

Add a first public planning command in `xlsx-cli`, exposed through the Office MCP server. Recommended surface:

- `Lucid skill-local planner`

Inputs:

- workbook path
- extracted table payloads or JSON file paths
- optional scenario name

Outputs:

- structured plan
- proposed `xlsx_edit` operations
- warnings and confidence markers

This keeps planning explicit and reviewable before mutation.

### 4. Sample-Driven Tests

Create synthetic ANA14-like fixtures in `~/tc/mcp-servers` that mimic:

- a fixed-layout workbook with three sections
- source table identifiers for 280nm and 360nm
- derived third-section calculations
- note rows / instruction rows

Tests should verify:

- planner selects the correct source tables
- planner maps rows to the correct workbook regions
- planner computes derived values consistently
- generated `xlsx_edit` operations apply cleanly

### 5. Real-Sample Verification

After synthetic coverage exists, verify the planner against the real ANA14 assets:

- sample workbook
- 280nm PDF
- 360nm PDF

This verification should prove:

- the planner can identify the right candidate tables from the extracted set
- the target workbook regions are correct
- no placeholder assumptions remain

## Implementation Steps

### Step 1. Extend `xlsx_inspect`

Update `xlsx-cli` inspection output to include:

- used-range metadata per sheet
- merged cell regions
- literal/formula cell classification
- grouped non-empty regions
- candidate instruction rows

Files likely involved:

- `packages/xlsx-cli/Inspector.cs`
- `packages/xlsx-cli/Models.cs`
- `packages/xlsx-cli/Program.cs`

### Step 2. Define Planner Models

Add models for:

- extracted table references
- row matches
- section plans
- warnings / ambiguities
- generated edit operations

Files likely involved:

- `packages/xlsx-cli/Models.cs`
- new planner implementation file, likely `packages/xlsx-cli/Planner.cs`

### Step 3. Add `Lucid skill-local planner`

Implement a new command in `xlsx-cli` and expose it in the Office MCP server.

Files likely involved:

- `packages/xlsx-cli/Program.cs`
- `packages/xlsx-cli/Planner.cs`
- `servers/office/index.mjs`
- `servers/office/README.md`

### Step 4. Build Synthetic Tests

Create tests and fixtures covering:

- correct table selection
- correct row mapping
- correct derived calculations
- correct emitted edit operations

Files likely involved:

- `packages/xlsx-cli.tests/*`

### Step 5. Verify Against Real ANA14 Assets

Run:

- workbook inspection
- PDF extraction
- planning
- `xlsx_edit` application on a throwaway workbook copy

Then inspect the resulting workbook JSON to confirm edits landed in the right ranges.

## Verification Plan

Required checks:

```bash
dotnet build /Users/hugh/tc/mcp-servers/packages/xlsx-cli/xlsx.csproj
dotnet test /Users/hugh/tc/mcp-servers/packages/xlsx-cli.tests/xlsx-cli.tests.csproj
```

And scenario checks:

```bash
dotnet run --project /Users/hugh/tc/mcp-servers/packages/xlsx-cli/xlsx.csproj -- inspect '/path/to/sample.xlsx'
uv run --project /Users/hugh/tc/mcp-servers/packages/pdf-cli tiwater-pdf extract-tables '/path/to/280nm.pdf' --json
uv run --project /Users/hugh/tc/mcp-servers/packages/pdf-cli tiwater-pdf extract-tables '/path/to/360nm.pdf' --json
dotnet run --project /Users/hugh/tc/mcp-servers/packages/xlsx-cli/xlsx.csproj -- plan ...
dotnet run --project /Users/hugh/tc/mcp-servers/packages/xlsx-cli/xlsx.csproj -- edit ...
```

## Risks

- extracted PDF tables may not be stable enough across layouts for exact matching without heuristics
- workbook layout may drift between scenarios, so ANA14-specific assumptions must stay explicit
- formula preservation needs care when writing into partially computed regions
- planner output must stay reviewable; hidden auto-corrections will make debugging harder

## Recommendation

Implement this in two passes:

1. ANA14-specific planner with explicit assumptions and strong tests
2. only after it proves stable, generalize the planner shape for other fixed-layout spreadsheet workflows
