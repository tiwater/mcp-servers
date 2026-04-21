# ANA03 Comment-Driven DOCX Editing Implementation Plan

> **2026-04-21 update:** This historical implementation plan is superseded. The generic inspection/export/edit work remains in `mcp-servers`, but the scenario-specific ANA03 planning layer now lives in Lucid skill-local scripts instead of a `Lucid skill-local planner` server/API surface.

> **2026-04-21 update:** Any sentence-understanding stage in this workflow must use an LLM-based planner. Regex- or keyword-search heuristics must not be used to interpret comment meaning.

## Scope

Implement the first ANA03 slice described in:

- `docs/superpowers/specs/2026-04-16-ana03-comments-design.md`

This slice includes:

1. richer `docx_inspect` context for comments and nearby structure
2. a new `Lucid skill-local planner` command in `docx-cli`
3. Office MCP exposure for `Lucid skill-local planner`
4. synthetic tests for inspect/plan/edit behavior
5. `stability-report` skill updates to use `inspect -> plan -> edit -> inspect`

This slice does not include:

- image-slot planning
- generic cross-table planning
- visual diff tooling

## File Map

### `~/tc/mcp-servers`

- `packages/docx-cli/Models.cs`
  - Extend inspection and planning models.
- `packages/docx-cli/Inspector.cs`
  - Enrich comment context and table metadata in `docx_inspect`.
- `packages/docx-cli/Program.cs`
  - Add the `plan` command surface.
- `packages/docx-cli/Planner.cs`
  - New ANA03-oriented planner implementation.
- `packages/docx-cli/Editor.cs`
  - Reuse existing edit ops; add only if planner needs one missing primitive.
- `packages/docx-cli/README.md`
  - Document `Lucid skill-local planner` and the inspect/plan/edit workflow.
- `packages/docx-cli.tests/*.cs`
  - Add synthetic tests and fixtures for comment-driven planning.
- `servers/office/index.mjs`
  - Expose `Lucid skill-local planner` through the MCP server.
- `servers/office/README.md`
  - Document the new MCP tool.

### `~/tc/lucid`

- `skills/stability-report/SKILL.md`
  - Update the skill flow to require `Lucid skill-local planner` before `docx_edit`.

## Task 1: Extend DOCX Models

- [ ] Add planning model types in `packages/docx-cli/Models.cs`.

Add:

- `DocxPlanRequest`
- `DocxPlanResult`
- `DocxPlanItem`
- `DocxPlanCandidateTarget`

And extend `AnnotationAnchor` or add a richer structure so inspection includes:

- nearest heading
- nearby paragraph text
- current table preview metadata

- [ ] Keep existing model names stable where possible.

Do not break:

- `docx_inspect`
- `docx_edit`
- existing JSON property names that current skills already depend on

## Task 2: Enrich `docx_inspect`

- [ ] Update `packages/docx-cli/Inspector.cs` so annotation anchors include richer nearby context.

Implementation requirements:

- nearest heading text
- current paragraph text
- preceding paragraph text
- following paragraph text
- current table index if inside a table
- table row/cell coordinates if inside a table
- lightweight table preview:
  - row count
  - cell count for the first rows

- [ ] Keep output small enough for LLM use.

Do not dump whole tables. Include only small previews and structural hints.

- [ ] Verify `docx_inspect` still works on existing test fixtures.

Run:

```bash
dotnet build /Users/hugh/tc/mcp-servers/packages/docx-cli/docx.csproj
```

Expected:

- build succeeds

## Task 3: Add `Lucid skill-local planner` Models and Planner

- [ ] Create `packages/docx-cli/Planner.cs`.

Implement a planner that:

- loads the inspection report
- classifies each comment into:
  - `replace_anchor_text`
  - `fill_table_block`
  - `generate_paragraph`
  - `manual_only`
- resolves target scope:
  - `paragraph`
  - `current_table`
  - `section`
- emits `proposedEdits` using existing `docx_edit` operation types

- [ ] Restrict first-pass planner heuristics to ANA03-focused patterns.

Rules:

- if the comment anchor and comment text clearly point to local replacement, emit `replaceAnchoredText`
- if comment is in the first cell or title area of a table and references table content broadly, emit a `fill_table_block` plan item with explicit candidate table target
- if the comment describes a derived narrative paragraph, emit `replaceParagraphText`
- if no safe classification exists, emit `manual_only` with low confidence and no edit

- [ ] Ensure planner results include confidence and reasoning.

Required fields:

- `instructionType`
- `targetScope`
- `candidateTargets`
- `requiredSources`
- `confidence`
- `reasoning`
- `proposedEdits`

## Task 4: Add `plan` Command to `docx-cli`

- [ ] Update `packages/docx-cli/Program.cs`.

Add:

```text
plan <input.docx> <plan-data.json>
```

Behavior:

- parse `DocxPlanRequest`
- run `Planner.Plan(...)`
- print JSON

- [ ] Keep `inspect`, `compare`, and `edit` behavior unchanged.

## Task 5: Expose `Lucid skill-local planner` in Office MCP

- [ ] Update `servers/office/index.mjs`.

Add a new MCP tool:

- `Lucid skill-local planner`

Input:

- `input`
- `data` or `dataPath`

Output:

- `plan`

- [ ] Update `servers/office/README.md` to list `Lucid skill-local planner`.

## Task 6: Document `Lucid skill-local planner`

- [ ] Update `packages/docx-cli/README.md`.

Document:

- when to use `Lucid skill-local planner`
- how it differs from `docx_inspect`
- how it pairs with `docx_edit`

Show the intended workflow:

```text
inspect -> plan -> edit -> inspect
```

## Task 7: Add Synthetic Planner Tests

- [ ] Add or update tests in `packages/docx-cli.tests/`.

Create public-safe fixtures for:

1. direct replacement comment
2. table-scope comment anchored in a first cell
3. narrative-generation comment

- [ ] Add tests for `docx_inspect`.

Verify:

- richer anchor context exists
- table metadata is present

- [ ] Add tests for `Lucid skill-local planner`.

Verify:

- direct replacement comment -> `replace_anchor_text`
- table comment -> `fill_table_block`
- narrative comment -> `generate_paragraph`
- low-confidence cases -> `manual_only`

- [ ] Add tests that `docx_edit` can apply planner-emitted edits for supported cases.

Run:

```bash
dotnet test /Users/hugh/tc/mcp-servers/packages/docx-cli.tests/docx-cli.tests.csproj
```

Expected:

- all tests pass

## Task 8: Update `stability-report` Skill

- [ ] Update `~/tc/lucid/skills/stability-report/SKILL.md`.

Required changes:

- annotated DOCX path must use:
  1. `docx_inspect`
  2. `Lucid skill-local planner`
  3. `docx_edit`
  4. `docx_inspect` for review
- explicitly forbid skipping planning for comment-annotated report samples
- explicitly forbid improvising edits when plan confidence is low

- [ ] Update the installed copy if needed:

- `~/.supen/skills/stability-report/SKILL.md`

## Task 9: Real ANA03 Verification

- [ ] Run local verification against the real ANA03 sample assets without committing them.

Verify:

- `docx_inspect` returns richer comment context
- `Lucid skill-local planner` produces structured plan items
- table-scope comments do not get mistaken for direct anchor replacements
- low-confidence items are explicit

- [ ] If safe planner edits exist, apply them on a throwaway copy with `docx_edit`.

Do not mutate the original fixture.

## Task 10: Final Verification

- [ ] Run final build and tests:

```bash
dotnet build /Users/hugh/tc/mcp-servers/packages/docx-cli/docx.csproj
dotnet test /Users/hugh/tc/mcp-servers/packages/docx-cli.tests/docx-cli.tests.csproj
```

- [ ] Verify Office MCP still starts:

```bash
node /Users/hugh/tc/mcp-servers/servers/office/index.mjs
```

- [ ] Check `lucid` skill diff for ANA03 wording consistency.

## Commit Plan

- `mcp-servers`
  - `feat(docx): add ANA03 planning workflow`
- `lucid`
  - `docs(skills): route ANA03 through docx planning`

## Risks to Watch

- planner heuristics may overfit to one ANA03 comment pattern
- current `docx_edit` primitives may be too weak for some table-wide instructions
- inspection output may become too verbose unless previews stay small

## Success Criteria

This slice is complete when:

1. `docx_inspect` exposes enough context for planning
2. `Lucid skill-local planner` classifies direct, table-scope, and narrative comments on synthetic fixtures
3. Office MCP exposes `Lucid skill-local planner`
4. `stability-report` uses `Lucid skill-local planner` explicitly
5. real ANA03 verification shows the planner no longer equates comment anchor with edit target by default
