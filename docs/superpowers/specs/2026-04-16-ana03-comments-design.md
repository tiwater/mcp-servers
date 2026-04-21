# ANA03 Comment-Driven DOCX Editing Design

> **2026-04-21 update:** This historical design note is superseded for runtime ownership. ANA03 planning/resolution now lives in Lucid skill-local scripts, not in `mcp-servers`, and `Lucid skill-local planner` is no longer an active server/API surface. Natural-language sentence understanding for comment interpretation must be LLM-based. Regex- or keyword-driven semantic classification is prohibited for the Lucid-owned planning layer.

## Goal

Support ANA03-style report generation where the target DOCX is not a placeholder template but an annotated report sample with embedded Word comments that describe:

1. direct text replacements
2. table-wide or section-wide fill instructions
3. generated narrative requirements
4. review-only or manual instructions

The system must let the agent observe the annotated document, plan explicit edits, apply those edits through stable low-level operations, and review the result without conflating comment anchors with edit targets.

## Context

Current state:

- `docx_inspect` exposes comment anchors, but only as flat metadata.
- `docx_edit` assumes a comment anchor is usually the direct replacement target.
- ANA03 includes comments anchored inside a table cell that actually describe how to fill an entire table or larger section.
- The current model therefore lacks an explicit planning stage between “observe the comments” and “apply edits”.

This causes the exact failure mode we saw earlier: comments are real instructions, but their anchor location is only a hint, not always the final edit destination.

## Design Decision

Use a three-stage model:

1. `docx_inspect`
2. `Lucid skill-local planner`
3. `docx_edit`

And update the ANA03 `stability-report` skill to follow:

1. inspect
2. plan
3. review plan confidence and required sources
4. apply explicit edits
5. inspect again for review

## Why This Design

This keeps the boundaries clean:

- `docx_inspect` is the canonical read model
- `Lucid skill-local planner` is the interpretation layer
- `docx_edit` is the explicit mutation layer

It avoids:

- overloading `docx_edit` with planning logic
- requiring the skill to infer target scopes from raw comments every time
- hidden “finalization” semantics

## Tool Surface

### 1. `docx_inspect`

Keep the public name `docx_inspect`, but enrich its report so it is useful as the planner input.

New or improved output should include:

- comments with:
  - comment id
  - author
  - full comment text
  - anchor text
  - source part
  - nearest heading
  - paragraph/table/row/cell location
- structural neighbors for each comment:
  - current paragraph
  - current table id
  - row/column coordinates when inside a table
  - nearby preceding/following paragraph text
- richer table metadata:
  - table index
  - row and cell counts
  - first rows as structural preview

This is still an inspection tool, not a planner.

### 2. `Lucid skill-local planner`

Add a new planner tool that consumes:

- document path
- optional scenario name, starting with `stability-report`
- optional source-document hints

And returns structured plan items.

Each plan item should contain:

- `commentId`
- `commentText`
- `anchor`
- `instructionType`
  - `replace_anchor_text`
  - `fill_table_block`
  - `fill_table_row`
  - `fill_table_column`
  - `generate_paragraph`
  - `manual_only`
- `targetScope`
  - `paragraph`
  - `current_table`
  - `table_row`
  - `table_column`
  - `section`
- `candidateTargets`
- `requiredSources`
- `confidence`
- `reasoning`
- `proposedEdits`

The planner output must stay reviewable and explicit. It should not silently mutate the document.

### 3. `docx_edit`

Keep `docx_edit` explicit.

Supported operations remain mutation primitives only:

- `replaceAnchoredText`
- `replaceParagraphText`
- `replaceTableCellText`
- `deleteComment`
- `deleteComments`
- `markFieldsDirty`

Later additions can include:

- `replaceTableRowText`
- `replaceTableRegionText`

But those should remain explicit edit operations, not planning constructs.

## Planner Semantics

### Direct Replacement Comments

Pattern:

- comment is attached to text that is itself the intended replacement slot

Planner result:

- `instructionType = replace_anchor_text`
- `targetScope = paragraph` or `anchor`
- emit `replaceAnchoredText` or `replaceParagraphText`

### Table-Scope Comments

Pattern:

- comment is attached to one cell, often the first cell or title row
- comment text describes how to populate multiple rows/columns of the table

Planner result:

- `instructionType = fill_table_block`
- `targetScope = current_table`
- candidate targets identify the current table and relevant rows/columns
- emit multiple explicit table cell operations or, in a later pass, a table-region edit op

### Narrative Generation Comments

Pattern:

- comment describes how to derive a paragraph from data or from other sections

Planner result:

- `instructionType = generate_paragraph`
- `targetScope = paragraph` or `section`
- emit `replaceParagraphText`

### Manual or Unsupported Comments

Pattern:

- image insertion, signature, domain ambiguity, or missing source requirements

Planner result:

- `instructionType = manual_only`
- no edit emitted
- planner returns warning and unresolved requirement

## ANA03 Skill Changes

Update `stability-report` so it must:

1. inspect the report doc with `docx_inspect`
2. plan with `Lucid skill-local planner`
3. stop early if required items are low confidence or required sources are missing
4. apply only `proposedEdits` through `docx_edit`
5. run `docx_inspect` again for review
6. optionally remove comments and mark fields dirty only after the content pass is correct

The skill should explicitly forbid:

- treating every comment anchor as the direct edit target
- skipping the planning step for comment-annotated sample reports
- inventing edits when plan confidence is low

## First Implementation Slice

This design should be implemented in a narrow, ANA03-focused first pass:

Included:

- richer `docx_inspect` context for comments
- `Lucid skill-local planner` for:
  - anchored paragraph replacement
  - current-table / table-block instructions
  - generated narrative instructions
- skill update for `stability-report`

Deferred:

- image-slot planning
- cross-table multi-target planning
- generic multi-scenario planner generalization
- visual diff tooling

## Verification Strategy

### Synthetic Tests

Create public-safe synthetic fixtures in `mcp-servers` that model:

- comment anchored to direct replacement text
- comment anchored in a first table cell but describing the whole table
- comment describing a generated summary paragraph

Verify:

- `docx_inspect` returns the necessary context
- `Lucid skill-local planner` classifies instruction type and target scope correctly
- `docx_edit` can apply the emitted edits

### Real ANA03 Verification

Use the real ANA03 sample assets only for local verification, not as committed fixtures.

Verify:

- planner output is structurally correct
- low-confidence items are surfaced explicitly
- the `stability-report` skill can follow inspect -> plan -> edit without improvising

## Risks

- comment text may be too domain-specific for generic heuristics, requiring scenario-aware planning
- target-scope inference may be ambiguous when multiple nearby tables/paragraphs are plausible
- paragraph-only and single-cell edit primitives may still be too low-level for some table instructions

## Recommendation

Implement the planner now rather than extending anchor-only editing further.

The anchor-only model is already proven insufficient for ANA03. The next useful step is not “more edit ops first”; it is a planning layer that separates instruction interpretation from document mutation.
