using System.Text.Json;
using Dockit.Docx;

namespace Dockit.Docx.Cli;

internal static class Program
{
    public static Task<int> Main(string[] args) => Cli.RunAsync(args);
}

public static class Cli
{
    private static readonly string[] DiscoverableCommands =
    [
        "start-template-migration-decisions",
        "find-template-migration-targets",
        "record-template-migration-decision",
        "resolve-template-migration-decisions",
        "list-template-migration-choices",
        "resolve-template-migration-choices",
        "list-template-migration-options",
        "resolve-template-migration-semantic-candidate",
        "close-template-migration-reviews",
        "build-template-migration-operations",
        "apply-template-migration",
        "validate-template-migration-output",
        "preview-template-migration",
    ];

    public static Task<int> RunAsync(string[] args)
    {
        if (args.Length == 1 && args[0] == "--list-tools")
        {
            WriteJson(new { schema = "tiwater.provider-tool-list/v1", commands = DiscoverableCommands });
            return Task.FromResult(0);
        }

        if (args.Length == 1 && args[0] == "--describe-tool")
        {
            PrintToolDescription();
            return Task.FromResult(0);
        }

        if (args.Length == 1 && args[0] is "--help" or "-h")
        {
            PrintUsage();
            return Task.FromResult(0);
        }

        if (args.Length == 0)
        {
            PrintUsage();
            return Task.FromResult(1);
        }

        if (args.Length == 2 && args[1] is "--help" or "-h" && PrintCommandUsage(args[0]))
        {
            return Task.FromResult(0);
        }

        try
        {
            return args[0] switch
            {
                "inspect" => RunInspectAsync(args[1..]),
                "inspect-tables" => RunInspectTablesAsync(args[1..]),
                "compare" => RunCompareAsync(args[1..]),
                "validate-template-transform" => RunValidateTemplateTransformAsync(args[1..]),
                "validate-openxml" => Task.FromResult(OpenXmlValidation.Run(args[1..])),
                "analyze-template-migration" => Task.FromResult(TemplateMigration.RunAnalyze(args[1..])),
                "derive-template-migration-exact-text-plan" => Task.FromResult(TemplateMigration.RunDeriveExactTextPlan(args[1..])),
                "list-template-migration-options" or "find-template-migration-candidates" => Task.FromResult(TemplateMigration.RunFindCandidates(args[1..])),
                "start-template-migration-decisions" => Task.FromResult(TemplateMigration.RunStartDecisions(args[1..])),
                "find-template-migration-targets" => Task.FromResult(TemplateMigration.RunListDecisionTargets(args[1..])),
                "record-template-migration-decision" => Task.FromResult(TemplateMigration.RunRecordDecision(args[1..])),
                "resolve-template-migration-decisions" => Task.FromResult(TemplateMigration.RunResolveDecisionDraft(args[1..])),
                "list-template-migration-choices" => Task.FromResult(TemplateMigration.RunListChoices(args[1..])),
                "resolve-template-migration-choices" => Task.FromResult(TemplateMigration.RunResolveChoices(args[1..])),
                "derive-template-migration-anchor-gap-plan" => Task.FromResult(TemplateMigration.RunDeriveAnchorGapPlan(args[1..])),
                "resolve-template-migration-semantic-candidate" => Task.FromResult(TemplateMigration.RunResolveSemanticCandidate(args[1..])),
                "close-template-migration-reviews" => Task.FromResult(TemplateMigration.RunCloseReviews(args[1..])),
                "build-template-migration-operations" => Task.FromResult(TemplateMigration.RunBuildOperations(args[1..])),
                "apply-template-migration" => Task.FromResult(TemplateMigration.RunApply(args[1..])),
                "validate-template-migration-output" => Task.FromResult(TemplateMigration.RunValidateOutput(args[1..])),
                "preview-template-migration" => Task.FromResult(TemplateMigration.RunPreview(args[1..])),
                "strip-direct-formatting" => Task.FromResult(Transforms.RunStripDirectFormatting(args[1..])),
                "replace-style-ids" => Task.FromResult(Transforms.RunReplaceStyleIds(args[1..])),
                "export-json" => Task.FromResult(Transforms.RunExportJson(args[1..])),
                "fill-template" => Task.FromResult(Transforms.RunFillTemplate(args[1..])),
                "normalize-openxml" => Task.FromResult(DocxPackageNormalizer.RunNormalize(args[1..])),
                "edit" => Task.FromResult(Editor.RunEdit(args[1..])),
                "validate-font-policy" => Task.FromResult(FontPolicy.RunValidate(args[1..])),
                _ => FailUnknown(args[0]),
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return Task.FromResult(1);
        }
    }

    private static Task<int> RunInspectAsync(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("inspect requires <input.docx>");
        }

        var input = args[0];
        var json = args.Skip(1).Contains("--json", StringComparer.Ordinal);
        if (json)
        {
            WriteJson(new
            {
                document = Inspector.Inspect(input),
                tables = Inspector.InspectTables(input),
                flow = Inspector.InspectDocumentFlow(input),
                fonts = FontPolicy.Inspect(input),
            });
            return Task.FromResult(0);
        }
        var report = Inspector.Inspect(input);

        RenderInspect(report);

        return Task.FromResult(0);
    }

    private static Task<int> RunInspectTablesAsync(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("inspect-tables requires <input.docx>");
        }

        var input = args[0];
        var json = args.Skip(1).Contains("--json", StringComparer.Ordinal);
        var report = Inspector.InspectTables(input);

        if (json)
        {
            WriteJson(report);
        }
        else
        {
            RenderInspectTables(report);
        }

        return Task.FromResult(0);
    }

    private static Task<int> RunCompareAsync(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("compare requires <old.docx> <new.docx>");
        }

        var baseline = args[0];
        var updated = args[1];
        var json = args.Skip(2).Contains("--json", StringComparer.Ordinal);
        var report = Comparer.Compare(baseline, updated);

        if (json)
        {
            WriteJson(report);
        }
        else
        {
            RenderCompare(report);
        }

        return Task.FromResult(0);
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  inspect <input.docx> [--json]");
        Console.WriteLine("  inspect-tables <input.docx> [--json]");
        Console.WriteLine("  compare <old.docx> <new.docx> [--json]");
        Console.WriteLine("  validate-template-transform <source-template.docx> <target-template.docx> [--json]");
        Console.WriteLine("  validate-openxml <input.docx>");
        Console.WriteLine("Template migration: record one current-document decision at a time; read each next command's --help before using it.");
        Console.WriteLine("  start-template-migration-decisions <source.docx> <baseline.docx> <draft.json>");
        Console.WriteLine("  find-template-migration-targets <source.docx> <baseline.docx> <draft.json> <branch> [query|-] [offset] [limit]");
        Console.WriteLine("  record-template-migration-decision <source.docx> <baseline.docx> <draft.json> <branch> <branch arguments>");
        Console.WriteLine("  resolve-template-migration-decisions <source.docx> <baseline.docx> <draft.json>");
        Console.WriteLine("  list-template-migration-choices <source.docx> <baseline.docx>  (compatibility)");
        Console.WriteLine("  resolve-template-migration-choices <source.docx> <baseline.docx> <choices.json>  (compatibility)");
        Console.WriteLine("  list-template-migration-options <source.docx> <baseline.docx>  (compatibility)");
        Console.WriteLine("  resolve-template-migration-semantic-candidate <source.docx> <baseline.docx> <candidate.json>  (compatibility)");
        Console.WriteLine("  build-template-migration-operations <source.docx> <baseline.docx> <plan.json>");
        Console.WriteLine("  apply-template-migration <source.docx> <baseline.docx> <plan.json> <output.docx>");
        Console.WriteLine("  validate-template-migration-output <source.docx> <baseline.docx> <plan.json> <output.docx>");
        Console.WriteLine("  preview-template-migration <source.docx> <baseline.docx> <plan.json> <output.docx>");
        Console.WriteLine("  strip-direct-formatting <input.docx> <output.docx>");
        Console.WriteLine("  replace-style-ids <input.docx> <output.docx> <style-map.json>");
        Console.WriteLine("  export-json <input.docx> [<output.json>]");
        Console.WriteLine("  fill-template <template.docx> <data.json> <output.docx>");
        Console.WriteLine("  normalize-openxml <input.docx> <output.docx>");
        Console.WriteLine("  edit <input.docx> <operations.json> <output.docx>");
        Console.WriteLine("  validate-font-policy <input.docx> <policy.json>");
    }

    private static void PrintToolDescription()
    {
        Console.WriteLine("Purpose: Inspect, migrate, edit, normalize, and validate DOCX documents through the published DOCX provider.");
        Console.WriteLine("Consumes: DOCX files and the typed plans, candidates, operations, or policies required by the selected subcommand.");
        Console.WriteLine("Produces: Published DOCX observations, plans, edited documents, previews, and validation receipts.");
        Console.WriteLine("Use when: A scenario-declared capability requires DOCX inspection, template migration, editing, normalization, or validation.");
        Console.WriteLine("Do not use for: Choosing scenario semantics, inventing business values, deciding delivery, rendering Office pages, or OCR.");
        Console.WriteLine("Command discovery: Run tiwater-docx --help; start template migration with start-template-migration-decisions, then read each command's --help when its inputs exist.");
        Console.WriteLine("Usage: tiwater-docx <command> [arguments]");
    }

    private static bool PrintCommandUsage(string command)
    {
        var help = command switch
        {
            "analyze-template-migration" => """
                Purpose: Observe the current source and selected baseline as immutable migration object inventories.
                Consumes: One current source DOCX and one selected baseline DOCX.
                Produces: A hash-bound analysis with source objects, baseline objects, candidate-ready unique semantic selectors, and unresolved findings; use --json for machine output.
                Use when: Diagnosing low-level source and baseline object inventories for an existing caller.
                Do not use for: Starting semantic resolution, selecting mappings, building operations, editing a document, or validating output. New migration callers start with start-template-migration-decisions.
                Usage:
                  tiwater-docx analyze-template-migration <source.docx> <baseline.docx> [--json]
                """,
            "derive-template-migration-exact-text-plan" => """
                Purpose: Preserve the legacy exact-text diagnostic receipt for compatible callers.
                Consumes: One current source DOCX and the same selected baseline DOCX used for analysis.
                Produces: A hash-bound plan; Unresolved[].Source and Unresolved[].BaselineOptions carry current observations, and UnclaimedBaseline lists unclaimed baseline content and selectable child runs. Exact-text match missing or non-unique describes only this mechanical comparison; it does not mean that a semantic target is absent or ambiguous.
                Use when: An existing caller still consumes this diagnostic receipt; new semantic-resolution callers use start-template-migration-decisions.
                Do not use for: New tool discovery, treating Unresolved as review-required, building operations, inventing target mappings, editing, or output validation.
                Usage:
                  tiwater-docx derive-template-migration-exact-text-plan <source.docx> <baseline.docx>
                """,
            "list-template-migration-options" or "find-template-migration-candidates" => """
                Purpose: Preserve the selector-level migration discovery artifact for compatible callers.
                Consumes: One current source DOCX and one selected baseline DOCX; the provider performs its conservative exact comparison internally.
                Produces: Uniform RequiredDecisions and AvailableTargets observations bound to the current source and baseline hashes. Each distinguishable required source appears once with its structural context. Repeated sources that cannot be separated by a semantic selector appear as one Count > 1 group with RequiredCardinality=all. AvailableTargets groups selectable baseline paragraphs, cells, and media with the same contextual shape and selectable child runs. It does not produce a migration plan or target recommendation.
                Use when: An existing integration is explicitly bound to selector-level semantic candidates.
                Do not use for: New Agent-facing migration work, choosing copy, retain, exclude, or review semantics, building a migration plan, or executing document operations. New callers start with start-template-migration-decisions.
                Next for compatible callers: Use current scenario authority and AvailableTargets to propose one semantic disposition for every RequiredDecision, then call resolve-template-migration-semantic-candidate. A RequiredCardinality=all group can be closed only by an out-of-scope mapping with cardinality=all; if those repeated items are business facts, stop because they cannot be assigned individually.
                Output fields:
                  top level: Schema, Pass, SourceSha256, BaselineSha256, RequiredDecisions, AvailableTargets
                  RequiredDecisions[]: Source, Count, RequiredCardinality
                  Source and AvailableTargets[]: Kind, Scope, Text, Selector, Context
                  Context: PreviousText, NextText, SameRowTexts, SelectableChildren
                Usage:
                  tiwater-docx list-template-migration-options <source.docx> <baseline.docx>
                Compatibility alias: find-template-migration-candidates
                """,
            "start-template-migration-decisions" => """
                Purpose: Create a provider-owned run-local draft and return the first unresolved current source choice.
                Consumes: One current source DOCX, one selected baseline DOCX, and a new draft path.
                Produces: An atomic decision draft plus progress containing only counts and the next source choice.
                Use when: Starting Agent-guided template migration without asking the Agent to author candidate JSON.
                Do not use for: Choosing business meaning, overwriting an existing draft, building operations, or editing a document.
                Usage:
                  tiwater-docx start-template-migration-decisions <source.docx> <baseline.docx> <draft.json>
                Next: For the returned source choice, inspect targets when needed and record one decision.
                """,
            "find-template-migration-targets" => """
                Purpose: Page or filter the complete structurally eligible target choices for one current source and decision branch.
                Consumes: The same current source and baseline, the provider-owned draft, a branch, and optional literal query, offset, and limit.
                Produces: One bounded target page; ordering is stable and no target is ranked by presumed business meaning.
                Use when: The next business decision needs a target without loading the full baseline inventory into Agent context.
                Do not use for: Choosing a target, hiding targets outside a page, inferring scenario semantics, or editing a document.
                Usage:
                  tiwater-docx find-template-migration-targets <source.docx> <baseline.docx> <draft.json> <copy-text|copy-media|retain-target|retain-target-label|choice-selection|baseline-clear> [query|-] [offset] [limit]
                """,
            "record-template-migration-decision" => """
                Purpose: Validate and atomically record exactly one business decision in the provider-owned draft.
                Consumes: The same current source and baseline, the draft path, and one mapping, choice-selection, or baseline-clear decision. The draft supplies the current source identity.
                Produces: The updated draft plus progress containing counts and the next unresolved source choice.
                Use when: The current scenario and document observations determine one source disposition or one baseline cleanup.
                Do not use for: Authoring draft JSON, supplying selectors or coordinates, guessing a business decision, or editing a document.
                Usage:
                  tiwater-docx record-template-migration-decision <source.docx> <baseline.docx> <draft.json> mapping <disposition> <target-choice-id|-> [cardinality|-]
                  tiwater-docx record-template-migration-decision <source.docx> <baseline.docx> <draft.json> choice-selection <target-choice-id>
                  tiwater-docx record-template-migration-decision <source.docx> <baseline.docx> <draft.json> baseline-clear <target-choice-id> <cell|row>
                Mapping dispositions: copy-text, copy-media, retain-target, retain-target-label, out-of-scope, review-required. The last two use - as the target.
                Recording an explicit previously returned source choice replaces that source's earlier decision atomically for compatibility with a typed validation correction.
                """,
            "resolve-template-migration-decisions" => """
                Purpose: Re-read both current documents and expand the complete provider-owned decision draft into a validated migration plan.
                Consumes: The same current source DOCX, selected baseline DOCX, and provider-owned draft.
                Produces: A hash-bound passing migration plan, or a closed local-review receipt accepted directly by preview-template-migration; stale, missing, duplicate, or incompatible decisions fail without mutation.
                Use when: Progress reports no remaining source choices and any required baseline cleanup has been recorded.
                Do not use for: Supplying selectors, object ids, coordinates, values, operation payloads, bypassing unresolved choices, or calling close-template-migration-reviews on its result.
                Usage:
                  tiwater-docx resolve-template-migration-decisions <source.docx> <baseline.docx> <draft.json>
                """,
            "list-template-migration-choices" => """
                Purpose: Present unresolved source observations and selectable baseline targets as concise, current-document-bound choices.
                Consumes: One current source DOCX and one selected baseline DOCX.
                Produces: Opaque source and target choice ids with visible text and local context. It does not expose selectors or recommend business semantics.
                Use when: An existing integration consumes the complete choice catalog.
                Do not use for: New Agent-facing migration work, inventing business values, choosing a target automatically, building operations, or persisting choices across document revisions. New callers start with start-template-migration-decisions.
                Usage:
                  tiwater-docx list-template-migration-choices <source.docx> <baseline.docx>
                Compatibility next: Submit only the chosen ids and business dispositions to resolve-template-migration-choices.
                """,
            "resolve-template-migration-choices" => """
                Purpose: Expand current-document-bound semantic choices into the provider's validated migration plan.
                Consumes: The same current source DOCX and selected baseline DOCX plus a choice candidate returned from list-template-migration-choices.
                Produces: The existing hash-bound migration resolution and plan; unknown, stale, duplicate, or incompatible choices fail without mutation.
                Use when: An existing integration has already produced the complete compatibility candidate.
                Do not use for: New Agent-facing migration work, supplying selectors, object ids, coordinates, document values, or operation payloads. New callers use resolve-template-migration-decisions.
                Usage:
                  tiwater-docx resolve-template-migration-choices <source.docx> <baseline.docx> <choices.json>
                Candidate shape:
                  schema: tiwater.docx.template-migration-choice-candidate/v1
                  mappings[]: sourceChoiceId, targetChoiceId unless out-of-scope, disposition, optional cardinality
                  choiceSelections[]: sourceChoiceId, targetChoiceId
                  baselineClears[]: targetChoiceId, mode (cell or row)
                Allowed dispositions: copy-text, copy-media, retain-target, retain-target-label, out-of-scope.
                """,
            "derive-template-migration-anchor-gap-plan" => """
                Purpose: Legacy compatibility alias for the former mixed plan-and-candidate receipt.
                Consumes: One current source DOCX and one selected baseline DOCX.
                Produces: The legacy anchor-gap plan receipt.
                Use when: An existing caller still consumes the legacy receipt; new callers use start-template-migration-decisions.
                Do not use for: New tool discovery, approving mappings, editing, or output validation.
                Usage:
                  tiwater-docx derive-template-migration-anchor-gap-plan <source.docx> <baseline.docx>
                """,
            "build-template-migration-operations" => """
                Purpose: Deterministically build DOCX operations from one fully resolved, hash-bound migration plan.
                Consumes: The current source DOCX, selected baseline DOCX, and a plan returned by semantic resolution.
                Produces: An operation build receipt; any unresolved mapping is rejected with template-migration-semantic-resolution-required.
                Use when: resolve-template-migration-decisions, or a compatible choice/selector resolver, has returned Pass=true and no Unresolved items.
                Do not use for: Consuming an exact-text or anchor-gap plan that still has Unresolved items, selecting mappings, or editing.
                Usage:
                  tiwater-docx build-template-migration-operations <source.docx> <baseline.docx> <plan.json>
                """,
            "apply-template-migration" => """
                Purpose: Apply one fully resolved migration plan to the selected baseline.
                Consumes: The current source DOCX, selected baseline DOCX, passing resolved plan, and output path.
                Produces: An output DOCX and an apply receipt.
                Use when: Operation building has passed for the same source, baseline, and plan.
                Do not use for: Selecting mappings, resolving candidates, or validating the produced document.
                Usage:
                  tiwater-docx apply-template-migration <source.docx> <baseline.docx> <plan.json> <output.docx>
                """,
            "validate-template-migration-output" => """
                Purpose: Independently validate a migrated DOCX against the current source, selected baseline, and approved plan.
                Consumes: The current source DOCX, selected baseline DOCX, passing resolved plan, and produced output DOCX.
                Produces: A fresh readback verdict for content, structure, style, and migration coverage.
                Use when: Apply has completed and the output exists at a stable path.
                Do not use for: Building operations, editing, or accepting an unresolved plan.
                Usage:
                  tiwater-docx validate-template-migration-output <source.docx> <baseline.docx> <plan.json> <output.docx>
                """,
            _ => null,
        };
        if (help is null) return false;
        Console.WriteLine(help);
        return true;
    }

    private static Task<int> FailUnknown(string command)
    {
        Console.Error.WriteLine($"Unknown command: {command}");
        PrintUsage();
        return Task.FromResult(1);
    }

    private static void WriteJson<T>(T value)
    {
        Console.WriteLine(JsonSerializer.Serialize(value, Json.Options));
    }

    private static void RenderInspect(InspectionReport report)
    {
        Console.WriteLine($"File: {report.File}");
        Console.WriteLine($"Parts: {report.Package.PartCount}");
        Console.WriteLine($"Paragraphs: {report.Content.ParagraphCount}");
        Console.WriteLine($"Tables: {report.Content.TableCount}");
        Console.WriteLine($"Sections: {report.Content.SectionCount}");
        Console.WriteLine($"Headers: {report.Content.HeaderPartCount}");
        Console.WriteLine($"Footers: {report.Content.FooterPartCount}");
        Console.WriteLine($"Comments: {report.Annotations.CommentCount}");
        Console.WriteLine($"Footnotes: {report.Annotations.FootnoteCount}");
        Console.WriteLine($"Endnotes: {report.Annotations.EndnoteCount}");
        Console.WriteLine($"Tracked change elements: {report.Annotations.TrackedChangeElements}");
        Console.WriteLine($"Bookmarks: {report.Structure.BookmarkCount}");
        Console.WriteLine($"Hyperlinks: {report.Structure.HyperlinkCount}");
        Console.WriteLine($"Fields: {report.Structure.FieldCount}");
        Console.WriteLine($"Content controls: {report.Structure.ContentControlCount}");
        Console.WriteLine($"Drawings: {report.Structure.DrawingCount}");
        Console.WriteLine($"Annotation anchors: {report.Structure.AnnotationAnchors.Count}");
        Console.WriteLine($"Direct formatting paragraphs: {report.Formatting.ParagraphsWithDirectFormatting}");
        Console.WriteLine($"Direct formatting runs: {report.Formatting.RunsWithDirectFormatting}");

        Console.WriteLine("Paragraph styles in use:");
        foreach (var item in report.Styles.ParagraphStylesInUse)
        {
            Console.WriteLine($"  {item.Style}: {item.Count}");
        }

        if (report.Structure.AnnotationAnchors.Count > 0)
        {
            Console.WriteLine("Annotation anchors:");
            foreach (var anchor in report.Structure.AnnotationAnchors.Take(10))
            {
                Console.WriteLine($"  [{anchor.CommentId}] {anchor.TargetKind} {anchor.AnchorText}");
            }
        }
    }

    private static void RenderCompare(ComparisonReport report)
    {
        Console.WriteLine($"Old: {report.OldFile}");
        Console.WriteLine($"New: {report.NewFile}");
        Console.WriteLine($"Same parts: {report.PackageComparison.SamePartCount}");
        Console.WriteLine($"Different parts: {report.PackageComparison.DifferentPartCount}");

        if (report.PackageComparison.DifferentParts.Count > 0)
        {
            Console.WriteLine("Changed package parts:");
            foreach (var part in report.PackageComparison.DifferentParts)
            {
                Console.WriteLine($"  {part}");
            }
        }

        Console.WriteLine("Changed metrics:");
        foreach (var diff in report.MetricDiffs.Where(d => d.OldValue != d.NewValue))
        {
            Console.WriteLine($"  {diff.Name}: {diff.OldValue} -> {diff.NewValue}");
        }
    }

    private static void RenderInspectTables(TableInspectionReport report)
    {
        Console.WriteLine($"File: {report.File}");
        Console.WriteLine($"Tables: {report.Tables.Count}");
        foreach (var table in report.Tables)
        {
            Console.WriteLine($"Table {table.TableIndex}: {table.RowCount} row(s), {table.ColumnCount} column(s)");
            foreach (var row in table.Rows.Take(5))
            {
                var cells = row.Cells
                    .Take(5)
                    .Select(cell => $"[{cell.GridColumnStart}-{cell.GridColumnEnd} {cell.VMerge ?? "-"}] {cell.Text}")
                    .ToArray();
                Console.WriteLine($"  Row {row.RowIndex}: {string.Join(" | ", cells)}");
            }
        }
    }

    private static Task<int> RunValidateTemplateTransformAsync(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("validate-template-transform requires <source-template.docx> <target-template.docx>");
        }

        var source = args[0];
        var target = args[1];
        var json = args.Skip(2).Contains("--json", StringComparer.Ordinal);
        var report = TemplateTransformValidator.Validate(source, target);

        if (json)
        {
            WriteJson(report);
        }
        else
        {
            Console.WriteLine($"Source template: {report.SourceTemplate}");
            Console.WriteLine($"Target template: {report.TargetTemplate}");
            Console.WriteLine($"Compatible: {report.IsCompatible}");
        }

        return Task.FromResult(report.IsCompatible ? 0 : 2);
    }
}
