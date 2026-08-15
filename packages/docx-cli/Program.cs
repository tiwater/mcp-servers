using System.Text.Json;
using Dockit.Docx;

namespace Dockit.Docx.Cli;

internal static class Program
{
    public static Task<int> Main(string[] args) => Cli.RunAsync(args);
}

public static class Cli
{
    public static Task<int> RunAsync(string[] args)
    {
        if (args.Length == 1 && args[0] == "--describe-tool")
        {
            PrintToolDescription();
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
                "derive-template-migration-anchor-gap-plan" => Task.FromResult(TemplateMigration.RunDeriveAnchorGapPlan(args[1..])),
                "resolve-template-migration-semantic-candidate" => Task.FromResult(TemplateMigration.RunResolveSemanticCandidate(args[1..])),
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
        Console.WriteLine("  analyze-template-migration <source.docx> <baseline.docx> [--json]");
        Console.WriteLine("  derive-template-migration-exact-text-plan <source.docx> <baseline.docx>");
        Console.WriteLine("  derive-template-migration-anchor-gap-plan <source.docx> <baseline.docx>");
        Console.WriteLine("  resolve-template-migration-semantic-candidate <source.docx> <baseline.docx> <candidate.json>  (append --help for candidate shape)");
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
        Console.WriteLine("Command discovery: Run tiwater-docx with no arguments; run a listed template-migration command with --help for its exact contract.");
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
                Use when: Starting a template migration before selecting any semantic mapping.
                Do not use for: Selecting mappings, building operations, editing a document, or validating output.
                Usage:
                  tiwater-docx analyze-template-migration <source.docx> <baseline.docx> [--json]
                """,
            "derive-template-migration-exact-text-plan" => """
                Purpose: Derive the conservative automatic portion of a template-migration plan from unique current text and topology.
                Consumes: One current source DOCX and the same selected baseline DOCX used for analysis.
                Produces: A hash-bound plan; Unresolved[].Source and Unresolved[].BaselineOptions carry current observations, and UnclaimedBaseline lists unclaimed baseline content and selectable child runs.
                Use when: The source and baseline have been observed and automatic exact matches are needed.
                Do not use for: Treating Unresolved as review-required, inventing target mappings, editing, or output validation.
                Usage:
                  tiwater-docx derive-template-migration-exact-text-plan <source.docx> <baseline.docx>
                """,
            "derive-template-migration-anchor-gap-plan" => """
                Purpose: Add conservative paragraph candidates found between reciprocal exact-text anchors.
                Consumes: One current source DOCX and the same selected baseline DOCX used by the exact-text plan.
                Produces: A hash-bound plan; Unresolved[].Source and Unresolved[].Baseline carry the current anchor-gap observations, and UnclaimedBaseline remains available for semantic resolution.
                Use when: Exact matching leaves paragraph gaps that may have a unique current semantic target.
                Do not use for: Approving a candidate, building operations, editing, or output validation.
                Usage:
                  tiwater-docx derive-template-migration-anchor-gap-plan <source.docx> <baseline.docx>
                """,
            "build-template-migration-operations" => """
                Purpose: Deterministically build DOCX operations from one fully resolved, hash-bound migration plan.
                Consumes: The current source DOCX, selected baseline DOCX, and a plan returned by semantic resolution.
                Produces: An operation build receipt; any unresolved mapping is rejected with template-migration-semantic-resolution-required.
                Use when: resolve-template-migration-semantic-candidate has returned Pass=true and no Unresolved items.
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
