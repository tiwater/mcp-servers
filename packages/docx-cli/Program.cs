using System.Text.Json;
using Dockit.Docx;

namespace Dockit.Docx.Cli;

internal static class Program
{
    public static Task<int> Main(string[] args) => Cli.RunAsync(args);
}

internal static class Cli
{
    public static Task<int> RunAsync(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();
            return Task.FromResult(1);
        }

        try
        {
            return args[0] switch
            {
                "inspect" => RunInspectAsync(args[1..]),
                "plan" => Task.FromResult(Planner.RunPlan(args[1..])),
                "resolve" => Task.FromResult(Resolver.RunResolve(args[1..])),
                "compare" => RunCompareAsync(args[1..]),
                "validate-template-transform" => RunValidateTemplateTransformAsync(args[1..]),
                "strip-direct-formatting" => Task.FromResult(Transforms.RunStripDirectFormatting(args[1..])),
                "replace-style-ids" => Task.FromResult(Transforms.RunReplaceStyleIds(args[1..])),
                "export-json" => Task.FromResult(Transforms.RunExportJson(args[1..])),
                "fill-template" => Task.FromResult(Transforms.RunFillTemplate(args[1..])),
                "edit" => Task.FromResult(Editor.RunEdit(args[1..])),
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
        var report = Inspector.Inspect(input);

        if (json)
        {
            WriteJson(report);
        }
        else
        {
            RenderInspect(report);
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
        Console.WriteLine("  plan <input.docx> <plan-data.json>");
        Console.WriteLine("  resolve <plan.json> <resolve-data.json>");
        Console.WriteLine("  compare <old.docx> <new.docx> [--json]");
        Console.WriteLine("  validate-template-transform <source-template.docx> <target-template.docx> [--json]");
        Console.WriteLine("  strip-direct-formatting <input.docx> <output.docx>");
        Console.WriteLine("  replace-style-ids <input.docx> <output.docx> <style-map.json>");
        Console.WriteLine("  export-json <input.docx> [<output.json>]");
        Console.WriteLine("  fill-template <template.docx> <data.json> <output.docx>");
        Console.WriteLine("  edit <input.docx> <operations.json> <output.docx>");
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
