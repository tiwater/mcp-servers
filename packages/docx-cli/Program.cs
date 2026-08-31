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
        "inspect",
        "compare",
        "validate-openxml",
        "strip-direct-formatting",
        "replace-style-ids",
        "export-json",
        "normalize-openxml",
        "validate-font-policy",
        "validate-toc-style-policy",
        .. ObservationCommand.Commands,
        NativeContentCopy.Command,
        NativeTextMutation.Command,
        NativeObjectMutation.InsertCommand,
        NativeObjectMutation.DeleteCommand,
        NativeCellMutation.MergeCommand,
        NativeCellMutation.SplitCommand,
        NativeTableColumnMutation.InsertCommand,
        NativeTableColumnMutation.DeleteCommand,
        NativePolicyMutation.FontCommand,
        NativePolicyMutation.TocCommand,
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
                "compare" => RunCompareAsync(args[1..]),
                "validate-openxml" => Task.FromResult(OpenXmlValidation.Run(args[1..])),
                "strip-direct-formatting" => Task.FromResult(Transforms.RunStripDirectFormatting(args[1..])),
                "replace-style-ids" => Task.FromResult(Transforms.RunReplaceStyleIds(args[1..])),
                "export-json" => Task.FromResult(Transforms.RunExportJson(args[1..])),
                "normalize-openxml" => Task.FromResult(DocxPackageNormalizer.RunNormalize(args[1..])),
                "validate-font-policy" => Task.FromResult(FontPolicy.RunValidate(args[1..])),
                "validate-toc-style-policy" => Task.FromResult(TocStylePolicy.RunValidate(args[1..])),
                _ when ObservationCommand.IsCommand(args[0]) => Task.FromResult(ObservationCommand.Run(args[0], args[1..])),
                _ when args[0] == NativeContentCopy.Command => Task.FromResult(NativeContentCopy.Run(args[1..])),
                _ when args[0] == NativeTextMutation.Command => Task.FromResult(NativeTextMutation.Run(args[1..])),
                _ when args[0] is NativeObjectMutation.InsertCommand or NativeObjectMutation.DeleteCommand
                    => Task.FromResult(NativeObjectMutation.Run(args[0], args[1..])),
                _ when args[0] is NativeCellMutation.MergeCommand or NativeCellMutation.SplitCommand
                    => Task.FromResult(NativeCellMutation.Run(args[0], args[1..])),
                _ when args[0] is NativeTableColumnMutation.InsertCommand or NativeTableColumnMutation.DeleteCommand
                    => Task.FromResult(NativeTableColumnMutation.Run(args[0], args[1..])),
                _ when args[0] is NativePolicyMutation.FontCommand or NativePolicyMutation.TocCommand
                    => Task.FromResult(NativePolicyMutation.Run(args[0], args[1..])),
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
                tables = Observation.TableIndex(input),
                flow = Inspector.InspectDocumentFlow(input),
                fonts = FontPolicy.Inspect(input),
            });
            return Task.FromResult(0);
        }
        var report = Inspector.Inspect(input);

        RenderInspect(report);

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
        Console.WriteLine("  docx_list_objects <request.json>");
        Console.WriteLine("  docx_find_literal <request.json>");
        Console.WriteLine("  docx_read_object <request.json>");
        Console.WriteLine("  compare <old.docx> <new.docx> [--json]");
        Console.WriteLine("  validate-openxml <input.docx>");
        Console.WriteLine("  strip-direct-formatting <input.docx> <output.docx>");
        Console.WriteLine("  replace-style-ids <input.docx> <output.docx> <style-map.json>");
        Console.WriteLine("  export-json <input.docx> [<output.json>]");
        Console.WriteLine("  normalize-openxml <input.docx> <output.docx>");
        Console.WriteLine("  validate-font-policy <input.docx> <policy.json>");
        Console.WriteLine("  validate-toc-style-policy <input.docx> <italic> <indent-characters-per-level>");
        Console.WriteLine("  docx_* <request.json>  (fixed published mutation commands)");
    }

    private static void PrintToolDescription()
    {
        Console.WriteLine("Purpose: Inspect, edit, normalize, and validate DOCX Open XML documents.");
        Console.WriteLine("Consumes: DOCX files and fixed technical arguments for the selected command.");
        Console.WriteLine("Produces: Technical observations, edited documents, and validation receipts.");
        Console.WriteLine("Do not use for: Business mappings, workflow decisions, delivery status, rendering, or OCR.");
        Console.WriteLine("Command discovery: Use --list-tools and each command's --help machine contract.");
        Console.WriteLine("Usage: tiwater-docx <command> [arguments]");
    }

    private static bool PrintCommandUsage(string command)
    {
        var usage = command switch
        {
            "inspect" => "tiwater-docx inspect <input.docx> [--json]",
            _ when ObservationCommand.IsCommand(command) => $"tiwater-docx {command} <request.json>",
            _ when command is NativeContentCopy.Command or NativeTextMutation.Command
                or NativeObjectMutation.InsertCommand or NativeObjectMutation.DeleteCommand
                or NativeCellMutation.MergeCommand or NativeCellMutation.SplitCommand
                or NativeTableColumnMutation.InsertCommand or NativeTableColumnMutation.DeleteCommand
                or NativePolicyMutation.FontCommand or NativePolicyMutation.TocCommand
                => $"tiwater-docx {command} <request.json>",
            "compare" => "tiwater-docx compare <old.docx> <new.docx> [--json]",
            "validate-openxml" => "tiwater-docx validate-openxml <input.docx>",
            "strip-direct-formatting" => "tiwater-docx strip-direct-formatting <input.docx> <output.docx>",
            "replace-style-ids" => "tiwater-docx replace-style-ids <input.docx> <output.docx> <style-map.json>",
            "export-json" => "tiwater-docx export-json <input.docx> [<output.json>]",
            "normalize-openxml" => "tiwater-docx normalize-openxml <input.docx> <output.docx>",
            "validate-font-policy" => "tiwater-docx validate-font-policy <input.docx> <policy.json>",
            "validate-toc-style-policy" => "tiwater-docx validate-toc-style-policy <input.docx> <italic> <indent-characters-per-level>",
            _ => null,
        };
        if (usage is null) return false;
        Console.WriteLine(usage);
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

}
