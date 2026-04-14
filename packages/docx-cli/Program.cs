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
                "compare" => RunCompareAsync(args[1..]),
                "validate-template-transform" => RunValidateTemplateTransformAsync(args[1..]),
                "strip-direct-formatting" => Task.FromResult(Transforms.RunStripDirectFormatting(args[1..])),
                "replace-style-ids" => Task.FromResult(Transforms.RunReplaceStyleIds(args[1..])),
                "export-json" => Task.FromResult(Transforms.RunExportJson(args[1..])),
                "fill-template" => Task.FromResult(Transforms.RunFillTemplate(args[1..])),
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
        Console.WriteLine("  compare <old.docx> <new.docx> [--json]");
        Console.WriteLine("  validate-template-transform <source-template.docx> <target-template.docx> [--json]");
        Console.WriteLine("  strip-direct-formatting <input.docx> <output.docx>");
        Console.WriteLine("  replace-style-ids <input.docx> <output.docx> <style-map.json>");
        Console.WriteLine("  export-json <input.docx> [<output.json>]");
        Console.WriteLine("  fill-template <template.docx> <data.json> <output.docx>");
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
        Console.WriteLine($"Direct formatting paragraphs: {report.Formatting.ParagraphsWithDirectFormatting}");
        Console.WriteLine($"Direct formatting runs: {report.Formatting.RunsWithDirectFormatting}");

        Console.WriteLine("Paragraph styles in use:");
        foreach (var item in report.Styles.ParagraphStylesInUse)
        {
            Console.WriteLine($"  {item.Style}: {item.Count}");
        }

        Console.WriteLine("Run styles in use:");
        foreach (var item in report.Styles.RunStylesInUse)
        {
            Console.WriteLine($"  {item.Style}: {item.Count}");
        }

        Console.WriteLine("Headings:");
        foreach (var heading in report.Content.Headings)
        {
            Console.WriteLine($"  [{heading.Style}] {heading.Text}");
        }

        if (report.Content.Placeholders.Count > 0)
        {
            Console.WriteLine("Placeholders:");
            foreach (var placeholder in report.Content.Placeholders)
            {
                Console.WriteLine($"  {placeholder}");
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

        if (report.StyleDiffs.AddedParagraphStyles.Count > 0 || report.StyleDiffs.RemovedParagraphStyles.Count > 0)
        {
            Console.WriteLine("Paragraph style usage changes:");
            foreach (var item in report.StyleDiffs.AddedParagraphStyles)
            {
                Console.WriteLine($"  + {item.Style}: {item.Count}");
            }

            foreach (var item in report.StyleDiffs.RemovedParagraphStyles)
            {
                Console.WriteLine($"  - {item.Style}: {item.Count}");
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
            RenderTemplateValidation(report);
        }

        return Task.FromResult(report.IsCompatible ? 0 : 2);
    }

    private static void RenderTemplateValidation(TemplateTransformValidationReport report)
    {
        Console.WriteLine($"Source template: {report.SourceTemplate}");
        Console.WriteLine($"Target template: {report.TargetTemplate}");
        Console.WriteLine($"Compatible: {report.IsCompatible}");
        Console.WriteLine($"Source body field slots: {report.SourceBodyFieldSlotCount}");
        Console.WriteLine($"Target body field slots: {report.TargetBodyFieldSlotCount}");
        Console.WriteLine($"Source empty input slots: {report.SourceEmptyInputSlotCount}");
        Console.WriteLine($"Target empty input slots: {report.TargetEmptyInputSlotCount}");

        if (report.Errors.Count > 0)
        {
            Console.WriteLine("Errors:");
            foreach (var error in report.Errors)
            {
                Console.WriteLine($"  {error}");
            }
        }

        if (report.Warnings.Count > 0)
        {
            Console.WriteLine("Warnings:");
            foreach (var warning in report.Warnings)
            {
                Console.WriteLine($"  {warning}");
            }
        }

        if (report.MismatchedBodySlots.Count > 0)
        {
            Console.WriteLine("Mismatched body slots:");
            foreach (var slot in report.MismatchedBodySlots)
            {
                Console.WriteLine($"  {slot.Path}: '{slot.SourceText}' -> '{slot.TargetText}'");
            }
        }
    }
}
