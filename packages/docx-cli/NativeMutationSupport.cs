using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

internal static class NativeMutationSupport
{
    public sealed record PathsResult(string Input, string Output, string Receipt, bool InPlace);

    public static PathsResult Paths(
        string input,
        string output,
        string receiptOutput)
    {
        var inputPath = Path.GetFullPath(input);
        var outputPath = Path.GetFullPath(output);
        var receiptPath = Path.GetFullPath(receiptOutput);
        if (!File.Exists(inputPath) || Directory.Exists(inputPath))
            throw new InvalidOperationException("input-file-not-found");
        var inPlace = StringComparer.OrdinalIgnoreCase.Equals(inputPath, outputPath);
        if (!inPlace) RequireNewPath(outputPath, "output");
        RequireNewPath(receiptPath, "receiptOutput");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, receiptPath))
            throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
        if (!inPlace) Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        Directory.CreateDirectory(Path.GetDirectoryName(receiptPath)!);
        return new PathsResult(inputPath, outputPath, receiptPath, inPlace);
    }

    public static void Commit(string temporaryPath, PathsResult paths)
        => File.Move(temporaryPath, paths.Output, paths.InPlace);

    public static void CleanupFailure(string temporaryPath, PathsResult paths)
    {
        Cleanup(temporaryPath, paths.Receipt);
        if (!paths.InPlace) Cleanup(paths.Output);
    }

    public static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document)
            .GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    public static void RejectAddedValidationIssues(
        WordprocessingDocument document,
        IReadOnlyDictionary<string, int> baseline)
    {
        var added = ValidationIssueCounts(document)
            .FirstOrDefault(item => item.Value > baseline.GetValueOrDefault(item.Key));
        if (added.Key is not null)
            throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
    }

    public static ObjectArtifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new ObjectArtifact(
            Path.GetFullPath(path),
            Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(),
            stream.Length);
    }

    public static string JsonSha256<T>(T value)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(value, Json.CamelCaseOptions);
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    public static string ContentSha256(string value)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    public static string PlainText(OpenXmlElement target)
        => target is Text directText ? directText.Text : string.Concat(target.Descendants<OpenXmlElement>().Select(element => element switch
        {
            Text text => text.Text,
            Break => "\n",
            CarriageReturn => "\n",
            TabChar => "\t",
            _ => string.Empty,
        }));

    public static void RejectOverlappingTargets(IReadOnlyList<OpenXmlElement> targets)
    {
        for (var left = 0; left < targets.Count; left++)
        for (var right = left + 1; right < targets.Count; right++)
            if (targets[left].Ancestors().Contains(targets[right])
                || targets[right].Ancestors().Contains(targets[left]))
                throw new InvalidOperationException("target-refs-must-not-overlap");
    }

    public static void RequirePlainTextContainer(OpenXmlElement target)
    {
        if (target is TableCell cell)
        {
            if (cell.ChildElements.Any(child => child is not TableCellProperties and not Paragraph))
                throw new InvalidOperationException("target-cell-must-contain-only-plain-text-paragraphs");
            foreach (var cellParagraph in cell.Elements<Paragraph>()) RequirePlainTextParagraph(cellParagraph, true);
            return;
        }
        if (target is Paragraph paragraph)
        {
            RequirePlainTextParagraph(paragraph, true);
            return;
        }
        if (target is Text) return;
        throw new InvalidOperationException("target-ref-must-be-paragraph-cell-or-text");
    }

    private static void RequirePlainTextParagraph(Paragraph paragraph, bool allowBookmarks)
    {
        if (paragraph.ChildElements.Any(child => child is not ParagraphProperties and not Run and not ProofError
                && (!allowBookmarks || child is not BookmarkStart and not BookmarkEnd))
            || paragraph.Elements<Run>().Any(run =>
                run.ChildElements.Any(child => child is not RunProperties and not Text and not LastRenderedPageBreak
                    and not Break and not CarriageReturn and not TabChar)))
            throw new InvalidOperationException("target-paragraph-must-contain-only-plain-text-runs");
    }

    public static void Cleanup(params string[] paths)
    {
        foreach (var path in paths) if (File.Exists(path)) File.Delete(path);
    }

    private static void RequireNewPath(string path, string name)
    {
        if (File.Exists(path) || Directory.Exists(path))
            throw new InvalidOperationException($"{name}-already-exists");
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory))
            throw new InvalidOperationException($"{name}-directory-invalid");
    }
}
