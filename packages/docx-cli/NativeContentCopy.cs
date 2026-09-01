using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeContentCopy
{
    public const string Command = "docx_replace_content_from_source";
    private const string MainStory = "/word/document.xml";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<CopyContentRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("replace-content-from-source-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = Describe(request.ReceiptOutput),
            output = Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static CopyContentReceipt Apply(CopyContentRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var targetPath = paths.Input;
        var outputPath = paths.Output;
        var receiptPath = paths.Receipt;

        var targetRefs = Observation.ResolveAddresses(
            targetPath,
            request.Changes.Select(change => change.Target).ToArray(),
            "changes.target");
        if (targetRefs.Any(item => item.Kind is not "paragraph" and not "cell" || item.StoryPart != MainStory))
            throw new InvalidOperationException("target-address-must-be-main-document-paragraph-or-cell");
        if (targetRefs.Select(item => item.Address).Distinct().Count() != targetRefs.Count)
            throw new InvalidOperationException("target-address-duplicate");

        var prepared = PrepareSources(request, targetRefs);
        IReadOnlyDictionary<string, int> baselineIssues;
        using (var targetDocument = WordprocessingDocument.Open(targetPath, false))
        {
            var targets = targetRefs.Select(item =>
                Observation.ResolveNativePath(targetDocument, item.StoryPart, item.NativePath)).ToArray();
            NativeMutationSupport.RejectOverlappingTargets(targets);
            foreach (var target in targets) NativeMutationSupport.RequirePlainTextContainer(target);
            baselineIssues = ValidationIssueCounts(targetDocument);
            PreflightImports(targetDocument, prepared);
        }

        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(targetPath, temporaryPath, overwrite: false);
            using (var outputDocument = WordprocessingDocument.Open(temporaryPath, true))
            {
                ApplyChanges(outputDocument, prepared);
                outputDocument.MainDocumentPart?.Document?.Save();
                var issues = ValidationIssueCounts(outputDocument);
                var added = issues.FirstOrDefault(item => item.Value > baselineIssues.GetValueOrDefault(item.Key));
                if (added.Key is not null)
                    throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            var readback = ReadBack(outputPath, prepared);
            var receipt = new CopyContentReceipt(
                "tiwater.docx-replace-content-from-source-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                readback,
                outputPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static IReadOnlyList<PreparedChange> PrepareSources(
        CopyContentRequest request,
        IReadOnlyList<ResolvedDocxAddress> targetRefs)
    {
        var result = new List<PreparedChange>(request.Changes.Count);
        for (var index = 0; index < request.Changes.Count; index++)
        {
            var change = request.Changes[index];
            var sourcePath = Path.GetFullPath(change.SourceInput);
            var sourceRefs = Observation.ResolveAddresses(
                sourcePath,
                change.SourceSelections.Select(selection => selection.Address).ToArray(),
                $"changes[{index}].sourceSelections.address");
            if (sourceRefs.Any(item => item.StoryPart != MainStory))
                throw new InvalidOperationException("source-address-must-be-main-document-object");
            using var sourceDocument = WordprocessingDocument.Open(sourcePath, false);
            var paragraphs = new List<Paragraph>();
            for (var sourceIndex = 0; sourceIndex < sourceRefs.Count; sourceIndex++)
            {
                var resolved = sourceRefs[sourceIndex];
                var element = Observation.ResolveNativePath(sourceDocument, resolved.StoryPart, resolved.NativePath);
                paragraphs.AddRange(CopySelection(element, change.SourceSelections[sourceIndex]));
            }
            if (paragraphs.Count == 0) paragraphs.Add(new Paragraph());
            result.Add(new PreparedChange(targetRefs[index], sourcePath, paragraphs));
        }
        return result;
    }

    private static void PreflightImports(WordprocessingDocument targetDocument, IReadOnlyList<PreparedChange> changes)
    {
        var targetMain = targetDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        foreach (var group in changes.GroupBy(change => change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(change => change.Paragraphs).Cast<OpenXmlElement>().ToArray();
            foreach (var relationshipId in DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal))
                if (!DocxObjectActions.CanCopyRelationship(sourceMain, relationshipId, out var error))
                    throw new InvalidOperationException(error);
            if (!DocxObjectActions.TryImportStyles(sourceMain, targetMain, roots, apply: false, out var styleError))
                throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, targetMain, roots, apply: false, out var numberingError))
                throw new InvalidOperationException(numberingError);
        }
    }

    private static void ApplyChanges(WordprocessingDocument outputDocument, IReadOnlyList<PreparedChange> changes)
    {
        var outputMain = outputDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        var outputBody = outputMain.Document?.Body ?? throw new InvalidOperationException("target-body-not-found");
        var targets = changes.Select(change =>
            Observation.ResolveNativePath(outputDocument, change.Target.StoryPart, change.Target.NativePath)).ToArray();

        foreach (var group in changes.Select((change, index) => (change, index))
                     .GroupBy(item => item.change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(item => item.change.Paragraphs).Cast<OpenXmlElement>().ToArray();
            if (!DocxObjectActions.TryImportStyles(sourceMain, outputMain, roots, apply: true, out var styleError))
                throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, outputMain, roots, apply: true, out var numberingError))
                throw new InvalidOperationException(numberingError);
            var relationshipMap = DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal)
                .ToDictionary(id => id, id => DocxObjectActions.CopyRelationship(sourceMain, outputMain, id), StringComparer.Ordinal);
            foreach (var item in group)
            {
                var paragraphs = item.change.Paragraphs.Select(paragraph => (Paragraph)paragraph.CloneNode(true)).ToArray();
                foreach (var paragraph in paragraphs) DocxObjectActions.RewriteRelationships(paragraph, relationshipMap);
                DocxObjectActions.RemapDrawingIds(outputBody, paragraphs.Cast<OpenXmlElement>().ToArray());
                ReplaceTargetContent(targets[item.index], paragraphs);
            }
        }
    }

    internal static IReadOnlyList<Paragraph> CopySelection(OpenXmlElement element, CopyContentSelection selection)
    {
        if (selection.Range is not null)
            return [ParagraphForTextRange(element, selection.Range.Start, selection.Range.Length)];
        return element switch
        {
            TableCell cell => cell.Elements<Paragraph>().Select(CloneParagraph).ToArray(),
            Paragraph paragraph => [CloneParagraph(paragraph)],
            DocumentFormat.OpenXml.Wordprocessing.Run run => [new Paragraph((DocumentFormat.OpenXml.Wordprocessing.Run)run.CloneNode(true))],
            Text text => [new Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run((Text)text.CloneNode(true)))],
            _ => throw new InvalidOperationException($"source-ref-kind-not-supported-for-content-copy: {element.LocalName}"),
        };
    }

    internal static Paragraph CloneParagraph(Paragraph paragraph)
    {
        var clone = (Paragraph)paragraph.CloneNode(true);
        clone.ParagraphProperties?.Remove();
        foreach (var bookmark in clone.Descendants<BookmarkStart>().Cast<OpenXmlElement>()
                     .Concat(clone.Descendants<BookmarkEnd>()).ToArray())
            bookmark.Remove();
        foreach (var cached in clone.Descendants<ProofError>().Cast<OpenXmlElement>()
                     .Concat(clone.Descendants<LastRenderedPageBreak>()).ToArray())
            cached.Remove();
        return clone;
    }

    private static Paragraph ParagraphForTextRange(OpenXmlElement element, int start, int length)
    {
        if (element is not DocumentFormat.OpenXml.Wordprocessing.Run and not Text)
            throw new InvalidOperationException("text-range-requires-run-or-text-ref");
        if (element is DocumentFormat.OpenXml.Wordprocessing.Run run && run.ChildElements.Any(child => child is not RunProperties and not Text))
            throw new InvalidOperationException("text-range-run-must-contain-only-text");
        var value = element is Text text ? text.Text : string.Concat(element.Descendants<Text>().Select(item => item.Text));
        var selected = SliceScalars(value, start, length);
        var outputRun = element is DocumentFormat.OpenXml.Wordprocessing.Run sourceRun
            ? (DocumentFormat.OpenXml.Wordprocessing.Run)sourceRun.CloneNode(true)
            : new DocumentFormat.OpenXml.Wordprocessing.Run();
        outputRun.RemoveAllChildren<Text>();
        outputRun.Append(new Text(selected) { Space = SpaceProcessingModeValues.Preserve });
        return new Paragraph(outputRun);
    }

    private static string SliceScalars(string value, int start, int length)
    {
        var runes = value.EnumerateRunes().ToArray();
        if (start < 0 || length <= 0 || start > runes.Length - length)
            throw new InvalidOperationException("text-range-out-of-bounds");
        var builder = new StringBuilder();
        foreach (var rune in runes.Skip(start).Take(length)) builder.Append(rune.ToString());
        return builder.ToString();
    }

    internal static void ReplaceCellContent(TableCell target, IReadOnlyList<Paragraph> sourceParagraphs)
    {
        var template = target.Elements<Paragraph>().FirstOrDefault();
        var paragraphProperties = template?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var targetRunProperties = template?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        target.RemoveAllChildren<Paragraph>();
        foreach (var source in sourceParagraphs)
        {
            var paragraph = (Paragraph)source.CloneNode(true);
            paragraph.ParagraphProperties?.Remove();
            if (paragraphProperties is not null)
                paragraph.PrependChild((ParagraphProperties)paragraphProperties.CloneNode(true));
            foreach (var run in paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                run.RunProperties = MergeRunProperties(targetRunProperties, run.RunProperties);
            target.Append(paragraph);
        }
    }

    private static void ReplaceTargetContent(OpenXmlElement target, IReadOnlyList<Paragraph> sourceParagraphs)
    {
        if (target is TableCell cell)
        {
            ReplaceCellContent(cell, sourceParagraphs);
            return;
        }
        if (target is not Paragraph paragraph)
            throw new InvalidOperationException("target-ref-must-be-main-document-paragraph-or-cell");
        var targetRunProperties = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
            .FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        var insertionIndex = paragraph.ChildElements
            .TakeWhile(child => child is not DocumentFormat.OpenXml.Wordprocessing.Run and not ProofError).Count();
        foreach (var child in paragraph.ChildElements
                     .Where(child => child is DocumentFormat.OpenXml.Wordprocessing.Run or ProofError).ToArray())
            child.Remove();
        var nextIndex = Math.Min(insertionIndex, paragraph.ChildElements.Count);
        foreach (var source in sourceParagraphs)
        {
            foreach (var child in source.ChildElements.Where(child => child is not ParagraphProperties))
            {
                var clone = child.CloneNode(true);
                foreach (var run in clone.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                    run.RunProperties = MergeRunProperties(targetRunProperties, run.RunProperties);
                if (clone is DocumentFormat.OpenXml.Wordprocessing.Run directRun)
                    directRun.RunProperties = MergeRunProperties(targetRunProperties, directRun.RunProperties);
                paragraph.InsertAt(clone, nextIndex++);
            }
        }
    }

    private static RunProperties? MergeRunProperties(RunProperties? target, RunProperties? source)
    {
        var result = target?.CloneNode(true) as RunProperties ?? new RunProperties();
        if (source is not null)
        {
            foreach (var semantic in source.ChildElements.Where(element => element is Bold or Italic or Underline
                or Strike or DoubleStrike or Caps or SmallCaps or VerticalTextAlignment).ToArray())
            {
                SetRunProperty(result, semantic);
            }
        }
        return result.ChildElements.Count == 0 ? null : result;
    }

    private static void SetRunProperty(RunProperties properties, OpenXmlElement property)
    {
        switch (property)
        {
            case Bold value:
                properties.Bold = (Bold)value.CloneNode(true);
                break;
            case Italic value:
                properties.Italic = (Italic)value.CloneNode(true);
                break;
            case Underline value:
                properties.Underline = (Underline)value.CloneNode(true);
                break;
            case Strike value:
                properties.Strike = (Strike)value.CloneNode(true);
                break;
            case DoubleStrike value:
                properties.DoubleStrike = (DoubleStrike)value.CloneNode(true);
                break;
            case Caps value:
                properties.Caps = (Caps)value.CloneNode(true);
                break;
            case SmallCaps value:
                properties.SmallCaps = (SmallCaps)value.CloneNode(true);
                break;
            case VerticalTextAlignment value:
                properties.VerticalTextAlignment = (VerticalTextAlignment)value.CloneNode(true);
                break;
            default:
                throw new InvalidOperationException("run-property-not-supported-for-content-copy");
        }
    }

    private static IReadOnlyList<CopyContentReadback> ReadBack(
        string output,
        IReadOnlyList<PreparedChange> changes)
    {
        using var document = WordprocessingDocument.Open(output, false);
        return changes.Select(change =>
        {
            var target = Observation.ResolveNativePath(document, change.Target.StoryPart, change.Target.NativePath);
            var expected = string.Concat(change.Paragraphs.Select(paragraph => paragraph.InnerText));
            if (!StringComparer.Ordinal.Equals(target.InnerText, expected))
                throw new InvalidOperationException("output-readback-content-mismatch");
            return new CopyContentReadback(
                change.Target.Address,
                target.InnerText);
        }).ToArray();
    }

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document)
            .GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static CopyContentArtifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new CopyContentArtifact(
            Path.GetFullPath(path),
            Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(),
            stream.Length);
    }

    private sealed record PreparedChange(
        ResolvedDocxAddress Target,
        string SourcePath,
        IReadOnlyList<Paragraph> Paragraphs);
}

public sealed record CopyContentSelection(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    CopyContentRange? Range = null);
public sealed record CopyContentRange(int Start, int Length);
public sealed record CopyContentChange(
    DocxObjectAddress Target,
    string SourceInput,
    IReadOnlyList<CopyContentSelection> SourceSelections);
public sealed record CopyContentRequest(
    string Input,
    IReadOnlyList<CopyContentChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record CopyContentArtifact(string Path, string Sha256, long Bytes);
public sealed record CopyContentReadback(DocxObjectAddress Address, string Text);
public sealed record CopyContentReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<CopyContentReadback> Changes,
    string Output);
