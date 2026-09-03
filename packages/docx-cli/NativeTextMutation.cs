using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTextMutation
{
    public const string Command = "docx_set_text";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetTextRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-text-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static SetTextReceipt Apply(SetTextRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var addresses = request.Changes.Select(change => change.Target).ToArray();
        var duplicate = request.Changes
            .Select((change, index) => new { change.Target, Index = index })
            .GroupBy(item => item.Target)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null)
            throw new InvalidOperationException(
                $"target-address-duplicate: changes=[{string.Join(',', duplicate.Select(item => item.Index))}]");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var resolved = Observation.ResolveAddresses(paths.Input, addresses, "changes.target");
        for (var index = 0; index < resolved.Count; index++)
            if (resolved[index].Kind is not "paragraph" and not "cell")
                throw new InvalidOperationException(
                    $"target-must-be-paragraph-or-cell: changes[{index}].target; kind={resolved[index].Kind}");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            var targets = resolved.Select(item =>
                Observation.ResolveNativePath(input, item.StoryPart, item.NativePath)).ToArray();
            NativeMutationSupport.RejectOverlappingTargets(targets);
            for (var index = 0; index < targets.Length; index++)
            {
                try
                {
                    NativeMutationSupport.RequirePlainTextContainer(targets[index]);
                }
                catch (InvalidOperationException error)
                {
                    throw new InvalidOperationException(
                        $"{error.Message}: changes[{index}].target", error);
                }
            }
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                for (var index = 0; index < resolved.Count; index++)
                {
                    var target = Observation.ResolveNativePath(output, resolved[index].StoryPart, resolved[index].NativePath);
                    SetText(target, request.Changes[index].Text);
                }
                SaveChangedStories(output, resolved.Select(item => item.StoryPart));
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            IReadOnlyList<SetTextReadback> readback;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                readback = resolved.Select(item =>
                {
                    var target = Observation.ResolveNativePath(output, item.StoryPart, item.NativePath);
                    var text = NativeMutationSupport.PlainText(target);
                    return new SetTextReadback(
                        item.Address,
                        item.Kind,
                        text);
                }).ToArray();
            }
            for (var index = 0; index < readback.Count; index++)
                if (!StringComparer.Ordinal.Equals(readback[index].Text, request.Changes[index].Text))
                    throw new InvalidOperationException("output-readback-content-mismatch");
            var receipt = new SetTextReceipt(
                "tiwater.docx-set-text-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                readback,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    internal static void SetText(OpenXmlElement target, string text)
    {
        switch (target)
        {
            case Paragraph paragraph:
                var runProperties = paragraph.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
                var insertionIndex = paragraph.ChildElements
                    .TakeWhile(child => child is not DocumentFormat.OpenXml.Wordprocessing.Run and not ProofError).Count();
                foreach (var child in paragraph.ChildElements
                             .Where(child => child is DocumentFormat.OpenXml.Wordprocessing.Run or ProofError).ToArray())
                    child.Remove();
                var replacementRun = TextRun(runProperties, text);
                if (replacementRun is not null)
                    paragraph.InsertAt(replacementRun, Math.Min(insertionIndex, paragraph.ChildElements.Count));
                break;
            case TableCell cell:
                var template = cell.Elements<Paragraph>().FirstOrDefault();
                var paragraphProperties = template?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
                var cellRunProperties = template?.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
                var bookmarkStarts = cell.Elements<Paragraph>()
                    .SelectMany(paragraph => paragraph.Elements<BookmarkStart>())
                    .Select(bookmark => (BookmarkStart)bookmark.CloneNode(true))
                    .ToArray();
                var bookmarkEnds = cell.Elements<Paragraph>()
                    .SelectMany(paragraph => paragraph.Elements<BookmarkEnd>())
                    .Select(bookmark => (BookmarkEnd)bookmark.CloneNode(true))
                    .ToArray();
                foreach (var child in cell.ChildElements.Where(child => child is not TableCellProperties).ToArray())
                    child.Remove();
                var replacement = new Paragraph();
                if (paragraphProperties is not null) replacement.Append((ParagraphProperties)paragraphProperties.CloneNode(true));
                foreach (var bookmark in bookmarkStarts) replacement.Append(bookmark);
                AppendText(replacement, cellRunProperties, text);
                foreach (var bookmark in bookmarkEnds) replacement.Append(bookmark);
                cell.Append(replacement);
                break;
            default:
                throw new InvalidOperationException("target-ref-must-be-paragraph-or-cell");
        }
    }

    private static void AppendText(OpenXmlCompositeElement parent, RunProperties? properties, string text)
    {
        var run = TextRun(properties, text);
        if (run is not null) parent.Append(run);
    }

    private static Run? TextRun(RunProperties? properties, string text)
    {
        if (text.Length == 0) return null;
        var run = new Run();
        if (properties is not null) run.Append((RunProperties)properties.CloneNode(true));
        var start = 0;
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character is not '\r' and not '\n' and not '\t') continue;
            if (index > start)
                run.Append(new Text(text[start..index]) { Space = SpaceProcessingModeValues.Preserve });
            if (character == '\t')
                run.Append(new TabChar());
            else
            {
                run.Append(new Break());
                if (character == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
            }
            start = index + 1;
        }
        if (start < text.Length)
            run.Append(new Text(text[start..]) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void SaveChangedStories(WordprocessingDocument document, IEnumerable<string> storyParts)
    {
        var parts = storyParts.Distinct(StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
        if (main.Document is not null && parts.Contains(PartUri(main.Uri))) main.Document.Save();
        foreach (var part in main.HeaderParts.Where(part => part.Header is not null && parts.Contains(PartUri(part.Uri)))) part.Header!.Save();
        foreach (var part in main.FooterParts.Where(part => part.Footer is not null && parts.Contains(PartUri(part.Uri)))) part.Footer!.Save();
        if (main.FootnotesPart?.Footnotes is not null && parts.Contains(PartUri(main.FootnotesPart.Uri))) main.FootnotesPart.Footnotes.Save();
        if (main.EndnotesPart?.Endnotes is not null && parts.Contains(PartUri(main.EndnotesPart.Uri))) main.EndnotesPart.Endnotes.Save();
        if (main.WordprocessingCommentsPart?.Comments is not null && parts.Contains(PartUri(main.WordprocessingCommentsPart.Uri))) main.WordprocessingCommentsPart.Comments.Save();
    }

    private static string PartUri(Uri uri)
        => uri.OriginalString.StartsWith("/", StringComparison.Ordinal) ? uri.OriginalString : "/" + uri.OriginalString;
}

public sealed record SetTextChange(DocxObjectAddress Target, string Text);
public sealed record SetTextRequest(string Input, IReadOnlyList<SetTextChange> Changes, string Output, string ReceiptOutput);
public sealed record SetTextReadback(DocxObjectAddress Address, string Kind, string Text);
public sealed record SetTextReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<SetTextReadback> Changes,
    string Output);
