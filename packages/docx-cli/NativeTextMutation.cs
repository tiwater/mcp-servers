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
        var refs = request.Changes.Select(change => change.TargetRef).ToArray();
        if (refs.Distinct(StringComparer.Ordinal).Count() != refs.Length)
            throw new InvalidOperationException("target-ref-duplicate");
        var paths = NativeMutationSupport.Paths(request.TargetDocument, request.Output, request.ReceiptOutput);
        var resolved = Observation.ResolveReferences(paths.Input, request.TargetDocument.Revision, refs);
        if (resolved.Any(item => item.Kind is not "paragraph" and not "cell"))
            throw new InvalidOperationException("target-ref-must-be-paragraph-or-cell");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            var targets = resolved.Select(item =>
                Observation.ResolveNativePath(input, item.StoryPart, item.NativePath)).ToArray();
            NativeMutationSupport.RejectOverlappingTargets(targets);
            foreach (var target in targets) NativeMutationSupport.RequirePlainTextContainer(target);
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(paths.Input, temporaryPath, false);
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
            File.Move(temporaryPath, paths.Output);
            var outputRevision = Observation.CurrentRevision(paths.Output);
            IReadOnlyList<SetTextReadback> readback;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                readback = resolved.Select(item =>
                {
                    var target = Observation.ResolveNativePath(output, item.StoryPart, item.NativePath);
                    var text = target.InnerText;
                    return new SetTextReadback(
                        Observation.MakeReference(outputRevision, item.Kind, item.StoryPart, item.NativePath),
                        item.Kind,
                        item.StoryPart,
                        item.NativePath,
                        text,
                        NativeMutationSupport.ContentSha256(target.OuterXml));
                }).ToArray();
            }
            for (var index = 0; index < readback.Count; index++)
                if (!StringComparer.Ordinal.Equals(readback[index].Text, request.Changes[index].Text))
                    throw new InvalidOperationException("output-readback-content-mismatch");
            var receipt = new SetTextReceipt(
                "tiwater.docx-set-text-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                Observation.CurrentRevision(paths.Input),
                outputRevision,
                readback,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.Cleanup(temporaryPath, paths.Output, paths.Receipt);
            throw;
        }
    }

    private static void SetText(OpenXmlElement target, string text)
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
                foreach (var child in cell.ChildElements.Where(child => child is not TableCellProperties).ToArray())
                    child.Remove();
                var replacement = new Paragraph();
                if (paragraphProperties is not null) replacement.Append((ParagraphProperties)paragraphProperties.CloneNode(true));
                AppendText(replacement, cellRunProperties, text);
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
        run.Append(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
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

public sealed record SetTextChange(string TargetRef, string Text);
public sealed record SetTextRequest(ObjectDocument TargetDocument, IReadOnlyList<SetTextChange> Changes, string Output, string ReceiptOutput);
public sealed record SetTextReadback(string Ref, string Kind, string StoryPart, string NativePath, string Text, string ContentSha256);
public sealed record SetTextReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    DocxRevision TargetRevision,
    DocxRevision OutputRevision,
    IReadOnlyList<SetTextReadback> Changes,
    string Output);
