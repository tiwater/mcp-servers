using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeParagraphPaginationMutation
{
    public const string Command = "docx_set_paragraph_pagination";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetParagraphPaginationRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-paragraph-pagination-request-invalid");
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

    public static SetParagraphPaginationReceipt Apply(SetParagraphPaginationRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        if (request.Changes.Any(change => change.KeepWithNext is null
            && change.KeepLinesTogether is null
            && change.PageBreakBefore is null
            && change.PreventWidowOrphanLines is null))
            throw new InvalidOperationException("paragraph-pagination-change-must-set-a-property");
        var duplicate = request.Changes.GroupBy(change => change.Paragraph).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) throw new InvalidOperationException("paragraph-address-duplicate");

        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var resolved = Observation.ResolveAddresses(paths.Input, request.Changes.Select(change => change.Paragraph).ToArray(), "changes.paragraph");
        if (resolved.Any(item => item.Kind != "paragraph"))
            throw new InvalidOperationException("target-must-be-paragraph");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
            baseline = NativeMutationSupport.ValidationIssueCounts(input);

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(paths.Input, temporaryPath, false);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                for (var index = 0; index < resolved.Count; index++)
                {
                    var paragraph = Observation.ResolveNativePath(output, resolved[index].StoryPart, resolved[index].NativePath) as Paragraph
                        ?? throw new InvalidOperationException("output-paragraph-not-found");
                    ApplyChange(paragraph, request.Changes[index]);
                }
                SaveStories(output, resolved.Select(item => item.StoryPart));
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);

            IReadOnlyList<ParagraphPaginationReadback> readback;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                readback = resolved.Select(item =>
                {
                    var paragraph = Observation.ResolveNativePath(output, item.StoryPart, item.NativePath) as Paragraph
                        ?? throw new InvalidOperationException("readback-paragraph-not-found");
                    return Readback(item.Address, paragraph);
                }).ToArray();
            }
            for (var index = 0; index < request.Changes.Count; index++) Verify(request.Changes[index], readback[index]);
            var receipt = new SetParagraphPaginationReceipt(
                "tiwater.docx-set-paragraph-pagination-receipt/v1",
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

    private static void ApplyChange(Paragraph paragraph, SetParagraphPaginationChange change)
    {
        var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
        if (change.KeepWithNext is not null)
            properties.KeepNext = new KeepNext { Val = change.KeepWithNext.Value };
        if (change.KeepLinesTogether is not null)
            properties.KeepLines = new KeepLines { Val = change.KeepLinesTogether.Value };
        if (change.PageBreakBefore is not null)
            properties.PageBreakBefore = new PageBreakBefore { Val = change.PageBreakBefore.Value };
        if (change.PreventWidowOrphanLines is not null)
            properties.WidowControl = new WidowControl { Val = change.PreventWidowOrphanLines.Value };
    }

    private static ParagraphPaginationReadback Readback(DocxObjectAddress address, Paragraph paragraph)
    {
        var properties = paragraph.ParagraphProperties;
        return new ParagraphPaginationReadback(
            address,
            Value(properties?.GetFirstChild<KeepNext>()),
            Value(properties?.GetFirstChild<KeepLines>()),
            Value(properties?.GetFirstChild<PageBreakBefore>()),
            Value(properties?.GetFirstChild<WidowControl>()));
    }

    private static bool Value(OnOffType? property) => property is not null && (property.Val?.Value ?? true);

    private static void Verify(SetParagraphPaginationChange requested, ParagraphPaginationReadback actual)
    {
        if (requested.KeepWithNext is not null && requested.KeepWithNext != actual.KeepWithNext
            || requested.KeepLinesTogether is not null && requested.KeepLinesTogether != actual.KeepLinesTogether
            || requested.PageBreakBefore is not null && requested.PageBreakBefore != actual.PageBreakBefore
            || requested.PreventWidowOrphanLines is not null && requested.PreventWidowOrphanLines != actual.PreventWidowOrphanLines)
            throw new InvalidOperationException("output-readback-pagination-mismatch");
    }

    private static void SaveStories(WordprocessingDocument document, IEnumerable<string> storyParts)
    {
        var parts = storyParts.Distinct(StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
        static string Uri(OpenXmlPart part) => part.Uri.OriginalString.StartsWith('/') ? part.Uri.OriginalString : "/" + part.Uri.OriginalString;
        if (main.Document is not null && parts.Contains(Uri(main))) main.Document.Save();
        foreach (var part in main.HeaderParts.Where(part => part.Header is not null && parts.Contains(Uri(part)))) part.Header!.Save();
        foreach (var part in main.FooterParts.Where(part => part.Footer is not null && parts.Contains(Uri(part)))) part.Footer!.Save();
        if (main.FootnotesPart?.Footnotes is not null && parts.Contains(Uri(main.FootnotesPart))) main.FootnotesPart.Footnotes.Save();
        if (main.EndnotesPart?.Endnotes is not null && parts.Contains(Uri(main.EndnotesPart))) main.EndnotesPart.Endnotes.Save();
        if (main.WordprocessingCommentsPart?.Comments is not null && parts.Contains(Uri(main.WordprocessingCommentsPart))) main.WordprocessingCommentsPart.Comments.Save();
    }
}

public sealed record SetParagraphPaginationChange(
    DocxObjectAddress Paragraph,
    bool? KeepWithNext,
    bool? KeepLinesTogether,
    bool? PageBreakBefore,
    bool? PreventWidowOrphanLines);
public sealed record SetParagraphPaginationRequest(
    string Input,
    IReadOnlyList<SetParagraphPaginationChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record ParagraphPaginationReadback(
    DocxObjectAddress Paragraph,
    bool KeepWithNext,
    bool KeepLinesTogether,
    bool PageBreakBefore,
    bool PreventWidowOrphanLines);
public sealed record SetParagraphPaginationReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<ParagraphPaginationReadback> Changes,
    string Output);
