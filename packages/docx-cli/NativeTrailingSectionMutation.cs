using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTrailingSectionMutation
{
    public const string Command = "docx_collapse_trailing_empty_section";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<CollapseTrailingEmptySectionRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("collapse-trailing-empty-section-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = 1, appliedCount = 1 },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static CollapseTrailingEmptySectionReceipt Apply(CollapseTrailingEmptySectionRequest request)
    {
        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        IReadOnlyDictionary<string, int> baseline;
        int sectionsBefore;
        string visibleText;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
            var body = input.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("main-document-body-not-found");
            RequireTrailingEmptySection(body);
            sectionsBefore = body.Descendants<SectionProperties>().Count();
            visibleText = body.InnerText;
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            int removedParagraphs;
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var body = output.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("main-document-body-not-found");
                var state = RequireTrailingEmptySection(body);
                var promoted = (SectionProperties)state.Boundary.CloneNode(true);
                state.Boundary.Remove();
                foreach (var paragraph in state.TrailingParagraphs) paragraph.Remove();
                state.Final.Parent!.ReplaceChild(promoted, state.Final);
                removedParagraphs = state.TrailingParagraphs.Count;
                output.MainDocumentPart!.Document.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);

            int sectionsAfter;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                var body = output.MainDocumentPart?.Document.Body ?? throw new InvalidOperationException("main-document-body-not-found");
                sectionsAfter = body.Descendants<SectionProperties>().Count();
                if (sectionsAfter != sectionsBefore - 1) throw new InvalidOperationException("output-readback-section-count-mismatch");
                if (HasTrailingEmptySection(body)) throw new InvalidOperationException("output-readback-trailing-empty-section-remains");
                if (!StringComparer.Ordinal.Equals(body.InnerText, visibleText)) throw new InvalidOperationException("output-readback-visible-content-changed");
            }
            var receipt = new CollapseTrailingEmptySectionReceipt(
                "tiwater.docx-collapse-trailing-empty-section-receipt/v1", "tiwater.docx.cli", RuntimeIdentity.Version,
                sectionsBefore, sectionsAfter, removedParagraphs, false, paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static TrailingSectionState RequireTrailingEmptySection(Body body)
    {
        var children = body.ChildElements.ToList();
        if (children.LastOrDefault() is not SectionProperties final)
            throw new InvalidOperationException("trailing-empty-section-not-found");
        for (var index = children.Count - 2; index >= 0; index--)
        {
            if (children[index] is not Paragraph paragraph || paragraph.ParagraphProperties?.SectionProperties is not { } boundary) continue;
            var trailing = children.Skip(index + 1).Take(children.Count - index - 2)
                .Where(child => child is not BookmarkStart and not BookmarkEnd).ToArray();
            if (trailing.Length == 0 || trailing.Any(child => child is not Paragraph item || !IsEmpty(item)))
                throw new InvalidOperationException("trailing-empty-section-not-found");
            return new TrailingSectionState(boundary, final, trailing.Cast<Paragraph>().ToArray());
        }
        throw new InvalidOperationException("trailing-empty-section-not-found");
    }

    private static bool HasTrailingEmptySection(Body body)
    {
        try { RequireTrailingEmptySection(body); return true; }
        catch (InvalidOperationException error) when (error.Message == "trailing-empty-section-not-found") { return false; }
    }

    private static bool IsEmpty(Paragraph paragraph)
        => string.IsNullOrWhiteSpace(paragraph.InnerText)
           && !paragraph.Descendants().Any(element => element is Drawing or Break or TabChar or CarriageReturn
               or FieldChar or FieldCode or FootnoteReference or EndnoteReference or CommentReference
               or BookmarkStart or BookmarkEnd or Hyperlink);

    private sealed record TrailingSectionState(SectionProperties Boundary, SectionProperties Final, IReadOnlyList<Paragraph> TrailingParagraphs);
}

public sealed record CollapseTrailingEmptySectionRequest(string Input, string Output, string ReceiptOutput);
public sealed record CollapseTrailingEmptySectionReceipt(
    string Schema, string Provider, string ToolVersion, int SectionCountBefore, int SectionCountAfter,
    int RemovedTrailingParagraphCount, bool HasTrailingEmptySection, string Output);
