using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeCommentMutation
{
    public const string Command = "docx_delete_comments";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<DeleteCommentsRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("delete-comments-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = 1, appliedCount = receipt.DeletedCommentCount },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static DeleteCommentsReceipt Apply(DeleteCommentsRequest request)
    {
        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        int commentCount;
        int markerCount;
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            var main = input.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
            commentCount = main.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Count() ?? 0;
            markerCount = StoryRoots(main).Sum(root => CommentMarkers(root).Count());
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var main = output.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
                foreach (var root in StoryRoots(main))
                {
                    foreach (var marker in CommentMarkers(root).ToArray()) marker.Remove();
                    root.Save();
                }
                if (main.WordprocessingCommentsPart is { } commentsPart) main.DeletePart(commentsPart);
                foreach (var part in main.Parts.Select(pair => pair.OpenXmlPart)
                    .Where(part => part.Uri.OriginalString.EndsWith("/commentsExtended.xml", StringComparison.OrdinalIgnoreCase)
                        || part.Uri.OriginalString.EndsWith("/people.xml", StringComparison.OrdinalIgnoreCase))
                    .ToArray())
                    main.DeletePart(part);
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);

            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                var main = output.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
                if (main.WordprocessingCommentsPart is not null
                    || StoryRoots(main).SelectMany(CommentMarkers).Any()
                    || main.Parts.Any(pair => pair.OpenXmlPart.Uri.OriginalString.EndsWith("/commentsExtended.xml", StringComparison.OrdinalIgnoreCase)
                        || pair.OpenXmlPart.Uri.OriginalString.EndsWith("/people.xml", StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("output-readback-comments-remain");
            }
            var receipt = new DeleteCommentsReceipt(
                "tiwater.docx-delete-comments-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                commentCount,
                markerCount,
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

    private static IEnumerable<OpenXmlPartRootElement> StoryRoots(MainDocumentPart main)
    {
        if (main.Document is not null) yield return main.Document;
        foreach (var part in main.HeaderParts) if (part.Header is not null) yield return part.Header;
        foreach (var part in main.FooterParts) if (part.Footer is not null) yield return part.Footer;
        if (main.FootnotesPart?.Footnotes is not null) yield return main.FootnotesPart.Footnotes;
        if (main.EndnotesPart?.Endnotes is not null) yield return main.EndnotesPart.Endnotes;
    }

    private static IEnumerable<OpenXmlElement> CommentMarkers(OpenXmlElement root)
        => root.Descendants().Where(element => element is CommentRangeStart or CommentRangeEnd or CommentReference);
}

public sealed record DeleteCommentsRequest(string Input, string Output, string ReceiptOutput);

public sealed record DeleteCommentsReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    int DeletedCommentCount,
    int DeletedMarkerCount,
    string Output);
