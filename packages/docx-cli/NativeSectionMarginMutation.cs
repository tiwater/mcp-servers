using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeSectionMarginMutation
{
    public const string Command = "docx_set_section_margins";
    private const int MaxTwips = 31680;

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetSectionMarginsRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-section-margins-request-invalid");
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

    public static SetSectionMarginsReceipt Apply(SetSectionMarginsRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var duplicate = request.Changes.GroupBy(change => change.SectionIndex).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) throw new InvalidOperationException($"section-index-duplicate: {duplicate.Key}");
        foreach (var (change, index) in request.Changes.Select((change, index) => (change, index)))
        {
            if (change.SectionIndex < 0) throw new InvalidOperationException($"section-index-invalid: changes[{index}]");
            if (change.Unit != "twip") throw new InvalidOperationException($"section-margin-unit-unsupported: changes[{index}].unit");
            var values = change.Margins.Values().ToArray();
            if (values.Length == 0) throw new InvalidOperationException($"section-margins-empty: changes[{index}]");
            if (values.Any(value => value < 0 || value > MaxTwips))
                throw new InvalidOperationException($"section-margin-out-of-range: changes[{index}]");
        }

        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
            RequireSectionIndexes(input, request.Changes);
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            List<SetSectionMarginsReadback> changes;
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var sections = Sections(output);
                changes = request.Changes.Select(change =>
                {
                    var margin = sections[change.SectionIndex].GetFirstChild<PageMargin>()
                        ?? throw new InvalidOperationException($"section-page-margin-missing: {change.SectionIndex}");
                    var before = Read(margin);
                    Apply(margin, change.Margins);
                    return new SetSectionMarginsReadback(change.SectionIndex, before, Read(margin));
                }).ToList();
                output.MainDocumentPart!.Document.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);

            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                var sections = Sections(output);
                foreach (var change in changes)
                    if (Read(sections[change.SectionIndex].GetFirstChild<PageMargin>()!) != change.After)
                        throw new InvalidOperationException("output-readback-section-margin-mismatch");
            }
            var receipt = new SetSectionMarginsReceipt(
                "tiwater.docx-set-section-margins-receipt/v1", "tiwater.docx.cli", RuntimeIdentity.Version, changes, paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static void RequireSectionIndexes(WordprocessingDocument document, IReadOnlyList<SetSectionMarginsChange> changes)
    {
        var count = Sections(document).Count;
        var invalid = changes.FirstOrDefault(change => change.SectionIndex >= count);
        if (invalid is not null) throw new InvalidOperationException($"section-index-out-of-range: {invalid.SectionIndex}; count={count}");
    }

    private static List<SectionProperties> Sections(WordprocessingDocument document)
        => document.MainDocumentPart?.Document.Body?.Descendants<SectionProperties>().ToList()
           ?? throw new InvalidOperationException("main-document-body-not-found");

    private static SectionMarginValues Read(PageMargin margin) => new(
        margin.Top?.Value ?? throw new InvalidOperationException("section-page-margin-top-missing"),
        margin.Bottom?.Value ?? throw new InvalidOperationException("section-page-margin-bottom-missing"),
        checked((int)(margin.Left?.Value ?? throw new InvalidOperationException("section-page-margin-left-missing"))),
        checked((int)(margin.Right?.Value ?? throw new InvalidOperationException("section-page-margin-right-missing"))),
        checked((int)(margin.Header?.Value ?? throw new InvalidOperationException("section-page-margin-header-missing"))),
        checked((int)(margin.Footer?.Value ?? throw new InvalidOperationException("section-page-margin-footer-missing"))));

    private static void Apply(PageMargin margin, SectionMarginPatch patch)
    {
        if (patch.Top is not null) margin.Top = patch.Top.Value;
        if (patch.Bottom is not null) margin.Bottom = patch.Bottom.Value;
        if (patch.Left is not null) margin.Left = (uint)patch.Left.Value;
        if (patch.Right is not null) margin.Right = (uint)patch.Right.Value;
        if (patch.Header is not null) margin.Header = (uint)patch.Header.Value;
        if (patch.Footer is not null) margin.Footer = (uint)patch.Footer.Value;
    }
}

public sealed record SectionMarginPatch(int? Top = null, int? Bottom = null, int? Left = null, int? Right = null, int? Header = null, int? Footer = null)
{
    public IEnumerable<int> Values()
    {
        if (Top is not null) yield return Top.Value;
        if (Bottom is not null) yield return Bottom.Value;
        if (Left is not null) yield return Left.Value;
        if (Right is not null) yield return Right.Value;
        if (Header is not null) yield return Header.Value;
        if (Footer is not null) yield return Footer.Value;
    }
}
public sealed record SectionMarginValues(int Top, int Bottom, int Left, int Right, int Header, int Footer);
public sealed record SetSectionMarginsChange(int SectionIndex, string Unit, SectionMarginPatch Margins);
public sealed record SetSectionMarginsRequest(string Input, IReadOnlyList<SetSectionMarginsChange> Changes, string Output, string ReceiptOutput);
public sealed record SetSectionMarginsReadback(int SectionIndex, SectionMarginValues Before, SectionMarginValues After);
public sealed record SetSectionMarginsReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<SetSectionMarginsReadback> Changes, string Output);
