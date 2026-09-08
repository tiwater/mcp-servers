using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeSectionHeaderFooterReferenceCopy
{
    public const string Command = "docx_copy_section_header_footer_references";
    private static readonly string[] AllowedTypes = ["default", "even", "first"];
    private static readonly string[] AllowedStories = ["header", "footer"];

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<CopySectionHeaderFooterReferencesRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("copy-section-header-footer-references-request-invalid");
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

    public static CopySectionHeaderFooterReferencesReceipt Apply(CopySectionHeaderFooterReferencesRequest request)
    {
        ValidateRequest(request);
        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
            RequireSectionIndexes(input, request.Changes);
            foreach (var change in request.Changes) RequireSourceReferences(input, change);
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            List<CopySectionHeaderFooterReferencesReadback> changes;
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var sections = Sections(output);
                changes = request.Changes.Select(change =>
                {
                    var source = sections[change.SourceSectionIndex];
                    var target = sections[change.TargetSectionIndex];
                    var before = Read(target);
                    CopySelected<HeaderReference>(source, target,
                        change.References.Where(reference => reference.Story == "header").Select(reference => reference.Type).ToList(),
                        "header");
                    CopySelected<FooterReference>(source, target,
                        change.References.Where(reference => reference.Story == "footer").Select(reference => reference.Type).ToList(),
                        "footer");
                    return new CopySectionHeaderFooterReferencesReadback(
                        change.SourceSectionIndex,
                        change.TargetSectionIndex,
                        before,
                        Read(target));
                }).ToList();
                output.MainDocumentPart!.Document.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);

            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                var sections = Sections(output);
                foreach (var change in changes)
                    if (!Same(Read(sections[change.TargetSectionIndex]), change.After))
                        throw new InvalidOperationException("output-readback-section-header-footer-reference-mismatch");
            }
            var receipt = new CopySectionHeaderFooterReferencesReceipt(
                "tiwater.docx-copy-section-header-footer-references-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                changes,
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

    private static void ValidateRequest(CopySectionHeaderFooterReferencesRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var duplicate = request.Changes.GroupBy(change => change.TargetSectionIndex).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) throw new InvalidOperationException($"target-section-index-duplicate: {duplicate.Key}");
        foreach (var (change, index) in request.Changes.Select((change, index) => (change, index)))
        {
            if (change.SourceSectionIndex < 0 || change.TargetSectionIndex < 0)
                throw new InvalidOperationException($"section-index-invalid: changes[{index}]");
            if (change.SourceSectionIndex == change.TargetSectionIndex)
                throw new InvalidOperationException($"source-target-section-same: changes[{index}]");
            if (change.References is null || change.References.Count == 0)
                throw new InvalidOperationException($"section-header-footer-references-empty: changes[{index}]");
            var duplicateReference = change.References.GroupBy(
                reference => (reference.Story, reference.Type)).FirstOrDefault(group => group.Count() > 1);
            if (duplicateReference is not null)
                throw new InvalidOperationException($"section-header-footer-reference-duplicate: changes[{index}]");
            foreach (var reference in change.References)
            {
                if (!AllowedStories.Contains(reference.Story, StringComparer.Ordinal))
                    throw new InvalidOperationException($"section-header-footer-story-unsupported: changes[{index}]: {reference.Story}");
                if (!AllowedTypes.Contains(reference.Type, StringComparer.Ordinal))
                    throw new InvalidOperationException($"section-header-footer-type-unsupported: changes[{index}]: {reference.Type}");
            }
        }
    }

    private static void RequireSectionIndexes(
        WordprocessingDocument document,
        IReadOnlyList<CopySectionHeaderFooterReferencesChange> changes)
    {
        var count = Sections(document).Count;
        var invalid = changes.FirstOrDefault(change => change.SourceSectionIndex >= count || change.TargetSectionIndex >= count);
        if (invalid is not null)
            throw new InvalidOperationException(
                $"section-index-out-of-range: source={invalid.SourceSectionIndex}; target={invalid.TargetSectionIndex}; count={count}");
    }

    private static void RequireSourceReferences(
        WordprocessingDocument document,
        CopySectionHeaderFooterReferencesChange change)
    {
        var source = Sections(document)[change.SourceSectionIndex];
        foreach (var reference in change.References)
        {
            if (reference.Story == "header") RequireUniqueReference<HeaderReference>(source, reference.Type, reference.Story);
            else RequireUniqueReference<FooterReference>(source, reference.Type, reference.Story);
        }
    }

    private static T RequireUniqueReference<T>(SectionProperties section, string type, string story)
        where T : HeaderFooterReferenceType
    {
        var matches = section.Elements<T>().Where(reference => Type(reference.Type?.Value) == type).ToArray();
        if (matches.Length != 1)
            throw new InvalidOperationException($"source-section-{story}-reference-{(matches.Length == 0 ? "missing" : "ambiguous")}: {type}");
        return matches[0];
    }

    private static void CopySelected<T>(
        SectionProperties source,
        SectionProperties target,
        IReadOnlyList<string> types,
        string story)
        where T : HeaderFooterReferenceType
    {
        var requested = types;
        if (requested.Count == 0) return;
        var replacements = requested.ToDictionary(
            type => type,
            type => (T)RequireUniqueReference<T>(source, type, story).CloneNode(true),
            StringComparer.Ordinal);
        var rewritten = new List<T>();
        foreach (var existing in target.Elements<T>())
        {
            var type = Type(existing.Type?.Value);
            if (replacements.Remove(type, out var replacement)) rewritten.Add(replacement);
            else rewritten.Add((T)existing.CloneNode(true));
        }
        rewritten.AddRange(requested.Where(replacements.ContainsKey).Select(type => replacements[type]));
        target.RemoveAllChildren<T>();
        var insertionIndex = typeof(T) == typeof(HeaderReference)
            ? 0
            : target.Elements<HeaderReference>().Count();
        foreach (var reference in rewritten) target.InsertAt(reference, insertionIndex++);
    }

    private static List<SectionProperties> Sections(WordprocessingDocument document)
        => document.MainDocumentPart?.Document.Body?.Descendants<SectionProperties>().ToList()
           ?? throw new InvalidOperationException("main-document-body-not-found");

    private static string Type(HeaderFooterValues? type)
    {
        if (type is null) throw new InvalidOperationException("section-header-footer-reference-type-missing");
        if (type.Value == HeaderFooterValues.Default) return "default";
        if (type.Value == HeaderFooterValues.Even) return "even";
        if (type.Value == HeaderFooterValues.First) return "first";
        throw new InvalidOperationException("section-header-footer-reference-type-unsupported");
    }

    private static SectionHeaderFooterReferences Read(SectionProperties section) => new(
        HeaderReferences: section.Elements<HeaderReference>().Select(Reference).ToList(),
        FooterReferences: section.Elements<FooterReference>().Select(Reference).ToList());

    private static bool Same(SectionHeaderFooterReferences left, SectionHeaderFooterReferences right)
        => left.HeaderReferences.SequenceEqual(right.HeaderReferences)
           && left.FooterReferences.SequenceEqual(right.FooterReferences);

    private static SectionStoryReference Reference(HeaderFooterReferenceType reference) => new(
        Type: Type(reference.Type?.Value),
        RelationshipId: reference.Id?.Value
                        ?? throw new InvalidOperationException("section-header-footer-reference-id-missing"));
}

public sealed record CopySectionHeaderFooterReferencesChange(
    int SourceSectionIndex,
    int TargetSectionIndex,
    IReadOnlyList<SectionHeaderFooterReferenceSelection> References);
public sealed record SectionHeaderFooterReferenceSelection(string Story, string Type);
public sealed record CopySectionHeaderFooterReferencesRequest(
    string Input,
    IReadOnlyList<CopySectionHeaderFooterReferencesChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record SectionStoryReference(string Type, string RelationshipId);
public sealed record SectionHeaderFooterReferences(
    IReadOnlyList<SectionStoryReference> HeaderReferences,
    IReadOnlyList<SectionStoryReference> FooterReferences);
public sealed record CopySectionHeaderFooterReferencesReadback(
    int SourceSectionIndex,
    int TargetSectionIndex,
    SectionHeaderFooterReferences Before,
    SectionHeaderFooterReferences After);
public sealed record CopySectionHeaderFooterReferencesReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<CopySectionHeaderFooterReferencesReadback> Changes,
    string Output);
