using System.Globalization;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

// Native paragraph layout only. Caller selects addresses and desired layout;
// page-field instructions, run formatting and body content remain untouched.
public static class StoryLayout
{
    public const string InspectCommand = "docx_read_story_layout";
    public const string SetCommand = "docx_set_story_layout";
    public static int Run(string command, string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires one input");
        if (command == InspectCommand)
        {
            Console.WriteLine(JsonSerializer.Serialize(Inspect(args[0]), Json.CamelCaseOptions));
            return 0;
        }
        var request = JsonSerializer.Deserialize<StoryLayoutRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("story-layout-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new { tool = SetCommand, output = NativeMutationSupport.Describe(request.Output), receipt = NativeMutationSupport.Describe(request.ReceiptOutput), summary = new { pass = true, operationCount = receipt.Changes.Count } }, Json.CamelCaseOptions));
        return 0;
    }

    public static object Inspect(string input)
    {
        using var doc = WordprocessingDocument.Open(input, false);
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-missing");
        var sections = main.Document.Body!.Descendants<SectionProperties>().Select((s, i) => new {
            sectionIndex = i, pageWidthTwips = s.GetFirstChild<PageSize>()?.Width?.Value,
            marginLeftTwips = s.GetFirstChild<PageMargin>()?.Left?.Value,
            marginRightTwips = s.GetFirstChild<PageMargin>()?.Right?.Value,
            usableWidthTwips = UsableWidth(s),
            references = s.ChildElements.Where(x => x is HeaderReference or FooterReference).Select(x => new { story = x.LocalName == "headerReference" ? "header" : "footer", type = x.GetAttributes().FirstOrDefault(a => a.LocalName == "type").Value, part = main.GetPartById(x.GetAttributes().First(a => a.LocalName == "id").Value!).Uri.OriginalString }).ToArray()
        }).ToArray();
        var stories = Stories(main).Select(s => new {
            part = s.part.Uri.OriginalString,
            paragraphs = s.root.Descendants<Paragraph>().Where(p => !p.Ancestors<Table>().Any()).Select(p => new {
                address = new DocxObjectAddress(s.part.Uri.OriginalString, Observation.NativePathFor(p)),
                text = NativeMutationSupport.PlainText(p),
                fieldInstructions = p.Descendants<FieldCode>().Select(f => f.Text).Concat(p.Descendants<SimpleField>().Select(f => f.Instruction?.Value ?? "")).ToArray(),
                rightIndentTwips = p.ParagraphProperties?.Indentation?.Right?.Value,
                justification = p.ParagraphProperties?.Justification?.GetAttributes().FirstOrDefault(a => a.LocalName == "val").Value,
                leadingWhitespace = HasLeadingWhitespace(p)
            }).ToArray()
        }).ToArray();
        return new { schema = "tiwater.docx-story-layout/v1", input = NativeMutationSupport.Describe(input), sections, stories };
    }

    public static StoryLayoutReceipt Apply(StoryLayoutRequest request)
    {
        if (request.Changes.Count == 0 || request.Changes.Select(c => c.Paragraph).Distinct().Count() != request.Changes.Count)
            throw new InvalidOperationException("story-layout-changes-empty-or-duplicate");
        if (request.Changes.Any(c => !double.IsFinite(c.RightInsetFraction) || c.RightInsetFraction < 0 || c.RightInsetFraction > 0.2))
            throw new InvalidOperationException("story-layout-inset-invalid");
        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var inputArtifact = NativeMutationSupport.Describe(paths.Input);
        var resolved = Observation.ResolveAddresses(paths.Input, request.Changes.Select(c => c.Paragraph).ToArray(), "changes.paragraph");
        HashSet<string> storyParts;
        using (var preflight = WordprocessingDocument.Open(paths.Input, false))
            storyParts = Stories(preflight.MainDocumentPart!).Select(s => s.part.Uri.OriginalString).ToHashSet(StringComparer.Ordinal);
        if (resolved.Any(r => r.Kind != "paragraph" || !storyParts.Contains(r.StoryPart)))
            throw new InvalidOperationException("story-layout-target-must-be-header-or-footer-paragraph");
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false)) baseline = NativeMutationSupport.ValidationIssueCounts(input);
        var temporary = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporary);
            var changes = new List<StoryLayoutReadback>();
            using (var output = WordprocessingDocument.Open(temporary, true))
            {
                var main = output.MainDocumentPart!;
                var bodyXml = main.Document.Body!.OuterXml;
                var sections = main.Document.Body.Descendants<SectionProperties>().ToArray();
                for (var i = 0; i < resolved.Count; i++)
                {
                    var r = resolved[i]; var c = request.Changes[i];
                    var p = (Paragraph)Observation.ResolveNativePath(output, r.StoryPart, r.NativePath);
                    if (p.Ancestors<Table>().Any()) throw new InvalidOperationException("story-layout-table-paragraph-not-supported");
                    var fieldCodes = Instructions(p);
                    var before = NativeMutationSupport.PlainText(p);
                    if (before != c.ExpectedText) throw new InvalidOperationException("story-layout-expected-text-mismatch");
                    var widths = sections.Where(s => s.ChildElements.Where(x => x is HeaderReference or FooterReference)
                        .Any(x => main.GetPartById(x.GetAttributes().First(a => a.LocalName == "id").Value!).Uri.OriginalString == r.StoryPart)).Select(UsableWidth).ToArray();
                    // A later section may inherit a story. Use the narrowest section
                    // as the safe bound rather than guessing which pages consume it.
                    widths = widths.Concat(sections.Select(UsableWidth)).ToArray();
                    if (widths.Length == 0 || widths.Any(w => w <= 0)) throw new InvalidOperationException("story-layout-page-width-unavailable");
                    var inset = (int)Math.Ceiling(widths.Min() * c.RightInsetFraction);
                    var properties = p.ParagraphProperties ?? p.PrependChild(new ParagraphProperties());
                    var indent = properties.Indentation;
                    if (indent is null) { indent = new Indentation(); properties.AddChild(indent, true); }
                    indent.Right = inset.ToString(CultureInfo.InvariantCulture);
                    indent.RightChars = null;
                    properties.Justification = new Justification { Val = JustificationValues.Right };
                    if (c.TrimLeadingWhitespace) TrimLeadingWhitespace(p);
                    if (Instructions(p) != fieldCodes) throw new InvalidOperationException("story-layout-field-instructions-changed");
                    var after = NativeMutationSupport.PlainText(p);
                    if (after != (c.TrimLeadingWhitespace ? before.TrimStart() : before)) throw new InvalidOperationException("story-layout-content-changed");
                    changes.Add(new StoryLayoutReadback(c.Paragraph, before, after, inset, "right", fieldCodes));
                }
                if (main.Document.Body.OuterXml != bodyXml) throw new InvalidOperationException("story-layout-body-changed");
                foreach (var story in Stories(main)) story.root.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporary, paths);
            using (var output = WordprocessingDocument.Open(paths.Output, false))
                foreach (var c in changes)
                {
                    var p = (Paragraph)Observation.ResolveNativePath(output, c.Paragraph.Part, c.Paragraph.Path);
                    if (NativeMutationSupport.PlainText(p) != c.AfterText || Instructions(p) != c.FieldInstructions
                        || p.ParagraphProperties?.Indentation?.Right?.Value != c.RightIndentTwips.ToString(CultureInfo.InvariantCulture)
                        || p.ParagraphProperties?.Justification?.Val?.Value != JustificationValues.Right)
                        throw new InvalidOperationException("story-layout-fresh-readback-mismatch");
                }
            var receipt = new StoryLayoutReceipt("tiwater.docx-story-layout-receipt/v1", RuntimeIdentity.Version, inputArtifact, NativeMutationSupport.Describe(paths.Output), changes);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch { NativeMutationSupport.CleanupFailure(temporary, paths); throw; }
    }

    private static long UsableWidth(SectionProperties s)
    {
        var width = s.GetFirstChild<PageSize>()?.Width?.Value;
        var m = s.GetFirstChild<PageMargin>();
        if (width is null || m?.Left is null || m.Right is null) throw new InvalidOperationException("story-layout-section-dimensions-missing");
        return (long)width.Value - m.Left.Value - m.Right.Value - (long)(m.Gutter?.Value ?? 0);
    }
    private static IEnumerable<(OpenXmlPart part, OpenXmlPartRootElement root)> Stories(MainDocumentPart main)
        => main.HeaderParts.Select(p => ((OpenXmlPart)p, (OpenXmlPartRootElement)p.Header!)).Concat(main.FooterParts.Select(p => ((OpenXmlPart)p, (OpenXmlPartRootElement)p.Footer!)));
    private static string Instructions(Paragraph p) => JsonSerializer.Serialize(p.Descendants<FieldCode>().Select(f => f.Text).Concat(p.Descendants<SimpleField>().Select(f => f.Instruction?.Value ?? "")));
    private static bool HasLeadingWhitespace(Paragraph p) => NativeMutationSupport.PlainText(p) is { Length: > 0 } t && char.IsWhiteSpace(t[0]);
    private static void TrimLeadingWhitespace(Paragraph p)
    {
        foreach (var text in p.Descendants<Text>())
        {
            text.Text = text.Text.TrimStart();
            if (text.Text.Length > 0) break;
        }
    }
}
public sealed record StoryLayoutChange(DocxObjectAddress Paragraph, string ExpectedText, double RightInsetFraction, bool TrimLeadingWhitespace);
public sealed record StoryLayoutRequest(string Input, string Output, string ReceiptOutput, IReadOnlyList<StoryLayoutChange> Changes);
public sealed record StoryLayoutReadback(DocxObjectAddress Paragraph, string BeforeText, string AfterText, int RightIndentTwips, string Justification, string FieldInstructions);
public sealed record StoryLayoutReceipt(string Schema, string Version, ObjectArtifact Input, ObjectArtifact Output, IReadOnlyList<StoryLayoutReadback> Changes);
