using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class TocStylePolicy
{
    public static int Apply(WordprocessingDocument document, bool italic, int indentCharactersPerLevel)
    {
        if (indentCharactersPerLevel < 0)
            throw new InvalidOperationException("indent-characters-per-level-must-be-nonnegative");
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles
            ?? throw new InvalidOperationException("document-styles-not-found");
        var entries = TocEntries(document, styles);
        var matched = 0;
        if (entries.Count == 0)
        {
            foreach (var style in TocStyles(styles))
            {
                var level = TocLevel(style);
                var paragraph = style.StyleParagraphProperties;
                if (paragraph is null)
                {
                    paragraph = new StyleParagraphProperties();
                    style.AddChild(paragraph, true);
                }
                paragraph.RemoveAllChildren<Indentation>();
                paragraph.AddChild(new Indentation { LeftChars = (level - 1) * indentCharactersPerLevel * 100 }, true);
                var run = style.StyleRunProperties;
                if (run is null)
                {
                    run = new StyleRunProperties();
                    style.AddChild(run, true);
                }
                run.RemoveAllChildren<Italic>();
                run.RemoveAllChildren<ItalicComplexScript>();
                run.AddChild(new Italic { Val = italic }, true);
                run.AddChild(new ItalicComplexScript { Val = italic }, true);
                matched++;
            }
        }
        else
        {
            foreach (var entry in entries)
            {
                var paragraph = entry.Paragraph.ParagraphProperties ?? entry.Paragraph.PrependChild(new ParagraphProperties());
                paragraph.RemoveAllChildren<Indentation>();
                paragraph.AddChild(new Indentation { LeftChars = (entry.Level - 1) * indentCharactersPerLevel * 100 }, true);
                foreach (var run in entry.Paragraph.Descendants<Run>())
                {
                    var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
                    properties.RemoveAllChildren<Italic>();
                    properties.RemoveAllChildren<ItalicComplexScript>();
                    properties.AddChild(new Italic { Val = italic }, true);
                    properties.AddChild(new ItalicComplexScript { Val = italic }, true);
                }
            }
            matched = entries.Select(entry => entry.Level).Distinct().Count();
        }
        if (matched == 0) throw new InvalidOperationException("toc-styles-not-found");
        return matched;
    }

    public static int RunValidate(string[] args)
    {
        if (args.Length != 3 || !bool.TryParse(args[1], out var italic)
            || !int.TryParse(args[2], out var indentCharactersPerLevel) || indentCharactersPerLevel < 0)
            throw new InvalidOperationException("validate-toc-style-policy requires <input.docx> <italic> <nonnegative-indent-characters-per-level>");
        var report = Validate(Path.GetFullPath(args[0]), italic, indentCharactersPerLevel);
        Console.WriteLine(JsonSerializer.Serialize(report, Json.Options));
        return report.Pass ? 0 : 1;
    }

    public static DocxTocStyleValidationReport Validate(string input, bool italic, int indentCharactersPerLevel)
    {
        using var document = WordprocessingDocument.Open(input, false);
        var styles = document.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        var findings = new List<DocxTocStyleFinding>();
        var matched = 0;
        if (styles is not null)
        {
            var entries = TocEntries(document, styles);
            if (entries.Count == 0)
            {
                foreach (var style in TocStyles(styles))
                {
                    var id = style.StyleId?.Value ?? string.Empty;
                    var level = TocLevel(style);
                    matched += 1;
                    var expectedIndent = (level - 1) * indentCharactersPerLevel * 100;
                    var actualIndent = style.StyleParagraphProperties?.GetFirstChild<Indentation>()?.LeftChars?.Value;
                    if (actualIndent != expectedIndent) findings.Add(new(id, level, "indent-characters", expectedIndent.ToString(), actualIndent?.ToString()));
                    var actualItalic = OnOff(style.StyleRunProperties?.GetFirstChild<Italic>());
                    if (actualItalic != italic) findings.Add(new(id, level, "italic", italic.ToString().ToLowerInvariant(), actualItalic?.ToString().ToLowerInvariant()));
                    var actualComplexItalic = OnOff(style.StyleRunProperties?.GetFirstChild<ItalicComplexScript>());
                    if (actualComplexItalic != italic) findings.Add(new(id, level, "italic-complex-script", italic.ToString().ToLowerInvariant(), actualComplexItalic?.ToString().ToLowerInvariant()));
                }
            }
            else
            {
                matched = entries.Select(entry => entry.Level).Distinct().Count();
                foreach (var entry in entries)
                {
                    var expectedIndent = (entry.Level - 1) * indentCharactersPerLevel * 100;
                    var actualIndent = entry.Paragraph.ParagraphProperties?.GetFirstChild<Indentation>()?.LeftChars?.Value;
                    if (actualIndent != expectedIndent)
                        findings.Add(new(entry.StyleId, entry.Level, "indent-characters", expectedIndent.ToString(), actualIndent?.ToString()));
                    foreach (var run in entry.Paragraph.Descendants<Run>().Where(run => !string.IsNullOrEmpty(run.InnerText)))
                    {
                        var directItalic = OnOff(run.RunProperties?.GetFirstChild<Italic>());
                        var directComplexItalic = OnOff(run.RunProperties?.GetFirstChild<ItalicComplexScript>());
                        if (directItalic != italic)
                            findings.Add(new(entry.StyleId, entry.Level, "direct-italic", italic.ToString().ToLowerInvariant(), directItalic?.ToString().ToLowerInvariant()));
                        if (directComplexItalic != italic)
                            findings.Add(new(entry.StyleId, entry.Level, "direct-italic-complex-script", italic.ToString().ToLowerInvariant(), directComplexItalic?.ToString().ToLowerInvariant()));
                    }
                }
            }
        }
        if (matched == 0) findings.Add(new("", 0, "toc-styles", "at-least-one", "none"));
        return new("tiwater.docx-toc-style-validation/v1", RuntimeIdentity.Version, findings.Count == 0,
            input, italic, indentCharactersPerLevel, matched, findings);
    }

    private static IEnumerable<Style> TocStyles(Styles styles)
        => styles.Elements<Style>()
            .Where(style => style.Type?.Value == StyleValues.Paragraph && TocLevel(style) >= 1);

    private static IReadOnlyList<TocEntry> TocEntries(WordprocessingDocument document, Styles styles)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null) return [];
        var bookmarks = body.Descendants<BookmarkStart>()
            .Where(start => start.Name?.Value?.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase) == true)
            .GroupBy(start => start.Name!.Value!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var entries = new List<TocEntry>();
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var names = paragraph.Descendants<FieldCode>()
                .SelectMany(code => Regex.Matches(code.Text ?? string.Empty, @"\b_Toc[^\s\""\\]+", RegexOptions.IgnoreCase)
                    .Select(match => match.Value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (names.Count == 0) continue;
            if (names.Count != 1 || !bookmarks.TryGetValue(names[0], out var starts) || starts.Count != 1)
                throw new InvalidOperationException("toc-entry-heading-binding-invalid");
            var heading = starts[0].Ancestors<Paragraph>().SingleOrDefault()
                ?? throw new InvalidOperationException("toc-entry-heading-not-found");
            var level = OutlineLevel(heading, styles);
            if (level < 1) continue;
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId))
                throw new InvalidOperationException("toc-entry-style-binding-invalid");
            entries.Add(new TocEntry(paragraph, styleId, level));
        }
        return entries;
    }

    private static int OutlineLevel(Paragraph paragraph, Styles styles)
    {
        var direct = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct is not null) return direct.Value + 1;
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId))
        {
            var style = styles.Elements<Style>().SingleOrDefault(candidate => candidate.StyleId?.Value == styleId);
            if (style is null) break;
            var outline = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (outline is not null) return outline.Value + 1;
            styleId = style.BasedOn?.Val?.Value;
        }
        return 0;
    }

    private static int TocLevel(Style style)
    {
        var id = style.StyleId?.Value ?? string.Empty;
        var name = style.StyleName?.Val?.Value ?? string.Empty;
        var token = id.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) ? id[3..]
            : name.StartsWith("toc ", StringComparison.OrdinalIgnoreCase) ? name[4..] : string.Empty;
        return int.TryParse(token, out var level) ? level : 0;
    }

    private static bool? OnOff(OnOffType? value) => value is null ? null : value.Val?.Value ?? true;

    private sealed record TocEntry(Paragraph Paragraph, string StyleId, int Level);
}

public sealed record DocxTocStyleValidationReport(string Schema, string ToolVersion, bool Pass, string File,
    bool Italic, int IndentCharactersPerLevel, int MatchedStyleCount, IReadOnlyList<DocxTocStyleFinding> Findings);
public sealed record DocxTocStyleFinding(string StyleId, int Level, string Property, string Expected, string? Actual);
