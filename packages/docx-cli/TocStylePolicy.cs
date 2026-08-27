using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class TocStylePolicy
{
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
            foreach (var style in styles.Elements<Style>().Where(style => style.Type?.Value == StyleValues.Paragraph))
            {
                var id = style.StyleId?.Value ?? string.Empty;
                var name = style.StyleName?.Val?.Value ?? string.Empty;
                var token = id.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) ? id[3..]
                    : name.StartsWith("toc ", StringComparison.OrdinalIgnoreCase) ? name[4..] : string.Empty;
                if (!int.TryParse(token, out var level) || level < 1) continue;
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
        if (matched == 0) findings.Add(new("", 0, "toc-styles", "at-least-one", "none"));
        return new("tiwater.docx-toc-style-validation/v1", RuntimeIdentity.Version, findings.Count == 0,
            input, italic, indentCharactersPerLevel, matched, findings);
    }

    private static bool? OnOff(OnOffType? value) => value is null ? null : value.Val?.Value ?? true;
}

public sealed record DocxTocStyleValidationReport(string Schema, string ToolVersion, bool Pass, string File,
    bool Italic, int IndentCharactersPerLevel, int MatchedStyleCount, IReadOnlyList<DocxTocStyleFinding> Findings);
public sealed record DocxTocStyleFinding(string StyleId, int Level, string Property, string Expected, string? Actual);
