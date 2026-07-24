using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class FontPolicy
{
    public const string Schema = "tiwater.docx-font-policy/v1";
    private const string ReportSchema = "tiwater.docx-font-validation/v1";
    private static string ToolVersion => RuntimeIdentity.Version;

    public static int RunValidate(string[] args)
    {
        if (args.Length != 2) throw new InvalidOperationException("validate-font-policy requires <input.docx> <policy.json>");
        var input = Path.GetFullPath(args[0]);
        var policyBytes = File.ReadAllBytes(Path.GetFullPath(args[1]));
        var policy = ReadPolicy(policyBytes);
        if (!TryNormalize(policy, out var normalized, out var error)) throw new InvalidOperationException(error);
        var report = Validate(input, normalized, Convert.ToHexString(SHA256.HashData(policyBytes)).ToLowerInvariant());
        Console.WriteLine(JsonSerializer.Serialize(report, Json.Options));
        return report.Pass ? 0 : 1;
    }

    public static DocxFontPolicy ReadPolicy(byte[] bytes)
    {
        using var document = JsonDocument.Parse(bytes);
        Closed(document.RootElement, new[] { "schema", "body", "table" }, "font-policy");
        foreach (var name in new[] { "body", "table" }) Closed(document.RootElement.GetProperty(name), new[] { "eastAsia", "latin", "size" }, $"font-policy-{name}");
        var value = document.RootElement.Deserialize<DocxFontPolicy>(Json.Options) ?? throw new InvalidOperationException("font-policy-invalid");
        if (value.Schema != Schema) throw new InvalidOperationException("font-policy-schema-invalid");
        return value;
    }

    public static bool TryNormalize(DocxFontPolicy policy, out DocxFontPolicy normalized, out string? error)
    {
        normalized = policy;
        error = null;
        if (policy.Schema != Schema || !TryRule(policy.Body, out var body) || !TryRule(policy.Table, out var table))
        {
            error = "font-policy-rule-invalid";
            return false;
        }
        normalized = policy with { Body = body, Table = table };
        return true;
    }

    public static bool HasText(Run run) => run.Descendants<Text>().Any(text => !string.IsNullOrEmpty(text.Text));

    public static void Apply(Run run, DocxFontRule rule)
    {
        var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
        properties.RemoveAllChildren<RunFonts>();
        properties.PrependChild(new RunFonts { Ascii = rule.Latin, HighAnsi = rule.Latin, EastAsia = rule.EastAsia, ComplexScript = rule.Latin });
        if (rule.Size != "preserve")
        {
            properties.RemoveAllChildren<FontSize>();
            properties.RemoveAllChildren<FontSizeComplexScript>();
            properties.AppendChild(new FontSize { Val = rule.Size });
            properties.AppendChild(new FontSizeComplexScript { Val = rule.Size });
        }
    }

    public static DocxFontValidationReport Validate(string input, DocxFontPolicy policy, string policySha256)
    {
        using var document = WordprocessingDocument.Open(input, false);
        var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var findings = new List<DocxFontFinding>();
        var bodyOrdinal = 0;
        var tableOrdinal = 0;
        foreach (var run in body.Descendants<Run>().Where(HasText))
        {
            var inTable = run.Ancestors<Table>().Any();
            var ordinal = inTable ? tableOrdinal++ : bodyOrdinal++;
            var rule = inTable ? policy.Table : policy.Body;
            var properties = run.RunProperties;
            var fonts = properties?.RunFonts;
            var size = properties?.FontSize?.Val?.Value;
            var complexSize = properties?.FontSizeComplexScript?.Val?.Value;
            var sizeMatches = rule.Size == "preserve" || size == rule.Size && complexSize == rule.Size;
            if (fonts?.Ascii?.Value == rule.Latin && fonts.HighAnsi?.Value == rule.Latin && fonts.EastAsia?.Value == rule.EastAsia && fonts.ComplexScript?.Value == rule.Latin && sizeMatches) continue;
            findings.Add(new DocxFontFinding(inTable ? "table" : "body", ordinal, "font-policy-mismatch", fonts?.Ascii?.Value, fonts?.HighAnsi?.Value, fonts?.EastAsia?.Value, fonts?.ComplexScript?.Value, size, complexSize));
        }
        return new DocxFontValidationReport(ReportSchema, ToolVersion, findings.Count == 0, Path.GetFullPath(input), Sha256(input), policySha256, bodyOrdinal, tableOrdinal, findings);
    }

    public static DocxFontInspectionReport Inspect(string input)
    {
        using var document = WordprocessingDocument.Open(input, false);
        var body = document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var runs = new List<DocxFontRunObservation>();
        var bodyParagraphs = body.Descendants<Paragraph>().Where(paragraph => !paragraph.Ancestors<Table>().Any()).ToList();
        var tables = body.Descendants<Table>().ToList();
        var bodyOrdinal = 0;
        var tableOrdinal = 0;
        foreach (var run in body.Descendants<Run>())
        {
            var inTable = run.Ancestors<Table>().Any();
            var ordinal = inTable ? tableOrdinal++ : bodyOrdinal++;
            var paragraph = run.Ancestors<Paragraph>().First();
            var paragraphRuns = paragraph.Descendants<Run>().ToList();
            var runIndex = paragraphRuns.IndexOf(run);
            string container;
            if (inTable)
            {
                var table = run.Ancestors<Table>().First();
                var row = run.Ancestors<TableRow>().First();
                var cell = run.Ancestors<TableCell>().First();
                container = $"table:{tables.IndexOf(table)}:row:{table.Elements<TableRow>().ToList().IndexOf(row)}:cell:{row.Elements<TableCell>().ToList().IndexOf(cell)}:paragraph:{cell.Descendants<Paragraph>().ToList().IndexOf(paragraph)}";
            }
            else container = $"body:paragraph:{bodyParagraphs.IndexOf(paragraph)}";
            var properties = run.RunProperties;
            var fonts = properties?.RunFonts;
            var runText = string.Concat(run.Descendants<Text>().Select(value => value.Text));
            runs.Add(new DocxFontRunObservation(inTable ? "table" : "body", ordinal, container, runIndex, runText, HasText(run), fonts?.Ascii?.Value, fonts?.HighAnsi?.Value, fonts?.EastAsia?.Value, fonts?.ComplexScript?.Value, properties?.FontSize?.Val?.Value, properties?.FontSizeComplexScript?.Val?.Value));
        }
        return new DocxFontInspectionReport("tiwater.docx-font-inspection/v2", ToolVersion, bodyOrdinal, tableOrdinal, runs);
    }

    private static bool TryRule(DocxFontRule rule, out DocxFontRule normalized)
    {
        normalized = rule;
        var size = rule.Size?.Trim();
        if (string.Equals(size, "preserve", StringComparison.Ordinal))
        {
            if (string.IsNullOrWhiteSpace(rule.EastAsia) || string.IsNullOrWhiteSpace(rule.Latin)) return false;
            normalized = rule with { EastAsia = rule.EastAsia.Trim(), Latin = rule.Latin.Trim(), Size = "preserve" };
            return true;
        }
        if (string.IsNullOrWhiteSpace(rule.EastAsia) || string.IsNullOrWhiteSpace(rule.Latin) || size is null || !TryHalfPoints(size, out var halfPoints)) return false;
        normalized = rule with { EastAsia = rule.EastAsia.Trim(), Latin = rule.Latin.Trim(), Size = halfPoints };
        return true;
    }

    private static bool TryHalfPoints(string value, out string normalized)
    {
        normalized = string.Empty;
        var text = value?.Trim() ?? string.Empty;
        if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            if (!decimal.TryParse(text[..^2].Trim(), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var points) || points <= 0 || points * 2 != decimal.Truncate(points * 2)) return false;
            normalized = decimal.Truncate(points * 2).ToString("0", CultureInfo.InvariantCulture);
            return true;
        }
        if (!uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var halfPoints) || halfPoints == 0) return false;
        normalized = halfPoints.ToString(CultureInfo.InvariantCulture);
        return true;
    }

    private static void Closed(JsonElement value, IReadOnlyCollection<string> fields, string label)
    {
        if (value.ValueKind != JsonValueKind.Object) throw new InvalidOperationException($"{label}-object-invalid");
        foreach (var property in value.EnumerateObject()) if (!fields.Contains(property.Name)) throw new InvalidOperationException($"{label}-unknown-field:{property.Name}");
        foreach (var field in fields) if (!value.TryGetProperty(field, out _)) throw new InvalidOperationException($"{label}-missing-field:{field}");
    }

    private static string Sha256(string file) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file))).ToLowerInvariant();
}
