using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class Transforms
{
    public static int RunStripDirectFormatting(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("strip-direct-formatting requires <input.docx> <output.docx>");
        }

        var input = Path.GetFullPath(args[0]);
        var output = Path.GetFullPath(args[1]);
        File.Copy(input, output, overwrite: true);

        using var doc = WordprocessingDocument.Open(output, true);
        var body = doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found.");

        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            NormalizeParagraph(paragraph);
        }

        doc.MainDocumentPart!.Document.Save();
        Console.WriteLine(output);
        return 0;
    }

    public static int RunReplaceStyleIds(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("replace-style-ids requires <input.docx> <output.docx> <style-map.json>");
        }

        var input = Path.GetFullPath(args[0]);
        var output = Path.GetFullPath(args[1]);
        var mapPath = Path.GetFullPath(args[2]);
        var styleMap = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(mapPath))
            ?? throw new InvalidOperationException("Could not parse style map JSON.");

        File.Copy(input, output, overwrite: true);

        using var doc = WordprocessingDocument.Open(output, true);
        var body = doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found.");

        var changed = 0;
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var pStyle = paragraph.ParagraphProperties?.ParagraphStyleId;
            if (pStyle?.Val?.Value is string currentParagraphStyle && styleMap.TryGetValue(currentParagraphStyle, out var newParagraphStyle))
            {
                pStyle.Val = newParagraphStyle;
                changed++;
            }

            foreach (var runStyle in paragraph.Descendants<RunStyle>())
            {
                if (runStyle.Val?.Value is string currentRunStyle && styleMap.TryGetValue(currentRunStyle, out var newRunStyle))
                {
                    runStyle.Val = newRunStyle;
                    changed++;
                }
            }
        }

        doc.MainDocumentPart!.Document.Save();
        Console.WriteLine($"Updated {changed} style references in {output}");
        return 0;
    }

    private static void NormalizeParagraph(Paragraph paragraph)
    {
        if (paragraph.ParagraphProperties is { } pPr)
        {
            var keep = new List<OpenXmlElement>();

            if (pPr.ParagraphStyleId is not null)
            {
                keep.Add((OpenXmlElement)pPr.ParagraphStyleId.CloneNode(true));
            }

            if (pPr.NumberingProperties is not null)
            {
                keep.Add((OpenXmlElement)pPr.NumberingProperties.CloneNode(true));
            }

            if (pPr.SectionProperties is not null)
            {
                keep.Add((OpenXmlElement)pPr.SectionProperties.CloneNode(true));
            }

            paragraph.ParagraphProperties = new ParagraphProperties();
            foreach (var item in keep)
            {
                paragraph.ParagraphProperties.Append(item);
            }
        }

        foreach (var run in paragraph.Descendants<Run>())
        {
            if (run.RunProperties is { } rPr)
            {
                var keep = new List<OpenXmlElement>();
                if (rPr.RunStyle is not null)
                {
                    keep.Add((OpenXmlElement)rPr.RunStyle.CloneNode(true));
                }

                run.RunProperties = keep.Count == 0 ? null : new RunProperties(keep);
            }
        }
    }

    public static int RunExportJson(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("export-json requires <input.docx> [<output.json>]");
        }

        var input = Path.GetFullPath(args[0]);
        var output = args.Length > 1 ? Path.GetFullPath(args[1]) : null;

        var nodes = Inspector.InspectDocumentFlow(input);

        var json = JsonSerializer.Serialize(nodes, Json.CamelCaseOptions);
        if (output != null)
        {
            File.WriteAllText(output, json);
            Console.WriteLine(output);
        }
        else
        {
            Console.WriteLine(json);
        }

        return 0;
    }

    public class TemplateData 
    {
        public Dictionary<string, string>? CellValues { get; set; } = new();
        public Dictionary<string, string>? TableSlots { get; set; } = new();
        public Dictionary<string, List<Dictionary<string, string>>>? RowGroups { get; set; } = new();
        public List<string>? SelectedOptions { get; set; } = [];
        public List<string>? RemoveRowsContaining { get; set; } = [];
        public Dictionary<string, string>? UniqueTextValues { get; set; } = new();
    }

    public static int RunFillTemplate(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("fill-template requires <template.docx> <data.json> <output.docx>");
        }

        var template = Path.GetFullPath(args[0]);
        var dataJson = Path.GetFullPath(args[1]);
        var output = Path.GetFullPath(args[2]);

        var data = JsonSerializer.Deserialize<TemplateData>(File.ReadAllText(dataJson), new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) 
            ?? new TemplateData();

        File.Copy(template, output, overwrite: true);

        using var doc = WordprocessingDocument.Open(output, true);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document body not found.");

        if (data.RowGroups != null)
        {
            ExpandRowGroups(doc, data.RowGroups);
        }
        if (data.SelectedOptions != null)
        {
            MarkSelectedOptions(doc, data.SelectedOptions);
        }
        if (data.RemoveRowsContaining != null)
        {
            RemoveRowsContaining(doc, data.RemoveRowsContaining);
        }
        if (data.UniqueTextValues != null)
        {
            ReplaceUniqueTextValues(doc, data.UniqueTextValues);
        }

        if (data.CellValues != null)
        {
            ReplaceCellValuePlaceholders(doc, data.CellValues);
        }

        if (data.TableSlots != null)
        {
            var tables = body.Elements<Table>().ToList();
            foreach (var kvp in data.TableSlots)
            {
                try 
                {
                    var match = System.Text.RegularExpressions.Regex.Match(kvp.Key, @"table\[(\d+)\]\.row\[(\d+)\]\.cell\[(\d+)\]");
                    if (match.Success)
                    {
                        int tIdx = int.Parse(match.Groups[1].Value);
                        int rIdx = int.Parse(match.Groups[2].Value);
                        int cIdx = int.Parse(match.Groups[3].Value);

                        if (tIdx < tables.Count)
                        {
                            var rows = tables[tIdx].Elements<TableRow>().ToList();
                            if (rIdx < rows.Count)
                            {
                                var cells = rows[rIdx].Elements<TableCell>().ToList();
                                if (cIdx < cells.Count)
                                {
                                    var cell = cells[cIdx];
                                    cell.RemoveAllChildren<Paragraph>();
                                    cell.Append(new Paragraph(new Run(new Text(kvp.Value))));
                                }
                            }
                        }
                    }
                }
                catch 
                {
                    // Ignore invalid slot paths
                }
            }
        }

        doc.MainDocumentPart!.Document.Save();
        Console.WriteLine($"Filled template saved to {output}");
        return 0;
    }

    private static void ExpandRowGroups(
        WordprocessingDocument doc,
        IReadOnlyDictionary<string, List<Dictionary<string, string>>> groups)
    {
        foreach (var group in groups)
        {
            var marker = "{{" + group.Key + "}}";
            var matches = Inspector.GetRoots(doc)
                .SelectMany(root => root.Descendants<Table>())
                .SelectMany(table => table.Elements<TableRow>())
                .Where(row => row.InnerText.Contains(marker, StringComparison.Ordinal))
                .ToList();
            if (matches.Count == 0)
            {
                throw new InvalidOperationException($"fill-template-row-group-marker-missing:{group.Key}");
            }
            foreach (var markerRow in matches)
            {
                var table = markerRow.Parent as Table
                    ?? throw new InvalidOperationException($"fill-template-row-group-table-missing:{group.Key}");
                var rows = table.Elements<TableRow>().ToList();
                var markerIndex = rows.IndexOf(markerRow);
                if (markerIndex < 0 || markerIndex + 1 >= rows.Count)
                {
                    throw new InvalidOperationException($"fill-template-row-group-prototype-missing:{group.Key}");
                }
                var prototype = rows[markerIndex + 1];
                ReplaceText(markerRow, marker, string.Empty);
                foreach (var values in group.Value)
                {
                    var clone = (TableRow)prototype.CloneNode(true);
                    foreach (var value in values)
                    {
                        ReplaceText(clone, "{{" + value.Key + "}}", value.Value ?? string.Empty);
                        ReplaceText(clone, "[" + value.Key + "]", value.Value ?? string.Empty);
                    }
                    table.InsertBefore(clone, prototype);
                }
                prototype.Remove();
            }
        }
    }

    private static void ReplaceUniqueTextValues(
        WordprocessingDocument doc,
        IReadOnlyDictionary<string, string> values)
    {
        foreach (var value in values)
        {
            var matches = Inspector.GetRoots(doc)
                .SelectMany(root => root.Descendants<Paragraph>())
                .Where(paragraph => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text))
                    .Contains(value.Key, StringComparison.Ordinal))
                .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException($"fill-template-unique-text-match-count:{value.Key}:{matches.Count}");
            }
            ReplaceText(matches[0], value.Key, value.Value ?? string.Empty);
        }
    }

    private static void MarkSelectedOptions(WordprocessingDocument doc, IReadOnlyList<string> options)
    {
        foreach (var option in options)
        {
            var expected = NormalizeVisibleText(option);
            var matches = Inspector.GetRoots(doc)
                .SelectMany(root => root.Descendants<Paragraph>())
                .Where(paragraph => string.Equals(
                    NormalizeVisibleText(string.Concat(paragraph.Descendants<Text>().Select(text => text.Text))),
                    expected,
                    StringComparison.Ordinal))
                .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException($"fill-template-selected-option-match-count:{option}:{matches.Count}");
            }
            var paragraph = matches[0];
            var drawings = paragraph.Descendants<Drawing>().ToList();
            if (drawings.Count == 0)
            {
                throw new InvalidOperationException($"fill-template-selected-option-marker-missing:{option}");
            }
            foreach (var drawing in drawings) drawing.Remove();
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            var marker = new Run(new Text("☒ ") { Space = SpaceProcessingModeValues.Preserve });
            if (firstRun is null) paragraph.Append(marker);
            else paragraph.InsertBefore(marker, firstRun);
        }
    }

    private static void RemoveRowsContaining(WordprocessingDocument doc, IReadOnlyList<string> markers)
    {
        foreach (var marker in markers)
        {
            var expected = NormalizeVisibleText(marker);
            var matches = Inspector.GetRoots(doc)
                .SelectMany(root => root.Descendants<TableRow>())
                .Where(row => NormalizeVisibleText(row.InnerText).Contains(expected, StringComparison.Ordinal))
                .ToList();
            if (matches.Count != 1)
            {
                throw new InvalidOperationException($"fill-template-remove-row-match-count:{marker}:{matches.Count}");
            }
            matches[0].Remove();
        }
    }

    private static string NormalizeVisibleText(string value)
        => string.Concat((value ?? string.Empty).Where(character => !char.IsWhiteSpace(character)));

    private static void ReplaceText(OpenXmlElement root, string token, string value)
    {
        var paragraphs = root is Paragraph paragraphRoot
            ? [paragraphRoot]
            : root.Descendants<Paragraph>().ToList();
        foreach (var paragraph in paragraphs)
        {
            var texts = paragraph.Descendants<Text>().ToList();
            if (texts.Count == 0) continue;
            var combined = string.Concat(texts.Select(text => text.Text));
            var updated = combined.Replace(token, value, StringComparison.Ordinal);
            if (string.Equals(combined, updated, StringComparison.Ordinal)) continue;
            texts[0].Text = updated;
            texts[0].Space = SpaceProcessingModeValues.Preserve;
            foreach (var extra in texts.Skip(1)) extra.Text = string.Empty;
        }
    }

    private static int ReplaceCellValuePlaceholders(WordprocessingDocument doc, IReadOnlyDictionary<string, string> values)
    {
        var replacedParagraphs = 0;
        foreach (var root in Inspector.GetRoots(doc))
        {
            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                var texts = paragraph.Descendants<Text>().ToList();
                if (texts.Count == 0)
                {
                    continue;
                }

                var combined = string.Concat(texts.Select(text => text.Text));
                var updated = ApplyPlaceholderReplacements(combined, values);
                if (updated == combined)
                {
                    continue;
                }

                texts[0].Text = updated;
                texts[0].Space = SpaceProcessingModeValues.Preserve;
                foreach (var extra in texts.Skip(1))
                {
                    extra.Text = string.Empty;
                }

                replacedParagraphs += 1;
            }
        }

        return replacedParagraphs;
    }

    private static string ApplyPlaceholderReplacements(string input, IReadOnlyDictionary<string, string> values)
    {
        var updated = input;
        foreach (var kvp in values)
        {
            updated = updated.Replace("{{" + kvp.Key + "}}", kvp.Value ?? string.Empty, StringComparison.Ordinal);
            updated = updated.Replace("[" + kvp.Key + "]", kvp.Value ?? string.Empty, StringComparison.Ordinal);
        }
        return updated;
    }
}
