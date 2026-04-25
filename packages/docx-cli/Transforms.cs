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

        using var doc = WordprocessingDocument.Open(input, false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document body not found.");

        var nodes = new List<object>();

        var paragraphIndex = 0;
        var tableIndex = 0;

        foreach (var element in body.ChildElements)
        {
            if (element is Paragraph p)
            {
                var text = string.Concat(p.Descendants<Text>().Select(t => t.Text)).Trim();
                var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (!string.IsNullOrEmpty(text))
                {
                    nodes.Add(new { Type = "paragraph", ParagraphIndex = paragraphIndex, Style = style, Text = text });
                }
                paragraphIndex += 1;
            }
            else if (element is Table t)
            {
                var tableData = new List<List<string>>();
                foreach (var row in t.Descendants<TableRow>())
                {
                    var rowData = new List<string>();
                    foreach (var cell in row.Descendants<TableCell>())
                    {
                        rowData.Add(string.Concat(cell.Descendants<Text>().Select(x => x.Text)).Trim());
                    }
                    tableData.Add(rowData);
                }
                nodes.Add(new { Type = "table", TableIndex = tableIndex, Rows = tableData });
                tableIndex += 1;
            }
        }

        var json = JsonSerializer.Serialize(nodes, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
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
        }
        return updated;
    }
}
