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

}
