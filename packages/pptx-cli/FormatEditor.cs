using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Pptx;

public static class FormatEditor
{
    public static FormatEditResult Apply(string inputPath, string planPath, string outputPath)
    {
        var plan = ReadPlan(planPath);
        return Apply(inputPath, plan, outputPath);
    }

    public static FormatEditResult Apply(string inputPath, FormatEditPlan plan, string outputPath)
    {
        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        File.Copy(inputPath, outputPath, overwrite: true);

        var changes = new List<FormatEditChange>();
        var issues = new List<FormatEditIssue>();
        using var presentation = PresentationDocument.Open(outputPath, true);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slides = EnumerateSlides(presentationPart).ToList();
        foreach (var operation in plan.Operations)
        {
            if (operation.SlideNumber < 1 || operation.SlideNumber > slides.Count)
            {
                issues.Add(new FormatEditIssue(operation.SlideNumber, operation.ShapeId, operation.RunIndex, "slide not found"));
                continue;
            }

            var slidePart = slides[operation.SlideNumber - 1];
            var shape = slidePart.Slide.Descendants<Shape>()
                .FirstOrDefault(candidate =>
                    candidate.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == operation.ShapeId);
            if (shape?.TextBody is null)
            {
                issues.Add(new FormatEditIssue(operation.SlideNumber, operation.ShapeId, operation.RunIndex, "shape with text body not found"));
                continue;
            }

            var runEntry = FindRun(shape.TextBody, operation.RunIndex);
            if (runEntry is null)
            {
                issues.Add(new FormatEditIssue(operation.SlideNumber, operation.ShapeId, operation.RunIndex, "run not found"));
                continue;
            }

            var changed = ApplyOperation(runEntry.Value.Run, runEntry.Value.Paragraph, operation);
            if (changed.Count == 0)
            {
                continue;
            }

            changes.Add(new FormatEditChange(operation.SlideNumber, operation.ShapeId, operation.RunIndex, changed));
            slidePart.Slide.Save();
        }

        return new FormatEditResult(
            Input: inputPath,
            Output: outputPath,
            OperationCount: plan.Operations.Count,
            ChangedCount: changes.Count,
            Changes: changes,
            Issues: issues);
    }

    private static FormatEditPlan ReadPlan(string planPath)
    {
        var plan = JsonSerializer.Deserialize<FormatEditPlan>(File.ReadAllText(planPath), Json.Options);
        return plan ?? throw new InvalidOperationException("Format edit plan could not be parsed.");
    }

    private static (A.Paragraph Paragraph, A.Run Run)? FindRun(OpenXmlElement textBody, int targetRunIndex)
    {
        var runIndex = 0;
        foreach (var paragraph in textBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                if (runIndex == targetRunIndex)
                {
                    return (paragraph, run);
                }

                runIndex++;
            }
        }

        return null;
    }

    private static List<string> ApplyOperation(A.Run run, A.Paragraph paragraph, FormatEditOperation operation)
    {
        var changed = new List<string>();
        var properties = run.RunProperties ?? run.PrependChild(new A.RunProperties());

        if (!string.IsNullOrWhiteSpace(operation.FontFamily))
        {
            SetTypeface<A.LatinFont>(properties, operation.FontFamily);
            SetTypeface<A.EastAsianFont>(properties, operation.FontFamily);
            SetTypeface<A.ComplexScriptFont>(properties, operation.FontFamily);
            changed.Add("fontFamily");
        }

        if (operation.FontSize is not null)
        {
            properties.FontSize = (int)Math.Round(operation.FontSize.Value * 100d);
            changed.Add("fontSize");
        }

        if (!string.IsNullOrWhiteSpace(operation.Color))
        {
            properties.RemoveAllChildren<A.SolidFill>();
            properties.AppendChild(new A.SolidFill(new A.RgbColorModelHex { Val = operation.Color.ToUpperInvariant() }));
            changed.Add("color");
        }

        if (operation.Bold is not null)
        {
            properties.Bold = operation.Bold.Value;
            changed.Add("bold");
        }

        if (!string.IsNullOrWhiteSpace(operation.ParagraphAlignment))
        {
            var paragraphProperties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new A.ParagraphProperties());
            paragraphProperties.Alignment = ParseAlignment(operation.ParagraphAlignment);
            changed.Add("paragraphAlignment");
        }

        return changed;
    }

    private static void SetTypeface<TFont>(A.RunProperties properties, string typeface)
        where TFont : OpenXmlElement, new()
    {
        var font = properties.GetFirstChild<TFont>() ?? properties.AppendChild(new TFont());
        switch (font)
        {
            case A.LatinFont latin:
                latin.Typeface = typeface;
                break;
            case A.EastAsianFont eastAsian:
                eastAsian.Typeface = typeface;
                break;
            case A.ComplexScriptFont complexScript:
                complexScript.Typeface = typeface;
                break;
        }
    }

    private static A.TextAlignmentTypeValues ParseAlignment(string alignment)
    {
        return alignment.Trim().ToLowerInvariant() switch
        {
            "center" => A.TextAlignmentTypeValues.Center,
            "right" => A.TextAlignmentTypeValues.Right,
            "justified" => A.TextAlignmentTypeValues.Justified,
            "distributed" => A.TextAlignmentTypeValues.Distributed,
            _ => A.TextAlignmentTypeValues.Left,
        };
    }

    private static IEnumerable<SlidePart> EnumerateSlides(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (presentationPart.GetPartById(relationshipId) is SlidePart slidePart)
            {
                yield return slidePart;
            }
        }
    }
}
