using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Pptx;

public static class Inspector
{
    private static readonly Regex PlaceholderRegex = new(@"{{\s*([^{}]+?)\s*}}", RegexOptions.Compiled);

    public static PresentationReport Inspect(string path)
    {
        using var presentation = PresentationDocument.Open(path, false);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slides = new List<SlideReport>();
        var allPlaceholders = new HashSet<string>(StringComparer.Ordinal);

        var slideParts = EnumerateSlides(presentationPart).ToList();
        for (var i = 0; i < slideParts.Count; i++)
        {
            var slidePart = slideParts[i];
            var texts = ExtractTexts(slidePart.Slide);
            var placeholders = ExtractPlaceholders(texts);
            foreach (var placeholder in placeholders)
            {
                allPlaceholders.Add(placeholder);
            }

            slides.Add(new SlideReport(
                SlideNumber: i + 1,
                Path: NormalizePartPath(slidePart.Uri),
                TextCount: texts.Count,
                Placeholders: placeholders));
        }

        return new PresentationReport(
            File: path,
            SlideCount: slides.Count,
            Placeholders: allPlaceholders.OrderBy(value => value, StringComparer.Ordinal).ToList(),
            Slides: slides);
    }

    public static PresentationDetailReport InspectDetail(string path)
    {
        using var presentation = PresentationDocument.Open(path, false);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slideSize = presentationPart.Presentation.SlideSize;
        var slides = new List<SlideDetailReport>();
        var slideParts = EnumerateSlides(presentationPart).ToList();
        for (var i = 0; i < slideParts.Count; i++)
        {
            var slidePart = slideParts[i];
            slides.Add(new SlideDetailReport(
                SlideNumber: i + 1,
                Path: NormalizePartPath(slidePart.Uri),
                Shapes: ExtractShapes(slidePart.Slide)));
        }

        return new PresentationDetailReport(
            File: path,
            SlideCount: slides.Count,
            SlideSize: new SlideSizeInfo(slideSize?.Cx ?? 0L, slideSize?.Cy ?? 0L),
            Slides: slides);
    }

    internal static List<string> ExtractTexts(OpenXmlPartRootElement? root)
    {
        if (root is null)
        {
            return [];
        }

        return root.Descendants<A.Text>()
            .Select(node => node.Text)
            .Where(text => !string.IsNullOrEmpty(text))
            .Select(text => text!)
            .ToList();
    }

    internal static List<string> ExtractPlaceholders(IEnumerable<string> texts)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var placeholders = new List<string>();

        foreach (var text in texts)
        {
            foreach (Match match in PlaceholderRegex.Matches(text))
            {
                var key = match.Groups[1].Value.Trim();
                if (key.Length == 0 || !seen.Add(key))
                {
                    continue;
                }

                placeholders.Add(key);
            }
        }

        return placeholders;
    }

    internal static string NormalizePartPath(Uri? uri)
    {
        if (uri is null)
        {
            return string.Empty;
        }

        var path = uri.OriginalString;
        return path.StartsWith('/') ? path[1..] : path;
    }

    private static List<ShapeDetail> ExtractShapes(Slide? slide)
    {
        if (slide is null)
        {
            return [];
        }

        var shapes = new List<ShapeDetail>();
        foreach (var shape in slide.Descendants<Shape>())
        {
            var id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            shapes.Add(new ShapeDetail(
                ShapeId: id,
                Name: name,
                Kind: "shape",
                Text: string.Concat(shape.TextBody?.Descendants<A.Text>().Select(text => text.Text) ?? []),
                Transform: ExtractTransform(shape.ShapeProperties?.Transform2D),
                Paragraphs: ExtractParagraphs(shape.TextBody),
                Runs: ExtractRuns(shape.TextBody)));
        }

        foreach (var picture in slide.Descendants<Picture>())
        {
            var id = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
            var name = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            shapes.Add(new ShapeDetail(
                ShapeId: id,
                Name: name,
                Kind: "picture",
                Text: string.Empty,
                Transform: ExtractTransform(picture.ShapeProperties?.Transform2D),
                Paragraphs: [],
                Runs: []));
        }

        foreach (var frame in slide.Descendants<GraphicFrame>())
        {
            var id = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
            var name = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            shapes.Add(new ShapeDetail(
                ShapeId: id,
                Name: name,
                Kind: "graphicFrame",
                Text: string.Empty,
                Transform: ExtractTransform(frame.Transform),
                Paragraphs: [],
                Runs: []));
        }

        return shapes.OrderBy(shape => shape.ShapeId).ToList();
    }

    private static TransformInfo? ExtractTransform(A.Transform2D? transform)
    {
        if (transform?.Offset is null && transform?.Extents is null)
        {
            return null;
        }

        return new TransformInfo(
            X: transform.Offset?.X ?? 0L,
            Y: transform.Offset?.Y ?? 0L,
            Cx: transform.Extents?.Cx ?? 0L,
            Cy: transform.Extents?.Cy ?? 0L);
    }

    private static TransformInfo? ExtractTransform(Transform? transform)
    {
        if (transform?.Offset is null && transform?.Extents is null)
        {
            return null;
        }

        return new TransformInfo(
            X: transform.Offset?.X ?? 0L,
            Y: transform.Offset?.Y ?? 0L,
            Cx: transform.Extents?.Cx ?? 0L,
            Cy: transform.Extents?.Cy ?? 0L);
    }

    private static List<ParagraphDetail> ExtractParagraphs(OpenXmlElement? textBody)
    {
        if (textBody is null)
        {
            return [];
        }

        return textBody.Elements<A.Paragraph>()
            .Select((paragraph, index) => new ParagraphDetail(
                ParagraphIndex: index,
                Text: string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text)),
                Alignment: ToAlignment(paragraph.ParagraphProperties?.Alignment?.Value)))
            .ToList();
    }

    private static List<TextRunDetail> ExtractRuns(OpenXmlElement? textBody)
    {
        if (textBody is null)
        {
            return [];
        }

        var runs = new List<TextRunDetail>();
        var runIndex = 0;
        var paragraphIndex = 0;
        var textBodyDefaultRunProperties = textBody.GetFirstChild<A.ListStyle>()?
            .Descendants<A.DefaultRunProperties>()
            .FirstOrDefault();
        foreach (var paragraph in textBody.Elements<A.Paragraph>())
        {
            var paragraphDefaultRunProperties = paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>()
                ?? textBodyDefaultRunProperties;
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var properties = run.RunProperties;
                runs.Add(new TextRunDetail(
                    RunIndex: runIndex,
                    ParagraphIndex: paragraphIndex,
                    Text: run.Text?.Text ?? string.Empty,
                    FontFamily: ExtractFontFamily(properties, paragraphDefaultRunProperties),
                    FontSize: ExtractFontSize(properties, paragraphDefaultRunProperties),
                    Color: ExtractColor(properties, paragraphDefaultRunProperties),
                    Bold: ExtractBold(properties, paragraphDefaultRunProperties)));
                runIndex++;
            }

            paragraphIndex++;
        }

        return runs;
    }

    private static string? ExtractFontFamily(params OpenXmlElement?[] propertyCandidates)
    {
        foreach (var properties in propertyCandidates)
        {
            var value = properties?.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value
                ?? properties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value
                ?? properties?.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value;
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static double? ExtractFontSize(params OpenXmlElement?[] propertyCandidates)
    {
        foreach (var properties in propertyCandidates)
        {
            var value = GetAttributeValue(properties, "sz");
            if (int.TryParse(value, out var fontSize))
            {
                return fontSize / 100d;
            }
        }

        return null;
    }

    private static string? ExtractColor(params OpenXmlElement?[] propertyCandidates)
    {
        foreach (var properties in propertyCandidates)
        {
            var value = properties?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.ToUpperInvariant();
            }
        }

        return null;
    }

    private static bool? ExtractBold(params OpenXmlElement?[] propertyCandidates)
    {
        foreach (var properties in propertyCandidates)
        {
            var value = GetAttributeValue(properties, "b");
            if (string.Equals(value, "1", StringComparison.Ordinal) || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(value, "0", StringComparison.Ordinal) || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return null;
    }

    private static string? GetAttributeValue(OpenXmlElement? element, string localName)
    {
        return element?.GetAttributes()
            .FirstOrDefault(attribute => string.Equals(attribute.LocalName, localName, StringComparison.Ordinal))
            .Value;
    }

    private static string? ToAlignment(A.TextAlignmentTypeValues? alignment)
    {
        if (alignment is null)
        {
            return null;
        }

        if (alignment.Value == A.TextAlignmentTypeValues.Center) return "center";
        if (alignment.Value == A.TextAlignmentTypeValues.Right) return "right";
        if (alignment.Value == A.TextAlignmentTypeValues.Justified) return "justified";
        if (alignment.Value == A.TextAlignmentTypeValues.Distributed) return "distributed";
        return "left";
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
