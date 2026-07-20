using System.Text.RegularExpressions;
using System.Security.Cryptography;
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
        var layoutOwners = presentationPart.SlideMasterParts
            .SelectMany(master => master.SlideLayoutParts.Select(layout => (LayoutPath: NormalizePartPath(layout.Uri), MasterPath: NormalizePartPath(master.Uri))))
            .ToDictionary(item => item.LayoutPath, item => item.MasterPath, StringComparer.Ordinal);
        for (var i = 0; i < slideParts.Count; i++)
        {
            var slidePart = slideParts[i];
            var layoutPath = slidePart.SlideLayoutPart is { } layout ? NormalizePartPath(layout.Uri) : null;
            slides.Add(new SlideDetailReport(
                SlideNumber: i + 1,
                Path: NormalizePartPath(slidePart.Uri),
                MasterPath: layoutPath is not null && layoutOwners.TryGetValue(layoutPath, out var masterPath) ? masterPath : null,
                LayoutPath: layoutPath,
                Shapes: ExtractShapes(slidePart)));
        }

        var masters = presentationPart.SlideMasterParts
            .Select(master => new MasterDetail(
                NormalizePartPath(master.Uri),
                master.SlideMaster?.CommonSlideData?.Name?.Value ?? string.Empty,
                HashText(master.SlideMaster?.OuterXml ?? string.Empty),
                master.ThemePart is null ? null : NormalizePartPath(master.ThemePart.Uri),
                master.ThemePart is null ? null : HashPart(master.ThemePart),
                ExtractShapes(master, master.SlideMaster?.CommonSlideData?.ShapeTree),
                master.SlideLayoutParts.Select(layout => new LayoutDetail(
                    NormalizePartPath(layout.Uri),
                    layout.SlideLayout?.CommonSlideData?.Name?.Value ?? string.Empty,
                    GetAttributeValue(layout.SlideLayout, "type"),
                    HashText(layout.SlideLayout?.OuterXml ?? string.Empty),
                    ExtractShapes(layout, layout.SlideLayout?.CommonSlideData?.ShapeTree))).OrderBy(layout => layout.Path, StringComparer.Ordinal).ToList()))
            .OrderBy(master => master.Path, StringComparer.Ordinal).ToList();

        return new PresentationDetailReport(
            File: path,
            SlideCount: slides.Count,
            SlideSize: new SlideSizeInfo(slideSize?.Cx ?? 0L, slideSize?.Cy ?? 0L),
            Masters: masters,
            Slides: slides);
    }

    private static string HashText(string value) => Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static string HashPart(OpenXmlPart part) { using var stream = part.GetStream(); return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(); }

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

    private static List<ShapeDetail> ExtractShapes(SlidePart slidePart)
        => ExtractShapes(slidePart, slidePart.Slide?.CommonSlideData?.ShapeTree);

    private static List<ShapeDetail> ExtractShapes(OpenXmlPart ownerPart, ShapeTree? shapeTree)
    {
        var shapes = new List<ShapeDetail>();
        var zOrder = 0;
        var seen = new HashSet<(string Kind, uint Id)>();
        foreach (var child in VisualChildren(shapeTree))
        {
            if (child is Shape shape)
            {
                var app = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
                var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("shape", shapeId))) continue;
                shapes.Add(new ShapeDetail(shapeId,
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "shape", zOrder++,
                    GetAttributeValue(app?.PlaceholderShape, "type"), null, null,
                    string.Concat(shape.TextBody?.Descendants<A.Text>().Select(text => text.Text) ?? []), ExtractTransform(shape.ShapeProperties?.Transform2D), ExtractParagraphs(shape.TextBody), ExtractRuns(shape.TextBody)));
            }
            else if (child is Picture picture)
            {
                var shapeId = picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("picture", shapeId))) continue;
                string? mediaPath = null; string? mediaHash = null;
                var relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
                if (!string.IsNullOrWhiteSpace(relationshipId) && ownerPart.GetPartById(relationshipId) is OpenXmlPart media)
                {
                    mediaPath = NormalizePartPath(media.Uri); using var stream = media.GetStream(); mediaHash = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
                }
                var app = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
                shapes.Add(new ShapeDetail(shapeId,
                    picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "picture", zOrder++,
                    GetAttributeValue(app?.PlaceholderShape, "type"), mediaPath, mediaHash, string.Empty,
                    ExtractTransform(picture.ShapeProperties?.Transform2D), [], []));
            }
            else if (child is GraphicFrame frame)
            {
                var app = frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties;
                var shapeId = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("graphicFrame", shapeId))) continue;
                shapes.Add(new ShapeDetail(shapeId,
                    frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "graphicFrame", zOrder++,
                    GetAttributeValue(app?.PlaceholderShape, "type"), null, null,
                    string.Concat(frame.Descendants<A.Text>().Select(value => value.Text)), ExtractTransform(frame.Transform),
                    ExtractDescendantParagraphs(frame), ExtractDescendantRuns(frame), ExtractTable(frame)));
            }
            else if (child is GroupShape group)
            {
                var shapeId = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("groupShape", shapeId))) continue;
                shapes.Add(new ShapeDetail(shapeId,
                    group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "groupShape", zOrder++,
                    null, null, null, string.Concat(group.Descendants<A.Text>().Select(value => value.Text)),
                    ExtractTransform(group.GroupShapeProperties?.TransformGroup),
                    ExtractDescendantParagraphs(group), ExtractDescendantRuns(group)));
            }
        }
        return shapes;
    }

    private static IEnumerable<OpenXmlElement> VisualChildren(ShapeTree? shapeTree)
    {
        foreach (var child in shapeTree?.ChildElements ?? [])
        {
            if (child.LocalName != "AlternateContent") { yield return child; continue; }
            foreach (var descendant in child.Descendants().Where(value => value is Shape or Picture or GraphicFrame or GroupShape))
                yield return descendant;
        }
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

    private static TransformInfo? ExtractTransform(A.TransformGroup? transform)
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

    private static List<ParagraphDetail> ExtractDescendantParagraphs(OpenXmlElement owner)
    {
        var result = new List<ParagraphDetail>();
        foreach (var textBody in owner.Descendants().Where(value => value.LocalName == "txBody"))
            foreach (var paragraph in ExtractParagraphs(textBody))
                result.Add(paragraph with { ParagraphIndex = result.Count });
        return result;
    }

    private static List<TextRunDetail> ExtractDescendantRuns(OpenXmlElement owner)
    {
        var result = new List<TextRunDetail>();
        var paragraphOffset = 0;
        foreach (var textBody in owner.Descendants().Where(value => value.LocalName == "txBody"))
        {
            var paragraphs = ExtractParagraphs(textBody);
            foreach (var run in ExtractRuns(textBody))
                result.Add(run with { RunIndex = result.Count, ParagraphIndex = paragraphOffset + run.ParagraphIndex });
            paragraphOffset += paragraphs.Count;
        }
        return result;
    }

    private static TableDetail? ExtractTable(GraphicFrame frame)
    {
        var table = frame.Descendants<A.Table>().SingleOrDefault();
        if (table is null) return null;
        var columnWidths = table.TableGrid?.Elements<A.GridColumn>()
            .Select(column => ParseLongAttribute(column, "w") ?? 0L).ToList() ?? [];
        var rows = table.Elements<A.TableRow>().ToList();
        var rowHeights = rows.Select(row => ParseLongAttribute(row, "h") ?? 0L).ToList();
        var cells = new List<TableCellDetail>();
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var rowCells = rows[rowIndex].Elements<A.TableCell>().ToList();
            for (var columnIndex = 0; columnIndex < rowCells.Count; columnIndex++)
            {
                var properties = rowCells[columnIndex].TableCellProperties;
                cells.Add(new TableCellDetail(rowIndex, columnIndex,
                    ParseLongAttribute(properties, "marL"), ParseLongAttribute(properties, "marR"),
                    ParseLongAttribute(properties, "marT"), ParseLongAttribute(properties, "marB")));
            }
        }
        return new TableDetail(columnWidths, rowHeights, cells);
    }

    private static long? ParseLongAttribute(OpenXmlElement? element, string localName)
        => long.TryParse(GetAttributeValue(element, localName), out var value) ? value : null;

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
