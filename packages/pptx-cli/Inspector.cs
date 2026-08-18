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
            ArtifactSha256: HashFile(path),
            SlideCount: slides.Count,
            SlideSize: new SlideSizeInfo(slideSize?.Cx ?? 0L, slideSize?.Cy ?? 0L),
            Masters: masters,
            Slides: slides);
    }

    private static string HashText(string value) => Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static string HashFile(string path) { using var stream = File.OpenRead(path); return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(); }
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
        => ExtractShapes(slidePart, slidePart.Slide?.CommonSlideData?.ShapeTree, slidePart);

    private static List<ShapeDetail> ExtractShapes(OpenXmlPart ownerPart, ShapeTree? shapeTree, SlidePart? slideContext = null)
    {
        var shapes = new List<ShapeDetail>();
        var zOrder = 0;
        var seen = new HashSet<(string Kind, uint Id)>();
        foreach (var child in VisualChildren(shapeTree))
        {
            if (child is Shape shape)
            {
                var app = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties;
                var placeholder = app?.PlaceholderShape;
                var layoutTextStyle = LayoutTextStyle(slideContext, placeholder);
                var masterTextStyle = MasterTextStyle(slideContext, placeholder);
                var shapeId = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("shape", shapeId))) continue;
                shapes.Add(new ShapeDetail(shapeId,
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "shape", zOrder++,
                    GetAttributeValue(app?.PlaceholderShape, "type"), null, null,
                    string.Concat(shape.TextBody?.Descendants<A.Text>().Select(text => text.Text) ?? []), ExtractTransform(shape.ShapeProperties?.Transform2D), ExtractParagraphs(shape.TextBody, layoutTextStyle, masterTextStyle),
                    ExtractRuns(shape.TextBody, layoutTextStyle, masterTextStyle, slideContext?.SlideLayoutPart?.SlideMasterPart?.ThemePart,
                        slideContext?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap))
                {
                    PlaceholderPresent = placeholder is not null,
                    PlaceholderIndex = placeholder?.Index?.Value
                });
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
                var placeholder = app?.PlaceholderShape;
                shapes.Add(new ShapeDetail(shapeId,
                    picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "picture", zOrder++,
                    GetAttributeValue(placeholder, "type"), mediaPath, mediaHash, string.Empty,
                    ExtractTransform(picture.ShapeProperties?.Transform2D), [], [])
                {
                    PlaceholderPresent = placeholder is not null,
                    PlaceholderIndex = placeholder?.Index?.Value
                });
            }
            else if (child is GraphicFrame frame)
            {
                var app = frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties;
                var placeholder = app?.PlaceholderShape;
                var shapeId = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("graphicFrame", shapeId))) continue;
                shapes.Add(new ShapeDetail(shapeId,
                    frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "graphicFrame", zOrder++,
                    GetAttributeValue(placeholder, "type"), null, null,
                    string.Concat(frame.Descendants<A.Text>().Select(value => value.Text)), ExtractTransform(frame.Transform),
                    ExtractDescendantParagraphs(frame), ExtractDescendantRuns(frame), ExtractTable(frame))
                {
                    PlaceholderPresent = placeholder is not null,
                    PlaceholderIndex = placeholder?.Index?.Value
                });
            }
            else if (child is GroupShape group)
            {
                var app = group.NonVisualGroupShapeProperties?.ApplicationNonVisualDrawingProperties;
                var placeholder = app?.PlaceholderShape;
                var shapeId = group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value ?? 0U;
                if (!seen.Add(("groupShape", shapeId))) continue;
                shapes.Add(new ShapeDetail(shapeId,
                    group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty, "groupShape", zOrder++,
                    GetAttributeValue(placeholder, "type"), null, null, string.Concat(group.Descendants<A.Text>().Select(value => value.Text)),
                    ExtractTransform(group.GroupShapeProperties?.TransformGroup),
                    ExtractDescendantParagraphs(group), ExtractDescendantRuns(group))
                {
                    PlaceholderPresent = placeholder is not null,
                    PlaceholderIndex = placeholder?.Index?.Value
                });
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

    private static List<ParagraphDetail> ExtractParagraphs(OpenXmlElement? textBody, OpenXmlElement? layoutTextStyle = null, OpenXmlElement? masterTextStyle = null)
    {
        if (textBody is null)
        {
            return [];
        }

        var listStyle = textBody.GetFirstChild<A.ListStyle>();
        return textBody.Elements<A.Paragraph>().Select((paragraph, index) =>
        {
            var level = int.TryParse(GetAttributeValue(paragraph.ParagraphProperties, "lvl"), out var parsed) ? parsed : 0;
            var alignment = Resolve(new[]
            {
                new FormatCandidate(paragraph.ParagraphProperties, "paragraph-direct"),
                new FormatCandidate(LevelProperties(listStyle, level), $"shape-list-level-{level + 1}"),
                new FormatCandidate(LevelProperties(layoutTextStyle, level), $"layout-list-level-{level + 1}"),
                new FormatCandidate(LevelProperties(masterTextStyle, level), $"master-text-style-level-{level + 1}"),
            }, value => value is null ? null : ToAlignment(GetAttributeValue(value, "algn") switch
            {
                "ctr" => A.TextAlignmentTypeValues.Center,
                "r" => A.TextAlignmentTypeValues.Right,
                "just" => A.TextAlignmentTypeValues.Justified,
                "dist" => A.TextAlignmentTypeValues.Distributed,
                "l" => A.TextAlignmentTypeValues.Left,
                _ => null,
            })).Value;
            return new ParagraphDetail(index, string.Concat(paragraph.Descendants<A.Text>().Select(text => text.Text)), alignment);
        }).ToList();
    }

    private static List<TextRunDetail> ExtractRuns(OpenXmlElement? textBody, OpenXmlElement? layoutTextStyle = null, OpenXmlElement? masterTextStyle = null, ThemePart? themePart = null, OpenXmlElement? colorMap = null)
    {
        if (textBody is null)
        {
            return [];
        }

        var runs = new List<TextRunDetail>();
        var runIndex = 0;
        var paragraphIndex = 0;
        foreach (var paragraph in textBody.Elements<A.Paragraph>())
        {
            var paragraphDefaultRunProperties = paragraph.ParagraphProperties?
                .GetFirstChild<A.DefaultRunProperties>();
            var paragraphLevel = int.TryParse(GetAttributeValue(paragraph.ParagraphProperties, "lvl"), out var level) ? level : 0;
            var listStyle = textBody.GetFirstChild<A.ListStyle>();
            var levelDefaultRunProperties = listStyle?.ChildElements
                .FirstOrDefault(value => string.Equals(value.LocalName, $"lvl{paragraphLevel + 1}pPr", StringComparison.Ordinal))
                ?.GetFirstChild<A.DefaultRunProperties>();
            var bodyDefaultRunProperties = listStyle?.ChildElements
                .FirstOrDefault(value => string.Equals(value.LocalName, "defPPr", StringComparison.Ordinal))
                ?.GetFirstChild<A.DefaultRunProperties>();
            var masterLevelDefaultRunProperties = masterTextStyle?.ChildElements
                .FirstOrDefault(value => string.Equals(value.LocalName, $"lvl{paragraphLevel + 1}pPr", StringComparison.Ordinal))
                ?.GetFirstChild<A.DefaultRunProperties>();
            var masterDefaultRunProperties = masterTextStyle?.ChildElements
                .FirstOrDefault(value => string.Equals(value.LocalName, "defPPr", StringComparison.Ordinal))
                ?.GetFirstChild<A.DefaultRunProperties>();
            var layoutLevelDefaultRunProperties = LevelProperties(layoutTextStyle, paragraphLevel)?.GetFirstChild<A.DefaultRunProperties>();
            var layoutDefaultRunProperties = layoutTextStyle?.ChildElements
                .FirstOrDefault(value => string.Equals(value.LocalName, "defPPr", StringComparison.Ordinal))
                ?.GetFirstChild<A.DefaultRunProperties>();
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var properties = run.RunProperties;
                var candidates = new[]
                {
                    new FormatCandidate(properties, "direct-run"),
                    new FormatCandidate(paragraphDefaultRunProperties, "paragraph-default"),
                    new FormatCandidate(levelDefaultRunProperties, $"shape-list-level-{paragraphLevel + 1}"),
                    new FormatCandidate(bodyDefaultRunProperties, "shape-list-default"),
                    new FormatCandidate(layoutLevelDefaultRunProperties, $"layout-list-level-{paragraphLevel + 1}"),
                    new FormatCandidate(layoutDefaultRunProperties, "layout-list-default"),
                    new FormatCandidate(masterLevelDefaultRunProperties, $"master-text-style-level-{paragraphLevel + 1}"),
                    new FormatCandidate(masterDefaultRunProperties, "master-text-style-default")
                };
                var fontFamily = Resolve(candidates, properties => ExtractFontFamily(properties, themePart));
                var fontSize = Resolve(candidates, ExtractFontSize);
                var color = Resolve(candidates, properties => ExtractColor(properties, themePart, colorMap));
                var bold = Resolve(candidates, ExtractBold);
                runs.Add(new TextRunDetail(
                    RunIndex: runIndex,
                    ParagraphIndex: paragraphIndex,
                    Text: run.Text?.Text ?? string.Empty,
                    FontFamily: fontFamily.Value,
                    FontSize: fontSize.Value,
                    Color: color.Value,
                    Bold: bold.Value,
                    DirectFontFamily: ExtractFontFamily(properties, themePart),
                    DirectFontSize: ExtractFontSize(properties),
                    DirectColor: ExtractDirectColor(properties),
                    DirectBold: ExtractBold(properties),
                    FontFamilySource: fontFamily.Source,
                    FontSizeSource: fontSize.Source,
                    ColorSource: color.Source,
                    BoldSource: bold.Source));
                runIndex++;
            }

            paragraphIndex++;
        }

        return runs;
    }

    private static OpenXmlElement? MasterTextStyle(SlidePart? slidePart, PlaceholderShape? placeholder)
    {
        var styles = slidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        if (styles is null) return null;
        var type = GetAttributeValue(placeholder, "type");
        if (type is "title" or "ctrTitle") return styles.TitleStyle;
        if (type is "body" or "subTitle") return styles.BodyStyle;
        return styles.OtherStyle;
    }

    private static OpenXmlElement? LayoutTextStyle(SlidePart? slidePart, PlaceholderShape? placeholder)
    {
        if (placeholder is null) return null;
        var type = GetAttributeValue(placeholder, "type");
        var index = GetAttributeValue(placeholder, "idx");
        var shape = slidePart?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree?
            .Elements<Shape>()
            .FirstOrDefault(candidate =>
            {
                var current = candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
                if (current is null || GetAttributeValue(current, "type") != type) return false;
                return string.IsNullOrWhiteSpace(index) || GetAttributeValue(current, "idx") == index;
            });
        return shape?.TextBody?.ListStyle;
    }

    private static OpenXmlElement? LevelProperties(OpenXmlElement? owner, int level)
        => owner?.ChildElements.FirstOrDefault(element => element.LocalName == $"lvl{level + 1}pPr")
            ?? owner?.ChildElements.FirstOrDefault(element => element.LocalName == "defPPr");

    private sealed record FormatCandidate(OpenXmlElement? Properties, string Source);
    private sealed record ResolvedFormat<T>(T? Value, string? Source);

    private static ResolvedFormat<T> Resolve<T>(IEnumerable<FormatCandidate> candidates, Func<OpenXmlElement?, T?> extract)
    {
        foreach (var candidate in candidates)
        {
            var value = extract(candidate.Properties);
            if (value is not null) return new ResolvedFormat<T>(value, candidate.Source);
        }
        return new ResolvedFormat<T>(default, null);
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

    private static string? ExtractFontFamily(OpenXmlElement? properties, ThemePart? themePart = null)
    {
        var value = properties?.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value
            ?? properties?.GetFirstChild<A.LatinFont>()?.Typeface?.Value
            ?? properties?.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value;
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (!value.StartsWith("+m", StringComparison.Ordinal)) return value;
        var family = value.StartsWith("+mj-", StringComparison.Ordinal) ? "majorFont" : "minorFont";
        var script = value.EndsWith("-ea", StringComparison.Ordinal) ? "ea"
            : value.EndsWith("-cs", StringComparison.Ordinal) ? "cs" : "latin";
        var fontScheme = themePart?.Theme?.ThemeElements?.FontScheme;
        var familyElement = fontScheme?.ChildElements.FirstOrDefault(element => element.LocalName == family);
        var resolved = familyElement?.ChildElements.FirstOrDefault(element => element.LocalName == script)
            ?.GetAttribute("typeface", string.Empty).Value;
        if (string.IsNullOrWhiteSpace(resolved) && script != "latin")
            resolved = familyElement?.ChildElements.FirstOrDefault(element => element.LocalName == "latin")
                ?.GetAttribute("typeface", string.Empty).Value;
        return string.IsNullOrWhiteSpace(resolved) ? null : resolved;
    }

    private static double? ExtractFontSize(OpenXmlElement? properties)
    {
        var value = GetAttributeValue(properties, "sz");
        return int.TryParse(value, out var fontSize) ? fontSize / 100d : null;
    }

    private static string? ExtractColor(OpenXmlElement? properties, ThemePart? themePart = null, OpenXmlElement? colorMap = null)
    {
        var fill = properties?.GetFirstChild<A.SolidFill>();
        var rgb = fill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(rgb)) return rgb.ToUpperInvariant();
        var system = fill?.GetFirstChild<A.SystemColor>();
        if (!string.IsNullOrWhiteSpace(system?.LastColor?.Value)) return system.LastColor.Value.ToUpperInvariant();
        var scheme = fill?.GetFirstChild<A.SchemeColor>();
        var schemeName = GetAttributeValue(scheme, "val");
        if (string.IsNullOrWhiteSpace(schemeName)) return null;
        if (scheme!.ChildElements.Count > 0) return null;
        var mappedName = GetAttributeValue(colorMap, schemeName) ?? schemeName;
        var themeColor = themePart?.Theme?.ThemeElements?.ColorScheme?.ChildElements
            .FirstOrDefault(element => string.Equals(element.LocalName, mappedName, StringComparison.Ordinal));
        var themeRgb = themeColor?.Descendants<A.RgbColorModelHex>().FirstOrDefault()?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(themeRgb)) return themeRgb.ToUpperInvariant();
        var themeSystem = themeColor?.Descendants<A.SystemColor>().FirstOrDefault()?.LastColor?.Value;
        return string.IsNullOrWhiteSpace(themeSystem) ? null : themeSystem.ToUpperInvariant();
    }

    private static string? ExtractDirectColor(OpenXmlElement? properties)
    {
        var fill = properties?.GetFirstChild<A.SolidFill>();
        var rgb = fill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(rgb)) return rgb.ToUpperInvariant();
        var system = fill?.GetFirstChild<A.SystemColor>();
        if (!string.IsNullOrWhiteSpace(system?.LastColor?.Value)) return $"system:{GetAttributeValue(system, "val")}";
        var scheme = fill?.GetFirstChild<A.SchemeColor>();
        var schemeName = GetAttributeValue(scheme, "val");
        return string.IsNullOrWhiteSpace(schemeName) ? null : $"scheme:{schemeName}";
    }

    private static bool? ExtractBold(OpenXmlElement? properties)
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
