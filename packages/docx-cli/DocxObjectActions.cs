using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Dockit.Docx;

internal static class DocxObjectActions
{
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static DocxEditAppliedOperation InsertBodyRange(WordprocessingDocument targetDocument, DocxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Source)
            || operation.SourceStartBodyIndex is null
            || operation.SourceEndBodyIndex is null
            || operation.TargetBodyIndex is null)
            return Failed(operation, "source, sourceStartBodyIndex, sourceEndBodyIndex, and targetBodyIndex are required");

        var sourcePath = Path.GetFullPath(operation.Source);
        if (!File.Exists(sourcePath)) return Failed(operation, $"source does not exist: {sourcePath}");

        using var sourceDocument = WordprocessingDocument.Open(sourcePath, false);
        var sourceMain = sourceDocument.MainDocumentPart;
        var targetMain = targetDocument.MainDocumentPart;
        var sourceBody = sourceMain?.Document?.Body;
        var targetBody = targetMain?.Document?.Body;
        if (sourceMain is null || sourceBody is null || targetMain is null || targetBody is null)
            return Failed(operation, "source and target must contain main document bodies");

        var sourceChildren = sourceBody.ChildElements.ToList();
        var targetChildren = targetBody.ChildElements.ToList();
        var start = operation.SourceStartBodyIndex.Value;
        var end = operation.SourceEndBodyIndex.Value;
        var targetIndex = operation.TargetBodyIndex.Value;
        if (start < 0 || end < start || end >= sourceChildren.Count)
            return Failed(operation, $"source body range [{start}, {end}] is out of range");
        if (targetIndex < 0 || targetIndex > targetChildren.Count)
            return Failed(operation, $"targetBodyIndex {targetIndex} is out of range");
        if (targetChildren.LastOrDefault() is SectionProperties && targetIndex == targetChildren.Count)
            return Failed(operation, "targetBodyIndex cannot be after the final body section properties");

        var selected = sourceChildren.Skip(start).Take(end - start + 1).ToList();
        if (selected.Any(element => element is not Paragraph and not Table and not SectionProperties))
            return Failed(operation, "source range contains an unsupported direct body object");
        if (selected.OfType<SectionProperties>().Any(section => section != sourceChildren.LastOrDefault()))
            return Failed(operation, "only the final direct body section properties can be copied");
        if (!ValidateSectionBoundary(sourceChildren, start, end, selected, out var boundaryError))
            return Failed(operation, boundaryError);
        if (selected.SelectMany(DescendantsAndSelf).Any(element => element is FootnoteReference or EndnoteReference
            || element is SdtElement
            || element is BookmarkStart or BookmarkEnd or CommentRangeStart or CommentRangeEnd or CommentReference
            || element.LocalName is "ins" or "del" or "moveFrom" or "moveTo" or "altChunk" or "object"))
            return Failed(operation, "source range contains unsupported annotation, linked, revision, content-control, or embedded content");

        var relationshipIds = RelationshipIds(selected).Distinct(StringComparer.Ordinal).ToList();
        foreach (var relationshipId in relationshipIds)
        {
            if (!CanCopyRelationship(sourceMain, relationshipId, out var relationshipError))
                return Failed(operation, relationshipError);
        }

        var styleRoots = new List<OpenXmlElement>(selected);
        foreach (var relationshipId in relationshipIds)
        {
            var relatedPart = PartByIdOrNull(sourceMain, relationshipId);
            if (relatedPart is HeaderPart { Header: not null } header) styleRoots.Add(header.Header);
            if (relatedPart is FooterPart { Footer: not null } footer) styleRoots.Add(footer.Footer);
        }
        styleRoots.AddRange(RequiredSourceStyles(sourceMain, styleRoots));
        if (!TryImportStyles(sourceMain, targetMain, styleRoots, apply: false, out var styleError)) return Failed(operation, styleError);
        if (!TryImportNumbering(sourceMain, targetMain, styleRoots, apply: false, out var numberingError)) return Failed(operation, numberingError);
        TryImportStyles(sourceMain, targetMain, styleRoots, apply: true, out _);
        TryImportNumbering(sourceMain, targetMain, styleRoots, apply: true, out _);

        var relationshipMap = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var relationshipId in relationshipIds)
            relationshipMap[relationshipId] = CopyRelationship(sourceMain, targetMain, relationshipId);

        var clones = selected.Select(element => element.CloneNode(true)).ToList();
        foreach (var clone in clones) RewriteRelationships(clone, relationshipMap);
        RemapDrawingIds(targetBody, clones);

        var anchor = targetIndex < targetChildren.Count ? targetChildren[targetIndex] : null;
        foreach (var clone in clones)
        {
            OpenXmlElement insertion = clone is SectionProperties section
                ? new Paragraph(new ParagraphProperties(section))
                : clone;
            if (anchor is null) targetBody.AppendChild(insertion);
            else targetBody.InsertBefore(insertion, anchor);
        }
        targetMain.Document.Save();
        return new DocxEditAppliedOperation(operation.Type, true,
            $"Inserted source body range [{start}, {end}] before target body index {targetIndex}");
    }

    internal static DocxEditAppliedOperation ReplaceDrawingImage(WordprocessingDocument document, DocxEditOperation operation)
    {
        if (operation.ParagraphIndex is null || operation.DrawingIndex is null || string.IsNullOrWhiteSpace(operation.Image))
            return Failed(operation, "paragraphIndex, drawingIndex, and image are required");
        var main = document.MainDocumentPart;
        var body = main?.Document?.Body;
        if (main is null || body is null) return Failed(operation, "target document body is missing");
        var paragraphs = body.Elements<Paragraph>().ToList();
        if (operation.ParagraphIndex < 0 || operation.ParagraphIndex >= paragraphs.Count)
            return Failed(operation, $"paragraphIndex {operation.ParagraphIndex} is out of range");
        var drawings = paragraphs[operation.ParagraphIndex.Value].Descendants<Drawing>().ToList();
        if (operation.DrawingIndex < 0 || operation.DrawingIndex >= drawings.Count)
            return Failed(operation, $"drawingIndex {operation.DrawingIndex} is out of range");
        var blips = drawings[operation.DrawingIndex.Value].Descendants<A.Blip>().ToList();
        if (blips.Count != 1 || string.IsNullOrWhiteSpace(blips[0].Embed?.Value) || !string.IsNullOrWhiteSpace(blips[0].Link?.Value))
            return Failed(operation, "selected drawing must contain exactly one embedded, non-linked image");

        if (!TryOpenImage(operation.Image, out var imagePath, out var imageType, out var imageError)) return Failed(operation, imageError);
        var imagePart = main.AddImagePart(imageType);
        using (var input = File.OpenRead(imagePath)) imagePart.FeedData(input);
        blips[0].Embed = main.GetIdOfPart(imagePart);
        main.Document.Save();
        return new DocxEditAppliedOperation(operation.Type, true,
            $"Replaced body paragraph[{operation.ParagraphIndex}].drawing[{operation.DrawingIndex}] image");
    }

    internal static DocxEditAppliedOperation InsertBodyImage(WordprocessingDocument document, DocxEditOperation operation)
    {
        if (operation.TargetBodyIndex is null || string.IsNullOrWhiteSpace(operation.Image)
            || operation.WidthEmu is null or <= 0 || operation.HeightEmu is null or <= 0)
            return Failed(operation, "targetBodyIndex, image, widthEmu, and heightEmu are required");
        var main = document.MainDocumentPart;
        var body = main?.Document?.Body;
        if (main is null || body is null) return Failed(operation, "target document body is missing");
        var children = body.ChildElements.ToList();
        var targetIndex = operation.TargetBodyIndex.Value;
        if (targetIndex < 0 || targetIndex > children.Count)
            return Failed(operation, $"targetBodyIndex {targetIndex} is out of range");
        if (children.LastOrDefault() is SectionProperties && targetIndex == children.Count)
            return Failed(operation, "targetBodyIndex cannot be after the final body section properties");
        if (!TryOpenImage(operation.Image, out var imagePath, out var imageType, out var imageError)) return Failed(operation, imageError);

        var imagePart = main.AddImagePart(imageType);
        using (var input = File.OpenRead(imagePath)) imagePart.FeedData(input);
        var relationshipId = main.GetIdOfPart(imagePart);
        var nextId = body.Descendants<DW.DocProperties>().Select(item => item.Id?.Value ?? 0U).DefaultIfEmpty().Max() + 1U;
        var name = string.IsNullOrWhiteSpace(operation.AltText) ? Path.GetFileName(imagePath) : operation.AltText;
        var drawing = BuildInlineDrawing(relationshipId, operation.WidthEmu.Value, operation.HeightEmu.Value, nextId, name!);
        var paragraph = new Paragraph(new Run(drawing));
        if (targetIndex == children.Count) body.AppendChild(paragraph);
        else body.InsertBefore(paragraph, children[targetIndex]);
        main.Document.Save();
        return new DocxEditAppliedOperation(operation.Type, true, $"Inserted image before target body index {targetIndex}");
    }

    private static Drawing BuildInlineDrawing(string relationshipId, long width, long height, uint id, string name)
    {
        var graphicData = new A.GraphicData(
            new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties { Id = id, Name = name },
                    new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(
                    new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
        var inline = new DW.Inline(
                new DW.Extent { Cx = width, Cy = height },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = id, Name = name, Description = name },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(graphicData))
        { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };
        return new Drawing(inline);
    }

    private static bool ValidateSectionBoundary(IReadOnlyList<OpenXmlElement> children, int start, int end, IReadOnlyList<OpenXmlElement> selected, out string error)
    {
        error = string.Empty;
        if (!selected.SelectMany(DescendantsAndSelf).OfType<SectionProperties>().Any()) return true;
        var boundaries = children.Select((element, index) => (element, index))
            .Where(item => item.element is SectionProperties
                || item.element is Paragraph paragraph && paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>() is not null)
            .Select(item => item.index).ToList();
        var sectionStart = boundaries.Where(index => index < start).DefaultIfEmpty(-1).Max() + 1;
        var sectionEnd = boundaries.FirstOrDefault(index => index >= start, -1);
        if (start == sectionStart && end == sectionEnd) return true;
        error = $"section-bearing ranges must select one whole section [{sectionStart}, {sectionEnd}]";
        return false;
    }

    internal static IEnumerable<OpenXmlElement> DescendantsAndSelf(OpenXmlElement element)
    {
        yield return element;
        foreach (var descendant in element.Descendants()) yield return descendant;
    }

    internal static IEnumerable<string> RelationshipIds(IEnumerable<OpenXmlElement> roots)
        => roots.SelectMany(DescendantsAndSelf)
            .SelectMany(element => element.GetAttributes())
            .Where(attribute => attribute.NamespaceUri == RelationshipsNamespace && !string.IsNullOrWhiteSpace(attribute.Value))
            .Select(attribute => attribute.Value!);

    internal static bool CanCopyRelationship(MainDocumentPart source, string id, out string error)
    {
        error = string.Empty;
        if (PartByIdOrNull(source, id) is ImagePart or HeaderPart or FooterPart) return true;
        if (source.HyperlinkRelationships.Any(relationship => relationship.Id == id)) return true;
        error = $"source relationship {id} is missing or has an unsupported part type";
        return false;
    }

    internal static string CopyRelationship(MainDocumentPart source, MainDocumentPart target, string id)
    {
        var part = PartByIdOrNull(source, id);
        return part switch
        {
            ImagePart image => target.GetIdOfPart(target.AddPart(image)),
            HeaderPart header => target.GetIdOfPart(target.AddPart(header)),
            FooterPart footer => target.GetIdOfPart(target.AddPart(footer)),
            _ => target.AddHyperlinkRelationship(source.HyperlinkRelationships.Single(item => item.Id == id).Uri, true).Id,
        };
    }

    private static OpenXmlPart? PartByIdOrNull(OpenXmlPartContainer container, string id)
    {
        try { return container.GetPartById(id); }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    internal static void RewriteRelationships(OpenXmlElement root, IReadOnlyDictionary<string, string> map)
    {
        foreach (var element in DescendantsAndSelf(root))
        foreach (var attribute in element.GetAttributes().Where(item => item.NamespaceUri == RelationshipsNamespace).ToList())
            if (attribute.Value is { } value && map.TryGetValue(value, out var replacement))
                element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, replacement));
    }

    internal static void RemapDrawingIds(Body targetBody, IReadOnlyList<OpenXmlElement> clones)
    {
        var nextId = targetBody.Descendants<DW.DocProperties>().Select(item => item.Id?.Value ?? 0U).DefaultIfEmpty().Max() + 1U;
        foreach (var drawing in clones.SelectMany(DescendantsAndSelf).OfType<Drawing>())
        {
            foreach (var properties in drawing.Descendants<DW.DocProperties>()) properties.Id = nextId++;
            foreach (var properties in drawing.Descendants<PIC.NonVisualDrawingProperties>()) properties.Id = nextId++;
        }
    }

    private static HashSet<string> RequiredStyleIds(IEnumerable<OpenXmlElement> roots)
        => roots.SelectMany(DescendantsAndSelf).SelectMany(element => element switch
        {
            ParagraphStyleId value => [value.Val?.Value],
            RunStyle value => [value.Val?.Value],
            TableStyle value => [value.Val?.Value],
            _ => Array.Empty<string?>(),
        }).Where(value => !string.IsNullOrWhiteSpace(value)).Cast<string>().ToHashSet(StringComparer.Ordinal);

    private static IReadOnlyList<Style> RequiredSourceStyles(MainDocumentPart source, IEnumerable<OpenXmlElement> roots)
    {
        var styles = source.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(style => style.StyleId?.Value is not null)
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal) ?? [];
        var queue = new Queue<string>(RequiredStyleIds(roots));
        var result = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (queue.TryDequeue(out var id))
        {
            if (!seen.Add(id) || !styles.TryGetValue(id, out var style)) continue;
            result.Add(style);
            foreach (var dependency in new[] { style.BasedOn?.Val?.Value, style.NextParagraphStyle?.Val?.Value, style.LinkedStyle?.Val?.Value })
                if (!string.IsNullOrWhiteSpace(dependency)) queue.Enqueue(dependency);
        }
        return result;
    }

    internal static bool TryImportStyles(MainDocumentPart source, MainDocumentPart target, IReadOnlyList<OpenXmlElement> roots, bool apply, out string error)
    {
        error = string.Empty;
        var requested = RequiredStyleIds(roots);
        if (requested.Count == 0) return true;
        var sourceStyles = source.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal) ?? [];
        var targetStylesPart = target.StyleDefinitionsPart;
        var targetStyles = (targetStylesPart?.Styles?.Elements<Style>() ?? [])
            .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal);
        var queue = new Queue<string>(requested);
        var ordered = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (queue.TryDequeue(out var id))
        {
            if (!seen.Add(id)) continue;
            if (!sourceStyles.TryGetValue(id, out var sourceStyle)) { error = $"source style is missing: {id}"; return false; }
            if (targetStyles.TryGetValue(id, out var targetStyle))
            {
                if (sourceStyle.OuterXml != targetStyle.OuterXml) { error = $"target style conflicts with source style: {id}"; return false; }
                continue;
            }
            ordered.Add(sourceStyle);
            foreach (var dependency in new[] { sourceStyle.BasedOn?.Val?.Value, sourceStyle.NextParagraphStyle?.Val?.Value, sourceStyle.LinkedStyle?.Val?.Value })
                if (!string.IsNullOrWhiteSpace(dependency)) queue.Enqueue(dependency);
        }
        if (!apply) return true;
        targetStylesPart ??= target.AddNewPart<StyleDefinitionsPart>();
        targetStylesPart.Styles ??= new Styles();
        foreach (var style in ordered.AsEnumerable().Reverse()) targetStylesPart.Styles.AppendChild((Style)style.CloneNode(true));
        targetStylesPart.Styles.Save();
        return true;
    }

    internal static bool TryImportNumbering(MainDocumentPart source, MainDocumentPart target, IReadOnlyList<OpenXmlElement> roots, bool apply, out string error)
    {
        error = string.Empty;
        var requested = roots.SelectMany(DescendantsAndSelf).OfType<NumberingId>()
            .Select(item => item.Val?.Value).Where(value => value is not null).Cast<int>().ToHashSet();
        if (requested.Count == 0) return true;
        var sourceNumbering = source.NumberingDefinitionsPart?.Numbering;
        if (sourceNumbering is null) { error = "source numbering definitions are missing"; return false; }
        var targetPart = target.NumberingDefinitionsPart;
        var targetNumbering = targetPart?.Numbering;
        var abstractsToAdd = new List<AbstractNum>();
        var instancesToAdd = new List<NumberingInstance>();
        foreach (var numId in requested)
        {
            var sourceInstance = sourceNumbering.Elements<NumberingInstance>().SingleOrDefault(item => item.NumberID?.Value == numId);
            if (sourceInstance is null) { error = $"source numbering instance is missing: {numId}"; return false; }
            var targetInstance = targetNumbering?.Elements<NumberingInstance>().SingleOrDefault(item => item.NumberID?.Value == numId);
            if (targetInstance is not null && targetInstance.OuterXml != sourceInstance.OuterXml) { error = $"target numbering instance conflicts: {numId}"; return false; }
            var abstractId = sourceInstance.AbstractNumId?.Val?.Value;
            if (abstractId is null) { error = $"source numbering instance {numId} has no abstract numbering id"; return false; }
            var sourceAbstract = sourceNumbering.Elements<AbstractNum>().SingleOrDefault(item => item.AbstractNumberId?.Value == abstractId);
            if (sourceAbstract is null) { error = $"source abstract numbering is missing: {abstractId}"; return false; }
            var targetAbstract = targetNumbering?.Elements<AbstractNum>().SingleOrDefault(item => item.AbstractNumberId?.Value == abstractId);
            if (targetAbstract is not null && targetAbstract.OuterXml != sourceAbstract.OuterXml) { error = $"target abstract numbering conflicts: {abstractId}"; return false; }
            if (targetAbstract is null && abstractsToAdd.All(item => item.AbstractNumberId?.Value != abstractId)) abstractsToAdd.Add(sourceAbstract);
            if (targetInstance is null) instancesToAdd.Add(sourceInstance);
        }
        if (!apply) return true;
        targetPart ??= target.AddNewPart<NumberingDefinitionsPart>();
        targetPart.Numbering ??= new Numbering();
        foreach (var abstractNumbering in abstractsToAdd) targetPart.Numbering.AddChild((AbstractNum)abstractNumbering.CloneNode(true), true);
        foreach (var instance in instancesToAdd) targetPart.Numbering.AddChild((NumberingInstance)instance.CloneNode(true), true);
        targetPart.Numbering.Save();
        return true;
    }

    private static bool TryOpenImage(string value, out string path, out PartTypeInfo type, out string error)
    {
        path = Path.GetFullPath(value);
        error = string.Empty;
        type = ImagePartType.Png;
        if (!File.Exists(path)) { error = $"image does not exist: {path}"; return false; }
        var extension = Path.GetExtension(path).ToLowerInvariant();
        type = extension switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            ".ico" => ImagePartType.Icon,
            ".svg" => ImagePartType.Svg,
            _ => default,
        };
        if (type == default) { error = $"unsupported image extension: {extension}"; return false; }
        var header = File.ReadAllBytes(path).Take(12).ToArray();
        var signatureMatches = extension switch
        {
            ".png" => header.AsSpan().StartsWith(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }),
            ".jpg" or ".jpeg" => header.AsSpan().StartsWith(new byte[] { 0xFF, 0xD8, 0xFF }),
            ".gif" => header.AsSpan().StartsWith("GIF8"u8),
            ".bmp" => header.AsSpan().StartsWith("BM"u8),
            ".tif" or ".tiff" => header.AsSpan().StartsWith(new byte[] { 0x49, 0x49, 0x2A, 0x00 }) || header.AsSpan().StartsWith(new byte[] { 0x4D, 0x4D, 0x00, 0x2A }),
            ".ico" => header.AsSpan().StartsWith(new byte[] { 0x00, 0x00, 0x01, 0x00 }),
            ".svg" => File.ReadAllText(path).Contains("<svg", StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
        if (!signatureMatches) { error = $"image content does not match extension: {extension}"; return false; }
        return true;
    }

    private static DocxEditAppliedOperation Failed(DocxEditOperation operation, string detail)
        => new(operation.Type, false, detail);
}
