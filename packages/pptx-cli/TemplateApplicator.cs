using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Text.Json;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Pptx;

public static class TemplateApplicator
{
    public static TemplateApplicationResult Apply(string inputPath, string templatePath, string planPath, string outputPath)
    {
        var plan = JsonSerializer.Deserialize<TemplateApplicationPlan>(File.ReadAllText(planPath), Json.Options)
            ?? throw new InvalidOperationException("Template application plan could not be parsed.");
        return Apply(inputPath, templatePath, plan, outputPath);
    }

    public static TemplateApplicationResult Apply(string inputPath, string templatePath, TemplateApplicationPlan plan, string outputPath)
    {
        var sourceEvidence = Inspector.InspectDetail(inputPath);
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        File.Copy(inputPath, outputPath, true);
        var issues = new List<TemplateApplicationIssue>();
        var materializedLayoutShapes = new List<MaterializedLayoutShape>();
        var removedSystemPlaceholders = new List<RemovedSystemPlaceholder>();
        var frozenPlaceholderCount = 0;
        var changed = 0;
        using var template = PresentationDocument.Open(templatePath, false);
        using var output = PresentationDocument.Open(outputPath, true);
        var templatePart = template.PresentationPart ?? throw new InvalidOperationException("Template presentation part not found.");
        var outputPart = output.PresentationPart ?? throw new InvalidOperationException("Output presentation part not found.");
        var previousMasters = outputPart.SlideMasterParts.ToList();
        var targetMaster = templatePart.SlideMasterParts.SingleOrDefault(part => PartPath(part) == plan.TargetMasterPath);
        if (targetMaster is null) return new TemplateApplicationResult(inputPath, templatePath, outputPath, 0, [new(null, "target master not found")], [], 0, []);
        if (plan.SystemPlaceholderPolicy is not ("preserve" or "target-template"))
            return new TemplateApplicationResult(inputPath, templatePath, outputPath, 0, [new(null, "system placeholder policy is invalid")], [], 0, []);

        var importedMaster = outputPart.AddPart(targetMaster);
        var masterRelationshipId = outputPart.GetIdOfPart(importedMaster);
        var masterIds = outputPart.Presentation.SlideMasterIdList ?? outputPart.Presentation.PrependChild(new SlideMasterIdList());
        var nextMasterId = masterIds.Elements<SlideMasterId>().Select(item => item.Id?.Value ?? 2147483647U).DefaultIfEmpty(2147483647U).Max() + 1U;
        masterIds.Append(new SlideMasterId { Id = nextMasterId, RelationshipId = masterRelationshipId });

        var slides = EnumerateSlides(outputPart).ToList();
        foreach (var assignment in plan.Slides)
        {
            if (assignment.SlideNumber < 1 || assignment.SlideNumber > slides.Count) { issues.Add(new(assignment.SlideNumber, "slide not found")); continue; }
            var targetLayout = targetMaster.SlideLayoutParts.SingleOrDefault(part => PartPath(part) == assignment.TargetLayoutPath);
            if (targetLayout is null) { issues.Add(new(assignment.SlideNumber, "target layout not found")); continue; }
            var targetRelationshipId = targetMaster.GetIdOfPart(targetLayout);
            if (importedMaster.GetPartById(targetRelationshipId) is not SlideLayoutPart importedLayout) { issues.Add(new(assignment.SlideNumber, "imported target layout not found")); continue; }
            var slide = slides[assignment.SlideNumber - 1];
            if (assignment.ContentBounds is not null && (assignment.ContentShapeIds is null || assignment.ContentShapeIds.Count == 0))
            {
                issues.Add(new(assignment.SlideNumber, "content shape ids are required when content bounds are specified"));
                continue;
            }
            if (assignment.ContentBounds is null && assignment.ContentShapeIds is { Count: > 0 })
            {
                issues.Add(new(assignment.SlideNumber, "content bounds are required when content shape ids are specified"));
                continue;
            }
            var sourceLayout = slide.SlideLayoutPart;
            var preserveIds = assignment.SourceLayoutShapeIdsToPreserve ?? [];
            if (preserveIds.Count != preserveIds.Distinct().Count())
            {
                issues.Add(new(assignment.SlideNumber, "source layout shape ids contain duplicates"));
                continue;
            }
            var sourceLayoutElements = VisualChildren(sourceLayout?.SlideLayout?.CommonSlideData?.ShapeTree)
                .Select(element => (Element: element, ShapeId: ShapeIdFor(element)))
                .Where(item => item.ShapeId is not null)
                .ToDictionary(item => item.ShapeId!.Value, item => item.Element);
            var missingPreserveIds = preserveIds.Where(id => !sourceLayoutElements.ContainsKey(id)).Order().ToList();
            if (missingPreserveIds.Count > 0)
            {
                issues.Add(new(assignment.SlideNumber, $"source layout shapes not found: {string.Join(",", missingPreserveIds)}"));
                continue;
            }
            if (assignment.ContentBounds is { } contentBounds)
            {
                try { FitSlideContent(slide, contentBounds, assignment.ContentShapeIds!); }
                catch (InvalidOperationException error) { issues.Add(new(assignment.SlideNumber, error.Message)); continue; }
            }
            try
            {
                frozenPlaceholderCount += FreezeSlidePlaceholders(slide, sourceEvidence.Slides[assignment.SlideNumber - 1], assignment.SlideNumber, plan.SystemPlaceholderPolicy, removedSystemPlaceholders);
            }
            catch (InvalidOperationException error)
            {
                issues.Add(new(assignment.SlideNumber, error.Message));
                continue;
            }
            foreach (var sourceShapeId in preserveIds)
            {
                var outputShapeId = MaterializeLayoutShape(slide, sourceLayout!, sourceLayoutElements[sourceShapeId]);
                materializedLayoutShapes.Add(new(assignment.SlideNumber, PartPath(sourceLayout!), sourceShapeId, outputShapeId));
            }
            if (slide.SlideLayoutPart is { } oldLayout) slide.DeletePart(oldLayout);
            slide.AddPart(importedLayout);
            changed++;
        }
        if (issues.Count == 0 && changed == slides.Count)
        {
            foreach (var previousMaster in previousMasters)
            {
                var relationshipId = outputPart.GetIdOfPart(previousMaster);
                masterIds.Elements<SlideMasterId>().Where(item => item.RelationshipId?.Value == relationshipId).ToList().ForEach(item => item.Remove());
                outputPart.DeletePart(previousMaster);
            }
            if (templatePart.Presentation.SlideSize is { } targetSize)
                outputPart.Presentation.SlideSize = (SlideSize)targetSize.CloneNode(true);
        }
        outputPart.Presentation.Save();
        return new TemplateApplicationResult(inputPath, templatePath, outputPath, changed, issues, materializedLayoutShapes, frozenPlaceholderCount, removedSystemPlaceholders);
    }

    private static int FreezeSlidePlaceholders(SlidePart slidePart, SlideDetailReport sourceEvidence, int slideNumber, string systemPlaceholderPolicy, List<RemovedSystemPlaceholder> removed)
    {
        var slideShapes = VisualChildren(slidePart.Slide?.CommonSlideData?.ShapeTree).Where(element => ShapeIdFor(element) is not null).GroupBy(ShapeIdFor).Select(group => group.First()).ToList();
        var layoutShapes = VisualChildren(slidePart.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree).Where(element => ShapeIdFor(element) is not null).GroupBy(ShapeIdFor).Select(group => group.First()).ToList();
        var masterShapes = VisualChildren(slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree).Where(element => ShapeIdFor(element) is not null).GroupBy(ShapeIdFor).Select(group => group.First()).ToList();
        var inheritedByShape = new Dictionary<OpenXmlElement, (OpenXmlElement? Style, OpenXmlElement? Geometry)>();
        foreach (var shape in slideShapes)
        {
            var placeholder = PlaceholderFor(shape);
            if (placeholder is null || systemPlaceholderPolicy == "target-template" && IsSystemPlaceholder(placeholder)) continue;
            var layoutMatch = FindPlaceholder(layoutShapes, placeholder, "layout");
            OpenXmlElement? masterMatch = null;
            if (layoutMatch is null || GeometryFor(layoutMatch) is null)
                masterMatch = FindPlaceholder(masterShapes, placeholder, "master");
            inheritedByShape[shape] = (layoutMatch ?? masterMatch,
                layoutMatch is not null && GeometryFor(layoutMatch) is not null ? layoutMatch : masterMatch);
        }
        var changed = 0;
        foreach (var shape in slideShapes)
        {
            var placeholder = PlaceholderFor(shape);
            if (placeholder is not null && systemPlaceholderPolicy == "target-template" && IsSystemPlaceholder(placeholder))
            {
                removed.Add(new(slideNumber, ShapeIdFor(shape) ?? throw new InvalidOperationException("system placeholder has no shape identity"), PlaceholderToken(placeholder)));
                shape.Remove();
                continue;
            }
            inheritedByShape.TryGetValue(shape, out var inherited);
            if (placeholder is not null) MaterializeInheritedTransform(shape, inherited.Geometry);
            if (shape is Shape textShape)
                MaterializeTextStyle(slidePart, textShape, inherited.Style as Shape, sourceEvidence.Shapes.Single(item => item.ShapeId == ShapeIdFor(shape)));
            if (placeholder is not null && !IsSystemPlaceholder(placeholder)) changed++;
        }
        slidePart.Slide?.Save();
        return changed;
    }

    private static void MaterializeTextStyle(SlidePart slidePart, Shape shape, Shape? layoutShape, ShapeDetail evidence)
    {
        var textBody = shape.TextBody;
        if (textBody is null) return;
        var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
        var masterStyle = MasterTextStyle(slidePart, placeholder);
        var layoutListStyle = layoutShape?.TextBody?.ListStyle;
        var slideListStyle = textBody.ListStyle;
        var paragraphs = textBody.Elements<A.Paragraph>().ToList();
        for (var paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
        {
            var paragraph = paragraphs[paragraphIndex];
            var level = paragraph.ParagraphProperties?.Level?.Value ?? 0;
            var merged = new A.ParagraphProperties();
            MergeProperties(merged, LevelProperties(masterStyle, level));
            MergeProperties(merged, LevelProperties(layoutListStyle, level));
            MergeProperties(merged, LevelProperties(slideListStyle, level));
            MergeProperties(merged, paragraph.ParagraphProperties);
            if (!merged.ChildElements.Any(IsBulletProperty)) merged.AddChild(new A.NoBullet(), true);
            paragraph.ParagraphProperties = merged;
        }

        var runs = paragraphs.SelectMany(paragraph => paragraph.Elements<A.Run>()).ToList();
        foreach (var detail in evidence.Runs)
        {
            if (detail.RunIndex < 0 || detail.RunIndex >= runs.Count) throw new InvalidOperationException($"source run evidence is inconsistent:shape={evidence.ShapeId}:run={detail.RunIndex}:count={runs.Count}");
            var properties = runs[detail.RunIndex].RunProperties ??= new A.RunProperties();
            if (!string.IsNullOrWhiteSpace(detail.FontFamily))
            {
                ReplaceChild(properties, new A.LatinFont { Typeface = detail.FontFamily });
                ReplaceChild(properties, new A.EastAsianFont { Typeface = detail.FontFamily });
                ReplaceChild(properties, new A.ComplexScriptFont { Typeface = detail.FontFamily });
            }
            if (detail.FontSize is { } fontSize) properties.FontSize = checked((int)Math.Round(fontSize * 100d, MidpointRounding.AwayFromZero));
            if (detail.Bold is { } bold) properties.Bold = bold;
            if (!string.IsNullOrWhiteSpace(detail.Color))
                ReplaceChild(properties, new A.SolidFill(new A.RgbColorModelHex { Val = detail.Color }));
        }
    }

    private static OpenXmlElement? MasterTextStyle(SlidePart slidePart, PlaceholderShape? placeholder)
    {
        var styles = slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.TextStyles;
        var type = placeholder?.Type?.Value;
        if (type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle) return styles?.TitleStyle;
        if (type == PlaceholderValues.Body || type == PlaceholderValues.SubTitle) return styles?.BodyStyle;
        return styles?.OtherStyle;
    }

    private static OpenXmlElement? LevelProperties(OpenXmlElement? owner, int level)
        => owner?.ChildElements.FirstOrDefault(element => element.LocalName == $"lvl{level + 1}pPr")
            ?? owner?.ChildElements.FirstOrDefault(element => element.LocalName == "defPPr");

    private static void MergeProperties(OpenXmlCompositeElement target, OpenXmlElement? source)
    {
        if (source is null) return;
        foreach (var attribute in source.GetAttributes()) target.SetAttribute(attribute);
        foreach (var child in source.ChildElements)
        {
            foreach (var existing in target.ChildElements.Where(item => item.LocalName == child.LocalName).ToList()) existing.Remove();
            target.AddChild(child.CloneNode(true), true);
        }
    }

    private static bool IsBulletProperty(DocumentFormat.OpenXml.OpenXmlElement element)
        => element.LocalName is "buNone" or "buChar" or "buAutoNum" or "buBlip";

    private static void ReplaceChild(OpenXmlCompositeElement owner, DocumentFormat.OpenXml.OpenXmlElement child)
    {
        foreach (var existing in owner.ChildElements.Where(item => item.GetType() == child.GetType()).ToList()) existing.Remove();
        owner.AddChild(child, true);
    }

    private static bool IsSystemPlaceholder(PlaceholderShape placeholder)
    {
        var value = placeholder.Type?.Value;
        return value == PlaceholderValues.DateAndTime || value == PlaceholderValues.Footer
            || value == PlaceholderValues.SlideNumber || value == PlaceholderValues.Header;
    }

    private static string PlaceholderToken(PlaceholderShape placeholder)
    {
        var value = placeholder.Type?.Value;
        if (value == PlaceholderValues.DateAndTime) return "dt";
        if (value == PlaceholderValues.Footer) return "ftr";
        if (value == PlaceholderValues.SlideNumber) return "sldNum";
        if (value == PlaceholderValues.Header) return "hdr";
        return placeholder.Type?.InnerText ?? "object";
    }

    private static PlaceholderShape? PlaceholderFor(OpenXmlElement element) => element switch
    {
        Shape shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        Picture picture => picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        GroupShape group => group.NonVisualGroupShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
        _ => null,
    };

    private static OpenXmlElement? FindPlaceholder(IEnumerable<OpenXmlElement> shapes, PlaceholderShape requested, string scope)
    {
        var requestedType = requested.Type?.Value ?? PlaceholderValues.Object;
        var requestedIndex = requested.Index?.Value ?? 0U;
        var matches = shapes.Where(shape =>
        {
            var candidate = PlaceholderFor(shape);
            return candidate is not null
                && (candidate.Type?.Value ?? PlaceholderValues.Object) == requestedType
                && (candidate.Index?.Value ?? 0U) == requestedIndex;
        }).ToList();
        if (matches.Count > 1)
            throw new InvalidOperationException($"source placeholder identity is ambiguous in {scope}: {requestedType}/{requestedIndex}");
        return matches.SingleOrDefault();
    }

    private static void MaterializeInheritedTransform(OpenXmlElement target, OpenXmlElement? inherited)
    {
        switch (target, inherited)
        {
            case (Shape current, Shape source) when current.ShapeProperties?.Transform2D is null && source.ShapeProperties?.Transform2D is { } transform:
                current.ShapeProperties ??= new ShapeProperties();
                current.ShapeProperties.Transform2D = (A.Transform2D)transform.CloneNode(true);
                return;
            case (Picture current, Picture source) when current.ShapeProperties?.Transform2D is null && source.ShapeProperties?.Transform2D is { } transform:
                current.ShapeProperties ??= new ShapeProperties();
                current.ShapeProperties.Transform2D = (A.Transform2D)transform.CloneNode(true);
                return;
            case (GraphicFrame current, GraphicFrame source) when current.Transform is null && source.Transform is { } transform:
                current.Transform = (Transform)transform.CloneNode(true);
                return;
            case (GroupShape current, GroupShape source) when current.GroupShapeProperties?.TransformGroup is null && source.GroupShapeProperties?.TransformGroup is { } transform:
                current.GroupShapeProperties ??= new GroupShapeProperties();
                current.GroupShapeProperties.TransformGroup = (A.TransformGroup)transform.CloneNode(true);
                return;
        }
        var geometry = inherited is null ? null : GeometryFor(inherited);
        if (geometry is null) return;
        var (x, y, cx, cy) = geometry.Value;
        switch (target)
        {
            case Shape current when current.ShapeProperties?.Transform2D is null:
                current.ShapeProperties ??= new ShapeProperties();
                current.ShapeProperties.Transform2D = Transform2D(x, y, cx, cy);
                break;
            case Picture current when current.ShapeProperties?.Transform2D is null:
                current.ShapeProperties ??= new ShapeProperties();
                current.ShapeProperties.Transform2D = Transform2D(x, y, cx, cy);
                break;
            case GraphicFrame current when current.Transform is null:
                current.Transform = new Transform(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy });
                break;
            case GroupShape current when current.GroupShapeProperties?.TransformGroup is null:
                current.GroupShapeProperties ??= new GroupShapeProperties();
                current.GroupShapeProperties.TransformGroup = new A.TransformGroup(
                    new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy },
                    new A.ChildOffset { X = 0L, Y = 0L }, new A.ChildExtents { Cx = cx, Cy = cy });
                break;
        }
    }

    private static (long X, long Y, long Cx, long Cy)? GeometryFor(OpenXmlElement element) => element switch
    {
        Shape shape when shape.ShapeProperties?.Transform2D is { } transform => GeometryFor(transform),
        Picture picture when picture.ShapeProperties?.Transform2D is { } transform => GeometryFor(transform),
        GraphicFrame frame when frame.Transform is { } transform => GeometryFor(transform),
        GroupShape group when group.GroupShapeProperties?.TransformGroup is { } transform => GeometryFor(transform),
        _ => null,
    };

    private static (long X, long Y, long Cx, long Cy)? GeometryFor(OpenXmlCompositeElement transform)
    {
        var offset = transform.GetFirstChild<A.Offset>();
        var extents = transform.GetFirstChild<A.Extents>();
        return offset is null || extents is null
            ? null
            : (offset.X?.Value ?? 0L, offset.Y?.Value ?? 0L, extents.Cx?.Value ?? 0L, extents.Cy?.Value ?? 0L);
    }

    private static A.Transform2D Transform2D(long x, long y, long cx, long cy)
        => new(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy });

    private static uint MaterializeLayoutShape(SlidePart slidePart, SlideLayoutPart sourceLayout, DocumentFormat.OpenXml.OpenXmlElement sourceElement)
    {
        var clone = sourceElement.CloneNode(true);
        RewriteRelationships(sourceLayout, slidePart, clone);
        var usedIds = VisualChildren(slidePart.Slide?.CommonSlideData?.ShapeTree).Select(ShapeIdFor).OfType<uint>().ToHashSet();
        var outputShapeId = usedIds.DefaultIfEmpty(1U).Max() + 1U;
        SetShapeId(clone, outputShapeId);
        slidePart.Slide?.CommonSlideData?.ShapeTree?.Append(clone);
        slidePart.Slide?.Save();
        return outputShapeId;
    }

    private static void RewriteRelationships(OpenXmlPart sourcePart, OpenXmlPart targetPart, DocumentFormat.OpenXml.OpenXmlElement clone)
    {
        var relationshipAttributes = SelfAndDescendants(clone)
            .SelectMany(element => element.GetAttributes().Select(attribute => (Element: element, Attribute: attribute)))
            .Where(item => item.Attribute.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                && !string.IsNullOrWhiteSpace(item.Attribute.Value))
            .ToList();
        var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var maybeRelationshipId in relationshipAttributes.Select(item => item.Attribute.Value).Distinct(StringComparer.Ordinal))
        {
            var relationshipId = maybeRelationshipId!;
            string replacement;
            var related = sourcePart.Parts.SingleOrDefault(item => item.RelationshipId == relationshipId);
            if (related.OpenXmlPart is not null)
            {
                targetPart.AddPart(related.OpenXmlPart);
                replacement = targetPart.GetIdOfPart(related.OpenXmlPart);
            }
            else
            {
                var external = sourcePart.ExternalRelationships.SingleOrDefault(item => item.Id == relationshipId)
                    ?? throw new InvalidOperationException($"layout shape relationship not found: {relationshipId}");
                replacement = targetPart.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
            }
            replacements[relationshipId] = replacement;
        }
        foreach (var item in relationshipAttributes)
        {
            var replacement = replacements[item.Attribute.Value!];
            item.Element.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(item.Attribute.Prefix, item.Attribute.LocalName, item.Attribute.NamespaceUri, replacement));
        }
    }

    private static void SetShapeId(DocumentFormat.OpenXml.OpenXmlElement element, uint shapeId)
    {
        var properties = element switch
        {
            Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
            Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
            GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
            GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            _ => null,
        } ?? throw new InvalidOperationException("layout element has no shape identity");
        properties.Id = shapeId;
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');

    private static void FitSlideContent(SlidePart slidePart, TransformInfo target, IReadOnlyList<uint> contentShapeIds)
    {
        if (target.X < 0 || target.Y < 0 || target.Cx <= 0 || target.Cy <= 0)
            throw new InvalidOperationException("content bounds are invalid");
        var requestedIds = contentShapeIds.ToHashSet();
        if (requestedIds.Count != contentShapeIds.Count) throw new InvalidOperationException("content shape ids contain duplicates");
        var selectedElements = VisualChildren(slidePart.Slide?.CommonSlideData?.ShapeTree)
            .Where(element => ShapeIdFor(element) is { } shapeId && requestedIds.Contains(shapeId))
            .ToList();
        var selectedIds = selectedElements.Select(ShapeIdFor).OfType<uint>().ToHashSet();
        var missingIds = requestedIds.Except(selectedIds).Order().ToList();
        if (missingIds.Count > 0) throw new InvalidOperationException($"content shapes not found: {string.Join(",", missingIds)}");
        var transforms = selectedElements
            .Select(MutableTransformFor)
            .Where(value => value is not null)
            .Cast<MutableTransform>()
            .Where(value => value.Cx > 0 && value.Cy > 0)
            .ToList();
        if (transforms.Count == 0) throw new InvalidOperationException("slide has no transformable content");

        var sourceX = transforms.Min(value => value.X);
        var sourceY = transforms.Min(value => value.Y);
        var sourceRight = transforms.Max(value => checked(value.X + value.Cx));
        var sourceBottom = transforms.Max(value => checked(value.Y + value.Cy));
        var sourceWidth = sourceRight - sourceX;
        var sourceHeight = sourceBottom - sourceY;
        if (sourceWidth <= 0 || sourceHeight <= 0) throw new InvalidOperationException("slide content bounds are empty");

        var scale = Math.Min((double)target.Cx / sourceWidth, (double)target.Cy / sourceHeight);
        if (!double.IsFinite(scale) || scale <= 0) throw new InvalidOperationException("slide content scale is invalid");
        var fittedWidth = Round(sourceWidth * scale);
        var fittedHeight = Round(sourceHeight * scale);
        var offsetX = target.X + (target.Cx - fittedWidth) / 2;
        var offsetY = target.Y + (target.Cy - fittedHeight) / 2;
        foreach (var transform in transforms)
        {
            transform.Apply(
                offsetX + Round((transform.X - sourceX) * scale),
                offsetY + Round((transform.Y - sourceY) * scale),
                Math.Max(1, Round(transform.Cx * scale)),
                Math.Max(1, Round(transform.Cy * scale)));
        }
        ScaleTableGeometry(selectedElements, scale);
        ScaleExplicitFontSizes(selectedElements, scale);
        slidePart.Slide?.Save();
    }

    private static void ScaleTableGeometry(IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> selectedElements, double scale)
    {
        if (Math.Abs(scale - 1d) < 0.0000001d) return;
        var scalableAttributes = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal)
        {
            ["gridCol"] = new(StringComparer.Ordinal) { "w" },
            ["tr"] = new(StringComparer.Ordinal) { "h" },
            ["tcPr"] = new(StringComparer.Ordinal) { "marL", "marR", "marT", "marB" },
        };
        foreach (var element in selectedElements.SelectMany(SelfAndDescendants).Where(value => scalableAttributes.ContainsKey(value.LocalName)))
        {
            var attributeNames = scalableAttributes[element.LocalName];
            foreach (var attribute in element.GetAttributes().Where(value => attributeNames.Contains(value.LocalName)).ToList())
            {
                if (!long.TryParse(attribute.Value, out var value) || value < 0) continue;
                var scaled = Math.Max(0, Round(value * scale));
                element.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, scaled.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            }
        }
    }

    private static void ScaleExplicitFontSizes(IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> selectedElements, double scale)
    {
        if (Math.Abs(scale - 1d) < 0.0000001d) return;
        var runPropertyNames = new HashSet<string>(StringComparer.Ordinal) { "rPr", "defRPr", "endParaRPr", "defaultRPr" };
        foreach (var element in selectedElements.SelectMany(SelfAndDescendants).Where(value => runPropertyNames.Contains(value.LocalName)))
        {
            var attribute = element.GetAttributes().FirstOrDefault(value => value.LocalName == "sz");
            if (string.IsNullOrWhiteSpace(attribute.Value) || !int.TryParse(attribute.Value, out var size) || size <= 0) continue;
            var scaled = Math.Max(100, checked((int)Math.Round(size * scale, MidpointRounding.AwayFromZero)));
            element.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, scaled.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
    }

    private static long Round(double value) => checked((long)Math.Round(value, MidpointRounding.AwayFromZero));

    private static IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> VisualChildren(ShapeTree? shapeTree)
    {
        foreach (var child in shapeTree?.ChildElements ?? [])
        {
            if (child.LocalName != "AlternateContent") { yield return child; continue; }
            foreach (var descendant in child.Descendants().Where(value => value is Shape or Picture or GraphicFrame or GroupShape))
                yield return descendant;
        }
    }

    private static MutableTransform? MutableTransformFor(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        if (element is Shape shape && shape.ShapeProperties?.Transform2D is { } shapeTransform)
            return FromTransform(shapeTransform);
        if (element is Picture picture && picture.ShapeProperties?.Transform2D is { } pictureTransform)
            return FromTransform(pictureTransform);
        if (element is GraphicFrame frame && frame.Transform is { } frameTransform)
            return FromTransform(frameTransform);
        if (element is GroupShape group && group.GroupShapeProperties?.TransformGroup is { } groupTransform)
            return FromTransform(groupTransform);
        return null;
    }

    private static uint? ShapeIdFor(DocumentFormat.OpenXml.OpenXmlElement element) => element switch
    {
        Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
        GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
        GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        _ => null,
    };

    private static IEnumerable<DocumentFormat.OpenXml.OpenXmlElement> SelfAndDescendants(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        yield return element;
        foreach (var descendant in element.Descendants()) yield return descendant;
    }

    private static MutableTransform? FromTransform(A.Transform2D transform)
    {
        if (transform.Offset is not { } offset || transform.Extents is not { } extents) return null;
        return new(offset.X ?? 0L, offset.Y ?? 0L, extents.Cx ?? 0L, extents.Cy ?? 0L, (x, y, cx, cy) =>
        {
            offset.X = x; offset.Y = y; extents.Cx = cx; extents.Cy = cy;
        });
    }

    private static MutableTransform? FromTransform(Transform transform)
    {
        if (transform.Offset is not { } offset || transform.Extents is not { } extents) return null;
        return new(offset.X ?? 0L, offset.Y ?? 0L, extents.Cx ?? 0L, extents.Cy ?? 0L, (x, y, cx, cy) =>
        {
            offset.X = x; offset.Y = y; extents.Cx = cx; extents.Cy = cy;
        });
    }

    private static MutableTransform? FromTransform(A.TransformGroup transform)
    {
        if (transform.Offset is not { } offset || transform.Extents is not { } extents) return null;
        return new(offset.X ?? 0L, offset.Y ?? 0L, extents.Cx ?? 0L, extents.Cy ?? 0L, (x, y, cx, cy) =>
        {
            offset.X = x; offset.Y = y; extents.Cx = cx; extents.Cy = cy;
        });
    }

    private sealed record MutableTransform(long X, long Y, long Cx, long Cy, Action<long, long, long, long> Apply);

    private static IEnumerable<SlidePart> EnumerateSlides(PresentationPart presentationPart)
    {
        foreach (var slideId in presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [])
            if (slideId.RelationshipId?.Value is { } id && presentationPart.GetPartById(id) is SlidePart slide) yield return slide;
    }
}
