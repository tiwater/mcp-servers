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
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        File.Copy(inputPath, outputPath, true);
        var issues = new List<TemplateApplicationIssue>();
        var changed = 0;
        using var template = PresentationDocument.Open(templatePath, false);
        using var output = PresentationDocument.Open(outputPath, true);
        var templatePart = template.PresentationPart ?? throw new InvalidOperationException("Template presentation part not found.");
        var outputPart = output.PresentationPart ?? throw new InvalidOperationException("Output presentation part not found.");
        var previousMasters = outputPart.SlideMasterParts.ToList();
        var targetMaster = templatePart.SlideMasterParts.SingleOrDefault(part => PartPath(part) == plan.TargetMasterPath);
        if (targetMaster is null) return new TemplateApplicationResult(inputPath, templatePath, outputPath, 0, [new(null, "target master not found")]);

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
            if (assignment.ContentBounds is { } contentBounds)
            {
                try { FitSlideContent(slide, contentBounds, assignment.ContentShapeIds!); }
                catch (InvalidOperationException error) { issues.Add(new(assignment.SlideNumber, error.Message)); continue; }
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
        return new TemplateApplicationResult(inputPath, templatePath, outputPath, changed, issues);
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
