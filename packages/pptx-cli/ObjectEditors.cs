using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Pptx;

public static class ShapeGeometryEditor
{
    public static ShapeGeometryResult Apply(string inputPath, string planPath, string outputPath)
        => Apply(inputPath, ObjectEditSupport.ReadPlan<ShapeGeometryPlan>(planPath), outputPath);

    public static ShapeGeometryResult Apply(string inputPath, ShapeGeometryPlan plan, string outputPath)
    {
        var issues = new List<PptxObjectEditIssue>();
        var prepared = new List<(ShapeGeometryChange Change, TransformInfo Before)>();
        var requested = plan.Changes ?? [];
        if (requested.Count == 0)
            issues.Add(new(1, 1U, "changes must contain at least one operation"));

        using (var document = PresentationDocument.Open(inputPath, false))
        {
            var slides = ObjectEditSupport.Slides(document);
            var duplicateKeys = requested.GroupBy(change => (change.SlideNumber, change.ShapeId)).Where(group => group.Count() > 1).Select(group => group.Key).ToHashSet();
            foreach (var change in requested)
            {
                if (change.SlideNumber < 1 || change.ShapeId == 0 || change.Cx < 1 || change.Cy < 1)
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "slideNumber, shapeId, cx, and cy must be positive"));
                    continue;
                }
                if (duplicateKeys.Contains((change.SlideNumber, change.ShapeId)))
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "shape reference is duplicated"));
                    continue;
                }
                var target = ObjectEditSupport.Resolve(slides, change.SlideNumber, change.ShapeId, out var message);
                if (target is null)
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, message));
                    continue;
                }
                var bounds = ObjectEditSupport.Bounds(target);
                if (bounds is null)
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "shape kind or complete explicit transform is not supported"));
                    continue;
                }
                prepared.Add((change, bounds));
            }
        }

        if (issues.Count > 0)
            return new(inputPath, outputPath, requested.Count, 0, [], issues);

        var applied = new List<AppliedShapeGeometryChange>();
        ObjectEditSupport.MutateAtomically(inputPath, outputPath, document =>
        {
            var slides = ObjectEditSupport.Slides(document);
            foreach (var item in prepared)
            {
                var target = ObjectEditSupport.Resolve(slides, item.Change.SlideNumber, item.Change.ShapeId, out _)!;
                ObjectEditSupport.SetBounds(target, item.Change.X, item.Change.Y, item.Change.Cx, item.Change.Cy);
                slides[item.Change.SlideNumber - 1].Slide.Save();
                applied.Add(new(item.Change.SlideNumber, item.Change.ShapeId, item.Before,
                    new(item.Change.X, item.Change.Y, item.Change.Cx, item.Change.Cy)));
            }
        });
        return new(inputPath, outputPath, requested.Count, applied.Count, applied, []);
    }
}

public static class PictureImageEditor
{
    public static PictureImageResult Apply(string inputPath, string planPath, string outputPath)
        => Apply(inputPath, ObjectEditSupport.ReadPlan<PictureImagePlan>(planPath), outputPath);

    public static PictureImageResult Apply(string inputPath, PictureImagePlan plan, string outputPath)
    {
        var issues = new List<PptxObjectEditIssue>();
        var prepared = new List<PreparedPicture>();
        var requested = plan.Changes ?? [];
        if (requested.Count == 0)
            issues.Add(new(1, 1U, "changes must contain at least one operation"));

        using (var document = PresentationDocument.Open(inputPath, false))
        {
            var slides = ObjectEditSupport.Slides(document);
            var duplicateKeys = requested.GroupBy(change => (change.SlideNumber, change.ShapeId)).Where(group => group.Count() > 1).Select(group => group.Key).ToHashSet();
            foreach (var change in requested)
            {
                if (change.SlideNumber < 1 || change.ShapeId == 0 || string.IsNullOrWhiteSpace(change.Image))
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "slideNumber, shapeId, and image are required"));
                    continue;
                }
                if (duplicateKeys.Contains((change.SlideNumber, change.ShapeId)))
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "shape reference is duplicated"));
                    continue;
                }
                var target = ObjectEditSupport.Resolve(slides, change.SlideNumber, change.ShapeId, out var message);
                if (target is not Picture picture)
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, target is null ? message : "shape is not a picture"));
                    continue;
                }
                var embed = picture.BlipFill?.Blip?.Embed?.Value;
                if (string.IsNullOrWhiteSpace(embed) || slides[change.SlideNumber - 1].GetPartById(embed) is not ImagePart imagePart)
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "picture has no supported embedded image relationship"));
                    continue;
                }
                byte[] bytes;
                try { bytes = File.ReadAllBytes(change.Image); }
                catch (Exception error) { issues.Add(new(change.SlideNumber, change.ShapeId, $"image could not be read: {error.Message}")); continue; }
                var detected = ObjectEditSupport.DetectImageContentType(change.Image, bytes);
                if (detected is null || !string.Equals(detected, imagePart.ContentType, StringComparison.OrdinalIgnoreCase))
                {
                    issues.Add(new(change.SlideNumber, change.ShapeId, "replacement image must be a valid PNG or JPEG with the same media type as the current picture"));
                    continue;
                }
                using var current = imagePart.GetStream();
                prepared.Add(new(change, bytes, imagePart.ContentType, ObjectEditSupport.Hash(current)));
            }
        }

        if (issues.Count > 0)
            return new(inputPath, outputPath, requested.Count, 0, [], issues);

        var applied = new List<AppliedPictureImageChange>();
        ObjectEditSupport.MutateAtomically(inputPath, outputPath, document =>
        {
            var slides = ObjectEditSupport.Slides(document);
            foreach (var item in prepared)
            {
                var slide = slides[item.Change.SlideNumber - 1];
                var picture = (Picture)ObjectEditSupport.Resolve(slides, item.Change.SlideNumber, item.Change.ShapeId, out _)!;
                var replacement = slide.AddImagePart(item.ContentType);
                using (var stream = new MemoryStream(item.Bytes, writable: false)) replacement.FeedData(stream);
                picture.BlipFill!.Blip!.Embed = slide.GetIdOfPart(replacement);
                slide.Slide.Save();
                applied.Add(new(item.Change.SlideNumber, item.Change.ShapeId, item.Change.Image, item.BeforeSha256, ObjectEditSupport.Hash(item.Bytes)));
            }
        });
        return new(inputPath, outputPath, requested.Count, applied.Count, applied, []);
    }

    private sealed record PreparedPicture(PictureImageChange Change, byte[] Bytes, string ContentType, string BeforeSha256);
}

internal static class ObjectEditSupport
{
    internal static T ReadPlan<T>(string path)
    {
        var options = Json.Options;
        options.UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow;
        return JsonSerializer.Deserialize<T>(File.ReadAllText(path), options)
            ?? throw new InvalidOperationException("Object edit plan could not be parsed.");
    }

    internal static IReadOnlyList<SlidePart> Slides(PresentationDocument document)
    {
        var part = document.PresentationPart ?? throw new InvalidOperationException("Presentation part not found.");
        var result = new List<SlidePart>();
        foreach (var id in part.Presentation?.SlideIdList?.Elements<SlideId>() ?? [])
            if (!string.IsNullOrWhiteSpace(id.RelationshipId?.Value) && part.GetPartById(id.RelationshipId!.Value!) is SlidePart slide) result.Add(slide);
        return result;
    }

    internal static OpenXmlElement? Resolve(IReadOnlyList<SlidePart> slides, int slideNumber, uint shapeId, out string message)
    {
        if (slideNumber < 1 || slideNumber > slides.Count) { message = "slide not found"; return null; }
        var matches = VisualObjects(slides[slideNumber - 1].Slide.CommonSlideData?.ShapeTree).Where(value => Id(value) == shapeId).ToList();
        if (matches.Count == 0) { message = "shape not found"; return null; }
        if (matches.Count != 1) { message = "shape reference is ambiguous"; return null; }
        message = string.Empty; return matches[0];
    }

    internal static TransformInfo? Bounds(OpenXmlElement element) => element switch
    {
        Shape shape => Read(shape.ShapeProperties?.Transform2D),
        Picture picture => Read(picture.ShapeProperties?.Transform2D),
        GraphicFrame frame => Read(frame.Transform),
        GroupShape group => Read(group.GroupShapeProperties?.TransformGroup),
        _ => null,
    };

    internal static void SetBounds(OpenXmlElement element, long x, long y, long cx, long cy)
    {
        OpenXmlCompositeElement transform = element switch
        {
            Shape shape => shape.ShapeProperties!.Transform2D!,
            Picture picture => picture.ShapeProperties!.Transform2D!,
            GraphicFrame frame => frame.Transform!,
            GroupShape group => group.GroupShapeProperties!.TransformGroup!,
            _ => throw new InvalidOperationException("Unsupported shape kind."),
        };
        var offset = transform.GetFirstChild<A.Offset>()!;
        var extents = transform.GetFirstChild<A.Extents>()!;
        offset.X = x; offset.Y = y; extents.Cx = cx; extents.Cy = cy;
    }

    internal static void MutateAtomically(string inputPath, string outputPath, Action<PresentationDocument> mutation)
    {
        var input = Path.GetFullPath(inputPath); var output = Path.GetFullPath(outputPath);
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Output must differ from input.");
        Directory.CreateDirectory(Path.GetDirectoryName(output) ?? ".");
        var temporary = Path.Combine(Path.GetDirectoryName(output) ?? ".", $".{Path.GetFileName(output)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.Copy(input, temporary, overwrite: false);
            using (var document = PresentationDocument.Open(temporary, true)) mutation(document);
            File.Move(temporary, output, overwrite: true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    internal static string? DetectImageContentType(string path, byte[] bytes)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();
        if (extension == ".png" && bytes.AsSpan().StartsWith(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 })) return "image/png";
        if ((extension == ".jpg" || extension == ".jpeg") && bytes.AsSpan().StartsWith(new byte[] { 255, 216, 255 })) return "image/jpeg";
        return null;
    }

    internal static string Hash(Stream stream) => Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    internal static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static IEnumerable<OpenXmlElement> VisualObjects(ShapeTree? tree)
    {
        foreach (var child in tree?.ChildElements ?? [])
        {
            if (child is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape or ContentPart) yield return child;
            else if (child.LocalName == "AlternateContent")
                foreach (var descendant in child.Descendants().Where(value => value is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape or ContentPart).Where(value => !value.Ancestors<GroupShape>().Any())) yield return descendant;
        }
    }

    private static uint? Id(OpenXmlElement element) => element switch
    {
        Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
        GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
        GroupShape group => group.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        ConnectionShape connector => connector.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
        _ => element.Descendants<NonVisualDrawingProperties>().FirstOrDefault()?.Id?.Value,
    };

    private static TransformInfo? Read(OpenXmlCompositeElement? transform)
    {
        var offset = transform?.GetFirstChild<A.Offset>(); var extents = transform?.GetFirstChild<A.Extents>();
        if (offset?.X?.Value is not long x || offset.Y?.Value is not long y || extents?.Cx?.Value is not long cx || extents.Cy?.Value is not long cy || cx < 1 || cy < 1) return null;
        return new(x, y, cx, cy);
    }
}
