using System.Text.Json;

namespace Dockit.Pptx;

public static class RenderedFindingValidator
{
    public static RenderFindingMapVerdict Validate(
        PresentationDetailReport inspection,
        RenderFindingManifest manifest,
        RenderFindingRequest request,
        RenderFindingMap actual)
    {
        var failures = new List<string>();
        RenderFindingMap? expected = null;
        try { expected = ReferenceMap(inspection, manifest, request); }
        catch (Exception error) { failures.Add(error.Message); }
        if (expected is not null && JsonSerializer.Serialize(expected, Json.Options) != JsonSerializer.Serialize(actual, Json.Options)) failures.Add("render-finding-map-drift");
        return new RenderFindingMapVerdict("tiwater.pptx-render-finding-map-verdict/v1", failures.Count == 0, failures);
    }

    public static RenderFindingMapVerdict ValidateFiles(string inspectionPath, string manifestPath, string requestPath, string mapPath)
    {
        var inspection = Read<PresentationDetailReport>(inspectionPath);
        var manifest = Read<RenderFindingManifest>(manifestPath);
        var request = Read<RenderFindingRequest>(requestPath);
        var actual = Read<RenderFindingMap>(mapPath);
        return Validate(inspection, manifest, request, actual);
    }

    private static RenderFindingMap ReferenceMap(PresentationDetailReport inspection, RenderFindingManifest manifest, RenderFindingRequest request)
    {
        if (request.Schema != "tiwater.pptx-render-findings/v1" || manifest.Artifact.Sha256 != request.ArtifactSha256 || manifest.Pages.Count != inspection.SlideCount) throw new InvalidOperationException("render-finding-reference-authority-invalid");
        var bindings = new List<RenderFindingBinding>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var finding in request.Findings)
        {
            if (!seen.Add(finding.Id)) throw new InvalidOperationException($"render-finding-id-duplicate:{finding.Id}");
            var page = manifest.Pages.SingleOrDefault(item => item.PageNumber == finding.PageNumber) ?? throw new InvalidOperationException($"render-finding-page-missing:{finding.Id}");
            if (page.Sha256 != finding.RasterSha256 || RenderedFindingMapper.FileSha256(page.Path) != page.Sha256) throw new InvalidOperationException($"render-finding-raster-binding-invalid:{finding.Id}");
            var pixels = RenderedFindingMapper.ReadPngDimensions(page.Path);
            if (finding.Region.X < 0 || finding.Region.Y < 0 || finding.Region.Width <= 0 || finding.Region.Height <= 0 || finding.Region.X + finding.Region.Width > pixels.Width || finding.Region.Y + finding.Region.Height > pixels.Height) throw new InvalidOperationException($"render-finding-region-invalid:{finding.Id}");
            var slide = inspection.Slides.Single(item => item.SlideNumber == finding.PageNumber);
            var master = inspection.Masters.Single(item => item.Path == slide.MasterPath);
            var layout = master.Layouts.Single(item => item.Path == slide.LayoutPath);
            var candidates = new List<RenderObjectLocator>();
            Add(candidates, "slide", slide.Path, slide.SlideNumber, slide.Shapes, inspection.SlideSize, pixels, finding);
            Add(candidates, "layout", layout.Path, slide.SlideNumber, layout.Shapes, inspection.SlideSize, pixels, finding);
            Add(candidates, "master", master.Path, slide.SlideNumber, master.Shapes, inspection.SlideSize, pixels, finding);
            candidates = candidates.OrderBy(item => item.Scope == "slide" ? 0 : item.Scope == "layout" ? 1 : 2).ThenBy(item => item.PartPath, StringComparer.Ordinal).ThenBy(item => item.ZOrder).ThenBy(item => item.ShapeId).ToList();
            bindings.Add(ReferenceBinding(finding, candidates));
        }
        return new RenderFindingMap("tiwater.pptx-render-finding-map/v1", request.ArtifactSha256, manifest.Pages.Count, bindings);
    }

    private static void Add(List<RenderObjectLocator> output, string scope, string partPath, int slideNumber, IReadOnlyList<ShapeDetail> shapes, SlideSizeInfo slideSize, (int Width, int Height) pixels, RenderFindingCandidate finding)
    {
        foreach (var shape in shapes)
        {
            if (shape.Transform is not { } transform) continue;
            var bounds = new PixelRegion(
                ConvertCoordinate(transform.X, slideSize.Cx, pixels.Width),
                ConvertCoordinate(transform.Y, slideSize.Cy, pixels.Height),
                Math.Max(1, ConvertCoordinate(transform.Cx, slideSize.Cx, pixels.Width)),
                Math.Max(1, ConvertCoordinate(transform.Cy, slideSize.Cy, pixels.Height)));
            if (!Overlap(bounds, finding.Region) || !TextMatch(shape.Text, finding.ObservedText)) continue;
            output.Add(new RenderObjectLocator(scope, partPath, slideNumber, shape.ShapeId, shape.Kind, shape.ZOrder, shape.Text, bounds));
        }
    }

    private static RenderFindingBinding ReferenceBinding(RenderFindingCandidate finding, IReadOnlyList<RenderObjectLocator> candidates)
    {
        if (candidates.Count == 0) return new(finding.Id, "unmapped", null, candidates, "not-generated", "render finding intersects no current rendered object");
        if (candidates.Count > 1) return new(finding.Id, "ambiguous", null, candidates, "not-generated", "render finding intersects multiple current rendered objects");
        var target = candidates[0];
        return new(finding.Id, "unique", target, candidates, "not-generated", "object identity is unique but the finding kind has no independently provable geometry-preserving automatic correction contract");
    }

    private static int ConvertCoordinate(long value, long source, int target) => checked((int)Math.Round((double)value * target / source, MidpointRounding.AwayFromZero));
    private static bool Overlap(PixelRegion a, PixelRegion b) => a.X < b.X + b.Width && a.X + a.Width > b.X && a.Y < b.Y + b.Height && a.Y + a.Height > b.Y;
    private static bool TextMatch(string candidate, string? observed)
    {
        if (string.IsNullOrWhiteSpace(observed)) return true;
        var needle = string.Concat(observed.Where(value => !char.IsWhiteSpace(value))).ToUpperInvariant();
        var source = string.Concat(candidate.Where(value => !char.IsWhiteSpace(value))).ToUpperInvariant();
        return needle.Length >= 2 && source.Contains(needle, StringComparison.Ordinal);
    }

    private static T Read<T>(string path) => JsonSerializer.Deserialize<T>(File.ReadAllText(path), Json.Options) ?? throw new InvalidOperationException($"render-finding-json-invalid:{Path.GetFileName(path)}");
}
