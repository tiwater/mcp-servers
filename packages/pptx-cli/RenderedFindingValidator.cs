using System.Text.Json;
using System.Security.Cryptography;

namespace Dockit.Pptx;

public static class RenderedFindingValidator
{
    private static readonly HashSet<string> SupportedKinds = new(StringComparer.Ordinal)
    {
        "edge-clipping", "occlusion", "text-overflow", "low-readability"
    };

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
        if (request.Schema != "tiwater.pptx-render-findings/v1"
            || !ValidSha(request.ArtifactSha256)
            || !ValidSha(manifest.Artifact.Sha256)
            || !ValidSha(inspection.ArtifactSha256)
            || manifest.Artifact.Sha256 != request.ArtifactSha256
            || inspection.ArtifactSha256 != request.ArtifactSha256
            || !File.Exists(inspection.File)
            || FileHash(inspection.File) != inspection.ArtifactSha256
            || inspection.SlideCount < 1
            || inspection.SlideSize.Cx <= 0
            || inspection.SlideSize.Cy <= 0
            || manifest.Pages.Count != inspection.SlideCount
            || !manifest.Pages.Select(page => page.PageNumber).SequenceEqual(Enumerable.Range(1, inspection.SlideCount)))
            throw new InvalidOperationException("render-finding-reference-authority-invalid");
        foreach (var page in manifest.Pages)
        {
            if (!File.Exists(page.Path) || !ValidSha(page.Sha256) || FileHash(page.Path) != page.Sha256)
                throw new InvalidOperationException($"render-finding-raster-binding-invalid:{page.PageNumber}");
            PngDimensions(page.Path);
        }
        var bindings = new List<RenderFindingBinding>(); var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var finding in request.Findings)
        {
            if (string.IsNullOrWhiteSpace(finding.Id)
                || !seen.Add(finding.Id)
                || !SupportedKinds.Contains(finding.Kind))
                throw new InvalidOperationException($"render-finding-invalid:{finding.Id}");
            var page = manifest.Pages.SingleOrDefault(item => item.PageNumber == finding.PageNumber) ?? throw new InvalidOperationException($"render-finding-page-missing:{finding.Id}");
            if (!ValidSha(finding.RasterSha256) || page.Sha256 != finding.RasterSha256 || FileHash(page.Path) != page.Sha256) throw new InvalidOperationException($"render-finding-raster-binding-invalid:{finding.Id}");
            var pixels = PngDimensions(page.Path);
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

    private static bool ValidSha(string? value)
        => value is { Length: 64 } && value.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f');

    private static string FileHash(string path)
        => Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(path)));

    private static (int Width, int Height) PngDimensions(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length < 24
            || !bytes.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 })
            || !bytes.AsSpan(12, 4).SequenceEqual("IHDR"u8))
            throw new InvalidOperationException("render-finding-png-invalid");
        var width = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(16, 4));
        var height = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(20, 4));
        if (width <= 0 || height <= 0) throw new InvalidOperationException("render-finding-png-dimensions-invalid");
        return (width, height);
    }

    private static T Read<T>(string path) => JsonSerializer.Deserialize<T>(File.ReadAllText(path), Json.Options) ?? throw new InvalidOperationException($"render-finding-json-invalid:{Path.GetFileName(path)}");
}
