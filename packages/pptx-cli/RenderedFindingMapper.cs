using System.Security.Cryptography;
using System.Text.Json;

namespace Dockit.Pptx;

public static class RenderedFindingMapper
{
    private static readonly HashSet<string> SupportedKinds = new(StringComparer.Ordinal)
    {
        "edge-clipping", "occlusion", "text-overflow", "low-readability"
    };

    public static RenderFindingMap Map(
        PresentationDetailReport inspection,
        RenderFindingManifest manifest,
        RenderFindingRequest request)
    {
        ValidateAuthority(inspection, manifest, request);
        var pages = manifest.Pages.ToDictionary(page => page.PageNumber);
        var bindings = new List<RenderFindingBinding>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var finding in request.Findings)
        {
            if (!ids.Add(finding.Id)) throw new InvalidOperationException($"render-finding-id-duplicate:{finding.Id}");
            ValidateFinding(finding, pages);
            var page = pages[finding.PageNumber];
            var dimensions = ReadPngDimensions(page.Path);
            var candidates = CandidateObjects(inspection, finding.PageNumber, dimensions)
                .Where(candidate => Intersects(candidate.PixelBounds, finding.Region))
                .Where(candidate => MatchesObservedText(candidate.Text, finding.ObservedText))
                .OrderBy(candidate => ScopeOrder(candidate.Scope))
                .ThenBy(candidate => candidate.PartPath, StringComparer.Ordinal)
                .ThenBy(candidate => candidate.ZOrder)
                .ThenBy(candidate => candidate.ShapeId)
                .ToList();
            bindings.Add(Bind(finding, candidates));
        }
        return new RenderFindingMap(
            "tiwater.pptx-render-finding-map/v1",
            request.ArtifactSha256,
            manifest.Pages.Count,
            bindings);
    }

    public static RenderFindingMap MapFiles(string inspectionPath, string manifestPath, string requestPath)
    {
        var inspection = Read<PresentationDetailReport>(inspectionPath);
        var manifest = Read<RenderFindingManifest>(manifestPath);
        var request = Read<RenderFindingRequest>(requestPath);
        return Map(inspection, manifest, request);
    }

    private static RenderFindingBinding Bind(RenderFindingCandidate finding, IReadOnlyList<RenderObjectLocator> candidates)
    {
        if (candidates.Count == 0)
            return new(finding.Id, "unmapped", null, candidates, "not-generated", "render finding intersects no current rendered object");
        if (candidates.Count > 1)
            return new(finding.Id, "ambiguous", null, candidates, "not-generated", "render finding intersects multiple current rendered objects");
        var target = candidates[0];
        return new(finding.Id, "unique", target, candidates, "not-generated",
            "object identity is unique but the finding kind has no independently provable geometry-preserving automatic correction contract");
    }

    private static IEnumerable<RenderObjectLocator> CandidateObjects(PresentationDetailReport inspection, int slideNumber, (int Width, int Height) pixels)
    {
        var slide = inspection.Slides.Single(item => item.SlideNumber == slideNumber);
        foreach (var shape in slide.Shapes)
            if (shape.Transform is not null)
                yield return Locator("slide", slide.Path, slideNumber, shape, shape.Transform, inspection.SlideSize, pixels);
        var master = inspection.Masters.Single(item => item.Path == slide.MasterPath);
        var layout = master.Layouts.Single(item => item.Path == slide.LayoutPath);
        foreach (var shape in layout.Shapes)
            if (shape.Transform is not null)
                yield return Locator("layout", layout.Path, slideNumber, shape, shape.Transform, inspection.SlideSize, pixels);
        foreach (var shape in master.Shapes)
            if (shape.Transform is not null)
                yield return Locator("master", master.Path, slideNumber, shape, shape.Transform, inspection.SlideSize, pixels);
    }

    private static RenderObjectLocator Locator(string scope, string partPath, int slideNumber, ShapeDetail shape, TransformInfo transform, SlideSizeInfo slideSize, (int Width, int Height) pixels) =>
        new(scope, partPath, slideNumber, shape.ShapeId, shape.Kind, shape.ZOrder, shape.Text,
            new PixelRegion(
                Scale(transform.X, slideSize.Cx, pixels.Width),
                Scale(transform.Y, slideSize.Cy, pixels.Height),
                Math.Max(1, Scale(transform.Cx, slideSize.Cx, pixels.Width)),
                Math.Max(1, Scale(transform.Cy, slideSize.Cy, pixels.Height))));

    private static int Scale(long value, long source, int target) => checked((int)Math.Round((double)value * target / source, MidpointRounding.AwayFromZero));
    private static int ScopeOrder(string scope) => scope switch { "slide" => 0, "layout" => 1, _ => 2 };
    private static bool Intersects(PixelRegion left, PixelRegion right) => left.Width > 0 && left.Height > 0 && right.Width > 0 && right.Height > 0 && left.X < right.X + right.Width && left.X + left.Width > right.X && left.Y < right.Y + right.Height && left.Y + left.Height > right.Y;
    private static string Normalize(string value) => string.Concat(value.Where(character => !char.IsWhiteSpace(character))).ToUpperInvariant();
    private static bool MatchesObservedText(string candidate, string? observed)
    {
        if (string.IsNullOrWhiteSpace(observed)) return true;
        var needle = Normalize(observed); var haystack = Normalize(candidate);
        return needle.Length >= 2 && haystack.Contains(needle, StringComparison.Ordinal);
    }

    private static void ValidateAuthority(PresentationDetailReport inspection, RenderFindingManifest manifest, RenderFindingRequest request)
    {
        if (request.Schema != "tiwater.pptx-render-findings/v1") throw new InvalidOperationException("render-finding-request-schema-invalid");
        if (!Sha(request.ArtifactSha256) || manifest.Artifact.Sha256 != request.ArtifactSha256) throw new InvalidOperationException("render-finding-artifact-binding-invalid");
        if (inspection.SlideCount < 1 || inspection.SlideSize.Cx <= 0 || inspection.SlideSize.Cy <= 0 || manifest.Pages.Count != inspection.SlideCount) throw new InvalidOperationException("render-finding-page-union-invalid");
        if (!manifest.Pages.Select(page => page.PageNumber).SequenceEqual(Enumerable.Range(1, inspection.SlideCount))) throw new InvalidOperationException("render-finding-page-sequence-invalid");
        foreach (var page in manifest.Pages)
        {
            if (!File.Exists(page.Path) || !Sha(page.Sha256) || FileSha256(page.Path) != page.Sha256) throw new InvalidOperationException($"render-finding-raster-binding-invalid:{page.PageNumber}");
            ReadPngDimensions(page.Path);
        }
    }

    private static void ValidateFinding(RenderFindingCandidate finding, IReadOnlyDictionary<int, RenderFindingPage> pages)
    {
        if (string.IsNullOrWhiteSpace(finding.Id) || !pages.TryGetValue(finding.PageNumber, out var page) || finding.RasterSha256 != page.Sha256 || !SupportedKinds.Contains(finding.Kind) || finding.Region.X < 0 || finding.Region.Y < 0 || finding.Region.Width <= 0 || finding.Region.Height <= 0) throw new InvalidOperationException($"render-finding-invalid:{finding.Id}");
        var dimensions = ReadPngDimensions(page.Path);
        if (finding.Region.X + finding.Region.Width > dimensions.Width || finding.Region.Y + finding.Region.Height > dimensions.Height) throw new InvalidOperationException($"render-finding-region-outside-page:{finding.Id}");
    }

    internal static (int Width, int Height) ReadPngDimensions(string path)
    {
        var bytes = File.ReadAllBytes(path);
        if (bytes.Length < 24 || !bytes.AsSpan(0, 8).SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }) || !bytes.AsSpan(12, 4).SequenceEqual("IHDR"u8)) throw new InvalidOperationException("render-finding-png-invalid");
        var width = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(16, 4));
        var height = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(bytes.AsSpan(20, 4));
        if (width <= 0 || height <= 0) throw new InvalidOperationException("render-finding-png-dimensions-invalid");
        return (width, height);
    }

    private static bool Sha(string value) => value.Length == 64 && value.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f');
    internal static string FileSha256(string path) => Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(path)));
    private static T Read<T>(string path) => JsonSerializer.Deserialize<T>(File.ReadAllText(path), Json.Options) ?? throw new InvalidOperationException($"render-finding-json-invalid:{Path.GetFileName(path)}");
}
