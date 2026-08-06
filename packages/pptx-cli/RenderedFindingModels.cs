namespace Dockit.Pptx;

public sealed record RenderFindingManifest(
    RenderFindingArtifact Artifact,
    IReadOnlyList<RenderFindingPage> Pages
);

public sealed record RenderFindingArtifact(string Sha256);

public sealed record RenderFindingPage(
    int PageNumber,
    string Path,
    string Sha256
);

public sealed record RenderFindingRequest(
    string Schema,
    string ArtifactSha256,
    IReadOnlyList<RenderFindingCandidate> Findings
);

public sealed record RenderFindingCandidate(
    string Id,
    int PageNumber,
    string RasterSha256,
    string Kind,
    PixelRegion Region,
    string? ObservedText = null
);

public sealed record PixelRegion(int X, int Y, int Width, int Height);

public sealed record RenderObjectLocator(
    string Scope,
    string PartPath,
    int SlideNumber,
    uint ShapeId,
    string Kind,
    int ZOrder,
    string Text,
    PixelRegion PixelBounds
);

public sealed record RenderFindingBinding(
    string FindingId,
    string Status,
    RenderObjectLocator? Target,
    IReadOnlyList<RenderObjectLocator> Candidates,
    string OperationDisposition,
    string Reason
);

public sealed record RenderFindingMap(
    string Schema,
    string ArtifactSha256,
    int PageCount,
    IReadOnlyList<RenderFindingBinding> Findings
);

public sealed record RenderFindingMapVerdict(
    string Schema,
    bool Pass,
    IReadOnlyList<string> Findings
);
