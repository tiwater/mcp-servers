using System.Text.Json;

namespace Dockit.Pptx;

public sealed record PresentationReport(
    string File,
    int SlideCount,
    IReadOnlyList<string> Placeholders,
    IReadOnlyList<SlideReport> Slides
);

public sealed record SlideReport(
    int SlideNumber,
    string Path,
    int TextCount,
    IReadOnlyList<string> Placeholders
);

public sealed record PresentationExport(
    string File,
    IReadOnlyList<SlideExport> Slides,
    IReadOnlyList<NoteExport> Notes
);

public sealed record SlideExport(
    int SlideNumber,
    string Path,
    IReadOnlyList<string> Texts,
    IReadOnlyList<string> Placeholders
);

public sealed record NoteExport(
    int NotesNumber,
    string Path,
    IReadOnlyList<string> Texts
);

public sealed record FillResult(
    string Template,
    string Output,
    int ChangedSlides,
    int ChangedNotes,
    int PlaceholderCount
);

internal static class Json
{
    public static JsonSerializerOptions Options => new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
    };
}
