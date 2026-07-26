using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

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

public sealed record PresentationDetailReport(
    string File,
    int SlideCount,
    SlideSizeInfo SlideSize,
    IReadOnlyList<MasterDetail> Masters,
    IReadOnlyList<SlideDetailReport> Slides
);

public sealed record MasterDetail(
    string Path,
    string Name,
    string XmlSha256,
    string? ThemePath,
    string? ThemeSha256,
    IReadOnlyList<ShapeDetail> Shapes,
    IReadOnlyList<LayoutDetail> Layouts
);

public sealed record LayoutDetail(
    string Path,
    string Name,
    string? Type,
    string XmlSha256,
    IReadOnlyList<ShapeDetail> Shapes
);

public sealed record SlideSizeInfo(long Cx, long Cy);

public sealed record SlideDetailReport(
    int SlideNumber,
    string Path,
    string? MasterPath,
    string? LayoutPath,
    IReadOnlyList<ShapeDetail> Shapes
);

public sealed record ShapeDetail(
    uint ShapeId,
    string Name,
    string Kind,
    int ZOrder,
    string? PlaceholderType,
    string? MediaPartPath,
    string? MediaSha256,
    string Text,
    TransformInfo? Transform,
    IReadOnlyList<ParagraphDetail> Paragraphs,
    IReadOnlyList<TextRunDetail> Runs,
    TableDetail? Table = null
);

public sealed record TableDetail(
    IReadOnlyList<long> ColumnWidths,
    IReadOnlyList<long> RowHeights,
    IReadOnlyList<TableCellDetail> Cells
);

public sealed record TableCellDetail(
    int RowIndex,
    int ColumnIndex,
    long? MarginLeft,
    long? MarginRight,
    long? MarginTop,
    long? MarginBottom
);

public sealed record TransformInfo(long X, long Y, long Cx, long Cy);

public sealed record ParagraphDetail(
    int ParagraphIndex,
    string Text,
    string? Alignment
);

public sealed record TextRunDetail(
    int RunIndex,
    int ParagraphIndex,
    string Text,
    string? FontFamily,
    double? FontSize,
    string? Color,
    bool? Bold,
    string? DirectFontFamily = null,
    double? DirectFontSize = null,
    string? DirectColor = null,
    bool? DirectBold = null,
    string? FontFamilySource = null,
    string? FontSizeSource = null,
    string? ColorSource = null,
    string? BoldSource = null
);

public sealed record FormatEditPlan(IReadOnlyList<FormatEditOperation> Operations);

public sealed record FormatEditOperation(
    int SlideNumber,
    uint ShapeId,
    int RunIndex,
    string? FontFamily,
    double? FontSize,
    string? Color,
    bool? Bold,
    string? ParagraphAlignment
);

public sealed record FormatEditResult(
    string Input,
    string Output,
    int OperationCount,
    int ChangedCount,
    IReadOnlyList<FormatEditChange> Changes,
    IReadOnlyList<FormatEditIssue> Issues
);

public sealed record FormatEditChange(
    int SlideNumber,
    uint ShapeId,
    int RunIndex,
    IReadOnlyList<string> Properties
);

public sealed record FormatEditIssue(
    int SlideNumber,
    uint ShapeId,
    int RunIndex,
    string Message
);

public sealed record TemplateApplicationPlan(
    string TargetMasterPath,
    IReadOnlyList<SlideLayoutAssignment> Slides
);

public sealed record SlideLayoutAssignment(
    int SlideNumber,
    string TargetLayoutPath,
    TransformInfo? ContentBounds = null,
    IReadOnlyList<uint>? ContentShapeIds = null
);

public sealed record TemplateApplicationResult(
    string Input,
    string Template,
    string Output,
    int ChangedSlideCount,
    IReadOnlyList<TemplateApplicationIssue> Issues
);

public sealed record TemplateApplicationIssue(int? SlideNumber, string Message);

internal static class Json
{
    public static JsonSerializerOptions Options => new()
    {
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
    };
}
