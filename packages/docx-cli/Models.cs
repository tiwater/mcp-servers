using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace Dockit.Docx;

public sealed record StyleCount(string Style, int Count);

public sealed record HeadingInfo(string Style, string Text, string Source);

public sealed record PackageSummary(int PartCount, IReadOnlyList<string> Parts);

public sealed record ContentSummary(
    int ParagraphCount,
    int TableCount,
    int SectionCount,
    int HeaderPartCount,
    int FooterPartCount,
    IReadOnlyList<HeadingInfo> Headings,
    IReadOnlyList<string> Placeholders);

public sealed record StyleSummary(
    int DefinedParagraphStyleCount,
    int DefinedCharacterStyleCount,
    int DefinedTableStyleCount,
    IReadOnlyList<StyleCount> ParagraphStylesInUse,
    IReadOnlyList<StyleCount> RunStylesInUse);

public sealed record AnnotationSummary(
    int CommentCount,
    int FootnoteCount,
    int EndnoteCount,
    int TrackedChangeElements);

public sealed record AnnotationAnchor(
    string CommentId,
    string? Author,
    string? CommentText,
    string AnchorText,
    string Source,
    string TargetKind,
    int? ParagraphIndex,
    int? TableIndex,
    int? RowIndex,
    int? CellIndex,
    string? NearestHeadingText = null,
    string? CurrentParagraphText = null,
    string? PreviousParagraphText = null,
    string? FollowingParagraphText = null,
    int? CurrentTableRowCount = null,
    int? CurrentTableColumnCount = null);

public sealed record TableMetadata(
    int TableIndex,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<int> RowWidths,
    IReadOnlyList<int> RowCellCounts,
    IReadOnlyList<IReadOnlyList<string>> PreviewRows);

public sealed record TableRunDetail(
    int RunIndex,
    string Text,
    string? Style,
    string? Color,
    string? Underline,
    bool Bold,
    bool Italic,
    string? FontAscii,
    string? FontHighAnsi,
    string? FontEastAsia,
    string? FontComplexScript,
    string? FontSize,
    bool HasTextFill);

public sealed record TableParagraphDetail(
    int ParagraphIndex,
    string Text,
    string? Style,
    string? Justification,
    IReadOnlyList<TableRunDetail> Runs);

public sealed record TableCellDetail(
    int CellIndex,
    int GridColumnStart,
    int GridColumnEnd,
    int GridSpan,
    string? VMerge,
    string? Width,
    string? WidthType,
    string? VerticalAlignment,
    string? ShadingFill,
    string Text,
    IReadOnlyList<TableParagraphDetail> Paragraphs,
    bool NoWrap = false);

public sealed record TableRowDetail(
    int RowIndex,
    int GridBefore,
    int GridAfter,
    int CellCount,
    int GridWidth,
    bool CantSplit,
    IReadOnlyList<TableCellDetail> Cells,
    string? Height = null,
    string? HeightRule = null);

public sealed record TableDetail(
    int TableIndex,
    int RowCount,
    int ColumnCount,
    IReadOnlyList<TableRowDetail> Rows,
    string? Width = null,
    string? WidthType = null,
    string? Layout = null);

public sealed record TableInspectionReport(
    string File,
    IReadOnlyList<TableDetail> Tables);

public sealed record DocumentParagraphDetail(
    string Id,
    int ParagraphIndex,
    string Text,
    string? Style,
    string? Justification,
    IReadOnlyList<TableRunDetail> Runs);

public sealed record BodyNodeDetail(
    string Id,
    int NodeIndex,
    string Kind,
    DocumentParagraphDetail? Paragraph = null,
    int? TableIndex = null);

public sealed record HeaderFooterPartDetail(
    string Id,
    string Kind,
    string RelationshipId,
    string PartUri,
    IReadOnlyList<DocumentParagraphDetail> Paragraphs);

public sealed record SectionPartBinding(
    string Kind,
    string Type,
    string PartId,
    string RelationshipId,
    bool LinkedToPrevious);

public sealed record SectionDetail(
    string Id,
    int SectionIndex,
    string? EndingParagraphId,
    IReadOnlyList<SectionPartBinding> Headers,
    IReadOnlyList<SectionPartBinding> Footers);

public sealed record DrawingPlacementDetail(
    string Kind,
    long? WidthEmu,
    long? HeightEmu,
    string? HorizontalRelativeFrom = null,
    string? VerticalRelativeFrom = null,
    string? HorizontalOffset = null,
    string? VerticalOffset = null,
    uint? DistanceFromTop = null,
    uint? DistanceFromBottom = null,
    uint? DistanceFromLeft = null,
    uint? DistanceFromRight = null);

public sealed record DrawingDetail(
    string Id,
    string Name,
    string RelationshipId,
    string PartUri,
    string ContentSha256,
    string AnchorKind,
    string AnchorText,
    string ParagraphId,
    string? SectionId,
    string? TableId,
    int? RowIndex,
    int? CellIndex,
    DrawingPlacementDetail Placement);

public sealed record StructureSummary(
    int BookmarkCount,
    int HyperlinkCount,
    int FieldCount,
    int ContentControlCount,
    int DrawingCount,
    IReadOnlyList<TableMetadata> Tables,
    IReadOnlyList<AnnotationAnchor> AnnotationAnchors,
    IReadOnlyList<BodyNodeDetail> BodyNodes,
    IReadOnlyList<SectionDetail> Sections,
    IReadOnlyList<HeaderFooterPartDetail> Headers,
    IReadOnlyList<HeaderFooterPartDetail> Footers,
    IReadOnlyList<DrawingDetail> Drawings);

public sealed record FormattingSummary(
    int ParagraphsWithDirectFormatting,
    int RunsWithDirectFormatting);

public sealed record InspectionReport(
    string File,
    PackageSummary Package,
    ContentSummary Content,
    StyleSummary Styles,
    AnnotationSummary Annotations,
    StructureSummary Structure,
    FormattingSummary Formatting);

public sealed record MetricDiff(string Name, object? OldValue, object? NewValue);

public sealed record PackageComparison(
    int SamePartCount,
    int DifferentPartCount,
    IReadOnlyList<string> DifferentParts);

public sealed record StyleDiffSummary(
    IReadOnlyList<StyleCount> AddedParagraphStyles,
    IReadOnlyList<StyleCount> RemovedParagraphStyles,
    IReadOnlyList<StyleCount> AddedRunStyles,
    IReadOnlyList<StyleCount> RemovedRunStyles);

public sealed record ComparisonReport(
    string OldFile,
    string NewFile,
    PackageComparison PackageComparison,
    IReadOnlyList<MetricDiff> MetricDiffs,
    StyleDiffSummary StyleDiffs,
    InspectionReport OldInspection,
    InspectionReport NewInspection);

public sealed record TemplateFieldSlot(
    string Scope,
    string Path,
    string Text,
    bool IsEmptyInputSlot);

public sealed record TemplateSlotMismatch(
    string Path,
    string SourceText,
    string TargetText);

public sealed record TemplateTransformValidationReport(
    string SourceTemplate,
    string TargetTemplate,
    bool IsCompatible,
    int SourceBodyFieldSlotCount,
    int TargetBodyFieldSlotCount,
    int SourceEmptyInputSlotCount,
    int TargetEmptyInputSlotCount,
    IReadOnlyList<TemplateSlotMismatch> MismatchedBodySlots,
    IReadOnlyList<string> Errors,
    IReadOnlyList<string> Warnings);

public sealed record DocxSemanticFillRule(
    IReadOnlyList<string> RowPatterns,
    IReadOnlyList<string> ColPatterns,
    string Text);

public sealed record DocxEditOperation(
    string Type,
    string? CommentId = null,
    string? Text = null,
    string? FindText = null,
    int? HeaderIndex = null,
    int? ParagraphIndex = null,
    int? TableIndex = null,
    int? RowIndex = null,
    int? CellIndex = null,
    IReadOnlyList<IReadOnlyList<DocxTableCellInput>>? Rows = null,
    IReadOnlyList<string>? CommentIds = null,
    int? StartCellIndex = null,
    int? EndCellIndex = null,
    int? StartRowIndex = null,
    int? EndRowIndex = null,
    int? TemplateRowIndex = null,
    int? ColumnIndex = null,
    int? ColumnCount = null,
    int? TemplateColumnIndex = null,
    IReadOnlyList<DocxSemanticFillRule>? Cells = null,
    string? Alignment = null,
    string? Width = null,
    string? WidthType = null,
    string? Orientation = null,
    string? FontSize = null,
    string? Height = null,
    string? HeightRule = null,
    bool? NoWrap = null,
    bool? CantSplit = null,
    IReadOnlyList<DocxRichTextSegment>? RichText = null);

public sealed record DocxTableCellInput(
    string? Text = null,
    int? GridSpan = null,
    string? VMerge = null,
    bool? Bold = null,
    bool? Header = null,
    string? Shading = null,
    string? Alignment = null,
    IReadOnlyList<DocxRichTextSegment>? RichText = null);

public sealed record DocxRichTextSegment(
    string Text,
    string? Color = null,
    bool? Underline = null,
    bool? Bold = null,
    string? FontName = null);

public sealed record DocxEditDocument(
    IReadOnlyList<DocxEditOperation> Operations);

public sealed record DocxEditAppliedOperation(
    string Type,
    bool Applied,
    string Detail);

public sealed record DocxEditResult(
    string Input,
    string Output,
    IReadOnlyList<DocxEditAppliedOperation> AppliedOperations);

public static class Json
{
    public static readonly JsonSerializerOptions Options = new()
    {
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true
    };

    public static readonly JsonSerializerOptions CamelCaseOptions = new()
    {
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };
}
