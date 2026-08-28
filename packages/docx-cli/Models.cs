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
    bool HasTrailingEmptySection,
    int TrailingEmptyBodyParagraphCount,
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
    string? VerticalAlignment,
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
    IReadOnlyList<TableParagraphDetail> Paragraphs);

public sealed record TableRowDetail(
    int RowIndex,
    int GridBefore,
    int GridAfter,
    int CellCount,
    int GridWidth,
    bool CantSplit,
    bool KeepNext,
    IReadOnlyList<TableCellDetail> Cells,
    bool RepeatAsHeader = false);

public sealed record TableStoryReference(
    int SectionIndex,
    string ReferenceType,
    string RelationshipId);

public sealed record TableStoryIdentity(
    string Kind,
    int? HeaderIndex,
    int? FooterIndex,
    IReadOnlyList<TableStoryReference> References);

public sealed record TableMutationAddress(
    string Kind,
    int TableIndex,
    int? HeaderIndex = null,
    int? FooterIndex = null);

public sealed record TableDetail(
    int TableIndex,
    IReadOnlyList<string> ContainmentPath,
    string? ParentCellAddress,
    int RowCount,
    int ColumnCount,
    int GridColumnCount,
    IReadOnlyList<string?> GridColumnWidths,
    string? Width,
    string? WidthType,
    IReadOnlyList<TableRowDetail> Rows,
    TableStoryIdentity? Story = null,
    TableMutationAddress? MutationAddress = null);

public sealed record TableInspectionReport(
    string Schema,
    string ToolVersion,
    IReadOnlyDictionary<string, string> ExtractionView,
    string File,
    IReadOnlyList<TableDetail> Tables,
    IReadOnlyList<TableDetail>? StoryTables = null);

public sealed record StructureSummary(
    int BookmarkCount,
    int HyperlinkCount,
    int FieldCount,
    int ContentControlCount,
    int DrawingCount,
    IReadOnlyList<TableMetadata> Tables,
    IReadOnlyList<AnnotationAnchor> AnnotationAnchors);

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

public sealed record DocxEditOperation(
    string Type,
    string? CommentId = null,
    string? Text = null,
    string? FindText = null,
    int? HeaderIndex = null,
    int? FooterIndex = null,
    int? ParagraphIndex = null,
    int? RunIndex = null,
    int? TableIndex = null,
    int? RowIndex = null,
    int? CellIndex = null,
    int? GridColumn = null,
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
    string? Alignment = null,
    string? Width = null,
    string? WidthType = null,
    string? Orientation = null,
    string? FontSize = null,
    string? Height = null,
    string? HeightRule = null,
    string? EndFindText = null,
    string? MatchMode = null,
    string? EndMatchMode = null,
    string? ParagraphStyle = null,
    string? EndParagraphStyle = null,
    bool? DeleteToBodyEnd = null,
    bool? RemovePrecedingPageBreak = null,
    bool? NoWrap = null,
    bool? CantSplit = null,
    bool? KeepNext = null,
    bool? KeepLines = null,
    bool? Italic = null,
    int? IndentCharactersPerLevel = null,
    bool? RepeatAsHeader = null,
    IReadOnlyList<DocxRichTextSegment>? RichText = null,
    DocxFontPolicy? FontPolicy = null,
    IReadOnlyList<string>? ParagraphTexts = null,
    int? ExpectedCount = null,
    string? Source = null,
    int? SourceStartBodyIndex = null,
    int? SourceEndBodyIndex = null,
    int? TargetBodyIndex = null,
    string? Image = null,
    int? DrawingIndex = null,
    long? WidthEmu = null,
    long? HeightEmu = null,
    string? AltText = null);

public sealed record DocxFontRule(string EastAsia, string Latin, string Size);

public sealed record DocxFontPolicy(string Schema, DocxFontRule Body, DocxFontRule Table);

public sealed record DocxFontFinding(
    string Scope,
    int RunOrdinal,
    string Reason,
    string? FontAscii,
    string? FontHighAnsi,
    string? FontEastAsia,
    string? FontComplexScript,
    string? FontSize,
    string? FontSizeComplexScript);

public sealed record DocxFontValidationReport(
    string Schema,
    string ToolVersion,
    bool Pass,
    string File,
    string FileSha256,
    string PolicySha256,
    int BodyRunCount,
    int TableRunCount,
    IReadOnlyList<DocxFontFinding> Findings);

public sealed record DocxFontRunObservation(
    string Scope,
    int RunOrdinal,
    string Container,
    int RunIndex,
    string Text,
    bool HasText,
    string? FontAscii,
    string? FontHighAnsi,
    string? FontEastAsia,
    string? FontComplexScript,
    string? FontSize,
    string? FontSizeComplexScript);

public sealed record DocxFontInspectionReport(
    string Schema,
    string ToolVersion,
    int BodyRunCount,
    int TableRunCount,
    IReadOnlyList<DocxFontRunObservation> Runs);

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
    string? FontName = null,
    bool? Italic = null,
    string? VerticalAlignment = null);

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
