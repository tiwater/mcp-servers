using System.Text.Json;
using System.Text.Json.Serialization;

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

public sealed record DocxEditOperation(
    string Type,
    string? CommentId = null,
    string? Text = null,
    int? ParagraphIndex = null,
    int? TableIndex = null,
    int? RowIndex = null,
    int? CellIndex = null,
    IReadOnlyList<string>? CommentIds = null);

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

public sealed record DocxPlanRequest(
    string? Scenario,
    IReadOnlyList<string>? SourceHints);

public sealed record DocxPlanCandidateTarget(
    string Kind,
    string Description,
    int? ParagraphIndex = null,
    int? TableIndex = null,
    int? RowIndex = null,
    int? CellIndex = null,
    int? RowCount = null,
    int? ColumnCount = null);

public sealed record DocxPlanItem(
    string CommentId,
    string CommentText,
    AnnotationAnchor Anchor,
    string InstructionType,
    string TargetScope,
    IReadOnlyList<DocxPlanCandidateTarget> CandidateTargets,
    IReadOnlyList<string> RequiredSources,
    string Confidence,
    string Reasoning,
    IReadOnlyList<DocxEditOperation> ProposedEdits);

public sealed record DocxPlanResult(
    string Input,
    string Scenario,
    IReadOnlyList<DocxPlanItem> Items,
    IReadOnlyList<DocxEditOperation> ProposedEdits,
    IReadOnlyList<string> Warnings,
    string Confidence);

public sealed record DocxResolveRequest(
    string? Scenario,
    string? StabilityDataPath,
    string? ProtocolPath,
    string? QualityStandardCnPath,
    string? ReportPath,
    string? InspectionReportPath = null,
    string? SamplingPlanPath = null);

public sealed record DocxResolveUnresolvedItem(
    string CommentId,
    string InstructionType,
    string Reason);

public sealed record DocxResolveResult(
    string Input,
    string Scenario,
    IReadOnlyList<DocxEditOperation> Operations,
    IReadOnlyList<string> ResolvedCommentIds,
    IReadOnlyList<DocxResolveUnresolvedItem> UnresolvedItems,
    IReadOnlyList<string> Warnings);

public static class Json
{
    public static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNameCaseInsensitive = true
    };
}
