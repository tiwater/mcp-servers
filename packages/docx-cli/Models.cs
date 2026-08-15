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
    IReadOnlyList<TableParagraphDetail> Paragraphs);

public sealed record TableRowDetail(
    int RowIndex,
    int GridBefore,
    int GridAfter,
    int CellCount,
    int GridWidth,
    bool CantSplit,
    bool KeepNext,
    IReadOnlyList<TableCellDetail> Cells);

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
    IReadOnlyList<TableRowDetail> Rows);

public sealed record TableInspectionReport(
    string Schema,
    string ToolVersion,
    IReadOnlyDictionary<string, string> ExtractionView,
    string File,
    IReadOnlyList<TableDetail> Tables);

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

public sealed record TemplateMigrationObject(
    string Id,
    string Kind,
    string Scope,
    string? ParentId,
    string? Text,
    string? Style,
    IReadOnlyDictionary<string, string> Provenance,
    TemplateMigrationTopology? Topology = null,
    TemplateMigrationSemanticSelector? Selector = null);

public sealed record TemplateMigrationTopology(
    string ContainerObjectId,
    int Row,
    int Column);

public sealed record TemplateMigrationInventory(
    string File,
    string Sha256,
    IReadOnlyList<TemplateMigrationObject> Objects);

public sealed record TemplateMigrationFinding(
    string Id,
    string Kind,
    string SourceObjectId,
    string? BaselineObjectId,
    string Disposition,
    IReadOnlyDictionary<string, string> Evidence);

public sealed record TemplateMigrationAnalysis(
    string Schema,
    TemplateMigrationInventory Source,
    TemplateMigrationInventory Baseline,
    IReadOnlyList<TemplateMigrationFinding> Findings,
    IReadOnlyList<string> UnsupportedObjectKinds);

public sealed record TemplateMigrationMapping(
    string SourceObjectId,
    string? BaselineObjectId,
    string Disposition,
    string? Reason = null);

/// <summary>
/// A hash-bound source body range appended after the baseline body. The range
/// is resolved from semantic selectors before it becomes this plan payload.
/// </summary>
public sealed record TemplateMigrationBodyAppend(
    string SourceStartObjectId,
    string SourceEndObjectId);

/// <summary>
/// A hash-bound contiguous source body range inserted before a target anchor.
/// The paired before/after anchors prove the target location and ordering.
/// </summary>
public sealed record TemplateMigrationBodyInsertion(
    string SourceStartObjectId,
    string SourceEndObjectId,
    string BaselineBeforeObjectId,
    string BaselineAfterObjectId,
    string StylePolicy);

/// <summary>
/// Clears baseline-owned placeholder content without supplying replacement
/// business facts. The object id is bound to the admitted baseline inventory.
/// </summary>
public sealed record TemplateMigrationBaselineClear(
    string BaselineObjectId,
    string Mode);

public sealed record TemplateMigrationPlan(
    string Schema,
    string SourceSha256,
    string BaselineSha256,
    IReadOnlyList<TemplateMigrationMapping> Mappings,
    IReadOnlyList<TemplateMigrationBodyAppend>? BodyAppends = null,
    IReadOnlyList<TemplateMigrationBaselineClear>? BaselineClears = null,
    IReadOnlyList<TemplateMigrationValueProjection>? ValueProjections = null,
    IReadOnlyList<TemplateMigrationBodyInsertion>? BodyInsertions = null,
    IReadOnlyList<TemplateMigrationChoiceSelection>? ChoiceSelections = null);

/// <summary>
/// A semantic scalar projection between two current, hash-attested parent
/// objects. The caller declares only the semantic identity and value shape;
/// the provider derives the value, target runs, and edit operations.
/// </summary>
public sealed record TemplateMigrationValueProjection(
    string SourceParentObjectId,
    string BaselineParentObjectId,
    string Semantic,
    string ValueKind,
    string Extraction);

public sealed record TemplateMigrationChoiceSelection(
    string SourceMemberObjectId,
    string BaselineLabelRunObjectId);

public sealed record TemplateMigrationSemanticSelector(
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("scope")] string? Scope = null,
    [property: JsonPropertyName("text")] string? Text = null,
    [property: JsonPropertyName("sha256")] string? Sha256 = null,
    [property: JsonPropertyName("parentText")] string? ParentText = null,
    [property: JsonPropertyName("previousText")] string? PreviousText = null,
    [property: JsonPropertyName("nextText")] string? NextText = null,
    [property: JsonPropertyName("descendantText")] string? DescendantText = null,
    [property: JsonPropertyName("textState")] string? TextState = null,
    [property: JsonPropertyName("sameRowText")] string? SameRowText = null,
    [property: JsonPropertyName("sameColumnText")] string? SameColumnText = null);

public sealed record TemplateMigrationSemanticCandidateMapping(
    TemplateMigrationSemanticSelector Source,
    TemplateMigrationSemanticSelector? Baseline,
    string Disposition,
    string? Cardinality = null);

public sealed record TemplateMigrationSemanticCandidateBodyAppend(
    TemplateMigrationSemanticSelector SourceStart,
    TemplateMigrationSemanticSelector SourceEnd);

public sealed record TemplateMigrationSemanticCandidateBodyInsertion(
    TemplateMigrationSemanticSelector SourceStart,
    TemplateMigrationSemanticSelector SourceEnd,
    TemplateMigrationSemanticSelector BaselineBefore,
    TemplateMigrationSemanticSelector BaselineAfter,
    string StylePolicy);

public sealed record TemplateMigrationSemanticCandidateValueProjection(
    TemplateMigrationSemanticSelector SourceParent,
    TemplateMigrationSemanticSelector BaselineParent,
    string Semantic,
    string ValueKind,
    string Extraction);

public sealed record TemplateMigrationSemanticCandidateChoiceSelection(
    TemplateMigrationSemanticSelector SourceMember,
    TemplateMigrationSemanticSelector BaselineLabel);

public sealed record TemplateMigrationSemanticCandidateBaselineClear(
    TemplateMigrationSemanticSelector Baseline,
    string Mode);

public sealed record TemplateMigrationSemanticCandidate(
    string Schema,
    IReadOnlyList<TemplateMigrationSemanticCandidateMapping> Mappings,
    IReadOnlyList<TemplateMigrationSemanticCandidateBodyAppend>? BodyAppends = null,
    IReadOnlyList<TemplateMigrationSemanticCandidateValueProjection>? ValueProjections = null,
    IReadOnlyList<TemplateMigrationSemanticCandidateBodyInsertion>? BodyInsertions = null,
    IReadOnlyList<TemplateMigrationSemanticCandidateChoiceSelection>? ChoiceSelections = null,
    IReadOnlyList<TemplateMigrationSemanticCandidateBaselineClear>? BaselineClears = null);

public sealed record TemplateMigrationMappingDerivation(
    string Schema,
    bool Pass,
    TemplateMigrationPlan Plan,
    IReadOnlyList<TemplateMigrationPlanFailure> Unresolved,
    IReadOnlyList<TemplateMigrationSemanticObservation>? UnclaimedBaseline = null);

public sealed record TemplateMigrationSemanticObservation(
    string Kind,
    string Scope,
    string? Text,
    TemplateMigrationSemanticSelector? Selector);

public sealed record TemplateMigrationSuggestedTarget(
    string Basis,
    TemplateMigrationSemanticObservation Baseline);

public sealed record TemplateMigrationRequiredDecision(
    TemplateMigrationSemanticObservation Source,
    IReadOnlyList<TemplateMigrationSuggestedTarget> SuggestedTargets);

public sealed record TemplateMigrationCandidateDiscovery(
    string Schema,
    bool Pass,
    string SourceSha256,
    string BaselineSha256,
    IReadOnlyList<TemplateMigrationRequiredDecision> RequiredDecisions,
    IReadOnlyList<TemplateMigrationSemanticObservation> UnclaimedBaseline);

public sealed record TemplateMigrationPlanFailure(
    string Reason,
    string? SourceObjectId = null,
    string? BaselineObjectId = null,
    string? Detail = null,
    TemplateMigrationSemanticObservation? Source = null,
    TemplateMigrationSemanticObservation? Baseline = null,
    IReadOnlyList<TemplateMigrationSemanticObservation>? BaselineOptions = null);

public sealed record TemplateMigrationMediaCopy(
    string SourceObjectId,
    string BaselineObjectId);

public sealed record TemplateMigrationOperationBuild(
    string Schema,
    bool Pass,
    bool ReviewRequired,
    string SourceSha256,
    string BaselineSha256,
    string? OperationsSha256,
    IReadOnlyList<DocxEditOperation> Operations,
    IReadOnlyList<TemplateMigrationMediaCopy> MediaCopies,
    IReadOnlyList<TemplateMigrationBodyAppend> BodyAppends,
    IReadOnlyList<TemplateMigrationBodyInsertion> BodyInsertions,
    string? PreviewOperationsSha256,
    IReadOnlyList<DocxEditOperation> PreviewOperations,
    IReadOnlyList<TemplateMigrationMediaCopy> PreviewMediaCopies,
    IReadOnlyList<TemplateMigrationPlanFailure> Failures);

public sealed record TemplateMigrationReadback(
    bool Pass,
    IReadOnlyList<TemplateMigrationPlanFailure> Failures);

public sealed record TemplateMigrationOutputValidation(
    string Schema,
    string ToolVersion,
    bool Pass,
    string Source,
    string SourceSha256,
    string Baseline,
    string BaselineSha256,
    string Output,
    string OutputSha256,
    string Plan,
    string PlanSha256,
    TemplateMigrationOperationBuild Build,
    TemplateMigrationReadback Readback,
    IReadOnlyList<TemplateMigrationPlanFailure> Failures);

public sealed record TemplateMigrationApplyResult(
    string Schema,
    bool Pass,
    string? Output,
    TemplateMigrationOperationBuild Build,
    DocxEditResult? Edit,
    IReadOnlyList<TemplateMigrationPlanFailure> MediaFailures,
    TemplateMigrationReadback? Readback);

public sealed record TemplateMigrationPreviewResult(
    string Schema,
    bool Pass,
    bool ReviewRequired,
    bool OutputVerified,
    string? Output,
    TemplateMigrationOperationBuild Build,
    DocxEditResult? Edit,
    IReadOnlyList<TemplateMigrationPlanFailure> MediaFailures,
    TemplateMigrationReadback? Readback);

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
    IReadOnlyList<DocxSemanticFillRule>? Cells = null,
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
    IReadOnlyList<DocxRichTextSegment>? RichText = null,
    DocxFontPolicy? FontPolicy = null,
    IReadOnlyList<string>? ParagraphTexts = null);

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
