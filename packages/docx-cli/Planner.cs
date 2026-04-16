using System.Text.Json;
using System.Text.RegularExpressions;

namespace Dockit.Docx;

public static class Planner
{
    private const string SupportedScenario = "stability-report";

    private static readonly Regex DirectReplacementPattern = new(
        @"(?i)^(?:replace(?:\s+with)?|use|set(?:\s+to)?|change(?:\s+to)?|替换(?:为|成)?|改为|设为|更改为|修改为|调整为|更新为)\s*(?:[:：-]\s*)?(?<text>.+)$",
        RegexOptions.Compiled);

    private static readonly Regex NarrativeDraftPattern = new(
        @"(?i)^(?:write|draft|generate|update(?:\s+to)?|写|撰写|起草|生成|编写|描述|说明|总结|改写).*?(?:[:：-]\s*)(?<text>.+)$",
        RegexOptions.Compiled);

    private static readonly Regex TableFillPattern = new(
        @"(?i)^(?:fill(?:\s+the\s+table|\s+table|\s+it)?(?:\s+with)?|填写(?:本表|表)?|填入(?:本表|表)?|补充(?:本表|表)?|完善(?:本表|表)?|录入(?:本表|表)?|输入(?:本表|表)?|填报(?:本表|表)?)\s*(?:[:：-]\s*)?(?<text>.+)$",
        RegexOptions.Compiled);

    public static int RunPlan(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("plan requires <input.docx> <plan-data.json>");
        }

        var input = Path.GetFullPath(args[0]);
        var dataPath = Path.GetFullPath(args[1]);
        if (!File.Exists(dataPath))
        {
            throw new InvalidOperationException($"Plan data file not found: {dataPath}");
        }

        var json = File.ReadAllText(dataPath);
        var request = JsonSerializer.Deserialize<DocxPlanRequest>(json, Json.Options)
            ?? new DocxPlanRequest(null, null);
        var result = Plan(input, request);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    public static DocxPlanResult Plan(string input, DocxPlanRequest request)
    {
        var scenario = string.IsNullOrWhiteSpace(request.Scenario) ? SupportedScenario : request.Scenario.Trim();
        if (!scenario.StartsWith(SupportedScenario, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"Unsupported planning scenario: {scenario}");
        }

        var inspection = Inspector.Inspect(input);
        var items = new List<DocxPlanItem>();
        var warnings = new List<string>();
        var proposedEdits = new List<DocxEditOperation>();

        foreach (var anchor in inspection.Structure.AnnotationAnchors
            .OrderBy(anchor => anchor.ParagraphIndex ?? int.MaxValue)
            .ThenBy(anchor => anchor.TableIndex ?? int.MaxValue)
            .ThenBy(anchor => anchor.RowIndex ?? int.MaxValue)
            .ThenBy(anchor => anchor.CellIndex ?? int.MaxValue)
            .ThenBy(anchor => anchor.CommentId, StringComparer.Ordinal))
        {
            var item = ClassifyAnchor(anchor, inspection.Structure.Tables, request.SourceHints, warnings);
            items.Add(item);
            proposedEdits.AddRange(item.ProposedEdits);
        }

        var confidence = AggregateConfidence(items);
        return new DocxPlanResult(
            Path.GetFullPath(input),
            scenario,
            items,
            proposedEdits,
            warnings,
            confidence);
    }

    private static DocxPlanItem ClassifyAnchor(
        AnnotationAnchor anchor,
        IReadOnlyList<TableMetadata> tables,
        IReadOnlyList<string>? sourceHints,
        List<string> warnings)
    {
        var commentText = anchor.CommentText ?? string.Empty;
        var candidateTargets = new List<DocxPlanCandidateTarget>();
        var requiredSources = new List<string>();

        if (IsValidationComment(commentText))
        {
            candidateTargets.Add(BuildParagraphTarget(anchor, "Anchor paragraph"));
            if (anchor.TableIndex is not null)
            {
                candidateTargets.Add(BuildTableTarget(anchor, tables));
            }

            requiredSources.Add("human review");
            warnings.Add($"Comment {anchor.CommentId} stayed manual_only because it is a validation or review instruction.");
            return new DocxPlanItem(
                anchor.CommentId,
                commentText,
                anchor,
                "manual_only",
                anchor.TableIndex is null ? "paragraph" : "current_table",
                candidateTargets,
                requiredSources,
                "low",
                "The comment asks for validation or review rather than an edit.",
                []);
        }

        if (IsTableScopeComment(anchor, commentText))
        {
            var fillValues = ExtractTableFillValues(commentText).ToList();
            if (fillValues.Count == 0)
            {
                var reason = "The table instruction uses language that this first-pass parser cannot safely execute.";
                warnings.Add($"Comment {anchor.CommentId} fell back to manual_only: {reason}");
                return new DocxPlanItem(
                    anchor.CommentId,
                    commentText,
                    anchor,
                    "manual_only",
                    "current_table",
                    [
                        BuildTableTarget(anchor, tables),
                        BuildAnchorCellTarget(anchor)
                    ],
                    ["current table structure", "source data"],
                    "low",
                    reason,
                []);
            }

            candidateTargets.Add(BuildTableTarget(anchor, tables));
            candidateTargets.Add(BuildAnchorCellTarget(anchor));
            requiredSources.Add("current table structure");
            requiredSources.Add("source data");
            AppendSourceHints(requiredSources, sourceHints);

            var proposedEdits = BuildTableBlockEdits(anchor, fillValues, tables, out var tableFallbackReason);
            if (proposedEdits.Count == 0)
            {
                var reason = tableFallbackReason ?? "The comment is table-scoped, but the planner could not derive a safe explicit table edit.";
                warnings.Add($"Comment {anchor.CommentId} fell back to manual_only: {reason}");
                return new DocxPlanItem(
                    anchor.CommentId,
                    commentText,
                    anchor,
                    "manual_only",
                    "current_table",
                    candidateTargets,
                    requiredSources,
                    "low",
                    reason,
                    []);
            }

            return new DocxPlanItem(
                anchor.CommentId,
                commentText,
                anchor,
                "fill_table_block",
                "current_table",
                candidateTargets,
                requiredSources,
                "medium",
                "The comment sits in the table title or first cell and describes filling the table as a block.",
                proposedEdits);
        }

        if (IsSourceDrivenTableComment(anchor, commentText))
        {
            candidateTargets.Add(BuildTableTarget(anchor, tables));
            candidateTargets.Add(BuildAnchorCellTarget(anchor));
            requiredSources.Add("current table structure");
            requiredSources.Add("source data");
            AppendSourceHints(requiredSources, sourceHints);

            return new DocxPlanItem(
                anchor.CommentId,
                commentText,
                anchor,
                "fill_table_block",
                "current_table",
                candidateTargets,
                requiredSources,
                "medium",
                "The comment identifies table or column content that must be resolved from external sources before explicit cell edits are generated.",
                []);
        }

        if (IsSourceMappingComment(anchor, commentText))
        {
            candidateTargets.Add(BuildParagraphTarget(anchor, "Source mapping paragraph"));
            if (anchor.TableIndex is not null)
            {
                candidateTargets.Add(BuildTableTarget(anchor, tables));
                candidateTargets.Add(BuildAnchorCellTarget(anchor));
            }
            if (!string.IsNullOrWhiteSpace(anchor.NearestHeadingText))
            {
                candidateTargets.Add(BuildSectionTarget(anchor));
            }

            requiredSources.Add("comment");
            requiredSources.Add(anchor.TableIndex is null ? "source data" : "current table structure");
            AppendSourceHints(requiredSources, sourceHints);

            return new DocxPlanItem(
                anchor.CommentId,
                commentText,
                anchor,
                "source_mapping",
                anchor.TableIndex is null ? "paragraph" : "current_table",
                candidateTargets,
                requiredSources,
                "medium",
                "The comment identifies where source data should come from, but it does not request a direct edit.",
                []);
        }

        if (TryExtractExplicitText(commentText, out var explicitText))
        {
            var isNarrative = IsNarrativeComment(anchor, commentText);
            if (isNarrative)
            {
                candidateTargets.Add(BuildParagraphTarget(anchor, "Narrative paragraph"));
                if (!string.IsNullOrWhiteSpace(anchor.NearestHeadingText))
                {
                    candidateTargets.Add(BuildSectionTarget(anchor));
                }

                requiredSources.Add("comment");
                if (!string.IsNullOrWhiteSpace(anchor.NearestHeadingText))
                {
                    requiredSources.Add("nearest heading");
                }
                if (string.IsNullOrWhiteSpace(explicitText))
                {
                    warnings.Add($"Comment {anchor.CommentId} looks narrative but did not yield an explicit paragraph draft.");
                    return new DocxPlanItem(
                        anchor.CommentId,
                        commentText,
                        anchor,
                        "generate_paragraph",
                        anchor.NearestHeadingText is null ? "paragraph" : "section",
                        candidateTargets,
                        requiredSources,
                        "medium",
                        "The comment describes a generated paragraph, but the first-pass planner could not extract a safe draft.",
                        []);
                }

                return new DocxPlanItem(
                    anchor.CommentId,
                    commentText,
                    anchor,
                    "generate_paragraph",
                    anchor.NearestHeadingText is null ? "paragraph" : "section",
                    candidateTargets,
                    requiredSources,
                    "high",
                    "The comment requests a generated narrative and includes an explicit draft text.",
                    [
                        new DocxEditOperation(
                            "replaceParagraphText",
                            Text: explicitText,
                            ParagraphIndex: anchor.ParagraphIndex)
                    ]);
            }

            candidateTargets.Add(BuildParagraphTarget(anchor, "Anchor paragraph"));
            if (anchor.TableIndex is not null)
            {
                candidateTargets.Add(BuildTableTarget(anchor, tables));
            }

            requiredSources.Add("comment");
            requiredSources.Add("anchor paragraph");
            return new DocxPlanItem(
                anchor.CommentId,
                commentText,
                anchor,
                "replace_anchor_text",
                "paragraph",
                candidateTargets,
                requiredSources,
                "high",
                "The comment points at the local text and includes an explicit replacement string.",
                [
                    new DocxEditOperation(
                        "replaceAnchoredText",
                        CommentId: anchor.CommentId,
                        Text: explicitText)
                ]);
        }

        if (IsNarrativeComment(anchor, commentText))
        {
            candidateTargets.Add(BuildParagraphTarget(anchor, "Narrative paragraph"));
            if (!string.IsNullOrWhiteSpace(anchor.NearestHeadingText))
            {
                candidateTargets.Add(BuildSectionTarget(anchor));
            }

            requiredSources.Add("comment");
            if (!string.IsNullOrWhiteSpace(anchor.NearestHeadingText))
            {
                requiredSources.Add("nearest heading");
            }
            requiredSources.Add("source data");
            AppendSourceHints(requiredSources, sourceHints);
            warnings.Add($"Comment {anchor.CommentId} needs source context before a safe narrative edit can be proposed.");
            return new DocxPlanItem(
                anchor.CommentId,
                commentText,
                anchor,
                "generate_paragraph",
                anchor.NearestHeadingText is null ? "paragraph" : "section",
                candidateTargets,
                requiredSources,
                "medium",
                "The comment describes a narrative request that depends on resolved source data rather than an explicit draft in the comment itself.",
                []);
        }

        candidateTargets.Add(BuildParagraphTarget(anchor, "Anchor paragraph"));
        if (anchor.TableIndex is not null)
        {
            candidateTargets.Add(BuildTableTarget(anchor, tables));
        }

        requiredSources.Add("human review");
        warnings.Add($"Comment {anchor.CommentId} was left manual_only because the first-pass ANA03 heuristics could not classify it safely.");
        return new DocxPlanItem(
            anchor.CommentId,
            commentText,
            anchor,
            "manual_only",
            anchor.TableIndex is null ? "paragraph" : "current_table",
            candidateTargets,
            requiredSources,
            "low",
            "The comment is ambiguous or outside the supported first-pass ANA03 patterns.",
            []);
    }

    private static bool IsValidationComment(string commentText)
    {
        return ContainsAny(commentText,
            "需要判断",
            "需要进行判断",
            "需要注意",
            "请核对",
            "核对",
            "请确认",
            "确认",
            "请审核",
            "审核",
            "请复核",
            "复核",
            "检查",
            "验证",
            "校对",
            "仅供",
            "参考",
            "勿改",
            "不要修改",
            "无需修改",
            "保持原文",
            "保留原文",
            "留空",
            "人工",
            "手动",
            "自行填写",
            "另行",
            "待确认",
            "待核对",
            "不应少于",
            "不得延后",
            "不得提前",
            "时间格式");
    }

    private static bool IsSourceMappingComment(AnnotationAnchor anchor, string commentText)
    {
        if (!HasSourceMappingKeywords(commentText))
        {
            return IsShortSourceReferenceComment(anchor, commentText);
        }

        if (IsValidationComment(commentText))
        {
            return false;
        }

        if (!HasSourceScopeSignals(anchor, commentText))
        {
            return false;
        }

        return !IsTableScopeComment(anchor, commentText);
    }

    private static bool IsSourceDrivenTableComment(AnnotationAnchor anchor, string commentText)
    {
        if (anchor.TableIndex is null || IsValidationComment(commentText))
        {
            return false;
        }

        if (!HasSourceMappingKeywords(commentText) && !IsShortSourceReferenceComment(anchor, commentText))
        {
            return false;
        }

        return ContainsAny(commentText,
            "表格",
            "本表",
            "该表",
            "此表",
            "该列",
            "右侧",
            "接受标准",
            "检测结果",
            "检项",
            "顺序",
            "图谱",
            "quality standard",
            "result");
    }

    private static bool HasSourceMappingKeywords(string commentText)
        => ContainsAny(commentText,
            "来源",
            "取自",
            "来自",
            "依据",
            "根据",
            "参照",
            "对应",
            "源自",
            "摘自",
            "出自",
            "按",
            "按照",
            "获取",
            "附件",
            "reference",
            "source");

    private static bool IsShortSourceReferenceComment(AnnotationAnchor anchor, string commentText)
    {
        if (string.IsNullOrWhiteSpace(commentText) || IsValidationComment(commentText))
        {
            return false;
        }

        if (commentText.Length > 32)
        {
            return false;
        }

        if (anchor.TableIndex is not null)
        {
            return ContainsAny(commentText,
                "附件",
                "方案",
                "报告",
                "质量标准",
                "检验报告");
        }

        return ContainsAny(commentText,
            "附件",
            "方案",
            "报告",
            "质量标准",
            "检验报告",
            "汇总表");
    }

    private static bool HasSourceScopeSignals(AnnotationAnchor anchor, string commentText)
    {
        if (anchor.TableIndex is not null || anchor.TargetKind.Equals("tableCell", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return ContainsAny(commentText,
            "表",
            "数据",
            "附件",
            "方案",
            "结果",
            "样品",
            "段落",
            "章节",
            "报告",
            "记录");
    }

    private static bool IsTableScopeComment(AnnotationAnchor anchor, string commentText)
    {
        if (anchor.TableIndex is null)
        {
            return false;
        }

        if (anchor.RowIndex is int rowIndex && rowIndex > 0)
        {
            return false;
        }

        if (anchor.CellIndex is int cellIndex && cellIndex > 0)
        {
            return false;
        }

        if (HasSourceMappingKeywords(commentText) && !ContainsAny(commentText,
            "table",
            "fill",
            "填写",
            "填入",
            "补充",
            "完善",
            "录入",
            "输入",
            "填报",
            "complete"))
        {
            return false;
        }

        return ContainsAny(commentText,
            "table",
            "rows",
            "row",
            "columns",
            "column",
            "populate",
            "fill",
            "complete",
            "all entries",
            "entire table",
            "summary table",
            "data table",
            "填写",
            "填入",
            "补充",
            "完善",
            "录入",
            "输入",
            "填报",
            "本表",
            "该表",
            "此表",
            "整表",
            "表格",
            "汇总表",
            "数据表");
    }

    private static bool IsNarrativeComment(AnnotationAnchor anchor, string commentText)
    {
        if (ContainsAny(commentText, "paragraph", "summary", "narrative", "write", "draft", "generate", "describe", "report"))
        {
            return true;
        }

        if (ContainsAny(commentText, "段落", "总结", "叙述", "说明", "描述", "报告", "结论", "起草", "撰写", "生成", "编写", "改写"))
        {
            return true;
        }

        return anchor.NearestHeadingText is not null &&
            ContainsAny(commentText, "section", "overview", "summary", "report", "章节", "小结", "概述", "总结");
    }

    private static bool TryExtractExplicitText(string commentText, out string explicitText)
    {
        explicitText = string.Empty;
        if (string.IsNullOrWhiteSpace(commentText))
        {
            return false;
        }

        var quoted = Regex.Match(commentText, "\"(?<text>[^\"]+)\"");
        if (quoted.Success)
        {
            explicitText = quoted.Groups["text"].Value.Trim();
            return !string.IsNullOrWhiteSpace(explicitText);
        }

        var directMatch = DirectReplacementPattern.Match(commentText);
        if (directMatch.Success)
        {
            explicitText = directMatch.Groups["text"].Value.Trim();
            explicitText = explicitText.TrimEnd('.', ';');
            return !string.IsNullOrWhiteSpace(explicitText);
        }

        var narrativeMatch = NarrativeDraftPattern.Match(commentText);
        if (narrativeMatch.Success)
        {
            explicitText = narrativeMatch.Groups["text"].Value.Trim();
            explicitText = explicitText.TrimEnd('.', ';');
            return !string.IsNullOrWhiteSpace(explicitText);
        }

        return false;
    }

    private static IReadOnlyList<DocxEditOperation> BuildTableBlockEdits(
        AnnotationAnchor anchor,
        IReadOnlyList<string> fillValues,
        IReadOnlyList<TableMetadata> tables,
        out string? failureReason)
    {
        failureReason = null;
        if (fillValues.Count == 0)
        {
            failureReason = "no safe table edit could be derived from the comment text";
            return [];
        }

        var tableInfo = anchor.TableIndex is null || anchor.TableIndex < 0 || anchor.TableIndex >= tables.Count
            ? null
            : tables[anchor.TableIndex.Value];
        var rowIndex = anchor.RowIndex ?? 0;
        var startCellIndex = anchor.CellIndex ?? 0;
        var rowWidth = GetRowWidth(tableInfo, rowIndex, fillValues.Count);
        var rowCellCount = GetRowCellCount(tableInfo, rowIndex, fillValues.Count);
        if (rowWidth <= 0 || rowCellCount <= 0 || rowWidth != rowCellCount || startCellIndex < 0 || startCellIndex >= rowCellCount || startCellIndex + fillValues.Count > rowCellCount)
        {
            failureReason = rowWidth != rowCellCount
                ? "the table row contains spans or omitted cells, so the planner cannot safely map raw table cell edits"
                : "the comment implies more table values than fit the available cells";
            return [];
        }

        var edits = new List<DocxEditOperation>();

        for (var offset = 0; offset < fillValues.Count; offset++)
        {
            edits.Add(new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: anchor.TableIndex,
                RowIndex: rowIndex,
                CellIndex: startCellIndex + offset,
                Text: fillValues[offset]));
        }

        return edits;
    }

    private static int GetRowWidth(TableMetadata? tableInfo, int rowIndex, int fallbackWidth)
    {
        if (tableInfo is null)
        {
            return fallbackWidth;
        }

        if (rowIndex >= 0 && rowIndex < tableInfo.RowWidths.Count)
        {
            return tableInfo.RowWidths[rowIndex];
        }

        return tableInfo.ColumnCount;
    }

    private static int GetRowCellCount(TableMetadata? tableInfo, int rowIndex, int fallbackCellCount)
    {
        if (tableInfo is null)
        {
            return fallbackCellCount;
        }

        if (rowIndex >= 0 && rowIndex < tableInfo.RowCellCounts.Count)
        {
            return tableInfo.RowCellCounts[rowIndex];
        }

        return tableInfo.ColumnCount;
    }

    private static IEnumerable<string> ExtractTableFillValues(string commentText)
    {
        if (TryExtractExplicitText(commentText, out var explicitText))
        {
            var split = SplitTableValues(explicitText);
            if (split.Count > 0)
            {
                return split;
            }

            return [explicitText];
        }

        var match = TableFillPattern.Match(commentText);
        if (!match.Success)
        {
            return [];
        }

        var tableText = match.Groups["text"].Value.Trim().TrimEnd('.', ';');
        var values = SplitTableValues(tableText);
        return values.Count > 0 ? values : [tableText];
    }

    private static IReadOnlyList<string> SplitTableValues(string text)
    {
        var normalized = text
            .Replace('｜', '|')
            .Replace('；', ';')
            .Replace('、', '|')
            .Replace('，', ',');

        var separators = new[] { "|", ";", "\n", "\r\n" };
        var values = normalized.Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToList();
        return values;
    }

    private static bool ContainsAny(string text, params string[] values)
    {
        foreach (var value in values)
        {
            if (text.Contains(value, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static DocxPlanCandidateTarget BuildParagraphTarget(AnnotationAnchor anchor, string description)
        => new(
            Kind: "paragraph",
            Description: description,
            ParagraphIndex: anchor.ParagraphIndex);

    private static DocxPlanCandidateTarget BuildSectionTarget(AnnotationAnchor anchor)
        => new(
            Kind: "section",
            Description: anchor.NearestHeadingText ?? "Section scope",
            ParagraphIndex: anchor.ParagraphIndex);

    private static DocxPlanCandidateTarget BuildAnchorCellTarget(AnnotationAnchor anchor)
        => new(
            Kind: "table_cell",
            Description: "Anchor cell",
            ParagraphIndex: anchor.ParagraphIndex,
            TableIndex: anchor.TableIndex,
            RowIndex: anchor.RowIndex,
            CellIndex: anchor.CellIndex);

    private static DocxPlanCandidateTarget BuildTableTarget(AnnotationAnchor anchor, IReadOnlyList<TableMetadata> tables)
    {
        var tableInfo = anchor.TableIndex is null || anchor.TableIndex < 0 || anchor.TableIndex >= tables.Count
            ? null
            : tables[anchor.TableIndex.Value];

        var preview = tableInfo is null
            ? "Current table"
            : $"Table {tableInfo.TableIndex} preview: {FormatPreview(tableInfo)}";

        return new DocxPlanCandidateTarget(
            Kind: "current_table",
            Description: preview,
            TableIndex: anchor.TableIndex,
            RowCount: tableInfo?.RowCount ?? anchor.CurrentTableRowCount,
            ColumnCount: tableInfo?.ColumnCount ?? anchor.CurrentTableColumnCount);
    }

    private static string FormatPreview(TableMetadata table)
    {
        var rows = table.PreviewRows
            .Select(row => $"[{string.Join(" | ", row)}]")
            .ToList();
        return rows.Count == 0 ? "no preview available" : string.Join("; ", rows);
    }

    private static void AppendSourceHints(List<string> requiredSources, IReadOnlyList<string>? sourceHints)
    {
        if (sourceHints is null)
        {
            return;
        }

        foreach (var hint in sourceHints)
        {
            if (!string.IsNullOrWhiteSpace(hint) && !requiredSources.Contains(hint, StringComparer.OrdinalIgnoreCase))
            {
                requiredSources.Add(hint);
            }
        }
    }

    private static string AggregateConfidence(IReadOnlyList<DocxPlanItem> items)
    {
        if (items.Count == 0)
        {
            return "medium";
        }

        if (items.Any(item => string.Equals(item.Confidence, "low", StringComparison.OrdinalIgnoreCase)))
        {
            return "low";
        }

        if (items.Any(item => string.Equals(item.Confidence, "medium", StringComparison.OrdinalIgnoreCase)))
        {
            return "medium";
        }

        return "high";
    }
}
