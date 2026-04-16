using Dockit.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Dockit.Docx.Tests;

public class PlannerTests
{
    [Fact]
    public void Inspect_includes_richer_comment_context_and_table_metadata()
    {
        var docPath = CreateAna03Fixture();

        var report = Inspector.Inspect(docPath);

        Assert.Equal(5, report.Annotations.CommentCount);
        Assert.Single(report.Structure.Tables);

        var table = report.Structure.Tables[0];
        Assert.Equal(1, table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        Assert.Equal(2, table.RowWidths[0]);
        Assert.Equal(2, table.RowCellCounts[0]);
        Assert.Equal("Label", table.PreviewRows[0][0]);

        var direct = Assert.Single(report.Structure.AnnotationAnchors, anchor => anchor.CommentId == "0");
        Assert.Equal("mainDocument", direct.Source);
        Assert.Equal("Outline Section", direct.NearestHeadingText);
        Assert.Equal("Project code XXXX", direct.CurrentParagraphText);
        Assert.Equal("Outline Section", direct.PreviousParagraphText);
        Assert.Equal("Label", direct.FollowingParagraphText);

        var tableAnchor = Assert.Single(report.Structure.AnnotationAnchors, anchor => anchor.CommentId == "1");
        Assert.Equal("mainDocument", tableAnchor.Source);
        Assert.Equal(0, tableAnchor.TableIndex);
        Assert.Equal(0, tableAnchor.RowIndex);
        Assert.Equal(0, tableAnchor.CellIndex);
        Assert.Equal(1, tableAnchor.CurrentTableRowCount);
        Assert.Equal(2, tableAnchor.CurrentTableColumnCount);
    }

    [Fact]
    public void Plan_classifies_ana03_comment_patterns_and_emits_supported_edits()
    {
        var docPath = CreateAna03Fixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", ["source workbook", "notes"]));

        Assert.Equal("stability-report", plan.Scenario);
        Assert.Equal(5, plan.Items.Count);
        Assert.Equal(1, plan.ProposedEdits.Count(edit => edit.Type == "replaceAnchoredText"));
        Assert.Equal(1, plan.ProposedEdits.Count(edit => edit.Type == "replaceParagraphText"));
        Assert.Equal(2, plan.ProposedEdits.Count(edit => edit.Type == "replaceTableCellText"));
        Assert.Contains(plan.Items, item => item.InstructionType == "replace_anchor_text" && item.CommentId == "0");
        Assert.Contains(plan.Items, item => item.InstructionType == "fill_table_block" && item.CommentId == "1");
        Assert.Contains(plan.Items, item => item.InstructionType == "generate_paragraph" && item.CommentId == "2");
        Assert.Contains(plan.Items, item => item.InstructionType == "generate_paragraph" && item.CommentId == "3");

        var tableItem = Assert.Single(plan.Items, item => item.CommentId == "1");
        Assert.Equal("current_table", tableItem.TargetScope);
        Assert.Contains(tableItem.CandidateTargets, target => target.Kind == "current_table");
        Assert.Equal(2, tableItem.ProposedEdits.Count);
        Assert.All(tableItem.ProposedEdits, edit => Assert.Equal("replaceTableCellText", edit.Type));

        var directItem = Assert.Single(plan.Items, item => item.CommentId == "0");
        Assert.Single(directItem.ProposedEdits);
        Assert.Equal("replaceAnchoredText", directItem.ProposedEdits[0].Type);
        Assert.Equal("high", directItem.Confidence);

        var narrativeItem = Assert.Single(plan.Items, item => item.CommentId == "2");
        Assert.Single(narrativeItem.ProposedEdits);
        Assert.Equal("replaceParagraphText", narrativeItem.ProposedEdits[0].Type);
        Assert.Equal("generate_paragraph", narrativeItem.InstructionType);
        Assert.Contains(narrativeItem.RequiredSources, source => source == "nearest heading");

        var manualNarrativeItem = Assert.Single(plan.Items, item => item.CommentId == "3");
        Assert.Empty(manualNarrativeItem.ProposedEdits);
        Assert.Equal("generate_paragraph", manualNarrativeItem.InstructionType);
        Assert.Equal("medium", manualNarrativeItem.Confidence);

        var manualItem = Assert.Single(plan.Items, item => item.CommentId == "4");
        Assert.Empty(manualItem.ProposedEdits);
        Assert.Equal("low", manualItem.Confidence);
        Assert.Equal("manual_only", manualItem.InstructionType);
    }

    [Fact]
    public void Edit_can_apply_planner_emitted_supported_edits()
    {
        var docPath = CreateAna03Fixture();
        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));
        var output = Path.Combine(Path.GetTempPath(), $"ana03-edited-{Guid.NewGuid():N}.docx");
        var supportedEdits = plan.Items.SelectMany(item => item.ProposedEdits).ToList();

        var result = Editor.Apply(docPath, output, supportedEdits);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));

        using var doc = WordprocessingDocument.Open(output, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        Assert.Contains(body.Descendants<Paragraph>(), paragraph =>
            Inspector.GetParagraphText(paragraph).Contains("Project code HSP001", StringComparison.Ordinal));
        var tableRow = body.Elements<Table>().Single().Elements<TableRow>().Single();
        Assert.Contains(tableRow.Elements<TableCell>().First().Descendants<Paragraph>(),
            paragraph => Inspector.GetParagraphText(paragraph).Contains("Batch HSP001-01", StringComparison.Ordinal));
        Assert.Contains(tableRow.Elements<TableCell>().ElementAt(1).Descendants<Paragraph>(),
            paragraph => Inspector.GetParagraphText(paragraph).Contains("Released", StringComparison.Ordinal));
        Assert.Contains(body.Descendants<Paragraph>(), paragraph =>
            Inspector.GetParagraphText(paragraph).Contains("Summary completed successfully", StringComparison.Ordinal));
    }

    [Fact]
    public void Grid_spanned_rows_are_counted_by_effective_width()
    {
        var docPath = CreateMergedTableFixture();

        var report = Inspector.Inspect(docPath);
        var table = Assert.Single(report.Structure.Tables);

        Assert.Equal(3, table.ColumnCount);
        Assert.Equal(3, table.RowWidths[0]);
        Assert.Equal(2, table.RowCellCounts[0]);

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));
        var item = Assert.Single(plan.Items, item => item.CommentId == "20");

        Assert.Equal("manual_only", item.InstructionType);
        Assert.Empty(item.ProposedEdits);

        var output = Path.Combine(Path.GetTempPath(), $"merged-safety-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(docPath, output, plan.ProposedEdits);
        Assert.Empty(result.AppliedOperations);
    }

    [Fact]
    public void Non_first_cell_table_anchors_fall_back_safely()
    {
        var docPath = CreateNonFirstCellTableFixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));
        var item = Assert.Single(plan.Items, item => item.CommentId == "30");

        Assert.Equal("manual_only", item.InstructionType);
        Assert.Empty(item.ProposedEdits);
        Assert.DoesNotContain(plan.ProposedEdits, edit => edit.Type == "replaceTableCellText");
    }

    [Fact]
    public void Sparse_rows_with_omitted_grid_cells_fall_back_safely()
    {
        var docPath = CreateSparseTableFixture();

        var report = Inspector.Inspect(docPath);
        var table = Assert.Single(report.Structure.Tables);
        Assert.Equal(4, table.RowWidths[0]);
        Assert.Equal(2, table.RowCellCounts[0]);

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));
        var item = Assert.Single(plan.Items, item => item.CommentId == "50");

        Assert.Equal("manual_only", item.InstructionType);
        Assert.Empty(item.ProposedEdits);
        Assert.DoesNotContain(plan.ProposedEdits, edit => edit.Type == "replaceTableCellText");
        Assert.Contains(plan.Warnings, warning => warning.Contains("manual_only", StringComparison.Ordinal));
    }

    [Fact]
    public void Invalid_planning_scenario_is_rejected()
    {
        var docPath = CreateAna03Fixture();

        Assert.Throws<InvalidOperationException>(() => Planner.Plan(docPath, new DocxPlanRequest("not-supported", null)));
    }

    [Fact]
    public void Narrative_plans_only_require_nearest_heading_when_one_exists()
    {
        var docPath = CreateNarrativeNoHeadingFixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));
        var item = Assert.Single(plan.Items, item => item.CommentId == "40");

        Assert.Equal("generate_paragraph", item.InstructionType);
        Assert.DoesNotContain(item.RequiredSources, source => source == "nearest heading");
    }

    [Fact]
    public void Table_block_edge_cases_fall_back_to_manual_only_without_truncation()
    {
        var docPath = CreateTableEdgeCaseFixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));

        var overflow = Assert.Single(plan.Items, item => item.CommentId == "10");
        Assert.Equal("manual_only", overflow.InstructionType);
        Assert.Empty(overflow.ProposedEdits);
        Assert.NotEqual("fill_table_block", overflow.InstructionType);
        Assert.Contains(plan.Warnings, warning => warning.Contains("manual_only", StringComparison.Ordinal));
        Assert.DoesNotContain(plan.ProposedEdits, edit => edit.Type == "replaceTableCellText");

        var empty = Assert.Single(plan.Items, item => item.CommentId == "11");
        Assert.Equal("manual_only", empty.InstructionType);
        Assert.Empty(empty.ProposedEdits);
    }

    [Fact]
    public void Broad_table_scope_language_without_safe_payload_falls_back_to_manual_only()
    {
        var docPath = CreateBroadTableScopeFallbackFixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", null));

        var item = Assert.Single(plan.Items, item => item.CommentId == "12");
        Assert.Equal("manual_only", item.InstructionType);
        Assert.NotEqual("fill_table_block", item.InstructionType);
        Assert.Empty(item.ProposedEdits);
        Assert.DoesNotContain(plan.ProposedEdits, edit => edit.Type == "replaceTableCellText");
    }

    [Fact]
    public void Chinese_comment_patterns_are_classified_safely_without_falling_back_everywhere()
    {
        var docPath = CreateChineseCommentFixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", ["稳定性数据汇总表", "结果说明"]));

        var sourceMapping = Assert.Single(plan.Items, item => item.CommentId == "60");
        Assert.Equal("source_mapping", sourceMapping.InstructionType);
        Assert.Equal("medium", sourceMapping.Confidence);
        Assert.Empty(sourceMapping.ProposedEdits);
        Assert.Contains(sourceMapping.CandidateTargets, target => target.Kind == "paragraph");
        Assert.Contains(sourceMapping.RequiredSources, source => source == "current table structure");

        var tableScope = Assert.Single(plan.Items, item => item.CommentId == "61");
        Assert.Equal("fill_table_block", tableScope.InstructionType);
        Assert.Equal("medium", tableScope.Confidence);
        Assert.Equal(3, tableScope.ProposedEdits.Count);
        Assert.All(tableScope.ProposedEdits, edit => Assert.Equal("replaceTableCellText", edit.Type));
        Assert.Contains(tableScope.CandidateTargets, target => target.Kind == "current_table");

        var narrative = Assert.Single(plan.Items, item => item.CommentId == "62");
        Assert.Equal("generate_paragraph", narrative.InstructionType);
        Assert.Single(narrative.ProposedEdits);
        Assert.Equal("replaceParagraphText", narrative.ProposedEdits[0].Type);
        Assert.Contains(narrative.RequiredSources, source => source == "nearest heading");

        var validation = Assert.Single(plan.Items, item => item.CommentId == "63");
        Assert.Equal("manual_only", validation.InstructionType);
        Assert.Equal("low", validation.Confidence);
        Assert.Empty(validation.ProposedEdits);
    }

    [Fact]
    public void Chinese_ana03_comments_are_classified_without_falling_back_entirely_to_manual_only()
    {
        var docPath = CreateChineseAna03Fixture();

        var plan = Planner.Plan(docPath, new DocxPlanRequest("stability-report", ["稳定性数据汇总表", "稳定性方案"]));

        Assert.Equal("stability-report", plan.Scenario);
        Assert.Equal(4, plan.Items.Count);

        var sourceMappedParagraph = Assert.Single(plan.Items, item => item.CommentId == "60");
        Assert.Equal("source_mapping", sourceMappedParagraph.InstructionType);
        Assert.Equal("paragraph", sourceMappedParagraph.TargetScope);
        Assert.Equal("medium", sourceMappedParagraph.Confidence);
        Assert.Empty(sourceMappedParagraph.ProposedEdits);
        Assert.Contains(sourceMappedParagraph.RequiredSources, source => source.Contains("source", StringComparison.OrdinalIgnoreCase));

        var sourceMappedTable = Assert.Single(plan.Items, item => item.CommentId == "61");
        Assert.Equal("fill_table_block", sourceMappedTable.InstructionType);
        Assert.Equal("current_table", sourceMappedTable.TargetScope);
        Assert.Equal("medium", sourceMappedTable.Confidence);
        Assert.Empty(sourceMappedTable.ProposedEdits);
        Assert.Contains(sourceMappedTable.CandidateTargets, target => target.Kind == "current_table");

        var generatedNarrative = Assert.Single(plan.Items, item => item.CommentId == "62");
        Assert.Equal("generate_paragraph", generatedNarrative.InstructionType);
        Assert.Equal("section", generatedNarrative.TargetScope);
        Assert.Equal("medium", generatedNarrative.Confidence);
        Assert.Empty(generatedNarrative.ProposedEdits);

        var validationOnly = Assert.Single(plan.Items, item => item.CommentId == "63");
        Assert.Equal("manual_only", validationOnly.InstructionType);
        Assert.Equal("low", validationOnly.Confidence);
        Assert.Empty(validationOnly.ProposedEdits);
    }

    private static string CreateAna03Fixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            CreateComment("0", "reviewer", "replace with Project code HSP001"),
            CreateComment("1", "reviewer", "fill the table with Batch HSP001-01 | Released"),
            CreateComment("2", "reviewer", "generate summary paragraph: Summary completed successfully"),
            CreateComment("3", "reviewer", "generate summary paragraph for this section"),
            CreateComment("4", "reviewer", "insert diagram from the slide deck"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Executive Summary"))));
        body.Append(new Paragraph(
            new ParagraphProperties(new OutlineLevel { Val = 0 }),
            new Run(new Text("Outline Section"))));
        body.Append(CreateParagraphWithComment("0", "Project code XXXX"));
        body.Append(CreateTableWithComment());
        body.Append(CreateParagraphWithComment("2", "Summary TBD"));
        body.Append(CreateParagraphWithComment("3", "Closing note TBD"));
        body.Append(CreateParagraphWithComment("4", "Image placeholder"));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateTableEdgeCaseFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-table-edge-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            CreateComment("10", "reviewer", "fill the table with Alpha | Beta | Gamma"),
            CreateComment("11", "reviewer", "fill the table"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Edge Cases"))));
        body.Append(CreateTableWithComment("10", "11"));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateBroadTableScopeFallbackFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-broad-table-fallback-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(CreateComment("12", "reviewer", "populate the table with Alpha | Beta"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Broad Table Scope"))));
        body.Append(CreateTableWithSingleComment("12"));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateChineseCommentFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-chinese-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            CreateComment("60", "reviewer", "来源：稳定性数据汇总表中的 280nm 数据"),
            CreateComment("61", "reviewer", "填写本表：样品编号 | 280nm | 360nm"),
            CreateComment("62", "reviewer", "生成结果说明：结果符合要求"),
            CreateComment("63", "reviewer", "请核对本段，不要修改格式"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Chinese Comments"))));
        body.Append(CreateTableWithSingleComment("60"));
        body.Append(CreateTableWithComment("61"));
        body.Append(CreateParagraphWithComment("62", "结果说明待补充"));
        body.Append(CreateParagraphWithComment("63", "格式保持不变"));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateNarrativeNoHeadingFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-narrative-noheading-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(CreateComment("40", "reviewer", "generate summary paragraph: Summary completed successfully"));

        var body = mainPart.Document.Body!;
        body.Append(CreateParagraphWithComment("40", "Summary TBD"));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateMergedTableFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-merged-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(CreateComment("20", "reviewer", "fill the table with Alpha | Beta"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Merged"))));
        body.Append(new Table(
            new TableRow(
                CreateMergedCellWithComment("20", "Merged Label", 2),
                CreateCell("Trailing"))));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateNonFirstCellTableFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-nonfirst-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(CreateComment("30", "reviewer", "fill the table with Alpha | Beta"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("NonFirst"))));
        body.Append(new Table(
            new TableRow(
                CreateCell("Leading"),
                CreateCellWithComment("30", "Target"))));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateSparseTableFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-sparse-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(CreateComment("50", "reviewer", "fill the table with Alpha | Beta"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("Sparse"))));
        body.Append(new Table(
            new TableRow(
                new TableRowProperties(new GridBefore { Val = 1 }, new GridAfter { Val = 1 }),
                CreateCellWithComment("50", "Label"),
                CreateCell("Batch YYYY"))));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateChineseAna03Fixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"ana03-zh-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            CreateComment("60", "reviewer", "项目号XXXX和批号YYYY来自于稳定性数据汇总表"),
            CreateComment("61", "reviewer", "右侧三列检测结果，来源于稳定性数据汇总表，根据具体条件和时间点进行索引"),
            CreateComment("62", "reviewer", "结果描述，基于表格中数据。需要将最后时间点和零时间点进行数据对比，并给出结论"),
            CreateComment("63", "reviewer", "需要进行判断效期填写是否正确，根据稳定性方案进行核对"));

        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new Run(new Text("结果与讨论"))));
        body.Append(CreateParagraphWithComment("60", "本报告汇总HSPXXXX工程批原液（YYYY）影响因素试验结果。"));
        body.Append(new Table(
            new TableRow(
                CreateCell("检项"),
                CreateCell("接受标准"),
                CreateCellWithComment("61", "检测结果"),
                CreateCell("备注")),
            new TableRow(
                CreateCell("SEC"),
                CreateCell("符合质量标准"),
                CreateCell("TBD"),
                CreateCell("TBD"))));
        body.Append(CreateParagraphWithComment("62", "结果描述待补充"));
        body.Append(CreateParagraphWithComment("63", "校验有效期至"));

        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static Paragraph CreateParagraphWithComment(string commentId, string text)
        => new(
            new CommentRangeStart { Id = commentId },
            new Run(new Text(text)),
            new CommentRangeEnd { Id = commentId },
            new Run(new CommentReference { Id = commentId }));

    private static Table CreateTableWithComment()
        => new(
            new TableRow(
                CreateCellWithComment("1", "Label"),
                CreateCell("Batch YYYY")));

    private static Table CreateTableWithComment(string overflowCommentId, string emptyCommentId)
        => new(
            new TableRow(
                CreateCellWithComment(overflowCommentId, "Label"),
                CreateCellWithComment(emptyCommentId, "Batch YYYY")));

    private static Table CreateTableWithSingleComment(string commentId)
        => new(
            new TableRow(
                CreateCellWithComment(commentId, "Label"),
                CreateCell("Batch YYYY")));

    private static Table CreateTableWithComment(string commentId)
        => new(
            new TableRow(
                CreateCellWithComment(commentId, "样品编号"),
                CreateCell("280nm"),
                CreateCell("360nm")));

    private static TableCell CreateCell(string text)
        => new(new Paragraph(new Run(new Text(text))));

    private static TableCell CreateCellWithComment(string commentId, string text)
        => new(new Paragraph(
            new CommentRangeStart { Id = commentId },
            new Run(new Text(text)),
            new CommentRangeEnd { Id = commentId },
            new Run(new CommentReference { Id = commentId })));

    private static TableCell CreateMergedCellWithComment(string commentId, string text, int gridSpan)
        => new(
            new TableCellProperties(new GridSpan { Val = gridSpan }),
            new Paragraph(
                new CommentRangeStart { Id = commentId },
                new Run(new Text(text)),
                new CommentRangeEnd { Id = commentId },
                new Run(new CommentReference { Id = commentId })));

    private static Comment CreateComment(string id, string author, string text)
    {
        var comment = new Comment
        {
            Id = id,
            Author = author,
            Initials = author,
            Date = DateTime.Parse("2026-04-15T00:00:00Z")
        };

        comment.Append(new Paragraph(new Run(new Text(text))));
        return comment;
    }
}
