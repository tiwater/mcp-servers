using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Docx;

namespace Dockit.Docx.Tests;

public class AnnotationToolsTests
{
    [Fact]
    public void Inspect_includes_annotation_anchors_in_unified_report()
    {
        var docPath = CreateAnnotatedFixture();

        var report = Inspector.Inspect(docPath);

        Assert.Equal(2, report.Annotations.CommentCount);
        Assert.Equal(2, report.Structure.AnnotationAnchors.Count);

        var paragraphAnchor = Assert.Single(report.Structure.AnnotationAnchors, anchor => anchor.CommentId == "0");
        Assert.Equal("paragraph", paragraphAnchor.TargetKind);
        Assert.Contains("Project code XXXX", paragraphAnchor.AnchorText);
        Assert.Equal("value comes from summary sheet", paragraphAnchor.CommentText);

        var tableAnchor = Assert.Single(report.Structure.AnnotationAnchors, anchor => anchor.CommentId == "1");
        Assert.Equal("tableCell", tableAnchor.TargetKind);
        Assert.Equal(0, tableAnchor.TableIndex);
        Assert.Equal(0, tableAnchor.RowIndex);
        Assert.Equal(1, tableAnchor.CellIndex);
        Assert.Contains("Batch YYYY", tableAnchor.AnchorText);
    }

    [Fact]
    public void Edit_applies_explicit_operations_and_preserves_other_content()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceAnchoredText", CommentId: "0", Text: "Project code HSP001"),
            new DocxEditOperation("replaceParagraphText", ParagraphIndex: 1, Text: "Top-level paragraph HSP001"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 1, Text: "Batch HSP001-01"),
            new DocxEditOperation("markFieldsDirty")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var topParagraph = body.Elements<Paragraph>().First();
        Assert.Contains("Project code HSP001", GetParagraphText(topParagraph));

        var tableCellParagraph = body.Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().ElementAt(1)
            .Elements<Paragraph>().Single();
        Assert.Contains("Batch HSP001-01", GetParagraphText(tableCellParagraph));

        var topLevelParagraphs = body.Elements<Paragraph>().ToList();
        Assert.Contains("Top-level paragraph HSP001", GetParagraphText(topLevelParagraphs[1]));
        Assert.DoesNotContain("Top-level paragraph HSP001", string.Concat(body.Elements<Table>().Single().Descendants<Text>().Select(text => text.Text)));
        Assert.True(doc.MainDocumentPart.DocumentSettingsPart?.Settings?.Elements<UpdateFieldsOnOpen>().Any() == true);
    }

    [Fact]
    public void Edit_can_replace_header_paragraph_text()
    {
        var docPath = CreateSplitPlaceholderFixture();
        var output = Path.Combine(Path.GetTempPath(), $"header-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceAllHeaderParagraphText", ParagraphIndex: 0, Text: "XX（客户项目代号）（与报告中HSPTEST对应）")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var headerText = string.Concat(
            doc.MainDocumentPart!.HeaderParts.SelectMany(part => part.Header!.Descendants<Text>()).Select(text => text.Text));
        Assert.Contains("XX（客户项目代号）（与报告中HSPTEST对应）", headerText);
        Assert.DoesNotContain("Header date", headerText);
    }

    [Fact]
    public void Edit_can_replace_header_text_without_overwriting_other_header_content()
    {
        var docPath = CreateHeaderLayoutFixture();
        var output = Path.Combine(Path.GetTempPath(), $"header-text-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceHeaderText", FindText: "XX（客户项目代号）（与报告中HSPTEST对应）", Text: "HSPTEST")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var headerParagraph = doc.MainDocumentPart!.HeaderParts.Single().Header!.Elements<Paragraph>().Single();
        var headerText = string.Concat(headerParagraph.Descendants<Text>().Select(text => text.Text));
        Assert.Contains("HSPTEST", headerText);
        Assert.Contains("3.2.S.7 稳定性", headerText);
        Assert.Contains("SN0000", headerText);
        Assert.DoesNotContain("XX（客户项目代号）（与报告中HSPTEST对应）", headerText);
        Assert.True(headerParagraph.Descendants<TabChar>().Count() >= 2);
    }

    [Fact]
    public void Edit_can_delete_comments_explicitly()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"clean-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("deleteComments"),
            new DocxEditOperation("markFieldsDirty")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var mainPart = doc.MainDocumentPart!;
        Assert.Null(mainPart.WordprocessingCommentsPart);
        Assert.Empty(mainPart.Document!.Descendants<CommentRangeStart>());
        Assert.Empty(mainPart.Document.Descendants<CommentRangeEnd>());
        Assert.Empty(mainPart.Document.Descendants<CommentReference>());
    }

    [Fact]
    public void ExportJson_includes_body_paragraph_and_table_indexes()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"export-{Guid.NewGuid():N}.json");

        Transforms.RunExportJson([docPath, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("Project code XXXX 峰面积", json, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\u5CF0", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u9762", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u79EF", json, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"paragraphIndex\": 0", json, StringComparison.Ordinal);
        Assert.Contains("\"tableIndex\": 0", json, StringComparison.Ordinal);
    }

    [Fact]
    public void ExportJson_includes_header_paragraphs()
    {
        var docPath = CreateSplitPlaceholderFixture();
        var output = Path.Combine(Path.GetTempPath(), $"export-header-{Guid.NewGuid():N}.json");

        Transforms.RunExportJson([docPath, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("\"type\": \"headerParagraph\"", json, StringComparison.Ordinal);
        Assert.Contains("\"headerIndex\": 0", json, StringComparison.Ordinal);
        Assert.Contains("Header date:", json, StringComparison.Ordinal);
    }

    [Fact]
    public void FillTemplate_replaces_split_placeholders_in_body_and_header()
    {
        var docPath = CreateSplitPlaceholderFixture();
        var dataPath = Path.Combine(Path.GetTempPath(), $"fill-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"filled-{Guid.NewGuid():N}.docx");

        File.WriteAllText(
            dataPath,
            """
            {
              "cellValues": {
                "effectiveDate": "2024-09-18"
              }
            }
            """,
            System.Text.Encoding.UTF8);

        Transforms.RunFillTemplate([docPath, dataPath, output]);

        var report = Inspector.Inspect(output);
        Assert.DoesNotContain("{{effectiveDate}}", report.Content.Placeholders);

        using var doc = WordprocessingDocument.Open(output, false);
        var bodyText = string.Concat(doc.MainDocumentPart!.Document!.Descendants<Text>().Select(text => text.Text));
        Assert.Contains("2024-09-18", bodyText);

        var headerText = string.Concat(
            doc.MainDocumentPart.HeaderParts.SelectMany(part => part.Header!.Descendants<Text>()).Select(text => text.Text));
        Assert.Contains("2024-09-18", headerText);
    }

    private static string CreateAnnotatedFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"annotated-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            CreateComment("0", "tester", "value comes from summary sheet"),
            CreateComment("1", "tester", "batch id comes from inspection report"));

        var body = mainPart.Document.Body!;
        body.Append(CreateParagraphWithComment("0", "Project code XXXX 峰面积"));
        body.Append(CreateTableWithComment());
        body.Append(CreateFieldParagraph());
        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static Paragraph CreateParagraphWithComment(string commentId, string text)
    {
        return new Paragraph(
            new CommentRangeStart { Id = commentId },
            new Run(new Text(text)),
            new CommentRangeEnd { Id = commentId },
            new Run(new CommentReference { Id = commentId }));
    }

    private static Table CreateTableWithComment()
    {
        return new Table(
            new TableRow(
                CreateCell("Label"),
                CreateCellWithComment("1", "Batch YYYY")));
    }

    private static TableCell CreateCell(string text)
        => new(new Paragraph(new Run(new Text(text))));

    private static TableCell CreateCellWithComment(string commentId, string text)
        => new(new Paragraph(
            new CommentRangeStart { Id = commentId },
            new Run(new Text(text)),
            new CommentRangeEnd { Id = commentId },
            new Run(new CommentReference { Id = commentId })));

    private static Paragraph CreateFieldParagraph()
    {
        return new Paragraph(
            new SimpleField { Instruction = "SEQ Figure \\* ARABIC", Dirty = false },
            new Run(new Text("Figure 1")));
    }

    private static string CreateSplitPlaceholderFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"split-placeholder-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(
            new Paragraph(
                new Run(new Text("Header date: ")),
                new Run(new Text("{{")),
                new Run(new Text("effectiveDate")),
                new Run(new Text("}}"))));

        var headerPartId = mainPart.GetIdOfPart(headerPart);
        var sectionProps = new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId });

        var body = mainPart.Document.Body!;
        body.Append(
            new Paragraph(
                new Run(new Text("Body date: ")),
                new Run(new Text("{{")),
                new Run(new Text("effectiveDate")),
                new Run(new Text("}}"))));
        body.Append(new Paragraph(new Run(new Text("after"))));
        body.Append(sectionProps);

        mainPart.Document.Save();
        headerPart.Header.Save();
        return path;
    }

    private static string CreateHeaderLayoutFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"header-layout-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(
            new Paragraph(
                new Run(new Text("XX（客户项目代号）（与报告中HSPTEST对应）")),
                new Run(new TabChar()),
                new Run(new Text("3.2.S.7 稳定性")),
                new Run(new TabChar()),
                new Run(new Text("SN0000"))));

        var headerPartId = mainPart.GetIdOfPart(headerPart);
        var sectionProps = new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId });
        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(new Run(new Text("body"))));
        body.Append(sectionProps);

        mainPart.Document.Save();
        headerPart.Header.Save();
        return path;
    }

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

    private static string GetParagraphText(Paragraph paragraph)
        => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
}
