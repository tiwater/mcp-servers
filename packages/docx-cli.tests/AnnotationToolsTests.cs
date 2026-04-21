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
        Assert.True(doc.MainDocumentPart.DocumentSettingsPart?.Settings?.Elements<UpdateFieldsOnOpen>().Any() == true);
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
        Assert.Contains("\"paragraphIndex\": 0", json, StringComparison.Ordinal);
        Assert.Contains("\"tableIndex\": 0", json, StringComparison.Ordinal);
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
        body.Append(CreateParagraphWithComment("0", "Project code XXXX"));
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
