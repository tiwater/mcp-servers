using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class DrawingEditTests
{
    [Fact]
    public void Deletes_the_single_drawing_only_body_paragraph_immediately_before_a_unique_anchor()
    {
        WithFixture(
            new Body(
                TextParagraph("retained before"),
                DrawingParagraph(),
                TextParagraph("Figure 1: current caption"),
                TextParagraph("retained after")),
            (input, output) =>
            {
                var result = Editor.Apply(
                    input,
                    output,
                    [new DocxEditOperation("deleteBodyDrawingBeforeParagraph", FindText: "Figure 1: current caption")]);

                Assert.True(Assert.Single(result.AppliedOperations).Applied);
                using var edited = WordprocessingDocument.Open(output, false);
                var paragraphs = edited.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
                Assert.Equal(["retained before", "Figure 1: current caption", "retained after"], paragraphs.Select(p => p.InnerText));
                Assert.Empty(edited.MainDocumentPart.Document.Body.Descendants<Drawing>());
            });
    }

    [Fact]
    public void Supports_starts_with_and_style_bound_anchor_selection()
    {
        var anchor = TextParagraph("Figure 2: variable suffix");
        anchor.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Caption" });
        WithFixture(
            new Body(DrawingParagraph(), anchor),
            (input, output) =>
            {
                var result = Editor.Apply(
                    input,
                    output,
                    [new DocxEditOperation(
                        "deleteBodyDrawingBeforeParagraph",
                        FindText: "Figure 2:",
                        MatchMode: "startsWith",
                        ParagraphStyle: "Caption")]);

                Assert.True(Assert.Single(result.AppliedOperations).Applied);
                Assert.Equal(0, Inspector.Inspect(output).Structure.DrawingCount);
            });
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("noDrawing")]
    [InlineData("drawingWithText")]
    [InlineData("multipleDrawings")]
    [InlineData("notDirectBodyChild")]
    public void Fails_closed_when_the_anchor_or_preceding_drawing_paragraph_is_not_unique_and_safe(string mutation)
    {
        var body = mutation switch
        {
            "missing" => new Body(DrawingParagraph(), TextParagraph("Other caption")),
            "duplicate" => new Body(DrawingParagraph(), TextParagraph("Figure 1"), DrawingParagraph(), TextParagraph("Figure 1")),
            "noDrawing" => new Body(TextParagraph("ordinary paragraph"), TextParagraph("Figure 1")),
            "drawingWithText" => new Body(DrawingParagraph("unexpected text"), TextParagraph("Figure 1")),
            "multipleDrawings" => new Body(DrawingParagraph(drawingCount: 2), TextParagraph("Figure 1")),
            "notDirectBodyChild" => new Body(
                new Table(new TableRow(new TableCell(DrawingParagraph(), TextParagraph("Figure 1"))))),
            _ => throw new ArgumentOutOfRangeException(nameof(mutation)),
        };

        WithFixture(
            body,
            (input, output) =>
            {
                var before = Inspector.Inspect(input).Structure.DrawingCount;
                var result = Editor.Apply(
                    input,
                    output,
                    [new DocxEditOperation("deleteBodyDrawingBeforeParagraph", FindText: "Figure 1")]);

                Assert.False(Assert.Single(result.AppliedOperations).Applied);
                Assert.Equal(before, Inspector.Inspect(output).Structure.DrawingCount);
            });
    }

    [Fact]
    public void Preserves_drawings_outside_the_selected_body_paragraph_and_exports_body_drawing_evidence()
    {
        WithFixture(
            new Body(
                DrawingParagraph(),
                TextParagraph("unrelated figure"),
                DrawingParagraph(),
                TextParagraph("Figure 1")),
            (input, output) =>
            {
                using (var document = WordprocessingDocument.Open(input, true))
                {
                    var header = document.MainDocumentPart!.AddNewPart<HeaderPart>();
                    header.Header = new Header(DrawingParagraph());
                    header.Header.Save();
                }

                var inputFlow = System.Text.Json.JsonSerializer.Serialize(Inspector.InspectDocumentFlow(input));
                Assert.Contains("\"drawingCount\":1", inputFlow, StringComparison.Ordinal);

                var result = Editor.Apply(
                    input,
                    output,
                    [new DocxEditOperation("deleteBodyDrawingBeforeParagraph", FindText: "Figure 1")]);

                Assert.True(Assert.Single(result.AppliedOperations).Applied);
                Assert.Equal(2, Inspector.Inspect(output).Structure.DrawingCount);
                var outputFlow = System.Text.Json.JsonSerializer.Serialize(Inspector.InspectDocumentFlow(output));
                Assert.Equal(1, CountOccurrences(outputFlow, "\"drawingCount\":1"));
            });
    }

    private static Paragraph TextParagraph(string text) => new(new Run(new Text(text)));

    private static Paragraph DrawingParagraph(string? text = null, int drawingCount = 1)
    {
        var paragraph = new Paragraph();
        for (var index = 0; index < drawingCount; index++)
        {
            paragraph.Append(new Run(new Drawing(new DW.Inline())));
        }
        if (text is not null) paragraph.Append(new Run(new Text(text)));
        return paragraph;
    }

    private static int CountOccurrences(string text, string value)
        => text.Split(value, StringSplitOptions.None).Length - 1;

    private static void WithFixture(Body body, Action<string, string> assertion)
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-drawing-edit-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(body);
                main.Document.Save();
            }

            assertion(input, output);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
