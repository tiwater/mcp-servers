using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class FontPolicyTests
{
    [Fact]
    public void Preserve_size_changes_only_font_channels_and_validation_accepts_the_result()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-font-preserve-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "等线", ComplexScript = "Calibri" },
                            new FontSize { Val = "24" },
                            new FontSizeComplexScript { Val = "26" },
                            new Bold()),
                        new Text("正文 Body"))),
                    new Table(new TableRow(new TableCell(new Paragraph(
                        new Run(new RunProperties(new FontSize { Val = "17" })),
                        new Run(
                        new RunProperties(
                            new RunFonts { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "黑体", ComplexScript = "Arial" },
                            new FontSize { Val = "19" },
                            new FontSizeComplexScript { Val = "21" },
                            new Italic()),
                        new Text("表格 Table"))))))));
                main.Document.Save();
            }

            var policy = new DocxFontPolicy(
                FontPolicy.Schema,
                new DocxFontRule("宋体", "Times New Roman", "preserve"),
                new DocxFontRule("宋体", "Times New Roman", "preserve"));
            var result = Editor.Apply(input, output, [new DocxEditOperation("applyDocumentFontPolicy", FontPolicy: policy)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            using (var document = WordprocessingDocument.Open(output, false))
            {
                var runs = document.MainDocumentPart!.Document!.Body!.Descendants<Run>().ToList();
                Assert.Equal(3, runs.Count);
                AssertRun(runs[0], "24", "26");
                Assert.NotNull(runs[0].RunProperties!.Bold);
                Assert.Equal("17", runs[1].RunProperties!.FontSize!.Val!.Value);
                AssertRun(runs[2], "19", "21");
                Assert.NotNull(runs[2].RunProperties!.Italic);
            }

            Assert.True(FontPolicy.Validate(output, policy, "policy-hash").Pass);
            Assert.False(FontPolicy.Validate(input, policy, "policy-hash").Pass);
            var inspection = FontPolicy.Inspect(output);
            Assert.Equal("tiwater.docx-font-inspection/v2", inspection.Schema);
            Assert.Collection(inspection.Runs,
                run =>
                {
                    Assert.Equal("body:paragraph:0", run.Container);
                    Assert.Equal(0, run.RunIndex);
                    Assert.Equal("正文 Body", run.Text);
                    Assert.True(run.HasText);
                },
                run =>
                {
                    Assert.Equal("table:0:row:0:cell:0:paragraph:0", run.Container);
                    Assert.Equal(0, run.RunIndex);
                    Assert.Equal(string.Empty, run.Text);
                    Assert.False(run.HasText);
                },
                run =>
                {
                    Assert.Equal("table:0:row:0:cell:0:paragraph:0", run.Container);
                    Assert.Equal(1, run.RunIndex);
                    Assert.Equal("表格 Table", run.Text);
                    Assert.True(run.HasText);
                });
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void AssertRun(Run run, string size, string complexSize)
    {
        var properties = Assert.IsType<RunProperties>(run.RunProperties);
        var fonts = Assert.IsType<RunFonts>(properties.RunFonts);
        Assert.Equal("Times New Roman", fonts.Ascii!.Value);
        Assert.Equal("Times New Roman", fonts.HighAnsi!.Value);
        Assert.Equal("宋体", fonts.EastAsia!.Value);
        Assert.Equal("Times New Roman", fonts.ComplexScript!.Value);
        Assert.Equal(size, properties.FontSize!.Val!.Value);
        Assert.Equal(complexSize, properties.FontSizeComplexScript!.Val!.Value);
    }
}
