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
                    new Table(new TableRow(new TableCell(new Paragraph(new Run(
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
                Assert.Equal(2, runs.Count);
                AssertRun(runs[0], "24", "26");
                Assert.NotNull(runs[0].RunProperties!.Bold);
                AssertRun(runs[1], "19", "21");
                Assert.NotNull(runs[1].RunProperties!.Italic);
            }

            Assert.True(FontPolicy.Validate(output, policy, "policy-hash").Pass);
            Assert.False(FontPolicy.Validate(input, policy, "policy-hash").Pass);
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
