using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Docx;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class TocStylePolicyTests
{
    [Fact]
    public void Applies_nonitalic_character_indentation_to_all_toc_levels_only()
    {
        var input = Path.Combine(Path.GetTempPath(), $"toc-style-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("body")))));
            var part = main.AddNewPart<StyleDefinitionsPart>();
            part.Styles = new Styles(
                Toc("TOC1", "toc 1"),
                Toc("TOC2", "toc 2"),
                Toc("TOC3", "toc 3"),
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal" });
            part.Styles.Save();
            main.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"toc-style-edited-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(input, output, [
            new DocxEditOperation("applyTocStylePolicy", Italic: false, IndentCharactersPerLevel: 2)
        ]);
        Assert.True(Assert.Single(result.AppliedOperations).Applied);

        using var edited = WordprocessingDocument.Open(output, false);
        var styles = edited.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Assert.Equal(0, Indent(styles, "TOC1"));
        Assert.Equal(200, Indent(styles, "TOC2"));
        Assert.Equal(400, Indent(styles, "TOC3"));
        Assert.All(styles.Elements<Style>().Where(style => style.StyleId!.Value!.StartsWith("TOC")), style =>
            Assert.False(style.StyleRunProperties!.GetFirstChild<Italic>()!.Val!.Value));
        Assert.Null(styles.Elements<Style>().Single(style => style.StyleId == "Normal").StyleParagraphProperties);
    }

    private static Style Toc(string id, string name) => new(
        new StyleName { Val = name },
        new StyleParagraphProperties(new Indentation { Left = "999" }),
        new StyleRunProperties(new Italic()))
    { Type = StyleValues.Paragraph, StyleId = id };

    private static int Indent(Styles styles, string id) => styles.Elements<Style>()
        .Single(style => style.StyleId == id).StyleParagraphProperties!
        .GetFirstChild<Indentation>()!.LeftChars!.Value;
}
