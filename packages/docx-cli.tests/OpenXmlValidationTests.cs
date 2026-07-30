using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Dockit.Docx.Tests;

public class OpenXmlValidationTests
{
    [Fact]
    public void Trailing_ui_priority_from_compatible_producers_is_a_warning()
    {
        var input = CreateFixture(
            """
            <w:style w:type="paragraph" w:styleId="PrimaryAfterQFormat">
              <w:name w:val="Primary after qFormat"/>
              <w:qFormat/>
              <w:uiPriority w:val="1"/>
              <w:rPr>
                <w14:textFill/>
              </w:rPr>
            </w:style>
            """,
            """
            <w:style w:type="table" w:styleId="PriorityAfterVisibility">
              <w:name w:val="Priority after visibility"/>
              <w:semiHidden/>
              <w:unhideWhenUsed/>
              <w:uiPriority w:val="99"/>
              <w:tblPr/>
            </w:style>
            """,
            """
            <w:style w:type="paragraph" w:styleId="PriorityAfterSemiHidden">
              <w:name w:val="Priority after semiHidden"/>
              <w:hidden/>
              <w:semiHidden/>
              <w:uiPriority w:val="99"/>
              <w:rPr/>
            </w:style>
            """);

        var result = OpenXmlValidation.Validate(input);

        Assert.True(
            result.Pass,
            string.Join(
                Environment.NewLine,
                result.Errors.Select(error => $"{error.Id}: {error.Description} ({error.Path})")));
        Assert.Empty(result.Errors);
        Assert.Equal(3, result.WarningCount);
        Assert.All(result.Warnings, warning =>
            Assert.Equal("wordprocessing-style-trailing-ui-priority", warning.CompatibilityCode));
    }

    [Fact]
    public void Other_ui_priority_ordering_errors_remain_hard_failures()
    {
        var input = CreateFixture(
            """
            <w:style w:type="paragraph" w:styleId="Compatible">
              <w:name w:val="Compatible"/>
              <w:qFormat/>
              <w:uiPriority w:val="1"/>
              <w:rPr/>
            </w:style>
            """,
            """
            <w:style w:type="paragraph" w:styleId="Invalid">
              <w:name w:val="Invalid"/>
              <w:rPr/>
              <w:uiPriority w:val="2"/>
            </w:style>
            """);

        var result = OpenXmlValidation.Validate(input);

        Assert.False(result.Pass);
        Assert.Equal(1, result.WarningCount);
        Assert.Single(result.Errors);
        Assert.Equal("Sch_UnexpectedElementContentExpectingComplex", result.Errors[0].Id);
        Assert.Null(result.Errors[0].CompatibilityCode);
    }

    [Fact]
    public void Wps_table_layout_in_a_style_is_a_compatibility_warning()
    {
        var input = CreateFixture(
            """
            <w:style w:type="table" w:styleId="WpsTableStyle">
              <w:name w:val="WPS table style"/>
              <w:qFormat/>
              <w:uiPriority w:val="1"/>
              <w:tblPr>
                <w:tblBorders/>
                <w:tblLayout w:type="fixed"/>
              </w:tblPr>
            </w:style>
            """);

        var result = OpenXmlValidation.Validate(input);

        Assert.True(
            result.Pass,
            string.Join(
                Environment.NewLine,
                result.Errors.Select(error => $"{error.Id}: {error.Description} ({error.Path})")));
        Assert.Empty(result.Errors);
        Assert.Equal(2, result.WarningCount);
        Assert.Contains(result.Warnings, warning =>
            warning.CompatibilityCode == "wordprocessing-style-trailing-ui-priority");
        Assert.Contains(result.Warnings, warning =>
            warning.CompatibilityCode == "wordprocessing-style-table-layout");
    }

    private static string CreateFixture(params string[] styles)
    {
        var path = Path.Combine(Path.GetTempPath(), $"openxml-validation-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("fixture")))));

        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var stylesXml =
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <w:styles
                 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                 mc:Ignorable="w14">
             {string.Join(Environment.NewLine, styles)}
             </w:styles>
             """;
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(stylesXml));
        stylesPart.FeedData(stream);
        return path;
    }
}
