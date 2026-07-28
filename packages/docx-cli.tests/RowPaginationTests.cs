using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class RowPaginationTests
{
    [Fact]
    public void Keep_next_is_written_in_schema_order_and_reported_by_inspection()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-row-pagination-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Table(
                        new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                        new TableGrid(new GridColumn { Width = "5000" }),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                                new Paragraph(
                                    new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                                    new Run(new Text("row"))))))));
                main.Document.Save();
            }

            var result = Editor.Apply(
                input,
                output,
                [new DocxEditOperation("setTableRowKeepNext", TableIndex: 0, RowIndex: 0, KeepNext: true)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            var row = Assert.Single(Assert.Single(Inspector.InspectTables(output).Tables).Rows);
            Assert.True(row.KeepNext);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
