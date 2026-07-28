using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using Xunit;

namespace Dockit.Xlsx.Tests;

public sealed class FontBoldEvidenceTests
{
    [Fact]
    public void Inspect_export_and_edit_report_effective_font_bold()
    {
        var input = CreateBoldWorkbook();
        var before = Inspector.InspectEvidence(input).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("cells");
        Assert.True(Cell(before, "A1").GetProperty("style").GetProperty("bold").GetBoolean());
        Assert.True(Cell(before, "B1").GetProperty("style").GetProperty("bold").GetBoolean());

        var output = Path.Combine(Path.GetTempPath(), $"xlsx-bold-output-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(input, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A1", Value: "direct", Bold: false),
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "B1", Value: "inherited", Bold: false),
        ]);
        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));

        var after = Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("cells");
        Assert.False(Cell(after, "A1").GetProperty("style").GetProperty("bold").GetBoolean());
        Assert.False(Cell(after, "B1").GetProperty("style").GetProperty("bold").GetBoolean());

        var exportPath = Path.Combine(Path.GetTempPath(), $"xlsx-bold-export-{Guid.NewGuid():N}.json");
        Assert.Equal(0, Extractor.RunExportJson([output, exportPath]));
        using var export = JsonDocument.Parse(File.ReadAllText(exportPath));
        var cells = export.RootElement[0].GetProperty("cells");
        Assert.False(Cell(cells, "A1").GetProperty("style").GetProperty("bold").GetBoolean());
        Assert.False(Cell(cells, "B1").GetProperty("style").GetProperty("bold").GetBoolean());

        using var edited = SpreadsheetDocument.Open(output, false);
        var styles = edited.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
        Assert.Equal((uint)2, styles.Fonts!.Count!.Value);
        Assert.Equal((uint)4, styles.CellFormats!.Count!.Value);
    }

    private static JsonElement Cell(JsonElement cells, string reference)
        => cells.EnumerateArray().Single(cell => cell.GetProperty("reference").GetString() == reference);

    private static string CreateBoldWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-bold-input-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(
                new Font(),
                new Font(new Bold())
            ) { Count = 2 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(
                new CellFormat { FontId = 0 },
                new CellFormat { FontId = 1, ApplyFont = true }
            ) { Count = 2 },
            new CellFormats(
                new CellFormat { FontId = 0, ApplyFont = true, FormatId = 0 },
                new CellFormat { FontId = 1, ApplyFont = true, FormatId = 0 },
                new CellFormat { FontId = 0, ApplyFont = false, FormatId = 1 }
            ) { Count = 3 }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            new Row(
                new Cell { CellReference = "A1", StyleIndex = 1, DataType = CellValues.InlineString, InlineString = new InlineString(new Text("direct bold")) },
                new Cell { CellReference = "B1", StyleIndex = 2, DataType = CellValues.InlineString, InlineString = new InlineString(new Text("inherited bold")) }
            ) { RowIndex = 1 }
        ));
        workbookPart.Workbook.AppendChild(new Sheets()).Append(
            new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }
}
