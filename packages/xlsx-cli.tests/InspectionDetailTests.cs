using Dockit.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class InspectionDetailTests
{
    [Fact]
    public void Inspect_exposes_visible_text_formulas_dimensions_and_merges()
    {
        var path = CreateAna14LikeWorkbook();

        var report = Inspector.Inspect(path);

        var sheet = Assert.Single(report.Sheets);
        Assert.Equal("RP", sheet.Name);
        Assert.Contains(sheet.TextCells!, cell => cell.Reference == "A5" && cell.Text == "280 nm峰面积");
        Assert.Contains(sheet.TextCells!, cell => cell.Reference == "A8" && cell.Text == "360 nm峰面积");
        Assert.Contains(sheet.FormulaCells!, cell => cell.Reference == "B12" && cell.Formula == "B6-B9*0.784");
        Assert.Contains("A15:L15", sheet.MergedRanges!);
        Assert.Contains(sheet.RowHeights!, row => row.Row == 15 && row.Height == 42);
        Assert.Contains(sheet.ColumnWidths!, column => column.Column == 1 && column.Width > 20);
    }

    private static string CreateAna14LikeWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-inspection-detail-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        var sheetData = new SheetData(
            CreateInlineStringRow(5, ("A5", "280 nm峰面积")),
            CreateInlineStringRow(8, ("A8", "360 nm峰面积")),
            CreateInlineStringRow(11, ("A11", "杂质峰面积")),
            CreateFormulaRow(12, "B12", "B6-B9*0.784", "10"),
            CreateFormulaRow(13, "B13", "B7-B10*0.784", "11"),
            CreateInlineStringRow(15, ("A15", "merged title"))
        );
        sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 15).Height = 42;
        sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 15).CustomHeight = true;

        worksheetPart.Worksheet = new Worksheet(
            new Columns(new Column { Min = 1, Max = 1, Width = 24, CustomWidth = true }),
            sheetData,
            new MergeCells(new MergeCell { Reference = "A15:L15" })
        );

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "RP" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static Row CreateInlineStringRow(uint rowIndex, params (string Reference, string Value)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach (var (reference, value) in cells)
        {
            row.Append(new Cell
            {
                CellReference = reference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value))
            });
        }

        return row;
    }

    private static Row CreateFormulaRow(uint rowIndex, string reference, string formula, string cachedValue)
    {
        return new Row(
            new Cell
            {
                CellReference = reference,
                CellFormula = new CellFormula(formula),
                CellValue = new CellValue(cachedValue)
            })
        { RowIndex = rowIndex };
    }
}
