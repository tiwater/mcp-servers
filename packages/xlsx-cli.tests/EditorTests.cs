using Xunit;
using Dockit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx.Tests;

public class EditorTests
{
    [Fact]
    public void Inspect_reports_no_placeholders_for_fixed_layout_fixture()
    {
        var path = CreateWorkbookFixture();
        var report = Inspector.Inspect(path);

        Assert.Single(report.Sheets);
        Assert.Empty(report.Sheets[0].Placeholders);
        Assert.Empty(report.Sheets[0].TablePlaceholders);
    }

    [Fact]
    public void Edit_sets_single_cell_and_range_values()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "D2", Value: "260359-01"),
            new XlsxEditOperation("setRangeValues", Sheet: "Sheet1", StartCell: "E2", Values: [["233988", "383789"], ["252353", "341366"]])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var workbookPart = spreadsheet.WorkbookPart!;
        var sharedStrings = workbookPart.SharedStringTablePart!.SharedStringTable;
        var worksheet = workbookPart.WorksheetParts.Single().Worksheet;
        Assert.Equal("260359-01", GetCellText(worksheet, sharedStrings, "D2"));
        Assert.Equal("233988", GetCellText(worksheet, sharedStrings, "E2"));
        Assert.Equal("383789", GetCellText(worksheet, sharedStrings, "F2"));
        Assert.Equal("252353", GetCellText(worksheet, sharedStrings, "E3"));
        Assert.Equal("341366", GetCellText(worksheet, sharedStrings, "F3"));
    }

    [Fact]
    public void ExportJson_preserves_inline_string_headers_and_labels()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([path, output]);

        var json = File.ReadAllText(output);
        var parsed = System.Text.Json.JsonDocument.Parse(json);
        var rows = parsed.RootElement[0].GetProperty("rows");
        Assert.Equal("280 nm峰面积", rows[0][3].GetString());
        Assert.Equal("sample", rows[1][3].GetString());
        Assert.Equal("std", rows[2][3].GetString());
    }

    private static string CreateWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("D1", "280 nm峰面积"), ("E1", "LC"), ("F1", "LC_1d")),
            CreateRow(2, ("D2", "sample"), ("E2", "old"), ("F2", "old")),
            CreateRow(3, ("D3", "std"), ("E3", "old"), ("F3", "old"))
        ));
        var sheets = spreadsheet.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static Row CreateRow(uint rowIndex, params (string Ref, string Value)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach (var (cellRef, value) in cells)
        {
            row.Append(new Cell { CellReference = cellRef, DataType = CellValues.InlineString, InlineString = new InlineString(new Text(value)) });
        }
        return row;
    }

    private static string GetCellText(Worksheet worksheet, SharedStringTable sharedStrings, string cellRef)
    {
        var cell = worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == cellRef);
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return sharedStrings.ElementAt(int.Parse(cell.CellValue!.Text)).InnerText;
        }
        return cell.InnerText;
    }
}
