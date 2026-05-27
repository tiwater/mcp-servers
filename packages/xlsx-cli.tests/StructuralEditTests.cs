using Dockit.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class StructuralEditTests
{
    [Fact]
    public void Edit_insertRows_moves_following_rows_and_merged_ranges()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-insert-rows-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("insertRows", Sheet: "RP", StartRow: 8, Count: 2)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);
        Assert.Equal("RP", operation.Sheet);
        Assert.Equal("8:9", operation.ChangedRange);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var workbookPart = spreadsheet.WorkbookPart!;
        var worksheet = GetWorksheet(workbookPart, "RP");

        Assert.Equal("280 nm peak area", TestWorkbookReader.GetCellText(workbookPart, worksheet, "A5"));
        Assert.Equal("360 nm peak area", TestWorkbookReader.GetCellText(workbookPart, worksheet, "A10"));
        Assert.Equal("impurity peak area", TestWorkbookReader.GetCellText(workbookPart, worksheet, "A13"));
        Assert.Equal("B6-B9*0.784", TestWorkbookReader.GetCell(worksheet, "B14").CellFormula!.Text);
        Assert.Equal("B7-B10*0.784", TestWorkbookReader.GetCell(worksheet, "B15").CellFormula!.Text);

        var mergeCells = worksheet.Elements<MergeCells>().Single();
        Assert.Contains(mergeCells.Elements<MergeCell>(), merge => merge.Reference?.Value == "A17:L17");
        Assert.DoesNotContain(mergeCells.Elements<MergeCell>(), merge => merge.Reference?.Value == "A15:L15");
    }

    [Fact]
    public void Edit_copyRow_copies_styles_formulas_and_row_height()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);
        Assert.Equal("RP", operation.Sheet);
        Assert.Equal("14:14", operation.ChangedRange);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var sourceCell = TestWorkbookReader.GetCell(worksheet, "B12");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "B14");
        var sourceRow = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single(row => row.RowIndex?.Value == 12);
        var targetRow = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single(row => row.RowIndex?.Value == 14);

        Assert.Equal(sourceCell.StyleIndex, targetCell.StyleIndex);
        Assert.True(targetRow.CustomHeight?.Value);
        Assert.Equal(sourceRow.Height, targetRow.Height);
        Assert.Equal("B8-B11*0.784", targetCell.CellFormula!.Text);
    }

    private static Worksheet GetWorksheet(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook.Descendants<Sheet>().Single(s => s.Name?.Value == sheetName);
        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
        return worksheetPart.Worksheet;
    }
}

internal static class WorkbookFixtures
{
    public static string CreateAna14LikeWorkbookWithStyles()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-ana14-like-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        AddStyles(workbookPart);

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(
            CreateInlineStringRow(5, ("A5", "280 nm peak area"), ("B5", "area"), ("C5", "ratio")),
            CreateInlineStringRow(6, ("A6", "sample 280"), ("B6", "125.000")),
            CreateInlineStringRow(7, ("A7", "sample 280 repeat"), ("B7", "130.000")),
            CreateInlineStringRow(8, ("A8", "360 nm peak area"), ("B8", "area")),
            CreateInlineStringRow(9, ("A9", "blank 360"), ("B9", "12.500")),
            CreateInlineStringRow(10, ("A10", "blank 360 repeat"), ("B10", "13.000")),
            CreateInlineStringRow(11, ("A11", "impurity peak area"), ("B11", "area")),
            CreateFormulaRow(12, "B12", "B6-B9*0.784", "115.2"),
            CreateFormulaRow(13, "B13", "B7-B10*0.784", "119.808"),
            CreateInlineStringRow(15, ("A15", "calculation note spans report width"))
        );

        var row12 = sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 12);
        row12.Height = 31.5;
        row12.CustomHeight = true;

        var row15 = sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 15);
        row15.Height = 42;
        row15.CustomHeight = true;

        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A5:L15" },
            new Columns(new Column { Min = 1, Max = 12, Width = 14, CustomWidth = true }),
            sheetData,
            new MergeCells(new MergeCell { Reference = "A15:L15" })
        );

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "RP" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static void AddStyles(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(
                new Font(),
                new Font(new Bold())
            )
            { Count = 2 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(
                new CellFormat { FontId = 0, FillId = 0, BorderId = 0 },
                new CellFormat { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }
            )
            { Count = 2 }
        );
        stylesPart.Stylesheet.Save();
    }

    private static Row CreateInlineStringRow(uint rowIndex, params (string Reference, string Value)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach (var (reference, value) in cells)
        {
            row.Append(new Cell
            {
                CellReference = reference,
                StyleIndex = 1,
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
                StyleIndex = 1,
                CellFormula = new CellFormula(formula),
                CellValue = new CellValue(cachedValue)
            })
        { RowIndex = rowIndex };
    }
}

internal static class TestWorkbookReader
{
    public static string GetCellText(WorkbookPart workbookPart, Worksheet worksheet, string reference)
    {
        var cell = GetCell(worksheet, reference);
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sharedStrings = workbookPart.SharedStringTablePart!.SharedStringTable;
            return sharedStrings.ElementAt(int.Parse(cell.CellValue!.Text)).InnerText;
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString!.InnerText;
        }

        return cell.CellValue?.Text ?? cell.CellFormula?.Text ?? string.Empty;
    }

    public static Cell GetCell(Worksheet worksheet, string reference)
    {
        return worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == reference);
    }
}
