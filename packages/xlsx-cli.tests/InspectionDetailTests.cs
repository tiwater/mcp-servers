using Dockit.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
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
        Assert.Contains(sheet.TextCells!, cell => cell.Reference == "C5" && cell.Text == "shared label");
        Assert.Contains(sheet.TextCells!, cell => cell.Reference == "D5" && cell.Text == "TRUE");
        Assert.Contains(sheet.TextCells!, cell => cell.Reference == "E5" && cell.Text == "123.45");
        var inlineRichCell = Assert.Single(sheet.TextCells!, cell => cell.Reference == "F5");
        Assert.Equal("QVQLVQSGAEVK", inlineRichCell.Text);
        Assert.Contains(inlineRichCell.RichTextRuns!, run => run.Text == "Q" && run.Color == "FFFF0000" && run.Underline == "single");
        var sharedRichCell = Assert.Single(sheet.TextCells!, cell => cell.Reference == "G5");
        Assert.Equal("QAPGQGLEWMGWIYPGSANTK", sharedRichCell.Text);
        Assert.Contains(sharedRichCell.RichTextRuns!, run => run.Text == "N" && run.Color == "FFFF0000" && run.Underline == "single");
        Assert.Contains(sheet.FormulaCells!, cell => cell.Reference == "B12" && cell.Formula == "B6-B9*0.784" && cell.CachedValue == "10");
        Assert.Contains(sheet.FormulaCells!, cell => cell.Reference == "B14" && cell.Formula == "B12*2" && cell.CachedValue is null);
        Assert.DoesNotContain(sheet.TextCells!, cell => cell.Reference == "B14");
        Assert.Contains("A15:L15", sheet.MergedRanges!);
        Assert.Contains(sheet.RowHeights!, row => row.Row == 15 && row.Height == 42);
        Assert.DoesNotContain(sheet.RowHeights!, row => row.Row == 16);
        Assert.Contains(sheet.ColumnWidths!, column => column.Column == 1 && column.Width > 20);
    }

    [Fact]
    public void ExportJson_exposes_rich_text_runs_for_openxml_cells()
    {
        var path = CreateAna14LikeWorkbook();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-export-rich-text-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([path, output]);

        using var document = JsonDocument.Parse(File.ReadAllText(output));
        var cells = document.RootElement[0].GetProperty("cells").EnumerateArray().ToList();
        var inlineRichCell = cells.Single(cell => cell.GetProperty("reference").GetString() == "F5");
        var inlineRuns = inlineRichCell.GetProperty("richTextRuns").EnumerateArray().ToList();
        Assert.Contains(inlineRuns, run =>
            run.GetProperty("text").GetString() == "Q" &&
            run.GetProperty("color").GetString() == "FFFF0000" &&
            run.GetProperty("underline").GetString() == "single");

        var sharedRichCell = cells.Single(cell => cell.GetProperty("reference").GetString() == "G5");
        var sharedRuns = sharedRichCell.GetProperty("richTextRuns").EnumerateArray().ToList();
        Assert.Contains(sharedRuns, run =>
            run.GetProperty("text").GetString() == "N" &&
            run.GetProperty("color").GetString() == "FFFF0000" &&
            run.GetProperty("underline").GetString() == "single");
    }

    private static string CreateAna14LikeWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-inspection-detail-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringPart.SharedStringTable = new SharedStringTable(
            new SharedStringItem(new Text("shared label")),
            CreateRichSharedString());
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        var sheetData = new SheetData(
            CreateMixedValueRow(),
            CreateInlineStringRow(8, ("A8", "360 nm峰面积")),
            CreateInlineStringRow(11, ("A11", "杂质峰面积")),
            CreateFormulaRow(12, "B12", "B6-B9*0.784", "10"),
            CreateFormulaRow(13, "B13", "B7-B10*0.784", "11"),
            CreateFormulaWithoutCachedValueRow(),
            CreateInlineStringRow(15, ("A15", "merged title")),
            CreateInlineStringRow(16, ("A16", "default height flag"))
        );
        sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 15).Height = 42;
        sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 15).CustomHeight = true;
        sheetData.Elements<Row>().Single(row => row.RowIndex?.Value == 16).Height = 36;

        worksheetPart.Worksheet = new Worksheet(
            new Columns(new Column { Min = 1, Max = 1, Width = 24, CustomWidth = true }),
            sheetData,
            new MergeCells(new MergeCell { Reference = "A15:L15" })
        );

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "RP" });
        workbookPart.Workbook.Save();
        sharedStringPart.SharedStringTable.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static Row CreateMixedValueRow()
    {
        return new Row(
            new Cell { CellReference = "A5", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("280 nm峰面积")) },
            new Cell { CellReference = "C5", DataType = CellValues.SharedString, CellValue = new CellValue("0") },
            new Cell { CellReference = "D5", DataType = CellValues.Boolean, CellValue = new CellValue("1") },
            new Cell { CellReference = "E5", CellValue = new CellValue("123.45") },
            new Cell { CellReference = "F5", DataType = CellValues.InlineString, InlineString = CreateRichInlineString() },
            new Cell { CellReference = "G5", DataType = CellValues.SharedString, CellValue = new CellValue("1") })
        { RowIndex = 5 };
    }

    private static InlineString CreateRichInlineString()
    {
        return new InlineString(
            new Run(new Text("QV")),
            CreateRedUnderlinedRun("Q"),
            new Run(new Text("LVQSGAEVK")));
    }

    private static SharedStringItem CreateRichSharedString()
    {
        return new SharedStringItem(
            new Run(new Text("QAPGQGLEWMGWIYPGSA")),
            CreateRedUnderlinedRun("N"),
            new Run(new Text("TK")));
    }

    private static Run CreateRedUnderlinedRun(string text)
    {
        return new Run(
            new RunProperties(
                new Color { Rgb = "FFFF0000" },
                new Underline { Val = UnderlineValues.Single }),
            new Text(text));
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

    private static Row CreateFormulaWithoutCachedValueRow()
    {
        return new Row(
            new Cell
            {
                CellReference = "B14",
                CellFormula = new CellFormula("B12*2")
            })
        { RowIndex = 14 };
    }
}
