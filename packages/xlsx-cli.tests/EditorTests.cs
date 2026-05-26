using Xunit;
using Dockit.Xlsx;
using DocumentFormat.OpenXml;
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
    public void Edit_stores_numeric_text_as_number_while_preserving_target_style()
    {
        var path = CreateFormattedWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-number-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A2", Value: "10.2"),
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "C2", Value: "10.2")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var styledCell = GetCell(worksheet, "A2");
        var generalCell = GetCell(worksheet, "C2");
        Assert.Null(styledCell.DataType);
        Assert.Equal("10.2", styledCell.CellValue!.Text);
        Assert.Equal<UInt32Value>(1, styledCell.StyleIndex!);
        Assert.Null(generalCell.DataType);
        Assert.Equal("10.2", generalCell.CellValue!.Text);
    }

    [Fact]
    public void Edit_keeps_numeric_text_as_text_when_target_cell_is_text_formatted()
    {
        var path = CreateTextFormattedWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-text-format-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A2", Value: "10.2")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var workbookPart = spreadsheet.WorkbookPart!;
        var worksheet = workbookPart.WorksheetParts.Single().Worksheet;
        var cell = GetCell(worksheet, "A2");
        Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
        Assert.Equal("10.2", workbookPart.SharedStringTablePart!.SharedStringTable.ElementAt(int.Parse(cell.CellValue!.Text)).InnerText);
        Assert.Equal<UInt32Value>(1, cell.StyleIndex!);
    }

    [Fact]
    public void Edit_converts_percent_text_when_target_cell_uses_percent_format()
    {
        var path = CreatePercentFormattedWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-percent-format-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A2", Value: "99.1%")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var cell = GetCell(worksheet, "A2");
        Assert.Null(cell.DataType);
        Assert.Equal("0.991", cell.CellValue!.Text);
        Assert.Equal<UInt32Value>(1, cell.StyleIndex!);
    }

    [Fact]
    public void ExportJson_preserves_inline_string_headers_and_labels()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([path, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("280 nm峰面积", json, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\u5CF0", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u9762", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u79EF", json, StringComparison.OrdinalIgnoreCase);
        var parsed = System.Text.Json.JsonDocument.Parse(json);
        var rows = parsed.RootElement[0].GetProperty("rows");
        Assert.Equal("280 nm峰面积", rows[0][3].GetString());
        Assert.Equal("sample", rows[1][3].GetString());
        Assert.Equal("std", rows[2][3].GetString());
    }

    [Fact]
    public void ExportJson_uses_display_format_for_numeric_cells()
    {
        var path = CreateFormattedWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-formatted-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([path, output]);

        var parsed = System.Text.Json.JsonDocument.Parse(File.ReadAllText(output));
        var rows = parsed.RootElement[0].GetProperty("rows");
        var formattedRows = parsed.RootElement[0].GetProperty("formattedRows");
        Assert.Equal("0.393", rows[1][0].GetString());
        Assert.Equal("32.299999999999997", rows[1][1].GetString());
        Assert.Equal("0.4", formattedRows[1][0].GetString());
        Assert.Equal("32.3", formattedRows[1][1].GetString());
    }

    [Fact]
    public void ExportJson_includes_addressed_cells_with_formulas()
    {
        var path = CreateFormulaWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-addressed-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([path, output]);

        var parsed = System.Text.Json.JsonDocument.Parse(File.ReadAllText(output));
        var cells = parsed.RootElement[0].GetProperty("cells");
        Assert.Equal("A1", cells[0].GetProperty("reference").GetString());
        Assert.Equal(1, cells[0].GetProperty("row").GetInt32());
        Assert.Equal(1, cells[0].GetProperty("column").GetInt32());
        Assert.Equal("Sample", cells[0].GetProperty("value").GetString());
        var formulaCell = cells.EnumerateArray().Single(cell => cell.GetProperty("reference").GetString() == "C2");
        Assert.Equal("A2+B2", formulaCell.GetProperty("formula").GetString());
        Assert.Equal("3", formulaCell.GetProperty("value").GetString());
    }

    [Fact]
    public void ExportJson_expands_shared_formulas_in_addressed_cells()
    {
        var path = CreateSharedFormulaWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-shared-formula-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([path, output]);

        var parsed = System.Text.Json.JsonDocument.Parse(File.ReadAllText(output));
        var cells = parsed.RootElement[0].GetProperty("cells");
        var firstFormula = cells.EnumerateArray().Single(cell => cell.GetProperty("reference").GetString() == "C2");
        var sharedFormula = cells.EnumerateArray().Single(cell => cell.GetProperty("reference").GetString() == "C3");
        Assert.Equal("A2+B2", firstFormula.GetProperty("formula").GetString());
        Assert.Equal("A3+B3", sharedFormula.GetProperty("formula").GetString());
        Assert.Equal("7", sharedFormula.GetProperty("value").GetString());
    }

    [Fact]
    public void Edit_inserts_rows_from_template_row_and_translates_formulas()
    {
        var path = CreateRowInsertionWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-insert-row-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("insertRows", Sheet: "Sheet1", SourceRow: 2, TargetRow: 3, Count: 2),
            new XlsxEditOperation("setRangeValues", Sheet: "Sheet1", StartCell: "A3", Values: [["3", "4"], ["5", "6"]])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.Equal("Footer", GetCell(worksheet, "A5").InnerText);
        Assert.Equal<UInt32Value>(5, worksheet.Descendants<Row>().Single(row => row.RowIndex?.Value == 5).RowIndex!);
        Assert.Equal("A3+B3", GetCell(worksheet, "C3").CellFormula!.Text);
        Assert.Equal("A4+B4", GetCell(worksheet, "C4").CellFormula!.Text);
        Assert.Equal<UInt32Value>(1, GetCell(worksheet, "A3").StyleIndex!);
        Assert.Equal<UInt32Value>(1, GetCell(worksheet, "A4").StyleIndex!);
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

    private static string CreateFormattedWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-formatted-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new NumberingFormats(
                new NumberingFormat { NumberFormatId = 164, FormatCode = "0.0_);[Red]\\(0.0\\)" }
            ) { Count = 1 },
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellFormats(
                new CellFormat { NumberFormatId = 0, ApplyNumberFormat = false },
                new CellFormat { NumberFormatId = 164, ApplyNumberFormat = true }
            ) { Count = 2 }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var dataRow = new Row { RowIndex = 2 };
        dataRow.Append(
            new Cell { CellReference = "A2", StyleIndex = 1, CellValue = new CellValue("0.393") },
            new Cell { CellReference = "B2", CellValue = new CellValue("32.299999999999997") }
        );
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("A1", "Rounded"), ("B1", "General")),
            dataRow
        ));
        var sheets = spreadsheet.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreateTextFormattedWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-text-formatted-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellFormats(
                new CellFormat { NumberFormatId = 0, ApplyNumberFormat = false },
                new CellFormat { NumberFormatId = 49, ApplyNumberFormat = true }
            ) { Count = 2 }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var dataRow = new Row { RowIndex = 2 };
        dataRow.Append(new Cell { CellReference = "A2", StyleIndex = 1, DataType = CellValues.InlineString, InlineString = new InlineString(new Text("old")) });
        worksheetPart.Worksheet = new Worksheet(new SheetData(CreateRow(1, ("A1", "Text")), dataRow));
        var sheets = spreadsheet.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreatePercentFormattedWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-percent-formatted-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellFormats(
                new CellFormat { NumberFormatId = 0, ApplyNumberFormat = false },
                new CellFormat { NumberFormatId = 10, ApplyNumberFormat = true }
            ) { Count = 2 }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var dataRow = new Row { RowIndex = 2 };
        dataRow.Append(new Cell { CellReference = "A2", StyleIndex = 1, CellValue = new CellValue("0") });
        worksheetPart.Worksheet = new Worksheet(new SheetData(CreateRow(1, ("A1", "Percent")), dataRow));
        var sheets = spreadsheet.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreateFormulaWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-formula-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var dataRow = new Row { RowIndex = 2 };
        dataRow.Append(
            new Cell { CellReference = "A2", CellValue = new CellValue("1") },
            new Cell { CellReference = "B2", CellValue = new CellValue("2") },
            new Cell { CellReference = "C2", CellFormula = new CellFormula("A2+B2"), CellValue = new CellValue("3") }
        );
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("A1", "Sample"), ("B1", "Input"), ("C1", "Formula")),
            dataRow
        ));
        var sheets = spreadsheet.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreateSharedFormulaWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-shared-formula-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var firstDataRow = new Row { RowIndex = 2 };
        firstDataRow.Append(
            new Cell { CellReference = "A2", CellValue = new CellValue("1") },
            new Cell { CellReference = "B2", CellValue = new CellValue("2") },
            new Cell { CellReference = "C2", CellFormula = new CellFormula("A2+B2") { FormulaType = CellFormulaValues.Shared, Reference = "C2:C3", SharedIndex = 0 }, CellValue = new CellValue("3") }
        );
        var secondDataRow = new Row { RowIndex = 3 };
        secondDataRow.Append(
            new Cell { CellReference = "A3", CellValue = new CellValue("3") },
            new Cell { CellReference = "B3", CellValue = new CellValue("4") },
            new Cell { CellReference = "C3", CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 }, CellValue = new CellValue("7") }
        );
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("A1", "Left"), ("B1", "Right"), ("C1", "Total")),
            firstDataRow,
            secondDataRow
        ));
        var sheets = spreadsheet.WorkbookPart!.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreateRowInsertionWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-row-insertion-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellFormats(
                new CellFormat { NumberFormatId = 0, ApplyNumberFormat = false },
                new CellFormat { NumberFormatId = 2, ApplyNumberFormat = true }
            ) { Count = 2 }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var templateRow = new Row { RowIndex = 2, CustomHeight = true, Height = 20 };
        templateRow.Append(
            new Cell { CellReference = "A2", StyleIndex = 1, CellValue = new CellValue("1") },
            new Cell { CellReference = "B2", StyleIndex = 1, CellValue = new CellValue("2") },
            new Cell { CellReference = "C2", StyleIndex = 1, CellFormula = new CellFormula("A2+B2"), CellValue = new CellValue("3") }
        );
        var footerRow = new Row { RowIndex = 3 };
        footerRow.Append(new Cell { CellReference = "A3", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("Footer")) });
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("A1", "Left"), ("B1", "Right"), ("C1", "Total")),
            templateRow,
            footerRow
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
        var cell = GetCell(worksheet, cellRef);
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            return sharedStrings.ElementAt(int.Parse(cell.CellValue!.Text)).InnerText;
        }
        return cell.InnerText;
    }

    private static Cell GetCell(Worksheet worksheet, string cellRef)
        => worksheet.Descendants<Cell>().Single(c => c.CellReference?.Value == cellRef);
}
