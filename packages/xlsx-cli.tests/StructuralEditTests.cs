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
        var summary = GetWorksheet(workbookPart, "Summary");

        Assert.Equal("280 nm peak area", TestWorkbookReader.GetCellText(workbookPart, worksheet, "A5"));
        Assert.Equal("360 nm peak area", TestWorkbookReader.GetCellText(workbookPart, worksheet, "A10"));
        Assert.Equal("impurity peak area", TestWorkbookReader.GetCellText(workbookPart, worksheet, "A13"));
        Assert.Equal("B6-B11*0.784", TestWorkbookReader.GetCell(worksheet, "B14").CellFormula!.Text);
        Assert.Equal("B7-B12*0.784", TestWorkbookReader.GetCell(worksheet, "B15").CellFormula!.Text);
        Assert.Equal("B11+$B$11+\"B9\"+'Q1'!B9+Q1!A9", TestWorkbookReader.GetCell(worksheet, "D5").CellFormula!.Text);
        Assert.Equal("RP!B11", TestWorkbookReader.GetCell(summary, "A1").CellFormula!.Text);

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

    [Fact]
    public void Edit_expandSectionRows_expands_existing_template_section_in_place()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-expand-section-rows-{Guid.NewGuid():N}.xlsx");
        const string anchorText = "280 nm峰面积-360 nm峰面积*0.784";

        using (var template = SpreadsheetDocument.Open(path, true))
        {
            var templateWorksheet = GetWorksheet(template.WorkbookPart!, "RP");
            var anchorCell = TestWorkbookReader.GetCell(templateWorksheet, "A11");
            anchorCell.InlineString = new InlineString(new Text(anchorText));
            templateWorksheet.Save();
        }

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation(
                "expandSectionRows",
                Sheet: "RP",
                AnchorText: anchorText,
                ExampleRows: 2,
                TargetRows: 4,
                PreserveStyle: true,
                PreserveFormulas: true,
                PreserveMergedRanges: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);
        Assert.Equal("RP", operation.Sheet);
        Assert.Equal("12:15", operation.ChangedRange);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var sheetData = worksheet.GetFirstChild<SheetData>()!;

        Assert.Equal(anchorText, TestWorkbookReader.GetCellText(spreadsheet.WorkbookPart!, worksheet, "A11"));
        Assert.Contains(sheetData.Elements<Row>(), row => row.RowIndex?.Value == 14);
        Assert.Contains(sheetData.Elements<Row>(), row => row.RowIndex?.Value == 15);
        Assert.Equal("B8-B11*0.784", TestWorkbookReader.GetCell(worksheet, "B14").CellFormula!.Text);

        var mergeCells = worksheet.Elements<MergeCells>().Single();
        Assert.Contains(mergeCells.Elements<MergeCell>(), merge => merge.Reference?.Value == "A17:L17");
        Assert.DoesNotContain(mergeCells.Elements<MergeCell>(), merge => merge.Reference?.Value == "A15:L15");
    }

    [Fact]
    public void Edit_expandSectionRows_copies_original_example_formulas_and_preserves_merged_ranges_by_default()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-expand-section-formulas-merges-{Guid.NewGuid():N}.xlsx");
        const string anchorText = "280 nm峰面积-360 nm峰面积*0.784";

        using (var template = SpreadsheetDocument.Open(path, true))
        {
            var templateWorksheet = GetWorksheet(template.WorkbookPart!, "RP");
            var anchorCell = TestWorkbookReader.GetCell(templateWorksheet, "A11");
            anchorCell.InlineString = new InlineString(new Text(anchorText));

            var exemplarFormula = TestWorkbookReader.GetCell(templateWorksheet, "B12");
            exemplarFormula.CellFormula = new CellFormula("SUM(B15)");
            exemplarFormula.CellValue = new CellValue("10");

            var templateMergeCells = templateWorksheet.Elements<MergeCells>().Single();
            templateMergeCells.AppendChild(new MergeCell { Reference = "A12:C12" });
            templateWorksheet.Save();
        }

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation(
                "expandSectionRows",
                Sheet: "RP",
                AnchorText: anchorText,
                ExampleRows: 2,
                TargetRows: 4,
                PreserveStyle: true,
                PreserveFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");

        Assert.Equal("SUM(B17)", TestWorkbookReader.GetCell(worksheet, "B12").CellFormula!.Text);
        Assert.Equal("SUM(B17)", TestWorkbookReader.GetCell(worksheet, "B14").CellFormula!.Text);

        var mergeCells = worksheet.Elements<MergeCells>().Single();
        Assert.Contains(mergeCells.Elements<MergeCell>(), merge => merge.Reference?.Value == "A12:C12");
        Assert.Contains(mergeCells.Elements<MergeCell>(), merge => merge.Reference?.Value == "A14:C14");
    }

    [Fact]
    public void Edit_expandSectionRows_materializes_shared_formulas_before_copying_rows()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-expand-section-shared-formulas-{Guid.NewGuid():N}.xlsx");
        const string anchorText = "280 nm峰面积-360 nm峰面积*0.784";

        using (var template = SpreadsheetDocument.Open(path, true))
        {
            var templateWorksheet = GetWorksheet(template.WorkbookPart!, "RP");
            var anchorCell = TestWorkbookReader.GetCell(templateWorksheet, "A11");
            anchorCell.InlineString = new InlineString(new Text(anchorText));

            var master = TestWorkbookReader.GetCell(templateWorksheet, "F12");
            master.CellFormula = new CellFormula("F6-F11*0.784")
            {
                FormulaType = CellFormulaValues.Shared,
                Reference = "F12:G12",
                SharedIndex = 0
            };
            master.CellValue = new CellValue("10");

            var follower = TestWorkbookReader.GetCell(templateWorksheet, "G12");
            follower.CellFormula = new CellFormula
            {
                FormulaType = CellFormulaValues.Shared,
                SharedIndex = 0
            };
            follower.CellValue = new CellValue("20");

            templateWorksheet.Save();
        }

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation(
                "expandSectionRows",
                Sheet: "RP",
                AnchorText: anchorText,
                ExampleRows: 2,
                TargetRows: 4,
                PreserveStyle: true,
                PreserveFormulas: true,
                PreserveMergedRanges: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        var validation = Validator.Validate(output);
        Assert.True(validation.Valid, string.Join(Environment.NewLine, validation.Errors));

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        Assert.Equal("F6-F11*0.784", TestWorkbookReader.GetCell(worksheet, "F12").CellFormula!.Text);
        Assert.Equal("G6-G11*0.784", TestWorkbookReader.GetCell(worksheet, "G12").CellFormula!.Text);
        Assert.Equal("F8-F13*0.784", TestWorkbookReader.GetCell(worksheet, "F14").CellFormula!.Text);
        Assert.Null(TestWorkbookReader.GetCell(worksheet, "F14").CellFormula!.FormulaType);
    }

    [Fact]
    public void Edit_expandSectionRows_shrink_noop_reports_no_changed_range()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-expand-section-shrink-{Guid.NewGuid():N}.xlsx");
        const string anchorText = "280 nm峰面积-360 nm峰面积*0.784";

        using (var template = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = GetWorksheet(template.WorkbookPart!, "RP");
            var anchorCell = TestWorkbookReader.GetCell(worksheet, "A11");
            anchorCell.InlineString = new InlineString(new Text(anchorText));
            worksheet.Save();
        }

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation(
                "expandSectionRows",
                Sheet: "RP",
                AnchorText: anchorText,
                ExampleRows: 2,
                TargetRows: 1)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);
        Assert.Null(operation.ChangedRange);
        Assert.Contains(operation.Warnings ?? [], warning => warning.Contains("Shrink unsupported", StringComparison.Ordinal));
    }

    [Fact]
    public void Edit_copyRow_translates_formula_references_without_rewriting_functions_strings_or_absolute_rows()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-formula-guards-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "C14");

        Assert.Equal("LOG10(A3)+\"A12\"+$B$12", targetCell.CellFormula!.Text);
    }

    [Fact]
    public void Edit_copyRow_clears_cached_value_when_formula_is_translated()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-clear-cache-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "C14");

        Assert.Null(targetCell.CellValue);
    }

    [Fact]
    public void Edit_copyRow_preserves_cached_value_when_formula_is_unchanged_by_translation()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-preserve-cache-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "D14");

        Assert.Equal("$B$12", targetCell.CellFormula!.Text);
        Assert.Equal("115.2", targetCell.CellValue!.Text);
    }

    [Fact]
    public void Edit_copyRow_rejects_translation_that_would_create_invalid_row_references()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-invalid-upward-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 1, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.False(operation.Applied);
        Assert.Contains("row < 1", operation.Detail, StringComparison.Ordinal);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        Assert.DoesNotContain(worksheet.Descendants<Cell>(), cell => cell.CellReference?.Value == "C1");
    }

    [Fact]
    public void Edit_copyRow_does_not_rewrite_single_quoted_sheet_names()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-sheet-name-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "E14");

        Assert.Equal("='Q1'!A3", targetCell.CellFormula!.Text);
    }

    [Fact]
    public void Edit_copyRow_preserves_unquoted_sheet_names_while_translating_references()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-unquoted-sheet-name-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "G14");

        Assert.Equal("Q1!A3", targetCell.CellFormula!.Text);
    }

    [Fact]
    public void Edit_copyRow_translates_lowercase_references_and_clears_cached_value()
    {
        var path = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-copy-row-lowercase-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("copyRow", Sheet: "RP", SourceRow: 12, TargetRow: 14, TranslateFormulas: true)
        ]);

        var operation = Assert.Single(result.AppliedOperations);
        Assert.True(operation.Applied, operation.Detail);

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = GetWorksheet(spreadsheet.WorkbookPart!, "RP");
        var targetCell = TestWorkbookReader.GetCell(worksheet, "F14");

        Assert.Equal("a3+b3", targetCell.CellFormula!.Text);
        Assert.Null(targetCell.CellValue);
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
            CreateMixedRow(5, ("A5", "280 nm peak area"), ("B5", "area"), ("C5", "ratio")),
            CreateInlineStringRow(6, ("A6", "sample 280"), ("B6", "125.000")),
            CreateInlineStringRow(7, ("A7", "sample 280 repeat"), ("B7", "130.000")),
            CreateInlineStringRow(8, ("A8", "360 nm peak area"), ("B8", "area")),
            CreateInlineStringRow(9, ("A9", "blank 360"), ("B9", "12.500")),
            CreateInlineStringRow(10, ("A10", "blank 360 repeat"), ("B10", "13.000")),
            CreateInlineStringRow(11, ("A11", "impurity peak area"), ("B11", "area")),
            CreateFormulaRow(12, ("B12", "B6-B9*0.784", "115.2"), ("C12", "LOG10(A1)+\"A12\"+$B$12", "7.5"), ("D12", "$B$12", "115.2"), ("E12", "='Q1'!A1", "11"), ("F12", "a1+b1", "22"), ("G12", "Q1!A1", "33")),
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

        var summaryPart = workbookPart.AddNewPart<WorksheetPart>();
        summaryPart.Worksheet = new Worksheet(new SheetData(
            CreateFormulaRow(1, "A1", "RP!B9", "12.5")
        ));

        var q1Part = workbookPart.AddNewPart<WorksheetPart>();
        q1Part.Worksheet = new Worksheet(new SheetData(
            CreateInlineStringRow(1, ("A1", "Q1 value"))
        ));

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "RP" });
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(summaryPart), SheetId = 2, Name = "Summary" });
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(q1Part), SheetId = 3, Name = "Q1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        summaryPart.Worksheet.Save();
        q1Part.Worksheet.Save();
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

    private static Row CreateMixedRow(uint rowIndex, params (string Reference, string Value)[] cells)
    {
        var row = CreateInlineStringRow(rowIndex, cells);
        row.Append(new Cell
        {
            CellReference = "D5",
            StyleIndex = 1,
            CellFormula = new CellFormula("B9+$B$9+\"B9\"+'Q1'!B9+Q1!A9"),
            CellValue = new CellValue("25")
        });

        return row;
    }

    private static Row CreateFormulaRow(uint rowIndex, string reference, string formula, string cachedValue)
    {
        return CreateFormulaRow(rowIndex, (reference, formula, cachedValue));
    }

    private static Row CreateFormulaRow(uint rowIndex, params (string Reference, string Formula, string CachedValue)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach (var (reference, formula, cachedValue) in cells)
        {
            row.Append(new Cell
            {
                CellReference = reference,
                StyleIndex = 1,
                CellFormula = new CellFormula(formula),
                CellValue = new CellValue(cachedValue)
            });
        }

        return row;
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
