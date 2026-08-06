using Xunit;
using Dockit.Xlsx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace Dockit.Xlsx.Tests;

public class EditorTests
{
    [Fact]
    public void Operation_positional_value_parameter_remains_compatible_with_pre_range_api()
    {
        var operation = new XlsxEditOperation("setCellValue", "Sheet1", "A1", "legacy-value");
        Assert.Equal("legacy-value", operation.Value);
        Assert.Null(operation.Range);
    }
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
    public void Edit_can_enable_text_fitting_while_preserving_existing_cell_styles_and_read_it_back()
    {
        var path = CreateFormattedWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-wrap-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A2", Value: "complete customer text", WrapText: true),
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "B2", Value: "complete compact text", ShrinkToFit: true)
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var workbookPart = spreadsheet.WorkbookPart!;
        var cell = GetCell(workbookPart.WorksheetParts.Single().Worksheet, "A2");
        var style = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!
            .Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.True(style.Alignment!.WrapText!.Value);
        Assert.Equal<UInt32Value>(164, style.NumberFormatId!);
        var evidence = Inspector.InspectEvidence(output);
        var cells = evidence.GetProperty("evidence").GetProperty("sheets")[0].GetProperty("cells");
        Assert.True(cells.EnumerateArray().Single(item => item.GetProperty("reference").GetString() == "A2").GetProperty("style").GetProperty("wrapText").GetBoolean());
        Assert.True(cells.EnumerateArray().Single(item => item.GetProperty("reference").GetString() == "B2").GetProperty("style").GetProperty("shrinkToFit").GetBoolean());
    }

    [Fact]
    public void Edit_sets_a_sheet_local_print_area()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-print-area-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPrintArea", Sheet: "Sheet1", Range: "A1:F3")
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var definedName = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>());
        Assert.Equal("_xlnm.Print_Area", definedName.Name!.Value);
        Assert.Equal<uint>(0, definedName.LocalSheetId!.Value);
        Assert.Equal("'Sheet1'!$A$1:$F$3", definedName.Text);
        Assert.Equal("'Sheet1'!$A$1:$F$3", Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("print").GetProperty("area").GetString());
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_sets_fit_to_page_dimensions()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-page-setup-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", FitToPagesWide: 1)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.True(worksheet.GetFirstChild<SheetProperties>()!.GetFirstChild<PageSetupProperties>()!.FitToPage!.Value);
        var setup = worksheet.GetFirstChild<PageSetup>()!;
        Assert.Equal<uint>(1, setup.FitToWidth!.Value);
        Assert.Equal<uint>(0, setup.FitToHeight!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_sets_fit_to_page_height_without_constraining_width()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-page-height-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", FitToPagesTall: 1)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var setup = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<PageSetup>()!;
        Assert.Equal<uint>(0, setup.FitToWidth!.Value);
        Assert.Equal<uint>(1, setup.FitToHeight!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_sets_page_orientation_without_changing_fit_dimensions()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-page-orientation-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", Orientation: "landscape", PaperSize: "a3")
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var setup = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<PageSetup>()!;
        Assert.Equal(OrientationValues.Landscape, setup.Orientation!.Value);
        Assert.Equal<uint>(8, setup.PaperSize!.Value);
        Assert.Null(setup.FitToWidth);
        Assert.Null(setup.FitToHeight);
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_sets_repeating_title_rows_and_preserves_other_defined_names()
    {
        var path = CreateWorkbookFixture();
        using (var source = SpreadsheetDocument.Open(path, true))
        {
            var workbook = source.WorkbookPart!.Workbook;
            workbook.Sheets!.Elements<Sheet>().Single().Name = "O'Brien, East";
            workbook.Append(new DefinedNames(
                new DefinedName("0.25") { Name = "GlobalRate" },
                new DefinedName("'O''Brien, East'!$A$1:$F$9") { Name = "_xlnm.Print_Area", LocalSheetId = 0 },
                new DefinedName("'O''Brien, East'!$A:$B,'O''Brien, East'!$1:$1") { Name = "_xlnm.Print_Titles", LocalSheetId = 0 }));
            workbook.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-repeat-rows-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "O'Brien, East", RepeatRowsStart: 2, RepeatRowsEnd: 3)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var names = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
        Assert.Contains(names, name => name.Name?.Value == "GlobalRate" && name.Text == "0.25");
        Assert.Contains(names, name => name.Name?.Value == "_xlnm.Print_Area" && name.Text == "'O''Brien, East'!$A$1:$F$9");
        var titles = Assert.Single(names, name => name.Name?.Value == "_xlnm.Print_Titles");
        Assert.Equal<uint>(0, titles.LocalSheetId!.Value);
        Assert.Equal("'O''Brien, East'!$A:$B,'O''Brien, East'!$2:$3", titles.Text);
        var print = Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("print");
        Assert.Equal("'O''Brien, East'!$2:$3", print.GetProperty("repeatRows").GetString());
        Assert.Equal("'O''Brien, East'!$2:$3", print.GetProperty("normalizedRepeatRows").GetString());
        Assert.Null(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<PageSetup>());
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_sets_repeating_title_rows_and_columns_with_standard_cross_page_openxml()
    {
        var path = CreateWorkbookFixture();
        using (var source = SpreadsheetDocument.Open(path, true))
        {
            var workbook = source.WorkbookPart!.Workbook;
            workbook.Sheets!.Elements<Sheet>().Single().Name = "Neutral, West";
            workbook.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-repeat-titles-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation(
                "setPageSetup",
                Sheet: "Neutral, West",
                RepeatRowsStart: 3,
                RepeatRowsEnd: 4,
                RepeatColsStart: 2,
                RepeatColsEnd: 3)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var titles = Assert.Single(
            spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>(),
            name => name.Name?.Value == "_xlnm.Print_Titles");
        Assert.Equal("'Neutral, West'!$B:$C,'Neutral, West'!$3:$4", titles.Text);
        var print = Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("print");
        Assert.Equal("'Neutral, West'!$B:$C", print.GetProperty("repeatCols").GetString());
        Assert.Equal("'Neutral, West'!$B:$C", print.GetProperty("normalizedRepeatCols").GetString());
        Assert.Equal("'Neutral, West'!$3:$4", print.GetProperty("repeatRows").GetString());
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_cli_accepts_the_repeat_column_json_contract()
    {
        var path = CreateWorkbookFixture();
        var operations = Path.Combine(Path.GetTempPath(), $"xlsx-repeat-column-contract-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-repeat-column-contract-{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(operations,
            """{"operations":[{"type":"setPageSetup","sheet":"Sheet1","repeatColsStart":1,"repeatColsEnd":3}]}""");

        var exitCode = Editor.RunEdit([path, operations, output]);

        Assert.Equal(0, exitCode);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var titles = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>());
        Assert.Equal("'Sheet1'!$A:$C", titles.Text);
    }

    [Theory]
    [InlineData(1, 1, "$A:$A")]
    [InlineData(16384, 16384, "$XFD:$XFD")]
    public void Edit_accepts_boundary_repeating_title_columns(int startColumn, int endColumn, string expectedRange)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-repeat-column-boundary-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", RepeatColsStart: startColumn, RepeatColsEnd: endColumn)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var titles = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>());
        Assert.Equal($"'Sheet1'!{expectedRange}", titles.Text);
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_replaces_repeating_title_columns_without_removing_existing_title_rows()
    {
        var path = CreateWorkbookFixture();
        using (var source = SpreadsheetDocument.Open(path, true))
        {
            source.WorkbookPart!.Workbook.Append(new DefinedNames(
                new DefinedName("'Sheet1'!$D:$E,'Sheet1'!$5:$6") { Name = "_xlnm.Print_Titles", LocalSheetId = 0 }));
            source.WorkbookPart.Workbook.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-replace-repeat-columns-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", RepeatColsStart: 1, RepeatColsEnd: 2)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var titles = Assert.Single(spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>());
        Assert.Equal("'Sheet1'!$5:$6,'Sheet1'!$A:$B", titles.Text);
        var print = Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("print");
        Assert.Equal("'Sheet1'!$A:$B", print.GetProperty("repeatCols").GetString());
        Assert.Equal("'Sheet1'!$5:$6", print.GetProperty("repeatRows").GetString());
    }

    [Theory]
    [InlineData(null, 1)]
    [InlineData(1, null)]
    [InlineData(0, 1)]
    [InlineData(2, 1)]
    [InlineData(1, 16385)]
    public void Edit_rejects_invalid_repeating_title_columns(int? startColumn, int? endColumn)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-repeat-columns-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", RepeatColsStart: startColumn, RepeatColsEnd: endColumn)
        ]);

        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Theory]
    [InlineData(null, 1)]
    [InlineData(1, null)]
    [InlineData(0, 1)]
    [InlineData(2, 1)]
    [InlineData(1, 1048577)]
    public void Edit_rejects_invalid_repeating_title_rows(int? startRow, int? endRow)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-repeat-rows-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", RepeatRowsStart: startRow, RepeatRowsEnd: endRow)
        ]);

        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Edit_sets_exact_manual_row_page_breaks_and_exposes_them_as_evidence()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-row-breaks-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setRowPageBreaks", Sheet: "Sheet1", BreakBeforeRows: [12, 27])
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var rowBreaks = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<RowBreaks>()!;
        Assert.Equal<uint>(2, rowBreaks.Count!.Value);
        Assert.Equal<uint>(2, rowBreaks.ManualBreakCount!.Value);
        Assert.Equal([11U, 26U], rowBreaks.Elements<Break>().Select(item => item.Id!.Value));
        Assert.All(rowBreaks.Elements<Break>(), item => Assert.True(item.ManualPageBreak!.Value));
        var evidence = Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("print");
        Assert.Equal([12U, 27U], evidence.GetProperty("breakBeforeRows").EnumerateArray().Select(item => item.GetUInt32()));
        Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
    }

    [Fact]
    public void Edit_rejects_missing_duplicate_descending_or_unbounded_row_page_breaks()
    {
        var path = CreateWorkbookFixture();
        var invalid = new IReadOnlyList<int>?[] { null, [], [1], [12, 12], [27, 12], [1_048_577] };
        foreach (var rows in invalid)
        {
            var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-row-breaks-{Guid.NewGuid():N}.xlsx");
            var result = Editor.Apply(path, output, [
                new XlsxEditOperation("setRowPageBreaks", Sheet: "Sheet1", BreakBeforeRows: rows)
            ]);
            Assert.False(result.AppliedOperations.Single().Applied);
            Assert.False(File.Exists(output));
        }
    }

    [Fact]
    public void Edit_rejects_invalid_page_orientation()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-page-orientation-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", Orientation: "diagonal")
        ]);

        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Edit_rejects_invalid_paper_size()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-paper-size-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", PaperSize: "poster")
        ]);

        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    [InlineData(1, -1)]
    [InlineData(32768, 0)]
    public void Edit_rejects_invalid_fit_to_page_dimensions(int wide, int tall)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-page-setup-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", FitToPagesWide: wide, FitToPagesTall: tall)
        ]);

        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Edit_sets_one_column_width_and_preserves_adjacent_column_geometry()
    {
        var path = CreateWorkbookFixture();
        using (var spreadsheet = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
            worksheet.GetFirstChild<Columns>()?.Remove();
            worksheet.InsertBefore(
                new Columns(new Column { Min = 1, Max = 3, Width = 12, CustomWidth = true }),
                worksheet.GetFirstChild<SheetData>());
            worksheet.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-column-width-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setColumnWidth", Sheet: "Sheet1", Column: "B", Width: 24.5)
        ]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var edited = SpreadsheetDocument.Open(output, false);
        var columns = edited.WorkbookPart!.WorksheetParts.Single().Worksheet
            .GetFirstChild<Columns>()!.Elements<Column>().ToList();
        Assert.Collection(columns,
            column => { Assert.Equal<uint>(1, column.Min!.Value); Assert.Equal<uint>(1, column.Max!.Value); Assert.Equal(12, column.Width!.Value); },
            column => { Assert.Equal<uint>(2, column.Min!.Value); Assert.Equal<uint>(2, column.Max!.Value); Assert.Equal(24.5, column.Width!.Value); Assert.True(column.CustomWidth!.Value); },
            column => { Assert.Equal<uint>(3, column.Min!.Value); Assert.Equal<uint>(3, column.Max!.Value); Assert.Equal(12, column.Width!.Value); });
        Assert.Empty(new OpenXmlValidator().Validate(edited));
    }

    [Theory]
    [InlineData("A", 0)]
    [InlineData("A", 256)]
    [InlineData("XFE", 10)]
    [InlineData("A1", 10)]
    public void Edit_rejects_invalid_column_width_coordinates_before_mutation(string column, double width)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-column-width-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setColumnWidth", Sheet: "Sheet1", Column: column, Width: width)
        ]);

        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Theory]
    [InlineData("A0")]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("Sheet1!A1")]
    public void Edit_rejects_unbounded_cell_coordinates_before_mutation(string cell)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-cell-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: cell, Value: "must-not-write")]);
        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Edit_rejects_range_values_whose_derived_end_exceeds_sheet_bounds()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-derived-range-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [new XlsxEditOperation("setRangeValues", Sheet: "Sheet1", StartCell: "XFD1048576", Values: [["one", "overflow"]])]);
        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void Edit_preflights_the_entire_batch_without_overwriting_an_existing_output()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-preflight-existing-{Guid.NewGuid():N}.xlsx");
        var marker = new byte[] { 1, 2, 3, 4 };
        File.WriteAllBytes(output, marker);

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A1", Value: "would-be-partial"),
            new XlsxEditOperation("setPrintArea", Sheet: "Sheet1", Range: "F3:A1")
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.False(operation.Applied));
        Assert.Equal(marker, File.ReadAllBytes(output));
    }

    [Theory]
    [InlineData("Sheet1!A1:F3")]
    [InlineData("A0:F3")]
    [InlineData("A1:XFE3")]
    [InlineData("A1:F1048577")]
    [InlineData("A1:F999999999999999999999999")]
    [InlineData("F3:A1")]
    [InlineData("B1:A3")]
    [InlineData("A3:B1")]
    [InlineData("A1")]
    [InlineData("A1:")]
    [InlineData(" A1:F3")]
    public void Edit_rejects_invalid_print_areas_and_cli_returns_nonzero(string range)
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-print-area-{Guid.NewGuid():N}.xlsx");
        var operations = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-print-area-{Guid.NewGuid():N}.json");
        File.WriteAllText(operations, $$"""[{"type":"setPrintArea","sheet":"Sheet1","range":"{{range}}"}]""");

        Assert.Equal(1, Editor.RunEdit([path, operations, output]));
    }

    [Fact]
    public void Edit_replaces_only_the_target_sheet_print_area_and_quotes_sheet_names()
    {
        var path = CreateWorkbookFixture();
        using (var spreadsheet = SpreadsheetDocument.Open(path, true))
        {
            var workbookPart = spreadsheet.WorkbookPart!;
            var secondPart = workbookPart.AddNewPart<WorksheetPart>();
            secondPart.Worksheet = new Worksheet(new SheetData(CreateRow(1, ("A1", "Other"))));
            workbookPart.Workbook.Sheets!.Append(new Sheet { Id = workbookPart.GetIdOfPart(secondPart), SheetId = 2, Name = "O'Brien" });
            workbookPart.Workbook.DefinedNames = new DefinedNames(
                new DefinedName("'Sheet1'!$A$1:$B$2") { Name = "_xlnm.Print_Area", LocalSheetId = 0 },
                new DefinedName("'O''Brien'!$A$1:$C$3") { Name = "_xlnm.Print_Area", LocalSheetId = 1 },
                new DefinedName("'Sheet1'!$A$1") { Name = "keep_me" });
            workbookPart.Workbook.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-print-area-replace-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [new XlsxEditOperation("setPrintArea", Sheet: "O'Brien", Range: "$B$2:D4")]);

        Assert.True(result.AppliedOperations.Single().Applied);
        using var edited = SpreadsheetDocument.Open(output, false);
        var names = edited.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
        Assert.Equal(3, names.Count);
        Assert.Contains(names, name => name.Name!.Value == "keep_me");
        Assert.Contains(names, name => name.LocalSheetId?.Value == 0 && name.Text == "'Sheet1'!$A$1:$B$2");
        Assert.Contains(names, name => name.LocalSheetId?.Value == 1 && name.Text == "'O''Brien'!$B$2:$D$4");
    }

    [Fact]
    public void Edit_materializes_inherited_alignment_preserves_other_style_and_reuses_equivalent_cell_xf()
    {
        var path = CreateInheritedAlignmentWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-inherited-alignment-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A2", Value: "one", WrapText: false),
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "B2", Value: "two", WrapText: false)
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var styles = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
        var formats = styles.CellFormats!;
        Assert.Equal((uint)2, formats.Count!.Value);
        Assert.Equal(2, formats.Elements<CellFormat>().Count());
        var worksheet = spreadsheet.WorkbookPart.WorksheetParts.Single().Worksheet;
        Assert.Equal(GetCell(worksheet, "A2").StyleIndex!.Value, GetCell(worksheet, "B2").StyleIndex!.Value);
        var target = formats.Elements<CellFormat>().ElementAt(1);
        Assert.Equal(HorizontalAlignmentValues.Center, target.Alignment!.Horizontal!.Value);
        Assert.Equal(VerticalAlignmentValues.Top, target.Alignment.Vertical!.Value);
        Assert.False(target.Alignment.WrapText!.Value);
        Assert.True(target.Protection!.Locked!.Value);

        var cells = Inspector.InspectEvidence(output).GetProperty("evidence").GetProperty("sheets")[0].GetProperty("cells");
        var cell = cells.EnumerateArray().Single(item => item.GetProperty("reference").GetString() == "A2");
        Assert.Equal("center", cell.GetProperty("style").GetProperty("horizontalAlignment").GetString());
        Assert.Equal("top", cell.GetProperty("style").GetProperty("verticalAlignment").GetString());
        Assert.False(cell.GetProperty("style").GetProperty("wrapText").GetBoolean());
    }

    [Fact]
    public void Edit_sets_one_rich_text_value_and_explicitly_clears_bold_on_every_run()
    {
        var path = CreateWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-rich-value-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [new XlsxEditOperation("setRichTextCellValue", Sheet: "Sheet1", Cell: "D2", Value: "current value", Bold: false)]);
        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var cell = GetCell(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet, "D2");
        var run = Assert.Single(cell.InlineString!.Elements<Run>());
        Assert.Equal("current value", run.Text!.Text);
        Assert.False(run.RunProperties!.GetFirstChild<Bold>()!.Val!.Value);
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
    public void Edit_accepts_numeric_json_value_without_stringifying_it_upstream()
    {
        var path = CreateFormattedWorkbookFixture();
        var operations = Path.Combine(Path.GetTempPath(), $"xlsx-numeric-operations-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-number-json-{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(operations, "[{\"type\":\"setCellValue\",\"sheet\":\"Sheet1\",\"cell\":\"A2\",\"value\":11}]");

        Assert.Equal(0, Editor.RunEdit([path, operations, output]));

        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var cell = GetCell(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet, "A2");
        Assert.Null(cell.DataType);
        Assert.Equal("11", cell.CellValue!.Text);
    }

    [Fact]
    public void Edit_writes_iso_date_as_excel_serial_while_preserving_target_style()
    {
        var path = CreateFormattedWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-edited-date-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "A2", Value: "2026-07-19", ValueType: "date")]);
        Assert.True(result.AppliedOperations.Single().Applied);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var cell = GetCell(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet, "A2");
        Assert.Null(cell.DataType);
        Assert.Equal(DateTime.Parse("2026-07-19").ToOADate(), double.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal<UInt32Value>(1, cell.StyleIndex!);
    }

    [Fact]
    public void Edit_sets_explicit_number_format_on_existing_target_without_changing_value_or_other_style_components()
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-explicit-number-format-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellValue", Sheet: "Sheet1", Cell: "B2", Value: "2026-02-14", ValueType: "date"),
            new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", NumberFormat: "yyyy-mm-dd")
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var workbookPart = spreadsheet.WorkbookPart!;
        var cell = GetCell(workbookPart.WorksheetParts.Single().Worksheet, "B2");
        Assert.Null(cell.DataType);
        Assert.Equal(DateTime.Parse("2026-02-14").ToOADate(), double.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture));
        var format = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        Assert.Equal<UInt32Value>(2, format.FontId!);
        Assert.Equal<UInt32Value>(1, format.FillId!);
        Assert.Equal<UInt32Value>(1, format.BorderId!);
        Assert.Equal(HorizontalAlignmentValues.Center, format.Alignment!.Horizontal!.Value);
        var numberFormat = workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats!.Elements<NumberingFormat>()
            .Single(item => item.NumberFormatId!.Value == format.NumberFormatId!.Value);
        Assert.Equal("yyyy-mm-dd", numberFormat.FormatCode!.Value);
    }

    [Fact]
    public void Edit_copies_observed_number_format_from_same_workbook_peer()
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-peer-number-format-{Guid.NewGuid():N}.xlsx");

        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", SourceSheet: "Sheet1", SourceCell: "A2")
        ]);

        Assert.True(result.AppliedOperations.Single().Applied, result.AppliedOperations.Single().Detail);
        using var spreadsheet = SpreadsheetDocument.Open(output, false);
        var workbookPart = spreadsheet.WorkbookPart!;
        var worksheet = workbookPart.WorksheetParts.Single().Worksheet;
        var source = GetCell(worksheet, "A2");
        var target = GetCell(worksheet, "B2");
        var formats = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ToList();
        Assert.Equal(formats[(int)source.StyleIndex!.Value].NumberFormatId!.Value, formats[(int)target.StyleIndex!.Value].NumberFormatId!.Value);
        Assert.Equal("123", target.CellValue!.Text);
        Assert.Equal<UInt32Value>(2, formats[(int)target.StyleIndex.Value].FontId!);
        Assert.Equal<UInt32Value>(1, formats[(int)target.StyleIndex.Value].BorderId!);
    }

    [Theory]
    [InlineData("missing-target")]
    [InlineData("missing-source")]
    [InlineData("ambiguous-selector")]
    public void Edit_number_format_mutation_fails_closed_when_cell_binding_is_unproven(string variant)
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-number-format-{Guid.NewGuid():N}.xlsx");
        var operation = variant switch
        {
            "missing-target" => new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "C3", NumberFormat: "yyyy-mm-dd"),
            "missing-source" => new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", SourceSheet: "Sheet1", SourceCell: "C3"),
            _ => new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", NumberFormat: "yyyy-mm-dd", SourceSheet: "Sheet1", SourceCell: "A2"),
        };

        var result = Editor.Apply(path, output, [operation]);

        Assert.False(result.AppliedOperations.Single().Applied);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("yyyy-mm-dd\n")]
    [InlineData("yyyy-mm-dd\\")]
    [InlineData("yyyy-mm-dd_")]
    [InlineData("yyyy-mm-dd*")]
    [InlineData("[Red")]
    [InlineData("[]0")]
    [InlineData("[[Red]]0")]
    [InlineData("\"unterminated")]
    [InlineData("0;0;0;0;0")]
    public void Edit_rejects_invalid_explicit_number_format_grammar(string formatCode)
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-format-grammar-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", NumberFormat: formatCode)
        ]);
        Assert.False(result.AppliedOperations.Single().Applied);
    }

    [Theory]
    [InlineData("yyyy-mm-dd")]
    [InlineData("0.00;[Red]-0.00")]
    [InlineData("0.0 \"kg\"")]
    [InlineData("[$-409]mmm\\ d,\\ yyyy")]
    public void Edit_accepts_balanced_excel_number_format_grammar(string formatCode)
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-valid-format-grammar-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", NumberFormat: formatCode)
        ]);
        Assert.True(result.AppliedOperations.Single().Applied, result.AppliedOperations.Single().Detail);
    }

    [Theory]
    [InlineData("XFE1")]
    [InlineData("A1048577")]
    [InlineData("ZZZ9999999")]
    public void Edit_rejects_out_of_bounds_number_format_source_cell(string sourceCell)
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-format-source-coordinate-{Guid.NewGuid():N}.xlsx");
        var result = Editor.Apply(path, output, [
            new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", SourceSheet: "Sheet1", SourceCell: sourceCell)
        ]);
        Assert.False(result.AppliedOperations.Single().Applied);
    }

    [Theory]
    [InlineData("target")]
    [InlineData("source")]
    public void Edit_rejects_out_of_range_style_indices_in_number_format_bindings(string invalidCell)
    {
        var path = CreateNumberFormatMutationWorkbookFixture();
        using (var spreadsheet = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet;
            GetCell(worksheet, invalidCell == "target" ? "B2" : "A2").StyleIndex = 999;
            worksheet.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-format-style-{Guid.NewGuid():N}.xlsx");
        var operation = invalidCell == "target"
            ? new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", NumberFormat: "yyyy-mm-dd")
            : new XlsxEditOperation("setCellNumberFormat", Sheet: "Sheet1", Cell: "B2", SourceSheet: "Sheet1", SourceCell: "A2");
        var result = Editor.Apply(path, output, [operation]);
        Assert.False(result.AppliedOperations.Single().Applied);
        Assert.Contains("style invalid", result.AppliedOperations.Single().Detail, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_operation_keeps_new_number_format_fields_after_the_legacy_positional_surface()
    {
        var parameters = typeof(XlsxEditOperation).GetConstructors().Single().GetParameters().Select(parameter => parameter.Name).ToList();
        var breakBeforeRows = parameters.FindIndex(name => string.Equals(name, "BreakBeforeRows", StringComparison.OrdinalIgnoreCase));
        Assert.True(parameters.FindIndex(name => string.Equals(name, "NumberFormat", StringComparison.OrdinalIgnoreCase)) > breakBeforeRows);
        Assert.True(parameters.FindIndex(name => string.Equals(name, "SourceSheet", StringComparison.OrdinalIgnoreCase)) > breakBeforeRows);
        Assert.True(parameters.FindIndex(name => string.Equals(name, "SourceCell", StringComparison.OrdinalIgnoreCase)) > breakBeforeRows);
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

    private static string CreateNumberFormatMutationWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-number-format-mutation-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new NumberingFormats(new NumberingFormat { NumberFormatId = 164, FormatCode = "dd-mmm-yyyy" }) { Count = 1 },
            new Fonts(new Font(), new Font(), new Font(new Bold())) { Count = 3 },
            new Fills(new Fill(), new Fill(new PatternFill { PatternType = PatternValues.Solid })) { Count = 2 },
            new Borders(new Border(), new Border(new LeftBorder { Style = BorderStyleValues.Thin })) { Count = 2 },
            new CellFormats(
                new CellFormat { NumberFormatId = 0, ApplyNumberFormat = false },
                new CellFormat { NumberFormatId = 164, ApplyNumberFormat = true, FontId = 1 },
                new CellFormat {
                    NumberFormatId = 0,
                    ApplyNumberFormat = false,
                    FontId = 2,
                    FillId = 1,
                    BorderId = 1,
                    ApplyAlignment = true,
                    Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center }
                }
            ) { Count = 3 }
        );
        stylesPart.Stylesheet.Save();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("A1", "Peer date"), ("B1", "Target")),
            new Row(
                new Cell { CellReference = "A2", StyleIndex = 1, CellValue = new CellValue("46067") },
                new Cell { CellReference = "B2", StyleIndex = 2, CellValue = new CellValue("123") }
            ) { RowIndex = 2 }
        ));
        workbookPart.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
        workbookPart.Workbook.Save();
        worksheetPart.Worksheet.Save();
        return path;
    }

    private static string CreateInheritedAlignmentWorkbookFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-inherited-alignment-fixture-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat {
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true }
            }) { Count = 1 },
            new CellFormats(new CellFormat {
                FormatId = 0,
                Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Right },
                Protection = new Protection { Locked = true },
                ApplyAlignment = false,
                ApplyProtection = true
            }) { Count = 1 }
        );
        stylesPart.Stylesheet.Save();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(
            CreateRow(1, ("A1", "A"), ("B1", "B")),
            new Row(new Cell { CellReference = "A2", StyleIndex = 0 }, new Cell { CellReference = "B2", StyleIndex = 0 }) { RowIndex = 2 }
        ));
        workbookPart.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
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
