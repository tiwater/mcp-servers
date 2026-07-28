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
            new XlsxEditOperation("setPageSetup", Sheet: "Sheet1", FitToPagesWide: 1, FitToPagesTall: 0)
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

    [Theory]
    [InlineData(0, 0)]
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
