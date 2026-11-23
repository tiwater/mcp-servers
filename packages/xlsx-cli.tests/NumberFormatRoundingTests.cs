using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using Xunit;

namespace Dockit.Xlsx.Tests;

public sealed class NumberFormatRoundingTests
{
    [Theory]
    [InlineData(1.25, "0.0", "1.3")]
    [InlineData(1.25, @"0.0_);[Red]\(0.0\)", "1.3")]
    [InlineData(-1.25, "0.0", "-1.3")]
    [InlineData(-1.25, @"0.0_);[Red]\(0.0\)", "-1.3")]
    [InlineData(0.05, "0.0", "0.1")]
    [InlineData(-0.05, "0.0", "-0.1")]
    [InlineData(1.15, "0.0", "1.2")]
    [InlineData(-1.15, "0.0", "-1.2")]
    [InlineData(1.005, "0.00", "1.01")]
    [InlineData(2.675, "0.00", "2.68")]
    [InlineData(1.255, "0.00", "1.26")]
    [InlineData(1.245, "0.00", "1.25")]
    [InlineData(0.5, "0", "1")]
    [InlineData(2.5, "0", "3")]
    [InlineData(-0.5, "0", "-1")]
    [InlineData(-2.5, "0", "-3")]
    [InlineData(0.5, "#,##0", "1")]
    [InlineData(2.5, "#,##0", "3")]
    [InlineData(1.25, "0.000", "1.250")]
    [InlineData(0.0, "0.0", "0.0")]
    [InlineData(-0.04, "0.0", "0.0")]
    [InlineData(9.95, "0.0", "10.0")]
    [InlineData(-9.95, "0.0", "-10.0")]
    [InlineData(1.24999999999999, "0.0", "1.2")]
    [InlineData(1.2499999999999998, "0.0", "1.3")]
    [InlineData(1.255, @"0.00"" mg/mL""", "1.26")]
    [InlineData(-1.255, "0.0;0.00", "-1.26")]
    [InlineData(0.0, "0.0;0.0;0.000", "0.000")]
    public void CustomAndStandardNumericFormats_UseAwayFromZeroRounding(double value, string formatCode, string expected)
    {
        var path = CreateWorkbookWithSingleNumber(value, customFormatCode: formatCode);
        try
        {
            var exported = ExportFirstSheetCell(path, "A1");
            Assert.Equal(expected, exported.FormattedValue);

            var range = RangeReader.Read(path, "Sheet1", "A1:A1", 0, 10);
            Assert.Equal(expected, range.Cells.Single().FormattedValue);

            var inventory = RegionInventory.Inspect(path);
            Assert.Equal(expected, inventory.Sheets.Single().Regions.Single().Rows.Single().Cells.Single().FormattedValue);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Theory]
    [InlineData(0.5, 1U, "1")]       // built-in 1: "0"
    [InlineData(2.5, 1U, "3")]       // built-in 1: "0"
    [InlineData(-2.5, 1U, "-3")]     // built-in 1: "0"
    [InlineData(1.005, 2U, "1.01")]  // built-in 2: "0.00"
    [InlineData(1.255, 2U, "1.26")]  // built-in 2: "0.00"
    [InlineData(2.5, 3U, "3")]       // built-in 3: "#,##0"
    [InlineData(1.255, 4U, "1.26")]  // built-in 4: "#,##0.00"
    [InlineData(0.005, 9U, "1%")]    // built-in 9: "0%"
    [InlineData(0.025, 9U, "3%")]    // built-in 9: "0%"
    [InlineData(0.145, 9U, "15%")]   // built-in 9: "0%"
    [InlineData(0.01255, 10U, "1.26%")]  // built-in 10: "0.00%"
    [InlineData(-0.01255, 10U, "-1.26%")] // built-in 10: "0.00%"
    [InlineData(1255.0, 11U, "1.26E+03")] // built-in 11: "0.00E+00"
    [InlineData(-1255.0, 11U, "-1.26E+03")] // built-in 11: "0.00E+00"
    [InlineData(12550.0, 48U, "12.6E+3")] // built-in 48: "##0.0E+0"
    [InlineData(999.95, 48U, "1.0E+3")]   // built-in 48 carry
    [InlineData(1.25, 0U, "1.25")]   // built-in 0: "General"
    [InlineData(1.25, 49U, "1.25")]  // built-in 49: "@"
    public void BuiltInNumericFormats_UseAwayFromZeroRounding(double value, uint builtInFormatId, string expected)
    {
        var path = CreateWorkbookWithSingleNumber(value, builtInFormatId: builtInFormatId);
        try
        {
            var exported = ExportFirstSheetCell(path, "A1");
            Assert.Equal(expected, exported.FormattedValue);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Theory]
    [InlineData(0.0125, "0.0%", "1.3%")]
    [InlineData(-0.0125, "0.0%", "-1.3%")]
    [InlineData(0.0005, "0.0%", "0.1%")]
    [InlineData(-0.0005, "0.0%", "-0.1%")]
    [InlineData(0.005, "0%", "1%")]
    [InlineData(0.145, "0%", "15%")]
    [InlineData(0.12555, "0.00%", "12.56%")]
    [InlineData(1.25, @"0.0""%""", "1.3")] // quoted literal '%' must not scale by 100
    public void PercentageFormats_UseAwayFromZeroRounding(double value, string formatCode, string expected)
    {
        var path = CreateWorkbookWithSingleNumber(value, customFormatCode: formatCode);
        try
        {
            var exported = ExportFirstSheetCell(path, "A1");
            Assert.Equal(expected, exported.FormattedValue);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Theory]
    [InlineData(1250.0, "0.0E+00", "1.3E+03")]
    [InlineData(-1250.0, "0.0E+00", "-1.3E+03")]
    [InlineData(1255.0, "0.00E+00", "1.26E+03")]
    [InlineData(0.00125, "0.0E+00", "1.3E-03")]
    [InlineData(-0.00125, "0.0E+00", "-1.3E-03")]
    [InlineData(9.95, "0.0E+00", "1.0E+01")]
    [InlineData(-9.95, "0.0E+00", "-1.0E+01")]
    [InlineData(0.0, "0.00E+00", "0.00E+00")]
    [InlineData(1250.0, "0.0E-00", "1.3E03")]
    [InlineData(0.00125, "0.0E-00", "1.3E-03")]
    [InlineData(1.25e30, "0.0E+00", "1.3E+30")]
    [InlineData(1.25e-30, "0.0E+00", "1.3E-30")]
    [InlineData(500.0, "0E+00", "5E+02")]
    [InlineData(2500.0, "0E+00", "3E+03")]
    public void ScientificFormats_UseAwayFromZeroRounding(double value, string formatCode, string expected)
    {
        var path = CreateWorkbookWithSingleNumber(value, customFormatCode: formatCode);
        try
        {
            var exported = ExportFirstSheetCell(path, "A1");
            Assert.Equal(expected, exported.FormattedValue);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Theory]
    [InlineData(1.25, "0.0", "1.3")]
    [InlineData(1.25, @"0.0_);[Red]\(0.0\)", "1.3")]
    [InlineData(-1.25, "0.0", "-1.3")]
    [InlineData(0.05, "0.0", "0.1")]
    [InlineData(0.0125, "0.0%", "1.3%")]
    [InlineData(1250.0, "0.0E+00", "1.3E+03")]
    public void LegacyXlsNumericFormats_UseAwayFromZeroRounding(double value, string formatCode, string expected)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"xlsx-cli-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "legacy.xls");
        try
        {
            var wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);
            var style = wb.CreateCellStyle();
            var fmt = wb.CreateDataFormat();
            style.DataFormat = fmt.GetFormat(formatCode);
            cell.CellStyle = style;
            cell.SetCellValue(value);

            using (var fs = File.Create(path))
            {
                wb.Write(fs);
            }

            var exported = ExportFirstSheetCell(path, "A1");
            Assert.Equal(expected, exported.FormattedValue);
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }
    }

    private static (string Value, string FormattedValue) ExportFirstSheetCell(string path, string reference)
    {
        var result = Extractor.Export(path);
        var json = JsonSerializer.Serialize(result, Json.Options);
        using var doc = JsonDocument.Parse(json);
        var cells = doc.RootElement[0].GetProperty("cells");
        foreach (var cell in cells.EnumerateArray())
        {
            if (string.Equals(cell.GetProperty("reference").GetString(), reference, StringComparison.OrdinalIgnoreCase))
            {
                return (
                    cell.GetProperty("value").GetString() ?? string.Empty,
                    cell.GetProperty("formattedValue").GetString() ?? string.Empty);
            }
        }
        throw new InvalidOperationException($"Cell {reference} not found.");
    }

    private static string CreateWorkbookWithSingleNumber(
        double value,
        string? customFormatCode = null,
        uint? builtInFormatId = null)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"xlsx-cli-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "test.xlsx");

        using (var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var numberingFormats = new NumberingFormats();
            uint numFmtId;
            if (customFormatCode is not null)
            {
                numFmtId = 164U;
                numberingFormats.Append(new NumberingFormat
                {
                    NumberFormatId = numFmtId,
                    FormatCode = customFormatCode
                });
                numberingFormats.Count = 1U;
            }
            else
            {
                numFmtId = builtInFormatId ?? 0U;
            }

            stylesPart.Stylesheet = new Stylesheet(
                numberingFormats,
                new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })),
                new Borders(new Border()),
                new CellFormats(
                    new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U },
                    new CellFormat
                    {
                        NumberFormatId = numFmtId,
                        FontId = 0U,
                        FillId = 0U,
                        BorderId = 0U,
                        ApplyNumberFormat = true
                    }));
            stylesPart.Stylesheet.Save();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData(
                new Row(
                    new Cell
                    {
                        CellReference = "A1",
                        StyleIndex = 1U,
                        CellValue = new CellValue(value.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
                    })
                { RowIndex = 1U });
            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Save();

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1U,
                Name = "Sheet1"
            });
            workbookPart.Workbook.Save();
        }

        return path;
    }

    private static void Cleanup(string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (dir is not null && Directory.Exists(dir)) Directory.Delete(dir, true);
    }
}
