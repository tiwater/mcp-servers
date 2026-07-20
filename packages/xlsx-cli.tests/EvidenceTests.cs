using Dockit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.Json;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class EvidenceTests
{
    private static readonly JsonSerializerOptions Options = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    [Fact]
    public void Evidence_is_versioned_and_exposes_style_formula_merge_view_and_print_facts()
    {
        var path = Fixture();
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));
        var root = json.RootElement;
        Assert.Equal("tiwater.xlsx.evidence/v1", root.GetProperty("schema").GetString());
        Assert.Equal("1904", root.GetProperty("dateSystem").GetString());
        var sheet = root.GetProperty("sheets")[0];
        Assert.Equal("A1:B2", sheet.GetProperty("dimension").GetString());
        Assert.Equal("A1:B1", sheet.GetProperty("mergedRanges")[0].GetString());
        Assert.Equal("pageBreakPreview", sheet.GetProperty("sheetView").GetProperty("view").GetString(), ignoreCase: true);
        Assert.Equal("landscape", sheet.GetProperty("print").GetProperty("orientation").GetString(), ignoreCase: true);
        var cell = sheet.GetProperty("cells").EnumerateArray().Single(x => x.GetProperty("reference").GetString() == "B2");
        Assert.Equal((uint)1, cell.GetProperty("styleIndex").GetUInt32());
        Assert.Equal("0.00", cell.GetProperty("style").GetProperty("numberFormat").GetString());
        var numberFormat = cell.GetProperty("style").GetProperty("numberFormatEvidence");
        Assert.Equal("builtIn", numberFormat.GetProperty("source").GetString());
        Assert.Equal("number", numberFormat.GetProperty("kind").GetString());
        Assert.False(numberFormat.GetProperty("isDateLike").GetBoolean());
        Assert.Equal(64, cell.GetProperty("style").GetProperty("numberFormatFingerprint").GetString()!.Length);
        Assert.Equal("center", cell.GetProperty("style").GetProperty("horizontalAlignment").GetString(), ignoreCase: true);
        Assert.Equal("A2*2", cell.GetProperty("formula").GetProperty("text").GetString());
        var date1904 = sheet.GetProperty("cells").EnumerateArray().Single(x => x.GetProperty("reference").GetString() == "A2");
        Assert.Equal("date", date1904.GetProperty("style").GetProperty("numberFormatEvidence").GetProperty("kind").GetString());
        Assert.Equal("1904-01-03", date1904.GetProperty("normalizedValue").GetProperty("iso8601").GetString());
    }

    [Fact]
    public void Evidence_normalizes_builtin_and_custom_dates_with_effective_inherited_styles()
    {
        var path = DateFixture();
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));
        var root = json.RootElement;
        Assert.Equal("1900", root.GetProperty("dateSystem").GetString());
        var cells = root.GetProperty("sheets")[0].GetProperty("cells").EnumerateArray().ToDictionary(x => x.GetProperty("reference").GetString()!);

        var builtin = cells["A1"];
        Assert.Equal("date", builtin.GetProperty("style").GetProperty("numberFormatEvidence").GetProperty("kind").GetString());
        Assert.Equal("builtIn", builtin.GetProperty("style").GetProperty("numberFormatEvidence").GetProperty("source").GetString());
        Assert.Equal("2026-07-19", builtin.GetProperty("normalizedValue").GetProperty("iso8601").GetString());
        Assert.False(string.IsNullOrWhiteSpace(builtin.GetProperty("formattedValue").GetString()));

        var inherited = cells["B1"];
        var inheritedStyle = inherited.GetProperty("style");
        Assert.Equal((uint)1, inheritedStyle.GetProperty("baseStyleIndex").GetUInt32());
        Assert.Equal((uint)164, inheritedStyle.GetProperty("numberFormatId").GetUInt32());
        Assert.Equal("custom", inheritedStyle.GetProperty("numberFormatEvidence").GetProperty("source").GetString());
        Assert.Equal("datetime", inheritedStyle.GetProperty("numberFormatEvidence").GetProperty("kind").GetString());
        Assert.Equal("2026-07-19T12:30:00", inherited.GetProperty("normalizedValue").GetProperty("iso8601").GetString());
    }

    [Fact]
    public void Evidence_changes_when_a_baseline_fact_is_tampered()
    {
        var path = Fixture();
        var before = JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options);
        using (var doc = SpreadsheetDocument.Open(path, true)) doc.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!.ShowGridLines = true;
        var after = JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options);
        Assert.NotEqual(before, after);
    }

    [Fact]
    public void Resolved_style_fingerprints_ignore_component_ids_but_detect_component_changes()
    {
        var baseline = CellStyleEvidence(StyleFixture(reordered: false, changed: false));
        var reordered = CellStyleEvidence(StyleFixture(reordered: true, changed: false));
        var changed = CellStyleEvidence(StyleFixture(reordered: true, changed: true));

        Assert.NotEqual(baseline.GetProperty("fontId").GetUInt32(), reordered.GetProperty("fontId").GetUInt32());
        Assert.NotEqual(baseline.GetProperty("fillId").GetUInt32(), reordered.GetProperty("fillId").GetUInt32());
        Assert.NotEqual(baseline.GetProperty("borderId").GetUInt32(), reordered.GetProperty("borderId").GetUInt32());
        foreach (var component in new[] { "font", "fill", "border", "protection" })
        {
            var fingerprint = $"{component}Fingerprint";
            Assert.Equal(baseline.GetProperty(fingerprint).GetString(), reordered.GetProperty(fingerprint).GetString());
            Assert.NotEqual(baseline.GetProperty(fingerprint).GetString(), changed.GetProperty(fingerprint).GetString());
        }
        Assert.False(baseline.GetProperty("protection").GetProperty("locked").GetBoolean());
        Assert.True(baseline.GetProperty("protection").GetProperty("hidden").GetBoolean());
    }

    [Fact]
    public void Evidence_exposes_all_global_and_sheet_local_defined_names()
    {
        var path = DefinedNamesFixture();
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));

        var names = json.RootElement.GetProperty("definedNames").EnumerateArray().ToList();
        Assert.Equal(3, names.Count);
        Assert.Contains(names, name =>
            name.GetProperty("name").GetString() == "GlobalRate"
            && name.GetProperty("localSheetName").ValueKind == JsonValueKind.Null
            && name.GetProperty("text").GetString() == "0.25"
            && !name.GetProperty("hidden").GetBoolean());
        Assert.Contains(names, name =>
            name.GetProperty("name").GetString() == "LocalInput"
            && name.GetProperty("localSheetName").GetString() == "Second"
            && name.GetProperty("text").GetString() == "'Second'!$A$1"
            && name.GetProperty("hidden").GetBoolean());
        Assert.Contains(names, name =>
            name.GetProperty("name").GetString() == "_xlnm.Print_Area"
            && name.GetProperty("localSheetName").GetString() == "First"
            && name.GetProperty("text").GetString() == "'First'!$A$1:$B$2");
    }

    private static JsonElement CellStyleEvidence(string path)
    {
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));
        return json.RootElement.GetProperty("sheets")[0].GetProperty("cells")[0].GetProperty("style").Clone();
    }

    private static string StyleFixture(bool reordered, bool changed)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-style-evidence-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var styles = wb.AddNewPart<WorkbookStylesPart>();
        var targetFont = new Font(new Bold(), new FontName { Val = "Evidence Font" }, new Color { Rgb = changed ? "FF0000FF" : "FFFF0000" });
        var targetFill = new Fill(new PatternFill(new ForegroundColor { Rgb = changed ? "FF00FFFF" : "FFFFFF00" }) { PatternType = PatternValues.Solid });
        var targetBorder = new Border(new BottomBorder(new Color { Rgb = changed ? "FFFF00FF" : "FF00FF00" }) { Style = BorderStyleValues.Thin });
        var fonts = reordered ? new Fonts(new Font(), new Font(new Italic()), targetFont) : new Fonts(new Font(), targetFont);
        var fills = reordered ? new Fills(new Fill(), new Fill(new PatternFill { PatternType = PatternValues.Gray125 }), targetFill) : new Fills(new Fill(), targetFill);
        var borders = reordered ? new Borders(new Border(), new Border(new LeftBorder { Style = BorderStyleValues.Dashed }), targetBorder) : new Borders(new Border(), targetBorder);
        fonts.Count = (uint)fonts.ChildElements.Count;
        fills.Count = (uint)fills.ChildElements.Count;
        borders.Count = (uint)borders.ChildElements.Count;
        var targetId = reordered ? 2U : 1U;
        styles.Stylesheet = new Stylesheet(
            fonts,
            fills,
            borders,
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(
                new CellFormat(),
                new CellFormat {
                    FontId = targetId,
                    FillId = targetId,
                    BorderId = targetId,
                    ApplyProtection = true,
                    Protection = new Protection { Locked = changed, Hidden = !changed }
                }) { Count = 2 });
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetData(new Row(new Cell { CellReference = "A1", StyleIndex = 1, CellValue = new CellValue("1") }) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Styles" });
        wb.Workbook.Save();
        styles.Stylesheet.Save();
        ws.Worksheet.Save();
        return path;
    }

    private static string DefinedNamesFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-defined-names-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var first = wb.AddNewPart<WorksheetPart>();
        first.Worksheet = new Worksheet(new SheetData());
        var second = wb.AddNewPart<WorksheetPart>();
        second.Worksheet = new Worksheet(new SheetData());
        wb.Workbook.Append(
            new Sheets(
                new Sheet { Id = wb.GetIdOfPart(first), SheetId = 1, Name = "First" },
                new Sheet { Id = wb.GetIdOfPart(second), SheetId = 2, Name = "Second" }),
            new DefinedNames(
                new DefinedName("0.25") { Name = "GlobalRate" },
                new DefinedName("'Second'!$A$1") { Name = "LocalInput", LocalSheetId = 1, Hidden = true },
                new DefinedName("'First'!$A$1:$B$2") { Name = "_xlnm.Print_Area", LocalSheetId = 0 }));
        wb.Workbook.Save();
        first.Worksheet.Save();
        second.Worksheet.Save();
        return path;
    }

    private static string Fixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-evidence-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart(); wb.Workbook = new Workbook(new WorkbookProperties { Date1904 = true });
        var styles = wb.AddNewPart<WorkbookStylesPart>(); styles.Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 2, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }, new CellFormat { NumberFormatId = 14 }));
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetDimension { Reference = "A1:B2" }, new SheetViews(new SheetView { WorkbookViewId = 0, View = SheetViewValues.PageBreakPreview, ShowGridLines = false }), new SheetData(new Row(new Cell { CellReference = "A2", StyleIndex = 2, CellValue = new CellValue("2") }, new Cell { CellReference = "B2", StyleIndex = 1, CellFormula = new CellFormula("A2*2"), CellValue = new CellValue("4") }) { RowIndex = 2 }), new MergeCells(new MergeCell { Reference = "A1:B1" }), new PageMargins { Left = .7, Right = .7, Top = .75, Bottom = .75, Header = .3, Footer = .3 }, new PageSetup { Orientation = OrientationValues.Landscape });
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Report" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }

    private static string DateFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-date-evidence-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart(); wb.Workbook = new Workbook(new WorkbookProperties { Date1904 = false });
        var styles = wb.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new NumberingFormats(new NumberingFormat { NumberFormatId = 164, FormatCode = "yyyy-mm-dd hh:mm" }) { Count = 1 },
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(
                new CellFormat { NumberFormatId = 0 },
                new CellFormat { NumberFormatId = 164, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 2 },
            new CellFormats(
                new CellFormat(),
                new CellFormat { NumberFormatId = 14 },
                new CellFormat { FormatId = 1 }) { Count = 3 });
        var ws = wb.AddNewPart<WorksheetPart>();
        var date = new DateTime(2026, 7, 19);
        var dateTime = date.AddHours(12.5);
        ws.Worksheet = new Worksheet(new SheetData(new Row(
            new Cell { CellReference = "A1", StyleIndex = 1, CellValue = new CellValue(ExcelSerial(date).ToString(CultureInfo.InvariantCulture)) },
            new Cell { CellReference = "B1", StyleIndex = 2, CellValue = new CellValue(ExcelSerial(dateTime).ToString(CultureInfo.InvariantCulture)) }) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Dates" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }

    private static double ExcelSerial(DateTime value)
    {
        var days = (value - new DateTime(1899, 12, 31)).TotalDays;
        return days >= 60 ? days + 1 : days;
    }
}
