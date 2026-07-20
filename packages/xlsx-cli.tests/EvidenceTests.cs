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
    public void Effective_style_components_inherit_from_base_xf_when_apply_flags_are_false()
    {
        var path = InheritedStyleFixture();
        using (var document = SpreadsheetDocument.Open(path, false))
        {
            var raw = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt(1);
            var rawCells = document.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToList();
            Assert.Equal(1U, rawCells[0].StyleIndex!.Value);
            Assert.Equal(2U, rawCells[1].StyleIndex!.Value);
            Assert.Equal(2U, raw.FontId!.Value);
            Assert.Equal(2U, raw.FillId!.Value);
            Assert.Equal(2U, raw.BorderId!.Value);
            Assert.Equal(165U, raw.NumberFormatId!.Value);
            Assert.False(raw.ApplyFont!.Value);
            Assert.False(raw.ApplyFill!.Value);
            Assert.False(raw.ApplyBorder!.Value);
            Assert.False(raw.ApplyNumberFormat!.Value);
        }
        var styles = CellStylesEvidence(path);
        var inherited = styles["A1"];
        var direct = styles["B1"];

        Assert.Equal(1U, inherited.GetProperty("fontId").GetUInt32());
        Assert.Equal(1U, inherited.GetProperty("fillId").GetUInt32());
        Assert.Equal(1U, inherited.GetProperty("borderId").GetUInt32());
        Assert.Equal(164U, inherited.GetProperty("numberFormatId").GetUInt32());
        foreach (var component in new[] { "font", "fill", "border", "protection", "numberFormat", "alignment" })
            Assert.Equal(inherited.GetProperty($"{component}Fingerprint").GetString(), direct.GetProperty($"{component}Fingerprint").GetString());
        Assert.False(inherited.GetProperty("protection").GetProperty("applyProtection").GetBoolean());
        Assert.True(direct.GetProperty("protection").GetProperty("applyProtection").GetBoolean());
    }

    [Fact]
    public void Number_format_fingerprint_ignores_id_but_preserves_literal_case_and_whitespace()
    {
        var baseline = CellStyleEvidence(NumberFormatFixture(164, "0 \"KG\""));
        var renumbered = CellStyleEvidence(NumberFormatFixture(200, "0 \"KG\""));
        var literalCaseChanged = CellStyleEvidence(NumberFormatFixture(201, "0 \"kg\""));
        var trailingSpaceChanged = CellStyleEvidence(NumberFormatFixture(202, "0 \"KG\" "));

        Assert.NotEqual(baseline.GetProperty("numberFormatId").GetUInt32(), renumbered.GetProperty("numberFormatId").GetUInt32());
        Assert.Equal(baseline.GetProperty("numberFormatFingerprint").GetString(), renumbered.GetProperty("numberFormatFingerprint").GetString());
        Assert.NotEqual(baseline.GetProperty("numberFormatFingerprint").GetString(), literalCaseChanged.GetProperty("numberFormatFingerprint").GetString());
        Assert.NotEqual(baseline.GetProperty("numberFormatFingerprint").GetString(), trailingSpaceChanged.GetProperty("numberFormatFingerprint").GetString());
    }

    [Fact]
    public void Wps_locale_short_date_ids_share_semantics_without_hiding_general_drift()
    {
        var shortDate = CellStyleEvidence(BuiltInNumberFormatFixture(14));
        var wpsSavedShortDate = CellStyleEvidence(BuiltInNumberFormatFixture(58));
        var general = CellStyleEvidence(BuiltInNumberFormatFixture(0));

        Assert.Equal(14U, shortDate.GetProperty("numberFormatId").GetUInt32());
        Assert.Equal(58U, wpsSavedShortDate.GetProperty("numberFormatId").GetUInt32());
        Assert.Equal("m/d/yy", shortDate.GetProperty("numberFormat").GetString());
        Assert.Equal("builtin:58", wpsSavedShortDate.GetProperty("numberFormat").GetString());
        Assert.Equal("date", shortDate.GetProperty("numberFormatEvidence").GetProperty("kind").GetString());
        Assert.Equal("date", wpsSavedShortDate.GetProperty("numberFormatEvidence").GetProperty("kind").GetString());
        Assert.Equal("wps-locale-short-date", shortDate.GetProperty("numberFormatEvidence").GetProperty("normalizedCode").GetString());
        Assert.Equal("wps-locale-short-date", wpsSavedShortDate.GetProperty("numberFormatEvidence").GetProperty("normalizedCode").GetString());
        Assert.Equal(shortDate.GetProperty("numberFormatFingerprint").GetString(), wpsSavedShortDate.GetProperty("numberFormatFingerprint").GetString());
        Assert.NotEqual(shortDate.GetProperty("numberFormatFingerprint").GetString(), general.GetProperty("numberFormatFingerprint").GetString());
    }

    [Theory]
    [InlineData("horizontal")]
    [InlineData("vertical")]
    [InlineData("textRotation")]
    [InlineData("wrapText")]
    [InlineData("shrinkToFit")]
    [InlineData("indent")]
    [InlineData("relativeIndent")]
    [InlineData("justifyLastLine")]
    [InlineData("readingOrder")]
    [InlineData("mergeCell")]
    public void Alignment_fingerprint_detects_every_openxml_alignment_semantic(string changedProperty)
    {
        var baseline = CellStyleEvidence(AlignmentFixture());
        var changed = CellStyleEvidence(AlignmentFixture(changedProperty));

        Assert.NotEqual(baseline.GetProperty("alignmentFingerprint").GetString(), changed.GetProperty("alignmentFingerprint").GetString());
        var alignment = baseline.GetProperty("alignment");
        foreach (var property in new[] { "horizontal", "vertical", "textRotation", "wrapText", "shrinkToFit", "indent", "relativeIndent", "justifyLastLine", "readingOrder", "mergeCell" })
            Assert.True(alignment.TryGetProperty(property, out _), $"Missing alignment semantic: {property}");
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

    [Fact]
    public void Defined_name_evidence_is_order_stable_and_detects_text_or_visibility_changes()
    {
        var baseline = DefinedNamesEvidence(DefinedNamesFixture());
        var reordered = DefinedNamesEvidence(DefinedNamesFixture(reversed: true));
        var nameChanged = DefinedNamesEvidence(DefinedNamesFixture(localName: "LocalOutput"));
        var scopeChanged = DefinedNamesEvidence(DefinedNamesFixture(localSheetId: 0));
        var textChanged = DefinedNamesEvidence(DefinedNamesFixture(localText: "'Second'!$B$2"));
        var visibilityChanged = DefinedNamesEvidence(DefinedNamesFixture(localHidden: false));

        Assert.Equal(baseline, reordered);
        Assert.NotEqual(baseline, nameChanged);
        Assert.NotEqual(baseline, scopeChanged);
        Assert.NotEqual(baseline, textChanged);
        Assert.NotEqual(baseline, visibilityChanged);
    }

    private static JsonElement CellStyleEvidence(string path)
    {
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));
        return json.RootElement.GetProperty("sheets")[0].GetProperty("cells")[0].GetProperty("style").Clone();
    }

    private static Dictionary<string, JsonElement> CellStylesEvidence(string path)
    {
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));
        return json.RootElement.GetProperty("sheets")[0].GetProperty("cells").EnumerateArray()
            .ToDictionary(cell => cell.GetProperty("reference").GetString()!, cell => cell.GetProperty("style").Clone());
    }

    private static string DefinedNamesEvidence(string path)
    {
        using var json = JsonDocument.Parse(JsonSerializer.Serialize(EvidenceInspector.Inspect(path), Options));
        return json.RootElement.GetProperty("definedNames").GetRawText();
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

    private static string InheritedStyleFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-inherited-style-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var styles = wb.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new NumberingFormats(
                new NumberingFormat { NumberFormatId = 164, FormatCode = "0 \"BASE\"" },
                new NumberingFormat { NumberFormatId = 165, FormatCode = "0 \"DIRECT\"" }) { Count = 2 },
            new Fonts(new Font(), new Font(new Bold()), new Font(new Italic())) { Count = 3 },
            new Fills(new Fill(), new Fill(new PatternFill { PatternType = PatternValues.Solid }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 3 },
            new Borders(new Border(), new Border(new BottomBorder { Style = BorderStyleValues.Thin }), new Border(new TopBorder { Style = BorderStyleValues.Dashed })) { Count = 3 },
            new CellStyleFormats(
                new CellFormat(),
                new CellFormat {
                    FontId = 1, FillId = 1, BorderId = 1, NumberFormatId = 164,
                    Alignment = AlignmentValue(),
                    Protection = new Protection { Locked = false, Hidden = true }
                }) { Count = 2 },
            new CellFormats(
                new CellFormat(),
                new CellFormat {
                    FormatId = 1, FontId = 2, FillId = 2, BorderId = 2, NumberFormatId = 165,
                    ApplyFont = false, ApplyFill = false, ApplyBorder = false, ApplyNumberFormat = false,
                    ApplyAlignment = false, Alignment = AlignmentValue("horizontal"),
                    ApplyProtection = false, Protection = new Protection { Locked = true, Hidden = false }
                },
                new CellFormat {
                    FontId = 1, FillId = 1, BorderId = 1, NumberFormatId = 164,
                    ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyNumberFormat = true,
                    ApplyAlignment = true, Alignment = AlignmentValue(),
                    ApplyProtection = true, Protection = new Protection { Locked = false, Hidden = true }
                }) { Count = 3 });
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetData(new Row(
            new Cell { CellReference = "A1", StyleIndex = 1, CellValue = new CellValue("1") },
            new Cell { CellReference = "B1", StyleIndex = 2, CellValue = new CellValue("1") }) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Styles" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }

    private static string AlignmentFixture(string? changedProperty = null)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-alignment-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var styles = wb.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(
                new CellFormat(),
                new CellFormat { ApplyAlignment = true, Alignment = AlignmentValue(changedProperty) }) { Count = 2 });
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetData(new Row(new Cell { CellReference = "A1", StyleIndex = 1, CellValue = new CellValue("1") }) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Alignment" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }

    private static Alignment AlignmentValue(string? changedProperty = null) => new()
    {
        Horizontal = changedProperty == "horizontal" ? HorizontalAlignmentValues.Left : HorizontalAlignmentValues.Center,
        Vertical = changedProperty == "vertical" ? VerticalAlignmentValues.Top : VerticalAlignmentValues.Center,
        TextRotation = changedProperty == "textRotation" ? 30U : 15U,
        WrapText = changedProperty != "wrapText",
        ShrinkToFit = changedProperty != "shrinkToFit",
        Indent = changedProperty == "indent" ? 3U : 2U,
        RelativeIndent = changedProperty == "relativeIndent" ? 2 : 1,
        JustifyLastLine = changedProperty != "justifyLastLine",
        ReadingOrder = changedProperty == "readingOrder" ? 2U : 1U,
        MergeCell = changedProperty == "mergeCell" ? "0" : "1"
    };

    private static string NumberFormatFixture(uint id, string code)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-number-format-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var styles = wb.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new NumberingFormats(new NumberingFormat { NumberFormatId = id, FormatCode = code }) { Count = 1 },
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = id, ApplyNumberFormat = true }) { Count = 2 });
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetData(new Row(new Cell { CellReference = "A1", StyleIndex = 1, CellValue = new CellValue("1") }) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Formats" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }

    private static string BuiltInNumberFormatFixture(uint id)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-builtin-number-format-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var styles = wb.AddNewPart<WorkbookStylesPart>();
        styles.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(new Fill()) { Count = 1 },
            new Borders(new Border()) { Count = 1 },
            new CellStyleFormats(new CellFormat()) { Count = 1 },
            new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = id, ApplyNumberFormat = true }) { Count = 2 });
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetData(new Row(new Cell { CellReference = "A1", StyleIndex = 1 }) { RowIndex = 1 }));
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Formats" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }

    private static string DefinedNamesFixture(
        bool reversed = false,
        string localName = "LocalInput",
        uint localSheetId = 1,
        string localText = "'Second'!$A$1",
        bool localHidden = true)
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-defined-names-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart();
        wb.Workbook = new Workbook();
        var first = wb.AddNewPart<WorksheetPart>();
        first.Worksheet = new Worksheet(new SheetData());
        var second = wb.AddNewPart<WorksheetPart>();
        second.Worksheet = new Worksheet(new SheetData());
        var definedNames = new[] {
            new DefinedName("0.25") { Name = "GlobalRate" },
            new DefinedName(localText) { Name = localName, LocalSheetId = localSheetId, Hidden = localHidden },
            new DefinedName("'First'!$A$1:$B$2") { Name = "_xlnm.Print_Area", LocalSheetId = 0 }
        };
        if (reversed) Array.Reverse(definedNames);
        wb.Workbook.Append(
            new Sheets(
                new Sheet { Id = wb.GetIdOfPart(first), SheetId = 1, Name = "First" },
                new Sheet { Id = wb.GetIdOfPart(second), SheetId = 2, Name = "Second" }),
            new DefinedNames(definedNames));
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
