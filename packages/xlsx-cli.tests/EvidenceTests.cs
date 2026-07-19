using Dockit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        Assert.Equal("center", cell.GetProperty("style").GetProperty("horizontalAlignment").GetString(), ignoreCase: true);
        Assert.Equal("A2*2", cell.GetProperty("formula").GetProperty("text").GetString());
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

    private static string Fixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-evidence-{Guid.NewGuid():N}.xlsx");
        using var doc = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var wb = doc.AddWorkbookPart(); wb.Workbook = new Workbook(new WorkbookProperties { Date1904 = true });
        var styles = wb.AddNewPart<WorkbookStylesPart>(); styles.Stylesheet = new Stylesheet(new Fonts(new Font()), new Fills(new Fill()), new Borders(new Border()), new CellStyleFormats(new CellFormat()), new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 2, Alignment = new Alignment { Horizontal = HorizontalAlignmentValues.Center } }));
        var ws = wb.AddNewPart<WorksheetPart>();
        ws.Worksheet = new Worksheet(new SheetDimension { Reference = "A1:B2" }, new SheetViews(new SheetView { WorkbookViewId = 0, View = SheetViewValues.PageBreakPreview, ShowGridLines = false }), new SheetData(new Row(new Cell { CellReference = "A2", CellValue = new CellValue("2") }, new Cell { CellReference = "B2", StyleIndex = 1, CellFormula = new CellFormula("A2*2"), CellValue = new CellValue("4") }) { RowIndex = 2 }), new MergeCells(new MergeCell { Reference = "A1:B1" }), new PageMargins { Left = .7, Right = .7, Top = .75, Bottom = .75, Header = .3, Footer = .3 }, new PageSetup { Orientation = OrientationValues.Landscape });
        wb.Workbook.AppendChild(new Sheets()).Append(new Sheet { Id = wb.GetIdOfPart(ws), SheetId = 1, Name = "Report" });
        wb.Workbook.Save(); styles.Stylesheet.Save(); ws.Worksheet.Save();
        return path;
    }
}
