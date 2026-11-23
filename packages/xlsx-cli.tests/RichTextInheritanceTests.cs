using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;
using CellValue = DocumentFormat.OpenXml.Spreadsheet.CellValue;
using Font = DocumentFormat.OpenXml.Spreadsheet.Font;
using NpoiIndexedColors = NPOI.SS.UserModel.IndexedColors;

namespace Dockit.Xlsx.Tests;

public sealed class RichTextInheritanceTests
{
    [Fact]
    public void UnstyledFirstRun_InheritsCellXfColorAndUnderline_WhileExplicitRunPreservesOwnFormatting()
    {
        var path = CreateRichTextWorkbook(
            cellFont: new Font(
                new FontName { Val = "Arial" },
                new Color { Rgb = "FFFF0000" },
                new Underline()),
            runs:
            [
                // Position-1 run omitting <rPr> completely (exact C126 pattern)
                new Run(new Text("N")),
                // Following run with explicit <rPr> specifying black color and no underline
                new Run(
                    new RunProperties(
                        new RunFont { Val = "Arial" },
                        new Color { Rgb = "FF000000" }),
                    new Text("QVSLTCLVK"))
            ]);
        try
        {
            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(2, exportedRuns.Count);

            Assert.Equal("N", exportedRuns[0].Text);
            Assert.Equal("FFFF0000", exportedRuns[0].Color);
            Assert.Equal("single", exportedRuns[0].Underline);

            Assert.Equal("QVSLTCLVK", exportedRuns[1].Text);
            Assert.Equal("FF000000", exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);

            // Verify RangeReader and Inspector match Extractor
            var range = RangeReader.Read(path, "Sheet1", "C126:C126", 0, 10);
            var rangeRuns = range.Cells.Single().RichTextRuns;
            Assert.NotNull(rangeRuns);
            Assert.Equal("FFFF0000", rangeRuns![0].Color);
            Assert.Equal("single", rangeRuns[0].Underline);
            Assert.Equal("FF000000", rangeRuns[1].Color);
            Assert.Null(rangeRuns[1].Underline);

            var inspected = Inspector.Inspect(path);
            var inspectedCell = inspected.Sheets.Single().Cells!.Single(c => c.Reference == "C126");
            Assert.NotNull(inspectedCell.RichTextRuns);
            Assert.Equal("FFFF0000", inspectedCell.RichTextRuns![0].Color);
            Assert.Equal("single", inspectedCell.RichTextRuns[0].Underline);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Fact]
    public void LibreOfficeConvertedC126Pattern_InheritsRedOnUnstyledFirstRun_AndDoesNotBleedIntoStyledBlackRun()
    {
        // Reproduces exact C126 XML from LibreOffice convertXlsToXlsx:
        // cell XF has red font; run 0 has NO <rPr>; run 1 has <rPr><sz val="10"/><rFont val="Arial"/></rPr> (no <color>).
        var path = CreateRichTextWorkbook(
            cellFont: new Font(
                new FontName { Val = "Arial" },
                new Color { Rgb = "FFFF0000" }),
            runs:
            [
                new Run(new Text("N")),
                new Run(
                    new RunProperties(
                        new FontSize { Val = 10 },
                        new RunFont { Val = "Arial" }),
                    new Text("QVSLTCLVK"))
            ]);
        try
        {
            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(2, exportedRuns.Count);
            Assert.Equal("N", exportedRuns[0].Text);
            Assert.Equal("FFFF0000", exportedRuns[0].Color);
            Assert.Null(exportedRuns[0].Underline);

            Assert.Equal("QVSLTCLVK", exportedRuns[1].Text);
            Assert.Null(exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Fact]
    public void ExplicitRunColorAndUnderlineNoneOverride_DoNotInheritCellXf()
    {
        var path = CreateRichTextWorkbook(
            cellFont: new Font(
                new FontName { Val = "Arial" },
                new Color { Rgb = "FFFF0000" },
                new Underline { Val = UnderlineValues.Double },
                new Bold(),
                new Italic()),
            runs:
            [
                new Run(new Text("N")),
                new Run(
                    new RunProperties(
                        new RunFont { Val = "Calibri" },
                        new Color { Rgb = "FF0000FF" },
                        new Underline { Val = UnderlineValues.None }),
                    new Text("QVSLTCLVK"))
            ]);
        try
        {
            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(2, exportedRuns.Count);

            Assert.Equal("N", exportedRuns[0].Text);
            Assert.Equal("Arial", exportedRuns[0].FontName);
            Assert.Equal("FFFF0000", exportedRuns[0].Color);
            Assert.Equal("double", exportedRuns[0].Underline);
            Assert.True(exportedRuns[0].Bold);
            Assert.True(exportedRuns[0].Italic);

            Assert.Equal("QVSLTCLVK", exportedRuns[1].Text);
            Assert.Equal("Calibri", exportedRuns[1].FontName);
            Assert.Equal("FF0000FF", exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
            Assert.False(exportedRuns[1].Bold);
            Assert.False(exportedRuns[1].Italic);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Fact]
    public void UnstyledRun_InheritsFontFromCellStyleFormatWhenApplyFontIsFalse()
    {
        var path = CreateRichTextWorkbook(
            cellFont: new Font(
                new Color { Indexed = 10U },
                new Underline { Val = UnderlineValues.Single }),
            runs:
            [
                new Run(new Text("Q")),
                new Run(
                    new RunProperties(new Color { Theme = 1U }),
                    new Text("VQLVESGG"))
            ],
            inheritFromCellStyleFormat: true);
        try
        {
            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(2, exportedRuns.Count);
            Assert.Equal("Q", exportedRuns[0].Text);
            Assert.Equal("indexed:10", exportedRuns[0].Color);
            Assert.Equal("single", exportedRuns[0].Underline);

            Assert.Equal("VQLVESGG", exportedRuns[1].Text);
            Assert.Equal("theme:1", exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Fact]
    public void InlineStringUnstyledRun_InheritsCellXfColorAndUnderline()
    {
        var path = CreateRichTextWorkbook(
            cellFont: new Font(
                new Color { Rgb = "FFFF0000" },
                new Underline()),
            runs:
            [
                new Run(new Text("N")),
                new Run(
                    new RunProperties(new Color { Rgb = "FF000000" }),
                    new Text("QV")),
                new Run(new Text("S"))
            ],
            inlineString: true);
        try
        {
            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(3, exportedRuns.Count);
            Assert.Equal("FFFF0000", exportedRuns[0].Color);
            Assert.Equal("single", exportedRuns[0].Underline);
            Assert.Equal("FF000000", exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
            Assert.Equal("FFFF0000", exportedRuns[2].Color);
            Assert.Equal("single", exportedRuns[2].Underline);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Fact]
    public void DefaultUncoloredCellXf_LeavesUnstyledRunColorAndUnderlineNull()
    {
        var path = CreateRichTextWorkbook(
            cellFont: new Font(new FontName { Val = "Arial" }),
            runs:
            [
                new Run(new Text("AB")),
                new Run(
                    new RunProperties(new Color { Rgb = "FFFF0000" }),
                    new Text("C")),
                new Run(new Text("DE"))
            ]);
        try
        {
            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(3, exportedRuns.Count);
            Assert.Null(exportedRuns[0].Color);
            Assert.Null(exportedRuns[0].Underline);
            Assert.Equal("FFFF0000", exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
            Assert.Null(exportedRuns[2].Color);
            Assert.Null(exportedRuns[2].Underline);
        }
        finally
        {
            Cleanup(path);
        }
    }

    [Fact]
    public void LegacyXlsUnstyledFirstRun_InheritsCellXfColorAndUnderline()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"xlsx-cli-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "c126.xls");
        try
        {
            var wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet("Sheet1");
            var row = sheet.CreateRow(125);
            var cell = row.CreateCell(2); // C126

            var redFont = wb.CreateFont();
            redFont.Color = NpoiIndexedColors.Red.Index;
            redFont.Underline = FontUnderlineType.Single;
            redFont.FontName = "Arial";

            var blackFont = wb.CreateFont();
            blackFont.Color = NpoiIndexedColors.Black.Index;
            blackFont.FontName = "Arial";

            var style = wb.CreateCellStyle();
            style.SetFont(redFont);
            cell.CellStyle = style;

            var richText = new HSSFRichTextString("NQVSLTCLVK");
            richText.ApplyFont(1, 10, blackFont);
            cell.SetCellValue(richText);

            using (var fs = File.Create(path))
            {
                wb.Write(fs);
            }

            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(2, exportedRuns.Count);
            Assert.Equal("N", exportedRuns[0].Text);
            Assert.Equal("FF0000", exportedRuns[0].Color);
            Assert.Equal("single", exportedRuns[0].Underline);

            Assert.Equal("QVSLTCLVK", exportedRuns[1].Text);
            Assert.Equal("000000", exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void LegacyXlsDefaultFontRun_DoesNotInheritCellXfRedColor()
    {
        var dir = Path.Combine(Path.GetTempPath(), $"xlsx-cli-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "c126-default-font.xls");
        try
        {
            var wb = new HSSFWorkbook();
            var sheet = wb.CreateSheet("Sheet1");
            var row = sheet.CreateRow(125);
            var cell = row.CreateCell(2); // C126

            var redFont = wb.CreateFont();
            redFont.Color = NpoiIndexedColors.Red.Index;
            redFont.FontName = "Arial";

            var defaultFont = wb.GetFontAt(0); // fontIndex == 0 (HSSFRichTextString.NO_FONT)

            var style = wb.CreateCellStyle();
            style.SetFont(redFont);
            cell.CellStyle = style;

            var richText = new HSSFRichTextString("NQVSLTCLVK");
            richText.ApplyFont(1, 10, defaultFont);
            cell.SetCellValue(richText);

            using (var fs = File.Create(path))
            {
                wb.Write(fs);
            }

            var exportedRuns = ExportCellRichTextRuns(path, "C126");
            Assert.Equal(2, exportedRuns.Count);
            Assert.Equal("N", exportedRuns[0].Text);
            Assert.Equal("FF0000", exportedRuns[0].Color);
            Assert.Null(exportedRuns[0].Underline);

            Assert.Equal("QVSLTCLVK", exportedRuns[1].Text);
            Assert.Null(exportedRuns[1].Color);
            Assert.Null(exportedRuns[1].Underline);
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }
    }

    private static IReadOnlyList<RichTextRunReport> ExportCellRichTextRuns(string path, string reference)
    {
        var result = Extractor.Export(path);
        var json = JsonSerializer.Serialize(result, Json.Options);
        using var doc = JsonDocument.Parse(json);
        var cells = doc.RootElement[0].GetProperty("cells");
        foreach (var cell in cells.EnumerateArray())
        {
            if (string.Equals(cell.GetProperty("reference").GetString(), reference, StringComparison.OrdinalIgnoreCase))
            {
                var runs = cell.GetProperty("richTextRuns");
                Assert.Equal(JsonValueKind.Array, runs.ValueKind);
                return JsonSerializer.Deserialize<List<RichTextRunReport>>(runs.GetRawText(), Json.Options)!;
            }
        }
        throw new InvalidOperationException($"Cell {reference} not found.");
    }

    private static string CreateRichTextWorkbook(
        Font cellFont,
        IEnumerable<Run> runs,
        bool inlineString = false,
        bool inheritFromCellStyleFormat = false)
    {
        var dir = Path.Combine(Path.GetTempPath(), $"xlsx-cli-tests-{Guid.NewGuid():N}");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "richtext.xlsx");

        using (var spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            var cellStyleFormats = new CellStyleFormats(
                new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U },
                new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U });
            var cellFormats = new CellFormats(
                new CellFormat { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U },
                inheritFromCellStyleFormat
                    ? new CellFormat { FormatId = 1U, ApplyFont = false }
                    : new CellFormat { NumberFormatId = 0U, FontId = 1U, FillId = 0U, BorderId = 0U, ApplyFont = true });

            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font(), cellFont),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })),
                new Borders(new Border()),
                cellStyleFormats,
                cellFormats);
            stylesPart.Stylesheet.Save();

            var cell = new Cell
            {
                CellReference = "C126",
                StyleIndex = 1U
            };

            if (inlineString)
            {
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(runs);
            }
            else
            {
                var sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sstPart.SharedStringTable = new SharedStringTable(new SharedStringItem(runs));
                sstPart.SharedStringTable.Save();
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue("0");
            }

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(new Row(cell) { RowIndex = 126U }));
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
