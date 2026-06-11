using Dockit.Xlsx;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System.Text.Json;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class LegacyXlsTests
{
    [Fact]
    public void Inspect_reads_legacy_xls_workbooks()
    {
        var path = CreateLegacyXlsFixture();

        var report = Inspector.Inspect(path);

        var sheet = Assert.Single(report.Sheets);
        Assert.Equal("Plan", sheet.Name);
        Assert.Equal(3, sheet.RowCount);
        Assert.True(sheet.ColumnCount >= 3);
        var richCell = Assert.Single(sheet.TextCells!, cell => cell.Reference == "D2");
        Assert.Equal("QVQLVQSGAEVK", richCell.Text);
        Assert.Contains(richCell.RichTextRuns!, run => run.Text == "Q" && run.Color == "FF0000" && run.Underline == "single");
    }

    [Fact]
    public void ExportJson_reads_legacy_xls_workbooks()
    {
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"legacy-xls-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([input, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("\"sheet\": \"Plan\"", json, StringComparison.Ordinal);
        Assert.Contains("\"2025-09-23\"", json, StringComparison.Ordinal);

        using var parsed = JsonDocument.Parse(json);
        var cells = parsed.RootElement[0].GetProperty("cells").EnumerateArray().ToList();
        var richCell = cells.Single(cell => cell.GetProperty("reference").GetString() == "D2");
        var runs = richCell.GetProperty("richTextRuns").EnumerateArray().ToList();
        Assert.Contains(runs, run =>
            run.GetProperty("text").GetString() == "Q" &&
            run.GetProperty("color").GetString() == "FF0000" &&
            run.GetProperty("underline").GetString() == "single");
    }

    private static string CreateLegacyXlsFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"legacy-xls-{Guid.NewGuid():N}.xls");

        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Plan");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Condition");
        header.CreateCell(1).SetCellValue("Placement");
        header.CreateCell(2).SetCellValue("Sampling");

        var row1 = sheet.CreateRow(1);
        row1.CreateCell(0).SetCellValue("High temperature");
        row1.CreateCell(1).SetCellValue("2025-09-23");
        row1.CreateCell(2).SetCellValue("2025-10-23");
        row1.CreateCell(3).SetCellValue(CreateLegacyRichString(workbook));

        var row2 = sheet.CreateRow(2);
        row2.CreateCell(0).SetCellValue("Freeze-thaw");
        row2.CreateCell(1).SetCellValue("2025-09-23");
        row2.CreateCell(2).SetCellValue("2025-10-09");

        using var stream = File.Create(path);
        workbook.Write(stream);
        return path;
    }

    private static HSSFRichTextString CreateLegacyRichString(HSSFWorkbook workbook)
    {
        var redUnderlinedFont = workbook.CreateFont();
        redUnderlinedFont.Color = HSSFColor.Red.Index;
        redUnderlinedFont.Underline = FontUnderlineType.Single;

        var richText = new HSSFRichTextString("QVQLVQSGAEVK");
        richText.ApplyFont(2, 3, redUnderlinedFont);
        return richText;
    }

    [Fact]
    public void ExportJson_resolves_merged_cells()
    {
        var input = CreateMergedXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"merged-xls-export-{Guid.NewGuid():N}.json");

        Extractor.RunExportJson([input, output, "--resolve-merged-cells"]);

        var json = File.ReadAllText(output);
        Assert.Contains("\"MergedValue\"", json, StringComparison.Ordinal);

        using var parsed = System.Text.Json.JsonDocument.Parse(json);
        var sheet = parsed.RootElement[0];
        Assert.Equal("MergedValue", sheet.GetProperty("rows")[1][0].GetString());
        Assert.Equal("MergedValue", sheet.GetProperty("rows")[2][0].GetString());
        Assert.Equal("MergedValue", sheet.GetProperty("formattedRows")[1][0].GetString());
        Assert.Equal("MergedValue", sheet.GetProperty("formattedRows")[2][0].GetString());
        var mergedCell = sheet.GetProperty("cells").EnumerateArray().Single(cell => cell.GetProperty("reference").GetString() == "A2");
        Assert.Equal("MergedValue", mergedCell.GetProperty("value").GetString());
    }

    private static string CreateMergedXlsFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"merged-xls-{Guid.NewGuid():N}.xls");

        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        
        var row1 = sheet.CreateRow(0);
        row1.CreateCell(0).SetCellValue("Header");
        row1.CreateCell(1).SetCellValue("Val");

        var row2 = sheet.CreateRow(1);
        row2.CreateCell(0).SetCellValue("MergedValue");
        row2.CreateCell(1).SetCellValue("A");

        var row3 = sheet.CreateRow(2);
        // Column 0 is left empty (merged)
        row3.CreateCell(1).SetCellValue("B");

        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 2, 0, 0));

        using var stream = File.Create(path);
        workbook.Write(stream);
        return path;
    }
}
