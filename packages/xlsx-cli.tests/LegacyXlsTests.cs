using Dockit.Xlsx;
using NPOI.HSSF.UserModel;
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

        var row2 = sheet.CreateRow(2);
        row2.CreateCell(0).SetCellValue("Freeze-thaw");
        row2.CreateCell(1).SetCellValue("2025-09-23");
        row2.CreateCell(2).SetCellValue("2025-10-09");

        using var stream = File.Create(path);
        workbook.Write(stream);
        return path;
    }
}
