using Dockit.Convert;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Convert.Tests;

public class ConvertCliTests
{
    [Fact]
    public void Xls_to_xlsx_conversion_preserves_sheet_and_values()
    {
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.xlsx");

        WorkbookConverter.ConvertXlsToXlsx(input, output);

        Assert.True(File.Exists(output));
        using var stream = File.OpenRead(output);
        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheetAt(0);
        Assert.Equal("Plan", sheet.SheetName);
        Assert.Equal("Condition", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("High temperature", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("2025-09-23", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal(BorderStyle.Thin, sheet.GetRow(1).GetCell(0).CellStyle.BorderBottom);
    }

    [Fact]
    public void Classify_open_error_marks_encrypted_workbooks_clearly()
    {
        var ex = WorkbookConverter.ClassifyOpenWorkbookError(
            "/tmp/protected.xls",
            new InvalidOperationException("Implement it based on poi 4.2 in the future"));

        Assert.Contains("Encrypted or password-protected XLS", ex.Message);
    }

    [Fact]
    public void Office_to_pdf_reports_clear_failure_when_soffice_is_missing()
    {
        var input = CreateDocxFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");

        var ex = Assert.Throws<InvalidOperationException>(
            () => OfficePdfConverter.ConvertToPdf(input, output, "docx", sofficePath: "/missing/soffice"));

        Assert.Contains("LibreOffice/soffice is required", ex.Message);
    }

    [Fact]
    public void Docx_to_pdf_conversion_creates_real_pdf_when_soffice_is_available()
    {
        var soffice = OfficePdfConverter.FindSofficeBinary();
        if (string.IsNullOrWhiteSpace(soffice))
        {
            return;
        }

        var input = CreateDocxFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");

        OfficePdfConverter.ConvertToPdf(input, output, "docx");

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 1_000);
        Assert.Equal("%PDF", File.ReadAllText(output)[..4]);
    }

    private static string CreateLegacyXlsFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"legacy-convert-{Guid.NewGuid():N}.xls");

        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Plan");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Condition");
        header.CreateCell(1).SetCellValue("Placement");

        var row1 = sheet.CreateRow(1);
        var borderedStyle = workbook.CreateCellStyle();
        borderedStyle.BorderBottom = BorderStyle.Thin;
        var styledCell = row1.CreateCell(0);
        styledCell.SetCellValue("High temperature");
        styledCell.CellStyle = borderedStyle;
        row1.CreateCell(1).SetCellValue("2025-09-23");

        using var output = File.Create(path);
        workbook.Write(output);
        return path;
    }

    private static string CreateDocxFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"office-convert-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(
            new Body(
                new Paragraph(
                    new Run(
                        new Text("Certificate of Analysis 260245 HSP1028")))));
        mainPart.Document.Save();
        return path;
    }
}
