using Dockit.Convert;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;

namespace Dockit.Convert.Tests;

public class ConvertCliTests
{
    [Fact]
    public void Xls_to_xlsx_conversion_preserves_sheet_and_values()
    {
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.xlsx");

        var result = WorkbookConverter.ConvertXlsToXlsx(input, output);

        Assert.True(File.Exists(output));
        Assert.Contains(result.Backend, new[] { "wps", "libreoffice", "npoi" });
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

        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "libreoffice");
        OfficePdfConversionResult result;
        try
        {
            result = OfficePdfConverter.ConvertToPdf(input, output, "docx");
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
        }

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 1_000);
        Assert.Equal("%PDF", File.ReadAllText(output)[..4]);
        Assert.Equal("libreoffice", result.Backend);
    }

    [Fact]
    public void Required_wps_writer_backend_fails_closed_when_runtime_is_unavailable()
    {
        if (WpsWriterPdfConverter.IsAvailable()) return;
        var input = CreateDocxFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "wps-writer");
        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() => OfficePdfConverter.ConvertToPdf(input, output, "docx"));
            Assert.Contains("WPS Writer PDF backend was required", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
        }
    }

    [Theory]
    [InlineData("getWpsApplication failed: 0x80000008")]
    [InlineData("Fatal IO error on X server :101")]
    [InlineData("X server has exited")]
    public void Wps_writer_recognizes_transient_rpc_startup_failures(string message)
    {
        Assert.True(WpsWriterPdfConverter.IsTransientStartupFailure(message));
    }

    [Fact]
    public void Wps_rpc_runs_in_a_fresh_dbus_session_around_isolated_xvfb()
    {
        var isolated = Path.Combine(Path.GetTempPath(), $"wps-working-{Guid.NewGuid():N}");
        Directory.CreateDirectory(isolated);
        var helper = Path.Combine(isolated, "helper.py");
        var input = Path.Combine(isolated, "input.docx");
        var output = Path.Combine(isolated, "output.pdf");

        var startInfo = WpsRpcSession.CreateProcessStartInfo(
            "/usr/bin/dbus-run-session",
            "/usr/bin/xvfb-run",
            "/venv/bin/python",
            helper,
            input,
            output,
            isolated);

        Assert.Equal("/usr/bin/dbus-run-session", startInfo.FileName);
        Assert.Equal(Path.GetFullPath(isolated), startInfo.WorkingDirectory);
        Assert.NotEqual(Directory.GetCurrentDirectory(), startInfo.WorkingDirectory);
        Assert.Equal(new[]
        {
            "--",
            "/usr/bin/xvfb-run",
            "-a",
            "/venv/bin/python",
            helper,
            Path.GetFullPath(input),
            Path.GetFullPath(output),
        }, startInfo.ArgumentList);
    }

    [Fact]
    public void Wps_writer_uses_the_supported_SaveAs2_pdf_api()
    {
        var helperScript = typeof(WpsWriterPdfConverter)
            .GetField("WpsHelperScript", BindingFlags.NonPublic | BindingFlags.Static)
            ?.GetRawConstantValue() as string;

        Assert.NotNull(helperScript);
        Assert.Contains("document.SaveAs2(output_path, FileFormat=wpsapi.wdFormatPDF)", helperScript);
        Assert.DoesNotContain("ExportAsFixedFormat", helperScript);
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
