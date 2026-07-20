using Dockit.Convert;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;
using System.Security.Cryptography;

namespace Dockit.Convert.Tests;

public class ConvertCliTests
{
    [Fact]
    public void Wps_xlsx_recalculation_is_writable_full_calculation_and_fresh_save()
    {
        var script = WpsSpreadsheetRecalculator.WpsHelperScript;
        Assert.Contains("ReadOnly=False", script);
        Assert.Contains("hr = app.CalculateFull()", script);
        Assert.Contains("Application.CalculateFull failed", script);
        Assert.Contains("book.SaveAs(output_path", script);
        Assert.Throws<InvalidOperationException>(() => WpsSpreadsheetRecalculator.Recalculate("/missing/input.xlsx", "/tmp/output.xlsx"));
        var input = CreateXlsxFixture();
        Assert.Throws<InvalidOperationException>(() => WpsSpreadsheetRecalculator.Recalculate(input, input));
    }

    [Fact]
    public void Lima_recalculation_transport_invokes_the_versioned_remote_command()
    {
        var start = LimaWpsWriterPdfConverter.CreateSpreadsheetConversionStartInfo("/usr/bin/limactl", "wps", "/shared/input.xlsx", "/shared/output.xlsx", "recalculate-xlsx");
        Assert.Equal("/usr/bin/limactl", start.FileName);
        Assert.Contains("recalculate-xlsx '/shared/input.xlsx' '/shared/output.xlsx'", start.ArgumentList.Last());
    }

    [Fact]
    public void Lima_recalculation_evidence_must_attest_actual_staged_bytes()
    {
        var input = Path.GetTempFileName();
        var output = Path.GetTempFileName();
        File.WriteAllText(input, "input bytes");
        File.WriteAllText(output, "output bytes");
        var inputHash = System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))).ToLowerInvariant();
        var outputHash = System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(output))).ToLowerInvariant();
        var valid = $$"""{"status":"ok","backend":"wps-spreadsheet","fallback_reason":null,"source_format":"xlsx","target_format":"xlsx","input_sha256":"{{inputHash}}","output_sha256":"{{outputHash}}"}""";

        LimaWpsWriterPdfConverter.ValidateSpreadsheetEvidence(valid, "recalculate-xlsx", input, output);

        var unattested = valid.Replace(outputHash, new string('0', 64), StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => LimaWpsWriterPdfConverter.ValidateSpreadsheetEvidence(unattested, "recalculate-xlsx", input, output));
    }

    [Fact]
    public void Xls_to_xlsx_conversion_preserves_sheet_and_values()
    {
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.xlsx");

        var result = WorkbookConverter.ConvertXlsToXlsx(input, output);

        Assert.True(File.Exists(output));
        Assert.Contains(result.Backend, new[] { "wps-spreadsheet", "libreoffice", "npoi" });
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
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");

        try
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "libreoffice");
            var ex = Assert.Throws<InvalidOperationException>(
                () => OfficePdfConverter.ConvertToPdf(input, output, "docx", sofficePath: "/missing/soffice"));

            Assert.Contains("LibreOffice/soffice is required", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
        }
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
        var originalLimaInstance = Environment.GetEnvironmentVariable("TIWATER_WPS_WRITER_LIMA_INSTANCE");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "wps-writer");
        Environment.SetEnvironmentVariable("TIWATER_WPS_WRITER_LIMA_INSTANCE", null);
        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() => OfficePdfConverter.ConvertToPdf(input, output, "docx"));
            Assert.Contains("WPS Writer PDF backend was required", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
            Environment.SetEnvironmentVariable("TIWATER_WPS_WRITER_LIMA_INSTANCE", originalLimaInstance);
        }
    }

    [Fact]
    public void Required_wps_spreadsheet_backend_fails_closed_when_runtime_is_unavailable()
    {
        if (WpsSpreadsheetPdfConverter.IsAvailable() || LimaWpsWriterPdfConverter.IsAvailable()) return;
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "wps-spreadsheet");
        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() => OfficePdfConverter.ConvertToPdf(input, output, "xls"));
            Assert.Contains("WPS Spreadsheets PDF backend was required", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
        }
    }

    [Fact]
    public void Wps_spreadsheet_uses_the_supported_ExportAsFixedFormat_pdf_api()
    {
        var helperScript = typeof(WpsSpreadsheetPdfConverter)
            .GetField("WpsHelperScript", BindingFlags.NonPublic | BindingFlags.Static)
            ?.GetRawConstantValue() as string;

        Assert.NotNull(helperScript);
        Assert.Contains("from pywpsrpc.rpcetapi import createEtRpcInstance, etapi", helperScript);
        Assert.Contains("book.ExportAsFixedFormat(", helperScript);
        Assert.Contains("etapi.XlFixedFormatType.xlTypePDF", helperScript);
        Assert.Contains("IgnorePrintAreas=False", helperScript);
    }

    [Fact]
    public void Wps_spreadsheet_pdf_conversion_creates_a_real_pdf_when_runtime_is_available()
    {
        if (!WpsSpreadsheetPdfConverter.IsAvailable()) return;
        var input = CreateXlsxFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "wps-spreadsheet");
        try
        {
            var result = OfficePdfConverter.ConvertToPdf(input, output, "xlsx");
            Assert.Equal("wps-spreadsheet", result.Backend);
            Assert.True(File.Exists(output));
            Assert.True(new FileInfo(output).Length > 1_000);
            Assert.Equal("%PDF", File.ReadAllText(output)[..4]);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
        }
    }

    [Theory]
    [InlineData("getWpsApplication failed: 0x80000008")]
    [InlineData("Fatal IO error on X server :101")]
    public void Wps_writer_recognizes_transient_rpc_startup_failures(string message)
    {
        Assert.True(WpsWriterPdfConverter.IsTransientStartupFailure(message));
    }

    [Fact]
    public void Wps_writer_runs_in_an_isolated_working_directory()
    {
        var isolated = Path.Combine(Path.GetTempPath(), $"wps-working-{Guid.NewGuid():N}");
        Directory.CreateDirectory(isolated);

        var startInfo = WpsWriterPdfConverter.CreateProcessStartInfo("xvfb-run", isolated);

        Assert.Equal(Path.GetFullPath(isolated), startInfo.WorkingDirectory);
        Assert.NotEqual(Directory.GetCurrentDirectory(), startInfo.WorkingDirectory);
        Assert.Equal(Path.Combine(Path.GetFullPath(isolated), "cache"), startInfo.Environment["XDG_CACHE_HOME"]);
        Assert.Equal(Path.Combine(Path.GetFullPath(isolated), "runtime"), startInfo.Environment["XDG_RUNTIME_DIR"]);
    }

    [Fact]
    public void Wps_writer_starts_an_isolated_dbus_session()
    {
        var arguments = WpsWriterPdfConverter.CreateHelperArguments(
            "dbus-run-session",
            "/tmp/wpsrpc-python",
            "/tmp/writer_to_pdf_wps.py",
            "/tmp/input.docx",
            "/tmp/output.pdf");

        Assert.Equal(new[]
        {
            "-a",
            "dbus-run-session",
            "--",
            "/tmp/wpsrpc-python",
            "/tmp/writer_to_pdf_wps.py",
            "/tmp/input.docx",
            "/tmp/output.pdf",
        }, arguments);
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

    private static string CreateXlsxFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-convert-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Results");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Batch");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("260245");
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
