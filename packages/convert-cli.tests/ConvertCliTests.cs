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
using System.Diagnostics;

namespace Dockit.Convert.Tests;

public class ConvertCliTests
{
    [Fact]
    public async Task Wps_office_lease_serializes_mixed_backends_across_processes()
    {
        var root = Path.Combine(Path.GetTempPath(), $"wps-lease-{Guid.NewGuid():N}");
        var lockPath = Path.Combine(root, "office.lock");
        Directory.CreateDirectory(root);
        var eventLog = Path.Combine(root, "events.log");
        try
        {
            var processes = new[] { "writer", "spreadsheet", "presentation", "lima" }
                .Select(route => StartLeaseProbe(route, lockPath, eventLog, 150, 5_000))
                .ToArray();
            await Task.WhenAll(processes.Select(WaitForExit));
            Assert.All(processes, process => Assert.Equal(0, process.ExitCode));
            var active = 0;
            var maximumActive = 0;
            foreach (var line in File.ReadAllLines(eventLog))
            {
                active += line[0] == '+' ? 1 : -1;
                maximumActive = Math.Max(maximumActive, active);
                Assert.InRange(active, 0, 1);
            }
            Assert.Equal(0, active);
            Assert.Equal(1, maximumActive);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task Wps_spreadsheet_lease_is_released_when_the_holding_process_is_killed()
    {
        var root = Path.Combine(Path.GetTempPath(), $"wps-lease-{Guid.NewGuid():N}");
        var lockPath = Path.Combine(root, "spreadsheet.lock");
        Directory.CreateDirectory(root);
        var eventLog = Path.Combine(root, "events.log");
        try
        {
            using var killed = StartLeaseProbe("writer", lockPath, eventLog, 60_000, 5_000);
            await WaitForEvent(eventLog, '+', TimeSpan.FromSeconds(5));
            killed.Kill(entireProcessTree: true);
            await WaitForExit(killed);
            using var successor = StartLeaseProbe("presentation", lockPath, eventLog, 0, 2_000);
            await WaitForExit(successor);
            Assert.Equal(0, successor.ExitCode);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task Wps_spreadsheet_lease_reports_a_bounded_busy_timeout_across_processes()
    {
        var root = Path.Combine(Path.GetTempPath(), $"wps-lease-{Guid.NewGuid():N}");
        var lockPath = Path.Combine(root, "spreadsheet.lock");
        var eventLog = Path.Combine(root, "events.log");
        Directory.CreateDirectory(root);
        try
        {
            using var holder = StartLeaseProbe("spreadsheet", lockPath, eventLog, 60_000, 5_000);
            await WaitForEvent(eventLog, '+', TimeSpan.FromSeconds(5));
            using var blocked = StartLeaseProbe("writer", lockPath, eventLog, 0, 150);
            await WaitForExit(blocked);
            Assert.Equal(23, blocked.ExitCode);
            Assert.Contains(File.ReadAllLines(eventLog), line => line.StartsWith('!'));
            holder.Kill(entireProcessTree: true);
            await WaitForExit(holder);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Local_and_lima_office_routes_share_one_host_lease()
    {
        var root = Path.Combine(Path.GetTempPath(), $"wps-lease-{Guid.NewGuid():N}");
        var lockPath = Path.Combine(root, "spreadsheet.lock");
        Directory.CreateDirectory(root);
        try
        {
            using (EtPdfConverter.AcquireRuntimeLease(TimeSpan.FromSeconds(1), lockPath))
            {
                Assert.Throws<TimeoutException>(() =>
                    WpsPdfConverter.AcquireRuntimeLease(TimeSpan.FromMilliseconds(150), lockPath));
                Assert.Throws<TimeoutException>(() =>
                    WppPdfConverter.AcquireRuntimeLease(TimeSpan.FromMilliseconds(150), lockPath));
                Assert.Throws<TimeoutException>(() =>
                    LimaWpsPdfConverter.AcquireOfficeHostLease(TimeSpan.FromMilliseconds(150), lockPath));
            }

            using var limaLease = LimaWpsPdfConverter.AcquireOfficeHostLease(TimeSpan.FromSeconds(1), lockPath);
            Assert.Throws<TimeoutException>(() =>
                EtPdfConverter.AcquireRuntimeLease(TimeSpan.FromMilliseconds(150), lockPath));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Wps_xlsx_recalculation_is_writable_full_calculation_and_fresh_save()
    {
        var script = EtRecalculator.EtHelperScript;
        Assert.Contains("ReadOnly=False", script);
        Assert.Contains("hr = app.CalculateFull()", script);
        Assert.Contains("Application.CalculateFull failed", script);
        Assert.Contains("book.SaveAs(output_path", script);
        Assert.Throws<InvalidOperationException>(() => EtRecalculator.Recalculate("/missing/input.xlsx", "/tmp/output.xlsx"));
        var input = CreateXlsxFixture();
        Assert.Throws<InvalidOperationException>(() => EtRecalculator.Recalculate(input, input));
    }

    [Fact]
    public void Wps_xls_conversion_suppresses_interactive_save_as_prompts()
    {
        Assert.Contains("app.DisplayAlerts = False", EtWorkbookConverter.EtHelperScript);
        Assert.Contains("book.SaveAs(output_path", EtWorkbookConverter.EtHelperScript);
    }

    [Fact]
    public void Wps_xlsx_recalculation_does_not_wait_for_inherited_diagnostic_pipes()
    {
        var inheritedPipe = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        var started = Stopwatch.StartNew();
        var details = WpsRpcSession.CollectDiagnosticOutput(
            inheritedPipe.Task,
            Task.FromResult("helper diagnostic"),
            TimeSpan.FromMilliseconds(20));

        Assert.Equal("helper diagnostic", details);
        Assert.InRange(started.ElapsedMilliseconds, 0, 500);
    }

    [Fact]
    public void Wps_rpc_completion_marker_ends_a_lingering_office_process_after_valid_output()
    {
        if (OperatingSystem.IsWindows()) return;
        var directory = Path.Combine(Path.GetTempPath(), $"wps-completion-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var output = Path.Combine(directory, "output.pdf");
        var marker = Path.Combine(directory, "output-complete");
        var startInfo = new ProcessStartInfo
        {
            FileName = "/bin/sh",
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        foreach (var argument in new[]
        {
            "-c", "printf '%s' '%PDF-current-output' > \"$1\"; printf '%s' complete > \"$2\"; sleep 30", "sh", output, marker,
        }) startInfo.ArgumentList.Add(argument);
        using var process = Process.Start(startInfo)!;
        var started = Stopwatch.StartNew();

        var completedOutput = WpsRpcSession.WaitForCompletedOutputOrExit(
            process,
            marker,
            () => File.Exists(output) && new FileInfo(output).Length > 4,
            TimeSpan.FromSeconds(10),
            "synthetic helper timeout",
            TimeSpan.FromMilliseconds(20));

        Assert.True(completedOutput);
        Assert.True(process.HasExited);
        Assert.InRange(started.ElapsedMilliseconds, 0, 3000);
    }

    [Fact]
    public void Lima_recalculation_transport_invokes_the_versioned_remote_command()
    {
        var start = LimaWpsPdfConverter.CreateSpreadsheetConversionStartInfo("/usr/bin/limactl", "wps", "/shared/input.xlsx", "/shared/output.xlsx", "recalculate-xlsx");
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
        var valid = $$"""{"status":"ok","backend":"et","fallback_reason":null,"source_format":"xlsx","target_format":"xlsx","input_sha256":"{{inputHash}}","output_sha256":"{{outputHash}}"}""";

        LimaWpsPdfConverter.ValidateSpreadsheetEvidence(valid, "recalculate-xlsx", input, output);

        var unattested = valid.Replace(outputHash, new string('0', 64), StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => LimaWpsPdfConverter.ValidateSpreadsheetEvidence(unattested, "recalculate-xlsx", input, output));
    }

    [Fact]
    public void Xls_to_xlsx_conversion_preserves_sheet_and_values()
    {
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.xlsx");

        var result = WorkbookConverter.ConvertXlsToXlsx(input, output);

        Assert.True(File.Exists(output));
        Assert.Contains(result.Backend, new[] { "et", "libreoffice", "npoi" });
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
        if (WpsPdfConverter.IsAvailable()) return;
        var input = CreateDocxFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        var originalLimaInstance = Environment.GetEnvironmentVariable("TIWATER_WPS_OFFICE_LIMA_INSTANCE");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "wps");
        Environment.SetEnvironmentVariable("TIWATER_WPS_OFFICE_LIMA_INSTANCE", null);
        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() => OfficePdfConverter.ConvertToPdf(input, output, "docx"));
            Assert.Contains("WPS PDF backend was required", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
            Environment.SetEnvironmentVariable("TIWATER_WPS_OFFICE_LIMA_INSTANCE", originalLimaInstance);
        }
    }

    [Fact]
    public void Required_wps_spreadsheet_backend_fails_closed_when_runtime_is_unavailable()
    {
        if (EtPdfConverter.IsAvailable() || LimaWpsPdfConverter.IsAvailable()) return;
        var input = CreateLegacyXlsFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "et");
        try
        {
            var ex = Assert.Throws<InvalidOperationException>(() => OfficePdfConverter.ConvertToPdf(input, output, "xls"));
            Assert.Contains("ET PDF backend was required", ex.Message);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", originalBackend);
        }
    }

    [Fact]
    public void Wps_spreadsheet_uses_the_supported_ExportAsFixedFormat_pdf_api()
    {
        var helperScript = typeof(EtPdfConverter)
            .GetField("EtHelperScript", BindingFlags.NonPublic | BindingFlags.Static)
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
        if (!EtPdfConverter.IsAvailable()) return;
        var input = CreateXlsxFixture();
        var output = Path.Combine(Path.GetTempPath(), $"converted-{Guid.NewGuid():N}.pdf");
        var originalBackend = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", "et");
        try
        {
            var result = OfficePdfConverter.ConvertToPdf(input, output, "xlsx");
            Assert.Equal("et", result.Backend);
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
        Assert.True(WpsPdfConverter.IsTransientStartupFailure(message));
    }

    [Fact]
    public void Wps_writer_runs_in_an_isolated_working_directory()
    {
        var isolated = Path.Combine(Path.GetTempPath(), $"wps-working-{Guid.NewGuid():N}");
        Directory.CreateDirectory(isolated);

        var startInfo = WpsPdfConverter.CreateProcessStartInfo("xvfb-run", isolated);

        Assert.Equal(Path.GetFullPath(isolated), startInfo.WorkingDirectory);
        Assert.NotEqual(Directory.GetCurrentDirectory(), startInfo.WorkingDirectory);
        Assert.Equal(Path.Combine(Path.GetFullPath(isolated), "cache"), startInfo.Environment["XDG_CACHE_HOME"]);
        Assert.Equal(Path.Combine(Path.GetFullPath(isolated), "runtime"), startInfo.Environment["XDG_RUNTIME_DIR"]);
    }

    [Fact]
    public void Wps_writer_starts_an_isolated_dbus_session()
    {
        var arguments = WpsPdfConverter.CreateHelperArguments(
            "dbus-run-session",
            "/tmp/wpsrpc-python",
            "/tmp/writer_to_pdf_wps.py",
            "/tmp/input.docx",
            "/tmp/output.pdf",
            "/tmp/output-complete");

        Assert.Equal(new[]
        {
            "-a",
            "dbus-run-session",
            "--",
            "/tmp/wpsrpc-python",
            "/tmp/writer_to_pdf_wps.py",
            "/tmp/input.docx",
            "/tmp/output.pdf",
            "/tmp/output-complete",
        }, arguments);
    }

    [Fact]
    public void Wps_writer_uses_the_supported_SaveAs2_pdf_api()
    {
        var helperScript = typeof(WpsPdfConverter)
            .GetField("EtHelperScript", BindingFlags.NonPublic | BindingFlags.Static)
            ?.GetRawConstantValue() as string;

        Assert.NotNull(helperScript);
        Assert.Contains("document.SaveAs2(output_path, FileFormat=wpsapi.wdFormatPDF)", helperScript);
        Assert.Contains("with open(completion_marker, \"x\"", helperScript);
        Assert.True(helperScript.IndexOf("document.SaveAs2", StringComparison.Ordinal)
            < helperScript.IndexOf("with open(completion_marker", StringComparison.Ordinal));
        Assert.True(helperScript.IndexOf("with open(completion_marker", StringComparison.Ordinal)
            < helperScript.IndexOf("document.Close(False)", StringComparison.Ordinal));
        Assert.DoesNotContain("ExportAsFixedFormat", helperScript);
    }

    [Fact]
    public void Wps_writer_refreshes_layout_dependent_indexes_and_saves_a_distinct_docx()
    {
        var script = WpsPdfConverter.RefreshFieldsHelperScript;
        Assert.Contains("document.get_TablesOfContents()", script);
        Assert.Contains("toc.Update()", script);
        Assert.Contains("document.get_TablesOfFigures()", script);
        Assert.Contains("figure.Update()", script);
        Assert.Contains("document.Repaginate()", script);
        Assert.Contains("toc.UpdatePageNumbers()", script);
        Assert.Contains("figure.UpdatePageNumbers()", script);
        Assert.Contains("document.SaveAs2(output_path, FileFormat=wpsapi.wdFormatXMLDocument)", script);
        Assert.DoesNotContain("document.get_Fields()", script);
    }

    [Fact]
    public void Document_field_refresh_rejects_invalid_identity_before_runtime_selection()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"missing-{Guid.NewGuid():N}.docx");
        Assert.Contains("Input file not found", Assert.Throws<InvalidOperationException>(
            () => DocumentFieldRefresher.RefreshDocxFields(missing, missing + ".out.docx")).Message);

        var input = CreateDocxFixture();
        Assert.Contains("distinct input and output", Assert.Throws<InvalidOperationException>(
            () => DocumentFieldRefresher.RefreshDocxFields(input, input)).Message);
        Assert.Contains("output must be a DOCX", Assert.Throws<InvalidOperationException>(
            () => DocumentFieldRefresher.RefreshDocxFields(input, input + ".pdf")).Message);
    }

    private static Process StartLeaseProbe(string route, string lockPath, string eventLog, int holdMilliseconds, int timeoutMilliseconds)
    {
        var configuration = new DirectoryInfo(AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar)).Parent?.Name
            ?? throw new InvalidOperationException("Test configuration directory is unavailable.");
        var probe = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "../../../../convert-cli.lease-probe/bin",
            configuration,
            "net9.0",
            "WpsLeaseProbe.dll"));
        if (!File.Exists(probe)) throw new InvalidOperationException($"WPS lease probe is unavailable: {probe}");
        var start = new ProcessStartInfo
        {
            FileName = "dotnet",
            UseShellExecute = false,
        };
        foreach (var argument in new[] { probe, route, lockPath, eventLog, holdMilliseconds.ToString(), timeoutMilliseconds.ToString() })
            start.ArgumentList.Add(argument);
        return Process.Start(start) ?? throw new InvalidOperationException("Failed to start WPS lease probe.");
    }

    private static async Task WaitForExit(Process process)
    {
        await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
    }

    private static async Task WaitForEvent(string eventLog, char prefix, TimeSpan timeout)
    {
        var started = Stopwatch.StartNew();
        while (started.Elapsed < timeout)
        {
            if (File.Exists(eventLog) && File.ReadAllLines(eventLog).Any(line => line.StartsWith(prefix))) return;
            await Task.Delay(25);
        }
        throw new TimeoutException($"Lease probe did not publish {prefix} within {timeout}.");
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
