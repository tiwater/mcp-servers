using System.Diagnostics;

namespace Dockit.Convert;

public static class WpsSpreadsheetPdfConverter
{
    internal static IDisposable AcquireRuntimeLease(TimeSpan? timeout = null, string? lockPath = null)
        => WpsRpcSession.AcquireSpreadsheetLease(timeout, lockPath);

    public static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindWpsRpcPython())
            && WpsRpcSession.IsAvailable()
            && !string.IsNullOrWhiteSpace(FindOnPath("et"));

    public static void ConvertToPdf(string input, string output)
    {
        if (!File.Exists(input)) throw new InvalidOperationException($"Input file not found: {input}");

        var python = FindWpsRpcPython()
            ?? throw new InvalidOperationException("WPS RPC python is required for WPS Spreadsheets PDF conversion. Set TIWATER_WPSRPC_PYTHON.");
        var xvfb = WpsRpcSession.RequireCommand("xvfb-run", "WPS Spreadsheets PDF conversion");
        var dbusRunSession = WpsRpcSession.RequireCommand("dbus-run-session", "WPS Spreadsheets PDF conversion");
        if (string.IsNullOrWhiteSpace(FindOnPath("et"))) throw new InvalidOperationException("WPS Spreadsheets command not found: et");

        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        var temporaryRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wps-spreadsheet-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporaryRoot);
        var helperPath = Path.Combine(temporaryRoot, "spreadsheet_to_pdf_wps.py");
        File.WriteAllText(helperPath, WpsHelperScript);

        try
        {
            using var lease = AcquireRuntimeLease();
            var startInfo = WpsRpcSession.CreateProcessStartInfo(
                dbusRunSession, xvfb, python, helperPath, input, output, temporaryRoot);
            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Failed to start WPS Spreadsheets RPC conversion.");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(TimeSpan.FromMinutes(3)))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                throw new TimeoutException("WPS Spreadsheets RPC PDF conversion timed out after 180 seconds.");
            }

            var stdout = stdoutTask.GetAwaiter().GetResult();
            var stderr = stderrTask.GetAwaiter().GetResult();
            if (process.ExitCode != 0 || !IsPdf(output))
            {
                var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static value => !string.IsNullOrWhiteSpace(value)));
                throw new InvalidOperationException("WPS Spreadsheets RPC failed to convert workbook to PDF."
                    + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
            }
        }
        finally
        {
            try { Directory.Delete(temporaryRoot, recursive: true); } catch { }
        }
    }

    private static bool IsPdf(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 5) return false;
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 && header.SequenceEqual("%PDF"u8);
    }

    private static string? FindWpsRpcPython()
    {
        foreach (var environmentName in new[] { "TIWATER_WPSRPC_PYTHON" })
        {
            var value = Environment.GetEnvironmentVariable(environmentName);
            if (!string.IsNullOrWhiteSpace(value) && File.Exists(value)) return Path.GetFullPath(value);
        }
        var candidate = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "share", "lucid-docs", "wpsrpc-venv", "bin", "python");
        return File.Exists(candidate) ? Path.GetFullPath(candidate) : null;
    }

    private static string? FindOnPath(string command)
    {
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, OperatingSystem.IsWindows() ? $"{command}.exe" : command);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        return null;
    }

    private const string WpsHelperScript = """
import os
import sys

from pywpsrpc.rpcetapi import createEtRpcInstance, etapi
from pywpsrpc.common import S_OK, QtApp

input_path = os.path.realpath(sys.argv[1])
output_path = os.path.realpath(sys.argv[2])
os.makedirs(os.path.dirname(output_path), exist_ok=True)

q_app = QtApp(sys.argv)
hr, rpc = createEtRpcInstance()
if hr != S_OK:
    raise SystemExit(f"createEtRpcInstance failed: {hex(hr & 0xffffffff)}")

hr, app = rpc.getEtApplication()
if hr != S_OK:
    raise SystemExit(f"getEtApplication failed: {hex(hr & 0xffffffff)}")

try:
    app.Visible = False
    hr, books = app.get_Workbooks()
    if hr != S_OK:
        raise SystemExit(f"get_Workbooks failed: {hex(hr & 0xffffffff)}")
    hr, book = books.Open(input_path, ReadOnly=True, IgnoreReadOnlyRecommended=True, AddToMru=False)
    if hr != S_OK:
        raise SystemExit(f"Workbooks.Open failed: {hex(hr & 0xffffffff)}")
    try:
        hr = book.ExportAsFixedFormat(
            etapi.XlFixedFormatType.xlTypePDF,
            output_path,
            Quality=etapi.XlFixedFormatQuality.xlQualityStandard,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False)
        if hr != S_OK:
            raise SystemExit(f"Workbook.ExportAsFixedFormat PDF failed: {hex(hr & 0xffffffff)}")
    finally:
        book.Close(False)
finally:
    app.Quit()
""";
}
