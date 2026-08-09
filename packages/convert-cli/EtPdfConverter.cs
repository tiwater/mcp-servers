using System.Diagnostics;

namespace Dockit.Convert;

public static class EtPdfConverter
{
    internal static IDisposable AcquireRuntimeLease(TimeSpan? timeout = null, string? lockPath = null)
        => WpsRpcSession.AcquireEtLease(timeout, lockPath);

    public static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindWpsRpcPython())
            && WpsRpcSession.IsAvailable()
            && !string.IsNullOrWhiteSpace(FindOnPath("et"));

    public static void ConvertToPdf(string input, string output)
    {
        if (!File.Exists(input)) throw new InvalidOperationException($"Input file not found: {input}");

        var python = FindWpsRpcPython()
            ?? throw new InvalidOperationException("WPS RPC python is required for ET PDF conversion. Set TIWATER_WPSRPC_PYTHON.");
        var xvfb = WpsRpcSession.RequireCommand("xvfb-run", "ET PDF conversion");
        var dbusRunSession = WpsRpcSession.RequireCommand("dbus-run-session", "ET PDF conversion");
        if (string.IsNullOrWhiteSpace(FindOnPath("et"))) throw new InvalidOperationException("ET command not found: et");

        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        var temporaryRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-et-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporaryRoot);
        var helperPath = Path.Combine(temporaryRoot, "spreadsheet_to_pdf_wps.py");
        File.WriteAllText(helperPath, EtHelperScript);

        try
        {
            using var lease = AcquireRuntimeLease();
            var completionMarker = Path.Combine(temporaryRoot, "spreadsheet-output-complete");
            var startInfo = WpsRpcSession.CreateProcessStartInfo(
                dbusRunSession, xvfb, python, helperPath, input, output, completionMarker, temporaryRoot);
            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Failed to start ET RPC conversion.");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            var completedOutput = WpsRpcSession.WaitForCompletedOutputOrExit(
                process, completionMarker, () => IsPdf(output), TimeSpan.FromMinutes(3),
                "ET RPC PDF conversion timed out after 180 seconds.");

            var details = WpsRpcSession.CollectDiagnosticOutput(
                stdoutTask, stderrTask, TimeSpan.FromMilliseconds(250));
            if ((!completedOutput && process.ExitCode != 0) || !IsPdf(output))
            {
                throw new InvalidOperationException("ET RPC failed to convert workbook to PDF."
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
        var candidate = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "share", "tiwater", "wpsrpc-venv", "bin", "python");
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

    private const string EtHelperScript = """
import os
import sys

from pywpsrpc.rpcetapi import createEtRpcInstance, etapi
from pywpsrpc.common import S_OK, QtApp

input_path = os.path.realpath(sys.argv[1])
output_path = os.path.realpath(sys.argv[2])
completion_marker = os.path.realpath(sys.argv[3])
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
        with open(completion_marker, "x", encoding="utf-8") as marker:
            marker.write("complete\n")
            marker.flush()
            os.fsync(marker.fileno())
    finally:
        book.Close(False)
finally:
    app.Quit()
""";
}
