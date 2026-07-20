using System.Diagnostics;
using NPOI.XSSF.UserModel;

namespace Dockit.Convert;

public static class WpsSpreadsheetConverter
{
    public static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindWpsRpcPython())
            && WpsRpcSession.IsAvailable()
            && !string.IsNullOrWhiteSpace(FindOnPath("et"));

    public static void ConvertXlsToXlsx(string input, string output)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var python = FindWpsRpcPython();
        if (string.IsNullOrWhiteSpace(python))
        {
            throw new InvalidOperationException(
                "WPS RPC python is required for WPS XLS conversion. Set TIWATER_WPSRPC_PYTHON or LUCID_WPSRPC_PYTHON.");
        }
        var xvfb = WpsRpcSession.RequireCommand("xvfb-run", "WPS XLS conversion");
        var dbusRunSession = WpsRpcSession.RequireCommand("dbus-run-session", "WPS XLS conversion");
        if (string.IsNullOrWhiteSpace(FindOnPath("et")))
        {
            throw new InvalidOperationException("WPS Spreadsheets command not found: et");
        }

        var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var tempRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wps-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        var helperPath = Path.Combine(tempRoot, "xls_to_xlsx_wps.py");
        File.WriteAllText(helperPath, WpsHelperScript);

        try
        {
            var startInfo = WpsRpcSession.CreateProcessStartInfo(
                dbusRunSession, xvfb, python, helperPath, input, output, tempRoot);

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Failed to start WPS RPC conversion.");
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(TimeSpan.FromMinutes(2)))
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // Ignore kill races; the timeout error is the actionable failure.
                }

                throw new TimeoutException("WPS RPC XLS conversion timed out after 120 seconds.");
            }

            var stdout = stdoutTask.GetAwaiter().GetResult();
            var stderr = stderrTask.GetAwaiter().GetResult();
            if (process.ExitCode != 0 || !File.Exists(output))
            {
                var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static s => !string.IsNullOrWhiteSpace(s)));
                throw new InvalidOperationException(
                    $"WPS RPC failed to convert {input} to XLSX."
                    + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
            }

            using var stream = File.OpenRead(output);
            using var workbook = new XSSFWorkbook(stream);
            if (workbook.NumberOfSheets < 1)
            {
                throw new InvalidOperationException($"WPS RPC produced an XLSX without worksheets: {output}");
            }
        }
        finally
        {
            try
            {
                Directory.Delete(tempRoot, recursive: true);
            }
            catch
            {
                // Temporary cleanup failure should not invalidate a successful conversion.
            }
        }
    }

    private static string? FindWpsRpcPython()
    {
        foreach (var envName in new[] { "TIWATER_WPSRPC_PYTHON", "LUCID_WPSRPC_PYTHON" })
        {
            var value = Environment.GetEnvironmentVariable(envName);
            if (!string.IsNullOrWhiteSpace(value) && File.Exists(value))
            {
                return Path.GetFullPath(value);
            }
        }

        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrWhiteSpace(home))
        {
            var candidate = Path.Combine(home, ".local", "share", "lucid-docs", "wpsrpc-venv", "bin", "python");
            if (File.Exists(candidate))
            {
                return Path.GetFullPath(candidate);
            }
        }

        return null;
    }

    private static string? FindOnPath(string command)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, OperatingSystem.IsWindows() ? $"{command}.exe" : command);
            if (File.Exists(candidate))
            {
                return Path.GetFullPath(candidate);
            }
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
        hr = book.SaveAs(output_path, FileFormat=etapi.XlFileFormat.xlOpenXMLWorkbook, AddToMru=False)
        if hr != S_OK:
            raise SystemExit(f"Workbook.SaveAs XLSX failed: {hex(hr & 0xffffffff)}")
    finally:
        book.Close(False)
finally:
    app.Quit()
""";
}
