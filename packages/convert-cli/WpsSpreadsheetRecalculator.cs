using System.Diagnostics;
using NPOI.XSSF.UserModel;

namespace Dockit.Convert;

public static class WpsSpreadsheetRecalculator
{
    public static bool IsAvailable()
        => FindPython() is not null && WpsRpcSession.IsAvailable() && FindOnPath("et") is not null;

    public static void Recalculate(string input, string output)
    {
        input = Path.GetFullPath(input); output = Path.GetFullPath(output);
        if (!File.Exists(input)) throw new InvalidOperationException($"Input file not found: {input}");
        if (!string.Equals(Path.GetExtension(input), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("WPS recalculation input must be an XLSX file.");
        if (string.Equals(input, output, StringComparison.Ordinal)) throw new InvalidOperationException("WPS recalculation requires a distinct output path.");
        var python = FindPython() ?? throw new InvalidOperationException("WPS RPC python is required for XLSX recalculation.");
        var xvfb = WpsRpcSession.RequireCommand("xvfb-run", "WPS XLSX recalculation");
        var dbus = WpsRpcSession.RequireCommand("dbus-run-session", "WPS XLSX recalculation");
        if (FindOnPath("et") is null) throw new InvalidOperationException("WPS Spreadsheets command not found: et");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        var root = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wps-recalculate-{Guid.NewGuid():N}"); Directory.CreateDirectory(root);
        var helper = Path.Combine(root, "recalculate_xlsx_wps.py"); File.WriteAllText(helper, WpsHelperScript);
        try
        {
            using var process = Process.Start(WpsRpcSession.CreateProcessStartInfo(dbus, xvfb, python, helper, input, output, root)) ?? throw new InvalidOperationException("Failed to start WPS XLSX recalculation.");
            var stdout = process.StandardOutput.ReadToEndAsync(); var stderr = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(TimeSpan.FromMinutes(10))) { try { process.Kill(entireProcessTree: true); } catch { } throw new TimeoutException("WPS XLSX recalculation timed out after 600 seconds."); }
            var details = string.Join(" ", new[] { stdout.GetAwaiter().GetResult().Trim(), stderr.GetAwaiter().GetResult().Trim() }.Where(value => value.Length > 0));
            if (process.ExitCode != 0 || !File.Exists(output)) throw new InvalidOperationException("WPS XLSX recalculation failed." + (details.Length > 0 ? $" {details}" : string.Empty));
            using var stream = File.OpenRead(output); using var workbook = new XSSFWorkbook(stream); if (workbook.NumberOfSheets < 1) throw new InvalidOperationException("WPS recalculation produced an XLSX without worksheets.");
        }
        finally { try { Directory.Delete(root, recursive: true); } catch { } }
    }

    private static string? FindPython()
    {
        foreach (var name in new[] { "TIWATER_WPSRPC_PYTHON", "LUCID_WPSRPC_PYTHON" }) { var value = Environment.GetEnvironmentVariable(name); if (!string.IsNullOrWhiteSpace(value) && File.Exists(value)) return Path.GetFullPath(value); }
        var candidate = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "share", "lucid-docs", "wpsrpc-venv", "bin", "python"); return File.Exists(candidate) ? candidate : null;
    }

    private static string? FindOnPath(string command) => (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries).Select(directory => Path.Combine(directory, OperatingSystem.IsWindows() ? $"{command}.exe" : command)).FirstOrDefault(File.Exists);

    internal const string WpsHelperScript = """
import os
import sys
from pywpsrpc.rpcetapi import createEtRpcInstance, etapi
from pywpsrpc.common import S_OK, QtApp
input_path = os.path.realpath(sys.argv[1])
output_path = os.path.realpath(sys.argv[2])
os.makedirs(os.path.dirname(output_path), exist_ok=True)
q_app = QtApp(sys.argv)
hr, rpc = createEtRpcInstance()
if hr != S_OK: raise SystemExit(f"createEtRpcInstance failed: {hex(hr & 0xffffffff)}")
hr, app = rpc.getEtApplication()
if hr != S_OK: raise SystemExit(f"getEtApplication failed: {hex(hr & 0xffffffff)}")
try:
    app.Visible = False
    hr, books = app.get_Workbooks()
    if hr != S_OK: raise SystemExit(f"get_Workbooks failed: {hex(hr & 0xffffffff)}")
    hr, book = books.Open(input_path, ReadOnly=False, IgnoreReadOnlyRecommended=True, AddToMru=False)
    if hr != S_OK: raise SystemExit(f"Workbooks.Open failed: {hex(hr & 0xffffffff)}")
    try:
        app.Calculation = etapi.XlCalculation.xlCalculationAutomatic
        hr = app.CalculateFull()
        if hr != S_OK: raise SystemExit(f"Application.CalculateFull failed: {hex(hr & 0xffffffff)}")
        hr = book.SaveAs(output_path, FileFormat=etapi.XlFileFormat.xlOpenXMLWorkbook, AddToMru=False)
        if hr != S_OK: raise SystemExit(f"Workbook.SaveAs XLSX failed: {hex(hr & 0xffffffff)}")
    finally: book.Close(False)
finally: app.Quit()
""";
}
