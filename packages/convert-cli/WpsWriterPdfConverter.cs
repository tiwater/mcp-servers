using System.Diagnostics;

namespace Dockit.Convert;

public static class WpsWriterPdfConverter
{
    public static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindWpsRpcPython())
            && WpsRpcSession.IsAvailable()
            && !string.IsNullOrWhiteSpace(WpsRpcSession.FindOnPath("wps"));

    public static void ConvertToPdf(string input, string output)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var python = FindWpsRpcPython()
            ?? throw new InvalidOperationException("WPS RPC python is required for WPS Writer PDF conversion. Set TIWATER_WPSRPC_PYTHON or LUCID_WPSRPC_PYTHON.");
        var xvfb = WpsRpcSession.RequireCommand("xvfb-run", "WPS Writer PDF conversion");
        var dbusRunSession = WpsRpcSession.RequireCommand("dbus-run-session", "WPS Writer PDF conversion");
        if (string.IsNullOrWhiteSpace(WpsRpcSession.FindOnPath("wps")))
        {
            throw new InvalidOperationException("WPS Writer command not found: wps");
        }

        var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDir)) Directory.CreateDirectory(outputDir);

        var tempRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wps-writer-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        var helperPath = Path.Combine(tempRoot, "writer_to_pdf_wps.py");
        File.WriteAllText(helperPath, WpsHelperScript);

        try
        {
            Exception? lastError = null;
            for (var attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    RunWpsHelper(xvfb, dbusRunSession, python, helperPath, input, output, tempRoot);
                    lastError = null;
                    break;
                }
                catch (InvalidOperationException error) when (attempt == 1 && IsTransientStartupFailure(error.Message))
                {
                    lastError = error;
                    if (File.Exists(output)) File.Delete(output);
                    Thread.Sleep(1000);
                }
            }
            if (lastError is not null) throw lastError;
        }
        finally
        {
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }

    private static void RunWpsHelper(string xvfb, string dbusRunSession, string python, string helperPath, string input, string output, string tempRoot)
    {
        var startInfo = WpsRpcSession.CreateProcessStartInfo(dbusRunSession, xvfb, python, helperPath, input, output, tempRoot);
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start WPS Writer RPC conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(3)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("WPS Writer RPC PDF conversion timed out after 180 seconds.");
        }
        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0 || !IsPdf(output))
        {
            var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static s => !string.IsNullOrWhiteSpace(s)));
            throw new InvalidOperationException($"WPS Writer RPC failed to convert {input} to PDF." + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
    }

    public static bool IsTransientStartupFailure(string message)
        => message.Contains("getWpsApplication failed", StringComparison.OrdinalIgnoreCase)
            || message.Contains("Fatal IO error on X server", StringComparison.OrdinalIgnoreCase);

    private static bool IsPdf(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 5) return false;
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 && header.SequenceEqual("%PDF"u8);
    }

    private static string? FindWpsRpcPython()
    {
        foreach (var envName in new[] { "TIWATER_WPSRPC_PYTHON", "LUCID_WPSRPC_PYTHON" })
        {
            var value = Environment.GetEnvironmentVariable(envName);
            if (!string.IsNullOrWhiteSpace(value) && File.Exists(value)) return Path.GetFullPath(value);
        }
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var candidate = Path.Combine(home, ".local", "share", "lucid-docs", "wpsrpc-venv", "bin", "python");
        return File.Exists(candidate) ? Path.GetFullPath(candidate) : null;
    }

    private const string WpsHelperScript = """
import os
import sys

from pywpsrpc.rpcwpsapi import createWpsRpcInstance, wpsapi
from pywpsrpc.common import S_OK, QtApp

input_path = os.path.realpath(sys.argv[1])
output_path = os.path.realpath(sys.argv[2])
os.makedirs(os.path.dirname(output_path), exist_ok=True)

q_app = QtApp(sys.argv)
hr, rpc = createWpsRpcInstance()
if hr != S_OK:
    raise SystemExit(f"createWpsRpcInstance failed: {hex(hr & 0xffffffff)}")

hr, app = rpc.getWpsApplication()
if hr != S_OK:
    raise SystemExit(f"getWpsApplication failed: {hex(hr & 0xffffffff)}")

try:
    app.Visible = False
    hr, documents = app.get_Documents()
    if hr != S_OK:
        raise SystemExit(f"get_Documents failed: {hex(hr & 0xffffffff)}")
    hr, document = documents.Open(input_path, ReadOnly=True, AddToRecentFiles=False, Visible=False)
    if hr != S_OK:
        raise SystemExit(f"Documents.Open failed: {hex(hr & 0xffffffff)}")
    try:
        hr = document.SaveAs2(output_path, FileFormat=wpsapi.wdFormatPDF)
        if hr != S_OK:
            raise SystemExit(f"Document.SaveAs2 PDF failed: {hex(hr & 0xffffffff)}")
    finally:
        document.Close(False)
finally:
    app.Quit()
""";
}
