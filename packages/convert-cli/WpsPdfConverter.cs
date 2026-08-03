using System.Diagnostics;

namespace Dockit.Convert;

public static class WpsPdfConverter
{
    internal static IDisposable AcquireRuntimeLease(TimeSpan? timeout = null, string? lockPath = null)
        => WpsRpcSession.AcquireOfficeLease(timeout, lockPath);

    public static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindWpsRpcPython())
            && !string.IsNullOrWhiteSpace(FindOnPath("xvfb-run"))
            && !string.IsNullOrWhiteSpace(FindOnPath("dbus-run-session"))
            && !string.IsNullOrWhiteSpace(FindOnPath("wps"));

    public static void ConvertToPdf(string input, string output)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var python = FindWpsRpcPython()
            ?? throw new InvalidOperationException("WPS RPC python is required for WPS PDF conversion. Set TIWATER_WPSRPC_PYTHON.");
        var xvfb = FindOnPath("xvfb-run")
            ?? throw new InvalidOperationException("xvfb-run is required for WPS PDF conversion.");
        var dbusRunSession = FindOnPath("dbus-run-session")
            ?? throw new InvalidOperationException("dbus-run-session is required for WPS PDF conversion.");
        if (string.IsNullOrWhiteSpace(FindOnPath("wps")))
        {
            throw new InvalidOperationException("WPS command not found: wps");
        }

        var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDir)) Directory.CreateDirectory(outputDir);

        var tempRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wps-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        var helperPath = Path.Combine(tempRoot, "writer_to_pdf_wps.py");
        File.WriteAllText(helperPath, EtHelperScript);

        try
        {
            using var lease = AcquireRuntimeLease();
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
        var startInfo = CreateProcessStartInfo(xvfb, tempRoot);
        foreach (var arg in CreateHelperArguments(dbusRunSession, python, helperPath, input, output)) startInfo.ArgumentList.Add(arg);
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start WPS RPC conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(3)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("WPS RPC PDF conversion timed out after 180 seconds.");
        }
        var details = WpsRpcSession.CollectDiagnosticOutput(
            stdoutTask, stderrTask, TimeSpan.FromMilliseconds(250));
        if (process.ExitCode != 0 || !IsPdf(output))
        {
            throw new InvalidOperationException($"WPS RPC failed to convert {input} to PDF." + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
    }

    internal static ProcessStartInfo CreateProcessStartInfo(string executable, string isolatedWorkingDirectory)
    {
        var workingDirectory = Path.GetFullPath(isolatedWorkingDirectory);
        var cacheDirectory = Path.Combine(workingDirectory, "cache");
        var runtimeDirectory = Path.Combine(workingDirectory, "runtime");
        Directory.CreateDirectory(cacheDirectory);
        Directory.CreateDirectory(runtimeDirectory);
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(runtimeDirectory, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            WorkingDirectory = workingDirectory,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        startInfo.Environment["XDG_CACHE_HOME"] = cacheDirectory;
        startInfo.Environment["XDG_RUNTIME_DIR"] = runtimeDirectory;
        return startInfo;
    }

    internal static string[] CreateHelperArguments(string dbusRunSession, string python, string helperPath, string input, string output)
        => new[]
        {
            "-a",
            dbusRunSession,
            "--",
            python,
            helperPath,
            Path.GetFullPath(input),
            Path.GetFullPath(output),
        };

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
        foreach (var envName in new[] { "TIWATER_WPSRPC_PYTHON" })
        {
            var value = Environment.GetEnvironmentVariable(envName);
            if (!string.IsNullOrWhiteSpace(value) && File.Exists(value)) return Path.GetFullPath(value);
        }
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var candidate = Path.Combine(home, ".local", "share", "tiwater", "wpsrpc-venv", "bin", "python");
        return File.Exists(candidate) ? Path.GetFullPath(candidate) : null;
    }

    private static string? FindOnPath(string command)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, OperatingSystem.IsWindows() ? $"{command}.exe" : command);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        return null;
    }

    private const string EtHelperScript = """
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
