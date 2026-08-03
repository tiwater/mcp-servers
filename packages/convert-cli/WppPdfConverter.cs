using System.Diagnostics;

namespace Dockit.Convert;

public static class WppPdfConverter
{
    internal static IDisposable AcquireRuntimeLease(TimeSpan? timeout = null, string? lockPath = null)
        => WpsRpcSession.AcquireOfficeLease(timeout, lockPath);

    public static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindWpsRpcPython())
            && !string.IsNullOrWhiteSpace(FindOnPath("xvfb-run"))
            && !string.IsNullOrWhiteSpace(FindOnPath("dbus-run-session"))
            && !string.IsNullOrWhiteSpace(FindOnPath("wpp"));

    public static void ConvertToPdf(string input, string output)
    {
        if (!File.Exists(input)) throw new InvalidOperationException($"Input file not found: {input}");

        var python = FindWpsRpcPython()
            ?? throw new InvalidOperationException("WPS RPC python is required for WPP PDF conversion. Set TIWATER_WPSRPC_PYTHON.");
        var xvfb = FindOnPath("xvfb-run")
            ?? throw new InvalidOperationException("xvfb-run is required for WPP PDF conversion.");
        var dbusRunSession = FindOnPath("dbus-run-session")
            ?? throw new InvalidOperationException("dbus-run-session is required for WPP PDF conversion.");
        if (string.IsNullOrWhiteSpace(FindOnPath("wpp"))) throw new InvalidOperationException("WPP command not found: wpp");

        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        var temporaryRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wpp-{Guid.NewGuid():N}");
        Directory.CreateDirectory(temporaryRoot);
        var helperPath = Path.Combine(temporaryRoot, "presentation_to_pdf_wps.py");
        File.WriteAllText(helperPath, EtHelperScript);

        try
        {
            using var lease = AcquireRuntimeLease();
            Exception? lastError = null;
            for (var attempt = 1; attempt <= 2; attempt++)
            {
                try
                {
                    RunWpsHelper(dbusRunSession, xvfb, python, helperPath, input, output, temporaryRoot);
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
            try { Directory.Delete(temporaryRoot, recursive: true); } catch { }
        }
    }

    private static void RunWpsHelper(string dbusRunSession, string xvfb, string python, string helperPath, string input, string output, string temporaryRoot)
    {
        var startInfo = CreateProcessStartInfo(dbusRunSession, xvfb, python, helperPath, input, output, temporaryRoot);
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start WPP RPC conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(3)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("WPP RPC PDF conversion timed out after 180 seconds.");
        }

        var details = WpsRpcSession.CollectDiagnosticOutput(
            stdoutTask, stderrTask, TimeSpan.FromMilliseconds(250));
        if (process.ExitCode != 0 || !IsPdf(output))
        {
            throw new InvalidOperationException("WPP RPC failed to convert presentation to PDF."
                + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
    }

    internal static ProcessStartInfo CreateProcessStartInfo(
        string dbusRunSession,
        string xvfb,
        string python,
        string helperPath,
        string input,
        string output,
        string isolatedWorkingDirectory)
    {
        var workingDirectory = Path.GetFullPath(isolatedWorkingDirectory);
        var cacheDirectory = Path.Combine(workingDirectory, "cache");
        var runtimeDirectory = Path.Combine(workingDirectory, "runtime");
        Directory.CreateDirectory(cacheDirectory);
        Directory.CreateDirectory(runtimeDirectory);
        if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(runtimeDirectory, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);

        var startInfo = new ProcessStartInfo
        {
            FileName = dbusRunSession,
            WorkingDirectory = workingDirectory,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        startInfo.Environment["XDG_CACHE_HOME"] = cacheDirectory;
        startInfo.Environment["XDG_RUNTIME_DIR"] = runtimeDirectory;
        foreach (var argument in new[] { "--", xvfb, "-a", python, helperPath, Path.GetFullPath(input), Path.GetFullPath(output) }) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }

    public static bool IsTransientStartupFailure(string message)
        => message.Contains("getWppApplication failed", StringComparison.OrdinalIgnoreCase)
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

from pywpsrpc.rpcwppapi import createWppRpcInstance, wppapi
from pywpsrpc.common import S_OK, QtApp

input_path = os.path.realpath(sys.argv[1])
output_path = os.path.realpath(sys.argv[2])
os.makedirs(os.path.dirname(output_path), exist_ok=True)

q_app = QtApp(sys.argv)
hr, rpc = createWppRpcInstance()
if hr != S_OK:
    raise SystemExit(f"createWppRpcInstance failed: {hex(hr & 0xffffffff)}")

hr, app = rpc.getWppApplication()
if hr != S_OK:
    raise SystemExit(f"getWppApplication failed: {hex(hr & 0xffffffff)}")

try:
    hr, presentations = app.get_Presentations()
    if hr != S_OK:
        raise SystemExit(f"get_Presentations failed: {hex(hr & 0xffffffff)}")
    hr, presentation = presentations.Open(input_path, ReadOnly=True, Untitled=False, WithWindow=False)
    if hr != S_OK:
        raise SystemExit(f"Presentations.Open failed: {hex(hr & 0xffffffff)}")
    try:
        hr = presentation.SaveAs(output_path, wppapi.ppSaveAsPDF)
        if hr != S_OK:
            raise SystemExit(f"Presentation.SaveAs PDF failed: {hex(hr & 0xffffffff)}")
    finally:
        presentation.Close()
finally:
    app.Quit()
""";
}
