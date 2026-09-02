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
        var renderInput = DocxWpsRenderNormalizer.Prepare(input, tempRoot);

        try
        {
            using var lease = AcquireRuntimeLease();
            RunWithTransientStartupRetry(
                () => RunWpsHelper(xvfb, dbusRunSession, python, helperPath, renderInput, output, tempRoot),
                () => { if (File.Exists(output)) File.Delete(output); });
        }
        finally
        {
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }

    public static void RefreshDocxFields(string input, string output)
    {
        if (!File.Exists(input))
            throw new InvalidOperationException($"Input file not found: {input}");

        var python = FindWpsRpcPython()
            ?? throw new InvalidOperationException("WPS RPC python is required for WPS document field refresh. Set TIWATER_WPSRPC_PYTHON.");
        var xvfb = FindOnPath("xvfb-run")
            ?? throw new InvalidOperationException("xvfb-run is required for WPS document field refresh.");
        var dbusRunSession = FindOnPath("dbus-run-session")
            ?? throw new InvalidOperationException("dbus-run-session is required for WPS document field refresh.");
        if (string.IsNullOrWhiteSpace(FindOnPath("wps")))
            throw new InvalidOperationException("WPS command not found: wps");

        var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDir)) Directory.CreateDirectory(outputDir);
        var tempRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wps-refresh-{Guid.NewGuid():N}");
        Directory.CreateDirectory(tempRoot);
        var helperPath = Path.Combine(tempRoot, "refresh_docx_fields_wps.py");
        var refreshedPath = Path.Combine(tempRoot, "wps-refreshed.docx");
        File.WriteAllText(helperPath, RefreshFieldsHelperScript);

        try
        {
            using var lease = AcquireRuntimeLease();
            RunWithTransientStartupRetry(
                () => RunFieldRefreshHelper(xvfb, dbusRunSession, python, helperPath, input, refreshedPath, tempRoot),
                () => { if (File.Exists(refreshedPath)) File.Delete(refreshedPath); });
            DocxFieldResultMerger.Merge(input, refreshedPath, output);
        }
        finally
        {
            try { Directory.Delete(tempRoot, recursive: true); } catch { }
        }
    }

    private static void RunFieldRefreshHelper(string xvfb, string dbusRunSession, string python, string helperPath, string input, string output, string tempRoot)
    {
        var completionMarker = Path.Combine(tempRoot, "writer-output-complete");
        if (File.Exists(completionMarker)) File.Delete(completionMarker);
        var startInfo = CreateProcessStartInfo(xvfb, tempRoot);
        foreach (var arg in CreateHelperArguments(dbusRunSession, python, helperPath, input, output, completionMarker))
            startInfo.ArgumentList.Add(arg);
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start WPS RPC document field refresh.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        var completedOutput = WpsRpcSession.WaitForCompletedOutputOrExit(
            process, completionMarker, () => IsDocx(output), WpsRpcSession.OfficeOperationTimeout,
            "WPS RPC document field refresh timed out after 600 seconds.");
        var details = WpsRpcSession.CollectDiagnosticOutput(stdoutTask, stderrTask, TimeSpan.FromMilliseconds(250));
        if ((!completedOutput && process.ExitCode != 0) || !IsDocx(output))
            throw new InvalidOperationException($"WPS RPC failed to refresh document fields for {input}." +
                (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
    }

    private static void RunWpsHelper(string xvfb, string dbusRunSession, string python, string helperPath, string input, string output, string tempRoot)
    {
        var completionMarker = Path.Combine(tempRoot, "writer-output-complete");
        var startInfo = CreateProcessStartInfo(xvfb, tempRoot);
        foreach (var arg in CreateHelperArguments(dbusRunSession, python, helperPath, input, output, completionMarker)) startInfo.ArgumentList.Add(arg);
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start WPS RPC conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        var completedOutput = WpsRpcSession.WaitForCompletedOutputOrExit(
            process, completionMarker, () => IsPdf(output), WpsRpcSession.OfficeOperationTimeout,
            "WPS RPC PDF conversion timed out after 600 seconds.");
        var details = WpsRpcSession.CollectDiagnosticOutput(
            stdoutTask, stderrTask, TimeSpan.FromMilliseconds(250));
        if ((!completedOutput && process.ExitCode != 0) || !IsPdf(output))
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

    internal static string[] CreateHelperArguments(string dbusRunSession, string python, string helperPath, string input, string output, string completionMarker)
        => new[]
        {
            "-a",
            dbusRunSession,
            "--",
            python,
            helperPath,
            Path.GetFullPath(input),
            Path.GetFullPath(output),
            Path.GetFullPath(completionMarker),
        };

    public static bool IsTransientStartupFailure(string message)
        => message.Contains("getWpsApplication failed", StringComparison.OrdinalIgnoreCase)
            || message.Contains("get_Documents failed", StringComparison.OrdinalIgnoreCase)
            || message.Contains("Fatal IO error on X server", StringComparison.OrdinalIgnoreCase);

    internal static void RunWithTransientStartupRetry(Action operation, Action cleanup, Action? delay = null)
    {
        try
        {
            operation();
        }
        catch (InvalidOperationException error) when (IsTransientStartupFailure(error.Message))
        {
            cleanup();
            (delay ?? (() => Thread.Sleep(1000)))();
            operation();
        }
    }

    private static bool IsPdf(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 5) return false;
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 && header.SequenceEqual("%PDF"u8);
    }

    private static bool IsDocx(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 4) return false;
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 && header.SequenceEqual("PK\u0003\u0004"u8);
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
completion_marker = os.path.realpath(sys.argv[3])
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
        with open(completion_marker, "x", encoding="utf-8") as marker:
            marker.write("complete\n")
            marker.flush()
            os.fsync(marker.fileno())
    finally:
        document.Close(False)
finally:
    app.Quit()
""";

    internal const string RefreshFieldsHelperScript = """
import os
import sys

from pywpsrpc.rpcwpsapi import createWpsRpcInstance, wpsapi
from pywpsrpc.common import S_OK, QtApp

input_path = os.path.realpath(sys.argv[1])
output_path = os.path.realpath(sys.argv[2])
completion_marker = os.path.realpath(sys.argv[3])
os.makedirs(os.path.dirname(output_path), exist_ok=True)

def require(label, result):
    hr = result[0] if isinstance(result, tuple) else result
    if hr != S_OK:
        raise SystemExit(f"{label} failed: {hex(hr & 0xffffffff)}")
    return result[1] if isinstance(result, tuple) and len(result) > 1 else None

q_app = QtApp(sys.argv)
rpc = require("createWpsRpcInstance", createWpsRpcInstance())
app = require("getWpsApplication", rpc.getWpsApplication())

try:
    app.Visible = False
    app.DisplayAlerts = False
    documents = require("get_Documents", app.get_Documents())
    document = require("Documents.Open", documents.Open(input_path, ReadOnly=False, AddToRecentFiles=False, Visible=False))
    try:
        tables_of_contents = require("get_TablesOfContents", document.get_TablesOfContents())
        toc_count = require("TablesOfContents.get_Count", tables_of_contents.get_Count())
        for index in range(1, toc_count + 1):
            toc = require("TablesOfContents.Item", tables_of_contents.Item(index))
            require("TableOfContents.Update", toc.Update())

        tables_of_figures = require("get_TablesOfFigures", document.get_TablesOfFigures())
        figure_count = require("TablesOfFigures.get_Count", tables_of_figures.get_Count())
        for index in range(1, figure_count + 1):
            figure = require("TablesOfFigures.Item", tables_of_figures.Item(index))
            require("TableOfFigures.Update", figure.Update())

        require("Document.Repaginate", document.Repaginate())
        for index in range(1, toc_count + 1):
            toc = require("TablesOfContents.Item", tables_of_contents.Item(index))
            require("TableOfContents.UpdatePageNumbers", toc.UpdatePageNumbers())
        for index in range(1, figure_count + 1):
            figure = require("TablesOfFigures.Item", tables_of_figures.Item(index))
            require("TableOfFigures.UpdatePageNumbers", figure.UpdatePageNumbers())
        require("Document.SaveAs2 DOCX", document.SaveAs2(output_path, FileFormat=wpsapi.wdFormatXMLDocument))
        with open(completion_marker, "x", encoding="utf-8") as marker:
            marker.write("complete\n")
            marker.flush()
            os.fsync(marker.fileno())
    finally:
        document.Close(False)
finally:
    app.Quit()
""";
}
