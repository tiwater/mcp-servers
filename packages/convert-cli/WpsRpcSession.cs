using System.Diagnostics;

namespace Dockit.Convert;

internal static class WpsRpcSession
{
    private static readonly string OfficeLeasePath = Path.Combine(Path.GetTempPath(), "tiwater-office.lock");

    internal static bool IsAvailable()
        => FindOnPath("dbus-run-session") is not null
            && FindOnPath("xvfb-run") is not null;

    internal static string RequireCommand(string command, string purpose)
        => FindOnPath(command)
            ?? throw new InvalidOperationException($"{command} is required for {purpose}.");

    internal static IDisposable AcquireOfficeLease(TimeSpan? timeout = null, string? lockPath = null)
        => AcquireFileLease(
            Path.GetFullPath(lockPath ?? OfficeLeasePath),
            timeout ?? TimeSpan.FromMinutes(5),
            "WPS Office runtime");

    internal static IDisposable AcquireEtLease(TimeSpan? timeout = null, string? lockPath = null)
        => AcquireOfficeLease(timeout, lockPath);

    internal static IDisposable AcquireContentLease(string lockPath, TimeSpan timeout)
        => AcquireFileLease(Path.GetFullPath(lockPath), timeout, "WPS spreadsheet content conversion");

    private static IDisposable AcquireFileLease(string absolute, TimeSpan wait, string label)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(absolute)!);
        var started = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                return new FileStream(absolute, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException) when (started.Elapsed < wait)
            {
                Thread.Sleep(100);
            }
            catch (IOException error)
            {
                throw new TimeoutException($"{label} remained busy for {wait}.", error);
            }
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
        => CreateProcessStartInfo(
            dbusRunSession, xvfb, python, helperPath, input, output, null, isolatedWorkingDirectory);

    internal static ProcessStartInfo CreateProcessStartInfo(
        string dbusRunSession,
        string xvfb,
        string python,
        string helperPath,
        string input,
        string output,
        string? completionMarker,
        string isolatedWorkingDirectory)
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
            FileName = dbusRunSession,
            WorkingDirectory = workingDirectory,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        startInfo.Environment["XDG_CACHE_HOME"] = cacheDirectory;
        startInfo.Environment["XDG_RUNTIME_DIR"] = runtimeDirectory;
        var arguments = new List<string>
        {
            "--", xvfb, "-a", python, helperPath,
            Path.GetFullPath(input), Path.GetFullPath(output),
        };
        if (!string.IsNullOrWhiteSpace(completionMarker)) arguments.Add(Path.GetFullPath(completionMarker));
        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }
        return startInfo;
    }

    internal static (string Stdout, string Stderr) CollectProcessOutput(
        Task<string> stdout,
        Task<string> stderr,
        TimeSpan wait)
    {
        try { Task.WhenAll(stdout, stderr).Wait(wait); } catch { }
        return (
            stdout.IsCompletedSuccessfully ? stdout.Result.Trim() : string.Empty,
            stderr.IsCompletedSuccessfully ? stderr.Result.Trim() : string.Empty);
    }

    internal static string CollectDiagnosticOutput(
        Task<string> stdout,
        Task<string> stderr,
        TimeSpan wait)
    {
        var output = CollectProcessOutput(stdout, stderr, wait);
        return string.Join(" ", new[] { output.Stdout, output.Stderr }.Where(value => value.Length > 0));
    }

    internal static bool WaitForCompletedOutputOrExit(
        Process process,
        string completionMarker,
        Func<bool> outputIsValid,
        TimeSpan timeout,
        string timeoutMessage,
        TimeSpan? gracefulExit = null)
    {
        var started = Stopwatch.StartNew();
        var grace = gracefulExit ?? TimeSpan.FromSeconds(2);
        while (started.Elapsed < timeout)
        {
            if (process.WaitForExit(100)) return false;
            if (!File.Exists(completionMarker)) continue;
            if (!outputIsValid())
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                throw new InvalidOperationException("WPS RPC reported completion without a valid output file.");
            }
            if (!process.WaitForExit(grace))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                process.WaitForExit(TimeSpan.FromSeconds(5));
            }
            return true;
        }
        try { process.Kill(entireProcessTree: true); } catch { }
        throw new TimeoutException(timeoutMessage);
    }

    private static string? FindOnPath(string command)
    {
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
                     .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, OperatingSystem.IsWindows() ? $"{command}.exe" : command);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        return null;
    }
}
