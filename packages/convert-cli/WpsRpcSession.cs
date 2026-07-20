using System.Diagnostics;

namespace Dockit.Convert;

internal static class WpsRpcSession
{
    internal static bool IsAvailable()
        => FindOnPath("dbus-run-session") is not null
            && FindOnPath("xvfb-run") is not null;

    internal static string RequireCommand(string command, string purpose)
        => FindOnPath(command)
            ?? throw new InvalidOperationException($"{command} is required for {purpose}.");

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
        foreach (var argument in new[]
        {
            "--", xvfb, "-a", python, helperPath,
            Path.GetFullPath(input), Path.GetFullPath(output),
        })
        {
            startInfo.ArgumentList.Add(argument);
        }
        return startInfo;
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
