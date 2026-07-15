using System.Diagnostics;

namespace Dockit.Convert;

internal static class WpsRpcSession
{
    internal static bool IsAvailable()
        => !string.IsNullOrWhiteSpace(FindOnPath("dbus-run-session"))
            && !string.IsNullOrWhiteSpace(FindOnPath("xvfb-run"));

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
        var startInfo = new ProcessStartInfo
        {
            FileName = dbusRunSession,
            WorkingDirectory = Path.GetFullPath(isolatedWorkingDirectory),
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        foreach (var arg in new[]
        {
            "--",
            xvfb,
            "-a",
            python,
            helperPath,
            Path.GetFullPath(input),
            Path.GetFullPath(output),
        })
        {
            startInfo.ArgumentList.Add(arg);
        }
        return startInfo;
    }

    internal static string? FindOnPath(string command)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, OperatingSystem.IsWindows() ? $"{command}.exe" : command);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        return null;
    }
}
