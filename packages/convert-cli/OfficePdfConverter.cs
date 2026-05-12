using System.Diagnostics;

namespace Dockit.Convert;

public static class OfficePdfConverter
{
    private static readonly string[] SupportedFormats =
    [
        "doc",
        "docx",
        "odt",
        "rtf",
        "xls",
        "xlsx",
        "ods",
        "ppt",
        "pptx",
        "odp",
    ];

    public static string? FindSofficeBinary()
    {
        foreach (var candidate in CandidatePaths())
        {
            if (IsExecutableFile(candidate))
            {
                return Path.GetFullPath(candidate);
            }
        }

        return null;
    }

    public static void ConvertToPdf(string input, string output, string sourceFormat, string? sofficePath = null)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var normalizedFormat = NormalizeFormat(sourceFormat);
        if (!SupportedFormats.Contains(normalizedFormat))
        {
            throw new InvalidOperationException($"Unsupported PDF source format: {sourceFormat}");
        }

        var inputExtension = NormalizeFormat(Path.GetExtension(input).TrimStart('.'));
        if (!string.IsNullOrWhiteSpace(inputExtension) && inputExtension != normalizedFormat)
        {
            throw new InvalidOperationException(
                $"Command source format {normalizedFormat} does not match input extension {inputExtension}: {input}");
        }

        var soffice = ResolveSofficeBinary(sofficePath);
        var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var tempRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-pdf-{Guid.NewGuid():N}");
        var exportDir = Path.Combine(tempRoot, "out");
        var profileDir = Path.Combine(tempRoot, "profile");
        Directory.CreateDirectory(exportDir);
        Directory.CreateDirectory(profileDir);

        try
        {
            var profileUri = new Uri(Path.GetFullPath(profileDir) + Path.DirectorySeparatorChar).AbsoluteUri;
            var startInfo = new ProcessStartInfo
            {
                FileName = soffice,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
            };
            startInfo.ArgumentList.Add($"-env:UserInstallation={profileUri}");
            startInfo.ArgumentList.Add("--headless");
            startInfo.ArgumentList.Add("--nologo");
            startInfo.ArgumentList.Add("--nofirststartwizard");
            startInfo.ArgumentList.Add("--convert-to");
            startInfo.ArgumentList.Add("pdf");
            startInfo.ArgumentList.Add("--outdir");
            startInfo.ArgumentList.Add(exportDir);
            startInfo.ArgumentList.Add(Path.GetFullPath(input));

            using var process = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Failed to start LibreOffice/soffice.");
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

                throw new TimeoutException("LibreOffice/soffice PDF conversion timed out after 120 seconds.");
            }

            var stdout = stdoutTask.GetAwaiter().GetResult();
            var stderr = stderrTask.GetAwaiter().GetResult();
            var generated = Path.Combine(exportDir, $"{Path.GetFileNameWithoutExtension(input)}.pdf");
            if (process.ExitCode != 0 || !File.Exists(generated))
            {
                var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static s => !string.IsNullOrWhiteSpace(s)));
                throw new InvalidOperationException(
                    $"LibreOffice/soffice failed to convert {input} to PDF."
                    + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
            }

            File.Copy(generated, output, overwrite: true);
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

    private static string ResolveSofficeBinary(string? sofficePath)
    {
        var resolved = string.IsNullOrWhiteSpace(sofficePath)
            ? FindSofficeBinary()
            : ResolveExplicitCandidate(sofficePath);

        if (string.IsNullOrWhiteSpace(resolved))
        {
            throw new InvalidOperationException(
                "LibreOffice/soffice is required for PDF conversion. Install LibreOffice or set TIWATER_SOFFICE, SOFFICE, or LIBREOFFICE_PATH.");
        }

        return resolved;
    }

    private static string? ResolveExplicitCandidate(string value)
    {
        foreach (var candidate in ExpandCandidate(value))
        {
            if (IsExecutableFile(candidate))
            {
                return Path.GetFullPath(candidate);
            }
        }

        return null;
    }

    private static IEnumerable<string> CandidatePaths()
    {
        foreach (var envName in new[] { "TIWATER_SOFFICE", "SOFFICE", "LIBREOFFICE_PATH", "LIBREOFFICE" })
        {
            var value = Environment.GetEnvironmentVariable(envName);
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            foreach (var candidate in ExpandCandidate(value))
            {
                yield return candidate;
            }
        }

        foreach (var command in new[] { "soffice", "libreoffice" })
        {
            foreach (var candidate in FindOnPath(command))
            {
                yield return candidate;
            }
        }

        foreach (var candidate in PlatformDefaultCandidates())
        {
            yield return candidate;
        }
    }

    private static IEnumerable<string> ExpandCandidate(string value)
    {
        yield return value;

        if (Directory.Exists(value))
        {
            yield return Path.Combine(value, OperatingSystem.IsWindows() ? "soffice.exe" : "soffice");
            yield return Path.Combine(value, "program", OperatingSystem.IsWindows() ? "soffice.exe" : "soffice");
            yield return Path.Combine(value, "Contents", "MacOS", "soffice");
        }
    }

    private static IEnumerable<string> PlatformDefaultCandidates()
    {
        if (OperatingSystem.IsMacOS())
        {
            yield return "/Applications/LibreOffice.app/Contents/MacOS/soffice";
        }
        else if (OperatingSystem.IsLinux())
        {
            yield return "/usr/bin/soffice";
            yield return "/usr/bin/libreoffice";
            yield return "/snap/bin/libreoffice";
        }
        else if (OperatingSystem.IsWindows())
        {
            var roots = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            };
            foreach (var root in roots.Where(static r => !string.IsNullOrWhiteSpace(r)))
            {
                yield return Path.Combine(root, "LibreOffice", "program", "soffice.exe");
            }
        }
    }

    private static IEnumerable<string> FindOnPath(string command)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            if (OperatingSystem.IsWindows())
            {
                yield return Path.Combine(directory, $"{command}.exe");
                yield return Path.Combine(directory, command);
            }
            else
            {
                yield return Path.Combine(directory, command);
            }
        }
    }

    private static bool IsExecutableFile(string? path)
    {
        return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
    }

    private static string NormalizeFormat(string value)
    {
        return (value ?? string.Empty).Trim().TrimStart('.').ToLowerInvariant();
    }
}
