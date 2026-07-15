using System.Diagnostics;
using NPOI.XSSF.UserModel;

namespace Dockit.Convert;

public sealed record OfficePdfConversionResult(string Backend, string? FallbackReason = null);

public static class OfficeConverter
{
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

    public static OfficePdfConversionResult ConvertToPdf(string input, string output, string sourceFormat, string? sofficePath = null)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var normalizedFormat = NormalizeFormat(sourceFormat);
        var supportedFormats = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
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
        };
        if (!supportedFormats.Contains(normalizedFormat))
        {
            throw new InvalidOperationException($"Unsupported PDF source format: {sourceFormat}");
        }

        var inputExtension = NormalizeFormat(Path.GetExtension(input).TrimStart('.'));
        if (!string.IsNullOrWhiteSpace(inputExtension) && inputExtension != normalizedFormat)
        {
            throw new InvalidOperationException(
                $"Command source format {normalizedFormat} does not match input extension {inputExtension}: {input}");
        }

        var writerFormats = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "doc", "docx", "odt", "rtf" };
        var spreadsheetFormats = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "xls", "xlsx" };
        var requestedBackend = (Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND") ?? "auto").Trim().ToLowerInvariant();
        if (!new[] { "auto", "wps-writer", "wps-spreadsheet", "libreoffice" }.Contains(requestedBackend))
        {
            throw new InvalidOperationException($"Unsupported TIWATER_OFFICE_PDF_BACKEND: {requestedBackend}");
        }

        if (requestedBackend == "wps-writer" && !writerFormats.Contains(normalizedFormat))
            throw new InvalidOperationException($"WPS Writer PDF backend does not support {normalizedFormat} input.");
        if (requestedBackend == "wps-spreadsheet" && !spreadsheetFormats.Contains(normalizedFormat))
            throw new InvalidOperationException($"WPS Spreadsheets PDF backend does not support {normalizedFormat} input.");

        if (writerFormats.Contains(normalizedFormat) && requestedBackend != "libreoffice")
        {
            if (WpsWriterPdfConverter.IsAvailable())
            {
                WpsWriterPdfConverter.ConvertToPdf(input, output);
                return new OfficePdfConversionResult("wps-writer");
            }
            if (requestedBackend == "wps-writer")
            {
                throw new InvalidOperationException("WPS Writer PDF backend was required but WPS Writer, xvfb-run, dbus-run-session, or pywpsrpc is unavailable.");
            }
        }

        if (spreadsheetFormats.Contains(normalizedFormat) && requestedBackend != "libreoffice")
        {
            if (WpsSpreadsheetPdfConverter.IsAvailable())
            {
                WpsSpreadsheetPdfConverter.ConvertToPdf(input, output);
                return new OfficePdfConversionResult("wps-spreadsheet");
            }
            if (requestedBackend == "wps-spreadsheet")
            {
                throw new InvalidOperationException("WPS Spreadsheets PDF backend was required but WPS Spreadsheets, xvfb-run, dbus-run-session, or pywpsrpc is unavailable.");
            }
        }

        ConvertWithSoffice(input, output, "pdf", sofficePath);
        var fallbackReason = writerFormats.Contains(normalizedFormat)
            ? "wps-writer-unavailable"
            : spreadsheetFormats.Contains(normalizedFormat) ? "wps-spreadsheet-unavailable" : null;
        return new OfficePdfConversionResult("libreoffice", fallbackReason);
    }

    public static void ConvertXlsToXlsx(string input, string output, string? sofficePath = null)
    {
        if (!File.Exists(input))
        {
            throw new InvalidOperationException($"Input file not found: {input}");
        }

        var inputExtension = NormalizeFormat(Path.GetExtension(input).TrimStart('.'));
        if (inputExtension != "xls")
        {
            throw new InvalidOperationException($"Input extension must be .xls for xls-to-xlsx conversion: {input}");
        }

        ConvertWithSoffice(input, output, "xlsx", sofficePath);
        using var stream = File.OpenRead(output);
        using var workbook = new XSSFWorkbook(stream);
        if (workbook.NumberOfSheets < 1)
        {
            throw new InvalidOperationException($"LibreOffice/soffice produced an XLSX without worksheets: {output}");
        }
    }

    private static void ConvertWithSoffice(string input, string output, string targetFormat, string? sofficePath)
    {
        var soffice = ResolveSofficeBinary(sofficePath);
        var outputDir = Path.GetDirectoryName(Path.GetFullPath(output));
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        var tempRoot = Path.Combine(Path.GetTempPath(), $"tiwater-convert-{targetFormat}-{Guid.NewGuid():N}");
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
            startInfo.ArgumentList.Add(targetFormat);
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

                throw new TimeoutException($"LibreOffice/soffice {targetFormat} conversion timed out after 120 seconds.");
            }

            var stdout = stdoutTask.GetAwaiter().GetResult();
            var stderr = stderrTask.GetAwaiter().GetResult();
            var generated = Path.Combine(exportDir, $"{Path.GetFileNameWithoutExtension(input)}.{targetFormat}");
            if (process.ExitCode != 0 || !File.Exists(generated))
            {
                var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static s => !string.IsNullOrWhiteSpace(s)));
                throw new InvalidOperationException(
                    $"LibreOffice/soffice failed to convert {input} to {targetFormat.ToUpperInvariant()}."
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
                "LibreOffice/soffice is required for this conversion. Install LibreOffice or set TIWATER_SOFFICE, SOFFICE, or LIBREOFFICE_PATH.");
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
