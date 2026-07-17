using System.Diagnostics;

namespace Dockit.Convert;

internal static class LimaWpsWriterPdfConverter
{
    private const string InstanceEnvironment = "TIWATER_WPS_WRITER_LIMA_INSTANCE";
    private const string SharedRoot = "/tmp/lucid-wps-render";

    internal static bool IsAvailable()
        => OperatingSystem.IsMacOS()
            && !string.IsNullOrWhiteSpace(InstanceName())
            && !string.IsNullOrWhiteSpace(FindOnPath("limactl"));

    internal static void ConvertToPdf(string input, string output)
    {
        var instance = InstanceName()
            ?? throw new InvalidOperationException($"{InstanceEnvironment} is required for the Lima WPS Writer PDF backend.");
        var limactl = FindOnPath("limactl")
            ?? throw new InvalidOperationException("limactl is required for the Lima WPS Writer PDF backend.");
        var extension = Path.GetExtension(input);
        if (string.IsNullOrWhiteSpace(extension)) throw new InvalidOperationException("WPS Writer PDF input must have an extension.");

        var staging = Path.Combine(SharedRoot, $"tiwater-convert-wps-writer-{Guid.NewGuid():N}");
        var stagedInput = Path.Combine(staging, $"input{extension}");
        var stagedOutput = Path.Combine(staging, "output.pdf");
        Directory.CreateDirectory(staging);
        File.Copy(input, stagedInput, overwrite: false);

        try
        {
            Run(limactl, instance, stagedInput, stagedOutput);
            if (!IsPdf(stagedOutput)) throw new InvalidOperationException("Lima WPS Writer did not produce a valid PDF.");
            var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
            if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
            File.Copy(stagedOutput, output, overwrite: true);
        }
        finally
        {
            try { Directory.Delete(staging, recursive: true); } catch { }
        }
    }

    private static void Run(string limactl, string instance, string input, string output)
    {
        var startInfo = CreateProcessStartInfo(limactl, instance, input, output);
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start Lima WPS Writer PDF conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(3)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("Lima WPS Writer PDF conversion timed out after 180 seconds.");
        }

        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static value => !string.IsNullOrWhiteSpace(value)));
            throw new InvalidOperationException("Lima WPS Writer PDF conversion failed." + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
    }

    internal static ProcessStartInfo CreateProcessStartInfo(string limactl, string instance, string input, string output)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = limactl,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        foreach (var argument in new[] { "shell", instance, "--", "bash", "-lc", RemoteCommand(input, output) }) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }

    private static string RemoteCommand(string input, string output)
        => $"set -e; export DOTNET_ROOT=\"$HOME/.dotnet\"; export PATH=\"$HOME/.dotnet:$HOME/.local/bin:$PATH\"; export TIWATER_WPSRPC_PYTHON=\"$HOME/.local/share/lucid-docs/wpsrpc-venv/bin/python\"; export TIWATER_OFFICE_PDF_BACKEND=wps-writer; tiwater-convert {SourceFormat(input)}-to-pdf '{input}' '{output}'";

    private static string SourceFormat(string input)
    {
        var format = Path.GetExtension(input).TrimStart('.').ToLowerInvariant();
        return format is "doc" or "docx" or "odt" or "rtf"
            ? format
            : throw new InvalidOperationException($"Unsupported Lima WPS Writer PDF source format: {format}");
    }

    private static string? InstanceName()
    {
        var value = Environment.GetEnvironmentVariable(InstanceEnvironment)?.Trim();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string? FindOnPath(string command)
    {
        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var directory in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, command);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        return null;
    }

    private static bool IsPdf(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 5) return false;
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        return stream.Read(header) == 4 && header.SequenceEqual("%PDF"u8);
    }
}
