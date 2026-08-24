using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;

namespace Dockit.Convert;

internal static class LimaWpsPdfConverter
{
    private const string InstanceEnvironment = "TIWATER_WPS_OFFICE_LIMA_INSTANCE";
    private const string SharedRoot = "/tmp/tiwater-wps-render";

    internal static bool IsAvailable()
        => OperatingSystem.IsMacOS()
            && !string.IsNullOrWhiteSpace(InstanceName())
            && !string.IsNullOrWhiteSpace(FindOnPath("limactl"));

    internal static IDisposable AcquireOfficeHostLease(TimeSpan? timeout = null, string? lockPath = null)
        => WpsRpcSession.AcquireOfficeLease(timeout, lockPath);

    internal static IDisposable AcquireEtHostLease(TimeSpan? timeout = null, string? lockPath = null)
        => AcquireOfficeHostLease(timeout, lockPath);

    internal static NativeRenderProvenance ConvertToPdf(string input, string output)
        => ConvertToPdf(input, output, "wps");

    internal static NativeRenderProvenance ConvertSpreadsheetToPdf(string input, string output)
        => ConvertToPdf(input, output, "et");

    internal static NativeRenderProvenance ConvertPresentationToPdf(string input, string output)
        => ConvertToPdf(input, output, "wpp");

    internal static void ConvertSpreadsheetToXlsx(string input, string output)
        => SaveSpreadsheetAsXlsx(input, output, requireLegacyInput: true);

    internal static void RecalculateXlsx(string input, string output)
        => SaveSpreadsheetAsXlsx(input, output, requireLegacyInput: false);

    internal static void RefreshDocxFields(string input, string output)
    {
        var instance = InstanceName()
            ?? throw new InvalidOperationException($"{InstanceEnvironment} is required for the Lima WPS backend.");
        var limactl = FindOnPath("limactl")
            ?? throw new InvalidOperationException("limactl is required for the Lima WPS backend.");
        var staging = Path.Combine(SharedRoot, $"tiwater-convert-wps-refresh-{Guid.NewGuid():N}");
        var stagedInput = Path.Combine(staging, "input.docx");
        var stagedOutput = Path.Combine(staging, "output.docx");
        Directory.CreateDirectory(staging);
        File.Copy(input, stagedInput, overwrite: false);

        try
        {
            using (AcquireOfficeHostLease())
            {
                RunDocumentFieldRefresh(limactl, instance, stagedInput, stagedOutput);
            }
            ValidateDocx(stagedOutput);
            var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
            if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
            File.Copy(stagedOutput, output, overwrite: true);
        }
        finally
        {
            try { Directory.Delete(staging, recursive: true); } catch { }
        }
    }

    private static void SaveSpreadsheetAsXlsx(string input, string output, bool requireLegacyInput)
    {
        var instance = InstanceName()
            ?? throw new InvalidOperationException($"{InstanceEnvironment} is required for the Lima ET backend.");
        var limactl = FindOnPath("limactl")
            ?? throw new InvalidOperationException("limactl is required for the Lima ET backend.");
        var expectedExtension = requireLegacyInput ? ".xls" : ".xlsx";
        if (!string.Equals(Path.GetExtension(input), expectedExtension, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Lima ET input must be an {expectedExtension.TrimStart('.').ToUpperInvariant()} file.");

        var staging = Path.Combine(SharedRoot, $"tiwater-convert-et-{Guid.NewGuid():N}");
        var stagedInput = Path.Combine(staging, $"input{expectedExtension}");
        var stagedOutput = Path.Combine(staging, "output.xlsx");
        Directory.CreateDirectory(staging);
        File.Copy(input, stagedInput, overwrite: false);

        try
        {
            using (AcquireEtHostLease())
            {
                RunSpreadsheetConversion(limactl, instance, stagedInput, stagedOutput, requireLegacyInput ? "xls-to-xlsx" : "recalculate-xlsx");
            }
            ValidateXlsx(stagedOutput);
            var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
            if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
            File.Copy(stagedOutput, output, overwrite: true);
        }
        finally
        {
            try { Directory.Delete(staging, recursive: true); } catch { }
        }
    }

    private static NativeRenderProvenance ConvertToPdf(string input, string output, string backend)
    {
        var instance = InstanceName()
            ?? throw new InvalidOperationException($"{InstanceEnvironment} is required for the Lima WPS PDF backend.");
        var limactl = FindOnPath("limactl")
            ?? throw new InvalidOperationException("limactl is required for the Lima WPS PDF backend.");
        var extension = Path.GetExtension(input);
        if (string.IsNullOrWhiteSpace(extension)) throw new InvalidOperationException("Lima WPS PDF input must have an extension.");

        var staging = Path.Combine(SharedRoot, $"tiwater-convert-{backend}-{Guid.NewGuid():N}");
        var stagedInput = Path.Combine(staging, $"input{extension}");
        var stagedOutput = Path.Combine(staging, "output.pdf");
        Directory.CreateDirectory(staging);
        File.Copy(input, stagedInput, overwrite: false);

        try
        {
            NativeRenderProvenance provenance;
            using (AcquireOfficeHostLease())
            {
                provenance = Run(limactl, instance, stagedInput, stagedOutput, backend);
            }
            if (!IsPdf(stagedOutput)) throw new InvalidOperationException($"Lima {backend} did not produce a valid PDF.");
            var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(output));
            if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
            File.Copy(stagedOutput, output, overwrite: true);
            NativeRenderProvenanceCollector.Validate(provenance, input, output, backend);
            return provenance;
        }
        finally
        {
            try { Directory.Delete(staging, recursive: true); } catch { }
        }
    }

    private static NativeRenderProvenance Run(string limactl, string instance, string input, string output, string backend)
    {
        var startInfo = CreateProcessStartInfo(limactl, instance, input, output, backend);
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"Failed to start Lima {backend} PDF conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(3)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"Lima {backend} PDF conversion timed out after 180 seconds.");
        }

        var outputText = WpsRpcSession.CollectProcessOutput(
            stdoutTask, stderrTask, TimeSpan.FromSeconds(1));
        var stdout = outputText.Stdout;
        var stderr = outputText.Stderr;
        if (process.ExitCode != 0)
        {
            var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static value => !string.IsNullOrWhiteSpace(value)));
            throw new InvalidOperationException($"Lima {backend} PDF conversion failed." + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
        try
        {
            using var document = JsonDocument.Parse(stdout);
            var provenance = document.RootElement.GetProperty("native_render_provenance")
                .Deserialize<NativeRenderProvenance>() ?? throw new InvalidOperationException();
            return provenance;
        }
        catch (Exception error)
        {
            throw new InvalidOperationException($"Lima {backend} native render provenance is missing or invalid.", error);
        }
    }

    internal static ProcessStartInfo CreateProcessStartInfo(string limactl, string instance, string input, string output)
        => CreateProcessStartInfo(limactl, instance, input, output, "wps");

    internal static ProcessStartInfo CreateSpreadsheetConversionStartInfo(string limactl, string instance, string input, string output, string command = "xls-to-xlsx")
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = limactl,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        foreach (var argument in new[] { "shell", instance, "--", "bash", "-lc", SpreadsheetConversionCommand(input, output, command) }) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }

    internal static ProcessStartInfo CreateDocumentFieldRefreshStartInfo(string limactl, string instance, string input, string output)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = limactl,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        foreach (var argument in new[] { "shell", instance, "--", "bash", "-lc", DocumentFieldRefreshCommand(input, output) })
            startInfo.ArgumentList.Add(argument);
        return startInfo;
    }

    private static ProcessStartInfo CreateProcessStartInfo(string limactl, string instance, string input, string output, string backend)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = limactl,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false,
        };
        foreach (var argument in new[] { "shell", instance, "--", "bash", "-lc", RemoteCommand(input, output, backend) }) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }

    private static string RemoteCommand(string input, string output, string backend)
        => $"set -e; export DOTNET_ROOT=\"$HOME/.dotnet\"; export PATH=\"$HOME/.dotnet:$HOME/.dotnet/tools:$HOME/.local/bin:$PATH\"; export TIWATER_WPSRPC_PYTHON=\"$HOME/.local/share/tiwater/wpsrpc-venv/bin/python\"; export TIWATER_OFFICE_PDF_BACKEND={backend}; tiwater-convert {SourceFormat(input, backend)}-to-pdf '{input}' '{output}'";

    private static string SpreadsheetConversionCommand(string input, string output, string command)
        => $"set -e; export DOTNET_ROOT=\"$HOME/.dotnet\"; export PATH=\"$HOME/.dotnet:$HOME/.dotnet/tools:$HOME/.local/bin:$PATH\"; export TIWATER_WPSRPC_PYTHON=\"$HOME/.local/share/tiwater/wpsrpc-venv/bin/python\"; export TIWATER_OFFICE_XLSX_BACKEND=et; tiwater-convert {command} '{input}' '{output}'";

    private static string DocumentFieldRefreshCommand(string input, string output)
        => $"set -e; export DOTNET_ROOT=\"$HOME/.dotnet\"; export PATH=\"$HOME/.dotnet:$HOME/.dotnet/tools:$HOME/.local/bin:$PATH\"; export TIWATER_WPSRPC_PYTHON=\"$HOME/.local/share/tiwater/wpsrpc-venv/bin/python\"; tiwater-convert refresh-docx-fields '{input}' '{output}'";

    private static void RunDocumentFieldRefresh(string limactl, string instance, string input, string output)
    {
        using var process = Process.Start(CreateDocumentFieldRefreshStartInfo(limactl, instance, input, output))
            ?? throw new InvalidOperationException("Failed to start Lima WPS document field refresh.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(4)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException("Lima WPS document field refresh timed out after 240 seconds.");
        }
        var outputText = WpsRpcSession.CollectProcessOutput(stdoutTask, stderrTask, TimeSpan.FromSeconds(1));
        if (process.ExitCode != 0)
        {
            var details = string.Join(" ", new[] { outputText.Stdout, outputText.Stderr }.Where(static value => !string.IsNullOrWhiteSpace(value)));
            throw new InvalidOperationException("Lima WPS document field refresh failed." +
                (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
        ValidateDocumentFieldRefreshEvidence(outputText.Stdout, input, output);
    }

    internal static void ValidateDocumentFieldRefreshEvidence(string stdout, string input, string output)
    {
        using var document = JsonDocument.Parse(stdout);
        var root = document.RootElement;
        if (root.GetProperty("schema").GetString() != "tiwater.convert-refresh-docx-fields/v1"
            || root.GetProperty("status").GetString() != "ok"
            || root.GetProperty("backend").GetString() != "wps"
            || root.GetProperty("source_format").GetString() != "docx"
            || root.GetProperty("target_format").GetString() != "docx"
            || !root.GetProperty("refresh_scope").EnumerateArray().Select(static item => item.GetString())
                .SequenceEqual(new[] { "table-of-contents", "table-of-figures" })
            || root.GetProperty("input_sha256").GetString() != FileSha256(input)
            || root.GetProperty("output_sha256").GetString() != FileSha256(output))
            throw new InvalidOperationException("Document field refresh evidence does not attest the staged input and output bytes.");
    }

    private static void RunSpreadsheetConversion(string limactl, string instance, string input, string output, string command)
    {
        using var process = Process.Start(CreateSpreadsheetConversionStartInfo(limactl, instance, input, output, command))
            ?? throw new InvalidOperationException("Failed to start Lima ET conversion.");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        if (!process.WaitForExit(TimeSpan.FromMinutes(10)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new TimeoutException($"Lima ET {command} timed out after 600 seconds.");
        }
        var outputText = WpsRpcSession.CollectProcessOutput(
            stdoutTask, stderrTask, TimeSpan.FromSeconds(1));
        var stdout = outputText.Stdout;
        var stderr = outputText.Stderr;
        if (process.ExitCode != 0)
        {
            var details = string.Join(" ", new[] { stdout.Trim(), stderr.Trim() }.Where(static value => !string.IsNullOrWhiteSpace(value)));
            throw new InvalidOperationException($"Lima ET {command} failed." + (string.IsNullOrWhiteSpace(details) ? string.Empty : $" {details}"));
        }
        try
        {
            ValidateSpreadsheetEvidence(stdout, command, input, output);
        }
        catch (Exception error)
        {
            throw new InvalidOperationException($"Lima ET {command} evidence is missing or invalid.", error);
        }
    }

    internal static void ValidateSpreadsheetEvidence(string stdout, string command, string input, string output)
    {
        using var document = JsonDocument.Parse(stdout);
        var root = document.RootElement;
        var expectedSourceFormat = command == "xls-to-xlsx" ? "xls" : "xlsx";
        if (root.GetProperty("status").GetString() != "ok"
            || root.GetProperty("backend").GetString() != "et"
            || root.GetProperty("fallback_reason").ValueKind != JsonValueKind.Null
            || root.GetProperty("source_format").GetString() != expectedSourceFormat
            || root.GetProperty("target_format").GetString() != "xlsx"
            || (command == "recalculate-xlsx" && (root.GetProperty("input_sha256").GetString() != FileSha256(input)
                || root.GetProperty("output_sha256").GetString() != FileSha256(output))))
            throw new InvalidOperationException("Spreadsheet evidence does not attest the staged input and output bytes.");
    }

    private static string FileSha256(string path)
    {
        using var stream = File.OpenRead(path);
        return System.Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static void ValidateXlsx(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 4) throw new InvalidOperationException("Lima ET did not produce an XLSX file.");
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        if (stream.Read(header) != 4 || !header.SequenceEqual("PK\u0003\u0004"u8)) throw new InvalidOperationException("Lima ET output is not an XLSX package.");
    }

    private static void ValidateDocx(string path)
    {
        if (!File.Exists(path) || new FileInfo(path).Length < 4)
            throw new InvalidOperationException("Lima WPS did not produce a DOCX file.");
        using var stream = File.OpenRead(path);
        Span<byte> header = stackalloc byte[4];
        if (stream.Read(header) != 4 || !header.SequenceEqual("PK\u0003\u0004"u8))
            throw new InvalidOperationException("Lima WPS output is not a DOCX package.");
    }

    private static string SourceFormat(string input, string backend)
    {
        var format = Path.GetExtension(input).TrimStart('.').ToLowerInvariant();
        var supported = backend == "wps"
            ? format is "doc" or "docx" or "odt" or "rtf"
            : backend == "et"
                ? format is "xls" or "xlsx"
                : backend == "wpp" && format is "ppt" or "pptx" or "odp";
        return supported ? format : throw new InvalidOperationException($"Unsupported Lima {backend} PDF source format: {format}");
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
