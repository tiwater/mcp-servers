using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
using UglyToad.PdfPig;

namespace Dockit.Convert;

public sealed record NativeRenderFileIdentity(
    [property: JsonPropertyName("sha256")] string Sha256,
    [property: JsonPropertyName("size_bytes")] long SizeBytes);

public sealed record NativeRenderWpsIdentity(
    [property: JsonPropertyName("package")] string Package,
    [property: JsonPropertyName("build_version")] string BuildVersion,
    [property: JsonPropertyName("executable_sha256")] string ExecutableSha256);

public sealed record NativeRenderRuntimeIdentity(
    [property: JsonPropertyName("os_description")] string OsDescription,
    [property: JsonPropertyName("os_architecture")] string OsArchitecture,
    [property: JsonPropertyName("process_architecture")] string ProcessArchitecture,
    [property: JsonPropertyName("framework_description")] string FrameworkDescription);

public sealed record NativeRenderFontInventory(
    [property: JsonPropertyName("source")] string Source,
    [property: JsonPropertyName("count")] int Count,
    [property: JsonPropertyName("sha256")] string Sha256);

public sealed record NativeRenderProvenance(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("backend")] string Backend,
    [property: JsonPropertyName("wps")] NativeRenderWpsIdentity Wps,
    [property: JsonPropertyName("runtime")] NativeRenderRuntimeIdentity Runtime,
    [property: JsonPropertyName("fonts")] NativeRenderFontInventory Fonts,
    [property: JsonPropertyName("input")] NativeRenderFileIdentity Input,
    [property: JsonPropertyName("output")] NativeRenderFileIdentity Output,
    [property: JsonPropertyName("page_count")] int PageCount);

internal static class NativeRenderProvenanceCollector
{
    internal static NativeRenderProvenance Capture(string input, string output, string backend)
    {
        if (backend is not ("wps-writer" or "wps-spreadsheet" or "wps-presentation"))
            throw new InvalidOperationException($"Native WPS provenance does not support backend: {backend}");

        var buildVersion = Run("dpkg-query", ["-W", "-f=${Version}", "wps-office"], "native-render-wps-build-unavailable");
        var executable = ResolveWpsExecutable(backend);
        var fontLines = Run("fc-list", ["--format=%{family[0]}\t%{style[0]}\t%{file}\n"], "native-render-font-inventory-unavailable")
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var normalizedFonts = fontLines.Select(NormalizeFont).Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        if (normalizedFonts.Length == 0) throw new InvalidOperationException("native-render-font-inventory-empty");

        int pageCount;
        try
        {
            using var pdf = PdfDocument.Open(output);
            pageCount = pdf.NumberOfPages;
        }
        catch (Exception error)
        {
            throw new InvalidOperationException($"native-render-page-count-unavailable:{error.Message}", error);
        }
        if (pageCount < 1) throw new InvalidOperationException("native-render-page-count-empty");

        return new NativeRenderProvenance(
            "tiwater.convert-native-render-provenance/v1",
            backend,
            new NativeRenderWpsIdentity("wps-office", buildVersion, FileSha256(executable)),
            new NativeRenderRuntimeIdentity(
                RuntimeInformation.OSDescription,
                RuntimeInformation.OSArchitecture.ToString().ToLowerInvariant(),
                RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
                RuntimeInformation.FrameworkDescription),
            new NativeRenderFontInventory("fontconfig-family-style-file-sha256", normalizedFonts.Length, Sha256(string.Join('\n', normalizedFonts))),
            FileIdentity(input),
            FileIdentity(output),
            pageCount);
    }

    internal static void Validate(NativeRenderProvenance provenance, string input, string output, string backend)
    {
        if (provenance.Schema != "tiwater.convert-native-render-provenance/v1"
            || provenance.Backend != backend
            || string.IsNullOrWhiteSpace(provenance.Wps.BuildVersion)
            || provenance.Wps.ExecutableSha256.Length != 64
            || string.IsNullOrWhiteSpace(provenance.Runtime.OsDescription)
            || string.IsNullOrWhiteSpace(provenance.Runtime.FrameworkDescription)
            || provenance.Fonts.Count < 1
            || provenance.Fonts.Sha256.Length != 64
            || provenance.Input != FileIdentity(input)
            || provenance.Output != FileIdentity(output)
            || provenance.PageCount < 1)
            throw new InvalidOperationException("native-render-provenance-invalid");
    }

    private static string NormalizeFont(string value)
    {
        var parts = value.Split('\t');
        if (parts.Length != 3 || string.IsNullOrWhiteSpace(parts[0]) || string.IsNullOrWhiteSpace(parts[2]) || !File.Exists(parts[2]))
            throw new InvalidOperationException("native-render-font-inventory-entry-invalid");
        return $"{parts[0].Trim()}\t{parts[1].Trim()}\t{FileSha256(parts[2])}";
    }

    private static string ResolveWpsExecutable(string backend)
    {
        var command = backend switch { "wps-writer" => "wps", "wps-spreadsheet" => "et", _ => "wpp" };
        var installed = Path.Combine("/opt/kingsoft/wps-office/office6", command);
        if (File.Exists(installed)) return installed;
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty).Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            var candidate = Path.Combine(directory, command);
            if (File.Exists(candidate)) return Path.GetFullPath(candidate);
        }
        throw new InvalidOperationException($"native-render-wps-executable-unavailable:{command}");
    }

    private static string Run(string command, IReadOnlyList<string> arguments, string failure)
    {
        var startInfo = new ProcessStartInfo { FileName = command, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
        foreach (var argument in arguments) startInfo.ArgumentList.Add(argument);
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException(failure);
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        if (!process.WaitForExit(TimeSpan.FromSeconds(30)))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new InvalidOperationException($"{failure}:timeout");
        }
        var value = stdout.Trim();
        if (process.ExitCode != 0 || string.IsNullOrWhiteSpace(value)) throw new InvalidOperationException($"{failure}:{stderr.Trim()}");
        return value;
    }

    private static NativeRenderFileIdentity FileIdentity(string path) => new(FileSha256(path), new FileInfo(path).Length);
    private static string FileSha256(string path) { using var stream = File.OpenRead(path); return System.Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(); }
    private static string Sha256(string value) => System.Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
