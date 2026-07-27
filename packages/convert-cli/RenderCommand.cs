using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Dockit.Convert;

public sealed record RenderInvocation(string Input, string Output, string Format, string Backend);

public delegate OfficePdfConversionResult RenderBackend(RenderInvocation invocation);

/// <summary>
/// Producer for the closed tiwater.convert-render-request/v1 -> tiwater.convert-render-result/v1
/// contract. The renderer is injectable so the contract surface can be proven without a native
/// Office runtime; the production default routes through the existing Office/WPS render path.
/// </summary>
public static class RenderCommand
{
    public const string RequestSchemaId = "tiwater.convert-render-request/v1";
    public const string ResultSchemaId = "tiwater.convert-render-result/v1";
    public const string ResultContractFile = "render-result-v1.schema.json";
    public const string ProducerId = "tiwater-convert-render-producer";
    public const string ProviderId = "tiwater-convert";

    private static readonly string[] Formats = ["docx", "xlsx", "xls", "pptx", "pdf"];
    private static readonly string[] Backends = ["wps-writer", "wps-spreadsheet", "wps-presentation", "libreoffice", "passthrough"];

    private sealed record Admitted(
        string RequestId,
        string Format,
        string Backend,
        string InputPath,
        string InputSha256,
        long InputSizeBytes,
        string InputArtifactVersionId,
        string OutputPath,
        string RequestSha256);

    public static int Run(string[] args, RenderBackend? renderer = null)
    {
        var requestPath = Required(args, "--request");
        var outputPath = Required(args, "--output");
        EnsureFresh(outputPath);
        var request = JsonNode.Parse(File.ReadAllText(requestPath)) as JsonObject
            ?? throw new InvalidOperationException("render request is not an object");
        var admitted = AdmitRequest(request);
        if (File.Exists(admitted.OutputPath))
            throw new InvalidOperationException("render output artifact must be fresh");
        var conversion = (renderer ?? DefaultRenderer)(
            new RenderInvocation(admitted.InputPath, admitted.OutputPath, admitted.Format, admitted.Backend));
        if (conversion.Backend != admitted.Backend)
            throw new InvalidOperationException(
                $"render backend mismatch: request requires {admitted.Backend} but the renderer reported {conversion.Backend}");
        var provenance = conversion.NativeRenderProvenance;
        if (admitted.Backend.StartsWith("wps-", StringComparison.Ordinal))
        {
            if (provenance is null)
                throw new InvalidOperationException("native render provenance is required for a WPS backend");
            NativeRenderProvenanceCollector.Validate(provenance, admitted.InputPath, admitted.OutputPath, admitted.Backend);
        }
        else if (provenance is not null)
        {
            throw new InvalidOperationException("native render provenance is not allowed for a non-WPS backend");
        }

        var pages = ExtractPages(admitted.OutputPath);
        var provenanceNode = provenance is null ? null : JsonSerializer.SerializeToNode(provenance);
        var result = new JsonObject
        {
            ["schema"] = ResultSchemaId,
            ["request_id"] = admitted.RequestId,
            ["request_sha256"] = admitted.RequestSha256,
            ["format"] = admitted.Format,
            ["input"] = new JsonObject
            {
                ["path"] = admitted.InputPath,
                ["sha256"] = admitted.InputSha256,
                ["size_bytes"] = admitted.InputSizeBytes,
                ["artifact_version_id"] = admitted.InputArtifactVersionId
            },
            ["output"] = new JsonObject
            {
                ["path"] = admitted.OutputPath,
                ["sha256"] = FileSha(admitted.OutputPath),
                ["size_bytes"] = new FileInfo(admitted.OutputPath).Length,
                ["media_type"] = "application/pdf"
            },
            ["renderer"] = new JsonObject
            {
                ["backend"] = admitted.Backend,
                ["provider"] = ProviderId,
                ["version"] = RuntimeIdentity.Version
            },
            ["runtime"] = CurrentRuntime(),
            ["options"] = new JsonObject(),
            ["page_count"] = pages.Count,
            ["pages"] = new JsonArray(pages.Select(page => (JsonNode)new JsonObject
            {
                ["page"] = page.Page,
                ["sha256"] = page.Sha256,
                ["size_bytes"] = page.SizeBytes
            }).ToArray()),
            ["native_render_provenance"] = provenanceNode,
            ["provenance_sha256"] = provenanceNode is null ? null : Sha(Canonical(provenanceNode)),
            ["producer"] = new JsonObject
            {
                ["id"] = ProducerId,
                ["version"] = RuntimeIdentity.Version,
                ["command"] = "render"
            }
        };
        Write(outputPath, result);
        return 0;
    }

    private static Admitted AdmitRequest(JsonObject request)
    {
        ExactKeys(request,
            ["schema", "request_id", "format", "input", "output", "renderer", "runtime", "target_format", "options", "result_contract"],
            "render request");
        if (request["schema"]!.GetValue<string>() != RequestSchemaId)
            throw new InvalidOperationException("render request schema identity invalid");
        var requestId = request["request_id"]!.GetValue<string>();
        if (string.IsNullOrWhiteSpace(requestId))
            throw new InvalidOperationException("render request id invalid");
        var format = request["format"]!.GetValue<string>();
        if (!Formats.Contains(format, StringComparer.Ordinal))
            throw new InvalidOperationException($"render request format unsupported: {format}");
        if (request["target_format"]!.GetValue<string>() != "pdf")
            throw new InvalidOperationException("render request target format invalid");

        var input = request["input"]!.AsObject();
        ExactKeys(input, ["path", "sha256", "size_bytes", "artifact_version_id"], "render request input");
        var inputPath = input["path"]!.GetValue<string>();
        var inputSha256 = input["sha256"]!.GetValue<string>();
        var inputSize = input["size_bytes"]!.GetValue<long>();
        var artifactVersionId = input["artifact_version_id"]!.GetValue<string>();
        RequireSha(inputSha256, "render request input");
        if (!Path.IsPathFullyQualified(inputPath) || !File.Exists(inputPath))
            throw new InvalidOperationException("render request input artifact unavailable");
        if (string.IsNullOrWhiteSpace(artifactVersionId))
            throw new InvalidOperationException("render request input artifact version invalid");
        if (inputSize < 1 || new FileInfo(inputPath).Length != inputSize || FileSha(inputPath) != inputSha256)
            throw new InvalidOperationException("render request input bytes do not match the declared identity");

        var output = request["output"]!.AsObject();
        ExactKeys(output, ["path", "media_type"], "render request output");
        var outputPath = output["path"]!.GetValue<string>();
        if (!Path.IsPathFullyQualified(outputPath))
            throw new InvalidOperationException("render request output path invalid");
        if (output["media_type"]!.GetValue<string>() != "application/pdf")
            throw new InvalidOperationException("render request output media type invalid");

        var renderer = request["renderer"]!.AsObject();
        ExactKeys(renderer, ["backend"], "render request renderer");
        var backend = renderer["backend"]!.GetValue<string>();
        if (!Backends.Contains(backend, StringComparer.Ordinal))
            throw new InvalidOperationException($"render request backend unsupported: {backend}");
        var compatible = backend switch
        {
            "wps-writer" => format is "docx",
            "wps-spreadsheet" => format is "xlsx" or "xls",
            "wps-presentation" => format is "pptx",
            "libreoffice" => format is "docx" or "xlsx" or "xls" or "pptx",
            "passthrough" => format is "pdf",
            _ => false
        };
        if (!compatible)
            throw new InvalidOperationException($"render request backend {backend} does not support format {format}");

        var runtime = request["runtime"]!.AsObject();
        ExactKeys(runtime, ["id", "version"], "render request runtime");
        if (runtime["id"]!.GetValue<string>() != ProviderId
            || runtime["version"]!.GetValue<string>() != RuntimeIdentity.Version)
            throw new InvalidOperationException("render request runtime identity mismatch");

        if (request["options"] is not JsonObject options || options.Count != 0)
            throw new InvalidOperationException("render request options must be an empty object");

        var contract = request["result_contract"]!.AsObject();
        ExactKeys(contract, ["id", "sha256"], "render request result contract");
        if (contract["id"]!.GetValue<string>() != ResultSchemaId
            || contract["sha256"]!.GetValue<string>() != DeployedSchemaSha(ResultContractFile))
            throw new InvalidOperationException("render request result contract mismatch");

        return new Admitted(
            requestId,
            format,
            backend,
            Path.GetFullPath(inputPath),
            inputSha256,
            inputSize,
            artifactVersionId,
            Path.GetFullPath(outputPath),
            Sha(Canonical(request)));
    }

    private static OfficePdfConversionResult DefaultRenderer(RenderInvocation invocation)
    {
        if (invocation.Backend == "passthrough")
        {
            File.Copy(invocation.Input, invocation.Output, overwrite: false);
            return new OfficePdfConversionResult("passthrough");
        }

        var previous = Environment.GetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND");
        Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", invocation.Backend);
        try
        {
            return OfficePdfConverter.ConvertToPdf(invocation.Input, invocation.Output, invocation.Format);
        }
        finally
        {
            Environment.SetEnvironmentVariable("TIWATER_OFFICE_PDF_BACKEND", previous);
        }
    }

    internal sealed record RenderPageArtifact(int Page, string Sha256, long SizeBytes);

    private static List<RenderPageArtifact> ExtractPages(string path)
    {
        using var pdf = PdfDocument.Open(path);
        var pages = new List<RenderPageArtifact>();
        foreach (var page in pdf.GetPages())
        {
            var material = PageMaterial(page);
            pages.Add(new RenderPageArtifact(
                page.Number,
                Sha(Encoding.UTF8.GetBytes(material)),
                Encoding.UTF8.GetByteCount(material)));
        }
        if (pages.Count == 0)
            throw new InvalidOperationException("render output has no pages");
        return pages;
    }

    private static string PageMaterial(Page page)
    {
        var builder = new StringBuilder();
        builder.Append("w=").Append(page.Width.ToString("R", CultureInfo.InvariantCulture)).Append('\n');
        builder.Append("h=").Append(page.Height.ToString("R", CultureInfo.InvariantCulture)).Append('\n');
        builder.Append("rot=").Append(page.Rotation.Value.ToString(CultureInfo.InvariantCulture));
        foreach (var letter in page.Letters)
        {
            builder.Append('\n').Append("L|").Append(letter.Value)
                .Append('|').Append(letter.Location.X.ToString("R", CultureInfo.InvariantCulture))
                .Append('|').Append(letter.Location.Y.ToString("R", CultureInfo.InvariantCulture))
                .Append('|').Append(letter.Width.ToString("R", CultureInfo.InvariantCulture))
                .Append('|').Append(letter.BoundingBox.Height.ToString("R", CultureInfo.InvariantCulture))
                .Append('|').Append(letter.PointSize.ToString("R", CultureInfo.InvariantCulture))
                .Append('|').Append(letter.FontName ?? string.Empty);
        }
        foreach (var image in page.GetImages())
        {
            builder.Append('\n').Append("I|").Append(Sha(image.RawBytes));
        }
        return builder.ToString();
    }

    private static JsonObject CurrentRuntime() =>
        new()
        {
            ["id"] = ProviderId,
            ["version"] = RuntimeIdentity.Version,
            ["os_description"] = RuntimeInformation.OSDescription,
            ["os_architecture"] = RuntimeInformation.OSArchitecture.ToString().ToLowerInvariant(),
            ["process_architecture"] = RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
            ["framework_description"] = RuntimeInformation.FrameworkDescription
        };

    internal static string DeployedSchemaSha(string file)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "schemas", file);
        if (!File.Exists(path)) throw new InvalidOperationException($"render contract schema missing: {file}");
        return Sha(File.ReadAllBytes(path));
    }

    private static void ExactKeys(JsonObject value, IReadOnlyCollection<string> expected, string label)
    {
        if (value.Count != expected.Count || expected.Any(name => !value.ContainsKey(name)))
            throw new InvalidOperationException($"{label} fields invalid");
    }

    private static void RequireSha(string value, string label)
    {
        if (value.Length != 64 || value.Any(character =>
                character is not (>= '0' and <= '9') and not (>= 'a' and <= 'f')))
            throw new InvalidOperationException($"{label} hash invalid");
    }

    private static string Required(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        if (index < 0 || index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]))
            throw new InvalidOperationException($"{name} is required");
        return Path.GetFullPath(args[index + 1]);
    }

    private static void EnsureFresh(string path)
    {
        if (File.Exists(path)) throw new InvalidOperationException("render result output must be fresh");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
    }

    private static void Write(string path, JsonNode value) =>
        File.WriteAllText(path, $"{Canonical(value)}\n");

    private static string Canonical(JsonNode? node) => node switch
    {
        null => "null",
        JsonObject value => $"{{{string.Join(",", value.OrderBy(
            item => item.Key,
            StringComparer.Ordinal).Select(item =>
                $"{JsonSerializer.Serialize(item.Key)}:{Canonical(item.Value)}"))}}}",
        JsonArray value => $"[{string.Join(",", value.Select(Canonical))}]",
        _ => node.ToJsonString()
    };

    private static string Sha(ReadOnlySpan<byte> value) =>
        System.Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();

    private static string Sha(string value) => Sha(Encoding.UTF8.GetBytes(value));

    private static string FileSha(string path) =>
        Sha(File.ReadAllBytes(path));
}
