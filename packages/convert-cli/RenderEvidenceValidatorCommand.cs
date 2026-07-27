using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Dockit.Convert;

/// <summary>
/// Physically independent validator for the render contract. It recomputes every identity,
/// byte hash, and the full ordered page closure from the current authority bytes. Any verdict
/// or decision self-attested by the producer is ignored; a result carrying such claims is
/// rejected by the closed result shape.
/// </summary>
public static class RenderEvidenceValidatorCommand
{
    private const string VerdictSchemaId = "tiwater.convert-render-verdict/v1";
    private const string RequestSchemaId = "tiwater.convert-render-request/v1";
    private const string ResultSchemaId = "tiwater.convert-render-result/v1";
    private const string ProvenanceSchemaId = "tiwater.convert-native-render-provenance/v1";
    private const string ValidatorId = "tiwater-convert-render-validator";
    private const string ProviderId = "tiwater-convert";

    private static readonly string[] Formats = ["docx", "xlsx", "xls", "pptx", "pdf"];
    private static readonly string[] Backends = ["wps-writer", "wps-spreadsheet", "wps-presentation", "libreoffice", "passthrough"];

    private sealed record Finding(string Code, string Path);

    private sealed record RequestFacts(
        string RequestId,
        string Format,
        string Backend,
        string InputPath,
        string InputSha256,
        long InputSizeBytes,
        string OutputPath,
        JsonNode InputEcho,
        JsonNode OptionsEcho);

    public static int Run(string[] args)
    {
        var requestPath = Required(args, "--request");
        var resultPath = Required(args, "--result");
        var outputPath = Required(args, "--output");
        EnsureFresh(outputPath);
        var findings = new List<Finding>();

        var request = JsonNode.Parse(File.ReadAllText(requestPath)) as JsonObject
            ?? throw new InvalidOperationException("render request is not an object");
        var requestSha256 = Sha(Canonical(request));
        var facts = CheckRequest(request, findings);

        var result = JsonNode.Parse(File.ReadAllText(resultPath)) as JsonObject
            ?? throw new InvalidOperationException("render result is not an object");

        var inputSha256 = RecomputeFileSha(facts?.InputPath);
        var claimedOutputPath = result["output"] is JsonObject claimedOutput
            ? Text(claimedOutput["path"])
            : null;
        var outputPathForRecompute = facts?.OutputPath ?? claimedOutputPath;
        var outputSha256 = RecomputeFileSha(outputPathForRecompute);
        var outputSizeBytes = RecomputeFileSize(outputPathForRecompute);
        var recomputedPages = RecomputePages(outputPathForRecompute);

        CheckResult(facts, result, requestSha256, outputSha256, outputSizeBytes, recomputedPages, findings);

        var provenanceNode = result["native_render_provenance"];
        var verdict = new JsonObject
        {
            ["schema"] = VerdictSchemaId,
            ["request_id"] = facts?.RequestId ?? Text(request["request_id"]),
            ["request_sha256"] = requestSha256,
            ["result"] = new JsonObject
            {
                ["path"] = resultPath,
                ["sha256"] = Sha(File.ReadAllBytes(resultPath))
            },
            ["validator"] = new JsonObject
            {
                ["id"] = ValidatorId,
                ["version"] = RuntimeIdentity.Version,
                ["command"] = "validate-render-evidence"
            },
            ["recomputed"] = new JsonObject
            {
                ["input_sha256"] = inputSha256,
                ["output_sha256"] = outputSha256,
                ["page_count"] = recomputedPages?.Count,
                ["pages_sha256"] = recomputedPages is null ? null : Sha(Canonical(recomputedPages)),
                ["provenance_sha256"] = provenanceNode is null ? null : Sha(Canonical(provenanceNode))
            },
            ["decision"] = findings.Count == 0 ? "pass" : "failed",
            ["findings"] = new JsonArray(findings.Select(finding => (JsonNode)new JsonObject
            {
                ["code"] = finding.Code,
                ["path"] = finding.Path
            }).ToArray())
        };
        Write(outputPath, verdict);
        return findings.Count == 0 ? 0 : 1;
    }

    private static RequestFacts? CheckRequest(JsonObject request, List<Finding> findings)
    {
        var valid = HasKeys(request,
            ["schema", "request_id", "format", "input", "output", "renderer", "runtime", "target_format", "options", "result_contract"])
            && Text(request["schema"]) == RequestSchemaId
            && Text(request["request_id"]) is { Length: > 0 }
            && Formats.Contains(Text(request["format"]), StringComparer.Ordinal)
            && Text(request["target_format"]) == "pdf"
            && request["input"] is JsonObject
            && HasKeys(request["input"]!.AsObject(), ["path", "sha256", "size_bytes", "artifact_version_id"])
            && Text(request["input"]!["path"]) is { Length: > 0 }
            && IsSha(Text(request["input"]!["sha256"]))
            && Number(request["input"]!["size_bytes"]) is >= 1
            && Text(request["input"]!["artifact_version_id"]) is { Length: > 0 }
            && request["output"] is JsonObject
            && HasKeys(request["output"]!.AsObject(), ["path", "media_type"])
            && Text(request["output"]!["path"]) is { Length: > 0 }
            && Text(request["output"]!["media_type"]) == "application/pdf"
            && request["renderer"] is JsonObject
            && HasKeys(request["renderer"]!.AsObject(), ["backend"])
            && Backends.Contains(Text(request["renderer"]!["backend"]), StringComparer.Ordinal)
            && request["runtime"] is JsonObject
            && HasKeys(request["runtime"]!.AsObject(), ["id", "version"])
            && request["options"] is JsonObject optionsObj
            && optionsObj.Count == 0
            && request["result_contract"] is JsonObject
            && HasKeys(request["result_contract"]!.AsObject(), ["id", "sha256"])
            && Text(request["result_contract"]!["id"]) == ResultSchemaId
            && IsSha(Text(request["result_contract"]!["sha256"]));
        if (!valid)
        {
            findings.Add(new Finding("render-request-invalid", "/"));
            return null;
        }

        var format = Text(request["format"])!;
        var backend = Text(request["renderer"]!["backend"])!;
        var inputPath = Path.GetFullPath(Text(request["input"]!["path"])!);
        var inputSha256 = Text(request["input"]!["sha256"])!;
        var inputSize = Number(request["input"]!["size_bytes"])!.Value;
        var outputPath = Path.GetFullPath(Text(request["output"]!["path"])!);

        if (!File.Exists(inputPath)
            || new FileInfo(inputPath).Length != inputSize
            || Sha(File.ReadAllBytes(inputPath)) != inputSha256)
            findings.Add(new Finding("render-request-input-mismatch", "/input"));
        if (Text(request["runtime"]!["id"]) != ProviderId
            || Text(request["runtime"]!["version"]) != RuntimeIdentity.Version)
            findings.Add(new Finding("render-request-runtime-mismatch", "/runtime"));
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
            findings.Add(new Finding("render-request-backend-mismatch", "/renderer"));
        if (Text(request["result_contract"]!["sha256"]) != DeployedSchemaSha())
            findings.Add(new Finding("render-request-result-contract-mismatch", "/result_contract"));

        return new RequestFacts(
            Text(request["request_id"])!,
            format,
            backend,
            inputPath,
            inputSha256,
            inputSize,
            outputPath,
            request["input"]!.DeepClone(),
            request["options"]!.DeepClone());
    }

    private static bool IsStructurallyValid(JsonObject result)
    {
        return HasKeys(result,
            ["schema", "request_id", "request_sha256", "format", "input", "output", "renderer", "runtime", "options",
             "page_count", "pages", "native_render_provenance", "provenance_sha256", "producer"])
            && Text(result["schema"]) == ResultSchemaId
            && Text(result["request_id"]) is { Length: > 0 }
            && IsSha(Text(result["request_sha256"]))
            && Formats.Contains(Text(result["format"]), StringComparer.Ordinal)
            && result["input"] is JsonObject
            && HasKeys(result["input"]!.AsObject(), ["path", "sha256", "size_bytes", "artifact_version_id"])
            && result["output"] is JsonObject outputObj
            && HasKeys(outputObj, ["path", "sha256", "size_bytes", "media_type"])
            && Text(outputObj["path"]) is { Length: > 0 }
            && IsSha(Text(outputObj["sha256"]))
            && Number(outputObj["size_bytes"]) is >= 1
            && Text(outputObj["media_type"]) == "application/pdf"
            && result["renderer"] is JsonObject
            && HasKeys(result["renderer"]!.AsObject(), ["backend", "provider", "version"])
            && result["runtime"] is JsonObject
            && HasKeys(result["runtime"]!.AsObject(),
                ["id", "version", "os_description", "os_architecture", "process_architecture", "framework_description"])
            && result["options"] is JsonObject
            && Number(result["page_count"]) is >= 1
            && result["pages"] is JsonArray pagesArray
            && pagesArray.Count >= 1
            && pagesArray.All(item =>
                item is JsonObject pageObj
                && HasKeys(pageObj, ["page", "sha256", "size_bytes"])
                && Number(pageObj["page"]) is >= 1
                && IsSha(Text(pageObj["sha256"]))
                && Number(pageObj["size_bytes"]) is >= 1)
            && (result["provenance_sha256"] is null || IsSha(Text(result["provenance_sha256"])))
            && result["producer"] is JsonObject
            && HasKeys(result["producer"]!.AsObject(), ["id", "version", "command"]);
    }

    private static void CheckResult(
        RequestFacts? facts,
        JsonObject result,
        string requestSha256,
        string? outputSha256,
        long? outputSizeBytes,
        JsonArray? recomputedPages,
        List<Finding> findings)
    {
        if (!IsStructurallyValid(result))
        {
            findings.Add(new Finding("render-result-invalid", "/"));
            return;
        }
        if (facts is null) return;

        if (Text(result["request_id"]) != facts.RequestId)
            findings.Add(new Finding("render-request-id-mismatch", "/request_id"));
        if (Text(result["request_sha256"]) != requestSha256)
            findings.Add(new Finding("render-request-hash-mismatch", "/request_sha256"));
        if (Text(result["format"]) != facts.Format)
            findings.Add(new Finding("render-format-mismatch", "/format"));
        if (Canonical(result["input"]) != Canonical(facts.InputEcho))
            findings.Add(new Finding("render-input-echo-mismatch", "/input"));
        var renderer = result["renderer"]!.AsObject();
        if (Text(renderer["backend"]) != facts.Backend
            || Text(renderer["provider"]) != ProviderId
            || Text(renderer["version"]) != RuntimeIdentity.Version)
            findings.Add(new Finding("render-renderer-identity-mismatch", "/renderer"));
        var runtime = result["runtime"]!.AsObject();
        if (Text(runtime["id"]) != ProviderId
            || Text(runtime["version"]) != RuntimeIdentity.Version
            || Text(runtime["os_description"]) is not { Length: > 0 }
            || Text(runtime["os_architecture"]) is not { Length: > 0 }
            || Text(runtime["process_architecture"]) is not { Length: > 0 }
            || Text(runtime["framework_description"]) is not { Length: > 0 })
            findings.Add(new Finding("render-runtime-mismatch", "/runtime"));
        if (Canonical(result["options"]) != Canonical(facts.OptionsEcho))
            findings.Add(new Finding("render-options-mismatch", "/options"));

        var output = result["output"]!.AsObject();
        if (Path.GetFullPath(Text(output["path"])!) != facts.OutputPath)
            findings.Add(new Finding("render-output-path-mismatch", "/output/path"));
        if (!File.Exists(facts.OutputPath))
            findings.Add(new Finding("render-output-missing", "/output"));
        else if (Text(output["sha256"]) != outputSha256 || Number(output["size_bytes"]) != outputSizeBytes)
            findings.Add(new Finding("render-output-bytes-mismatch", "/output"));

        if (recomputedPages is null)
        {
            if (File.Exists(facts.OutputPath))
                findings.Add(new Finding("render-output-unreadable", "/output"));
        }
        else
        {
            if (Number(result["page_count"]) != recomputedPages.Count)
                findings.Add(new Finding("render-page-count-mismatch", "/page_count"));
            if (Canonical(result["pages"]) != Canonical(recomputedPages))
                findings.Add(new Finding("render-page-closure-mismatch", "/pages"));
        }

        CheckProvenance(facts, result, outputSha256, outputSizeBytes, recomputedPages, findings);

        var producer = result["producer"]!.AsObject();
        if (Text(producer["id"]) != "tiwater-convert-render-producer"
            || Text(producer["command"]) != "render"
            || Text(producer["version"]) != RuntimeIdentity.Version)
            findings.Add(new Finding("render-producer-identity-mismatch", "/producer"));
    }

    private static void CheckProvenance(
        RequestFacts facts,
        JsonObject result,
        string? outputSha256,
        long? outputSizeBytes,
        JsonArray? recomputedPages,
        List<Finding> findings)
    {
        var provenance = result["native_render_provenance"];
        if (!facts.Backend.StartsWith("wps-", StringComparison.Ordinal))
        {
            if (provenance is not null)
                findings.Add(new Finding("render-provenance-unexpected", "/native_render_provenance"));
            if (result["provenance_sha256"] is not null)
                findings.Add(new Finding("render-provenance-unexpected", "/provenance_sha256"));
            return;
        }
        if (provenance is not JsonObject native)
        {
            findings.Add(new Finding("render-provenance-missing", "/native_render_provenance"));
            return;
        }
        var bound = HasKeys(native, ["schema", "backend", "wps", "runtime", "fonts", "input", "output", "page_count"])
            && Text(native["schema"]) == ProvenanceSchemaId
            && Text(native["backend"]) == facts.Backend
            && native["wps"] is JsonObject wps
            && HasKeys(wps, ["package", "build_version", "executable_sha256"])
            && Text(wps["package"]) == "wps-office"
            && Text(wps["build_version"]) is { Length: > 0 }
            && IsSha(Text(wps["executable_sha256"]))
            && native["runtime"] is JsonObject nativeRuntime
            && HasKeys(nativeRuntime, ["os_description", "os_architecture", "process_architecture", "framework_description"])
            && Text(nativeRuntime["os_description"]) is { Length: > 0 }
            && Text(nativeRuntime["os_architecture"]) is { Length: > 0 }
            && Text(nativeRuntime["process_architecture"]) is { Length: > 0 }
            && Text(nativeRuntime["framework_description"]) is { Length: > 0 }
            && native["fonts"] is JsonObject fonts
            && HasKeys(fonts, ["source", "count", "sha256"])
            && Text(fonts["source"]) == "fontconfig-family-style-file-sha256"
            && Number(fonts["count"]) is >= 1
            && IsSha(Text(fonts["sha256"]))
            && native["input"] is JsonObject nativeInput
            && HasKeys(nativeInput, ["sha256", "size_bytes"])
            && Text(nativeInput["sha256"]) == facts.InputSha256
            && Number(nativeInput["size_bytes"]) == facts.InputSizeBytes
            && native["output"] is JsonObject nativeOutput
            && HasKeys(nativeOutput, ["sha256", "size_bytes"])
            && Text(nativeOutput["sha256"]) == outputSha256
            && Number(nativeOutput["size_bytes"]) == outputSizeBytes
            && Number(native["page_count"]) == recomputedPages?.Count;
        if (!bound)
            findings.Add(new Finding("render-provenance-binding-mismatch", "/native_render_provenance"));
        if (Text(result["provenance_sha256"]) != Sha(Canonical(native)))
            findings.Add(new Finding("render-provenance-hash-mismatch", "/provenance_sha256"));
    }

    private static JsonArray? RecomputePages(string? outputPath)
    {
        if (outputPath is null || !File.Exists(outputPath)) return null;
        try
        {
            using var pdf = PdfDocument.Open(outputPath);
            var pages = new JsonArray();
            foreach (var page in pdf.GetPages())
            {
                var material = PageFingerprint(page);
                pages.Add(new JsonObject
                {
                    ["page"] = page.Number,
                    ["sha256"] = Sha(Encoding.UTF8.GetBytes(material)),
                    ["size_bytes"] = Encoding.UTF8.GetByteCount(material)
                });
            }
            return pages.Count == 0 ? null : pages;
        }
        catch
        {
            return null;
        }
    }

    private static string PageFingerprint(Page page)
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

    private static string? RecomputeFileSha(string? path)
    {
        if (path is null || !File.Exists(path)) return null;
        return Sha(File.ReadAllBytes(path));
    }

    private static long? RecomputeFileSize(string? path)
    {
        if (path is null || !File.Exists(path)) return null;
        return new FileInfo(path).Length;
    }

    private static string DeployedSchemaSha()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "schemas", "render-result-v1.schema.json");
        if (!File.Exists(path)) throw new InvalidOperationException("render result contract schema missing");
        return Sha(File.ReadAllBytes(path));
    }

    private static bool HasKeys(JsonObject value, IReadOnlyCollection<string> expected) =>
        value.Count == expected.Count && expected.All(value.ContainsKey);

    private static string? Text(JsonNode? node) =>
        node is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;

    private static long? Number(JsonNode? node) =>
        node is JsonValue value && value.TryGetValue<long>(out var number) ? number : null;

    private static bool IsSha(string? value) =>
        value is { Length: 64 } && value.All(character =>
            character is (>= '0' and <= '9') or (>= 'a' and <= 'f'));

    private static string Required(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        if (index < 0 || index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]))
            throw new InvalidOperationException($"{name} is required");
        return Path.GetFullPath(args[index + 1]);
    }

    private static void EnsureFresh(string path)
    {
        if (File.Exists(path)) throw new InvalidOperationException("render verdict output must be fresh");
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
}
