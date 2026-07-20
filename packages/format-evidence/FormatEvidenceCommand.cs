using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Tiwater.FormatEvidence;

public static class FormatEvidenceCommand
{
    public sealed record AdditionalObservation(string ObservationId, string SemanticField, string Use, object Value, string Pointer);
    public static int RunProducer(string[] args, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations = null)
        => Run(args, false, tool, version, format, inspect, additionalObservations);

    public static int RunValidator(string[] args, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations = null)
        => Run(args, true, tool, version, format, inspect, additionalObservations);

    private static int Run(string[] args, bool validator, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations)
    {
        var values = ParseArgs(args, validator);
        var request = JsonNode.Parse(File.ReadAllText(values["request"]))!.AsObject();
        var output = Path.GetFullPath(values["output"]);
        try
        {
            ValidateRequest(request, output, format, validator);
            var expected = BuildEvidence(request, tool, version, format, inspect, additionalObservations);
            JsonObject result;
            if (!validator) result = expected;
            else
            {
                var evidence = JsonNode.Parse(File.ReadAllText(values["evidence"]))!.AsObject();
                var pass = Canonical(evidence) == Canonical(expected);
                result = new JsonObject
                {
                    ["schema"] = "lucid.published-format-evidence-verdict/v1",
                    ["requestId"] = request["requestId"]!.GetValue<string>(),
                    ["subject"] = request["subject"]!.DeepClone(),
                    ["artifactVersionId"] = request["artifact"]!["artifactVersionId"]!.GetValue<string>(),
                    ["epochId"] = expected["epoch"]!["epochId"]!.GetValue<string>(),
                    ["evidenceRef"] = new JsonObject { ["evidenceId"] = evidence["evidenceId"]?.GetValue<string>() ?? "missing", ["sha256"] = Sha(Canonical(evidence)) },
                    ["validator"] = new JsonObject { ["tool"] = tool, ["toolVersion"] = version, ["capabilityId"] = "validate-inspect-evidence", ["capabilityVersion"] = "1" },
                    ["recomputedSemanticHash"] = Sha(Canonical(new JsonObject { ["entities"] = expected["entities"]!.DeepClone(), ["observations"] = expected["observations"]!.DeepClone() })),
                    ["pass"] = pass,
                    ["findings"] = pass ? new JsonArray() : new JsonArray(new JsonObject { ["code"] = "inspect-evidence-recomputation-mismatch", ["owner"] = "validator" })
                };
            }
            AtomicWrite(output, result);
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error.Message);
            var artifact = request["artifact"] as JsonObject;
            var result = new JsonObject
            {
                ["schema"] = "tiwater.format-evidence-error/v1",
                ["requestId"] = request["requestId"]?.GetValue<string>() ?? "unknown",
                ["subject"] = request["subject"]?.DeepClone() ?? new JsonObject { ["kind"] = "input", ["inputId"] = "unknown" },
                ["artifactVersionId"] = artifact?["artifactVersionId"]?.GetValue<string>() ?? "unknown",
                ["code"] = "inspect-evidence-invalid",
                ["category"] = "evidence",
                ["retryable"] = false,
                ["provider"] = new JsonObject { ["tool"] = tool, ["toolVersion"] = version, ["capabilityId"] = validator ? "validate-inspect-evidence" : "inspect-evidence" },
                ["refs"] = new JsonArray(),
                ["message"] = error.Message
            };
            result.Remove("message"); // Public error contract is closed and intentionally excludes free-form diagnostics.
            AtomicWrite(output, result);
            return 0;
        }
    }

    private static Dictionary<string, string> ParseArgs(string[] args, bool validator)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < args.Length; index += 2)
        {
            if (index + 1 >= args.Length || !args[index].StartsWith("--", StringComparison.Ordinal)) throw new InvalidOperationException("invalid arguments");
            values[args[index][2..]] = args[index + 1];
        }
        foreach (var name in validator ? new[] { "request", "evidence", "output" } : new[] { "request", "output" }) if (!values.ContainsKey(name)) throw new InvalidOperationException($"missing --{name}");
        return values;
    }

    private static void ValidateRequest(JsonObject request, string output, string format, bool validator)
    {
        var required = new[] { "schema", "requestId", "runId", "subject", "artifact", "extraction", "expectedEvidenceSchema", "outputPath" };
        if (request.Count != required.Length || required.Any(name => !request.ContainsKey(name)) || request["schema"]!.GetValue<string>() != "tiwater.format-evidence-request/v1" || request["expectedEvidenceSchema"]!.GetValue<string>() != "lucid.published-format-evidence/v1") throw new InvalidOperationException("request contract invalid");
        if (!validator && Path.GetFullPath(request["outputPath"]!.GetValue<string>()) != output) throw new InvalidOperationException("output path mismatch");
        var artifact = request["artifact"]!.AsObject(); var file = artifact["path"]!.GetValue<string>();
        if (!Path.IsPathFullyQualified(file) || artifact["format"]!.GetValue<string>() != format || FileSha(file) != artifact["bytesSha256"]!.GetValue<string>()) throw new InvalidOperationException("artifact authority mismatch");
        var extraction = request["extraction"]!.AsObject(); if (Sha(Canonical(extraction["options"]!)) != extraction["optionsSha256"]!.GetValue<string>()) throw new InvalidOperationException("extraction options mismatch");
    }

    private static JsonObject BuildEvidence(JsonObject request, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations)
    {
        var artifact = request["artifact"]!.AsObject(); var extraction = request["extraction"]!.AsObject(); var artifactPath = artifact["path"]!.GetValue<string>(); var inspection = JsonSerializer.SerializeToNode(inspect(artifactPath))!;
        var entity = new JsonObject { ["entityId"] = "document-1", ["kind"] = $"{format}-document", ["provenance"] = new JsonObject { ["source"] = "runtime", ["pointer"] = "/inspection" } };
        var observation = new JsonObject { ["observationId"] = "inspection-1", ["entityId"] = "document-1", ["semanticField"] = $"{format}.inspection", ["use"] = "structure", ["value"] = inspection, ["parentObservationIds"] = new JsonArray(), ["provenance"] = new JsonObject { ["source"] = "runtime", ["pointer"] = "/inspection" } };
        var observations = new JsonArray(observation);
        foreach (var item in additionalObservations?.Invoke(artifactPath) ?? []) observations.Add(new JsonObject { ["observationId"] = item.ObservationId, ["entityId"] = "document-1", ["semanticField"] = item.SemanticField, ["use"] = item.Use, ["value"] = JsonSerializer.SerializeToNode(item.Value), ["parentObservationIds"] = new JsonArray("inspection-1"), ["provenance"] = new JsonObject { ["source"] = "runtime", ["pointer"] = item.Pointer } });
        var epochMaterial = new JsonObject { ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(), ["runtimeTool"] = tool, ["runtimeSchema"] = extraction["schema"]!.GetValue<string>(), ["runtimeVersion"] = version, ["extractionOptions"] = extraction["options"]!.DeepClone() };
        var evidence = new JsonObject { ["schema"] = "lucid.published-format-evidence/v1", ["requestId"] = request["requestId"]!.GetValue<string>(), ["subject"] = request["subject"]!.DeepClone(), ["artifactVersionId"] = artifact["artifactVersionId"]!.GetValue<string>(), ["provider"] = new JsonObject { ["tool"] = tool, ["toolVersion"] = version, ["capabilityId"] = "inspect-evidence", ["capabilityVersion"] = "1", ["outputSchema"] = "lucid.published-format-evidence/v1" }, ["source"] = new JsonObject { ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(), ["format"] = format }, ["extraction"] = extraction.DeepClone(), ["epoch"] = new JsonObject { ["epochId"] = $"ep-{Sha(Canonical(epochMaterial))}", ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(), ["runtimeTool"] = tool, ["runtimeSchema"] = extraction["schema"]!.GetValue<string>(), ["runtimeVersion"] = version, ["extractionOptionsSha256"] = extraction["optionsSha256"]!.GetValue<string>() }, ["entities"] = new JsonArray(entity), ["observations"] = observations };
        evidence["evidenceId"] = $"evidence-{Sha(Canonical(evidence))}"; return evidence;
    }

    private static string Quote(string value)
    {
        var output = new StringBuilder(value.Length + 2).Append('"');
        for (var index = 0; index < value.Length; index++)
        {
            var current = value[index];
            switch (current)
            {
                case '"': output.Append("\\\""); break;
                case '\\': output.Append("\\\\"); break;
                case '\b': output.Append("\\b"); break;
                case '\f': output.Append("\\f"); break;
                case '\n': output.Append("\\n"); break;
                case '\r': output.Append("\\r"); break;
                case '\t': output.Append("\\t"); break;
                default:
                    if (current < 0x20 || (char.IsSurrogate(current) && !(char.IsHighSurrogate(current) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])))) output.Append($"\\u{(int)current:x4}");
                    else { output.Append(current); if (char.IsHighSurrogate(current)) output.Append(value[++index]); }
                    break;
            }
        }
        return output.Append('"').ToString();
    }

    private static string Canonical(JsonNode? node) => node switch
    {
        null => "null",
        JsonObject value => "{" + string.Join(",", value.OrderBy(item => item.Key, StringComparer.Ordinal).Select(item => Quote(item.Key) + ":" + Canonical(item.Value))) + "}",
        JsonArray value => "[" + string.Join(",", value.Select(Canonical)) + "]",
        JsonValue value when value.TryGetValue<string>(out var text) => Quote(text),
        _ => node.ToJsonString()
    };
    private static string Sha(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    private static string FileSha(string file) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file))).ToLowerInvariant();
    private static void AtomicWrite(string output, JsonNode value) { Directory.CreateDirectory(Path.GetDirectoryName(output)!); var temp = $"{output}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp"; using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { var bytes = Encoding.UTF8.GetBytes(Canonical(value) + "\n"); stream.Write(bytes); stream.Flush(true); } File.Move(temp, output, false); }
}
