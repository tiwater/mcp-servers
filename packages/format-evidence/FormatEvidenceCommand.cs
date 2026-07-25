using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Tiwater.FormatEvidence;

public static class FormatEvidenceCommand
{
    public sealed record AdditionalObservation(string ObservationId, string SemanticField, string Use, object Value, string Pointer);
    public sealed record ErrorClassification(string Code, string Category, bool Retryable);
    public static int RunProducer(string[] args, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations = null, IReadOnlySet<string>? acceptedSourceFormats = null, Func<Exception, ErrorClassification?>? classifyError = null)
        => Run(args, false, tool, version, format, inspect, additionalObservations, acceptedSourceFormats, classifyError);

    public static int RunValidator(string[] args, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations = null, IReadOnlySet<string>? acceptedSourceFormats = null, Func<Exception, ErrorClassification?>? classifyError = null)
        => Run(args, true, tool, version, format, inspect, additionalObservations, acceptedSourceFormats, classifyError);

    public static int RunProducerV2(string[] args, string tool, string version, string format, Func<string, object> inspect, IReadOnlySet<string>? acceptedSourceFormats = null, Func<Exception, ErrorClassification?>? classifyError = null)
        => RunV2(args, false, tool, version, format, inspect, acceptedSourceFormats, classifyError);

    public static int RunValidatorV2(string[] args, string tool, string version, string format, Func<string, object> inspect, IReadOnlySet<string>? acceptedSourceFormats = null, Func<Exception, ErrorClassification?>? classifyError = null)
        => RunV2(args, true, tool, version, format, inspect, acceptedSourceFormats, classifyError);

    private static int Run(string[] args, bool validator, string tool, string version, string format, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations, IReadOnlySet<string>? acceptedSourceFormats, Func<Exception, ErrorClassification?>? classifyError)
    {
        var values = ParseArgs(args, validator);
        var request = JsonNode.Parse(File.ReadAllText(values["request"]))!.AsObject();
        var output = Path.GetFullPath(values["output"]);
        try
        {
            var sourceFormat = ValidateRequest(request, output, format, validator, acceptedSourceFormats);
            var expected = BuildEvidence(request, tool, version, format, sourceFormat, inspect, additionalObservations);
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
            var classification = classifyError?.Invoke(error)
                ?? new ErrorClassification("inspect-evidence-invalid", "evidence", false);
            var result = new JsonObject
            {
                ["schema"] = "tiwater.format-evidence-error/v1",
                ["requestId"] = request["requestId"]?.GetValue<string>() ?? "unknown",
                ["subject"] = request["subject"]?.DeepClone() ?? new JsonObject { ["kind"] = "input", ["inputId"] = "unknown" },
                ["artifactVersionId"] = artifact?["artifactVersionId"]?.GetValue<string>() ?? "unknown",
                ["code"] = classification.Code,
                ["category"] = classification.Category,
                ["retryable"] = classification.Retryable,
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

    private static int RunV2(string[] args, bool validator, string tool, string version, string format, Func<string, object> inspect, IReadOnlySet<string>? acceptedSourceFormats, Func<Exception, ErrorClassification?>? classifyError)
    {
        var values = ParseArgs(args, validator);
        var request = JsonNode.Parse(File.ReadAllText(values["request"]))!.AsObject();
        var output = Path.GetFullPath(values["output"]);
        try
        {
            var sourceFormat = ValidateRequestV2(request, tool, version, format, acceptedSourceFormats);
            var expected = BuildEvidenceV2(request, sourceFormat, inspect);
            JsonObject result;
            if (!validator) result = expected;
            else
            {
                var evidence = JsonNode.Parse(File.ReadAllText(values["evidence"]))!.AsObject();
                var pass = Canonical(evidence) == Canonical(expected);
                result = new JsonObject
                {
                    ["schema"] = "tiwater.format-evidence-verdict/v2",
                    ["requestId"] = request["requestId"]!.GetValue<string>(),
                    ["subject"] = request["subject"]!.DeepClone(),
                    ["artifactVersionId"] = request["artifact"]!["artifactVersionId"]!.GetValue<string>(),
                    ["evidence"] = new JsonObject
                    {
                        ["evidenceId"] = evidence["evidenceId"]?.GetValue<string>() ?? "missing",
                        ["sha256"] = Sha(Canonical(evidence))
                    },
                    ["validator"] = request["validator"]!.DeepClone(),
                    ["recomputedSourceBytesSha256"] = request["artifact"]!["bytesSha256"]!.GetValue<string>(),
                    ["recomputedObservationSha256"] = expected["observation"]!["sha256"]!.GetValue<string>(),
                    ["recomputedProvenanceSha256"] = Sha(Canonical(expected["provenance"])),
                    ["decision"] = pass ? "pass" : "failed",
                    ["findings"] = pass
                        ? new JsonArray()
                        : new JsonArray(new JsonObject
                        {
                            ["code"] = "format-evidence-recomputation-mismatch",
                            ["severity"] = "error"
                        })
                };
            }
            AtomicWrite(output, result);
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error.Message);
            var artifact = request["artifact"] as JsonObject;
            var classification = classifyError?.Invoke(error)
                ?? new ErrorClassification("format-evidence-v2-invalid", "evidence", false);
            var result = new JsonObject
            {
                ["schema"] = "tiwater.format-evidence-error/v1",
                ["requestId"] = request["requestId"]?.GetValue<string>() ?? "unknown",
                ["subject"] = request["subject"]?.DeepClone() ?? new JsonObject { ["kind"] = "input", ["inputId"] = "unknown" },
                ["artifactVersionId"] = artifact?["artifactVersionId"]?.GetValue<string>() ?? "unknown",
                ["code"] = classification.Code,
                ["category"] = classification.Category,
                ["retryable"] = classification.Retryable,
                ["provider"] = new JsonObject { ["tool"] = tool, ["toolVersion"] = version, ["capabilityId"] = validator ? "validate-inspect-evidence-v2" : "inspect-evidence-v2" },
                ["refs"] = new JsonArray()
            };
            AtomicWrite(output, result);
            return 0;
        }
    }

    private static string ValidateRequestV2(JsonObject request, string tool, string version, string format, IReadOnlySet<string>? acceptedSourceFormats)
    {
        var required = new[] { "schema", "requestId", "runId", "subject", "artifact", "provider", "validator", "runtime", "extraction", "expectedEvidenceContract" };
        if (request.Count != required.Length || required.Any(name => !request.ContainsKey(name)) || request["schema"]!.GetValue<string>() != "tiwater.format-evidence-request/v2")
            throw new InvalidOperationException("v2 request contract invalid");
        var expectedProvider = new JsonObject { ["id"] = tool, ["version"] = version };
        var expectedValidator = new JsonObject { ["id"] = $"{tool}-validator", ["version"] = version };
        if (Canonical(request["provider"]) != Canonical(expectedProvider) || Canonical(request["validator"]) != Canonical(expectedValidator) || Canonical(request["runtime"]) != Canonical(expectedProvider))
            throw new InvalidOperationException("v2 provider identity mismatch");
        var expectedEvidence = ContractRef("tiwater.format-evidence-v2.schema.json", "tiwater.format-evidence/v2");
        if (Canonical(request["expectedEvidenceContract"]) != Canonical(expectedEvidence))
            throw new InvalidOperationException("v2 evidence contract mismatch");
        var extraction = request["extraction"]!.AsObject();
        var expectedExtraction = ContractRef("tiwater.format-extraction-options-v1.schema.json", "tiwater.format-extraction-options/v1");
        if (Canonical(extraction["schema"]) != Canonical(expectedExtraction) || extraction["sha256"]!.GetValue<string>() != Sha(Canonical(extraction["value"])))
            throw new InvalidOperationException("v2 extraction authority mismatch");
        var extractionValue = extraction["value"]!.AsObject();
        if (extractionValue.Count != 1 || extractionValue["facets"] is not JsonArray facets || facets.Count != 1 || facets[0]!.GetValue<string>() != "format-summary")
            throw new InvalidOperationException("v2 extraction options invalid");
        var artifact = request["artifact"]!.AsObject();
        var file = artifact["path"]!.GetValue<string>();
        if (!Path.IsPathFullyQualified(file) || FileSha(file) != artifact["bytesSha256"]!.GetValue<string>())
            throw new InvalidOperationException("v2 artifact authority mismatch");
        var sourceFormat = Path.GetExtension(file).Equals(".xls", StringComparison.OrdinalIgnoreCase) ? "xls" : format;
        var allowedFormats = acceptedSourceFormats ?? new HashSet<string>(StringComparer.Ordinal) { format };
        if (!allowedFormats.Contains(sourceFormat)) throw new InvalidOperationException("v2 source format invalid");
        return sourceFormat;
    }

    private static JsonObject BuildEvidenceV2(JsonObject request, string sourceFormat, Func<string, object> inspect)
    {
        var artifact = request["artifact"]!.AsObject();
        var inspection = JsonSerializer.SerializeToNode(inspect(artifact["path"]!.GetValue<string>()))!;
        var inspectionSha256 = Sha(Canonical(inspection));
        var facets = new JsonArray();
        if (inspection is JsonObject objectInspection)
        {
            foreach (var item in objectInspection.OrderBy(item => item.Key, StringComparer.Ordinal))
                facets.Add(new JsonObject { ["facetId"] = item.Key, ["sha256"] = Sha(Canonical(item.Value)) });
        }
        else facets.Add(new JsonObject { ["facetId"] = "inspection", ["sha256"] = inspectionSha256 });
        var observationValue = new JsonObject
        {
            ["format"] = sourceFormat,
            ["inspectionSha256"] = inspectionSha256,
            ["facets"] = facets
        };
        var observation = TypedValue("tiwater.format-observation-summary-v1.schema.json", "tiwater.format-observation-summary/v1", observationValue);
        var provenanceValue = new JsonObject
        {
            ["kind"] = "provider-inspection",
            ["artifactVersionId"] = artifact["artifactVersionId"]!.GetValue<string>(),
            ["sourceBytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(),
            ["inspectionSha256"] = inspectionSha256,
            ["provider"] = request["provider"]!.DeepClone(),
            ["runtime"] = request["runtime"]!.DeepClone(),
            ["extractionSha256"] = request["extraction"]!["sha256"]!.GetValue<string>()
        };
        var provenance = new JsonArray(TypedValue("tiwater.format-provenance-v1.schema.json", "tiwater.format-provenance/v1", provenanceValue));
        var evidence = new JsonObject
        {
            ["schema"] = "tiwater.format-evidence/v2",
            ["requestId"] = request["requestId"]!.GetValue<string>(),
            ["subject"] = request["subject"]!.DeepClone(),
            ["artifactVersionId"] = artifact["artifactVersionId"]!.GetValue<string>(),
            ["source"] = new JsonObject
            {
                ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(),
                ["mediaType"] = artifact["mediaType"]!.GetValue<string>()
            },
            ["format"] = sourceFormat,
            ["provider"] = request["provider"]!.DeepClone(),
            ["runtime"] = request["runtime"]!.DeepClone(),
            ["extractionSha256"] = request["extraction"]!["sha256"]!.GetValue<string>(),
            ["observation"] = observation,
            ["provenance"] = provenance
        };
        evidence["evidenceId"] = $"evidence-{Sha(Canonical(evidence))}";
        return evidence;
    }

    private static JsonObject TypedValue(string file, string id, JsonNode value)
        => new()
        {
            ["schema"] = ContractRef(file, id),
            ["value"] = value,
            ["sha256"] = Sha(Canonical(value))
        };

    private static JsonObject ContractRef(string file, string id)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        if (!File.Exists(path)) throw new InvalidOperationException($"provider contract missing: {file}");
        return new JsonObject { ["id"] = id, ["sha256"] = FileSha(path) };
    }

    private static string ValidateRequest(JsonObject request, string output, string format, bool validator, IReadOnlySet<string>? acceptedSourceFormats)
    {
        var required = new[] { "schema", "requestId", "runId", "subject", "artifact", "extraction", "expectedEvidenceSchema", "outputPath" };
        if (request.Count != required.Length || required.Any(name => !request.ContainsKey(name)) || request["schema"]!.GetValue<string>() != "tiwater.format-evidence-request/v1" || request["expectedEvidenceSchema"]!.GetValue<string>() != "lucid.published-format-evidence/v1") throw new InvalidOperationException("request contract invalid");
        if (!validator && Path.GetFullPath(request["outputPath"]!.GetValue<string>()) != output) throw new InvalidOperationException("output path mismatch");
        var artifact = request["artifact"]!.AsObject(); var file = artifact["path"]!.GetValue<string>(); var sourceFormat = artifact["format"]!.GetValue<string>();
        var allowedFormats = acceptedSourceFormats ?? new HashSet<string>(StringComparer.Ordinal) { format };
        if (!Path.IsPathFullyQualified(file) || !allowedFormats.Contains(sourceFormat) || FileSha(file) != artifact["bytesSha256"]!.GetValue<string>()) throw new InvalidOperationException("artifact authority mismatch");
        var extraction = request["extraction"]!.AsObject(); if (Sha(Canonical(extraction["options"]!)) != extraction["optionsSha256"]!.GetValue<string>()) throw new InvalidOperationException("extraction options mismatch");
        return sourceFormat;
    }

    private static JsonObject BuildEvidence(JsonObject request, string tool, string version, string format, string sourceFormat, Func<string, object> inspect, Func<string, IReadOnlyList<AdditionalObservation>>? additionalObservations)
    {
        var artifact = request["artifact"]!.AsObject(); var extraction = request["extraction"]!.AsObject(); var artifactPath = artifact["path"]!.GetValue<string>(); var inspection = JsonSerializer.SerializeToNode(inspect(artifactPath))!;
        var entity = new JsonObject { ["entityId"] = "document-1", ["kind"] = $"{format}-document", ["provenance"] = new JsonObject { ["source"] = "runtime", ["pointer"] = "/inspection" } };
        var observation = new JsonObject { ["observationId"] = "inspection-1", ["entityId"] = "document-1", ["semanticField"] = $"{format}.inspection", ["use"] = "structure", ["value"] = inspection, ["parentObservationIds"] = new JsonArray(), ["provenance"] = new JsonObject { ["source"] = "runtime", ["pointer"] = "/inspection" } };
        var observations = new JsonArray(observation);
        foreach (var item in additionalObservations?.Invoke(artifactPath) ?? []) observations.Add(new JsonObject { ["observationId"] = item.ObservationId, ["entityId"] = "document-1", ["semanticField"] = item.SemanticField, ["use"] = item.Use, ["value"] = JsonSerializer.SerializeToNode(item.Value), ["parentObservationIds"] = new JsonArray("inspection-1"), ["provenance"] = new JsonObject { ["source"] = "runtime", ["pointer"] = item.Pointer } });
        var epochMaterial = new JsonObject { ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(), ["runtimeTool"] = tool, ["runtimeSchema"] = extraction["schema"]!.GetValue<string>(), ["runtimeVersion"] = version, ["extractionOptions"] = extraction["options"]!.DeepClone() };
        var evidence = new JsonObject { ["schema"] = "lucid.published-format-evidence/v1", ["requestId"] = request["requestId"]!.GetValue<string>(), ["subject"] = request["subject"]!.DeepClone(), ["artifactVersionId"] = artifact["artifactVersionId"]!.GetValue<string>(), ["provider"] = new JsonObject { ["tool"] = tool, ["toolVersion"] = version, ["capabilityId"] = "inspect-evidence", ["capabilityVersion"] = "1", ["outputSchema"] = "lucid.published-format-evidence/v1" }, ["source"] = new JsonObject { ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(), ["format"] = sourceFormat }, ["extraction"] = extraction.DeepClone(), ["epoch"] = new JsonObject { ["epochId"] = $"ep-{Sha(Canonical(epochMaterial))}", ["bytesSha256"] = artifact["bytesSha256"]!.GetValue<string>(), ["runtimeTool"] = tool, ["runtimeSchema"] = extraction["schema"]!.GetValue<string>(), ["runtimeVersion"] = version, ["extractionOptionsSha256"] = extraction["optionsSha256"]!.GetValue<string>() }, ["entities"] = new JsonArray(entity), ["observations"] = observations };
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
