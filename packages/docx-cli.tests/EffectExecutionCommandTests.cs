using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

public sealed class EffectExecutionCommandTests
{
    private const string Hash = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

    [Fact]
    public void Exact_authorities_execute_and_independently_validate()
    {
        var fixture = Fixture();
        Assert.Equal(0, EffectExecutionCommand.RunProducer(
            [
                "--request", fixture.RequestPath,
                "--effect-bundle", fixture.BundlePath,
                "--effect-verdict", fixture.BundleVerdictPath,
                "--output", fixture.EvidencePath
            ],
            fixture.Contract));
        var evidence = Load(fixture.EvidencePath);
        Assert.Equal("applied", evidence["status"]!.GetValue<string>());
        Assert.Equal(FileSha(fixture.OutputArtifactPath), evidence["outputArtifact"]!["bytesSha256"]!.GetValue<string>());
        Assert.Equal(0, EffectExecutionCommand.RunValidator(
            [
                "--request", fixture.RequestPath,
                "--effect-bundle", fixture.BundlePath,
                "--effect-verdict", fixture.BundleVerdictPath,
                "--evidence", fixture.EvidencePath,
                "--output", fixture.ExecutionVerdictPath
            ],
            fixture.Contract));
        Assert.Equal("pass", Load(fixture.ExecutionVerdictPath)["decision"]!.GetValue<string>());
    }

    [Fact]
    public void Bundle_derivation_and_input_mutations_fail_before_execution()
    {
        var mutations = new Action<ExecutionFixture>[]
        {
            fixture => Mutate(fixture.BundleVerdictPath, node => node["payload"]!["decision"] = "failed"),
            fixture => Mutate(fixture.BundlePath, node => node["payload"]!["effects"] = new JsonArray()),
            fixture => Mutate(fixture.DerivationVerdictPath, node => node["decision"] = "fail"),
            fixture => File.WriteAllText(fixture.InputArtifactPath, "stale input")
        };
        foreach (var mutate in mutations)
        {
            var fixture = Fixture();
            mutate(fixture);
            Assert.ThrowsAny<Exception>(() => EffectExecutionCommand.RunProducer(
                [
                    "--request", fixture.RequestPath,
                    "--effect-bundle", fixture.BundlePath,
                    "--effect-verdict", fixture.BundleVerdictPath,
                    "--output", fixture.EvidencePath
                ],
                fixture.Contract));
            Assert.False(File.Exists(fixture.OutputArtifactPath));
        }
    }

    [Fact]
    public void Output_bytes_cannot_drift_after_provider_execution()
    {
        var fixture = Fixture();
        EffectExecutionCommand.RunProducer(
            [
                "--request", fixture.RequestPath,
                "--effect-bundle", fixture.BundlePath,
                "--effect-verdict", fixture.BundleVerdictPath,
                "--output", fixture.EvidencePath
            ],
            fixture.Contract);
        File.AppendAllText(fixture.OutputArtifactPath, "forged");
        EffectExecutionCommand.RunValidator(
            [
                "--request", fixture.RequestPath,
                "--effect-bundle", fixture.BundlePath,
                "--effect-verdict", fixture.BundleVerdictPath,
                "--evidence", fixture.EvidencePath,
                "--output", fixture.ExecutionVerdictPath
            ],
            fixture.Contract);
        Assert.Equal("fail", Load(fixture.ExecutionVerdictPath)["decision"]!.GetValue<string>());
    }

    private sealed record ExecutionFixture(
        string RequestPath,
        string BundlePath,
        string BundleVerdictPath,
        string DerivationVerdictPath,
        string EvidencePath,
        string ExecutionVerdictPath,
        string InputArtifactPath,
        string OutputArtifactPath,
        EffectExecutionCommand.Contract Contract);

    private static ExecutionFixture Fixture()
    {
        var root = Path.Combine(Path.GetTempPath(), $"effect-execution-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var inputPath = Path.Combine(root, "input.docx");
        var outputPath = Path.Combine(root, "output.docx");
        File.WriteAllText(inputPath, "provider input");
        var operation = Typed(
            "tiwater.docx-edit/v1",
            Ref("tiwater.docx-edit-v1.schema.json", "tiwater.docx-edit/v1")["sha256"]!.GetValue<string>(),
            new JsonObject
            {
                ["operations"] = new JsonArray(new JsonObject
                {
                    ["type"] = "replaceBodyText",
                    ["findText"] = "input",
                    ["text"] = "output"
                })
            });
        var effectValue = new JsonObject
        {
            ["effectId"] = "effect-1",
            ["operationId"] = "operation-1",
            ["planId"] = "plan-1",
            ["closureDeclarationId"] = "closure-1",
            ["effectIntentId"] = "intent-1",
            ["outputId"] = "primary",
            ["target"] = new JsonObject(),
            ["effectDescriptor"] = new JsonObject
            {
                ["effectType"] = "docx.edit",
                ["descriptorSha256"] = Hash,
                ["operationSchema"] = Ref("tiwater.docx-edit-v1.schema.json", "tiwater.docx-edit/v1"),
                ["executionAdapter"] = new JsonObject { ["id"] = "tiwater-docx-edit", ["version"] = "0.10.18" }
            },
            ["operation"] = operation,
            ["resourceSet"] = new JsonArray(),
            ["writeSet"] = new JsonArray(),
            ["topologicalIndex"] = 0,
            ["afterEffectIds"] = new JsonArray(),
            ["authorities"] = new JsonObject(),
            ["evidenceRefs"] = new JsonArray()
        };
        var effect = Typed(
            "lucid.composed-effect/v2",
            Ref("lucid.composed-effect-v2.schema.json", "lucid.composed-effect/v2")["sha256"]!.GetValue<string>(),
            effectValue);
        var derivationResult = new JsonObject
        {
            ["schema"] = "tiwater.operation-derivation-result/v1",
            ["derivationId"] = "derivation-1",
            ["requestId"] = "derive-1",
            ["effectDescriptor"] = new JsonObject(),
            ["output"] = new JsonObject(),
            ["targetCandidateId"] = "target-1",
            ["operation"] = operation.DeepClone(),
            ["resourceSet"] = Typed("tiwater.provider-resource-set/v1", Hash, new JsonArray()),
            ["writeSet"] = Typed("tiwater.provider-write-set/v1", Hash, new JsonArray()),
            ["provenance"] = Typed("tiwater.operation-derivation-provenance/v1", Hash, new JsonObject())
        };
        var derivationResultPath = Write(root, "derivation-result.json", derivationResult);
        var derivationVerdict = new JsonObject
        {
            ["schema"] = "tiwater.operation-derivation-verdict/v1",
            ["requestId"] = "derive-1",
            ["result"] = new JsonObject
            {
                ["schema"] = Ref("tiwater.operation-derivation-result-v1.schema.json", "tiwater.operation-derivation-result/v1"),
                ["sha256"] = FileSha(derivationResultPath)
            },
            ["validator"] = new JsonObject { ["id"] = "validator", ["version"] = "1" },
            ["recomputedOperationSha256"] = operation["sha256"]!.DeepClone(),
            ["recomputedResourceSetSha256"] = Hash,
            ["recomputedWriteSetSha256"] = Hash,
            ["recomputedProvenanceSha256"] = Hash,
            ["decision"] = "pass",
            ["findings"] = new JsonArray()
        };
        var derivationVerdictPath = Write(root, "derivation-verdict.json", derivationVerdict);
        var providerRequestValue = new JsonObject
        {
            ["requestId"] = "execute-1",
            ["runId"] = "run-1",
            ["outputId"] = "primary",
            ["mode"] = "mutation",
            ["operationDerivationResult"] = Artifact(
                derivationResultPath,
                "tiwater.operation-derivation-result-v1.schema.json",
                "tiwater.operation-derivation-result/v1"),
            ["operationDerivationVerdict"] = Artifact(
                derivationVerdictPath,
                "tiwater.operation-derivation-verdict-v1.schema.json",
                "tiwater.operation-derivation-verdict/v1"),
            ["inputArtifact"] = new JsonObject
            {
                ["artifactVersionId"] = "input-artifact",
                ["path"] = inputPath,
                ["bytesSha256"] = FileSha(inputPath),
                ["mediaType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            ["inputEpochId"] = "input-epoch",
            ["auxiliaryArtifacts"] = new JsonArray(),
            ["output"] = new JsonObject
            {
                ["path"] = outputPath,
                ["mediaType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            ["expectedReceiptContract"] = Ref("tiwater.docx-edit-result-v1.schema.json", "tiwater.docx-edit-result/v1"),
            ["expectedExecutionEvidenceContract"] = Ref("lucid.execution-evidence-v2.schema.json", "lucid.execution-evidence/v2"),
            ["lineageContract"] = Ref("tiwater.provider-artifact-lineage-v1.schema.json", "tiwater.provider-artifact-lineage/v1")
        };
        var bundlePayload = new JsonObject
        {
            ["effects"] = new JsonArray(effect.DeepClone()),
            ["resourceSet"] = new JsonArray(),
            ["writeSet"] = new JsonArray(),
            ["conflictVerdict"] = new JsonObject { ["decision"] = "pass", ["findings"] = new JsonArray() },
            ["bundleSha256"] = Hash
        };
        var bundle = CanonicalNode(
            "bundle-node",
            Ref("lucid.effect-bundle-v3.schema.json", "lucid.effect-bundle/v3"),
            bundlePayload);
        var bundleRef = NodeRef(bundle);
        var bundlePath = Write(root, "bundle.json", bundle);
        var bundleVerdictPayload = new JsonObject
        {
            ["subject"] = bundleRef.DeepClone(),
            ["decision"] = "pass",
            ["recomputedPayloadSha256"] = bundle["payloadSha256"]!.DeepClone(),
            ["recomputedAuthoritySetSha256"] = Hash,
            ["findings"] = new JsonArray(),
            ["evidenceRefs"] = new JsonArray()
        };
        var bundleVerdict = CanonicalNode(
            "bundle-verdict-node",
            Ref("lucid.operator-verdict-v1.schema.json", "lucid.operator-verdict/v1"),
            bundleVerdictPayload);
        var bundleVerdictPath = Write(root, "bundle-verdict.json", bundleVerdict);
        var request = new JsonObject
        {
            ["effectId"] = "effect-1",
            ["effectBundleAuthority"] = bundleRef,
            ["effectBundleVerdict"] = NodeRef(bundleVerdict),
            ["effect"] = effect,
            ["providerRequest"] = Typed(
                "tiwater.provider-effect-execution-request/v1",
                Ref("tiwater.provider-effect-execution-request-v1.schema.json", "tiwater.provider-effect-execution-request/v1")["sha256"]!.GetValue<string>(),
                providerRequestValue)
        };
        request["requestSha256"] = Sha(Canonical(request));
        var requestPath = Write(root, "request.json", request);
        var contract = new EffectExecutionCommand.Contract(
            "tiwater-docx",
            "0.10.18",
            "docx.edit",
            "1",
            "tiwater-docx-edit",
            "0.10.18",
            "tiwater.docx-edit-v1.schema.json",
            "tiwater.docx-edit/v1",
            "tiwater.docx-edit-result-v1.schema.json",
            "tiwater.docx-edit-result/v1",
            (typedOperation, providerRequest) =>
            {
                var target = providerRequest["output"]!["path"]!.GetValue<string>();
                File.WriteAllText(target, $"{File.ReadAllText(inputPath)}:applied");
                var receipt = new JsonObject
                {
                    ["input"] = inputPath,
                    ["output"] = target,
                    ["appliedOperations"] = new JsonArray(new JsonObject
                    {
                        ["type"] = "replaceBodyText",
                        ["applied"] = true,
                        ["detail"] = "applied"
                    })
                };
                return new(receipt, true);
            },
            receipt => Assert.True(receipt["appliedOperations"]![0]!["applied"]!.GetValue<bool>()),
            receipt => receipt["appliedOperations"]![0]!["applied"]!.GetValue<bool>());
        return new(
            requestPath,
            bundlePath,
            bundleVerdictPath,
            derivationVerdictPath,
            Path.Combine(root, "evidence.json"),
            Path.Combine(root, "execution-verdict.json"),
            inputPath,
            outputPath,
            contract);
    }

    private static JsonObject CanonicalNode(string nodeId, JsonObject contract, JsonObject payload)
    {
        return new JsonObject
        {
            ["schema"] = "lucid.canonical-node/v2",
            ["runId"] = "run-1",
            ["nodeId"] = nodeId,
            ["contract"] = contract,
            ["dependsOn"] = new JsonArray(),
            ["authorityRefs"] = new JsonArray(),
            ["producer"] = new JsonObject
            {
                ["capability"] = new JsonObject { ["id"] = "fixture", ["version"] = "1" },
                ["adapter"] = new JsonObject { ["id"] = "fixture", ["version"] = "1" },
                ["registrySha256"] = Hash
            },
            ["payload"] = payload,
            ["payloadSha256"] = Sha(Canonical(payload)),
            ["diagnostics"] = new JsonArray()
        };
    }

    private static JsonObject NodeRef(JsonObject node) => new()
    {
        ["nodeId"] = node["nodeId"]!.DeepClone(),
        ["contract"] = node["contract"]!.DeepClone(),
        ["sha256"] = Sha(Canonical(node))
    };

    private static JsonObject Typed(string id, string schemaHash, JsonNode value) => new()
    {
        ["schema"] = new JsonObject { ["id"] = id, ["sha256"] = schemaHash },
        ["value"] = value.DeepClone(),
        ["sha256"] = Sha(Canonical(value))
    };

    private static JsonObject Artifact(string path, string schemaFile, string schemaId) => new()
    {
        ["schema"] = Ref(schemaFile, schemaId),
        ["path"] = path,
        ["sha256"] = FileSha(path)
    };

    private static JsonObject Ref(string file, string id) => new()
    {
        ["id"] = id,
        ["sha256"] = FileSha(Path.Combine(AppContext.BaseDirectory, "contracts", file))
    };

    private static string Write(string root, string name, JsonNode value)
    {
        var path = Path.Combine(root, name);
        File.WriteAllText(path, $"{Canonical(value)}\n");
        return path;
    }

    private static void Mutate(string path, Action<JsonObject> mutation)
    {
        var value = Load(path);
        mutation(value);
        if (value["payload"] is JsonObject payload)
            value["payloadSha256"] = Sha(Canonical(payload));
        File.WriteAllText(path, $"{Canonical(value)}\n");
    }

    private static JsonObject Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))!.AsObject();

    private static string FileSha(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

    private static string Sha(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

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
}
