using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

public static class EffectExecutionCommand
{
    public sealed record ExecutionResult(JsonNode Receipt, bool Applied);

    public sealed record Contract(
        string ProviderId,
        string ProviderVersion,
        string EffectTypeId,
        string EffectTypeVersion,
        string AdapterId,
        string AdapterVersion,
        string OperationSchemaFile,
        string OperationSchemaId,
        string ReceiptSchemaFile,
        string ReceiptSchemaId,
        Func<JsonNode, JsonObject, ExecutionResult> Execute,
        Action<JsonNode> ValidateReceipt,
        Func<JsonNode, bool> ReceiptPassed);

    private sealed record Authority(
        JsonObject Request,
        JsonObject ProviderRequest,
        JsonObject Effect,
        JsonObject DerivationResult,
        JsonObject DerivationVerdict,
        JsonObject Bundle,
        JsonObject BundleVerdict,
        JsonObject BundleArtifact,
        JsonObject BundleVerdictArtifact);

    public static int RunProducer(string[] args, Contract contract)
    {
        var values = Arguments(args, validator: false);
        var authority = Admit(
            Load(values["request"]).AsObject(),
            Load(values["effect-bundle"]).AsObject(),
            values["effect-bundle"],
            Load(values["effect-verdict"]).AsObject(),
            values["effect-verdict"],
            contract);
        var outputPath = authority.ProviderRequest["output"]!["path"]!.GetValue<string>();
        if (File.Exists(outputPath))
            throw new InvalidOperationException("effect execution output must be fresh");
        var executed = contract.Execute(authority.Effect["operation"]!, authority.ProviderRequest);
        contract.ValidateReceipt(executed.Receipt);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("effect execution output artifact missing");
        var evidence = BuildEvidence(authority, executed, contract);
        Write(values["output"], evidence);
        return 0;
    }

    public static int RunValidator(string[] args, Contract contract)
    {
        var values = Arguments(args, validator: true);
        var authority = Admit(
            Load(values["request"]).AsObject(),
            Load(values["effect-bundle"]).AsObject(),
            values["effect-bundle"],
            Load(values["effect-verdict"]).AsObject(),
            values["effect-verdict"],
            contract);
        var evidence = Load(values["evidence"]).AsObject();
        var expected = IndependentlyRecomputeEvidence(authority, evidence, contract);
        var findings = new JsonArray();
        foreach (var field in new[]
        {
            "effectId", "outputId", "mode", "request", "receipt",
            "inputArtifact", "outputArtifact", "inputEpochId", "outputEpochId",
            "lineage", "status", "evidenceRefs"
        })
        {
            if (Canonical(evidence[field]) != Canonical(expected[field]))
                findings.Add(Finding("execution-evidence-recomputation-mismatch", $"/{field}"));
        }
        var requestId = authority.ProviderRequest["requestId"]!.GetValue<string>();
        var evidenceRef = new JsonObject
        {
            ["schema"] = ContractRef("lucid.execution-evidence-v2.schema.json", "lucid.execution-evidence/v2"),
            ["path"] = Path.GetFullPath(values["evidence"]),
            ["sha256"] = FileSha(values["evidence"])
        };
        var verdict = new JsonObject
        {
            ["requestId"] = requestId,
            ["evidence"] = evidenceRef,
            ["validator"] = new JsonObject
            {
                ["id"] = $"{contract.ProviderId}.execution-evidence-validator",
                ["version"] = contract.ProviderVersion
            },
            ["recomputedRequestSha256"] = authority.Request["requestSha256"]!.DeepClone(),
            ["recomputedReceiptSha256"] = expected["receipt"]!["sha256"]!.DeepClone(),
            ["recomputedLineageSha256"] = expected["lineage"]!["sha256"]!.DeepClone(),
            ["decision"] = findings.Count == 0 ? "pass" : "fail",
            ["findings"] = findings
        };
        Write(values["output"], verdict);
        return 0;
    }

    private static Authority Admit(
        JsonObject request,
        JsonObject bundle,
        string bundlePath,
        JsonObject bundleVerdict,
        string bundleVerdictPath,
        Contract contract)
    {
        ExactKeys(request, [
            "effectId", "effectBundleAuthority", "effectBundleVerdict",
            "effect", "providerRequest", "requestSha256"
        ], "effect request");
        var requestMaterial = request.DeepClone().AsObject();
        requestMaterial.Remove("requestSha256");
        if (request["requestSha256"]!.GetValue<string>() != Sha(Canonical(requestMaterial)))
            throw new InvalidOperationException("effect execution request hash mismatch");
        RequireTyped(request["effect"]!.AsObject(), "lucid.composed-effect/v2");
        RequireTyped(request["providerRequest"]!.AsObject(), "tiwater.provider-effect-execution-request/v1");
        var effect = request["effect"]!["value"]!.AsObject();
        var providerRequest = request["providerRequest"]!["value"]!.AsObject();
        ExactKeys(providerRequest, [
            "requestId", "runId", "outputId", "mode",
            "operationDerivationResult", "operationDerivationVerdict",
            "inputArtifact", "inputEpochId", "auxiliaryArtifacts", "output",
            "expectedReceiptContract", "expectedExecutionEvidenceContract", "lineageContract"
        ], "provider request");
        if (
            request["effectId"]!.GetValue<string>() != effect["effectId"]!.GetValue<string>() ||
            providerRequest["outputId"]!.GetValue<string>() != effect["outputId"]!.GetValue<string>() ||
            effect["effectDescriptor"]!["effectType"]!.GetValue<string>() != contract.EffectTypeId ||
            effect["effectDescriptor"]!["executionAdapter"]!["id"]!.GetValue<string>() != contract.AdapterId ||
            effect["effectDescriptor"]!["executionAdapter"]!["version"]!.GetValue<string>() != contract.AdapterVersion ||
            Canonical(effect["effectDescriptor"]!["operationSchema"]) != Canonical(
                ContractRef(contract.OperationSchemaFile, contract.OperationSchemaId)))
            throw new InvalidOperationException("effect execution effect descriptor mismatch");
        if (
            Canonical(providerRequest["expectedReceiptContract"]) != Canonical(
                ContractRef(contract.ReceiptSchemaFile, contract.ReceiptSchemaId)) ||
            Canonical(providerRequest["expectedExecutionEvidenceContract"]) != Canonical(
                ContractRef("lucid.execution-evidence-v2.schema.json", "lucid.execution-evidence/v2")) ||
            Canonical(providerRequest["lineageContract"]) != Canonical(
                ContractRef("tiwater.provider-artifact-lineage-v1.schema.json", "tiwater.provider-artifact-lineage/v1")))
            throw new InvalidOperationException("effect execution output contract mismatch");
        if (
            (providerRequest["mode"]!.GetValue<string>() == "mutation" && providerRequest["inputArtifact"] is null) ||
            (providerRequest["mode"]!.GetValue<string>() == "materialization" && providerRequest["inputArtifact"] is not null))
            throw new InvalidOperationException("effect execution mode input mismatch");
        if (providerRequest["inputArtifact"] is JsonObject input) RequireBlob(input, "input");
        foreach (var auxiliary in providerRequest["auxiliaryArtifacts"]!.AsArray())
            RequireBlob(auxiliary!.AsObject(), "auxiliary");
        var outputPath = providerRequest["output"]!["path"]!.GetValue<string>();
        if (!Path.IsPathFullyQualified(outputPath))
            throw new InvalidOperationException("effect execution output path invalid");
        var bundleRef = request["effectBundleAuthority"]!.AsObject();
        RequireNode(bundle, bundleRef, "lucid.effect-bundle/v3", "effect bundle");
        var verdictRef = request["effectBundleVerdict"]!.AsObject();
        RequireNode(bundleVerdict, verdictRef, "lucid.operator-verdict/v1", "effect bundle verdict");
        if (
            bundle["payload"]!["conflictVerdict"]!["decision"]!.GetValue<string>() != "pass" ||
            bundleVerdict["payload"]!["decision"]!.GetValue<string>() != "pass" ||
            Canonical(bundleVerdict["payload"]!["subject"]) != Canonical(bundleRef) ||
            bundleVerdict["payload"]!["recomputedPayloadSha256"]!.GetValue<string>() != bundle["payloadSha256"]!.GetValue<string>())
            throw new InvalidOperationException("effect execution bundle verdict invalid");
        var effectMatches = bundle["payload"]!["effects"]!.AsArray()
            .Where(candidate => Canonical(candidate) == Canonical(request["effect"]))
            .ToList();
        if (effectMatches.Count != 1)
            throw new InvalidOperationException("effect execution effect membership invalid");
        var derivationResult = LoadArtifact(
            providerRequest["operationDerivationResult"]!.AsObject(),
            "tiwater.operation-derivation-result/v1",
            "derivation result");
        var derivationVerdict = LoadArtifact(
            providerRequest["operationDerivationVerdict"]!.AsObject(),
            "tiwater.operation-derivation-verdict/v1",
            "derivation verdict");
        if (
            derivationVerdict["decision"]!.GetValue<string>() != "pass" ||
            derivationVerdict["result"]!["sha256"]!.GetValue<string>() !=
                providerRequest["operationDerivationResult"]!["sha256"]!.GetValue<string>() ||
            Canonical(derivationResult["operation"]) != Canonical(effect["operation"]))
            throw new InvalidOperationException("effect execution derivation authority invalid");
        return new Authority(
            request,
            providerRequest,
            effect,
            derivationResult,
            derivationVerdict,
            bundle,
            bundleVerdict,
            Artifact(
                "lucid.canonical-node-v2.schema.json",
                "lucid.canonical-node/v2",
                bundlePath),
            Artifact(
                "lucid.canonical-node-v2.schema.json",
                "lucid.canonical-node/v2",
                bundleVerdictPath));
    }

    private static JsonObject BuildEvidence(Authority authority, ExecutionResult executed, Contract contract)
    {
        var receipt = Typed(
            LucidContractRef(contract.ReceiptSchemaFile, contract.ReceiptSchemaId),
            executed.Receipt);
        var input = authority.ProviderRequest["inputArtifact"]?.DeepClone();
        var outputPath = authority.ProviderRequest["output"]!["path"]!.GetValue<string>();
        var outputBytesSha256 = FileSha(outputPath);
        var outputArtifactVersionId = $"artifact-{Sha(Canonical(new JsonObject
        {
            ["effectId"] = authority.Effect["effectId"]!.DeepClone(),
            ["inputArtifactVersionId"] = input?["artifactVersionId"]?.DeepClone(),
            ["outputBytesSha256"] = outputBytesSha256
        }))}";
        var outputEpochId = $"epoch-{Sha(Canonical(new JsonObject
        {
            ["effectId"] = authority.Effect["effectId"]!.DeepClone(),
            ["outputArtifactVersionId"] = outputArtifactVersionId,
            ["receiptSha256"] = receipt["sha256"]!.DeepClone()
        }))}";
        var outputArtifact = new JsonObject
        {
            ["artifactVersionId"] = outputArtifactVersionId,
            ["path"] = outputPath,
            ["bytesSha256"] = outputBytesSha256,
            ["mediaType"] = authority.ProviderRequest["output"]!["mediaType"]!.DeepClone()
        };
        var lineageValue = LineageValue(
            authority,
            receipt,
            outputArtifact,
            outputEpochId,
            contract);
        return Evidence(
            authority,
            receipt,
            outputArtifact,
            outputEpochId,
            Typed(
                LucidContractRef(
                    "tiwater.provider-artifact-lineage-v1.schema.json",
                    "tiwater.provider-artifact-lineage/v1"),
                lineageValue),
            executed.Applied ? "applied" : "failed");
    }

    private static JsonObject IndependentlyRecomputeEvidence(
        Authority authority,
        JsonObject evidence,
        Contract contract)
    {
        RequireTyped(evidence["receipt"]!.AsObject(), contract.ReceiptSchemaId);
        contract.ValidateReceipt(evidence["receipt"]!["value"]!);
        var outputPath = authority.ProviderRequest["output"]!["path"]!.GetValue<string>();
        var outputHash = FileSha(outputPath);
        var expectedVersion = $"artifact-{Sha(Canonical(new JsonObject
        {
            ["effectId"] = authority.Effect["effectId"]!.DeepClone(),
            ["inputArtifactVersionId"] = authority.ProviderRequest["inputArtifact"]?["artifactVersionId"]?.DeepClone(),
            ["outputBytesSha256"] = outputHash
        }))}";
        var receipt = Typed(
            LucidContractRef(contract.ReceiptSchemaFile, contract.ReceiptSchemaId),
            evidence["receipt"]!["value"]!);
        var expectedEpoch = $"epoch-{Sha(Canonical(new JsonObject
        {
            ["effectId"] = authority.Effect["effectId"]!.DeepClone(),
            ["outputArtifactVersionId"] = expectedVersion,
            ["receiptSha256"] = receipt["sha256"]!.DeepClone()
        }))}";
        var outputArtifact = new JsonObject
        {
            ["artifactVersionId"] = expectedVersion,
            ["path"] = outputPath,
            ["bytesSha256"] = outputHash,
            ["mediaType"] = authority.ProviderRequest["output"]!["mediaType"]!.DeepClone()
        };
        var lineage = Typed(
            LucidContractRef(
                "tiwater.provider-artifact-lineage-v1.schema.json",
                "tiwater.provider-artifact-lineage/v1"),
            LineageValue(authority, receipt, outputArtifact, expectedEpoch, contract));
        return Evidence(
            authority,
            receipt,
            outputArtifact,
            expectedEpoch,
            lineage,
            contract.ReceiptPassed(receipt["value"]!) ? "applied" : "failed");
    }

    private static JsonObject Evidence(
        Authority authority,
        JsonObject receipt,
        JsonObject outputArtifact,
        string outputEpochId,
        JsonObject lineage,
        string status) => new()
    {
        ["effectId"] = authority.Effect["effectId"]!.DeepClone(),
        ["outputId"] = authority.ProviderRequest["outputId"]!.DeepClone(),
        ["mode"] = authority.ProviderRequest["mode"]!.DeepClone(),
        ["request"] = Typed(
            LucidContractRef(
                "lucid.effect-execution-request-v1.schema.json",
                "lucid.effect-execution-request/v1"),
            authority.Request),
        ["receipt"] = receipt.DeepClone(),
        ["inputArtifact"] = authority.ProviderRequest["inputArtifact"]?.DeepClone(),
        ["outputArtifact"] = outputArtifact.DeepClone(),
        ["inputEpochId"] = authority.ProviderRequest["inputEpochId"]?.DeepClone(),
        ["outputEpochId"] = outputEpochId,
        ["lineage"] = lineage.DeepClone(),
        ["status"] = status,
        ["evidenceRefs"] = new JsonArray(
            EvidenceRef("authority", authority.BundleArtifact, authority.Request["effectBundleAuthority"]!["sha256"]!.GetValue<string>()),
            EvidenceRef("validator", authority.BundleVerdictArtifact, authority.Request["effectBundleVerdict"]!["sha256"]!.GetValue<string>()),
            EvidenceRef("producer", LucidArtifactRef(authority.ProviderRequest["operationDerivationResult"]!.AsObject()), authority.ProviderRequest["operationDerivationResult"]!["sha256"]!.GetValue<string>()),
            EvidenceRef("validator", LucidArtifactRef(authority.ProviderRequest["operationDerivationVerdict"]!.AsObject()), authority.ProviderRequest["operationDerivationVerdict"]!["sha256"]!.GetValue<string>()))
    };

    private static JsonObject LineageValue(
        Authority authority,
        JsonObject receipt,
        JsonObject outputArtifact,
        string outputEpochId,
        Contract contract) => new()
    {
        ["effectId"] = authority.Effect["effectId"]!.DeepClone(),
        ["outputId"] = authority.ProviderRequest["outputId"]!.DeepClone(),
        ["inputArtifactVersionId"] = authority.ProviderRequest["inputArtifact"]?["artifactVersionId"]?.DeepClone(),
        ["outputArtifactVersionId"] = outputArtifact["artifactVersionId"]!.DeepClone(),
        ["inputEpochId"] = authority.ProviderRequest["inputEpochId"]?.DeepClone(),
        ["outputEpochId"] = outputEpochId,
        ["inputBytesSha256"] = authority.ProviderRequest["inputArtifact"]?["bytesSha256"]?.DeepClone(),
        ["outputBytesSha256"] = outputArtifact["bytesSha256"]!.DeepClone(),
        ["operationSha256"] = authority.Effect["operation"]!["sha256"]!.DeepClone(),
        ["receiptSha256"] = receipt["sha256"]!.DeepClone(),
        ["provider"] = new JsonObject { ["id"] = contract.ProviderId, ["version"] = contract.ProviderVersion }
    };

    private static JsonObject LoadArtifact(JsonObject artifact, string schemaId, string label)
    {
        var path = artifact["path"]!.GetValue<string>();
        if (
            artifact["schema"]!["id"]!.GetValue<string>() != schemaId ||
            FileSha(path) != artifact["sha256"]!.GetValue<string>())
            throw new InvalidOperationException($"effect execution {label} artifact invalid");
        return Load(path).AsObject();
    }

    private static void RequireNode(JsonObject node, JsonObject reference, string contractId, string label)
    {
        if (
            node["nodeId"]!.GetValue<string>() != reference["nodeId"]!.GetValue<string>() ||
            node["contract"]!["id"]!.GetValue<string>() != contractId ||
            Canonical(node["contract"]) != Canonical(reference["contract"]) ||
            Sha(Canonical(node)) != reference["sha256"]!.GetValue<string>() ||
            node["payloadSha256"]!.GetValue<string>() != Sha(Canonical(node["payload"])))
            throw new InvalidOperationException($"effect execution {label} invalid");
    }

    private static void RequireBlob(JsonObject blob, string label)
    {
        ExactKeys(blob, ["artifactVersionId", "path", "bytesSha256", "mediaType"], $"{label} blob");
        var path = blob["path"]!.GetValue<string>();
        if (!Path.IsPathFullyQualified(path) || FileSha(path) != blob["bytesSha256"]!.GetValue<string>())
            throw new InvalidOperationException($"effect execution {label} blob invalid");
    }

    private static void RequireTyped(JsonObject typed, string? schemaId = null)
    {
        ExactKeys(typed, ["schema", "value", "sha256"], "typed value");
        if (
            typed["sha256"]!.GetValue<string>() != Sha(Canonical(typed["value"])) ||
            (schemaId is not null && typed["schema"]!["id"]!.GetValue<string>() != schemaId))
            throw new InvalidOperationException("effect execution typed value invalid");
    }

    private static JsonObject Typed(JsonObject schema, JsonNode value) => new()
    {
        ["schema"] = schema.DeepClone(),
        ["value"] = value.DeepClone(),
        ["sha256"] = Sha(Canonical(value))
    };

    private static JsonObject EvidenceRef(string kind, JsonObject artifact, string subjectSha256) => new()
    {
        ["kind"] = kind,
        ["artifact"] = artifact.DeepClone(),
        ["subjectSha256"] = subjectSha256
    };

    private static JsonObject Artifact(string schemaFile, string schemaId, string path) => new()
    {
        ["schema"] = LucidContractRef(schemaFile, schemaId),
        ["path"] = Path.GetFullPath(path),
        ["sha256"] = FileSha(path)
    };

    // A tiwater request references its artifacts with the schema bytes bound per
    // reference; the same artifact inside the execution evidence keeps only the
    // contract id.
    private static JsonObject LucidArtifactRef(JsonObject artifact)
    {
        var projected = artifact.DeepClone().AsObject();
        projected["schema"] = new JsonObject
        {
            ["id"] = artifact["schema"]!["id"]!.DeepClone()
        };
        return projected;
    }

    private static JsonObject Finding(string code, string path) =>
        new() { ["code"] = code, ["path"] = path };

    private static void ExactKeys(JsonObject value, IReadOnlyCollection<string> keys, string label)
    {
        if (value.Count != keys.Count || keys.Any(key => !value.ContainsKey(key)))
            throw new InvalidOperationException($"effect execution {label} fields invalid");
    }

    private static Dictionary<string, string> Arguments(string[] args, bool validator)
    {
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var index = 0; index < args.Length; index += 2)
        {
            if (index + 1 >= args.Length || !args[index].StartsWith("--", StringComparison.Ordinal))
                throw new InvalidOperationException("effect execution arguments invalid");
            values[args[index][2..]] = Path.GetFullPath(args[index + 1]);
        }
        var required = validator
            ? new[] { "request", "effect-bundle", "effect-verdict", "evidence", "output" }
            : new[] { "request", "effect-bundle", "effect-verdict", "output" };
        if (required.Any(key => !values.ContainsKey(key)))
            throw new InvalidOperationException("effect execution arguments missing");
        return values;
    }

    private static JsonNode Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))
        ?? throw new InvalidOperationException($"effect execution JSON invalid: {path}");

    private static JsonObject ContractRef(string file, string id)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        if (!File.Exists(path)) throw new InvalidOperationException($"provider contract missing: {file}");
        return new JsonObject { ["id"] = id, ["sha256"] = FileSha(path) };
    }

    // lucid.execution-evidence/v2 closes every schema reference it carries over
    // the contract id alone: the schema bytes are bound once by the Lucid schema
    // set. Tiwater-owned documents keep the per-reference sha256.
    private static JsonObject LucidContractRef(string file, string id)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        if (!File.Exists(path)) throw new InvalidOperationException($"provider contract missing: {file}");
        return new JsonObject { ["id"] = id };
    }

    private static void Write(string path, JsonNode value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        if (File.Exists(path))
            throw new InvalidOperationException("effect execution output evidence must be fresh");
        File.WriteAllText(path, $"{Canonical(value)}\n");
    }

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
