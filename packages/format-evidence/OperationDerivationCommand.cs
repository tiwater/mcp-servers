using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

public static class OperationDerivationCommand
{
    public sealed record Contract(
        string ProviderId,
        string ProviderVersion,
        string Format,
        string EffectTypeId,
        string EffectTypeVersion,
        string DerivationAdapterId,
        string DerivationAdapterVersion,
        string ExecutionAdapterId,
        string OperationSchemaFile,
        string OperationSchemaId,
        string TargetScope,
        string OperationShape,
        bool IncludeOperationKindField,
        Action<JsonNode> ValidateOperation);

    private sealed record DerivedValues(
        JsonNode Operation,
        JsonArray ResourceSet,
        JsonArray WriteSet);

    public static int RunProducer(string[] args, Contract contract)
    {
        var request = Load(Argument(args, "request")).AsObject();
        var output = Argument(args, "output");
        Write(output, Produce(request, contract));
        return 0;
    }

    public static int RunValidator(string[] args, Contract contract)
    {
        var request = Load(Argument(args, "request")).AsObject();
        var resultPath = Argument(args, "result");
        var result = Load(resultPath).AsObject();
        var output = Argument(args, "output");
        Write(output, Validate(request, result, resultPath, contract));
        return 0;
    }

    public static JsonObject Produce(JsonObject request, Contract contract)
    {
        var authority = AdmitRequest(request, contract);
        var values = DeriveProducerValues(authority.Target, authority.SourceFact, authority.Intent, contract);
        ValidatePublishedOperationShape(values.Operation, contract);
        contract.ValidateOperation(values.Operation);
        return BuildResult(request, authority, values, contract);
    }

    public static JsonObject Validate(
        JsonObject request,
        JsonObject result,
        string resultPath,
        Contract contract)
    {
        var authority = AdmitRequest(request, contract);
        var values = DeriveValidatorValues(authority.Target, authority.SourceFact, authority.Intent, contract);
        ValidatePublishedOperationShape(values.Operation, contract);
        contract.ValidateOperation(values.Operation);
        var expected = BuildResult(request, authority, values, contract);
        var findings = new JsonArray();
        CompareTyped(result, expected, "operation", findings);
        CompareTyped(result, expected, "resourceSet", findings);
        CompareTyped(result, expected, "writeSet", findings);
        CompareTyped(result, expected, "provenance", findings);
        foreach (var field in new[] { "requestId", "effectDescriptor", "output", "targetCandidateId" })
        {
            if (Canonical(result[field]) != Canonical(expected[field]))
                findings.Add(Finding("operation-derivation-result-mismatch", $"/{field}"));
        }
        var resultSchema = ContractRef(
            "tiwater.operation-derivation-result-v1.schema.json",
            "tiwater.operation-derivation-result/v1");
        var resultRef = new JsonObject
        {
            ["schema"] = resultSchema,
            ["sha256"] = Sha(File.ReadAllText(resultPath))
        };
        return new JsonObject
        {
            ["schema"] = "tiwater.operation-derivation-verdict/v1",
            ["requestId"] = request["requestId"]!.DeepClone(),
            ["result"] = resultRef,
            ["validator"] = new JsonObject
            {
                ["id"] = $"{contract.ProviderId}.operation-derivation-validator",
                ["version"] = contract.ProviderVersion
            },
            ["recomputedOperationSha256"] = expected["operation"]!["sha256"]!.DeepClone(),
            ["recomputedResourceSetSha256"] = expected["resourceSet"]!["sha256"]!.DeepClone(),
            ["recomputedWriteSetSha256"] = expected["writeSet"]!["sha256"]!.DeepClone(),
            ["recomputedProvenanceSha256"] = expected["provenance"]!["sha256"]!.DeepClone(),
            ["decision"] = findings.Count == 0 ? "pass" : "fail",
            ["findings"] = findings
        };
    }

    private sealed record Admitted(
        JsonObject Target,
        JsonObject SourceFact,
        JsonObject Intent,
        JsonObject OperationSchema,
        JsonObject ResourceSchema,
        JsonObject WriteSchema,
        string RequestSha256);

    private static Admitted AdmitRequest(JsonObject request, Contract contract)
    {
        var requestSchema = request["schema"]!.GetValue<string>();
        var isV2 = requestSchema == "tiwater.operation-derivation-request/v2";
        if (isV2)
            ExactKeys(request, [
                "schema", "requestId", "runId", "effectIntentId", "bindingId",
                "closureAuthority", "bindingAuthority", "normalizedFactsAuthority",
                "effectDescriptor", "output", "targetArtifact", "observation",
                "target", "sourceFacts", "effectIntent", "provider", "expectedResultContract"
            ], "request");
        else
            ExactKeys(request, [
                "schema", "requestId", "runId", "effectDescriptor", "output",
                "targetArtifact", "observation", "target", "sourceFact", "effectIntent",
                "bindingAuthority", "provider", "expectedResultContract"
            ], "request");
        if (!isV2 && requestSchema != "tiwater.operation-derivation-request/v1")
            throw new InvalidOperationException("operation derivation request schema invalid");
        var expectedResult = ContractRef(
            "tiwater.operation-derivation-result-v1.schema.json",
            "tiwater.operation-derivation-result/v1");
        if (Canonical(request["expectedResultContract"]) != Canonical(expectedResult))
            throw new InvalidOperationException("operation derivation result contract mismatch");
        var descriptor = request["effectDescriptor"]!.AsObject();
        var expectedOperation = ContractRef(contract.OperationSchemaFile, contract.OperationSchemaId);
        var resourceSchema = ContractRef(
            "tiwater.provider-resource-set-v1.schema.json",
            "tiwater.provider-resource-set/v1");
        var writeSchema = ContractRef(
            "tiwater.provider-write-set-v1.schema.json",
            "tiwater.provider-write-set/v1");
        if (isV2)
            ExactKeys(descriptor, [
                "identity", "descriptorSha256", "operationSchema", "resourceSetSchema",
                "writeSetSchema", "targetScope", "executionAdapter"
            ], "effect descriptor");
        if (
            descriptor["identity"]!["id"]!.GetValue<string>() != contract.EffectTypeId ||
            descriptor["identity"]!["version"]!.GetValue<string>() != contract.EffectTypeVersion ||
            Canonical(descriptor["operationSchema"]) != Canonical(expectedOperation) ||
            Canonical(descriptor["resourceSetSchema"]) != Canonical(resourceSchema) ||
            Canonical(descriptor["writeSetSchema"]) != Canonical(writeSchema) ||
            descriptor["targetScope"]!.GetValue<string>() != contract.TargetScope)
            throw new InvalidOperationException("operation derivation effect descriptor mismatch");
        var provider = request["provider"]!.AsObject();
        if (
            provider["identity"]!["id"]!.GetValue<string>() != contract.ProviderId ||
            provider["identity"]!["version"]!.GetValue<string>() != contract.ProviderVersion ||
            (isV2 && (
                provider["adapter"]!["id"]!.GetValue<string>() != contract.DerivationAdapterId ||
                provider["adapter"]!["version"]!.GetValue<string>() != contract.DerivationAdapterVersion ||
                descriptor["executionAdapter"]!["id"]!.GetValue<string>() != contract.ExecutionAdapterId ||
                descriptor["executionAdapter"]!["version"]!.GetValue<string>() != contract.ProviderVersion)))
            throw new InvalidOperationException("operation derivation provider mismatch");
        var output = request["output"]!.AsObject();
        var outputArtifact = output["artifact"]!.AsObject();
        var targetArtifact = request["targetArtifact"]!.AsObject();
        JsonObject target;
        JsonObject sourceFact;
        if (isV2)
        {
            RequireText(request, "effectIntentId");
            RequireText(request, "bindingId");
            RequireNodeRef(request["closureAuthority"]!.AsObject());
            RequireNodeRef(request["bindingAuthority"]!.AsObject());
            RequireNodeRef(request["normalizedFactsAuthority"]!.AsObject());
            var targetEnvelope = request["target"]!.AsObject();
            ExactKeys(targetEnvelope, [
                "targetAuthority", "targetId", "candidate", "resourceSet", "writeSet"
            ], "target");
            RequireNodeRef(targetEnvelope["targetAuthority"]!.AsObject());
            RequireText(targetEnvelope, "targetId");
            target = targetEnvelope["candidate"]!.AsObject();
            var sourceFacts = request["sourceFacts"]!.AsArray();
            if (sourceFacts.Count != 1)
                throw new InvalidOperationException("operation derivation source facts cardinality invalid");
            sourceFact = sourceFacts[0]!.AsObject();
            ExactKeys(sourceFact, ["factId", "value"], "source fact");
            RequireText(sourceFact, "factId");
        }
        else
        {
            target = request["target"]!.AsObject();
            sourceFact = request["sourceFact"]!.AsObject();
        }
        if (
            output["format"]!.GetValue<string>() != contract.Format ||
            target["artifactVersionId"]!.GetValue<string>() != targetArtifact["artifactVersionId"]!.GetValue<string>() ||
            (contract.TargetScope == "current-artifact" &&
                (Canonical(targetArtifact) != Canonical(outputArtifact) ||
                 target["epochId"]!.GetValue<string>() != output["epochId"]!.GetValue<string>())))
            throw new InvalidOperationException("operation derivation output target mismatch");
        RequireArtifact(outputArtifact, "output");
        RequireArtifact(targetArtifact, "target");
        RequireTyped(request["observation"]!.AsObject(), "tiwater.provider-document-observation/v2");
        var observation = request["observation"]!["value"]!.AsObject();
        if (
            observation["format"]!.GetValue<string>() != contract.Format ||
            observation["artifactVersionId"]!.GetValue<string>() != targetArtifact["artifactVersionId"]!.GetValue<string>() ||
            observation["epochId"]!.GetValue<string>() != target["epochId"]!.GetValue<string>() ||
            observation["inspectionSha256"]!.GetValue<string>() != target["inspectionSha256"]!.GetValue<string>())
            throw new InvalidOperationException("operation derivation observation identity mismatch");
        var candidates = observation["targetUniverse"]!["candidates"]!.AsArray()
            .Where(candidate => candidate!["candidateId"]!.GetValue<string>() == target["candidateId"]!.GetValue<string>())
            .ToList();
        if (candidates.Count != 1 || Canonical(candidates[0]) != Canonical(target))
            throw new InvalidOperationException("operation derivation target authority mismatch");
        var capabilityCount = target["capabilities"]!.AsArray().Count(capability =>
            capability!["id"]!.GetValue<string>() == contract.EffectTypeId &&
            capability!["version"]!.GetValue<string>() == contract.EffectTypeVersion);
        if (capabilityCount != 1)
            throw new InvalidOperationException("operation derivation target capability mismatch");
        RequireTyped(target["locator"]!.AsObject(), "tiwater.provider-json-pointer-locator/v1");
        RequireTyped(sourceFact["value"]!.AsObject());
        var intentTyped = request["effectIntent"]!.AsObject();
        RequireTyped(intentTyped, "tiwater.provider-effect-intent/v1");
        var intent = intentTyped["value"]!.AsObject();
        if (
            (isV2 && request["effectIntentId"]!.GetValue<string>() != intent["effectId"]!.GetValue<string>()) ||
            intent["effectType"]!["id"]!.GetValue<string>() != contract.EffectTypeId ||
            intent["effectType"]!["version"]!.GetValue<string>() != contract.EffectTypeVersion ||
            !target["supportedOperationKinds"]!.AsArray().Any(kind =>
                kind!.GetValue<string>() == intent["operationKind"]!.GetValue<string>()))
            throw new InvalidOperationException("operation derivation intent effect mismatch");
        if (!isV2) RequireTyped(request["bindingAuthority"]!.AsObject());
        foreach (var field in target["semanticIdentity"]!.AsArray())
            if (field!["sha256"]!.GetValue<string>() != Sha(Canonical(field["value"])))
                throw new InvalidOperationException("operation derivation semantic identity hash mismatch");
        var resources = target["resourceDeclarations"]!.AsArray();
        var resourceKeys = resources.Select(value => value!["resourceKey"]!.GetValue<string>()).ToHashSet(StringComparer.Ordinal);
        if (resources.Count == 0 || resourceKeys.Count != resources.Count)
            throw new InvalidOperationException("operation derivation resource declarations invalid");
        var writes = target["writeDeclarations"]!.AsArray();
        if (writes.Count == 0 || writes.Any(value => !resourceKeys.Contains(value!["resourceKey"]!.GetValue<string>())))
            throw new InvalidOperationException("operation derivation write declarations invalid");
        if (isV2)
        {
            var targetEnvelope = request["target"]!.AsObject();
            RequireClaim(
                targetEnvelope["resourceSet"]!.AsArray(),
                resourceSchema,
                resources,
                "resource set");
            RequireClaim(
                targetEnvelope["writeSet"]!.AsArray(),
                writeSchema,
                writes,
                "write set");
        }
        return new Admitted(
            target,
            sourceFact,
            intent,
            expectedOperation,
            resourceSchema,
            writeSchema,
            Sha(Canonical(request)));
    }

    private static DerivedValues DeriveProducerValues(
        JsonObject target,
        JsonObject sourceFact,
        JsonObject intent,
        Contract contract)
    {
        var operation = new JsonObject();
        if (contract.IncludeOperationKindField)
            operation["type"] = intent["operationKind"]!.DeepClone();
        if (contract.OperationShape == "single-edit")
            foreach (var field in target["semanticIdentity"]!.AsArray())
                AddExact(operation, field!["name"]!.GetValue<string>(), field["value"]);
        foreach (var argument in intent["arguments"]!.AsArray())
        {
            var item = argument!.AsObject();
            var value = ResolveProducerArgument(item, target, sourceFact);
            AddExact(operation, item["name"]!.GetValue<string>(), value);
        }
        return new DerivedValues(
            contract.OperationShape == "single-edit"
                ? new JsonObject { ["operations"] = new JsonArray(operation) }
                : operation,
            target["resourceDeclarations"]!.AsArray().DeepClone().AsArray(),
            target["writeDeclarations"]!.AsArray().DeepClone().AsArray());
    }

    private static DerivedValues DeriveValidatorValues(
        JsonObject target,
        JsonObject sourceFact,
        JsonObject intent,
        Contract contract)
    {
        var members = new SortedDictionary<string, JsonNode?>(StringComparer.Ordinal);
        if (contract.IncludeOperationKindField)
            members["type"] = intent["operationKind"]!.DeepClone();
        if (contract.OperationShape == "single-edit")
            foreach (var identity in target["semanticIdentity"]!.AsArray())
            {
                var name = identity!["name"]!.GetValue<string>();
                if (!members.TryAdd(name, identity["value"]!.DeepClone()))
                    throw new InvalidOperationException("operation derivation duplicate semantic field");
            }
        foreach (var entry in intent["arguments"]!.AsArray())
        {
            var argument = entry!.AsObject();
            var name = argument["name"]!.GetValue<string>();
            var value = ResolveValidatorArgument(argument, target, sourceFact);
            if (members.TryGetValue(name, out var prior) && Canonical(prior) != Canonical(value))
                throw new InvalidOperationException("operation derivation conflicting field");
            members[name] = value;
        }
        var operation = new JsonObject();
        foreach (var (name, value) in members) operation[name] = value;
        var resources = new JsonArray(target["resourceDeclarations"]!.AsArray()
            .Select(value => value!.DeepClone()).ToArray());
        var writes = new JsonArray(target["writeDeclarations"]!.AsArray()
            .Select(value => value!.DeepClone()).ToArray());
        return new DerivedValues(
            contract.OperationShape == "single-edit"
                ? new JsonObject { ["operations"] = new JsonArray(operation) }
                : operation,
            resources,
            writes);
    }

    private static JsonObject BuildResult(
        JsonObject request,
        Admitted authority,
        DerivedValues values,
        Contract contract)
    {
        var operation = Typed(authority.OperationSchema, values.Operation);
        var resources = Typed(authority.ResourceSchema, values.ResourceSet);
        var writes = Typed(authority.WriteSchema, values.WriteSet);
        var provenanceValue = new JsonObject
        {
            ["requestSha256"] = authority.RequestSha256,
            ["effectDescriptorSha256"] = request["effectDescriptor"]!["descriptorSha256"]!.DeepClone(),
            ["targetCandidateId"] = authority.Target["candidateId"]!.DeepClone(),
            ["targetLocatorSha256"] = authority.Target["locator"]!["sha256"]!.DeepClone(),
            ["sourceFactSha256"] = authority.SourceFact["value"]!["sha256"]!.DeepClone(),
            ["effectIntentSha256"] = request["effectIntent"]!["sha256"]!.DeepClone(),
            ["bindingAuthoritySha256"] = request["bindingAuthority"]!["sha256"]!.DeepClone(),
            ["operationSha256"] = operation["sha256"]!.DeepClone(),
            ["resourceSetSha256"] = resources["sha256"]!.DeepClone(),
            ["writeSetSha256"] = writes["sha256"]!.DeepClone(),
            ["provider"] = request["provider"]!["identity"]!.DeepClone()
        };
        var provenance = Typed(
            ContractRef(
                "tiwater.operation-derivation-provenance-v1.schema.json",
                "tiwater.operation-derivation-provenance/v1"),
            provenanceValue);
        var material = new JsonObject
        {
            ["requestId"] = request["requestId"]!.DeepClone(),
            ["operationSha256"] = operation["sha256"]!.DeepClone(),
            ["resourceSetSha256"] = resources["sha256"]!.DeepClone(),
            ["writeSetSha256"] = writes["sha256"]!.DeepClone(),
            ["provenanceSha256"] = provenance["sha256"]!.DeepClone()
        };
        return new JsonObject
        {
            ["schema"] = "tiwater.operation-derivation-result/v1",
            ["derivationId"] = $"derivation-{Sha(Canonical(material))}",
            ["requestId"] = request["requestId"]!.DeepClone(),
            ["effectDescriptor"] = new JsonObject
            {
                ["identity"] = request["effectDescriptor"]!["identity"]!.DeepClone(),
                ["descriptorSha256"] = request["effectDescriptor"]!["descriptorSha256"]!.DeepClone()
            },
            ["output"] = new JsonObject
            {
                ["outputId"] = request["output"]!["outputId"]!.DeepClone(),
                ["artifactVersionId"] = request["output"]!["artifact"]!["artifactVersionId"]!.DeepClone(),
                ["epochId"] = request["output"]!["epochId"]!.DeepClone()
            },
            ["targetCandidateId"] = authority.Target["candidateId"]!.DeepClone(),
            ["operation"] = operation,
            ["resourceSet"] = resources,
            ["writeSet"] = writes,
            ["provenance"] = provenance
        };
    }

    private static void CompareTyped(JsonObject actual, JsonObject expected, string field, JsonArray findings)
    {
        if (Canonical(actual[field]) != Canonical(expected[field]))
            findings.Add(Finding("operation-derivation-recomputation-mismatch", $"/{field}"));
    }

    private static JsonObject Finding(string code, string path) =>
        new() { ["code"] = code, ["path"] = path };

    private static void ValidatePublishedOperationShape(JsonNode value, Contract contract)
    {
        var schemaPath = Path.Combine(AppContext.BaseDirectory, "contracts", contract.OperationSchemaFile);
        var schema = Load(schemaPath).AsObject();
        var root = value.AsObject();
        var rootProperties = schema["properties"]!.AsObject();
        var rootRequired = schema["required"]!.AsArray()
            .Select(item => item!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);
        if (
            root.Any(item => !rootProperties.ContainsKey(item.Key)) ||
            rootRequired.Any(name => !root.ContainsKey(name)))
            throw new InvalidOperationException("operation derivation operation root invalid");
        if (contract.OperationShape == "root-object") return;
        var operations = root["operations"]!.AsArray();
        if (operations.Count != 1)
            throw new InvalidOperationException("operation derivation operation cardinality invalid");
        var itemSchema = schema["$defs"]?["operation"]?.AsObject()
            ?? rootProperties["operations"]!["items"]!.AsObject();
        var item = operations[0]!.AsObject();
        var itemProperties = itemSchema["properties"]!.AsObject();
        var itemRequired = itemSchema["required"]!.AsArray()
            .Select(entry => entry!.GetValue<string>())
            .ToHashSet(StringComparer.Ordinal);
        if (
            item.Any(entry => !itemProperties.ContainsKey(entry.Key)) ||
            itemRequired.Any(name => !item.ContainsKey(name)))
            throw new InvalidOperationException("operation derivation operation fields invalid");
        if (itemProperties["type"]?["enum"] is JsonArray kindValues)
        {
            var allowedKinds = kindValues.Select(entry => entry!.GetValue<string>())
                .ToHashSet(StringComparer.Ordinal);
            if (!allowedKinds.Contains(item["type"]!.GetValue<string>()))
                throw new InvalidOperationException("operation derivation operation kind invalid");
        }
    }

    private static void AddExact(JsonObject target, string name, JsonNode? value)
    {
        if (target.TryGetPropertyValue(name, out var prior) && Canonical(prior) != Canonical(value))
            throw new InvalidOperationException($"operation derivation conflicting field: {name}");
        target[name] = value?.DeepClone();
    }

    private static JsonNode? ResolveProducerArgument(JsonObject argument, JsonObject target, JsonObject sourceFact) =>
        argument["source"]!.GetValue<string>() switch
        {
            "source-fact" => sourceFact["value"]!["value"],
            "literal" => argument["value"]!["value"],
            "target-field" => target["semanticIdentity"]!.AsArray().Single(field =>
                field!["name"]!.GetValue<string>() == argument["fieldName"]!.GetValue<string>())!["value"],
            _ => throw new InvalidOperationException("operation derivation argument source invalid")
        };

    private static JsonNode? ResolveValidatorArgument(JsonObject argument, JsonObject target, JsonObject sourceFact)
    {
        var source = argument["source"]!.GetValue<string>();
        if (source == "source-fact") return sourceFact["value"]!["value"]!.DeepClone();
        if (source == "literal") return argument["value"]!["value"]!.DeepClone();
        if (source != "target-field")
            throw new InvalidOperationException("operation derivation validator argument source invalid");
        var name = argument["fieldName"]!.GetValue<string>();
        var matches = target["semanticIdentity"]!.AsArray()
            .Where(field => field!["name"]!.GetValue<string>() == name)
            .ToList();
        if (matches.Count != 1)
            throw new InvalidOperationException("operation derivation validator target field invalid");
        return matches[0]!["value"]!.DeepClone();
    }

    private static void ExactKeys(JsonObject value, IReadOnlyCollection<string> expected, string label)
    {
        if (value.Count != expected.Count || expected.Any(name => !value.ContainsKey(name)))
            throw new InvalidOperationException($"operation derivation {label} fields invalid");
    }

    private static void RequireTyped(JsonObject typed, string? expectedSchema = null)
    {
        ExactKeys(typed, ["schema", "value", "sha256"], "typed value");
        if (
            typed["sha256"]!.GetValue<string>() != Sha(Canonical(typed["value"])) ||
            (expectedSchema is not null && typed["schema"]!["id"]!.GetValue<string>() != expectedSchema))
            throw new InvalidOperationException("operation derivation typed value invalid");
    }

    private static void RequireClaim(
        JsonArray claims,
        JsonObject expectedSchema,
        JsonArray expectedValue,
        string label)
    {
        if (claims.Count != 1)
            throw new InvalidOperationException($"operation derivation {label} cardinality invalid");
        var claim = claims[0]!.AsObject();
        RequireTyped(claim);
        if (
            Canonical(claim["schema"]) != Canonical(expectedSchema) ||
            Canonical(claim["value"]) != Canonical(expectedValue))
            throw new InvalidOperationException($"operation derivation {label} authority mismatch");
    }

    private static void RequireNodeRef(JsonObject reference)
    {
        ExactKeys(reference, ["nodeId", "contract", "sha256"], "node ref");
        RequireText(reference, "nodeId");
        var contract = reference["contract"]!.AsObject();
        ExactKeys(contract, ["id", "sha256"], "schema ref");
        RequireText(contract, "id");
        RequireSha(contract["sha256"], "schema ref");
        RequireSha(reference["sha256"], "node ref");
    }

    private static void RequireText(JsonObject value, string field)
    {
        if (string.IsNullOrWhiteSpace(value[field]!.GetValue<string>()))
            throw new InvalidOperationException($"operation derivation {field} invalid");
    }

    private static void RequireSha(JsonNode? value, string label)
    {
        var sha256 = value!.GetValue<string>();
        if (sha256.Length != 64 || sha256.Any(character =>
                character is not (>= '0' and <= '9') and not (>= 'a' and <= 'f')))
            throw new InvalidOperationException($"operation derivation {label} hash invalid");
    }

    private static void RequireArtifact(JsonObject artifact, string label)
    {
        ExactKeys(artifact, ["artifactVersionId", "path", "bytesSha256", "mediaType"], $"{label} artifact");
        var path = artifact["path"]!.GetValue<string>();
        if (
            !Path.IsPathFullyQualified(path) ||
            !File.Exists(path) ||
            FileSha(path) != artifact["bytesSha256"]!.GetValue<string>())
            throw new InvalidOperationException($"operation derivation {label} artifact invalid");
    }

    private static JsonObject Typed(JsonObject schema, JsonNode value) =>
        new()
        {
            ["schema"] = schema.DeepClone(),
            ["value"] = value.DeepClone(),
            ["sha256"] = Sha(Canonical(value))
        };

    private static JsonObject ContractRef(string file, string id)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        if (!File.Exists(path)) throw new InvalidOperationException($"provider contract missing: {file}");
        return new JsonObject { ["id"] = id, ["sha256"] = Sha(File.ReadAllText(path)) };
    }

    private static string Argument(string[] args, string name)
    {
        var index = Array.IndexOf(args, $"--{name}");
        if (index < 0 || index + 1 >= args.Length)
            throw new InvalidOperationException($"operation derivation {name} required");
        return Path.GetFullPath(args[index + 1]);
    }

    private static JsonNode Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))
        ?? throw new InvalidOperationException($"operation derivation JSON invalid: {path}");

    private static void Write(string path, JsonNode value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, $"{Canonical(value)}\n");
    }

    private static string Sha(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private static string FileSha(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

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
