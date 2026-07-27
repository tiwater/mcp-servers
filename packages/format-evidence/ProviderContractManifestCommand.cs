using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

public static class ProviderContractManifestCommand
{
    public sealed record Contract(
        string ProviderId,
        string ProviderVersion,
        string EffectType,
        string EffectTypeVersion,
        string OperationSchemaFile,
        string OperationSchemaId,
        string ReceiptSchemaFile,
        string ReceiptSchemaId,
        string DerivationProducerCommand,
        string DerivationValidatorCommand,
        string ExecutionProducerCommand,
        string ExecutionValidatorCommand,
        string AdapterId,
        string AdapterVersion);

    private sealed record ContractSpec(string Role, string File, string Id);

    public static int RunProducer(string[] args, Contract contract)
    {
        var output = Required(args, "--output");
        EnsureFresh(output);
        var manifest = new JsonObject
        {
            ["schema"] = "tiwater.provider-contract-manifest/v1",
            ["provider"] = Identity(contract.ProviderId, contract.ProviderVersion),
            ["effectType"] = Identity(contract.EffectType, contract.EffectTypeVersion),
            ["contracts"] = new JsonArray(Specs(contract).Select(spec => (JsonNode)new JsonObject
            {
                ["role"] = spec.Role,
                ["schema"] = SchemaRef(spec.File, spec.Id)
            }).ToArray()),
            ["ports"] = new JsonArray(
                Port("format-observation", contract, "inspect-evidence-v2", "validate-inspect-evidence-v2"),
                Port("operation-derivation", contract, contract.DerivationProducerCommand, contract.DerivationValidatorCommand),
                Port("effect-execution", contract, contract.ExecutionProducerCommand, contract.ExecutionValidatorCommand)),
            ["executionAdapter"] = Identity(contract.AdapterId, contract.AdapterVersion)
        };
        manifest["manifestSha256"] = Sha(Canonical(manifest));
        Write(output, manifest);
        return 0;
    }

    public static int RunValidator(string[] args, Contract contract)
    {
        var manifestPath = Required(args, "--manifest");
        var output = Required(args, "--output");
        EnsureFresh(output);
        var manifest = JsonNode.Parse(File.ReadAllText(manifestPath))?.AsObject()
            ?? throw new InvalidOperationException("provider contract manifest is not an object");
        var findings = Validate(manifest, contract);
        Write(output, new JsonObject
        {
            ["schema"] = "tiwater.provider-contract-manifest-verdict/v1",
            ["manifestSha256"] = Sha(File.ReadAllBytes(manifestPath)),
            ["validator"] = Identity($"{contract.ProviderId}:{contract.EffectType}:manifest-validator", contract.ProviderVersion),
            ["decision"] = findings.Count == 0 ? "pass" : "fail",
            ["findings"] = new JsonArray(findings.Select(item => (JsonNode)new JsonObject
            {
                ["code"] = item.Code,
                ["message"] = item.Message
            }).ToArray())
        });
        return findings.Count == 0 ? 0 : 1;
    }

    private static List<(string Code, string Message)> Validate(JsonObject manifest, Contract contract)
    {
        var findings = new List<(string, string)>();
        Check(manifest["schema"]?.GetValue<string>() == "tiwater.provider-contract-manifest/v1",
            "manifest-schema", "manifest schema identity is wrong", findings);
        ValidateIdentity(manifest["provider"], contract.ProviderId, contract.ProviderVersion, "provider", findings);
        ValidateIdentity(manifest["effectType"], contract.EffectType, contract.EffectTypeVersion, "effect-type", findings);
        ValidateIdentity(manifest["executionAdapter"], contract.AdapterId, contract.AdapterVersion, "execution-adapter", findings);
        ValidateContracts(manifest["contracts"], contract, findings);
        ValidatePorts(manifest["ports"], contract, findings);
        var clone = manifest.DeepClone().AsObject();
        clone.Remove("manifestSha256");
        Check(manifest["manifestSha256"]?.GetValue<string>() == Sha(Canonical(clone)),
            "manifest-hash-mismatch", "manifestSha256 does not bind the exact manifest", findings);
        return findings;
    }

    private static void ValidateContracts(JsonNode? node, Contract contract, List<(string Code, string Message)> findings)
    {
        if (node is not JsonArray actual)
        {
            findings.Add(("contracts-invalid", "contracts must be an array"));
            return;
        }
        var expected = Specs(contract);
        Check(actual.Count == expected.Count, "contracts-count-mismatch",
            "contracts must contain every required role exactly once", findings);
        foreach (var spec in expected)
        {
            var matches = actual.OfType<JsonObject>()
                .Where(item => item["role"]?.GetValue<string>() == spec.Role).ToArray();
            Check(matches.Length == 1, "contract-role-missing",
                $"contract role {spec.Role} must occur exactly once", findings);
            if (matches.Length == 1)
                Check(JsonNode.DeepEquals(matches[0]["schema"], SchemaRef(spec.File, spec.Id)),
                    "contract-schema-mismatch", $"contract role {spec.Role} has wrong schema bytes or identity", findings);
        }
    }

    private static void ValidatePorts(JsonNode? node, Contract contract, List<(string Code, string Message)> findings)
    {
        if (node is not JsonArray actual)
        {
            findings.Add(("ports-invalid", "ports must be an array"));
            return;
        }
        var expected = new[]
        {
            ("format-observation", "inspect-evidence-v2", "validate-inspect-evidence-v2"),
            ("operation-derivation", contract.DerivationProducerCommand, contract.DerivationValidatorCommand),
            ("effect-execution", contract.ExecutionProducerCommand, contract.ExecutionValidatorCommand)
        };
        Check(actual.Count == expected.Length, "ports-count-mismatch",
            "ports must contain the three provider boundaries", findings);
        foreach (var (role, producer, validator) in expected)
        {
            var matches = actual.OfType<JsonObject>()
                .Where(item => item["role"]?.GetValue<string>() == role).ToArray();
            Check(matches.Length == 1, "port-role-missing", $"port role {role} must occur exactly once", findings);
            if (matches.Length != 1) continue;
            ValidatePort(matches[0]["producer"], contract, role, "producer", producer, findings);
            ValidatePort(matches[0]["validator"], contract, role, "validator", validator, findings);
        }
    }

    private static void ValidatePort(
        JsonNode? node,
        Contract contract,
        string role,
        string direction,
        string command,
        List<(string Code, string Message)> findings)
    {
        var valid = node is JsonObject port && port.Count == 3 &&
            port["id"]?.GetValue<string>() == $"{contract.ProviderId}:{contract.EffectType}:{role}:{direction}" &&
            port["version"]?.GetValue<string>() == contract.ProviderVersion &&
            port["command"]?.GetValue<string>() == command;
        Check(valid, "port-identity-mismatch", $"{role} {direction} does not match the package", findings);
    }

    private static void ValidateIdentity(
        JsonNode? node,
        string id,
        string version,
        string role,
        List<(string Code, string Message)> findings)
    {
        var valid = node is JsonObject value && value.Count == 2 &&
            value["id"]?.GetValue<string>() == id &&
            value["version"]?.GetValue<string>() == version;
        Check(valid, $"{role}-identity-mismatch", $"{role} identity does not match the package", findings);
    }

    private static void Check(
        bool condition,
        string code,
        string message,
        List<(string Code, string Message)> findings)
    {
        if (!condition) findings.Add((code, message));
    }

    private static IReadOnlyList<ContractSpec> Specs(Contract contract) =>
    [
        new("manifest", "tiwater.provider-contract-manifest-v1.schema.json", "tiwater.provider-contract-manifest/v1"),
        new("manifest-verdict", "tiwater.provider-contract-manifest-verdict-v1.schema.json", "tiwater.provider-contract-manifest-verdict/v1"),
        new("format-evidence-request", "tiwater.format-evidence-request-v2.schema.json", "tiwater.format-evidence-request/v2"),
        new("format-evidence", "tiwater.format-evidence-v2.schema.json", "tiwater.format-evidence/v2"),
        new("format-evidence-verdict", "tiwater.format-evidence-verdict-v2.schema.json", "tiwater.format-evidence-verdict/v2"),
        new("effect-intent", "tiwater.provider-effect-intent-v1.schema.json", "tiwater.provider-effect-intent/v1"),
        new("resource-set", "tiwater.provider-resource-set-v1.schema.json", "tiwater.provider-resource-set/v1"),
        new("write-set", "tiwater.provider-write-set-v1.schema.json", "tiwater.provider-write-set/v1"),
        new("derivation-request-v1-compatibility", "tiwater.operation-derivation-request-v1.schema.json", "tiwater.operation-derivation-request/v1"),
        new("derivation-request", "tiwater.operation-derivation-request-v2.schema.json", "tiwater.operation-derivation-request/v2"),
        new("derivation-result", "tiwater.operation-derivation-result-v1.schema.json", "tiwater.operation-derivation-result/v1"),
        new("derivation-verdict", "tiwater.operation-derivation-verdict-v1.schema.json", "tiwater.operation-derivation-verdict/v1"),
        new("derivation-provenance", "tiwater.operation-derivation-provenance-v1.schema.json", "tiwater.operation-derivation-provenance/v1"),
        new("canonical-node", "lucid.canonical-node-v2.schema.json", "lucid.canonical-node/v2"),
        new("operator-verdict", "lucid.operator-verdict-v1.schema.json", "lucid.operator-verdict/v1"),
        new("effect-bundle", "lucid.effect-bundle-v3.schema.json", "lucid.effect-bundle/v3"),
        new("composed-effect", "lucid.composed-effect-v2.schema.json", "lucid.composed-effect/v2"),
        new("effect-execution-request", "lucid.effect-execution-request-v1.schema.json", "lucid.effect-execution-request/v1"),
        new("provider-execution-request", "tiwater.provider-effect-execution-request-v1.schema.json", "tiwater.provider-effect-execution-request/v1"),
        new("operation", contract.OperationSchemaFile, contract.OperationSchemaId),
        new("receipt", contract.ReceiptSchemaFile, contract.ReceiptSchemaId),
        new("execution-evidence", "lucid.execution-evidence-v2.schema.json", "lucid.execution-evidence/v2"),
        new("artifact-lineage", "tiwater.provider-artifact-lineage-v1.schema.json", "tiwater.provider-artifact-lineage/v1"),
        new("execution-evidence-verdict", "tiwater.execution-evidence-verdict-v1.schema.json", "tiwater.execution-evidence-verdict/v1")
    ];

    private static JsonObject Port(string role, Contract contract, string producer, string validator) =>
        new()
        {
            ["role"] = role,
            ["producer"] = new JsonObject
            {
                ["id"] = $"{contract.ProviderId}:{contract.EffectType}:{role}:producer",
                ["version"] = contract.ProviderVersion,
                ["command"] = producer
            },
            ["validator"] = new JsonObject
            {
                ["id"] = $"{contract.ProviderId}:{contract.EffectType}:{role}:validator",
                ["version"] = contract.ProviderVersion,
                ["command"] = validator
            }
        };

    private static JsonObject Identity(string id, string version) =>
        new() { ["id"] = id, ["version"] = version };

    private static JsonObject SchemaRef(string file, string id)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        if (!File.Exists(path)) throw new InvalidOperationException($"provider contract missing: {file}");
        return new JsonObject { ["id"] = id, ["sha256"] = Sha(File.ReadAllBytes(path)) };
    }

    private static string Required(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        if (index < 0 || index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]))
            throw new InvalidOperationException($"{name} is required");
        return args[index + 1];
    }

    private static void EnsureFresh(string path)
    {
        if (File.Exists(path)) throw new InvalidOperationException("provider contract manifest output must be fresh");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
    }

    private static void Write(string path, JsonNode node) =>
        File.WriteAllText(path, node.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) + "\n");

    private static string Canonical(JsonNode node) =>
        node switch
        {
            JsonObject value => "{" + string.Join(",", value.OrderBy(item => item.Key, StringComparer.Ordinal)
                .Select(item => JsonSerializer.Serialize(item.Key) + ":" + Canonical(item.Value!))) + "}",
            JsonArray value => "[" + string.Join(",", value.Select(item => Canonical(item!))) + "]",
            _ => node.ToJsonString()
        };

    private static string Sha(string value) => Sha(Encoding.UTF8.GetBytes(value));
    private static string Sha(byte[] value) => Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();
}
