using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Convert;

/// <summary>
/// Producer and independent validator for the closed
/// tiwater.convert-provider-contract-manifest/v1 contract. The manifest binds every render
/// contract schema byte hash, the render producer/validator command identities, and the
/// runtime version of this package; the validator recomputes all of it from the deployed
/// package.
/// </summary>
public static class ConvertProviderContractManifestCommand
{
    private const string ManifestSchemaId = "tiwater.convert-provider-contract-manifest/v1";
    private const string VerdictSchemaId = "tiwater.convert-provider-contract-manifest-verdict/v1";
    private const string ProviderId = "tiwater-convert";

    private sealed record ContractSpec(string Role, string File, string Id);

    private sealed record PortSpec(string Role, string ProducerId, string ProducerCommand, string ValidatorId, string ValidatorCommand);

    private static readonly ContractSpec[] Specs =
    [
        new("render-request", "render-request-v1.schema.json", "tiwater.convert-render-request/v1"),
        new("render-result", "render-result-v1.schema.json", "tiwater.convert-render-result/v1"),
        new("render-verdict", "render-verdict-v1.schema.json", "tiwater.convert-render-verdict/v1"),
        new("native-render-provenance", "native-render-provenance-v1.schema.json", "tiwater.convert-native-render-provenance/v1"),
        new("provider-contract-manifest", "provider-contract-manifest-v1.schema.json", ManifestSchemaId),
        new("provider-contract-manifest-verdict", "provider-contract-manifest-verdict-v1.schema.json", VerdictSchemaId)
    ];

    private static readonly PortSpec[] Ports =
    [
        new("render",
            "tiwater-convert-render-producer", "render",
            "tiwater-convert-render-validator", "validate-render-evidence"),
        new("provider-contract-manifest",
            "tiwater-convert:provider-contract-manifest:producer", "provider-contract-manifest",
            "tiwater-convert:provider-contract-manifest:validator", "validate-provider-contract-manifest")
    ];

    public static int RunProducer(string[] args)
    {
        var output = Required(args, "--output");
        EnsureFresh(output);
        var manifest = new JsonObject
        {
            ["schema"] = ManifestSchemaId,
            ["provider"] = Identity(),
            ["contracts"] = new JsonArray(Specs.Select(spec => (JsonNode)new JsonObject
            {
                ["role"] = spec.Role,
                ["schema_ref"] = SchemaRef(spec)
            }).ToArray()),
            ["ports"] = new JsonArray(Ports.Select(port => (JsonNode)new JsonObject
            {
                ["role"] = port.Role,
                ["producer"] = Command(port.ProducerId, port.ProducerCommand),
                ["validator"] = Command(port.ValidatorId, port.ValidatorCommand)
            }).ToArray()),
            ["runtime"] = Identity()
        };
        manifest["manifest_sha256"] = Sha(Canonical(manifest));
        Write(output, manifest);
        return 0;
    }

    public static int RunValidator(string[] args)
    {
        var manifestPath = Required(args, "--manifest");
        var output = Required(args, "--output");
        EnsureFresh(output);
        var manifest = JsonNode.Parse(File.ReadAllText(manifestPath)) as JsonObject
            ?? throw new InvalidOperationException("provider contract manifest is not an object");
        var findings = Validate(manifest);
        Write(output, new JsonObject
        {
            ["schema"] = VerdictSchemaId,
            ["manifest_sha256"] = Sha(File.ReadAllBytes(manifestPath)),
            ["validator"] = new JsonObject
            {
                ["id"] = "tiwater-convert:provider-contract-manifest:validator",
                ["version"] = RuntimeIdentity.Version
            },
            ["decision"] = findings.Count == 0 ? "pass" : "failed",
            ["findings"] = new JsonArray(findings.Select(finding => (JsonNode)new JsonObject
            {
                ["code"] = finding.Code,
                ["message"] = finding.Message
            }).ToArray())
        });
        return findings.Count == 0 ? 0 : 1;
    }

    private static List<(string Code, string Message)> Validate(JsonObject manifest)
    {
        var findings = new List<(string, string)>();
        Check(
            manifest.Count == 6
            && new[] { "schema", "provider", "contracts", "ports", "runtime", "manifest_sha256" }.All(manifest.ContainsKey),
            "manifest-fields", "manifest must close over exactly the declared fields", findings);
        Check(manifest["schema"] is JsonValue schemaValue
            && schemaValue.TryGetValue<string>(out var schema) && schema == ManifestSchemaId,
            "manifest-schema", "manifest schema identity is wrong", findings);
        ValidateIdentity(manifest["provider"], "provider", findings);
        ValidateIdentity(manifest["runtime"], "runtime", findings);
        ValidateContracts(manifest["contracts"], findings);
        ValidatePorts(manifest["ports"], findings);
        var clone = manifest.DeepClone().AsObject();
        clone.Remove("manifest_sha256");
        Check(
            manifest["manifest_sha256"] is JsonValue hashValue
            && hashValue.TryGetValue<string>(out var hash) && hash == Sha(Canonical(clone)),
            "manifest-hash-mismatch", "manifest_sha256 does not bind the exact manifest", findings);
        return findings;
    }

    private static void ValidateContracts(JsonNode? node, List<(string Code, string Message)> findings)
    {
        if (node is not JsonArray actual)
        {
            findings.Add(("contracts-invalid", "contracts must be an array"));
            return;
        }
        Check(actual.Count == Specs.Length, "contracts-count-mismatch",
            "contracts must contain every render contract role exactly once", findings);
        foreach (var spec in Specs)
        {
            var matches = actual.OfType<JsonObject>()
                .Where(item => item["role"] is JsonValue role
                    && role.TryGetValue<string>(out var name) && name == spec.Role)
                .ToArray();
            Check(matches.Length == 1, "contract-role-missing",
                $"contract role {spec.Role} must occur exactly once", findings);
            if (matches.Length == 1)
                Check(matches[0].Count == 2 && JsonNode.DeepEquals(matches[0]["schema_ref"], SchemaRef(spec)),
                    "contract-schema-mismatch", $"contract role {spec.Role} has wrong schema bytes or identity", findings);
        }
    }

    private static void ValidatePorts(JsonNode? node, List<(string Code, string Message)> findings)
    {
        if (node is not JsonArray actual)
        {
            findings.Add(("ports-invalid", "ports must be an array"));
            return;
        }
        Check(actual.Count == Ports.Length, "ports-count-mismatch",
            "ports must contain the render and provider-contract-manifest boundaries", findings);
        foreach (var port in Ports)
        {
            var matches = actual.OfType<JsonObject>()
                .Where(item => item["role"] is JsonValue role
                    && role.TryGetValue<string>(out var name) && name == port.Role)
                .ToArray();
            Check(matches.Length == 1, "port-role-missing", $"port role {port.Role} must occur exactly once", findings);
            if (matches.Length != 1) continue;
            Check(matches[0].Count == 3, "port-fields", $"port role {port.Role} fields are invalid", findings);
            Check(JsonNode.DeepEquals(matches[0]["producer"], Command(port.ProducerId, port.ProducerCommand)),
                "port-identity-mismatch", $"{port.Role} producer does not match the package", findings);
            Check(JsonNode.DeepEquals(matches[0]["validator"], Command(port.ValidatorId, port.ValidatorCommand)),
                "port-identity-mismatch", $"{port.Role} validator does not match the package", findings);
        }
    }

    private static void ValidateIdentity(JsonNode? node, string role, List<(string Code, string Message)> findings)
    {
        Check(node is JsonObject value && value.Count == 2
            && value["id"] is JsonValue idValue
            && idValue.TryGetValue<string>(out var id) && id == ProviderId
            && value["version"] is JsonValue versionValue
            && versionValue.TryGetValue<string>(out var version) && version == RuntimeIdentity.Version,
            $"{role}-identity-mismatch", $"{role} identity does not match the package", findings);
    }

    private static void Check(
        bool condition,
        string code,
        string message,
        List<(string Code, string Message)> findings)
    {
        if (!condition) findings.Add((code, message));
    }

    private static JsonObject Identity() =>
        new() { ["id"] = ProviderId, ["version"] = RuntimeIdentity.Version };

    private static JsonObject Command(string id, string command) =>
        new() { ["id"] = id, ["version"] = RuntimeIdentity.Version, ["command"] = command };

    private static JsonObject SchemaRef(ContractSpec spec)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "schemas", spec.File);
        if (!File.Exists(path)) throw new InvalidOperationException($"render contract schema missing: {spec.File}");
        return new JsonObject { ["id"] = spec.Id, ["sha256"] = Sha(File.ReadAllBytes(path)) };
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
        if (File.Exists(path)) throw new InvalidOperationException("provider contract manifest output must be fresh");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
    }

    private static void Write(string path, JsonNode node) =>
        File.WriteAllText(path, $"{Canonical(node)}\n");

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
