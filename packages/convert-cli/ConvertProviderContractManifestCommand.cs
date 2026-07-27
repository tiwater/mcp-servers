using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Convert;

/// <summary>
/// Producer and independent validator for the Set-15 lucid.provider-contract-manifest
/// contract. The manifest declares exactly one read-only render port bound to the real
/// render / validate-render-evidence commands and the convert render contract family;
/// it never declares observation, derivation, reobservation, or execution ports (those
/// belong to the Office and PDF providers). manifestId is recomputed over the
/// provider/runtime/port declarations and the deployed bytes of every declared contract
/// schema, so the validator re-derives all of it from the deployed package rather than
/// trusting the manifest under validation.
/// </summary>
public static class ConvertProviderContractManifestCommand
{
    private const string ManifestSchemaId = "lucid.provider-contract-manifest";
    private const string PackagedManifestSchemaId = "tiwater.convert-provider-contract-manifest/v1";
    private const string VerdictSchemaId = "tiwater.convert-provider-contract-manifest-verdict/v1";
    private const string ProviderId = "tiwater-convert";
    private const string RenderProducerId = "tiwater-convert-render-producer";
    private const string RenderValidatorId = "tiwater-convert-render-validator";
    private const string ManifestValidatorId = "tiwater-convert:provider-contract-manifest:validator";
    private const string RenderRequestSchemaId = "tiwater.convert-render-request/v1";
    private const string RenderResultSchemaId = "tiwater.convert-render-result/v1";
    private const string RenderVerdictSchemaId = "tiwater.convert-render-verdict/v1";

    private static readonly string[] ManifestFields =
        ["schema", "schemaSetVersion", "manifestId", "provider", "runtime", "ports"];

    private static readonly string[] ForbiddenKinds = ["observe", "derive", "reobserve", "execute"];

    private sealed record ContractSpec(string File, string Id);

    // Declared contract family: every schema whose deployed bytes manifestId binds.
    private static readonly ContractSpec[] ContractSchemas =
    [
        new("render-request-v1.schema.json", RenderRequestSchemaId),
        new("render-result-v1.schema.json", RenderResultSchemaId),
        new("render-verdict-v1.schema.json", RenderVerdictSchemaId),
        new("native-render-provenance-v1.schema.json", "tiwater.convert-native-render-provenance/v1"),
        new("provider-contract-manifest-v1.schema.json", PackagedManifestSchemaId),
        new("provider-contract-manifest-verdict-v1.schema.json", VerdictSchemaId)
    ];

    public static int RunProducer(string[] args)
    {
        var output = RequiredPath(args, "--output");
        var schemaSetVersion = RequiredSchemaSetVersion(args);
        EnsureFresh(output);
        Write(output, BuildManifest(schemaSetVersion));
        return 0;
    }

    public static int RunValidator(string[] args)
    {
        var manifestPath = RequiredPath(args, "--manifest");
        var output = RequiredPath(args, "--output");
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
                ["id"] = ManifestValidatorId,
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

    private static JsonObject BuildManifest(int schemaSetVersion)
    {
        var producer = Identity(RenderProducerId);
        var validator = Identity(RenderValidatorId);
        var port = new JsonObject
        {
            ["kind"] = "render",
            ["producer"] = WithAdapterIdentity(producer),
            ["validator"] = WithAdapterIdentity(validator),
            ["requestSchema"] = RenderRequestSchemaId,
            ["validatorRequestSchema"] = RenderRequestSchemaId,
            ["resultSchema"] = RenderResultSchemaId,
            ["verdictSchema"] = RenderVerdictSchemaId,
            // tiwater.convert-render-request/v1 declares an explicitly empty option set.
            ["options"] = new JsonArray(),
            ["cacheKeyComposition"] = new JsonArray(
                "schemaSetVersion", "bytesSha256", "provider", "optionsSha256", "rendererBackend"),
            ["resourceDeclarations"] = new JsonArray((JsonNode)new JsonObject
            {
                ["resourceKey"] = "convert:document",
                ["access"] = "read"
            }),
            ["sideEffect"] = new JsonObject { ["kind"] = "read-only", ["idempotent"] = true },
            ["attemptBudget"] = 1
        };
        var manifest = new JsonObject
        {
            ["schema"] = ManifestSchemaId,
            ["schemaSetVersion"] = schemaSetVersion,
            ["provider"] = Identity(ProviderId),
            ["runtime"] = Identity(ProviderId),
            ["ports"] = new JsonArray(port)
        };
        var material = new JsonObject
        {
            ["contractSchemas"] = new JsonArray(ContractSchemas
                .OrderBy(spec => spec.Id, StringComparer.Ordinal)
                .Select(spec => (JsonNode)SchemaRef(spec))
                .ToArray()),
            ["manifest"] = manifest.DeepClone()
        };
        manifest["manifestId"] = $"manifest-{Sha(Canonical(material))}";
        return manifest;
    }

    private static List<(string Code, string Message)> Validate(JsonObject manifest)
    {
        var findings = new List<(string, string)>();
        Check(
            manifest.Count == ManifestFields.Length && ManifestFields.All(manifest.ContainsKey),
            "manifest-fields", "manifest must close over exactly schema, schemaSetVersion, manifestId, provider, runtime, ports", findings);
        Check(manifest["schema"] is JsonValue schemaValue
            && schemaValue.TryGetValue<string>(out var schema) && schema == ManifestSchemaId,
            "manifest-schema", "manifest schema identity is wrong", findings);
        var schemaSetVersion = ReadSchemaSetVersion(manifest, findings);
        if (manifest["ports"] is JsonArray declaredPorts)
        {
            var forbidden = declaredPorts.OfType<JsonObject>()
                .Select(port => port["kind"] is JsonValue kindValue
                    && kindValue.TryGetValue<string>(out var kind) ? kind : null)
                .Where(kind => kind is not null && ForbiddenKinds.Contains(kind))
                .ToArray();
            Check(forbidden.Length == 0, "port-kind-forbidden",
                "convert declares render only: observation, derivation, reobservation, and execution ports belong to the Office and PDF providers", findings);
        }
        if (findings.Count > 0 || schemaSetVersion is null) return findings;
        var expected = BuildManifest(schemaSetVersion.Value);
        Check(JsonNode.DeepEquals(manifest["provider"], expected["provider"]),
            "provider-identity-mismatch", "provider identity does not match the package", findings);
        Check(JsonNode.DeepEquals(manifest["runtime"], expected["runtime"]),
            "runtime-identity-mismatch", "runtime identity does not match the package", findings);
        Check(JsonNode.DeepEquals(manifest["ports"], expected["ports"]),
            "port-declaration-mismatch",
            "render port declaration does not match the package: expected exactly one read-only render port "
            + "with the deployed adapter identities, render schema family, empty options, cache composition, "
            + "resource declarations, and attempt budget", findings);
        Check(manifest["manifestId"] is JsonValue idValue
            && idValue.TryGetValue<string>(out var manifestId) && manifestId == expected["manifestId"]!.GetValue<string>(),
            "manifest-id-mismatch",
            "manifestId does not bind the package identities, port declarations, and deployed contract schema bytes", findings);
        return findings;
    }

    private static int? ReadSchemaSetVersion(JsonObject manifest, List<(string Code, string Message)> findings)
    {
        if (manifest["schemaSetVersion"] is JsonValue value
            && value.TryGetValue<int>(out var schemaSetVersion) && schemaSetVersion >= 1)
            return schemaSetVersion;
        Check(false, "schema-set-version", "schemaSetVersion must be a positive integer", findings);
        return null;
    }

    private static void Check(
        bool condition,
        string code,
        string message,
        List<(string Code, string Message)> findings)
    {
        if (!condition) findings.Add((code, message));
    }

    private static JsonObject Identity(string id) =>
        new() { ["id"] = id, ["version"] = RuntimeIdentity.Version };

    private static JsonObject WithAdapterIdentity(JsonObject identity)
    {
        var adapterIdentity = identity.DeepClone();
        var node = identity.DeepClone().AsObject();
        node["adapterIdentity"] = adapterIdentity;
        return node;
    }

    private static JsonObject SchemaRef(ContractSpec spec)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "schemas", spec.File);
        if (!File.Exists(path)) throw new InvalidOperationException($"render contract schema missing: {spec.File}");
        return new JsonObject { ["id"] = spec.Id, ["sha256"] = Sha(File.ReadAllBytes(path)) };
    }

    private static string RequiredPath(string[] args, string name) =>
        Path.GetFullPath(Required(args, name));

    private static int RequiredSchemaSetVersion(string[] args)
    {
        var raw = Required(args, "--schema-set-version");
        if (!int.TryParse(raw, out var value) || value < 1)
            throw new InvalidOperationException("--schema-set-version must be a positive integer");
        return value;
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
