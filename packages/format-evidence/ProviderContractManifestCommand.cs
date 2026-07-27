using System.Reflection;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

public static class ProviderContractManifestCommand
{
    // Directly admissible Lucid schema-set manifest shape (issue #89): the same
    // command publishes tiwater.provider-contract-manifest/v1 by default and the
    // closed lucid.provider-contract-manifest shape when --format/--schema-set-version
    // are given. Consumers distinguish the two authorities by the schema const.
    public const string SetManifestSchemaId = "lucid.provider-contract-manifest";

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
        string AdapterVersion,
        string Format);

    private sealed record ContractSpec(string Role, string File, string Id);

    // Every schema name a lucid.provider-contract-manifest port may declare must
    // resolve to a real contract schema deployed with the package (packed into the
    // nupkg from format-evidence/contracts). Unmapped or missing files fail closed.
    private static readonly IReadOnlyDictionary<string, string> SetSchemaFiles =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["tiwater.format-evidence-request/v2"] = "tiwater.format-evidence-request-v2.schema.json",
            ["tiwater.format-evidence/v2"] = "tiwater.format-evidence-v2.schema.json",
            ["tiwater.format-evidence-verdict/v2"] = "tiwater.format-evidence-verdict-v2.schema.json",
            ["tiwater.format-extraction-options/v1"] = "tiwater.format-extraction-options-v1.schema.json",
            ["tiwater.operation-derivation-request/v2"] = "tiwater.operation-derivation-request-v2.schema.json",
            ["tiwater.operation-derivation-result/v1"] = "tiwater.operation-derivation-result-v1.schema.json",
            ["tiwater.operation-derivation-verdict/v1"] = "tiwater.operation-derivation-verdict-v1.schema.json",
            ["tiwater.provider-effect-execution-request/v1"] = "tiwater.provider-effect-execution-request-v1.schema.json",
            ["lucid.execution-evidence/v2"] = "lucid.execution-evidence-v2.schema.json",
            ["tiwater.execution-evidence-verdict/v1"] = "tiwater.execution-evidence-verdict-v1.schema.json"
        };

    public static int RunProducer(string[] args, Contract contract)
    {
        var output = Required(args, "--output");
        EnsureFresh(output);
        var format = Optional(args, "--format");
        var schemaSetVersionText = Optional(args, "--schema-set-version");
        if (format is not null || schemaSetVersionText is not null)
        {
            if (format != SetManifestSchemaId)
                throw new InvalidOperationException(
                    "--format lucid.provider-contract-manifest is required to publish a schema set manifest");
            if (!int.TryParse(schemaSetVersionText, out var schemaSetVersion) || schemaSetVersion < 1)
                throw new InvalidOperationException(
                    "--schema-set-version <positive integer> is required to publish a schema set manifest");
            Write(output, SetManifest(contract, schemaSetVersion));
            return 0;
        }
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
        var findings = manifest["schema"]?.GetValue<string>() == SetManifestSchemaId
            ? ValidateSetManifest(manifest, contract, RequiredSchemaSetVersion(args))
            : Validate(manifest, contract);
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

    private static int RequiredSchemaSetVersion(string[] args)
    {
        if (!int.TryParse(Optional(args, "--schema-set-version"), out var schemaSetVersion) || schemaSetVersion < 1)
            throw new InvalidOperationException(
                "--schema-set-version <positive integer> is required to validate a schema set manifest");
        return schemaSetVersion;
    }

    private static JsonObject SetManifest(Contract contract, int schemaSetVersion) => new()
    {
        ["schema"] = SetManifestSchemaId,
        ["schemaSetVersion"] = schemaSetVersion,
        ["manifestId"] = $"manifest:{contract.ProviderId}:{contract.ProviderVersion}:set-{schemaSetVersion}",
        ["provider"] = Identity(contract.ProviderId, contract.ProviderVersion),
        ["runtime"] = RuntimeIdentity(),
        ["ports"] = new JsonArray(
            SetObservationPort(contract, "observe"),
            SetObservationPort(contract, "reobserve"),
            SetDerivationPort(contract),
            SetExecutionPort(contract))
    };

    private static JsonObject SetObservationPort(Contract contract, string kind)
    {
        // Reobservation reuses the format-observation commands and is declared
        // explicitly as its own reobserve port (issue #89); the observe cache key
        // is exactly content + options + provider, the reobserve key additionally
        // binds the schema set version.
        var cache = kind == "observe"
            ? new JsonArray("bytesSha256", "optionsSha256", "provider")
            : new JsonArray("bytesSha256", "optionsSha256", "provider", "schemaSetVersion");
        return new JsonObject
        {
            ["kind"] = kind,
            ["producer"] = SetAdapter(contract.ProviderId, contract.ProviderVersion, "inspect-evidence-v2", contract.ProviderVersion),
            ["validator"] = SetAdapter($"{contract.ProviderId}-validator", contract.ProviderVersion, "validate-inspect-evidence-v2", contract.ProviderVersion),
            ["requestSchema"] = DeployedSchema("tiwater.format-evidence-request/v2"),
            ["validatorRequestSchema"] = DeployedSchema("tiwater.format-evidence-request/v2"),
            ["resultSchema"] = DeployedSchema("tiwater.format-evidence/v2"),
            ["verdictSchema"] = DeployedSchema("tiwater.format-evidence-verdict/v2"),
            ["options"] = new JsonArray(new JsonObject
            {
                ["name"] = "facets",
                ["valueSchema"] = DeployedSchema("tiwater.format-extraction-options/v1")
            }),
            ["cacheKeyComposition"] = cache,
            ["resourceDeclarations"] = new JsonArray(new JsonObject
            {
                ["resourceKey"] = $"{contract.Format}-document-bytes",
                ["access"] = "read"
            }),
            ["sideEffect"] = new JsonObject { ["kind"] = "read-only", ["idempotent"] = true },
            ["attemptBudget"] = 1
        };
    }

    private static JsonObject SetDerivationPort(Contract contract) => new()
    {
        ["kind"] = "derive",
        ["producer"] = SetAdapter(contract.ProviderId, contract.ProviderVersion, contract.DerivationProducerCommand, contract.ProviderVersion),
        ["validator"] = SetAdapter($"{contract.ProviderId}.operation-derivation-validator", contract.ProviderVersion, contract.DerivationValidatorCommand, contract.ProviderVersion),
        ["requestSchema"] = DeployedSchema("tiwater.operation-derivation-request/v2"),
        ["validatorRequestSchema"] = DeployedSchema("tiwater.operation-derivation-request/v2"),
        ["resultSchema"] = DeployedSchema("tiwater.operation-derivation-result/v1"),
        ["verdictSchema"] = DeployedSchema("tiwater.operation-derivation-verdict/v1"),
        ["options"] = new JsonArray(),
        ["cacheKeyComposition"] = new JsonArray("bytesSha256", "optionsSha256", "provider", "schemaSetVersion"),
        ["resourceDeclarations"] = new JsonArray(new JsonObject
        {
            ["resourceKey"] = $"{contract.Format}-document-bytes",
            ["access"] = "read"
        }),
        ["sideEffect"] = new JsonObject { ["kind"] = "read-only", ["idempotent"] = true },
        ["attemptBudget"] = 1
    };

    private static JsonObject SetExecutionPort(Contract contract) => new()
    {
        ["kind"] = "execute",
        ["producer"] = SetAdapter(contract.ProviderId, contract.ProviderVersion, contract.ExecutionProducerCommand, contract.ProviderVersion),
        ["validator"] = SetAdapter($"{contract.ProviderId}.execution-evidence-validator", contract.ProviderVersion, contract.ExecutionValidatorCommand, contract.ProviderVersion),
        ["requestSchema"] = DeployedSchema("tiwater.provider-effect-execution-request/v1"),
        ["validatorRequestSchema"] = DeployedSchema("tiwater.provider-effect-execution-request/v1"),
        ["resultSchema"] = DeployedSchema("lucid.execution-evidence/v2"),
        ["verdictSchema"] = DeployedSchema("tiwater.execution-evidence-verdict/v1"),
        ["options"] = new JsonArray(),
        ["cacheKeyComposition"] = new JsonArray("bytesSha256", "optionsSha256", "provider", "schemaSetVersion"),
        ["resourceDeclarations"] = new JsonArray(new JsonObject
        {
            ["resourceKey"] = $"{contract.Format}-document-bytes",
            ["access"] = "exclusive-write"
        }),
        ["sideEffect"] = new JsonObject { ["kind"] = "mutating", ["idempotent"] = false },
        ["attemptBudget"] = 1
    };

    private static JsonObject SetAdapter(string id, string version, string command, string commandVersion) =>
        new()
        {
            ["id"] = id,
            ["version"] = version,
            ["adapterIdentity"] = new JsonObject { ["id"] = command, ["version"] = commandVersion }
        };

    // The runtime this package honestly runs on: the .NET target framework the
    // deployed assembly was built for (net9.0), read from the assembly attribute
    // rather than hard-coded; producer and independent validator probe it the
    // same way on the same deployment.
    public static JsonObject RuntimeIdentity()
    {
        var framework = typeof(ProviderContractManifestCommand).Assembly
            .GetCustomAttribute<TargetFrameworkAttribute>()?.FrameworkName
            ?? throw new InvalidOperationException("provider runtime identity is unavailable");
        const string prefix = ".NETCoreApp,Version=v";
        if (!framework.StartsWith(prefix, StringComparison.Ordinal) || framework.Length == prefix.Length)
            throw new InvalidOperationException("provider runtime identity is unavailable");
        return Identity("dotnet", framework[prefix.Length..]);
    }

    private static string DeployedSchema(string schemaId)
    {
        if (!SetSchemaFiles.TryGetValue(schemaId, out var file))
            throw new InvalidOperationException($"provider contract schema unmapped: {schemaId}");
        if (!File.Exists(Path.Combine(AppContext.BaseDirectory, "contracts", file)))
            throw new InvalidOperationException($"provider contract schema not deployed: {file}");
        return schemaId;
    }

    private static List<(string Code, string Message)> ValidateSetManifest(
        JsonObject manifest,
        Contract contract,
        int schemaSetVersion)
    {
        var findings = new List<(string, string)>();
        Check(manifest.Count == 6 &&
            new[] { "schema", "schemaSetVersion", "manifestId", "provider", "runtime", "ports" }.All(manifest.ContainsKey),
            "manifest-fields", "manifest must close over exactly the lucid.provider-contract-manifest fields", findings);
        Check(manifest["schemaSetVersion"] is JsonValue versionValue &&
            versionValue.TryGetValue<int>(out var declaredVersion) && declaredVersion == schemaSetVersion,
            "manifest-schema-set-version-mismatch", "schemaSetVersion does not match the validation admission", findings);
        Check(manifest["manifestId"]?.GetValue<string>() ==
            $"manifest:{contract.ProviderId}:{contract.ProviderVersion}:set-{schemaSetVersion}",
            "manifest-id-mismatch", "manifestId does not match the package", findings);
        ValidateIdentity(manifest["provider"], contract.ProviderId, contract.ProviderVersion, "provider", findings);
        var runtime = RuntimeIdentity();
        ValidateIdentity(
            manifest["runtime"],
            runtime["id"]!.GetValue<string>(),
            runtime["version"]!.GetValue<string>(),
            "runtime",
            findings);
        ValidateSetPorts(manifest["ports"], contract, findings);
        return findings;
    }

    private static void ValidateSetPorts(JsonNode? node, Contract contract, List<(string Code, string Message)> findings)
    {
        if (node is not JsonArray actual)
        {
            findings.Add(("ports-invalid", "ports must be an array"));
            return;
        }
        Check(actual.Count == 4, "ports-count-mismatch",
            "ports must declare observe, reobserve, derive and execute exactly once", findings);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in actual.OfType<JsonObject>())
        {
            var kind = item["kind"]?.GetValue<string>();
            if (kind is null || !seen.Add(kind))
            {
                findings.Add(("port-kind-invalid", "port kinds must be unique and declared"));
                continue;
            }
            switch (kind)
            {
                case "observe":
                    ValidateSetObservationPort(item, contract, "observe", findings);
                    break;
                case "reobserve":
                    ValidateSetObservationPort(item, contract, "reobserve", findings);
                    break;
                case "derive":
                    ValidateSetDerivationPort(item, contract, findings);
                    break;
                case "execute":
                    ValidateSetExecutionPort(item, contract, findings);
                    break;
                default:
                    findings.Add(("port-kind-invalid", $"port kind {kind} is not declared by this package"));
                    break;
            }
        }
        foreach (var kind in new[] { "observe", "reobserve", "derive", "execute" })
            Check(seen.Contains(kind), "port-missing", $"port kind {kind} must be declared", findings);
    }

    private static void ValidateSetPortKeys(JsonObject port, string kind, List<(string Code, string Message)> findings) =>
        Check(port.Count == 12 &&
            new[]
            {
                "kind", "producer", "validator", "requestSchema", "validatorRequestSchema",
                "resultSchema", "verdictSchema", "options", "cacheKeyComposition",
                "resourceDeclarations", "sideEffect", "attemptBudget"
            }.All(port.ContainsKey),
            "port-fields", $"{kind} port must close over exactly the declared fields", findings);

    private static void ValidateSetObservationPort(
        JsonObject port,
        Contract contract,
        string kind,
        List<(string Code, string Message)> findings)
    {
        ValidateSetPortKeys(port, kind, findings);
        ValidateSetAdapter(port["producer"], contract.ProviderId, contract.ProviderVersion,
            "inspect-evidence-v2", contract.ProviderVersion, $"{kind} producer", findings);
        ValidateSetAdapter(port["validator"], $"{contract.ProviderId}-validator", contract.ProviderVersion,
            "validate-inspect-evidence-v2", contract.ProviderVersion, $"{kind} validator", findings);
        ValidateSetSchema(port["requestSchema"], "tiwater.format-evidence-request/v2", $"{kind} request", findings);
        ValidateSetSchema(port["validatorRequestSchema"], "tiwater.format-evidence-request/v2", $"{kind} validator request", findings);
        ValidateSetSchema(port["resultSchema"], "tiwater.format-evidence/v2", $"{kind} result", findings);
        ValidateSetSchema(port["verdictSchema"], "tiwater.format-evidence-verdict/v2", $"{kind} verdict", findings);
        ValidateSetOptions(port["options"],
            [("facets", "tiwater.format-extraction-options/v1")], kind, findings);
        ValidateSetCache(port["cacheKeyComposition"],
            kind == "observe"
                ? ["bytesSha256", "optionsSha256", "provider"]
                : ["bytesSha256", "optionsSha256", "provider", "schemaSetVersion"],
            kind, findings);
        ValidateSetResources(port["resourceDeclarations"], $"{contract.Format}-document-bytes", "read", kind, findings);
        ValidateSetSideEffect(port["sideEffect"], "read-only", true, kind, findings);
        Check(port["attemptBudget"] is JsonValue budget && budget.TryGetValue<int>(out var attempts) && attempts == 1,
            "port-attempt-budget-mismatch", $"{kind} attempt budget does not match the package", findings);
    }

    private static void ValidateSetDerivationPort(JsonObject port, Contract contract, List<(string Code, string Message)> findings)
    {
        ValidateSetPortKeys(port, "derive", findings);
        ValidateSetAdapter(port["producer"], contract.ProviderId, contract.ProviderVersion,
            contract.DerivationProducerCommand, contract.ProviderVersion, "derive producer", findings);
        ValidateSetAdapter(port["validator"], $"{contract.ProviderId}.operation-derivation-validator", contract.ProviderVersion,
            contract.DerivationValidatorCommand, contract.ProviderVersion, "derive validator", findings);
        ValidateSetSchema(port["requestSchema"], "tiwater.operation-derivation-request/v2", "derive request", findings);
        ValidateSetSchema(port["validatorRequestSchema"], "tiwater.operation-derivation-request/v2", "derive validator request", findings);
        ValidateSetSchema(port["resultSchema"], "tiwater.operation-derivation-result/v1", "derive result", findings);
        ValidateSetSchema(port["verdictSchema"], "tiwater.operation-derivation-verdict/v1", "derive verdict", findings);
        ValidateSetOptions(port["options"], [], "derive", findings);
        ValidateSetCache(port["cacheKeyComposition"],
            ["bytesSha256", "optionsSha256", "provider", "schemaSetVersion"], "derive", findings);
        ValidateSetResources(port["resourceDeclarations"], $"{contract.Format}-document-bytes", "read", "derive", findings);
        ValidateSetSideEffect(port["sideEffect"], "read-only", true, "derive", findings);
        Check(port["attemptBudget"] is JsonValue budget && budget.TryGetValue<int>(out var attempts) && attempts == 1,
            "port-attempt-budget-mismatch", "derive attempt budget does not match the package", findings);
    }

    private static void ValidateSetExecutionPort(JsonObject port, Contract contract, List<(string Code, string Message)> findings)
    {
        ValidateSetPortKeys(port, "execute", findings);
        ValidateSetAdapter(port["producer"], contract.ProviderId, contract.ProviderVersion,
            contract.ExecutionProducerCommand, contract.ProviderVersion, "execute producer", findings);
        ValidateSetAdapter(port["validator"], $"{contract.ProviderId}.execution-evidence-validator", contract.ProviderVersion,
            contract.ExecutionValidatorCommand, contract.ProviderVersion, "execute validator", findings);
        ValidateSetSchema(port["requestSchema"], "tiwater.provider-effect-execution-request/v1", "execute request", findings);
        ValidateSetSchema(port["validatorRequestSchema"], "tiwater.provider-effect-execution-request/v1", "execute validator request", findings);
        ValidateSetSchema(port["resultSchema"], "lucid.execution-evidence/v2", "execute result", findings);
        ValidateSetSchema(port["verdictSchema"], "tiwater.execution-evidence-verdict/v1", "execute verdict", findings);
        ValidateSetOptions(port["options"], [], "execute", findings);
        ValidateSetCache(port["cacheKeyComposition"],
            ["bytesSha256", "optionsSha256", "provider", "schemaSetVersion"], "execute", findings);
        ValidateSetResources(port["resourceDeclarations"], $"{contract.Format}-document-bytes", "exclusive-write", "execute", findings);
        ValidateSetSideEffect(port["sideEffect"], "mutating", false, "execute", findings);
        Check(port["attemptBudget"] is JsonValue budget && budget.TryGetValue<int>(out var attempts) && attempts == 1,
            "port-attempt-budget-mismatch", "execute attempt budget does not match the package", findings);
    }

    private static void ValidateSetAdapter(
        JsonNode? node,
        string id,
        string version,
        string command,
        string commandVersion,
        string role,
        List<(string Code, string Message)> findings)
    {
        var valid = node is JsonObject value && value.Count == 3 &&
            value["id"]?.GetValue<string>() == id &&
            value["version"]?.GetValue<string>() == version &&
            value["adapterIdentity"] is JsonObject adapter && adapter.Count == 2 &&
            adapter["id"]?.GetValue<string>() == command &&
            adapter["version"]?.GetValue<string>() == commandVersion;
        Check(valid, "port-adapter-mismatch", $"{role} identity does not match the package", findings);
    }

    private static void ValidateSetSchema(
        JsonNode? node,
        string expectedId,
        string role,
        List<(string Code, string Message)> findings)
    {
        Check(node?.GetValue<string>() == expectedId,
            "port-schema-mismatch", $"{role} schema does not match the package contract", findings);
        DeployedSchema(expectedId);
    }

    private static void ValidateSetOptions(
        JsonNode? node,
        IReadOnlyList<(string Name, string ValueSchema)> expected,
        string kind,
        List<(string Code, string Message)> findings)
    {
        if (node is not JsonArray actual || actual.Count != expected.Count)
        {
            findings.Add(("port-options-mismatch", $"{kind} options do not match the package"));
            return;
        }
        for (var index = 0; index < expected.Count; index += 1)
        {
            var valid = actual[index] is JsonObject option && option.Count == 2 &&
                option["name"]?.GetValue<string>() == expected[index].Name &&
                option["valueSchema"]?.GetValue<string>() == expected[index].ValueSchema;
            Check(valid, "port-options-mismatch", $"{kind} options do not match the package", findings);
            DeployedSchema(expected[index].ValueSchema);
        }
    }

    private static void ValidateSetCache(
        JsonNode? node,
        IReadOnlyList<string> expected,
        string kind,
        List<(string Code, string Message)> findings)
    {
        var valid = node is JsonArray actual && actual.Count == expected.Count &&
            actual.Select((item, index) => item?.GetValue<string>() == expected[index]).All(match => match);
        Check(valid, "port-cache-composition-mismatch", $"{kind} cache key composition does not match the package", findings);
    }

    private static void ValidateSetResources(
        JsonNode? node,
        string resourceKey,
        string access,
        string kind,
        List<(string Code, string Message)> findings)
    {
        var valid = node is JsonArray actual && actual.Count == 1 &&
            actual[0] is JsonObject resource && resource.Count == 2 &&
            resource["resourceKey"]?.GetValue<string>() == resourceKey &&
            resource["access"]?.GetValue<string>() == access;
        Check(valid, "port-resources-mismatch", $"{kind} resource declarations do not match the package", findings);
    }

    private static void ValidateSetSideEffect(
        JsonNode? node,
        string kind,
        bool idempotent,
        string port,
        List<(string Code, string Message)> findings)
    {
        var valid = node is JsonObject value && value.Count == 2 &&
            value["kind"]?.GetValue<string>() == kind &&
            value["idempotent"] is JsonValue flag && flag.TryGetValue<bool>(out var declared) && declared == idempotent;
        Check(valid, "port-side-effect-mismatch", $"{port} side effect does not match the package", findings);
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
        new("format-extraction-options", "tiwater.format-extraction-options-v1.schema.json", "tiwater.format-extraction-options/v1"),
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

    private static string? Optional(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        if (index < 0) return null;
        if (index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]))
            throw new InvalidOperationException($"{name} requires a value");
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
