using System.Security.Cryptography;
using System.Text.Json.Nodes;
using Dockit.Pptx.Cli;
using Json.Schema;
using Xunit;

/// <summary>
/// Set-15 (<c>lucid.provider-contract-manifest</c>) tests for the PPTX package.
/// The manifest under test is produced from the package's real production
/// contract (<see cref="Cli.ManifestContract"/>), validated against the
/// byte-copied Lucid schema-set 15 schema with a real Draft 2020-12 evaluator,
/// and independently recomputed by the package validator.
/// </summary>
public sealed class Set15ProviderContractManifestTests
{
    private const int SchemaSetVersion = 15;
    // Byte-copied from lucid-docs plugins/lucid/workflow/schema-sets/15/provider-contract-manifest.schema.json
    // (see packages/format-evidence/test-fixtures/README.md); pinned so the
    // fixture cannot drift from the real Lucid schema-set 15 bytes unnoticed.
    private const string FixtureSchemaSha256 = "7eddd13c38eb9b61d82787292b3b46433caafc29ea31f04ffae14725d60e14bc";

    [Fact]
    public void Manifest_satisfies_real_lucid_schema_and_independent_validator()
    {
        var fixture = Fixture();
        Produce(fixture);
        var manifest = Load(fixture.Manifest);
        Assert.Equal("lucid.provider-contract-manifest", manifest["schema"]!.GetValue<string>());
        Assert.Equal(SchemaSetVersion, manifest["schemaSetVersion"]!.GetValue<int>());
        Assert.Equal(
            $"manifest:{fixture.Contract.ProviderId}:{fixture.Contract.ProviderVersion}:set-{SchemaSetVersion}",
            manifest["manifestId"]!.GetValue<string>());
        Assert.Equal(fixture.Contract.ProviderId, manifest["provider"]!["id"]!.GetValue<string>());
        Assert.Equal(fixture.Contract.ProviderVersion, manifest["provider"]!["version"]!.GetValue<string>());
        Assert.Equal("dotnet", manifest["runtime"]!["id"]!.GetValue<string>());
        Assert.False(string.IsNullOrWhiteSpace(manifest["runtime"]!["version"]!.GetValue<string>()));
        Assert.Equal(
            ["observe", "reobserve", "derive", "execute"],
            manifest["ports"]!.AsArray().Select(port => port!["kind"]!.GetValue<string>()));
        foreach (var port in manifest["ports"]!.AsArray().Cast<JsonObject>())
        {
            Assert.Equal(12, port.Count);
            Assert.Equal(1, port["attemptBudget"]!.GetValue<int>());
            Assert.NotEqual(
                port["producer"]!["id"]!.GetValue<string>(),
                port["validator"]!["id"]!.GetValue<string>());
        }
        Assert.Equal("mutating", manifest["ports"]![3]!["sideEffect"]!["kind"]!.GetValue<string>());
        Assert.Equal("read-only", manifest["ports"]![0]!["sideEffect"]!["kind"]!.GetValue<string>());

        // The fixture is the exact Lucid schema-set 15 schema byte copy.
        Assert.Equal(FixtureSchemaSha256, Sha(File.ReadAllBytes(FixtureSchemaPath)));
        var schema = JsonSchema.FromText(File.ReadAllText(FixtureSchemaPath));
        var results = schema.Evaluate(
            JsonNode.Parse(File.ReadAllText(fixture.Manifest)),
            new EvaluationOptions { OutputFormat = OutputFormat.List });
        Assert.True(results.IsValid, Describe(results));

        Assert.Equal(0, Validate(fixture));
        Assert.Equal("pass", Load(fixture.Verdict)["decision"]!.GetValue<string>());
    }

    [Fact]
    public void Manifest_binds_real_ports_commands_and_deployed_schema_bytes()
    {
        var fixture = Fixture();
        Produce(fixture);
        var manifest = Load(fixture.Manifest);
        var ports = manifest["ports"]!.AsArray().Cast<JsonObject>().ToArray();
        // Render is deliberately absent: rendering belongs to the convert
        // package, not to the Office format packages.
        Assert.Equal(
            ["observe", "reobserve", "derive", "execute"],
            ports.Select(port => port["kind"]!.GetValue<string>()));

        // Every adapterIdentity is a real command surface of this package.
        var realCommands = new HashSet<string>(StringComparer.Ordinal)
        {
            "inspect-evidence-v2",
            "validate-inspect-evidence-v2",
            fixture.Contract.DerivationProducerCommand,
            fixture.Contract.DerivationValidatorCommand,
            fixture.Contract.ExecutionProducerCommand,
            fixture.Contract.ExecutionValidatorCommand
        };
        Assert.Equal(6, realCommands.Count);
        foreach (var port in ports)
            foreach (var role in new[] { "producer", "validator" })
            {
                var adapter = port[role]!["adapterIdentity"]!.AsObject();
                Assert.Contains(adapter["id"]!.GetValue<string>(), realCommands);
                Assert.Equal(fixture.Contract.ProviderVersion, adapter["version"]!.GetValue<string>());
            }

        // Observation reuse is explicit: the reobserve port reuses the
        // observation commands and format-evidence v2 family.
        var reobserve = ports[1];
        Assert.Equal("inspect-evidence-v2", reobserve["producer"]!["adapterIdentity"]!["id"]!.GetValue<string>());
        Assert.Equal("tiwater.format-evidence-request/v2", reobserve["requestSchema"]!.GetValue<string>());

        // Every declared schema name is a real deployed contract whose bytes
        // match the sha256 attested by the v1 manifest contracts list.
        var deployed = DeployedSchemaIndex();
        var attested = V1ContractHashes(fixture);
        var declared = DeclaredSchemaNames(manifest).Distinct(StringComparer.Ordinal).ToArray();
        Assert.Equal(10, declared.Length);
        foreach (var name in declared)
        {
            Assert.True(deployed.TryGetValue(name, out var file), $"schema not deployed: {name}");
            Assert.True(attested.TryGetValue(name, out var sha), $"schema not attested: {name}");
            Assert.Equal(sha, file.Sha256);
        }
    }

    [Fact]
    public void Independent_validator_rejects_schema_set_port_schema_and_adapter_mutations()
    {
        var mutations = new Action<JsonObject>[]
        {
            value => value["schemaSetVersion"] = SchemaSetVersion + 1,
            value => value["ports"]![0]!["producer"]!["id"] = "tiwater-forgery",
            value => value["ports"]![0]!["requestSchema"] = "tiwater.format-evidence-request/v9",
            value => value["ports"]![2]!["validator"]!["adapterIdentity"]!["id"] = "invented-command",
            value => value["ports"]![3]!["sideEffect"]!["kind"] = "read-only",
            value => value["injected"] = true,
            value => ((JsonArray)value["ports"]!).RemoveAt(1)
        };
        foreach (var mutate in mutations)
        {
            var fixture = Fixture();
            Produce(fixture);
            var manifest = Load(fixture.Manifest);
            mutate(manifest);
            File.WriteAllText(fixture.Manifest, manifest.ToJsonString());
            Assert.Equal(1, Validate(fixture));
            Assert.Equal("fail", Load(fixture.Verdict)["decision"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Manifest_and_verdict_outputs_are_immutable()
    {
        var fixture = Fixture();
        File.WriteAllText(fixture.Manifest, "occupied");
        Assert.ThrowsAny<Exception>(() => Produce(fixture));
        fixture = Fixture();
        Produce(fixture);
        File.WriteAllText(fixture.Verdict, "occupied");
        Assert.ThrowsAny<Exception>(() => Validate(fixture));
    }

    [Fact]
    public void Producer_and_validator_fail_closed_without_format_or_schema_set_version()
    {
        var contract = Cli.ManifestContract();
        var root = Path.Combine(Path.GetTempPath(), $"provider-manifest-set15-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var manifest = Path.Combine(root, "manifest.json");
        var verdict = Path.Combine(root, "verdict.json");
        Assert.ThrowsAny<Exception>(() => ProviderContractManifestCommand.RunProducer(
            ["--output", manifest, "--format", "lucid.provider-contract-manifest"], contract));
        Assert.ThrowsAny<Exception>(() => ProviderContractManifestCommand.RunProducer(
            ["--output", manifest, "--schema-set-version", "15"], contract));
        Assert.ThrowsAny<Exception>(() => ProviderContractManifestCommand.RunProducer(
            ["--output", manifest, "--format", "tiwater.invented/v1", "--schema-set-version", "15"], contract));
        Assert.ThrowsAny<Exception>(() => ProviderContractManifestCommand.RunProducer(
            ["--output", manifest, "--format", "lucid.provider-contract-manifest", "--schema-set-version", "abc"], contract));
        Assert.Equal(0, ProviderContractManifestCommand.RunProducer(
            ["--output", manifest, "--format", "lucid.provider-contract-manifest", "--schema-set-version", "15"], contract));
        Assert.ThrowsAny<Exception>(() => ProviderContractManifestCommand.RunValidator(
            ["--manifest", manifest, "--output", verdict], contract));
    }

    private static string FixtureSchemaPath =>
        Path.Combine(AppContext.BaseDirectory, "test-fixtures", "lucid.provider-contract-manifest.schema.json");

    private static void Produce(FixtureValue fixture) =>
        Assert.Equal(0, ProviderContractManifestCommand.RunProducer(
            ["--output", fixture.Manifest, "--format", "lucid.provider-contract-manifest", "--schema-set-version", "15"],
            fixture.Contract));

    private static int Validate(FixtureValue fixture) =>
        ProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict, "--schema-set-version", "15"],
            fixture.Contract);

    private static IReadOnlyDictionary<string, string> V1ContractHashes(FixtureValue fixture)
    {
        Assert.Equal(0, ProviderContractManifestCommand.RunProducer(["--output", fixture.V1Manifest], fixture.Contract));
        return Load(fixture.V1Manifest)["contracts"]!.AsArray().Cast<JsonObject>()
            .ToDictionary(
                item => item["schema"]!["id"]!.GetValue<string>(),
                item => item["schema"]!["sha256"]!.GetValue<string>(),
                StringComparer.Ordinal);
    }

    private static IReadOnlyDictionary<string, (string Path, string Sha256)> DeployedSchemaIndex()
    {
        var index = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
        foreach (var file in Directory.EnumerateFiles(
                     Path.Combine(AppContext.BaseDirectory, "contracts"), "*.schema.json"))
        {
            var id = JsonNode.Parse(File.ReadAllText(file))?["$id"]?.GetValue<string>();
            if (id is not null) index[id] = (file, Sha(File.ReadAllBytes(file)));
        }
        return index;
    }

    private static IEnumerable<string> DeclaredSchemaNames(JsonObject manifest)
    {
        foreach (var port in manifest["ports"]!.AsArray().Cast<JsonObject>())
        {
            yield return port["requestSchema"]!.GetValue<string>();
            yield return port["validatorRequestSchema"]!.GetValue<string>();
            yield return port["resultSchema"]!.GetValue<string>();
            yield return port["verdictSchema"]!.GetValue<string>();
            foreach (var option in port["options"]!.AsArray().Cast<JsonObject>())
                yield return option["valueSchema"]!.GetValue<string>();
        }
    }

    private static string Describe(EvaluationResults results)
    {
        var errors = new List<string>();
        void Walk(EvaluationResults node)
        {
            if (node.Errors is not null)
                foreach (var (key, value) in node.Errors)
                    errors.Add($"{node.EvaluationPath}: {key}: {value}");
            foreach (var detail in node.Details) Walk(detail);
        }
        Walk(results);
        return string.Join("; ", errors);
    }

    private sealed record FixtureValue(
        string Manifest,
        string Verdict,
        string V1Manifest,
        ProviderContractManifestCommand.Contract Contract);

    private static FixtureValue Fixture()
    {
        var root = Path.Combine(Path.GetTempPath(), $"provider-manifest-set15-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        return new(
            Path.Combine(root, "manifest-set15.json"),
            Path.Combine(root, "verdict.json"),
            Path.Combine(root, "manifest-v1.json"),
            Cli.ManifestContract());
    }

    private static JsonObject Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))!.AsObject();

    private static string Sha(byte[] value) =>
        Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();
}
