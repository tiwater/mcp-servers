using System.Text.Json;
using System.Text.Json.Nodes;
using Dockit.Convert;
using Json.Schema;
using Xunit;

namespace Dockit.Convert.Tests;

public sealed class ConvertProviderContractManifestCommandTests
{
    private static readonly JsonSchema Set15Schema = JsonSchema.FromText(
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "Fixtures", "provider-contract-manifest.schema.json")));

    private static readonly JsonSchema PackagedManifestSchema = JsonSchema.FromText(
        File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "schemas", "provider-contract-manifest-v1.schema.json")));

    private static readonly EvaluationOptions Evaluation = new() { OutputFormat = OutputFormat.List };

    [Fact]
    public void Manifest_satisfies_the_real_set15_schema_and_declares_render_only()
    {
        var fixture = Fixture();
        Assert.Equal(0, ConvertProviderContractManifestCommand.RunProducer(
            ["--schema-set-version", "15", "--output", fixture.Manifest]));
        var manifest = Load(fixture.Manifest);

        AssertValid(Set15Schema, manifest);
        AssertValid(PackagedManifestSchema, manifest);
        Assert.Equal("lucid.provider-contract-manifest", manifest["schema"]!.GetValue<string>());
        Assert.Equal(15, manifest["schemaSetVersion"]!.GetValue<int>());
        Assert.Matches("^manifest-[a-f0-9]{64}$", manifest["manifestId"]!.GetValue<string>());
        Assert.Equal("tiwater-convert", manifest["provider"]!["id"]!.GetValue<string>());
        Assert.Equal(RuntimeIdentity.Version, manifest["provider"]!["version"]!.GetValue<string>());
        Assert.Equal("tiwater-convert", manifest["runtime"]!["id"]!.GetValue<string>());

        var ports = manifest["ports"]!.AsArray();
        Assert.Single(ports);
        var port = ports[0]!;
        Assert.Equal("render", port["kind"]!.GetValue<string>());
        // Observation, derivation, reobservation, and execution ports belong to the Office and PDF providers.
        Assert.DoesNotContain(
            ports.Select(item => item!["kind"]!.GetValue<string>()),
            kind => kind is "observe" or "derive" or "reobserve" or "execute");
        Assert.Equal("tiwater-convert-render-producer", port["producer"]!["id"]!.GetValue<string>());
        Assert.Equal("tiwater-convert-render-producer", port["producer"]!["adapterIdentity"]!["id"]!.GetValue<string>());
        Assert.Equal(RuntimeIdentity.Version, port["producer"]!["adapterIdentity"]!["version"]!.GetValue<string>());
        Assert.Equal("tiwater-convert-render-validator", port["validator"]!["adapterIdentity"]!["id"]!.GetValue<string>());
        Assert.Equal("tiwater.convert-render-request/v1", port["requestSchema"]!.GetValue<string>());
        Assert.Equal("tiwater.convert-render-request/v1", port["validatorRequestSchema"]!.GetValue<string>());
        Assert.Equal("tiwater.convert-render-result/v1", port["resultSchema"]!.GetValue<string>());
        Assert.Equal("tiwater.convert-render-verdict/v1", port["verdictSchema"]!.GetValue<string>());
        Assert.Empty(port["options"]!.AsArray());
        Assert.Equal("read-only", port["sideEffect"]!["kind"]!.GetValue<string>());
        Assert.True(port["sideEffect"]!["idempotent"]!.GetValue<bool>());
        Assert.Equal(1, port["attemptBudget"]!.GetValue<int>());

        Assert.Equal(0, ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
        var verdict = Load(fixture.Verdict);
        Assert.Equal("tiwater.convert-provider-contract-manifest-verdict/v1", verdict["schema"]!.GetValue<string>());
        Assert.Equal("pass", verdict["decision"]!.GetValue<string>());
        Assert.Empty(verdict["findings"]!.AsArray());
    }

    [Fact]
    public void Producer_is_deterministic_for_the_same_package_and_schema_set()
    {
        var first = Fixture();
        var second = Fixture();
        ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "15", "--output", first.Manifest]);
        ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "15", "--output", second.Manifest]);
        Assert.Equal(File.ReadAllText(first.Manifest), File.ReadAllText(second.Manifest));
    }

    [Fact]
    public void Producer_fails_closed_without_schema_set_version()
    {
        var fixture = Fixture();
        Assert.ThrowsAny<Exception>(() =>
            ConvertProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest]));
        Assert.False(File.Exists(fixture.Manifest));
        Assert.ThrowsAny<Exception>(() =>
            ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "abc", "--output", fixture.Manifest]));
        Assert.ThrowsAny<Exception>(() =>
            ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "0", "--output", fixture.Manifest]));
    }

    public static IEnumerable<object[]> Mutations()
    {
        yield return Mutation(value => value["ports"]![0]!["kind"] = "observe", "render-port-becomes-observe");
        yield return Mutation(value => value["ports"]!.AsArray().Add(JsonNode.Parse(value["ports"]![0]!.ToJsonString())), "duplicate-render-port");
        yield return Mutation(value =>
        {
            var derive = value["ports"]![0]!.DeepClone();
            derive["kind"] = "derive";
            value["ports"]!.AsArray().Add(derive);
        }, "derive-port-added");
        yield return Mutation(value =>
        {
            var execute = value["ports"]![0]!.DeepClone();
            execute["kind"] = "execute";
            value["ports"]!.AsArray().Add(execute);
        }, "execute-port-added");
        yield return Mutation(value => value["ports"]![0]!["producer"]!["version"] = "0.0.1", "producer-version-drift");
        yield return Mutation(value => value["ports"]![0]!["producer"]!["adapterIdentity"]!["id"] = "invented-adapter", "producer-adapter-identity-drift");
        yield return Mutation(value => value["ports"]![0]!["validator"]!["adapterIdentity"]!["version"] = "0.0.1", "validator-adapter-identity-drift");
        yield return Mutation(value => value["ports"]![0]!["requestSchema"] = "tiwater.invented-request/v9", "request-schema-drift");
        yield return Mutation(value => value["ports"]![0]!["validatorRequestSchema"] = "tiwater.invented-request/v9", "validator-request-schema-drift");
        yield return Mutation(value => value["ports"]![0]!["resultSchema"] = "tiwater.invented-result/v9", "result-schema-drift");
        yield return Mutation(value => value["ports"]![0]!["verdictSchema"] = "tiwater.invented-verdict/v9", "verdict-schema-drift");
        yield return Mutation(value => value["ports"]![0]!["options"]!.AsArray().Add(new JsonObject
        {
            ["name"] = "quality",
            ["valueSchema"] = "tiwater.convert-render-request/v1"
        }), "undeclared-option");
        yield return Mutation(value => value["ports"]![0]!["cacheKeyComposition"]!.AsArray().RemoveAt(1), "cache-component-dropped");
        yield return Mutation(value => value["ports"]![0]!["resourceDeclarations"]![0]!["access"] = "exclusive-write", "resource-access-escalation");
        yield return Mutation(value => value["ports"]![0]!["sideEffect"]!["idempotent"] = false, "side-effect-drift");
        yield return Mutation(value => value["ports"]![0]!["attemptBudget"] = 99, "attempt-budget-drift");
        yield return Mutation(value => value["schemaSetVersion"] = 14, "schema-set-version-drift");
        yield return Mutation(value => value["manifestId"] = $"manifest-{new string('0', 64)}", "manifest-id-forgery");
        yield return Mutation(value => value["provider"]!["version"] = "0.0.1", "provider-version-drift");
        yield return Mutation(value => value["runtime"]!["id"] = "invented-runtime", "runtime-identity-drift");
    }

    [Theory]
    [MemberData(nameof(Mutations))]
    public void Independent_validator_rejects_mutations(Action<JsonObject> mutate, string _)
    {
        var fixture = Fixture();
        ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "15", "--output", fixture.Manifest]);
        var manifest = Load(fixture.Manifest);
        mutate(manifest);
        File.WriteAllText(fixture.Manifest, manifest.ToJsonString());
        Assert.Equal(1, ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
        var verdict = Load(fixture.Verdict);
        Assert.Equal("failed", verdict["decision"]!.GetValue<string>());
        Assert.NotEmpty(verdict["findings"]!.AsArray());
    }

    [Fact]
    public void Independent_validator_rejects_a_dropped_port()
    {
        var fixture = Fixture();
        ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "15", "--output", fixture.Manifest]);
        var manifest = Load(fixture.Manifest);
        manifest["ports"]!.AsArray().RemoveAt(0);
        File.WriteAllText(fixture.Manifest, manifest.ToJsonString());
        Assert.Equal(1, ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
        Assert.Equal("failed", Load(fixture.Verdict)["decision"]!.GetValue<string>());
    }

    [Fact]
    public void Manifest_and_verdict_outputs_are_immutable()
    {
        var fixture = Fixture();
        File.WriteAllText(fixture.Manifest, "occupied");
        Assert.ThrowsAny<Exception>(() =>
            ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "15", "--output", fixture.Manifest]));
        fixture = Fixture();
        ConvertProviderContractManifestCommand.RunProducer(["--schema-set-version", "15", "--output", fixture.Manifest]);
        File.WriteAllText(fixture.Verdict, "occupied");
        Assert.ThrowsAny<Exception>(() => ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
    }

    private static object[] Mutation(Action<JsonObject> mutate, string name) => [mutate, name];

    private static void AssertValid(JsonSchema schema, JsonObject instance)
    {
        using var document = JsonDocument.Parse(instance.ToJsonString());
        var results = schema.Evaluate(document.RootElement, Evaluation);
        Assert.True(results.IsValid, string.Join("; ", Errors(results)));
    }

    private static IEnumerable<string> Errors(EvaluationResults results)
    {
        if (results.Errors is { Count: > 0 })
            yield return $"{results.EvaluationPath}: {JsonSerializer.Serialize(results.Errors)}";
        foreach (var detail in results.Details ?? [])
            foreach (var error in Errors(detail))
                yield return error;
    }

    private sealed record FixtureValue(string Manifest, string Verdict);

    private static FixtureValue Fixture()
    {
        var root = Path.Combine(Path.GetTempPath(), $"convert-manifest-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        return new FixtureValue(Path.Combine(root, "manifest.json"), Path.Combine(root, "verdict.json"));
    }

    private static JsonObject Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))!.AsObject();
}
