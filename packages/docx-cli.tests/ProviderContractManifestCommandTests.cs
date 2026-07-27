using System.Text.Json.Nodes;
using Xunit;

public sealed class ProviderContractManifestCommandTests
{
    [Fact]
    public void Package_manifest_closes_contracts_ports_and_adapter()
    {
        var fixture = Fixture();
        Assert.Equal(0, ProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest], fixture.Contract));
        var manifest = Load(fixture.Manifest);
        Assert.Equal(25, manifest["contracts"]!.AsArray().Count);
        Assert.Equal(
            ["format-observation", "operation-derivation", "effect-execution"],
            manifest["ports"]!.AsArray().Select(item => item!["role"]!.GetValue<string>()));
        Assert.Equal(0, ProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict], fixture.Contract));
        Assert.Equal("pass", Load(fixture.Verdict)["decision"]!.GetValue<string>());
    }

    [Fact]
    public void Independent_validator_rejects_contract_port_adapter_and_hash_mutations()
    {
        var mutations = new Action<JsonObject>[]
        {
            value => value["contracts"]![18]!["schema"]!["sha256"] = new string('0', 64),
            value => value["ports"]![1]!["producer"]!["command"] = "invented-command",
            value => value["executionAdapter"]!["id"] = "wrong-adapter",
            value => value["manifestSha256"] = new string('f', 64)
        };
        foreach (var mutate in mutations)
        {
            var fixture = Fixture();
            ProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest], fixture.Contract);
            var manifest = Load(fixture.Manifest);
            mutate(manifest);
            File.WriteAllText(fixture.Manifest, manifest.ToJsonString());
            Assert.Equal(1, ProviderContractManifestCommand.RunValidator(
                ["--manifest", fixture.Manifest, "--output", fixture.Verdict], fixture.Contract));
            Assert.Equal("fail", Load(fixture.Verdict)["decision"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Manifest_and_verdict_outputs_are_immutable()
    {
        var fixture = Fixture();
        File.WriteAllText(fixture.Manifest, "occupied");
        Assert.ThrowsAny<Exception>(() =>
            ProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest], fixture.Contract));
        fixture = Fixture();
        ProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest], fixture.Contract);
        File.WriteAllText(fixture.Verdict, "occupied");
        Assert.ThrowsAny<Exception>(() => ProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict], fixture.Contract));
    }

    private sealed record FixtureValue(
        string Manifest,
        string Verdict,
        ProviderContractManifestCommand.Contract Contract);

    private static FixtureValue Fixture()
    {
        var root = Path.Combine(Path.GetTempPath(), $"provider-manifest-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        return new(
            Path.Combine(root, "manifest.json"),
            Path.Combine(root, "verdict.json"),
            new(
                "tiwater-docx", "0.10.18", "docx.edit", "1",
                "tiwater.docx-edit-v1.schema.json", "tiwater.docx-edit/v1",
                "tiwater.docx-edit-result-v1.schema.json", "tiwater.docx-edit-result/v1",
                "derive-operation", "validate-derived-operation",
                "execute-effect", "validate-execution-evidence",
                "tiwater-docx-edit", "0.10.18", "docx"));
    }

    private static JsonObject Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))!.AsObject();
}
