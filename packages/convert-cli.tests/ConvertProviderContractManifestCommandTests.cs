using System.Text.Json.Nodes;
using Dockit.Convert;
using Xunit;

namespace Dockit.Convert.Tests;

public sealed class ConvertProviderContractManifestCommandTests
{
    [Fact]
    public void Package_manifest_closes_contracts_ports_and_runtime()
    {
        var fixture = Fixture();
        Assert.Equal(0, ConvertProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest]));
        var manifest = Load(fixture.Manifest);
        Assert.Equal("tiwater.convert-provider-contract-manifest/v1", manifest["schema"]!.GetValue<string>());
        Assert.Equal(
            ["render-request", "render-result", "render-verdict", "native-render-provenance",
             "provider-contract-manifest", "provider-contract-manifest-verdict"],
            manifest["contracts"]!.AsArray().Select(item => item!["role"]!.GetValue<string>()).ToArray());
        Assert.Equal(
            ["render", "provider-contract-manifest"],
            manifest["ports"]!.AsArray().Select(item => item!["role"]!.GetValue<string>()).ToArray());
        Assert.Equal("tiwater-convert", manifest["runtime"]!["id"]!.GetValue<string>());
        Assert.Equal(RuntimeIdentity.Version, manifest["runtime"]!["version"]!.GetValue<string>());

        Assert.Equal(0, ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
        var verdict = Load(fixture.Verdict);
        Assert.Equal("tiwater.convert-provider-contract-manifest-verdict/v1", verdict["schema"]!.GetValue<string>());
        Assert.Equal("pass", verdict["decision"]!.GetValue<string>());
        Assert.Empty(verdict["findings"]!.AsArray());
    }

    [Fact]
    public void Independent_validator_rejects_contract_port_runtime_and_hash_mutations()
    {
        var mutations = new Action<JsonObject>[]
        {
            value => value["contracts"]![0]!["schema_ref"]!["sha256"] = new string('0', 64),
            value => value["ports"]![0]!["producer"]!["command"] = "invented-command",
            value => value["ports"]![1]!["validator"]!["id"] = "wrong-validator",
            value => value["runtime"]!["version"] = "0.0.1",
            value => value["manifest_sha256"] = new string('f', 64)
        };
        foreach (var mutate in mutations)
        {
            var fixture = Fixture();
            ConvertProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest]);
            var manifest = Load(fixture.Manifest);
            mutate(manifest);
            File.WriteAllText(fixture.Manifest, manifest.ToJsonString());
            Assert.Equal(1, ConvertProviderContractManifestCommand.RunValidator(
                ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
            Assert.Equal("failed", Load(fixture.Verdict)["decision"]!.GetValue<string>());
        }
    }

    [Fact]
    public void Independent_validator_rejects_a_dropped_contract_role()
    {
        var fixture = Fixture();
        ConvertProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest]);
        var manifest = Load(fixture.Manifest);
        manifest["contracts"]!.AsArray().RemoveAt(2);
        // The forger recomputes the manifest hash honestly after dropping the role.
        var clone = manifest.DeepClone().AsObject();
        clone.Remove("manifest_sha256");
        manifest["manifest_sha256"] = Sha(Canonical(clone));
        File.WriteAllText(fixture.Manifest, manifest.ToJsonString());
        Assert.Equal(1, ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
        var verdict = Load(fixture.Verdict);
        Assert.Equal("failed", verdict["decision"]!.GetValue<string>());
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "contract-role-missing");
    }

    [Fact]
    public void Manifest_and_verdict_outputs_are_immutable()
    {
        var fixture = Fixture();
        File.WriteAllText(fixture.Manifest, "occupied");
        Assert.ThrowsAny<Exception>(() =>
            ConvertProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest]));
        fixture = Fixture();
        ConvertProviderContractManifestCommand.RunProducer(["--output", fixture.Manifest]);
        File.WriteAllText(fixture.Verdict, "occupied");
        Assert.ThrowsAny<Exception>(() => ConvertProviderContractManifestCommand.RunValidator(
            ["--manifest", fixture.Manifest, "--output", fixture.Verdict]));
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

    private static string Sha(string value) =>
        System.Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(value)))
        .ToLowerInvariant();

    private static string Canonical(JsonNode? node) => node switch
    {
        null => "null",
        JsonObject value => $"{{{string.Join(",", value.OrderBy(
            item => item.Key,
            StringComparer.Ordinal).Select(item =>
                $"{System.Text.Json.JsonSerializer.Serialize(item.Key)}:{Canonical(item.Value)}"))}}}",
        JsonArray value => $"[{string.Join(",", value.Select(Canonical))}]",
        _ => node.ToJsonString()
    };
}
