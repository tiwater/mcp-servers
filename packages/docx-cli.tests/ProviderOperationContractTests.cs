using System.Security.Cryptography;
using System.Text.Json.Nodes;
using Xunit;

public sealed class ProviderOperationContractTests
{
    private static readonly IReadOnlyDictionary<string, string[]> Required =
        new Dictionary<string, string[]>
        {
            ["tiwater.operation-derivation-request/v1"] =
            [
                "schema", "requestId", "runId", "effectDescriptor", "output",
                "targetArtifact", "observation", "target", "sourceFact", "effectIntent", "bindingAuthority",
                "provider", "expectedResultContract"
            ],
            ["tiwater.operation-derivation-request/v2"] =
            [
                "schema", "requestId", "runId", "effectIntentId", "bindingId",
                "closureAuthority", "bindingAuthority", "normalizedFactsAuthority",
                "effectDescriptor", "output", "targetArtifact", "observation",
                "target", "sourceFacts", "effectIntent", "provider", "expectedResultContract"
            ],
            ["tiwater.operation-derivation-result/v1"] =
            [
                "schema", "derivationId", "requestId", "effectDescriptor",
                "output", "targetCandidateId", "operation", "resourceSet",
                "writeSet", "provenance"
            ],
            ["tiwater.operation-derivation-verdict/v1"] =
            [
                "schema", "requestId", "result", "validator",
                "recomputedOperationSha256", "recomputedResourceSetSha256",
                "recomputedWriteSetSha256", "recomputedProvenanceSha256",
                "decision", "findings"
            ]
        };

    [Fact]
    public void Public_derivation_envelopes_are_closed_and_exact()
    {
        foreach (var (id, required) in Required)
        {
            var schema = Load(id);
            Assert.Equal(id, schema["$id"]!.GetValue<string>());
            Assert.False(schema["additionalProperties"]!.GetValue<bool>());
            Assert.Equal(
                required,
                schema["required"]!.AsArray().Select(value => value!.GetValue<string>()));
        }
    }

    [Theory]
    [InlineData("tiwater.provider-effect-intent/v1")]
    [InlineData("tiwater.provider-resource-set/v1")]
    [InlineData("tiwater.provider-write-set/v1")]
    [InlineData("tiwater.operation-derivation-provenance/v1")]
    [InlineData("lucid.effect-execution-request/v1")]
    [InlineData("lucid.execution-evidence/v2")]
    [InlineData("tiwater.provider-effect-execution-request/v1")]
    [InlineData("tiwater.provider-artifact-lineage/v1")]
    [InlineData("tiwater.execution-evidence-verdict/v1")]
    [InlineData("lucid.canonical-node/v2")]
    [InlineData("lucid.operator-verdict/v1")]
    [InlineData("lucid.effect-bundle/v3")]
    [InlineData("lucid.composed-effect/v2")]
    [InlineData("tiwater.provider-contract-manifest/v1")]
    [InlineData("tiwater.provider-contract-manifest-verdict/v1")]
    [InlineData("tiwater.provider-document-observation/v2")]
    public void Supporting_value_contracts_are_package_owned(string id)
    {
        Assert.Equal(id, Load(id)["$id"]!.GetValue<string>());
    }

    [Fact]
    public void Published_observation_v1_bytes_remain_compatible()
    {
        var path = Path.Combine(
            AppContext.BaseDirectory,
            "contracts",
            "tiwater.provider-document-observation-v1.schema.json");
        var sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
        Assert.Equal("f61eeab31c67fba295192b817986ed30833010df85ab81a17ee34c95fd35a005", sha256);
    }

    [Fact]
    public void Published_derivation_request_v1_bytes_remain_compatible()
    {
        var path = Path.Combine(
            AppContext.BaseDirectory,
            "contracts",
            "tiwater.operation-derivation-request-v1.schema.json");
        var sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
        Assert.Equal("18deaab46eacb64b53385f95afb11917820adc9f7be427435cd9dd23f629fee2", sha256);
    }

    private static JsonObject Load(string id)
    {
        var file = id.Replace("/", "-", StringComparison.Ordinal) + ".schema.json";
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        return JsonNode.Parse(File.ReadAllText(path))!.AsObject();
    }
}
