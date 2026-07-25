using System.Text.Json.Nodes;

public sealed class ProviderOperationContractTests
{
    private static readonly IReadOnlyDictionary<string, string[]> Required =
        new Dictionary<string, string[]>
        {
            ["tiwater.operation-derivation-request/v1"] =
            [
                "schema", "requestId", "runId", "effectDescriptor", "output",
                "observation", "target", "sourceFact", "effectIntent", "bindingAuthority",
                "provider", "expectedResultContract"
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
    public void Supporting_value_contracts_are_package_owned(string id)
    {
        Assert.Equal(id, Load(id)["$id"]!.GetValue<string>());
    }

    private static JsonObject Load(string id)
    {
        var file = id.Replace("/", "-", StringComparison.Ordinal) + ".schema.json";
        var path = Path.Combine(AppContext.BaseDirectory, "contracts", file);
        return JsonNode.Parse(File.ReadAllText(path))!.AsObject();
    }
}
