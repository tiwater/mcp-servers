using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Dockit.Docx;
using Xunit;

public sealed class OperationDerivationCommandTests
{
    private const string Hash = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

    [Fact]
    public void Scalar_fact_derives_one_provider_operation_and_independent_pass()
    {
        var request = Request();
        var result = OperationDerivationCommand.Produce(request, Contract());
        Assert.Equal("replaceTableCellText", result["operation"]!["value"]!["operations"]![0]!["type"]!.GetValue<string>());
        Assert.Equal("approved", result["operation"]!["value"]!["operations"]![0]!["text"]!.GetValue<string>());
        Assert.Equal(request["target"]!["resourceDeclarations"]!.ToJsonString(), result["resourceSet"]!["value"]!.ToJsonString());
        var path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, $"{Canonical(result)}\n");
            var verdict = OperationDerivationCommand.Validate(request, result, path, Contract());
            Assert.Equal("pass", verdict["decision"]!.GetValue<string>());
            Assert.Empty(verdict["findings"]!.AsArray());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Structured_fact_remains_typed_and_derives_table_rows()
    {
        var rows = new JsonArray(
            new JsonArray(
                new JsonObject { ["text"] = "A" },
                new JsonObject { ["text"] = "B" }));
        var request = Request(
            operationKind: "replaceTable",
            sourceValue: rows,
            sourceArgument: "rows",
            semanticFields:
            [
                Scalar("tableIndex", JsonValue.Create(0))
            ]);
        var result = OperationDerivationCommand.Produce(request, Contract());
        Assert.Equal(Canonical(rows), Canonical(result["operation"]!["value"]!["operations"]![0]!["rows"]));
    }

    [Fact]
    public void Input_authority_mutations_fail_before_derivation()
    {
        var mutations = new Action<JsonObject>[]
        {
            request => request["target"]!["epochId"] = "stale-epoch",
            request => request["sourceFact"]!["value"]!["sha256"] = Hash,
            request => request["effectDescriptor"]!["operationSchema"]!["sha256"] = Hash,
            request => request["target"]!["resourceDeclarations"] = new JsonArray(),
            request => request["observation"]!["value"]!["targetUniverse"]!["candidates"] = new JsonArray()
        };
        foreach (var mutate in mutations)
        {
            var request = Request();
            mutate(request);
            RetagObservation(request);
            Assert.ThrowsAny<Exception>(() => OperationDerivationCommand.Produce(request, Contract()));
        }
    }

    [Fact]
    public void Result_cannot_self_attest_after_operation_tampering()
    {
        var request = Request();
        var result = OperationDerivationCommand.Produce(request, Contract());
        result["operation"]!["value"]!["operations"]![0]!["text"] = "forged";
        result["operation"]!["sha256"] = Sha(Canonical(result["operation"]!["value"]));
        var path = Path.GetTempFileName();
        try
        {
            File.WriteAllText(path, $"{Canonical(result)}\n");
            var verdict = OperationDerivationCommand.Validate(request, result, path, Contract());
            Assert.Equal("fail", verdict["decision"]!.GetValue<string>());
            Assert.Contains(
                verdict["findings"]!.AsArray(),
                finding => finding!["path"]!.GetValue<string>() == "/operation");
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static OperationDerivationCommand.Contract Contract() => new(
        "tiwater-docx",
        "0.10.14",
        "docx",
        "docx.edit",
        "1",
        "tiwater.docx-edit-v1.schema.json",
        "tiwater.docx-edit/v1",
        value =>
        {
            var document = JsonSerializer.Deserialize<DocxEditDocument>(value.ToJsonString(), Json.Options);
            Assert.NotNull(document);
            Assert.Single(document.Operations);
        });

    private static JsonObject Request(
        string operationKind = "replaceTableCellText",
        JsonNode? sourceValue = null,
        string sourceArgument = "text",
        JsonArray? semanticFields = null)
    {
        sourceValue ??= JsonValue.Create("approved");
        semanticFields ??=
        [
            Scalar("tableIndex", JsonValue.Create(0)),
            Scalar("rowIndex", JsonValue.Create(1)),
            Scalar("cellIndex", JsonValue.Create(2))
        ];
        var target = new JsonObject
        {
            ["candidateId"] = "candidate-1",
            ["artifactVersionId"] = "artifact-1",
            ["epochId"] = "epoch-1",
            ["semanticIdentity"] = semanticFields,
            ["locator"] = Typed("tiwater.provider-json-pointer-locator/v1", new JsonObject
            {
                ["format"] = "docx",
                ["pointer"] = "/tables/0/rows/1/cells/2",
                ["candidateValueSha256"] = Hash
            }),
            ["capabilities"] = new JsonArray(new JsonObject { ["id"] = "docx.edit", ["version"] = "1" }),
            ["resourceDeclarations"] = new JsonArray(new JsonObject
            {
                ["resourceKey"] = "docx:artifact-1:document",
                ["access"] = "shared-write"
            }),
            ["writeDeclarations"] = new JsonArray(new JsonObject
            {
                ["resourceKey"] = "docx:artifact-1:document",
                ["writeKey"] = "/tables/0/rows/1/cells/2"
            }),
            ["candidateValueSha256"] = Hash,
            ["inspectionSha256"] = Hash
        };
        var observationValue = new JsonObject
        {
            ["format"] = "docx",
            ["artifactVersionId"] = "artifact-1",
            ["epochId"] = "epoch-1",
            ["inspectionSha256"] = Hash,
            ["facets"] = new JsonArray(),
            ["inventoryUniverse"] = new JsonObject(),
            ["targetUniverse"] = new JsonObject
            {
                ["candidates"] = new JsonArray(target.DeepClone())
            }
        };
        var sourceTyped = Typed("lucid.typed-value/v1", sourceValue);
        var request = new JsonObject
        {
            ["schema"] = "tiwater.operation-derivation-request/v1",
            ["requestId"] = "request-1",
            ["runId"] = "run-1",
            ["effectDescriptor"] = new JsonObject
            {
                ["identity"] = new JsonObject { ["id"] = "docx.edit", ["version"] = "1" },
                ["descriptorSha256"] = Hash,
                ["operationSchema"] = Ref("tiwater.docx-edit-v1.schema.json", "tiwater.docx-edit/v1"),
                ["resourceSetSchema"] = Ref("tiwater.provider-resource-set-v1.schema.json", "tiwater.provider-resource-set/v1"),
                ["writeSetSchema"] = Ref("tiwater.provider-write-set-v1.schema.json", "tiwater.provider-write-set/v1")
            },
            ["output"] = new JsonObject
            {
                ["outputId"] = "primary",
                ["artifactVersionId"] = "artifact-1",
                ["epochId"] = "epoch-1",
                ["format"] = "docx"
            },
            ["observation"] = Typed("tiwater.provider-document-observation/v1", observationValue),
            ["target"] = target,
            ["sourceFact"] = new JsonObject
            {
                ["ref"] = new JsonObject
                {
                    ["nodeId"] = "facts",
                    ["contract"] = new JsonObject { ["id"] = "lucid.normalized-facts/v4", ["sha256"] = Hash },
                    ["sha256"] = Hash
                },
                ["factId"] = "fact-1",
                ["value"] = sourceTyped
            },
            ["effectIntent"] = Typed("tiwater.provider-effect-intent/v1", new JsonObject
            {
                ["effectId"] = "effect-1",
                ["effectType"] = new JsonObject { ["id"] = "docx.edit", ["version"] = "1" },
                ["operationKind"] = operationKind,
                ["arguments"] = new JsonArray(new JsonObject
                {
                    ["name"] = sourceArgument,
                    ["source"] = "source-fact"
                })
            }),
            ["bindingAuthority"] = Typed("lucid.binding-authority/v1", new JsonObject
            {
                ["bindingId"] = "binding-1"
            }),
            ["provider"] = new JsonObject
            {
                ["identity"] = new JsonObject { ["id"] = "tiwater-docx", ["version"] = "0.10.14" },
                ["adapter"] = new JsonObject { ["id"] = "tiwater-docx-operation-derivation", ["version"] = "1" },
                ["runtime"] = new JsonObject { ["id"] = "dotnet", ["version"] = "9" }
            },
            ["expectedResultContract"] = Ref(
                "tiwater.operation-derivation-result-v1.schema.json",
                "tiwater.operation-derivation-result/v1")
        };
        return request;
    }

    private static JsonObject Scalar(string name, JsonNode? value) =>
        new()
        {
            ["name"] = name,
            ["kind"] = value is null ? "null" : value.GetValueKind() switch
            {
                JsonValueKind.String => "string",
                JsonValueKind.Number => "number",
                JsonValueKind.True or JsonValueKind.False => "boolean",
                _ => throw new InvalidOperationException("unsupported scalar")
            },
            ["value"] = value?.DeepClone(),
            ["sha256"] = Sha(Canonical(value))
        };

    private static JsonObject Typed(string id, JsonNode value) =>
        new()
        {
            ["schema"] = new JsonObject { ["id"] = id, ["sha256"] = Hash },
            ["value"] = value.DeepClone(),
            ["sha256"] = Sha(Canonical(value))
        };

    private static JsonObject Ref(string file, string id) =>
        new()
        {
            ["id"] = id,
            ["sha256"] = Sha(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "contracts", file)))
        };

    private static void RetagObservation(JsonObject request)
    {
        var observation = request["observation"]!.AsObject();
        observation["sha256"] = Sha(Canonical(observation["value"]));
    }

    private static string Sha(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

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
}
