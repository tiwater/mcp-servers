using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Dockit.Pptx;
using Xunit;

namespace Dockit.Pptx.Tests;

public sealed class TemplateOperationDerivationTests
{
    private const string Hash = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

    [Fact]
    public void External_template_target_derives_and_independently_validates()
    {
        var request = Request();
        var result = OperationDerivationCommand.Produce(request, Contract());
        var operation = result["operation"]!["value"]!;
        Assert.Equal(
            "/ppt/slideMasters/slideMaster1.xml",
            operation["targetMasterPath"]!.GetValue<string>());
        Assert.Equal(2, operation["slides"]!.AsArray().Count);
        var resultPath = Path.GetTempFileName();
        try
        {
            File.WriteAllText(resultPath, $"{Canonical(result)}\n");
            Assert.Equal(
                "pass",
                OperationDerivationCommand.Validate(request, result, resultPath, Contract())["decision"]!.GetValue<string>());
            result["operation"]!["value"]!["slides"]!.AsArray().RemoveAt(1);
            result["operation"]!["sha256"] = Sha(Canonical(result["operation"]!["value"]));
            File.WriteAllText(resultPath, $"{Canonical(result)}\n");
            Assert.Equal(
                "fail",
                OperationDerivationCommand.Validate(request, result, resultPath, Contract())["decision"]!.GetValue<string>());
        }
        finally
        {
            File.Delete(resultPath);
        }
    }

    [Fact]
    public void External_template_bytes_are_authority()
    {
        var request = Request();
        File.WriteAllText(request["targetArtifact"]!["path"]!.GetValue<string>(), "different template");
        Assert.ThrowsAny<Exception>(() => OperationDerivationCommand.Produce(request, Contract()));
    }

    private static OperationDerivationCommand.Contract Contract() => new(
        "tiwater-pptx",
        "0.2.7",
        "pptx",
        "pptx.template-apply",
        "1",
        "tiwater.pptx-template-apply-v1.schema.json",
        "tiwater.pptx-template-apply/v1",
        "external-artifact",
        "root-object",
        false,
        value =>
        {
            var plan = JsonSerializer.Deserialize<TemplateApplicationPlan>(
                value.ToJsonString(),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            Assert.NotNull(plan);
            Assert.Equal(2, plan.Slides.Count);
        });

    private static JsonObject Request()
    {
        var sourcePath = Path.GetTempFileName();
        var templatePath = Path.GetTempFileName();
        File.WriteAllText(sourcePath, "source presentation");
        File.WriteAllText(templatePath, "template presentation");
        var sourceArtifact = Artifact("source-artifact", sourcePath);
        var templateArtifact = Artifact("template-artifact", templatePath);
        var target = new JsonObject
        {
            ["candidateId"] = "template-master-1",
            ["artifactVersionId"] = "template-artifact",
            ["epochId"] = "template-epoch",
            ["semanticIdentity"] = new JsonArray(Scalar("path", "/ppt/slideMasters/slideMaster1.xml")),
            ["locator"] = Typed("tiwater.provider-json-pointer-locator/v1", new JsonObject
            {
                ["format"] = "pptx",
                ["pointer"] = "/detail/Masters/0",
                ["candidateValueSha256"] = Hash
            }),
            ["capabilities"] = new JsonArray(new JsonObject { ["id"] = "pptx.template-apply", ["version"] = "1" }),
            ["supportedOperationKinds"] = new JsonArray(JsonValue.Create("applyTemplate")),
            ["resourceDeclarations"] = new JsonArray(new JsonObject
            {
                ["resourceKey"] = "pptx:template-artifact:master",
                ["access"] = "shared-write"
            }),
            ["writeDeclarations"] = new JsonArray(new JsonObject
            {
                ["resourceKey"] = "pptx:template-artifact:master",
                ["writeKey"] = "/detail/Masters/0"
            }),
            ["candidateValueSha256"] = Hash,
            ["inspectionSha256"] = Hash
        };
        var observationValue = new JsonObject
        {
            ["format"] = "pptx",
            ["artifactVersionId"] = "template-artifact",
            ["epochId"] = "template-epoch",
            ["inspectionSha256"] = Hash,
            ["facets"] = new JsonArray(),
            ["inventoryUniverse"] = new JsonObject(),
            ["targetUniverse"] = new JsonObject { ["candidates"] = new JsonArray(target.DeepClone()) }
        };
        var slides = new JsonArray(
            new JsonObject { ["slideNumber"] = 1, ["targetLayoutPath"] = "/ppt/slideLayouts/slideLayout1.xml" },
            new JsonObject { ["slideNumber"] = 2, ["targetLayoutPath"] = "/ppt/slideLayouts/slideLayout1.xml" });
        return new JsonObject
        {
            ["schema"] = "tiwater.operation-derivation-request/v1",
            ["requestId"] = "template-request",
            ["runId"] = "run-1",
            ["effectDescriptor"] = new JsonObject
            {
                ["identity"] = new JsonObject { ["id"] = "pptx.template-apply", ["version"] = "1" },
                ["descriptorSha256"] = Hash,
                ["operationSchema"] = Ref("tiwater.pptx-template-apply-v1.schema.json", "tiwater.pptx-template-apply/v1"),
                ["resourceSetSchema"] = Ref("tiwater.provider-resource-set-v1.schema.json", "tiwater.provider-resource-set/v1"),
                ["writeSetSchema"] = Ref("tiwater.provider-write-set-v1.schema.json", "tiwater.provider-write-set/v1"),
                ["targetScope"] = "external-artifact"
            },
            ["output"] = new JsonObject
            {
                ["outputId"] = "primary",
                ["artifact"] = sourceArtifact,
                ["epochId"] = "source-epoch",
                ["format"] = "pptx"
            },
            ["targetArtifact"] = templateArtifact,
            ["observation"] = Typed("tiwater.provider-document-observation/v2", observationValue),
            ["target"] = target,
            ["sourceFact"] = new JsonObject
            {
                ["ref"] = new JsonObject
                {
                    ["nodeId"] = "facts",
                    ["contract"] = new JsonObject { ["id"] = "lucid.normalized-facts/v4", ["sha256"] = Hash },
                    ["sha256"] = Hash
                },
                ["factId"] = "slides",
                ["value"] = Typed("lucid.slide-assignments/v1", slides)
            },
            ["effectIntent"] = Typed("tiwater.provider-effect-intent/v1", new JsonObject
            {
                ["effectId"] = "template-effect",
                ["effectType"] = new JsonObject { ["id"] = "pptx.template-apply", ["version"] = "1" },
                ["operationKind"] = "applyTemplate",
                ["arguments"] = new JsonArray(
                    new JsonObject
                    {
                        ["name"] = "targetMasterPath",
                        ["source"] = "target-field",
                        ["fieldName"] = "path"
                    },
                    new JsonObject
                    {
                        ["name"] = "slides",
                        ["source"] = "source-fact"
                    })
            }),
            ["bindingAuthority"] = Typed("lucid.binding-authority/v1", new JsonObject { ["bindingId"] = "binding-1" }),
            ["provider"] = new JsonObject
            {
                ["identity"] = new JsonObject { ["id"] = "tiwater-pptx", ["version"] = "0.2.7" },
                ["adapter"] = new JsonObject { ["id"] = "tiwater-pptx-template-derivation", ["version"] = "1" },
                ["runtime"] = new JsonObject { ["id"] = "dotnet", ["version"] = "9" }
            },
            ["expectedResultContract"] = Ref(
                "tiwater.operation-derivation-result-v1.schema.json",
                "tiwater.operation-derivation-result/v1")
        };
    }

    private static JsonObject Artifact(string id, string path) => new()
    {
        ["artifactVersionId"] = id,
        ["path"] = path,
        ["bytesSha256"] = FileSha(path),
        ["mediaType"] = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    };

    private static JsonObject Scalar(string name, string value) => new()
    {
        ["name"] = name,
        ["kind"] = "string",
        ["value"] = value,
        ["sha256"] = Sha(Canonical(JsonValue.Create(value)))
    };

    private static JsonObject Typed(string id, JsonNode value) => new()
    {
        ["schema"] = new JsonObject { ["id"] = id, ["sha256"] = Hash },
        ["value"] = value.DeepClone(),
        ["sha256"] = Sha(Canonical(value))
    };

    private static JsonObject Ref(string file, string id) => new()
    {
        ["id"] = id,
        ["sha256"] = FileSha(Path.Combine(AppContext.BaseDirectory, "contracts", file))
    };

    private static string FileSha(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

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
