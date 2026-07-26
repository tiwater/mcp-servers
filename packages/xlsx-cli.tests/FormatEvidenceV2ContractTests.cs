using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Tiwater.FormatEvidence;
using Xunit;
using XlsxCli = Dockit.Xlsx.Cli.Cli;

namespace Dockit.Xlsx.Tests;

public sealed class FormatEvidenceV2ContractTests
{
    [Fact]
    public void V2_evidence_exposes_each_worksheet_as_a_complete_edit_target()
    {
        var root = Path.Combine(Path.GetTempPath(), $"xlsx-format-evidence-v2-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "current.xlsx");
        var requestPath = Path.Combine(root, "request.json");
        var evidencePath = Path.Combine(root, "evidence.json");
        File.WriteAllText(input, "current provider bytes");

        var extractionValue = new { facets = new[] { "format-summary" } };
        var extractionJson = JsonSerializer.Serialize(extractionValue);
        File.WriteAllText(requestPath, JsonSerializer.Serialize(new
        {
            schema = "tiwater.format-evidence-request/v2",
            requestId = "request-1",
            runId = "run-1",
            subject = new { kind = "input", inputId = "input-1" },
            artifact = new
            {
                artifactVersionId = "artifact-1",
                path = input,
                bytesSha256 = FileSha(input),
                mediaType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            },
            provider = new { id = "tiwater-xlsx", version = "1.0.0" },
            validator = new { id = "tiwater-xlsx-validator", version = "1.0.0" },
            runtime = new { id = "tiwater-xlsx", version = "1.0.0" },
            extraction = new
            {
                schema = ContractRef("tiwater.format-extraction-options-v1.schema.json", "tiwater.format-extraction-options/v1"),
                value = extractionValue,
                sha256 = Sha(extractionJson)
            },
            expectedEvidenceContract = ContractRef("tiwater.format-evidence-v2.schema.json", "tiwater.format-evidence/v2")
        }));

        Assert.Equal(0, FormatEvidenceCommand.RunProducerV2(
            ["--request", requestPath, "--output", evidencePath],
            "tiwater-xlsx",
            "1.0.0",
            "xlsx",
            _ => new
            {
                export = new object[]
                {
                    new { sheet = "Inputs", cells = new[] { new { reference = "A1", value = "before" } } },
                    new { sheet = "Results", cells = Array.Empty<object>() }
                }
            },
            candidateCapabilities: XlsxCli.CandidateCapabilities));

        var evidence = JsonNode.Parse(File.ReadAllText(evidencePath))!;
        Assert.True(evidence["observation"] is not null, evidence.ToJsonString());
        var targets = evidence["observation"]!["value"]!
            ["targetUniverse"]!["candidates"]!.AsArray();
        Assert.Equal(2, targets.Count);
        Assert.Equal(
            new[] { "Inputs", "Results" },
            targets.Select(target => target!["semanticIdentity"]![0]!["value"]!.GetValue<string>()).ToArray());
        var expectedOperations = new[]
        {
            "copyRow",
            "expandSectionRows",
            "insertRows",
            "setCellValue",
            "setPrintArea",
            "setRangeValues",
            "setRichTextCellValue"
        };
        foreach (var target in targets)
        {
            Assert.Equal(
                expectedOperations,
                target!["supportedOperationKinds"]!.AsArray()
                    .Select(value => value!.GetValue<string>())
                    .ToArray());
            Assert.Equal("xlsx.edit", target["capabilities"]![0]!["id"]!.GetValue<string>());
            Assert.Equal("1", target["capabilities"]![0]!["version"]!.GetValue<string>());
        }
    }

    private static object ContractRef(string file, string id)
        => new { id, sha256 = FileSha(Path.Combine(AppContext.BaseDirectory, "contracts", file)) };

    private static string FileSha(string path)
        => System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

    private static string Sha(string value)
        => System.Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
