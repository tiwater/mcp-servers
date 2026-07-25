using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Tiwater.FormatEvidence;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class FormatEvidenceV2ContractTests
{
    [Fact]
    public void V2_evidence_is_typed_recomputed_and_tamper_evident()
    {
        var root = Path.Combine(Path.GetTempPath(), $"format-evidence-v2-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "current.docx");
        var requestPath = Path.Combine(root, "request.json");
        var evidencePath = Path.Combine(root, "evidence.json");
        var secondEvidencePath = Path.Combine(root, "evidence-second.json");
        var verdictPath = Path.Combine(root, "verdict.json");
        var tamperedVerdictPath = Path.Combine(root, "verdict-tampered.json");
        File.WriteAllText(input, "current provider bytes");

        var extractionValue = new { facets = new[] { "format-summary" } };
        var extractionJson = JsonSerializer.Serialize(extractionValue);
        var request = new
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
                mediaType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            },
            provider = new { id = "tiwater-docx", version = "1.0.0" },
            validator = new { id = "tiwater-docx-validator", version = "1.0.0" },
            runtime = new { id = "tiwater-docx", version = "1.0.0" },
            extraction = new
            {
                schema = ContractRef("tiwater.format-extraction-options-v1.schema.json", "tiwater.format-extraction-options/v1"),
                value = extractionValue,
                sha256 = Sha(extractionJson)
            },
            expectedEvidenceContract = ContractRef("tiwater.format-evidence-v2.schema.json", "tiwater.format-evidence/v2")
        };
        File.WriteAllText(requestPath, JsonSerializer.Serialize(request));

        Assert.Equal(0, FormatEvidenceCommand.RunProducerV2(
            ["--request", requestPath, "--output", evidencePath],
            "tiwater-docx",
            "1.0.0",
            "docx",
            _ => new { document = new { paragraphs = 1 }, tables = new { count = 0 } },
            supportedOperationKinds: SupportedKinds));
        Assert.Equal(0, FormatEvidenceCommand.RunValidatorV2(
            ["--request", requestPath, "--evidence", evidencePath, "--output", verdictPath],
            "tiwater-docx",
            "1.0.0",
            "docx",
            _ => new { document = new { paragraphs = 1 }, tables = new { count = 0 } },
            supportedOperationKinds: SupportedKinds));
        Assert.Equal("pass", JsonNode.Parse(File.ReadAllText(verdictPath))!["decision"]!.GetValue<string>());

        Assert.Equal(0, FormatEvidenceCommand.RunProducerV2(
            ["--request", requestPath, "--output", secondEvidencePath],
            "tiwater-docx",
            "1.0.0",
            "docx",
            _ => new { document = new { paragraphs = 2 }, tables = new { count = 1 } },
            supportedOperationKinds: SupportedKinds));
        var firstObservation = JsonNode.Parse(File.ReadAllText(evidencePath))!["observation"]!["sha256"]!.GetValue<string>();
        var secondObservation = JsonNode.Parse(File.ReadAllText(secondEvidencePath))!["observation"]!["sha256"]!.GetValue<string>();
        Assert.NotEqual(firstObservation, secondObservation);
        var firstUniverse = JsonNode.Parse(File.ReadAllText(evidencePath))!["observation"]!["value"]!["inventoryUniverse"]!;
        var secondUniverse = JsonNode.Parse(File.ReadAllText(secondEvidencePath))!["observation"]!["value"]!["inventoryUniverse"]!;
        Assert.NotEqual(
            firstUniverse["universeSha256"]!.GetValue<string>(),
            secondUniverse["universeSha256"]!.GetValue<string>());
        Assert.Equal(
            firstUniverse["candidates"]!.AsArray().Count,
            secondUniverse["candidates"]!.AsArray().Count);
        Assert.NotEqual(
            firstUniverse["candidates"]!.ToJsonString(),
            secondUniverse["candidates"]!.ToJsonString());
        var targets = JsonNode.Parse(File.ReadAllText(evidencePath))!["observation"]!["value"]!["targetUniverse"]!["candidates"]!.AsArray();
        Assert.Single(targets);
        Assert.Equal(
            "replaceParagraphText",
            targets[0]!["supportedOperationKinds"]![0]!.GetValue<string>());

        var tampered = JsonNode.Parse(File.ReadAllText(evidencePath))!.AsObject();
        tampered["observation"]!["value"]!["inventoryUniverse"]!["candidates"]!.AsArray().RemoveAt(0);
        File.WriteAllText(evidencePath, tampered.ToJsonString());
        Assert.Equal(0, FormatEvidenceCommand.RunValidatorV2(
            ["--request", requestPath, "--evidence", evidencePath, "--output", tamperedVerdictPath],
            "tiwater-docx",
            "1.0.0",
            "docx",
            _ => new { document = new { paragraphs = 1 }, tables = new { count = 0 } },
            supportedOperationKinds: SupportedKinds));
        Assert.Equal("failed", JsonNode.Parse(File.ReadAllText(tamperedVerdictPath))!["decision"]!.GetValue<string>());
    }

    private static object ContractRef(string file, string id)
        => new { id, sha256 = FileSha(Path.Combine(AppContext.BaseDirectory, "contracts", file)) };

    private static string FileSha(string path)
        => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

    private static string Sha(string canonicalJson)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonicalJson))).ToLowerInvariant();

    private static IReadOnlyList<string> SupportedKinds(IReadOnlySet<string> fields) =>
        fields.Contains("paragraphs") ? ["replaceParagraphText"] : [];
}
