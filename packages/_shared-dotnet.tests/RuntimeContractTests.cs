using System.Text;
using System.Text.Json;
using Tiwater.RuntimeContracts;
using Xunit;

namespace Tiwater.RuntimeContracts.Tests;

public sealed class RuntimeContractTests
{
    [Fact]
    public void File_identity_binds_path_size_hash_and_content_id()
    {
        var identity = FileIdentity.IdentifyBytes("/runs/inputs/source.bin", Encoding.UTF8.GetBytes("abc"));

        Assert.Equal("/runs/inputs/source.bin", identity.Path);
        Assert.Equal(3, identity.SizeBytes);
        Assert.Equal("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", identity.Sha256);
        Assert.Equal($"sha256:{identity.Sha256}", identity.ContentId);
    }

    [Fact]
    public void Canonical_artifact_identity_is_independent_of_object_key_order()
    {
        using var first = JsonDocument.Parse("{\"b\":2,\"a\":1}");
        using var second = JsonDocument.Parse("{\"a\":1,\"b\":2}");
        var schema = new SchemaIdentity("tiwater.test-payload", "1.0.0");

        var firstArtifact = EvidenceEnvelope.IdentifyCanonicalJson(first.RootElement, schema);
        var secondArtifact = EvidenceEnvelope.IdentifyCanonicalJson(second.RootElement, schema);

        Assert.Equal(firstArtifact, secondArtifact);
        Assert.Equal("canonical-json", firstArtifact.Encoding);
        Assert.Equal($"sha256:{firstArtifact.Sha256}", firstArtifact.ArtifactId);
    }

    [Fact]
    public void Canonical_json_matches_the_cross_language_utf8_fixture()
    {
        using var payload = JsonDocument.Parse("{\"text\":\"中文\",\"items\":[2,1],\"nested\":{\"b\":true,\"a\":null}}");

        var bytes = EvidenceEnvelope.CanonicalJsonBytes(payload.RootElement);

        Assert.Equal("{\"items\":[2,1],\"nested\":{\"a\":null,\"b\":true},\"text\":\"中文\"}", Encoding.UTF8.GetString(bytes));
        Assert.Equal("be6d16a737da3afc2ab5eb06b725a397ab7c1a462eb5d4e221c8ecdd6b1264ec", Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes)));
    }

    [Fact]
    public void Canonical_json_matches_shared_adversarial_vectors()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "fixtures", "canonical-json-vectors.json");
        using var fixture = JsonDocument.Parse(File.ReadAllText(fixturePath));

        foreach (var vector in fixture.RootElement.GetProperty("vectors").EnumerateArray())
        {
            var bytes = EvidenceEnvelope.CanonicalJsonBytes(vector.GetProperty("value"));
            Assert.Equal(Encoding.UTF8.GetBytes(vector.GetProperty("canonical").GetString()!), bytes);
        }
    }

    [Fact]
    public void Canonical_json_rejects_duplicate_object_keys()
    {
        using var payload = JsonDocument.Parse("{\"a\":1,\"a\":2}");

        var error = Assert.Throws<InvalidOperationException>(() => EvidenceEnvelope.CanonicalJsonBytes(payload.RootElement));

        Assert.Contains("duplicate", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Canonical_json_rejects_non_integer_numbers()
    {
        using var payload = JsonDocument.Parse("{\"value\":1.5}");

        var error = Assert.Throws<InvalidOperationException>(() => EvidenceEnvelope.CanonicalJsonBytes(payload.RootElement));

        Assert.Contains("integer", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Canonical_json_rejects_shared_lossy_numeric_lexemes()
    {
        var fixturePath = Path.Combine(AppContext.BaseDirectory, "fixtures", "canonical-json-negative-vectors.json");
        using var fixture = JsonDocument.Parse(File.ReadAllText(fixturePath));

        foreach (var vector in fixture.RootElement.GetProperty("vectors").EnumerateArray())
        {
            using var payload = JsonDocument.Parse(vector.GetProperty("json").GetString()!);
            var error = Assert.Throws<InvalidOperationException>(() => EvidenceEnvelope.CanonicalJsonBytes(payload.RootElement));
            Assert.Contains("integer", error.Message, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void Derived_identity_serialization_cannot_claim_a_native_id()
    {
        EvidenceIdentity identity = new DerivedEvidenceIdentity(
            "source-version-structural-locator",
            ["sha256:source", "table[0]"]);

        var json = JsonSerializer.Serialize(identity, RuntimeJson.Options);

        Assert.Contains("\"kind\":\"derived\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("nativeId", json, StringComparison.Ordinal);
    }

    [Fact]
    public void Edit_summary_is_derived_from_the_complete_ordered_results()
    {
        var operations = new[]
        {
            EditOperationResult.ForApplied(0, "setCellValue", Json("{\"type\":\"setCellValue\"}"), Json("{\"value\":\"42\"}"), []),
            EditOperationResult.ForNoop(1, "setCellValue", Json("{\"type\":\"setCellValue\"}"), Json("{\"value\":\"42\"}"), []),
            EditOperationResult.ForRejected(2, "setCellValue", Json("{\"type\":\"setCellValue\"}"), [], [new ContractFinding("target-not-found", "Target was not found.")]),
        };

        var summary = EditReportSummary.FromOperations(operations);

        Assert.Equal(new EditReportSummary(3, 1, 1, 1, 0), summary);
        Assert.Equal([0, 1, 2], operations.Select(operation => operation.Index));
        Assert.Null(operations[2].AppliedPayload);
    }

    [Fact]
    public void Required_null_contract_fields_are_not_omitted_during_serialization()
    {
        var root = new EvidenceObject(
            "document:root",
            "document",
            true,
            null,
            new NativeEvidenceIdentity("package-part", "/document.xml"));
        var rejected = EditOperationResult.ForRejected(
            0,
            "setValue",
            Json("{\"type\":\"setValue\"}"),
            [],
            [new ContractFinding("target-not-found", "Target was not found.")]);

        var rootJson = JsonSerializer.Serialize(root, RuntimeJson.Options);
        var rejectedJson = JsonSerializer.Serialize(rejected, RuntimeJson.Options);

        Assert.Contains("\"parentObjectId\":null", rootJson, StringComparison.Ordinal);
        Assert.Contains("\"appliedPayload\":null", rejectedJson, StringComparison.Ordinal);
    }

    [Fact]
    public void Source_read_failure_serializes_without_inventing_source_identity()
    {
        using var payload = JsonDocument.Parse("{\"failureClass\":\"source-read-error\"}");
        var payloadSchema = new SchemaIdentity("tiwater.runtime.identify-payload", "1.0.0");
        var envelope = new RuntimeEvidenceEnvelope(
            "1.0.0",
            "runtime-evidence",
            "identify",
            "failed",
            "source-read",
            new PackageIdentity("tiwater.docx.cli", "0.4.0"),
            new RuntimeIdentity("office", "tiwater-docx", "0.4.0"),
            new SchemaIdentity("https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json", "1.0.0"),
            null,
            new RuntimeFileEvidence(null, null, new SignatureEvidence("not-checked", "ooxml-content-type", [])),
            EvidenceEnvelope.IdentifyCanonicalJson(payload.RootElement, payloadSchema),
            payload.RootElement.Clone(),
            [],
            [],
            [new ContractFinding("source-read-failed", "The source bytes could not be read.")]);

        var json = JsonSerializer.Serialize(envelope, RuntimeJson.Options);

        Assert.Contains("\"source\":null", json, StringComparison.Ordinal);
        Assert.Contains("\"failureStage\":\"source-read\"", json, StringComparison.Ordinal);
        Assert.Contains("\"status\":\"not-checked\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("contentId", json, StringComparison.Ordinal);
    }

    [Fact]
    public void Capability_descriptor_serializes_both_non_mutating_discovery_commands()
    {
        var schema = new SchemaIdentity("https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json", "1.0.0");
        var descriptor = new RuntimeCapabilityDescriptor(
            RuntimeContractVersions.Capabilities,
            "runtime-capabilities",
            new PackageIdentity("tiwater.test.cli", "1.2.3"),
            new RuntimeIdentity("test", "tiwater-test", "1.2.3"),
            schema,
            new DiscoveryCommand("capabilities", ["--json"], false),
            new IdentifyProbe("identify", ["<input>", "--json"], false, ["supported", "unsupported", "failed"]),
            [new SupportedKind("test", ["application/x-test"], ["test-signature"])],
            [
                new RuntimeCommand("capabilities", false, new SchemaIdentity("https://tiwater.dev/contracts/runtime/runtime-capabilities.schema.json", "1.0.0")),
                new RuntimeCommand("identify", false, schema),
            ],
            new IdentityPolicy("runtime-native-only", "deterministic-and-explicit", "parent-object-id-required-for-non-root"));

        var json = JsonSerializer.Serialize(descriptor, RuntimeJson.Options);

        Assert.Contains("\"descriptorCommand\":{\"command\":\"capabilities\"", json, StringComparison.Ordinal);
        Assert.Contains("\"identifyProbe\":{\"command\":\"identify\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("requiredProbeSet", json, StringComparison.Ordinal);
    }

    [Fact]
    public void Normalized_extraction_nodes_are_stable_contained_and_canonical()
    {
        using var report = JsonDocument.Parse("""
            {
              "file":"/renamed/source.docx",
              "tables":[{"rows":[{"cells":[{"text":"结果","confidence":0.95}]}]}]
            }
            """);

        var extraction = NormalizedEvidence.Build(report.RootElement);
        var nodes = extraction.Payload.GetProperty("nodes").EnumerateArray().ToArray();
        var text = nodes.Single(node => node.GetProperty("runtimeNodeId").GetString() == "/tables/0/rows/0/cells/0/text");
        var decimalNode = nodes.Single(node => node.GetProperty("runtimeNodeId").GetString() == "/tables/0/rows/0/cells/0/confidence");

        Assert.Equal("text", text.GetProperty("kind").GetString());
        Assert.Equal("结果", text.GetProperty("value").GetString());
        Assert.Equal("/tables/0/rows/0/cells/0", text.GetProperty("containedBy").GetString());
        Assert.Equal("0.95", decimalNode.GetProperty("value").GetString());
        Assert.DoesNotContain(nodes, node => node.GetProperty("runtimeNodeId").GetString() == "/file");
        Assert.Equal(extraction.Objects.Count, extraction.Objects.Select(node => node.ObjectId).Distinct().Count());
        Assert.All(extraction.Objects.Where(node => !node.Root), node =>
            Assert.Contains(extraction.Objects, parent => parent.ObjectId == node.ParentObjectId));
        _ = EvidenceEnvelope.IdentifyCanonicalJson(
            extraction.Payload,
            new SchemaIdentity("tiwater.runtime.normalized-evidence", "1.0.0"));
    }

    [Fact]
    public void Extraction_envelope_reuses_exact_source_identity_and_hashes_normalized_payload()
    {
        var identify = SupportedIdentify();
        using var report = JsonDocument.Parse("{\"rows\":[{\"text\":\"value\"}]}");

        var extraction = EvidenceEnvelope.CreateExtraction(identify, report.RootElement);

        Assert.Equal("extract-evidence", extraction.Probe);
        Assert.Equal(identify.Source, extraction.Source);
        Assert.Equal(identify.Runtime, extraction.Runtime);
        Assert.NotEmpty(extraction.Objects);
        Assert.Equal(
            EvidenceEnvelope.IdentifyCanonicalJson(extraction.Payload, extraction.Artifact.Schema),
            extraction.Artifact);
    }

    [Fact]
    public void Edit_report_binds_authoritative_request_source_output_and_complete_payloads()
    {
        var directory = Directory.CreateTempSubdirectory("tiwater-edit-report-");
        try
        {
            var source = Path.Combine(directory.FullName, "source.bin");
            var output = Path.Combine(directory.FullName, "output.bin");
            File.WriteAllText(source, "before");
            File.WriteAllText(output, "after");
            var request = Json("{\"operations\":[{\"type\":\"setValue\",\"value\":\"42\"}]}");
            var requestedOperation = request.GetProperty("operations")[0];
            var operations = new[]
            {
                EditOperationResult.ForApplied(
                    0,
                    "setValue",
                    requestedOperation,
                    requestedOperation,
                    [new TargetReference("cell:A1")]),
            };

            var report = EditReports.Create(
                new PackageIdentity("tiwater.test.cli", "1.0.0"),
                new RuntimeIdentity("test", "tiwater-test", "1.0.0"),
                source,
                output,
                request,
                [requestedOperation],
                operations);

            Assert.Equal("runtime-edit-report", report.ReportType);
            Assert.Equal(FileIdentity.IdentifyFile(source), report.Source);
            Assert.Equal(FileIdentity.IdentifyFile(output), report.Output);
            Assert.Equal(EvidenceEnvelope.IdentifyCanonicalJson(request, report.RequestArtifact.Schema), report.RequestArtifact);
            Assert.Equal(new EditReportSummary(1, 1, 0, 0, 0), report.Summary);
            Assert.True(JsonElement.DeepEquals(requestedOperation, report.Operations[0].AppliedPayload!.Value));
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    private static JsonElement Json(string value)
    {
        using var document = JsonDocument.Parse(value);
        return document.RootElement.Clone();
    }

    private static RuntimeEvidenceEnvelope SupportedIdentify()
    {
        var payload = Json("{\"recognized\":true}");
        return new RuntimeEvidenceEnvelope(
            "1.0.0",
            "runtime-evidence",
            "identify",
            "supported",
            null,
            new PackageIdentity("tiwater.test.cli", "1.0.0"),
            new RuntimeIdentity("test", "tiwater-test", "1.0.0"),
            new SchemaIdentity("https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json", "1.0.0"),
            FileIdentity.IdentifyBytes("/source", Encoding.UTF8.GetBytes("source")),
            new RuntimeFileEvidence("test", "application/x-test", new SignatureEvidence("matched", "test", ["matched"])),
            EvidenceEnvelope.IdentifyCanonicalJson(payload, new SchemaIdentity("tiwater.runtime.identify-payload", "1.0.0")),
            payload,
            [],
            [],
            []);
    }
}
