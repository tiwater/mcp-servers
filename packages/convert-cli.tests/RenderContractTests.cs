using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Dockit.Convert;
using Xunit;

namespace Dockit.Convert.Tests;

public sealed class RenderContractTests
{
    [Fact]
    public void Passthrough_render_and_independent_validator_pass_for_semantically_different_inputs()
    {
        var inputs = new[]
        {
            new[] { (612.0, 792.0, "Alpha certificate 260245") },
            new[] { (612.0, 792.0, "Beta batch HSP1028"), (300.0, 400.0, "Second page") }
        };
        foreach (var pages in inputs)
        {
            var fixture = Fixture(".pdf");
            File.WriteAllBytes(fixture.Input, BuildPdf(pages));
            WriteRequest(fixture, "pdf", "passthrough");

            Assert.Equal(0, RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result]));
            var result = Load(fixture.Result);
            Assert.Equal("tiwater.convert-render-result/v1", result["schema"]!.GetValue<string>());
            Assert.Equal(pages.Length, result["page_count"]!.GetValue<int>());
            Assert.Equal(pages.Length, result["pages"]!.AsArray().Count);
            Assert.Null(result["native_render_provenance"]);
            Assert.Null(result["provenance_sha256"]);

            Assert.Equal(0, RenderEvidenceValidatorCommand.Run(
                ["--request", fixture.Request, "--result", fixture.Result, "--output", fixture.Verdict]));
            var verdict = Load(fixture.Verdict);
            Assert.Equal("pass", verdict["decision"]!.GetValue<string>());
            Assert.Empty(verdict["findings"]!.AsArray());
            Assert.Equal(FileSha(fixture.Input), verdict["recomputed"]!["input_sha256"]!.GetValue<string>());
            Assert.Equal(FileSha(fixture.Output), verdict["recomputed"]!["output_sha256"]!.GetValue<string>());
            Assert.Equal(pages.Length, verdict["recomputed"]!["page_count"]!.GetValue<int>());
        }
    }

    [Fact]
    public void Injected_renderer_result_passes_independent_validation()
    {
        var fixture = Fixture(".docx");
        File.WriteAllBytes(fixture.Input, "synthetic docx bytes"u8.ToArray());
        WriteRequest(fixture, "docx", "libreoffice");
        var pdf = BuildPdf((612.0, 792.0, "Rendered page one"), (612.0, 792.0, "Rendered page two"));
        RenderBackend fake = invocation =>
        {
            Assert.Equal("libreoffice", invocation.Backend);
            File.WriteAllBytes(invocation.Output, pdf);
            return new OfficePdfConversionResult("libreoffice");
        };

        Assert.Equal(0, RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result], fake));
        var result = Load(fixture.Result);
        Assert.Equal(2, result["page_count"]!.GetValue<int>());
        Assert.Equal("tiwater-convert-render-producer", result["producer"]!["id"]!.GetValue<string>());

        Assert.Equal(0, RenderEvidenceValidatorCommand.Run(
            ["--request", fixture.Request, "--result", fixture.Result, "--output", fixture.Verdict]));
        Assert.Equal("pass", Load(fixture.Verdict)["decision"]!.GetValue<string>());
    }

    [Fact]
    public void Injected_wps_renderer_with_provenance_passes_independent_validation()
    {
        var (fixture, result) = ProduceWpsResult();
        Assert.Equal("wps-writer", result["renderer"]!["backend"]!.GetValue<string>());
        Assert.NotNull(result["native_render_provenance"]);
        Assert.Equal(0, RenderEvidenceValidatorCommand.Run(
            ["--request", fixture.Request, "--result", fixture.Result, "--output", fixture.Verdict]));
        var verdict = Load(fixture.Verdict);
        Assert.Equal("pass", verdict["decision"]!.GetValue<string>());
        Assert.Equal(
            verdict["recomputed"]!["provenance_sha256"]!.GetValue<string>(),
            result["provenance_sha256"]!.GetValue<string>());
    }

    [Fact]
    public void Validator_rejects_a_missing_page()
    {
        var fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result =>
        {
            result["pages"]!.AsArray().RemoveAt(1);
            result["page_count"] = 1;
        });
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-page-closure-mismatch");
    }

    [Fact]
    public void Validator_rejects_an_extra_page()
    {
        var fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result =>
        {
            result["pages"]!.AsArray().Add(new JsonObject
            {
                ["page"] = 3,
                ["sha256"] = new string('c', 64),
                ["size_bytes"] = 10
            });
            result["page_count"] = 3;
        });
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-page-closure-mismatch");
    }

    [Fact]
    public void Validator_rejects_swapped_pages()
    {
        var fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result =>
        {
            var pages = result["pages"]!.AsArray();
            (pages[0], pages[1]) = (pages[1]!.DeepClone(), pages[0]!.DeepClone());
        });
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-page-closure-mismatch");
    }

    [Fact]
    public void Validator_rejects_a_stale_input()
    {
        var fixture = ProduceTwoPageResult(out _);
        File.WriteAllBytes(fixture.Input, BuildPdf((200.0, 200.0, "replaced input bytes")));
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-request-input-mismatch");
        Assert.Equal(FileSha(fixture.Input), verdict["recomputed"]!["input_sha256"]!.GetValue<string>());
    }

    [Fact]
    public void Validator_rejects_tampered_output_bytes()
    {
        var fixture = ProduceTwoPageResult(out _);
        File.WriteAllBytes(fixture.Output, BuildPdf((612.0, 792.0, "Forged single page output")));
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-output-bytes-mismatch");
    }

    [Fact]
    public void Validator_rejects_a_wrong_runtime_version()
    {
        var fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result => result["runtime"]!["version"] = "0.0.1");
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-runtime-mismatch");
    }

    [Fact]
    public void Validator_rejects_forged_provenance_on_a_non_wps_backend()
    {
        var fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result =>
        {
            var forged = new JsonObject
            {
                ["schema"] = "tiwater.convert-native-render-provenance/v1",
                ["backend"] = "wps-writer",
                ["wps"] = new JsonObject
                {
                    ["package"] = "wps-office",
                    ["build_version"] = "12.1.0",
                    ["executable_sha256"] = new string('a', 64)
                },
                ["runtime"] = new JsonObject
                {
                    ["os_description"] = RuntimeInformation.OSDescription,
                    ["os_architecture"] = "x64",
                    ["process_architecture"] = "x64",
                    ["framework_description"] = RuntimeInformation.FrameworkDescription
                },
                ["fonts"] = new JsonObject
                {
                    ["source"] = "fontconfig-family-style-file-sha256",
                    ["count"] = 1,
                    ["sha256"] = new string('b', 64)
                },
                ["input"] = new JsonObject
                {
                    ["sha256"] = FileSha(fixture.Input),
                    ["size_bytes"] = new FileInfo(fixture.Input).Length
                },
                ["output"] = new JsonObject
                {
                    ["sha256"] = FileSha(fixture.Output),
                    ["size_bytes"] = new FileInfo(fixture.Output).Length
                },
                ["page_count"] = 2
            };
            result["native_render_provenance"] = forged;
            result["provenance_sha256"] = Sha(Canonical(forged));
        });
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-provenance-unexpected");
    }

    [Fact]
    public void Validator_rejects_provenance_with_a_rebound_input_even_when_rehashed_honestly()
    {
        var (fixture, _) = ProduceWpsResult();
        var mutated = Load(fixture.Result);
        mutated["native_render_provenance"]!["input"]!["sha256"] = new string('0', 64);
        // The forger honestly recomputes the provenance hash over the forged object.
        mutated["provenance_sha256"] = Sha(Canonical(mutated["native_render_provenance"]));
        File.WriteAllText(fixture.Result, mutated.ToJsonString());
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-provenance-binding-mismatch");
        Assert.DoesNotContain(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-provenance-hash-mismatch");
    }

    [Fact]
    public void Validator_rejects_provenance_with_a_stale_hash()
    {
        var (fixture, _) = ProduceWpsResult();
        var mutated = Load(fixture.Result);
        mutated["native_render_provenance"]!["input"]!["sha256"] = new string('0', 64);
        File.WriteAllText(fixture.Result, mutated.ToJsonString());
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-provenance-hash-mismatch");
    }

    [Fact]
    public void Validator_ignores_producer_self_attested_verdicts()
    {
        var withVerdictClaim = ProduceTwoPageResult(out _);
        MutateResult(withVerdictClaim, result => result["verdict"] = "pass");
        var verdictClaim = ValidateExpectingFailure(withVerdictClaim);
        Assert.Contains(verdictClaim["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-result-invalid");

        var withDecisionClaim = ProduceTwoPageResult(out _);
        MutateResult(withDecisionClaim, result =>
        {
            result["decision"] = "pass";
            result["output"]!["sha256"] = new string('f', 64);
        });
        var decisionClaim = ValidateExpectingFailure(withDecisionClaim);
        Assert.Equal("failed", decisionClaim["decision"]!.GetValue<string>());
    }

    [Fact]
    public void Render_and_verdict_outputs_are_immutable()
    {
        var fixture = Fixture(".pdf");
        File.WriteAllBytes(fixture.Input, BuildPdf((612.0, 792.0, "immutable")));
        WriteRequest(fixture, "pdf", "passthrough");

        File.WriteAllText(fixture.Result, "occupied");
        Assert.ThrowsAny<Exception>(() =>
            RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result]));

        fixture = Fixture(".pdf");
        File.WriteAllBytes(fixture.Input, BuildPdf((612.0, 792.0, "immutable")));
        File.WriteAllText(fixture.Output, "occupied");
        WriteRequest(fixture, "pdf", "passthrough");
        Assert.ThrowsAny<Exception>(() =>
            RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result]));

        fixture = Fixture(".pdf");
        File.WriteAllBytes(fixture.Input, BuildPdf((612.0, 792.0, "immutable")));
        WriteRequest(fixture, "pdf", "passthrough");
        Assert.Equal(0, RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result]));
        File.WriteAllText(fixture.Verdict, "occupied");
        Assert.ThrowsAny<Exception>(() => RenderEvidenceValidatorCommand.Run(
            ["--request", fixture.Request, "--result", fixture.Result, "--output", fixture.Verdict]));
    }

    [Fact]
    public void Closed_schemas_reject_extra_and_missing_fields()
    {
        var fixture = Fixture(".pdf");
        File.WriteAllBytes(fixture.Input, BuildPdf((612.0, 792.0, "closure")));
        WriteRequest(fixture, "pdf", "passthrough");
        var request = Load(fixture.Request);
        request["unexpected"] = true;
        File.WriteAllText(fixture.Request, request.ToJsonString());
        Assert.Throws<InvalidOperationException>(() =>
            RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result]));

        fixture = Fixture(".pdf");
        File.WriteAllBytes(fixture.Input, BuildPdf((612.0, 792.0, "closure")));
        WriteRequest(fixture, "pdf", "passthrough");
        request = Load(fixture.Request);
        request.Remove("options");
        File.WriteAllText(fixture.Request, request.ToJsonString());
        Assert.Throws<InvalidOperationException>(() =>
            RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result]));

        fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result => result["extra_field"] = "not-closed");
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-result-invalid");
    }

    [Fact]
    public void Validator_rejects_a_result_bound_to_a_different_request()
    {
        var fixture = ProduceTwoPageResult(out _);
        MutateResult(fixture, result => result["request_sha256"] = new string('e', 64));
        var verdict = ValidateExpectingFailure(fixture);
        Assert.Contains(verdict["findings"]!.AsArray(), finding =>
            finding!["code"]!.GetValue<string>() == "render-request-hash-mismatch");
    }

    private sealed record RenderFixture(string Root, string Input, string Output, string Request, string Result, string Verdict);

    private static RenderFixture Fixture(string inputExtension)
    {
        var root = Path.Combine(Path.GetTempPath(), $"render-contract-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        return new RenderFixture(
            root,
            Path.Combine(root, $"input{inputExtension}"),
            Path.Combine(root, "output.pdf"),
            Path.Combine(root, "request.json"),
            Path.Combine(root, "result.json"),
            Path.Combine(root, "verdict.json"));
    }

    private static RenderFixture ProduceTwoPageResult(out (double, double, string)[] pages)
    {
        pages = [(612.0, 792.0, "First rendered page"), (612.0, 792.0, "Second rendered page")];
        var fixture = Fixture(".docx");
        File.WriteAllBytes(fixture.Input, "synthetic two page source"u8.ToArray());
        WriteRequest(fixture, "docx", "libreoffice");
        var pdf = BuildPdf(pages);
        RenderBackend fake = invocation =>
        {
            File.WriteAllBytes(invocation.Output, pdf);
            return new OfficePdfConversionResult("libreoffice");
        };
        Assert.Equal(0, RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result], fake));
        return fixture;
    }

    private static (RenderFixture Fixture, JsonObject Result) ProduceWpsResult()
    {
        var fixture = Fixture(".docx");
        File.WriteAllBytes(fixture.Input, "synthetic wps source"u8.ToArray());
        WriteRequest(fixture, "docx", "wps-writer");
        var pdf = BuildPdf((612.0, 792.0, "Wps rendered page one"), (612.0, 792.0, "Wps rendered page two"));
        RenderBackend fake = invocation =>
        {
            File.WriteAllBytes(invocation.Output, pdf);
            var provenance = new NativeRenderProvenance(
                "tiwater.convert-native-render-provenance/v1",
                "wps-writer",
                new NativeRenderWpsIdentity("wps-office", "12.1.0-r1", new string('a', 64)),
                new NativeRenderRuntimeIdentity(
                    RuntimeInformation.OSDescription,
                    RuntimeInformation.OSArchitecture.ToString().ToLowerInvariant(),
                    RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant(),
                    RuntimeInformation.FrameworkDescription),
                new NativeRenderFontInventory("fontconfig-family-style-file-sha256", 1, new string('b', 64)),
                new NativeRenderFileIdentity(FileSha(invocation.Input), new FileInfo(invocation.Input).Length),
                new NativeRenderFileIdentity(FileSha(invocation.Output), new FileInfo(invocation.Output).Length),
                2);
            return new OfficePdfConversionResult("wps-writer", NativeRenderProvenance: provenance);
        };
        Assert.Equal(0, RenderCommand.Run(["--request", fixture.Request, "--output", fixture.Result], fake));
        return (fixture, Load(fixture.Result));
    }

    private static void WriteRequest(RenderFixture fixture, string format, string backend)
    {
        var inputBytes = File.ReadAllBytes(fixture.Input);
        var request = new JsonObject
        {
            ["schema"] = "tiwater.convert-render-request/v1",
            ["request_id"] = $"req-{Guid.NewGuid():N}",
            ["format"] = format,
            ["input"] = new JsonObject
            {
                ["path"] = fixture.Input,
                ["sha256"] = Sha(inputBytes),
                ["size_bytes"] = inputBytes.Length,
                ["artifact_version_id"] = "v1"
            },
            ["output"] = new JsonObject
            {
                ["path"] = fixture.Output,
                ["media_type"] = "application/pdf"
            },
            ["renderer"] = new JsonObject { ["backend"] = backend },
            ["runtime"] = new JsonObject
            {
                ["id"] = "tiwater-convert",
                ["version"] = RuntimeIdentity.Version
            },
            ["target_format"] = "pdf",
            ["options"] = new JsonObject(),
            ["result_contract"] = new JsonObject
            {
                ["id"] = "tiwater.convert-render-result/v1",
                ["sha256"] = DeployedResultSchemaSha()
            }
        };
        File.WriteAllText(fixture.Request, $"{request.ToJsonString()}\n");
    }

    private static string DeployedResultSchemaSha()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "schemas", "render-result-v1.schema.json");
        Assert.True(File.Exists(path), $"deployed render result schema missing: {path}");
        return Sha(File.ReadAllBytes(path));
    }

    private static void MutateResult(RenderFixture fixture, Action<JsonObject> mutate)
    {
        var result = Load(fixture.Result);
        mutate(result);
        File.WriteAllText(fixture.Result, result.ToJsonString());
    }

    private static JsonObject ValidateExpectingFailure(RenderFixture fixture)
    {
        Assert.Equal(1, RenderEvidenceValidatorCommand.Run(
            ["--request", fixture.Request, "--result", fixture.Result, "--output", fixture.Verdict]));
        var verdict = Load(fixture.Verdict);
        Assert.Equal("failed", verdict["decision"]!.GetValue<string>());
        Assert.NotEmpty(verdict["findings"]!.AsArray());
        return verdict;
    }

    private static JsonObject Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path))!.AsObject();

    private static byte[] BuildPdf(params (double Width, double Height, string Text)[] pages)
    {
        var objects = new SortedDictionary<int, byte[]>
        {
            [1] = Ascii("<< /Type /Catalog /Pages 2 0 R >>"),
            [2] = Ascii($"<< /Type /Pages /Kids [{string.Join(" ", pages.Select((_, index) => $"{4 + 2 * index} 0 R"))}] /Count {pages.Length} >>"),
            [3] = Ascii("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        };
        for (var index = 0; index < pages.Length; index++)
        {
            var (width, height, text) = pages[index];
            objects[4 + 2 * index] = Ascii(
                $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {Number(width)} {Number(height)}] " +
                $"/Resources << /Font << /F1 3 0 R >> >> /Contents {5 + 2 * index} 0 R >>");
            var stream = Ascii($"BT /F1 24 Tf 72 {Number(height - 72)} Td ({text}) Tj ET");
            objects[5 + 2 * index] = Ascii($"<< /Length {stream.Length} >>\nstream\n")
                .Concat(stream)
                .Concat(Ascii("\nendstream"))
                .ToArray();
        }

        var output = new List<byte>(Ascii("%PDF-1.4\n"));
        var offsets = new Dictionary<int, int>();
        foreach (var (number, body) in objects)
        {
            offsets[number] = output.Count;
            output.AddRange(Ascii($"{number} 0 obj\n"));
            output.AddRange(body);
            output.AddRange(Ascii("\nendobj\n"));
        }
        var xref = output.Count;
        output.AddRange(Ascii($"xref\n0 {objects.Count + 1}\n"));
        output.AddRange(Ascii("0000000000 65535 f \n"));
        foreach (var number in objects.Keys)
            output.AddRange(Ascii($"{offsets[number]:D10} 00000 n \n"));
        output.AddRange(Ascii($"trailer\n<< /Size {objects.Count + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n"));
        return output.ToArray();
    }

    private static string Number(double value) =>
        value.ToString("R", CultureInfo.InvariantCulture);

    private static byte[] Ascii(string value) => Encoding.ASCII.GetBytes(value);

    private static string FileSha(string path) => Sha(File.ReadAllBytes(path));

    private static string Sha(byte[] value) =>
        System.Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();

    private static string Sha(string value) => Sha(Encoding.UTF8.GetBytes(value));

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
