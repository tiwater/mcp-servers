using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Xlsx.Cli;
using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using Tiwater.RuntimeContracts;
using Xunit;

namespace Dockit.Xlsx.Tests;

public sealed class RuntimeIdentityTests
{
    private static readonly SchemaIdentity IdentifyPayloadSchema =
        new("tiwater.runtime.identify-payload", "1.0.0");

    [Fact]
    public void Capabilities_describe_non_mutating_xlsx_and_xls_identity_commands()
    {
        var descriptor = XlsxRuntimeIdentity.Capabilities();

        Assert.Equal("1.0.0", descriptor.SchemaVersion);
        Assert.Equal("runtime-capabilities", descriptor.DescriptorType);
        Assert.Equal(new PackageIdentity("tiwater.xlsx.cli", "0.1.35"), descriptor.Package);
        Assert.Equal(new RuntimeIdentity("office", "tiwater-xlsx", "0.1.35"), descriptor.Runtime);
        Assert.Equal(
            new SchemaIdentity(
                "https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json",
                "1.0.0"),
            descriptor.EvidenceSchema);
        Assert.Equal("capabilities", descriptor.DescriptorCommand.Command);
        Assert.Equal(["--json"], descriptor.DescriptorCommand.Arguments);
        Assert.False(descriptor.DescriptorCommand.Mutates);
        Assert.Equal("identify", descriptor.IdentifyProbe.Command);
        Assert.Equal(["<input>", "--json"], descriptor.IdentifyProbe.Arguments);
        Assert.False(descriptor.IdentifyProbe.Mutates);
        Assert.Equal(["supported", "unsupported", "failed"], descriptor.IdentifyProbe.Outcomes);
        Assert.Contains(descriptor.SupportedKinds, kind =>
            kind.FileKind == "xlsx"
            && kind.MediaTypes.SequenceEqual([
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ])
            && kind.SignatureKinds.SequenceEqual(["ooxml-spreadsheet-main-part"]));
        Assert.Contains(descriptor.SupportedKinds, kind =>
            kind.FileKind == "xls"
            && kind.MediaTypes.SequenceEqual(["application/vnd.ms-excel"])
            && kind.SignatureKinds.SequenceEqual(["ole-compound-hssf-workbook"]));
        Assert.Contains(descriptor.Commands, command => command.Name == "capabilities" && !command.Mutates);
        Assert.Contains(descriptor.Commands, command => command.Name == "identify" && !command.Mutates);
        Assert.Equal("runtime-native-only", descriptor.IdentityPolicy.NativeIds);
        Assert.Equal("deterministic-and-explicit", descriptor.IdentityPolicy.DerivedIds);
        Assert.Equal("parent-object-id-required-for-non-root", descriptor.IdentityPolicy.Containment);
    }

    [Fact]
    public async Task Cli_capabilities_emits_the_versioned_descriptor_as_json()
    {
        var result = await RunXlsxCliAsync("capabilities", "--json");

        Assert.Equal(0, result.ExitCode);
        Assert.Equal(string.Empty, result.Stderr);
        using var document = JsonDocument.Parse(result.Stdout);
        var root = document.RootElement;
        Assert.Equal("1.0.0", root.GetProperty("schemaVersion").GetString());
        Assert.Equal("runtime-capabilities", root.GetProperty("descriptorType").GetString());
        Assert.Equal("tiwater.xlsx.cli", root.GetProperty("package").GetProperty("name").GetString());
        Assert.Equal("0.1.35", root.GetProperty("package").GetProperty("version").GetString());
        Assert.False(root.GetProperty("descriptorCommand").GetProperty("mutates").GetBoolean());
        Assert.Equal(
            ["supported", "unsupported", "failed"],
            root.GetProperty("identifyProbe").GetProperty("outcomes")
                .EnumerateArray().Select(item => item.GetString()!).ToArray());
        Assert.Contains(
            root.GetProperty("supportedKinds").EnumerateArray(),
            kind => kind.GetProperty("fileKind").GetString() == "xlsx");
        Assert.Contains(
            root.GetProperty("supportedKinds").EnumerateArray(),
            kind => kind.GetProperty("fileKind").GetString() == "xls");
    }

    [Fact]
    public void Renamed_valid_xlsx_is_supported_from_package_bytes()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateSpreadsheet(path);

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertSupportedEvidence(evidence, path, "xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ooxml-spreadsheet-main-part");
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.Contains("/xl/workbook.xml", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Renamed_valid_legacy_xls_is_supported_only_after_hssf_open()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateLegacySpreadsheet(path);

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertSupportedEvidence(
                evidence,
                path,
                "xls",
                "application/vnd.ms-excel",
                "ole-compound-hssf-workbook");
            Assert.Contains("ole-compound:magic=d0cf11e0a1b11ae1", evidence.File.Signature.Evidence);
            Assert.Contains("npoi-hssf:workbook-opened", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fake_xlsx_extension_is_unsupported_not_failed()
    {
        var path = TemporaryPath(".xlsx");
        try
        {
            File.WriteAllText(path, "not a spreadsheet package");

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Zip_without_spreadsheet_main_part_is_unsupported()
    {
        var path = TemporaryPath(".xlsx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteZipEntry(
                    archive,
                    "[Content_Types].xml",
                    """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="xml" ContentType="application/xml" />
                    </Types>
                    """);
            }

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.Contains("spreadsheet-main-part", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Ole_magic_without_an_openable_hssf_workbook_is_unsupported()
    {
        var path = TemporaryPath(".xls");
        try
        {
            File.WriteAllBytes(path, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
            Assert.Contains("ole-compound:magic=d0cf11e0a1b11ae1", evidence.File.Signature.Evidence);
            Assert.Contains("npoi-hssf:workbook-rejected", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Ole_container_with_a_corrupt_workbook_stream_is_unsupported_not_an_untyped_exception()
    {
        var path = TemporaryPath(".bin");
        try
        {
            var compoundFile = new NPOIFSFileSystem();
            try
            {
                using var invalidWorkbook = new MemoryStream([0x01, 0x02, 0x03, 0x04]);
                compoundFile.CreateDocument(invalidWorkbook, "Workbook");
                using var output = File.Create(path);
                compoundFile.WriteFileSystem(output);
            }
            finally
            {
                compoundFile.Close();
            }

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
            Assert.Contains("npoi-hssf:workbook-rejected", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Docx_renamed_to_xlsx_is_unsupported()
    {
        var path = TemporaryPath(".xlsx");
        try
        {
            CreateWordDocument(path);

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.Contains("spreadsheet-main-part", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("[Content_Types].xml", "[content_types].xml")]
    [InlineData("_rels/.rels", "_RELS/.RELS")]
    [InlineData("xl/workbook.xml", "XL/WORKBOOK.XML")]
    [InlineData("custom/part.bin", "CUSTOM/PART.BIN")]
    [InlineData("custom/a.xml", "CUSTOM/%61.XML")]
    public void Package_wide_opc_equivalent_entries_are_unsupported(
        string originalName,
        string collidingName)
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateSpreadsheet(path);
            AddOpcEquivalentZipEntry(path, originalName, collidingName);

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.Contains("opc-part-uri-collision", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("/absolute.xml")]
    [InlineData("xl//invalid.xml")]
    [InlineData("xl/../invalid.xml")]
    [InlineData("xl\\invalid.xml")]
    [InlineData("xl/%ZZ-invalid.xml")]
    public void Invalid_opc_part_uri_entry_names_are_unsupported(string invalidName)
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateSpreadsheet(path);
            AddZipEntry(path, invalidName, "invalid package entry");

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertUnsupportedEvidence(evidence, path);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.Contains("opc-part-uri-invalid", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Non_colliding_directory_entries_do_not_change_a_renamed_workbook_identity()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateSpreadsheet(path);
            AddZipEntry(path, "custom/", string.Empty);
            AddZipEntry(path, "custom/evidence.bin", "fixture");

            var evidence = XlsxRuntimeIdentity.Identify(path);

            AssertSupportedEvidence(
                evidence,
                path,
                "xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "ooxml-spreadsheet-main-part");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Missing_source_is_typed_source_read_failure_without_fake_hash()
    {
        var path = TemporaryPath(".xlsx");

        var evidence = XlsxRuntimeIdentity.Identify(path);

        Assert.Equal("failed", evidence.Status);
        Assert.Equal("source-read", evidence.FailureStage);
        Assert.Null(evidence.Source);
        Assert.Null(evidence.File.FileKind);
        Assert.Null(evidence.File.MediaType);
        Assert.Equal("not-checked", evidence.File.Signature.Status);
        Assert.Empty(evidence.File.Signature.Evidence);
        Assert.Empty(evidence.Objects);
        Assert.NotEmpty(evidence.Errors);
        Assert.Equal(
            EvidenceEnvelope.IdentifyCanonicalJson(evidence.Payload, IdentifyPayloadSchema),
            evidence.Artifact);
    }

    [Fact]
    public async Task Cli_identify_emits_failed_evidence_for_a_missing_source()
    {
        var path = TemporaryPath(".xlsx");

        var result = await RunXlsxCliAsync("identify", path, "--json");

        Assert.Equal(1, result.ExitCode);
        using var document = JsonDocument.Parse(result.Stdout);
        Assert.Equal("failed", document.RootElement.GetProperty("status").GetString());
        Assert.Equal("source-read", document.RootElement.GetProperty("failureStage").GetString());
        Assert.Equal(JsonValueKind.Null, document.RootElement.GetProperty("source").ValueKind);
        Assert.Equal(
            "not-checked",
            document.RootElement.GetProperty("file").GetProperty("signature").GetProperty("status").GetString());
    }

    [Fact]
    public void Repeated_identify_is_byte_for_byte_deterministic()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateSpreadsheet(path);

            var first = JsonSerializer.Serialize(XlsxRuntimeIdentity.Identify(path), RuntimeJson.Options);
            var second = JsonSerializer.Serialize(XlsxRuntimeIdentity.Identify(path), RuntimeJson.Options);

            Assert.Equal(first, second);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static void AssertSupportedEvidence(
        RuntimeEvidenceEnvelope evidence,
        string path,
        string expectedKind,
        string expectedMediaType,
        string expectedSignatureKind)
    {
        Assert.Equal("1.0.0", evidence.SchemaVersion);
        Assert.Equal("runtime-evidence", evidence.EnvelopeType);
        Assert.Equal("identify", evidence.Probe);
        var diagnostic = JsonSerializer.Serialize(evidence, RuntimeJson.Options);
        if (expectedKind == "xlsx")
        {
            diagnostic += Environment.NewLine + ReadZipEntry(path, "[Content_Types].xml");
        }
        Assert.True(evidence.Status == "supported", diagnostic);
        Assert.Null(evidence.FailureStage);
        Assert.Equal(new PackageIdentity("tiwater.xlsx.cli", "0.1.35"), evidence.Package);
        Assert.Equal(new RuntimeIdentity("office", "tiwater-xlsx", "0.1.35"), evidence.Runtime);
        Assert.Equal(
            new SchemaIdentity(
                "https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json",
                "1.0.0"),
            evidence.EvidenceSchema);
        Assert.NotNull(evidence.Source);
        var bytes = File.ReadAllBytes(path);
        var sha256 = Convert.ToHexStringLower(SHA256.HashData(bytes));
        Assert.Equal(Path.GetFullPath(path), evidence.Source.Path);
        Assert.Equal(bytes.Length, evidence.Source.SizeBytes);
        Assert.Equal(sha256, evidence.Source.Sha256);
        Assert.Equal($"sha256:{sha256}", evidence.Source.ContentId);
        Assert.Equal(expectedKind, evidence.File.FileKind);
        Assert.Equal(expectedMediaType, evidence.File.MediaType);
        Assert.Equal("matched", evidence.File.Signature.Status);
        Assert.Equal(expectedSignatureKind, evidence.File.Signature.Kind);
        Assert.NotEmpty(evidence.File.Signature.Evidence);
        Assert.Empty(evidence.Objects);
        Assert.Empty(evidence.Warnings);
        Assert.Empty(evidence.Errors);
        Assert.Equal(
            EvidenceEnvelope.IdentifyCanonicalJson(evidence.Payload, IdentifyPayloadSchema),
            evidence.Artifact);
    }

    private static void AssertUnsupportedEvidence(RuntimeEvidenceEnvelope evidence, string path)
    {
        Assert.Equal("unsupported", evidence.Status);
        Assert.Null(evidence.FailureStage);
        Assert.NotNull(evidence.Source);
        Assert.Equal(Path.GetFullPath(path), evidence.Source.Path);
        Assert.Null(evidence.File.FileKind);
        Assert.Null(evidence.File.MediaType);
        Assert.Equal("mismatched", evidence.File.Signature.Status);
        Assert.Empty(evidence.Objects);
        Assert.Empty(evidence.Errors);
        Assert.Equal(
            EvidenceEnvelope.IdentifyCanonicalJson(evidence.Payload, IdentifyPayloadSchema),
            evidence.Artifact);
    }

    private static string TemporaryPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"xlsx-runtime-{Guid.NewGuid():N}{extension}");

    private static void CreateSpreadsheet(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());
        workbookPart.Workbook.Save();
    }

    private static void CreateLegacySpreadsheet(string path)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("identity fixture");
        using var stream = File.Create(path);
        workbook.Write(stream);
    }

    private static void CreateWordDocument(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(
            new DocumentFormat.OpenXml.Wordprocessing.Run(
                new DocumentFormat.OpenXml.Wordprocessing.Text("identity fixture")))));
        mainPart.Document.Save();
    }

    private static void WriteZipEntry(ZipArchive archive, string name, string value)
    {
        var entry = archive.CreateEntry(name);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream);
        writer.Write(value);
    }

    private static void AddOpcEquivalentZipEntry(string path, string originalName, string collidingName)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var original = archive.GetEntry(originalName);
        if (original is null)
        {
            original = archive.CreateEntry(originalName);
            using var originalStream = original.Open();
            originalStream.Write([0x01, 0x02, 0x03]);
        }

        byte[] originalBytes;
        using (var originalStream = original.Open())
        using (var copy = new MemoryStream())
        {
            originalStream.CopyTo(copy);
            originalBytes = copy.ToArray();
        }

        var collision = archive.CreateEntry(collidingName);
        using var collisionStream = collision.Open();
        collisionStream.Write(originalBytes);
    }

    private static void AddZipEntry(string path, string name, string value)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        WriteZipEntry(archive, name, value);
    }

    private static string ReadZipEntry(string path, string name)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry(name)!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunXlsxCliAsync(params string[] args)
    {
        var executable = Path.Combine(AppContext.BaseDirectory, OperatingSystem.IsWindows() ? "xlsx.exe" : "xlsx");
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        foreach (var argument in args) startInfo.ArgumentList.Add(argument);

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start xlsx CLI.");
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return (process.ExitCode, stdout.Trim(), stderr.Trim());
    }
}
