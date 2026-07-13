using System.IO.Compression;
using System.Reflection;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using NPOI.HSSF.UserModel;
using Tiwater.RuntimeContracts;

namespace Dockit.Xlsx.Cli;

public static class XlsxRuntimeIdentity
{
    private const string PackageName = "tiwater.xlsx.cli";
    private const string RuntimeName = "tiwater-xlsx";
    private const string EvidenceSchemaId = "https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json";
    private const string CapabilitiesSchemaId = "https://tiwater.dev/contracts/runtime/runtime-capabilities.schema.json";
    private const string PayloadSchemaId = "tiwater.runtime.identify-payload";
    private const string XlsxMediaType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    private const string XlsMediaType = "application/vnd.ms-excel";
    private const string SpreadsheetMainPartContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string XlsxSignatureKind = "ooxml-spreadsheet-main-part";
    private const string XlsSignatureKind = "ole-compound-hssf-workbook";
    private const string GenericSignatureKind = "spreadsheet-package-signature";

    private static readonly byte[] ZipLocalFileMagic = [0x50, 0x4b, 0x03, 0x04];
    private static readonly byte[] OleCompoundMagic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    private static readonly string PackageVersion = ResolvePackageVersion();
    private static readonly SchemaIdentity EvidenceSchema =
        new(EvidenceSchemaId, RuntimeContractVersions.EvidenceEnvelope);
    private static readonly SchemaIdentity PayloadSchema =
        new(PayloadSchemaId, RuntimeContractVersions.EvidenceEnvelope);

    public static RuntimeCapabilityDescriptor Capabilities() => new(
        RuntimeContractVersions.Capabilities,
        "runtime-capabilities",
        PackageIdentity(),
        RuntimeIdentity(),
        EvidenceSchema,
        new DiscoveryCommand("capabilities", ["--json"], false),
        new IdentifyProbe("identify", ["<input>", "--json"], false, ["supported", "unsupported", "failed"]),
        [
            new SupportedKind("xlsx", [XlsxMediaType], [XlsxSignatureKind]),
            new SupportedKind("xls", [XlsMediaType], [XlsSignatureKind]),
        ],
        [
            new RuntimeCommand(
                "capabilities",
                false,
                new SchemaIdentity(CapabilitiesSchemaId, RuntimeContractVersions.Capabilities)),
            new RuntimeCommand("identify", false, EvidenceSchema),
        ],
        new IdentityPolicy(
            "runtime-native-only",
            "deterministic-and-explicit",
            "parent-object-id-required-for-non-root"));

    public static RuntimeEvidenceEnvelope Identify(string filePath)
    {
        string resolvedPath;
        byte[] sourceBytes;
        try
        {
            resolvedPath = Path.GetFullPath(filePath);
            sourceBytes = File.ReadAllBytes(resolvedPath);
        }
        catch (Exception exception) when (
            exception is IOException
                or UnauthorizedAccessException
                or ArgumentException
                or NotSupportedException)
        {
            return CreateEnvelope(
                status: "failed",
                failureStage: "source-read",
                source: null,
                file: new RuntimeFileEvidence(
                    null,
                    null,
                    new SignatureEvidence("not-checked", GenericSignatureKind, [])),
                payload: JsonSerializer.SerializeToElement(
                    new { failureClass = "source-read-error" },
                    RuntimeJson.Options),
                errors: [new ContractFinding("source-read-failed", "The source bytes could not be read.")]);
        }

        var source = FileIdentity.IdentifyBytes(resolvedPath, sourceBytes);
        SignatureInspection inspection;
        try
        {
            inspection = InspectSignature(sourceBytes);
        }
        catch (Exception exception) when (
            exception is IOException
                or InvalidDataException
                or XmlException
                or OpenXmlPackageException)
        {
            return CreateEnvelope(
                status: "failed",
                failureStage: "signature-inspection",
                source,
                file: new RuntimeFileEvidence(
                    null,
                    null,
                    new SignatureEvidence("unknown", GenericSignatureKind, [])),
                payload: JsonSerializer.SerializeToElement(
                    new { failureClass = "signature-inspection-error" },
                    RuntimeJson.Options),
                errors: [new ContractFinding(
                    "signature-inspection-failed",
                    "The spreadsheet signature could not be inspected.")]);
        }

        if (!inspection.Matched)
        {
            return CreateEnvelope(
                status: "unsupported",
                failureStage: null,
                source,
                file: new RuntimeFileEvidence(
                    null,
                    null,
                    new SignatureEvidence("mismatched", inspection.SignatureKind, inspection.Evidence)),
                payload: JsonSerializer.SerializeToElement(
                    new { recognized = false, reason = inspection.Reason },
                    RuntimeJson.Options),
                errors: []);
        }

        return CreateEnvelope(
            status: "supported",
            failureStage: null,
            source,
            file: new RuntimeFileEvidence(
                inspection.FileKind,
                inspection.MediaType,
                new SignatureEvidence("matched", inspection.SignatureKind, inspection.Evidence)),
            payload: JsonSerializer.SerializeToElement(
                new { recognized = true, fileKind = inspection.FileKind },
                RuntimeJson.Options),
            errors: []);
    }

    private static SignatureInspection InspectSignature(ReadOnlyMemory<byte> sourceBytes)
    {
        if (StartsWith(sourceBytes.Span, ZipLocalFileMagic))
        {
            return InspectOpenXmlSpreadsheet(sourceBytes);
        }
        if (StartsWith(sourceBytes.Span, OleCompoundMagic))
        {
            return InspectLegacySpreadsheet(sourceBytes);
        }
        return SignatureInspection.Unsupported(
            GenericSignatureKind,
            "unrecognized-container-signature",
            ["container-signature:unrecognized"]);
    }

    private static SignatureInspection InspectOpenXmlSpreadsheet(ReadOnlyMemory<byte> sourceBytes)
    {
        var evidence = new List<string> { "zip:local-file-header" };
        try
        {
            using var archiveStream = new MemoryStream(sourceBytes.ToArray(), writable: false);
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: false);
            var contentTypesEntries = archive.Entries
                .Where(entry => entry.FullName == "[Content_Types].xml")
                .ToArray();
            if (contentTypesEntries.Length != 1)
            {
                evidence.Add($"[Content_Types].xml:count={contentTypesEntries.Length}");
                return SignatureInspection.Unsupported(
                    XlsxSignatureKind,
                    "content-types-part-missing-or-ambiguous",
                    evidence);
            }

            XDocument contentTypes;
            using (var contentTypesStream = contentTypesEntries[0].Open())
            using (var reader = XmlReader.Create(contentTypesStream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
            }))
            {
                contentTypes = XDocument.Load(reader, LoadOptions.None);
            }

            XNamespace contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
            string mainPartName;
            using (var packageStream = new MemoryStream(sourceBytes.ToArray(), writable: false))
            using (var spreadsheet = SpreadsheetDocument.Open(packageStream, false))
            {
                if (spreadsheet.WorkbookPart is null)
                {
                    evidence.Add("spreadsheet-main-part:openxml-workbook-part-missing");
                    return SignatureInspection.Unsupported(
                        XlsxSignatureKind,
                        "spreadsheet-main-part-not-openable",
                        evidence);
                }
                mainPartName = spreadsheet.WorkbookPart.Uri.ToString();
            }

            var overrideContentTypes = contentTypes
                .Descendants(contentTypesNamespace + "Override")
                .Where(element => string.Equals(
                    (string?)element.Attribute("PartName"),
                    mainPartName,
                    StringComparison.Ordinal))
                .Select(element => (string?)element.Attribute("ContentType"))
                .ToArray();
            if (overrideContentTypes.Length > 1)
            {
                evidence.Add($"spreadsheet-main-part:override-count={overrideContentTypes.Length}");
                return SignatureInspection.Unsupported(
                    XlsxSignatureKind,
                    "spreadsheet-main-part-content-type-ambiguous",
                    evidence);
            }

            var effectiveContentType = overrideContentTypes.SingleOrDefault();
            if (effectiveContentType is null)
            {
                var extension = Path.GetExtension(mainPartName).TrimStart('.');
                var defaultContentTypes = contentTypes
                    .Descendants(contentTypesNamespace + "Default")
                    .Where(element => string.Equals(
                        (string?)element.Attribute("Extension"),
                        extension,
                        StringComparison.OrdinalIgnoreCase))
                    .Select(element => (string?)element.Attribute("ContentType"))
                    .ToArray();
                if (defaultContentTypes.Length == 1) effectiveContentType = defaultContentTypes[0];
            }
            if (!string.Equals(effectiveContentType, SpreadsheetMainPartContentType, StringComparison.Ordinal))
            {
                evidence.Add("spreadsheet-main-part:content-type-mismatch");
                return SignatureInspection.Unsupported(
                    XlsxSignatureKind,
                    "spreadsheet-main-part-content-type-mismatch",
                    evidence);
            }

            var archiveEntryName = mainPartName.TrimStart('/');
            var mainPartCount = archive.Entries.Count(entry => entry.FullName == archiveEntryName);
            if (mainPartCount != 1)
            {
                evidence.Add($"spreadsheet-main-part:entry-count={mainPartCount}");
                return SignatureInspection.Unsupported(
                    XlsxSignatureKind,
                    "spreadsheet-main-part-missing-or-ambiguous",
                    evidence);
            }

            evidence.Add($"[Content_Types].xml:{mainPartName}={SpreadsheetMainPartContentType}");
            evidence.Add($"part:{archiveEntryName}");
            evidence.Add("openxml:workbook-part-opened");
            return SignatureInspection.Supported(
                "xlsx",
                XlsxMediaType,
                XlsxSignatureKind,
                evidence);
        }
        catch (Exception exception) when (
            exception is InvalidDataException
                or XmlException
                or OpenXmlPackageException
                or IOException)
        {
            evidence.Add("spreadsheet-main-part:package-rejected");
            return SignatureInspection.Unsupported(
                XlsxSignatureKind,
                "spreadsheet-package-invalid",
                evidence);
        }
    }

    private static SignatureInspection InspectLegacySpreadsheet(ReadOnlyMemory<byte> sourceBytes)
    {
        var evidence = new List<string> { "ole-compound:magic=d0cf11e0a1b11ae1" };
        try
        {
            using var stream = new MemoryStream(sourceBytes.ToArray(), writable: false);
            using var workbook = new HSSFWorkbook(stream);
            _ = workbook.NumberOfSheets;
            evidence.Add("npoi-hssf:workbook-opened");
            return SignatureInspection.Supported("xls", XlsMediaType, XlsSignatureKind, evidence);
        }
        catch (Exception exception) when (
            exception is IOException
                or InvalidDataException
                or ArgumentException
                or InvalidOperationException
                or IndexOutOfRangeException
                or NPOI.Util.RecordFormatException)
        {
            evidence.Add("npoi-hssf:workbook-rejected");
            return SignatureInspection.Unsupported(
                XlsSignatureKind,
                "ole-compound-not-hssf-workbook",
                evidence);
        }
    }

    private static bool StartsWith(ReadOnlySpan<byte> bytes, ReadOnlySpan<byte> prefix) =>
        bytes.Length >= prefix.Length && bytes[..prefix.Length].SequenceEqual(prefix);

    private static RuntimeEvidenceEnvelope CreateEnvelope(
        string status,
        string? failureStage,
        FileContentIdentity? source,
        RuntimeFileEvidence file,
        JsonElement payload,
        IReadOnlyList<ContractFinding> errors)
    {
        var retainedPayload = payload.Clone();
        return new RuntimeEvidenceEnvelope(
            RuntimeContractVersions.EvidenceEnvelope,
            "runtime-evidence",
            "identify",
            status,
            failureStage,
            PackageIdentity(),
            RuntimeIdentity(),
            EvidenceSchema,
            source,
            file,
            EvidenceEnvelope.IdentifyCanonicalJson(retainedPayload, PayloadSchema),
            retainedPayload,
            [],
            [],
            errors);
    }

    private static PackageIdentity PackageIdentity() => new(PackageName, PackageVersion);

    private static RuntimeIdentity RuntimeIdentity() => new("office", RuntimeName, PackageVersion);

    private static string ResolvePackageVersion()
    {
        var informationalVersion = typeof(XlsxRuntimeIdentity).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (string.IsNullOrWhiteSpace(informationalVersion))
        {
            throw new InvalidOperationException("XLSX package version metadata is unavailable.");
        }
        return informationalVersion.Split('+', 2)[0];
    }

    private sealed record SignatureInspection(
        bool Matched,
        string? FileKind,
        string? MediaType,
        string SignatureKind,
        string Reason,
        IReadOnlyList<string> Evidence)
    {
        public static SignatureInspection Supported(
            string fileKind,
            string mediaType,
            string signatureKind,
            IReadOnlyList<string> evidence) =>
            new(true, fileKind, mediaType, signatureKind, "matched", evidence);

        public static SignatureInspection Unsupported(
            string signatureKind,
            string reason,
            IReadOnlyList<string> evidence) =>
            new(false, null, null, signatureKind, reason, evidence);
    }
}
