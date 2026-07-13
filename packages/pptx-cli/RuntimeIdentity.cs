using System.IO.Compression;
using System.IO.Packaging;
using System.Reflection;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Tiwater.RuntimeContracts;

namespace Dockit.Pptx.Cli;

public static class PptxRuntimeIdentity
{
    private const string PackageName = "tiwater.pptx.cli";
    private const string RuntimeName = "tiwater-pptx";
    private const string EvidenceSchemaId = "https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json";
    private const string CapabilitiesSchemaId = "https://tiwater.dev/contracts/runtime/runtime-capabilities.schema.json";
    private const string PayloadSchemaId = "tiwater.runtime.identify-payload";
    private const string PptxMediaType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    private const string PresentationMainPartContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    private const string OfficeDocumentRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string StrictOfficeDocumentRelationshipType =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string SignatureKind = "ooxml-presentation-main-part";
    private const long MaxXmlCharacters = 8 * 1024 * 1024;

    private static readonly string PackageVersion = ResolvePackageVersion();
    private static readonly SchemaIdentity EvidenceSchema = new(EvidenceSchemaId, RuntimeContractVersions.EvidenceEnvelope);
    private static readonly SchemaIdentity PayloadSchema = new(PayloadSchemaId, RuntimeContractVersions.EvidenceEnvelope);
    private static readonly Uri RootRelationshipsPartUri = CreatePartUri("/_rels/.rels");

    public static RuntimeCapabilityDescriptor Capabilities() => new(
        RuntimeContractVersions.Capabilities,
        "runtime-capabilities",
        PackageIdentity(),
        RuntimeIdentity(),
        EvidenceSchema,
        new DiscoveryCommand("capabilities", ["--json"], false),
        new IdentifyProbe("identify", ["<input>", "--json"], false, ["supported", "unsupported", "failed"]),
        [new SupportedKind("pptx", [PptxMediaType], [SignatureKind])],
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
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return CreateEnvelope(
                status: "failed",
                failureStage: "source-read",
                source: null,
                file: new RuntimeFileEvidence(null, null, new SignatureEvidence("not-checked", SignatureKind, [])),
                payload: JsonSerializer.SerializeToElement(new { failureClass = "source-read-error" }, RuntimeJson.Options),
                errors: [new ContractFinding("source-read-failed", "The source bytes could not be read.")]);
        }

        var source = FileIdentity.IdentifyBytes(resolvedPath, sourceBytes);
        SignatureInspection inspection;
        try
        {
            inspection = InspectSignature(sourceBytes);
        }
        catch (Exception)
        {
            return CreateEnvelope(
                status: "failed",
                failureStage: "signature-inspection",
                source,
                file: new RuntimeFileEvidence(null, null, new SignatureEvidence("unknown", SignatureKind, [])),
                payload: JsonSerializer.SerializeToElement(new { failureClass = "signature-inspection-error" }, RuntimeJson.Options),
                errors: [new ContractFinding("signature-inspection-failed", "The PPTX signature could not be inspected.")]);
        }

        if (!inspection.Matched)
        {
            return CreateEnvelope(
                status: "unsupported",
                failureStage: null,
                source,
                file: new RuntimeFileEvidence(null, null, new SignatureEvidence("mismatched", SignatureKind, inspection.Evidence)),
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
                "pptx",
                PptxMediaType,
                new SignatureEvidence("matched", SignatureKind, inspection.Evidence)),
            payload: JsonSerializer.SerializeToElement(new { recognized = true }, RuntimeJson.Options),
            errors: []);
    }

    private static SignatureInspection InspectSignature(ReadOnlyMemory<byte> sourceBytes)
    {
        try
        {
            using var stream = new MemoryStream(sourceBytes.ToArray(), writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            var evidence = new List<string> { "zip:central-directory-readable" };

            var contentTypesEntries = archive.Entries
                .Where(entry => entry.FullName == "[Content_Types].xml")
                .ToArray();
            if (contentTypesEntries.Length != 1)
            {
                evidence.Add($"[Content_Types].xml:count={contentTypesEntries.Length}");
                return Mismatch("content-types-part-missing-or-ambiguous", evidence);
            }

            if (!TryIndexPackageParts(archive, out var packageParts, out var indexReason))
            {
                evidence.Add($"package-parts:{indexReason}");
                return Mismatch("package-parts-invalid-or-ambiguous", evidence);
            }

            var rootRelationships = packageParts
                .Where(part => PartUrisEquivalent(part.Uri, RootRelationshipsPartUri))
                .ToArray();
            if (rootRelationships.Length != 1)
            {
                evidence.Add($"root-relationships-part:count={rootRelationships.Length}");
                return Mismatch("root-relationships-part-missing-or-ambiguous", evidence);
            }

            if (!TryLoadXml(contentTypesEntries[0], out var contentTypes))
            {
                evidence.Add("[Content_Types].xml:invalid-xml");
                return Mismatch("content-types-invalid", evidence);
            }
            if (!TryReadContentTypes(contentTypes!, out var contentTypeMap, out var contentTypesReason))
            {
                evidence.Add($"[Content_Types].xml:{contentTypesReason}");
                return Mismatch("content-types-invalid", evidence);
            }

            if (!TryLoadXml(rootRelationships[0].Entry, out var relationships))
            {
                evidence.Add("root-relationships:invalid-xml");
                return Mismatch("root-relationships-invalid", evidence);
            }
            if (!TryResolvePresentationRelationship(
                    relationships!,
                    out var targetPartUri,
                    out var relationshipReason))
            {
                evidence.Add($"office-document-relationship:{relationshipReason}");
                return Mismatch("office-document-relationship-invalid", evidence);
            }

            var targetParts = packageParts
                .Where(part => PartUrisEquivalent(part.Uri, targetPartUri!))
                .ToArray();
            if (targetParts.Length != 1)
            {
                evidence.Add($"presentation-main-part:count={targetParts.Length}");
                return Mismatch("presentation-main-part-missing-or-ambiguous", evidence);
            }

            if (!contentTypeMap!.TryResolve(targetPartUri!, out var effectiveContentType)
                || !string.Equals(
                    effectiveContentType,
                    PresentationMainPartContentType,
                    StringComparison.Ordinal))
            {
                evidence.Add("[Content_Types].xml:presentation-main-part-mismatch");
                return Mismatch("presentation-main-part-content-type-mismatch", evidence);
            }

            if (!IsReadablePresentation(sourceBytes, targetPartUri!))
            {
                evidence.Add("openxml:presentation-unreadable");
                return Mismatch("presentation-openxml-unreadable", evidence);
            }

            var targetName = targetPartUri!.ToString().TrimStart('/');
            evidence.Add($"office-document-relationship:target={targetName}");
            evidence.Add($"[Content_Types].xml:/{targetName}={effectiveContentType}");
            evidence.Add($"part:{targetParts[0].Entry.FullName}");
            evidence.Add("openxml:presentation-readable");
            return new SignatureInspection(true, "matched", evidence);
        }
        catch (InvalidDataException)
        {
            return Mismatch("not-a-zip-package", ["zip:invalid"]);
        }
    }

    private static bool TryIndexPackageParts(
        ZipArchive archive,
        out IReadOnlyList<PackagePartEntry> packageParts,
        out string reason)
    {
        var parts = new List<PackagePartEntry>();
        foreach (var entry in archive.Entries)
        {
            if (entry.FullName == "[Content_Types].xml") continue;
            if (IsDirectoryEntry(entry))
            {
                if (!TryCreatePartUri(entry.FullName[..^1], requireLeadingSlash: false, out _))
                {
                    packageParts = [];
                    reason = "invalid-directory-entry";
                    return false;
                }
                continue;
            }
            if (!TryCreatePartUri(entry.FullName, requireLeadingSlash: false, out var partUri))
            {
                packageParts = [];
                reason = "invalid-part-uri";
                return false;
            }
            if (parts.Any(existing => PartUrisEquivalent(existing.Uri, partUri!)))
            {
                packageParts = [];
                reason = "duplicate-equivalent-part-uri";
                return false;
            }
            parts.Add(new PackagePartEntry(partUri!, entry));
        }
        packageParts = parts;
        reason = "valid";
        return true;
    }

    private static bool IsDirectoryEntry(ZipArchiveEntry entry) =>
        entry.FullName.EndsWith("/", StringComparison.Ordinal) && entry.Name.Length == 0;

    private static bool TryLoadXml(ZipArchiveEntry entry, out XDocument? document)
    {
        try
        {
            using var entryStream = entry.Open();
            using var reader = XmlReader.Create(entryStream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = MaxXmlCharacters,
            });
            document = XDocument.Load(reader, LoadOptions.None);
            return true;
        }
        catch (Exception exception) when (exception is XmlException or InvalidDataException or IOException)
        {
            document = null;
            return false;
        }
    }

    private static bool TryReadContentTypes(
        XDocument document,
        out ContentTypeMap? contentTypeMap,
        out string reason)
    {
        XNamespace contentTypesNamespace = ContentTypesNamespace;
        if (document.Root?.Name != contentTypesNamespace + "Types")
        {
            contentTypeMap = null;
            reason = "invalid-root";
            return false;
        }

        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new List<(Uri PartUri, string ContentType)>();
        foreach (var element in document.Root.Elements())
        {
            if (element.Name == contentTypesNamespace + "Default")
            {
                var extension = (string?)element.Attribute("Extension");
                var contentType = (string?)element.Attribute("ContentType");
                if (string.IsNullOrWhiteSpace(extension)
                    || extension.Contains('.')
                    || extension.Contains('/')
                    || string.IsNullOrWhiteSpace(contentType)
                    || !defaults.TryAdd(extension, contentType))
                {
                    contentTypeMap = null;
                    reason = "invalid-or-duplicate-default";
                    return false;
                }
                continue;
            }

            if (element.Name == contentTypesNamespace + "Override")
            {
                var partName = (string?)element.Attribute("PartName");
                var contentType = (string?)element.Attribute("ContentType");
                if (!TryCreatePartUri(partName, requireLeadingSlash: true, out var partUri)
                    || string.IsNullOrWhiteSpace(contentType)
                    || overrides.Any(existing => PartUrisEquivalent(existing.PartUri, partUri!)))
                {
                    contentTypeMap = null;
                    reason = "invalid-or-duplicate-override";
                    return false;
                }
                overrides.Add((partUri!, contentType));
                continue;
            }

            contentTypeMap = null;
            reason = "invalid-direct-declaration";
            return false;
        }

        contentTypeMap = new ContentTypeMap(defaults, overrides);
        reason = "valid";
        return true;
    }

    private static bool TryResolvePresentationRelationship(
        XDocument document,
        out Uri? targetPartUri,
        out string reason)
    {
        XNamespace relationshipsNamespace = RelationshipsNamespace;
        if (document.Root?.Name != relationshipsNamespace + "Relationships")
        {
            targetPartUri = null;
            reason = "invalid-root";
            return false;
        }

        var relationships = document.Root.Elements().ToArray();
        if (relationships.Any(element => element.Name != relationshipsNamespace + "Relationship"))
        {
            targetPartUri = null;
            reason = "invalid-direct-declaration";
            return false;
        }

        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in relationships)
        {
            var id = (string?)relationship.Attribute("Id");
            var type = (string?)relationship.Attribute("Type");
            var target = (string?)relationship.Attribute("Target");
            var targetModeAttribute = relationship.Attribute("TargetMode");
            if (string.IsNullOrWhiteSpace(id)
                || string.IsNullOrWhiteSpace(type)
                || string.IsNullOrWhiteSpace(target)
                || !ids.Add(id))
            {
                targetPartUri = null;
                reason = "invalid-required-attributes";
                return false;
            }
            if (targetModeAttribute is not null
                && (string)targetModeAttribute is not ("Internal" or "External"))
            {
                targetPartUri = null;
                reason = "invalid-target-mode";
                return false;
            }
        }

        var candidates = relationships
            .Where(relationship =>
                string.Equals((string?)relationship.Attribute("Type"), OfficeDocumentRelationshipType, StringComparison.Ordinal)
                || string.Equals(
                    (string?)relationship.Attribute("Type"),
                    StrictOfficeDocumentRelationshipType,
                    StringComparison.Ordinal))
            .ToArray();
        if (candidates.Length != 1)
        {
            targetPartUri = null;
            reason = $"count={candidates.Length}";
            return false;
        }

        var candidate = candidates[0];
        var targetMode = candidate.Attribute("TargetMode");
        if (targetMode is not null && (string)targetMode != "Internal")
        {
            targetPartUri = null;
            reason = "not-internal";
            return false;
        }
        if (!TryCreatePartUri((string?)candidate.Attribute("Target"), requireLeadingSlash: false, out targetPartUri))
        {
            reason = "invalid-target";
            return false;
        }

        reason = "valid";
        return true;
    }

    private static bool TryCreatePartUri(string? value, bool requireLeadingSlash, out Uri? partUri)
    {
        partUri = null;
        if (string.IsNullOrWhiteSpace(value)
            || (requireLeadingSlash && !value.StartsWith("/", StringComparison.Ordinal))
            || value.StartsWith("//", StringComparison.Ordinal)
            || value.Contains('\\', StringComparison.Ordinal)
            || value.Contains('?', StringComparison.Ordinal)
            || value.Contains('#', StringComparison.Ordinal)
            || value.Any(character => char.IsControl(character) || character > 0x7f))
        {
            return false;
        }

        var absoluteProbe = value.StartsWith("/", StringComparison.Ordinal) ? value[1..] : value;
        if (Uri.TryCreate(absoluteProbe, UriKind.Absolute, out _)) return false;
        string decoded;
        try
        {
            decoded = Uri.UnescapeDataString(value);
        }
        catch (UriFormatException)
        {
            return false;
        }
        if (decoded.Contains('\\', StringComparison.Ordinal)
            || decoded.Contains('?', StringComparison.Ordinal)
            || decoded.Contains('#', StringComparison.Ordinal)
            || decoded.Any(character => char.IsControl(character) || character > 0x7f))
        {
            return false;
        }

        var packagePath = decoded.StartsWith("/", StringComparison.Ordinal) ? decoded : $"/{decoded}";
        var segments = packagePath[1..].Split('/');
        if (segments.Length == 0
            || segments.Any(segment => string.IsNullOrEmpty(segment) || segment is "." or ".."))
        {
            return false;
        }

        try
        {
            partUri = CreatePartUri(packagePath);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (UriFormatException)
        {
            return false;
        }
    }

    private static Uri CreatePartUri(string packagePath) =>
        PackUriHelper.CreatePartUri(new Uri(packagePath, UriKind.Relative));

    private static bool PartUrisEquivalent(Uri first, Uri second) =>
        PackUriHelper.ComparePartUri(first, second) == 0;

    private static bool IsReadablePresentation(ReadOnlyMemory<byte> sourceBytes, Uri expectedMainPartUri)
    {
        try
        {
            using var stream = new MemoryStream(sourceBytes.ToArray(), writable: false);
            using var presentation = PresentationDocument.Open(stream, false);
            var presentationPart = presentation.PresentationPart;
            if (presentationPart is null
                || !PartUrisEquivalent(presentationPart.Uri, expectedMainPartUri)
                || presentationPart.Presentation is null)
            {
                return false;
            }
            _ = presentationPart.Presentation.ChildElements.Count;
            return true;
        }
        catch (Exception exception) when (
            exception is InvalidDataException
                or IOException
                or XmlException
                or ArgumentException
                or InvalidOperationException
                or OpenXmlPackageException)
        {
            return false;
        }
    }

    private static SignatureInspection Mismatch(string reason, IReadOnlyList<string> evidence) =>
        new(false, reason, evidence);

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
        var informationalVersion = typeof(PptxRuntimeIdentity).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (string.IsNullOrWhiteSpace(informationalVersion))
        {
            throw new InvalidOperationException("PPTX package version metadata is unavailable.");
        }
        return informationalVersion.Split('+', 2)[0];
    }

    private sealed record SignatureInspection(bool Matched, string Reason, IReadOnlyList<string> Evidence);

    private sealed record PackagePartEntry(Uri Uri, ZipArchiveEntry Entry);

    private sealed record ContentTypeMap(
        IReadOnlyDictionary<string, string> Defaults,
        IReadOnlyList<(Uri PartUri, string ContentType)> Overrides)
    {
        public bool TryResolve(Uri partUri, out string? contentType)
        {
            var matchingOverrides = Overrides
                .Where(item => PartUrisEquivalent(item.PartUri, partUri))
                .ToArray();
            if (matchingOverrides.Length == 1)
            {
                contentType = matchingOverrides[0].ContentType;
                return true;
            }
            if (matchingOverrides.Length > 1)
            {
                contentType = null;
                return false;
            }

            var path = Uri.UnescapeDataString(partUri.ToString());
            var extension = Path.GetExtension(path).TrimStart('.');
            return Defaults.TryGetValue(extension, out contentType);
        }
    }
}
