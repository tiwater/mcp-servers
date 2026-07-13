using System.IO.Compression;
using System.IO.Packaging;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using Tiwater.RuntimeContracts;

namespace Dockit.Docx.Cli;

public static class DocxRuntimeIdentity
{
    private const string PackageName = "tiwater.docx.cli";
    private const string RuntimeName = "tiwater-docx";
    private const string EvidenceSchemaId = "https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json";
    private const string CapabilitiesSchemaId = "https://tiwater.dev/contracts/runtime/runtime-capabilities.schema.json";
    private const string PayloadSchemaId = "tiwater.runtime.identify-payload";
    private const string DocxMediaType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const string WordMainPartContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    private const string OfficeDocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string StrictOfficeDocumentRelationshipType = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
    private const string SignatureKind = "ooxml-word-main-part";

    private static readonly string PackageVersion = ResolvePackageVersion();
    private static readonly SchemaIdentity EvidenceSchema = new(EvidenceSchemaId, RuntimeContractVersions.EvidenceEnvelope);
    private static readonly SchemaIdentity PayloadSchema = new(PayloadSchemaId, RuntimeContractVersions.EvidenceEnvelope);

    public static RuntimeCapabilityDescriptor Capabilities() => new(
        RuntimeContractVersions.Capabilities,
        "runtime-capabilities",
        PackageIdentity(),
        RuntimeIdentity(),
        EvidenceSchema,
        new DiscoveryCommand("capabilities", ["--json"], false),
        new IdentifyProbe("identify", ["<input>", "--json"], false, ["supported", "unsupported", "failed"]),
        [new SupportedKind("docx", [DocxMediaType], [SignatureKind])],
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
                errors: [new ContractFinding("signature-inspection-failed", "The DOCX signature could not be inspected.")]);
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
                "docx",
                DocxMediaType,
                new SignatureEvidence("matched", SignatureKind, inspection.Evidence)),
            payload: JsonSerializer.SerializeToElement(new { recognized = true }, RuntimeJson.Options),
            errors: []);
    }

    private static SignatureInspection InspectSignature(ReadOnlyMemory<byte> sourceBytes)
    {
        try
        {
            using var stream = new MemoryStream(sourceBytes.ToArray(), writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            var evidence = new List<string> { "zip:central-directory-readable" };
            var equivalentContentTypesEntries = new List<ZipArchiveEntry>();
            var exactContentTypesEntries = new List<ZipArchiveEntry>();
            var packageEntries = new List<PackageEntry>();
            foreach (var entry in archive.Entries)
            {
                if (entry.Name.Length == 0 && entry.FullName.EndsWith("/", StringComparison.Ordinal))
                {
                    if (!TryParsePartUri(entry.FullName[..^1], false, false, out _))
                    {
                        evidence.Add("package-directory:invalid-uri");
                        return new SignatureInspection(false, "package-directory-uri-invalid", evidence);
                    }
                    continue;
                }

                if (IsEquivalentContentTypesItem(entry.FullName))
                {
                    equivalentContentTypesEntries.Add(entry);
                    if (string.Equals(entry.FullName, "[Content_Types].xml", StringComparison.Ordinal))
                    {
                        exactContentTypesEntries.Add(entry);
                    }
                    continue;
                }

                if (!TryParsePartUri(entry.FullName, false, false, out var part))
                {
                    evidence.Add("package-part:invalid-uri");
                    return new SignatureInspection(false, "package-part-uri-invalid", evidence);
                }
                packageEntries.Add(new PackageEntry(entry, part));
            }

            var packagePartGroups = packageEntries
                .GroupBy(item => item.Part.Key, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);

            if (equivalentContentTypesEntries.Count != 1 || exactContentTypesEntries.Count != 1)
            {
                evidence.Add($"[Content_Types].xml:exact-count={exactContentTypesEntries.Count}");
                evidence.Add($"[Content_Types].xml:equivalent-count={equivalentContentTypesEntries.Count}");
                return new SignatureInspection(false, "content-types-part-missing-or-ambiguous", evidence);
            }

            XDocument contentTypes;
            try
            {
                using var contentTypesStream = exactContentTypesEntries[0].Open();
                using var reader = XmlReader.Create(contentTypesStream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                });
                contentTypes = XDocument.Load(reader, LoadOptions.None);
            }
            catch (XmlException)
            {
                evidence.Add("[Content_Types].xml:invalid-xml");
                return new SignatureInspection(false, "content-types-invalid", evidence);
            }

            XNamespace contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
            if (contentTypes.Root?.Name != contentTypesNamespace + "Types")
            {
                evidence.Add("[Content_Types].xml:invalid-root");
                return new SignatureInspection(false, "content-types-invalid", evidence);
            }
            if (!TryParseContentTypeManifest(contentTypes, contentTypesNamespace, out var contentTypeManifest))
            {
                evidence.Add("[Content_Types].xml:invalid-declaration");
                return new SignatureInspection(false, "content-types-invalid", evidence);
            }

            XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            if (!TryParsePartUri("_rels/.rels", false, false, out var rootRelationshipsPart))
            {
                throw new InvalidOperationException("The canonical root relationships part URI is invalid.");
            }
            var relationshipsKey = rootRelationshipsPart.Key;
            packagePartGroups.TryGetValue(relationshipsKey, out var relationshipsEntries);
            relationshipsEntries ??= [];
            if (relationshipsEntries.Length != 1)
            {
                evidence.Add("office-document-relationship:count=0");
                evidence.Add($"_rels/.rels:count={relationshipsEntries.Length}");
                return new SignatureInspection(false, "root-relationships-part-missing-or-ambiguous", evidence);
            }

            XDocument relationships;
            try
            {
                using var relationshipsStream = relationshipsEntries[0].Entry.Open();
                using var reader = XmlReader.Create(relationshipsStream, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                });
                relationships = XDocument.Load(reader, LoadOptions.None);
            }
            catch (XmlException)
            {
                evidence.Add("_rels/.rels:invalid-xml");
                return new SignatureInspection(false, "root-relationships-invalid", evidence);
            }

            if (relationships.Root?.Name != relationshipsNamespace + "Relationships")
            {
                evidence.Add("_rels/.rels:invalid-root");
                return new SignatureInspection(false, "root-relationships-invalid", evidence);
            }
            if (relationships.Root.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
                || relationships.Root.Nodes().Any(node =>
                    node is not XElement
                    && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
            {
                evidence.Add("_rels/.rels:invalid-declaration");
                return new SignatureInspection(false, "root-relationships-invalid", evidence);
            }

            var relationshipElements = relationships.Root
                .Elements()
                .ToArray();
            if (relationshipElements.Any(element => element.Name != relationshipsNamespace + "Relationship")
                || relationships
                    .Descendants(relationshipsNamespace + "Relationship")
                    .Any(element => element.Parent != relationships.Root))
            {
                evidence.Add("_rels/.rels:invalid-declaration");
                return new SignatureInspection(false, "root-relationships-invalid", evidence);
            }

            var relationshipIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var relationship in relationshipElements)
            {
                var id = (string?)relationship.Attribute("Id");
                var type = (string?)relationship.Attribute("Type");
                var targetValue = (string?)relationship.Attribute("Target");
                var allowedAttributes = new HashSet<XName>
                {
                    "Id",
                    "Type",
                    "Target",
                    "TargetMode",
                };
                if (string.IsNullOrWhiteSpace(id)
                    || string.IsNullOrWhiteSpace(type)
                    || string.IsNullOrWhiteSpace(targetValue)
                    || !IsValidRelationshipId(id)
                    || !IsValidAbsoluteUri(type)
                    || relationship.Attributes().Any(attribute =>
                        !attribute.IsNamespaceDeclaration && !allowedAttributes.Contains(attribute.Name))
                    || relationship.Nodes().Any()
                    || !relationshipIds.Add(id))
                {
                    evidence.Add("_rels/.rels:invalid-declaration");
                    return new SignatureInspection(false, "root-relationships-invalid", evidence);
                }

                var targetModeAttribute = relationship.Attribute("TargetMode");
                if (targetModeAttribute is not null
                    && targetModeAttribute.Value is not "Internal" and not "External")
                {
                    var isOfficeDocument = IsOfficeDocumentRelationship(type);
                    evidence.Add(isOfficeDocument
                        ? "office-document-relationship:invalid-target-mode"
                        : "_rels/.rels:invalid-declaration");
                    return new SignatureInspection(false, "root-relationships-invalid", evidence);
                }

                var isExternal = string.Equals(
                    targetModeAttribute?.Value,
                    "External",
                    StringComparison.Ordinal);
                var targetIsValid = isExternal
                    ? IsValidUriReference(targetValue)
                    : TryParsePartUri(targetValue, false, true, out _);
                if (!targetIsValid)
                {
                    evidence.Add(IsOfficeDocumentRelationship(type)
                        ? "office-document-relationship:invalid-target"
                        : "_rels/.rels:invalid-declaration");
                    return new SignatureInspection(false, "root-relationships-invalid", evidence);
                }
            }

            var officeDocumentRelationships = relationshipElements
                .Where(element => string.Equals(
                    (string?)element.Attribute("Type"),
                    OfficeDocumentRelationshipType,
                    StringComparison.Ordinal)
                    || string.Equals(
                        (string?)element.Attribute("Type"),
                        StrictOfficeDocumentRelationshipType,
                        StringComparison.Ordinal))
                .ToArray();
            if (officeDocumentRelationships.Length != 1)
            {
                evidence.Add($"office-document-relationship:count={officeDocumentRelationships.Length}");
                return new SignatureInspection(false, "office-document-relationship-missing-or-ambiguous", evidence);
            }

            var officeDocumentRelationship = officeDocumentRelationships[0];
            var officeTargetModeAttribute = officeDocumentRelationship.Attribute("TargetMode");
            if (officeTargetModeAttribute is not null
                && !string.Equals(officeTargetModeAttribute.Value, "Internal", StringComparison.Ordinal))
            {
                evidence.Add(string.Equals(officeTargetModeAttribute.Value, "External", StringComparison.Ordinal)
                    ? "office-document-relationship:external"
                    : "office-document-relationship:invalid-target-mode");
                return new SignatureInspection(false, "office-document-relationship-not-internal", evidence);
            }

            if (!TryParsePartUri(
                    (string?)officeDocumentRelationship.Attribute("Target"),
                    false,
                    true,
                    out var target))
            {
                evidence.Add("office-document-relationship:invalid-target");
                return new SignatureInspection(false, "office-document-relationship-target-invalid", evidence);
            }

            string? effectiveMainPartContentType = null;
            if (contentTypeManifest.Overrides.TryGetValue(target.Key, out var overrideContentType))
            {
                effectiveMainPartContentType = overrideContentType;
            }
            else if (TryGetExtension(target.CanonicalPath, out var extension)
                && contentTypeManifest.Defaults.TryGetValue(AsciiFold(extension), out var defaultContentType))
            {
                effectiveMainPartContentType = defaultContentType;
            }

            if (!string.Equals(effectiveMainPartContentType, WordMainPartContentType, StringComparison.Ordinal))
            {
                evidence.Add("[Content_Types].xml:word-main-part-mismatch");
                return new SignatureInspection(false, "word-main-part-content-type-mismatch", evidence);
            }

            var mainParts = packageEntries
                .Where(item => item.Part.Key == target.Key)
                .ToArray();
            if (mainParts.Length != 1)
            {
                evidence.Add($"{AsciiFold(target.Path)}:count={mainParts.Length}");
                return new SignatureInspection(false, "word-main-part-missing-or-ambiguous", evidence);
            }

            var mainPartPath = mainParts[0].Entry.FullName;
            var unrelatedCollision = packagePartGroups
                .FirstOrDefault(group => group.Value.Length > 1);
            if (unrelatedCollision.Value is not null)
            {
                evidence.Add(
                    $"package-part:case-equivalent-collision={AsciiFold(unrelatedCollision.Value[0].Entry.FullName)}");
                return new SignatureInspection(false, "package-part-uri-ambiguous", evidence);
            }

            evidence.Add($"office-document-relationship:target={mainPartPath}");
            evidence.Add($"[Content_Types].xml:/{mainPartPath}={effectiveMainPartContentType}");
            evidence.Add($"part:{mainPartPath}");
            return new SignatureInspection(true, "matched", evidence);
        }
        catch (InvalidDataException)
        {
            return new SignatureInspection(false, "not-a-zip-package", ["zip:invalid"]);
        }
    }

    private static bool TryParseContentTypeManifest(
        XDocument contentTypes,
        XNamespace contentTypesNamespace,
        out ContentTypeManifest manifest)
    {
        var defaults = new Dictionary<string, string>(StringComparer.Ordinal);
        var overrides = new Dictionary<string, string>(StringComparer.Ordinal);
        manifest = new ContentTypeManifest(defaults, overrides);
        var root = contentTypes.Root!;
        if (root.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration)
            || root.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))
            || contentTypes
                .Descendants()
                .Any(element =>
                    (element.Name == contentTypesNamespace + "Default"
                        || element.Name == contentTypesNamespace + "Override")
                    && element.Parent != root))
        {
            return false;
        }

        foreach (var declaration in root.Elements())
        {
            if (declaration.Elements().Any()
                || declaration.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            {
                return false;
            }

            if (declaration.Name == contentTypesNamespace + "Default")
            {
                var extension = (string?)declaration.Attribute("Extension");
                var contentType = (string?)declaration.Attribute("ContentType");
                if (!HasOnlyAttributes(declaration, "Extension", "ContentType")
                    || !IsValidExtension(extension)
                    || !IsValidMediaType(contentType)
                    || !defaults.TryAdd(AsciiFold(extension!), contentType!))
                {
                    return false;
                }
            }
            else if (declaration.Name == contentTypesNamespace + "Override")
            {
                var partName = (string?)declaration.Attribute("PartName");
                var contentType = (string?)declaration.Attribute("ContentType");
                if (!HasOnlyAttributes(declaration, "PartName", "ContentType")
                    || !TryParsePartUri(partName, true, true, out var part)
                    || !IsValidMediaType(contentType)
                    || !overrides.TryAdd(part.Key, contentType!))
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        return true;
    }

    private static bool HasOnlyAttributes(XElement element, params XName[] expectedNames)
    {
        var attributes = element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration)
            .ToArray();
        return attributes.Length == expectedNames.Length
            && expectedNames.All(name => attributes.Count(attribute => attribute.Name == name) == 1);
    }

    private static bool IsValidExtension(string? value) =>
        !string.IsNullOrWhiteSpace(value)
        && string.Equals(value, value.Trim(), StringComparison.Ordinal)
        && value.All(character =>
            !char.IsControl(character)
            && !char.IsWhiteSpace(character)
            && character is not '.' and not '/' and not '\\' and not '?' and not '#');

    private static bool TryParsePartUri(
        string? value,
        bool requireLeadingSlash,
        bool allowLeadingSlash,
        out PackagePartUri part)
    {
        part = null!;
        if (string.IsNullOrWhiteSpace(value)
            || !string.Equals(value, value.Trim(), StringComparison.Ordinal)
            || value.Contains('\\', StringComparison.Ordinal)
            || value.Contains('?', StringComparison.Ordinal)
            || value.Contains('#', StringComparison.Ordinal))
        {
            return false;
        }

        var hasLeadingSlash = value.StartsWith("/", StringComparison.Ordinal);
        if (requireLeadingSlash != hasLeadingSlash && (requireLeadingSlash || !allowLeadingSlash))
        {
            return false;
        }
        if (value.StartsWith("//", StringComparison.Ordinal)) return false;

        var packagePath = hasLeadingSlash ? value[1..] : value;
        if (packagePath.Any(char.IsWhiteSpace)
            || Uri.TryCreate(packagePath, UriKind.Absolute, out _))
        {
            return false;
        }

        Uri normalizedPartUri;
        try
        {
            var candidate = PackUriHelper.CreatePartUri(new Uri($"/{packagePath}", UriKind.Relative));
            normalizedPartUri = PackUriHelper.GetNormalizedPartUri(candidate);
        }
        catch (Exception exception) when (exception is ArgumentException or UriFormatException)
        {
            return false;
        }

        if (!normalizedPartUri.OriginalString.StartsWith("/", StringComparison.Ordinal)) return false;
        var rawSegments = packagePath.Split('/');
        if (rawSegments.Length == 0 || rawSegments.Any(string.IsNullOrEmpty)) return false;

        foreach (var rawSegment in rawSegments)
        {
            if (!HasValidPercentEncoding(rawSegment)) return false;
            var safetyDecoded = DecodeAsciiPercentEscapesForSafety(rawSegment);
            if (safetyDecoded.Length == 0
                || safetyDecoded is "." or ".."
                || safetyDecoded.EndsWith(".", StringComparison.Ordinal)
                || safetyDecoded.Any(character =>
                    char.IsControl(character)
                    || character is '/' or '\\' or '\0'))
            {
                return false;
            }
        }

        var canonicalPath = normalizedPartUri.OriginalString.TrimStart('/');
        part = new PackagePartUri(packagePath, canonicalPath, normalizedPartUri.OriginalString);
        return true;
    }

    private static string DecodeAsciiPercentEscapesForSafety(string value)
    {
        var decoded = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] != '%')
            {
                decoded.Append(value[index]);
                continue;
            }

            var byteValue = Convert.ToByte(value.Substring(index + 1, 2), 16);
            decoded.Append(byteValue <= 0x7f ? (char)byteValue : '\ufffd');
            index += 2;
        }
        return decoded.ToString();
    }

    private static bool HasValidPercentEncoding(string value)
    {
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] != '%') continue;
            if (index + 2 >= value.Length
                || !Uri.IsHexDigit(value[index + 1])
                || !Uri.IsHexDigit(value[index + 2]))
            {
                return false;
            }
            index += 2;
        }
        return true;
    }

    private static bool IsEquivalentContentTypesItem(string itemName)
    {
        if (string.IsNullOrEmpty(itemName)
            || itemName.Contains('/', StringComparison.Ordinal)
            || itemName.Contains('\\', StringComparison.Ordinal)
            || itemName.Any(char.IsControl)
            || !HasValidPercentEncoding(itemName))
        {
            return false;
        }

        try
        {
            return string.Equals(
                Uri.UnescapeDataString(itemName),
                "[Content_Types].xml",
                StringComparison.OrdinalIgnoreCase);
        }
        catch (UriFormatException)
        {
            return false;
        }
    }

    private static bool IsValidAbsoluteUri(string value)
    {
        if (!HasValidUriLexicalForm(value)
            || value.StartsWith("//", StringComparison.Ordinal)
            || !Uri.TryCreate(value, UriKind.Absolute, out var uri))
        {
            return false;
        }

        return uri.IsAbsoluteUri
            && !string.IsNullOrWhiteSpace(uri.Scheme)
            && uri.IsWellFormedOriginalString();
    }

    private static bool IsValidUriReference(string value)
    {
        if (!HasValidUriLexicalForm(value)
            || !Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out var uri))
        {
            return false;
        }

        return uri.IsWellFormedOriginalString();
    }

    private static bool HasValidUriLexicalForm(string value) =>
        !string.IsNullOrWhiteSpace(value)
        && string.Equals(value, value.Trim(), StringComparison.Ordinal)
        && !value.Contains('\\', StringComparison.Ordinal)
        && !value.Any(character => char.IsControl(character) || char.IsWhiteSpace(character))
        && HasValidPercentEncoding(value);

    private static bool IsValidMediaType(string? value)
    {
        if (string.IsNullOrEmpty(value)
            || !string.Equals(value, value.Trim(), StringComparison.Ordinal))
        {
            return false;
        }

        var slashIndex = value.IndexOf('/');
        return slashIndex > 0
            && slashIndex == value.LastIndexOf('/')
            && slashIndex < value.Length - 1
            && IsMediaTypeToken(value[..slashIndex])
            && IsMediaTypeToken(value[(slashIndex + 1)..]);
    }

    private static bool IsMediaTypeToken(string value) => value.All(character =>
        char.IsAsciiLetterOrDigit(character)
        || character is '!' or '#' or '$' or '%' or '&' or '\'' or '*'
            or '+' or '-' or '.' or '^' or '_' or '`' or '|' or '~');

    private static bool TryGetExtension(string path, out string extension)
    {
        var lastSegment = path[(path.LastIndexOf('/') + 1)..];
        var dotIndex = lastSegment.LastIndexOf('.');
        extension = dotIndex >= 0 && dotIndex < lastSegment.Length - 1
            ? lastSegment[(dotIndex + 1)..]
            : string.Empty;
        return extension.Length > 0;
    }

    private static string AsciiFold(string value) => string.Create(
        value.Length,
        value,
        static (characters, source) =>
        {
            for (var index = 0; index < source.Length; index++)
            {
                var character = source[index];
                characters[index] = character is >= 'A' and <= 'Z'
                    ? (char)(character + ('a' - 'A'))
                    : character;
            }
        });

    private static bool IsOfficeDocumentRelationship(string type) =>
        string.Equals(type, OfficeDocumentRelationshipType, StringComparison.Ordinal)
        || string.Equals(type, StrictOfficeDocumentRelationshipType, StringComparison.Ordinal);

    private static bool IsValidRelationshipId(string value)
    {
        try
        {
            XmlConvert.VerifyNCName(value);
            return true;
        }
        catch (XmlException)
        {
            return false;
        }
    }

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
        var informationalVersion = typeof(DocxRuntimeIdentity).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        if (string.IsNullOrWhiteSpace(informationalVersion))
        {
            throw new InvalidOperationException("DOCX package version metadata is unavailable.");
        }
        return informationalVersion.Split('+', 2)[0];
    }

    private sealed record SignatureInspection(
        bool Matched,
        string Reason,
        IReadOnlyList<string> Evidence);

    private sealed record PackagePartUri(string Path, string CanonicalPath, string Key);

    private sealed record PackageEntry(ZipArchiveEntry Entry, PackagePartUri Part);

    private sealed record ContentTypeManifest(
        IReadOnlyDictionary<string, string> Defaults,
        IReadOnlyDictionary<string, string> Overrides);
}
