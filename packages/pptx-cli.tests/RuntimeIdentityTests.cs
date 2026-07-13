using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Dockit.Pptx.Cli;
using Tiwater.RuntimeContracts;
using Xunit;

namespace Dockit.Pptx.Tests;

public sealed class RuntimeIdentityTests
{
    private const string TransitionalOfficeDocument =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    private const string StrictOfficeDocument =
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
    private const string PresentationContentType =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";

    [Fact]
    public void Capabilities_describe_non_mutating_pptx_identity_commands()
    {
        var descriptor = PptxRuntimeIdentity.Capabilities();

        Assert.Equal("runtime-capabilities", descriptor.DescriptorType);
        Assert.Equal(new PackageIdentity("tiwater.pptx.cli", "0.2.0"), descriptor.Package);
        Assert.Equal(new RuntimeIdentity("office", "tiwater-pptx", "0.2.0"), descriptor.Runtime);
        Assert.Equal("capabilities", descriptor.DescriptorCommand.Command);
        Assert.Equal(["--json"], descriptor.DescriptorCommand.Arguments);
        Assert.False(descriptor.DescriptorCommand.Mutates);
        Assert.Equal("identify", descriptor.IdentifyProbe.Command);
        Assert.Equal(["<input>", "--json"], descriptor.IdentifyProbe.Arguments);
        Assert.Equal(["supported", "unsupported", "failed"], descriptor.IdentifyProbe.Outcomes);
        Assert.False(descriptor.IdentifyProbe.Mutates);
        Assert.Contains(descriptor.SupportedKinds, kind => kind.FileKind == "pptx");
        Assert.Contains(descriptor.Commands, command => command.Name == "capabilities" && !command.Mutates);
        Assert.Contains(descriptor.Commands, command => command.Name == "identify" && !command.Mutates);
    }

    [Fact]
    public void Renamed_valid_presentation_is_supported_from_exact_unchanged_bytes()
    {
        var path = TemporaryPath(".payload");
        try
        {
            WritePackage(path, ValidEntries());
            var before = File.ReadAllBytes(path);

            var evidence = PptxRuntimeIdentity.Identify(path);

            Assert.True(
                evidence.Status == "supported",
                JsonSerializer.Serialize(evidence, RuntimeJson.Options));
            Assert.Null(evidence.FailureStage);
            Assert.Equal(Path.GetFullPath(path), evidence.Source!.Path);
            Assert.Equal(before.Length, evidence.Source.SizeBytes);
            var sha256 = Convert.ToHexStringLower(SHA256.HashData(before));
            Assert.Equal(sha256, evidence.Source.Sha256);
            Assert.Equal($"sha256:{sha256}", evidence.Source.ContentId);
            Assert.Equal(new PackageIdentity("tiwater.pptx.cli", "0.2.0"), evidence.Package);
            Assert.Equal(new RuntimeIdentity("office", "tiwater-pptx", "0.2.0"), evidence.Runtime);
            Assert.Equal(
                new SchemaIdentity("https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json", "1.0.0"),
                evidence.EvidenceSchema);
            Assert.Equal("pptx", evidence.File.FileKind);
            Assert.Equal(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                evidence.File.MediaType);
            Assert.Equal("matched", evidence.File.Signature.Status);
            Assert.NotEmpty(evidence.File.Signature.Evidence);
            Assert.Empty(evidence.Errors);
            Assert.Equal(
                EvidenceEnvelope.IdentifyCanonicalJson(
                    evidence.Payload,
                    new SchemaIdentity("tiwater.runtime.identify-payload", "1.0.0")),
                evidence.Artifact);
            Assert.Equal(before, File.ReadAllBytes(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fake_pptx_extension_is_unsupported_not_failed()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            File.WriteAllText(path, "not a zip package");

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal("not-a-zip-package", evidence.Payload.GetProperty("reason").GetString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")]
    [InlineData("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")]
    public void Other_office_packages_renamed_to_pptx_are_unsupported(string partName, string contentType)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(
                ContentTypesXml(partName, contentType),
                RelationshipsXml(partName),
                [(partName, Encoding.UTF8.GetBytes("<root />"))]));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string> InvalidContentTypesDocuments => new()
    {
        "<Types",
        "<!DOCTYPE Types [<!ENTITY x 'presentation'>]><Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types' />",
        $"<Bogus xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Bogus>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Wrapper><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Wrapper></Types>",
    };

    [Theory]
    [MemberData(nameof(InvalidContentTypesDocuments))]
    public void Malformed_dtd_wrong_root_or_nested_content_types_fail_closed(string contentTypes)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(contentTypes, RelationshipsXml("ppt/presentation.xml"), MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("[CONTENT_TYPES].XML")]
    [InlineData("%5BContent_Types%5D.xml")]
    public void Canonically_equivalent_content_types_items_are_ambiguous(string duplicateName)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            entries.Add((duplicateName, Encoding.UTF8.GetBytes(ContentTypesXml())));
            WritePackage(path, entries);

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal(
                "content-types-part-missing-or-ambiguous",
                evidence.Payload.GetProperty("reason").GetString());
            Assert.Contains("[Content_Types].xml:exact-count=1", evidence.File.Signature.Evidence);
            Assert.Contains("[Content_Types].xml:equivalent-count=2", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("[CONTENT_TYPES].XML")]
    [InlineData("%5BContent_Types%5D.xml")]
    public void Lone_content_types_alias_does_not_replace_the_exact_special_item(string alias)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            RenameEntry(entries, "[Content_Types].xml", alias);
            WritePackage(path, entries);

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal(
                "content-types-part-missing-or-ambiguous",
                evidence.Payload.GetProperty("reason").GetString());
            Assert.Contains("[Content_Types].xml:exact-count=0", evidence.File.Signature.Evidence);
            Assert.Contains("[Content_Types].xml:equivalent-count=1", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string> StructurallyInvalidContentTypesDocuments => new()
    {
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types' extra='x'><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>unexpected<Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml' Extra='x' /><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default ContentType='application/xml' /><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override ContentType='{PresentationContentType}' /></Types>",
        "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/ppt/presentation.xml' /></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' /></Default></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}'><Child /></Override></Types>",
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}'>unexpected</Override></Types>",
    };

    [Theory]
    [MemberData(nameof(StructurallyInvalidContentTypesDocuments))]
    public void Content_type_declarations_require_exact_attributes_and_empty_content(string contentTypes)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(contentTypes, RelationshipsXml("ppt/presentation.xml"), MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal("content-types-invalid", evidence.Payload.GetProperty("reason").GetString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("not a mime type")]
    [InlineData("application")]
    [InlineData("application/")]
    [InlineData("/xml")]
    [InlineData("application/xml; charset=utf-8")]
    public void Every_content_type_requires_parameter_free_rfc_media_type_tokens(string contentType)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var contentTypes = ContentTypesXml().Replace(
                "</Types>",
                $"<Default Extension='probe' ContentType='{contentType}' /></Types>",
                StringComparison.Ordinal);
            WritePackage(path, PackageEntries(
                contentTypes,
                RelationshipsXml("ppt/presentation.xml"),
                MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal("content-types-invalid", evidence.Payload.GetProperty("reason").GetString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string> InvalidRelationshipsDocuments => new()
    {
        "<Relationships",
        $"<!DOCTYPE Relationships [<!ENTITY x 'ppt/presentation.xml'>]><Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='&x;' /></Relationships>",
        $"<Bogus xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml' /></Bogus>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Wrapper><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml' /></Wrapper></Relationships>",
    };

    [Theory]
    [MemberData(nameof(InvalidRelationshipsDocuments))]
    public void Malformed_dtd_wrong_root_or_nested_root_relationships_fail_closed(string relationships)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string> StructurallyInvalidRelationshipsDocuments => new()
    {
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships' extra='x'>{Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml")}</Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>unexpected{Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml")}</Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml' Extra='x' /></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml' /></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='ppt/presentation.xml' /></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' /></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='1invalid' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml' /></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='invalid:id' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml' /></Relationships>",
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='relative' Target='ppt/presentation.xml' /></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml'><Child /></Relationship></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml'>unexpected</Relationship></Relationships>",
        $"<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='{TransitionalOfficeDocument}' Target='ppt/presentation.xml'>{Relationship("rId2", TransitionalOfficeDocument, "ppt/presentation.xml")}</Relationship></Relationships>",
    };

    [Theory]
    [MemberData(nameof(StructurallyInvalidRelationshipsDocuments))]
    public void Relationships_require_exact_attributes_ncname_ids_absolute_types_and_empty_content(string relationships)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal(
                "office-document-relationship-invalid",
                evidence.Payload.GetProperty("reason").GetString());
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Root_relationship_part_is_unique_by_case_insensitive_opc_equivalence()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            entries.Add(("_RELS/.RELS", Encoding.UTF8.GetBytes(RelationshipsXml("ppt/presentation.xml"))));
            WritePackage(path, entries);

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Case_variant_root_relationship_part_is_recognized()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            RenameEntry(entries, "_rels/.rels", "_RELS/.RELS");
            WritePackage(path, entries);

            var evidence = PptxRuntimeIdentity.Identify(path);
            Assert.True(
                evidence.Status == "supported",
                JsonSerializer.Serialize(evidence, RuntimeJson.Options));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Transitional_and_strict_office_document_relationships_are_ambiguous_together()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml"),
                Relationship("rId2", StrictOfficeDocument, "ppt/presentation.xml"));
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("Internal")]
    public void Default_and_explicit_internal_relationships_are_supported(string? targetMode)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var relationships = RelationshipsXml("ppt/presentation.xml", targetMode: targetMode);
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);
            Assert.True(
                evidence.Status == "supported",
                JsonSerializer.Serialize(evidence, RuntimeJson.Options));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Strict_office_document_relationship_is_supported()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(
                ContentTypesXml(),
                RelationshipsXml("ppt/presentation.xml", StrictOfficeDocument),
                MainPartEntries()));

            Assert.Equal("supported", PptxRuntimeIdentity.Identify(path).Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    public static TheoryData<string, string?> InvalidTargets => new()
    {
        { "ppt/presentation.xml", "" },
        { "ppt/presentation.xml", "External" },
        { "ppt/presentation.xml", "internal" },
        { "https://example.test/presentation.xml", null },
        { "//example.test/presentation.xml", null },
        { "../ppt/presentation.xml", null },
        { "%2e%2e/ppt/presentation.xml", null },
        { "ppt/presentation.xml?download=1", null },
        { "ppt/presentation.xml#slide", null },
        { "ppt\\presentation.xml", null },
        { "ppt%5cpresentation.xml", null },
    };

    [Theory]
    [MemberData(nameof(InvalidTargets))]
    public void External_invalid_or_unsafe_relationship_targets_are_unsupported(string target, string? targetMode)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(
                ContentTypesXml(),
                RelationshipsXml(target, targetMode: targetMode),
                MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("relative")]
    [InlineData("//example.test/type")]
    [InlineData("https://example.test/a b")]
    [InlineData("https://example.test/%ZZ")]
    [InlineData("https:\\example.test/type")]
    public void Every_relationship_type_requires_a_strict_absolute_uri(string type)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml"),
                Relationship("rId2", type, "ppt/presentation.xml", "Internal"));
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal(
                "office-document-relationship-invalid",
                evidence.Payload.GetProperty("reason").GetString());
            Assert.Contains(
                "office-document-relationship:invalid-relationship-type",
                evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("https://example.test/metadata.xml")]
    [InlineData("//example.test/metadata.xml")]
    [InlineData("../metadata.xml")]
    [InlineData("%2e%2e/metadata.xml")]
    [InlineData("metadata/%ZZ.xml")]
    [InlineData("metadata/info file.xml")]
    [InlineData("metadata/info.xml?query=1")]
    [InlineData("metadata/info.xml#fragment")]
    [InlineData("metadata\\info.xml")]
    public void Every_internal_root_relationship_target_requires_a_safe_package_part_uri(string target)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml"),
                Relationship("rId2", "urn:example:metadata", target, "Internal"));
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal(
                "office-document-relationship-invalid",
                evidence.Payload.GetProperty("reason").GetString());
            Assert.Contains(
                "office-document-relationship:invalid-relationship-target",
                evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("https://example.test/a b")]
    [InlineData("https://example.test/%ZZ")]
    [InlineData("https:\\example.test/resource")]
    public void Every_external_root_relationship_target_requires_a_well_formed_uri_reference(string target)
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml"),
                Relationship("rId2", "urn:example:metadata", target, "External"));
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            var evidence = PptxRuntimeIdentity.Identify(path);

            AssertUnsupported(evidence);
            Assert.Equal(
                "office-document-relationship-invalid",
                evidence.Payload.GetProperty("reason").GetString());
            Assert.Contains(
                "office-document-relationship:invalid-relationship-target",
                evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Valid_unrelated_content_type_and_internal_and_external_relationships_are_supported()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var contentTypes = ContentTypesXml().Replace(
                "</Types>",
                "<Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml' /></Types>",
                StringComparison.Ordinal);
            var relationships = RelationshipsDocument(
                Relationship("rId1", TransitionalOfficeDocument, "ppt/presentation.xml"),
                Relationship(
                    "rId2",
                    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                    "docProps/core.xml",
                    "Internal"),
                Relationship(
                    "rId3",
                    "https://example.test/relationships/reference",
                    "https://example.test/resource?query=1#fragment",
                    "External"));
            var parts = MainPartEntries();
            parts.Add((
                "docProps/core.xml",
                Encoding.UTF8.GetBytes(
                    "<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' />")));
            WritePackage(path, PackageEntries(contentTypes, relationships, parts));

            var evidence = PptxRuntimeIdentity.Identify(path);

            Assert.True(
                evidence.Status == "supported",
                JsonSerializer.Serialize(evidence, RuntimeJson.Options));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Relationship_requires_nonempty_unique_id()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship("", TransitionalOfficeDocument, "ppt/presentation.xml"));
            WritePackage(path, PackageEntries(ContentTypesXml(), relationships, MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Main_part_is_unique_by_case_insensitive_opc_equivalence()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            entries.Add(("PPT/PRESENTATION.XML", MainPartBytes()));
            WritePackage(path, entries);

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Every_package_part_is_unique_by_case_insensitive_opc_equivalence()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            entries[0] = (
                "[Content_Types].xml",
                Encoding.UTF8.GetBytes(ContentTypesXml().Replace(
                    "</Types>",
                    "<Default Extension='bin' ContentType='application/octet-stream' /></Types>",
                    StringComparison.Ordinal)));
            entries.Add(("media/probe.bin", [1]));
            entries.Add(("MEDIA/PROBE.BIN", [2]));
            WritePackage(path, entries);

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Directory_zip_entries_are_not_package_parts_or_case_collisions()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            entries.Add(("ppt/", []));
            entries.Add(("PPT/", []));
            WritePackage(path, entries);

            Assert.Equal("supported", PptxRuntimeIdentity.Identify(path).Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Unsafe_directory_zip_entries_are_unsupported()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = ValidEntries();
            entries.Add(("../", []));
            WritePackage(path, entries);

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Opc_equivalent_content_type_still_requires_openxml_readability()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var entries = PackageEntries(
                ContentTypesXml("PPT/PRESENTATION.XML", PresentationContentType),
                RelationshipsXml("ppt/presentation.xml"),
                [("ppt/presentation.xml", MainPartBytes())]);
            WritePackage(path, entries);

            var evidence = PptxRuntimeIdentity.Identify(path);
            AssertUnsupported(evidence);
            Assert.Contains("openxml:presentation-unreadable", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Duplicate_case_equivalent_content_type_overrides_are_unsupported()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            var contentTypes =
                $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
                $"<Override PartName='/ppt/presentation.xml' ContentType='{PresentationContentType}' />" +
                $"<Override PartName='/PPT/PRESENTATION.XML' ContentType='{PresentationContentType}' />" +
                "</Types>";
            WritePackage(path, PackageEntries(contentTypes, RelationshipsXml("ppt/presentation.xml"), MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Missing_part_or_wrong_content_type_is_unsupported()
    {
        var missingPath = TemporaryPath(".pptx");
        var wrongTypePath = TemporaryPath(".pptx");
        try
        {
            WritePackage(missingPath, PackageEntries(ContentTypesXml(), RelationshipsXml("ppt/presentation.xml"), []));
            WritePackage(wrongTypePath, PackageEntries(
                ContentTypesXml("ppt/presentation.xml", "application/xml"),
                RelationshipsXml("ppt/presentation.xml"),
                MainPartEntries()));

            AssertUnsupported(PptxRuntimeIdentity.Identify(missingPath));
            AssertUnsupported(PptxRuntimeIdentity.Identify(wrongTypePath));
        }
        finally
        {
            File.Delete(missingPath);
            File.Delete(wrongTypePath);
        }
    }

    [Fact]
    public void Unreadable_presentation_main_part_is_unsupported()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, PackageEntries(
                ContentTypesXml(),
                RelationshipsXml("ppt/presentation.xml"),
                [("ppt/presentation.xml", Encoding.UTF8.GetBytes("<not-presentation />"))]));

            AssertUnsupported(PptxRuntimeIdentity.Identify(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Missing_source_is_typed_source_read_failure_without_invented_identity()
    {
        var evidence = PptxRuntimeIdentity.Identify(TemporaryPath(".pptx"));

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
            EvidenceEnvelope.IdentifyCanonicalJson(
                evidence.Payload,
                new SchemaIdentity("tiwater.runtime.identify-payload", "1.0.0")),
            evidence.Artifact);
    }

    [Fact]
    public void Repeated_identify_is_byte_for_byte_deterministic()
    {
        var path = TemporaryPath(".pptx");
        try
        {
            WritePackage(path, ValidEntries());

            var first = JsonSerializer.Serialize(PptxRuntimeIdentity.Identify(path), RuntimeJson.Options);
            var second = JsonSerializer.Serialize(PptxRuntimeIdentity.Identify(path), RuntimeJson.Options);

            Assert.Equal(first, second);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static void AssertUnsupported(RuntimeEvidenceEnvelope evidence)
    {
        Assert.Equal("unsupported", evidence.Status);
        Assert.Null(evidence.FailureStage);
        Assert.NotNull(evidence.Source);
        Assert.Null(evidence.File.FileKind);
        Assert.Null(evidence.File.MediaType);
        Assert.Equal("mismatched", evidence.File.Signature.Status);
        Assert.Empty(evidence.Errors);
    }

    private static List<(string Name, byte[] Bytes)> ValidEntries() => PackageEntries(
        ContentTypesXml(),
        RelationshipsXml("ppt/presentation.xml"),
        MainPartEntries());

    private static List<(string Name, byte[] Bytes)> MainPartEntries() =>
        [("ppt/presentation.xml", MainPartBytes())];

    private static byte[] MainPartBytes() => Encoding.UTF8.GetBytes(
        "<p:presentation xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' " +
        "xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main'>" +
        "<p:sldMasterIdLst/><p:sldIdLst/><p:sldSz cx='9144000' cy='6858000'/><p:notesSz cx='6858000' cy='9144000'/>" +
        "</p:presentation>");

    private static string ContentTypesXml(
        string partName = "ppt/presentation.xml",
        string contentType = PresentationContentType) =>
        $"<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>" +
        "<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml' />" +
        $"<Override PartName='/{partName}' ContentType='{contentType}' />" +
        "</Types>";

    private static string RelationshipsXml(
        string target,
        string relationshipType = TransitionalOfficeDocument,
        string? targetMode = null) =>
        RelationshipsDocument(Relationship("rId1", relationshipType, target, targetMode));

    private static string RelationshipsDocument(params string[] relationships) =>
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
        string.Concat(relationships) +
        "</Relationships>";

    private static string Relationship(string id, string type, string target, string? targetMode = null) =>
        $"<Relationship Id='{id}' Type='{type}' Target='{target}'" +
        (targetMode is null ? "" : $" TargetMode='{targetMode}'") +
        " />";

    private static List<(string Name, byte[] Bytes)> PackageEntries(
        string contentTypes,
        string relationships,
        IEnumerable<(string Name, byte[] Bytes)> parts)
    {
        var entries = new List<(string Name, byte[] Bytes)>
        {
            ("[Content_Types].xml", Encoding.UTF8.GetBytes(contentTypes)),
            ("_rels/.rels", Encoding.UTF8.GetBytes(relationships)),
        };
        entries.AddRange(parts);
        return entries;
    }

    private static void RenameEntry(
        List<(string Name, byte[] Bytes)> entries,
        string oldName,
        string newName)
    {
        var index = entries.FindIndex(entry => entry.Name == oldName);
        entries[index] = (newName, entries[index].Bytes);
    }

    private static void WritePackage(string path, IEnumerable<(string Name, byte[] Bytes)> entries)
    {
        using var stream = File.Create(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        foreach (var item in entries)
        {
            var entry = archive.CreateEntry(item.Name);
            using var entryStream = entry.Open();
            entryStream.Write(item.Bytes);
        }
    }

    private static string TemporaryPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"pptx-runtime-{Guid.NewGuid():N}{extension}");
}
