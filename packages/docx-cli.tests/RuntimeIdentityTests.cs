using System.IO.Compression;
using System.IO.Packaging;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Docx.Cli;
using Tiwater.RuntimeContracts;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class RuntimeIdentityTests
{
    [Fact]
    public void Capabilities_describe_non_mutating_docx_identity_commands()
    {
        var descriptor = DocxRuntimeIdentity.Capabilities();

        Assert.Equal("runtime-capabilities", descriptor.DescriptorType);
        Assert.Equal(new PackageIdentity("tiwater.docx.cli", "0.4.0"), descriptor.Package);
        Assert.Equal(new RuntimeIdentity("office", "tiwater-docx", "0.4.0"), descriptor.Runtime);
        Assert.Equal("capabilities", descriptor.DescriptorCommand.Command);
        Assert.Equal(["--json"], descriptor.DescriptorCommand.Arguments);
        Assert.False(descriptor.DescriptorCommand.Mutates);
        Assert.Equal("identify", descriptor.IdentifyProbe.Command);
        Assert.Equal(["<input>", "--json"], descriptor.IdentifyProbe.Arguments);
        Assert.Equal(["supported", "unsupported", "failed"], descriptor.IdentifyProbe.Outcomes);
        Assert.False(descriptor.IdentifyProbe.Mutates);
        Assert.Contains(descriptor.SupportedKinds, kind => kind.FileKind == "docx");
        Assert.Contains(descriptor.Commands, command => command.Name == "capabilities" && !command.Mutates);
        Assert.Contains(descriptor.Commands, command => command.Name == "identify" && !command.Mutates);
    }

    [Fact]
    public void Renamed_valid_docx_is_supported_from_package_bytes()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateWordDocument(path);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.True(
                evidence.Status == "supported",
                JsonSerializer.Serialize(evidence, RuntimeJson.Options)
                + Environment.NewLine
                + ReadZipEntry(path, "[Content_Types].xml"));
            Assert.Null(evidence.FailureStage);
            Assert.NotNull(evidence.Source);
            Assert.Equal(Path.GetFullPath(path), evidence.Source.Path);
            var sourceBytes = File.ReadAllBytes(path);
            var sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(sourceBytes));
            Assert.Equal(sourceBytes.Length, evidence.Source.SizeBytes);
            Assert.Equal(sourceSha256, evidence.Source.Sha256);
            Assert.Equal($"sha256:{sourceSha256}", evidence.Source.ContentId);
            Assert.Equal(new PackageIdentity("tiwater.docx.cli", "0.4.0"), evidence.Package);
            Assert.Equal(new RuntimeIdentity("office", "tiwater-docx", "0.4.0"), evidence.Runtime);
            Assert.Equal(
                new SchemaIdentity("https://tiwater.dev/contracts/runtime/runtime-evidence-envelope.schema.json", "1.0.0"),
                evidence.EvidenceSchema);
            Assert.Equal("docx", evidence.File.FileKind);
            Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", evidence.File.MediaType);
            Assert.Equal("matched", evidence.File.Signature.Status);
            Assert.NotEmpty(evidence.File.Signature.Evidence);
            Assert.Empty(evidence.Errors);
            Assert.Equal(
                EvidenceEnvelope.IdentifyCanonicalJson(
                    evidence.Payload,
                    new SchemaIdentity("tiwater.runtime.identify-payload", "1.0.0")),
                evidence.Artifact);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Fake_docx_extension_is_unsupported_not_failed()
    {
        var path = TemporaryPath(".docx");
        try
        {
            File.WriteAllText(path, "not a zip package");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.NotNull(evidence.Source);
            Assert.Null(evidence.File.FileKind);
            Assert.Null(evidence.File.MediaType);
            Assert.NotEqual("matched", evidence.File.Signature.Status);
            Assert.Empty(evidence.Errors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Xlsx_renamed_to_docx_is_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateSpreadsheet(path);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.NotNull(evidence.Source);
            Assert.Null(evidence.File.FileKind);
            Assert.Equal("mismatched", evidence.File.Signature.Status);
            Assert.Empty(evidence.Errors);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Word_like_parts_without_office_document_relationship_are_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteZipEntry(
                    archive,
                    "[Content_Types].xml",
                    """
                    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                      <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
                    </Types>
                    """);
                WriteZipEntry(archive, "word/document.xml", "<document />");
            }

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.Contains("office-document-relationship", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<NotTypes xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" />")]
    [InlineData("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Wrapper><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\" /></Wrapper></Types>")]
    [InlineData("<Types")]
    [InlineData("<!DOCTYPE Types [<!ENTITY part \"/word/document.xml\">]><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"&part;\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\" /></Types>")]
    public void Invalid_content_types_xml_fails_closed(string contentTypesXml)
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreatePackage(
                path,
                contentTypesXml,
                RootOfficeDocumentRelationship("word/document.xml"),
                "word/document.xml");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item.StartsWith("[Content_Types].xml:invalid", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("[CONTENT_TYPES].XML")]
    [InlineData("%5BContent_Types%5D.xml")]
    public void Lone_content_types_alias_is_not_the_exact_special_item(string alias)
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreatePackageWithContentTypeItems(path, [alias]);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("[Content_Types].xml:exact-count=0", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("[CONTENT_TYPES].XML")]
    [InlineData("%5BContent_Types%5D.xml")]
    public void Exact_content_types_item_and_canonical_alias_are_ambiguous(string alias)
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreatePackageWithContentTypeItems(path, ["[Content_Types].xml", alias]);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("[Content_Types].xml:equivalent-count=2", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<Default Extension=\"xml\" />")]
    [InlineData("<Override PartName=\"/word/document.xml\" />")]
    [InlineData("<Override ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\" />")]
    [InlineData("<Default Extension=\"XML\" ContentType=\"application/xml\" /><Default Extension=\"xml\" ContentType=\"application/xml\" />")]
    [InlineData("<Override PartName=\"/WORD/document.xml\" ContentType=\"application/xml\" /><Override PartName=\"/word/DOCUMENT.xml\" ContentType=\"application/xml\" />")]
    public void Invalid_or_duplicate_content_type_declarations_fail_closed(string declarations)
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreatePackage(
                path,
                $"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">{declarations}</Types>",
                RootOfficeDocumentRelationship("word/document.xml"),
                "word/document.xml");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("[Content_Types].xml:invalid-declaration", evidence.File.Signature.Evidence);
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
    public void Content_types_require_parameter_free_rfc_media_type_tokens(string contentType)
    {
        var path = TemporaryPath(".docx");
        try
        {
            var contentTypes = $$"""
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="bin" ContentType="{{contentType}}" />
                  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
                </Types>
                """;
            CreatePackage(
                path,
                contentTypes,
                RootOfficeDocumentRelationship("word/document.xml"),
                "word/document.xml");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("[Content_Types].xml:invalid-declaration", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Ambiguous_office_document_relationships_are_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains(
                evidence.File.Signature.Evidence,
                item => item == "office-document-relationship:count=2");
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(" xml:base=\"https://example.test/\"")]
    [InlineData(" extra=\"value\"")]
    public void Root_relationships_reject_non_namespace_attributes(string rootAttribute)
    {
        var path = TemporaryPath(".docx");
        try
        {
            var relationships = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"{rootAttribute}>{Relationship("rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "word/document.xml")}</Relationships>";
            CreateWordLikePackage(path, "word/document.xml", relationships);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-declaration", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("unexpected")]
    [InlineData("<!-- comment -->")]
    public void Root_relationships_allow_only_direct_relationships_and_whitespace(string rootContent)
    {
        var path = TemporaryPath(".docx");
        try
        {
            var relationships = RelationshipsDocument(
                rootContent,
                Relationship(
                    "rId1",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "word/document.xml"));
            CreateWordLikePackage(path, "word/document.xml", relationships);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-declaration", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Case_equivalent_duplicate_root_relationship_parts_are_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteZipEntry(archive, "[Content_Types].xml", WordContentTypes("/word/document.xml"));
                WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "_RELS/.RELS", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "word/document.xml", "<document />");
            }

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:count=2", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void External_office_document_relationship_is_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="https://example.test/document.xml" TargetMode="External" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("office-document-relationship:external", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Traversing_office_document_relationship_target_is_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="../word/document.xml" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("office-document-relationship:invalid-target", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Relationship_targeted_renamed_word_main_part_is_supported()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreateWordLikePackage(
                path,
                "custom/main.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/main.xml" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("supported", evidence.Status);
            Assert.Contains("office-document-relationship:target=custom/main.xml", evidence.File.Signature.Evidence);
            Assert.Contains("part:custom/main.xml", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Part_uri_case_differences_across_relationship_override_and_zip_entry_are_supported()
    {
        var path = TemporaryPath(".payload");
        try
        {
            CreatePackage(
                path,
                WordContentTypes("/word/DOCUMENT.xml"),
                RootOfficeDocumentRelationship("WoRd/document.XML"),
                "WORD/Document.xml");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("supported", evidence.Status);
            Assert.Contains("part:WORD/Document.xml", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("custom/a.xml", "CUSTOM/%61.XML")]
    [InlineData("custom/a%3ab.xml", "CUSTOM/A%3AB.XML")]
    public void Pack_uri_equivalent_relationship_override_and_zip_entry_forms_are_supported(
        string declaredPartName,
        string zipEntryName)
    {
        Assert.Equal(
            0,
            PackUriHelper.ComparePartUri(
                PackPartUri(declaredPartName),
                PackPartUri(zipEntryName)));
        var path = TemporaryPath(".payload");
        try
        {
            CreatePackage(
                path,
                WordContentTypes($"/{declaredPartName}"),
                RootOfficeDocumentRelationship(declaredPartName),
                zipEntryName);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("supported", evidence.Status);
            Assert.Contains($"part:{zipEntryName}", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Reserved_percent_encoding_is_not_an_alias_for_the_literal_character()
    {
        const string declaredPartName = "custom/a:b.xml";
        const string zipEntryName = "custom/a%3Ab.xml";
        Assert.NotEqual(
            0,
            PackUriHelper.ComparePartUri(
                PackPartUri(declaredPartName),
                PackPartUri(zipEntryName)));
        var path = TemporaryPath(".payload");
        try
        {
            CreatePackage(
                path,
                WordContentTypes($"/{declaredPartName}"),
                RootOfficeDocumentRelationship(declaredPartName),
                zipEntryName);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.DoesNotContain($"part:{zipEntryName}", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Literal_and_reserved_percent_encoded_unrelated_parts_remain_distinct()
    {
        Assert.NotEqual(
            0,
            PackUriHelper.ComparePartUri(
                PackPartUri("custom/a:b.xml"),
                PackPartUri("custom/a%3Ab.xml")));
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteZipEntry(archive, "[Content_Types].xml", WordContentTypes("/word/document.xml"));
                WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "word/document.xml", "<document />");
                WriteZipEntry(archive, "custom/a:b.xml", "<metadata />");
                WriteZipEntry(archive, "custom/a%3Ab.xml", "<metadata />");
            }

            Assert.Equal("supported", DocxRuntimeIdentity.Identify(path).Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Case_equivalent_duplicate_main_parts_are_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreatePackage(
                path,
                WordContentTypes("/word/document.xml"),
                RootOfficeDocumentRelationship("word/document.xml"),
                "word/document.xml",
                "WORD/DOCUMENT.XML");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("word/document.xml:count=2", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Case_equivalent_collision_in_an_unrelated_package_part_is_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteZipEntry(archive, "[Content_Types].xml", WordContentTypes("/word/document.xml"));
                WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "word/document.xml", "<document />");
                WriteZipEntry(archive, "custom/foo.xml", "<metadata />");
                WriteZipEntry(archive, "CUSTOM/FOO.XML", "<metadata />");
            }

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("package-part:case-equivalent-collision=custom/foo.xml", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Invalid_unrelated_package_part_uri_is_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteZipEntry(archive, "[Content_Types].xml", WordContentTypes("/word/document.xml"));
                WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "word/document.xml", "<document />");
                WriteZipEntry(archive, "custom/../invalid.xml", "<metadata />");
            }

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("package-part:invalid-uri", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Zip_directory_entries_are_not_treated_as_package_parts()
    {
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                archive.CreateEntry("custom/");
                WriteZipEntry(archive, "[Content_Types].xml", WordContentTypes("/word/document.xml"));
                WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "word/document.xml", "<document />");
            }

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("supported", evidence.Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("custom/../")]
    [InlineData("custom/%2e%2e/")]
    [InlineData("custom/%ZZ/")]
    [InlineData("custom\\/")]
    [InlineData("custom//")]
    public void Unsafe_zip_directory_entries_are_unsupported(string directoryName)
    {
        var path = TemporaryPath(".docx");
        try
        {
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                archive.CreateEntry(directoryName);
                WriteZipEntry(archive, "[Content_Types].xml", WordContentTypes("/word/document.xml"));
                WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
                WriteZipEntry(archive, "word/document.xml", "<document />");
            }

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("package-directory:invalid-uri", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Office_document_relationship_target_must_exist_in_the_package()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/missing.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/missing.xml" />
                </Relationships>
                """,
                includeMainPart: false);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("word/missing.xml:count=0", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Explicit_internal_office_document_relationship_is_supported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" TargetMode="Internal" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("supported", evidence.Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("internal")]
    [InlineData("INTERNAL")]
    [InlineData("invalid")]
    public void Non_exact_target_mode_values_fail_closed(string targetMode)
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                $$"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" TargetMode="{{targetMode}}" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("office-document-relationship:invalid-target-mode", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("https://example.test/metadata.xml")]
    [InlineData("../metadata.xml")]
    [InlineData("%2e%2e/metadata.xml")]
    [InlineData("metadata/%ZZ.xml")]
    [InlineData("metadata/info.xml?query=1")]
    [InlineData("metadata/info.xml#fragment")]
    [InlineData("metadata\\info.xml")]
    public void Every_internal_root_relationship_target_must_be_a_safe_package_part_uri(string target)
    {
        var path = TemporaryPath(".docx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship(
                    "rId1",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "word/document.xml"),
                Relationship("rId2", "urn:example:metadata", target, "Internal"));
            CreateWordLikePackage(path, "word/document.xml", relationships);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-declaration", evidence.File.Signature.Evidence);
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
    public void Every_external_root_relationship_target_must_be_a_well_formed_uri_reference(string target)
    {
        var path = TemporaryPath(".docx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship(
                    "rId1",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "word/document.xml"),
                Relationship("rId2", "urn:example:metadata", target, "External"));
            CreateWordLikePackage(path, "word/document.xml", relationships);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-declaration", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Safe_internal_and_external_non_office_relationships_are_supported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship(
                    "rId1",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "word/document.xml"),
                Relationship("rId2", "urn:example:metadata", "word/document.xml", "Internal"),
                Relationship(
                    "rId3",
                    "https://example.test/relationships/reference",
                    "https://example.test/resource?query=1#fragment",
                    "External"),
                Relationship("rId4", "urn:example:relative-reference", "relative/resource", "External"));
            CreateWordLikePackage(path, "word/document.xml", relationships);

            Assert.Equal("supported", DocxRuntimeIdentity.Identify(path).Status);
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
    public void Every_relationship_type_must_be_a_well_formed_absolute_uri(string type)
    {
        var path = TemporaryPath(".docx");
        try
        {
            var relationships = RelationshipsDocument(
                Relationship(
                    "rId1",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                    "word/document.xml"),
                Relationship("rId2", type, "word/document.xml", "Internal"));
            CreateWordLikePackage(path, "word/document.xml", relationships);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-declaration", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\" />")]
    [InlineData("<Relationship Id=\"rId1\" Target=\"word/document.xml\" />")]
    [InlineData("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" />")]
    [InlineData("<Relationship Id=\"not valid\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\" />")]
    [InlineData("<Relationship Id=\"rId1\" Type=\"not an absolute URI\" Target=\"word/document.xml\" />")]
    [InlineData("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\" /><Relationship Id=\"rId1\" Type=\"urn:example\" Target=\"metadata.xml\" />")]
    public void Missing_relationship_attributes_or_duplicate_ids_fail_closed(string declarations)
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{declarations}</Relationships>");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-declaration", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Strict_ooxml_office_document_relationship_is_supported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument" Target="word/document.xml" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("supported", evidence.Status);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Malformed_relationship_xml_fails_closed_as_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(path, "word/document.xml", "<Relationships");

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-xml", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Relationship_xml_with_a_doctype_fails_closed_as_unsupported()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordLikePackage(
                path,
                "word/document.xml",
                """
                <!DOCTYPE Relationships [<!ENTITY probe "word/document.xml">]>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="&probe;" />
                </Relationships>
                """);

            var evidence = DocxRuntimeIdentity.Identify(path);

            Assert.Equal("unsupported", evidence.Status);
            Assert.Contains("_rels/.rels:invalid-xml", evidence.File.Signature.Evidence);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Missing_source_is_typed_source_read_failure_without_fake_hash()
    {
        var path = TemporaryPath(".docx");

        var evidence = DocxRuntimeIdentity.Identify(path);

        Assert.Equal("failed", evidence.Status);
        Assert.Equal("source-read", evidence.FailureStage);
        Assert.Null(evidence.Source);
        Assert.Null(evidence.File.FileKind);
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

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("invalid\0path.docx")]
    public void Invalid_source_path_is_a_typed_source_read_failure(string? path)
    {
        var evidence = DocxRuntimeIdentity.Identify(path!);

        Assert.Equal("failed", evidence.Status);
        Assert.Equal("source-read", evidence.FailureStage);
        Assert.Null(evidence.Source);
        Assert.Equal("not-checked", evidence.File.Signature.Status);
        Assert.NotEmpty(evidence.Errors);
    }

    [Fact]
    public void Repeated_identify_is_byte_for_byte_deterministic()
    {
        var path = TemporaryPath(".docx");
        try
        {
            CreateWordDocument(path);

            var first = JsonSerializer.Serialize(DocxRuntimeIdentity.Identify(path), RuntimeJson.Options);
            var second = JsonSerializer.Serialize(DocxRuntimeIdentity.Identify(path), RuntimeJson.Options);

            Assert.Equal(first, second);
        }
        finally
        {
            File.Delete(path);
        }
    }

    private static string TemporaryPath(string extension) =>
        Path.Combine(Path.GetTempPath(), $"docx-runtime-{Guid.NewGuid():N}{extension}");

    private static Uri PackPartUri(string packagePath) =>
        PackUriHelper.CreatePartUri(new Uri($"/{packagePath}", UriKind.Relative));

    private static void CreateWordDocument(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("identity fixture")))));
        mainPart.Document.Save();
    }

    private static void CreateSpreadsheet(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook(
            new DocumentFormat.OpenXml.Spreadsheet.Sheets());
        workbookPart.Workbook.Save();
    }

    private static void CreateWordLikePackage(
        string path,
        string mainPartName,
        string relationshipsXml,
        bool includeMainPart = true)
    {
        CreatePackage(
            path,
            WordContentTypes($"/{mainPartName}"),
            relationshipsXml,
            includeMainPart ? [mainPartName] : []);
    }

    private static void CreatePackage(
        string path,
        string contentTypesXml,
        string relationshipsXml,
        params string[] mainPartNames)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteZipEntry(archive, "[Content_Types].xml", contentTypesXml);
        WriteZipEntry(archive, "_rels/.rels", relationshipsXml);
        foreach (var mainPartName in mainPartNames)
        {
            WriteZipEntry(archive, mainPartName, "<document />");
        }
    }

    private static void CreatePackageWithContentTypeItems(string path, IEnumerable<string> contentTypeItemNames)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var itemName in contentTypeItemNames)
        {
            WriteZipEntry(archive, itemName, WordContentTypes("/word/document.xml"));
        }
        WriteZipEntry(archive, "_rels/.rels", RootOfficeDocumentRelationship("word/document.xml"));
        WriteZipEntry(archive, "word/document.xml", "<document />");
    }

    private static string WordContentTypes(string partName) =>
        $$"""
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Override PartName="{{partName}}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
        </Types>
        """;

    private static string RootOfficeDocumentRelationship(string target) =>
        $$"""
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{{target}}" />
        </Relationships>
        """;

    private static string RelationshipsDocument(params string[] declarations) =>
        $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{string.Concat(declarations)}</Relationships>";

    private static string Relationship(
        string id,
        string type,
        string target,
        string? targetMode = null) =>
        $"<Relationship Id=\"{id}\" Type=\"{type}\" Target=\"{target}\"" +
        (targetMode is null ? string.Empty : $" TargetMode=\"{targetMode}\"") +
        " />";

    private static string ReadZipEntry(string path, string entryName)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry(entryName)!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static void WriteZipEntry(ZipArchive archive, string entryName, string value)
    {
        var entry = archive.CreateEntry(entryName);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream);
        writer.Write(value);
    }
}
