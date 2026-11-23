using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Dockit.Convert;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Dockit.Convert.Tests;

public sealed class Wave3cConvertTests : IDisposable
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"tiwater-convert-wave3c-{Guid.NewGuid():N}");

    public Wave3cConvertTests()
    {
        Directory.CreateDirectory(_root);
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    [Fact]
    public void Merge_WithActualPageCount_RefreshesHeaderComplexNumPagesAppPagesAndSettingsUpdateFields()
    {
        var sourcePath = Path.Combine(_root, "ana08-f4-source.docx");
        var refreshedPath = Path.Combine(_root, "ana08-f4-refreshed.docx");
        var outputPath = Path.Combine(_root, "ana08-f4-output.docx");

        // Source has stale cached NUMPAGES="17" inside a header table, app.xml Pages=17, and settings without updateFields.
        CreateFullDocx(sourcePath, headerNumPages: "17", footerNumPages: "17", appPages: "17", includeUpdateFields: false);
        // Simulate WPS saving refreshed copy where header table NUMPAGES was still stale "17", while actualPageCount=19 is known.
        CreateFullDocx(refreshedPath, headerNumPages: "17", footerNumPages: "17", appPages: "19", includeUpdateFields: false);

        DocxFieldResultMerger.Merge(sourcePath, refreshedPath, outputPath, actualPageCount: 19);

        var header = XDocument.Parse(ReadPart(outputPath, "word/header1.xml"));
        Assert.Equal("页码(Page)：1 / 19", ExtractText(header));
        // Verify run formatting on the cached NUMPAGES value was preserved
        var numPagesTextNode = header.Descendants(W + "t").Single(t => t.Value == "19");
        Assert.Equal("SimSun", (string?)numPagesTextNode.Parent?.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "eastAsia"));

        var footer = XDocument.Parse(ReadPart(outputPath, "word/footer1.xml"));
        Assert.Equal("Total: 19 | Ref: Table 3", ExtractText(footer));

        var app = XDocument.Parse(ReadPart(outputPath, "docProps/app.xml"));
        Assert.Equal("19", app.Descendants().Single(e => e.Name.LocalName == "Pages").Value);

        var settings = XDocument.Parse(ReadPart(outputPath, "word/settings.xml"));
        var updateFields = settings.Root!.Element(W + "updateFields");
        Assert.NotNull(updateFields);
        Assert.Equal("true", (string?)updateFields!.Attribute(W + "val"));

        // Verify schema validity via OpenXmlValidator (including updateFields placement before compat/rsids/themeFontLang)
        using var openXml = WordprocessingDocument.Open(outputPath, false);
        var errors = new OpenXmlValidator().Validate(openXml).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Id}: {e.Description}")));
    }

    [Fact]
    public void Merge_WithoutActualPageCount_UsesRefreshedHeaderFieldsAndAppPagesFallback()
    {
        // Case A: refreshed header1.xml has updated NUMPAGES="3" and PAGE="1" (e.g., record output).
        var sourceA = Path.Combine(_root, "ana08-f4-record-source.docx");
        var refreshedA = Path.Combine(_root, "ana08-f4-record-refreshed.docx");
        var outputA = Path.Combine(_root, "ana08-f4-record-output.docx");

        CreateFullDocx(sourceA, headerNumPages: "17", footerNumPages: "17", appPages: "17", includeUpdateFields: false);
        CreateFullDocx(refreshedA, headerNumPages: "3", footerNumPages: "3", appPages: "3", includeUpdateFields: false, headerPage: "1");

        DocxFieldResultMerger.Merge(sourceA, refreshedA, outputA);

        Assert.Equal("页码(Page)：1 / 3", ExtractText(XDocument.Parse(ReadPart(outputA, "word/header1.xml"))));
        Assert.Equal("Total: 3 | Ref: Table 3", ExtractText(XDocument.Parse(ReadPart(outputA, "word/footer1.xml"))));
        Assert.Equal("true", (string?)XDocument.Parse(ReadPart(outputA, "word/settings.xml")).Root!.Element(W + "updateFields")?.Attribute(W + "val"));

        // Case B: refreshed header still says "17", but refreshed docProps/app.xml reports Pages=24 (unseen synthetic variant).
        var sourceB = Path.Combine(_root, "ana08-f4-app-source.docx");
        var refreshedB = Path.Combine(_root, "ana08-f4-app-refreshed.docx");
        var outputB = Path.Combine(_root, "ana08-f4-app-output.docx");

        CreateFullDocx(sourceB, headerNumPages: "17", footerNumPages: "17", appPages: "17", includeUpdateFields: true, updateFieldsVal: "false");
        CreateFullDocx(refreshedB, headerNumPages: "17", footerNumPages: "17", appPages: "24", includeUpdateFields: false);

        DocxFieldResultMerger.Merge(sourceB, refreshedB, outputB);

        Assert.Equal("页码(Page)：1 / 24", ExtractText(XDocument.Parse(ReadPart(outputB, "word/header1.xml"))));
        Assert.Equal("Total: 24 | Ref: Table 3", ExtractText(XDocument.Parse(ReadPart(outputB, "word/footer1.xml"))));
        Assert.Equal("24", XDocument.Parse(ReadPart(outputB, "docProps/app.xml")).Descendants().Single(e => e.Name.LocalName == "Pages").Value);
        Assert.Equal("true", (string?)XDocument.Parse(ReadPart(outputB, "word/settings.xml")).Root!.Element(W + "updateFields")?.Attribute(W + "val"));
    }

    [Fact]
    public void Merge_HandlesComplexNumPagesFieldWithoutPriorSeparateRun()
    {
        var sourcePath = Path.Combine(_root, "ana08-f4-nosep-source.docx");
        var refreshedPath = Path.Combine(_root, "ana08-f4-nosep-refreshed.docx");
        var outputPath = Path.Combine(_root, "ana08-f4-nosep-output.docx");

        const string headerWithoutSeparate = """
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t>Pages: </w:t></w:r>
    <w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:hdr>
""";
        CreateFullDocx(sourcePath, headerNumPages: "17", footerNumPages: "17", appPages: "17", includeUpdateFields: false, customHeaderXml: headerWithoutSeparate);
        CreateFullDocx(refreshedPath, headerNumPages: "17", footerNumPages: "17", appPages: "11", includeUpdateFields: false, customHeaderXml: headerWithoutSeparate);

        DocxFieldResultMerger.Merge(sourcePath, refreshedPath, outputPath, actualPageCount: 11);

        var header = XDocument.Parse(ReadPart(outputPath, "word/header1.xml"));
        Assert.Equal("Pages: 11", ExtractText(header));
        Assert.Contains(header.Descendants(W + "fldChar"), e => (string?)e.Attribute(W + "fldCharType") == "separate");

        using var openXml = WordprocessingDocument.Open(outputPath, false);
        var errors = new OpenXmlValidator().Validate(openXml).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Id}: {e.Description}")));
    }

    [Fact]
    public void RefreshFieldsHelperScript_RefreshesHeaderFooterFieldsAndPreservesIndexPaginationOrder()
    {
        var script = WpsPdfConverter.RefreshFieldsHelperScript;

        var firstFigures = script.IndexOf("TableOfFigures.Update", StringComparison.Ordinal);
        var lastContents = script.LastIndexOf("TableOfContents.Update\"", StringComparison.Ordinal);
        var finalPagination = script.LastIndexOf("document.Repaginate()", StringComparison.Ordinal);
        var lastFigures = script.LastIndexOf("TableOfFigures.Update\"", StringComparison.Ordinal);
        var headerFooterStage = script.IndexOf("header-footer-fields", StringComparison.Ordinal);

        Assert.True(firstFigures >= 0 && lastContents > firstFigures);
        Assert.True(finalPagination > lastContents && lastFigures > finalPagination);
        Assert.True(headerFooterStage > lastFigures);
        Assert.Contains("get_Headers", script, StringComparison.Ordinal);
        Assert.Contains("get_Footers", script, StringComparison.Ordinal);
        Assert.Contains("wdStatisticPages", script, StringComparison.Ordinal);
        Assert.DoesNotContain("UpdatePageNumbers", script, StringComparison.Ordinal);
        Assert.DoesNotContain("SaveAs2", script, StringComparison.Ordinal);

        var marker = Path.Combine(_root, "marker.txt");
        File.WriteAllText(marker, "complete\npages=19\n");
        Assert.Equal(19, WpsPdfConverter.ReadCompletionMarkerPageCount(marker));

        File.WriteAllText(marker, "complete\n");
        Assert.Null(WpsPdfConverter.ReadCompletionMarkerPageCount(marker));
    }

    private static string ExtractText(XDocument doc)
        => string.Concat(doc.Descendants(W + "t").Select(e => e.Value));

    private static string ReadPart(string path, string partName)
    {
        using var zip = ZipFile.OpenRead(path);
        using var reader = new StreamReader(zip.GetEntry(partName)!.Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    private static void CreateFullDocx(
        string path,
        string headerNumPages,
        string footerNumPages,
        string appPages,
        bool includeUpdateFields,
        string updateFieldsVal = "true",
        string headerPage = "1",
        string? customHeaderXml = null)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        Write(zip, "[Content_Types].xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
""");
        Write(zip, "_rels/.rels", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
""");
        Write(zip, "word/_rels/document.xml.rels", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>
""");
        Write(zip, "word/styles.xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>
""");
        var updateFieldsXml = includeUpdateFields ? $"<w:updateFields w:val=\"{updateFieldsVal}\"/>" : string.Empty;
        Write(zip, "word/settings.xml", $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  {updateFieldsXml}
  <w:compat/>
  <w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>
</w:settings>
""");
        Write(zip, "docProps/app.xml", $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Pages>{appPages}</Pages>
</Properties>
""");
        Write(zip, "word/header1.xml", customHeaderXml ?? $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:tbl>
    <w:tblPr/>
    <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
    <w:tr>
      <w:tc>
        <w:p>
          <w:r><w:t>页码(Page)：</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:t>{headerPage}</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
          <w:r><w:t xml:space="preserve"> / </w:t></w:r>
          <w:r><w:fldChar w:fldCharType="begin"/></w:r>
          <w:r><w:instrText xml:space="preserve"> NUMPAGES \* MERGEFORMAT </w:instrText></w:r>
          <w:r><w:fldChar w:fldCharType="separate"/></w:r>
          <w:r><w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>{headerNumPages}</w:t></w:r>
          <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
</w:hdr>
""");
        Write(zip, "word/footer1.xml", $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Total: </w:t></w:r>
    <w:fldSimple w:instr=" NUMPAGES \* ARABIC ">
      <w:r><w:t>{footerNumPages}</w:t></w:r>
    </w:fldSimple>
    <w:r><w:t xml:space="preserve"> | Ref: </w:t></w:r>
    <w:fldSimple w:instr=" PAGEREF _Ref1 ">
      <w:r><w:t>Table 3</w:t></w:r>
    </w:fldSimple>
  </w:p>
</w:ftr>
""");
        Write(zip, "word/document.xml", """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Document body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId3"/>
      <w:footerReference w:type="default" r:id="rId4"/>
    </w:sectPr>
  </w:body>
</w:document>
""");
    }

    private static void Write(ZipArchive zip, string name, string content)
    {
        using var stream = zip.CreateEntry(name).Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(content);
    }
}
