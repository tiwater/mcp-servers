using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Dockit.Convert;

if (args is ["--merge-probe", var sourcePath, var refreshedPath, var outputPath])
{
    DocxFieldResultMerger.Merge(sourcePath, refreshedPath, outputPath);
    return 0;
}

var root = Path.Combine(Path.GetTempPath(), "tiwater-convert-integration-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    var input = Path.Combine(root, "input.docx");
    CreatePackage(input, PageNumberFooter(), BodyControl());
    var sourceBytes = File.ReadAllBytes(input);
    var prepared = DocxWpsRenderNormalizer.Prepare(input, root);
    Require(prepared != input, "page-number wrapper was not admitted");
    Require(sourceBytes.SequenceEqual(File.ReadAllBytes(input)), "source DOCX was modified");

    var footer = ReadPart(prepared, "word/footer1.xml");
    Require(!footer.Contains("Page Numbers (Bottom of Page)", StringComparison.Ordinal),
        "outer page-number wrapper remains");
    Require(!footer.Contains("Page Numbers (Top of Page)", StringComparison.Ordinal),
        "nested page-number wrapper remains");
    Require(footer.Contains("<w:instrText>PAGE</w:instrText>", StringComparison.Ordinal)
            && footer.Contains("<w:instrText>NUMPAGES</w:instrText>", StringComparison.Ordinal),
        "dynamic page fields were removed");
    Require(footer.Contains("Unrelated footer control", StringComparison.Ordinal),
        "unrelated footer content control was removed");
    Require(ReadPart(prepared, "word/document.xml").Contains("Page Numbers (Body fixture)", StringComparison.Ordinal),
        "main-document content control was modified");

    var sourceToc = Path.Combine(root, "source-toc.docx");
    var refreshedToc = Path.Combine(root, "refreshed-toc.docx");
    var mergedToc = Path.Combine(root, "merged-toc.docx");
    CreateDocxPackage(sourceToc, SourceTocDocument(), TocStyles());
    CreateDocxPackage(refreshedToc, RefreshedTocDocument(), TocStyles());
    DocxFieldResultMerger.Merge(sourceToc, refreshedToc, mergedToc);
    VerifyTemplateTocStyles(mergedToc);

    var unchanged = Path.Combine(root, "unchanged.docx");
    CreatePackage(unchanged, UnrelatedFooter(), BodyControl());
    Require(DocxWpsRenderNormalizer.Prepare(unchanged, root) == unchanged,
        "DOCX without a page-number story wrapper was copied");
    Console.WriteLine("convert integration passed");
    return 0;
}
finally
{
    try { Directory.Delete(root, recursive: true); } catch { }
}

static void CreatePackage(string path, string footer, string document)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    Write(archive, "word/footer1.xml", footer);
    Write(archive, "word/document.xml", document);
}

static void CreateDocxPackage(string path, string document, string styles)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    Write(archive, "word/document.xml", document);
    Write(archive, "word/styles.xml", styles);
}

static void Write(ZipArchive archive, string name, string value)
{
    using var stream = archive.CreateEntry(name).Open();
    using var writer = new StreamWriter(stream, new UTF8Encoding(false));
    writer.Write(value);
}

static string ReadPart(string path, string name)
{
    using var archive = ZipFile.OpenRead(path);
    using var reader = new StreamReader(archive.GetEntry(name)!.Open());
    return reader.ReadToEnd();
}

static string PageNumberFooter() => """
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Page Numbers (Bottom of Page)"/></w:docPartObj></w:sdtPr><w:sdtContent>
    <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Page Numbers (Top of Page)"/></w:docPartObj></w:sdtPr><w:sdtContent>
      <w:p><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:t> / </w:t></w:r><w:r><w:instrText>NUMPAGES</w:instrText></w:r></w:p>
    </w:sdtContent></w:sdt>
  </w:sdtContent></w:sdt>
  <w:sdt><w:sdtPr><w:tag w:val="unrelated"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Unrelated footer control</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:ftr>
""";

static string UnrelatedFooter() => """
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:sdt><w:sdtPr><w:tag w:val="unrelated"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Unrelated footer control</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:ftr>
""";

static string BodyControl() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Page Numbers (Body fixture)"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:instrText>PAGE</w:instrText></w:r></w:p></w:sdtContent></w:sdt>
</w:body></w:document>
""";

static string TocStyles() => """
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="HeadingOne"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="HeadingThree"><w:name w:val="heading 3"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="TemplateTocOne"><w:name w:val="toc 1"/><w:pPr><w:ind w:leftChars="0"/></w:pPr><w:rPr><w:i w:val="0"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TemplateTocThree"><w:name w:val="toc 3"/><w:pPr><w:ind w:leftChars="400"/></w:pPr><w:rPr><w:i w:val="0"/></w:rPr></w:style>
</w:styles>
""";

static string SourceTocDocument() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
  <w:p w14:paraId="HEAD0001"><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>Top heading</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="TemplateTocOne"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \o "1-3" \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p>
  <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p w14:paraId="HEAD0003"><w:pPr><w:pStyle w:val="HeadingThree"/></w:pPr><w:r><w:t>Nested heading</w:t></w:r></w:p>
</w:body></w:document>
""";

static string RefreshedTocDocument() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
  <w:p w14:paraId="HEAD0001"><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:bookmarkStart w:id="41" w:name="_TocFresh1"/><w:r><w:rPr><w:i/></w:rPr><w:t>Top heading</w:t></w:r><w:bookmarkEnd w:id="41"/></w:p>
  <w:p><w:pPr><w:pStyle w:val="WrongListStyle"/><w:ind w:firstLine="420"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \o "1-3" \h \z \u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \l _TocFresh1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Top entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="WrongListStyle"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \l _TocFresh3 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Nested entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p w14:paraId="HEAD0003"><w:pPr><w:pStyle w:val="HeadingThree"/></w:pPr><w:bookmarkStart w:id="42" w:name="_TocFresh3"/><w:r><w:t>Nested heading</w:t></w:r><w:bookmarkEnd w:id="42"/></w:p>
</w:body></w:document>
""";

static void VerifyTemplateTocStyles(string path)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var document = XDocument.Parse(ReadPart(path, "word/document.xml"));
    var entries = document.Descendants(w + "p")
        .Where(paragraph => paragraph.Descendants(w + "t").Any(text => text.Value.EndsWith("entry", StringComparison.Ordinal)))
        .ToDictionary(paragraph => paragraph.Descendants(w + "t").Single().Value, StringComparer.Ordinal);
    Require((string?)entries["Top entry"].Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val") == "TemplateTocOne",
        "level-one TOC entry did not retain the template TOC style");
    Require((string?)entries["Nested entry"].Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val") == "TemplateTocThree",
        "level-three TOC entry did not retain the template TOC style");
    Require(entries.Values.All(paragraph => paragraph.Element(w + "pPr")?.Elements().Count() == 1),
        "refreshed TOC paragraph direct formatting overrides the template style");
    Require(entries.Values.SelectMany(paragraph => paragraph.Descendants(w + "r"))
            .Where(run => run.Descendants(w + "t").Any()).All(run => run.Element(w + "rPr") is null),
        "refreshed TOC text direct formatting overrides the template style");
}

static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
