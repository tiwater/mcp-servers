using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Dockit.Convert;

if (args is ["--verify-native-index-pages", var indexedDocx, var nativePdf, var expectedEntryCount])
{
    var w = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    var document = XDocument.Parse(ReadPart(indexedDocx, "word/document.xml"));
    using var pdf = UglyToad.PdfPig.PdfDocument.Open(nativePdf);
    static string Compact(string value) => string.Concat(value.Where(character => !char.IsWhiteSpace(character)));
    var pages = pdf.GetPages().Select(page => (page.Number, Text: Compact(page.Text))).ToList();
    var checkedCount = 0;
    var mismatches = new List<string>();
    foreach (var paragraph in document.Descendants(w + "p"))
    {
        if (!paragraph.Descendants(w + "hyperlink").Any()
            && !paragraph.Descendants(w + "instrText").Any(node => node.Value.Contains("PAGEREF", StringComparison.Ordinal))) continue;
        var texts = paragraph.Descendants(w + "t").Select(node => node.Value).ToList();
        if (texts.Count < 2 || !int.TryParse(texts[^1].Trim(), out var cachedPage)) continue;
        var label = Compact(string.Concat(texts.Take(texts.Count - 1)));
        if (label.Length == 0) continue;
        var matches = pages.Where(page => page.Text.Contains(label, StringComparison.Ordinal)).ToList();
        Require(matches.Count > 0, "native PDF does not contain index label: " + label);
        var actualPage = matches[^1].Number;
        checkedCount++;
        if (cachedPage != actualPage) mismatches.Add($"{label}: cached={cachedPage}, native={actualPage}");
    }
    Require(int.TryParse(expectedEntryCount, out var expectedCount) && expectedCount > 0
        && checkedCount == expectedCount, "native index comparison count differs from the declared fixture");
    foreach (var mismatch in mismatches) Console.WriteLine(mismatch);
    Console.WriteLine($"native index entries={checkedCount}; mismatches={mismatches.Count}");
    Require(mismatches.Count == 0, "cached index pages differ from native body pagination");
    return 0;
}

if (args is ["--index-pagination-order-probe"])
{
    var script = WpsPdfConverter.RefreshFieldsHelperScript;
    var lastContents = script.LastIndexOf("TableOfContents.Update\"", StringComparison.Ordinal);
    var finalPagination = script.LastIndexOf("document.Repaginate()", StringComparison.Ordinal);
    var lastFigures = script.LastIndexOf("TableOfFigures.Update\"", StringComparison.Ordinal);
    Require(lastContents >= 0 && finalPagination > lastContents && lastFigures > finalPagination,
        "figure index pages must be recomputed after contents expansion and final pagination");
    Console.WriteLine("index pagination ordering regression passed");
    return 0;
}

if (args is ["--index-pagination-fixture", var fixturePath])
{
    var body = new StringBuilder();
    body.Append("<w:p><w:r><w:t>Independent index pagination fixture</w:t></w:r></w:p>");
    foreach (var code in new[] { " TOC \\o &quot;1-1&quot; \\h ", " TOC \\c &quot;Figure&quot; \\h " })
        body.Append($"<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>{code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>Unrefreshed index</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
    for (var index = 1; index <= 26; index++)
    {
        body.Append($"<w:p><w:pPr><w:pStyle w:val=\"HeadingOne\"/></w:pPr><w:r><w:t>Independent section {index:D2}</w:t></w:r></w:p>");
        body.Append($"<w:p><w:r><w:t>Figure </w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> SEQ Figure \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>{index}</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:t> Independent marker {index:D2}</w:t></w:r></w:p>");
    }
    using var package = ZipFile.Open(fixturePath, ZipArchiveMode.Create);
    Write(package, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/></Types>");
    Write(package, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
    Write(package, "word/_rels/document.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
    Write(package, "word/styles.xml", TocStyles());
    Write(package, "word/document.xml", $"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>");
    return 0;
}

if (args is ["--lima-guest-timeout-probe"])
{
    var docx = LimaWpsPdfConverter.CreateDocumentFieldRefreshStartInfo(
        "/usr/bin/limactl", "isolated-wps", "/shared/input.docx", "/shared/output.docx");
    var spreadsheet = LimaWpsPdfConverter.CreateSpreadsheetConversionStartInfo(
        "/usr/bin/limactl", "isolated-wps", "/shared/input.xlsx", "/shared/output.xlsx", "recalculate-xlsx");
    var pdf = LimaWpsPdfConverter.CreateProcessStartInfo(
        "/usr/bin/limactl", "isolated-wps", "/shared/input.docx", "/shared/output.pdf");

    Require(docx.ArgumentList[^1].Contains(
        "timeout --kill-after=5s 230s tiwater-convert refresh-docx-fields '/shared/input.docx' '/shared/output.docx'",
        StringComparison.Ordinal), "DOCX field refresh is not bounded inside the Lima guest");
    Require(spreadsheet.ArgumentList[^1].Contains(
        "timeout --kill-after=5s 590s tiwater-convert recalculate-xlsx '/shared/input.xlsx' '/shared/output.xlsx'",
        StringComparison.Ordinal), "spreadsheet conversion is not bounded inside the Lima guest");
    Require(pdf.ArgumentList[^1].Contains(
        "timeout --kill-after=5s 650s tiwater-convert docx-to-pdf '/shared/input.docx' '/shared/output.pdf'",
        StringComparison.Ordinal), "PDF conversion is not bounded inside the Lima guest");
    Console.WriteLine("Lima guest timeout integration passed");
    return 0;
}

if (args is ["--lima-guest-version-probe"])
{
    var versionProbeRoot = Path.Combine(Path.GetTempPath(), "tiwater-lima-version-" + Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(versionProbeRoot);
    try
    {
        var input = Path.Combine(versionProbeRoot, "input.docx");
        var output = Path.Combine(versionProbeRoot, "output.docx");
        File.WriteAllText(input, "current input");
        File.WriteAllText(output, "current output");
        var inputHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(input))).ToLowerInvariant();
        var outputHash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(output))).ToLowerInvariant();
        string Receipt(string version) => $$"""
        {"schema":"tiwater.convert-refresh-docx-fields/v1","status":"ok","input_sha256":"{{inputHash}}","output_sha256":"{{outputHash}}","source_format":"docx","target_format":"docx","version":"{{version}}","backend":"wps","refresh_scope":["table-of-contents","table-of-figures"]}
        """;
        LimaWpsPdfConverter.ValidateDocumentFieldRefreshEvidence(Receipt("0.9.37"), input, output);
        var rejected = false;
        try
        {
            LimaWpsPdfConverter.ValidateDocumentFieldRefreshEvidence(Receipt("0.9.36"), input, output);
        }
        catch (InvalidOperationException error) when (error.Message ==
            "Lima WPS guest runtime version mismatch: expected 0.9.37, received 0.9.36.")
        {
            rejected = true;
        }
        Require(rejected, "Lima host accepted field-refresh evidence from a stale guest runtime");
        Console.WriteLine("Lima guest version integration passed");
        return 0;
    }
    finally
    {
        try { Directory.Delete(versionProbeRoot, recursive: true); } catch { }
    }
}

if (args is ["--wps-automation-process-probe"])
{
    Require(WpsRpcSession.DocumentFieldRefreshTimeout == TimeSpan.FromSeconds(210),
        "document refresh must finish cleanup before the 230-second Lima guest deadline");
    Require(WpsRpcSession.IsWpsAutomationCommandLine(new[] {
        "/opt/kingsoft/wps-office/office6/wps", "-automation", "-rpcserverport=/wpsrpc-123-456"
    }), "direct WPS automation process was not identified");
    Require(WpsRpcSession.IsWpsAutomationCommandLine(new[] {
        "/bin/bash", "/usr/bin/wps", "-automation", "-rpcserverport=/wpsrpc-123-456"
    }), "WPS launcher automation process was not identified");
    Require(!WpsRpcSession.IsWpsAutomationCommandLine(new[] {
        "/opt/kingsoft/wps-office/office6/wps", "/home/customer/report.docx"
    }), "interactive WPS process must not be identified as an automation session");
    Require(!WpsRpcSession.IsWpsAutomationCommandLine(new[] {
        "/usr/bin/wpp", "-automation", "-rpcserverport=/wpsrpc-123-456"
    }), "a different Office application must not be identified as Writer automation");
    var refreshScript = WpsPdfConverter.RefreshFieldsHelperScript;
    var figureUpdate = refreshScript.IndexOf("TableOfFigures.Update", StringComparison.Ordinal);
    var contentsUpdate = refreshScript.IndexOf("TableOfContents.Update", StringComparison.Ordinal);
    Require(figureUpdate >= 0 && contentsUpdate > figureUpdate,
        "WPS must update figure indexes before contents indexes");
    Require(!refreshScript.Contains("UpdatePageNumbers", StringComparison.Ordinal),
        "full index updates must not be followed by a duplicate WPS page-number update");
    Require(refreshScript.Contains("shutil.copy2(input_path, output_path)", StringComparison.Ordinal)
            && refreshScript.Contains("documents.Open(output_path", StringComparison.Ordinal)
            && refreshScript.Contains("Document.Save\", document.Save()", StringComparison.Ordinal)
            && !refreshScript.Contains("SaveAs2", StringComparison.Ordinal),
        "WPS refresh must save an isolated working copy in place");
    Console.WriteLine("WPS automation process integration passed");
    return 0;
}

if (args is ["--merge-probe", var sourcePath, var refreshedPath, var outputPath])
{
    DocxFieldResultMerger.Merge(sourcePath, refreshedPath, outputPath);
    return 0;
}

if (args is ["--inline-boundary-probe", var probeRoot])
{
    Directory.CreateDirectory(probeRoot);
    var inlineSourcePath = Path.Combine(probeRoot, "inline-source.docx");
    var inlineRefreshedPath = Path.Combine(probeRoot, "inline-refreshed.docx");
    var inlineOutputPath = Path.Combine(probeRoot, "inline-output.docx");
    CreateDocxPackage(inlineSourcePath, InlineBoundarySource(), TocStyles());
    CreateDocxPackage(inlineRefreshedPath, InlineBoundaryRefreshed(), TocStyles());
    DocxFieldResultMerger.Merge(inlineSourcePath, inlineRefreshedPath, inlineOutputPath);
    VerifyInlineBoundary(inlineOutputPath);
    Console.WriteLine("inline field boundary integration passed");
    return 0;
}

if (args is ["--inline-toc-end-probe", var inlineTocRoot])
{
    Directory.CreateDirectory(inlineTocRoot);
    var inlineTocSourcePath = Path.Combine(inlineTocRoot, "inline-toc-end-source.docx");
    var inlineTocRefreshedPath = Path.Combine(inlineTocRoot, "inline-toc-end-refreshed.docx");
    var inlineTocOutputPath = Path.Combine(inlineTocRoot, "inline-toc-end-output.docx");
    CreateDocxPackage(inlineTocSourcePath, InlineTocEndSource(), TocStyles());
    CreateDocxPackage(inlineTocRefreshedPath, InlineTocEndRefreshed(), TocStyles());
    DocxFieldResultMerger.Merge(inlineTocSourcePath, inlineTocRefreshedPath, inlineTocOutputPath);
    VerifyInlineTocEndStyle(inlineTocOutputPath);
    Console.WriteLine("inline TOC end boundary integration passed");
    return 0;
}

if (args is ["--duplicate-bookmark-end-probe", var duplicateBookmarkRoot])
{
    Directory.CreateDirectory(duplicateBookmarkRoot);
    var duplicateSourcePath = Path.Combine(duplicateBookmarkRoot, "source.docx");
    var duplicateRefreshedPath = Path.Combine(duplicateBookmarkRoot, "refreshed.docx");
    var duplicateOutputPath = Path.Combine(duplicateBookmarkRoot, "output.docx");
    var source = SourceTocDocument().Replace(
        "</w:body>", "<w:bookmarkEnd w:id=\"0\"/></w:body>", StringComparison.Ordinal);
    var refreshed = RefreshedTocDocument()
        .Replace("w:id=\"41\" w:name=\"_TocFresh1\"", "w:id=\"0\" w:name=\"_TocFresh1\"", StringComparison.Ordinal)
        .Replace("<w:bookmarkEnd w:id=\"41\"/>", "<w:bookmarkEnd w:id=\"0\"/>", StringComparison.Ordinal)
        .Replace("</w:body>", "<w:bookmarkEnd w:id=\"0\"/></w:body>", StringComparison.Ordinal);
    CreateDocxPackage(duplicateSourcePath, source, TocStyles());
    CreateDocxPackage(duplicateRefreshedPath, refreshed, TocStyles());
    DocxFieldResultMerger.Merge(duplicateSourcePath, duplicateRefreshedPath, duplicateOutputPath);
    VerifyDuplicateBookmarkEndRecovery(duplicateOutputPath);
    Console.WriteLine("duplicate bookmark end integration passed");
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

static string InlineTocEndSource() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
  <w:p><w:pPr><w:pStyle w:val="TemplateTocOne"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \\l _TocInlineTarget </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Top entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="TemplateTocOne"/><w:ind w:firstLine="315"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \\l _TocSecondTarget </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Second entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p w14:paraId="ABCD0001"><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkStart w:id="51" w:name="_TocInlineTarget"/><w:r><w:t>Top heading</w:t></w:r><w:bookmarkEnd w:id="51"/></w:p>
  <w:p w14:paraId="ABCD0002"><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:bookmarkStart w:id="52" w:name="_TocSecondTarget"/><w:r><w:t>Second heading</w:t></w:r><w:bookmarkEnd w:id="52"/></w:p>
</w:body></w:document>
""";

static string InlineTocEndRefreshed() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>
  <w:p><w:pPr><w:pStyle w:val="WrongListStyle"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \\l _TocRefreshedTarget </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Top entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="WrongListStyle"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK \\l _TocRefreshedSecond </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Second entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p w14:paraId="ABCD0001"><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkStart w:id="61" w:name="_TocRefreshedTarget"/><w:r><w:t>Top heading</w:t></w:r><w:bookmarkEnd w:id="61"/></w:p>
  <w:p w14:paraId="ABCD0002"><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:bookmarkStart w:id="62" w:name="_TocRefreshedSecond"/><w:r><w:t>Second heading</w:t></w:r><w:bookmarkEnd w:id="62"/></w:p>
</w:body></w:document>
""";

static string InlineBoundarySource() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:pStyle w:val="SourceBoundary"/></w:pPr><w:bookmarkStart w:id="90" w:name="_TocOutsideBefore"/><w:bookmarkEnd w:id="90"/><w:r><w:rPr><w:b/></w:rPr><w:t>BEFORE</w:t><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \c "Figure" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>OLD RESULT</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/><w:t>AFTER</w:t></w:r><w:bookmarkStart w:id="91" w:name="_TocOutsideAfter"/><w:bookmarkEnd w:id="91"/></w:p>
</w:body></w:document>
""";

static string InlineBoundaryRefreshed() => """
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p><w:pPr><w:pStyle w:val="RefreshedBoundary"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>WPS BEFORE</w:t><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC \c "Figure" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>NEW RESULT</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/><w:t>WPS AFTER</w:t></w:r></w:p>
</w:body></w:document>
""";

static void VerifyInlineBoundary(string path)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var document = XDocument.Parse(ReadPart(path, "word/document.xml"));
    var paragraph = document.Descendants(w + "p").Single();
    var text = string.Concat(paragraph.Descendants(w + "t").Select(element => element.Value));
    Require(text == "BEFORENEW RESULTAFTER", "field merge did not replace only the inline result boundary");
    Require((string?)paragraph.Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val") == "SourceBoundary",
        "field merge replaced the source boundary paragraph properties");
    Require(paragraph.Elements(w + "r").Single(run => run.Elements(w + "t").Any(text => text.Value == "BEFORE"))
            .Element(w + "rPr")?.Element(w + "b") is not null,
        "field merge lost source run formatting before an inline field boundary");
    Require(paragraph.Descendants(w + "fldChar")
            .Single(element => (string?)element.Attribute(w + "fldCharType") == "begin")
            .Parent?.Element(w + "rPr")?.Element(w + "i") is not null,
        "field merge lost refreshed run formatting at an inline field boundary");
    Require(paragraph.Elements(w + "bookmarkStart").Select(element => (string?)element.Attribute(w + "name"))
            .SequenceEqual(["_TocOutsideBefore", "_TocOutsideAfter"]),
        "field merge removed source bookmarks outside the field boundary");
    Require(paragraph.Elements(w + "bookmarkEnd").Count() == 2,
        "field merge removed source bookmark ends outside the field boundary");
    Require(!text.Contains("OLD RESULT", StringComparison.Ordinal)
            && !text.Contains("WPS BEFORE", StringComparison.Ordinal)
            && !text.Contains("WPS AFTER", StringComparison.Ordinal),
        "field merge retained content outside the selected refreshed field boundary");
}

static void VerifyInlineTocEndStyle(string path)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var document = XDocument.Parse(ReadPart(path, "word/document.xml"));
    var entries = document.Descendants(w + "p")
        .Where(paragraph => paragraph.Descendants(w + "t").Any(text => text.Value.EndsWith("entry", StringComparison.Ordinal)))
        .ToList();
    Require(entries.Count == 2 && entries.All(entry =>
            (string?)entry.Element(w + "pPr")?.Element(w + "pStyle")?.Attribute(w + "val") == "TemplateTocOne"),
        "TOC entry did not use the source template style when its target follows the field end in the same paragraph");
    var bookmarkNames = document.Descendants(w + "bookmarkStart")
        .Select(start => (string?)start.Attribute(w + "name"))
        .Where(name => !string.IsNullOrWhiteSpace(name))
        .ToHashSet(StringComparer.OrdinalIgnoreCase);
    Require(bookmarkNames.Contains("_TocRefreshedTarget") && bookmarkNames.Contains("_TocRefreshedSecond"),
        "TOC target bookmark after an inline field end was not copied back to the source document");
}

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

static void VerifyDuplicateBookmarkEndRecovery(string path)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var document = XDocument.Parse(ReadPart(path, "word/document.xml"));
    var start = document.Descendants(w + "bookmarkStart")
        .Single(element => (string?)element.Attribute(w + "name") == "_TocFresh1");
    var id = (string?)start.Attribute(w + "id");
    var ends = document.Descendants(w + "bookmarkEnd")
        .Where(element => (string?)element.Attribute(w + "id") == id)
        .ToList();
    Require(id != "0", "copied TOC bookmark reused an existing bookmark-end identity");
    Require(ends.Count == 1, "copied TOC bookmark does not have one unique end");
    Require(ReferenceEquals(start.Ancestors(w + "p").FirstOrDefault(), ends[0].Ancestors(w + "p").FirstOrDefault()),
        "copied TOC bookmark did not retain its same-paragraph end");
    Require(document.Descendants(w + "bookmarkEnd")
            .Select(element => (string?)element.Attribute(w + "id"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .GroupBy(value => value, StringComparer.Ordinal)
            .All(group => group.Count() == 1),
        "merged DOCX contains duplicate bookmark-end identities");
}

static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
