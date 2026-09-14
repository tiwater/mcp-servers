using System.IO.Compression;
using System.Security.Cryptography;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Dockit.Convert;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

if (args is ["--xlsx-formula-cache-merge-probe", var xlsxMergeRoot])
{
    RunXlsxFormulaCacheMergeProbe(xlsxMergeRoot);
    Console.WriteLine("xlsx formula cache merge integration passed");
    return 0;
}

if (args is ["--legacy-xls-font-recalculation-probe", var legacyXlsRoot])
{
    RunLegacyXlsFontRecalculationProbe(legacyXlsRoot);
    Console.WriteLine("legacy XLS font recalculation integration passed");
    return 0;
}

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

static void RunXlsxFormulaCacheMergeProbe(string root)
{
    Directory.CreateDirectory(root);
    var source = Path.Combine(root, "source.xlsx");
    var recalculated = Path.Combine(root, "recalculated.xlsx");
    var output = Path.Combine(root, "output.xlsx");
    const string sourceStyles = """
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2">
  <font><sz val="11"/><color rgb="FFFF0000"/><name val="Times New Roman"/><family val="1"/></font>
  <font><b/><sz val="9.5"/><color rgb="FF0000FF"/><name val="Arial"/><family val="2"/></font>
</fonts></styleSheet>
""";
    const string recalculatedStyles = """
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2">
  <font><sz val="14"/><color rgb="FF00FF00"/><name val="Calibri"/></font>
  <font><sz val="9.5"/><color rgb="FF0000FF"/><name val="Arial"/></font>
</fonts></styleSheet>
""";
    const string sourceSheet = """
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">
  <c r="A1" s="1"><v>41</v></c><c r="B1" s="2"><f>A1*2</f><v>0</v></c><c r="C1" s="2"><f t="shared" ref="C1:C2" si="0">A1+1</f><v>0</v></c>
</row><row r="2"><c r="C2" s="2"><f t="shared" si="0"></f><v>0</v></c></row></sheetData></worksheet>
""";
    const string recalculatedSheet = """
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">
  <c r="A1" s="99"><v>999</v></c><c r="B1" s="99"><f>A1*2</f><v>82</v></c><c r="C1" s="99" t="str"><f t="shared" ref="C1:C2" si="0">A1+1</f><v>42</v></c>
</row><row r="2"><c r="C2" s="99"><f t="shared" si="0"></f><v>43</v></c></row></sheetData></worksheet>
""";
    CreateSyntheticXlsx(source, sourceStyles, sourceSheet);
    CreateSyntheticXlsx(recalculated, recalculatedStyles, recalculatedSheet);
    XlsxFormulaCacheMerger.Merge(source, recalculated, output);

    Require(ReadPart(output, "xl/styles.xml") == sourceStyles, "formula-cache merge changed distinct font family, name, size, weight, or color semantics");
    Require(ReadPart(output, "docProps/custom.xml") == "<properties><marker>source-authoritative</marker></properties>",
        "formula-cache merge changed a non-calculation package part");
    var merged = XDocument.Parse(ReadPart(output, "xl/worksheets/sheet1.xml"));
    XNamespace x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    var cells = merged.Descendants(x + "c").ToDictionary(cell => (string)cell.Attribute("r")!, StringComparer.Ordinal);
    Require((string?)cells["A1"].Element(x + "v") == "41" && (string?)cells["A1"].Attribute("s") == "1",
        "formula-cache merge imported an unplanned non-formula mutation");
    Require((string?)cells["B1"].Element(x + "v") == "82" && (string?)cells["B1"].Attribute("s") == "2",
        "formula-cache merge did not import a numeric cache while preserving style");
    Require((string?)cells["C1"].Element(x + "v") == "42" && (string?)cells["C1"].Attribute("t") == "str",
        "formula-cache merge did not import the refreshed result type");
    Require((string?)cells["C2"].Element(x + "v") == "43", "formula-cache merge did not import a shared-formula cache");

    var changedFormula = Path.Combine(root, "changed-formula.xlsx");
    CreateSyntheticXlsx(changedFormula, recalculatedStyles, recalculatedSheet.Replace("A1*2", "A1*3", StringComparison.Ordinal));
    RequireThrows(() => XlsxFormulaCacheMerger.Merge(source, changedFormula, Path.Combine(root, "changed-formula-output.xlsx")), "changed a formula");

    var missingFormula = Path.Combine(root, "missing-formula.xlsx");
    CreateSyntheticXlsx(missingFormula, recalculatedStyles, recalculatedSheet.Replace("<f>A1*2</f>", "", StringComparison.Ordinal));
    RequireThrows(() => XlsxFormulaCacheMerger.Merge(source, missingFormula, Path.Combine(root, "missing-formula-output.xlsx")), "formula cell inventory");

    var addedFormula = Path.Combine(root, "added-formula.xlsx");
    CreateSyntheticXlsx(addedFormula, recalculatedStyles, recalculatedSheet.Replace("<c r=\"A1\" s=\"99\"><v>999</v></c>", "<c r=\"A1\" s=\"99\"><f>1+1</f><v>2</v></c>", StringComparison.Ordinal));
    RequireThrows(() => XlsxFormulaCacheMerger.Merge(source, addedFormula, Path.Combine(root, "added-formula-output.xlsx")), "formula cell inventory");
}

static void RunLegacyXlsFontRecalculationProbe(string root)
{
    Directory.CreateDirectory(root);
    var legacy = Path.Combine(root, "font-variants.xls");
    var converted = Path.Combine(root, "font-variants.xlsx");
    var recalculated = Path.Combine(root, "font-variants-recalculated.xlsx");
    using (var workbook = new HSSFWorkbook())
    {
        var sheet = workbook.CreateSheet("Font variants");
        var roman = workbook.CreateFont();
        roman.FontName = "Times New Roman";
        roman.FontHeightInPoints = 12;
        roman.IsBold = false;
        roman.Color = HSSFColor.Red.Index;
        SetLegacyFontFamily(workbook, roman, 1);
        var romanStyle = workbook.CreateCellStyle();
        romanStyle.SetFont(roman);

        var swiss = workbook.CreateFont();
        swiss.FontName = "Arial";
        swiss.FontHeightInPoints = 9;
        swiss.IsBold = true;
        swiss.IsItalic = true;
        swiss.Color = HSSFColor.Blue.Index;
        SetLegacyFontFamily(workbook, swiss, 2);
        var swissStyle = workbook.CreateCellStyle();
        swissStyle.SetFont(swiss);

        var first = sheet.CreateRow(0);
        first.CreateCell(0).SetCellValue(21);
        var firstFormula = first.CreateCell(1);
        firstFormula.SetCellFormula("A1*2");
        firstFormula.CellStyle = romanStyle;
        var second = sheet.CreateRow(1);
        second.CreateCell(0).SetCellValue(9);
        var secondFormula = second.CreateCell(1);
        secondFormula.SetCellFormula("A2+3");
        secondFormula.CellStyle = swissStyle;
        using var stream = File.Create(legacy);
        workbook.Write(stream, leaveOpen: false);
    }

    var conversion = WorkbookConverter.ConvertXlsToXlsx(legacy, converted);
    Require(conversion.Backend == "et", "legacy XLS font regression did not use ET conversion");
    Require(XDocument.Parse(ReadPart(converted, "xl/styles.xml")).Descendants()
            .Count(element => element.Name.LocalName == "family") >= 2,
        "legacy XLS conversion did not produce the font-family semantics under regression");
    var recalculation = WorkbookRecalculator.RecalculateXlsx(converted, recalculated);
    Require(recalculation.Backend == "et", "legacy XLS font regression did not use ET recalculation");
    Require(ReadPart(converted, "xl/styles.xml") == ReadPart(recalculated, "xl/styles.xml"),
        "legacy XLS font semantics changed during recalculation");

    using var resultStream = File.OpenRead(recalculated);
    using var result = new XSSFWorkbook(resultStream);
    var resultSheet = result.GetSheet("Font variants");
    Require(resultSheet.GetRow(0).GetCell(1).NumericCellValue == 42, "first legacy formula cache was not materialized");
    Require(resultSheet.GetRow(1).GetCell(1).NumericCellValue == 12, "second legacy formula cache was not materialized");
}

static void SetLegacyFontFamily(HSSFWorkbook workbook, IFont font, byte family)
{
    var internalWorkbookField = typeof(HSSFWorkbook).GetFields(BindingFlags.Instance | BindingFlags.NonPublic)
        .Single(field => field.FieldType.FullName == "NPOI.HSSF.Model.InternalWorkbook");
    var internalWorkbook = internalWorkbookField.GetValue(workbook)
        ?? throw new InvalidOperationException("NPOI legacy workbook internals are unavailable.");
    var fontRecord = internalWorkbook.GetType().GetMethod("GetFontRecordAt")!.Invoke(internalWorkbook, [Convert.ToInt32(font.Index)])
        ?? throw new InvalidOperationException("NPOI legacy font record is unavailable.");
    fontRecord.GetType().GetProperty("Family")!.SetValue(fontRecord, family);
}

static void CreateSyntheticXlsx(string path, string styles, string worksheet)
{
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    Write(archive, "xl/styles.xml", styles);
    Write(archive, "xl/worksheets/sheet1.xml", worksheet);
    Write(archive, "docProps/custom.xml", "<properties><marker>source-authoritative</marker></properties>");
}

static void RequireThrows(Action action, string expectedMessage)
{
    try { action(); }
    catch (InvalidOperationException error) when (error.Message.Contains(expectedMessage, StringComparison.Ordinal)) { return; }
    throw new InvalidOperationException($"Expected failure containing '{expectedMessage}'.");
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
