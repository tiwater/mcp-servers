using System.IO.Compression;
using System.Text;
using Dockit.Convert;

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

static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
