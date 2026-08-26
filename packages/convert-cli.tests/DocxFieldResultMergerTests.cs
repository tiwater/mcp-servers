using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace Dockit.Convert.Tests;

public sealed class DocxFieldResultMergerTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    [Fact]
    public void Merge_updates_index_results_and_bookmarks_but_preserves_business_body_and_other_parts()
    {
        var source = Package(
            Paragraph("BODY0001", "HSP6142DDS (250810S) source text"),
            Index("table-of-contents", "Old heading", "2", "_TocOld"),
            Heading("HEAD0001", "Source heading", "_TocOld", "7"),
            Heading("KEEP0001", "Unrelated target", "_TocKeep", "8"));
        var refreshed = Package(new[]
        {
            Paragraph("BODY0001", "HSP6142DDS(250810S) source text"),
            Index("table-of-contents", "New heading", "8", "_TocNew"),
            Heading("HEAD0001", "Rewritten heading", "_TocNew", "41"),
            Heading("KEEP0001", "Unrelated target rewritten", "_TocKeep", "42")
        }, "wps-rewritten-settings");
        var output = TemporaryDocx();

        DocxFieldResultMerger.Merge(source, refreshed, output);

        var document = ReadDocument(output);
        Assert.Contains("HSP6142DDS (250810S) source text", document.DescendantNodes().OfType<XText>().Select(text => text.Value));
        Assert.DoesNotContain("HSP6142DDS(250810S) source text", document.DescendantNodes().OfType<XText>().Select(text => text.Value));
        Assert.Contains("New heading", document.Root!.Value);
        Assert.Contains("8", document.Root.Value);
        Assert.Contains("Source heading", document.Root.Value);
        Assert.DoesNotContain("Rewritten heading", document.Root.Value);
        Assert.Contains(document.Descendants(W + "bookmarkStart"), element => (string?)element.Attribute(W + "name") == "_TocNew");
        Assert.DoesNotContain(document.Descendants(W + "bookmarkStart"), element => (string?)element.Attribute(W + "name") == "_TocOld");
        Assert.Contains(document.Descendants(W + "bookmarkStart"), element => (string?)element.Attribute(W + "name") == "_TocKeep");
        Assert.Contains("Unrelated target", document.Root.Value);
        Assert.DoesNotContain("Unrelated target rewritten", document.Root.Value);
        Assert.Equal("source-settings", ReadEntry(output, "word/settings.xml"));
    }

    [Fact]
    public void Merge_supports_more_than_one_semantically_distinct_index()
    {
        var source = Package(
            Index("table-of-contents", "Old contents", "1", "_TocA"),
            Index("table-of-figures", "Old figure", "3", "_TocB"),
            Heading("HEADA001", "Contents heading", "_TocA", "1"),
            Heading("HEADB001", "Figure heading", "_TocB", "2"));
        var refreshed = Package(
            Index("table-of-contents", "New contents", "2", "_TocC"),
            Index("table-of-figures", "New figure", "9", "_TocD"),
            Heading("HEADA001", "Contents heading", "_TocC", "3"),
            Heading("HEADB001", "Figure heading", "_TocD", "4"));
        var output = TemporaryDocx();

        DocxFieldResultMerger.Merge(source, refreshed, output);

        var text = ReadDocument(output).Root!.Value;
        Assert.Contains("New contents", text);
        Assert.Contains("New figure", text);
        Assert.DoesNotContain("Old contents", text);
        Assert.DoesNotContain("Old figure", text);
    }

    [Fact]
    public void Merge_preserves_a_uniform_explicit_body_font_policy_on_refreshed_index_results()
    {
        var sourceIndex = Index("table-of-contents", "Old contents", "1", "_TocA");
        var sourceHeading = Heading("HEAD0001", "Contents heading", "_TocA", "1");
        ApplyFont(sourceIndex, "宋体", "24");
        ApplyFont(sourceHeading, "宋体", "24");
        var refreshedIndex = Index("table-of-contents", "New contents", "8", "_TocB");
        var refreshedHeading = Heading("HEAD0001", "Contents heading", "_TocB", "2");
        ApplyFont(refreshedIndex, "SimSun", "22");
        ApplyFont(refreshedHeading, "SimSun", "22");
        var source = Package(sourceIndex, sourceHeading);
        var refreshed = Package(refreshedIndex, refreshedHeading);
        var output = TemporaryDocx();

        DocxFieldResultMerger.Merge(source, refreshed, output);

        var resultRuns = ReadDocument(output).Descendants(W + "r")
            .Where(run => run.Descendants(W + "t").Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            .ToList();
        Assert.All(resultRuns, run =>
        {
            var properties = run.Element(W + "rPr")!;
            var fonts = properties.Element(W + "rFonts")!;
            Assert.Equal("Times New Roman", (string?)fonts.Attribute(W + "ascii"));
            Assert.Equal("Times New Roman", (string?)fonts.Attribute(W + "hAnsi"));
            Assert.Equal("宋体", (string?)fonts.Attribute(W + "eastAsia"));
            Assert.Equal("Times New Roman", (string?)fonts.Attribute(W + "cs"));
            Assert.Equal("24", (string?)properties.Element(W + "sz")?.Attribute(W + "val"));
            Assert.Equal("24", (string?)properties.Element(W + "szCs")?.Attribute(W + "val"));
        });
    }

    [Fact]
    public void Merge_with_no_index_fields_preserves_the_source_document()
    {
        var source = Package(Paragraph("BODY0001", "source body"));
        var refreshed = Package(Paragraph("BODY0001", "rewritten body"));
        var output = TemporaryDocx();

        DocxFieldResultMerger.Merge(source, refreshed, output);

        Assert.Equal("source body", ReadDocument(output).Root!.Value);
    }

    [Fact]
    public void Merge_rejects_missing_extra_or_reordered_index_fields()
    {
        var one = Package(Index("table-of-contents", "Contents", "1", "_TocA"));
        var two = Package(
            Index("table-of-contents", "Contents", "1", "_TocA"),
            Index("table-of-figures", "Figure", "2", "_TocB"));
        var figure = Package(Index("table-of-figures", "Figure", "2", "_TocB"));

        Assert.Contains("number", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(one, two, TemporaryDocx())).Message);
        Assert.Contains("number", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(two, one, TemporaryDocx())).Message);
        Assert.Contains("order or kind", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(one, figure, TemporaryDocx())).Message);
    }

    [Fact]
    public void Merge_rejects_unbalanced_fields_and_incomplete_bookmarks()
    {
        var malformedField = Package(new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
            new XElement(W + "r", new XElement(W + "instrText", "TOC \\o \"1-3\""))));
        var valid = Package(Index("table-of-contents", "Contents", "1", "_TocA"));
        Assert.Contains("unclosed field", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(malformedField, valid, TemporaryDocx())).Message);

        var source = Package(
            Index("table-of-contents", "Contents", "1", "_TocA"),
            Paragraph("HEAD0001", "Heading"));
        var refreshed = Package(
            Index("table-of-contents", "Contents", "1", "_TocB"),
            new XElement(W + "p", new XAttribute(W14 + "paraId", "HEAD0001"),
                new XElement(W + "bookmarkStart", new XAttribute(W + "id", "9"), new XAttribute(W + "name", "_TocB")),
                new XElement(W + "r", new XElement(W + "t", "Heading"))));
        Assert.Contains("incomplete TOC bookmark", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(source, refreshed, TemporaryDocx())).Message);
    }

    [Fact]
    public void Merge_preserves_a_valid_bookmark_range_that_spans_paragraphs()
    {
        var source = Package(
            Index("table-of-contents", "Old heading", "1", "_TocOld"),
            BookmarkStartParagraph("HEAD0001", "First source paragraph", "_TocOld", "7"),
            BookmarkEndParagraph("HEAD0002", "Second source paragraph", "7"));
        var refreshed = Package(
            Index("table-of-contents", "New heading", "2", "_TocNew"),
            BookmarkStartParagraph("HEAD0001", "First rewritten paragraph", "_TocNew", "19"),
            BookmarkEndParagraph("HEAD0002", "Second rewritten paragraph", "19"));
        var output = TemporaryDocx();

        DocxFieldResultMerger.Merge(source, refreshed, output);

        var document = ReadDocument(output);
        var start = Assert.Single(document.Descendants(W + "bookmarkStart"),
            element => (string?)element.Attribute(W + "name") == "_TocNew");
        var end = Assert.Single(document.Descendants(W + "bookmarkEnd"),
            element => (string?)element.Attribute(W + "id") == (string?)start.Attribute(W + "id"));
        Assert.Equal("HEAD0001", (string?)start.Ancestors(W + "p").Single().Attribute(W14 + "paraId"));
        Assert.Equal("HEAD0002", (string?)end.Ancestors(W + "p").Single().Attribute(W14 + "paraId"));
        Assert.Contains("First source paragraph", document.Root!.Value);
        Assert.Contains("Second source paragraph", document.Root.Value);
    }

    [Fact]
    public void Merge_rejects_ambiguous_paragraph_identity_and_missing_document_part()
    {
        var source = Package(
            Index("table-of-contents", "Contents", "1", "_TocA"),
            Paragraph("DUPL0001", "First"),
            Paragraph("DUPL0001", "Second"));
        var refreshed = Package(
            Index("table-of-contents", "Contents", "1", "_TocB"),
            Heading("DUPL0001", "Heading", "_TocB", "4"));
        Assert.Contains("duplicate paragraph identity", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(source, refreshed, TemporaryDocx())).Message);

        var missing = TemporaryDocx();
        using (var archive = ZipFile.Open(missing, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry("[Content_Types].xml");
            using var writer = new StreamWriter(entry.Open());
            writer.Write("<Types/>");
        }
        Assert.Contains("missing word/document.xml", Assert.Throws<InvalidOperationException>(
            () => DocxFieldResultMerger.Merge(missing, refreshed, TemporaryDocx())).Message);
    }

    private static XElement Paragraph(string id, string text)
        => new(W + "p", new XAttribute(W14 + "paraId", id),
            new XElement(W + "r", new XElement(W + "t", text)));

    private static XElement Heading(string id, string text, string bookmark, string bookmarkId)
        => new(W + "p", new XAttribute(W14 + "paraId", id),
            new XElement(W + "bookmarkStart", new XAttribute(W + "id", bookmarkId), new XAttribute(W + "name", bookmark)),
            new XElement(W + "r", new XElement(W + "t", text)),
            new XElement(W + "bookmarkEnd", new XAttribute(W + "id", bookmarkId)));

    private static XElement BookmarkStartParagraph(string id, string text, string bookmark, string bookmarkId)
        => new(W + "p", new XAttribute(W14 + "paraId", id),
            new XElement(W + "bookmarkStart", new XAttribute(W + "id", bookmarkId), new XAttribute(W + "name", bookmark)),
            new XElement(W + "r", new XElement(W + "t", text)));

    private static XElement BookmarkEndParagraph(string id, string text, string bookmarkId)
        => new(W + "p", new XAttribute(W14 + "paraId", id),
            new XElement(W + "r", new XElement(W + "t", text)),
            new XElement(W + "bookmarkEnd", new XAttribute(W + "id", bookmarkId)));

    private static XElement Index(string kind, string text, string page, string bookmark)
    {
        var instruction = kind == "table-of-figures" ? "TOC \\c \"Figure\" \\h" : "TOC \\o \"1-3\" \\h \\u";
        return new XElement(W + "p",
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
            new XElement(W + "r", new XElement(W + "instrText", instruction)),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
            new XElement(W + "r", new XElement(W + "instrText", $" HYPERLINK \\l {bookmark} ")),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            new XElement(W + "r", new XElement(W + "t", text)),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))),
            new XElement(W + "r", new XElement(W + "t", page)),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
    }

    private static void ApplyFont(XElement element, string eastAsia, string size)
    {
        foreach (var run in element.Descendants(W + "r"))
        {
            run.AddFirst(new XElement(W + "rPr",
                new XElement(W + "rFonts",
                    new XAttribute(W + "ascii", "Times New Roman"),
                    new XAttribute(W + "hAnsi", "Times New Roman"),
                    new XAttribute(W + "eastAsia", eastAsia),
                    new XAttribute(W + "cs", "Times New Roman")),
                new XElement(W + "sz", new XAttribute(W + "val", size)),
                new XElement(W + "szCs", new XAttribute(W + "val", size))));
        }
    }

    private static string Package(XElement first, params XElement[] rest)
        => Package(new[] { first }.Concat(rest).ToArray(), "source-settings");

    private static string Package(XElement first, XElement second, XElement third, string settings)
        => Package(new[] { first, second, third }, settings);

    private static string Package(XElement[] body, string settings)
    {
        var path = TemporaryDocx();
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteEntry(archive, "word/document.xml", new XDocument(
            new XElement(W + "document",
                new XAttribute(XNamespace.Xmlns + "w", W),
                new XAttribute(XNamespace.Xmlns + "w14", W14),
                new XElement(W + "body", body))).ToString(SaveOptions.DisableFormatting));
        WriteEntry(archive, "word/settings.xml", settings);
        WriteEntry(archive, "custom/preserved.bin", "preserved");
        return path;
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    private static XDocument ReadDocument(string path)
        => XDocument.Parse(ReadEntry(path, "word/document.xml"), LoadOptions.PreserveWhitespace);

    private static string ReadEntry(string path, string name)
    {
        using var archive = ZipFile.OpenRead(path);
        using var reader = new StreamReader(archive.GetEntry(name)!.Open());
        return reader.ReadToEnd();
    }

    private static string TemporaryDocx()
        => Path.Combine(Path.GetTempPath(), $"field-merge-{Guid.NewGuid():N}.docx");
}
