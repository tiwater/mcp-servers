using System.IO.Compression;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class ObjectActionTests
{
    [Fact]
    public void Insert_body_range_preserves_an_unseen_section_table_style_media_and_header_footer_topology()
    {
        using var fixture = new Fixture();
        var source = fixture.Path("source.docx");
        var target = fixture.Path("target.docx");
        var output = fixture.Path("output.docx");
        CreateSectionSource(source);
        CreateTarget(target);

        var result = Editor.Apply(target, output, [new DocxEditOperation(
            "insertBodyRange", Source: source, SourceStartBodyIndex: 0, SourceEndBodyIndex: 2, TargetBodyIndex: 0)]);

        Assert.True(Assert.Single(result.AppliedOperations).Applied);
        using var document = WordprocessingDocument.Open(output, false);
        var main = document.MainDocumentPart!;
        var body = main.Document!.Body!;
        Assert.Equal(["styled linked", "section end", "target remains"],
            body.Elements<Paragraph>().Select(Inspector.GetParagraphText));
        Assert.Contains("cell image", Assert.Single(body.Elements<Table>()).InnerText);
        Assert.Equal("SourceTable", Assert.Single(body.Elements<Table>()).GetFirstChild<TableProperties>()!.TableStyle!.Val!.Value);
        Assert.Contains(main.StyleDefinitionsPart!.Styles!.Elements<Style>(), style => style.StyleId == "SourceParagraph");
        Assert.Contains(main.StyleDefinitionsPart.Styles.Elements<Style>(), style => style.StyleId == "SourceTable");
        Assert.Contains(main.NumberingDefinitionsPart!.Numbering!.Elements<NumberingInstance>(), item => item.NumberID?.Value == 42);
        Assert.Equal(new Uri("https://example.test/source"), Assert.Single(main.HyperlinkRelationships).Uri);
        Assert.Single(main.ImageParts);
        Assert.Single(main.HeaderParts);
        Assert.Single(main.FooterParts);
        var insertedSection = body.Elements<Paragraph>().Single(paragraph => paragraph.InnerText == "section end").ParagraphProperties!.SectionProperties!;
        Assert.NotNull(insertedSection.GetFirstChild<HeaderReference>()?.Id?.Value);
        Assert.NotNull(insertedSection.GetFirstChild<FooterReference>()?.Id?.Value);
        Assert.Contains("header image", Assert.Single(main.HeaderParts).Header!.InnerText);
        Assert.Single(Assert.Single(main.HeaderParts).ImageParts);
        Assert.Contains("footer text", Assert.Single(main.FooterParts).Footer!.InnerText);
        var validationErrors = new OpenXmlValidator().Validate(document).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors.Select(error => error.Description)));
    }

    [Fact]
    public void Insert_body_range_copies_one_whole_table_and_its_image_relationship()
    {
        using var fixture = new Fixture();
        var source = fixture.Path("table-source.docx");
        var target = fixture.Path("table-target.docx");
        var output = fixture.Path("table-output.docx");
        CreateTableOnlySource(source);
        CreateTarget(target);

        var result = Editor.Apply(target, output, [new DocxEditOperation(
            "insertBodyRange", Source: source, SourceStartBodyIndex: 0, SourceEndBodyIndex: 0, TargetBodyIndex: 1)]);

        Assert.True(Assert.Single(result.AppliedOperations).Applied);
        using var document = WordprocessingDocument.Open(output, false);
        var table = Assert.Single(document.MainDocumentPart!.Document!.Body!.Elements<Table>());
        Assert.Equal("preserved table", table.InnerText);
        var blip = Assert.Single(table.Descendants<A.Blip>());
        Assert.IsType<ImagePart>(document.MainDocumentPart.GetPartById(blip.Embed!.Value!));
    }

    [Theory]
    [InlineData(1, 2, "whole section")]
    [InlineData(0, 99, "out of range")]
    public void Insert_body_range_fails_closed_for_partial_sections_and_invalid_bounds(int start, int end, string reason)
    {
        using var fixture = new Fixture();
        var source = fixture.Path("source.docx");
        var target = fixture.Path("target.docx");
        var output = fixture.Path("output.docx");
        CreateSectionSource(source);
        CreateTarget(target);

        var result = Editor.Apply(target, output, [new DocxEditOperation(
            "insertBodyRange", Source: source, SourceStartBodyIndex: start, SourceEndBodyIndex: end, TargetBodyIndex: 0)]);

        var applied = Assert.Single(result.AppliedOperations);
        Assert.False(applied.Applied);
        Assert.Contains(reason, applied.Detail, StringComparison.OrdinalIgnoreCase);
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Equal("target remains", document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().InnerText);
    }

    [Fact]
    public void Insert_body_range_fails_closed_when_a_same_id_style_has_different_definition()
    {
        using var fixture = new Fixture();
        var source = fixture.Path("source.docx");
        var target = fixture.Path("conflict.docx");
        var output = fixture.Path("output.docx");
        CreateTableOnlySource(source);
        CreateTarget(target, conflictingTableStyle: true);

        var result = Editor.Apply(target, output, [new DocxEditOperation(
            "insertBodyRange", Source: source, SourceStartBodyIndex: 0, SourceEndBodyIndex: 0, TargetBodyIndex: 0)]);

        var applied = Assert.Single(result.AppliedOperations);
        Assert.False(applied.Applied);
        Assert.Contains("style conflicts", applied.Detail);
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Empty(document.MainDocumentPart!.Document!.Body!.Elements<Table>());
    }

    [Fact]
    public void Insert_and_replace_image_preserve_explicit_geometry_and_do_not_mutate_shared_media()
    {
        using var fixture = new Fixture();
        var target = fixture.Path("target.docx");
        var firstImage = fixture.Image("first.png", 0x11);
        var secondImage = fixture.Image("second.png", 0x22);
        var inserted = fixture.Path("inserted.docx");
        var replaced = fixture.Path("replaced.docx");
        CreateTarget(target);

        var insert = Editor.Apply(target, inserted, [new DocxEditOperation(
            "insertBodyImage", TargetBodyIndex: 0, Image: firstImage, WidthEmu: 1234567, HeightEmu: 765432, AltText: "evidence image")]);
        Assert.True(Assert.Single(insert.AppliedOperations).Applied);
        using (var document = WordprocessingDocument.Open(inserted, false))
        {
            var inline = Assert.Single(document.MainDocumentPart!.Document!.Body!.Descendants<DW.Inline>());
            Assert.Equal(1234567L, inline.Extent!.Cx!.Value);
            Assert.Equal(765432L, inline.Extent.Cy!.Value);
            Assert.Equal("evidence image", inline.DocProperties!.Description!.Value);
        }
        using (var document = WordprocessingDocument.Open(inserted, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            var duplicate = (Drawing)paragraph.Descendants<Drawing>().Single().CloneNode(true);
            duplicate.Descendants<DW.DocProperties>().Single().Id = 99U;
            duplicate.Descendants<PIC.NonVisualDrawingProperties>().Single().Id = 99U;
            paragraph.AppendChild(new Run(duplicate));
            document.MainDocumentPart.Document.Save();
        }

        var replace = Editor.Apply(inserted, replaced, [new DocxEditOperation(
            "replaceDrawingImage", ParagraphIndex: 0, DrawingIndex: 0, Image: secondImage)]);
        Assert.True(Assert.Single(replace.AppliedOperations).Applied);
        using var replacedDocument = WordprocessingDocument.Open(replaced, false);
        var main = replacedDocument.MainDocumentPart!;
        var replacedInline = main.Document!.Body!.Descendants<DW.Inline>().First();
        Assert.Equal(1234567L, replacedInline.Extent!.Cx!.Value);
        var blips = main.Document.Body.Descendants<A.Blip>().ToList();
        Assert.Equal(2, blips.Count);
        using var active = main.GetPartById(blips[0].Embed!.Value!).GetStream();
        using var retained = main.GetPartById(blips[1].Embed!.Value!).GetStream();
        Assert.Equal(File.ReadAllBytes(secondImage), ReadAll(active));
        Assert.Equal(File.ReadAllBytes(firstImage), ReadAll(retained));
        Assert.Equal(2, main.ImageParts.Count());
        Assert.Empty(new OpenXmlValidator().Validate(replacedDocument));
    }

    [Fact]
    public void Image_actions_reject_unsafe_targets_and_unknown_media_without_changing_body_content()
    {
        using var fixture = new Fixture();
        var target = fixture.Path("target.docx");
        var output = fixture.Path("output.docx");
        var unknown = fixture.Path("image.bin");
        File.WriteAllBytes(unknown, [1, 2, 3]);
        CreateTarget(target);

        var result = Editor.Apply(target, output, [
            new DocxEditOperation("replaceDrawingImage", ParagraphIndex: 0, DrawingIndex: 0, Image: unknown),
            new DocxEditOperation("insertBodyImage", TargetBodyIndex: 2, Image: unknown, WidthEmu: 1, HeightEmu: 1)
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.False(operation.Applied));
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Equal("target remains", document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().InnerText);
        Assert.Empty(document.MainDocumentPart.ImageParts);
    }

    [Fact]
    public void Edit_contract_accepts_only_the_new_technical_fields()
    {
        var schema = JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "docx-cli", "contracts", "tiwater.docx-edit-v1.schema.json")));
        var operation = schema.RootElement.GetProperty("$defs").GetProperty("operation");
        var names = operation.GetProperty("properties").EnumerateObject().Select(property => property.Name).ToHashSet();
        foreach (var required in new[] { "source", "sourceStartBodyIndex", "sourceEndBodyIndex", "targetBodyIndex", "image", "drawingIndex", "widthEmu", "heightEmu", "altText" })
            Assert.Contains(required, names);
        Assert.DoesNotContain("scenarioId", names);
        Assert.DoesNotContain("templateId", names);
        Assert.True(operation.GetProperty("additionalProperties").ValueKind == JsonValueKind.False);
    }

    private static void CreateSectionSource(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(
            new Style(
                new StyleName { Val = "Source paragraph" },
                new StyleParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 42 })))
            { Type = StyleValues.Paragraph, StyleId = "SourceParagraph" },
            new Style(new StyleName { Val = "Source table" }) { Type = StyleValues.Table, StyleId = "SourceTable" });
        var numbering = main.AddNewPart<NumberingDefinitionsPart>();
        numbering.Numbering = new Numbering(
            new AbstractNum(
                new Level(new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "%1." }) { LevelIndex = 0 })
            { AbstractNumberId = 7 },
            new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 42 });
        var image = main.AddImagePart(ImagePartType.Png);
        image.FeedData(new MemoryStream(Png(0x33)));
        var hyperlink = main.AddHyperlinkRelationship(new Uri("https://example.test/source"), true);
        var first = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "SourceParagraph" }),
            new Hyperlink(new Run(new Text("styled linked"))) { Id = hyperlink.Id });
        var table = new Table(
            new TableProperties(new TableStyle { Val = "SourceTable" }),
            new TableGrid(new GridColumn { Width = "2400" }),
            new TableRow(new TableCell(new Paragraph(new Run(Drawing(main.GetIdOfPart(image), 1U)), new Run(new Text("cell image"))))));

        var header = main.AddNewPart<HeaderPart>();
        var headerImage = header.AddImagePart(ImagePartType.Png);
        headerImage.FeedData(new MemoryStream(Png(0x44)));
        header.Header = new Header(new Paragraph(new Run(Drawing(header.GetIdOfPart(headerImage), 2U)), new Run(new Text("header image"))));
        var footer = main.AddNewPart<FooterPart>();
        footer.Footer = new Footer(new Paragraph(new Run(new Text("footer text"))));
        var section = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) },
            new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) },
            new PageSize { Width = 16838U, Height = 11906U, Orient = PageOrientationValues.Landscape });
        var boundary = new Paragraph(new ParagraphProperties(section), new Run(new Text("section end")));
        main.Document = new Document(new Body(first, table, boundary, new Paragraph(new Run(new Text("unseen second section"))), new SectionProperties()));
        main.Document.Save();
        header.Header.Save();
        footer.Footer.Save();
        styles.Styles.Save();
        numbering.Numbering.Save();
    }

    private static void CreateTableOnlySource(string path)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var styles = main.AddNewPart<StyleDefinitionsPart>();
        styles.Styles = new Styles(new Style(new StyleName { Val = "Source table" }) { Type = StyleValues.Table, StyleId = "SourceTable" });
        var image = main.AddImagePart(ImagePartType.Png);
        image.FeedData(new MemoryStream(Png(0x55)));
        var table = new Table(
            new TableProperties(new TableStyle { Val = "SourceTable" }),
            new TableGrid(new GridColumn { Width = "2400" }),
            new TableRow(new TableCell(new Paragraph(
                new Run(Drawing(main.GetIdOfPart(image), 1U)),
                new Run(new Text("preserved table"))))));
        main.Document = new Document(new Body(table, new SectionProperties()));
        main.Document.Save();
        styles.Styles.Save();
    }

    private static void CreateTarget(string path, bool conflictingTableStyle = false)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        if (conflictingTableStyle)
        {
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(new Style(new StyleName { Val = "Conflicting table" }) { Type = StyleValues.Table, StyleId = "SourceTable" });
            styles.Styles.Save();
        }
        main.Document = new Document(new Body(new Paragraph(new Run(new Text("target remains"))), new SectionProperties()));
        main.Document.Save();
    }

    private static Drawing Drawing(string relationshipId, uint id)
    {
        var graphicData = new A.GraphicData(
            new PIC.Picture(
                new PIC.NonVisualPictureProperties(new PIC.NonVisualDrawingProperties { Id = id, Name = $"image-{id}" }, new PIC.NonVisualPictureDrawingProperties()),
                new PIC.BlipFill(new A.Blip { Embed = relationshipId }, new A.Stretch(new A.FillRectangle())),
                new PIC.ShapeProperties(new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 990000L, Cy = 990000L }), new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
        return new Drawing(new DW.Inline(
            new DW.Extent { Cx = 990000L, Cy = 990000L },
            new DW.DocProperties { Id = id, Name = $"image-{id}" },
            new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(graphicData)));
    }

    private static byte[] Png(byte marker)
    {
        var bytes = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        bytes[^12] = marker;
        return bytes;
    }

    private static byte[] ReadAll(Stream stream)
    {
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    private sealed class Fixture : IDisposable
    {
        private readonly string root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"docx-object-actions-{Guid.NewGuid():N}");
        internal Fixture() => Directory.CreateDirectory(root);
        internal string Path(string name) => System.IO.Path.Combine(root, name);
        internal string Image(string name, byte marker) { var path = Path(name); File.WriteAllBytes(path, Png(marker)); return path; }
        public void Dispose() => Directory.Delete(root, recursive: true);
    }
}
