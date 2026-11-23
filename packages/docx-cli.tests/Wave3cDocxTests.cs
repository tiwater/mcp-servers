using System.Buffers.Binary;
using System.IO.Compression;
using System.IO.Packaging;
using System.Xml.Linq;
using Dockit.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class Wave3cDocxTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"tiwater-docx-wave3c-{Guid.NewGuid():N}");

    public Wave3cDocxTests()
    {
        Directory.CreateDirectory(_root);
    }

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    // -------------------------------------------------------------------------
    // ANA08-F1: docx_replace_content_from_source numbering preservation
    // -------------------------------------------------------------------------

    [Fact]
    public void ReplaceContentFromSource_PreservesBilingualAlternatingNumbering()
    {
        var sourcePath = Path.Combine(_root, "ana08-f1-source.docx");
        var targetPath = Path.Combine(_root, "ana08-f1-target.docx");
        var outputPath = Path.Combine(_root, "ana08-f1-output.docx");
        var receiptPath = Path.Combine(_root, "ana08-f1-receipt.json");

        CreateBilingualNumberedCellDoc(sourcePath, numId: 2, pairCount: 6, textPrefix: "Source", explicitZeroOnUnnumbered: true);
        CreateBilingualNumberedCellDoc(targetPath, numId: 4, pairCount: 6, textPrefix: "Baseline", explicitZeroOnUnnumbered: true);

        var cellAddress = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]");
        NativeContentCopy.Apply(new CopyContentRequest(
            targetPath,
            [new CopyContentChange(cellAddress, sourcePath, [new CopyContentSelection(cellAddress)])],
            outputPath,
            receiptPath));

        using var doc = WordprocessingDocument.Open(outputPath, false);
        AssertNoValidationErrors(doc);
        var paragraphs = doc.MainDocumentPart!.Document.Body!
            .Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().Single()
            .Elements<Paragraph>().ToArray();

        Assert.Equal(12, paragraphs.Length);
        for (var i = 0; i < 12; i++)
        {
            var isZh = i % 2 == 0;
            var numId = paragraphs[i].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value;
            if (isZh)
            {
                Assert.Equal(4, numId);
                Assert.Equal("425", paragraphs[i].ParagraphProperties?.Indentation?.Left?.Value);
                Assert.Equal("425", paragraphs[i].ParagraphProperties?.Indentation?.Hanging?.Value);
                Assert.StartsWith("Source ZH ", paragraphs[i].InnerText);
            }
            else
            {
                Assert.True(numId is null or 0, $"EN paragraph {i} unexpectedly inherited numId={numId}");
                Assert.Equal("227", paragraphs[i].ParagraphProperties?.Indentation?.Left?.Value);
                Assert.Null(paragraphs[i].ParagraphProperties?.Indentation?.Hanging?.Value);
                Assert.StartsWith("Source EN ", paragraphs[i].InnerText);
            }
        }
    }

    [Fact]
    public void ReplaceContentFromSource_UnseenExpandedSourceWithoutNumPr_DoesNotInheritTargetNumbering()
    {
        var sourcePath = Path.Combine(_root, "ana08-f1-expanded-source.docx");
        var targetPath = Path.Combine(_root, "ana08-f1-short-target.docx");
        var outputPath = Path.Combine(_root, "ana08-f1-expanded-output.docx");
        var receiptPath = Path.Combine(_root, "ana08-f1-expanded-receipt.json");

        // Unseen synthetic variant: source has 8 bilingual pairs (16 paragraphs) where EN paragraphs omit w:numPr entirely,
        // while target only has 2 pairs (4 paragraphs).
        CreateBilingualNumberedCellDoc(sourcePath, numId: 9, pairCount: 8, textPrefix: "Expanded", explicitZeroOnUnnumbered: false);
        CreateBilingualNumberedCellDoc(targetPath, numId: 4, pairCount: 2, textPrefix: "Template", explicitZeroOnUnnumbered: true);

        var cellAddress = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]");
        NativeContentCopy.Apply(new CopyContentRequest(
            targetPath,
            [new CopyContentChange(cellAddress, sourcePath, [new CopyContentSelection(cellAddress)])],
            outputPath,
            receiptPath));

        using var doc = WordprocessingDocument.Open(outputPath, false);
        AssertNoValidationErrors(doc);
        var paragraphs = doc.MainDocumentPart!.Document.Body!
            .Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().Single()
            .Elements<Paragraph>().ToArray();

        Assert.Equal(16, paragraphs.Length);
        var numberedCount = 0;
        for (var i = 0; i < 16; i++)
        {
            var effective = NativeContentCopy.ResolveEffectiveNumbering(paragraphs[i], doc.MainDocumentPart.StyleDefinitionsPart?.Styles);
            if (i % 2 == 0)
            {
                Assert.True(effective.IsNumbered);
                Assert.Equal(4, paragraphs[i].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
                numberedCount++;
            }
            else
            {
                Assert.False(effective.IsNumbered);
                Assert.Equal("227", paragraphs[i].ParagraphProperties?.Indentation?.Left?.Value);
            }
        }
        Assert.Equal(8, numberedCount);
    }

    [Fact]
    public void ReplaceContentFromSource_SingleNumberedTargetTemplate_SuppressesNumberingForUnnumberedSourceParagraphs()
    {
        var sourcePath = Path.Combine(_root, "ana08-f1-single-tpl-source.docx");
        var targetPath = Path.Combine(_root, "ana08-f1-single-tpl-target.docx");
        var outputPath = Path.Combine(_root, "ana08-f1-single-tpl-output.docx");
        var receiptPath = Path.Combine(_root, "ana08-f1-single-tpl-receipt.json");

        CreateBilingualNumberedCellDoc(sourcePath, numId: 2, pairCount: 3, textPrefix: "Mix", explicitZeroOnUnnumbered: false);

        // Target has ONLY one numbered paragraph (no unnumbered paragraph template in the cell at all).
        using (var target = WordprocessingDocument.Create(targetPath, WordprocessingDocumentType.Document))
        {
            var main = target.AddMainDocumentPart();
            AddSimpleNumbering(main, 4);
            var cell = new TableCell(
                new Paragraph(
                    new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = 0 },
                            new NumberingId { Val = 4 }),
                        new Indentation { Left = "425", Hanging = "425" }),
                    new Run(new Text("Only numbered template"))));
            main.Document = new Document(new Body(CreateValidTable(cell)));
            main.Document.Save();
        }

        var cellAddress = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]");
        NativeContentCopy.Apply(new CopyContentRequest(
            targetPath,
            [new CopyContentChange(cellAddress, sourcePath, [new CopyContentSelection(cellAddress)])],
            outputPath,
            receiptPath));

        using var doc = WordprocessingDocument.Open(outputPath, false);
        AssertNoValidationErrors(doc);
        var paragraphs = doc.MainDocumentPart!.Document.Body!
            .Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().Single()
            .Elements<Paragraph>().ToArray();

        Assert.Equal(6, paragraphs.Length);
        for (var i = 0; i < 6; i++)
        {
            var effective = NativeContentCopy.ResolveEffectiveNumbering(paragraphs[i], doc.MainDocumentPart.StyleDefinitionsPart?.Styles);
            if (i % 2 == 0)
            {
                Assert.True(effective.IsNumbered);
                Assert.Equal(4, paragraphs[i].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
                Assert.Equal("425", paragraphs[i].ParagraphProperties?.Indentation?.Hanging?.Value);
            }
            else
            {
                Assert.False(effective.IsNumbered);
                Assert.Null(paragraphs[i].ParagraphProperties?.NumberingProperties);
                Assert.Null(paragraphs[i].ParagraphProperties?.Indentation?.Hanging?.Value);
            }
        }
    }

    [Fact]
    public void ReplaceContentFromSource_StyleInheritedNumbering_WritesExplicitNumIdZeroForUnnumberedSource()
    {
        var sourcePath = Path.Combine(_root, "ana08-f1-style-source.docx");
        var targetPath = Path.Combine(_root, "ana08-f1-style-target.docx");
        var outputPath = Path.Combine(_root, "ana08-f1-style-output.docx");
        var receiptPath = Path.Combine(_root, "ana08-f1-style-receipt.json");

        CreateBilingualNumberedCellDoc(sourcePath, numId: 2, pairCount: 2, textPrefix: "Styled", explicitZeroOnUnnumbered: false);

        using (var target = WordprocessingDocument.Create(targetPath, WordprocessingDocumentType.Document))
        {
            var main = target.AddMainDocumentPart();
            AddSimpleNumbering(main, 4);
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new Style(
                    new StyleName { Val = "NumberedItem" },
                    new StyleParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = 0 },
                            new NumberingId { Val = 4 })))
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "NumberedItem",
                });
            stylesPart.Styles.Save();

            var cell = new TableCell(
                new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = "NumberedItem" }),
                    new Run(new Text("Style numbered template"))));
            main.Document = new Document(new Body(CreateValidTable(cell)));
            main.Document.Save();
        }

        var cellAddress = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]");
        NativeContentCopy.Apply(new CopyContentRequest(
            targetPath,
            [new CopyContentChange(cellAddress, sourcePath, [new CopyContentSelection(cellAddress)])],
            outputPath,
            receiptPath));

        using var doc = WordprocessingDocument.Open(outputPath, false);
        AssertNoValidationErrors(doc);
        var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles;
        var paragraphs = doc.MainDocumentPart.Document.Body!
            .Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().Single()
            .Elements<Paragraph>().ToArray();

        Assert.Equal(4, paragraphs.Length);
        Assert.True(NativeContentCopy.ResolveEffectiveNumbering(paragraphs[0], styles).IsNumbered);
        Assert.False(NativeContentCopy.ResolveEffectiveNumbering(paragraphs[1], styles).IsNumbered);
        Assert.Equal(0, paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
        Assert.True(NativeContentCopy.ResolveEffectiveNumbering(paragraphs[2], styles).IsNumbered);
        Assert.False(NativeContentCopy.ResolveEffectiveNumbering(paragraphs[3], styles).IsNumbered);
        Assert.Equal(0, paragraphs[3].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
    }

    // -------------------------------------------------------------------------
    // ANA08-F6: docx_set_drawing_checkbox_state standard word/media/ & deterministic rIds
    // -------------------------------------------------------------------------

    [Fact]
    public void SetDrawingCheckboxState_UsesStandardWordMediaRelativeTargetAndDeterministicRelIds()
    {
        var inputPath = Path.Combine(_root, "ana08-f6-input.docx");
        var output1 = Path.Combine(_root, "ana08-f6-output-1.docx");
        var output2 = Path.Combine(_root, "ana08-f6-output-2.docx");
        var receipt1 = Path.Combine(_root, "ana08-f6-receipt-1.json");
        var receipt2 = Path.Combine(_root, "ana08-f6-receipt-2.json");

        CreateStandardMediaCheckboxDoc(inputPath, checkboxCount: 3);

        var drawing1 = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]");
        var drawing2 = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[2]/w:r[1]/w:drawing[1]");

        NativeDrawingCheckboxMutation.Apply(new SetDrawingCheckboxStateRequest(
            inputPath,
            [
                new SetDrawingCheckboxStateChange(drawing1, true),
                new SetDrawingCheckboxStateChange(drawing2, true),
            ],
            output1,
            receipt1));

        NativeDrawingCheckboxMutation.Apply(new SetDrawingCheckboxStateRequest(
            inputPath,
            [
                new SetDrawingCheckboxStateChange(drawing1, true),
                new SetDrawingCheckboxStateChange(drawing2, true),
            ],
            output2,
            receipt2));

        using (var doc = WordprocessingDocument.Open(output1, false))
        {
            AssertNoValidationErrors(doc);
            var drawings = doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>().ToArray();
            Assert.Equal(3, drawings.Length);
            Assert.Equal(true, NativeDrawingCheckboxMutation.ObserveCheckboxState(doc, "/word/document.xml", drawings[0]));
            Assert.Equal(true, NativeDrawingCheckboxMutation.ObserveCheckboxState(doc, "/word/document.xml", drawings[1]));
            Assert.Equal(false, NativeDrawingCheckboxMutation.ObserveCheckboxState(doc, "/word/document.xml", drawings[2]));

            // Both checked drawings share the same deterministic deduplicated rId!
            var embed1 = drawings[0].Descendants<A.Blip>().Single().Embed!.Value!;
            var embed2 = drawings[1].Descendants<A.Blip>().Single().Embed!.Value!;
            Assert.Equal(embed1, embed2);
            Assert.Matches(@"^rId\d+$", embed1);
        }

        // Verify ZIP entries and document.xml.rels directly
        using (var zip1 = ZipFile.OpenRead(output1))
        using (var zip2 = ZipFile.OpenRead(output2))
        {
            Assert.DoesNotContain(zip1.Entries, e => e.FullName.StartsWith("media/", StringComparison.OrdinalIgnoreCase));
            var mediaEntries1 = zip1.Entries
                .Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                .Select(e => e.FullName)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToArray();
            var mediaEntries2 = zip2.Entries
                .Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
                .Select(e => e.FullName)
                .OrderBy(x => x, StringComparer.Ordinal)
                .ToArray();

            Assert.Equal(["word/media/image1.wmf", "word/media/image2.wmf"], mediaEntries1);
            Assert.Equal(mediaEntries1, mediaEntries2);

            var rels1 = XDocument.Load(zip1.GetEntry("word/_rels/document.xml.rels")!.Open());
            var rels2 = XDocument.Load(zip2.GetEntry("word/_rels/document.xml.rels")!.Open());
            Assert.Equal(rels1.ToString(SaveOptions.DisableFormatting), rels2.ToString(SaveOptions.DisableFormatting));

            var imageRels = rels1.Root!.Elements()
                .Where(e => ((string?)e.Attribute("Type"))?.EndsWith("/image", StringComparison.Ordinal) == true)
                .ToArray();
            Assert.All(imageRels, rel =>
            {
                var id = (string)rel.Attribute("Id")!;
                var target = (string)rel.Attribute("Target")!;
                Assert.Matches(@"^rId\d+$", id);
                Assert.StartsWith("media/", target, StringComparison.Ordinal);
                Assert.DoesNotContain("..", target, StringComparison.Ordinal);
            });
        }
    }

    [Fact]
    public void SetDrawingCheckboxState_InHeaderPart_StoresUnderWordMediaWithRelativeTarget()
    {
        var inputPath = Path.Combine(_root, "ana08-f6-header-input.docx");
        var outputPath = Path.Combine(_root, "ana08-f6-header-output.docx");
        var receiptPath = Path.Combine(_root, "ana08-f6-header-receipt.json");

        CreateStandardMediaHeaderCheckboxDoc(inputPath);
        var headerDrawing = new DocxObjectAddress("/word/header1.xml", "/w:hdr[1]/w:p[1]/w:r[1]/w:drawing[1]");

        NativeDrawingCheckboxMutation.Apply(new SetDrawingCheckboxStateRequest(
            inputPath,
            [new SetDrawingCheckboxStateChange(headerDrawing, true)],
            outputPath,
            receiptPath));

        using (var doc = WordprocessingDocument.Open(outputPath, false))
        {
            AssertNoValidationErrors(doc);
            var headerPart = doc.MainDocumentPart!.HeaderParts.Single();
            var drawing = headerPart.Header!.Descendants<Drawing>().Single();
            Assert.Equal(true, NativeDrawingCheckboxMutation.ObserveCheckboxState(doc, "/word/header1.xml", drawing));
            Assert.Matches(@"^rId\d+$", drawing.Descendants<A.Blip>().Single().Embed!.Value!);
        }

        using var zip = ZipFile.OpenRead(outputPath);
        Assert.DoesNotContain(zip.Entries, e => e.FullName.StartsWith("media/", StringComparison.OrdinalIgnoreCase));
        var headerRels = XDocument.Load(zip.GetEntry("word/_rels/header1.xml.rels")!.Open());
        var targets = headerRels.Root!.Elements().Select(e => (string)e.Attribute("Target")!).ToArray();
        Assert.All(targets, t => Assert.StartsWith("media/", t, StringComparison.Ordinal));
    }

    // -------------------------------------------------------------------------
    // ANA03-F09: docx_delete_object garbage-collects unreferenced word/media/*
    // -------------------------------------------------------------------------

    [Fact]
    public void DeleteObject_GarbageCollectsUnreferencedWordMediaAndShrinksFileSize()
    {
        var inputPath = Path.Combine(_root, "ana03-f09-input.docx");
        var outputPath = Path.Combine(_root, "ana03-f09-output.docx");
        var receiptPath = Path.Combine(_root, "ana03-f09-receipt.json");

        CreateMultiImageDocument(inputPath, imagePayloadBytes: 150_000);
        var inputSize = new FileInfo(inputPath).Length;
        Assert.True(inputSize > 500_000, $"Expected large input DOCX (>500KB), got {inputSize}");

        // Delete drawings 1, 2, 3 (each referencing unique big images); keep drawing 4 (shared with header) and drawing 5 (referencing image4).
        NativeObjectMutation.Delete(new DeleteObjectRequest(
            inputPath,
            [
                new DeleteObjectChange([
                    new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]"),
                    new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[2]/w:r[1]/w:drawing[1]"),
                    new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[3]/w:r[1]/w:drawing[1]"),
                    // Also delete the body drawing that shares header-logo.png with header1.xml!
                    new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[4]/w:r[1]/w:drawing[1]"),
                ])
            ],
            outputPath,
            receiptPath));

        using (var doc = WordprocessingDocument.Open(outputPath, false))
        {
            AssertNoValidationErrors(doc);
        }

        var outputSize = new FileInfo(outputPath).Length;
        Assert.True(outputSize < inputSize / 2, $"Expected significant size reduction: input={inputSize}, output={outputSize}");

        using var zip = ZipFile.OpenRead(outputPath);
        var remainingMedia = zip.Entries
            .Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .OrderBy(x => x, StringComparer.Ordinal)
            .ToArray();

        // Only header-logo.png (still referenced by word/header1.xml) and kept-body.png (referenced by p[5]) remain!
        // big1.png, big2.png, big3.png, and already-orphaned orphan.png MUST all be garbage-collected!
        Assert.Equal(["word/media/header-logo.png", "word/media/kept-body.png"], remainingMedia);

        var docRels = XDocument.Load(zip.GetEntry("word/_rels/document.xml.rels")!.Open());
        var docMediaTargets = docRels.Root!.Elements()
            .Where(e => ((string?)e.Attribute("Type"))?.EndsWith("/image", StringComparison.Ordinal) == true)
            .Select(e => (string)e.Attribute("Target")!)
            .ToArray();
        Assert.Equal(["media/kept-body.png"], docMediaTargets);
    }

    [Fact]
    public void DeleteObject_DeletingOneOfTwoDrawingsSharingMedia_PreservesSharedMediaPart()
    {
        var inputPath = Path.Combine(_root, "ana03-f09-shared-input.docx");
        var outputPath = Path.Combine(_root, "ana03-f09-shared-output.docx");
        var receiptPath = Path.Combine(_root, "ana03-f09-shared-receipt.json");

        CreateSharedImageDocument(inputPath);

        // Delete p[1] drawing (uses rId10 -> media/shared-a.png, shared with p[2] using same rId10)
        // and p[3] drawing (uses rId11 -> media/shared-b.png, shared with p[4] using separate rId12 -> media/shared-b.png).
        NativeObjectMutation.Delete(new DeleteObjectRequest(
            inputPath,
            [
                new DeleteObjectChange([
                    new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:drawing[1]"),
                    new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[3]/w:r[1]/w:drawing[1]"),
                ])
            ],
            outputPath,
            receiptPath));

        using (var doc = WordprocessingDocument.Open(outputPath, false))
        {
            AssertNoValidationErrors(doc);
        }

        using var zip = ZipFile.OpenRead(outputPath);
        var remainingMedia = zip.Entries
            .Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase))
            .Select(e => e.FullName)
            .OrderBy(x => x, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(["word/media/shared-a.png", "word/media/shared-b.png"], remainingMedia);

        var docRels = XDocument.Load(zip.GetEntry("word/_rels/document.xml.rels")!.Open());
        var relIds = docRels.Root!.Elements().Select(e => (string)e.Attribute("Id")!).ToHashSet(StringComparer.Ordinal);
        Assert.Contains("rId10", relIds);
        Assert.DoesNotContain("rId11", relIds);
        Assert.Contains("rId12", relIds);
    }

    // -------------------------------------------------------------------------
    // 0.20.58: docx_insert_objects numId=0 sentinel & Body append before w:sectPr
    // -------------------------------------------------------------------------

    [Fact]
    public void InsertObjects_SourceParagraphWithExplicitNumIdZero_SucceedsWithAndWithoutNumberingPart()
    {
        foreach (var withNumberingPart in new[] { false, true })
        {
            var tag = withNumberingPart ? "with-part" : "no-part";
            var inputPath = Path.Combine(_root, $"numid0-{tag}-input.docx");
            var outputPath = Path.Combine(_root, $"numid0-{tag}-output.docx");
            var receiptPath = Path.Combine(_root, $"numid0-{tag}-receipt.json");

            using (var doc = WordprocessingDocument.Create(inputPath, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                if (withNumberingPart) AddSimpleNumbering(main, 3);
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(new NumberingProperties(new NumberingId { Val = 0 })),
                        new Run(new Text("11. Relevant Files"))),
                    new Paragraph(new Run(new Text("Anchor"))),
                    new SectionProperties()));
                main.Document.Save();
            }

            var sourceHeading = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[1]");
            var anchor = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[2]");
            var bodyParent = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]");

            var receipt = NativeObjectMutation.Insert(new InsertObjectsRequest(
                inputPath,
                [new InsertObjectsChange(inputPath, [sourceHeading], bodyParent, anchor, Repeat: 2)],
                outputPath,
                receiptPath));

            Assert.Equal(2, receipt.Changes.Count);
            using var outDoc = WordprocessingDocument.Open(outputPath, false);
            AssertNoValidationErrors(outDoc);
            var bodyParagraphs = outDoc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToArray();
            Assert.Equal(4, bodyParagraphs.Length);
            Assert.Equal(0, bodyParagraphs[1].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
            Assert.Equal(0, bodyParagraphs[2].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
            Assert.Equal("Anchor", bodyParagraphs[3].InnerText);
        }
    }

    [Fact]
    public void InsertObjects_AppendToBodyWithBeforeNull_InsertsBeforeTrailingSectionPropertiesInOrder()
    {
        var inputPath = Path.Combine(_root, "append-sectpr-input.docx");
        var outputPath = Path.Combine(_root, "append-sectpr-output.docx");
        var receiptPath = Path.Combine(_root, "append-sectpr-receipt.json");

        using (var doc = WordprocessingDocument.Create(inputPath, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            AddSimpleNumbering(main, 5);
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Body prototype"))),
                new Paragraph(
                    new ParagraphProperties(new NumberingProperties(new NumberingId { Val = 0 })),
                    new Run(new Text("11. Relevant Files"))),
                new SectionProperties(new PageSize { Width = 11906U, Height = 16838U })));
            main.Document.Save();
        }

        var prototype = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[1]");
        var heading = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]/w:p[2]");
        var bodyParent = new DocxObjectAddress("/word/document.xml", "/w:document[1]/w:body[1]");

        // Append 3 body paragraphs after the terminal heading (before: null), including one with numId=0.
        var receipt = NativeObjectMutation.Insert(new InsertObjectsRequest(
            inputPath,
            [
                new InsertObjectsChange(inputPath, [prototype], bodyParent, Before: null, Repeat: 2),
                new InsertObjectsChange(inputPath, [heading], bodyParent, Before: null, Repeat: 1),
            ],
            outputPath,
            receiptPath));

        Assert.Equal(3, receipt.Changes.Count);
        Assert.Equal("/w:document[1]/w:body[1]/w:p[3]", receipt.Changes[0].Address.Path);
        Assert.Equal("/w:document[1]/w:body[1]/w:p[4]", receipt.Changes[1].Address.Path);
        Assert.Equal("/w:document[1]/w:body[1]/w:p[5]", receipt.Changes[2].Address.Path);

        using var outDoc = WordprocessingDocument.Open(outputPath, false);
        AssertNoValidationErrors(outDoc);
        var children = outDoc.MainDocumentPart!.Document.Body!.ChildElements.ToArray();
        Assert.Equal(6, children.Length);
        Assert.IsType<SectionProperties>(children[^1]);
        var paragraphs = children.Take(5).Cast<Paragraph>().Select(p => p.InnerText).ToArray();
        Assert.Equal(
            ["Body prototype", "11. Relevant Files", "Body prototype", "Body prototype", "11. Relevant Files"],
            paragraphs);
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static void AssertNoValidationErrors(WordprocessingDocument document)
    {
        var errors = new OpenXmlValidator().Validate(document).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Id}: {e.Description} ({e.Path?.XPath})")));
    }

    private static void CreateBilingualNumberedCellDoc(
        string path,
        int numId,
        int pairCount,
        string textPrefix,
        bool explicitZeroOnUnnumbered)
    {
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        AddSimpleNumbering(main, numId);

        var cell = new TableCell();
        for (var i = 1; i <= pairCount; i++)
        {
            cell.Append(new Paragraph(
                new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = numId }),
                    new Indentation { Left = "425", Hanging = "425" }),
                new Run(new Text($"{textPrefix} ZH {i}."))));

            var enPPr = new ParagraphProperties();
            if (explicitZeroOnUnnumbered)
            {
                enPPr.Append(new NumberingProperties(new NumberingId { Val = 0 }));
            }
            enPPr.Append(new Indentation { Left = "227" });
            cell.Append(new Paragraph(
                enPPr,
                new Run(new Text($"{textPrefix} EN {i}."))));
        }

        main.Document = new Document(new Body(CreateValidTable(cell)));
        main.Document.Save();
    }

    private static Table CreateValidTable(TableCell cell)
        => new(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "5000" }),
            new TableRow(cell));

    private static void AddSimpleNumbering(MainDocumentPart main, int numId)
    {
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left })
                { LevelIndex = 0 })
            { AbstractNumberId = numId },
            new NumberingInstance(new AbstractNumId { Val = numId }) { NumberID = numId });
        numberingPart.Numbering.Save();
    }

    private static void CreateStandardMediaCheckboxDoc(string path, int checkboxCount)
    {
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body();
            for (var i = 1; i <= checkboxCount; i++)
            {
                body.Append(new Paragraph(
                    new Run(CreateDrawingElement("rId1", (uint)i, $"Checkbox {i}")),
                    new Run(new Text($" Option {i}"))));
            }
            main.Document = new Document(body);
            main.Document.Save();
        }

        using var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        var mediaUri = new Uri("/word/media/image1.wmf", UriKind.Relative);
        var mediaPart = package.CreatePart(mediaUri, "image/x-wmf", CompressionOption.Normal);
        var bytes = UncheckedCheckboxWmf();
        using (var stream = mediaPart.GetStream(FileMode.Create, FileAccess.Write))
            stream.Write(bytes, 0, bytes.Length);
        var mainPart = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
        mainPart.CreateRelationship(new Uri("media/image1.wmf", UriKind.Relative), TargetMode.Internal,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "rId1");
        package.Flush();
    }

    private static void CreateStandardMediaHeaderCheckboxDoc(string path)
    {
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var headerPart = main.AddNewPart<HeaderPart>("rId1");
            headerPart.Header = new Header(
                new Paragraph(
                    new Run(CreateDrawingElement("rId1", 10, "Header Checkbox")),
                    new Run(new Text(" Header option"))));
            headerPart.Header.Save();
            main.Document = new Document(
                new Body(
                    new Paragraph(new Run(new Text("Body text"))),
                    new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rId1" })));
            main.Document.Save();
        }

        using var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        var mediaUri = new Uri("/word/media/image1.wmf", UriKind.Relative);
        var mediaPart = package.CreatePart(mediaUri, "image/x-wmf", CompressionOption.Normal);
        var bytes = UncheckedCheckboxWmf();
        using (var stream = mediaPart.GetStream(FileMode.Create, FileAccess.Write))
            stream.Write(bytes, 0, bytes.Length);
        var headerPackagePart = package.GetPart(new Uri("/word/header1.xml", UriKind.Relative));
        headerPackagePart.CreateRelationship(new Uri("media/image1.wmf", UriKind.Relative), TargetMode.Internal,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "rId1");
        package.Flush();
    }

    private static void CreateMultiImageDocument(string path, int imagePayloadBytes)
    {
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var headerPart = main.AddNewPart<HeaderPart>("rId1");
            headerPart.Header = new Header(
                new Paragraph(new Run(CreateDrawingElement("rId1", 100, "Header Logo"))));
            headerPart.Header.Save();

            main.Document = new Document(new Body(
                new Paragraph(new Run(CreateDrawingElement("rId10", 1, "Big 1"))),
                new Paragraph(new Run(CreateDrawingElement("rId11", 2, "Big 2"))),
                new Paragraph(new Run(CreateDrawingElement("rId12", 3, "Big 3"))),
                new Paragraph(new Run(CreateDrawingElement("rId13", 4, "Body Copy Of Header Logo"))),
                new Paragraph(new Run(CreateDrawingElement("rId14", 5, "Kept Body Image"))),
                new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = "rId1" })));
            main.Document.Save();
        }

        var rng = new Random(42);
        byte[] RandomBytes(int size)
        {
            var buffer = new byte[size];
            rng.NextBytes(buffer);
            return buffer;
        }

        using var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        var mainPart = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
        var headerPackagePart = package.GetPart(new Uri("/word/header1.xml", UriKind.Relative));
        const string imageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        void AddMedia(string name, int size, PackagePart? owner, string? relId)
        {
            var uri = new Uri($"/word/media/{name}", UriKind.Relative);
            var part = package.CreatePart(uri, "image/png", CompressionOption.Normal);
            var payload = RandomBytes(size);
            using (var stream = part.GetStream(FileMode.Create, FileAccess.Write))
                stream.Write(payload, 0, payload.Length);
            if (owner is not null && relId is not null)
                owner.CreateRelationship(new Uri($"media/{name}", UriKind.Relative), TargetMode.Internal, imageRel, relId);
        }

        AddMedia("big1.png", imagePayloadBytes, mainPart, "rId10");
        AddMedia("big2.png", imagePayloadBytes, mainPart, "rId11");
        AddMedia("big3.png", imagePayloadBytes, mainPart, "rId12");
        AddMedia("orphan.png", imagePayloadBytes, null, null);
        AddMedia("header-logo.png", 1024, mainPart, "rId13");
        headerPackagePart.CreateRelationship(new Uri("media/header-logo.png", UriKind.Relative), TargetMode.Internal, imageRel, "rId1");
        AddMedia("kept-body.png", 1024, mainPart, "rId14");
        package.Flush();
    }

    private static void CreateSharedImageDocument(string path)
    {
        using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(CreateDrawingElement("rId10", 1, "Shared A 1"))),
                new Paragraph(new Run(CreateDrawingElement("rId10", 2, "Shared A 2"))),
                new Paragraph(new Run(CreateDrawingElement("rId11", 3, "Shared B 1"))),
                new Paragraph(new Run(CreateDrawingElement("rId12", 4, "Shared B 2")))));
            main.Document.Save();
        }

        using var package = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
        var mainPart = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
        const string imageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        foreach (var name in new[] { "shared-a.png", "shared-b.png" })
        {
            var part = package.CreatePart(new Uri($"/word/media/{name}", UriKind.Relative), "image/png", CompressionOption.Normal);
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            stream.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        }
        mainPart.CreateRelationship(new Uri("media/shared-a.png", UriKind.Relative), TargetMode.Internal, imageRel, "rId10");
        mainPart.CreateRelationship(new Uri("media/shared-b.png", UriKind.Relative), TargetMode.Internal, imageRel, "rId11");
        mainPart.CreateRelationship(new Uri("media/shared-b.png", UriKind.Relative), TargetMode.Internal, imageRel, "rId12");
        package.Flush();
    }

    private static Drawing CreateDrawingElement(string embedRelId, uint id, string name)
        => new(
            new DW.Inline(
                new DW.Extent { Cx = 190500L, Cy = 190500L },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.DocProperties { Id = id, Name = name },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = id + 1000U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = embedRelId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = 190500L, Cy = 190500L }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
            });

    private static byte[] UncheckedCheckboxWmf()
    {
        const int width = 16;
        const int height = 16;
        var stride = ((width * 24 + 31) / 32) * 4;
        var dib = new byte[40 + stride * height];
        BinaryPrimitives.WriteUInt32LittleEndian(dib.AsSpan(0, 4), 40);
        BinaryPrimitives.WriteInt32LittleEndian(dib.AsSpan(4, 4), width);
        BinaryPrimitives.WriteInt32LittleEndian(dib.AsSpan(8, 4), height);
        BinaryPrimitives.WriteUInt16LittleEndian(dib.AsSpan(12, 2), 1);
        BinaryPrimitives.WriteUInt16LittleEndian(dib.AsSpan(14, 2), 24);
        BinaryPrimitives.WriteUInt32LittleEndian(dib.AsSpan(16, 4), 0);

        for (var y = 0; y < height; y++)
        for (var x = 0; x < width; x++)
        {
            var isBorder = x < 2 || y < 2 || x >= width - 2 || y >= height - 2;
            byte color = isBorder ? (byte)0 : (byte)255;
            var offset = 40 + (height - 1 - y) * stride + x * 3;
            dib[offset] = dib[offset + 1] = dib[offset + 2] = color;
        }
        return dib;
    }
}
