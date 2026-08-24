using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Dockit.Docx.Tests;

public sealed class RowPaginationTests
{
    [Fact]
    public void Trailing_empty_body_paragraphs_are_counted_and_removed_with_an_exact_precondition()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-trailing-empty-paragraphs-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("approval block"))),
                    new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "120" })),
                    new Paragraph(new Run(new RunProperties(new FontSize { Val = "24" }))),
                    new SectionProperties()));
                main.Document.Save();
            }

            Assert.Equal(2, Inspector.Inspect(input).Content.TrailingEmptyBodyParagraphCount);
            var result = Editor.Apply(input, output,
                [new DocxEditOperation("collapseTrailingEmptyBodyParagraphs", ExpectedCount: 2)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            Assert.Equal(0, Inspector.Inspect(output).Content.TrailingEmptyBodyParagraphCount);
            using var edited = WordprocessingDocument.Open(output, false);
            var body = edited.MainDocumentPart!.Document!.Body!;
            Assert.Single(body.Elements<Paragraph>());
            Assert.Equal("approval block", body.Elements<Paragraph>().Single().InnerText);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Trailing_empty_body_paragraph_collapse_fails_closed_on_count_drift_or_meaningful_content()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-trailing-empty-paragraph-negative-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var mismatch = Path.Combine(root, "mismatch.docx");
        var noTrailingEmpty = Path.Combine(root, "no-trailing-empty.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("content"))),
                    new Paragraph(),
                    new SectionProperties()));
                main.Document.Save();
            }

            var countDrift = Editor.Apply(input, mismatch,
                [new DocxEditOperation("collapseTrailingEmptyBodyParagraphs", ExpectedCount: 2)]);
            Assert.False(Assert.Single(countDrift.AppliedOperations).Applied);
            Assert.Equal(1, Inspector.Inspect(mismatch).Content.TrailingEmptyBodyParagraphCount);

            var missingPrecondition = Editor.Apply(input, Path.Combine(root, "missing-precondition.docx"),
                [new DocxEditOperation("collapseTrailingEmptyBodyParagraphs")]);
            Assert.False(Assert.Single(missingPrecondition.AppliedOperations).Applied);

            using (var document = WordprocessingDocument.Create(noTrailingEmpty, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(),
                    new Paragraph(new Run(new Text("final content"))),
                    new SectionProperties()));
                main.Document.Save();
            }
            Assert.Equal(0, Inspector.Inspect(noTrailingEmpty).Content.TrailingEmptyBodyParagraphCount);
            var noCandidate = Editor.Apply(noTrailingEmpty, Path.Combine(root, "no-candidate-output.docx"),
                [new DocxEditOperation("collapseTrailingEmptyBodyParagraphs", ExpectedCount: 1)]);
            Assert.False(Assert.Single(noCandidate.AppliedOperations).Applied);

            var pageBreak = Path.Combine(root, "page-break.docx");
            using (var document = WordprocessingDocument.Create(pageBreak, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("content"))),
                    new Paragraph(new Run(new Break { Type = BreakValues.Page })),
                    new SectionProperties()));
                main.Document.Save();
            }
            Assert.Equal(0, Inspector.Inspect(pageBreak).Content.TrailingEmptyBodyParagraphCount);
            var preservesPageBreak = Editor.Apply(pageBreak, Path.Combine(root, "page-break-output.docx"),
                [new DocxEditOperation("collapseTrailingEmptyBodyParagraphs", ExpectedCount: 1)]);
            Assert.False(Assert.Single(preservesPageBreak.AppliedOperations).Applied);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Body_paragraph_keep_lines_is_set_without_changing_paragraph_content()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-paragraph-keep-lines-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(new Run(new Text("a long semantic conclusion"))),
                    new Paragraph(new Run(new Text("following note")))));
                main.Document.Save();
            }

            var result = Editor.Apply(
                input,
                output,
                [new DocxEditOperation("setBodyParagraphKeepLines", ParagraphIndex: 0, KeepLines: true)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            using var edited = WordprocessingDocument.Open(output, false);
            var paragraphs = edited.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
            Assert.Equal("a long semantic conclusion", paragraphs[0].InnerText);
            Assert.NotNull(paragraphs[0].ParagraphProperties?.GetFirstChild<KeepLines>());
            Assert.Null(paragraphs[1].ParagraphProperties?.GetFirstChild<KeepLines>());
            var flowJson = System.Text.Json.JsonSerializer.Serialize(Inspector.InspectDocumentFlow(output));
            Assert.Contains("\"keepLines\":true", flowJson, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Header_paragraph_font_size_updates_all_runs_and_preserves_tabs_and_text()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-header-font-size-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(new Run(new Text("body")))));
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new Header(new Paragraph(
                    new Run(new RunProperties(new FontSize { Val = "24" }), new Text("Record")),
                    new Run(new TabChar()),
                    new Run(new RunProperties(new FontSize { Val = "22" }), new Text("R-0042"))));
                var section = main.Document.Body!.AppendChild(new SectionProperties());
                section.AppendChild(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) });
                main.Document.Save();
                header.Header.Save();
            }

            var result = Editor.Apply(
                input,
                output,
                [new DocxEditOperation("setHeaderParagraphFontSize", HeaderIndex: 0, ParagraphIndex: 0, FontSize: "20")]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            using var edited = WordprocessingDocument.Open(output, false);
            var paragraph = edited.MainDocumentPart!.HeaderParts.Single().Header!.Elements<Paragraph>().Single();
            Assert.Equal("RecordR-0042", paragraph.InnerText);
            Assert.Single(paragraph.Descendants<TabChar>());
            Assert.All(paragraph.Descendants<Run>(), run =>
            {
                Assert.Equal("20", run.RunProperties?.FontSize?.Val?.Value);
                Assert.Equal("20", run.RunProperties?.FontSizeComplexScript?.Val?.Value);
            });
            var flowJson = System.Text.Json.JsonSerializer.Serialize(Inspector.InspectDocumentFlow(output));
            Assert.Contains("\"fontSizes\":[\"20\"]", flowJson, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Theory]
    [InlineData("setBodyParagraphKeepLines")]
    [InlineData("setHeaderParagraphFontSize")]
    public void Paragraph_layout_operations_fail_closed_when_required_values_are_missing(string type)
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-paragraph-layout-negative-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(new Paragraph(new Run(new Text("neutral")))));
                main.Document.Save();
            }
            var result = Editor.Apply(input, output, [new DocxEditOperation(type)]);
            Assert.False(Assert.Single(result.AppliedOperations).Applied);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Body_paragraph_keep_next_is_set_without_changing_paragraph_content()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-paragraph-pagination-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                        new Run(new Text("first note"))),
                    new Paragraph(new Run(new Text("second note")))));
                main.Document.Save();
            }

            var result = Editor.Apply(
                input,
                output,
                [new DocxEditOperation("setBodyParagraphKeepNext", ParagraphIndex: 0, KeepNext: true)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            using var edited = WordprocessingDocument.Open(output, false);
            var paragraphs = edited.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
            Assert.Equal("first note", paragraphs[0].InnerText);
            Assert.NotNull(paragraphs[0].ParagraphProperties?.GetFirstChild<KeepNext>());
            Assert.Null(paragraphs[1].ParagraphProperties?.GetFirstChild<KeepNext>());
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Keep_next_is_written_in_schema_order_and_reported_by_inspection()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-row-pagination-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Table(
                        new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                        new TableGrid(new GridColumn { Width = "5000" }),
                        new TableRow(
                            new TableCell(
                                new TableCellProperties(new TableCellWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
                                new Paragraph(
                                    new ParagraphProperties(new Justification { Val = JustificationValues.Left }),
                                    new Run(new Text("row"))))))));
                main.Document.Save();
            }

            var result = Editor.Apply(
                input,
                output,
                [new DocxEditOperation("setTableRowKeepNext", TableIndex: 0, RowIndex: 0, KeepNext: true)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            var row = Assert.Single(Assert.Single(Inspector.InspectTables(output).Tables).Rows);
            Assert.True(row.KeepNext);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Body_paragraph_keep_next_is_written_and_reported_by_document_flow()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-paragraph-pagination-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                Paragraph Numbered(string text) => new(
                    new ParagraphProperties(
                        new NumberingProperties(
                            new NumberingLevelReference { Val = 0 },
                            new NumberingId { Val = 2 })),
                    new Run(new Text(text)));
                main.Document = new Document(new Body(Numbered("first"), Numbered("second"), Numbered("last")));
                main.Document.Save();
            }

            var result = Editor.Apply(
                input,
                output,
                [new DocxEditOperation("setBodyParagraphKeepNext", ParagraphIndex: 1, KeepNext: true)]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            using (var edited = WordprocessingDocument.Open(output, false))
            {
                var paragraphs = edited.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
                Assert.Null(paragraphs[0].ParagraphProperties!.GetFirstChild<KeepNext>());
                Assert.NotNull(paragraphs[1].ParagraphProperties!.GetFirstChild<KeepNext>());
                Assert.Null(paragraphs[2].ParagraphProperties!.GetFirstChild<KeepNext>());
            }
            var flowJson = System.Text.Json.JsonSerializer.Serialize(Inspector.InspectDocumentFlow(output));
            Assert.Contains("\"paragraphIndex\":1", flowJson, StringComparison.Ordinal);
            Assert.Contains("\"numberingId\":2", flowJson, StringComparison.Ordinal);
            Assert.Contains("\"keepNext\":true", flowJson, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Trailing_empty_section_is_collapsed_without_changing_the_content_section()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-empty-section-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new SectionProperties(
                                new PageSize { Width = 16838, Height = 11906, Orient = PageOrientationValues.Landscape })),
                        new Run(new Text("content"))),
                    new Paragraph(),
                    new SectionProperties(new PageSize { Width = 11906, Height = 16838 })));
                main.Document.Save();
            }

            Assert.True(Inspector.Inspect(input).Content.HasTrailingEmptySection);
            var result = Editor.Apply(input, output, [new DocxEditOperation("collapseTrailingEmptySection")]);

            Assert.True(Assert.Single(result.AppliedOperations).Applied);
            Assert.False(Inspector.Inspect(output).Content.HasTrailingEmptySection);
            var validation = OpenXmlValidation.Validate(output);
            Assert.True(validation.Pass, string.Join(Environment.NewLine, validation.Errors.Select(error => error.Description)));
            using var edited = WordprocessingDocument.Open(output, false);
            var body = edited.MainDocumentPart!.Document!.Body!;
            Assert.Single(body.Elements<Paragraph>());
            var section = Assert.Single(body.Elements<SectionProperties>());
            Assert.Equal(PageOrientationValues.Landscape, section.GetFirstChild<PageSize>()!.Orient!.Value);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }
}
