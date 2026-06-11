using Xunit;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Docx;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace Dockit.Docx.Tests;

public class AnnotationToolsTests
{
    [Fact]
    public void Inspect_includes_annotation_anchors_in_unified_report()
    {
        var docPath = CreateAnnotatedFixture();

        var report = Inspector.Inspect(docPath);

        Assert.Equal(2, report.Annotations.CommentCount);
        Assert.Equal(2, report.Structure.AnnotationAnchors.Count);

        var paragraphAnchor = Assert.Single(report.Structure.AnnotationAnchors, anchor => anchor.CommentId == "0");
        Assert.Equal("paragraph", paragraphAnchor.TargetKind);
        Assert.Contains("Project code XXXX", paragraphAnchor.AnchorText);
        Assert.Equal("value comes from summary sheet", paragraphAnchor.CommentText);

        var tableAnchor = Assert.Single(report.Structure.AnnotationAnchors, anchor => anchor.CommentId == "1");
        Assert.Equal("tableCell", tableAnchor.TargetKind);
        Assert.Equal(0, tableAnchor.TableIndex);
        Assert.Equal(0, tableAnchor.RowIndex);
        Assert.Equal(1, tableAnchor.CellIndex);
        Assert.Contains("Batch YYYY", tableAnchor.AnchorText);
    }

    [Fact]
    public void Edit_applies_explicit_operations_and_preserves_other_content()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceAnchoredText", CommentId: "0", Text: "Project code HSP001"),
            new DocxEditOperation("replaceParagraphText", ParagraphIndex: 1, Text: "Top-level paragraph HSP001"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 1, Text: "Batch HSP001-01"),
            new DocxEditOperation("markFieldsDirty")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        var topParagraph = body.Elements<Paragraph>().First();
        Assert.Contains("Project code HSP001", GetParagraphText(topParagraph));

        var tableCellParagraph = body.Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().ElementAt(1)
            .Elements<Paragraph>().Single();
        Assert.Contains("Batch HSP001-01", GetParagraphText(tableCellParagraph));

        var topLevelParagraphs = body.Elements<Paragraph>().ToList();
        Assert.Contains("Top-level paragraph HSP001", GetParagraphText(topLevelParagraphs[1]));
        Assert.DoesNotContain("Top-level paragraph HSP001", string.Concat(body.Elements<Table>().Single().Descendants<Text>().Select(text => text.Text)));
        Assert.True(doc.MainDocumentPart.DocumentSettingsPart?.Settings?.Elements<UpdateFieldsOnOpen>().Any() == true);
    }

    [Fact]
    public void Edit_can_replace_table_with_formatted_rows()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-replaced-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation(
                "replaceTable",
                TableIndex: 0,
                Rows: [
                    [
                        new DocxTableCellInput("检测项目", Bold: true),
                        new DocxTableCellInput("时间点", GridSpan: 2, Bold: true)
                    ],
                    [
                        new DocxTableCellInput("颜色"),
                        new DocxTableCellInput("1月"),
                        new DocxTableCellInput("3月")
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        Assert.Equal("5000", table.GetFirstChild<TableProperties>()!.GetFirstChild<TableWidth>()!.Width);
        Assert.True(table.Elements<TableRow>().First().GetFirstChild<TableRowProperties>()!.Elements<TableHeader>().Any());
        Assert.True(table.Elements<TableRow>().First().Descendants<Bold>().Any());
        Assert.Equal(2, table.Elements<TableRow>().First().Elements<TableCell>().ElementAt(1).GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
        Assert.Contains("颜色", string.Concat(table.Descendants<Text>().Select(text => text.Text)));
        Assert.DoesNotContain(
            new OpenXmlValidator().Validate(doc).Select(error => error.Description),
            description => description.Contains("unexpected child element", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Edit_can_replace_table_with_rich_text_cells()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-rich-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation(
                "replaceTable",
                TableIndex: 0,
                Rows: [
                    [
                        new DocxTableCellInput("序号", Bold: true),
                        new DocxTableCellInput(
                            RichText: [
                                new DocxRichTextSegment("QV"),
                                new DocxRichTextSegment("Q", Color: "FF0000", Underline: true),
                                new DocxRichTextSegment("LV"),
                                new DocxRichTextSegment("Q", Color: "FF0000", Underline: true),
                                new DocxRichTextSegment("SGAEVK")
                            ])
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var richCell = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().ElementAt(1);
        var runs = richCell.Descendants<Run>().ToList();
        Assert.Equal(["QV", "Q", "LV", "Q", "SGAEVK"], runs.Select(run => string.Concat(run.Descendants<Text>().Select(text => text.Text))).ToArray());
        Assert.All(runs.Where(run => string.Concat(run.Descendants<Text>().Select(text => text.Text)) == "Q"), run =>
        {
            var properties = run.RunProperties;
            Assert.NotNull(properties);
            Assert.Equal("FF0000", properties!.GetFirstChild<Color>()!.Val!.Value);
            Assert.Equal(UnderlineValues.Single, properties.GetFirstChild<Underline>()!.Val!.Value);
        });
        Assert.DoesNotContain(
            new OpenXmlValidator().Validate(doc).Select(error => error.Description),
            description => description.Contains("unexpected child element", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Edit_can_replace_table_with_advanced_formatting()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-advanced-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation(
                "replaceTable",
                TableIndex: 0,
                Rows: [
                    [
                        new DocxTableCellInput("Header 1", Bold: true, Shading: "F2F2F2", Alignment: "center"),
                        new DocxTableCellInput("Header 2", Bold: true, Shading: "F2F2F2", Alignment: "center")
                    ],
                    [
                        new DocxTableCellInput("Merged Row", VMerge: "restart"),
                        new DocxTableCellInput("Value 1", Alignment: "right")
                    ],
                    [
                        new DocxTableCellInput("", VMerge: "continue"),
                        new DocxTableCellInput("Value 2", Alignment: "right")
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        
        var rows = table.Elements<TableRow>().ToList();
        Assert.Equal(3, rows.Count);

        var cell1 = rows[0].Elements<TableCell>().First();
        var shading = cell1.GetFirstChild<TableCellProperties>()!.GetFirstChild<Shading>();
        Assert.NotNull(shading);
        Assert.Equal("F2F2F2", shading.Fill!.Value);

        var p1 = cell1.Elements<Paragraph>().First();
        var jc = p1.GetFirstChild<ParagraphProperties>()!.GetFirstChild<Justification>();
        Assert.NotNull(jc);
        Assert.Equal(JustificationValues.Center, jc.Val!.Value);
        AssertChildOrder(cell1.GetFirstChild<TableCellProperties>()!, nameof(Shading), nameof(TableCellVerticalAlignment));

        var cell2_1 = rows[1].Elements<TableCell>().First();
        var vm1 = cell2_1.GetFirstChild<TableCellProperties>()!.GetFirstChild<VerticalMerge>();
        Assert.NotNull(vm1);
        Assert.Equal(MergedCellValues.Restart, vm1.Val!.Value);

        var cell3_1 = rows[2].Elements<TableCell>().First();
        var vm2 = cell3_1.GetFirstChild<TableCellProperties>()!.GetFirstChild<VerticalMerge>();
        Assert.NotNull(vm2);
        Assert.Equal(MergedCellValues.Continue, vm2.Val!.Value);

        var cell2_2 = rows[1].Elements<TableCell>().ElementAt(1);
        var p2_2 = cell2_2.Elements<Paragraph>().First();
        var jc2_2 = p2_2.GetFirstChild<ParagraphProperties>()!.GetFirstChild<Justification>();
        Assert.NotNull(jc2_2);
        Assert.Equal(JustificationValues.Right, jc2_2.Val!.Value);

        Assert.DoesNotContain(
            new OpenXmlValidator().Validate(doc).Select(error => error.Description),
            description => description.Contains("unexpected child element", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Edit_can_insert_and_replace_table_rows_using_existing_row_style()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-row-edit-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                    new TableGrid(
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "2000" }),
                    new TableRow(
                        new TableRowProperties(new TableHeader()),
                        new TableCell(
                            new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9EAF7" }),
                            new Paragraph(new Run(new RunProperties(new Bold()), new Text("序号")))),
                        new TableCell(
                            new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "D9EAF7" }),
                            new Paragraph(new Run(new RunProperties(new Bold()), new Text("肽段序列"))))
                    ),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
                            new Paragraph(new Run(new RunProperties(new RunFonts { Ascii = "Times New Roman" }), new Text("1")))),
                        new TableCell(
                            new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
                            new Paragraph(new Run(new RunProperties(new RunFonts { Ascii = "Times New Roman" }), new Text("QVQLVQSGAEVK"))))
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("footer")))),
                        new TableCell(new Paragraph(new Run(new Text("keep"))))
                    )
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"row-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "replaceTableRows",
                TableIndex: 0,
                StartRowIndex: 1,
                EndRowIndex: 1,
                TemplateRowIndex: 1,
                Rows: [
                    [
                        new DocxTableCellInput("1"),
                        new DocxTableCellInput(
                            RichText: [
                                new DocxRichTextSegment("QV"),
                                new DocxRichTextSegment("Q", Color: "FF0000", Underline: true),
                                new DocxRichTextSegment("LVQSGAEVK")
                            ])
                    ],
                    [
                        new DocxTableCellInput("2"),
                        new DocxTableCellInput("KPGASVK")
                    ]
                ]),
            new DocxEditOperation(
                "insertTableRows",
                TableIndex: 0,
                RowIndex: 3,
                TemplateRowIndex: 1,
                Rows: [
                    [
                        new DocxTableCellInput("3"),
                        new DocxTableCellInput("PGASVK")
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var table = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        var rows = table.Elements<TableRow>().ToList();

        Assert.Equal(5, rows.Count);
        Assert.Equal("序号肽段序列", string.Concat(rows[0].Descendants<Text>().Select(t => t.Text)));
        Assert.Equal("1QVQLVQSGAEVK", string.Concat(rows[1].Descendants<Text>().Select(t => t.Text)));
        Assert.Equal("2KPGASVK", string.Concat(rows[2].Descendants<Text>().Select(t => t.Text)));
        Assert.Equal("3PGASVK", string.Concat(rows[3].Descendants<Text>().Select(t => t.Text)));
        Assert.Equal("footerkeep", string.Concat(rows[4].Descendants<Text>().Select(t => t.Text)));

        var copiedCellProperties = rows[2].Elements<TableCell>().First().GetFirstChild<TableCellProperties>();
        Assert.NotNull(copiedCellProperties?.GetFirstChild<TableCellVerticalAlignment>());
        var markedRun = rows[1].Elements<TableCell>().ElementAt(1).Descendants<Run>().Single(run => string.Concat(run.Descendants<Text>().Select(t => t.Text)) == "Q");
        Assert.Equal("FF0000", markedRun.RunProperties!.GetFirstChild<Color>()!.Val!.Value);
        Assert.Equal(UnderlineValues.Single, markedRun.RunProperties.GetFirstChild<Underline>()!.Val!.Value);
        Assert.DoesNotContain(
            new OpenXmlValidator().Validate(edited).Select(error => error.Description),
            description => description.Contains("unexpected child element", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Edit_can_replace_table_cell_with_rich_text_runs_and_remove_text_fill()
    {
        var path = CreateRichTextTableFixture();
        var output = Path.Combine(Path.GetTempPath(), $"rich-cell-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: 0,
                CellIndex: 0,
                RichText: [
                    new DocxRichTextSegment("QV"),
                    new DocxRichTextSegment("Q", Color: "FF0000", Underline: true, FontName: "Times New Roman"),
                    new DocxRichTextSegment("LV"),
                    new DocxRichTextSegment("Q", Color: "FF0000", Underline: true, FontName: "Times New Roman"),
                    new DocxRichTextSegment("SGAEVK")
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var cell = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().Single();
        var runs = cell.Descendants<Run>().ToList();
        Assert.Equal(["QV", "Q", "LV", "Q", "SGAEVK"], runs.Select(run => string.Concat(run.Descendants<Text>().Select(text => text.Text))).ToArray());

        var markedRuns = runs.Where(run => string.Concat(run.Descendants<Text>().Select(text => text.Text)) == "Q").ToList();
        Assert.Equal(2, markedRuns.Count);
        Assert.All(markedRuns, run =>
        {
            var properties = run.RunProperties;
            Assert.NotNull(properties);
            Assert.Equal("FF0000", properties!.GetFirstChild<Color>()!.Val!.Value);
            Assert.Equal(UnderlineValues.Single, properties.GetFirstChild<Underline>()!.Val!.Value);
            var fonts = properties.GetFirstChild<RunFonts>();
            Assert.NotNull(fonts);
            Assert.Equal("Times New Roman", fonts!.Ascii!.Value);
            Assert.Equal("Times New Roman", fonts.HighAnsi!.Value);
            Assert.Empty(properties.Elements<W14.FillTextEffect>());
        });

        var xml = ReadZipEntry(output, "word/document.xml");
        Assert.DoesNotContain("w14:textFill", xml, StringComparison.Ordinal);
        Assert.DoesNotContain(
            new OpenXmlValidator().Validate(doc).Select(error => error.Description),
            description => description.Contains("unexpected child element", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Edit_can_set_table_cell_font_size_and_row_height()
    {
        var docPath = CreateTwoCellTableFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-layout-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("setTableCellFontSize", TableIndex: 0, RowIndex: 0, CellIndex: 1, FontSize: "9pt"),
            new DocxEditOperation("setTableRowHeight", TableIndex: 0, RowIndex: 0, Height: "240", HeightRule: "exact")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var row = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().Single();
        var targetCell = row.Elements<TableCell>().ElementAt(1);
        Assert.All(targetCell.Descendants<Run>(), run =>
        {
            var properties = run.RunProperties;
            Assert.NotNull(properties);
            Assert.Equal("18", properties!.GetFirstChild<FontSize>()!.Val!.Value);
            Assert.Equal("18", properties.GetFirstChild<FontSizeComplexScript>()!.Val!.Value);
        });

        var height = row.GetFirstChild<TableRowProperties>()!.GetFirstChild<TableRowHeight>();
        Assert.NotNull(height);
        Assert.Equal((UInt32Value)240U, height!.Val!);
        Assert.Equal(HeightRuleValues.Exact, height.HeightType!.Value);
        var validationErrors = new OpenXmlValidator().Validate(doc).Select(error => error.Description).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
    }

    [Fact]
    public void InspectTables_exports_cell_merge_and_run_format_details()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-detail-{Guid.NewGuid():N}.docx");

        Editor.Apply(docPath, output, [
            new DocxEditOperation(
                "replaceTable",
                TableIndex: 0,
                Rows: [
                    [
                        new DocxTableCellInput("序号", Bold: true, Shading: "F2F2F2", Alignment: "center"),
                        new DocxTableCellInput("EIC 比例", GridSpan: 2, Bold: true, Shading: "F2F2F2", Alignment: "center")
                    ],
                    [
                        new DocxTableCellInput("1", VMerge: "restart", Alignment: "center"),
                        new DocxTableCellInput(
                            VMerge: "restart",
                            Alignment: "center",
                            RichText: [
                                new DocxRichTextSegment("QV"),
                                new DocxRichTextSegment("Q", Color: "FF0000", Underline: true, FontName: "Times New Roman"),
                                new DocxRichTextSegment("LVQSGAEVK")
                            ]),
                        new DocxTableCellInput("/", Alignment: "center")
                    ],
                    [
                        new DocxTableCellInput("", VMerge: "continue", Alignment: "center"),
                        new DocxTableCellInput("", VMerge: "continue", Alignment: "center"),
                        new DocxTableCellInput("99.7", Alignment: "center")
                    ]
                ])
        ]);

        var report = Inspector.InspectTables(output);

        var table = Assert.Single(report.Tables);
        Assert.Equal(3, table.RowCount);
        Assert.Equal(3, table.ColumnCount);

        var headerCells = table.Rows[0].Cells;
        Assert.Equal(2, headerCells.Count);
        Assert.Equal(1, headerCells[1].GridColumnStart);
        Assert.Equal(2, headerCells[1].GridColumnEnd);
        Assert.Equal(2, headerCells[1].GridSpan);
        Assert.Equal("F2F2F2", headerCells[1].ShadingFill);
        Assert.Equal("center", headerCells[1].Paragraphs[0].Justification);

        var sequenceCell = table.Rows[1].Cells[1];
        Assert.Equal("restart", sequenceCell.VMerge);
        Assert.Equal("QVQLVQSGAEVK", sequenceCell.Text);
        var markedRun = Assert.Single(sequenceCell.Paragraphs[0].Runs, run => run.Text == "Q");
        Assert.Equal("FF0000", markedRun.Color);
        Assert.Equal("single", markedRun.Underline);
        Assert.Equal("Times New Roman", markedRun.FontAscii);
        Assert.Equal("Times New Roman", markedRun.FontHighAnsi);
        Assert.False(markedRun.HasTextFill);

        var sequenceContinue = table.Rows[2].Cells[1];
        Assert.Equal("continue", sequenceContinue.VMerge);
    }

    [Fact]
    public void NormalizeOpenXml_canonicalizes_prefixes_and_property_order()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"normalized-{Guid.NewGuid():N}.docx");
        File.Copy(docPath, output);
        ReplaceZipEntry(
            output,
            "word/document.xml",
            """
            <?xml version="1.0" encoding="utf-8"?>
            <ns0:document xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:ns1="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:ns2="http://schemas.microsoft.com/office/word/2010/wordml" ns1:Ignorable="w14 wp14">
              <ns0:body>
                <ns0:p ns2:paraId="11111111" ns2:textId="22222222">
                  <ns0:r>
                    <ns0:rPr><ns0:b/><ns0:rFonts ns0:ascii="Times New Roman"/></ns0:rPr>
                    <ns0:t>Text</ns0:t>
                  </ns0:r>
                </ns0:p>
              </ns0:body>
            </ns0:document>
            """);

        DocxPackageNormalizer.Normalize(output, output);

        var xml = ReadZipEntry(output, "word/document.xml");
        Assert.Contains("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"", xml);
        Assert.Contains("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"", xml);
        Assert.Contains("xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"", xml);
        Assert.Contains("xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\"", xml);
        Assert.Contains("mc:Ignorable=\"w14 wp14\"", xml);
        Assert.DoesNotContain("<ns0:", xml);
        Assert.True(xml.IndexOf("<w:rFonts", StringComparison.Ordinal) < xml.IndexOf("<w:b", xml.IndexOf("<w:rPr", StringComparison.Ordinal), StringComparison.Ordinal));
    }

    [Fact]
    public void Edit_can_replace_header_paragraph_text()
    {
        var docPath = CreateSplitPlaceholderFixture();
        var output = Path.Combine(Path.GetTempPath(), $"header-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceAllHeaderParagraphText", ParagraphIndex: 0, Text: "XX（客户项目代号）（与报告中HSPTEST对应）")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var headerText = string.Concat(
            doc.MainDocumentPart!.HeaderParts.SelectMany(part => part.Header!.Descendants<Text>()).Select(text => text.Text));
        Assert.Contains("XX（客户项目代号）（与报告中HSPTEST对应）", headerText);
        Assert.DoesNotContain("Header date", headerText);
    }

    [Fact]
    public void Edit_can_replace_header_text_without_overwriting_other_header_content()
    {
        var docPath = CreateHeaderLayoutFixture();
        var output = Path.Combine(Path.GetTempPath(), $"header-text-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceHeaderText", FindText: "XX（客户项目代号）（与报告中HSPTEST对应）", Text: "HSPTEST")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var headerParagraph = doc.MainDocumentPart!.HeaderParts.Single().Header!.Elements<Paragraph>().Single();
        var headerText = string.Concat(headerParagraph.Descendants<Text>().Select(text => text.Text));
        Assert.Contains("HSPTEST", headerText);
        Assert.Contains("3.2.S.7 稳定性", headerText);
        Assert.Contains("SN0000", headerText);
        Assert.DoesNotContain("XX（客户项目代号）（与报告中HSPTEST对应）", headerText);
        Assert.True(headerParagraph.Descendants<TabChar>().Count() >= 2);
    }

    [Fact]
    public void Edit_can_replace_body_text_without_rewriting_paragraph_structure()
    {
        var docPath = CreateSplitBodyTextFixture();
        var output = Path.Combine(Path.GetTempPath(), $"body-text-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceBodyText", FindText: "HSPXXX", Text: "HSP-PTMs")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var paragraph = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single();
        var runs = paragraph.Elements<Run>().ToList();
        var bodyText = string.Concat(runs.Select(run => string.Concat(run.Descendants<Text>().Select(text => text.Text))));
        Assert.Equal("表 11. HSP-PTMs 样品翻译后修饰结果", bodyText);
        Assert.Equal(3, runs.Count);
        Assert.Equal("Times New Roman", runs[0].RunProperties!.RunFonts!.Ascii!.Value);
        Assert.Equal("000000", runs[0].RunProperties!.Color!.Val!.Value);
    }

    [Fact]
    public void Edit_can_freeze_fields_to_current_display_text()
    {
        var docPath = CreateFieldFixture();
        var output = Path.Combine(Path.GetTempPath(), $"freeze-fields-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("freezeFields")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Empty(body.Descendants<SimpleField>());
        var fieldCodes = body.Descendants<FieldCode>().Select(code => code.Text).ToList();
        Assert.DoesNotContain(fieldCodes, code => code.Contains("REF", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(fieldCodes, code => code.Contains("SEQ", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(fieldCodes, code => code.Contains("PAGE", StringComparison.OrdinalIgnoreCase));
        Assert.NotEmpty(body.Descendants<FieldChar>());

        var text = string.Concat(body.Descendants<Text>().Select(text => text.Text));
        Assert.Contains("见表 11。", text);
        Assert.Contains("表 11. HSP-PTMs样品翻译后修饰结果", text);
    }

    [Fact]
    public void Edit_can_delete_comments_explicitly()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"clean-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("deleteComments"),
            new DocxEditOperation("markFieldsDirty")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var mainPart = doc.MainDocumentPart!;
        Assert.Null(mainPart.WordprocessingCommentsPart);
        Assert.Empty(mainPart.Document!.Descendants<CommentRangeStart>());
        Assert.Empty(mainPart.Document.Descendants<CommentRangeEnd>());
        Assert.Empty(mainPart.Document.Descendants<CommentReference>());
    }

    [Fact]
    public void ExportJson_includes_body_paragraph_and_table_indexes()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"export-{Guid.NewGuid():N}.json");

        Transforms.RunExportJson([docPath, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("Project code XXXX 峰面积", json, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\u5CF0", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u9762", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u79EF", json, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\"paragraphIndex\": 0", json, StringComparison.Ordinal);
        Assert.Contains("\"tableIndex\": 0", json, StringComparison.Ordinal);
    }

    [Fact]
    public void ExportJson_includes_header_paragraphs()
    {
        var docPath = CreateSplitPlaceholderFixture();
        var output = Path.Combine(Path.GetTempPath(), $"export-header-{Guid.NewGuid():N}.json");

        Transforms.RunExportJson([docPath, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("\"type\": \"headerParagraph\"", json, StringComparison.Ordinal);
        Assert.Contains("\"headerIndex\": 0", json, StringComparison.Ordinal);
        Assert.Contains("Header date:", json, StringComparison.Ordinal);
    }

    [Fact]
    public void FillTemplate_replaces_split_placeholders_in_body_and_header()
    {
        var docPath = CreateSplitPlaceholderFixture();
        var dataPath = Path.Combine(Path.GetTempPath(), $"fill-{Guid.NewGuid():N}.json");
        var output = Path.Combine(Path.GetTempPath(), $"filled-{Guid.NewGuid():N}.docx");

        File.WriteAllText(
            dataPath,
            """
            {
              "cellValues": {
                "effectiveDate": "2024-09-18"
              }
            }
            """,
            System.Text.Encoding.UTF8);

        Transforms.RunFillTemplate([docPath, dataPath, output]);

        var report = Inspector.Inspect(output);
        Assert.DoesNotContain("{{effectiveDate}}", report.Content.Placeholders);

        using var doc = WordprocessingDocument.Open(output, false);
        var bodyText = string.Concat(doc.MainDocumentPart!.Document!.Descendants<Text>().Select(text => text.Text));
        Assert.Contains("2024-09-18", bodyText);

        var headerText = string.Concat(
            doc.MainDocumentPart.HeaderParts.SelectMany(part => part.Header!.Descendants<Text>()).Select(text => text.Text));
        Assert.Contains("2024-09-18", headerText);
    }

    private static string CreateAnnotatedFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"annotated-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
        var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments = new Comments(
            CreateComment("0", "tester", "value comes from summary sheet"),
            CreateComment("1", "tester", "batch id comes from inspection report"));

        var body = mainPart.Document.Body!;
        body.Append(CreateParagraphWithComment("0", "Project code XXXX 峰面积"));
        body.Append(CreateTableWithComment());
        body.Append(CreateFieldParagraph());
        mainPart.Document.Save();
        commentsPart.Comments.Save();
        return path;
    }

    private static string CreateRichTextTableFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"rich-text-table-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Table(
                new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                new TableGrid(new GridColumn { Width = "2400" }),
                new TableRow(
                    new TableCell(
                        new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
                        new Paragraph(
                            new Run(
                                new RunProperties(
                                    new Color { Val = "000000" },
                                    new W14.FillTextEffect()),
                                new Text("QVQLVQSGAEVK"))))))));
        mainPart.Document.Save();
        return path;
    }

    private static string CreateTwoCellTableFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"two-cell-table-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Table(
                new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                new TableGrid(
                    new GridColumn { Width = "2400" },
                    new GridColumn { Width = "2400" }),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Label")))),
                    new TableCell(new Paragraph(new Run(new Text("Batch YYYY"))))))));
        mainPart.Document.Save();
        return path;
    }

    private static string CreateSplitBodyTextFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"split-body-text-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Paragraph(
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                        new Color { Val = "000000" }),
                    new Text("表 11. H")),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                        new Color { Val = "000000" }),
                    new Text("SPXXX")),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                        new Color { Val = "000000" }),
                    new Text(" 样品翻译后修饰结果")))));
        mainPart.Document.Save();
        return path;
    }

    private static Paragraph CreateParagraphWithComment(string commentId, string text)
    {
        return new Paragraph(
            new CommentRangeStart { Id = commentId },
            new Run(new Text(text)),
            new CommentRangeEnd { Id = commentId },
            new Run(new CommentReference { Id = commentId }));
    }

    private static Table CreateTableWithComment()
    {
        return new Table(
            new TableRow(
                CreateCell("Label"),
                CreateCellWithComment("1", "Batch YYYY")));
    }

    private static TableCell CreateCell(string text)
        => new(
            new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
            new Paragraph(new Run(new Text(text))));

    private static TableCell CreateCellWithComment(string commentId, string text)
        => new(
            new TableCellProperties(new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
            new Paragraph(
                new CommentRangeStart { Id = commentId },
                new Run(new Text(text)),
                new CommentRangeEnd { Id = commentId },
                new Run(new CommentReference { Id = commentId })));

    private static string CreateFieldFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"field-freeze-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        var body = new Body();
        mainPart.Document = new Document(body);

        body.Append(new Paragraph(
            new Run(new Text("见")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" REF _RefTable11 \\h ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("表 11")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
            new Run(new Text("。"))));

        body.Append(new Paragraph(
            new BookmarkStart { Id = "1", Name = "_RefTable11" },
            new Run(new Text("表 ")),
            new SimpleField(
                new Run(new Text("11")))
            { Instruction = "SEQ 表 \\* ARABIC", Dirty = false },
            new BookmarkEnd { Id = "1" },
            new Run(new Text(". HSP-PTMs样品翻译后修饰结果"))));

        body.Append(new Paragraph(
            new Run(new Text("页码：")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("1")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End })));

        mainPart.Document.Save();
        return path;
    }

    private static Paragraph CreateFieldParagraph()
    {
        return new Paragraph(
            new SimpleField { Instruction = "SEQ Figure \\* ARABIC", Dirty = false },
            new Run(new Text("Figure 1")));
    }

    private static string CreateSplitPlaceholderFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"split-placeholder-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(
            new Paragraph(
                new Run(new Text("Header date: ")),
                new Run(new Text("{{")),
                new Run(new Text("effectiveDate")),
                new Run(new Text("}}"))));

        var headerPartId = mainPart.GetIdOfPart(headerPart);
        var sectionProps = new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId });

        var body = mainPart.Document.Body!;
        body.Append(
            new Paragraph(
                new Run(new Text("Body date: ")),
                new Run(new Text("{{")),
                new Run(new Text("effectiveDate")),
                new Run(new Text("}}"))));
        body.Append(new Paragraph(new Run(new Text("after"))));
        body.Append(sectionProps);

        mainPart.Document.Save();
        headerPart.Header.Save();
        return path;
    }

    private static string CreateHeaderLayoutFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"header-layout-{Guid.NewGuid():N}.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var headerPart = mainPart.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(
            new Paragraph(
                new Run(new Text("XX（客户项目代号）（与报告中HSPTEST对应）")),
                new Run(new TabChar()),
                new Run(new Text("3.2.S.7 稳定性")),
                new Run(new TabChar()),
                new Run(new Text("SN0000"))));

        var headerPartId = mainPart.GetIdOfPart(headerPart);
        var sectionProps = new SectionProperties(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerPartId });
        var body = mainPart.Document.Body!;
        body.Append(new Paragraph(new Run(new Text("body"))));
        body.Append(sectionProps);

        mainPart.Document.Save();
        headerPart.Header.Save();
        return path;
    }

    [Fact]
    public void Edit_can_merge_table_cells_horizontally_and_vertically()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-merge-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("R1C1")))),
                        new TableCell(new Paragraph(new Run(new Text("R1C2"))))
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("R2C1")))),
                        new TableCell(new Paragraph(new Run(new Text("R2C2"))))
                    )
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"merged-cells-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("mergeTableCells", TableIndex: 0, RowIndex: 0, StartCellIndex: 0, EndCellIndex: 1),
            new DocxEditOperation("mergeTableCells", TableIndex: 0, CellIndex: 0, StartRowIndex: 0, EndRowIndex: 1)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
            var rows = table.Elements<TableRow>().ToList();

            var r1Cell = rows[0].Elements<TableCell>().Single();
            var span = r1Cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>();
            Assert.NotNull(span);
            Assert.Equal(2, span.Val!.Value);

            var vm1 = r1Cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>();
            Assert.NotNull(vm1);
            Assert.Equal(MergedCellValues.Restart, vm1.Val!.Value);

            var r2Cell = rows[1].Elements<TableCell>().ElementAt(0);
            var vm2 = r2Cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>();
            Assert.NotNull(vm2);
            Assert.Equal(MergedCellValues.Continue, vm2.Val!.Value);
        }
    }

    [Fact]
    public void Edit_can_insert_table_columns_and_expand_crossing_grid_spans()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-column-insert-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableGrid(
                        new GridColumn { Width = "1800" },
                        new GridColumn { Width = "800" },
                        new GridColumn { Width = "800" },
                        new GridColumn { Width = "800" }
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("条件")))),
                        new TableCell(new Paragraph(new Run(new Text("T0")))),
                        new TableCell(new Paragraph(new Run(new Text("1月")))),
                        new TableCell(new Paragraph(new Run(new Text("3月"))))
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("反复冻融试验")))),
                        new TableCell(
                            new TableCellProperties(new GridSpan { Val = 3 }),
                            new Paragraph(new Run(new Text("冻融3个循环、5个循环，取样检测按A进行测定")))
                        )
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("长期")))),
                        new TableCell(new Paragraph(new Run(new Text("--")))),
                        new TableCell(new Paragraph(new Run(new Text("A")))),
                        new TableCell(new Paragraph(new Run(new Text("B"))))
                    )
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"column-inserted-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("insertTableColumns", TableIndex: 0, ColumnIndex: 3, ColumnCount: 2, TemplateColumnIndex: 2),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 3, Text: "6月"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 4, Text: "9月"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 2, CellIndex: 3, Text: "A"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 2, CellIndex: 4, Text: "--")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
            Assert.Equal(6, table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().Count());

            var rows = table.Elements<TableRow>().ToList();
            Assert.Equal(["条件", "T0", "1月", "6月", "9月", "3月"], rows[0].Elements<TableCell>().Select(GetCellText).ToArray());

            var freezeThawCells = rows[1].Elements<TableCell>().ToList();
            Assert.Equal(2, freezeThawCells.Count);
            var span = freezeThawCells[1].GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value;
            Assert.Equal(5, span);
            Assert.Contains("冻融3个循环", GetCellText(freezeThawCells[1]));

            Assert.Equal(["长期", "--", "A", "A", "--", "B"], rows[2].Elements<TableCell>().Select(GetCellText).ToArray());
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
    }

    [Fact]
    public void Edit_can_unmerge_table_column_vertical_cells_and_fill_continuations()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-unmerge-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("A1")))),
                        new TableCell(
                            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }),
                            new Paragraph(new Run(new Text("Ratio")))
                        )
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("A2")))),
                        new TableCell(
                            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Continue }),
                            new Paragraph()
                        )
                    )
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"unmerged-cells-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("unmergeTableColumnVerticalCells", TableIndex: 0, CellIndex: 1, StartRowIndex: 0, EndRowIndex: 1)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
            var rows = table.Elements<TableRow>().ToList();

            foreach (var row in rows)
            {
                var cell = row.Elements<TableCell>().ElementAt(1);
                Assert.Null(cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>());
                Assert.Equal("Ratio", string.Concat(cell.Descendants<Text>().Select(t => t.Text)));
            }
        }
    }

    [Fact]
    public void Edit_can_unmerge_table_row_horizontal_cells()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-unmerge-horizontal-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(),
                    new TableGrid(
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "1000" }
                    ),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new GridSpan { Val = 3 }),
                            new Paragraph(new Run(new Text("高温试验")))
                        )
                    )
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"unmerged-horizontal-cells-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("unmergeTableRowHorizontalCells", TableIndex: 0, RowIndex: 0, CellIndex: 0)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
            var cells = table.Elements<TableRow>().Single().Elements<TableCell>().ToList();
            Assert.Equal(3, cells.Count);
            Assert.Null(cells[0].GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>());
            Assert.Equal("高温试验", GetCellText(cells[0]));
            Assert.Equal("", GetCellText(cells[1]));
            Assert.Equal("", GetCellText(cells[2]));
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
    }

    [Fact]
    public void Edit_applies_fillTableSemantically_correctly()
    {
        var path = CreateSemanticTableFixture();
        var output = Path.Combine(Path.GetTempPath(), $"semantic-filled-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("fillTableSemantically", TableIndex: 0, Cells: [
                new DocxSemanticFillRule(RowPatterns: ["pH"], ColPatterns: ["1个月"], Text: "5.3"),
                new DocxSemanticFillRule(RowPatterns: ["主峰"], ColPatterns: ["1个月"], Text: "98.6")
            ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
            var gridMap = new TableGridMap(table);
            
            Assert.Equal("5.3", string.Concat(gridMap.Grid[1, 3]!.Descendants<Text>().Select(t => t.Text)).Trim());
            Assert.Equal("98.6", string.Concat(gridMap.Grid[2, 3]!.Descendants<Text>().Select(t => t.Text)).Trim());
        }
    }

    private static string CreateSemanticTableFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"semantic-template-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("检测项目")))),
                        new TableCell(new Paragraph(new Run(new Text("参考标准")))),
                        new TableCell(new Paragraph(new Run(new Text("T0")))),
                        new TableCell(new Paragraph(new Run(new Text("1个月"))))
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("pH")))),
                        new TableCell(new Paragraph(new Run(new Text("5.1±0.3")))),
                        new TableCell(new Paragraph(new Run(new Text("5.2")))),
                        new TableCell(new Paragraph(new Run(new Text(""))))
                    ),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("主峰")))),
                        new TableCell(new Paragraph(new Run(new Text("≥95.0%")))),
                        new TableCell(new Paragraph(new Run(new Text("98.4")))),
                        new TableCell(new Paragraph(new Run(new Text(""))))
                    )
                )
            ));
            mainPart.Document.Save();
        }
        return path;
    }

    private static Comment CreateComment(string id, string author, string text)
    {
        var comment = new Comment
        {
            Id = id,
            Author = author,
            Initials = author,
            Date = DateTime.Parse("2026-04-15T00:00:00Z")
        };
        comment.Append(new Paragraph(new Run(new Text(text))));
        return comment;
    }

    private static string GetParagraphText(Paragraph paragraph)
        => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));

    private static string GetCellText(TableCell cell)
        => string.Concat(cell.Descendants<Text>().Select(text => text.Text));

    private static void AssertChildOrder(OpenXmlElement parent, string beforeTypeName, string afterTypeName)
    {
        var children = parent.ChildElements.ToList();
        var beforeIndex = children.FindIndex(child => child.GetType().Name == beforeTypeName);
        var afterIndex = children.FindIndex(child => child.GetType().Name == afterTypeName);
        Assert.True(beforeIndex >= 0, $"{beforeTypeName} was not found under {parent.GetType().Name}");
        Assert.True(afterIndex >= 0, $"{afterTypeName} was not found under {parent.GetType().Name}");
        Assert.True(beforeIndex < afterIndex, $"{beforeTypeName} should appear before {afterTypeName}");
    }

    private static string ReadZipEntry(string path, string entryName)
    {
        using var archive = ZipFile.OpenRead(path);
        using var stream = archive.GetEntry(entryName)!.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    private static void ReplaceZipEntry(string path, string entryName, string text)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(entryName)?.Delete();
        var entry = archive.CreateEntry(entryName);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream);
        writer.Write(text);
    }
}
