using Xunit;
using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Dockit.Docx;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using System.Security.Cryptography;
using System.Text.Json;

namespace Dockit.Docx.Tests;

public class AnnotationToolsTests
{
    [Fact]
    public void Edit_command_exits_nonzero_when_any_operation_is_not_applied()
    {
        var source = CreateSemanticTableFixture();
        var root = Path.Combine(Path.GetTempPath(), $"docx-edit-fail-closed-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var operations = Path.Combine(root, "operations.json");
        var output = Path.Combine(root, "output.docx");
        File.WriteAllText(operations, JsonSerializer.Serialize(new DocxEditDocument([
            new DocxEditOperation("replaceTableCellText", TableIndex: 99, RowIndex: 0, CellIndex: 0, Text: "unreachable")
        ]), Json.Options));

        Assert.Equal(1, Editor.RunEdit([source, operations, output]));
        Assert.True(File.Exists(output));
    }

    [Fact]
    public void OpenXmlValidation_accepts_a_valid_document()
    {
        var input = CreateTextMigrationFixture("valid document");
        Assert.Equal(0, OpenXmlValidation.Run([input]));
    }

    [Fact]
    public void TemplateMigration_analysis_exports_hash_attested_object_inventories_without_guessing_mapping()
    {
        var source = CreateAnnotatedFixture();
        var baseline = Path.Combine(Path.GetTempPath(), $"migration-baseline-{Guid.NewGuid():N}.docx");

        Editor.Apply(source, baseline, [
            new DocxEditOperation("replaceParagraphText", ParagraphIndex: 0, Text: "Target format heading"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 1, Text: "Target placeholder")
        ]);

        var analysis = TemplateMigration.Analyze(source, baseline);

        Assert.Equal("tiwater.docx.template-migration-analysis/v1", analysis.Schema);
        Assert.Matches("^[A-F0-9]{64}$", analysis.Source.Sha256);
        Assert.Matches("^[A-F0-9]{64}$", analysis.Baseline.Sha256);
        Assert.Contains(analysis.Source.Objects, item => item.Id == "body:paragraph:0" && item.Kind == "paragraph");
        Assert.Contains(analysis.Source.Objects, item => item.Id == "body:table:0:row:0:cell:1" && item.Kind == "table-cell");
        Assert.Contains(analysis.Findings, item => item.SourceObjectId == "body:paragraph:0" && item.Kind == "object-content-differs");
        Assert.Contains(analysis.Findings, item => item.SourceObjectId == "body:table:0:row:0:cell:1" && item.Kind == "object-content-differs");
        Assert.All(analysis.Findings, item => Assert.Equal("requires-declared-mapping", item.Disposition));
    }

    [Fact]
    public void TemplateMigration_analysis_inventories_nested_tables_as_distinct_objects()
    {
        var source = Path.Combine(Path.GetTempPath(), $"migration-nested-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var inner = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("inner"))))));
            var outer = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("outer"))), inner)));
            main.Document = new Document(new Body(outer));
            main.Document.Save();
        }

        var analysis = TemplateMigration.Analyze(source, source);

        Assert.Equal(2, analysis.Source.Objects.Count(item => item.Kind == "table"));
        Assert.Contains(analysis.Source.Objects, item => item.Id == "body:table:0:row:0:cell:0:table:0" && item.ParentId == "body:table:0:row:0:cell:0");
    }

    [Fact]
    public void TemplateMigration_analysis_publishes_canonical_table_cell_topology()
    {
        var source = Path.Combine(Path.GetTempPath(), $"migration-topology-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var table = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("a")))),
                    new TableCell(new Paragraph(new Run(new Text("b"))))),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("c")))),
                    new TableCell(new Paragraph(new Run(new Text("d"))))));
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }

        var analysis = TemplateMigration.Analyze(source, source);
        var cell = Assert.Single(analysis.Source.Objects, item => item.Id == "body:table:0:row:1:cell:1");
        Assert.NotNull(cell.Topology);
        Assert.Equal("body:table:0", cell.Topology!.ContainerObjectId);
        Assert.Equal(1, cell.Topology.Row);
        Assert.Equal(1, cell.Topology.Column);
        Assert.All(analysis.Source.Objects.Where(item => item.Kind == "table-cell"), item => Assert.NotNull(item.Topology));
        Assert.All(analysis.Source.Objects.Where(item => item.Kind != "table-cell"), item => Assert.Null(item.Topology));
    }

    [Fact]
    public void Edit_replace_table_cell_text_preserves_embedded_drawings()
    {
        var source = Path.Combine(Path.GetTempPath(), $"cell-drawing-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var image = main.AddImagePart(ImagePartType.Png);
            image.FeedData(new MemoryStream([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]));
            var drawing = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = 990000L, Cy = 990000L },
                    new DW.DocProperties { Id = 1U, Name = "cell-image" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "cell-image" },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = main.GetIdOfPart(image) },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 990000L, Cy = 990000L }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
            var cell = new TableCell(
                new Paragraph(new Run(drawing)),
                new Paragraph(new Run(new Text("old label"))));
            main.Document = new Document(new Body(new Table(new TableRow(cell, new TableCell(new Paragraph(new Run(new Text("other"))))))));
            main.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"cell-drawing-out-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(source, output, [
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 0, Text: "new label")
        ]);
        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var editedCell = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().Single().Elements<TableCell>().First();
        Assert.Contains(editedCell.Descendants<Text>(), text => text.Text == "new label");
        Assert.Single(editedCell.Descendants<Drawing>());
        Assert.Equal(mainRelationshipId(source), editedCell.Descendants<A.Blip>().Single().Embed!.Value);

        static string mainRelationshipId(string path)
        {
            using var document = WordprocessingDocument.Open(path, false);
            return document.MainDocumentPart!.GetIdOfPart(document.MainDocumentPart.ImageParts.Single());
        }
    }

    [Fact]
    public void InspectTables_versions_the_view_and_addresses_nested_tables_without_leaking_nested_text()
    {
        var source = Path.Combine(Path.GetTempPath(), $"inspect-nested-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var inner = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("inner"))))));
            var outer = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("outer"))), inner)));
            main.Document = new Document(new Body(outer)); main.Document.Save();
        }

        var report = Inspector.InspectTables(source);

        Assert.Equal("tiwater.docx.inspect-tables/v1", report.Schema);
        Assert.NotEmpty(report.ToolVersion);
        Assert.Equal("direct-cell-paragraphs-excluding-nested-tables", report.ExtractionView["cellText"]);
        Assert.Equal(2, report.Tables.Count);
        Assert.Equal(report.Tables[0].ColumnCount, report.Tables[0].GridColumnCount);
        Assert.Equal(report.Tables[0].GridColumnCount, report.Tables[0].GridColumnWidths.Count);
        Assert.Equal("outer", report.Tables[0].Rows[0].Cells[0].Text);
        Assert.Equal(["body", "table:0", "row:0", "cell:0", "table:0"], report.Tables[1].ContainmentPath);
        Assert.Equal("table:0:row:0:cell:0", report.Tables[1].ParentCellAddress);
        Assert.Equal("inner", report.Tables[1].Rows[0].Cells[0].Text);
    }

    [Fact]
    public void TemplateMigration_inventory_captures_runs_sections_revisions_and_media_without_document_specific_rules()
    {
        var source = Path.Combine(Path.GetTempPath(), $"migration-rich-inventory-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new Paragraph(new Run(new RunProperties(new Bold()), new Text("stable content")));
            paragraph.Append(new InsertedRun(new Run(new Text("tracked content")))
            {
                Author = "reviewer",
                Date = new DateTimeValue(DateTime.Parse("2026-07-19T00:00:00Z", null, System.Globalization.DateTimeStyles.RoundtripKind))
            });
            main.Document = new Document(new Body(
                paragraph,
                new SectionProperties(new PageSize { Width = 11906, Height = 16838 }, new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            var image = main.AddImagePart(ImagePartType.Png);
            using var imageBytes = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
            image.FeedData(imageBytes);
            main.Document.Save();
        }

        var inventory = TemplateMigration.Analyze(source, source).Source.Objects;

        var run = Assert.Single(inventory, item => item.Kind == "run" && item.Text == "stable content");
        Assert.Equal("body:paragraph:0", run.ParentId);
        Assert.Matches("^[A-F0-9]{64}$", run.Provenance["runPropertiesSha256"]);
        var section = Assert.Single(inventory, item => item.Kind == "section");
        Assert.Matches("^[A-F0-9]{64}$", section.Provenance["pageMarginSha256"]);
        var revision = Assert.Single(inventory, item => item.Kind == "revision");
        Assert.Equal("ins", revision.Provenance["revisionType"]);
        Assert.Equal("reviewer", revision.Provenance["author"]);
        var media = Assert.Single(inventory, item => item.Kind == "media");
        Assert.Equal("image/png", media.Provenance["contentType"]);
        Assert.Matches("^[A-F0-9]{64}$", media.Provenance["sha256"]);

        var derived = TemplateMigration.DeriveExactTextPlan(source, source);
        Assert.False(derived.Pass);
        Assert.Contains(derived.Unresolved, item => item.SourceObjectId == revision.Id && item.Reason == "template-migration-automatic-strategy-unsupported");
        Assert.Contains(derived.Unresolved, item => item.SourceObjectId == media.Id && item.Reason == "template-migration-automatic-strategy-unsupported");
    }

    [Fact]
    public void TemplateMigration_inventory_requires_review_for_present_unsupported_document_objects()
    {
        var annotated = CreateAnnotatedFixture();
        var annotatedAnalysis = TemplateMigration.Analyze(annotated, annotated);
        var annotatedPlan = TemplateMigration.DeriveExactTextPlan(annotated, annotated);
        Assert.Contains("comments", annotatedAnalysis.UnsupportedObjectKinds);
        Assert.Contains(annotatedPlan.Plan.Mappings, item => item.SourceObjectId == "mainDocument:comments" && item.Disposition == "review-required");

        var controlled = Path.Combine(Path.GetTempPath(), $"migration-content-control-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(controlled, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new SdtBlock(
                new SdtProperties(new Tag { Val = "controlled" }),
                new SdtContentBlock(new Paragraph(new Run(new Text("controlled source content")))))));
            main.Document.Save();
        }
        var controlledAnalysis = TemplateMigration.Analyze(controlled, controlled);
        var controlledPlan = TemplateMigration.DeriveExactTextPlan(controlled, controlled);
        Assert.Contains("content-control", controlledAnalysis.UnsupportedObjectKinds);
        Assert.Contains(controlledPlan.Plan.Mappings, item => item.Disposition == "review-required" && item.Reason == "template-migration-automatic-strategy-unsupported");
    }

    [Fact]
    public void TemplateMigration_copies_declared_media_into_a_current_baseline_slot_and_proves_readback()
    {
        var source = CreateMediaMigrationFixture("source text", [1, 2, 3, 4]);
        var baseline = CreateMediaMigrationFixture("baseline placeholder", [9, 8, 7, 6]);
        var analysis = TemplateMigration.Analyze(source, baseline);
        var sourceMedia = Assert.Single(analysis.Source.Objects, item => item.Kind == "media");
        var baselineMedia = Assert.Single(analysis.Baseline.Objects, item => item.Kind == "media");
        var plan = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v1",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            [
                new TemplateMigrationMapping("body:paragraph:0", "body:paragraph:0", "copy-text"),
                new TemplateMigrationMapping(sourceMedia.Id, baselineMedia.Id, "copy-media")
            ]);

        var build = TemplateMigration.BuildOperations(source, baseline, plan);
        Assert.True(build.Pass, string.Join("; ", build.Failures.Select(item => item.Reason)));
        Assert.Single(build.MediaCopies);
        var output = Path.Combine(Path.GetTempPath(), $"migration-media-output-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, plan, output);

        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));
        Assert.Empty(applied.MediaFailures);
        var outputMedia = Assert.Single(TemplateMigration.Analyze(output, output).Source.Objects, item => item.Id == baselineMedia.Id);
        Assert.Equal(sourceMedia.Provenance["sha256"], outputMedia.Provenance["sha256"]);
    }

    [Fact]
    public void TemplateMigration_resolves_semantic_text_selectors_to_current_ids_without_accepting_coordinates()
    {
        var source = CreateTextMigrationFixture("legacy factual content");
        var baseline = CreateTextMigrationFixture("target format placeholder");
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("paragraph", Scope: "body", Text: "legacy factual content"),
                    new TemplateMigrationSemanticSelector("paragraph", Scope: "body", Text: "target format placeholder"),
                    "copy-text")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);

        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        var mapping = Assert.Single(resolved.Plan.Mappings);
        Assert.Equal("body:paragraph:0", mapping.SourceObjectId);
        Assert.Equal("body:paragraph:0", mapping.BaselineObjectId);
        var output = Path.Combine(Path.GetTempPath(), $"migration-semantic-output-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));

        var duplicateSource = Path.Combine(Path.GetTempPath(), $"migration-semantic-duplicate-{Guid.NewGuid():N}.docx");
        using (var document = WordprocessingDocument.Create(duplicateSource, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("legacy factual content"))), new Paragraph(new Run(new Text("legacy factual content")))));
            main.Document.Save();
        }
        var rejected = TemplateMigration.ResolveSemanticCandidate(duplicateSource, baseline, candidate);
        Assert.False(rejected.Pass);
        Assert.Contains(rejected.Unresolved, item => item.Reason == "template-migration-semantic-source-ambiguous");

        var contextualSource = CreateTextMigrationFixture("before source", "repeated label", "after source", "repeated label");
        var contextualBaseline = CreateTextMigrationFixture("before target", "target slot one", "after target", "target slot two");
        var contextualCandidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("paragraph", Scope: "body", Text: "repeated label", PreviousText: "before source"),
                    new TemplateMigrationSemanticSelector("paragraph", Scope: "body", Text: "target slot one", PreviousText: "before target"),
                    "copy-text")
            ]);
        var contextResolved = TemplateMigration.ResolveSemanticCandidate(contextualSource, contextualBaseline, contextualCandidate);
        Assert.Contains(contextResolved.Plan.Mappings, item => item.SourceObjectId == "body:paragraph:1" && item.BaselineObjectId == "body:paragraph:1" && item.Disposition == "copy-text");

        var invalidCandidate = Path.Combine(Path.GetTempPath(), $"migration-semantic-invalid-{Guid.NewGuid():N}.json");
        File.WriteAllText(invalidCandidate, """
        {"schema":"tiwater.docx.template-migration-semantic-candidate/v1","mappings":[{"source":{"kind":"paragraph","text":"legacy factual content","sourceObjectId":"body:paragraph:0"},"baseline":{"kind":"paragraph","text":"target format placeholder"},"disposition":"copy-text"}]}
        """);
        var error = Assert.Throws<InvalidOperationException>(() => TemplateMigration.RunResolveSemanticCandidate([source, baseline, invalidCandidate]));
        Assert.Equal("template-migration-semantic-candidate-source-unknown-field:sourceObjectId", error.Message);
    }

    [Fact]
    public void TemplateMigration_projects_a_multi_run_source_value_into_one_target_run()
    {
        var source = CreateSemanticValueProjectionFixture(["Source caption: ", "AX", "-", "17"]);
        var baseline = CreateSemanticValueProjectionFixture(["Destination label: ", "{{currentValue}}"]);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v2",
            [],
            ValueProjections:
            [
                new TemplateMigrationSemanticCandidateValueProjection(
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Source caption: AX-17"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Destination label: {{currentValue}}"),
                    "controlled-identifier",
                    "token",
                    "after-first-delimiter")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        Assert.Equal("tiwater.docx.template-migration-plan/v4", resolved.Plan.Schema);
        var projection = Assert.Single(resolved.Plan.ValueProjections!);
        Assert.Equal("body:paragraph:0", projection.SourceParentObjectId);
        Assert.Equal("body:paragraph:0", projection.BaselineParentObjectId);

        var build = TemplateMigration.BuildOperations(source, baseline, resolved.Plan);
        var operation = Assert.Single(build.Operations);
        Assert.Equal("replaceParagraphRunText", operation.Type);
        Assert.Equal(1, operation.RunIndex);
        Assert.Equal("AX-17", operation.Text);

        var output = Path.Combine(Path.GetTempPath(), $"semantic-value-multi-single-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));
        Assert.Equal("Destination label: AX-17", ReadOnlyParagraphText(output));
    }

    [Fact]
    public void TemplateMigration_projects_one_source_run_across_a_split_target_placeholder()
    {
        var source = CreateSemanticValueProjectionFixture(["Origin: ", "2026-08-06"]);
        var baseline = CreateSemanticValueProjectionFixture(["Target wording: ", "{{", "effective", "Date}}"]);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v2",
            [],
            ValueProjections:
            [
                new TemplateMigrationSemanticCandidateValueProjection(
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Origin: 2026-08-06"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Target wording: {{effectiveDate}}"),
                    "effective-date",
                    "date",
                    "after-first-delimiter")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        var build = TemplateMigration.BuildOperations(source, baseline, resolved.Plan);
        Assert.Equal(3, build.Operations.Count);
        Assert.Equal(["2026-08-06", "", ""], build.Operations.Select(item => item.Text).ToArray());

        var output = Path.Combine(Path.GetTempPath(), $"semantic-value-single-multi-{Guid.NewGuid():N}.docx");
        Assert.True(TemplateMigration.Apply(source, baseline, resolved.Plan, output).Pass);
        Assert.Equal("Target wording: 2026-08-06", ReadOnlyParagraphText(output));
        Assert.True(TemplateMigration.ValidateReadback(source, baseline, output, resolved.Plan).Pass);
    }

    [Fact]
    public void TemplateMigration_selects_one_typed_value_group_without_touching_sibling_fields()
    {
        var source = CreateMultiFieldProjectionFixture(
            ["Source id: ", "ZX-44"],
            ["Revision caption: ", "0", "2"],
            ["Page: ", "1/8"]);
        var baseline = CreateMultiFieldProjectionFixture(
            ["Destination identifier: ", "OLD-1"],
            ["Version wording: ", "1", ".", "0"],
            ["Page: ", "1/6"]);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v2",
            [],
            ValueProjections:
            [
                new TemplateMigrationSemanticCandidateValueProjection(
                    new TemplateMigrationSemanticSelector("table-cell", "body", "Source id: ZX-44Revision caption: 02Page: 1/8"),
                    new TemplateMigrationSemanticSelector("table-cell", "body", "Destination identifier: OLD-1Version wording: 1.0Page: 1/6"),
                    "document-identifier",
                    "identifier",
                    "unique-delimited-run-group"),
                new TemplateMigrationSemanticCandidateValueProjection(
                    new TemplateMigrationSemanticSelector("table-cell", "body", "Source id: ZX-44Revision caption: 02Page: 1/8"),
                    new TemplateMigrationSemanticSelector("table-cell", "body", "Destination identifier: OLD-1Version wording: 1.0Page: 1/6"),
                    "revision-version",
                    "version",
                    "unique-delimited-run-group")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        var build = TemplateMigration.BuildOperations(source, baseline, resolved.Plan);
        Assert.Equal(["ZX-44", "02", "", ""], build.Operations.Select(item => item.Text ?? string.Empty).ToArray());
        Assert.Equal([0, 1, 1, 1], build.Operations.Select(operation => operation.ParagraphIndex).ToArray());

        var output = Path.Combine(Path.GetTempPath(), $"semantic-value-multi-field-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback?.Failures.Select(item => $"{item.Reason}:{item.Detail}") ?? []));
        using var document = WordprocessingDocument.Open(output, false);
        var cell = document.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().Single();
        Assert.Equal(
            ["Destination identifier: ZX-44", "Version wording: 02", "Page: 1/6"],
            cell.Elements<Paragraph>().Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(item => item.Text))).ToArray());
    }

    [Fact]
    public void TemplateMigration_projects_a_typed_value_from_one_paragraph_into_a_multi_field_table_cell()
    {
        var source = CreateSemanticValueProjectionFixture(["Revision: ", "0", "0", "\tPage: ", "1/1"]);
        var baseline = CreateMultiFieldProjectionFixture(
            ["Document No.: ", "OLD-1"],
            ["Version: ", "1", ".", "0"],
            ["Page: ", "1/17"]);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v2",
            [],
            ValueProjections:
            [
                new TemplateMigrationSemanticCandidateValueProjection(
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Revision: 00\tPage: 1/1"),
                    new TemplateMigrationSemanticSelector("table-cell", "body", "Document No.: OLD-1Version: 1.0Page: 1/17"),
                    "revision-version",
                    "version",
                    "unique-delimited-value")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        var output = Path.Combine(Path.GetTempPath(), $"semantic-value-cross-parent-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback?.Failures.Select(item => $"{item.Reason}:{item.Detail}") ?? []));
        using var document = WordprocessingDocument.Open(output, false);
        var cell = document.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().Single();
        Assert.Equal(
            ["Document No.: OLD-1", "Version: 00", "Page: 1/17"],
            cell.Elements<Paragraph>().Select(paragraph => string.Concat(paragraph.Descendants<Text>().Select(item => item.Text))).ToArray());
    }

    [Fact]
    public void TemplateMigration_semantic_value_projection_fails_closed_for_empty_ambiguous_duplicate_and_wrong_type_inputs()
    {
        var baseline = CreateSemanticValueProjectionFixture(["Target: ", "{{value}}"]);
        var empty = CreateSemanticValueProjectionFixture(["Source: ", "   "]);
        var emptyCandidate = SemanticValueCandidate("Source:", "Target: {{value}}", "text");
        var emptyResult = TemplateMigration.ResolveSemanticCandidate(empty, baseline, emptyCandidate);
        Assert.False(emptyResult.Pass);
        Assert.Contains(emptyResult.Unresolved, item => item.Reason == "template-migration-semantic-value-source-empty");

        var ambiguous = CreateSemanticValueProjectionFixture(["Source: ", "A"], duplicateParagraph: true);
        var ambiguousResult = TemplateMigration.ResolveSemanticCandidate(ambiguous, baseline, SemanticValueCandidate("Source: A", "Target: {{value}}", "token"));
        Assert.False(ambiguousResult.Pass);
        Assert.Contains(ambiguousResult.Unresolved, item => item.Reason == "template-migration-semantic-value-source-ambiguous");

        var source = CreateSemanticValueProjectionFixture(["Source: ", "A"]);
        var duplicateCandidate = SemanticValueCandidate("Source: A", "Target: {{value}}", "token") with
        {
            ValueProjections =
            [
                SemanticValueCandidate("Source: A", "Target: {{value}}", "token").ValueProjections!.Single(),
                SemanticValueCandidate("Source: A", "Target: {{value}}", "token").ValueProjections!.Single()
            ]
        };
        var duplicateResult = TemplateMigration.ResolveSemanticCandidate(source, baseline, duplicateCandidate);
        Assert.False(duplicateResult.Pass);
        Assert.Contains(duplicateResult.Unresolved, item => item.Reason == "template-migration-semantic-value-binding-duplicate");

        var wrongTarget = CreateSemanticValueProjectionFixture(["Target: ", "{{value}}"]);
        var wrongType = SemanticValueCandidate("Source: A", "Target: {{value}}", "token") with
        {
            ValueProjections =
            [
                SemanticValueCandidate("Source: A", "Target: {{value}}", "token").ValueProjections!.Single() with
                {
                    BaselineParent = new TemplateMigrationSemanticSelector("run", "body", "{{value}}", ParentText: "Target: {{value}}")
                }
            ]
        };
        var wrongTypeResult = TemplateMigration.ResolveSemanticCandidate(source, wrongTarget, wrongType);
        Assert.False(wrongTypeResult.Pass);
        Assert.Contains(wrongTypeResult.Unresolved, item => item.Reason == "template-migration-semantic-value-parent-kind-mismatch");

        var kindMismatch = TemplateMigration.ResolveSemanticCandidate(source, baseline, SemanticValueCandidate("Source: A", "Target: {{value}}", "date"));
        Assert.False(kindMismatch.Pass);
        Assert.Contains(kindMismatch.Unresolved, item => item.Reason == "template-migration-semantic-value-source-kind-mismatch");

        var emptyTarget = CreateSemanticValueProjectionFixture(["Target: ", "   "]);
        var emptyTargetResult = TemplateMigration.ResolveSemanticCandidate(source, emptyTarget, SemanticValueCandidate("Source: A", "Target:", "token"));
        Assert.False(emptyTargetResult.Pass);
        Assert.Contains(emptyTargetResult.Unresolved, item => item.Reason == "template-migration-semantic-value-baseline-empty");
    }

    [Fact]
    public void TemplateMigration_semantic_value_projection_binds_candidate_plan_hashes_and_independent_readback()
    {
        var source = CreateSemanticValueProjectionFixture(["Source: ", "R", "9"]);
        var baseline = CreateSemanticValueProjectionFixture(["Target: ", "old"]);
        Assert.Throws<InvalidOperationException>(() => TemplateMigration.ResolveSemanticCandidate(
            source,
            baseline,
            SemanticValueCandidate("Source: R9", "Target: old", "token") with { Schema = "tiwater.docx.template-migration-semantic-candidate/v4" }));

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, SemanticValueCandidate("Source: R9", "Target: old", "token"));
        var stale = resolved.Plan with { SourceSha256 = new string('0', 64) };
        Assert.Contains(TemplateMigration.BuildOperations(source, baseline, stale).Failures, item => item.Reason == "template-migration-source-hash-mismatch");
        var deleted = resolved.Plan with { ValueProjections = [] };
        Assert.Contains(TemplateMigration.BuildOperations(source, baseline, deleted).Failures, item => item.Reason == "template-migration-plan-v4-value-projection-required");
        var changedType = resolved.Plan with
        {
            ValueProjections = [resolved.Plan.ValueProjections!.Single() with { ValueKind = "date" }]
        };
        Assert.Contains(TemplateMigration.BuildOperations(source, baseline, changedType).Failures, item => item.Reason == "template-migration-semantic-value-source-kind-mismatch");

        var output = Path.Combine(Path.GetTempPath(), $"semantic-value-readback-{Guid.NewGuid():N}.docx");
        Assert.True(TemplateMigration.Apply(source, baseline, resolved.Plan, output).Pass);
        var tampered = Path.Combine(Path.GetTempPath(), $"semantic-value-tampered-{Guid.NewGuid():N}.docx");
        Editor.Apply(output, tampered, [new DocxEditOperation("replaceParagraphRunText", ParagraphIndex: 0, RunIndex: 1, Text: "tampered")]);
        var validation = TemplateMigration.ValidateReadback(source, baseline, tampered, resolved.Plan);
        Assert.False(validation.Pass);
        Assert.Contains(validation.Failures, item => item.Reason == "template-migration-readback-semantic-value-mismatch");
    }

    [Fact]
    public void TemplateMigration_inserts_a_contiguous_source_range_between_unique_target_anchors_with_target_context_style()
    {
        var source = CreateStyledTextMigrationFixture(
            ("stable before", "Source"),
            ("new first", "Source"),
            ("new second", "Source"),
            ("stable after", "Source"));
        var baseline = CreateStyledTextMigrationFixture(
            ("stable before", "Before"),
            ("stable after", "After"));
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v3",
            [],
            BodyInsertions:
            [
                new TemplateMigrationSemanticCandidateBodyInsertion(
                    new TemplateMigrationSemanticSelector("paragraph", "body", "new first"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "new second"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "stable before"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "stable after"),
                    "target-after-context")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        Assert.Equal("tiwater.docx.template-migration-plan/v5", resolved.Plan.Schema);
        var output = Path.Combine(Path.GetTempPath(), $"migration-body-insertion-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback?.Failures.Select(item => item.Reason) ?? []));
        using var document = WordprocessingDocument.Open(output, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        Assert.Equal(["stable before", "new first", "new second", "stable after"], paragraphs.Select(item => item.InnerText).ToArray());
        Assert.Equal(["Before", "After", "After", "After"], paragraphs.Select(item => item.ParagraphProperties?.ParagraphStyleId?.Val?.Value).ToArray());

        var tampered = Path.Combine(Path.GetTempPath(), $"migration-body-insertion-tampered-{Guid.NewGuid():N}.docx");
        Editor.Apply(output, tampered, [new DocxEditOperation("replaceParagraphText", ParagraphIndex: 1, Text: "changed")]);
        Assert.Contains(TemplateMigration.ValidateReadback(source, baseline, tampered, resolved.Plan).Failures, item => item.Reason == "template-migration-readback-body-insertion-content-mismatch");
        Assert.Contains(TemplateMigration.BuildOperations(source, baseline, resolved.Plan with { BodyInsertions = [] }).Failures, item => item.Reason == "template-migration-plan-v5-body-insertion-required");
    }

    [Fact]
    public void TemplateMigration_body_insertion_fails_closed_for_ambiguous_or_non_adjacent_target_anchors()
    {
        var source = CreateTextMigrationFixture("before", "source addition", "after");
        var ambiguous = CreateTextMigrationFixture("before", "after", "after");
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v3",
            [],
            BodyInsertions:
            [new TemplateMigrationSemanticCandidateBodyInsertion(
                new TemplateMigrationSemanticSelector("paragraph", "body", "source addition"),
                new TemplateMigrationSemanticSelector("paragraph", "body", "source addition"),
                new TemplateMigrationSemanticSelector("paragraph", "body", "before"),
                new TemplateMigrationSemanticSelector("paragraph", "body", "after"),
                "target-after-context")]);
        var ambiguousResult = TemplateMigration.ResolveSemanticCandidate(source, ambiguous, candidate);
        Assert.False(ambiguousResult.Pass);
        Assert.Contains(ambiguousResult.Unresolved, item => item.Reason == "template-migration-semantic-body-insertion-anchor-not-unique");

        var nonAdjacent = CreateTextMigrationFixture("before", "target-owned", "after");
        var nonAdjacentResult = TemplateMigration.ResolveSemanticCandidate(source, nonAdjacent, candidate);
        Assert.False(nonAdjacentResult.Pass);
        Assert.Contains(nonAdjacentResult.Unresolved, item => item.Reason == "template-migration-semantic-body-insertion-range-invalid");
    }

    [Fact]
    public void TemplateMigration_resolves_an_explicit_source_exclusion_without_a_fake_target()
    {
        var source = CreateTextMigrationFixture("obsolete source section");
        var baseline = CreateTextMigrationFixture("current target section");
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("paragraph", Scope: "body", Text: "obsolete source section"),
                    null,
                    "out-of-scope")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);

        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        var mapping = Assert.Single(resolved.Plan.Mappings);
        Assert.Equal("out-of-scope", mapping.Disposition);
        Assert.Null(mapping.BaselineObjectId);
        Assert.Equal("semantic-candidate-out-of-scope", mapping.Reason);

        var missingTarget = candidate with
        {
            Mappings =
            [
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("paragraph", Scope: "body", Text: "obsolete source section"),
                    null,
                    "copy-text")
            ]
        };
        var error = Assert.Throws<InvalidOperationException>(() => TemplateMigration.ResolveSemanticCandidate(source, baseline, missingTarget));
        Assert.Equal("template-migration-semantic-candidate-baseline-missing", error.Message);
    }

    [Fact]
    public void TemplateMigration_pairs_only_equal_paragraph_gaps_between_unique_text_anchors()
    {
        var source = CreateTextMigrationFixture("anchor start", "legacy heading", "anchor end");
        var baseline = CreateTextMigrationFixture("anchor start", "target heading", "anchor end");

        var paired = TemplateMigration.DeriveAnchorGapPlan(source, baseline);

        Assert.False(paired.Pass);
        Assert.Contains(paired.Plan.Mappings, item => item.SourceObjectId == "body:paragraph:1" && item.Disposition == "review-required");
        Assert.Contains(paired.Unresolved, item => item.Reason == "template-migration-anchor-gap-candidate-review-required" && item.SourceObjectId == "body:paragraph:1" && item.BaselineObjectId == "body:paragraph:1");

        var unequalSource = CreateTextMigrationFixture("anchor start", "legacy heading one", "legacy heading two", "anchor end");
        var unequal = TemplateMigration.DeriveAnchorGapPlan(unequalSource, baseline);
        Assert.False(unequal.Pass);
        Assert.Contains(unequal.Plan.Mappings, item => item.SourceObjectId == "body:paragraph:1" && item.Disposition == "review-required");
        Assert.Contains(unequal.Plan.Mappings, item => item.SourceObjectId == "body:paragraph:2" && item.Disposition == "review-required");
    }

    [Fact]
    public void TemplateMigration_preview_emits_a_verified_review_candidate_without_claiming_pass()
    {
        var source = CreateTextMigrationFixture("shared verified content", "source content pending review");
        var baseline = CreateTextMigrationFixture("shared verified content", "target-owned format label");
        var plan = TemplateMigration.DeriveExactTextPlan(source, baseline).Plan;
        var output = Path.Combine(Path.GetTempPath(), $"migration-review-preview-{Guid.NewGuid():N}.docx");

        var preview = TemplateMigration.Preview(source, baseline, plan, output);

        Assert.False(preview.Pass);
        Assert.True(preview.ReviewRequired);
        Assert.True(preview.OutputVerified);
        Assert.Equal(output, preview.Output);
        Assert.True(File.Exists(output));
        using var document = WordprocessingDocument.Open(output, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToList();
        Assert.Equal("shared verified content", GetParagraphText(paragraphs[0]));
        Assert.Equal("target-owned format label", GetParagraphText(paragraphs[1]));
    }

    [Fact]
    public void TemplateMigration_preview_accepts_a_legacy_baseline_only_when_it_introduces_no_new_openxml_errors()
    {
        var source = CreateTextMigrationFixture("verified source value");
        var baseline = CreateTextMigrationFixture("baseline placeholder");
        ReplaceZipEntry(
            baseline,
            "word/document.xml",
            ReadZipEntry(baseline, "word/document.xml").Replace(
                "<w:p>",
                "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"-1\"/></w:numPr></w:pPr>",
                StringComparison.Ordinal));
        var analysis = TemplateMigration.Analyze(source, baseline);
        var plan = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v1",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            [new TemplateMigrationMapping("body:paragraph:0", "body:paragraph:0", "copy-text")]);
        var output = Path.Combine(Path.GetTempPath(), $"migration-legacy-openxml-{Guid.NewGuid():N}.docx");

        var preview = TemplateMigration.Preview(source, baseline, plan, output);

        Assert.NotEqual(0, OpenXmlValidation.Run([baseline]));
        Assert.True(preview.OutputVerified, string.Join("; ", preview.Readback!.Failures.Select(item => item.Reason)));
        Assert.True(preview.Pass);
        Assert.True(File.Exists(output));
        Assert.Equal(1, OpenXmlValidation.Run([output]));
    }

    [Fact]
    public void TemplateMigration_semantic_selectors_migrate_shifted_body_header_and_footer_content_without_coordinates()
    {
        var source = CreateCrossTemplateMigrationFixture("source header", "source opening", "source fact", "source closing", "source footer", false);
        var baseline = CreateCrossTemplateMigrationFixture("target header", "target opening", "target fact", "target closing", "target footer", true);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [
                new TemplateMigrationSemanticCandidateMapping(new TemplateMigrationSemanticSelector("paragraph", "header", "source header"), new TemplateMigrationSemanticSelector("paragraph", "header", "target header"), "copy-text"),
                new TemplateMigrationSemanticCandidateMapping(new TemplateMigrationSemanticSelector("paragraph", "body", "source opening"), new TemplateMigrationSemanticSelector("paragraph", "body", "target opening"), "copy-text"),
                new TemplateMigrationSemanticCandidateMapping(new TemplateMigrationSemanticSelector("table-cell", "body", "source fact"), new TemplateMigrationSemanticSelector("table-cell", "body", "target fact"), "copy-text"),
                new TemplateMigrationSemanticCandidateMapping(new TemplateMigrationSemanticSelector("paragraph", "body", "source closing"), new TemplateMigrationSemanticSelector("paragraph", "body", "target closing"), "copy-text"),
                new TemplateMigrationSemanticCandidateMapping(new TemplateMigrationSemanticSelector("table-cell", "footer", "source footer"), new TemplateMigrationSemanticSelector("table-cell", "footer", "target footer"), "copy-text")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        var output = Path.Combine(Path.GetTempPath(), $"migration-shifted-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);

        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Contains("source header", document.MainDocumentPart!.HeaderParts.Single().Header!.InnerText);
        Assert.Contains("source fact", document.MainDocumentPart.Document!.Body!.InnerText);
        Assert.Contains("source footer", document.MainDocumentPart.FooterParts.Single().Footer!.InnerText);
    }

    [Fact]
    public void TemplateMigration_builds_hash_bound_operations_only_from_complete_declared_mapping()
    {
        var source = CreatePlainMigrationFixture();
        var baseline = Path.Combine(Path.GetTempPath(), $"migration-plan-baseline-{Guid.NewGuid():N}.docx");
        Editor.Apply(source, baseline, [
            new DocxEditOperation("replaceParagraphText", ParagraphIndex: 0, Text: "Target heading"),
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 1, Text: "Target cell")
        ]);
        var analysis = TemplateMigration.Analyze(source, baseline);
        var mappings = analysis.Source.Objects
            .Where(item => (item.Kind == "paragraph" || item.Kind == "table-cell") && !string.IsNullOrWhiteSpace(item.Text))
            .Select(item => new TemplateMigrationMapping(item.Id, item.Id, "copy-text"))
            .ToList();
        var plan = new TemplateMigrationPlan("tiwater.docx.template-migration-plan/v1", analysis.Source.Sha256, analysis.Baseline.Sha256, mappings);

        var result = TemplateMigration.BuildOperations(source, baseline, plan);

        Assert.True(result.Pass, string.Join("; ", result.Failures.Select(item => item.Reason)));
        Assert.False(result.ReviewRequired);
        Assert.NotNull(result.OperationsSha256);
        Assert.Equal(mappings.Count, result.Operations.Count);
        Assert.Contains(result.Operations, operation => operation.Type == "replaceParagraphText" && operation.ParagraphIndex == 0 && operation.Text == "Project code XXXX 峰面积");
        Assert.Contains(result.Operations, operation => operation.Type == "replaceTableCellText" && operation.TableIndex == 0 && operation.RowIndex == 0 && operation.CellIndex == 1 && operation.Text == "Batch YYYY");

        var incomplete = plan with { Mappings = mappings.Skip(1).ToList() };
        var rejected = TemplateMigration.BuildOperations(source, baseline, incomplete);
        Assert.False(rejected.Pass);
        Assert.Contains(rejected.Failures, item => item.Reason == "template-migration-source-object-unmapped");
        Assert.Empty(rejected.Operations);

        var blockedOutput = Path.Combine(Path.GetTempPath(), $"migration-blocked-{Guid.NewGuid():N}.docx");
        var blockedApply = TemplateMigration.Apply(source, baseline, incomplete, blockedOutput);
        Assert.False(blockedApply.Pass);
        Assert.Null(blockedApply.Output);
        Assert.False(File.Exists(blockedOutput));

        var stale = plan with { SourceSha256 = new string('0', 64) };
        var staleRejected = TemplateMigration.BuildOperations(source, baseline, stale);
        Assert.False(staleRejected.Pass);
        Assert.Contains(staleRejected.Failures, item => item.Reason == "template-migration-source-hash-mismatch");
        Assert.Empty(staleRejected.Operations);

        var duplicate = plan with { Mappings = [mappings[0], mappings[0], .. mappings.Skip(1)] };
        var duplicateRejected = TemplateMigration.BuildOperations(source, baseline, duplicate);
        Assert.False(duplicateRejected.Pass);
        Assert.Contains(duplicateRejected.Failures, item => item.Reason == "template-migration-source-object-duplicate");
        Assert.Empty(duplicateRejected.Operations);
    }

    [Fact]
    public void TemplateMigration_copies_attested_run_content_without_replacing_target_owned_label()
    {
        var source = CreateLabeledRunMigrationFixture("Legacy document no.: ", "SOP03-5-0014");
        var baseline = CreateLabeledRunMigrationFixture("Document No.: ", "PLACEHOLDER");
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Legacy document no.: SOP03-5-0014"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Document No.: PLACEHOLDER"),
                    "retain-target"),
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("run", "body", "SOP03-5-0014", ParentText: "Legacy document no.: SOP03-5-0014"),
                    new TemplateMigrationSemanticSelector("run", "body", "PLACEHOLDER", ParentText: "Document No.: PLACEHOLDER"),
                    "copy-text")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        Assert.Contains(resolved.Plan.Mappings, item => item.Disposition == "retain-target" && item.SourceObjectId == "body:paragraph:0" && item.BaselineObjectId == "body:paragraph:0");
        var missingFactRun = resolved.Plan with { Mappings = resolved.Plan.Mappings.Where(item => item.Disposition != "copy-text").ToList() };
        var rejected = TemplateMigration.BuildOperations(source, baseline, missingFactRun);
        Assert.Contains(rejected.Failures, item => item.Reason == "template-migration-retain-target-fact-run-required");
        var build = TemplateMigration.BuildOperations(source, baseline, resolved.Plan);
        Assert.True(build.Pass, string.Join("; ", build.Failures.Select(item => item.Reason)));
        var operation = Assert.Single(build.Operations);
        Assert.Equal("replaceParagraphRunText", operation.Type);
        Assert.Equal(1, operation.RunIndex);

        var output = Path.Combine(Path.GetTempPath(), $"migration-run-output-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));
        using (var document = WordprocessingDocument.Open(output, false))
        {
            var runs = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().Elements<Run>().ToList();
            Assert.Equal("Document No.: ", string.Concat(runs[0].Descendants<Text>().Select(text => text.Text)));
            Assert.Equal("SOP03-5-0014", string.Concat(runs[1].Descendants<Text>().Select(text => text.Text)));
        }
        using (var document = WordprocessingDocument.Open(output, true))
        {
            document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().Elements<Run>().First().GetFirstChild<Text>()!.Text = "Tampered label: ";
            document.MainDocumentPart.Document.Save();
        }
        var tampered = TemplateMigration.ValidateReadback(source, baseline, output, resolved.Plan);
        Assert.Contains(tampered.Failures, item => item.Reason == "template-migration-readback-retained-target-run-mismatch");
    }

    [Fact]
    public void TemplateMigration_retains_an_explicitly_selected_target_label_without_emitting_an_edit()
    {
        var source = CreateTextMigrationFixture("Legacy purpose label");
        var baseline = CreateTextMigrationFixture("Objective:");
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [
                new TemplateMigrationSemanticCandidateMapping(
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Legacy purpose label"),
                    new TemplateMigrationSemanticSelector("paragraph", "body", "Objective:"),
                    "retain-target-label")
            ]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        Assert.Contains(resolved.Plan.Mappings, item => item.Disposition == "retain-target-label" && item.Reason == "semantic-candidate-retain-target-label");
        var build = TemplateMigration.BuildOperations(source, baseline, resolved.Plan);
        Assert.True(build.Pass, string.Join("; ", build.Failures.Select(item => item.Reason)));
        Assert.Empty(build.Operations);

        var output = Path.Combine(Path.GetTempPath(), $"migration-label-output-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));
        using (var document = WordprocessingDocument.Open(output, false))
        {
            Assert.Equal("Objective:", string.Concat(document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().Descendants<Text>().Select(text => text.Text)));
        }
        using (var document = WordprocessingDocument.Open(output, true))
        {
            document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text = "Tampered label:";
            document.MainDocumentPart.Document.Save();
        }
        var tampered = TemplateMigration.ValidateReadback(source, baseline, output, resolved.Plan);
        Assert.Contains(tampered.Failures, item => item.Reason == "template-migration-readback-retained-target-run-mismatch");
    }

    [Fact]
    public void TemplateMigration_output_validator_rebuilds_authority_and_rejects_tampering_without_apply_result()
    {
        var source = CreateTextMigrationFixture("source fact"); var baseline = CreateTextMigrationFixture("target slot");
        var analysis = TemplateMigration.Analyze(source, baseline);
        var plan = new TemplateMigrationPlan("tiwater.docx.template-migration-plan/v1", analysis.Source.Sha256, analysis.Baseline.Sha256,
            [new TemplateMigrationMapping("body:paragraph:0", "body:paragraph:0", "copy-text")]);
        var planPath = Path.Combine(Path.GetTempPath(), $"migration-plan-{Guid.NewGuid():N}.json");
        File.WriteAllText(planPath, System.Text.Json.JsonSerializer.Serialize(plan, Json.Options));
        var output = Path.Combine(Path.GetTempPath(), $"migration-validated-{Guid.NewGuid():N}.docx");
        Assert.True(TemplateMigration.Apply(source, baseline, plan, output).Pass);

        var valid = TemplateMigration.ValidateOutput(source, baseline, planPath, output, plan);
        Assert.True(valid.Pass); Assert.Equal("tiwater.docx.template-migration-output-validation/v1", valid.Schema); Assert.Matches("^[A-F0-9]{64}$", valid.OutputSha256);

        using (var document = WordprocessingDocument.Open(output, true))
        {
            document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().Single().GetFirstChild<Run>()!.GetFirstChild<Text>()!.Text = "tampered";
            document.MainDocumentPart.Document.Save();
        }
        var tampered = TemplateMigration.ValidateOutput(source, baseline, planPath, output, plan);
        Assert.False(tampered.Pass); Assert.Contains(tampered.Failures, item => item.Reason == "template-migration-readback-content-mismatch");
    }

    [Fact]
    public void TemplateMigration_appends_a_semantically_selected_body_range_without_coordinates()
    {
        var source = CreateBodyAppendFixture(includeDuplicateRevisionTable: false, baseline: false);
        var baseline = CreateBodyAppendFixture(includeDuplicateRevisionTable: false, baseline: true);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [],
            [new TemplateMigrationSemanticCandidateBodyAppend(
                new TemplateMigrationSemanticSelector("paragraph", "body", "Revision history"),
                new TemplateMigrationSemanticSelector("table", "body", DescendantText: "Revision No."))]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.True(resolved.Pass, string.Join("; ", resolved.Unresolved.Select(item => item.Reason)));
        Assert.Equal("tiwater.docx.template-migration-plan/v2", resolved.Plan.Schema);
        Assert.Single(resolved.Plan.BodyAppends!);

        var output = Path.Combine(Path.GetTempPath(), $"migration-body-append-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, resolved.Plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => item.Reason)));
        using var document = WordprocessingDocument.Open(output, false);
        var body = document.MainDocumentPart!.Document!.Body!;
        Assert.Equal(["before", "after", "Revision history"], body.Elements<Paragraph>().Select(item => item.InnerText).ToArray());
        var table = Assert.Single(body.Elements<Table>());
        Assert.Contains("Revision No.", table.InnerText);
        Assert.Contains("R1", table.InnerText);
    }

    [Fact]
    public void TemplateMigration_preview_clears_attested_baseline_placeholders_and_remains_review_required()
    {
        var source = CreateTextMigrationFixture("source fact");
        var baseline = CreateBaselineClearFixture("{{approval}}", "{{effectiveDate}}");
        var analysis = TemplateMigration.Analyze(source, baseline);
        var plan = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v3",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            [new TemplateMigrationMapping("body:paragraph:0", null, "review-required", "source-layout-not-representable")],
            BaselineClears:
            [
                new TemplateMigrationBaselineClear("body:table:0:row:0:cell:0", "cell"),
                new TemplateMigrationBaselineClear("body:table:0:row:1:cell:0", "row")
            ]);
        var output = Path.Combine(Path.GetTempPath(), $"migration-review-preview-{Guid.NewGuid():N}.docx");

        var preview = TemplateMigration.Preview(source, baseline, plan, output);

        Assert.False(preview.Pass);
        Assert.True(preview.ReviewRequired);
        Assert.True(preview.OutputVerified, string.Join("; ", preview.Readback?.Failures.Select(item => item.Reason) ?? []));
        using var document = WordprocessingDocument.Open(output, false);
        var rows = document.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(string.Empty, rows[0].Elements<TableCell>().First().InnerText);
        Assert.All(rows[1].Elements<TableCell>(), cell => Assert.Equal(string.Empty, cell.InnerText));
    }

    [Fact]
    public void TemplateMigration_rejects_unbound_or_conflicting_baseline_clear()
    {
        var source = CreateTextMigrationFixture("source fact");
        var baseline = CreateBaselineClearFixture("target", "other");
        var analysis = TemplateMigration.Analyze(source, baseline);
        var unknown = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v3",
            analysis.Source.Sha256,
            analysis.Baseline.Sha256,
            [],
            BaselineClears: [new TemplateMigrationBaselineClear("body:table:9:row:9:cell:9", "cell")]);
        Assert.Contains(TemplateMigration.BuildOperations(source, baseline, unknown).Failures,
            failure => failure.Reason == "template-migration-baseline-clear-object-invalid");

        var conflictSource = CreateBaselineClearFixture("source fact", "other");
        var conflictAnalysis = TemplateMigration.Analyze(conflictSource, baseline);
        var conflict = new TemplateMigrationPlan(
            "tiwater.docx.template-migration-plan/v3",
            conflictAnalysis.Source.Sha256,
            conflictAnalysis.Baseline.Sha256,
            [new TemplateMigrationMapping("body:table:0:row:0:cell:0", "body:table:0:row:0:cell:0", "copy-text")],
            BaselineClears: [new TemplateMigrationBaselineClear("body:table:0:row:0:cell:0", "cell")]);
        Assert.Contains(TemplateMigration.BuildOperations(conflictSource, baseline, conflict).Failures,
            failure => failure.Reason == "template-migration-baseline-clear-copy-conflict");
    }

    [Fact]
    public void TemplateMigration_rejects_an_ambiguous_semantic_body_append_selector()
    {
        var source = CreateBodyAppendFixture(includeDuplicateRevisionTable: true, baseline: false);
        var baseline = CreateBodyAppendFixture(includeDuplicateRevisionTable: false, baseline: true);
        var candidate = new TemplateMigrationSemanticCandidate(
            "tiwater.docx.template-migration-semantic-candidate/v1",
            [],
            [new TemplateMigrationSemanticCandidateBodyAppend(
                new TemplateMigrationSemanticSelector("paragraph", "body", "Revision history"),
                new TemplateMigrationSemanticSelector("table", "body", DescendantText: "Revision No."))]);

        var resolved = TemplateMigration.ResolveSemanticCandidate(source, baseline, candidate);
        Assert.False(resolved.Pass);
        Assert.Contains(resolved.Unresolved, item => item.Reason == "template-migration-semantic-append-end-ambiguous");
    }

    [Fact]
    public void TemplateMigration_and_editor_support_header_and_footer_table_cells()
    {
        var source = CreateHeaderFooterTableFixture("source-header", "source-footer");
        var baseline = CreateHeaderFooterTableFixture("baseline-header", "baseline-footer");
        var analysis = TemplateMigration.Analyze(source, baseline);
        var mappings = analysis.Source.Objects
            .Where(item => (item.Kind == "paragraph" || item.Kind == "table-cell" || item.Kind == "run") && !string.IsNullOrWhiteSpace(item.Text))
            .Select(item => new TemplateMigrationMapping(item.Id, item.Id, "copy-text"))
            .ToList();
        var plan = new TemplateMigrationPlan("tiwater.docx.template-migration-plan/v1", analysis.Source.Sha256, analysis.Baseline.Sha256, mappings);

        var build = TemplateMigration.BuildOperations(source, baseline, plan);
        Assert.True(build.Pass, string.Join("; ", build.Failures.Select(item => item.Reason)));
        Assert.Contains(build.Operations, item => item.Type == "replaceHeaderTableCellText" && item.Text == "source-header");
        Assert.Contains(build.Operations, item => item.Type == "replaceFooterTableCellText" && item.Text == "source-footer");
        Assert.Contains(build.Operations, item => item.Type == "replaceHeaderParagraphRunText" && item.Text == "header-paragraph-source-header");
        Assert.Contains(build.Operations, item => item.Type == "replaceFooterParagraphRunText" && item.Text == "footer-paragraph-source-footer");
        Assert.Contains(build.Operations, item => item.Type == "replaceHeaderTableCellRunText" && item.Text == "source-header");
        Assert.Contains(build.Operations, item => item.Type == "replaceFooterTableCellRunText" && item.Text == "source-footer");

        var output = Path.Combine(Path.GetTempPath(), $"migration-header-footer-{Guid.NewGuid():N}.docx");
        var applied = TemplateMigration.Apply(source, baseline, plan, output);
        Assert.True(applied.Pass, string.Join("; ", applied.Readback!.Failures.Select(item => $"{item.Reason}:{item.Detail}")));
        Assert.All(applied.Edit!.AppliedOperations, item => Assert.True(item.Applied, item.Detail));
        using var document = WordprocessingDocument.Open(output, false);
        Assert.Contains("source-header", document.MainDocumentPart!.HeaderParts.Single().Header!.Descendants<Text>().Select(item => item.Text));
        Assert.Contains("source-footer", document.MainDocumentPart.FooterParts.Single().Footer!.Descendants<Text>().Select(item => item.Text));
    }

    [Fact]
    public void TemplateMigration_exact_text_derivation_maps_only_unique_same_kind_content()
    {
        var source = CreateExactTextMappingFixture(includeDuplicateBaselineText: false, baseline: false);
        var baseline = CreateExactTextMappingFixture(includeDuplicateBaselineText: false, baseline: true);

        var derived = TemplateMigration.DeriveExactTextPlan(source, baseline);

        Assert.True(derived.Pass, string.Join("; ", derived.Unresolved.Select(item => item.Reason)));
        Assert.All(derived.Plan.Mappings, mapping => Assert.Equal("copy-text", mapping.Disposition));
        Assert.Contains(derived.Plan.Mappings, mapping => mapping.SourceObjectId == "body:paragraph:0" && mapping.BaselineObjectId == "body:paragraph:1");
        Assert.Contains(derived.Plan.Mappings, mapping => mapping.SourceObjectId == "body:table:0:row:0:cell:0" && mapping.BaselineObjectId == "body:table:0:row:0:cell:1");

        var duplicateBaseline = CreateExactTextMappingFixture(includeDuplicateBaselineText: true, baseline: true);
        var rejected = TemplateMigration.DeriveExactTextPlan(source, duplicateBaseline);
        Assert.False(rejected.Pass);
        Assert.Contains(rejected.Unresolved, item => item.Reason == "template-migration-exact-text-ambiguous");
    }

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
    public void Edit_preserves_table_cell_paragraph_properties_when_replacing_content()
    {
        var docPath = CreateAnnotatedFixture();
        var output = Path.Combine(Path.GetTempPath(), $"edited-cell-paragraph-{Guid.NewGuid():N}.docx");

        using (var doc = WordprocessingDocument.Open(docPath, true))
        {
            var cellParagraph = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
                .Elements<TableRow>().Single()
                .Elements<TableCell>().ElementAt(1)
                .Elements<Paragraph>().First();
            cellParagraph.ParagraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { Before = "0", After = "0" });
        }

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("replaceTableCellText", TableIndex: 0, RowIndex: 0, CellIndex: 1, Text: "Batch HSP001-01")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var replacedParagraph = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().Single()
            .Elements<TableCell>().ElementAt(1)
            .Elements<Paragraph>().Single();

        var properties = replacedParagraph.ParagraphProperties;
        Assert.NotNull(properties);
        Assert.Equal(JustificationValues.Center, properties!.GetFirstChild<Justification>()!.Val!.Value);
        var spacing = properties.GetFirstChild<SpacingBetweenLines>();
        Assert.NotNull(spacing);
        Assert.Equal("0", spacing!.Before!.Value);
        Assert.Equal("0", spacing.After!.Value);
        Assert.Contains("Batch HSP001-01", GetParagraphText(replacedParagraph));
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
    public void Edit_insert_table_rows_inherits_complete_style_from_text_bearing_template_run()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-insert-row-complete-style-{Guid.NewGuid():N}.docx");
        CreateInsertRowStyleFixture(path);
        var output = Path.Combine(Path.GetTempPath(), $"insert-row-complete-style-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "insertTableRows",
                TableIndex: 0,
                RowIndex: 3,
                TemplateRowIndex: 1,
                Rows: [[
                    new DocxTableCellInput("alpha"),
                    new DocxTableCellInput(RichText: [new DocxRichTextSegment("check")])
                ]])
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var inserted = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ElementAt(3);
        AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(0), "DDEBF7", JustificationValues.Center, "Aptos", "22", "1F4E78", UnderlineValues.Single);
        AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(1), "E2F0D9", JustificationValues.Right, "Aptos", "22", "1F4E78", UnderlineValues.Single);
    }

    [Fact]
    public void Edit_insert_table_rows_honors_template_index_mutation_for_every_inserted_row_and_segment()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-insert-row-template-mutation-{Guid.NewGuid():N}.docx");
        CreateInsertRowStyleFixture(path);
        var output = Path.Combine(Path.GetTempPath(), $"insert-row-template-mutation-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "insertTableRows",
                TableIndex: 0,
                RowIndex: 3,
                TemplateRowIndex: 2,
                Rows: [
                    [
                        new DocxTableCellInput("first"),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("one"), new DocxRichTextSegment("two")])
                    ],
                    [
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("second")]),
                        new DocxTableCellInput("third")
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var rows = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(6, rows.Count);
        foreach (var inserted in rows.Skip(3).Take(2))
        {
            AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(0), "FCE4D6", JustificationValues.Left, "Courier New", "28", "C00000", UnderlineValues.Double);
            AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(1), "FFF2CC", JustificationValues.Both, "Courier New", "28", "C00000", UnderlineValues.Double);
        }
    }

    [Fact]
    public void Edit_insert_empty_table_rows_preserves_paragraph_mark_and_empty_run_style_for_later_values()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-insert-empty-row-style-{Guid.NewGuid():N}.docx");
        CreateEmptyInsertRowStyleFixture(path);
        var output = Path.Combine(Path.GetTempPath(), $"insert-empty-row-style-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "insertTableRows",
                TableIndex: 0,
                RowIndex: 3,
                TemplateRowIndex: 1,
                Rows: [[new DocxTableCellInput(), new DocxTableCellInput()]]),
            new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: 3,
                CellIndex: 0,
                RichText: [new DocxRichTextSegment("alpha"), new DocxRichTextSegment("beta")]),
            new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: 3,
                CellIndex: 1,
                RichText: [new DocxRichTextSegment("check")])
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var inserted = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ElementAt(3);
        AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(0), "DDEBF7", JustificationValues.Center, "Aptos", "22", "1F4E78", UnderlineValues.Single);
        AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(1), "E2F0D9", JustificationValues.Right, "Aptos", "22", "1F4E78", UnderlineValues.Single);
    }

    [Fact]
    public void Edit_insert_empty_table_rows_honors_template_index_mutation_across_multiple_rows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-insert-empty-row-mutation-{Guid.NewGuid():N}.docx");
        CreateEmptyInsertRowStyleFixture(path);
        var output = Path.Combine(Path.GetTempPath(), $"insert-empty-row-mutation-{Guid.NewGuid():N}.docx");

        var operations = new List<DocxEditOperation>
        {
            new(
                "insertTableRows",
                TableIndex: 0,
                RowIndex: 3,
                TemplateRowIndex: 2,
                Rows: [
                    [new DocxTableCellInput(), new DocxTableCellInput()],
                    [new DocxTableCellInput(), new DocxTableCellInput()]
                ])
        };
        foreach (var rowIndex in new[] { 3, 4 })
        {
            operations.Add(new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: rowIndex,
                CellIndex: 0,
                RichText: [new DocxRichTextSegment($"row-{rowIndex}-left")]));
            operations.Add(new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: rowIndex,
                CellIndex: 1,
                RichText: [new DocxRichTextSegment("one"), new DocxRichTextSegment("two")]));
        }

        var result = Editor.Apply(path, output, operations);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var rows = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal(6, rows.Count);
        foreach (var inserted in rows.Skip(3).Take(2))
        {
            AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(0), "FCE4D6", JustificationValues.Left, "Courier New", "28", "C00000", UnderlineValues.Double);
            AssertInheritedCellStyle(inserted.Elements<TableCell>().ElementAt(1), "FFF2CC", JustificationValues.Both, "Courier New", "28", "C00000", UnderlineValues.Double);
        }
    }

    private static void CreateEmptyInsertRowStyleFixture(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Table(
                new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                new TableGrid(new GridColumn { Width = "2500" }, new GridColumn { Width = "2500" }),
                new TableRow(
                    CreateInsertRowStyleCell("header-a", "D9EAD3", JustificationValues.Center, "Aptos", "20", "000000", UnderlineValues.None),
                    CreateInsertRowStyleCell("header-b", "D9EAD3", JustificationValues.Center, "Aptos", "20", "000000", UnderlineValues.None)),
                new TableRow(
                    CreateParagraphMarkOnlyStyleCell("DDEBF7", JustificationValues.Center, "Aptos", "22", "1F4E78", UnderlineValues.Single),
                    CreateEmptyRunStyleCell("E2F0D9", JustificationValues.Right, "Aptos", "22", "1F4E78", UnderlineValues.Single)),
                new TableRow(
                    CreateParagraphMarkOnlyStyleCell("FCE4D6", JustificationValues.Left, "Courier New", "28", "C00000", UnderlineValues.Double),
                    CreateEmptyRunStyleCell("FFF2CC", JustificationValues.Both, "Courier New", "28", "C00000", UnderlineValues.Double)),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("footer-a")))),
                    new TableCell(new Paragraph(new Run(new Text("footer-b"))))))));
        mainPart.Document.Save();
    }

    private static TableCell CreateParagraphMarkOnlyStyleCell(
        string fill,
        JustificationValues justification,
        string font,
        string fontSize,
        string color,
        UnderlineValues underline)
        => new(
            CreateInsertRowCellProperties(fill),
            new Paragraph(new ParagraphProperties(
                new Justification { Val = justification },
                new SpacingBetweenLines { Before = "40", After = "80" },
                new KeepNext(),
                new ParagraphMarkRunProperties(CreateInsertRowRunStyle(font, fontSize, color, underline).ChildElements.Select(element => element.CloneNode(true))))));

    private static TableCell CreateEmptyRunStyleCell(
        string fill,
        JustificationValues justification,
        string font,
        string fontSize,
        string color,
        UnderlineValues underline)
        => new(
            CreateInsertRowCellProperties(fill),
            new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = justification },
                    new SpacingBetweenLines { Before = "40", After = "80" },
                    new KeepNext()),
                new Run(CreateInsertRowRunStyle(font, fontSize, color, underline))));

    private static TableCellProperties CreateInsertRowCellProperties(string fill)
        => new(
            new TableCellWidth { Width = "2500", Type = TableWidthUnitValues.Dxa },
            new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });

    private static RunProperties CreateInsertRowRunStyle(string font, string fontSize, string color, UnderlineValues underline)
        => new(
            new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font },
            new FontSize { Val = fontSize },
            new FontSizeComplexScript { Val = fontSize },
            new Color { Val = color },
            new Underline { Val = underline });

    private static void CreateInsertRowStyleFixture(string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(
            new Table(
                new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                new TableGrid(new GridColumn { Width = "2500" }, new GridColumn { Width = "2500" }),
                new TableRow(
                    CreateInsertRowStyleCell("header-a", "D9EAD3", JustificationValues.Center, "Aptos", "20", "000000", UnderlineValues.None),
                    CreateInsertRowStyleCell("header-b", "D9EAD3", JustificationValues.Center, "Aptos", "20", "000000", UnderlineValues.None)),
                new TableRow(
                    CreateInsertRowStyleCell("template-a1", "DDEBF7", JustificationValues.Center, "Aptos", "22", "1F4E78", UnderlineValues.Single),
                    CreateInsertRowStyleCell("template-a2", "E2F0D9", JustificationValues.Right, "Aptos", "22", "1F4E78", UnderlineValues.Single)),
                new TableRow(
                    CreateInsertRowStyleCell("template-b1", "FCE4D6", JustificationValues.Left, "Courier New", "28", "C00000", UnderlineValues.Double),
                    CreateInsertRowStyleCell("template-b2", "FFF2CC", JustificationValues.Both, "Courier New", "28", "C00000", UnderlineValues.Double)),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("footer-a")))),
                    new TableCell(new Paragraph(new Run(new Text("footer-b"))))))));
        mainPart.Document.Save();
    }

    private static TableCell CreateInsertRowStyleCell(
        string text,
        string fill,
        JustificationValues justification,
        string font,
        string fontSize,
        string color,
        UnderlineValues underline)
        => new(
            new TableCellProperties(
                new TableCellWidth { Width = "2500", Type = TableWidthUnitValues.Dxa },
                new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
            new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = justification },
                    new SpacingBetweenLines { Before = "40", After = "80" },
                    new KeepNext()),
                new Run(new TabChar()),
                new Run(
                    new RunProperties(
                        new RunFonts { Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font },
                        new FontSize { Val = fontSize },
                        new FontSizeComplexScript { Val = fontSize },
                        new Color { Val = color },
                        new Underline { Val = underline }),
                    new Text(text))));

    private static void AssertInheritedCellStyle(
        TableCell cell,
        string fill,
        JustificationValues justification,
        string font,
        string fontSize,
        string color,
        UnderlineValues underline)
    {
        var cellProperties = cell.GetFirstChild<TableCellProperties>();
        Assert.Equal(fill, cellProperties!.GetFirstChild<Shading>()!.Fill!.Value);
        Assert.Equal(TableVerticalAlignmentValues.Center, cellProperties.GetFirstChild<TableCellVerticalAlignment>()!.Val!.Value);
        var paragraphProperties = cell.Elements<Paragraph>().Single().ParagraphProperties;
        Assert.Equal(justification, paragraphProperties!.Justification!.Val!.Value);
        Assert.Equal("40", paragraphProperties.SpacingBetweenLines!.Before!.Value);
        Assert.Equal("80", paragraphProperties.SpacingBetweenLines.After!.Value);
        Assert.NotNull(paragraphProperties.KeepNext);
        var runs = cell.Descendants<Run>().Where(run => run.Descendants<Text>().Any()).ToList();
        Assert.NotEmpty(runs);
        Assert.All(runs, run =>
        {
            var properties = run.RunProperties;
            Assert.NotNull(properties);
            Assert.NotNull(properties.RunFonts);
            Assert.NotNull(properties.FontSize);
            Assert.NotNull(properties.FontSizeComplexScript);
            Assert.NotNull(properties.Color);
            Assert.NotNull(properties.Underline);
            Assert.Equal(font, properties.RunFonts.Ascii!.Value);
            Assert.Equal(font, properties.RunFonts.HighAnsi!.Value);
            Assert.Equal(fontSize, properties.FontSize.Val!.Value);
            Assert.Equal(fontSize, properties.FontSizeComplexScript.Val!.Value);
            Assert.Equal(color, properties.Color.Val!.Value);
            Assert.Equal(underline, properties.Underline.Val!.Value);
        });
    }

    [Fact]
    public void Edit_replace_table_rows_matches_template_row_shape_and_preserves_widths()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-mixed-row-shapes-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                    new TableGrid(
                        new GridColumn { Width = "433" },
                        new GridColumn { Width = "255" },
                        new GridColumn { Width = "600" },
                        new GridColumn { Width = "650" },
                        new GridColumn { Width = "550" },
                        new GridColumn { Width = "2512" }),
                    new TableRow(
                        CreateSizedCenteredCell("试验类型", "433"),
                        CreateSizedCenteredCell("考察条件", "855", gridSpan: 2),
                        CreateSizedCenteredCell("拟考察批次", "650"),
                        CreateSizedCenteredCell("拟考察时间", "550"),
                        CreateSizedCenteredCell("试验结果与结论", "2512")),
                    new TableRow(
                        CreateSizedCenteredCell("", "433"),
                        CreateSizedCenteredCell("", "855", gridSpan: 2),
                        CreateSizedCenteredCell("", "650"),
                        CreateSizedCenteredCell("", "550"),
                        CreateSizedCenteredCell("", "2512")),
                    new TableRow(
                        CreateSizedCenteredCell("", "433"),
                        CreateSizedCenteredCell("", "255"),
                        CreateSizedCenteredCell("", "600"),
                        CreateSizedCenteredCell("", "650"),
                        CreateSizedCenteredCell("", "550"),
                        CreateSizedCenteredCell("", "2512")))));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"mixed-row-shapes-edited-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "replaceTableRows",
                TableIndex: 0,
                StartRowIndex: 1,
                EndRowIndex: 2,
                TemplateRowIndex: 1,
                Rows: [
                    [
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("长期试验", Color: "FF0000")]),
                        new DocxTableCellInput(GridSpan: 2, RichText: [new DocxRichTextSegment("≤-60℃，正置", Color: "FF0000")]),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("202401S\n202402S", Color: "FF0000")]),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("1、3、6、9、12、18、24、36 月\n1、3、6 月", Color: "FF0000")]),
                        new DocxTableCellInput(Text: "")
                    ],
                    [
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("影响因素", Color: "FF0000")]),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("高温试验", Color: "FF0000")]),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("25℃±2℃，60%RH±5%RH，避光，正置", Color: "FF0000")]),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("202401S", Color: "FF0000")]),
                        new DocxTableCellInput(RichText: [new DocxRichTextSegment("1、2、3 月", Color: "FF0000")]),
                        new DocxTableCellInput(Text: "")
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var rows = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();

        var mergedRowCells = rows[1].Elements<TableCell>().ToList();
        Assert.Equal(5, mergedRowCells.Count);
        Assert.Equal(2, mergedRowCells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
        Assert.Equal("855", mergedRowCells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
        Assert.Equal(JustificationValues.Center, mergedRowCells[3].Elements<Paragraph>().Single().ParagraphProperties!.Justification!.Val!.Value);
        Assert.Equal(1, mergedRowCells[3].Descendants<Break>().Count());

        var splitRowCells = rows[2].Elements<TableCell>().ToList();
        Assert.Equal(6, splitRowCells.Count);
        Assert.Equal("255", splitRowCells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
        Assert.Equal("600", splitRowCells[2].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
        Assert.Equal(JustificationValues.Center, splitRowCells[2].Elements<Paragraph>().Single().ParagraphProperties!.Justification!.Val!.Value);
        var validationErrors = new OpenXmlValidator().Validate(edited).Select(error => error.Description).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
    }

    [Fact]
    public void Edit_insert_table_rows_inherits_template_cell_merges_when_not_overridden()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-insert-row-merge-inheritance-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }),
                    new TableGrid(
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "1000" },
                        new GridColumn { Width = "1000" }),
                    new TableRow(
                        CreateSizedCenteredCell("检验项目", "2000", gridSpan: 2),
                        CreateSizedCenteredCell("可接受标准", "1000"),
                        CreateSizedCenteredCell("检验结果", "2000", gridSpan: 2)),
                    new TableRow(
                        CreateSizedCenteredCell("pH", "2000", gridSpan: 2),
                        CreateSizedCenteredCell("5.5±0.3", "1000"),
                        CreateSizedCenteredCell("待补充", "2000", gridSpan: 2)))));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"insert-row-merge-inheritance-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "insertTableRows",
                TableIndex: 0,
                RowIndex: 2,
                TemplateRowIndex: 1,
                Rows: [
                    [
                        new DocxTableCellInput("渗透压摩尔浓度"),
                        new DocxTableCellInput("240 - 360 mOsmol/kg"),
                        new DocxTableCellInput("待补充检测记录")
                    ]
                ])
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var insertedCells = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().ElementAt(2)
            .Elements<TableCell>().ToList();

        Assert.Equal(3, insertedCells.Count);
        Assert.Equal(2, insertedCells[0].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
        Assert.Equal("渗透压摩尔浓度", GetCellText(insertedCells[0]));
        Assert.Null(insertedCells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>());
        Assert.Equal(2, insertedCells[2].GetFirstChild<TableCellProperties>()!.GetFirstChild<GridSpan>()!.Val!.Value);
        var validationErrors = new OpenXmlValidator().Validate(edited).Select(error => error.Description).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
    }

    [Fact]
    public void Edit_set_table_width_preserves_existing_table_layout()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-table-layout-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(
                        new TableWidth { Width = "4200", Type = TableWidthUnitValues.Pct },
                        new TableLayout { Type = TableLayoutValues.Fixed }),
                    new TableRow(
                        new TableCell(new Paragraph(new Run(new Text("A"))))))));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"table-layout-edited-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(path, output, [
            new DocxEditOperation("setTableWidth", TableIndex: 0, Width: "5000", WidthType: "pct")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        var inspection = Inspector.InspectTables(output);
        Assert.Equal("5000", inspection.Tables.Single().Width);
        Assert.Equal("pct", inspection.Tables.Single().WidthType);
        using var edited = WordprocessingDocument.Open(output, false);
        var properties = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().GetFirstChild<TableProperties>()!;
        Assert.Equal("5000", properties.GetFirstChild<TableWidth>()!.Width!.Value);
        Assert.Equal(TableLayoutValues.Fixed, properties.GetFirstChild<TableLayout>()!.Type!.Value);
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
    public void Edit_rich_text_bold_false_overrides_inherited_paragraph_bold()
    {
        var path = Path.Combine(Path.GetTempPath(), $"rich-cell-inherited-bold-{Guid.NewGuid():N}.docx");
        using (var source = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = source.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn { Width = "2400" }),
                    new TableRow(
                        new TableCell(
                            new Paragraph(
                                new ParagraphProperties(
                                    new ParagraphMarkRunProperties(new Bold(), new BoldComplexScript())),
                                new Run(new Text("inherited bold"))))))));
            main.Document.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"rich-cell-inherited-bold-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: 0,
                CellIndex: 0,
                RichText: [new DocxRichTextSegment("normal text", Bold: false, FontName: "Times New Roman")])
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var paragraph = edited.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().Single()
            .Elements<Paragraph>().Single();
        Assert.NotNull(paragraph.ParagraphProperties!.ParagraphMarkRunProperties!.GetFirstChild<Bold>());
        var runProperties = paragraph.Elements<Run>().Single().RunProperties!;
        Assert.False(runProperties.GetFirstChild<Bold>()!.Val!.Value);
        Assert.False(runProperties.GetFirstChild<BoldComplexScript>()!.Val!.Value);
        var validationErrors = new OpenXmlValidator().Validate(edited).Select(error => error.Description).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
    }

    [Theory]
    [InlineData("replaceTableCellText")]
    [InlineData("replaceTableCellRichText")]
    public void Edit_blank_table_cell_text_overrides_inherited_paragraph_superscript(string operationType)
    {
        var path = Path.Combine(Path.GetTempPath(), $"blank-cell-inherited-superscript-{Guid.NewGuid():N}.docx");
        using (var source = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = source.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn { Width = "2400" }),
                    new TableRow(
                        new TableCell(
                            new Paragraph(
                                new ParagraphProperties(
                                    new ParagraphMarkRunProperties(
                                        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                        new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }))))))));
            main.Document.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"blank-cell-inherited-superscript-edited-{Guid.NewGuid():N}.docx");
        var operation = operationType == "replaceTableCellText"
            ? new DocxEditOperation(operationType, TableIndex: 0, RowIndex: 0, CellIndex: 0, Text: "ordinary text")
            : new DocxEditOperation(
                operationType,
                TableIndex: 0,
                RowIndex: 0,
                CellIndex: 0,
                RichText: [new DocxRichTextSegment("ordinary text")]);

        var result = Editor.Apply(path, output, [operation]);

        Assert.All(result.AppliedOperations, applied => Assert.True(applied.Applied, applied.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var paragraph = edited.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().Single()
            .Elements<Paragraph>().Single();
        Assert.Equal(
            VerticalPositionValues.Superscript,
            paragraph.ParagraphProperties!.ParagraphMarkRunProperties!
                .GetFirstChild<VerticalTextAlignment>()!.Val!.Value);
        var runProperties = paragraph.Elements<Run>().Single().RunProperties!;
        Assert.Equal(
            VerticalPositionValues.Baseline,
            runProperties.GetFirstChild<VerticalTextAlignment>()!.Val!.Value);
        var validationErrors = new OpenXmlValidator().Validate(edited).Select(error => error.Description).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
    }

    [Fact]
    public void Edit_nonblank_table_cell_preserves_existing_superscript()
    {
        var path = Path.Combine(Path.GetTempPath(), $"nonblank-cell-superscript-{Guid.NewGuid():N}.docx");
        using (var source = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var main = source.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn { Width = "2400" }),
                    new TableRow(
                        new TableCell(
                            new Paragraph(
                                new Run(
                                    new RunProperties(
                                        new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                                    new Text("existing marker"))))))));
            main.Document.Save();
        }
        var output = Path.Combine(Path.GetTempPath(), $"nonblank-cell-superscript-edited-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: 0,
                CellIndex: 0,
                RichText: [new DocxRichTextSegment("replacement marker")])
        ]);

        Assert.All(result.AppliedOperations, applied => Assert.True(applied.Applied, applied.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var runProperties = edited.MainDocumentPart!.Document!.Body!.Descendants<TableCell>().Single()
            .Descendants<Run>().Single().RunProperties!;
        Assert.Equal(
            VerticalPositionValues.Superscript,
            runProperties.GetFirstChild<VerticalTextAlignment>()!.Val!.Value);
    }

    [Fact]
    public void Edit_and_inspect_preserve_line_breaks_in_plain_and_rich_table_cell_text()
    {
        var source = CreateTwoCellTableFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-line-breaks-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(source, output, [
            new DocxEditOperation(
                "replaceTableCellText",
                TableIndex: 0,
                RowIndex: 0,
                CellIndex: 0,
                Text: "plain first\r\nplain second"),
            new DocxEditOperation(
                "replaceTableCellRichText",
                TableIndex: 0,
                RowIndex: 0,
                CellIndex: 1,
                RichText: [new DocxRichTextSegment("rich first\rrich second", Color: "FF0000")])
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        var cells = Assert.Single(Inspector.InspectTables(output).Tables).Rows[0].Cells;
        Assert.Equal("plain first\nplain second", cells[0].Text);
        Assert.Equal("plain first\nplain second", Assert.Single(cells[0].Paragraphs).Text);
        Assert.Equal("plain first\nplain second", Assert.Single(cells[0].Paragraphs[0].Runs).Text);
        Assert.Equal("rich first\nrich second", cells[1].Text);
        Assert.Equal("rich first\nrich second", Assert.Single(cells[1].Paragraphs).Text);
        Assert.Equal("rich first\nrich second", Assert.Single(cells[1].Paragraphs[0].Runs).Text);
    }

    [Fact]
    public void Edit_can_set_table_cell_font_size_and_row_height()
    {
        var docPath = CreateTwoCellTableFixture();
        var output = Path.Combine(Path.GetTempPath(), $"table-layout-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(docPath, output, [
            new DocxEditOperation("setTableCellFontSize", TableIndex: 0, RowIndex: 0, CellIndex: 1, FontSize: "9pt"),
            new DocxEditOperation("setTableCellNoWrap", TableIndex: 0, RowIndex: 0, CellIndex: 1),
            new DocxEditOperation("setTableRowHeight", TableIndex: 0, RowIndex: 0, Height: "240", HeightRule: "exact"),
            new DocxEditOperation("setTableRowCantSplit", TableIndex: 0, RowIndex: 0, CantSplit: true)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var doc = WordprocessingDocument.Open(output, false);
        var row = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().Single();
        var targetCell = row.Elements<TableCell>().ElementAt(1);
        Assert.NotNull(targetCell.GetFirstChild<TableCellProperties>()!.GetFirstChild<NoWrap>());
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
        Assert.NotNull(row.GetFirstChild<TableRowProperties>()!.GetFirstChild<CantSplit>());
        var validationErrors = new OpenXmlValidator().Validate(doc).Select(error => error.Description).ToList();
        Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
    }

    [Fact]
    public void Edit_can_remove_table_row_cant_split_and_inspector_reports_state()
    {
        var docPath = CreateTwoCellTableFixture();
        var withCantSplit = Path.Combine(Path.GetTempPath(), $"table-cant-split-{Guid.NewGuid():N}.docx");
        var withoutCantSplit = Path.Combine(Path.GetTempPath(), $"table-can-split-{Guid.NewGuid():N}.docx");

        var addResult = Editor.Apply(docPath, withCantSplit, [
            new DocxEditOperation("setTableRowCantSplit", TableIndex: 0, RowIndex: 0, CantSplit: true)
        ]);
        Assert.True(Assert.Single(addResult.AppliedOperations).Applied);
        Assert.True(Assert.Single(Inspector.InspectTables(withCantSplit).Tables).Rows[0].CantSplit);

        var removeResult = Editor.Apply(withCantSplit, withoutCantSplit, [
            new DocxEditOperation("setTableRowCantSplit", TableIndex: 0, RowIndex: 0, CantSplit: false)
        ]);
        Assert.True(Assert.Single(removeResult.AppliedOperations).Applied);
        Assert.False(Assert.Single(Inspector.InspectTables(withoutCantSplit).Tables).Rows[0].CantSplit);
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
    public void NormalizeOpenXml_repairs_wps_no_numbering_and_settings_order()
    {
        var output = Path.Combine(Path.GetTempPath(), $"normalized-wps-{Guid.NewGuid():N}.docx");
        File.Copy(CreateAnnotatedFixture(), output);
        ReplaceZipEntry(
            output,
            "word/document.xml",
            """
            <?xml version="1.0" encoding="utf-8"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:pPr>
                    <w:numPr>
                      <w:ilvl w:val="-1"/>
                      <w:numId w:val="0"/>
                    </w:numPr>
                  </w:pPr>
                  <w:r><w:t>Not numbered</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            """);
        ReplaceZipEntry(
            output,
            "word/settings.xml",
            """
            <?xml version="1.0" encoding="utf-8"?>
            <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:doNotAutoCompressPictures/>
              <w:themeFontLang/>
              <w:clrSchemeMapping/>
              <w:doNotIncludeSubdocsInStats/>
            </w:settings>
            """);

        DocxPackageNormalizer.Normalize(output, output);

        var documentXml = ReadZipEntry(output, "word/document.xml");
        var settingsXml = ReadZipEntry(output, "word/settings.xml");
        Assert.DoesNotContain("<w:numPr", documentXml);
        Assert.True(settingsXml.IndexOf("<w:themeFontLang", StringComparison.Ordinal)
            < settingsXml.IndexOf("<w:clrSchemeMapping", StringComparison.Ordinal));
        Assert.True(settingsXml.IndexOf("<w:clrSchemeMapping", StringComparison.Ordinal)
            < settingsXml.IndexOf("<w:doNotIncludeSubdocsInStats", StringComparison.Ordinal));
        Assert.True(settingsXml.IndexOf("<w:doNotIncludeSubdocsInStats", StringComparison.Ordinal)
            < settingsXml.IndexOf("<w:doNotAutoCompressPictures", StringComparison.Ordinal));
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
    public void Edit_can_start_landscape_section_before_anchored_paragraph()
    {
        var path = Path.Combine(Path.GetTempPath(), $"section-anchor-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("缩略词表"))),
                new Paragraph(new Run(new Text("3.2.S.7.1.1 试验样品"))),
                new Paragraph(
                    new ParagraphProperties(
                        new SectionProperties(
                            new PageSize { Width = 11906, Height = 16838 },
                            new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })),
                    new Run(new Text("end portrait marker"))),
                new Paragraph(new Run(new Text("after"))),
                new SectionProperties(
                    new PageSize { Width = 16838, Height = 11906, Orient = PageOrientationValues.Landscape },
                    new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"section-anchor-edited-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(path, output, [
            new DocxEditOperation(
                "startSectionBeforeParagraph",
                FindText: "3.2.S.7.1.1 试验样品",
                Orientation: "landscape")
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var bodyChildren = edited.MainDocumentPart!.Document!.Body!.ChildElements.ToList();
        var insertedBreak = (Paragraph)bodyChildren[1];
        var insertedPageSize = insertedBreak.ParagraphProperties!.GetFirstChild<SectionProperties>()!.GetFirstChild<PageSize>()!;
        Assert.Equal(11906U, insertedPageSize.Width!.Value);
        Assert.Equal(16838U, insertedPageSize.Height!.Value);

        var originalSectionParagraph = bodyChildren.OfType<Paragraph>().Single(paragraph => GetParagraphText(paragraph).Contains("end portrait marker", StringComparison.Ordinal));
        var updatedPageSize = originalSectionParagraph.ParagraphProperties!.GetFirstChild<SectionProperties>()!.GetFirstChild<PageSize>()!;
        Assert.Equal(PageOrientationValues.Landscape, updatedPageSize.Orient!.Value);
        Assert.Equal(16838U, updatedPageSize.Width!.Value);
        Assert.Equal(11906U, updatedPageSize.Height!.Value);
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

    private static string CreateHeaderFooterTableFixture(string headerText, string footerText)
    {
        var path = Path.Combine(Path.GetTempPath(), $"header-footer-table-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var header = main.AddNewPart<HeaderPart>();
        var headerCell = new TableCell();
        headerCell.Append(new Paragraph(new Run(new Text(headerText))));
        var headerTable = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "2400" }),
            new TableRow(headerCell));
        header.Header = new Header(new Paragraph(new Run(new Text($"header-paragraph-{headerText}"))), headerTable);
        var footer = main.AddNewPart<FooterPart>();
        var footerCell = new TableCell();
        footerCell.Append(new Paragraph(new Run(new Text(footerText))));
        var footerTable = new Table(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "2400" }),
            new TableRow(footerCell));
        footer.Footer = new Footer(new Paragraph(new Run(new Text($"footer-paragraph-{footerText}"))), footerTable);
        var section = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) },
            new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
        main.Document = new Document(new Body(new Paragraph(new Run(new Text("body"))), section));
        main.Document.Save();
        header.Header.Save();
        footer.Footer.Save();
        return path;
    }

    private static string CreateExactTextMappingFixture(bool includeDuplicateBaselineText, bool baseline)
    {
        var path = Path.Combine(Path.GetTempPath(), $"exact-text-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body();
        if (baseline) body.Append(new Paragraph(new Run(new Text("baseline heading"))));
        body.Append(new Paragraph(new Run(new Text("unique paragraph"))));
        if (baseline && includeDuplicateBaselineText) body.Append(new Paragraph(new Run(new Text("unique paragraph"))));
        var row = baseline
            ? new TableRow(new TableCell(new Paragraph(new Run(new Text("baseline cell")))), new TableCell(new Paragraph(new Run(new Text("unique cell")))))
            : new TableRow(new TableCell(new Paragraph(new Run(new Text("unique cell")))));
        body.Append(new Table(new TableProperties(), new TableGrid(new GridColumn { Width = "2400" }, new GridColumn { Width = "2400" }), row));
        main.Document = new Document(body);
        main.Document.Save();
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

    private static TableCell CreateSizedCenteredCell(string text, string width, int? gridSpan = null)
    {
        var properties = new TableCellProperties(
            new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa },
            new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center });
        if (gridSpan is > 1)
        {
            properties.AppendChild(new GridSpan { Val = gridSpan.Value });
        }

        var paragraph = new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }));
        if (!string.IsNullOrEmpty(text))
        {
            paragraph.AppendChild(new Run(new Text(text)));
        }

        return new TableCell(properties, paragraph);
    }

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
    public void Edit_vertical_merge_resolves_logical_grid_column_across_different_row_spans()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-grid-merge-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Table(
                new TableRow(
                    new TableCell(new TableCellProperties(new GridSpan { Val = 2 }), new Paragraph(new Run(new Text("wide")))),
                    new TableCell(new Paragraph(new Run(new Text("top"))))),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("a")))),
                    new TableCell(new Paragraph(new Run(new Text("b")))),
                    new TableCell(new Paragraph(new Run(new Text("bottom"))))))));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"grid-merged-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(path, output, [
            new DocxEditOperation("mergeTableCells", TableIndex: 0, GridColumn: 2, StartRowIndex: 0, EndRowIndex: 1)
        ]);

        Assert.All(result.AppliedOperations, operation => Assert.True(operation.Applied, operation.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var rows = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
        var top = rows[0].Elements<TableCell>().ElementAt(1).GetFirstChild<TableCellProperties>()!.GetFirstChild<VerticalMerge>();
        var bottom = rows[1].Elements<TableCell>().ElementAt(2).GetFirstChild<TableCellProperties>()!.GetFirstChild<VerticalMerge>();
        Assert.Equal(MergedCellValues.Restart, top!.Val!.Value);
        Assert.Equal(MergedCellValues.Continue, bottom!.Val!.Value);
    }

    [Fact]
    public void Edit_vertical_merge_promotes_continuation_with_owner_paragraph_properties()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-merge-properties-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Restart }),
                            new Paragraph(
                                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                                new Run(new Text("owner")))),
                        new TableCell(new Paragraph(new Run(new Text("A"))))),
                    new TableRow(
                        new TableCell(
                            new TableCellProperties(new VerticalMerge { Val = MergedCellValues.Continue }),
                            new Paragraph()),
                        new TableCell(new Paragraph(new Run(new Text("B"))))),
                    new TableRow(
                        new TableCell(new Paragraph()),
                        new TableCell(new Paragraph(new Run(new Text("C")))))
                )));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"merged-cell-properties-{Guid.NewGuid():N}.docx");
        var result = Editor.Apply(path, output, [
            new DocxEditOperation("mergeTableCells", TableIndex: 0, CellIndex: 0, StartRowIndex: 1, EndRowIndex: 2)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var cell = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().ElementAt(1)
            .Elements<TableCell>().First();
        var merge = cell.GetFirstChild<TableCellProperties>()!.GetFirstChild<VerticalMerge>();
        Assert.Equal(MergedCellValues.Restart, merge!.Val!.Value);
        var justification = cell.Elements<Paragraph>().Single()
            .GetFirstChild<ParagraphProperties>()!.GetFirstChild<Justification>();
        Assert.Equal(JustificationValues.Center, justification!.Val!.Value);
        Assert.Empty(cell.Descendants<Text>());
    }

    [Fact]
    public void Edit_can_delete_table_rows()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-delete-rows-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(),
                    new TableGrid(new GridColumn()),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("header"))))),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("keep-1"))))),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("delete-1"))))),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("delete-2"))))),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("keep-2")))))
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"deleted-rows-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("deleteTableRows", TableIndex: 0, StartRowIndex: 2, EndRowIndex: 3)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using (var doc = WordprocessingDocument.Open(output, false))
        {
            var rows = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single().Elements<TableRow>().ToList();
            Assert.Equal(["header", "keep-1", "keep-2"], rows.Select(row => GetCellText(row.Elements<TableCell>().Single())).ToArray());
            Assert.Empty(new OpenXmlValidator().Validate(doc));
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
            Assert.Equal("1000", cells[0].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
            Assert.Equal("1000", cells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
            Assert.Equal("1000", cells[2].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
            Assert.Equal("高温试验", GetCellText(cells[0]));
            Assert.Equal("", GetCellText(cells[1]));
            Assert.Equal("", GetCellText(cells[2]));
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
    }

    [Fact]
    public void Edit_unmerge_table_row_horizontal_cells_uses_visible_reference_widths_before_grid_widths()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fixture-unmerge-reference-widths-{Guid.NewGuid():N}.docx");
        using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Table(
                    new TableProperties(),
                    new TableGrid(
                        new GridColumn { Width = "265" },
                        new GridColumn { Width = "1841" },
                        new GridColumn { Width = "1560" },
                        new GridColumn { Width = "311" }
                    ),
                    new TableRow(
                        CreateSizedCenteredCell("影响因素", "265"),
                        CreateSizedCenteredCell("反复冻融试验", "1007"),
                        CreateSizedCenteredCell("≤-60℃-室温", "853"),
                        CreateSizedCenteredCell("放行数据", "311")
                    ),
                    new TableRow(
                        CreateSizedCenteredCell("", "265"),
                        CreateSizedCenteredCell("高温试验", "1860", gridSpan: 2),
                        CreateSizedCenteredCell("", "311")
                    )
                )
            ));
            mainPart.Document.Save();
        }

        var output = Path.Combine(Path.GetTempPath(), $"unmerged-reference-widths-{Guid.NewGuid():N}.docx");

        var result = Editor.Apply(path, output, [
            new DocxEditOperation("unmergeTableRowHorizontalCells", TableIndex: 0, RowIndex: 1, CellIndex: 1)
        ]);

        Assert.All(result.AppliedOperations, op => Assert.True(op.Applied, op.Detail));
        using var edited = WordprocessingDocument.Open(output, false);
        var cells = edited.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().ElementAt(1)
            .Elements<TableCell>().ToList();
        Assert.Equal("1007", cells[1].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
        Assert.Equal("853", cells[2].GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellWidth>()!.Width!.Value);
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

    private static string CreateMediaMigrationFixture(string text, byte[] mediaBytes)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-media-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var image = main.AddImagePart(ImagePartType.Png);
        using var stream = new MemoryStream(mediaBytes);
        image.FeedData(stream);
        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = 990000L, Cy = 990000L },
                new DW.DocProperties { Id = 1U, Name = "migration-image" },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "migration-image" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = main.GetIdOfPart(image) },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 990000L, Cy = 990000L }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })));
        main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)), new Run(drawing))));
        main.Document.Save();
        return path;
    }

    private static string CreateTextMigrationFixture(params string[] text)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-text-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(text.Select(value => new Paragraph(new Run(new Text(value))))));
        main.Document.Save();
        return path;
    }

    private static string CreateStyledTextMigrationFixture(params (string Text, string StyleId)[] paragraphs)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-styled-text-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(paragraphs.Select(item => item.StyleId).Distinct(StringComparer.Ordinal).Select(styleId =>
            new Style(new StyleName { Val = styleId }) { Type = StyleValues.Paragraph, StyleId = styleId }));
        main.Document = new Document(new Body(paragraphs.Select(item =>
            new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = item.StyleId }), new Run(new Text(item.Text))))));
        main.Document.Save();
        stylesPart.Styles.Save();
        return path;
    }

    private static TemplateMigrationSemanticCandidate SemanticValueCandidate(string sourceText, string baselineText, string valueKind)
        => new(
            "tiwater.docx.template-migration-semantic-candidate/v2",
            [],
            ValueProjections:
            [
                new TemplateMigrationSemanticCandidateValueProjection(
                    new TemplateMigrationSemanticSelector("paragraph", "body", sourceText),
                    new TemplateMigrationSemanticSelector("paragraph", "body", baselineText),
                    "declared-fact",
                    valueKind,
                    "after-first-delimiter")
            ]);

    private static string CreateSemanticValueProjectionFixture(IReadOnlyList<string> runs, bool duplicateParagraph = false, bool useTableCell = false)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-semantic-value-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        Paragraph ParagraphFromRuns() => new(runs.Select(text => new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));
        OpenXmlElement root = useTableCell
            ? new Table(new TableRow(new TableCell(ParagraphFromRuns())))
            : ParagraphFromRuns();
        var body = new Body(root);
        if (duplicateParagraph) body.AppendChild(ParagraphFromRuns());
        main.Document = new Document(body);
        main.Document.Save();
        return path;
    }

    private static string CreateMultiFieldProjectionFixture(params IReadOnlyList<string>[] fields)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-semantic-multi-field-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var cell = new TableCell(fields.Select(field => new Paragraph(field.Select(text => new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })))));
        main.Document = new Document(new Body(new Table(new TableRow(cell))));
        main.Document.Save();
        return path;
    }

    private static string ReadOnlyParagraphText(string path)
    {
        using var document = WordprocessingDocument.Open(path, false);
        return string.Concat(document.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>().First().Descendants<Text>().Select(item => item.Text));
    }

    private static string CreateBaselineClearFixture(string first, string second)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-baseline-clear-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Table(
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text(first)))),
                new TableCell(new Paragraph(new Run(new Text("header"))))),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text(second)))),
                new TableCell(new Paragraph(new Run(new Text("baseline default"))))))));
        main.Document.Save();
        return path;
    }

    private static string CreateBodyAppendFixture(bool includeDuplicateRevisionTable, bool baseline)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-body-append-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var body = new Body(new Paragraph(new Run(new Text("before"))), new Paragraph(new Run(new Text("after"))));
        if (!baseline)
        {
            var after = body.Elements<Paragraph>().Last();
            body.InsertBefore(new Paragraph(new Run(new Text("Revision history"))), after);
            body.InsertBefore(CreateRevisionTable(), after);
            if (includeDuplicateRevisionTable) body.InsertBefore(CreateRevisionTable(), after);
        }
        main.Document = new Document(body);
        main.Document.Save();
        return path;
    }

    private static Table CreateRevisionTable()
        => new(
            new TableProperties(),
            new TableGrid(new GridColumn { Width = "2400" }, new GridColumn { Width = "2400" }),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("Revision No.")))), new TableCell(new Paragraph(new Run(new Text("Description"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("R1")))), new TableCell(new Paragraph(new Run(new Text("Current source fact"))))));

    private static string CreateLabeledRunMigrationFixture(string label, string value)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-labeled-run-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(
            new Run(new RunProperties(new Bold()), new Text(label)),
            new Run(new RunProperties(new Italic()), new Text(value)))));
        main.Document.Save();
        return path;
    }

    private static string CreateCrossTemplateMigrationFixture(string headerText, string openingText, string factText, string closingText, string footerText, bool shifted)
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-cross-template-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var header = main.AddNewPart<HeaderPart>();
        header.Header = new Header(new Paragraph(new Run(new Text(headerText))));
        var footer = main.AddNewPart<FooterPart>();
        footer.Footer = new Footer(new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text(footerText)))))));
        var section = new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) },
            new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
        var factTable = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text(factText))))));
        main.Document = shifted
            ? new Document(new Body(
                new Paragraph(new Run(new Text("template-owned banner"))),
                factTable,
                new Paragraph(new Run(new Text(closingText))),
                new Paragraph(new Run(new Text(openingText))),
                section))
            : new Document(new Body(
                new Paragraph(new Run(new Text(openingText))),
                factTable,
                new Paragraph(new Run(new Text(closingText))),
                section));
        main.Document.Save();
        header.Header.Save();
        footer.Footer.Save();
        return path;
    }

    private static string CreatePlainMigrationFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"migration-plain-{Guid.NewGuid():N}.docx");
        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text("Project code XXXX 峰面积"))),
            new Table(new TableRow(
                new TableCell(new Paragraph(new Run(new Text("Batch")))),
                new TableCell(new Paragraph(new Run(new Text("Batch YYYY")))))),
            new Paragraph(new Run(new Text("Top-level paragraph XXXX")))));
        main.Document.Save();
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
