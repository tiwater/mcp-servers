using Dockit.Pptx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System.Security.Cryptography;
using System.Reflection;
using System.Text.Json;

namespace Dockit.Pptx.Tests;

public class PptxCliTests
{
    [Fact]
    public void Inspect_reports_slide_metrics_and_placeholders()
    {
        var path = CreateFixture();

        var report = Inspector.Inspect(path);

        Assert.Equal(path, report.File);
        Assert.Equal(2, report.SlideCount);
        Assert.Equal(["batch", "title"], report.Placeholders);
        Assert.Equal(2, report.Slides.Count);
        Assert.Equal(1, report.Slides[0].SlideNumber);
        Assert.Equal("ppt/slides/slide1.xml", report.Slides[0].Path);
        Assert.Equal(["title"], report.Slides[0].Placeholders);
        Assert.Equal(2, report.Slides[1].SlideNumber);
        Assert.Equal(["batch"], report.Slides[1].Placeholders);
    }

    [Fact]
    public void ExportJson_includes_slide_text_and_notes()
    {
        var path = CreateFixture();

        var report = Extractor.Export(path);

        Assert.Equal(path, report.File);
        Assert.Equal(2, report.Slides.Count);
        Assert.Contains("Project {{title}} 峰面积", report.Slides[0].Texts);
        Assert.Equal(["title"], report.Slides[0].Placeholders);
        Assert.Contains("Batch {{batch}}", report.Slides[1].Texts);
        Assert.Equal(["batch"], report.Slides[1].Placeholders);
        Assert.Single(report.Notes);
        Assert.Equal(1, report.Notes[0].NotesNumber);
        Assert.Equal("ppt/notesSlides/notesSlide1.xml", report.Notes[0].Path);
        Assert.Contains("Notes {{title}}", report.Notes[0].Texts);

        var output = Path.Combine(Path.GetTempPath(), $"pptx-export-{Guid.NewGuid():N}.json");
        Extractor.RunExportJson([path, output]);

        var json = File.ReadAllText(output);
        Assert.Contains("Project {{title}} 峰面积", json, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\u5CF0", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u9762", json, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(@"\u79EF", json, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Fill_replaces_tokens_in_slides_and_notes()
    {
        var template = CreateFixture();
        var output = Path.Combine(Path.GetTempPath(), $"pptx-filled-{Guid.NewGuid():N}.pptx");

        var result = TemplateFiller.Fill(template, new Dictionary<string, string>
        {
            ["title"] = "Q2 Review",
            ["batch"] = "B-001"
        }, output);

        Assert.Equal(template, result.Template);
        Assert.Equal(output, result.Output);
        Assert.Equal(2, result.ChangedSlides);
        Assert.Equal(1, result.ChangedNotes);
        Assert.Equal(2, result.PlaceholderCount);

        using var presentation = PresentationDocument.Open(output, false);
        var slideTexts = presentation.PresentationPart!
            .SlideParts
            .SelectMany(part => part.Slide.Descendants<A.Text>().Select(text => text.Text))
            .ToList();
        Assert.Contains("Project Q2 Review 峰面积", slideTexts);
        Assert.Contains("Batch B-001", slideTexts);

        var notesTexts = presentation.PresentationPart!
            .SlideParts
            .Where(part => part.NotesSlidePart is not null)
            .SelectMany(part => part.NotesSlidePart!.NotesSlide.Descendants<A.Text>().Select(text => text.Text))
            .ToList();
        Assert.Contains("Notes Q2 Review", notesTexts);
    }

    [Fact]
    public void InspectDetail_reports_shape_positions_and_run_formatting()
    {
        var path = CreateFixture();

        var report = Inspector.InspectDetail(path);

        Assert.Equal(path, report.File);
        Assert.Equal(2, report.SlideCount);
        Assert.Equal(9144000L, report.SlideSize.Cx);
        Assert.Equal(6858000L, report.SlideSize.Cy);

        var firstShape = Assert.Single(report.Slides[0].Shapes);
        Assert.Equal(2U, firstShape.ShapeId);
        Assert.Equal("TextBox 1", firstShape.Name);
        Assert.Equal("shape", firstShape.Kind);
        Assert.Equal("Project {{title}} 峰面积", firstShape.Text);
        Assert.Equal(914400L, firstShape.Transform?.X);
        Assert.Equal(457200L, firstShape.Transform?.Y);
        Assert.Equal(3657600L, firstShape.Transform?.Cx);
        Assert.Equal(457200L, firstShape.Transform?.Cy);

        var run = Assert.Single(firstShape.Runs);
        Assert.Equal(0, run.RunIndex);
        Assert.Equal("Project {{title}} 峰面积", run.Text);
        Assert.Equal("微软雅黑", run.FontFamily);
        Assert.Equal(16d, run.FontSize);
        Assert.Equal("287341", run.Color);
        Assert.True(run.Bold);
        Assert.Single(report.Masters);
        Assert.Equal("ppt/slideMasters/slideMaster1.xml", report.Masters[0].Path);
        Assert.Equal("ppt/slideLayouts/slideLayout1.xml", report.Masters[0].Layouts[0].Path);
        Assert.Equal(report.Masters[0].Path, report.Slides[0].MasterPath);
        Assert.Equal(report.Masters[0].Layouts[0].Path, report.Slides[0].LayoutPath);
        Assert.Equal(0, firstShape.ZOrder);
    }

    [Fact]
    public void InspectDetail_reports_placeholder_presence_and_index_without_inferring_from_type_or_text()
    {
        var path = CreatePlaceholderEvidenceFixture();

        var report = Inspector.InspectDetail(path);
        var shapes = report.Slides[0].Shapes;

        var indexed = shapes.Single(shape => shape.ShapeId == 3U);
        Assert.True(indexed.PlaceholderPresent);
        Assert.Null(indexed.PlaceholderType);
        Assert.Equal(12U, indexed.PlaceholderIndex);

        var typed = shapes.Single(shape => shape.ShapeId == 4U);
        Assert.True(typed.PlaceholderPresent);
        Assert.Equal("title", typed.PlaceholderType);
        Assert.Null(typed.PlaceholderIndex);

        var textOnly = shapes.Single(shape => shape.ShapeId == 5U);
        Assert.False(textOnly.PlaceholderPresent);
        Assert.Null(textOnly.PlaceholderIndex);
    }

    [Fact]
    public void InspectDetail_placeholder_evidence_is_stable_and_tracks_openxml_mutation()
    {
        var path = CreatePlaceholderEvidenceFixture();

        var before = Inspector.InspectDetail(path).Slides[0].Shapes.Single(shape => shape.ShapeId == 3U);
        var repeat = Inspector.InspectDetail(path).Slides[0].Shapes.Single(shape => shape.ShapeId == 3U);
        Assert.Equal((before.PlaceholderPresent, before.PlaceholderIndex), (repeat.PlaceholderPresent, repeat.PlaceholderIndex));

        using (var presentation = PresentationDocument.Open(path, true))
        {
            var shape = presentation.PresentationPart!.SlideParts.First().Slide.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().Single(shape => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value == 3U);
            shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Type = P.PlaceholderValues.Body;
            shape.TextBody!.Descendants<A.Text>().Single().Text = "mutated visible text";
            presentation.PresentationPart.SlideParts.First().Slide.Save();
        }

        var after = Inspector.InspectDetail(path).Slides[0].Shapes.Single(shape => shape.ShapeId == 3U);
        Assert.True(after.PlaceholderPresent);
        Assert.Equal(12U, after.PlaceholderIndex);
        Assert.Equal("body", after.PlaceholderType);
    }

    [Fact]
    public void InspectDetail_serializes_all_shape_kinds_against_the_published_shape_schema()
    {
        var path = CreateAllShapeKindsFixture();
        var report = Inspector.InspectDetail(path);
        var json = JsonDocument.Parse(JsonSerializer.Serialize(report, Json.Options));
        var schema = JsonDocument.Parse(File.ReadAllText(RepositoryPath("packages/pptx-cli/contracts/tiwater.pptx-inspect-shape-v1.schema.json")));
        var slideShapes = json.RootElement.GetProperty("slides")[0].GetProperty("shapes");
        var expected = new[]
        {
            (Kind: "shape", Index: (uint?)31U),
            (Kind: "picture", Index: (uint?)32U),
            (Kind: "graphicFrame", Index: (uint?)33U),
            (Kind: "groupShape", Index: (uint?)34U),
        };

        foreach (var item in expected)
        {
            var shape = Assert.Single(report.Slides[0].Shapes, value => value.Kind == item.Kind && value.PlaceholderIndex == item.Index);
            Assert.True(shape.PlaceholderPresent);
            Assert.Equal(item.Index, shape.PlaceholderIndex);
            var serialized = Assert.Single(slideShapes.EnumerateArray(), value => value.GetProperty("kind").GetString() == item.Kind
                && value.GetProperty("placeholderIndex").ValueKind == JsonValueKind.Number
                && value.GetProperty("placeholderIndex").GetUInt32() == item.Index);
            AssertShapeEvidenceMatchesSchema(serialized, schema.RootElement);
            Assert.Contains(serialized.GetProperty("placeholderPresent").ValueKind, new[] { JsonValueKind.True, JsonValueKind.False });
            Assert.Equal(JsonValueKind.Number, serialized.GetProperty("placeholderIndex").ValueKind);
        }
    }

    [Fact]
    public void InspectDetail_placeholder_mutations_remove_and_change_index_evidence()
    {
        var path = CreatePlaceholderEvidenceFixture();
        var shapeId = 3U;

        using (var presentation = PresentationDocument.Open(path, true))
        {
            var shape = presentation.PresentationPart!.SlideParts.First().Slide.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().Single(value => value.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value == shapeId);
            shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Index = 0U;
            presentation.PresentationPart.SlideParts.First().Slide.Save();
        }

        var zeroIndex = Inspector.InspectDetail(path).Slides[0].Shapes.Single(value => value.ShapeId == shapeId);
        Assert.True(zeroIndex.PlaceholderPresent);
        Assert.Equal(0U, zeroIndex.PlaceholderIndex);

        using (var presentation = PresentationDocument.Open(path, true))
        {
            var shape = presentation.PresentationPart!.SlideParts.First().Slide.CommonSlideData!.ShapeTree!
                .Elements<P.Shape>().Single(value => value.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value == shapeId);
            shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Remove();
            presentation.PresentationPart.SlideParts.First().Slide.Save();
        }

        var removed = Inspector.InspectDetail(path).Slides[0].Shapes.Single(value => value.ShapeId == shapeId);
        Assert.False(removed.PlaceholderPresent);
        Assert.Null(removed.PlaceholderIndex);
        Assert.Null(removed.PlaceholderType);
    }

    [Fact]
    public void InspectDetail_deserializes_legacy_shape_json_with_additive_placeholder_fields_absent()
    {
        const string legacy = """
            {
              "shapeId": 2,
              "name": "legacy",
              "kind": "shape",
              "zOrder": 0,
              "placeholderType": null,
              "mediaPartPath": null,
              "mediaSha256": null,
              "text": "legacy text",
              "transform": null,
              "paragraphs": [],
              "runs": [],
              "table": null
            }
            """;

        var shape = JsonSerializer.Deserialize<ShapeDetail>(legacy, Json.Options);

        Assert.NotNull(shape);
        Assert.Equal(2U, shape!.ShapeId);
        Assert.False(shape.PlaceholderPresent);
        Assert.Null(shape.PlaceholderIndex);
        Assert.Null(shape.PlaceholderType);
    }

    [Fact]
    public void ShapeDetail_preserves_the_legacy_public_constructor_and_deconstruct_arity()
    {
        var legacyParameters = new[]
        {
            typeof(uint), typeof(string), typeof(string), typeof(int), typeof(string), typeof(string),
            typeof(string), typeof(string), typeof(TransformInfo), typeof(IReadOnlyList<ParagraphDetail>),
            typeof(IReadOnlyList<TextRunDetail>), typeof(TableDetail),
        };

        Assert.NotNull(typeof(ShapeDetail).GetConstructor(legacyParameters));
        Assert.Contains(typeof(ShapeDetail).GetMethods(BindingFlags.Public | BindingFlags.Instance), method =>
            method.Name == "Deconstruct" && method.GetParameters().Length == legacyParameters.Length);

        var shape = new ShapeDetail(2U, "legacy", "shape", 0, null, null, null, "text", null, [], [], null);
        shape.Deconstruct(out var id, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _, out _);
        Assert.Equal(2U, id);
    }

    [Fact]
    public void InspectDetail_resolves_the_current_paragraph_level_without_fabricating_direct_formatting()
    {
        var path = CreateInheritedFormattingFixture();

        var run = Assert.Single(Inspector.InspectDetail(path).Slides[0].Shapes[0].Runs);

        Assert.Equal("Inherited Sans", run.FontFamily);
        Assert.Equal(24d, run.FontSize);
        Assert.Equal("FFFFFF", run.Color);
        Assert.True(run.Bold);
        Assert.Null(run.DirectFontFamily);
        Assert.Null(run.DirectFontSize);
        Assert.Null(run.DirectColor);
        Assert.Null(run.DirectBold);
        Assert.Equal("shape-list-level-1", run.FontFamilySource);
        Assert.Equal("shape-list-level-1", run.FontSizeSource);
        Assert.Equal("shape-list-default", run.ColorSource);
        Assert.Equal("shape-list-level-1", run.BoldSource);
    }

    [Fact]
    public void InspectDetail_keeps_an_explicit_wrong_color_distinct_from_inherited_color()
    {
        var path = CreateInheritedFormattingFixture("287341");

        var run = Assert.Single(Inspector.InspectDetail(path).Slides[0].Shapes[0].Runs);

        Assert.Equal("287341", run.Color);
        Assert.Equal("287341", run.DirectColor);
        Assert.Equal("direct-run", run.ColorSource);
        Assert.Equal(24d, run.FontSize);
    }

    [Fact]
    public void InspectDetail_uses_master_other_text_style_when_the_shape_has_no_local_default()
    {
        var path = CreateInheritedFormattingFixture(useLocalStyle: false);

        var run = Assert.Single(Inspector.InspectDetail(path).Slides[0].Shapes[0].Runs);

        Assert.Equal("Master Sans", run.FontFamily);
        Assert.Equal(18d, run.FontSize);
        Assert.Equal("000000", run.Color);
        Assert.Equal("master-text-style-level-1", run.FontFamilySource);
        Assert.Equal("master-text-style-level-1", run.FontSizeSource);
        Assert.Equal("master-text-style-level-1", run.ColorSource);
    }

    [Fact]
    public void ApplyTemplate_preserves_slide_content_and_switches_every_slide_to_target_master()
    {
        var source = CreateFixture();
        var template = CreateFixture();
        using (var target = PresentationDocument.Open(template, true))
        {
            target.PresentationPart!.SlideMasterParts.Single().SlideMaster.CommonSlideData!.Name = "Approved Master";
            target.PresentationPart.SlideMasterParts.Single().SlideMaster.Save();
        }
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var before = Inspector.InspectDetail(source).Slides.SelectMany(slide => slide.Shapes).Select(shape => shape.Text).ToList();

        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout), new SlideLayoutAssignment(2, targetLayout)]), output);

        Assert.Empty(result.Issues);
        Assert.Equal(2, result.ChangedSlideCount);
        var after = Inspector.InspectDetail(output);
        Assert.All(after.Slides, slide => Assert.Equal("Approved Master", after.Masters.Single(master => master.Path == slide.MasterPath).Name));
        Assert.Equal(before, after.Slides.SelectMany(slide => slide.Shapes).Select(shape => shape.Text).ToList());
    }

    [Fact]
    public void ApplyTemplate_rejects_unscoped_content_fitting_before_mutating_slide()
    {
        var source = CreateFixture();
        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var before = Inspector.InspectDetail(source).Slides[0];

        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout, new TransformInfo(0, 0, 1000000, 1000000))]), output);

        Assert.Equal(0, result.ChangedSlideCount);
        Assert.Equal("content shape ids are required when content bounds are specified", Assert.Single(result.Issues).Message);
        var after = Inspector.InspectDetail(output).Slides[0];
        Assert.Equal(before.LayoutPath, after.LayoutPath);
        Assert.Equal(before.Shapes.Select(shape => (shape.ShapeId, shape.Text, shape.Transform)), after.Shapes.Select(shape => (shape.ShapeId, shape.Text, shape.Transform)));
    }

    [Fact]
    public void ApplyTemplate_fits_only_explicitly_selected_shapes()
    {
        var source = CreateFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var slide = presentation.PresentationPart!.SlideParts.First().Slide;
            slide.CommonSlideData!.ShapeTree!.Append(CreateTextShape(3U, "Fixed logo", 7000000L, 100000L, 1000000L, 300000L));
            slide.Save();
        }
        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var fixedBefore = Inspector.InspectDetail(source).Slides[0].Shapes.Single(shape => shape.ShapeId == 3U);

        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout, new TransformInfo(1000000, 1000000, 2000000, 1000000), [2U])]), output);

        Assert.Empty(result.Issues);
        var after = Inspector.InspectDetail(output).Slides[0];
        Assert.Equal(new TransformInfo(1000000, 1375000, 2000000, 250000), after.Shapes.Single(shape => shape.ShapeId == 2U).Transform);
        var fixedAfter = after.Shapes.Single(shape => shape.ShapeId == 3U);
        Assert.Equal((fixedBefore.ShapeId, fixedBefore.Text, fixedBefore.Transform), (fixedAfter.ShapeId, fixedAfter.Text, fixedAfter.Transform));
    }

    [Fact]
    public void ApplyTemplate_freezes_placeholder_geometry_and_materializes_only_declared_layout_content()
    {
        var source = CreateFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var master = presentation.PresentationPart!.SlideMasterParts.Single();
            var layout = master.SlideLayoutParts.Single();
            layout.SlideLayout.CommonSlideData!.ShapeTree!.Append(
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 8U, Name = "Inherited title" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                    new P.ShapeProperties(new A.Transform2D(
                        new A.Offset { X = 123400L, Y = 234500L },
                        new A.Extents { Cx = 3456000L, Cy = 456700L })),
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph())));
            layout.SlideLayout.CommonSlideData.ShapeTree.Append(CreateTextShape(9U, "Current visible layout content", 900000L, 1000000L, 3000000L, 400000L));
            layout.SlideLayout.Save();
            master.SlideMaster.TextStyles = new P.TextStyles(
                new P.TitleStyle(new A.Level1ParagraphProperties(
                    new A.NoBullet(),
                    new A.DefaultRunProperties(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "112233" }),
                        new A.LatinFont { Typeface = "Source Title" },
                        new A.EastAsianFont { Typeface = "Source Title" }) { FontSize = 2400, Bold = false })),
                new P.BodyStyle(),
                new P.OtherStyle());
            master.SlideMaster.Save();

            foreach (var slide in presentation.PresentationPart.SlideParts)
            {
                slide.Slide.CommonSlideData!.ShapeTree!.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 3U, Name = "Slide title" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                        new P.ShapeProperties(),
                        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Current title"))))));
                slide.Slide.CommonSlideData.ShapeTree.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 4U, Name = "Source footer" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Footer })),
                        new P.ShapeProperties(),
                        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Old footer"))))));
                slide.Slide.Save();
            }
        }
        var template = CreateFixture();
        using (var target = PresentationDocument.Open(template, true))
        {
            var master = target.PresentationPart!.SlideMasterParts.Single();
            master.SlideMaster.TextStyles = new P.TextStyles(
                new P.TitleStyle(new A.Level1ParagraphProperties(
                    new A.CharacterBullet { Char = "•" },
                    new A.DefaultRunProperties { FontSize = 4000, Bold = true })),
                new P.BodyStyle(),
                new P.OtherStyle());
            master.SlideMaster.Save();
        }
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");

        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
            [
                new SlideLayoutAssignment(1, targetLayout, SourceLayoutShapeIdsToPreserve: [9U]),
                new SlideLayoutAssignment(2, targetLayout, SourceLayoutShapeIdsToPreserve: [9U]),
            ], "target-template"), output);

        Assert.Empty(result.Issues);
        Assert.Equal(2, result.FrozenPlaceholderCount);
        Assert.Equal(2, result.MaterializedLayoutShapes?.Count);
        Assert.Equal(2, result.RemovedSystemPlaceholders?.Count);
        Assert.All(result.RemovedSystemPlaceholders!, entry => Assert.Equal("ftr", entry.PlaceholderType));
        Assert.All(result.MaterializedLayoutShapes!, entry => Assert.True(entry.OutputShapeId >= 2U));
        var outputEvidence = Inspector.InspectDetail(output);
        Assert.All(outputEvidence.Slides, slide =>
        {
            var title = slide.Shapes.Single(shape => shape.Text == "Current title");
            Assert.Equal("title", title.PlaceholderType);
            Assert.Equal(new TransformInfo(123400L, 234500L, 3456000L, 456700L), title.Transform);
            var titleRun = Assert.Single(title.Runs);
            Assert.Equal("Source Title", titleRun.FontFamily);
            Assert.Equal(24d, titleRun.FontSize);
            Assert.Equal("112233", titleRun.Color);
            Assert.False(titleRun.Bold);
            Assert.Single(slide.Shapes, shape => shape.Text == "Current visible layout content");
            Assert.DoesNotContain(slide.Shapes, shape => shape.Text == "Old footer");
        });
        using (var rendered = PresentationDocument.Open(output, false))
            Assert.All(rendered.PresentationPart!.SlideParts, slide => Assert.Contains(slide.Slide.Descendants<A.NoBullet>(), _ => true));
    }

    [Fact]
    public void ApplyTemplate_counts_untyped_placeholders_and_excludes_system_placeholders()
    {
        var source = CreateFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            foreach (var slide in presentation.PresentationPart!.SlideParts)
            {
                slide.Slide.CommonSlideData!.ShapeTree!.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 3U, Name = "Indexed body" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Index = 12U })),
                        new P.ShapeProperties(),
                        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Indexed body"))))));
                slide.Slide.CommonSlideData.ShapeTree.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 4U, Name = "Title" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Title })),
                        new P.ShapeProperties(),
                        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Title"))))));
                slide.Slide.CommonSlideData.ShapeTree.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 5U, Name = "Footer" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Footer })),
                        new P.ShapeProperties(),
                        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Footer"))))));
                slide.Slide.Save();
            }
        }

        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout), new SlideLayoutAssignment(2, targetLayout)],
                "target-template"), output);

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.FrozenPlaceholderCount);
        Assert.Equal(2, result.RemovedSystemPlaceholders?.Count);
        Assert.All(result.RemovedSystemPlaceholders!, entry => Assert.Equal("ftr", entry.PlaceholderType));
        var outputEvidence = Inspector.InspectDetail(output);
        Assert.Equal(2, outputEvidence.Slides.Count(slide => slide.Shapes.Any(shape => shape.Text == "Indexed body")));
        Assert.Equal(2, outputEvidence.Slides.Count(slide => slide.Shapes.Any(shape => shape.Text == "Title")));
        Assert.DoesNotContain(outputEvidence.Slides.SelectMany(slide => slide.Shapes), shape => shape.Text == "Footer");
    }

    [Fact]
    public void ApplyTemplate_freezes_all_published_placeholder_shape_kinds()
    {
        var source = CreateAllShapeKindsFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var slidePart = presentation.PresentationPart!.SlideParts.First();
            var slideTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
            var sourceFrame = slideTree.Elements<P.GraphicFrame>().Single();
            sourceFrame.Transform!.Remove();
            sourceFrame.NonVisualGraphicFrameProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Index = null;
            slideTree.Append(new P.GraphicFrame(
                new P.NonVisualGraphicFrameProperties(
                    new P.NonVisualDrawingProperties { Id = 9U, Name = "System footer frame" },
                    new P.NonVisualGraphicFrameDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Footer })),
                new P.Transform(),
                new A.Graphic(new A.GraphicData { Uri = "urn:tiwater:system" })));
            slidePart.Slide.Save();

            var layout = slidePart.SlideLayoutPart!;
            var layoutTree = layout.SlideLayout.CommonSlideData!.ShapeTree!;
            layoutTree.Append(CreatePlaceholderShapeWithTransform(61U, "Picture placeholder", 32U, 320L, 321L, 322L, 323L, P.PlaceholderValues.Object));
            layoutTree.Append(CreatePlaceholderShapeWithTransform(71U, "Frame placeholder", 0U, 330L, 331L, 332L, 333L, P.PlaceholderValues.Object));
            layoutTree.Append(new P.GroupShape(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 80U, Name = "Layout group without geometry" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Index = 34U })),
                new P.GroupShapeProperties(), new P.ShapeTree()));
            var masterTree = layout.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!;
            masterTree.Append(CreatePlaceholderShapeWithTransform(60U, "Lower-priority picture geometry", 32U, 900L, 901L, 902L, 903L, P.PlaceholderValues.Object));
            masterTree.Append(new P.GroupShape(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 81U, Name = "Inherited group" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Object, Index = 34U })),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = 340L, Y = 341L }, new A.Extents { Cx = 342L, Cy = 343L },
                    new A.ChildOffset { X = 44L, Y = 45L }, new A.ChildExtents { Cx = 46L, Cy = 47L })),
                new P.ShapeTree()));
            layout.SlideLayout.Save();
            layout.SlideMasterPart.SlideMaster.Save();
        }

        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout), new SlideLayoutAssignment(2, targetLayout)],
                "target-template"), output);

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.FrozenPlaceholderCount);
        var removed = Assert.Single(result.RemovedSystemPlaceholders!);
        Assert.Equal((1, 9U, "ftr"), (removed.SlideNumber, removed.ShapeId, removed.PlaceholderType));
        var shapes = Inspector.InspectDetail(output).Slides[0].Shapes;
        foreach (var item in new (string Kind, uint? Index)[] { ("shape", 31U), ("picture", 32U), ("graphicFrame", null), ("groupShape", 34U) })
        {
            var shape = Assert.Single(shapes, value => value.Kind == item.Kind && value.PlaceholderIndex == item.Index);
            Assert.True(shape.PlaceholderPresent);
        }
        Assert.Equal(new TransformInfo(320L, 321L, 322L, 323L), shapes.Single(value => value.Kind == "picture").Transform);
        Assert.Equal(new TransformInfo(330L, 331L, 332L, 333L), shapes.Single(value => value.Kind == "graphicFrame").Transform);
        Assert.Equal(new TransformInfo(340L, 341L, 342L, 343L), shapes.Single(value => value.Kind == "groupShape").Transform);
        Assert.DoesNotContain(shapes, value => value.ShapeId == 9U);
        using (var applied = PresentationDocument.Open(output, false))
        {
            var groupTransform = applied.PresentationPart!.SlideParts.First().Slide.Descendants<P.GroupShape>().Single().GroupShapeProperties!.TransformGroup!;
            Assert.Equal((44L, 45L, 46L, 47L), (groupTransform.ChildOffset!.X!.Value, groupTransform.ChildOffset.Y!.Value,
                groupTransform.ChildExtents!.Cx!.Value, groupTransform.ChildExtents.Cy!.Value));
        }
    }

    [Fact]
    public void ApplyTemplate_non_shape_without_placeholder_identity_is_not_frozen()
    {
        var source = CreateAllShapeKindsFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var picture = presentation.PresentationPart!.SlideParts.First().Slide.CommonSlideData!.ShapeTree!.Elements<P.Picture>().Single();
            picture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Remove();
            picture.Ancestors<P.Slide>().Single().Save();
        }

        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout), new SlideLayoutAssignment(2, targetLayout)]), output);

        Assert.Empty(result.Issues);
        Assert.Equal(3, result.FrozenPlaceholderCount);
        var outputPicture = Inspector.InspectDetail(output).Slides[0].Shapes.Single(value => value.Kind == "picture");
        Assert.False(outputPicture.PlaceholderPresent);
        Assert.Null(outputPicture.PlaceholderIndex);
    }

    [Fact]
    public void ApplyTemplate_rejects_ambiguous_effective_placeholder_identity_before_layout_reassignment()
    {
        var source = CreateAllShapeKindsFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var layout = presentation.PresentationPart!.SlideParts.First().SlideLayoutPart!;
            var layoutTree = layout.SlideLayout.CommonSlideData!.ShapeTree!;
            layoutTree.Append(CreatePlaceholderShapeWithTransform(61U, "First object", 32U, 10L, 11L, 12L, 13L, P.PlaceholderValues.Object));
            layoutTree.Append(CreatePlaceholderShapeWithTransform(62U, "Second object", 32U, 20L, 21L, 22L, 23L));
            layout.SlideLayout.Save();
        }

        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetEvidence.Masters.Single().Layouts.Single().Path)]), output);

        Assert.Equal(0, result.ChangedSlideCount);
        var issue = Assert.Single(result.Issues).Message;
        Assert.StartsWith("source placeholder identity is ambiguous in layout:", issue);
        Assert.EndsWith("/32", issue);
    }

    [Fact]
    public void ApplyTemplate_rejects_ambiguous_master_geometry_when_layout_identity_has_no_geometry()
    {
        var source = CreateAllShapeKindsFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var slide = presentation.PresentationPart!.SlideParts.First();
            var layoutTree = slide.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!;
            layoutTree.Append(CreatePlaceholderShape(61U, "Layout without geometry", new P.PlaceholderShape { Index = 32U }));
            slide.SlideLayoutPart.SlideLayout.Save();
            var masterTree = slide.SlideLayoutPart.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!;
            masterTree.Append(CreatePlaceholderShapeWithTransform(71U, "First master object", 32U, 10L, 11L, 12L, 13L));
            masterTree.Append(CreatePlaceholderShapeWithTransform(72U, "Second master object", 32U, 20L, 21L, 22L, 23L, P.PlaceholderValues.Object));
            slide.SlideLayoutPart.SlideMasterPart.SlideMaster.Save();
        }

        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetEvidence.Masters.Single().Layouts.Single().Path)]), output);

        Assert.Equal(0, result.ChangedSlideCount);
        var issue = Assert.Single(result.Issues).Message;
        Assert.StartsWith("source placeholder identity is ambiguous in master:", issue);
        Assert.EndsWith("/32", issue);
    }

    [Fact]
    public void ApplyTemplate_default_preserve_keeps_system_placeholders_and_rejects_invalid_policy()
    {
        var source = CreateFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            foreach (var slide in presentation.PresentationPart!.SlideParts)
            {
                slide.Slide.CommonSlideData!.ShapeTree!.Append(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 3U, Name = "Footer" },
                            new P.NonVisualShapeDrawingProperties(),
                            new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Type = P.PlaceholderValues.Footer })),
                        new P.ShapeProperties(),
                        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Footer"))))));
                slide.Slide.Save();
            }
        }

        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var preserveOutput = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var preserved = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout), new SlideLayoutAssignment(2, targetLayout)]), preserveOutput);

        Assert.Empty(preserved.Issues);
        Assert.Equal(0, preserved.FrozenPlaceholderCount);
        Assert.Empty(preserved.RemovedSystemPlaceholders!);
        Assert.Equal(2, Inspector.InspectDetail(preserveOutput).Slides.Count(slide => slide.Shapes.Any(shape => shape.Text == "Footer")));

        var invalidOutput = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");
        var invalid = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
                [new SlideLayoutAssignment(1, targetLayout), new SlideLayoutAssignment(2, targetLayout)],
                "remove-all-placeholders"), invalidOutput);

        Assert.Equal(0, invalid.ChangedSlideCount);
        Assert.Equal("system placeholder policy is invalid", Assert.Single(invalid.Issues).Message);
    }

    [Fact]
    public void ApplyTemplate_rejects_unknown_or_duplicate_source_layout_shape_ids()
    {
        var source = CreateFixture();
        var template = CreateFixture();
        var targetEvidence = Inspector.InspectDetail(template);
        var targetLayout = targetEvidence.Masters.Single().Layouts.Single().Path;
        var output = Path.Combine(Path.GetTempPath(), $"pptx-template-{Guid.NewGuid():N}.pptx");

        var result = TemplateApplicator.Apply(source, template,
            new TemplateApplicationPlan(targetEvidence.Masters.Single().Path,
            [
                new SlideLayoutAssignment(1, targetLayout, SourceLayoutShapeIdsToPreserve: [999U]),
                new SlideLayoutAssignment(2, targetLayout, SourceLayoutShapeIdsToPreserve: [1U, 1U]),
            ]), output);

        Assert.Equal(0, result.ChangedSlideCount);
        Assert.Equal(["source layout shapes not found: 999", "source layout shape ids contain duplicates"], result.Issues.Select(issue => issue.Message));
    }

    [Fact]
    public void InspectDetail_includes_top_level_group_geometry_and_descendant_text()
    {
        var source = CreateFixture();
        using (var presentation = PresentationDocument.Open(source, true))
        {
            var slide = presentation.PresentationPart!.SlideParts.First().Slide;
            slide.CommonSlideData!.ShapeTree!.Append(new P.GroupShape(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 99U, Name = "Grouped content" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(new A.TransformGroup(
                    new A.Offset { X = -1090367L, Y = 1696710L },
                    new A.Extents { Cx = 4738136L, Cy = 4400065L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 1000L, Cy = 1000L })),
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties { Id = 100U, Name = "Child" },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.ShapeProperties(),
                    new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Grouped text")))))));
            slide.Save();
        }

        var group = Assert.Single(Inspector.InspectDetail(source).Slides[0].Shapes, shape => shape.Kind == "groupShape");
        Assert.Equal(99U, group.ShapeId);
        Assert.Equal("Grouped text", group.Text);
        Assert.Equal(new TransformInfo(-1090367L, 1696710L, 4738136L, 4400065L), group.Transform);
    }

    [Fact]
    public void ApplyFormatEdits_updates_targeted_run_formatting_and_reports_changes()
    {
        var template = CreateFixture();
        var output = Path.Combine(Path.GetTempPath(), $"pptx-format-{Guid.NewGuid():N}.pptx");
        var plan = new FormatEditPlan([
            new FormatEditOperation(
                SlideNumber: 1,
                ShapeId: 2U,
                RunIndex: 0,
                FontFamily: "Arial",
                FontSize: 20d,
                Color: "F59B0F",
                Bold: false,
                ParagraphAlignment: "center")
        ]);

        var result = FormatEditor.Apply(template, plan, output);

        Assert.Equal(template, result.Input);
        Assert.Equal(output, result.Output);
        Assert.Equal(1, result.OperationCount);
        Assert.Equal(1, result.ChangedCount);
        var change = Assert.Single(result.Changes);
        Assert.Equal(1, change.SlideNumber);
        Assert.Equal(2U, change.ShapeId);
        Assert.Equal(0, change.RunIndex);

        var detail = Inspector.InspectDetail(output);
        var run = Assert.Single(detail.Slides[0].Shapes[0].Runs);
        Assert.Equal("Arial", run.FontFamily);
        Assert.Equal(20d, run.FontSize);
        Assert.Equal("F59B0F", run.Color);
        Assert.False(run.Bold);
        Assert.Equal("center", detail.Slides[0].Shapes[0].Paragraphs[0].Alignment);
    }

    private static string CreateFixture()
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-fixture-{Guid.NewGuid():N}.pptx");
        using var presentation = PresentationDocument.Create(path, PresentationDocumentType.Presentation);

        var presentationPart = presentation.AddPresentationPart();
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster1");
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>("rIdLayout1");
        slideLayoutPart.SlideLayout = new P.SlideLayout(
            new P.CommonSlideData(CreateShapeTree()),
            new P.ColorMapOverride(new A.MasterColorMapping()));
        slideLayoutPart.SlideLayout.Save();

        slideMasterPart.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(CreateShapeTree()),
            new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 1U, RelationshipId = "rIdLayout1" }),
            new P.TextStyles());
        slideMasterPart.SlideMaster.Save();
        slideLayoutPart.AddPart(slideMasterPart, "rIdMaster");

        var slidePart1 = presentationPart.AddNewPart<SlidePart>("rIdSlide1");
        slidePart1.Slide = CreateSlide("Project {{title}} 峰面积");
        slidePart1.AddPart(slideLayoutPart);
        slidePart1.Slide.Save();

        var notesPart = slidePart1.AddNewPart<NotesSlidePart>("rIdNotes1");
        notesPart.NotesSlide = CreateNotesSlide("Notes {{title}}");
        notesPart.NotesSlide.Save();

        var slidePart2 = presentationPart.AddNewPart<SlidePart>("rIdSlide2");
        slidePart2.Slide = CreateSlide("Batch {{batch}}");
        slidePart2.AddPart(slideLayoutPart);
        slidePart2.Slide.Save();

        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rIdMaster1" }),
            new P.SlideIdList(
                new P.SlideId { Id = 256U, RelationshipId = "rIdSlide1" },
                new P.SlideId { Id = 257U, RelationshipId = "rIdSlide2" }),
            new P.SlideSize { Cx = 9144000, Cy = 6858000, Type = P.SlideSizeValues.Screen4x3 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.Save();

        return path;
    }

    private static string CreatePlaceholderEvidenceFixture()
    {
        var path = CreateFixture();
        using var presentation = PresentationDocument.Open(path, true);
        var shapeTree = presentation.PresentationPart!.SlideParts.First().Slide.CommonSlideData!.ShapeTree!;
        shapeTree.Append(CreatePlaceholderShape(3U, "Indexed body", new P.PlaceholderShape { Index = 12U }));
        shapeTree.Append(CreatePlaceholderShape(4U, "Typed title", new P.PlaceholderShape { Type = P.PlaceholderValues.Title }));
        shapeTree.Append(CreatePlaceholderShape(5U, "Text {{looks-like-placeholder}}", null));
        presentation.PresentationPart.SlideParts.First().Slide.Save();
        return path;
    }

    private static string CreateAllShapeKindsFixture()
    {
        var path = CreateFixture();
        using var presentation = PresentationDocument.Open(path, true);
        var slidePart = presentation.PresentationPart!.SlideParts.First();
        var shapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        shapeTree.Append(CreatePlaceholderShape(3U, "Indexed shape", new P.PlaceholderShape { Index = 31U }));

        var imagePart = slidePart.AddImagePart(ImagePartType.Png, "rIdImageEvidence");
        imagePart.FeedData(new MemoryStream([137, 80, 78, 71, 13, 10, 26, 10]));
        shapeTree.Append(new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = 6U, Name = "Indexed picture" },
                new P.NonVisualPictureDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Index = 32U })),
            new A.BlipFill(new A.Blip { Embed = "rIdImageEvidence" }, new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties()));

        shapeTree.Append(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = 7U, Name = "Indexed graphic frame" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Index = 33U })),
            new P.Transform(),
            new A.Graphic(new A.GraphicData { Uri = "urn:tiwater:test" })));

        shapeTree.Append(new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 8U, Name = "Indexed group" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties(new P.PlaceholderShape { Index = 34U })),
            new P.GroupShapeProperties(),
            new P.ShapeTree()));

        slidePart.Slide.Save();
        return path;
    }

    private static void AssertShapeEvidenceMatchesSchema(JsonElement shape, JsonElement schema)
    {
        var required = schema.GetProperty("required").EnumerateArray().Select(value => value.GetString()!).ToList();
        var properties = schema.GetProperty("properties").EnumerateObject().Select(value => value.Name).ToHashSet(StringComparer.Ordinal);
        foreach (var property in required)
            Assert.True(shape.TryGetProperty(property, out _), $"published shape schema requires {property}");
        foreach (var property in shape.EnumerateObject())
            Assert.Contains(property.Name, properties);
        Assert.Equal(JsonValueKind.Number, shape.GetProperty("shapeId").ValueKind);
        Assert.Equal(JsonValueKind.String, shape.GetProperty("name").ValueKind);
        Assert.Equal(JsonValueKind.String, shape.GetProperty("kind").ValueKind);
        Assert.Equal(JsonValueKind.Number, shape.GetProperty("zOrder").ValueKind);
        Assert.Contains(shape.GetProperty("placeholderType").ValueKind, new[] { JsonValueKind.String, JsonValueKind.Null });
        Assert.Contains(shape.GetProperty("placeholderPresent").ValueKind, new[] { JsonValueKind.True, JsonValueKind.False });
        Assert.Contains(shape.GetProperty("placeholderIndex").ValueKind, new[] { JsonValueKind.Number, JsonValueKind.Null });
        Assert.Equal(JsonValueKind.Array, shape.GetProperty("paragraphs").ValueKind);
        Assert.Equal(JsonValueKind.Array, shape.GetProperty("runs").ValueKind);
    }

    private static string RepositoryPath(string relativePath)
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var candidate = Path.Combine(directory.FullName, relativePath);
            if (File.Exists(candidate)) return candidate;
        }

        throw new FileNotFoundException(relativePath);
    }

    private static P.Shape CreatePlaceholderShape(uint id, string text, P.PlaceholderShape? placeholder)
    {
        var app = placeholder is null
            ? new P.ApplicationNonVisualDrawingProperties()
            : new P.ApplicationNonVisualDrawingProperties(placeholder);
        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = text },
                new P.NonVisualShapeDrawingProperties(),
                app),
            new P.ShapeProperties(),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(text)))));
    }

    private static P.Shape CreatePlaceholderShapeWithTransform(uint id, string text, uint index, long x, long y, long cx, long cy, P.PlaceholderValues? type = null)
    {
        var shape = CreatePlaceholderShape(id, text, new P.PlaceholderShape { Index = index, Type = type });
        shape.ShapeProperties!.Transform2D = new A.Transform2D(
            new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy });
        return shape;
    }

    private static string CreateInheritedFormattingFixture(string? directColor = null, bool useLocalStyle = true)
    {
        var path = Path.Combine(Path.GetTempPath(), $"pptx-inherited-format-{Guid.NewGuid():N}.pptx");
        using var presentation = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = presentation.AddPresentationPart();
        var master = presentationPart.AddNewPart<SlideMasterPart>("rIdMaster1");
        var layout = master.AddNewPart<SlideLayoutPart>("rIdLayout1");
        layout.SlideLayout = new P.SlideLayout(new P.CommonSlideData(CreateShapeTree()), new P.ColorMapOverride(new A.MasterColorMapping()));
        layout.SlideLayout.Save();
        var masterRunProperties = new A.DefaultRunProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }),
            new A.LatinFont { Typeface = "Master Sans" },
            new A.EastAsianFont { Typeface = "Master Sans" })
        {
            FontSize = 1800
        };
        master.SlideMaster = new P.SlideMaster(
            new P.CommonSlideData(CreateShapeTree()),
            new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 1U, RelationshipId = "rIdLayout1" }),
            new P.TextStyles(new P.TitleStyle(), new P.BodyStyle(), new P.OtherStyle(new A.Level1ParagraphProperties(masterRunProperties))));
        master.SlideMaster.Save();
        layout.AddPart(master, "rIdMaster");

        var runProperties = new A.RunProperties();
        if (directColor is not null) runProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = directColor }));
        var levelProperties = new A.Level1ParagraphProperties(
            new A.DefaultRunProperties(
                new A.LatinFont { Typeface = "Inherited Sans" },
                new A.EastAsianFont { Typeface = "Inherited Sans" })
            {
                FontSize = 2400,
                Bold = true
            });
        var defaultProperties = new A.DefaultParagraphProperties(
            new A.DefaultRunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" })));
        var shape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = 2U, Name = "Inherited title" },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(),
            new P.TextBody(
                new A.BodyProperties(),
                useLocalStyle ? new A.ListStyle(defaultProperties, levelProperties) : new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties { Level = 0 },
                    new A.Run(runProperties, new A.Text("01")))));
        var slide = presentationPart.AddNewPart<SlidePart>("rIdSlide1");
        slide.Slide = new P.Slide(new P.CommonSlideData(CreateShapeTree(shape)), new P.ColorMapOverride(new A.MasterColorMapping()));
        slide.AddPart(layout);
        slide.Slide.Save();
        presentationPart.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rIdMaster1" }),
            new P.SlideIdList(new P.SlideId { Id = 256U, RelationshipId = "rIdSlide1" }),
            new P.SlideSize { Cx = 9144000, Cy = 6858000 },
            new P.NotesSize { Cx = 6858000, Cy = 9144000 });
        presentationPart.Presentation.Save();
        return path;
    }

    private static P.Slide CreateSlide(string text)
    {
        return new P.Slide(
            new P.CommonSlideData(
                CreateShapeTree(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2U, Name = "TextBox 1" },
                            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 914400L, Y = 457200L },
                                new A.Extents { Cx = 3657600L, Cy = 457200L })),
                        new P.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(
                                new A.ParagraphProperties(
                                    new A.DefaultRunProperties(
                                        new A.SolidFill(new A.RgbColorModelHex { Val = "287341" }),
                                        new A.LatinFont { Typeface = "微软雅黑" },
                                        new A.EastAsianFont { Typeface = "微软雅黑" })
                                    {
                                        FontSize = 1600,
                                        Bold = true
                                    }),
                                new A.Run(new A.Text(text))))))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
    }

    private static P.Shape CreateTextShape(uint id, string text, long x, long y, long cx, long cy) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = text },
                new P.NonVisualShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy })),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(text)))));

    private static P.NotesSlide CreateNotesSlide(string text)
    {
        return new P.NotesSlide(
            new P.CommonSlideData(
                CreateShapeTree(
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties { Id = 2U, Name = "Notes Placeholder 1" },
                            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                            new P.ApplicationNonVisualDrawingProperties()),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text(text))))))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
    }

    private static P.ShapeTree CreateShapeTree(params OpenXmlElement[] extraChildren)
    {
        var shapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup()));

        foreach (var child in extraChildren)
        {
            shapeTree.Append(child);
        }

        return shapeTree;
    }
}
