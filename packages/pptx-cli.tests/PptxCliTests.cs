using Dockit.Pptx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Tiwater.FormatEvidence;
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
