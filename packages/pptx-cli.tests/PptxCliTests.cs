using Dockit.Pptx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
        Assert.Contains("Project {{title}}", report.Slides[0].Texts);
        Assert.Equal(["title"], report.Slides[0].Placeholders);
        Assert.Contains("Batch {{batch}}", report.Slides[1].Texts);
        Assert.Equal(["batch"], report.Slides[1].Placeholders);
        Assert.Single(report.Notes);
        Assert.Equal(1, report.Notes[0].NotesNumber);
        Assert.Equal("ppt/notesSlides/notesSlide1.xml", report.Notes[0].Path);
        Assert.Contains("Notes {{title}}", report.Notes[0].Texts);
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
        Assert.Contains("Project Q2 Review", slideTexts);
        Assert.Contains("Batch B-001", slideTexts);

        var notesTexts = presentation.PresentationPart!
            .SlideParts
            .Where(part => part.NotesSlidePart is not null)
            .SelectMany(part => part.NotesSlidePart!.NotesSlide.Descendants<A.Text>().Select(text => text.Text))
            .ToList();
        Assert.Contains("Notes Q2 Review", notesTexts);
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

        var slidePart1 = presentationPart.AddNewPart<SlidePart>("rIdSlide1");
        slidePart1.Slide = CreateSlide("Project {{title}}");
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
                        new P.ShapeProperties(),
                        new P.TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text(text))))))),
            new P.ColorMapOverride(new A.MasterColorMapping()));
    }

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
