using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Text.Json;

namespace Dockit.Pptx;

public static class TemplateApplicator
{
    public static TemplateApplicationResult Apply(string inputPath, string templatePath, string planPath, string outputPath)
    {
        var plan = JsonSerializer.Deserialize<TemplateApplicationPlan>(File.ReadAllText(planPath), Json.Options)
            ?? throw new InvalidOperationException("Template application plan could not be parsed.");
        return Apply(inputPath, templatePath, plan, outputPath);
    }

    public static TemplateApplicationResult Apply(string inputPath, string templatePath, TemplateApplicationPlan plan, string outputPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");
        File.Copy(inputPath, outputPath, true);
        var issues = new List<TemplateApplicationIssue>();
        var changed = 0;
        using var template = PresentationDocument.Open(templatePath, false);
        using var output = PresentationDocument.Open(outputPath, true);
        var templatePart = template.PresentationPart ?? throw new InvalidOperationException("Template presentation part not found.");
        var outputPart = output.PresentationPart ?? throw new InvalidOperationException("Output presentation part not found.");
        var previousMasters = outputPart.SlideMasterParts.ToList();
        var targetMaster = templatePart.SlideMasterParts.SingleOrDefault(part => PartPath(part) == plan.TargetMasterPath);
        if (targetMaster is null) return new TemplateApplicationResult(inputPath, templatePath, outputPath, 0, [new(null, "target master not found")]);

        var importedMaster = outputPart.AddPart(targetMaster);
        var masterRelationshipId = outputPart.GetIdOfPart(importedMaster);
        var masterIds = outputPart.Presentation.SlideMasterIdList ?? outputPart.Presentation.PrependChild(new SlideMasterIdList());
        var nextMasterId = masterIds.Elements<SlideMasterId>().Select(item => item.Id?.Value ?? 2147483647U).DefaultIfEmpty(2147483647U).Max() + 1U;
        masterIds.Append(new SlideMasterId { Id = nextMasterId, RelationshipId = masterRelationshipId });

        var slides = EnumerateSlides(outputPart).ToList();
        foreach (var assignment in plan.Slides)
        {
            if (assignment.SlideNumber < 1 || assignment.SlideNumber > slides.Count) { issues.Add(new(assignment.SlideNumber, "slide not found")); continue; }
            var targetLayout = targetMaster.SlideLayoutParts.SingleOrDefault(part => PartPath(part) == assignment.TargetLayoutPath);
            if (targetLayout is null) { issues.Add(new(assignment.SlideNumber, "target layout not found")); continue; }
            var targetRelationshipId = targetMaster.GetIdOfPart(targetLayout);
            if (importedMaster.GetPartById(targetRelationshipId) is not SlideLayoutPart importedLayout) { issues.Add(new(assignment.SlideNumber, "imported target layout not found")); continue; }
            var slide = slides[assignment.SlideNumber - 1];
            if (slide.SlideLayoutPart is { } oldLayout) slide.DeletePart(oldLayout);
            slide.AddPart(importedLayout);
            changed++;
        }
        if (issues.Count == 0 && changed == slides.Count)
        {
            foreach (var previousMaster in previousMasters)
            {
                var relationshipId = outputPart.GetIdOfPart(previousMaster);
                masterIds.Elements<SlideMasterId>().Where(item => item.RelationshipId?.Value == relationshipId).ToList().ForEach(item => item.Remove());
                outputPart.DeletePart(previousMaster);
            }
            if (templatePart.Presentation.SlideSize is { } targetSize)
                outputPart.Presentation.SlideSize = (SlideSize)targetSize.CloneNode(true);
        }
        outputPart.Presentation.Save();
        return new TemplateApplicationResult(inputPath, templatePath, outputPath, changed, issues);
    }

    private static string PartPath(OpenXmlPart part) => part.Uri.OriginalString.TrimStart('/');
    private static IEnumerable<SlidePart> EnumerateSlides(PresentationPart presentationPart)
    {
        foreach (var slideId in presentationPart.Presentation.SlideIdList?.Elements<SlideId>() ?? [])
            if (slideId.RelationshipId?.Value is { } id && presentationPart.GetPartById(id) is SlidePart slide) yield return slide;
    }
}
