using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Dockit.Pptx;

public static class Extractor
{
    public static int RunExportJson(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("export-json requires <input.pptx> [<output.json>]");
        }

        var input = args[0];
        var output = args.Length > 1 ? args[1] : null;
        var report = Export(input);

        if (output is null)
        {
            Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(report, Json.Options));
            return 0;
        }

        var outputDir = Path.GetDirectoryName(output);
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        File.WriteAllText(output, System.Text.Json.JsonSerializer.Serialize(report, Json.Options));
        return 0;
    }

    public static PresentationExport Export(string path)
    {
        using var presentation = PresentationDocument.Open(path, false);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slides = new List<SlideExport>();
        var noteCandidates = new List<(string Path, List<string> Texts)>();

        var index = 0;
        foreach (var slidePart in EnumerateSlides(presentationPart))
        {
            index++;
            var texts = Inspector.ExtractTexts(slidePart.Slide);
            slides.Add(new SlideExport(
                SlideNumber: index,
                Path: Inspector.NormalizePartPath(slidePart.Uri),
                Texts: texts,
                Placeholders: Inspector.ExtractPlaceholders(texts)));

            var notesPart = slidePart.NotesSlidePart;
            if (notesPart?.NotesSlide is null)
            {
                continue;
            }

            noteCandidates.Add((Inspector.NormalizePartPath(notesPart.Uri), Inspector.ExtractTexts(notesPart.NotesSlide)));
        }

        var notes = noteCandidates
            .OrderBy(candidate => candidate.Path, StringComparer.Ordinal)
            .Select((candidate, idx) => new NoteExport(
                NotesNumber: idx + 1,
                Path: candidate.Path,
                Texts: candidate.Texts))
            .ToList();

        return new PresentationExport(
            File: path,
            Slides: slides,
            Notes: notes);
    }

    private static IEnumerable<SlidePart> EnumerateSlides(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<SlideId>() ?? [];
        foreach (var slideId in slideIds)
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId))
            {
                continue;
            }

            if (presentationPart.GetPartById(relationshipId) is SlidePart slidePart)
            {
                yield return slidePart;
            }
        }
    }
}
