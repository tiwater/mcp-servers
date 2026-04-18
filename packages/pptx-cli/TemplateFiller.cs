using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Pptx;

public static class TemplateFiller
{
    public static FillResult Fill(string templatePath, string dataPath, string outputPath)
    {
        var mapping = ReadMapping(dataPath);
        return Fill(templatePath, mapping, outputPath);
    }

    public static FillResult Fill(string templatePath, IReadOnlyDictionary<string, string> mapping, string outputPath)
    {
        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        File.Copy(templatePath, outputPath, overwrite: true);

        var changedSlides = 0;
        var changedNotes = 0;
        using var presentation = PresentationDocument.Open(outputPath, true);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        foreach (var slidePart in EnumerateSlides(presentationPart))
        {
            if (slidePart.Slide is not null)
            {
                var changed = ReplaceTokens(slidePart.Slide.Descendants<A.Text>(), mapping);
                if (changed > 0)
                {
                    changedSlides++;
                    slidePart.Slide.Save();
                }
            }

            var notesPart = slidePart.NotesSlidePart;
            if (notesPart?.NotesSlide is null)
            {
                continue;
            }

            var notesChanged = ReplaceTokens(notesPart.NotesSlide.Descendants<A.Text>(), mapping);
            if (notesChanged > 0)
            {
                changedNotes++;
                notesPart.NotesSlide.Save();
            }
        }

        return new FillResult(
            Template: templatePath,
            Output: outputPath,
            ChangedSlides: changedSlides,
            ChangedNotes: changedNotes,
            PlaceholderCount: mapping.Count);
    }

    private static Dictionary<string, string> ReadMapping(string dataPath)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(dataPath));
        if (document.RootElement.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException("Fill data must be a JSON object or an object with textValues");
        }

        var source = document.RootElement;
        if (source.TryGetProperty("textValues", out var textValuesElement))
        {
            if (textValuesElement.ValueKind != JsonValueKind.Object)
            {
                throw new InvalidOperationException("textValues must be a JSON object");
            }

            source = textValuesElement;
        }

        var mapping = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var property in source.EnumerateObject())
        {
            mapping[property.Name] = property.Value.ValueKind == JsonValueKind.String
                ? property.Value.GetString() ?? string.Empty
                : property.Value.ToString();
        }

        return mapping;
    }

    private static int ReplaceTokens(IEnumerable<A.Text> textNodes, IReadOnlyDictionary<string, string> mapping)
    {
        var changed = 0;

        foreach (var node in textNodes)
        {
            var text = node.Text;
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            var updated = ReplaceInlineTokens(text, mapping);
            if (updated == text)
            {
                continue;
            }

            node.Text = updated;
            changed++;
        }

        return changed;
    }

    private static string ReplaceInlineTokens(string text, IReadOnlyDictionary<string, string> mapping)
    {
        var updated = text;
        foreach (var entry in mapping)
        {
            updated = updated.Replace($"{{{{{entry.Key}}}}}", entry.Value, StringComparison.Ordinal);
        }

        return updated;
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
