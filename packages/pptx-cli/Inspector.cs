using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Pptx;

public static class Inspector
{
    private static readonly Regex PlaceholderRegex = new(@"{{\s*([^{}]+?)\s*}}", RegexOptions.Compiled);

    public static PresentationReport Inspect(string path)
    {
        using var presentation = PresentationDocument.Open(path, false);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation part not found.");

        var slides = new List<SlideReport>();
        var allPlaceholders = new HashSet<string>(StringComparer.Ordinal);

        var slideParts = EnumerateSlides(presentationPart).ToList();
        for (var i = 0; i < slideParts.Count; i++)
        {
            var slidePart = slideParts[i];
            var texts = ExtractTexts(slidePart.Slide);
            var placeholders = ExtractPlaceholders(texts);
            foreach (var placeholder in placeholders)
            {
                allPlaceholders.Add(placeholder);
            }

            slides.Add(new SlideReport(
                SlideNumber: i + 1,
                Path: NormalizePartPath(slidePart.Uri),
                TextCount: texts.Count,
                Placeholders: placeholders));
        }

        return new PresentationReport(
            File: path,
            SlideCount: slides.Count,
            Placeholders: allPlaceholders.OrderBy(value => value, StringComparer.Ordinal).ToList(),
            Slides: slides);
    }

    internal static List<string> ExtractTexts(OpenXmlPartRootElement? root)
    {
        if (root is null)
        {
            return [];
        }

        return root.Descendants<A.Text>()
            .Select(node => node.Text)
            .Where(text => !string.IsNullOrEmpty(text))
            .Select(text => text!)
            .ToList();
    }

    internal static List<string> ExtractPlaceholders(IEnumerable<string> texts)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var placeholders = new List<string>();

        foreach (var text in texts)
        {
            foreach (Match match in PlaceholderRegex.Matches(text))
            {
                var key = match.Groups[1].Value.Trim();
                if (key.Length == 0 || !seen.Add(key))
                {
                    continue;
                }

                placeholders.Add(key);
            }
        }

        return placeholders;
    }

    internal static string NormalizePartPath(Uri? uri)
    {
        if (uri is null)
        {
            return string.Empty;
        }

        var path = uri.OriginalString;
        return path.StartsWith('/') ? path[1..] : path;
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
