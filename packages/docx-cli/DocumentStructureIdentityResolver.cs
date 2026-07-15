using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

internal sealed record DocumentPartIdentity(
    string Id,
    string Kind,
    OpenXmlPart Part);

internal sealed record DocumentPartBindingIdentity(
    string Kind,
    string Type,
    string PartId,
    string RelationshipId,
    bool LinkedToPrevious,
    OpenXmlPart Part)
{
    public SectionPartBinding ToEvidence()
        => new(Kind, Type, PartId, RelationshipId, LinkedToPrevious);
}

internal sealed record DocumentSectionIdentity(
    string Id,
    int SectionIndex,
    SectionProperties Properties,
    Paragraph? EndingParagraph,
    IReadOnlyList<DocumentPartBindingIdentity> Headers,
    IReadOnlyList<DocumentPartBindingIdentity> Footers);

internal sealed record DocumentStructureIdentity(
    IReadOnlyList<DocumentSectionIdentity> Sections,
    IReadOnlyList<DocumentPartIdentity> Headers,
    IReadOnlyList<DocumentPartIdentity> Footers);

internal static class DocumentStructureIdentityResolver
{
    public static DocumentStructureIdentity Resolve(MainDocumentPart mainPart, Body body)
    {
        var sectionProperties = body.Elements<Paragraph>()
            .Select(paragraph => (Properties: paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>(), Paragraph: paragraph))
            .Where(item => item.Properties is not null)
            .Select(item => (item.Properties!, (Paragraph?)item.Paragraph))
            .Concat(body.Elements<SectionProperties>().Select(properties => (properties, (Paragraph?)null)))
            .ToList();

        var headerIds = new Dictionary<HeaderPart, string>();
        var footerIds = new Dictionary<FooterPart, string>();
        var currentHeaders = new Dictionary<string, DocumentPartBindingIdentity>(StringComparer.Ordinal);
        var currentFooters = new Dictionary<string, DocumentPartBindingIdentity>(StringComparer.Ordinal);
        var sections = new List<DocumentSectionIdentity>(sectionProperties.Count);

        for (var sectionIndex = 0; sectionIndex < sectionProperties.Count; sectionIndex++)
        {
            var section = sectionProperties[sectionIndex];
            var headers = ResolveBindings<HeaderReference, HeaderPart>(
                mainPart,
                section.Item1.Elements<HeaderReference>(),
                currentHeaders,
                headerIds,
                "header");
            var footers = ResolveBindings<FooterReference, FooterPart>(
                mainPart,
                section.Item1.Elements<FooterReference>(),
                currentFooters,
                footerIds,
                "footer");
            sections.Add(new DocumentSectionIdentity(
                $"section-{sectionIndex}",
                sectionIndex,
                section.Item1,
                section.Item2,
                headers,
                footers));
        }

        foreach (var pair in mainPart.Parts)
        {
            if (pair.OpenXmlPart is HeaderPart header && !headerIds.ContainsKey(header))
            {
                headerIds.Add(header, $"header-{headerIds.Count}");
            }
            else if (pair.OpenXmlPart is FooterPart footer && !footerIds.ContainsKey(footer))
            {
                footerIds.Add(footer, $"footer-{footerIds.Count}");
            }
        }

        return new DocumentStructureIdentity(
            sections,
            headerIds.OrderBy(pair => ParseStableIndex(pair.Value)).Select(pair => new DocumentPartIdentity(pair.Value, "header", pair.Key)).ToList(),
            footerIds.OrderBy(pair => ParseStableIndex(pair.Value)).Select(pair => new DocumentPartIdentity(pair.Value, "footer", pair.Key)).ToList());
    }

    private static IReadOnlyList<DocumentPartBindingIdentity> ResolveBindings<TReference, TPart>(
        MainDocumentPart mainPart,
        IEnumerable<TReference> references,
        Dictionary<string, DocumentPartBindingIdentity> current,
        Dictionary<TPart, string> stableIds,
        string kind)
        where TReference : OpenXmlElement
        where TPart : OpenXmlPart
    {
        var explicitlyBoundTypes = new HashSet<string>(StringComparer.Ordinal);
        foreach (var reference in references)
        {
            var type = GetAttribute(reference, "type") ?? "default";
            var relationshipId = GetAttribute(reference, "id")
                ?? throw new InvalidDataException($"{kind} reference is missing a relationship id.");
            OpenXmlPart relatedPart;
            try
            {
                relatedPart = mainPart.GetPartById(relationshipId);
            }
            catch (Exception ex) when (ex is ArgumentOutOfRangeException or KeyNotFoundException)
            {
                throw new InvalidDataException($"{kind} relationship '{relationshipId}' was not found.", ex);
            }

            if (relatedPart is not TPart typedPart)
            {
                throw new InvalidDataException($"Relationship '{relationshipId}' does not target a {kind} part.");
            }
            if (!stableIds.TryGetValue(typedPart, out var partId))
            {
                partId = $"{kind}-{stableIds.Count}";
                stableIds.Add(typedPart, partId);
            }

            current[type] = new DocumentPartBindingIdentity(kind, type, partId, relationshipId, false, typedPart);
            explicitlyBoundTypes.Add(type);
        }

        return current.Values
            .Select(binding => explicitlyBoundTypes.Contains(binding.Type)
                ? binding
                : binding with { LinkedToPrevious = true })
            .ToList();
    }

    private static string? GetAttribute(OpenXmlElement element, string localName)
    {
        foreach (var attribute in element.GetAttributes())
        {
            if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal))
            {
                return attribute.Value;
            }
        }
        return null;
    }

    private static int ParseStableIndex(string id)
        => int.Parse(id.AsSpan(id.LastIndexOf('-') + 1), System.Globalization.CultureInfo.InvariantCulture);
}
