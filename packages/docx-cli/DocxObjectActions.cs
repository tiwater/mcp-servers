using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Dockit.Docx;

internal static class DocxObjectActions
{
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    internal static IEnumerable<OpenXmlElement> DescendantsAndSelf(OpenXmlElement element)
    {
        yield return element;
        foreach (var descendant in element.Descendants()) yield return descendant;
    }

    internal static IEnumerable<string> RelationshipIds(IEnumerable<OpenXmlElement> roots)
        => roots.SelectMany(DescendantsAndSelf)
            .SelectMany(element => element.GetAttributes())
            .Where(attribute => attribute.NamespaceUri == RelationshipsNamespace && !string.IsNullOrWhiteSpace(attribute.Value))
            .Select(attribute => attribute.Value!);

    internal static bool CanCopyRelationship(MainDocumentPart source, string id, out string error)
    {
        error = string.Empty;
        if (PartByIdOrNull(source, id) is ImagePart or HeaderPart or FooterPart) return true;
        if (source.HyperlinkRelationships.Any(relationship => relationship.Id == id)) return true;
        error = $"source relationship {id} is missing or has an unsupported part type";
        return false;
    }

    internal static string CopyRelationship(MainDocumentPart source, MainDocumentPart target, string id)
    {
        var part = PartByIdOrNull(source, id);
        return part switch
        {
            ImagePart image => target.GetIdOfPart(target.AddPart(image)),
            HeaderPart header => target.GetIdOfPart(target.AddPart(header)),
            FooterPart footer => target.GetIdOfPart(target.AddPart(footer)),
            _ => target.AddHyperlinkRelationship(source.HyperlinkRelationships.Single(item => item.Id == id).Uri, true).Id,
        };
    }

    private static OpenXmlPart? PartByIdOrNull(OpenXmlPartContainer container, string id)
    {
        try { return container.GetPartById(id); }
        catch (ArgumentOutOfRangeException) { return null; }
    }

    internal static void RewriteRelationships(OpenXmlElement root, IReadOnlyDictionary<string, string> map)
    {
        foreach (var element in DescendantsAndSelf(root))
        foreach (var attribute in element.GetAttributes().Where(item => item.NamespaceUri == RelationshipsNamespace).ToList())
            if (attribute.Value is { } value && map.TryGetValue(value, out var replacement))
                element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, replacement));
    }

    internal static void RemapDrawingIds(Body targetBody, IReadOnlyList<OpenXmlElement> clones)
    {
        var nextId = targetBody.Descendants<DW.DocProperties>().Select(item => item.Id?.Value ?? 0U).DefaultIfEmpty().Max() + 1U;
        foreach (var drawing in clones.SelectMany(DescendantsAndSelf).OfType<Drawing>())
        {
            foreach (var properties in drawing.Descendants<DW.DocProperties>()) properties.Id = nextId++;
            foreach (var properties in drawing.Descendants<PIC.NonVisualDrawingProperties>()) properties.Id = nextId++;
        }
    }

    private static HashSet<string> RequiredStyleIds(IEnumerable<OpenXmlElement> roots)
        => roots.SelectMany(DescendantsAndSelf).SelectMany(element => element switch
        {
            ParagraphStyleId value => [value.Val?.Value],
            RunStyle value => [value.Val?.Value],
            TableStyle value => [value.Val?.Value],
            _ => Array.Empty<string?>(),
        }).Where(value => !string.IsNullOrWhiteSpace(value)).Cast<string>().ToHashSet(StringComparer.Ordinal);

    private static IReadOnlyList<Style> RequiredSourceStyles(MainDocumentPart source, IEnumerable<OpenXmlElement> roots)
    {
        var styles = source.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(style => style.StyleId?.Value is not null)
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal) ?? [];
        var queue = new Queue<string>(RequiredStyleIds(roots));
        var result = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (queue.TryDequeue(out var id))
        {
            if (!seen.Add(id) || !styles.TryGetValue(id, out var style)) continue;
            result.Add(style);
            foreach (var dependency in new[] { style.BasedOn?.Val?.Value, style.NextParagraphStyle?.Val?.Value, style.LinkedStyle?.Val?.Value })
                if (!string.IsNullOrWhiteSpace(dependency)) queue.Enqueue(dependency);
        }
        return result;
    }

    internal static bool TryImportStyles(MainDocumentPart source, MainDocumentPart target, IReadOnlyList<OpenXmlElement> roots, bool apply, out string error)
    {
        error = string.Empty;
        var requested = RequiredStyleIds(roots);
        if (requested.Count == 0) return true;
        var sourceStyles = source.StyleDefinitionsPart?.Styles?.Elements<Style>()
            .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal) ?? [];
        var targetStylesPart = target.StyleDefinitionsPart;
        var targetStyles = (targetStylesPart?.Styles?.Elements<Style>() ?? [])
            .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .ToDictionary(style => style.StyleId!.Value!, StringComparer.Ordinal);
        var queue = new Queue<string>(requested);
        var ordered = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        while (queue.TryDequeue(out var id))
        {
            if (!seen.Add(id)) continue;
            if (!sourceStyles.TryGetValue(id, out var sourceStyle)) { error = $"source style is missing: {id}"; return false; }
            ordered.Add(sourceStyle);
            foreach (var dependency in new[] { sourceStyle.BasedOn?.Val?.Value, sourceStyle.NextParagraphStyle?.Val?.Value, sourceStyle.LinkedStyle?.Val?.Value })
                if (!string.IsNullOrWhiteSpace(dependency)) queue.Enqueue(dependency);
        }
        if (!apply) return true;

        targetStylesPart ??= target.AddNewPart<StyleDefinitionsPart>();
        targetStylesPart.Styles ??= new Styles();
        var usedIds = targetStyles.Keys.Concat(sourceStyles.Keys).ToHashSet(StringComparer.Ordinal);
        var remapped = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var sourceStyle in ordered)
        {
            var id = sourceStyle.StyleId!.Value!;
            if (!targetStyles.TryGetValue(id, out var targetStyle)
                || sourceStyle.OuterXml == targetStyle.OuterXml)
            {
                remapped[id] = id;
                continue;
            }

            var digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sourceStyle.OuterXml)))
                .ToLowerInvariant()[..16];
            var candidate = $"tw_{digest}";
            for (var suffix = 2; usedIds.Contains(candidate); suffix++) candidate = $"tw_{digest}_{suffix}";
            usedIds.Add(candidate);
            remapped[id] = candidate;
        }

        foreach (var sourceStyle in ordered.AsEnumerable().Reverse())
        {
            var sourceId = sourceStyle.StyleId!.Value!;
            var targetId = remapped[sourceId];
            if (targetStyles.TryGetValue(sourceId, out var identical)
                && sourceId == targetId
                && sourceStyle.OuterXml == identical.OuterXml) continue;

            var clone = NormalizeStyle(sourceStyle);
            clone.StyleId = targetId;
            RewriteStyleDependencies(clone, remapped);
            if (clone.Default?.Value == true && targetStylesPart.Styles.Elements<Style>()
                .Any(style => style.Type?.Value == clone.Type?.Value && style.Default?.Value == true))
                clone.Default = null;
            targetStylesPart.Styles.AppendChild(clone);
        }
        foreach (var root in roots) RewriteStyleReferences(root, remapped);
        targetStylesPart.Styles.Save();
        return true;
    }

    private static Style NormalizeStyle(Style source)
    {
        var normalized = (Style)source.CloneNode(false);
        foreach (var child in source.ChildElements)
            if (!normalized.AddChild(child.CloneNode(true), true))
                throw new InvalidOperationException($"source style child is unsupported: {child.LocalName}");
        return normalized;
    }

    private static void RewriteStyleReferences(OpenXmlElement root, IReadOnlyDictionary<string, string> remapped)
    {
        foreach (var element in DescendantsAndSelf(root))
        {
            switch (element)
            {
                case ParagraphStyleId paragraph when paragraph.Val?.Value is { } id && remapped.TryGetValue(id, out var replacement):
                    paragraph.Val = replacement;
                    break;
                case RunStyle run when run.Val?.Value is { } id && remapped.TryGetValue(id, out var replacement):
                    run.Val = replacement;
                    break;
                case TableStyle table when table.Val?.Value is { } id && remapped.TryGetValue(id, out var replacement):
                    table.Val = replacement;
                    break;
            }
        }
    }

    private static void RewriteStyleDependencies(Style style, IReadOnlyDictionary<string, string> remapped)
    {
        if (style.BasedOn?.Val?.Value is { } basedOn && remapped.TryGetValue(basedOn, out var mappedBasedOn))
            style.BasedOn.Val = mappedBasedOn;
        if (style.NextParagraphStyle?.Val?.Value is { } next && remapped.TryGetValue(next, out var mappedNext))
            style.NextParagraphStyle.Val = mappedNext;
        if (style.LinkedStyle?.Val?.Value is { } linked && remapped.TryGetValue(linked, out var mappedLinked))
            style.LinkedStyle.Val = mappedLinked;
    }

    internal static bool TryImportNumbering(MainDocumentPart source, MainDocumentPart target, IReadOnlyList<OpenXmlElement> roots, bool apply, out string error)
    {
        error = string.Empty;
        var requested = roots.SelectMany(DescendantsAndSelf).OfType<NumberingId>()
            .Select(item => item.Val?.Value).Where(value => value is not null).Cast<int>().ToHashSet();
        if (requested.Count == 0) return true;
        var sourceNumbering = source.NumberingDefinitionsPart?.Numbering;
        if (sourceNumbering is null) { error = "source numbering definitions are missing"; return false; }
        var targetPart = target.NumberingDefinitionsPart;
        var targetNumbering = targetPart?.Numbering;
        var abstractsToAdd = new List<AbstractNum>();
        var instancesToAdd = new List<NumberingInstance>();
        foreach (var numId in requested)
        {
            var sourceInstance = sourceNumbering.Elements<NumberingInstance>().SingleOrDefault(item => item.NumberID?.Value == numId);
            if (sourceInstance is null) { error = $"source numbering instance is missing: {numId}"; return false; }
            var targetInstance = targetNumbering?.Elements<NumberingInstance>().SingleOrDefault(item => item.NumberID?.Value == numId);
            if (targetInstance is not null && targetInstance.OuterXml != sourceInstance.OuterXml) { error = $"target numbering instance conflicts: {numId}"; return false; }
            var abstractId = sourceInstance.AbstractNumId?.Val?.Value;
            if (abstractId is null) { error = $"source numbering instance {numId} has no abstract numbering id"; return false; }
            var sourceAbstract = sourceNumbering.Elements<AbstractNum>().SingleOrDefault(item => item.AbstractNumberId?.Value == abstractId);
            if (sourceAbstract is null) { error = $"source abstract numbering is missing: {abstractId}"; return false; }
            var targetAbstract = targetNumbering?.Elements<AbstractNum>().SingleOrDefault(item => item.AbstractNumberId?.Value == abstractId);
            if (targetAbstract is not null && targetAbstract.OuterXml != sourceAbstract.OuterXml) { error = $"target abstract numbering conflicts: {abstractId}"; return false; }
            if (targetAbstract is null && abstractsToAdd.All(item => item.AbstractNumberId?.Value != abstractId)) abstractsToAdd.Add(sourceAbstract);
            if (targetInstance is null) instancesToAdd.Add(sourceInstance);
        }
        if (!apply) return true;
        targetPart ??= target.AddNewPart<NumberingDefinitionsPart>();
        targetPart.Numbering ??= new Numbering();
        foreach (var abstractNumbering in abstractsToAdd) targetPart.Numbering.AddChild((AbstractNum)abstractNumbering.CloneNode(true), true);
        foreach (var instance in instancesToAdd) targetPart.Numbering.AddChild((NumberingInstance)instance.CloneNode(true), true);
        targetPart.Numbering.Save();
        return true;
    }

}
