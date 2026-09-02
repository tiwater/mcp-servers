using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Dockit.Convert;

internal static class DocxFieldResultMerger
{
    private const string DocumentPart = "word/document.xml";
    private const string StylesPart = "word/styles.xml";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    internal static string PrepareSourceParagraphIdentities(string sourcePath, string outputDirectory)
    {
        var document = LoadDocument(sourcePath);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var changed = false;
        uint nextId = 1;
        foreach (var paragraph in document.Descendants(W + "p"))
        {
            var id = (string?)paragraph.Attribute(W14 + "paraId");
            if (!string.IsNullOrWhiteSpace(id)
                && Regex.IsMatch(id, "^[0-9A-Fa-f]{8}$")
                && used.Add(id))
                continue;

            string replacement;
            do replacement = nextId++.ToString("X8");
            while (!used.Add(replacement));
            paragraph.SetAttributeValue(W14 + "paraId", replacement);
            changed = true;
        }
        if (!changed) return sourcePath;

        var preparedPath = Path.Combine(outputDirectory, "field-refresh-source.docx");
        WriteDocumentPart(sourcePath, document, preparedPath);
        return preparedPath;
    }

    internal static void Merge(string sourcePath, string refreshedPath, string outputPath)
    {
        var source = LoadDocument(sourcePath);
        var refreshed = LoadDocument(refreshedPath);
        var sourceRegions = FindIndexRegions(source);
        var refreshedRegions = FindIndexRegions(refreshed);

        if (sourceRegions.Count != refreshedRegions.Count)
            throw new InvalidOperationException("WPS field refresh changed the number of DOCX index fields.");
        for (var index = 0; index < sourceRegions.Count; index++)
        {
            if (sourceRegions[index].Kind != refreshedRegions[index].Kind)
                throw new InvalidOperationException("WPS field refresh changed the order or kind of DOCX index fields.");
        }

        var bodyFontPolicy = UniformExplicitBodyFontPolicy(source);
        NormalizeTocResultStyles(sourcePath, source, refreshed, refreshedRegions);
        CopyTocBookmarks(source, refreshed, sourceRegions, refreshedRegions);
        ReplaceIndexRegions(source, refreshed, sourceRegions, refreshedRegions, bodyFontPolicy);

        WriteDocumentPart(sourcePath, source, outputPath);
    }

    private static void WriteDocumentPart(string sourcePath, XDocument document, string outputPath)
    {
        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        var temporaryOutput = Path.Combine(
            outputDirectory ?? Path.GetTempPath(), $".{Path.GetFileName(outputPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.Copy(sourcePath, temporaryOutput, overwrite: false);
            using (var archive = ZipFile.Open(temporaryOutput, ZipArchiveMode.Update))
            {
                var existing = archive.GetEntry(DocumentPart)
                    ?? throw new InvalidOperationException($"DOCX package is missing {DocumentPart}.");
                existing.Delete();
                var replacement = archive.CreateEntry(DocumentPart, CompressionLevel.Optimal);
                using var stream = replacement.Open();
                document.Save(stream, SaveOptions.DisableFormatting);
            }
            File.Move(temporaryOutput, outputPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
        }
    }

    private static XDocument LoadDocument(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(DocumentPart)
            ?? throw new InvalidOperationException($"DOCX package is missing {DocumentPart}.");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static XDocument LoadPart(string path, string part)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(part)
            ?? throw new InvalidOperationException($"DOCX package is missing {part}.");
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static List<IndexRegion> FindIndexRegions(XDocument document)
    {
        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidOperationException("DOCX document.xml is missing w:body.");
        var blocks = body.Elements().ToList();
        var stack = new Stack<FieldFrame>();
        var regions = new List<IndexRegion>();

        for (var blockIndex = 0; blockIndex < blocks.Count; blockIndex++)
        {
            foreach (var element in blocks[blockIndex].DescendantsAndSelf())
            {
                if (element.Name == W + "fldChar")
                {
                    var type = (string?)element.Attribute(W + "fldCharType");
                    if (type == "begin")
                    {
                        stack.Push(new FieldFrame(blockIndex));
                    }
                    else if (type == "separate")
                    {
                        if (stack.Count == 0) throw new InvalidOperationException("DOCX contains an unmatched field separator.");
                        stack.Peek().InstructionComplete = true;
                    }
                    else if (type == "end")
                    {
                        if (stack.Count == 0) throw new InvalidOperationException("DOCX contains an unmatched field end.");
                        var field = stack.Pop();
                        var kind = IndexKind(field.Instruction.ToString());
                        if (kind is not null)
                            regions.Add(new IndexRegion(field.StartBlock, blockIndex, kind));
                    }
                }
                else if (element.Name == W + "instrText" && stack.Count > 0 && !stack.Peek().InstructionComplete)
                {
                    stack.Peek().Instruction.Append(element.Value);
                }
            }
        }
        if (stack.Count != 0) throw new InvalidOperationException("DOCX contains an unclosed field.");
        return regions.OrderBy(region => region.StartBlock).ToList();
    }

    private static string? IndexKind(string instruction)
    {
        var normalized = string.Join(' ', instruction.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        if (!normalized.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)) return null;
        return normalized.Contains("\\c", StringComparison.OrdinalIgnoreCase) ? "table-of-figures" : "table-of-contents";
    }

    private static void ReplaceIndexRegions(
        XDocument source,
        XDocument refreshed,
        IReadOnlyList<IndexRegion> sourceRegions,
        IReadOnlyList<IndexRegion> refreshedRegions,
        RunFontPolicy? bodyFontPolicy)
    {
        var sourceBlocks = source.Root!.Element(W + "body")!.Elements().ToList();
        var refreshedBlocks = refreshed.Root!.Element(W + "body")!.Elements().ToList();
        for (var regionIndex = sourceRegions.Count - 1; regionIndex >= 0; regionIndex--)
        {
            var sourceRegion = sourceRegions[regionIndex];
            var refreshedRegion = refreshedRegions[regionIndex];
            var replacements = refreshedBlocks
                .Skip(refreshedRegion.StartBlock)
                .Take(refreshedRegion.EndBlock - refreshedRegion.StartBlock + 1)
                .Select(element => new XElement(element))
                .ToList();
            if (bodyFontPolicy is not null)
                foreach (var replacement in replacements) ApplyFontPolicy(replacement, bodyFontPolicy);
            sourceBlocks[sourceRegion.StartBlock].AddBeforeSelf(replacements);
            for (var block = sourceRegion.EndBlock; block >= sourceRegion.StartBlock; block--)
                sourceBlocks[block].Remove();
        }
    }

    private static void NormalizeTocResultStyles(
        string sourcePath,
        XDocument source,
        XDocument refreshed,
        IReadOnlyList<IndexRegion> refreshedRegions)
    {
        var headingRegions = refreshedRegions.Where(region => region.Kind == "table-of-contents").ToList();
        if (headingRegions.Count == 0) return;
        var styles = LoadPart(sourcePath, StylesPart);
        var tocStyles = TocStylesByLevel(styles);
        var sourceParagraphs = ParagraphsById(source);
        var refreshedBody = refreshed.Root!.Element(W + "body")!;
        var refreshedBlocks = refreshedBody.Elements().ToList();
        var bookmarkParagraphIds = refreshed.Descendants(W + "bookmarkStart")
            .Where(IsTocBookmark)
            .Select(start => new
            {
                Name = (string?)start.Attribute(W + "name"),
                ParagraphId = (string?)start.Ancestors(W + "p").FirstOrDefault()?.Attribute(W14 + "paraId")
            })
            .Where(item => !string.IsNullOrWhiteSpace(item.Name) && !string.IsNullOrWhiteSpace(item.ParagraphId))
            .GroupBy(item => item.Name!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Select(item => item.ParagraphId!).Distinct().ToList(),
                StringComparer.OrdinalIgnoreCase);

        foreach (var region in headingRegions)
        {
            foreach (var paragraph in refreshedBlocks
                         .Skip(region.StartBlock)
                         .Take(region.EndBlock - region.StartBlock + 1)
                         .SelectMany(block => block.DescendantsAndSelf(W + "p")))
            {
                var bookmarkNames = ReferencedBookmarkNames([paragraph]);
                if (bookmarkNames.Count == 0) continue;
                if (bookmarkNames.Count != 1
                    || !bookmarkParagraphIds.TryGetValue(bookmarkNames.Single(), out var paragraphIds)
                    || paragraphIds.Count != 1
                    || !sourceParagraphs.TryGetValue(paragraphIds[0], out var sourceHeading))
                    throw new InvalidOperationException("WPS field refresh produced a TOC entry without a unique source heading.");
                var level = OutlineLevel(sourceHeading, styles);
                if (level <= 0 || !tocStyles.TryGetValue(level, out var tocStyleId))
                    throw new InvalidOperationException("Source template does not define the TOC style required by a refreshed heading.");
                ApplyTemplateTocStyle(paragraph, tocStyleId);
            }
        }
    }

    private static Dictionary<int, string> TocStylesByLevel(XDocument styles)
    {
        var result = new Dictionary<int, string>();
        foreach (var style in styles.Descendants(W + "style")
                     .Where(style => string.Equals((string?)style.Attribute(W + "type"), "paragraph", StringComparison.OrdinalIgnoreCase)))
        {
            var id = (string?)style.Attribute(W + "styleId") ?? string.Empty;
            var name = (string?)style.Element(W + "name")?.Attribute(W + "val") ?? string.Empty;
            var token = id.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) ? id[3..]
                : name.StartsWith("toc ", StringComparison.OrdinalIgnoreCase) ? name[4..] : string.Empty;
            if (int.TryParse(token, out var level) && level > 0 && !result.TryAdd(level, id))
                throw new InvalidOperationException($"Source template defines duplicate TOC level {level} styles.");
        }
        return result;
    }

    private static int OutlineLevel(XElement paragraph, XDocument styles)
    {
        var direct = (int?)paragraph.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val");
        if (direct is not null) return direct.Value + 1;
        var styleId = (string?)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId))
        {
            var style = styles.Descendants(W + "style")
                .FirstOrDefault(candidate => (string?)candidate.Attribute(W + "styleId") == styleId);
            if (style is null) break;
            var outline = (int?)style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val");
            if (outline is not null) return outline.Value + 1;
            styleId = (string?)style.Element(W + "basedOn")?.Attribute(W + "val");
        }
        return 0;
    }

    private static void ApplyTemplateTocStyle(XElement paragraph, string styleId)
    {
        var properties = paragraph.Element(W + "pPr");
        if (properties is null)
        {
            properties = new XElement(W + "pPr");
            paragraph.AddFirst(properties);
        }
        properties.RemoveNodes();
        properties.Add(new XElement(W + "pStyle", new XAttribute(W + "val", styleId)));
        foreach (var run in paragraph.Descendants(W + "r").Where(run => run.Descendants(W + "t").Any()))
            run.Element(W + "rPr")?.Remove();
    }

    private static RunFontPolicy? UniformExplicitBodyFontPolicy(XDocument document)
    {
        var policies = document.Descendants(W + "r")
            .Where(run => !run.Ancestors(W + "tbl").Any())
            .Where(run => run.Descendants(W + "t").Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            .Select(ExplicitFontPolicy)
            .ToList();
        if (policies.Count == 0 || policies.Any(policy => policy is null)) return null;
        var first = policies[0];
        return policies.All(policy => policy == first) ? first : null;
    }

    private static RunFontPolicy? ExplicitFontPolicy(XElement run)
    {
        var properties = run.Element(W + "rPr");
        var fonts = properties?.Element(W + "rFonts");
        var policy = new RunFontPolicy(
            (string?)fonts?.Attribute(W + "ascii"), (string?)fonts?.Attribute(W + "hAnsi"),
            (string?)fonts?.Attribute(W + "eastAsia"), (string?)fonts?.Attribute(W + "cs"),
            (string?)properties?.Element(W + "sz")?.Attribute(W + "val"),
            (string?)properties?.Element(W + "szCs")?.Attribute(W + "val"));
        return policy.Values.All(value => !string.IsNullOrWhiteSpace(value)) ? policy : null;
    }

    private static void ApplyFontPolicy(XElement block, RunFontPolicy policy)
    {
        foreach (var run in block.Descendants(W + "r")
                     .Where(run => run.Descendants(W + "t").Any()))
        {
            var properties = run.Element(W + "rPr");
            if (properties is null) { properties = new XElement(W + "rPr"); run.AddFirst(properties); }
            properties.Elements(W + "rFonts").Remove();
            properties.AddFirst(new XElement(W + "rFonts",
                new XAttribute(W + "ascii", policy.Ascii!), new XAttribute(W + "hAnsi", policy.HighAnsi!),
                new XAttribute(W + "eastAsia", policy.EastAsia!), new XAttribute(W + "cs", policy.ComplexScript!)));
            properties.Elements(W + "sz").Remove();
            properties.Elements(W + "szCs").Remove();
            properties.Add(new XElement(W + "sz", new XAttribute(W + "val", policy.Size!)));
            properties.Add(new XElement(W + "szCs", new XAttribute(W + "val", policy.ComplexSize!)));
        }
    }

    private static void CopyTocBookmarks(
        XDocument source,
        XDocument refreshed,
        IReadOnlyList<IndexRegion> sourceRegions,
        IReadOnlyList<IndexRegion> refreshedRegions)
    {
        var sourceParagraphs = ParagraphsById(source);
        var sourceBody = source.Root!.Element(W + "body")!;
        var sourceBlocks = sourceBody.Elements().ToList();
        var refreshedBody = refreshed.Root!.Element(W + "body")!;
        var refreshedBlocks = refreshedBody.Elements().ToList();
        var sourceIndexBlocks = sourceRegions
            .SelectMany(region => sourceBlocks.Skip(region.StartBlock).Take(region.EndBlock - region.StartBlock + 1))
            .ToHashSet();
        var refreshedIndexBlocks = refreshedRegions
            .SelectMany(region => refreshedBlocks.Skip(region.StartBlock).Take(region.EndBlock - region.StartBlock + 1))
            .ToHashSet();
        var sourceReferencedBookmarkNames = ReferencedBookmarkNames(sourceIndexBlocks);
        var referencedBookmarkNames = ReferencedBookmarkNames(refreshedIndexBlocks);
        var refreshedEnds = refreshed.Descendants(W + "bookmarkEnd")
            .Where(end => !string.IsNullOrWhiteSpace((string?)end.Attribute(W + "id")))
            .GroupBy(end => (string)end.Attribute(W + "id")!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        var matchingStarts = refreshed.Descendants(W + "bookmarkStart")
            .Where(IsTocBookmark)
            .Where(start => referencedBookmarkNames.Contains((string)start.Attribute(W + "name")!))
            .ToList();
        var foundBookmarkNames = matchingStarts
            .Select(start => (string)start.Attribute(W + "name")!)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (!referencedBookmarkNames.SetEquals(foundBookmarkNames))
            throw new InvalidOperationException("WPS field refresh produced an index hyperlink without a unique TOC bookmark.");
        var refreshedStarts = matchingStarts
            .Where(start => !refreshedIndexBlocks.Any(block => start.AncestorsAndSelf().Contains(block)))
            .ToList();

        var oldStarts = source.Descendants(W + "bookmarkStart")
            .Where(IsTocBookmark)
            .Where(start => sourceReferencedBookmarkNames.Contains((string)start.Attribute(W + "name")!))
            .ToList();
        var oldIds = oldStarts.Select(start => (string?)start.Attribute(W + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var element in source.Descendants(W + "bookmarkEnd")
                     .Where(end => oldIds.Contains((string?)end.Attribute(W + "id"))).ToList())
            element.Remove();
        foreach (var element in oldStarts) element.Remove();

        var usedIds = source.Descendants(W + "bookmarkStart")
            .Select(start => (string?)start.Attribute(W + "id"))
            .Where(id => int.TryParse(id, out _))
            .Select(id => int.Parse(id!))
            .ToHashSet();
        var nextId = usedIds.Count == 0 ? 0 : usedIds.Max() + 1;

        foreach (var start in refreshedStarts)
        {
            var oldId = (string?)start.Attribute(W + "id");
            if (string.IsNullOrWhiteSpace(oldId)
                || !refreshedEnds.TryGetValue(oldId, out var matchingEnds)
                || matchingEnds.Count != 1)
                throw new InvalidOperationException("WPS field refresh produced an incomplete TOC bookmark pair.");
            var refreshedStartParagraph = start.Ancestors(W + "p").FirstOrDefault();
            var refreshedEndParagraph = matchingEnds[0].Ancestors(W + "p").FirstOrDefault();
            var startParagraphId = (string?)refreshedStartParagraph?.Attribute(W14 + "paraId");
            var endParagraphId = (string?)refreshedEndParagraph?.Attribute(W14 + "paraId");
            if (string.IsNullOrWhiteSpace(startParagraphId)
                || string.IsNullOrWhiteSpace(endParagraphId)
                || !sourceParagraphs.TryGetValue(startParagraphId, out var sourceStartParagraph)
                || !sourceParagraphs.TryGetValue(endParagraphId, out var sourceEndParagraph))
                throw new InvalidOperationException("WPS field refresh produced a TOC bookmark without unique source paragraph identities.");

            while (usedIds.Contains(nextId)) nextId++;
            var newId = nextId++.ToString();
            usedIds.Add(int.Parse(newId));
            var copiedStart = new XElement(start);
            copiedStart.SetAttributeValue(W + "id", newId);
            var copiedEnd = new XElement(matchingEnds[0]);
            copiedEnd.SetAttributeValue(W + "id", newId);
            var paragraphProperties = sourceStartParagraph.Element(W + "pPr");
            if (paragraphProperties is null) sourceStartParagraph.AddFirst(copiedStart);
            else paragraphProperties.AddAfterSelf(copiedStart);
            sourceEndParagraph.Add(copiedEnd);
        }
    }

    private static HashSet<string> ReferencedBookmarkNames(IEnumerable<XElement> indexBlocks)
        => indexBlocks
            .SelectMany(block => block.Descendants(W + "instrText"))
            .SelectMany(instruction => Regex.Matches(instruction.Value, @"\b_Toc[^\s\""\\]+", RegexOptions.IgnoreCase)
                .Select(match => match.Value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static Dictionary<string, XElement> ParagraphsById(XDocument document)
    {
        var result = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var paragraph in document.Descendants(W + "p"))
        {
            var id = (string?)paragraph.Attribute(W14 + "paraId");
            if (string.IsNullOrWhiteSpace(id)) continue;
            if (!result.TryAdd(id, paragraph))
                throw new InvalidOperationException($"DOCX contains duplicate paragraph identity {id}.");
        }
        return result;
    }

    private static bool IsTocBookmark(XElement element)
        => ((string?)element.Attribute(W + "name"))?.StartsWith("_Toc", StringComparison.OrdinalIgnoreCase) == true;

    private sealed class FieldFrame(int startBlock)
    {
        internal int StartBlock { get; } = startBlock;
        internal System.Text.StringBuilder Instruction { get; } = new();
        internal bool InstructionComplete { get; set; }
    }

    private sealed record IndexRegion(int StartBlock, int EndBlock, string Kind);
    private sealed record RunFontPolicy(string? Ascii, string? HighAnsi, string? EastAsia, string? ComplexScript, string? Size, string? ComplexSize)
    {
        internal IEnumerable<string?> Values => [Ascii, HighAnsi, EastAsia, ComplexScript, Size, ComplexSize];
    }
}
