using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Dockit.Convert;

internal static class DocxFieldResultMerger
{
    private const string DocumentPart = "word/document.xml";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

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

        CopyTocBookmarks(source, refreshed, sourceRegions, refreshedRegions);
        ReplaceIndexRegions(source, refreshed, sourceRegions, refreshedRegions);

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
                source.Save(stream, SaveOptions.DisableFormatting);
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
        IReadOnlyList<IndexRegion> refreshedRegions)
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
            sourceBlocks[sourceRegion.StartBlock].AddBeforeSelf(replacements);
            for (var block = sourceRegion.EndBlock; block >= sourceRegion.StartBlock; block--)
                sourceBlocks[block].Remove();
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
}
