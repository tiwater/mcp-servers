using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Dockit.Convert;

internal static class DocxFieldResultMerger
{
    private const string DocumentPart = "word/document.xml";
    private const string StylesPart = "word/styles.xml";
    private const string SettingsPart = "word/settings.xml";
    private const string AppPropertiesPart = "docProps/app.xml";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly HashSet<string> SettingsLocalNamesAfterUpdateFields = new(StringComparer.Ordinal)
    {
        "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids",
        "mathPr", "uiCompat97To2003", "attachedSchema", "themeFontLang", "clrSchemeMapping",
        "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "compatSetting",
        "decimalSymbol", "listSeparator", "docId", "chartTrackingRefBased", "conflictMode",
        "discardImageEditingData", "defaultImageDpi"
    };

    internal static string PrepareSourceParagraphIdentities(string sourcePath, string outputDirectory, bool excludeCaptionHeadings = false)
    {
        var document = LoadDocument(sourcePath);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var changed = false;
        if (excludeCaptionHeadings)
        {
            var instructions = document.Descendants(W + "p")
                .Select(paragraph => string.Concat(paragraph.Descendants(W + "instrText").Select(node => node.Value)))
                .ToArray();
            var labels = instructions
                .SelectMany(instruction => Regex.Matches(instruction, @"\bTOC\b.*?\\c\s+""([^""]+)""", RegexOptions.IgnoreCase)
                    .Select(match => match.Groups[1].Value))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var paragraph in document.Descendants(W + "p"))
            {
                var instruction = string.Concat(paragraph.Descendants(W + "instrText").Select(node => node.Value));
                var sequence = Regex.Match(instruction, @"\bSEQ\s+(?:""([^""]+)""|([^\s\\]+))", RegexOptions.IgnoreCase);
                var label = sequence.Groups[1].Success ? sequence.Groups[1].Value : sequence.Groups[2].Value;
                if (!sequence.Success || !labels.Contains(label)) continue;
                var properties = paragraph.Element(W + "pPr");
                if (properties is null) { properties = new XElement(W + "pPr"); paragraph.AddFirst(properties); }
                properties.Elements(W + "outlineLvl").Remove();
                var outline = new XElement(W + "outlineLvl", new XAttribute(W + "val", 9));
                var later = properties.Elements().FirstOrDefault(element => element.Name == W + "divId"
                    || element.Name == W + "cnfStyle" || element.Name == W + "rPr"
                    || element.Name == W + "sectPr" || element.Name == W + "pPrChange");
                if (later is null) properties.Add(outline); else later.AddBeforeSelf(outline);
                changed = true;
            }
        }
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

    internal static void Merge(string sourcePath, string refreshedPath, string outputPath, int? actualPageCount = null)
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
        NormalizeTocResultStyles(sourcePath, source, refreshed, sourceRegions, refreshedRegions);
        CopyTocBookmarks(source, refreshed, sourceRegions, refreshedRegions);
        ReplaceIndexRegions(source, refreshed, sourceRegions, refreshedRegions, bodyFontPolicy);

        var updatedParts = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase)
        {
            [DocumentPart] = source,
        };
        MergeHeaderFooterFieldsAndSettings(sourcePath, refreshedPath, actualPageCount, updatedParts);
        WriteParts(sourcePath, updatedParts, outputPath);
    }

    private static void WriteDocumentPart(string sourcePath, XDocument document, string outputPath)
        => WriteParts(sourcePath, new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase) { [DocumentPart] = document }, outputPath);

    private static void WriteParts(string sourcePath, IReadOnlyDictionary<string, XDocument> updatedParts, string outputPath)
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
                foreach (var (partName, document) in updatedParts)
                {
                    var existing = archive.GetEntry(partName)
                        ?? throw new InvalidOperationException($"DOCX package is missing {partName}.");
                    var actualName = existing.FullName;
                    existing.Delete();
                    var replacement = archive.CreateEntry(actualName, CompressionLevel.Optimal);
                    using var stream = replacement.Open();
                    document.Save(stream, SaveOptions.DisableFormatting);
                }
            }
            File.Move(temporaryOutput, outputPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryOutput)) File.Delete(temporaryOutput);
        }
    }

    private static void MergeHeaderFooterFieldsAndSettings(
        string sourcePath,
        string refreshedPath,
        int? actualPageCount,
        Dictionary<string, XDocument> updatedParts)
    {
        using var sourceArchive = ZipFile.OpenRead(sourcePath);
        using var refreshedArchive = ZipFile.OpenRead(refreshedPath);

        var refreshedAppPages = TryReadAppPageCount(refreshedArchive);
        int? observedRefreshedNumPages = null;

        foreach (var sourceEntry in sourceArchive.Entries
                     .Where(entry => IsHeaderOrFooterPart(entry.FullName))
                     .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase))
        {
            var sourcePart = LoadEntry(sourceEntry);
            var sourceFields = FindStoryPageFields(sourcePart);
            if (sourceFields.Count == 0) continue;

            var refreshedEntry = refreshedArchive.GetEntry(sourceEntry.FullName);
            var refreshedFields = refreshedEntry is null
                ? []
                : FindStoryPageFields(LoadEntry(refreshedEntry));
            var refreshedNumPages = refreshedFields.Where(field => field.Kind == "NUMPAGES").ToList();
            var refreshedPages = refreshedFields.Where(field => field.Kind == "PAGE").ToList();

            var numPagesIndex = 0;
            var pageIndex = 0;
            var partChanged = false;

            foreach (var sourceField in sourceFields)
            {
                if (sourceField.Kind == "NUMPAGES")
                {
                    var refreshedField = numPagesIndex < refreshedNumPages.Count
                        ? refreshedNumPages[numPagesIndex]
                        : refreshedNumPages.FirstOrDefault();
                    numPagesIndex++;

                    string? targetValue = null;
                    if (actualPageCount is > 0)
                    {
                        targetValue = actualPageCount.Value.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        var sourceText = ReadStoryFieldCachedText(sourceField).Trim();
                        var refreshedText = refreshedField is null ? null : ReadStoryFieldCachedText(refreshedField).Trim();
                        if (!string.IsNullOrWhiteSpace(refreshedText)
                            && (!StringComparer.Ordinal.Equals(refreshedText, sourceText) || refreshedAppPages is null))
                        {
                            targetValue = refreshedText;
                        }
                        else if (refreshedAppPages is > 0)
                        {
                            targetValue = refreshedAppPages.Value.ToString(CultureInfo.InvariantCulture);
                        }
                        else if (!string.IsNullOrWhiteSpace(refreshedText))
                        {
                            targetValue = refreshedText;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(targetValue))
                    {
                        if (int.TryParse(targetValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                            && parsed > 0)
                            observedRefreshedNumPages ??= parsed;
                        if (SetStoryFieldCachedValue(sourceField, targetValue))
                            partChanged = true;
                    }
                }
                else if (sourceField.Kind == "PAGE")
                {
                    var refreshedField = pageIndex < refreshedPages.Count
                        ? refreshedPages[pageIndex]
                        : null;
                    pageIndex++;
                    var refreshedText = refreshedField is null ? null : ReadStoryFieldCachedText(refreshedField).Trim();
                    if (!string.IsNullOrWhiteSpace(refreshedText)
                        && SetStoryFieldCachedValue(sourceField, refreshedText))
                        partChanged = true;
                }
            }

            if (partChanged)
                updatedParts[sourceEntry.FullName] = sourcePart;
        }

        var resolvedPageCount = actualPageCount ?? refreshedAppPages ?? observedRefreshedNumPages;
        if (resolvedPageCount is > 0 && sourceArchive.GetEntry(AppPropertiesPart) is { } appEntry)
        {
            var appDoc = LoadEntry(appEntry);
            var pagesElement = appDoc.Descendants().FirstOrDefault(element => element.Name.LocalName == "Pages");
            var formatted = resolvedPageCount.Value.ToString(CultureInfo.InvariantCulture);
            if (pagesElement is not null && !StringComparer.Ordinal.Equals(pagesElement.Value.Trim(), formatted))
            {
                pagesElement.Value = formatted;
                updatedParts[appEntry.FullName] = appDoc;
            }
        }

        if (sourceArchive.GetEntry(SettingsPart) is { } settingsEntry)
        {
            var settingsDoc = LoadEntry(settingsEntry);
            if (EnsureUpdateFieldsTrue(settingsDoc))
                updatedParts[settingsEntry.FullName] = settingsDoc;
        }
    }

    private static bool IsHeaderOrFooterPart(string fullName)
        => Regex.IsMatch(fullName, @"^word/(?:header|footer)\d*\.xml$", RegexOptions.IgnoreCase);

    private static XDocument LoadEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static int? TryReadAppPageCount(ZipArchive archive)
    {
        var entry = archive.GetEntry(AppPropertiesPart);
        if (entry is null) return null;
        try
        {
            var doc = LoadEntry(entry);
            var pagesValue = doc.Descendants().FirstOrDefault(element => element.Name.LocalName == "Pages")?.Value?.Trim();
            if (int.TryParse(pagesValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0)
                return parsed;
        }
        catch
        {
            // Ignore malformed extended properties and rely on header/footer fields or WPS statistics.
        }
        return null;
    }

    private static bool EnsureUpdateFieldsTrue(XDocument settingsDoc)
    {
        var root = settingsDoc.Root;
        if (root is null || root.Name != W + "settings") return false;
        var existing = root.Element(W + "updateFields");
        if (existing is not null)
        {
            var current = (string?)existing.Attribute(W + "val");
            if (StringComparer.OrdinalIgnoreCase.Equals(current, "true")) return false;
            existing.SetAttributeValue(W + "val", "true");
            return true;
        }

        var updateFields = new XElement(W + "updateFields", new XAttribute(W + "val", "true"));
        var anchor = root.Elements().FirstOrDefault(element => SettingsLocalNamesAfterUpdateFields.Contains(element.Name.LocalName));
        if (anchor is null) root.Add(updateFields);
        else anchor.AddBeforeSelf(updateFields);
        return true;
    }

    private static List<StoryPageField> FindStoryPageFields(XDocument document)
    {
        var root = document.Root;
        if (root is null) return [];

        var fields = new List<StoryPageField>();
        var stack = new Stack<StoryFieldFrame>();

        foreach (var element in root.Descendants())
        {
            if (element.Name == W + "fldSimple")
            {
                var kind = HeaderFooterFieldKind((string?)element.Attribute(W + "instr") ?? string.Empty);
                if (kind is not null)
                    fields.Add(new StoryPageField(kind, root, element, null, null, null));
            }
            else if (element.Name == W + "fldChar")
            {
                var type = (string?)element.Attribute(W + "fldCharType");
                if (type == "begin")
                {
                    stack.Push(new StoryFieldFrame(element));
                }
                else if (type == "separate")
                {
                    if (stack.Count > 0)
                        stack.Peek().Separate ??= element;
                }
                else if (type == "end")
                {
                    if (stack.Count == 0) continue;
                    var frame = stack.Pop();
                    var kind = HeaderFooterFieldKind(frame.Instruction.ToString());
                    if (kind is not null)
                        fields.Add(new StoryPageField(kind, root, null, frame.Begin, frame.Separate, element));
                }
            }
            else if (element.Name == W + "instrText" && stack.Count > 0 && stack.Peek().Separate is null)
            {
                stack.Peek().Instruction.Append(element.Value);
            }
        }

        return fields;
    }

    private static string? HeaderFooterFieldKind(string instruction)
    {
        var match = Regex.Match(instruction.Trim(), @"^([A-Za-z]+)\b");
        if (!match.Success) return null;
        var keyword = match.Groups[1].Value.ToUpperInvariant();
        return keyword is "PAGE" or "NUMPAGES" ? keyword : null;
    }

    private static List<XElement> StoryFieldResultTexts(StoryPageField field)
    {
        if (field.SimpleField is not null)
            return field.SimpleField.Descendants(W + "t").ToList();
        if (field.Separate is null || field.End is null)
            return [];
        var order = XNode.DocumentOrderComparer;
        var scope = field.Separate.Ancestors(W + "p").FirstOrDefault();
        if (scope is null || !ReferenceEquals(scope, field.End.Ancestors(W + "p").FirstOrDefault()))
            scope = field.Root;
        return scope.Descendants(W + "t")
            .Where(text => order.Compare(text, field.Separate) > 0 && order.Compare(text, field.End) < 0)
            .ToList();
    }

    private static string ReadStoryFieldCachedText(StoryPageField field)
        => string.Concat(StoryFieldResultTexts(field).Select(element => element.Value));

    private static bool SetStoryFieldCachedValue(StoryPageField field, string newValue)
    {
        if (field.SimpleField is not null)
        {
            var texts = field.SimpleField.Descendants(W + "t").ToList();
            if (texts.Count > 0)
            {
                if (texts.Count == 1 && StringComparer.Ordinal.Equals(texts[0].Value, newValue))
                    return false;
                texts[0].Value = newValue;
                RemoveExtraTextNodes(texts.Skip(1));
                return true;
            }
            field.SimpleField.Add(new XElement(W + "r", new XElement(W + "t", newValue)));
            return true;
        }

        if (field.Begin is null || field.End is null) return false;
        if (field.Separate is not null)
        {
            var texts = StoryFieldResultTexts(field);
            if (texts.Count > 0)
            {
                if (texts.Count == 1 && StringComparer.Ordinal.Equals(texts[0].Value, newValue))
                    return false;
                texts[0].Value = newValue;
                RemoveExtraTextNodes(texts.Skip(1));
                return true;
            }

            var separateRun = field.Separate.Ancestors(W + "r").FirstOrDefault();
            var endRun = field.End.Ancestors(W + "r").FirstOrDefault();
            if (ReferenceEquals(separateRun, endRun) && separateRun is not null)
            {
                field.Separate.AddAfterSelf(new XElement(W + "t", newValue));
                return true;
            }
            if (separateRun is not null)
            {
                var valueRun = new XElement(W + "r");
                if (separateRun.Element(W + "rPr") is { } rPr)
                    valueRun.Add(new XElement(rPr));
                valueRun.Add(new XElement(W + "t", newValue));
                separateRun.AddAfterSelf(valueRun);
                return true;
            }
            return false;
        }

        var beginRun = field.Begin.Ancestors(W + "r").FirstOrDefault();
        var endRunNoSep = field.End.Ancestors(W + "r").FirstOrDefault();
        if (endRunNoSep is null) return false;
        if (ReferenceEquals(beginRun, endRunNoSep))
        {
            field.End.AddBeforeSelf(
                new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate")),
                new XElement(W + "t", newValue));
            return true;
        }
        var sepRun = new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate")));
        var insertedRun = new XElement(W + "r");
        if (beginRun?.Element(W + "rPr") is { } beginRPr)
            insertedRun.Add(new XElement(beginRPr));
        insertedRun.Add(new XElement(W + "t", newValue));
        endRunNoSep.AddBeforeSelf(sepRun, insertedRun);
        return true;
    }

    private static void RemoveExtraTextNodes(IEnumerable<XElement> extras)
    {
        foreach (var extra in extras.ToList())
        {
            var parentRun = extra.Parent;
            extra.Remove();
            if (parentRun?.Name == W + "r" && !parentRun.Elements().Any(child => child.Name != W + "rPr"))
                parentRun.Remove();
        }
    }

    private sealed class StoryFieldFrame(XElement begin)
    {
        internal XElement Begin { get; } = begin;
        internal XElement? Separate { get; set; }
        internal System.Text.StringBuilder Instruction { get; } = new();
    }

    private sealed record StoryPageField(
        string Kind,
        XElement Root,
        XElement? SimpleField,
        XElement? Begin,
        XElement? Separate,
        XElement? End);

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
                        stack.Push(new FieldFrame(blockIndex, element));
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
                            regions.Add(new IndexRegion(field.StartBlock, blockIndex, kind, field.Start, element));
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
            if (bodyFontPolicy is not null)
                foreach (var refreshedBlock in refreshedBlocks
                             .Skip(refreshedRegion.StartBlock)
                             .Take(refreshedRegion.EndBlock - refreshedRegion.StartBlock + 1))
                    ApplyFontPolicy(refreshedBlock, bodyFontPolicy);
            var replacements = InlineRegionReplacement(sourceBlocks, refreshedBlocks, sourceRegion, refreshedRegion);
            sourceBlocks[sourceRegion.StartBlock].AddBeforeSelf(replacements);
            for (var block = sourceRegion.EndBlock; block >= sourceRegion.StartBlock; block--)
                sourceBlocks[block].Remove();
        }
    }

    private static List<XElement> InlineRegionReplacement(
        IReadOnlyList<XElement> sourceBlocks,
        IReadOnlyList<XElement> refreshedBlocks,
        IndexRegion sourceRegion,
        IndexRegion refreshedRegion)
    {
        var sourceStart = BoundaryParagraph(sourceBlocks, sourceRegion.StartBlock, sourceRegion.Start);
        var sourceEnd = BoundaryParagraph(sourceBlocks, sourceRegion.EndBlock, sourceRegion.End);
        var refreshedStart = BoundaryParagraph(refreshedBlocks, refreshedRegion.StartBlock, refreshedRegion.Start);
        var refreshedEnd = BoundaryParagraph(refreshedBlocks, refreshedRegion.EndBlock, refreshedRegion.End);

        var sourcePrefix = SliceBefore(sourceStart, sourceRegion.Start);
        var sourceSuffix = SliceAfter(sourceEnd, sourceRegion.End);
        if (ReferenceEquals(refreshedStart, refreshedEnd))
        {
            if (!ReferenceEquals(sourceStart, sourceEnd))
                throw new InvalidOperationException("WPS field refresh collapsed a multi-paragraph DOCX index field boundary.");
            AppendParagraphContent(sourcePrefix, SliceRange(refreshedStart, refreshedRegion.Start, refreshedRegion.End));
            AppendParagraphContent(sourcePrefix, sourceSuffix);
            return [sourcePrefix];
        }

        var replacements = refreshedBlocks
            .Skip(refreshedRegion.StartBlock)
            .Take(refreshedRegion.EndBlock - refreshedRegion.StartBlock + 1)
            .Select(element => new XElement(element))
            .ToList();
        replacements[0] = sourcePrefix;
        AppendParagraphContent(replacements[0], SliceFrom(refreshedStart, refreshedRegion.Start));

        var last = sourceSuffix;
        PrependParagraphContent(last, SliceThrough(refreshedEnd, refreshedRegion.End));
        replacements[^1] = last;
        return replacements;
    }

    private static XElement BoundaryParagraph(
        IReadOnlyList<XElement> blocks,
        int blockIndex,
        XElement boundary)
    {
        var paragraph = boundary.Ancestors(W + "p").FirstOrDefault()
            ?? throw new InvalidOperationException("DOCX index field boundary is not inside a paragraph.");
        if (!ReferenceEquals(paragraph, blocks[blockIndex]))
            throw new InvalidOperationException("DOCX index field boundary paragraph is not a body-level block.");
        return paragraph;
    }

    private static void AppendParagraphContent(XElement target, XElement fragment)
        => target.Add(fragment.Elements().Where(element => element.Name != W + "pPr").Select(element => new XElement(element)));

    private static void PrependParagraphContent(XElement target, XElement fragment)
    {
        var content = fragment.Elements().Where(element => element.Name != W + "pPr")
            .Select(element => new XElement(element)).ToList();
        var properties = target.Element(W + "pPr");
        if (properties is null) target.AddFirst(content); else properties.AddAfterSelf(content);
    }

    private static XElement SliceBefore(XElement paragraph, XElement boundary)
    {
        var reached = false;
        return Slice(paragraph, element =>
        {
            if (ReferenceEquals(element, boundary)) { reached = true; return false; }
            return !reached;
        }, boundary) ?? new XElement(paragraph.Name, paragraph.Attributes());
    }

    private static XElement SliceAfter(XElement paragraph, XElement boundary)
    {
        var reached = false;
        var result = Slice(paragraph, element =>
        {
            if (ReferenceEquals(element, boundary)) { reached = true; return false; }
            return reached;
        }, boundary) ?? new XElement(paragraph.Name, paragraph.Attributes());
        var properties = paragraph.Element(W + "pPr");
        if (properties is not null && result.Element(W + "pPr") is null)
            result.AddFirst(new XElement(properties));
        return result;
    }

    private static XElement SliceFrom(XElement paragraph, XElement boundary)
    {
        var reached = false;
        return Slice(paragraph, element =>
        {
            if (ReferenceEquals(element, boundary)) reached = true;
            return reached;
        }, boundary) ?? throw new InvalidOperationException("Refreshed DOCX index field start boundary could not be sliced.");
    }

    private static XElement SliceThrough(XElement paragraph, XElement boundary)
    {
        var reached = false;
        return Slice(paragraph, element =>
        {
            if (reached) return false;
            if (ReferenceEquals(element, boundary)) reached = true;
            return true;
        }, boundary) ?? throw new InvalidOperationException("Refreshed DOCX index field end boundary could not be sliced.");
    }

    private static XElement SliceRange(XElement paragraph, XElement start, XElement end)
    {
        var inside = false;
        return Slice(paragraph, element =>
        {
            if (ReferenceEquals(element, start)) inside = true;
            var include = inside;
            if (ReferenceEquals(element, end)) inside = false;
            return include;
        }, start, end) ?? throw new InvalidOperationException("Refreshed DOCX index field boundary could not be sliced.");
    }

    private static XElement? Slice(XElement element, Func<XElement, bool> include, params XElement[] boundaries)
    {
        if (boundaries.Contains(element)) return include(element) ? new XElement(element) : null;
        if (!boundaries.Any(boundary => element.Descendants().Contains(boundary)))
            return include(element) ? new XElement(element) : null;

        var result = new XElement(element.Name, element.Attributes());
        foreach (var child in element.Elements())
        {
            var sliced = Slice(child, include, boundaries);
            if (sliced is not null) result.Add(sliced);
        }
        var propertiesName = element.Name == W + "p" ? W + "pPr"
            : element.Name == W + "r" ? W + "rPr" : null;
        if (!result.Elements().Any(child => propertiesName is null || child.Name != propertiesName))
            return element.Name == W + "p" ? result : null;
        var properties = propertiesName is null ? null : element.Element(propertiesName);
        if (properties is not null && result.Element(propertiesName!) is null)
            result.AddFirst(new XElement(properties));
        return result;
    }

    private static void NormalizeTocResultStyles(
        string sourcePath,
        XDocument source,
        XDocument refreshed,
        IReadOnlyList<IndexRegion> sourceRegions,
        IReadOnlyList<IndexRegion> refreshedRegions)
    {
        var headingRegions = refreshedRegions.Where(region => region.Kind == "table-of-contents").ToList();
        if (headingRegions.Count == 0) return;
        var styles = LoadPart(sourcePath, StylesPart);
        var tocStyles = SourceTocParagraphsByLevel(source, sourceRegions, styles);
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
                    throw new InvalidOperationException(
                        $"WPS field refresh produced a TOC entry without a unique source heading: " +
                        $"references=[{string.Join(",", bookmarkNames)}], " +
                        $"paragraphIds=[{(bookmarkNames.Count == 1 && bookmarkParagraphIds.TryGetValue(bookmarkNames.Single(), out var ids) ? string.Join(",", ids) : string.Empty)}].");
                var level = OutlineLevel(sourceHeading, styles);
                if (level <= 0)
                    throw new InvalidOperationException("Source template TOC heading has no outline level.");
                if (tocStyles.TryGetValue(level, out var tocParagraphProperties))
                    ApplyTemplateTocStyle(paragraph, tocParagraphProperties);
                else
                    RemoveRefreshedTocTextFormatting(paragraph);
            }
        }
    }

    private static Dictionary<int, XElement> SourceTocParagraphsByLevel(
        XDocument source,
        IReadOnlyList<IndexRegion> sourceRegions,
        XDocument styles)
    {
        var body = source.Root!.Element(W + "body")!;
        var blocks = body.Elements().ToList();
        var targets = source.Descendants(W + "bookmarkStart")
            .Where(IsTocBookmark)
            .Where(start => !sourceRegions.Any(region => IsWithinIndexRegion(start, blocks, region)))
            .GroupBy(start => (string)start.Attribute(W + "name")!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<int, XElement>();
        foreach (var region in sourceRegions.Where(region => region.Kind == "table-of-contents"))
        {
            foreach (var paragraph in blocks
                         .Skip(region.StartBlock)
                         .Take(region.EndBlock - region.StartBlock + 1)
                         .SelectMany(block => block.DescendantsAndSelf(W + "p")))
            {
                var bookmarkNames = ReferencedBookmarkNames([paragraph]);
                if (bookmarkNames.Count == 0) continue;
                if (bookmarkNames.Count != 1
                    || !targets.TryGetValue(bookmarkNames.Single(), out var starts)
                    || starts.Count != 1)
                    throw new InvalidOperationException("Source template TOC entry does not identify one heading.");
                var target = starts[0].Ancestors(W + "p").FirstOrDefault()
                    ?? throw new InvalidOperationException("Source template TOC bookmark has no heading paragraph.");
                var level = OutlineLevel(target, styles);
                var properties = paragraph.Element(W + "pPr")
                    ?? throw new InvalidOperationException("Source template TOC entry has no paragraph properties.");
                if (level <= 0) throw new InvalidOperationException("Source template TOC heading has no outline level.");
                if (result.TryGetValue(level, out var existing))
                {
                    if (!XNode.DeepEquals(existing, properties))
                    {
                        var existingStyle = (string?)existing.Element(W + "pStyle")?.Attribute(W + "val");
                        var currentStyle = (string?)properties.Element(W + "pStyle")?.Attribute(W + "val");
                        if (string.IsNullOrWhiteSpace(existingStyle)
                            || !string.Equals(existingStyle, currentStyle, StringComparison.Ordinal))
                            throw new InvalidOperationException($"Source template defines inconsistent TOC level {level} formatting.");
                    }
                }
                else result.Add(level, new XElement(properties));
            }
        }
        return result;
    }

    private static bool IsWithinIndexRegion(
        XElement element,
        IReadOnlyList<XElement> blocks,
        IndexRegion region)
    {
        var blockIndex = -1;
        for (var index = region.StartBlock; index <= region.EndBlock; index++)
        {
            if (!element.AncestorsAndSelf().Contains(blocks[index])) continue;
            blockIndex = index;
            break;
        }
        if (blockIndex < 0) return false;

        var order = XNode.DocumentOrderComparer;
        if (blockIndex == region.StartBlock && order.Compare(element, region.Start) < 0) return false;
        if (blockIndex == region.EndBlock && order.Compare(element, region.End) > 0) return false;
        return true;
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

    private static void ApplyTemplateTocStyle(XElement paragraph, XElement templateProperties)
    {
        var properties = paragraph.Element(W + "pPr");
        if (properties is null) paragraph.AddFirst(new XElement(templateProperties));
        else properties.ReplaceWith(new XElement(templateProperties));
        RemoveRefreshedTocTextFormatting(paragraph);
    }

    private static void RemoveRefreshedTocTextFormatting(XElement paragraph)
    {
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
            .Where(start => !refreshedRegions.Any(region => IsWithinIndexRegion(start, refreshedBlocks, region)))
            .ToList();

        var oldStarts = source.Descendants(W + "bookmarkStart")
            .Where(IsTocBookmark)
            .Where(start => sourceReferencedBookmarkNames.Contains((string)start.Attribute(W + "name")!))
            .Where(start => !sourceRegions.Any(region => IsWithinIndexRegion(start, sourceBlocks, region)))
            .ToList();
        var oldIds = oldStarts.Select(start => (string?)start.Attribute(W + "id"))
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var element in source.Descendants(W + "bookmarkEnd")
                     .Where(end => oldIds.Contains((string?)end.Attribute(W + "id"))).ToList())
            element.Remove();
        foreach (var element in oldStarts) element.Remove();

        var usedIds = source.Descendants()
            .Where(element => element.Name == W + "bookmarkStart" || element.Name == W + "bookmarkEnd")
            .Select(start => (string?)start.Attribute(W + "id"))
            .Where(id => int.TryParse(id, out _))
            .Select(id => int.Parse(id!))
            .ToHashSet();
        var nextId = usedIds.Count == 0 ? 0 : usedIds.Max() + 1;

        foreach (var start in refreshedStarts)
        {
            var oldId = (string?)start.Attribute(W + "id");
            refreshedEnds.TryGetValue(oldId ?? string.Empty, out var matchingEnds);
            var targetEnds = matchingEnds?
                .Where(end => !refreshedRegions.Any(region => IsWithinIndexRegion(end, refreshedBlocks, region)))
                .ToList();
            var refreshedStartParagraph = start.Ancestors(W + "p").FirstOrDefault();
            var sameParagraphEnds = targetEnds?
                .Where(end => ReferenceEquals(end.Ancestors(W + "p").FirstOrDefault(), refreshedStartParagraph))
                .Where(end => XNode.DocumentOrderComparer.Compare(start, end) < 0)
                .ToList();
            var pairedEnd = sameParagraphEnds?.Count == 1
                ? sameParagraphEnds[0]
                : targetEnds?.Count == 1 ? targetEnds[0] : null;
            if (string.IsNullOrWhiteSpace(oldId) || pairedEnd is null)
                throw new InvalidOperationException(
                    $"WPS field refresh produced an incomplete TOC bookmark pair for " +
                    $"{(string?)start.Attribute(W + "name") ?? "<unnamed>"} " +
                    $"(id={oldId ?? "<missing>"}, targetEnds={targetEnds?.Count ?? 0}, allEnds={matchingEnds?.Count ?? 0}).");
            var refreshedEndParagraph = pairedEnd.Ancestors(W + "p").FirstOrDefault();
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
            var copiedEnd = new XElement(pairedEnd);
            copiedEnd.SetAttributeValue(W + "id", newId);
            var inlineEndRegionIndex = Enumerable.Range(0, refreshedRegions.Count)
                .FirstOrDefault(index =>
                    ReferenceEquals(refreshedStartParagraph, refreshedBlocks[refreshedRegions[index].EndBlock])
                    && XNode.DocumentOrderComparer.Compare(start, refreshedRegions[index].End) > 0,
                    -1);
            if (inlineEndRegionIndex >= 0)
            {
                var sourceBoundary = sourceRegions[inlineEndRegionIndex].End.AncestorsAndSelf()
                    .First(element => ReferenceEquals(element.Parent, sourceStartParagraph));
                sourceBoundary.AddAfterSelf(copiedStart);
            }
            else
            {
                var paragraphProperties = sourceStartParagraph.Element(W + "pPr");
                if (paragraphProperties is null) sourceStartParagraph.AddFirst(copiedStart);
                else paragraphProperties.AddAfterSelf(copiedStart);
            }
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

    private sealed class FieldFrame(int startBlock, XElement start)
    {
        internal int StartBlock { get; } = startBlock;
        internal XElement Start { get; } = start;
        internal System.Text.StringBuilder Instruction { get; } = new();
        internal bool InstructionComplete { get; set; }
    }

    private sealed record IndexRegion(int StartBlock, int EndBlock, string Kind, XElement Start, XElement End);
    private sealed record RunFontPolicy(string? Ascii, string? HighAnsi, string? EastAsia, string? ComplexScript, string? Size, string? ComplexSize)
    {
        internal IEnumerable<string?> Values => [Ascii, HighAnsi, EastAsia, ComplexScript, Size, ComplexSize];
    }
}
