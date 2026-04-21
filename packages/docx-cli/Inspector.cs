using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class Inspector
{
    private static readonly Regex PlaceholderPattern = new(@"\{\{[^{}]+\}\}|<<[^<>]+>>", RegexOptions.Compiled);

    public static InspectionReport Inspect(string input)
    {
        var path = Path.GetFullPath(input);
        using var doc = WordprocessingDocument.Open(path, false);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body not found.");

        var allRoots = GetRoots(doc).ToList();
        var allParagraphs = allRoots.SelectMany(root => root.Descendants<Paragraph>()).ToList();
        var allTables = allRoots.SelectMany(root => root.Descendants<Table>()).ToList();
        var bodyParagraphs = body.Descendants<Paragraph>().ToList();
        var bodyParagraphTexts = bodyParagraphs.Select(GetParagraphText).ToList();
        var bodyTables = body.Elements<Table>().ToList();
        var tableMetadata = BuildTableMetadata(bodyTables);
        var allTexts = allParagraphs.Select(GetParagraphText).Where(text => !string.IsNullOrWhiteSpace(text)).ToList();

        var paragraphStyles = new Dictionary<string, int>(StringComparer.Ordinal);
        var runStyles = new Dictionary<string, int>(StringComparer.Ordinal);
        var headings = new List<HeadingInfo>();

        foreach (var paragraph in allParagraphs)
        {
            var text = GetParagraphText(paragraph);
            var pStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(pStyle))
            {
                paragraphStyles[pStyle] = paragraphStyles.GetValueOrDefault(pStyle) + 1;
                if (LooksLikeHeading(paragraph, pStyle) && !string.IsNullOrWhiteSpace(text))
                {
                    headings.Add(new HeadingInfo(pStyle, Clip(text, 160), GetParagraphSource(paragraph)));
                }
            }

            foreach (var runStyle in paragraph.Descendants<RunStyle>())
            {
                var value = runStyle.Val?.Value;
                if (!string.IsNullOrWhiteSpace(value))
                {
                    runStyles[value] = runStyles.GetValueOrDefault(value) + 1;
                }
            }
        }

        var styleDefinitions = mainPart.StyleDefinitionsPart?.Styles?.Elements<Style>().ToList() ?? [];
        var placeholders = PlaceholderPattern
            .Matches(string.Join("\n", allTexts))
            .Select(match => match.Value)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .Take(50)
            .ToList();

        var trackedChanges = allRoots.Sum(root =>
            root.Descendants<InsertedRun>().Count()
            + root.Descendants<DeletedRun>().Count()
            + root.Descendants<MoveFromRun>().Count()
            + root.Descendants<MoveToRun>().Count()
            + root.Descendants<Inserted>().Count()
            + root.Descendants<Deleted>().Count());

        var annotationAnchors = BuildAnnotationAnchors(body, mainPart, bodyParagraphs, bodyParagraphTexts, bodyTables, tableMetadata);

        return new InspectionReport(
            File: path,
            Package: BuildPackageSummary(path),
            Content: new ContentSummary(
                ParagraphCount: allParagraphs.Count,
                TableCount: allTables.Count,
                SectionCount: body.Descendants<SectionProperties>().Count(),
                HeaderPartCount: mainPart.HeaderParts.Count(),
                FooterPartCount: mainPart.FooterParts.Count(),
                Headings: headings.Take(50).ToList(),
                Placeholders: placeholders),
            Styles: new StyleSummary(
                DefinedParagraphStyleCount: styleDefinitions.Count(s => s.Type?.Value == StyleValues.Paragraph),
                DefinedCharacterStyleCount: styleDefinitions.Count(s => s.Type?.Value == StyleValues.Character),
                DefinedTableStyleCount: styleDefinitions.Count(s => s.Type?.Value == StyleValues.Table),
                ParagraphStylesInUse: paragraphStyles.OrderByDescending(kv => kv.Value).ThenBy(kv => kv.Key, StringComparer.Ordinal).Take(50).Select(kv => new StyleCount(kv.Key, kv.Value)).ToList(),
                RunStylesInUse: runStyles.OrderByDescending(kv => kv.Value).ThenBy(kv => kv.Key, StringComparer.Ordinal).Take(50).Select(kv => new StyleCount(kv.Key, kv.Value)).ToList()),
            Annotations: new AnnotationSummary(
                CommentCount: mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Count() ?? 0,
                FootnoteCount: mainPart.FootnotesPart?.Footnotes?.Elements<Footnote>().Count() ?? 0,
                EndnoteCount: mainPart.EndnotesPart?.Endnotes?.Elements<Endnote>().Count() ?? 0,
                TrackedChangeElements: trackedChanges),
            Structure: new StructureSummary(
                BookmarkCount: allRoots.Sum(root => root.Descendants<BookmarkStart>().Count()),
                HyperlinkCount: allRoots.Sum(root => root.Descendants<Hyperlink>().Count()),
                FieldCount: allRoots.Sum(root => root.Descendants<SimpleField>().Count() + root.Descendants<FieldCode>().Count()),
                ContentControlCount: allRoots.Sum(root => root.Descendants<SdtElement>().Count()),
                DrawingCount: allRoots.Sum(root => root.Descendants<Drawing>().Count()),
                Tables: tableMetadata,
                AnnotationAnchors: annotationAnchors),
            Formatting: new FormattingSummary(
                ParagraphsWithDirectFormatting: allParagraphs.Count(HasParagraphDirectFormatting),
                RunsWithDirectFormatting: allRoots.SelectMany(root => root.Descendants<Run>()).Count(HasRunDirectFormatting)));
    }

    public static IReadOnlyList<AnnotationAnchor> BuildAnnotationAnchors(
        Body body,
        MainDocumentPart mainPart,
        IReadOnlyList<Paragraph> bodyParagraphs,
        IReadOnlyList<string> bodyParagraphTexts,
        IReadOnlyList<Table> bodyTables,
        IReadOnlyList<TableMetadata> tableMetadata)
    {
        var comments = mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>()?.ToDictionary(
            comment => comment.Id?.Value ?? string.Empty,
            comment => comment,
            StringComparer.Ordinal) ?? new Dictionary<string, Comment>(StringComparer.Ordinal);

        var anchors = new List<AnnotationAnchor>();

        for (var paragraphIndex = 0; paragraphIndex < bodyParagraphs.Count; paragraphIndex++)
        {
            var paragraph = bodyParagraphs[paragraphIndex];
            var paragraphText = bodyParagraphTexts[paragraphIndex];
            var previousParagraphText = paragraphIndex > 0 ? bodyParagraphTexts[paragraphIndex - 1] : null;
            var followingParagraphText = paragraphIndex + 1 < bodyParagraphTexts.Count ? bodyParagraphTexts[paragraphIndex + 1] : null;
            var nearestHeadingText = GetNearestHeadingText(bodyParagraphs, bodyParagraphTexts, paragraphIndex);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var start in paragraph.Descendants<CommentRangeStart>())
            {
                var commentId = start.Id?.Value;
                if (string.IsNullOrWhiteSpace(commentId) || !seen.Add(commentId))
                {
                    continue;
                }

                comments.TryGetValue(commentId, out var comment);
                var anchorText = GetParagraphText(paragraph);
                var cell = paragraph.Ancestors<TableCell>().FirstOrDefault();
                var row = cell?.Parent as TableRow;
                var table = cell?.Ancestors<Table>().FirstOrDefault();
                var targetKind = cell is null ? "paragraph" : "tableCell";
                var tableIndex = table is null ? null : GetIndexWithinParent(bodyTables, table);
                var tableInfo = tableIndex is null || tableIndex < 0 || tableIndex >= tableMetadata.Count ? null : tableMetadata[tableIndex.Value];

                anchors.Add(new AnnotationAnchor(
                    CommentId: commentId,
                    Author: comment?.Author?.Value,
                    CommentText: GetCommentText(comment),
                    AnchorText: Clip(anchorText, 240),
                    Source: GetPartSource(paragraph),
                    TargetKind: targetKind,
                    ParagraphIndex: paragraphIndex,
                    TableIndex: tableIndex,
                    RowIndex: row is null ? null : GetIndexWithinParent(table?.Elements<TableRow>().ToList(), row),
                    CellIndex: cell is null ? null : GetIndexWithinParent(row?.Elements<TableCell>().ToList(), cell),
                    NearestHeadingText: nearestHeadingText,
                    CurrentParagraphText: Clip(paragraphText, 240),
                    PreviousParagraphText: ClipNullable(previousParagraphText, 160),
                    FollowingParagraphText: ClipNullable(followingParagraphText, 160),
                    CurrentTableRowCount: tableInfo?.RowCount,
                    CurrentTableColumnCount: tableInfo?.ColumnCount));
            }
        }

        return anchors;
    }

    public static IReadOnlyDictionary<string, string> GetPartHashes(string input)
    {
        var hashes = new Dictionary<string, string>(StringComparer.Ordinal);
        using var stream = File.OpenRead(Path.GetFullPath(input));
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries.OrderBy(e => e.FullName, StringComparer.Ordinal))
        {
            using var entryStream = entry.Open();
            using var sha = SHA256.Create();
            hashes[entry.FullName] = Convert.ToHexString(sha.ComputeHash(entryStream));
        }

        return hashes;
    }

    public static IEnumerable<OpenXmlPartRootElement> GetRoots(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart;
        if (mainPart?.Document is not null)
        {
            yield return mainPart.Document;
        }

        foreach (var header in mainPart?.HeaderParts ?? [])
        {
            if (header.Header is not null)
            {
                yield return header.Header;
            }
        }

        foreach (var footer in mainPart?.FooterParts ?? [])
        {
            if (footer.Footer is not null)
            {
                yield return footer.Footer;
            }
        }

        if (mainPart?.FootnotesPart?.Footnotes is not null)
        {
            yield return mainPart.FootnotesPart.Footnotes;
        }

        if (mainPart?.EndnotesPart?.Endnotes is not null)
        {
            yield return mainPart.EndnotesPart.Endnotes;
        }

        if (mainPart?.WordprocessingCommentsPart?.Comments is not null)
        {
            yield return mainPart.WordprocessingCommentsPart.Comments;
        }
    }

    public static string GetParagraphText(Paragraph paragraph)
        => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Trim();

    public static int? GetIndexWithinParent<T>(IReadOnlyList<T>? list, T value) where T : class
    {
        if (list is null)
        {
            return null;
        }

        for (var i = 0; i < list.Count; i++)
        {
            if (ReferenceEquals(list[i], value))
            {
                return i;
            }
        }

        return null;
    }

    public static string? GetCommentText(Comment? comment)
    {
        if (comment is null)
        {
            return null;
        }

        var text = string.Concat(comment.Descendants<Text>().Select(node => node.Text)).Trim();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static PackageSummary BuildPackageSummary(string input)
    {
        using var stream = File.OpenRead(input);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var parts = archive.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal).ToList();
        return new PackageSummary(parts.Count, parts);
    }

    private static bool LooksLikeHeading(Paragraph paragraph, string styleId)
    {
        if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return paragraph.ParagraphProperties?.OutlineLevel is not null;
    }

    private static string GetParagraphSource(Paragraph paragraph)
        => GetPartSource(paragraph);

    private static string GetPartSource(Paragraph paragraph)
    {
        var root = paragraph.Ancestors<OpenXmlPartRootElement>().LastOrDefault();
        return root switch
        {
            Document => "mainDocument",
            Header => "header",
            Footer => "footer",
            Footnotes => "footnotes",
            Endnotes => "endnotes",
            Comments => "comments",
            null => "unknown",
            _ => root.LocalName
        };
    }

    private static bool HasParagraphDirectFormatting(Paragraph paragraph)
    {
        var pPr = paragraph.ParagraphProperties;
        if (pPr is null)
        {
            return false;
        }

        return pPr.ChildElements.Any(child =>
            child is not ParagraphStyleId &&
            child is not NumberingProperties &&
            child is not SectionProperties);
    }

    private static bool HasRunDirectFormatting(Run run)
    {
        var rPr = run.RunProperties;
        if (rPr is null)
        {
            return false;
        }

        return rPr.ChildElements.Any(child => child is not RunStyle);
    }

    private static IReadOnlyList<TableMetadata> BuildTableMetadata(IReadOnlyList<Table> tables)
    {
        var result = new List<TableMetadata>(tables.Count);

        for (var tableIndex = 0; tableIndex < tables.Count; tableIndex++)
        {
            var table = tables[tableIndex];
            var rows = table.Elements<TableRow>().ToList();
            var previewRows = new List<IReadOnlyList<string>>();
            var rowWidths = new List<int>(rows.Count);
            var rowCellCounts = new List<int>(rows.Count);
            var columnCount = 0;

            foreach (var row in rows.Take(3))
            {
                var cells = row.Elements<TableCell>()
                    .Select(cell => Clip(string.Concat(cell.Descendants<Text>().Select(text => text.Text)).Trim(), 80))
                    .Take(4)
                    .ToList();
                var rowWidth = GetTableRowWidth(row);
                var rowCellCount = row.Elements<TableCell>().Count();
                rowWidths.Add(rowWidth);
                rowCellCounts.Add(rowCellCount);
                columnCount = Math.Max(columnCount, rowWidth);
                previewRows.Add(cells);
            }

            foreach (var row in rows.Skip(previewRows.Count))
            {
                var rowWidth = GetTableRowWidth(row);
                var rowCellCount = row.Elements<TableCell>().Count();
                rowWidths.Add(rowWidth);
                rowCellCounts.Add(rowCellCount);
                columnCount = Math.Max(columnCount, rowWidth);
            }

            if (columnCount == 0)
            {
                columnCount = rowWidths.DefaultIfEmpty(0).Max();
            }

            result.Add(new TableMetadata(tableIndex, rows.Count, columnCount, rowWidths, rowCellCounts, previewRows));
        }

        return result;
    }

    internal static int GetTableRowWidth(TableRow row)
    {
        var width = GetTableRowOffset(row.TableRowProperties, "gridBefore");
        foreach (var cell in row.Elements<TableCell>())
        {
            width += GetTableCellWidth(cell);
        }

        width += GetTableRowOffset(row.TableRowProperties, "gridAfter");

        return width;
    }

    private static int GetTableRowOffset(OpenXmlElement? rowProperties, string localName)
    {
        if (rowProperties is null)
        {
            return 0;
        }

        var offset = rowProperties.ChildElements.FirstOrDefault(child => child.LocalName == localName);
        if (offset is null)
        {
            return 0;
        }

        var valAttribute = offset.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val");
        if (string.IsNullOrWhiteSpace(valAttribute.Value))
        {
            return 0;
        }

        return int.TryParse(valAttribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Max(0, value)
            : 0;
    }

    private static int GetTableCellWidth(TableCell cell)
    {
        var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
        return Math.Max(1, span);
    }

    private static string? GetNearestHeadingText(
        IReadOnlyList<Paragraph> bodyParagraphs,
        IReadOnlyList<string> bodyParagraphTexts,
        int paragraphIndex)
    {
        for (var index = paragraphIndex; index >= 0; index--)
        {
            var paragraph = bodyParagraphs[index];
            var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if ((styleId is not null && LooksLikeHeading(paragraph, styleId)) || paragraph.ParagraphProperties?.OutlineLevel is not null)
            {
                var text = bodyParagraphTexts[index];
                return string.IsNullOrWhiteSpace(text) ? null : Clip(text, 160);
            }
        }

        return null;
    }

    private static string Clip(string text, int max)
        => text.Length <= max ? text : text[..max] + "...";

    private static string? ClipNullable(string? text, int max)
        => string.IsNullOrWhiteSpace(text) ? null : Clip(text, max);
}
