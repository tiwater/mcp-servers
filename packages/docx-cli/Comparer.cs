namespace Dockit.Docx;

public static class Comparer
{
    public static ComparisonReport Compare(string oldPath, string newPath)
    {
        var oldInspection = Inspector.Inspect(oldPath);
        var newInspection = Inspector.Inspect(newPath);
        var oldHashes = Inspector.GetPartHashes(oldPath);
        var newHashes = Inspector.GetPartHashes(newPath);

        var commonParts = oldHashes.Keys.Intersect(newHashes.Keys, StringComparer.Ordinal).ToList();
        var differentParts = commonParts
            .Where(part => !StringComparer.Ordinal.Equals(oldHashes[part], newHashes[part]))
            .OrderBy(part => part, StringComparer.Ordinal)
            .ToList();

        return new ComparisonReport(
            OldFile: oldInspection.File,
            NewFile: newInspection.File,
            PackageComparison: new PackageComparison(
                SamePartCount: commonParts.Count - differentParts.Count,
                DifferentPartCount: differentParts.Count,
                DifferentParts: differentParts),
            MetricDiffs: BuildMetricDiffs(oldInspection, newInspection),
            StyleDiffs: BuildStyleDiffs(oldInspection, newInspection),
            OldInspection: oldInspection,
            NewInspection: newInspection);
    }

    private static IReadOnlyList<MetricDiff> BuildMetricDiffs(InspectionReport oldInspection, InspectionReport newInspection)
        =>
        [
            new("partCount", oldInspection.Package.PartCount, newInspection.Package.PartCount),
            new("paragraphCount", oldInspection.Content.ParagraphCount, newInspection.Content.ParagraphCount),
            new("tableCount", oldInspection.Content.TableCount, newInspection.Content.TableCount),
            new("sectionCount", oldInspection.Content.SectionCount, newInspection.Content.SectionCount),
            new("headerPartCount", oldInspection.Content.HeaderPartCount, newInspection.Content.HeaderPartCount),
            new("footerPartCount", oldInspection.Content.FooterPartCount, newInspection.Content.FooterPartCount),
            new("commentCount", oldInspection.Annotations.CommentCount, newInspection.Annotations.CommentCount),
            new("footnoteCount", oldInspection.Annotations.FootnoteCount, newInspection.Annotations.FootnoteCount),
            new("endnoteCount", oldInspection.Annotations.EndnoteCount, newInspection.Annotations.EndnoteCount),
            new("trackedChangeElements", oldInspection.Annotations.TrackedChangeElements, newInspection.Annotations.TrackedChangeElements),
            new("bookmarkCount", oldInspection.Structure.BookmarkCount, newInspection.Structure.BookmarkCount),
            new("hyperlinkCount", oldInspection.Structure.HyperlinkCount, newInspection.Structure.HyperlinkCount),
            new("fieldCount", oldInspection.Structure.FieldCount, newInspection.Structure.FieldCount),
            new("contentControlCount", oldInspection.Structure.ContentControlCount, newInspection.Structure.ContentControlCount),
            new("drawingCount", oldInspection.Structure.DrawingCount, newInspection.Structure.DrawingCount),
            new("paragraphsWithDirectFormatting", oldInspection.Formatting.ParagraphsWithDirectFormatting, newInspection.Formatting.ParagraphsWithDirectFormatting),
            new("runsWithDirectFormatting", oldInspection.Formatting.RunsWithDirectFormatting, newInspection.Formatting.RunsWithDirectFormatting),
        ];

    private static StyleDiffSummary BuildStyleDiffs(InspectionReport oldInspection, InspectionReport newInspection)
        => new(
            AddedParagraphStyles: DiffAdded(oldInspection.Styles.ParagraphStylesInUse, newInspection.Styles.ParagraphStylesInUse),
            RemovedParagraphStyles: DiffAdded(newInspection.Styles.ParagraphStylesInUse, oldInspection.Styles.ParagraphStylesInUse),
            AddedRunStyles: DiffAdded(oldInspection.Styles.RunStylesInUse, newInspection.Styles.RunStylesInUse),
            RemovedRunStyles: DiffAdded(newInspection.Styles.RunStylesInUse, oldInspection.Styles.RunStylesInUse));

    private static IReadOnlyList<StyleCount> DiffAdded(IReadOnlyList<StyleCount> before, IReadOnlyList<StyleCount> after)
    {
        var beforeMap = before.ToDictionary(x => x.Style, x => x.Count, StringComparer.Ordinal);
        return after
            .Where(item => !beforeMap.TryGetValue(item.Style, out var count) || count != item.Count)
            .OrderBy(item => item.Style, StringComparer.Ordinal)
            .ToList();
    }
}
