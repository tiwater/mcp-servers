using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTableRowCopy
{
    public const string Command = "docx_copy_table_rows";
    private const string MainStory = "/word/document.xml";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var options = new JsonSerializerOptions(Json.Options)
        {
            PropertyNameCaseInsensitive = false,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        };
        var request = JsonSerializer.Deserialize<TableRowCopyRequest>(File.ReadAllText(args[0]), options)
            ?? throw new InvalidOperationException("table-row-copy-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = Describe(request.ReceiptOutput),
            output = Describe(receipt.Output),
            summary = new
            {
                pass = true,
                operationCount = request.Changes.Count,
                appliedCount = receipt.Changes.Count,
            },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static TableRowCopyReceipt Apply(TableRowCopyRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        RequireAbsolutePath(request.TargetDocument.Input, "targetDocument.input");
        RequireAbsolutePath(request.Output, "output");
        RequireAbsolutePath(request.ReceiptOutput, "receiptOutput");
        var targetPath = request.TargetDocument.Input;
        var outputPath = request.Output;
        var receiptPath = request.ReceiptOutput;
        RequireNewPath(outputPath, "output");
        RequireNewPath(receiptPath, "receiptOutput");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, receiptPath))
            throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, targetPath))
            throw new InvalidOperationException("output-must-not-overwrite-input");

        var prepared = Prepare(request, targetPath);
        if (prepared.Select(change => change.TargetTable.Reference).Distinct(StringComparer.Ordinal).Count() != prepared.Count)
            throw new InvalidOperationException("target-table-must-have-one-change");
        IReadOnlyDictionary<string, int> baselineIssues;
        using (var targetDocument = WordprocessingDocument.Open(targetPath, false))
        {
            baselineIssues = ValidationIssueCounts(targetDocument);
            PreflightImports(targetDocument, prepared);
        }

        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            if (!StringComparer.Ordinal.Equals(Observation.CurrentRevision(targetPath).Id, request.TargetDocument.Revision))
                throw new InvalidOperationException("stale-revision");
            foreach (var source in prepared.Select(change => (change.SourcePath, change.SourceRevision)).Distinct())
                if (!StringComparer.Ordinal.Equals(Observation.CurrentRevision(source.SourcePath).Id, source.SourceRevision))
                    throw new InvalidOperationException("stale-revision");
            File.Copy(targetPath, temporaryPath, overwrite: false);
            using (var outputDocument = WordprocessingDocument.Open(temporaryPath, true))
            {
                ApplyChanges(outputDocument, prepared);
                outputDocument.MainDocumentPart?.Document?.Save();
                var issues = ValidationIssueCounts(outputDocument);
                var added = issues.FirstOrDefault(item => item.Value > baselineIssues.GetValueOrDefault(item.Key));
                if (added.Key is not null)
                    throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
            }
            File.Move(temporaryPath, outputPath);
            var outputRevision = Observation.CurrentRevision(outputPath);
            var readback = ReadBack(outputPath, prepared);
            var receipt = new TableRowCopyReceipt(
                "tiwater.docx-copy-table-rows-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                Observation.CurrentRevision(targetPath),
                outputRevision,
                readback,
                outputPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
            if (File.Exists(receiptPath)) File.Delete(receiptPath);
            throw;
        }
    }

    private static IReadOnlyList<PreparedChange> Prepare(TableRowCopyRequest request, string targetPath)
    {
        var result = new List<PreparedChange>(request.Changes.Count);
        using var targetDocument = WordprocessingDocument.Open(targetPath, false);
        for (var changeIndex = 0; changeIndex < request.Changes.Count; changeIndex++)
        {
            var change = request.Changes[changeIndex];
            var targetReferences = new[]
            {
                change.TargetTableRef,
                change.TargetRows.FirstRef,
                change.TargetRows.LastRef,
            }.Concat(change.Columns.Select(column => column.TargetHeaderRef)).ToArray();
            var targetResolved = Observation.ResolveReferences(targetPath, request.TargetDocument.Revision, targetReferences);
            RequireKinds(targetResolved, "table", "row", "row", change.Columns.Count, "cell", "target");
            var targetTable = Resolve<Table>(targetDocument, targetResolved[0], "target-table");
            var allTargetRows = targetTable.Elements<TableRow>().ToArray();
            var targetRows = Range(allTargetRows,
                Resolve<TableRow>(targetDocument, targetResolved[1], "target-first-row"),
                Resolve<TableRow>(targetDocument, targetResolved[2], "target-last-row"), "target");
            var targetHeaderCells = targetResolved.Skip(3).Select(item =>
                Resolve<TableCell>(targetDocument, item, "target-header-cell")).ToArray();
            RequireOneHeaderRow(targetHeaderCells, targetRows, "target");
            var targetColumns = targetHeaderCells.Select(cell => GridStart(targetTable, cell)).ToArray();
            var targetGridCount = GridColumnCount(targetTable);
            RequireCompleteColumns(targetColumns, targetGridCount, "target");
            RequireClosedRangeBoundary(allTargetRows, targetRows, targetGridCount, "target");
            if (targetRows.SelectMany(row => row.Elements<TableCell>()).Any(cell =>
                    cell.ChildElements.Any(element => element is not TableCellProperties and not Paragraph)))
                throw new InvalidOperationException("target-row-contains-non-paragraph-cell-content");

            RequireAbsolutePath(change.SourceDocument.Input, "sourceDocument.input");
            var sourcePath = change.SourceDocument.Input;
            using var sourceDocument = WordprocessingDocument.Open(sourcePath, false);
            var sourceReferences = new[]
            {
                change.SourceTableRef,
                change.SourceRows.FirstRef,
                change.SourceRows.LastRef,
            }.Concat(change.SourceRows.ExcludeRefs ?? [])
             .Concat(change.Columns.Select(column => column.SourceHeaderRef)).ToArray();
            var sourceResolved = Observation.ResolveReferences(sourcePath, change.SourceDocument.Revision, sourceReferences);
            RequireKinds(sourceResolved, "table", "row", "row", (change.SourceRows.ExcludeRefs?.Count ?? 0) + change.Columns.Count, null, "source");
            var sourceTable = Resolve<Table>(sourceDocument, sourceResolved[0], "source-table");
            var allSourceRows = sourceTable.Elements<TableRow>().ToArray();
            var sourceRange = Range(allSourceRows,
                Resolve<TableRow>(sourceDocument, sourceResolved[1], "source-first-row"),
                Resolve<TableRow>(sourceDocument, sourceResolved[2], "source-last-row"), "source");
            var excludedCount = change.SourceRows.ExcludeRefs?.Count ?? 0;
            if (change.SourceRows.ExcludeRefs is { } excludeRefs
                && excludeRefs.Distinct(StringComparer.Ordinal).Count() != excludeRefs.Count)
                throw new InvalidOperationException("source-excluded-row-ref-duplicate");
            var excluded = sourceResolved.Skip(3).Take(excludedCount)
                .Select(item => Resolve<TableRow>(sourceDocument, item, "source-excluded-row"))
                .ToHashSet(ReferenceEqualityComparer.Instance);
            if (excluded.Any(row => !sourceRange.Contains(row, ReferenceEqualityComparer.Instance)))
                throw new InvalidOperationException("source-excluded-row-outside-range");
            var sourceRows = sourceRange.Where(row => !excluded.Contains(row)).ToArray();
            if (sourceRows.Length == 0) throw new InvalidOperationException("source-row-selection-must-not-be-empty");
            var sourceHeaders = sourceResolved.Skip(3 + excludedCount).ToArray();
            if (sourceHeaders.Any(item => item.Kind != "cell"))
                throw new InvalidOperationException("source-header-ref-must-be-cell");
            var sourceHeaderCells = sourceHeaders.Select(item =>
                Resolve<TableCell>(sourceDocument, item, "source-header-cell")).ToArray();
            RequireOneHeaderRow(sourceHeaderCells, sourceRange, "source");
            var sourceColumns = sourceHeaderCells.Select(cell => GridStart(sourceTable, cell)).ToArray();
            var sourceGridCount = GridColumnCount(sourceTable);
            RequireCompleteColumns(sourceColumns, sourceGridCount, "source");

            var preparedRows = sourceRows.Select(row => PrepareRow(row, sourceColumns, targetColumns, change.Columns)).ToArray();
            RequireClosedMergeChains(preparedRows, targetGridCount, "source");
            if (sourceRows.Length > targetRows.Count && targetRows.Count > 1
                && PresentationSignature(targetRows[^1]) != PresentationSignature(targetRows[^2]))
                throw new InvalidOperationException("target-last-row-is-not-repeatable");
            result.Add(new PreparedChange(changeIndex, targetResolved[0],
                targetRows.Select(Observation.NativePathFor).ToArray(), sourcePath, change.SourceDocument.Revision,
                targetGridCount, preparedRows));
        }
        return result;
    }

    private static PreparedRow PrepareRow(
        TableRow sourceRow,
        IReadOnlyList<int> sourceColumns,
        IReadOnlyList<int> targetColumns,
        IReadOnlyList<TableRowCopyColumn> columns)
    {
        var mappings = columns.Select((column, index) => new ColumnMapping(
            sourceColumns[index], targetColumns[index], column.Content)).ToArray();
        var prepared = new List<PreparedCell>();
        foreach (var cell in Cells(sourceRow))
        {
            var selected = mappings.Where(mapping => mapping.SourceColumn >= cell.Start
                && mapping.SourceColumn < cell.Start + cell.Span).ToArray();
            if (selected.Length != cell.Span
                || !selected.Select(mapping => mapping.SourceColumn).Order().SequenceEqual(Enumerable.Range(cell.Start, cell.Span)))
                throw new InvalidOperationException("source-cell-span-is-not-fully-mapped");
            var targetPositions = selected.Select(mapping => mapping.TargetColumn).Order().ToArray();
            if (!targetPositions.SequenceEqual(Enumerable.Range(targetPositions[0], targetPositions.Length)))
                throw new InvalidOperationException("source-cell-maps-to-noncontiguous-target-columns");
            var contentModes = selected.Select(mapping => mapping.Content).Distinct(StringComparer.Ordinal).ToArray();
            if (contentModes.Length != 1) throw new InvalidOperationException("source-cell-content-mode-conflict");
            var paragraphs = cell.Cell.Elements<Paragraph>().ToArray();
            if (cell.Cell.ChildElements.Any(element => element is not TableCellProperties and not Paragraph))
                throw new InvalidOperationException("source-cell-contains-non-paragraph-content");
            if (contentModes[0] == "first-paragraph") paragraphs = paragraphs.Take(1).ToArray();
            else if (contentModes[0] != "all-paragraphs") throw new InvalidOperationException("column-content-invalid");
            prepared.Add(new PreparedCell(
                targetPositions[0],
                paragraphs.Select(NativeContentCopy.CloneParagraph).ToArray(),
                cell.Cell.TableCellProperties?.VerticalMerge?.CloneNode(true) as VerticalMerge,
                targetPositions.Length));
        }
        return new PreparedRow(prepared.OrderBy(cell => cell.TargetColumn).ToArray());
    }

    private static void PreflightImports(WordprocessingDocument targetDocument, IReadOnlyList<PreparedChange> changes)
    {
        var targetMain = targetDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        foreach (var group in changes.GroupBy(change => change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(change => change.Rows).SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Paragraphs).Cast<OpenXmlElement>().ToArray();
            foreach (var relationshipId in DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal))
                if (!DocxObjectActions.CanCopyRelationship(sourceMain, relationshipId, out var error))
                    throw new InvalidOperationException(error);
            if (!DocxObjectActions.TryImportStyles(sourceMain, targetMain, roots, apply: false, out var styleError))
                throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, targetMain, roots, apply: false, out var numberingError))
                throw new InvalidOperationException(numberingError);
        }
    }

    private static void ApplyChanges(WordprocessingDocument outputDocument, IReadOnlyList<PreparedChange> changes)
    {
        var outputMain = outputDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        var outputBody = outputMain.Document?.Body ?? throw new InvalidOperationException("target-body-not-found");
        foreach (var group in changes.GroupBy(change => change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(change => change.Rows).SelectMany(row => row.Cells)
                .SelectMany(cell => cell.Paragraphs).Cast<OpenXmlElement>().ToArray();
            if (!DocxObjectActions.TryImportStyles(sourceMain, outputMain, roots, apply: true, out var styleError))
                throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, outputMain, roots, apply: true, out var numberingError))
                throw new InvalidOperationException(numberingError);
            var relationshipMap = DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal)
                .ToDictionary(id => id, id => DocxObjectActions.CopyRelationship(sourceMain, outputMain, id), StringComparer.Ordinal);

            foreach (var change in group)
            {
                var table = Resolve<Table>(outputDocument, change.TargetTable, "output-target-table");
                var originalRows = change.TargetRowPaths.Select(path =>
                    Observation.ResolveNativePath(outputDocument, MainStory, path) as TableRow
                    ?? throw new InvalidOperationException("output-target-row-not-found")).ToArray();
                var insertionPoint = originalRows[^1].NextSibling();
                var templates = originalRows.Select(row => (TableRow)row.CloneNode(true)).ToArray();
                foreach (var row in originalRows) row.Remove();
                for (var rowIndex = 0; rowIndex < change.Rows.Count; rowIndex++)
                {
                    var outputRow = (TableRow)templates[Math.Min(rowIndex, templates.Length - 1)].CloneNode(true);
                    var templateCells = Cells(outputRow);
                    outputRow.RemoveAllChildren<TableCell>();
                    foreach (var cellChange in change.Rows[rowIndex].Cells)
                    {
                        var targetSlot = templateCells.SingleOrDefault(item =>
                            cellChange.TargetColumn >= item.Start && cellChange.TargetColumn < item.Start + item.Span);
                        if (targetSlot is null) throw new InvalidOperationException("target-row-does-not-cover-selected-column");
                        var targetCell = (TableCell)targetSlot.Cell.CloneNode(true);
                        SetGridSpan(targetCell, cellChange.TargetSpan);
                        var paragraphs = cellChange.Paragraphs.Select(paragraph => (Paragraph)paragraph.CloneNode(true)).ToArray();
                        foreach (var paragraph in paragraphs) DocxObjectActions.RewriteRelationships(paragraph, relationshipMap);
                        DocxObjectActions.RemapDrawingIds(outputBody, paragraphs.Cast<OpenXmlElement>().ToArray());
                        NativeContentCopy.ReplaceCellContent(targetCell, paragraphs);
                        SetVerticalMerge(targetCell, cellChange.VerticalMerge);
                        outputRow.Append(targetCell);
                    }
                    if (Cells(outputRow).Sum(cell => cell.Span) != change.TargetGridCount)
                        throw new InvalidOperationException("output-row-grid-coverage-mismatch");
                    if (insertionPoint is null) table.Append(outputRow);
                    else table.InsertBefore(outputRow, insertionPoint);
                }
            }
        }
    }

    private static IReadOnlyList<TableRowCopyReadback> ReadBack(string output, IReadOnlyList<PreparedChange> changes)
    {
        using var document = WordprocessingDocument.Open(output, false);
        return changes.Select(change =>
        {
            var table = Resolve<Table>(document, change.TargetTable, "output-target-table");
            var allRows = table.Elements<TableRow>().ToArray();
            var firstOriginalPath = change.TargetRowPaths[0];
            var firstIndex = NativeSiblingIndex(firstOriginalPath) - 1;
            var rows = allRows.Skip(firstIndex).Take(change.Rows.Count).ToArray();
            if (rows.Length != change.Rows.Count) throw new InvalidOperationException("output-readback-row-count-mismatch");
            for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                foreach (var cellChange in change.Rows[rowIndex].Cells)
                {
                    var slot = Cells(rows[rowIndex]).Single(item => item.Start == cellChange.TargetColumn);
                    if (slot.Span != cellChange.TargetSpan)
                        throw new InvalidOperationException("output-readback-grid-span-mismatch");
                    var cell = slot.Cell;
                    var expected = string.Concat(cellChange.Paragraphs.Select(paragraph => paragraph.InnerText));
                    if (!StringComparer.Ordinal.Equals(cell.InnerText, expected))
                        throw new InvalidOperationException("output-readback-content-mismatch");
                    var expectedMerge = cellChange.VerticalMerge;
                    var actualMerge = cell.TableCellProperties?.VerticalMerge;
                    if ((expectedMerge is null) != (actualMerge is null)
                        || expectedMerge?.Val?.Value != actualMerge?.Val?.Value)
                        throw new InvalidOperationException("output-readback-vertical-merge-mismatch");
                }
            }
            return new TableRowCopyReadback(change.Index, change.Rows.Count, rows.Length,
                rows.Select(row => row.InnerText).ToArray());
        }).ToArray();
    }

    private static int NativeSiblingIndex(string nativePath)
    {
        var segment = nativePath.TrimEnd('/').Split('/')[^1];
        var start = segment.LastIndexOf('[') + 1;
        return int.Parse(segment[start..^1]);
    }

    private static void SetVerticalMerge(TableCell target, VerticalMerge? source)
    {
        var properties = target.TableCellProperties ?? target.PrependChild(new TableCellProperties());
        properties.RemoveAllChildren<VerticalMerge>();
        if (source is not null && !properties.AddChild(source.CloneNode(true), true))
            throw new InvalidOperationException("target-vertical-merge-not-supported");
    }

    private static void SetGridSpan(TableCell target, int span)
    {
        var properties = target.TableCellProperties ?? target.PrependChild(new TableCellProperties());
        properties.RemoveAllChildren<GridSpan>();
        if (span > 1 && !properties.AddChild(new GridSpan { Val = span }, true))
            throw new InvalidOperationException("target-grid-span-not-supported");
    }

    private static IReadOnlyList<CellSlot> Cells(TableRow row)
    {
        var start = 0;
        var result = new List<CellSlot>();
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            result.Add(new CellSlot(cell, start, span));
            start += span;
        }
        return result;
    }

    private static int GridStart(Table table, TableCell cell)
    {
        if (!cell.Ancestors<Table>().Any(item => ReferenceEquals(item, table)))
            throw new InvalidOperationException("header-cell-does-not-belong-to-selected-table");
        var row = cell.Ancestors<TableRow>().FirstOrDefault()
            ?? throw new InvalidOperationException("header-cell-row-not-found");
        return Cells(row).Single(item => ReferenceEquals(item.Cell, cell)).Start;
    }

    private static int GridColumnCount(Table table)
    {
        var count = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
        if (count <= 0) throw new InvalidOperationException("table-grid-must-declare-columns");
        if (table.Elements<TableRow>().Any(row => Cells(row).Sum(cell => cell.Span) != count))
            throw new InvalidOperationException("table-row-grid-coverage-mismatch");
        return count;
    }

    private static void RequireCompleteColumns(IReadOnlyList<int> columns, int gridCount, string name)
    {
        if (columns.Count != gridCount
            || !columns.Order().SequenceEqual(Enumerable.Range(0, gridCount)))
            throw new InvalidOperationException($"{name}-header-columns-must-cover-table-grid");
    }

    private static void RequireOneHeaderRow(
        IReadOnlyList<TableCell> headerCells,
        IReadOnlyList<TableRow> selectedRows,
        string name)
    {
        var rows = headerCells.Select(cell => cell.Ancestors<TableRow>().First()).Distinct(ReferenceEqualityComparer.Instance).ToArray();
        if (rows.Length != 1) throw new InvalidOperationException($"{name}-header-cells-must-share-one-row");
        if (selectedRows.Contains(rows[0], ReferenceEqualityComparer.Instance))
            throw new InvalidOperationException($"{name}-header-row-must-be-outside-selected-rows");
    }

    private static void RequireClosedRangeBoundary(
        IReadOnlyList<TableRow> allRows,
        IReadOnlyList<TableRow> selectedRows,
        int gridCount,
        string name)
    {
        var firstIndex = allRows.IndexOfReference(selectedRows[0]);
        var lastIndex = allRows.IndexOfReference(selectedRows[^1]);
        for (var column = 0; column < gridCount; column++)
        {
            if (MergeValue(selectedRows[0], column) == MergedCellValues.Continue)
                throw new InvalidOperationException($"{name}-row-range-starts-inside-vertical-merge");
            if (lastIndex + 1 < allRows.Count && MergeValue(allRows[lastIndex + 1], column) == MergedCellValues.Continue)
                throw new InvalidOperationException($"{name}-row-range-ends-inside-vertical-merge");
        }
    }

    private static void RequireClosedMergeChains(IReadOnlyList<PreparedRow> rows, int gridCount, string name)
    {
        for (var column = 0; column < gridCount; column++)
        {
            var active = false;
            foreach (var row in rows)
            {
                var cell = row.Cells.Single(item => column >= item.TargetColumn && column < item.TargetColumn + item.TargetSpan);
                if (cell.VerticalMerge is null) active = false;
                else if (cell.VerticalMerge.Val?.Value == MergedCellValues.Restart) active = true;
                else if (!active) throw new InvalidOperationException($"{name}-vertical-merge-continue-without-restart");
            }
        }
    }

    private static MergedCellValues? MergeValue(TableRow row, int column)
    {
        var merge = Cells(row).Single(cell => column >= cell.Start && column < cell.Start + cell.Span)
            .Cell.TableCellProperties?.VerticalMerge;
        if (merge is null) return null;
        return merge.Val?.Value ?? MergedCellValues.Continue;
    }

    private static string PresentationSignature(TableRow row)
    {
        var parts = new List<string> { row.TableRowProperties?.OuterXml ?? string.Empty };
        foreach (var slot in Cells(row))
        {
            var properties = slot.Cell.TableCellProperties?.CloneNode(true) as TableCellProperties ?? new TableCellProperties();
            properties.RemoveAllChildren<GridSpan>();
            properties.RemoveAllChildren<VerticalMerge>();
            parts.Add(properties.OuterXml);
            var paragraph = slot.Cell.Elements<Paragraph>().FirstOrDefault();
            parts.Add(paragraph?.ParagraphProperties?.OuterXml ?? string.Empty);
            parts.Add(paragraph?.Descendants<Run>().FirstOrDefault()?.RunProperties?.OuterXml ?? string.Empty);
        }
        return string.Join("\0", parts);
    }

    private static IReadOnlyList<TableRow> Range(IReadOnlyList<TableRow> rows, TableRow first, TableRow last, string name)
    {
        var firstIndex = rows.IndexOfReference(first);
        var lastIndex = rows.IndexOfReference(last);
        if (firstIndex < 0 || lastIndex < firstIndex) throw new InvalidOperationException($"{name}-row-range-invalid");
        return rows.Skip(firstIndex).Take(lastIndex - firstIndex + 1).ToArray();
    }

    private static T Resolve<T>(WordprocessingDocument document, ResolvedDocxReference reference, string name)
        where T : OpenXmlElement
    {
        if (reference.StoryPart != MainStory) throw new InvalidOperationException($"{name}-must-be-main-document-object");
        return Observation.ResolveNativePath(document, reference.StoryPart, reference.NativePath) as T
            ?? throw new InvalidOperationException($"{name}-kind-invalid");
    }

    private static void RequireKinds(IReadOnlyList<ResolvedDocxReference> refs, string first, string second,
        string third, int remainingCount, string? remainingKind, string name)
    {
        if (refs.Count != 3 + remainingCount || refs[0].Kind != first || refs[1].Kind != second || refs[2].Kind != third)
            throw new InvalidOperationException($"{name}-reference-kind-invalid");
        if (remainingKind is not null && refs.Skip(3).Any(item => item.Kind != remainingKind))
            throw new InvalidOperationException($"{name}-reference-kind-invalid");
    }

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document)
            .GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static void RequireNewPath(string path, string name)
    {
        if (File.Exists(path) || Directory.Exists(path)) throw new InvalidOperationException($"{name}-already-exists");
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            throw new InvalidOperationException($"{name}-directory-not-found");
    }

    private static void RequireAbsolutePath(string path, string name)
    {
        if (!Path.IsPathFullyQualified(path)) throw new InvalidOperationException($"{name}-must-be-absolute");
    }

    private static CopyContentArtifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new CopyContentArtifact(Path.GetFullPath(path),
            Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), stream.Length);
    }

    private sealed record CellSlot(TableCell Cell, int Start, int Span);
    private sealed record ColumnMapping(int SourceColumn, int TargetColumn, string Content);
    private sealed record PreparedCell(int TargetColumn, IReadOnlyList<Paragraph> Paragraphs, VerticalMerge? VerticalMerge, int TargetSpan);
    private sealed record PreparedRow(IReadOnlyList<PreparedCell> Cells);
    private sealed record PreparedChange(int Index, ResolvedDocxReference TargetTable,
        IReadOnlyList<string> TargetRowPaths, string SourcePath, string SourceRevision,
        int TargetGridCount, IReadOnlyList<PreparedRow> Rows);
}

internal static class TableRowCopyReferenceExtensions
{
    internal static int IndexOfReference<T>(this IReadOnlyList<T> values, T expected) where T : class
    {
        for (var index = 0; index < values.Count; index++) if (ReferenceEquals(values[index], expected)) return index;
        return -1;
    }
}

public sealed record TableRowCopyDocument(string Input, string Revision);
public sealed record TableRowCopyRange(string FirstRef, string LastRef);
public sealed record TableRowCopySourceRange(string FirstRef, string LastRef, IReadOnlyList<string>? ExcludeRefs = null);
public sealed record TableRowCopyColumn(string SourceHeaderRef, string TargetHeaderRef, string Content);
public sealed record TableRowCopyChange(string TargetTableRef, TableRowCopyRange TargetRows,
    TableRowCopyDocument SourceDocument, string SourceTableRef, TableRowCopySourceRange SourceRows,
    IReadOnlyList<TableRowCopyColumn> Columns);
public sealed record TableRowCopyRequest(TableRowCopyDocument TargetDocument,
    IReadOnlyList<TableRowCopyChange> Changes, string Output, string ReceiptOutput);
public sealed record TableRowCopyReadback(int ChangeIndex, int SourceRowCount, int OutputRowCount,
    IReadOnlyList<string> RowTexts);
public sealed record TableRowCopyReceipt(string Schema, string Provider, string ToolVersion,
    DocxRevision TargetRevision, DocxRevision OutputRevision, IReadOnlyList<TableRowCopyReadback> Changes, string Output);
