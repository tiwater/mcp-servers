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
            var targetHeaderSlots = targetHeaderCells.Select(cell => GridSlot(targetTable, cell)).ToArray();
            var targetGridCount = GridColumnCount(targetTable);
            RequireCompleteHeader(targetHeaderSlots, targetGridCount, "target");
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
            var sourceHeaderSlots = sourceHeaderCells.Select(cell => GridSlot(sourceTable, cell)).ToArray();
            var sourceGridCount = GridColumnCount(sourceTable);
            RequireCompleteHeader(sourceHeaderSlots, sourceGridCount, "source");

            var mappings = BuildMappings(sourceHeaderSlots, targetHeaderSlots, change.Columns);
            var targetGridWidths = BuildTargetGridWidths(sourceTable, targetTable, mappings);

            var preparedRows = NormalizeMergeStarts(allSourceRows, sourceRows,
                sourceRows.Select(row => PrepareRow(row, mappings, targetGridWidths.Count)).ToArray(),
                mappings, targetGridWidths.Count);
            RequireClosedMergeChains(preparedRows, targetGridWidths.Count, "source");
            result.Add(new PreparedChange(changeIndex, targetResolved[0],
                targetRows.Select(Observation.NativePathFor).ToArray(), sourcePath, change.SourceDocument.Revision,
                targetGridWidths, mappings.Select(mapping => new TargetColumnReshape(
                    mapping.OldTargetStart, mapping.OldTargetSpan, mapping.TargetStart, mapping.TargetSpan)).ToArray(),
                preparedRows));
        }
        return result;
    }

    private static PreparedRow PrepareRow(
        TableRow sourceRow,
        IReadOnlyList<ColumnMapping> mappings,
        int targetGridCount)
    {
        var prepared = new List<PreparedCell>();
        foreach (var cell in Cells(sourceRow))
        {
            var sourcePositions = Enumerable.Range(cell.Start, cell.Span).ToArray();
            var selected = mappings.Where(mapping => sourcePositions.Any(position =>
                position >= mapping.SourceStart && position < mapping.SourceStart + mapping.SourceSpan)).ToArray();
            var mappedSourcePositions = selected.SelectMany(mapping =>
                Enumerable.Range(mapping.SourceStart, mapping.SourceSpan)).Intersect(sourcePositions).Order().ToArray();
            if (!mappedSourcePositions.SequenceEqual(sourcePositions))
                throw new InvalidOperationException("source-cell-span-is-not-fully-mapped");
            var targetPositions = selected.SelectMany(mapping => MapInterval(
                Math.Max(cell.Start, mapping.SourceStart),
                Math.Min(cell.Start + cell.Span, mapping.SourceStart + mapping.SourceSpan),
                mapping.SourceStart, mapping.SourceSpan, mapping.TargetStart, mapping.TargetSpan)).Order().ToArray();
            if (!targetPositions.SequenceEqual(Enumerable.Range(targetPositions[0], targetPositions.Length)))
                throw new InvalidOperationException("source-cell-maps-to-noncontiguous-target-columns");
            var contentModes = selected.Select(mapping => mapping.Content).Distinct(StringComparer.Ordinal).ToArray();
            if (contentModes.Length != 1) throw new InvalidOperationException("source-cell-content-mode-conflict");
            var paragraphs = CloneCellParagraphs(cell.Cell, contentModes[0]);
            prepared.Add(new PreparedCell(
                cell.Start,
                cell.Span,
                targetPositions[0],
                paragraphs,
                cell.Cell.TableCellProperties?.VerticalMerge?.CloneNode(true) as VerticalMerge,
                targetPositions.Length));
        }
        var ordered = prepared.OrderBy(cell => cell.TargetColumn).ToArray();
        var occupied = ordered.SelectMany(cell => Enumerable.Range(cell.TargetColumn, cell.TargetSpan)).ToArray();
        if (occupied.Length == 0 || !occupied.SequenceEqual(Enumerable.Range(occupied[0], occupied.Length)))
            throw new InvalidOperationException("source-row-maps-to-noncontiguous-target-grid");
        return new PreparedRow(ordered, occupied[0], targetGridCount - occupied[^1] - 1);
    }

    private static IReadOnlyList<PreparedRow> NormalizeMergeStarts(
        IReadOnlyList<TableRow> allRows,
        IReadOnlyList<TableRow> selectedRows,
        IReadOnlyList<PreparedRow> preparedRows,
        IReadOnlyList<ColumnMapping> mappings,
        int targetGridCount)
    {
        var active = new bool[targetGridCount];
        var result = new List<PreparedRow>(preparedRows.Count);
        var previousSourceIndex = -2;
        for (var rowIndex = 0; rowIndex < preparedRows.Count; rowIndex++)
        {
            var sourceIndex = allRows.IndexOfReference(selectedRows[rowIndex]);
            if (sourceIndex != previousSourceIndex + 1) Array.Fill(active, false);
            previousSourceIndex = sourceIndex;
            var cells = new List<PreparedCell>(preparedRows[rowIndex].Cells.Count);
            foreach (var original in preparedRows[rowIndex].Cells)
            {
                var cell = original;
                var positions = Enumerable.Range(cell.TargetColumn, cell.TargetSpan).ToArray();
                if (cell.VerticalMerge is null)
                {
                    foreach (var position in positions) active[position] = false;
                }
                else if (cell.VerticalMerge.Val?.Value == MergedCellValues.Restart)
                {
                    foreach (var position in positions) active[position] = true;
                }
                else
                {
                    var activeStates = positions.Select(position => active[position]).Distinct().ToArray();
                    if (activeStates.Length != 1)
                        throw new InvalidOperationException("source-vertical-merge-maps-to-inconsistent-target-columns");
                    if (!activeStates[0])
                    {
                        var origin = FindMergeOrigin(allRows, sourceIndex, cell.SourceColumn);
                        var originCell = PrepareRow(origin, mappings, targetGridCount).Cells.Single(item =>
                            cell.SourceColumn >= item.SourceColumn
                            && cell.SourceColumn < item.SourceColumn + item.SourceSpan);
                        if (originCell.TargetColumn != cell.TargetColumn || originCell.TargetSpan != cell.TargetSpan)
                            throw new InvalidOperationException("source-vertical-merge-span-changes-at-selection-boundary");
                        cell = cell with
                        {
                            Paragraphs = originCell.Paragraphs,
                            VerticalMerge = new VerticalMerge { Val = MergedCellValues.Restart },
                        };
                    }
                    foreach (var position in positions) active[position] = true;
                }
                cells.Add(cell);
            }
            result.Add(preparedRows[rowIndex] with { Cells = cells });
        }
        return result;
    }

    private static TableRow FindMergeOrigin(IReadOnlyList<TableRow> rows, int beforeIndex, int sourceColumn)
    {
        for (var index = beforeIndex - 1; index >= 0; index--)
        {
            var slot = Cells(rows[index]).SingleOrDefault(cell =>
                sourceColumn >= cell.Start && sourceColumn < cell.Start + cell.Span);
            var merge = slot?.Cell.TableCellProperties?.VerticalMerge;
            if (merge is null) break;
            if (merge.Val?.Value == MergedCellValues.Restart) return rows[index];
        }
        throw new InvalidOperationException("source-vertical-merge-origin-not-found");
    }

    private static IReadOnlyList<Paragraph> CloneCellParagraphs(TableCell cell, string contentMode)
    {
        if (cell.ChildElements.Any(element => element is not TableCellProperties
            and not Paragraph and not BookmarkStart and not BookmarkEnd))
            throw new InvalidOperationException("source-cell-contains-non-paragraph-content");
        IEnumerable<Paragraph> paragraphs = cell.Elements<Paragraph>();
        if (contentMode == "first-paragraph") paragraphs = paragraphs.Take(1);
        else if (contentMode != "all-paragraphs") throw new InvalidOperationException("column-content-invalid");
        return paragraphs.Select(NativeContentCopy.CloneParagraph).ToArray();
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
                ReshapeTargetTable(table, change.TargetColumns, change.TargetGridWidths);
                var originalRows = change.TargetRowPaths.Select(path =>
                    Observation.ResolveNativePath(outputDocument, MainStory, path) as TableRow
                    ?? throw new InvalidOperationException("output-target-row-not-found")).ToArray();
                var insertionPoint = originalRows[^1].NextSibling();
                var templates = originalRows.Select(row => (TableRow)row.CloneNode(true)).ToArray();
                foreach (var row in originalRows) row.Remove();
                for (var rowIndex = 0; rowIndex < change.Rows.Count; rowIndex++)
                {
                    var outputRow = (TableRow)SelectTemplate(templates, rowIndex, change.Rows.Count).CloneNode(true);
                    var templateCells = Cells(outputRow);
                    outputRow.RemoveAllChildren<TableCell>();
                    SetGridOmissions(outputRow, change.Rows[rowIndex].GridBefore, change.Rows[rowIndex].GridAfter);
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
                    if (RowGridBefore(outputRow) + Cells(outputRow).Sum(cell => cell.Span) + RowGridAfter(outputRow)
                        != change.TargetGridWidths.Count)
                        throw new InvalidOperationException("output-row-grid-coverage-mismatch");
                    if (insertionPoint is null) table.Append(outputRow);
                    else table.InsertBefore(outputRow, insertionPoint);
                }
            }
        }
    }

    private static TableRow SelectTemplate(IReadOnlyList<TableRow> templates, int rowIndex, int outputCount)
    {
        if (rowIndex < templates.Count - 1) return templates[rowIndex];
        if (rowIndex == outputCount - 1) return templates[^1];
        if (templates.Count >= 3) return templates[^2];
        if (templates.Count == 1) return templates[0];
        if (PresentationSignature(templates[0]) == PresentationSignature(templates[1])) return templates[0];
        throw new InvalidOperationException("target-row-range-has-no-repeatable-interior-style");
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
        var start = RowGridBefore(row);
        var result = new List<CellSlot>();
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            result.Add(new CellSlot(cell, start, span));
            start += span;
        }
        return result;
    }

    private static CellSlot GridSlot(Table table, TableCell cell)
    {
        if (!cell.Ancestors<Table>().Any(item => ReferenceEquals(item, table)))
            throw new InvalidOperationException("header-cell-does-not-belong-to-selected-table");
        var row = cell.Ancestors<TableRow>().FirstOrDefault()
            ?? throw new InvalidOperationException("header-cell-row-not-found");
        return Cells(row).Single(item => ReferenceEquals(item.Cell, cell));
    }

    private static int GridColumnCount(Table table)
    {
        var count = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
        if (count <= 0) throw new InvalidOperationException("table-grid-must-declare-columns");
        if (table.Elements<TableRow>().Any(row =>
                RowGridBefore(row) + Cells(row).Sum(cell => cell.Span) + RowGridAfter(row) != count))
            throw new InvalidOperationException("table-row-grid-coverage-mismatch");
        return count;
    }

    private static void RequireCompleteHeader(IReadOnlyList<CellSlot> cells, int gridCount, string name)
    {
        var positions = cells.SelectMany(cell => Enumerable.Range(cell.Start, cell.Span)).Order().ToArray();
        if (positions.Distinct().Count() != positions.Length
            || !positions.SequenceEqual(Enumerable.Range(0, gridCount)))
            throw new InvalidOperationException($"{name}-header-columns-must-cover-table-grid");
    }

    private static IReadOnlyList<ColumnMapping> BuildMappings(
        IReadOnlyList<CellSlot> sourceHeaders,
        IReadOnlyList<CellSlot> targetHeaders,
        IReadOnlyList<TableRowCopyColumn> columns)
    {
        var indexed = columns.Select((column, index) => new
        {
            Column = column,
            Source = sourceHeaders[index],
            Target = targetHeaders[index],
        }).OrderBy(item => item.Target.Start).ToArray();
        var targetStart = 0;
        var result = new List<ColumnMapping>(indexed.Length);
        foreach (var item in indexed)
        {
            var targetSpan = Math.Max(item.Source.Span, item.Target.Span);
            result.Add(new ColumnMapping(item.Source.Start, item.Source.Span,
                item.Target.Start, item.Target.Span, targetStart, targetSpan, item.Column.Content));
            targetStart += targetSpan;
        }
        return result;
    }

    private static IReadOnlyList<long> BuildTargetGridWidths(
        Table sourceTable,
        Table targetTable,
        IReadOnlyList<ColumnMapping> mappings)
    {
        var sourceWidths = GridWidths(sourceTable);
        var targetWidths = GridWidths(targetTable);
        var result = new List<long>();
        foreach (var mapping in mappings.OrderBy(item => item.OldTargetStart))
        {
            var targetTotal = targetWidths.Skip(mapping.OldTargetStart).Take(mapping.OldTargetSpan).Sum();
            if (mapping.TargetSpan == mapping.OldTargetSpan)
                result.AddRange(targetWidths.Skip(mapping.OldTargetStart).Take(mapping.OldTargetSpan));
            else
                result.AddRange(ScaleWidths(
                    sourceWidths.Skip(mapping.SourceStart).Take(mapping.SourceSpan).ToArray(), targetTotal));
        }
        return result;
    }

    private static IReadOnlyList<long> GridWidths(Table table)
    {
        var columns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToArray()
            ?? throw new InvalidOperationException("table-grid-must-declare-columns");
        return columns.Select(column =>
        {
            if (!long.TryParse(column.Width?.Value, out var width) || width <= 0)
                throw new InvalidOperationException("table-grid-column-width-invalid");
            return width;
        }).ToArray();
    }

    private static IReadOnlyList<long> ScaleWidths(IReadOnlyList<long> weights, long total)
    {
        if (weights.Count == 0 || total < weights.Count)
            throw new InvalidOperationException("target-grid-width-cannot-be-subdivided");
        var result = new List<long>(weights.Count);
        var remainingTotal = total;
        var remainingWeight = weights.Sum();
        for (var index = 0; index < weights.Count; index++)
        {
            var remainingSlots = weights.Count - index - 1;
            var width = index == weights.Count - 1
                ? remainingTotal
                : Math.Clamp((long)Math.Round((decimal)remainingTotal * weights[index] / remainingWeight),
                    1, remainingTotal - remainingSlots);
            result.Add(width);
            remainingTotal -= width;
            remainingWeight -= weights[index];
        }
        return result;
    }

    private static void ReshapeTargetTable(
        Table table,
        IReadOnlyList<TargetColumnReshape> columns,
        IReadOnlyList<long> widths)
    {
        foreach (var row in table.Elements<TableRow>()) ReshapeTargetRow(row, columns, widths.Count);
        var grid = table.GetFirstChild<TableGrid>()
            ?? throw new InvalidOperationException("table-grid-must-declare-columns");
        grid.RemoveAllChildren<GridColumn>();
        foreach (var width in widths) grid.Append(new GridColumn { Width = width.ToString() });
    }

    private static void ReshapeTargetRow(
        TableRow row,
        IReadOnlyList<TargetColumnReshape> columns,
        int newGridCount)
    {
        var reshaped = new List<(TableCell Cell, int Start, int Span)>();
        foreach (var cell in Cells(row))
        {
            var oldPositions = Enumerable.Range(cell.Start, cell.Span).ToArray();
            var selected = columns.Where(column => oldPositions.Any(position =>
                position >= column.OldStart && position < column.OldStart + column.OldSpan)).ToArray();
            var selectedOld = selected.SelectMany(column => Enumerable.Range(
                Math.Max(cell.Start, column.OldStart),
                Math.Min(cell.Start + cell.Span, column.OldStart + column.OldSpan)
                    - Math.Max(cell.Start, column.OldStart))).Order().ToArray();
            if (!selectedOld.SequenceEqual(oldPositions))
                throw new InvalidOperationException("target-cell-must-align-with-header-columns");
            var newPositions = selected.SelectMany(column => MapInterval(
                Math.Max(cell.Start, column.OldStart),
                Math.Min(cell.Start + cell.Span, column.OldStart + column.OldSpan),
                column.OldStart, column.OldSpan, column.NewStart, column.NewSpan)).Order().ToArray();
            if (!newPositions.SequenceEqual(Enumerable.Range(newPositions[0], newPositions.Length)))
                throw new InvalidOperationException("target-cell-maps-to-noncontiguous-columns");
            reshaped.Add((cell.Cell, newPositions[0], newPositions.Length));
        }
        var occupied = reshaped.SelectMany(item => Enumerable.Range(item.Start, item.Span)).Order().ToArray();
        if (occupied.Length == 0 || !occupied.SequenceEqual(Enumerable.Range(occupied[0], occupied.Length)))
            throw new InvalidOperationException("target-row-maps-to-noncontiguous-grid");
        foreach (var item in reshaped) SetGridSpan(item.Cell, item.Span);
        SetGridOmissions(row, occupied[0], newGridCount - occupied[^1] - 1);
    }

    private static int RowGridBefore(TableRow row)
        => row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;

    private static int RowGridAfter(TableRow row)
        => row.TableRowProperties?.GetFirstChild<GridAfter>()?.Val?.Value ?? 0;

    private static void SetGridOmissions(TableRow row, int before, int after)
    {
        if (before < 0 || after < 0) throw new InvalidOperationException("row-grid-omission-invalid");
        var properties = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<GridBefore>();
        properties.RemoveAllChildren<GridAfter>();
        if (before > 0 && !properties.AddChild(new GridBefore { Val = before }, true))
            throw new InvalidOperationException("row-grid-before-not-supported");
        if (after > 0 && !properties.AddChild(new GridAfter { Val = after }, true))
            throw new InvalidOperationException("row-grid-after-not-supported");
    }

    private static IEnumerable<int> MapInterval(
        int start,
        int end,
        int sourceStart,
        int sourceSpan,
        int targetStart,
        int targetSpan)
    {
        var mappedStart = targetStart + (start - sourceStart) * targetSpan / sourceSpan;
        var mappedEnd = targetStart + (end - sourceStart) * targetSpan / sourceSpan;
        return Enumerable.Range(mappedStart, mappedEnd - mappedStart);
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
                var cell = row.Cells.SingleOrDefault(item =>
                    column >= item.TargetColumn && column < item.TargetColumn + item.TargetSpan);
                if (cell is null)
                {
                    active = false;
                    continue;
                }
                if (cell.VerticalMerge is null) active = false;
                else if (cell.VerticalMerge.Val?.Value == MergedCellValues.Restart) active = true;
                else if (!active) throw new InvalidOperationException($"{name}-vertical-merge-continue-without-restart");
            }
        }
    }

    private static MergedCellValues? MergeValue(TableRow row, int column)
    {
        var slot = Cells(row).SingleOrDefault(cell => column >= cell.Start && column < cell.Start + cell.Span);
        var merge = slot?.Cell.TableCellProperties?.VerticalMerge;
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
    private sealed record ColumnMapping(int SourceStart, int SourceSpan,
        int OldTargetStart, int OldTargetSpan, int TargetStart, int TargetSpan, string Content);
    private sealed record TargetColumnReshape(int OldStart, int OldSpan, int NewStart, int NewSpan);
    private sealed record PreparedCell(int SourceColumn, int SourceSpan, int TargetColumn,
        IReadOnlyList<Paragraph> Paragraphs, VerticalMerge? VerticalMerge, int TargetSpan);
    private sealed record PreparedRow(IReadOnlyList<PreparedCell> Cells, int GridBefore, int GridAfter);
    private sealed record PreparedChange(int Index, ResolvedDocxReference TargetTable,
        IReadOnlyList<string> TargetRowPaths, string SourcePath, string SourceRevision,
        IReadOnlyList<long> TargetGridWidths, IReadOnlyList<TargetColumnReshape> TargetColumns,
        IReadOnlyList<PreparedRow> Rows);
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
