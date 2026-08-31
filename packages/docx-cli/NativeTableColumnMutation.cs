using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTableColumnMutation
{
    public const string InsertCommand = "docx_insert_table_columns";
    public const string DeleteCommand = "docx_delete_table_columns";
    private const string MainStory = "/word/document.xml";

    public static int Run(string command, string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        var json = File.ReadAllText(args[0]);
        TableColumnMutationReceipt receipt;
        int operationCount;
        if (command == InsertCommand)
        {
            var request = JsonSerializer.Deserialize<InsertTableColumnsRequest>(json, Json.Options)
                ?? throw new InvalidOperationException("insert-table-columns-request-invalid");
            receipt = Insert(request);
            operationCount = request.Changes.Count;
        }
        else if (command == DeleteCommand)
        {
            var request = JsonSerializer.Deserialize<DeleteTableColumnsRequest>(json, Json.Options)
                ?? throw new InvalidOperationException("delete-table-columns-request-invalid");
            receipt = Delete(request);
            operationCount = request.Changes.Count;
        }
        else throw new InvalidOperationException("table-column-command-invalid");
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = command,
            receipt = NativeMutationSupport.Describe(receipt.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static TableColumnMutationReceipt Insert(InsertTableColumnsRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var refs = request.Changes.SelectMany(change => new[] { change.Table, change.SourceColumn }
                .Concat(change.Before is null ? [] : [change.Before]))
            .ToArray();
        var resolved = Observation.ResolveAddresses(paths.Input, refs, "changes.addresses");
        var prepared = new List<PreparedInsert>();
        var offset = 0;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            foreach (var change in request.Changes)
            {
                var tableRef = resolved[offset++];
                var sourceRef = resolved[offset++];
                var beforeRef = change.Before is null ? null : resolved[offset++];
                if (tableRef.Kind != "table" || sourceRef.Kind != "gridColumn"
                    || beforeRef is not null && beforeRef.Kind != "gridColumn"
                    || tableRef.StoryPart != MainStory || sourceRef.StoryPart != MainStory
                    || beforeRef is not null && beforeRef.StoryPart != MainStory)
                    throw new InvalidOperationException("table-column-address-kind-invalid");
                var table = Observation.ResolveNativePath(input, tableRef.StoryPart, tableRef.NativePath) as Table
                    ?? throw new InvalidOperationException("table-address-not-found");
                var source = Observation.ResolveNativePath(input, sourceRef.StoryPart, sourceRef.NativePath) as GridColumn
                    ?? throw new InvalidOperationException("source-column-address-not-found");
                var before = beforeRef is null ? null
                    : Observation.ResolveNativePath(input, beforeRef.StoryPart, beforeRef.NativePath) as GridColumn
                      ?? throw new InvalidOperationException("before-column-address-not-found");
                RequireSameGrid(table, source, before);
                var columns = table.GetFirstChild<TableGrid>()!.Elements<GridColumn>().ToList();
                var sourceIndex = columns.IndexOf(source);
                var insertionIndex = before is null ? columns.Count : columns.IndexOf(before);
                var prototypes = table.Elements<TableRow>()
                    .Select(row => PrototypeFor(row, sourceIndex, insertionIndex))
                    .ToArray();
                prepared.Add(new PreparedInsert(tableRef, sourceRef, beforeRef, change.Repeat ?? 1, prototypes));
            }
        }
        if (prepared.Any(change => change.Repeat < 1)) throw new InvalidOperationException("repeat-must-be-positive");
        return Apply(paths, InsertCommand, document =>
        {
            var active = prepared.Select(change =>
            {
                var table = (Table)Observation.ResolveNativePath(document, change.Table.StoryPart, change.Table.NativePath);
                var source = (GridColumn)Observation.ResolveNativePath(document, change.Source.StoryPart, change.Source.NativePath);
                var before = change.Before is null ? null
                    : (GridColumn)Observation.ResolveNativePath(document, change.Before.StoryPart, change.Before.NativePath);
                RequireSameGrid(table, source, before);
                return new ActiveInsert(table, source, before, change.Repeat, change.Prototypes);
            }).ToArray();
            var changes = new List<TableColumnMutationReadback>();
            foreach (var change in active)
            {
                var grid = change.Table.GetFirstChild<TableGrid>() ?? throw new InvalidOperationException("table-grid-not-found");
                for (var repeat = 0; repeat < change.Repeat; repeat++)
                {
                    var sourceIndex = grid.Elements<GridColumn>().ToList().IndexOf(change.Source);
                    var insertionIndex = change.Before is null
                        ? grid.Elements<GridColumn>().Count()
                        : grid.Elements<GridColumn>().ToList().IndexOf(change.Before);
                    InsertOne(change.Table, grid, change.Source, sourceIndex, insertionIndex, change.Prototypes);
                }
                changes.Add(Readback(change.Table));
            }
            return changes;
        });
    }

    public static TableColumnMutationReceipt Delete(DeleteTableColumnsRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var prepared = new List<PreparedDelete>();
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            foreach (var (change, changeIndex) in request.Changes.Select((value, index) => (value, index)))
            {
                if (change.Columns.Count == 0) throw new InvalidOperationException($"changes[{changeIndex}].columns-must-not-be-empty");
                if (change.Columns.Distinct().Count() != change.Columns.Count)
                    throw new InvalidOperationException($"changes[{changeIndex}].columns-must-not-contain-duplicates");
                var refs = Observation.ResolveAddresses(paths.Input, new[] { change.Table }.Concat(change.Columns).ToArray(), $"changes[{changeIndex}]");
                var tableRef = refs[0];
                if (tableRef.Kind != "table" || tableRef.StoryPart != MainStory
                    || refs.Skip(1).Any(item => item.Kind != "gridColumn" || item.StoryPart != MainStory))
                    throw new InvalidOperationException("table-column-address-kind-invalid");
                var table = (Table)Observation.ResolveNativePath(input, tableRef.StoryPart, tableRef.NativePath);
                var grid = table.GetFirstChild<TableGrid>() ?? throw new InvalidOperationException("table-grid-not-found");
                var columns = refs.Skip(1).Select(item =>
                    (GridColumn)Observation.ResolveNativePath(input, item.StoryPart, item.NativePath)).ToArray();
                if (columns.Any(column => !ReferenceEquals(column.Parent, grid)))
                    throw new InvalidOperationException("columns-must-belong-to-selected-table");
                if (columns.Length >= grid.Elements<GridColumn>().Count())
                    throw new InvalidOperationException("delete-must-not-remove-all-table-columns");
                prepared.Add(new PreparedDelete(tableRef, refs.Skip(1).ToArray()));
            }
        }
        return Apply(paths, DeleteCommand, document =>
        {
            var active = prepared.Select(change =>
            {
                var table = (Table)Observation.ResolveNativePath(document, change.Table.StoryPart, change.Table.NativePath);
                var grid = table.GetFirstChild<TableGrid>() ?? throw new InvalidOperationException("table-grid-not-found");
                var columns = change.Columns.Select(item =>
                    (GridColumn)Observation.ResolveNativePath(document, item.StoryPart, item.NativePath)).ToArray();
                if (columns.Any(column => !ReferenceEquals(column.Parent, grid)))
                    throw new InvalidOperationException("columns-must-belong-to-selected-table");
                return new ActiveDelete(table, columns);
            }).ToArray();
            var changes = new List<TableColumnMutationReadback>();
            foreach (var change in active)
            {
                var grid = change.Table.GetFirstChild<TableGrid>() ?? throw new InvalidOperationException("table-grid-not-found");
                foreach (var column in change.Columns.OrderByDescending(item => grid.Elements<GridColumn>().ToList().IndexOf(item)))
                    DeleteOne(change.Table, grid, column);
                changes.Add(Readback(change.Table));
            }
            return changes;
        });
    }

    private static TableColumnMutationReceipt Apply(
        NativeMutationSupport.PathsResult paths,
        string command,
        Func<WordprocessingDocument, IReadOnlyList<TableColumnMutationReadback>> mutate)
    {
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
        var temporary = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(paths.Input, temporary, false);
            IReadOnlyList<TableColumnMutationReadback> changes;
            using (var output = WordprocessingDocument.Open(temporary, true))
            {
                changes = mutate(output);
                output.MainDocumentPart?.Document?.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporary, paths);
            var receipt = new TableColumnMutationReceipt(
                command == InsertCommand ? "tiwater.docx-insert-table-columns-receipt/v1" : "tiwater.docx-delete-table-columns-receipt/v1",
                "tiwater.docx.cli", RuntimeIdentity.Version, changes, paths.Output, paths.Receipt);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporary, paths);
            throw;
        }
    }

    private static void InsertOne(
        Table table,
        TableGrid grid,
        GridColumn source,
        int sourceIndex,
        int insertionIndex,
        IReadOnlyList<TableCell> prototypes)
    {
        var clone = (GridColumn)source.CloneNode(true);
        var before = grid.Elements<GridColumn>().ElementAtOrDefault(insertionIndex);
        if (before is null) grid.Append(clone); else grid.InsertBefore(clone, before);
        var rows = table.Elements<TableRow>().ToArray();
        if (rows.Length != prototypes.Count) throw new InvalidOperationException("table-row-set-changed-during-column-insert");
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = rows[rowIndex];
            var sourcePosition = CellAt(row, sourceIndex);
            var beforeCount = RowOffset(row.TableRowProperties, "gridBefore");
            var afterCount = RowOffset(row.TableRowProperties, "gridAfter");
            var oldColumnCount = grid.Elements<GridColumn>().Count() - 1;
            if (insertionIndex < beforeCount
                || insertionIndex == beforeCount && sourcePosition is null && sourceIndex < beforeCount)
            {
                SetRowOffset(row, "gridBefore", beforeCount + 1);
                continue;
            }
            if (afterCount > 0 && insertionIndex >= oldColumnCount - afterCount)
            {
                SetRowOffset(row, "gridAfter", afterCount + 1);
                continue;
            }
            var insertionCell = CellAt(row, insertionIndex);
            var sourceEndsAtInsertion = sourcePosition is not null
                && sourcePosition.Start + sourcePosition.Span == insertionIndex;
            var sourceStartsAfterInsertion = sourcePosition?.Start == insertionIndex;
            if (insertionCell is not null && insertionIndex > insertionCell.Start)
            {
                SetSpan(insertionCell.Cell, insertionCell.Span + 1);
                continue;
            }
            if (sourcePosition is not null && sourcePosition.Span > 1 && (sourceEndsAtInsertion || sourceStartsAfterInsertion))
            {
                SetSpan(sourcePosition.Cell, sourcePosition.Span + 1);
                continue;
            }
            var prototype = EmptyCell(prototypes[rowIndex]);
            SetSpan(prototype, 1);
            var target = row.Elements<TableCell>().FirstOrDefault(cell => Position(row, cell).Start >= insertionIndex);
            if (target is null) row.Append(prototype); else row.InsertBefore(prototype, target);
        }
        RefreshCellWidths(table);
    }

    private static void DeleteOne(Table table, TableGrid grid, GridColumn column)
    {
        var columns = grid.Elements<GridColumn>().ToList();
        var index = columns.IndexOf(column);
        if (index < 0) throw new InvalidOperationException("column-address-not-found-in-table");
        var oldColumnCount = columns.Count;
        foreach (var row in table.Elements<TableRow>())
        {
            var before = RowOffset(row.TableRowProperties, "gridBefore");
            var after = RowOffset(row.TableRowProperties, "gridAfter");
            if (index < before)
            {
                SetRowOffset(row, "gridBefore", before - 1);
                continue;
            }
            if (index >= oldColumnCount - after)
            {
                SetRowOffset(row, "gridAfter", after - 1);
                continue;
            }
            var position = CellAt(row, index) ?? throw new InvalidOperationException("column-not-represented-in-row");
            if (position.Span > 1) SetSpan(position.Cell, position.Span - 1);
            else position.Cell.Remove();
            if (!row.Elements<TableCell>().Any()) throw new InvalidOperationException("delete-left-row-without-cells");
        }
        column.Remove();
        RefreshCellWidths(table);
    }

    private static void RequireSameGrid(Table table, GridColumn source, GridColumn? before)
    {
        var grid = table.GetFirstChild<TableGrid>() ?? throw new InvalidOperationException("table-grid-not-found");
        if (!ReferenceEquals(source.Parent, grid) || before is not null && !ReferenceEquals(before.Parent, grid))
            throw new InvalidOperationException("columns-must-belong-to-selected-table");
    }

    private static TableCell EmptyCell(TableCell prototype)
    {
        var result = new TableCell();
        if (prototype.TableCellProperties is { } cellProperties)
            result.Append((TableCellProperties)cellProperties.CloneNode(true));
        var sourceParagraph = prototype.Elements<Paragraph>().FirstOrDefault();
        var paragraph = new Paragraph();
        if (sourceParagraph?.ParagraphProperties is { } paragraphProperties)
            paragraph.Append((ParagraphProperties)paragraphProperties.CloneNode(true));
        var sourceRunProperties = sourceParagraph?.Descendants<Run>().FirstOrDefault()?.RunProperties;
        if (sourceRunProperties is not null)
            paragraph.Append(new Run((RunProperties)sourceRunProperties.CloneNode(true)));
        result.Append(paragraph);
        return result;
    }

    private static TableCell PrototypeFor(TableRow row, int sourceIndex, int insertionIndex)
    {
        var positions = Positions(row);
        var prototype = CellAt(row, sourceIndex)?.Cell
            ?? positions.FirstOrDefault(position => position.Start >= insertionIndex)?.Cell
            ?? positions.LastOrDefault()?.Cell
            ?? throw new InvalidOperationException("table-row-has-no-cell-prototype");
        return (TableCell)prototype.CloneNode(true);
    }

    private static CellPosition? CellAt(TableRow row, int column)
        => Positions(row).SingleOrDefault(item => column >= item.Start && column < item.Start + item.Span);

    private static CellPosition Position(TableRow row, TableCell cell)
        => Positions(row).Single(item => ReferenceEquals(item.Cell, cell));

    private static IReadOnlyList<CellPosition> Positions(TableRow row)
    {
        var cursor = RowOffset(row.TableRowProperties, "gridBefore");
        return row.Elements<TableCell>().Select(cell =>
        {
            var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            var result = new CellPosition(cell, cursor, span);
            cursor += span;
            return result;
        }).ToArray();
    }

    private static int RowOffset(TableRowProperties? properties, string localName)
    {
        var value = properties?.ChildElements.FirstOrDefault(child => child.LocalName == localName)
            ?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return int.TryParse(value, out var result) ? result : 0;
    }

    private static void SetRowOffset(TableRow row, string localName, int value)
    {
        var properties = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
        var existing = properties.ChildElements.FirstOrDefault(child => child.LocalName == localName);
        if (value == 0)
        {
            existing?.Remove();
            return;
        }
        var element = existing ?? (localName == "gridBefore" ? new GridBefore() : new GridAfter());
        element.SetAttribute(new OpenXmlAttribute("w", "val", element.NamespaceUri, value.ToString()));
        if (existing is null) properties.Append(element);
    }

    private static void SetSpan(TableCell cell, int span)
    {
        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        properties.GridSpan = span == 1 ? null : new GridSpan { Val = span };
    }

    private static void RefreshCellWidths(Table table)
    {
        var widths = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
            .Select(column => int.TryParse(column.Width?.Value, out var width) ? width : 0).ToArray() ?? [];
        if (widths.Length == 0 || widths.Any(width => width <= 0)) return;
        foreach (var row in table.Elements<TableRow>())
        foreach (var position in Positions(row))
        {
            if (position.Start + position.Span > widths.Length) continue;
            var properties = position.Cell.TableCellProperties ?? position.Cell.PrependChild(new TableCellProperties());
            if (properties.TableCellWidth?.Type?.Value != TableWidthUnitValues.Dxa) continue;
            properties.TableCellWidth.Width = widths.Skip(position.Start).Take(position.Span).Sum().ToString();
        }
    }

    private static TableColumnMutationReadback Readback(Table table)
        => new(
            Observation.Address(MainStory, Observation.NativePathFor(table)),
            table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0,
            table.Elements<TableRow>().Count());

    private sealed record CellPosition(TableCell Cell, int Start, int Span);
    private sealed record PreparedInsert(
        ResolvedDocxAddress Table,
        ResolvedDocxAddress Source,
        ResolvedDocxAddress? Before,
        int Repeat,
        IReadOnlyList<TableCell> Prototypes);
    private sealed record PreparedDelete(ResolvedDocxAddress Table, IReadOnlyList<ResolvedDocxAddress> Columns);
    private sealed record ActiveInsert(
        Table Table,
        GridColumn Source,
        GridColumn? Before,
        int Repeat,
        IReadOnlyList<TableCell> Prototypes);
    private sealed record ActiveDelete(Table Table, IReadOnlyList<GridColumn> Columns);
}

public sealed record InsertTableColumnsChange(
    DocxObjectAddress Table,
    DocxObjectAddress SourceColumn,
    DocxObjectAddress? Before = null,
    int? Repeat = null);
public sealed record InsertTableColumnsRequest(
    string Input,
    IReadOnlyList<InsertTableColumnsChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record DeleteTableColumnsChange(
    DocxObjectAddress Table,
    IReadOnlyList<DocxObjectAddress> Columns);
public sealed record DeleteTableColumnsRequest(
    string Input,
    IReadOnlyList<DeleteTableColumnsChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record TableColumnMutationReadback(DocxObjectAddress Table, int ColumnCount, int RowCount);
public sealed record TableColumnMutationReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<TableColumnMutationReadback> Changes,
    string Output,
    string ReceiptOutput);
