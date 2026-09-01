using System.Text.Json;

namespace Dockit.Docx;

public static class NativeTableFillMutation
{
    public const string Command = "docx_fill_table_from_table";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<FillTableFromTableRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("fill-table-from-table-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = 1, appliedCount = 1 },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static FillTableFromTableReceipt Apply(FillTableFromTableRequest request)
    {
        if (request.ColumnMappings.Count == 0)
            throw new InvalidOperationException("columnMappings-must-not-be-empty");

        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var sourceInput = Path.GetFullPath(request.SourceInput);
        if (!File.Exists(sourceInput) || Directory.Exists(sourceInput))
            throw new InvalidOperationException("sourceInput-file-not-found");

        var source = Observation.ReadTable(sourceInput, request.SourceTable);
        var target = Observation.ReadTable(paths.Input, request.Table);
        var sourceRows = SelectRows(source.Rows, request.SourceRows, "sourceRows");
        RequireClosedVerticalMerges(source.Rows, sourceRows, "sourceRows");

        var sourceRecordColumn = ColumnIndex(source.GridColumns, request.SourceRecordColumn,
            "sourceRecordColumn");
        var prototype = target.Rows.SingleOrDefault(row => row.Address == request.PrototypeRow)
            ?? throw new InvalidOperationException("prototypeRow-not-found-in-target-table");
        var targetSlots = TargetSlots(prototype, target.ColumnCount);

        var mappings = request.ColumnMappings.Select((mapping, mappingIndex) =>
        {
            if (mapping.TargetColumns.Count == 0)
                throw new InvalidOperationException($"columnMappings[{mappingIndex}].targetColumns-must-not-be-empty");
            var targets = mapping.TargetColumns
                .Select(column => ColumnIndex(target.GridColumns, column, "columnMappings.targetColumns"))
                .ToArray();
            if (!targets.SequenceEqual(targets.Order())
                || !targets.SequenceEqual(Enumerable.Range(targets[0], targets.Length)))
                throw new InvalidOperationException($"columnMappings[{mappingIndex}].targetColumns-must-be-contiguous-in-native-order");
            var slots = targetSlots.Select((slot, index) => (slot, index))
                .Where(item => targets.Contains(item.slot.Start))
                .ToArray();
            var covered = slots.SelectMany(item => Enumerable.Range(item.slot.Start, item.slot.Span)).ToArray();
            if (!covered.SequenceEqual(targets))
                throw new InvalidOperationException($"columnMappings[{mappingIndex}].targetColumns-must-cover-whole-prototype-cells");
            return new PreparedMapping(
                ColumnIndex(source.GridColumns, mapping.SourceColumn, "columnMappings.sourceColumn"),
                slots.Select(item => item.index).ToArray());
        }).ToArray();
        var mappedSlots = mappings.SelectMany(mapping => mapping.TargetSlots).ToArray();
        if (mappedSlots.Distinct().Count() != mappedSlots.Length)
            throw new InvalidOperationException("target-cells-must-be-unique");
        if (!mappedSlots.Order().SequenceEqual(Enumerable.Range(0, targetSlots.Count)))
            throw new InvalidOperationException("target-columns-must-cover-table-grid");

        var records = GroupRecords(sourceRows, sourceRecordColumn);
        if (records.Count == 0) throw new InvalidOperationException("sourceRows-must-contain-a-record");
        var values = records.Select(record => BuildRecord(record, mappings, targetSlots.Count)).ToArray();
        var rows = BuildTargetRows(values, request.PrototypeRow, targetSlots);

        var internalReceipt = paths.Receipt + ".table-body-" + Guid.NewGuid().ToString("N");
        try
        {
            var bodyReceipt = NativeTableBodyMutation.Apply(new SetTableBodyRequest(
                paths.Input,
                request.Table,
                request.ExistingRows,
                target.GridColumns.Select((column, index) =>
                    new SetTableBodyColumn("c" + index, column.Address)).ToArray(),
                rows,
                paths.Output,
                internalReceipt));
            var receipt = new FillTableFromTableReceipt(
                "tiwater.docx-fill-table-from-table-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                request.SourceTable,
                request.Table,
                records.Count,
                bodyReceipt.Rows,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        finally
        {
            NativeMutationSupport.Cleanup(internalReceipt);
        }
    }

    private static IReadOnlyList<DocxTableReadRow> SelectRows(
        IReadOnlyList<DocxTableReadRow> rows,
        SetTableBodyRowRange range,
        string name)
    {
        var first = IndexOf(rows, range.First);
        var last = IndexOf(rows, range.Last);
        if (first < 0 || last < first)
            throw new InvalidOperationException($"{name}-must-be-one-forward-range");
        return rows.Skip(first).Take(last - first + 1).ToArray();
    }

    private static int IndexOf(IReadOnlyList<DocxTableReadRow> rows, DocxObjectAddress address)
    {
        for (var index = 0; index < rows.Count; index++)
            if (rows[index].Address == address) return index;
        return -1;
    }

    private static int ColumnIndex(
        IReadOnlyList<DocxTableReadGridColumn> columns,
        DocxObjectAddress address,
        string name)
    {
        for (var index = 0; index < columns.Count; index++)
            if (columns[index].Address == address) return index;
        throw new InvalidOperationException(name + "-not-found-in-table-grid");
    }

    private static void RequireClosedVerticalMerges(
        IReadOnlyList<DocxTableReadRow> allRows,
        IReadOnlyList<DocxTableReadRow> selectedRows,
        string name)
    {
        var selected = selectedRows.Select(row => row.Address).ToHashSet();
        var groups = allRows
            .SelectMany(row => row.Cells
                .Where(cell => cell.VerticalMergeOwner is not null)
                .Select(cell => (row.Address, Owner: cell.VerticalMergeOwner!)))
            .GroupBy(item => item.Owner);
        foreach (var group in groups)
        {
            var rows = group.Select(item => item.Address).Distinct().ToArray();
            var count = rows.Count(selected.Contains);
            if (count > 0 && count != rows.Length)
                throw new InvalidOperationException(name + "-split-vertical-merge");
        }
    }

    private static IReadOnlyList<IReadOnlyList<DocxTableReadRow>> GroupRecords(
        IReadOnlyList<DocxTableReadRow> rows,
        int recordColumn)
    {
        var result = new List<IReadOnlyList<DocxTableReadRow>>();
        var current = new List<DocxTableReadRow>();
        DocxObjectAddress? currentOwner = null;
        foreach (var row in rows)
        {
            var cell = CellAt(row, recordColumn);
            var owner = LogicalOwner(cell);
            if (current.Count > 0 && owner != currentOwner)
            {
                result.Add(current.ToArray());
                current.Clear();
            }
            currentOwner = owner;
            current.Add(row);
        }
        if (current.Count > 0) result.Add(current.ToArray());
        return result;
    }

    private static PreparedRecord BuildRecord(
        IReadOnlyList<DocxTableReadRow> rows,
        IReadOnlyList<PreparedMapping> mappings,
        int targetSlotCount)
    {
        var result = Enumerable.Range(0, targetSlotCount)
            .Select(_ => new PreparedValue(string.Empty, null)).ToArray();
        foreach (var mapping in mappings)
        {
            var logicalCells = rows.Select(row => CellAt(row, mapping.SourceColumn))
                .GroupAdjacentBy(LogicalOwner)
                .Select(group => group.First())
                .ToArray();
            if (mapping.TargetSlots.Count == 1)
            {
                var distinctTexts = logicalCells.Select(cell => cell.LogicalText)
                    .Distinct(StringComparer.Ordinal).ToArray();
                if (distinctTexts.Length != 1)
                    throw new InvalidOperationException("scalar-source-values-differ-within-record");
                var owner = logicalCells.Length == 1 ? LogicalOwner(logicalCells[0]) : null;
                result[mapping.TargetSlots[0]] = new PreparedValue(distinctTexts[0], owner);
                continue;
            }

            if (logicalCells.Length > mapping.TargetSlots.Count)
                throw new InvalidOperationException("source-record-has-more-values-than-target-columns");
            for (var index = 0; index < logicalCells.Length; index++)
                result[mapping.TargetSlots[index]] = new PreparedValue(logicalCells[index].LogicalText, null);
        }
        return new PreparedRecord(result);
    }

    private static IReadOnlyList<SetTableBodyRow> BuildTargetRows(
        IReadOnlyList<PreparedRecord> records,
        DocxObjectAddress prototypeRow,
        IReadOnlyList<TargetSlot> targetSlots)
    {
        var rows = new List<SetTableBodyRow>(records.Count);
        for (var rowIndex = 0; rowIndex < records.Count; rowIndex++)
        {
            var cells = new List<SetTableBodyCell>();
            for (var slotIndex = 0; slotIndex < targetSlots.Count; slotIndex++)
            {
                var value = records[rowIndex].Values[slotIndex];
                if (value.Owner is not null && rowIndex > 0
                    && records[rowIndex - 1].Values[slotIndex].Owner == value.Owner) continue;
                var rowSpan = 1;
                if (value.Owner is not null)
                    while (rowIndex + rowSpan < records.Count
                           && records[rowIndex + rowSpan].Values[slotIndex].Owner == value.Owner) rowSpan++;
                var slot = targetSlots[slotIndex];
                cells.Add(new SetTableBodyCell(
                    Enumerable.Range(slot.Start, slot.Span).Select(column => "c" + column).ToArray(),
                    value.Text,
                    rowSpan > 1 ? rowSpan : null));
            }
            rows.Add(new SetTableBodyRow(prototypeRow, cells));
        }
        return rows;
    }

    private static DocxTableReadCell CellAt(DocxTableReadRow row, int gridColumn)
        => row.Cells.SingleOrDefault(cell =>
               cell.GridColumnStart <= gridColumn && gridColumn < cell.GridColumnStart + cell.GridSpan)
           ?? throw new InvalidOperationException("source-row-does-not-cover-selected-column");

    private static DocxObjectAddress LogicalOwner(DocxTableReadCell cell)
        => cell.VerticalMergeOwner ?? cell.Address;

    private static IReadOnlyList<TargetSlot> TargetSlots(DocxTableReadRow prototype, int columnCount)
    {
        if (prototype.GridBefore != 0 || prototype.GridAfter != 0)
            throw new InvalidOperationException("prototypeRow-must-cover-target-grid");
        var slots = prototype.Cells.Select(cell => new TargetSlot(cell.GridColumnStart, cell.GridSpan)).ToArray();
        var covered = slots.SelectMany(slot => Enumerable.Range(slot.Start, slot.Span)).ToArray();
        if (!covered.SequenceEqual(Enumerable.Range(0, columnCount)))
            throw new InvalidOperationException("prototypeRow-must-cover-target-grid");
        return slots;
    }

    private sealed record PreparedMapping(int SourceColumn, IReadOnlyList<int> TargetSlots);
    private sealed record TargetSlot(int Start, int Span);
    private sealed record PreparedValue(string Text, DocxObjectAddress? Owner);
    private sealed record PreparedRecord(IReadOnlyList<PreparedValue> Values);
}

internal static class AdjacentGroupingExtensions
{
    public static IEnumerable<IReadOnlyList<T>> GroupAdjacentBy<T, TKey>(
        this IEnumerable<T> source,
        Func<T, TKey> keySelector)
    {
        var group = new List<T>();
        var comparer = EqualityComparer<TKey>.Default;
        TKey? currentKey = default;
        var hasKey = false;
        foreach (var item in source)
        {
            var key = keySelector(item);
            if (hasKey && !comparer.Equals(currentKey!, key))
            {
                yield return group.ToArray();
                group.Clear();
            }
            currentKey = key;
            hasKey = true;
            group.Add(item);
        }
        if (group.Count > 0) yield return group.ToArray();
    }
}

public sealed record FillTableColumnMapping(
    DocxObjectAddress SourceColumn,
    IReadOnlyList<DocxObjectAddress> TargetColumns);

public sealed record FillTableFromTableRequest(
    string Input,
    DocxObjectAddress Table,
    SetTableBodyRowRange ExistingRows,
    DocxObjectAddress PrototypeRow,
    string SourceInput,
    DocxObjectAddress SourceTable,
    SetTableBodyRowRange SourceRows,
    DocxObjectAddress SourceRecordColumn,
    IReadOnlyList<FillTableColumnMapping> ColumnMappings,
    string Output,
    string ReceiptOutput);

public sealed record FillTableFromTableReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    DocxObjectAddress SourceTable,
    DocxObjectAddress Table,
    int RecordCount,
    IReadOnlyList<SetTableBodyRowReadback> Rows,
    string Output);
