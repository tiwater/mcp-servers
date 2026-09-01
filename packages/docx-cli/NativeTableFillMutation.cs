using System.Text.Json;

namespace Dockit.Docx;

public static class NativeTableFillMutation
{
    public const string Command = "docx_fill_table_from_tables";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<FillTableFromTablesRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("fill-table-from-tables-request-invalid");
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

    public static FillTableFromTablesReceipt Apply(FillTableFromTablesRequest request)
    {
        if (request.Sources.Count == 0)
            throw new InvalidOperationException("sources-must-not-be-empty");

        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var target = Observation.ReadTable(paths.Input, request.Table);
        var prototype = target.Rows.SingleOrDefault(row => row.Address == request.PrototypeRow)
            ?? throw new InvalidOperationException("prototypeRow-not-found-in-target-table");
        var targetSlots = TargetSlots(prototype, target.ColumnCount);
        var values = new List<PreparedRecord>();
        var sourceReadbacks = new List<FillTableSourceReadback>();
        for (var sourceIndex = 0; sourceIndex < request.Sources.Count; sourceIndex++)
        {
            var requestSource = request.Sources[sourceIndex];
            var sourceInput = Path.GetFullPath(requestSource.Input);
            if (!File.Exists(sourceInput) || Directory.Exists(sourceInput))
                throw new InvalidOperationException($"sources[{sourceIndex}].input-file-not-found");
            var source = Observation.ReadTable(sourceInput, requestSource.Table);
            var sourceRows = SelectRows(source.Rows, requestSource.Rows, $"sources[{sourceIndex}].rows");
            var sourceRecordColumn = ColumnIndex(source.GridColumns, requestSource.RecordColumn,
                $"sources[{sourceIndex}].recordColumn");
            var mappings = PrepareMappings(source, target, targetSlots, requestSource.ColumnMappings, sourceIndex);
            var records = GroupRecords(sourceRows, sourceRecordColumn);
            if (records.Count == 0)
                throw new InvalidOperationException($"sources[{sourceIndex}].rows-must-contain-a-record");
            var ownerScope = "source-" + sourceIndex;
            values.AddRange(records.Select((record, recordIndex) =>
                BuildRecord(record, mappings, targetSlots.Count,
                    ownerScope + "-record-" + recordIndex)));
            sourceReadbacks.Add(new FillTableSourceReadback(
                sourceInput, requestSource.Table, requestSource.Rows, records.Count));
        }
        var rows = BuildTargetRows(values, request.PrototypeRow, targetSlots);

        var identity = Guid.NewGuid().ToString("N");
        var internalReceipt = paths.Receipt + ".table-body-" + identity;
        var temporaryOutput = paths.Output + ".projection-" + identity;
        try
        {
            var bodyReceipt = NativeTableBodyMutation.Apply(new SetTableBodyRequest(
                paths.Input,
                request.Table,
                request.ExistingRows,
                target.GridColumns.Select((column, index) =>
                    new SetTableBodyColumn("c" + index, column.Address)).ToArray(),
                rows,
                temporaryOutput,
                internalReceipt));
            var receipt = new FillTableFromTablesReceipt(
                "tiwater.docx-fill-table-from-tables-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                sourceReadbacks,
                request.Table,
                values.Count,
                bodyReceipt.Rows,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            NativeMutationSupport.Commit(temporaryOutput, paths);
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryOutput, paths);
            throw;
        }
        finally
        {
            NativeMutationSupport.Cleanup(internalReceipt);
        }
    }

    private static IReadOnlyList<PreparedMapping> PrepareMappings(
        DocxTableReadResult source,
        DocxTableReadResult target,
        IReadOnlyList<TargetSlot> targetSlots,
        IReadOnlyList<FillTableColumnMapping> requested,
        int sourceIndex)
    {
        if (requested.Count == 0)
            throw new InvalidOperationException($"sources[{sourceIndex}].columnMappings-must-not-be-empty");
        var mappings = requested.Select((mapping, mappingIndex) =>
        {
            var prefix = $"sources[{sourceIndex}].columnMappings[{mappingIndex}]";
            if (mapping.SourceColumns.Count == 0)
                throw new InvalidOperationException($"{prefix}.sourceColumns-must-not-be-empty");
            if (mapping.TargetColumns.Count == 0)
                throw new InvalidOperationException($"{prefix}.targetColumns-must-not-be-empty");
            var sources = mapping.SourceColumns
                .Select(column => ColumnIndex(source.GridColumns, column, prefix + ".sourceColumns"))
                .ToArray();
            if (sources.Distinct().Count() != sources.Length
                || !sources.SequenceEqual(sources.Order())
                || !sources.SequenceEqual(Enumerable.Range(sources[0], sources.Length)))
                throw new InvalidOperationException($"{prefix}.sourceColumns-must-be-unique-contiguous-in-native-order");
            var targets = mapping.TargetColumns
                .Select(column => ColumnIndex(target.GridColumns, column, prefix + ".targetColumns"))
                .ToArray();
            if (!targets.SequenceEqual(targets.Order())
                || !targets.SequenceEqual(Enumerable.Range(targets[0], targets.Length)))
                throw new InvalidOperationException($"{prefix}.targetColumns-must-be-contiguous-in-native-order");
            var slots = targetSlots.Select((slot, index) => (slot, index))
                .Where(item => targets.Contains(item.slot.Start))
                .ToArray();
            var covered = slots.SelectMany(item => Enumerable.Range(item.slot.Start, item.slot.Span)).ToArray();
            if (!covered.SequenceEqual(targets))
                throw new InvalidOperationException($"{prefix}.targetColumns-must-cover-whole-prototype-cells");
            var targetSlotIndexes = slots.Select(item => item.index).ToArray();
            if (sources.Length > 1 && targetSlotIndexes.Length > 1)
                throw new InvalidOperationException($"{prefix}-cannot-compose-and-spread");
            return new PreparedMapping(sources, targetSlotIndexes, mapping.JoinWith ?? "\n");
        }).ToArray();
        var mappedSlots = mappings.SelectMany(mapping => mapping.TargetSlots).ToArray();
        if (mappedSlots.Distinct().Count() != mappedSlots.Length)
            throw new InvalidOperationException($"sources[{sourceIndex}].target-cells-must-be-unique");
        if (!mappedSlots.Order().SequenceEqual(Enumerable.Range(0, targetSlots.Count)))
            throw new InvalidOperationException($"sources[{sourceIndex}].target-columns-must-cover-table-grid");
        return mappings;
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
        int targetSlotCount,
        string ownerScope)
    {
        var result = Enumerable.Range(0, targetSlotCount)
            .Select(_ => new PreparedValue(string.Empty, null)).ToArray();
        foreach (var mapping in mappings)
        {
            if (mapping.TargetSlots.Count == 1)
            {
                var components = new List<DocxTableReadCell>();
                foreach (var sourceColumn in mapping.SourceColumns)
                {
                    var scalarCells = LogicalCells(rows, sourceColumn);
                    var distinctTexts = scalarCells.Select(cell => cell.LogicalText)
                        .Distinct(StringComparer.Ordinal).ToArray();
                    if (distinctTexts.Length != 1)
                        throw new InvalidOperationException("scalar-source-values-differ-within-record");
                    components.Add(scalarCells[0]);
                }
                var unique = components.GroupAdjacentBy(LogicalOwner).Select(group => group.First()).ToArray();
                var text = string.Join(mapping.JoinWith,
                    unique.Select(cell => cell.LogicalText).Where(value => !string.IsNullOrEmpty(value)));
                var owner = components.All(cell => LogicalCells(rows, cell.GridColumnStart).Length == 1)
                    ? string.Join("\u001f", unique.Select(cell => OwnerKey(ownerScope, LogicalOwner(cell))))
                    : null;
                result[mapping.TargetSlots[0]] = new PreparedValue(text, owner);
                continue;
            }

            var logicalCells = LogicalCells(rows, mapping.SourceColumns.Single());
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

    private static string OwnerKey(string scope, DocxObjectAddress owner)
        => scope + "\u001d" + owner.Part + "\u001e" + owner.Path;

    private static DocxTableReadCell[] LogicalCells(IReadOnlyList<DocxTableReadRow> rows, int sourceColumn)
        => rows.Select(row => CellAt(row, sourceColumn))
            .GroupAdjacentBy(LogicalOwner)
            .Select(group => group.First())
            .ToArray();

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

    private sealed record PreparedMapping(
        IReadOnlyList<int> SourceColumns,
        IReadOnlyList<int> TargetSlots,
        string JoinWith);
    private sealed record TargetSlot(int Start, int Span);
    private sealed record PreparedValue(string Text, string? Owner);
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
    IReadOnlyList<DocxObjectAddress> SourceColumns,
    IReadOnlyList<DocxObjectAddress> TargetColumns,
    string? JoinWith = null);

public sealed record FillTableSource(
    string Input,
    DocxObjectAddress Table,
    SetTableBodyRowRange Rows,
    DocxObjectAddress RecordColumn,
    IReadOnlyList<FillTableColumnMapping> ColumnMappings);

public sealed record FillTableFromTablesRequest(
    string Input,
    DocxObjectAddress Table,
    SetTableBodyRowRange ExistingRows,
    DocxObjectAddress PrototypeRow,
    IReadOnlyList<FillTableSource> Sources,
    string Output,
    string ReceiptOutput);

public sealed record FillTableSourceReadback(
    string Input,
    DocxObjectAddress Table,
    SetTableBodyRowRange Rows,
    int RecordCount);

public sealed record FillTableFromTablesReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<FillTableSourceReadback> Sources,
    DocxObjectAddress Table,
    int RecordCount,
    IReadOnlyList<SetTableBodyRowReadback> Rows,
    string Output);
