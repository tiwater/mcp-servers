using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTableBodyMutation
{
    public const string Command = "docx_set_table_body";
    private const string MainStory = "/word/document.xml";
    private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetTableBodyRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-table-body-request-invalid");
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

    public static SetTableBodyReceipt Apply(SetTableBodyRequest request)
    {
        if (request.Columns.Count == 0) throw new InvalidOperationException("columns-must-not-be-empty");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var prepared = Prepare(paths.Input, request);
        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
            baseline = NativeMutationSupport.ValidationIssueCounts(input);

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                ApplyPrepared(output, prepared);
                output.MainDocumentPart?.Document?.Save();
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            var readback = ReadBack(temporaryPath, prepared);
            var receipt = new SetTableBodyReceipt(
                "tiwater.docx-set-table-body-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                prepared.Table.Address,
                readback,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            NativeMutationSupport.Commit(temporaryPath, paths);
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static PreparedTable Prepare(string input, SetTableBodyRequest request)
    {
        var tableRef = Observation.ResolveAddresses(input, [request.Table], "table").Single();
        var bodyRefs = Observation.ResolveAddresses(
            input, [request.ExistingRows.First, request.ExistingRows.Last], "existingRows");
        var columnRefs = Observation.ResolveAddresses(
            input, request.Columns.Select(column => column.GridColumn).ToArray(), "columns.gridColumn");
        var prototypeRefs = Observation.ResolveAddresses(
            input, request.Rows.Select(row => row.PrototypeRow).ToArray(), "rows.prototypeRow");
        if (tableRef.Kind != "table" || tableRef.StoryPart != MainStory
            || bodyRefs.Any(item => item.Kind != "row" || item.StoryPart != MainStory)
            || columnRefs.Any(item => item.Kind != "gridColumn" || item.StoryPart != MainStory)
            || prototypeRefs.Any(item => item.Kind != "row" || item.StoryPart != MainStory))
            throw new InvalidOperationException("set-table-body-address-kind-invalid");

        var ids = request.Columns.Select(column => column.Id).ToArray();
        if (ids.Any(string.IsNullOrWhiteSpace) || ids.Distinct(StringComparer.Ordinal).Count() != ids.Length)
            throw new InvalidOperationException("column-ids-must-be-nonempty-and-unique");

        using var document = WordprocessingDocument.Open(input, false);
        var table = Resolve<Table>(document, tableRef, "table");
        var rows = table.Elements<TableRow>().ToArray();
        var first = Resolve<TableRow>(document, bodyRefs[0], "existingRows.first");
        var last = Resolve<TableRow>(document, bodyRefs[1], "existingRows.last");
        var firstIndex = Array.IndexOf(rows, first);
        var lastIndex = Array.IndexOf(rows, last);
        if (firstIndex < 0 || lastIndex < firstIndex)
            throw new InvalidOperationException("existingRows-must-be-one-forward-range");
        var selectedRows = rows[firstIndex..(lastIndex + 1)];
        if (selectedRows.Length == rows.Length && request.Rows.Count == 0)
            throw new InvalidOperationException("table-must-retain-at-least-one-row");
        RequireClosedVerticalMerges(table, selectedRows);

        var grid = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToArray()
            ?? throw new InvalidOperationException("table-grid-not-found");
        var selectedColumns = columnRefs.Select(item => Resolve<GridColumn>(document, item, "columns.gridColumn")).ToArray();
        if (selectedColumns.Length != grid.Length || !selectedColumns.SequenceEqual(grid))
            throw new InvalidOperationException("columns-must-cover-table-grid-in-native-order");

        var prototypeRows = prototypeRefs.Select(item => Resolve<TableRow>(document, item, "rows.prototypeRow")).ToArray();
        if (prototypeRows.Any(row => !selectedRows.Contains(row)))
            throw new InvalidOperationException("prototypeRow-must-be-within-existingRows");
        foreach (var prototype in prototypeRows)
        foreach (var cell in prototype.Elements<TableCell>())
            NativeMutationSupport.RequirePlainTextContainer(cell);

        var idToColumn = request.Columns.Select((column, index) => (column.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var preparedRows = PrepareRows(request.Rows, prototypeRows, selectedRows, idToColumn, grid.Length);
        return new PreparedTable(
            tableRef,
            selectedRows.Select(Observation.NativePathFor).ToArray(),
            firstIndex,
            grid.Select(GridWidth).ToArray(),
            preparedRows);
    }

    private static IReadOnlyList<PreparedRow> PrepareRows(
        IReadOnlyList<SetTableBodyRow> requests,
        IReadOnlyList<TableRow> prototypes,
        IReadOnlyList<TableRow> existingRows,
        IReadOnlyDictionary<string, int> idToColumn,
        int columnCount)
    {
        var result = new List<PreparedRow>(requests.Count);
        var active = new Dictionary<int, ActiveVerticalCell>();
        for (var rowIndex = 0; rowIndex < requests.Count; rowIndex++)
        {
            var existingCells = rowIndex < existingRows.Count
                ? CellPositions(existingRows[rowIndex]) : [];
            var cells = active.Values
                .Select(item => new PreparedCell(
                    item.Start,
                    item.Span,
                    string.Empty,
                    "continue",
                    UnchangedContent(existingCells, item.Start, item.Span, string.Empty)))
                .ToList();
            var occupied = active.Values
                .SelectMany(item => Enumerable.Range(item.Start, item.Span)).ToHashSet();
            var next = active.Values.Where(item => item.Remaining > 1)
                .ToDictionary(item => item.Start, item => item with { Remaining = item.Remaining - 1 });

            foreach (var (cell, cellIndex) in requests[rowIndex].Cells.Select((value, index) => (value, index)))
            {
                if (cell.Columns.Count == 0 || cell.Columns.Distinct(StringComparer.Ordinal).Count() != cell.Columns.Count)
                    throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].columns-invalid");
                int[] positions;
                try { positions = cell.Columns.Select(id => idToColumn[id]).ToArray(); }
                catch (KeyNotFoundException)
                {
                    throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].column-unknown");
                }
                Array.Sort(positions);
                if (!positions.SequenceEqual(Enumerable.Range(positions[0], positions.Length)))
                    throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].columns-not-contiguous");
                if (positions.Any(position => !occupied.Add(position)))
                    throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].columns-overlap");
                var rowSpan = cell.RowSpan ?? 1;
                if (rowSpan < 1 || rowSpan > requests.Count - rowIndex)
                    throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].rowSpan-invalid");
                var start = positions[0];
                var span = positions.Length;
                cells.Add(new PreparedCell(
                    start,
                    span,
                    cell.Text,
                    rowSpan > 1 ? "restart" : null,
                    UnchangedContent(existingCells, start, span, cell.Text)));
                if (rowSpan > 1) next.Add(start, new ActiveVerticalCell(start, span, rowSpan - 1));
            }
            if (!occupied.SetEquals(Enumerable.Range(0, columnCount)))
                throw new InvalidOperationException($"rows[{rowIndex}]-does-not-cover-table-grid");
            result.Add(new PreparedRow(
                (TableRow)prototypes[rowIndex].CloneNode(true),
                cells.OrderBy(cell => cell.Start).ToArray(),
                requests[rowIndex].CantSplit));
            active = next;
        }
        if (active.Count != 0) throw new InvalidOperationException("rowSpan-exceeds-final-row");
        return result;
    }

    private static void ApplyPrepared(WordprocessingDocument document, PreparedTable prepared)
    {
        var table = Resolve<Table>(document, prepared.Table, "output-table");
        var originalRows = prepared.BodyRowPaths.Select(path =>
            Observation.ResolveNativePath(document, MainStory, path) as TableRow
                ?? throw new InvalidOperationException("output-existing-row-not-found")).ToArray();
        var insertionPoint = originalRows[^1].NextSibling();
        foreach (var row in originalRows) row.Remove();

        foreach (var rowChange in prepared.Rows)
        {
            var row = (TableRow)rowChange.Prototype.CloneNode(true);
            var templateCells = CellPositions(row);
            row.RemoveAllChildren<TableCell>();
            RemoveGridOmissions(row);
            RemoveCopiedIdentities(row);
            SetCantSplit(row, rowChange.CantSplit);
            foreach (var cellChange in rowChange.Cells)
            {
                var template = templateCells.FirstOrDefault(item =>
                    item.Start <= cellChange.Start && cellChange.Start < item.Start + item.Span)?.Cell
                    ?? throw new InvalidOperationException("prototype-row-does-not-cover-output-column");
                var cell = (TableCell)template.CloneNode(true);
                SetGridSpan(cell, cellChange.Span);
                SetCellWidth(cell, prepared.GridWidths, cellChange.Start, cellChange.Span);
                SetVerticalMerge(cell, cellChange.VerticalMerge);
                if (cellChange.UnchangedContent is null)
                {
                    NativeTextMutation.SetText(cell, cellChange.Text);
                }
                else
                {
                    foreach (var child in cell.ChildElements
                                 .Where(child => child is not TableCellProperties).ToArray()) child.Remove();
                    foreach (var child in cellChange.UnchangedContent)
                        cell.Append(child.CloneNode(true));
                }
                RemoveCopiedIdentities(cell);
                row.Append(cell);
            }
            if (CellPositions(row).Sum(item => item.Span) != prepared.GridWidths.Count)
                throw new InvalidOperationException("output-row-grid-coverage-mismatch");
            if (insertionPoint is null) table.Append(row);
            else table.InsertBefore(row, insertionPoint);
        }
    }

    private static IReadOnlyList<SetTableBodyRowReadback> ReadBack(string output, PreparedTable prepared)
    {
        using var document = WordprocessingDocument.Open(output, false);
        var table = Resolve<Table>(document, prepared.Table, "readback-table");
        var rows = table.Elements<TableRow>().Skip(prepared.FirstRowIndex).Take(prepared.Rows.Count).ToArray();
        if (rows.Length != prepared.Rows.Count) throw new InvalidOperationException("output-readback-row-count-mismatch");
        var result = new List<SetTableBodyRowReadback>();
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var actual = CellPositions(rows[rowIndex]);
            var expected = prepared.Rows[rowIndex].Cells;
            if (actual.Count != expected.Count) throw new InvalidOperationException("output-readback-cell-count-mismatch");
            var cells = new List<SetTableBodyCellReadback>();
            for (var cellIndex = 0; cellIndex < actual.Count; cellIndex++)
            {
                var merge = MergeValue(actual[cellIndex].Cell);
                var text = NativeMutationSupport.PlainText(actual[cellIndex].Cell);
                if (actual[cellIndex].Start != expected[cellIndex].Start
                    || actual[cellIndex].Span != expected[cellIndex].Span
                    || !StringComparer.Ordinal.Equals(merge, expected[cellIndex].VerticalMerge)
                    || !StringComparer.Ordinal.Equals(text, expected[cellIndex].Text))
                    throw new InvalidOperationException("output-readback-cell-mismatch");
                cells.Add(new SetTableBodyCellReadback(actual[cellIndex].Start, actual[cellIndex].Span, merge, text));
            }
            result.Add(new SetTableBodyRowReadback(
                Observation.Address(MainStory, Observation.NativePathFor(rows[rowIndex])),
                rows[rowIndex].TableRowProperties?.GetFirstChild<CantSplit>() is not null,
                cells));
        }
        return result;
    }

    private static T Resolve<T>(WordprocessingDocument document, ResolvedDocxAddress address, string name)
        where T : OpenXmlElement
        => Observation.ResolveNativePath(document, address.StoryPart, address.NativePath) as T
            ?? throw new InvalidOperationException(name + "-address-not-found");

    private static void RequireClosedVerticalMerges(Table table, IReadOnlyList<TableRow> selectedRows)
    {
        var allRows = table.Elements<TableRow>().ToArray();
        var selected = selectedRows.ToHashSet();
        var active = new Dictionary<(int Start, int Span), List<TableRow>>();
        foreach (var row in allRows)
        {
            var continued = new HashSet<(int Start, int Span)>();
            foreach (var cell in CellPositions(row))
            {
                var key = (cell.Start, cell.Span);
                var merge = cell.Cell.TableCellProperties?.VerticalMerge;
                if (merge?.Val?.Value == MergedCellValues.Restart)
                {
                    Close(key);
                    active[key] = [row];
                    continued.Add(key);
                }
                else if (merge is not null)
                {
                    if (!active.TryGetValue(key, out var group))
                        throw new InvalidOperationException("vertical-merge-continuation-without-restart");
                    group.Add(row);
                    continued.Add(key);
                }
            }
            foreach (var key in active.Keys.Where(key => !continued.Contains(key)).ToArray()) Close(key);
        }
        foreach (var key in active.Keys.ToArray()) Close(key);

        void Close((int Start, int Span) key)
        {
            if (!active.Remove(key, out var group)) return;
            var count = group.Count(selected.Contains);
            if (count > 0 && count != group.Count)
                throw new InvalidOperationException("existingRows-split-vertical-merge");
        }
    }

    private static IReadOnlyList<CellPosition> CellPositions(TableRow row)
    {
        var result = new List<CellPosition>();
        var cursor = RowOffset(row, "gridBefore");
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            result.Add(new CellPosition(cell, cursor, span));
            cursor += span;
        }
        return result;
    }

    private static IReadOnlyList<OpenXmlElement>? UnchangedContent(
        IReadOnlyList<CellPosition> existingCells,
        int start,
        int span,
        string text)
    {
        var existing = existingCells.SingleOrDefault(cell => cell.Start == start && cell.Span == span);
        if (existing is null
            || !StringComparer.Ordinal.Equals(NativeMutationSupport.PlainText(existing.Cell), text)) return null;
        return existing.Cell.ChildElements
            .Where(child => child is not TableCellProperties)
            .Select(child => child.CloneNode(true))
            .ToArray();
    }

    private static int RowOffset(TableRow row, string localName)
    {
        var value = row.TableRowProperties?.ChildElements.FirstOrDefault(child => child.LocalName == localName)
            ?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return int.TryParse(value, out var result) ? result : 0;
    }

    private static int? GridWidth(GridColumn column)
        => int.TryParse(column.Width?.Value, out var value) ? value : null;

    private static void RemoveGridOmissions(TableRow row)
    {
        foreach (var child in row.TableRowProperties?.ChildElements
                     .Where(child => child.LocalName is "gridBefore" or "gridAfter").ToArray() ?? []) child.Remove();
    }

    private static void SetGridSpan(TableCell cell, int span)
    {
        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        properties.GridSpan = span > 1 ? new GridSpan { Val = span } : null;
    }

    private static void SetCellWidth(TableCell cell, IReadOnlyList<int?> widths, int start, int span)
    {
        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        int? width = widths.Skip(start).Take(span).All(value => value.HasValue)
            ? widths.Skip(start).Take(span).Sum(value => value!.Value) : null;
        if (width is null) return;
        properties.TableCellWidth = new TableCellWidth
        {
            Type = TableWidthUnitValues.Dxa,
            Width = width.Value.ToString(),
        };
    }

    private static void SetVerticalMerge(TableCell cell, string? value)
    {
        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        properties.VerticalMerge = value switch
        {
            "restart" => new VerticalMerge { Val = MergedCellValues.Restart },
            "continue" => new VerticalMerge { Val = MergedCellValues.Continue },
            _ => null,
        };
    }

    private static void SetCantSplit(TableRow row, bool? value)
    {
        if (value is null) return;
        var properties = row.TableRowProperties ?? row.PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<CantSplit>();
        if (value.Value) properties.Append(new CantSplit());
    }

    private static string? MergeValue(TableCell cell)
    {
        var merge = cell.TableCellProperties?.VerticalMerge;
        if (merge is null) return null;
        return merge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
    }

    private static void RemoveCopiedIdentities(OpenXmlElement root)
    {
        foreach (var item in new[] { root }.Concat(root.Descendants()).ToArray())
        {
            item.RemoveAttribute("paraId", Word2010Namespace);
            item.RemoveAttribute("textId", Word2010Namespace);
            if (item is BookmarkStart or BookmarkEnd) item.Remove();
        }
    }

    private sealed record CellPosition(TableCell Cell, int Start, int Span);
    private sealed record ActiveVerticalCell(int Start, int Span, int Remaining);
    private sealed record PreparedCell(
        int Start,
        int Span,
        string Text,
        string? VerticalMerge,
        IReadOnlyList<OpenXmlElement>? UnchangedContent);
    private sealed record PreparedRow(
        TableRow Prototype,
        IReadOnlyList<PreparedCell> Cells,
        bool? CantSplit);
    private sealed record PreparedTable(
        ResolvedDocxAddress Table,
        IReadOnlyList<string> BodyRowPaths,
        int FirstRowIndex,
        IReadOnlyList<int?> GridWidths,
        IReadOnlyList<PreparedRow> Rows);
}

public sealed record SetTableBodyRowRange(DocxObjectAddress First, DocxObjectAddress Last);
public sealed record SetTableBodyColumn(string Id, DocxObjectAddress GridColumn);
public sealed record SetTableBodyCell(IReadOnlyList<string> Columns, string Text, int? RowSpan = null);
public sealed record SetTableBodyRow(
    DocxObjectAddress PrototypeRow,
    IReadOnlyList<SetTableBodyCell> Cells,
    bool? CantSplit = null);
public sealed record SetTableBodyRequest(
    string Input,
    DocxObjectAddress Table,
    SetTableBodyRowRange ExistingRows,
    IReadOnlyList<SetTableBodyColumn> Columns,
    IReadOnlyList<SetTableBodyRow> Rows,
    string Output,
    string ReceiptOutput);
public sealed record SetTableBodyCellReadback(int GridColumnStart, int GridSpan, string? VerticalMerge, string Text);
public sealed record SetTableBodyRowReadback(
    DocxObjectAddress Address,
    bool CantSplit,
    IReadOnlyList<SetTableBodyCellReadback> Cells);
public sealed record SetTableBodyReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    DocxObjectAddress Table,
    IReadOnlyList<SetTableBodyRowReadback> Rows,
    string Output);
