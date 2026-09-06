using System.Text.Json;

namespace Dockit.Docx;

public static class NativeSetTableMutation
{
    public const string Command = "docx_set_table";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetTableRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-table-request-invalid");
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

    public static SetTableReceipt Apply(SetTableRequest request)
    {
        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        ValidateContentModes(request.Rows);
        var token = Guid.NewGuid().ToString("N");
        var shapeOutput = paths.Output + $".shape-{token}.docx";
        var shapeReceipt = paths.Output + $".shape-{token}.json";
        var contentOutput = paths.Output + $".content-{token}.docx";
        var contentReceipt = paths.Output + $".content-{token}.json";
        var finalTemporary = paths.Output + $".tmp-{token}";
        try
        {
            var retainedSourceText = RetainedSourceText(request);
            var shapeRequest = new SetTableBodyRequest(
                paths.Input,
                request.Table,
                request.ExistingRows,
                request.Columns,
                request.Rows.Select((row, rowIndex) => new SetTableBodyRow(
                    row.PrototypeRow,
                    row.Cells.Select((cell, cellIndex) => new SetTableBodyCell(
                        cell.Columns,
                        cell.Text ?? (cell.TextRuns is null
                            ? retainedSourceText.GetValueOrDefault((rowIndex, cellIndex), string.Empty)
                            : string.Concat(cell.TextRuns.Select(run => run.Text))),
                        cell.RowSpan)).ToArray(),
                    row.CantSplit)).ToArray(),
                shapeOutput,
                shapeReceipt);
            var shaped = NativeTableBodyMutation.Apply(shapeRequest);
            ApplyRichText(request, shaped, shapeOutput);
            var changes = BuildContentChanges(request, shaped, shapeOutput);
            var completed = shapeOutput;
            if (changes.Count > 0)
            {
                NativeContentCopy.Apply(new CopyContentRequest(
                    shapeOutput,
                    changes,
                    contentOutput,
                    contentReceipt));
                completed = contentOutput;
            }

            var readback = ReadBack(completed, shaped);
            File.Copy(completed, finalTemporary, false);
            var receipt = new SetTableReceipt(
                "tiwater.docx-set-table-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                request.Table,
                readback,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            NativeMutationSupport.Commit(finalTemporary, paths);
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(finalTemporary, paths);
            throw;
        }
        finally
        {
            NativeMutationSupport.Cleanup(shapeOutput, shapeReceipt, contentOutput, contentReceipt);
        }
    }

    private static void ValidateContentModes(IReadOnlyList<SetTableRow> rows)
    {
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        for (var cellIndex = 0; cellIndex < rows[rowIndex].Cells.Count; cellIndex++)
        {
            var cell = rows[rowIndex].Cells[cellIndex];
            var hasText = cell.Text is not null;
            var hasRuns = cell.TextRuns is not null;
            var hasSource = cell.SourceInput is not null || cell.SourceSelections is not null;
            if ((hasText ? 1 : 0) + (hasRuns ? 1 : 0) + (hasSource ? 1 : 0) != 1)
                throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}]-requires-exactly-one-content-mode");
            if (hasSource && (string.IsNullOrWhiteSpace(cell.SourceInput)
                || cell.SourceSelections is null || cell.SourceSelections.Count == 0))
                throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}]-source-content-incomplete");
            if (hasRuns)
            {
                if (cell.TextRuns!.Count == 0)
                    throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}]-text-runs-empty");
                for (var runIndex = 0; runIndex < cell.TextRuns.Count; runIndex++)
                {
                    var run = cell.TextRuns[runIndex];
                    if (run.Text.Length == 0)
                        throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].textRuns[{runIndex}]-text-empty");
                    if (run.Color is not null && ((run.Color.Length != 6 && run.Color.Length != 8)
                        || run.Color.Any(character => !Uri.IsHexDigit(character))
                        || (run.Color.Length == 8 && !run.Color.StartsWith("FF", StringComparison.OrdinalIgnoreCase))))
                        throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].textRuns[{runIndex}]-color-invalid");
                    if (!TryUnderline(run.Underline, out _))
                        throw new InvalidOperationException($"rows[{rowIndex}].cells[{cellIndex}].textRuns[{runIndex}]-underline-invalid");
                }
            }
        }
    }

    private static void ApplyRichText(SetTableRequest request, SetTableBodyReceipt shaped, string shapeOutput)
    {
        var columnStarts = request.Columns.Select((column, index) => (column.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var table = Observation.ReadTable(shapeOutput, shaped.Table);
        var rowsByAddress = table.Rows.ToDictionary(row => row.Address, row => row);
        var changes = new List<(DocxObjectAddress Address, IReadOnlyList<SetTableTextRun> Runs)>();
        for (var rowIndex = 0; rowIndex < request.Rows.Count; rowIndex++)
        {
            if (!rowsByAddress.TryGetValue(shaped.Rows[rowIndex].Address, out var observedRow))
                throw new InvalidOperationException("set-table-shaped-row-not-found");
            foreach (var cell in request.Rows[rowIndex].Cells.Where(cell => cell.TextRuns is not null))
            {
                int start;
                try { start = cell.Columns.Select(id => columnStarts[id]).Min(); }
                catch (KeyNotFoundException) { throw new InvalidOperationException("set-table-rich-text-column-unknown"); }
                var target = observedRow.Cells.SingleOrDefault(item => item.GridColumnStart == start)
                    ?? throw new InvalidOperationException("set-table-rich-text-shaped-cell-not-found");
                changes.Add((target.Address, cell.TextRuns!));
            }
        }
        if (changes.Count == 0) return;

        var resolved = Observation.ResolveAddresses(shapeOutput, changes.Select(change => change.Address).ToArray(), "rows.cells.textRuns");
        using (var output = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(shapeOutput, true))
        {
            for (var index = 0; index < resolved.Count; index++)
            {
                var target = Observation.ResolveNativePath(output, resolved[index].StoryPart, resolved[index].NativePath)
                    as DocumentFormat.OpenXml.Wordprocessing.TableCell
                    ?? throw new InvalidOperationException("set-table-rich-text-target-not-cell");
                NativeTextMutation.SetTextRuns(target, changes[index].Runs);
            }
            output.MainDocumentPart?.Document?.Save();
        }

        using var readback = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(shapeOutput, false);
        for (var index = 0; index < resolved.Count; index++)
        {
            var target = Observation.ResolveNativePath(readback, resolved[index].StoryPart, resolved[index].NativePath)
                as DocumentFormat.OpenXml.Wordprocessing.TableCell
                ?? throw new InvalidOperationException("set-table-rich-text-readback-target-not-cell");
            var actual = target.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
                .Where(run => NativeMutationSupport.PlainText(run).Length > 0).ToArray();
            var expected = changes[index].Runs;
            if (actual.Length != expected.Count) throw new InvalidOperationException("set-table-rich-text-readback-run-count-mismatch");
            for (var runIndex = 0; runIndex < actual.Length; runIndex++)
            {
                var color = actual[runIndex].RunProperties?.Color?.Val?.Value;
                var underlineValue = actual[runIndex].RunProperties?.Underline?.Val?.Value;
                var underline = underlineValue == DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single ? "single"
                    : underlineValue == DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Double ? "double"
                    : null;
                if (!StringComparer.Ordinal.Equals(NativeMutationSupport.PlainText(actual[runIndex]), expected[runIndex].Text)
                    || !StringComparer.OrdinalIgnoreCase.Equals(color, ColorValue(expected[runIndex].Color))
                    || !StringComparer.Ordinal.Equals(underline, UnderlineValue(expected[runIndex].Underline)))
                    throw new InvalidOperationException(
                        $"set-table-rich-text-readback-mismatch: change={index}; run={runIndex}; "
                        + $"text={NativeMutationSupport.PlainText(actual[runIndex])}; color={color ?? "null"}; underline={underline ?? "null"}");
            }
        }
    }

    internal static string? ColorValue(string? color)
        => color?.Length == 8 ? color[2..] : color;

    internal static string? UnderlineValue(JsonElement underline)
        => TryUnderline(underline, out var value) ? value : throw new InvalidOperationException("text-run-underline-invalid");

    private static bool TryUnderline(JsonElement underline, out string? value)
    {
        value = null;
        if (underline.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null or JsonValueKind.False) return true;
        if (underline.ValueKind != JsonValueKind.String) return false;
        value = underline.GetString();
        return value is "single" or "double";
    }

    private static IReadOnlyList<CopyContentChange> BuildContentChanges(
        SetTableRequest request,
        SetTableBodyReceipt shaped,
        string shapeOutput)
    {
        var columnStarts = request.Columns.Select((column, index) => (column.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var table = Observation.ReadTable(shapeOutput, shaped.Table);
        var rowsByAddress = table.Rows.ToDictionary(row => row.Address, row => row);
        var result = new List<CopyContentChange>();
        for (var rowIndex = 0; rowIndex < request.Rows.Count; rowIndex++)
        {
            if (!rowsByAddress.TryGetValue(shaped.Rows[rowIndex].Address, out var observedRow))
                throw new InvalidOperationException("set-table-shaped-row-not-found");
            foreach (var cell in request.Rows[rowIndex].Cells.Where(cell => cell.SourceInput is not null))
            {
                int start;
                try { start = cell.Columns.Select(id => columnStarts[id]).Min(); }
                catch (KeyNotFoundException) { throw new InvalidOperationException("set-table-content-column-unknown"); }
                var target = observedRow.Cells.SingleOrDefault(item => item.GridColumnStart == start)
                    ?? throw new InvalidOperationException("set-table-shaped-cell-not-found");
                if (IsExactRetainedSourceCell(shapeOutput, cell, target.Address)) continue;
                result.Add(new CopyContentChange(
                    target.Address,
                    cell.SourceInput!,
                    cell.SourceSelections!));
            }
        }
        return result;
    }

    private static IReadOnlyDictionary<(int Row, int Cell), string> RetainedSourceText(SetTableRequest request)
    {
        var table = Observation.ReadTable(request.Input, request.Table);
        var first = table.Rows.Select((row, index) => (row, index))
            .SingleOrDefault(item => item.row.Address == request.ExistingRows.First);
        var last = table.Rows.Select((row, index) => (row, index))
            .SingleOrDefault(item => item.row.Address == request.ExistingRows.Last);
        if (first.row is null || last.row is null || last.index < first.index) return new Dictionary<(int, int), string>();

        var selectedRows = table.Rows.Skip(first.index).Take(last.index - first.index + 1).ToArray();
        var columnStarts = request.Columns.Select((column, index) => (column.Id, index))
            .ToDictionary(item => item.Id, item => item.index, StringComparer.Ordinal);
        var result = new Dictionary<(int, int), string>();
        for (var rowIndex = 0; rowIndex < request.Rows.Count && rowIndex < selectedRows.Length; rowIndex++)
        for (var cellIndex = 0; cellIndex < request.Rows[rowIndex].Cells.Count; cellIndex++)
        {
            var cell = request.Rows[rowIndex].Cells[cellIndex];
            if (cell.SourceInput is null) continue;
            int start;
            try { start = cell.Columns.Select(id => columnStarts[id]).Min(); }
            catch (KeyNotFoundException) { continue; }
            var target = selectedRows[rowIndex].Cells.SingleOrDefault(item => item.GridColumnStart == start);
            if (target is not null && IsExactRetainedSourceCell(request.Input, cell, target.Address))
                result[(rowIndex, cellIndex)] = target.LogicalText;
        }
        return result;
    }

    private static bool IsExactRetainedSourceCell(string input, SetTableCell cell, DocxObjectAddress target)
    {
        if (cell.SourceInput is null || cell.SourceSelections is not [var selection] || selection.Range is not null)
            return false;
        var targetRef = Observation.ResolveAddresses(input, [target], "retainedSource.target").Single();
        var sourceRef = Observation.ResolveAddresses(cell.SourceInput, [selection.Address], "retainedSource.source").Single();
        if (targetRef.Kind != "cell" || sourceRef.Kind != "cell") return false;
        using var targetDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(input, false);
        using var sourceDocument = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(cell.SourceInput, false);
        var targetCell = Observation.ResolveNativePath(targetDocument, targetRef.StoryPart, targetRef.NativePath)
            as DocumentFormat.OpenXml.Wordprocessing.TableCell;
        var sourceCell = Observation.ResolveNativePath(sourceDocument, sourceRef.StoryPart, sourceRef.NativePath)
            as DocumentFormat.OpenXml.Wordprocessing.TableCell;
        if (targetCell is null || sourceCell is null) return false;
        return NativeContent(targetCell).Equals(NativeContent(sourceCell), StringComparison.Ordinal);
    }

    private static string NativeContent(DocumentFormat.OpenXml.Wordprocessing.TableCell cell)
        => string.Concat(cell.ChildElements
            .Where(child => child is not DocumentFormat.OpenXml.Wordprocessing.TableCellProperties)
            .Select(child => child.OuterXml));

    private static IReadOnlyList<SetTableBodyRowReadback> ReadBack(
        string output,
        SetTableBodyReceipt shaped)
    {
        var table = Observation.ReadTable(output, shaped.Table);
        var rowsByAddress = table.Rows.ToDictionary(row => row.Address, row => row);
        return shaped.Rows.Select(row =>
        {
            if (!rowsByAddress.TryGetValue(row.Address, out var observed))
                throw new InvalidOperationException("set-table-final-row-not-found");
            return new SetTableBodyRowReadback(
                observed.Address,
                observed.CantSplit,
                observed.Cells.Select(cell => new SetTableBodyCellReadback(
                    cell.GridColumnStart,
                    cell.GridSpan,
                    cell.VerticalMerge,
                    cell.LogicalText)).ToArray());
        }).ToArray();
    }
}

public sealed record SetTableCell(
    IReadOnlyList<string> Columns,
    string? Text,
    IReadOnlyList<SetTableTextRun>? TextRuns,
    string? SourceInput,
    IReadOnlyList<CopyContentSelection>? SourceSelections,
    int? RowSpan = null);
public sealed record SetTableTextRun(string Text, string? Color = null, JsonElement Underline = default);
public sealed record SetTableRow(
    DocxObjectAddress PrototypeRow,
    IReadOnlyList<SetTableCell> Cells,
    bool? CantSplit = null);
public sealed record SetTableRequest(
    string Input,
    DocxObjectAddress Table,
    SetTableBodyRowRange ExistingRows,
    IReadOnlyList<SetTableBodyColumn> Columns,
    IReadOnlyList<SetTableRow> Rows,
    string Output,
    string ReceiptOutput);
public sealed record SetTableReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    DocxObjectAddress Table,
    IReadOnlyList<SetTableBodyRowReadback> Rows,
    string Output);
