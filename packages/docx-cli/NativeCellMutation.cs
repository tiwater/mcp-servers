using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeCellMutation
{
    public const string MergeCommand = "docx_merge_cells";
    public const string SplitCommand = "docx_split_cells";
    private const string MainStory = "/word/document.xml";

    public static int Run(string command, string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        var request = JsonSerializer.Deserialize<CellMutationRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("cell-mutation-request-invalid");
        var receipt = Apply(command, request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = command,
            receipt = Describe(request.ReceiptOutput),
            output = Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static CellMutationReceipt Apply(string command, CellMutationRequest request)
    {
        if (command is not MergeCommand and not SplitCommand) throw new InvalidOperationException("cell-mutation-command-invalid");
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var targetPath = Path.GetFullPath(request.Input);
        var outputPath = Path.GetFullPath(request.Output);
        var receiptPath = Path.GetFullPath(request.ReceiptOutput);
        RequireNewPath(outputPath, "output");
        RequireNewPath(receiptPath, "receiptOutput");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, receiptPath)) throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
        if (StringComparer.OrdinalIgnoreCase.Equals(outputPath, targetPath)) throw new InvalidOperationException("output-must-not-overwrite-input");

        var addresses = request.Changes.SelectMany(change => change.Cells).ToArray();
        if (addresses.Length == 0 || addresses.Distinct().Count() != addresses.Length)
            throw new InvalidOperationException("cell-addresses-empty-or-duplicate");
        var resolved = Observation.ResolveAddresses(targetPath, addresses, "changes.cells");
        if (resolved.Any(item => item.Kind != "cell" || item.StoryPart != MainStory))
            throw new InvalidOperationException("cell-address-must-be-main-document-cell");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(targetPath, false)) baseline = ValidationIssueCounts(input);
        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            File.Copy(targetPath, temporaryPath, false);
            var changedPaths = new List<string>();
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                var offset = 0;
                foreach (var change in request.Changes)
                {
                    var cells = resolved.Skip(offset).Take(change.Cells.Count)
                        .Select(item => Observation.ResolveNativePath(output, item.StoryPart, item.NativePath) as TableCell
                            ?? throw new InvalidOperationException("cell-address-not-found-in-output"))
                        .ToArray();
                    offset += change.Cells.Count;
                    var changed = command == MergeCommand ? Merge(cells) : Split(cells);
                    changedPaths.AddRange(changed.Select(Observation.NativePathFor));
                }
                output.MainDocumentPart?.Document?.Save();
                RejectAddedValidationIssues(output, baseline);
            }
            File.Move(temporaryPath, outputPath);
            IReadOnlyList<CellMutationReadback> changes;
            using (var output = WordprocessingDocument.Open(outputPath, false))
                changes = changedPaths.Select(path => Readback(
                    Observation.ResolveNativePath(output, MainStory, path) as TableCell
                        ?? throw new InvalidOperationException("changed-cell-readback-failed"))).ToArray();
            var receipt = new CellMutationReceipt(
                command == MergeCommand ? "tiwater.docx-merge-cells-receipt" : "tiwater.docx-split-cells-receipt",
                "tiwater.docx.cli", RuntimeIdentity.Version, changes, outputPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            foreach (var path in new[] { temporaryPath, outputPath, receiptPath }) if (File.Exists(path)) File.Delete(path);
            throw;
        }
    }

    private static IReadOnlyList<TableCell> Merge(IReadOnlyList<TableCell> cells)
    {
        if (cells.Count < 2) throw new InvalidOperationException("merge-requires-at-least-two-cells");
        var table = cells[0].Ancestors<Table>().FirstOrDefault() ?? throw new InvalidOperationException("cell-table-not-found");
        if (cells.Any(cell => !ReferenceEquals(cell.Ancestors<Table>().FirstOrDefault(), table)))
            throw new InvalidOperationException("merge-cells-must-share-one-table");
        var rows = table.Elements<TableRow>().ToArray();
        var selected = cells.Select(cell => Position(rows, cell)).OrderBy(item => item.Row).ThenBy(item => item.Start).ToArray();
        var minRow = selected.Min(item => item.Row);
        var maxRow = selected.Max(item => item.Row);
        var minColumn = selected.Min(item => item.Start);
        var maxColumn = selected.Max(item => item.End);
        var selectedSet = cells.ToHashSet();
        for (var rowIndex = minRow; rowIndex <= maxRow; rowIndex++)
        {
            var rowCells = RowPositions(rows[rowIndex]).Where(item => item.End >= minColumn && item.Start <= maxColumn).ToArray();
            if (rowCells.Length == 0 || rowCells.Min(item => item.Start) != minColumn || rowCells.Max(item => item.End) != maxColumn
                || rowCells.Any(item => !selectedSet.Contains(item.Cell)))
                throw new InvalidOperationException("merge-cell-selection-must-be-one-closed-rectangle");
        }

        var owners = new List<TableCell>();
        for (var rowIndex = minRow; rowIndex <= maxRow; rowIndex++)
        {
            var rowSelection = RowPositions(rows[rowIndex]).Where(item => item.Start >= minColumn && item.End <= maxColumn).ToArray();
            var owner = rowSelection[0].Cell;
            foreach (var removed in rowSelection.Skip(1))
            {
                foreach (var paragraph in removed.Cell.Elements<Paragraph>().ToArray())
                {
                    paragraph.Remove();
                    owner.Append(paragraph);
                }
                removed.Cell.Remove();
            }
            var properties = owner.TableCellProperties ?? owner.PrependChild(new TableCellProperties());
            properties.GridSpan = null;
            var span = maxColumn - minColumn + 1;
            if (span > 1) properties.GridSpan = new GridSpan { Val = span };
            SetCellWidth(owner, table, minColumn, span);
            owners.Add(owner);
        }

        if (owners.Count > 1)
        {
            for (var index = 0; index < owners.Count; index++)
            {
                var properties = owners[index].TableCellProperties ?? owners[index].PrependChild(new TableCellProperties());
                properties.VerticalMerge = new VerticalMerge { Val = index == 0 ? MergedCellValues.Restart : MergedCellValues.Continue };
                if (index > 0)
                {
                    owners[index].RemoveAllChildren<Paragraph>();
                    owners[index].Append(new Paragraph());
                }
            }
        }
        return [owners[0]];
    }

    private static IReadOnlyList<TableCell> Split(IReadOnlyList<TableCell> cells)
    {
        var result = new List<TableCell>();
        foreach (var cell in cells)
        {
            var row = cell.Parent as TableRow ?? throw new InvalidOperationException("cell-row-not-found");
            var table = row.Parent as Table ?? throw new InvalidOperationException("cell-table-not-found");
            var rows = table.Elements<TableRow>().ToArray();
            var ownerPosition = Position(rows, cell);
            var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
            var span = Math.Max(1, properties.GridSpan?.Val?.Value ?? 1);
            var vertical = properties.VerticalMerge;
            if (span == 1 && vertical?.Val?.Value != MergedCellValues.Restart)
                throw new InvalidOperationException("split-cell-must-be-a-merge-owner");
            result.AddRange(SplitHorizontal(cell, table, ownerPosition.Start, span));
            if (vertical?.Val?.Value == MergedCellValues.Restart)
            {
                properties.VerticalMerge = null;
                for (var rowIndex = ownerPosition.Row + 1; rowIndex < rows.Length; rowIndex++)
                {
                    var continuation = RowPositions(rows[rowIndex]).FirstOrDefault(item => item.Start == ownerPosition.Start);
                    var merge = continuation?.Cell.TableCellProperties?.VerticalMerge;
                    if (continuation is null || merge is null || merge.Val?.Value == MergedCellValues.Restart) break;
                    continuation.Cell.TableCellProperties!.VerticalMerge = null;
                    result.Add(continuation.Cell);
                    result.AddRange(SplitHorizontal(continuation.Cell, table, ownerPosition.Start, span));
                }
            }
            result.Add(cell);
        }
        return result;
    }

    private static IReadOnlyList<TableCell> SplitHorizontal(TableCell cell, Table table, int startColumn, int span)
    {
        var result = new List<TableCell>();
        var row = cell.Parent as TableRow ?? throw new InvalidOperationException("cell-row-not-found");
        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        properties.GridSpan = null;
        SetCellWidth(cell, table, startColumn, 1);
        var cursor = cell;
        for (var index = 1; index < span; index++)
        {
            var clone = (TableCell)cell.CloneNode(true);
            clone.RemoveAllChildren<Paragraph>();
            clone.Append(new Paragraph());
            var cloneProperties = clone.TableCellProperties ?? clone.PrependChild(new TableCellProperties());
            cloneProperties.GridSpan = null;
            cloneProperties.VerticalMerge = null;
            SetCellWidth(clone, table, startColumn + index, 1);
            row.InsertAfter(clone, cursor);
            cursor = clone;
            result.Add(clone);
        }
        return result;
    }

    private static void SetCellWidth(TableCell cell, Table table, int startColumn, int count)
    {
        var columns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToArray() ?? [];
        if (startColumn < 0 || count < 1 || startColumn + count > columns.Length) return;
        var widths = columns.Skip(startColumn).Take(count).Select(column =>
            int.TryParse(column.Width?.Value, out var width) ? width : 0).ToArray();
        if (widths.Any(width => width <= 0)) return;
        var properties = cell.TableCellProperties ?? cell.PrependChild(new TableCellProperties());
        properties.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = widths.Sum().ToString() };
    }

    private static CellPosition Position(IReadOnlyList<TableRow> rows, TableCell cell)
    {
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var found = RowPositions(rows[rowIndex]).FirstOrDefault(item => ReferenceEquals(item.Cell, cell));
            if (found is not null) return new CellPosition(cell, rowIndex, found.Start, found.End);
        }
        throw new InvalidOperationException("cell-position-not-found");
    }

    private static IReadOnlyList<RowCellPosition> RowPositions(TableRow row)
    {
        var cursor = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        var result = new List<RowCellPosition>();
        foreach (var cell in row.Elements<TableCell>())
        {
            var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            result.Add(new RowCellPosition(cell, cursor, cursor + span - 1));
            cursor += span;
        }
        return result;
    }

    private static CellMutationReadback Readback(TableCell cell)
    {
        var path = Observation.NativePathFor(cell);
        var properties = cell.TableCellProperties;
        var verticalMerge = properties?.VerticalMerge is not { } merge
            ? null
            : merge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
        return new CellMutationReadback(
            Observation.Address(MainStory, path),
            Math.Max(1, properties?.GridSpan?.Val?.Value ?? 1),
            verticalMerge,
            cell.InnerText);
    }

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document).GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static void RejectAddedValidationIssues(WordprocessingDocument document, IReadOnlyDictionary<string, int> baseline)
    {
        var added = ValidationIssueCounts(document).FirstOrDefault(item => item.Value > baseline.GetValueOrDefault(item.Key));
        if (added.Key is not null) throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
    }

    private static void RequireNewPath(string path, string name)
    {
        if (File.Exists(path) || Directory.Exists(path)) throw new InvalidOperationException($"{name}-already-exists");
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory)) throw new InvalidOperationException($"{name}-directory-not-found");
    }

    private static ObjectArtifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new ObjectArtifact(Path.GetFullPath(path), Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), stream.Length);
    }

    private sealed record CellPosition(TableCell Cell, int Row, int Start, int End);
    private sealed record RowCellPosition(TableCell Cell, int Start, int End);
}

public sealed record CellMutationChange(IReadOnlyList<DocxObjectAddress> Cells);
public sealed record CellMutationRequest(string Input, IReadOnlyList<CellMutationChange> Changes, string Output, string ReceiptOutput);
public sealed record CellMutationReadback(DocxObjectAddress Address, int GridSpan, string? VerticalMerge, string Text);
public sealed record CellMutationReceipt(string Schema, string Provider, string ToolVersion, IReadOnlyList<CellMutationReadback> Changes, string Output);
