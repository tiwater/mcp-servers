using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace Dockit.Docx;

internal static partial class Editor
{
    private static DocxEditAppliedOperation ReplacePartTableCellText(WordprocessingDocument doc, DocxEditOperation operation, string partKind)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and text are required");
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var partIndex = partKind == "header" ? operation.HeaderIndex : operation.FooterIndex;
        if (partIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"{partKind}Index is required");
        }

        var roots = partKind == "header"
            ? mainPart.HeaderParts.Where(part => part.Header is not null).OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => (OpenXmlPartRootElement)part.Header!).ToList()
            : mainPart.FooterParts.Where(part => part.Footer is not null).OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => (OpenXmlPartRootElement)part.Footer!).ToList();
        if (partIndex.Value < 0 || partIndex.Value >= roots.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"{partKind}Index {partIndex} is out of range");
        }

        var tables = roots[partIndex.Value].Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range for {partKind} {partIndex}");
        }
        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }
        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        var fallbackRun = FindNearestTableRun(rows, operation.RowIndex.Value, operation.CellIndex.Value);
        ReplaceTableCellText(cells[operation.CellIndex.Value], operation.Text, operation.Alignment, fallbackRun, operation.ParagraphTexts);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated {partKind}[{partIndex}].table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}]");
    }

    private static DocxEditAppliedOperation ReplaceTableCellRunText(Body body, DocxEditOperation operation)
        => ReplaceTableCellRunText(body.Elements<Table>().ToList(), operation, "body", null);

    private static DocxEditAppliedOperation ReplacePartTableCellRunText(WordprocessingDocument doc, DocxEditOperation operation, string partKind)
    {
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var partIndex = partKind == "header" ? operation.HeaderIndex : operation.FooterIndex;
        if (partIndex is null) return new DocxEditAppliedOperation(operation.Type, false, $"{partKind}Index is required");
        var roots = partKind == "header"
            ? mainPart.HeaderParts.Where(part => part.Header is not null).OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => (OpenXmlPartRootElement)part.Header!).ToList()
            : mainPart.FooterParts.Where(part => part.Footer is not null).OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => (OpenXmlPartRootElement)part.Footer!).ToList();
        if (partIndex.Value < 0 || partIndex.Value >= roots.Count) return new DocxEditAppliedOperation(operation.Type, false, $"{partKind}Index {partIndex} is out of range");
        return ReplaceTableCellRunText(roots[partIndex.Value].Elements<Table>().ToList(), operation, partKind, partIndex);
    }

    private static DocxEditAppliedOperation ReplaceTableCellRunText(IReadOnlyList<Table> tables, DocxEditOperation operation, string scope, int? partIndex)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || operation.ParagraphIndex is null || operation.RunIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, paragraphIndex, runIndex, and text are required");
        }
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count) return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count) return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count) return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        if (!TryReplaceRunText(cells[operation.CellIndex.Value].Elements<Paragraph>().ToList(), operation.ParagraphIndex.Value, operation.RunIndex.Value, operation.Text, out var error))
        {
            return new DocxEditAppliedOperation(operation.Type, false, error);
        }
        var location = partIndex is null ? scope : $"{scope}[{partIndex}]";
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated {location}.table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}].paragraph[{operation.ParagraphIndex}].run[{operation.RunIndex}]");
    }

    private static DocxEditAppliedOperation ReplaceTableCellText(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and text are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        var fallbackRun = FindNearestTableRun(rows, operation.RowIndex.Value, operation.CellIndex.Value);
        ReplaceTableCellText(cells[operation.CellIndex.Value], operation.Text, operation.Alignment, fallbackRun, operation.ParagraphTexts);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}]");
    }

    private static DocxEditAppliedOperation SetTableWidth(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex is required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var properties = tables[operation.TableIndex.Value].GetFirstChild<TableProperties>() ?? tables[operation.TableIndex.Value].PrependChild(new TableProperties());
        properties.RemoveAllChildren<TableWidth>();
        var widthType = string.Equals(operation.WidthType, "dxa", StringComparison.OrdinalIgnoreCase)
            ? TableWidthUnitValues.Dxa
            : TableWidthUnitValues.Pct;
        properties.PrependChild(new TableWidth
        {
            Width = string.IsNullOrWhiteSpace(operation.Width) ? "5000" : operation.Width,
            Type = widthType,
        });
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}] width");
    }

    private static DocxEditAppliedOperation SetTableCellAlignment(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || string.IsNullOrWhiteSpace(operation.Alignment))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and alignment are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        ApplyCellAlignment(cells[operation.CellIndex.Value], operation.Alignment);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}] alignment");
    }

    private static DocxEditAppliedOperation SetTableCellNoWrap(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and cellIndex are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        var properties = cells[operation.CellIndex.Value].GetFirstChild<TableCellProperties>()
            ?? cells[operation.CellIndex.Value].PrependChild(new TableCellProperties());
        properties.RemoveAllChildren<NoWrap>();
        var noWrap = operation.NoWrap is not false;
        if (noWrap)
        {
            properties.AppendChild(new NoWrap());
        }
        NormalizeTableCellProperties(properties);

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}] noWrap={noWrap}");
    }

    private static DocxEditAppliedOperation SetTableCellFontSize(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || string.IsNullOrWhiteSpace(operation.FontSize))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and fontSize are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        var normalizedSize = NormalizeFontSize(operation.FontSize);
        if (normalizedSize is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid fontSize: {operation.FontSize}");
        }
        foreach (var run in cells[operation.CellIndex.Value].Descendants<Run>())
        {
            var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
            properties.RemoveAllChildren<FontSize>();
            properties.RemoveAllChildren<FontSizeComplexScript>();
            properties.AppendChild(new FontSize { Val = normalizedSize });
            properties.AppendChild(new FontSizeComplexScript { Val = normalizedSize });
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}] font size");
    }

    private static DocxEditAppliedOperation ApplyDocumentFontPolicy(Body body, DocxEditOperation operation)
    {
        if (operation.FontPolicy is null)
            return new DocxEditAppliedOperation(operation.Type, false, "fontPolicy is required");
        if (!FontPolicy.TryNormalize(operation.FontPolicy, out var normalized, out var error))
            return new DocxEditAppliedOperation(operation.Type, false, error ?? "fontPolicy is invalid");

        var bodyCount = 0;
        var tableCount = 0;
        foreach (var run in body.Descendants<Run>().Where(FontPolicy.HasText))
        {
            var inTable = run.Ancestors<Table>().Any();
            FontPolicy.Apply(run, inTable ? normalized.Table : normalized.Body);
            if (inTable) tableCount++; else bodyCount++;
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Applied document font policy to {bodyCount} body run(s) and {tableCount} table run(s)");
    }

    private static DocxEditAppliedOperation SetTableRowHeight(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || string.IsNullOrWhiteSpace(operation.Height))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and height are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        if (!uint.TryParse(operation.Height, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var height))
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid height: {operation.Height}");
        }

        var heightRule = ParseHeightRule(operation.HeightRule);
        if (heightRule is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid heightRule: {operation.HeightRule}");
        }

        var properties = rows[operation.RowIndex.Value].GetFirstChild<TableRowProperties>() ?? rows[operation.RowIndex.Value].PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<TableRowHeight>();
        properties.AppendChild(new TableRowHeight
        {
            Val = UInt32Value.FromUInt32(height),
            HeightType = heightRule.Value,
        });

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}] height");
    }

    private static DocxEditAppliedOperation SetTableRowCantSplit(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CantSplit is null)
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and cantSplit are required");

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");

        var properties = rows[operation.RowIndex.Value].GetFirstChild<TableRowProperties>()
            ?? rows[operation.RowIndex.Value].PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<CantSplit>();
        if (operation.CantSplit.Value) properties.AppendChild(new CantSplit());

        return new DocxEditAppliedOperation(operation.Type, true,
            $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}] cantSplit={operation.CantSplit.Value.ToString().ToLowerInvariant()}");
    }

    private static IReadOnlyList<DocxEditAppliedOperation>? ValidateRepeatHeaderBatch(
        WordprocessingDocument doc,
        IReadOnlyList<DocxEditOperation> operations)
    {
        var repeatOperations = operations.Where(operation => operation.Type == "setTableRowRepeatAsHeader").ToList();
        if (repeatOperations.Count == 0) return null;

        var errors = new string?[repeatOperations.Count];
        var seen = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var index = 0; index < repeatOperations.Count; index++)
        {
            var operation = repeatOperations[index];
            if (!TryResolveTableRow(doc, operation, out _, out var address, out var error))
            {
                errors[index] = error;
                continue;
            }
            if (!seen.TryAdd(address, index))
            {
                errors[index] = $"Duplicate table row target: {address}";
                errors[seen[address]] ??= $"Duplicate table row target: {address}";
            }
        }

        if (errors.All(error => error is null)) return null;
        var repeatIndex = 0;
        return operations.Select(operation =>
        {
            if (operation.Type == "setTableRowRepeatAsHeader")
            {
                var detail = errors[repeatIndex++] ?? "Batch rejected before mutation because another repeat-header target is invalid";
                return new DocxEditAppliedOperation(operation.Type, false, detail);
            }
            return new DocxEditAppliedOperation(operation.Type, false, "Batch rejected before mutation");
        }).ToList();
    }

    private static DocxEditAppliedOperation SetTableRowRepeatAsHeader(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (!TryResolveTableRow(doc, operation, out var row, out var address, out var error))
            return new DocxEditAppliedOperation(operation.Type, false, error);

        var properties = row!.GetFirstChild<TableRowProperties>() ?? row.PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<TableHeader>();
        if (operation.RepeatAsHeader!.Value) properties.AppendChild(new TableHeader());
        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Updated {address} repeatAsHeader={operation.RepeatAsHeader.Value.ToString().ToLowerInvariant()}");
    }

    private static bool TryResolveTableRow(
        WordprocessingDocument doc,
        DocxEditOperation operation,
        out TableRow? row,
        out string address,
        out string error)
    {
        row = null;
        address = string.Empty;
        error = string.Empty;
        if (operation.TableIndex is null || operation.RowIndex is null || operation.RepeatAsHeader is null)
        {
            error = "tableIndex, rowIndex, and repeatAsHeader are required";
            return false;
        }
        if (operation.TableIndex < 0 || operation.RowIndex < 0)
        {
            error = "tableIndex and rowIndex must be non-negative";
            return false;
        }
        if (operation.HeaderIndex is not null && operation.FooterIndex is not null)
        {
            error = "headerIndex and footerIndex are mutually exclusive";
            return false;
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        OpenXmlElement root;
        string scope;
        if (operation.HeaderIndex is { } headerIndex)
        {
            var headers = mainPart.HeaderParts.Where(part => part.Header is not null)
                .OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).ToList();
            if (headerIndex < 0 || headerIndex >= headers.Count)
            {
                error = $"headerIndex {headerIndex} is out of range";
                return false;
            }
            root = headers[headerIndex].Header!;
            scope = $"header[{headerIndex}]";
        }
        else if (operation.FooterIndex is { } footerIndex)
        {
            var footers = mainPart.FooterParts.Where(part => part.Footer is not null)
                .OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).ToList();
            if (footerIndex < 0 || footerIndex >= footers.Count)
            {
                error = $"footerIndex {footerIndex} is out of range";
                return false;
            }
            root = footers[footerIndex].Footer!;
            scope = $"footer[{footerIndex}]";
        }
        else
        {
            root = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
            scope = "body";
        }

        var tables = root.Elements<Table>().ToList();
        if (operation.TableIndex.Value >= tables.Count)
        {
            error = $"tableIndex {operation.TableIndex} is out of range for {scope}; nested tables are inspection-only targets";
            return false;
        }
        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value >= rows.Count)
        {
            error = $"rowIndex {operation.RowIndex} is out of range for {scope}.table[{operation.TableIndex}]";
            return false;
        }

        row = rows[operation.RowIndex.Value];
        address = $"{scope}.table[{operation.TableIndex}].row[{operation.RowIndex}]";
        return true;
    }

    private static DocxEditAppliedOperation ReplaceTable(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.Rows is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and rows are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var sourceTable = tables[operation.TableIndex.Value];
        var replacement = BuildReplacementTable(sourceTable, operation.Rows);
        sourceTable.InsertAfterSelf(replacement);
        sourceTable.Remove();
        return new DocxEditAppliedOperation(operation.Type, true, $"Replaced table[{operation.TableIndex}] with {operation.Rows.Count} row(s)");
    }

    private static Table BuildReplacementTable(Table sourceTable, IReadOnlyList<IReadOnlyList<DocxTableCellInput>> rows)
    {
        var table = new Table();
        var sourceProperties = sourceTable.GetFirstChild<TableProperties>();
        table.AppendChild(sourceProperties is null ? new TableProperties() : (TableProperties)sourceProperties.CloneNode(true));
        EnsureFullWidth(table.GetFirstChild<TableProperties>()!);

        var maxColumns = rows.Count == 0 ? 1 : rows.Max(row => row.Sum(cell => Math.Max(1, cell.GridSpan ?? 1)));
        var sourceGrid = sourceTable.GetFirstChild<TableGrid>();
        if (sourceGrid is not null)
        {
            var grid = (TableGrid)sourceGrid.CloneNode(true);
            while (grid.Elements<GridColumn>().Count() < maxColumns)
            {
                grid.AppendChild(new GridColumn { Width = "1200" });
            }
            table.AppendChild(grid);
        }
        else
        {
            var grid = new TableGrid();
            for (var i = 0; i < maxColumns; i++)
            {
                grid.AppendChild(new GridColumn { Width = "1200" });
            }
            table.AppendChild(grid);
        }

        var templateRows = sourceTable.Elements<TableRow>().ToList();
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var templateRow = templateRows.ElementAtOrDefault(Math.Min(rowIndex, Math.Max(0, templateRows.Count - 1)));
            var row = BuildReplacementRow(templateRow, rows[rowIndex], rowIndex == 0 || rows[rowIndex].Any(cell => cell.Header == true));
            table.AppendChild(row);
        }

        return table;
    }

    private static void EnsureFullWidth(TableProperties properties)
    {
        properties.RemoveAllChildren<TableWidth>();
        properties.PrependChild(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });
    }

    private static TableRow BuildReplacementRow(
        TableRow? templateRow,
        IReadOnlyList<DocxTableCellInput> cells,
        bool isHeader,
        bool inheritCompleteTemplateCellStyle = false)
    {
        var row = new TableRow();
        var templateProperties = templateRow?.GetFirstChild<TableRowProperties>();
        if (templateProperties is not null)
        {
            row.AppendChild((TableRowProperties)templateProperties.CloneNode(true));
        }
        if (isHeader)
        {
            var properties = row.GetFirstChild<TableRowProperties>() ?? row.PrependChild(new TableRowProperties());
            if (!properties.Elements<TableHeader>().Any())
            {
                properties.AppendChild(new TableHeader());
            }
        }

        var templateCells = templateRow?.Elements<TableCell>().ToList() ?? [];
        for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
        {
            var templateCell = templateCells.ElementAtOrDefault(Math.Min(cellIndex, Math.Max(0, templateCells.Count - 1)));
            row.AppendChild(BuildReplacementCell(templateCell, cells[cellIndex], isHeader, inheritCompleteTemplateCellStyle));
        }

        return row;
    }

    private static DocxEditAppliedOperation InsertTableRows(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.Rows is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and rows are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var existingRows = table.Elements<TableRow>().ToList();
        var insertBeforeIndex = operation.RowIndex.Value;
        if (insertBeforeIndex < 0 || insertBeforeIndex > existingRows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {insertBeforeIndex} is out of range");
        }

        var templateRowResult = ResolveTemplateRow(existingRows, operation.TemplateRowIndex, insertBeforeIndex);
        if (!templateRowResult.Valid)
        {
            return new DocxEditAppliedOperation(operation.Type, false, templateRowResult.Error ?? "Invalid template row");
        }

        var templateRow = templateRowResult.Row;
        InsertBuiltRows(
            table,
            existingRows.ElementAtOrDefault(insertBeforeIndex),
            templateRow,
            operation.Rows,
            inheritCompleteTemplateCellStyle: true);
        return new DocxEditAppliedOperation(operation.Type, true, $"Inserted {operation.Rows.Count} row(s) into table[{operation.TableIndex}] before row[{insertBeforeIndex}]");
    }

    private static DocxEditAppliedOperation ReplaceTableRows(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.StartRowIndex is null || operation.EndRowIndex is null || operation.Rows is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, startRowIndex, endRowIndex, and rows are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var existingRows = table.Elements<TableRow>().ToList();
        var start = operation.StartRowIndex.Value;
        var end = operation.EndRowIndex.Value;
        if (start < 0 || end >= existingRows.Count || end < start)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {start} to {end}");
        }

        var templateRowResult = ResolveTemplateRow(existingRows, operation.TemplateRowIndex, start);
        if (!templateRowResult.Valid)
        {
            return new DocxEditAppliedOperation(operation.Type, false, templateRowResult.Error ?? "Invalid template row");
        }

        var templateRow = templateRowResult.Row;
        var anchor = existingRows[start];
        var templateCandidates = existingRows.Skip(start).Take(end - start + 1).ToList();
        InsertBuiltRows(table, anchor, templateRow, operation.Rows, templateCandidates);
        foreach (var row in existingRows.Skip(start).Take(end - start + 1))
        {
            row.Remove();
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Replaced table[{operation.TableIndex}].rows[{start}..{end}] with {operation.Rows.Count} row(s)");
    }

    private static DocxEditAppliedOperation DeleteTableRows(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.StartRowIndex is null || operation.EndRowIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, startRowIndex, and endRowIndex are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var existingRows = table.Elements<TableRow>().ToList();
        var start = operation.StartRowIndex.Value;
        var end = operation.EndRowIndex.Value;
        if (start < 0 || end >= existingRows.Count || end < start)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {start} to {end}");
        }

        foreach (var row in existingRows.Skip(start).Take(end - start + 1))
        {
            row.Remove();
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Deleted table[{operation.TableIndex}].rows[{start}..{end}]");
    }

    private static (bool Valid, TableRow? Row, string? Error) ResolveTemplateRow(IReadOnlyList<TableRow> rows, int? templateRowIndex, int fallbackIndex)
    {
        if (rows.Count == 0)
        {
            return (true, null, null);
        }

        var index = templateRowIndex ?? Math.Clamp(fallbackIndex, 0, rows.Count - 1);
        if (index < 0 || index >= rows.Count)
        {
            return (false, null, $"templateRowIndex {index} is out of range");
        }

        return (true, rows[index], null);
    }

    private static void InsertBuiltRows(
        Table table,
        TableRow? beforeRow,
        TableRow? templateRow,
        IReadOnlyList<IReadOnlyList<DocxTableCellInput>> rows,
        IReadOnlyList<TableRow>? templateCandidates = null,
        bool inheritCompleteTemplateCellStyle = false)
    {
        foreach (var rowInput in rows)
        {
            var rowTemplate = ResolveReplacementRowTemplate(rowInput, templateCandidates, templateRow);
            var row = BuildReplacementRow(
                rowTemplate,
                rowInput,
                rowInput.Any(cell => cell.Header == true),
                inheritCompleteTemplateCellStyle);
            if (beforeRow is null)
            {
                table.AppendChild(row);
            }
            else
            {
                table.InsertBefore(row, beforeRow);
            }
        }
    }

    private static TableRow? ResolveReplacementRowTemplate(
        IReadOnlyList<DocxTableCellInput> rowInput,
        IReadOnlyList<TableRow>? candidates,
        TableRow? fallback)
    {
        if (candidates is not { Count: > 0 })
        {
            return fallback;
        }

        var inputPattern = rowInput.Select(cell => Math.Max(1, cell.GridSpan ?? 1)).ToArray();
        var exact = candidates.FirstOrDefault(row => row.Elements<TableCell>().Select(GetCellGridSpan).SequenceEqual(inputPattern));
        if (exact is not null)
        {
            return exact;
        }

        var sameCellCount = candidates.FirstOrDefault(row => row.Elements<TableCell>().Count() == rowInput.Count);
        return sameCellCount ?? fallback;
    }

    private static DocxEditAppliedOperation InsertTableColumns(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.ColumnIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and columnIndex are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var columnIndex = operation.ColumnIndex.Value;
        var columnCount = operation.ColumnCount ?? 1;
        if (columnIndex < 0 || columnCount < 1)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid columnIndex {columnIndex} or columnCount {columnCount}");
        }

        var table = tables[operation.TableIndex.Value];
        var existingGridWidth = GetTableGridWidth(table);
        if (columnIndex > existingGridWidth)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"columnIndex {columnIndex} is out of range for grid width {existingGridWidth}");
        }

        var templateColumnIndex = operation.TemplateColumnIndex ?? Math.Clamp(columnIndex == existingGridWidth ? columnIndex - 1 : columnIndex, 0, Math.Max(0, existingGridWidth - 1));
        if (existingGridWidth > 0 && (templateColumnIndex < 0 || templateColumnIndex >= existingGridWidth))
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"templateColumnIndex {templateColumnIndex} is out of range for grid width {existingGridWidth}");
        }

        InsertGridColumns(table, columnIndex, templateColumnIndex, columnCount, existingGridWidth);

        var rows = table.Elements<TableRow>().ToList();
        foreach (var row in rows)
        {
            InsertColumnsIntoRow(row, columnIndex, templateColumnIndex, columnCount);
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Inserted {columnCount} column(s) into table[{operation.TableIndex}] before grid column {columnIndex}");
    }

    private static int GetTableGridWidth(Table table)
    {
        var gridColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
        var rowWidth = table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>().Sum(GetCellGridSpan))
            .DefaultIfEmpty(0)
            .Max();
        return Math.Max(gridColumns, rowWidth);
    }

    private static void InsertGridColumns(Table table, int columnIndex, int templateColumnIndex, int columnCount, int existingGridWidth)
    {
        var properties = table.GetFirstChild<TableProperties>();
        if (properties is null)
        {
            properties = new TableProperties();
            table.PrependChild(properties);
        }

        var grid = table.GetFirstChild<TableGrid>();
        if (grid is null)
        {
            grid = new TableGrid();
            for (var i = 0; i < existingGridWidth; i++)
            {
                grid.AppendChild(new GridColumn { Width = "1200" });
            }

            table.InsertAfter(grid, properties);
        }

        var columns = grid.Elements<GridColumn>().ToList();
        var template = columns.ElementAtOrDefault(templateColumnIndex);
        for (var i = 0; i < columnCount; i++)
        {
            var inserted = template is null
                ? new GridColumn { Width = "1200" }
                : (GridColumn)template.CloneNode(true);
            var before = columns.ElementAtOrDefault(columnIndex + i);
            if (before is null)
            {
                grid.AppendChild(inserted);
            }
            else
            {
                grid.InsertBefore(inserted, before);
            }
            columns.Insert(Math.Min(columnIndex + i, columns.Count), inserted);
        }
    }

    private static void InsertColumnsIntoRow(TableRow row, int columnIndex, int templateColumnIndex, int columnCount)
    {
        var cells = row.Elements<TableCell>().ToList();
        var cursor = 0;
        TableCell? insertBefore = null;
        TableCell? templateCell = null;

        foreach (var cell in cells)
        {
            var span = GetCellGridSpan(cell);
            var start = cursor;
            var endExclusive = cursor + span;

            if (templateCell is null && templateColumnIndex >= start && templateColumnIndex < endExclusive)
            {
                templateCell = cell;
            }

            if (columnIndex > start && columnIndex < endExclusive)
            {
                SetCellGridSpan(cell, span + columnCount);
                return;
            }

            if (columnIndex <= start)
            {
                insertBefore = cell;
                break;
            }

            cursor = endExclusive;
        }

        templateCell ??= cells.LastOrDefault();
        for (var i = 0; i < columnCount; i++)
        {
            var cell = BuildInsertedColumnCell(templateCell);
            if (insertBefore is null)
            {
                row.AppendChild(cell);
            }
            else
            {
                row.InsertBefore(cell, insertBefore);
            }
        }
    }

    private static int GetCellGridSpan(TableCell cell)
        => Math.Max(1, cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1);

    private static void SetCellGridSpan(TableCell cell, int span)
    {
        var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
        properties.RemoveAllChildren<GridSpan>();
        if (span > 1)
        {
            properties.AppendChild(new GridSpan { Val = span });
        }
        NormalizeTableCellProperties(properties);
    }

    private static TableCell BuildInsertedColumnCell(TableCell? templateCell)
    {
        var cell = templateCell is null
            ? new TableCell(new TableCellProperties(), new Paragraph(new Run(new Text(string.Empty))))
            : (TableCell)templateCell.CloneNode(true);
        SetCellGridSpan(cell, 1);
        cell.GetFirstChild<TableCellProperties>()?.RemoveAllChildren<VerticalMerge>();
        ReplaceTableCellText(cell, string.Empty);
        return cell;
    }

    private static TableCell BuildReplacementCell(
        TableCell? templateCell,
        DocxTableCellInput input,
        bool rowIsHeader,
        bool inheritCompleteTemplateStyle = false)
    {
        var cell = inheritCompleteTemplateStyle && templateCell is not null
            ? (TableCell)templateCell.CloneNode(true)
            : new TableCell();
        var templateProperties = templateCell?.GetFirstChild<TableCellProperties>();
        if (cell.GetFirstChild<TableCellProperties>() is null && templateProperties is not null)
        {
            cell.AppendChild((TableCellProperties)templateProperties.CloneNode(true));
        }
        else if (cell.GetFirstChild<TableCellProperties>() is null)
        {
            cell.AppendChild(new TableCellProperties());
        }

        var properties = cell.GetFirstChild<TableCellProperties>()!;
        if (input.GridSpan is not null)
        {
            properties.RemoveAllChildren<GridSpan>();
            if (input.GridSpan is > 1)
            {
                properties.AppendChild(new GridSpan { Val = input.GridSpan.Value });
            }
        }

        if (input.VMerge is not null)
        {
            properties.RemoveAllChildren<VerticalMerge>();
            if (input.VMerge is { Length: > 0 } vMergeVal)
            {
                var vmVal = vMergeVal.ToLowerInvariant() == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue;
                properties.AppendChild(new VerticalMerge { Val = vmVal });
            }
        }

        if (input.Shading is { Length: > 0 } hexColor)
        {
            properties.RemoveAllChildren<Shading>();
            properties.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = hexColor });
        }
        NormalizeTableCellProperties(properties);

        if (inheritCompleteTemplateStyle && templateCell is not null)
        {
            if (input.RichText is { Count: > 0 } inheritedRichText)
            {
                ReplaceTableCellRichText(cell, inheritedRichText, input.Alignment);
            }
            else
            {
                ReplaceTableCellText(cell, input.Text ?? string.Empty, input.Alignment);
            }

            if (input.Bold == true || rowIsHeader)
            {
                foreach (var run in cell.Descendants<Run>().Where(run => run.Descendants<Text>().Any()))
                {
                    var runProperties = run.RunProperties ?? run.PrependChild(new RunProperties());
                    runProperties.RemoveAllChildren<Bold>();
                    runProperties.RemoveAllChildren<BoldComplexScript>();
                    runProperties.AppendChild(new Bold());
                    runProperties.AppendChild(new BoldComplexScript());
                }
            }
            return cell;
        }

        var paragraph = CreateParagraphLike(templateCell?.Elements<Paragraph>().FirstOrDefault());
        var paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());

        if (input.Alignment is { Length: > 0 } align)
        {
            var jcVal = align.ToLowerInvariant() switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                _ => JustificationValues.Left
            };
            paragraphProperties.RemoveAllChildren<Justification>();
            paragraphProperties.AppendChild(new Justification { Val = jcVal });
        }

        if (input.RichText is { Count: > 0 } richText)
        {
            foreach (var segment in richText)
            {
                paragraph.AppendChild(CreateRichRunLike(templateCell?.Descendants<Run>().FirstOrDefault(), segment, rowIsHeader));
            }
        }
        else
        {
            var run = CreateStyledRunLike(templateCell?.Descendants<Run>().FirstOrDefault(), input.Text ?? string.Empty);
            if (input.Bold == true || rowIsHeader)
            {
                var runProperties = run.RunProperties ?? run.PrependChild(new RunProperties());
                runProperties.RemoveAllChildren<Bold>();
                runProperties.AppendChild(new Bold());
            }
            paragraph.AppendChild(run);
        }
        cell.AppendChild(paragraph);
        return cell;
    }

    private static void AppendTextWithLineBreaks(Run run, string text)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                run.AppendChild(new Break());
            }
            run.AppendChild(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
        }
    }

    private static void ReplaceTableCellText(
        TableCell cell,
        string replacementText,
        string? alignment = null,
        Run? fallbackRun = null,
        IReadOnlyList<string>? paragraphTexts = null)
    {
        var lines = paragraphTexts?.ToArray()
            ?? replacementText.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');
        var paragraphs = cell.Elements<Paragraph>().ToList();
        if (paragraphs.Count == 0)
        {
            var paragraph = new Paragraph();
            paragraph.AppendChild(CreateStyledRunLike(fallbackRun, string.Empty));
            cell.Append(paragraph);
            paragraphs.Add(paragraph);
        }

        static bool Visible(string value) => !string.IsNullOrWhiteSpace(value);
        static string Comparable(string value) => Visible(value) ? value : string.Empty;
        var targetTexts = paragraphs.Select(Inspector.GetParagraphText).ToList();
        if (lines.Length == targetTexts.Count
            && lines.Select(Comparable).SequenceEqual(targetTexts.Select(Comparable), StringComparer.Ordinal))
        {
            if (!string.IsNullOrWhiteSpace(alignment))
                foreach (var paragraph in paragraphs) ApplyParagraphAlignment(paragraph, alignment);
            return;
        }

        var scores = new int[lines.Length + 1, paragraphs.Count + 1];
        for (var sourceIndex = lines.Length - 1; sourceIndex >= 0; sourceIndex -= 1)
        for (var targetIndex = paragraphs.Count - 1; targetIndex >= 0; targetIndex -= 1)
        {
            var skipSource = scores[sourceIndex + 1, targetIndex];
            var skipTarget = scores[sourceIndex, targetIndex + 1];
            var match = Visible(lines[sourceIndex]) == Visible(targetTexts[targetIndex])
                ? (Visible(lines[sourceIndex]) ? 3 : 2) + scores[sourceIndex + 1, targetIndex + 1]
                : int.MinValue;
            scores[sourceIndex, targetIndex] = Math.Max(match, Math.Max(skipSource, skipTarget));
        }

        var sourceToTarget = Enumerable.Repeat<int?>(null, lines.Length).ToArray();
        var sourceCursor = 0;
        var targetCursor = 0;
        while (sourceCursor < lines.Length && targetCursor < paragraphs.Count)
        {
            var weight = Visible(lines[sourceCursor]) ? 3 : 2;
            if (Visible(lines[sourceCursor]) == Visible(targetTexts[targetCursor])
                && scores[sourceCursor, targetCursor] == weight + scores[sourceCursor + 1, targetCursor + 1])
            {
                sourceToTarget[sourceCursor] = targetCursor;
                sourceCursor += 1;
                targetCursor += 1;
            }
            else if (scores[sourceCursor + 1, targetCursor] >= scores[sourceCursor, targetCursor + 1])
            {
                sourceCursor += 1;
            }
            else
            {
                targetCursor += 1;
            }
        }
        if (!targetTexts.Any(Visible) && lines.Any(Visible))
        {
            var promotedSource = Array.FindIndex(lines, Visible);
            var promotedTarget = paragraphs.FindIndex(paragraph => !paragraph.Descendants<Drawing>().Any());
            if (promotedTarget < 0) promotedTarget = 0;
            for (var index = 0; index < sourceToTarget.Length; index += 1)
                if (sourceToTarget[index] == promotedTarget) sourceToTarget[index] = null;
            sourceToTarget[promotedSource] = promotedTarget;
        }

        var mappedTargets = sourceToTarget.Where(index => index is not null).Select(index => index!.Value).ToHashSet();
        foreach (var (line, sourceIndex) in lines.Select((line, index) => (line, index)))
        {
            if (sourceToTarget[sourceIndex] is not { } mappedTarget) continue;
            ReplaceVisibleParagraphText(
                paragraphs[mappedTarget],
                line,
                fallbackRun,
                promoteBlank: paragraphTexts is null && Visible(line) && !Visible(targetTexts[mappedTarget]));
            if (!string.IsNullOrWhiteSpace(alignment)) ApplyParagraphAlignment(paragraphs[mappedTarget], alignment);
        }
        foreach (var (paragraph, targetIndex) in paragraphs.Select((paragraph, index) => (paragraph, index)))
        {
            if (mappedTargets.Contains(targetIndex) || !Visible(targetTexts[targetIndex])) continue;
            ReplaceVisibleParagraphText(paragraph, string.Empty, fallbackRun);
            if (!string.IsNullOrWhiteSpace(alignment)) ApplyParagraphAlignment(paragraph, alignment);
        }

        OpenXmlElement? lastInserted = null;
        for (var index = 0; index < lines.Length; index += 1)
        {
            if (sourceToTarget[index] is not null)
            {
                lastInserted = paragraphs[sourceToTarget[index]!.Value];
                continue;
            }
            var wantVisible = Visible(lines[index]);
            var template = paragraphs.LastOrDefault(paragraph => Visible(Inspector.GetParagraphText(paragraph)) == wantVisible)
                ?? paragraphs[^1];
            var inserted = (Paragraph)template.CloneNode(true);
            ReplaceVisibleParagraphText(inserted, lines[index], fallbackRun);
            if (!string.IsNullOrWhiteSpace(alignment)) ApplyParagraphAlignment(inserted, alignment);
            var nextTarget = sourceToTarget.Skip(index + 1).FirstOrDefault(candidate => candidate is not null);
            if (nextTarget is not null)
                cell.InsertBefore(inserted, paragraphs[nextTarget.Value]);
            else if (lastInserted is not null)
                cell.InsertAfter(inserted, lastInserted);
            else
                cell.Append(inserted);
            lastInserted = inserted;
        }
    }

    private static void ReplaceVisibleParagraphText(Paragraph paragraph, string text, Run? fallbackRun, bool promoteBlank = false)
    {
        if (string.IsNullOrWhiteSpace(text) && string.IsNullOrWhiteSpace(Inspector.GetParagraphText(paragraph))) return;
        if (promoteBlank)
        {
            foreach (var run in paragraph.Elements<Run>().Where(run => !run.Descendants<Drawing>().Any()).ToList()) run.Remove();
            paragraph.AppendChild(CreateStyledRunLike(
                fallbackRun,
                text,
                preserveEmphasis: false,
                forceBaselineVerticalAlignment: true));
            return;
        }
        var visibleRuns = paragraph.Descendants<Run>()
            .Where(run => run.Descendants<Text>().Any(value => !string.IsNullOrWhiteSpace(value.Text)))
            .ToList();
        var targetRun = visibleRuns.FirstOrDefault()
            ?? paragraph.Descendants<Run>().FirstOrDefault(run =>
                run.Descendants<Text>().Any()
                && !(run.RunProperties?.GetFirstChild<Underline>() is not null
                    && string.IsNullOrWhiteSpace(string.Concat(run.Descendants<Text>().Select(value => value.Text)))));
        if (targetRun is null)
        {
            targetRun = CreateStyledRunLike(fallbackRun, string.Empty);
            paragraph.AppendChild(targetRun);
        }
        if (!visibleRuns.Contains(targetRun)) visibleRuns.Add(targetRun);

        foreach (var run in visibleRuns)
        {
            foreach (var value in run.Descendants<Text>().ToList()) value.Remove();
            foreach (var lineBreak in run.Descendants<Break>().ToList()) lineBreak.Remove();
        }
        AppendTextWithLineBreaks(targetRun, text);
    }

    private static DocxEditAppliedOperation ReplaceTableCellRichText(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || operation.RichText is not { Count: > 0 })
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and richText are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        var fallbackRun = FindNearestTableRun(rows, operation.RowIndex.Value, operation.CellIndex.Value);
        ReplaceTableCellRichText(cells[operation.CellIndex.Value], operation.RichText, operation.Alignment, fallbackRun);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated rich text in table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}]");
    }

    private static void ReplaceTableCellRichText(TableCell cell, IReadOnlyList<DocxRichTextSegment> segments, string? alignment = null, Run? fallbackRun = null)
    {
        var graphicParagraphs = CaptureGraphicParagraphs(cell);
        var hasTextBearingRun = HasTextBearingRun(cell);
        var ownRun = FindCellStyleTemplateRun(cell);
        var firstRun = ownRun ?? fallbackRun;
        var firstParagraph = cell.Elements<Paragraph>().FirstOrDefault();
        cell.RemoveAllChildren<Paragraph>();
        var paragraph = CreateParagraphLike(firstParagraph);
        foreach (var segment in segments)
        {
            paragraph.AppendChild(CreateRichRunLike(
                firstRun,
                segment,
                preserveEmphasis: hasTextBearingRun,
                forceBaselineVerticalAlignment: !hasTextBearingRun));
        }
        if (!string.IsNullOrWhiteSpace(alignment))
        {
            ApplyParagraphAlignment(paragraph, alignment);
        }
        cell.Append(paragraph);
        foreach (var graphic in graphicParagraphs) cell.Append(graphic);
    }

    private static Run? FindCellStyleTemplateRun(TableCell cell)
    {
        var paragraph = cell.Elements<Paragraph>().FirstOrDefault();
        var run = cell.Descendants<Run>().FirstOrDefault(candidate => candidate.Descendants<Text>().Any(text => HasVisibleText(text.Text)))
            ?? cell.Descendants<Run>().FirstOrDefault();
        if (run is not null)
        {
            return run;
        }
        var paragraphMark = paragraph?.ParagraphProperties?.ParagraphMarkRunProperties;
        if (paragraphMark is null)
        {
            return null;
        }

        var effectiveProperties = new RunProperties();
        OverlayRunProperties(effectiveProperties, paragraphMark.ChildElements);
        NormalizeRunProperties(effectiveProperties);
        return new Run(effectiveProperties);
    }

    private static bool HasTextBearingRun(TableCell cell)
        => cell.Descendants<Run>()
            .Any(candidate => candidate.Descendants<Text>().Any(text => HasVisibleText(text.Text)));

    private static bool HasVisibleText(string? text)
        => text?.Any(character => !char.IsWhiteSpace(character) && character is not ('\u200B' or '\u200C' or '\u200D' or '\u2060' or '\uFEFF')) == true;

    private static void OverlayRunProperties(RunProperties target, IEnumerable<OpenXmlElement> source)
    {
        foreach (var property in source)
        {
            foreach (var existing in target.ChildElements.Where(candidate => candidate.GetType() == property.GetType()).ToList())
            {
                existing.Remove();
            }
            target.Append(property.CloneNode(true));
        }
    }

    private static List<Paragraph> CaptureGraphicParagraphs(TableCell cell)
    {
        var result = new List<Paragraph>();
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            var graphicRuns = paragraph.Elements<Run>()
                .Where(run => run.Descendants<Drawing>().Any() || run.Descendants<Picture>().Any())
                .Select(run => (Run)run.CloneNode(true))
                .ToList();
            if (graphicRuns.Count == 0) continue;
            var clone = CreateParagraphLike(paragraph);
            foreach (var run in graphicRuns) clone.AppendChild(run);
            result.Add(clone);
        }
        return result;
    }

    private static Run? FindNearestTableRun(IReadOnlyList<TableRow> rows, int rowIndex, int cellIndex)
    {
        var exactRow = FindNearestRunInRow(rows[rowIndex], cellIndex);
        if (exactRow is not null)
        {
            return exactRow;
        }

        for (var offset = 1; offset < rows.Count; offset++)
        {
            var previous = rowIndex - offset;
            if (previous >= 0)
            {
                var run = FindNearestRunInRow(rows[previous], cellIndex);
                if (run is not null)
                {
                    return run;
                }
            }

            var next = rowIndex + offset;
            if (next < rows.Count)
            {
                var run = FindNearestRunInRow(rows[next], cellIndex);
                if (run is not null)
                {
                    return run;
                }
            }
        }

        return null;
    }

    private static Run? FindNearestRunInRow(TableRow row, int cellIndex)
    {
        var cells = row.Elements<TableCell>().ToList();
        if (cells.Count == 0)
        {
            return null;
        }

        if (cellIndex >= 0 && cellIndex < cells.Count)
        {
            var run = cells[cellIndex].Descendants<Run>().FirstOrDefault();
            if (run is not null)
            {
                return run;
            }
        }

        for (var offset = 1; offset < cells.Count; offset++)
        {
            var previous = cellIndex - offset;
            if (previous >= 0)
            {
                var run = cells[previous].Descendants<Run>().FirstOrDefault();
                if (run is not null)
                {
                    return run;
                }
            }

            var next = cellIndex + offset;
            if (next < cells.Count)
            {
                var run = cells[next].Descendants<Run>().FirstOrDefault();
                if (run is not null)
                {
                    return run;
                }
            }
        }

        return null;
    }

    private static Paragraph CreateParagraphLike(Paragraph? templateParagraph)
    {
        var paragraph = new Paragraph();
        var templateProperties = templateParagraph?.GetFirstChild<ParagraphProperties>();
        if (templateProperties is not null)
        {
            paragraph.AppendChild((ParagraphProperties)templateProperties.CloneNode(true));
        }
        return paragraph;
    }

    private static void ApplyCellAlignment(TableCell cell, string alignment)
    {
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            ApplyParagraphAlignment(paragraph, alignment);
        }
    }

    private static void ApplyParagraphAlignment(Paragraph paragraph, string alignment)
    {
        var properties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
        properties.RemoveAllChildren<Justification>();
        properties.AppendChild(new Justification
        {
            Val = alignment.ToLowerInvariant() switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                _ => JustificationValues.Left,
            },
        });
    }

    private static Run CreateStyledRunLike(
        Run? templateRun,
        string text,
        bool preserveEmphasis = true,
        bool forceBaselineVerticalAlignment = false)
    {
        var run = new Run();
        if (templateRun?.RunProperties is not null)
        {
            run.RunProperties = (RunProperties)templateRun.RunProperties.CloneNode(true);
            if (!preserveEmphasis)
            {
                RemoveEmphasis(run.RunProperties);
            }
            if (forceBaselineVerticalAlignment)
            {
                ForceBaselineVerticalAlignment(run.RunProperties);
            }
            NormalizeRunProperties(run.RunProperties);
        }
        else if (forceBaselineVerticalAlignment)
        {
            run.RunProperties = new RunProperties();
            ForceBaselineVerticalAlignment(run.RunProperties);
            NormalizeRunProperties(run.RunProperties);
        }
        AppendTextWithLineBreaks(run, text);
        return run;
    }

    private static Run CreateRichRunLike(
        Run? templateRun,
        DocxRichTextSegment segment,
        bool forceBold = false,
        bool preserveEmphasis = true,
        bool forceBaselineVerticalAlignment = false)
    {
        var run = new Run();
        if (templateRun?.RunProperties is not null)
        {
            run.RunProperties = (RunProperties)templateRun.RunProperties.CloneNode(true);
        }

        var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
        RemoveTextFill(properties);
        if (!preserveEmphasis)
        {
            RemoveEmphasis(properties);
        }
        if (forceBaselineVerticalAlignment)
        {
            ForceBaselineVerticalAlignment(properties);
        }

        if (forceBold || segment.Bold == true)
        {
            properties.RemoveAllChildren<Bold>();
            properties.RemoveAllChildren<BoldComplexScript>();
            properties.AppendChild(new Bold());
            properties.AppendChild(new BoldComplexScript());
        }
        else if (segment.Bold == false)
        {
            properties.RemoveAllChildren<Bold>();
            properties.RemoveAllChildren<BoldComplexScript>();
            properties.AppendChild(new Bold { Val = false });
            properties.AppendChild(new BoldComplexScript { Val = false });
        }

        if (!string.IsNullOrWhiteSpace(segment.Color))
        {
            properties.RemoveAllChildren<Color>();
            properties.AppendChild(new Color { Val = segment.Color });
        }

        if (!string.IsNullOrWhiteSpace(segment.FontName))
        {
            properties.RemoveAllChildren<RunFonts>();
            properties.PrependChild(new RunFonts
            {
                Ascii = segment.FontName,
                HighAnsi = segment.FontName,
            });
        }

        if (segment.Underline == true)
        {
            properties.RemoveAllChildren<Underline>();
            properties.AppendChild(new Underline { Val = UnderlineValues.Single });
        }
        else if (segment.Underline == false)
        {
            properties.RemoveAllChildren<Underline>();
        }

        if (segment.Italic == true)
        {
            properties.RemoveAllChildren<Italic>();
            properties.RemoveAllChildren<ItalicComplexScript>();
            properties.AppendChild(new Italic());
            properties.AppendChild(new ItalicComplexScript());
        }
        else if (segment.Italic == false)
        {
            properties.RemoveAllChildren<Italic>();
            properties.RemoveAllChildren<ItalicComplexScript>();
        }

        if (!string.IsNullOrWhiteSpace(segment.VerticalAlignment))
        {
            var alignment = segment.VerticalAlignment switch
            {
                "baseline" => VerticalPositionValues.Baseline,
                "superscript" => VerticalPositionValues.Superscript,
                "subscript" => VerticalPositionValues.Subscript,
                _ => throw new InvalidOperationException("rich-text-vertical-alignment-invalid")
            };
            properties.RemoveAllChildren<VerticalTextAlignment>();
            properties.AppendChild(new VerticalTextAlignment { Val = alignment });
        }

        NormalizeRunProperties(properties);
        AppendTextWithLineBreaks(run, segment.Text);
        return run;
    }

    private static void RemoveEmphasis(RunProperties properties)
    {
        properties.RemoveAllChildren<Bold>();
        properties.RemoveAllChildren<BoldComplexScript>();
        properties.RemoveAllChildren<Italic>();
        properties.RemoveAllChildren<ItalicComplexScript>();
    }

    private static void ForceBaselineVerticalAlignment(RunProperties properties)
    {
        properties.RemoveAllChildren<VerticalTextAlignment>();
        properties.AppendChild(new VerticalTextAlignment { Val = VerticalPositionValues.Baseline });
    }

    private static void RemoveTextFill(RunProperties properties)
    {
        foreach (var textFill in properties.Elements<W14.FillTextEffect>().ToList())
        {
            textFill.Remove();
        }
    }

    private static DocxEditAppliedOperation MergeTableCells(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex is required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var rows = table.Elements<TableRow>().ToList();

        if (operation.RowIndex is not null)
        {
            var rowIndex = operation.RowIndex.Value;
            if (rowIndex < 0 || rowIndex >= rows.Count)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {rowIndex} is out of range");
            }

            var row = rows[rowIndex];
            var cells = row.Elements<TableCell>().ToList();
            var startCellIndex = operation.StartCellIndex ?? 0;
            var endCellIndex = operation.EndCellIndex ?? (cells.Count - 1);

            if (startCellIndex < 0 || endCellIndex >= cells.Count || endCellIndex <= startCellIndex)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"Invalid cell range {startCellIndex} to {endCellIndex}");
            }

            var selected = cells.Skip(startCellIndex).Take(endCellIndex - startCellIndex + 1).ToList();
            var totalSpan = selected.Sum(cell => {
                var span = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>();
                if (span?.Val?.Value is int val) return val;
                return 1;
            });

            var properties = selected[0].GetFirstChild<TableCellProperties>() ?? selected[0].PrependChild(new TableCellProperties());
            properties.RemoveAllChildren<GridSpan>();
            if (totalSpan > 1)
            {
                properties.AppendChild(new GridSpan { Val = totalSpan });
            }
            NormalizeTableCellProperties(properties);

            foreach (var cell in selected.Skip(1))
            {
                row.RemoveChild(cell);
            }

            return new DocxEditAppliedOperation(operation.Type, true, $"Merged table[{operation.TableIndex}].row[{rowIndex}].cells[{startCellIndex}..{endCellIndex}]");
        }
        else if (operation.CellIndex is not null || operation.GridColumn is not null)
        {
            if (operation.CellIndex is not null && operation.GridColumn is not null)
            {
                return new DocxEditAppliedOperation(operation.Type, false, "cellIndex and gridColumn are mutually exclusive");
            }
            var cellIndex = operation.CellIndex ?? -1;
            var startRowIndex = operation.StartRowIndex ?? 0;
            var endRowIndex = operation.EndRowIndex ?? (rows.Count - 1);

            if (startRowIndex < 0 || endRowIndex >= rows.Count || endRowIndex <= startRowIndex)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {startRowIndex} to {endRowIndex}");
            }

            var selectedCells = new List<(TableCell Cell, int CellIndex)>();
            for (var rIdx = startRowIndex; rIdx <= endRowIndex; rIdx++)
            {
                var rCells = rows[rIdx].Elements<TableCell>().ToList();
                var selected = operation.GridColumn is not null
                    ? CellAtGridColumn(rows[rIdx], operation.GridColumn.Value)
                    : cellIndex >= 0 && cellIndex < rCells.Count ? (rCells[cellIndex], cellIndex) : ((TableCell, int)?)null;
                if (selected is null)
                {
                    var selector = operation.GridColumn is not null ? $"gridColumn {operation.GridColumn}" : $"cellIndex {cellIndex}";
                    return new DocxEditAppliedOperation(operation.Type, false, $"{selector} is out of range in row {rIdx}");
                }
                selectedCells.Add(selected.Value);
            }

            for (var rIdx = startRowIndex; rIdx <= endRowIndex; rIdx++)
            {
                var selected = selectedCells[rIdx - startRowIndex];
                var cell = selected.Cell;
                var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
                if (rIdx == startRowIndex && IsVerticalMergeContinuation(properties))
                {
                    var ownerCell = operation.GridColumn is not null
                        ? FindPreviousVerticalMergeOwnerByGridColumn(rows, startRowIndex, operation.GridColumn.Value)
                        : FindPreviousVerticalMergeOwner(rows, startRowIndex, selected.CellIndex);
                    CopyMissingParagraphProperties(ownerCell, cell);
                }
                properties.RemoveAllChildren<VerticalMerge>();
                var mergeValue = rIdx == startRowIndex ? MergedCellValues.Restart : MergedCellValues.Continue;
                properties.AppendChild(new VerticalMerge { Val = mergeValue });
                NormalizeTableCellProperties(properties);
                if (rIdx != startRowIndex)
                {
                    cell.RemoveAllChildren<Paragraph>();
                    cell.AppendChild(new Paragraph());
                }
            }

            var detailSelector = operation.GridColumn is not null ? $"gridColumn[{operation.GridColumn}]" : $"cell[{cellIndex}]";
            return new DocxEditAppliedOperation(operation.Type, true, $"Vertically merged table[{operation.TableIndex}].{detailSelector}.rows[{startRowIndex}..{endRowIndex}]");
        }

        return new DocxEditAppliedOperation(operation.Type, false, "Either rowIndex (horizontal) or cellIndex/gridColumn (vertical) must be specified for merge");
    }

    private static (TableCell Cell, int CellIndex)? CellAtGridColumn(TableRow row, int gridColumn)
    {
        if (gridColumn < 0) return null;
        var cursor = 0;
        var cells = row.Elements<TableCell>().ToList();
        for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
        {
            var span = GetCellGridSpan(cells[cellIndex]);
            if (gridColumn >= cursor && gridColumn < cursor + span) return (cells[cellIndex], cellIndex);
            cursor += span;
        }
        return null;
    }

    private static TableCell? FindPreviousVerticalMergeOwnerByGridColumn(IReadOnlyList<TableRow> rows, int startRowIndex, int gridColumn)
    {
        for (var rowIndex = startRowIndex - 1; rowIndex >= 0; rowIndex--)
        {
            var selected = CellAtGridColumn(rows[rowIndex], gridColumn);
            if (selected is null) return null;
            var merge = selected.Value.Cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>();
            if (merge?.Val?.Value == MergedCellValues.Restart) return selected.Value.Cell;
            if (merge is null) return null;
        }
        return null;
    }

    private static bool IsVerticalMergeContinuation(TableCellProperties properties)
    {
        var merge = properties.GetFirstChild<VerticalMerge>();
        return merge is not null && (merge.Val is null || merge.Val.Value == MergedCellValues.Continue);
    }

    private static TableCell? FindPreviousVerticalMergeOwner(IReadOnlyList<TableRow> rows, int startRowIndex, int cellIndex)
    {
        for (var rowIndex = startRowIndex - 1; rowIndex >= 0; rowIndex--)
        {
            var cells = rows[rowIndex].Elements<TableCell>().ToList();
            if (cellIndex >= cells.Count)
            {
                continue;
            }

            var cell = cells[cellIndex];
            var properties = cell.GetFirstChild<TableCellProperties>();
            if (!IsVerticalMergeContinuation(properties ?? new TableCellProperties()))
            {
                return cell;
            }
        }

        return null;
    }

    private static void CopyMissingParagraphProperties(TableCell? sourceCell, TableCell targetCell)
    {
        var sourceProperties = sourceCell?.Elements<Paragraph>().FirstOrDefault()?.GetFirstChild<ParagraphProperties>();
        if (sourceProperties is null)
        {
            return;
        }

        var targetParagraph = targetCell.Elements<Paragraph>().FirstOrDefault();
        if (targetParagraph is null)
        {
            targetParagraph = targetCell.AppendChild(new Paragraph());
        }

        var targetProperties = targetParagraph.GetFirstChild<ParagraphProperties>() ?? targetParagraph.PrependChild(new ParagraphProperties());
        foreach (var sourceProperty in sourceProperties.ChildElements)
        {
            if (targetProperties.ChildElements.Any(existing => existing.GetType() == sourceProperty.GetType()))
            {
                continue;
            }

            targetProperties.AppendChild(sourceProperty.CloneNode(true));
        }
    }

    private static DocxEditAppliedOperation UnmergeTableRowHorizontalCells(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and cellIndex are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var rows = table.Elements<TableRow>().ToList();
        var rowIndex = operation.RowIndex.Value;
        if (rowIndex < 0 || rowIndex >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {rowIndex} is out of range");
        }

        var row = rows[rowIndex];
        var cells = row.Elements<TableCell>().ToList();
        var cellIndex = operation.CellIndex.Value;
        if (cellIndex < 0 || cellIndex >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {cellIndex} is out of range");
        }

        var cell = cells[cellIndex];
        var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
        var span = properties.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
        if (span <= 1)
        {
            return new DocxEditAppliedOperation(operation.Type, true, $"Cell table[{operation.TableIndex}].row[{rowIndex}].cell[{cellIndex}] is not horizontally merged");
        }

        var gridStart = cells.Take(cellIndex).Sum(GetCellGridSpan);
        var splitWidths = GetUnmergedCellWidths(table, row, gridStart, span);

        properties.RemoveAllChildren<GridSpan>();
        SetCellWidth(properties, splitWidths[0]);
        NormalizeTableCellProperties(properties);

        for (var i = 1; i < span; i++)
        {
            var newCell = (TableCell)cell.CloneNode(true);
            foreach (var child in newCell.ChildElements.Where(child => child is not TableCellProperties).ToList())
            {
                child.Remove();
            }
            newCell.AppendChild(new Paragraph());
            var newProperties = newCell.GetFirstChild<TableCellProperties>() ?? newCell.PrependChild(new TableCellProperties());
            newProperties.RemoveAllChildren<GridSpan>();
            SetCellWidth(newProperties, splitWidths[i]);
            NormalizeTableCellProperties(newProperties);
            row.InsertAfter(newCell, cell);
            cell = newCell;
        }

        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Unmerged horizontal cell in table[{operation.TableIndex}].row[{rowIndex}].cell[{cellIndex}], expanded {span} grid columns");
    }

    private static IReadOnlyList<string> GetUnmergedCellWidths(Table table, TableRow currentRow, int startColumn, int count)
    {
        var visibleReference = FindVisibleCellWidthsForGridRange(table, currentRow, startColumn, count);
        if (visibleReference is not null)
        {
            return visibleReference;
        }

        return GetTableGridWidths(table, startColumn, count);
    }

    private static IReadOnlyList<string>? FindVisibleCellWidthsForGridRange(Table table, TableRow currentRow, int startColumn, int count)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            if (ReferenceEquals(row, currentRow))
            {
                continue;
            }

            var cursor = 0;
            var widths = new List<string>();
            foreach (var cell in row.Elements<TableCell>())
            {
                var span = GetCellGridSpan(cell);
                if (cursor >= startColumn && cursor + span <= startColumn + count)
                {
                    if (span != 1)
                    {
                        widths.Clear();
                        break;
                    }

                    var width = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>()?.Width?.Value;
                    widths.Add(string.IsNullOrWhiteSpace(width) ? GetTableGridWidths(table, cursor, 1)[0] : width!);
                }
                else if (cursor < startColumn + count && cursor + span > startColumn)
                {
                    widths.Clear();
                    break;
                }

                cursor += span;
                if (cursor >= startColumn + count)
                {
                    break;
                }
            }

            if (widths.Count == count)
            {
                return widths;
            }
        }

        return null;
    }

    private static IReadOnlyList<string> GetTableGridWidths(Table table, int startColumn, int count)
    {
        var columns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList() ?? [];
        var widths = new List<string>();
        for (var offset = 0; offset < count; offset++)
        {
            var column = columns.ElementAtOrDefault(startColumn + offset);
            widths.Add(string.IsNullOrWhiteSpace(column?.Width) ? "1200" : column.Width!);
        }
        return widths;
    }

    private static void SetCellWidth(TableCellProperties properties, string width)
    {
        properties.RemoveAllChildren<TableCellWidth>();
        properties.PrependChild(new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa });
    }

    private static DocxEditAppliedOperation UnmergeTableColumnVerticalCells(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.CellIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and cellIndex are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var rows = table.Elements<TableRow>().ToList();
        var cellIndex = operation.CellIndex.Value;
        var startRowIndex = operation.StartRowIndex ?? 0;
        var endRowIndex = operation.EndRowIndex ?? (rows.Count - 1);

        if (startRowIndex < 0 || endRowIndex >= rows.Count || endRowIndex < startRowIndex)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {startRowIndex} to {endRowIndex}");
        }

        List<OpenXmlElement>? latestVisibleContent = null;
        var changed = 0;
        for (var rIdx = startRowIndex; rIdx <= endRowIndex; rIdx++)
        {
            var cells = rows[rIdx].Elements<TableCell>().ToList();
            if (cellIndex < 0 || cellIndex >= cells.Count)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {cellIndex} is out of range in row {rIdx}");
            }

            var cell = cells[cellIndex];
            var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
            var verticalMerge = properties.GetFirstChild<VerticalMerge>();
            var isContinuation = verticalMerge is not null
                && (verticalMerge.Val is null || verticalMerge.Val.Value == MergedCellValues.Continue);

            if (!isContinuation)
            {
                latestVisibleContent = cell.Elements<Paragraph>()
                    .Select(p => p.CloneNode(true))
                    .ToList();
            }

            if (verticalMerge is not null)
            {
                properties.RemoveAllChildren<VerticalMerge>();
                NormalizeTableCellProperties(properties);
                changed++;
            }

            if (isContinuation && latestVisibleContent is { Count: > 0 })
            {
                cell.RemoveAllChildren<Paragraph>();
                foreach (var paragraph in latestVisibleContent)
                {
                    cell.Append(paragraph.CloneNode(true));
                }
            }
        }

        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Unmerged vertical cells in table[{operation.TableIndex}].cell[{cellIndex}].rows[{startRowIndex}..{endRowIndex}], removed {changed} vMerge marker(s)");
    }
}
