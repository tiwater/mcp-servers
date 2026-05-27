using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

public static class Editor
{
    private static readonly Regex NumericTextPattern = new(@"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$", RegexOptions.Compiled);
    private static readonly Regex PercentTextPattern = new(@"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$", RegexOptions.Compiled);
    private static readonly Regex FormulaCellReferencePattern = new(@"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)", RegexOptions.Compiled);

    public static int RunEdit(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("edit requires <input.xlsx> <operations.json> <output.xlsx>");
        }

        var input = Path.GetFullPath(args[0]);
        var operationsPath = Path.GetFullPath(args[1]);
        var output = Path.GetFullPath(args[2]);
        var request = LoadOperations(operationsPath);
        var result = Apply(input, output, request.Operations);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return 0;
    }

    public static XlsxEditResult Apply(string input, string output, IReadOnlyList<XlsxEditOperation> operations)
    {
        File.Copy(input, output, overwrite: true);
        using var spreadsheet = SpreadsheetDocument.Open(output, true);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part not found.");
        var applied = new List<XlsxEditAppliedOperation>();

        foreach (var operation in operations)
        {
            applied.Add(ApplyOperation(workbookPart, operation));
        }

        workbookPart.Workbook.Save();
        spreadsheet.Save();
        return new XlsxEditResult(Path.GetFullPath(input), Path.GetFullPath(output), applied);
    }

    private static XlsxEditDocument LoadOperations(string path)
    {
        var json = File.ReadAllText(path);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new XlsxEditDocument([]);
        }

        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind == JsonValueKind.Array)
        {
            var ops = JsonSerializer.Deserialize<List<XlsxEditOperation>>(json, Json.Options) ?? [];
            return new XlsxEditDocument(ops);
        }

        return JsonSerializer.Deserialize<XlsxEditDocument>(json, Json.Options) ?? new XlsxEditDocument([]);
    }

    private static XlsxEditAppliedOperation ApplyOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        return operation.Type switch
        {
            "setCellValue" => SetCellValueOperation(workbookPart, operation),
            "setRangeValues" => SetRangeValuesOperation(workbookPart, operation),
            "insertRows" => InsertRowsOperation(workbookPart, operation),
            "copyRow" => CopyRowOperation(workbookPart, operation),
            "expandSectionRows" => ExpandSectionRowsOperation(workbookPart, operation),
            _ => new XlsxEditAppliedOperation(operation.Type, false, $"Unknown operation type: {operation.Type}"),
        };
    }

    private static XlsxEditAppliedOperation SetCellValueOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || string.IsNullOrWhiteSpace(operation.Cell) || operation.Value is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, cell, and value are required");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, error!);
        }

        var cell = GetOrCreateCell(worksheetPart, operation.Cell);
        SetCellValue(cell, operation.Value, workbookPart, operation.ValueType);
        if (operation.Bold.HasValue)
        {
            ApplyCellBold(workbookPart, cell, operation.Bold.Value);
        }
        worksheetPart.Worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Updated {operation.Sheet}!{operation.Cell}");
    }

    private static XlsxEditAppliedOperation SetRangeValuesOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || string.IsNullOrWhiteSpace(operation.StartCell) || operation.Values is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, startCell, and values are required");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, error!);
        }

        var (startColumn, startRow) = ParseCellReference(operation.StartCell);
        for (var rowOffset = 0; rowOffset < operation.Values.Count; rowOffset++)
        {
            var rowValues = operation.Values[rowOffset];
            for (var colOffset = 0; colOffset < rowValues.Count; colOffset++)
            {
                var cellReference = GetCellReference(startColumn + colOffset, startRow + rowOffset);
                var cell = GetOrCreateCell(worksheetPart, cellReference);
                SetCellValue(cell, rowValues[colOffset], workbookPart, operation.ValueType);
            }
        }

        worksheetPart.Worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Updated range from {operation.Sheet}!{operation.StartCell}");
    }

    private static XlsxEditAppliedOperation InsertRowsOperation(WorkbookPart workbookPart, XlsxEditOperation operation, bool preserveMergedRanges = true)
    {
        var startRow = operation.StartRow ?? operation.TargetRow;
        var legacyTemplateSourceRow = operation.StartRow is null ? operation.SourceRow : null;
        if (string.IsNullOrWhiteSpace(operation.Sheet) || startRow is null || operation.Count is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, startRow, and count are required");
        }

        if (startRow.Value < 1 || operation.Count.Value < 1)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "startRow and count must be positive");
        }

        if (legacyTemplateSourceRow is not null && legacyTemplateSourceRow.Value < 1)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sourceRow must be positive");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, error!);
        }

        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>();
        Row? legacyTemplateRow = null;
        if (legacyTemplateSourceRow is not null)
        {
            legacyTemplateRow = sheetData?.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == legacyTemplateSourceRow.Value);
            if (legacyTemplateRow is null)
            {
                return new XlsxEditAppliedOperation(operation.Type, false, $"Source row not found: {legacyTemplateSourceRow.Value}");
            }

            legacyTemplateRow = (Row)legacyTemplateRow.CloneNode(true);
        }

        if (sheetData is not null)
        {
            foreach (var row in sheetData.Elements<Row>()
                         .Where(row => row.RowIndex?.Value >= startRow.Value)
                         .OrderByDescending(row => row.RowIndex!.Value)
                         .ToList())
            {
                var targetRow = row.RowIndex!.Value + (uint)operation.Count.Value;
                row.RowIndex = targetRow;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value is string reference)
                    {
                        cell.CellReference = ShiftCellReference(reference, operation.Count.Value);
                    }
                }
            }
        }

        ShiftWorksheetDimension(worksheet, startRow.Value, operation.Count.Value);
        if (preserveMergedRanges)
        {
            ShiftMergedRanges(worksheet, startRow.Value, operation.Count.Value);
        }
        ShiftFormulasForInsertedRows(workbookPart, operation.Sheet, startRow.Value, operation.Count.Value);

        if (legacyTemplateRow is not null && sheetData is not null && legacyTemplateSourceRow is not null)
        {
            for (var offset = 0; offset < operation.Count.Value; offset++)
            {
                var targetRow = startRow.Value + offset;
                if (!TryCopyRow(
                        legacyTemplateRow,
                        sheetData,
                        worksheet,
                        legacyTemplateSourceRow.Value,
                        targetRow,
                        preserveStyle: true,
                        preserveFormulas: true,
                        translateFormulas: true,
                        out var copyError))
                {
                    return new XlsxEditAppliedOperation(operation.Type, false, copyError!, operation.Sheet);
                }
            }
        }

        worksheet.Save();
        var changedRange = $"{startRow}:{startRow + operation.Count - 1}";
        return new XlsxEditAppliedOperation(operation.Type, true, $"Inserted {operation.Count} row(s) at {operation.Sheet}!{startRow}", operation.Sheet, changedRange);
    }

    private static XlsxEditAppliedOperation CopyRowOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || operation.SourceRow is null || operation.TargetRow is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, sourceRow, and targetRow are required");
        }

        if (operation.SourceRow.Value < 1 || operation.TargetRow.Value < 1)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sourceRow and targetRow must be positive");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, error!);
        }

        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
        if (!TryCopyRow(
                sheetData,
                worksheet,
                operation.SourceRow.Value,
                operation.TargetRow.Value,
                preserveStyle: true,
                preserveFormulas: true,
                translateFormulas: operation.TranslateFormulas == true,
                out var copyError))
        {
            return new XlsxEditAppliedOperation(operation.Type, false, copyError!);
        }

        worksheet.Save();

        var changedRange = $"{operation.TargetRow}:{operation.TargetRow}";
        return new XlsxEditAppliedOperation(operation.Type, true, $"Copied row {operation.SourceRow} to {operation.Sheet}!{operation.TargetRow}", operation.Sheet, changedRange);
    }

    private static XlsxEditAppliedOperation ExpandSectionRowsOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || operation.AnchorText is null || operation.ExampleRows is null || operation.TargetRows is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, anchorText, exampleRows, and targetRows are required");
        }

        if (operation.ExampleRows.Value < 1 || operation.TargetRows.Value < 1)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "exampleRows and targetRows must be positive");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, error!);
        }

        var worksheet = worksheetPart.Worksheet;
        var sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
        var anchorCell = FindVisibleTextCell(workbookPart, worksheet, operation.AnchorText);
        if (anchorCell?.CellReference?.Value is not string anchorReference)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, $"Anchor text not found on sheet {operation.Sheet}: {operation.AnchorText}");
        }

        var (_, sectionHeaderRow) = ParseCellReference(anchorReference);
        var firstExampleRow = sectionHeaderRow + 1;
        var existingRows = operation.ExampleRows.Value;
        var targetRows = operation.TargetRows.Value;
        var changedRange = $"{firstExampleRow}:{firstExampleRow + targetRows - 1}";

        for (var sourceRowIndex = firstExampleRow; sourceRowIndex < firstExampleRow + existingRows; sourceRowIndex++)
        {
            if (!sheetData.Elements<Row>().Any(row => row.RowIndex?.Value == sourceRowIndex))
            {
                return new XlsxEditAppliedOperation(operation.Type, false, $"Example row not found: {sourceRowIndex}", operation.Sheet);
            }
        }

        if (targetRows <= existingRows)
        {
            return new XlsxEditAppliedOperation(
                operation.Type,
                true,
                $"Section already has {existingRows} example row(s); shrink to {targetRows} is unsupported",
                operation.Sheet,
                ChangedRange: null,
                Warnings: [$"Shrink unsupported for expandSectionRows; existing rows were left unchanged."]);
        }

        var preserveStyle = operation.PreserveStyle != false;
        var preserveFormulas = operation.PreserveFormulas != false;
        var preserveMergedRanges = operation.PreserveMergedRanges != false;
        var rowsToInsert = targetRows - existingRows;
        var firstInsertedRow = firstExampleRow + existingRows;
        var exemplarRows = sheetData.Elements<Row>()
            .Where(row => row.RowIndex?.Value >= firstExampleRow && row.RowIndex?.Value < firstExampleRow + existingRows)
            .ToDictionary(row => (int)row.RowIndex!.Value, row => (Row)row.CloneNode(true));
        var exemplarMergedRanges = preserveMergedRanges
            ? GetHorizontalMergedRangesOnRows(worksheet, firstExampleRow, existingRows).ToList()
            : [];

        for (var generatedRowIndex = firstInsertedRow; generatedRowIndex < firstExampleRow + targetRows; generatedRowIndex++)
        {
            var sourceRowIndex = firstExampleRow + ((generatedRowIndex - firstExampleRow) % existingRows);
            if (!CanCopyRow(exemplarRows[sourceRowIndex], sourceRowIndex, generatedRowIndex, preserveFormulas, out var copyError))
            {
                return new XlsxEditAppliedOperation(operation.Type, false, copyError!, operation.Sheet);
            }
        }

        var insertOperation = InsertRowsOperation(workbookPart, operation with
        {
            Type = "insertRows",
            StartRow = firstInsertedRow,
            Count = rowsToInsert
        }, preserveMergedRanges);
        if (!insertOperation.Applied)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, insertOperation.Detail, operation.Sheet);
        }

        for (var generatedRowIndex = firstInsertedRow; generatedRowIndex < firstExampleRow + targetRows; generatedRowIndex++)
        {
            var sourceRowIndex = firstExampleRow + ((generatedRowIndex - firstExampleRow) % existingRows);
            if (!TryCopyRow(exemplarRows[sourceRowIndex], sheetData, worksheet, sourceRowIndex, generatedRowIndex, preserveStyle, preserveFormulas, translateFormulas: preserveFormulas, out var copyError))
            {
                return new XlsxEditAppliedOperation(operation.Type, false, copyError!, operation.Sheet);
            }
        }

        if (preserveMergedRanges)
        {
            DuplicateMergedRangesForGeneratedRows(worksheet, exemplarMergedRanges, firstExampleRow, existingRows, targetRows);
        }

        worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Expanded section at {operation.Sheet}!{anchorReference} to {targetRows} row(s)", operation.Sheet, changedRange);
    }

    private static bool CanCopyRow(SheetData sheetData, int sourceRowIndex, int targetRowIndex, bool preserveFormulas, out string? error)
    {
        var sourceRow = sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == sourceRowIndex);
        if (sourceRow is null)
        {
            error = $"Source row not found: {sourceRowIndex}";
            return false;
        }

        return CanCopyRow(sourceRow, sourceRowIndex, targetRowIndex, preserveFormulas, out error);
    }

    private static bool CanCopyRow(Row sourceRow, int sourceRowIndex, int targetRowIndex, bool preserveFormulas, out string? error)
    {
        if (!preserveFormulas)
        {
            error = null;
            return true;
        }

        var rowDelta = targetRowIndex - sourceRowIndex;
        foreach (var cell in sourceRow.Elements<Cell>())
        {
            if (cell.CellFormula?.Text is not string formula)
            {
                continue;
            }

            if (!TryTranslateFormulaRows(formula, rowDelta, out _, out var formulaError))
            {
                error = $"Cannot copy row {sourceRowIndex} to {targetRowIndex}: {formulaError}";
                return false;
            }
        }

        error = null;
        return true;
    }

    private static bool TryCopyRow(
        SheetData sheetData,
        Worksheet worksheet,
        int sourceRowIndex,
        int targetRowIndex,
        bool preserveStyle,
        bool preserveFormulas,
        bool translateFormulas,
        out string? error)
    {
        var sourceRow = sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == sourceRowIndex);
        if (sourceRow is null)
        {
            error = $"Source row not found: {sourceRowIndex}";
            return false;
        }

        return TryCopyRow(sourceRow, sheetData, worksheet, sourceRowIndex, targetRowIndex, preserveStyle, preserveFormulas, translateFormulas, out error);
    }

    private static bool TryCopyRow(
        Row sourceRow,
        SheetData sheetData,
        Worksheet worksheet,
        int sourceRowIndex,
        int targetRowIndex,
        bool preserveStyle,
        bool preserveFormulas,
        bool translateFormulas,
        out string? error)
    {
        var rowDelta = targetRowIndex - sourceRowIndex;
        var translatedFormulasByReference = new Dictionary<string, string>(StringComparer.Ordinal);
        if (preserveFormulas && translateFormulas)
        {
            foreach (var cell in sourceRow.Elements<Cell>())
            {
                if (cell.CellFormula?.Text is not string formula)
                {
                    continue;
                }

                if (!TryTranslateFormulaRows(formula, rowDelta, out var translatedFormula, out var formulaError))
                {
                    error = $"Cannot copy row {sourceRowIndex} to {targetRowIndex}: {formulaError}";
                    return false;
                }

                if (cell.CellReference?.Value is string sourceReference)
                {
                    translatedFormulasByReference[sourceReference] = translatedFormula;
                }
            }
        }

        var existingTargetRow = sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == targetRowIndex);
        existingTargetRow?.Remove();

        var targetRow = (Row)sourceRow.CloneNode(true);
        targetRow.RowIndex = (uint)targetRowIndex;
        if (!preserveStyle)
        {
            targetRow.Height = null;
            targetRow.CustomHeight = null;
            targetRow.StyleIndex = null;
            targetRow.CustomFormat = null;
        }

        foreach (var cell in targetRow.Elements<Cell>())
        {
            var originalReference = cell.CellReference?.Value;
            if (cell.CellReference?.Value is string reference)
            {
                var (column, _) = ParseCellReference(reference);
                cell.CellReference = GetCellReference(column, targetRowIndex);
            }

            if (!preserveStyle)
            {
                cell.StyleIndex = null;
            }

            if (!preserveFormulas && cell.CellFormula is not null)
            {
                cell.CellFormula = null;
                cell.CellValue = null;
            }
            else if (translateFormulas && originalReference is not null && cell.CellFormula?.Text is string formula)
            {
                var translatedFormula = translatedFormulasByReference[originalReference];
                cell.CellFormula.Text = translatedFormula;
                if (!string.Equals(translatedFormula, formula, StringComparison.Ordinal))
                {
                    cell.CellValue = null;
                }
            }
        }

        InsertRow(sheetData, targetRow);
        ExpandWorksheetDimensionToRow(worksheet, targetRowIndex);
        error = null;
        return true;
    }

    private static WorksheetPart? GetWorksheetPart(WorkbookPart workbookPart, string sheetName, out string? error)
    {
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal));
        if (sheet?.Id?.Value is not string relationshipId)
        {
            error = $"Sheet not found: {sheetName}";
            return null;
        }

        error = null;
        return (WorksheetPart)workbookPart.GetPartById(relationshipId);
    }

    private static Cell? FindVisibleTextCell(WorkbookPart workbookPart, Worksheet worksheet, string text)
    {
        foreach (var row in worksheet.Descendants<Row>().OrderBy(row => row.RowIndex?.Value ?? 0))
        {
            if (row.Hidden?.Value == true)
            {
                continue;
            }

            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is not string reference || IsColumnHidden(worksheet, reference))
                {
                    continue;
                }

                if (string.Equals(GetCellText(workbookPart, cell), text, StringComparison.Ordinal))
                {
                    return cell;
                }
            }
        }

        return null;
    }

    private static string GetCellText(WorkbookPart workbookPart, Cell cell)
    {
        if (cell.DataType?.Value == CellValues.SharedString && cell.CellValue?.Text is string sharedStringIndexText)
        {
            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
            return int.TryParse(sharedStringIndexText, out var sharedStringIndex) && sharedStrings is not null
                ? sharedStrings.ElementAt(sharedStringIndex).InnerText
                : string.Empty;
        }

        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? string.Empty;
        }

        return cell.CellValue?.Text ?? string.Empty;
    }

    private static bool IsColumnHidden(Worksheet worksheet, string cellReference)
    {
        var (columnIndex, _) = ParseCellReference(cellReference);
        return worksheet.Elements<Columns>()
            .SelectMany(columns => columns.Elements<Column>())
            .Any(column => column.Hidden?.Value == true && column.Min?.Value <= columnIndex && column.Max?.Value >= columnIndex);
    }

    private static IEnumerable<(string Name, WorksheetPart Part)> GetWorksheetParts(WorkbookPart workbookPart)
    {
        foreach (var sheet in workbookPart.Workbook.Descendants<Sheet>())
        {
            if (sheet.Name?.Value is not string sheetName || sheet.Id?.Value is not string relationshipId)
            {
                continue;
            }

            yield return (sheetName, (WorksheetPart)workbookPart.GetPartById(relationshipId));
        }
    }

    private static Cell GetOrCreateCell(WorksheetPart worksheetPart, string cellReference)
    {
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? worksheetPart.Worksheet.AppendChild(new SheetData());
        var (_, rowIndex) = ParseCellReference(cellReference);
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
        if (row is null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            InsertRow(sheetData, row);
        }

        var cell = row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal));
        if (cell is null)
        {
            cell = new Cell { CellReference = cellReference };
            InsertCell(row, cell);
        }

        return cell;
    }

    private static void InsertRow(SheetData sheetData, Row row)
    {
        var nextRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value > row.RowIndex!.Value);
        if (nextRow is null)
        {
            sheetData.Append(row);
        }
        else
        {
            sheetData.InsertBefore(row, nextRow);
        }
    }

    private static void InsertCell(Row row, Cell cell)
    {
        var nextCell = row.Elements<Cell>().FirstOrDefault(existing => string.Compare(existing.CellReference?.Value, cell.CellReference?.Value, StringComparison.Ordinal) > 0);
        if (nextCell is null)
        {
            row.Append(cell);
        }
        else
        {
            row.InsertBefore(cell, nextCell);
        }
    }

    private static string ShiftCellReference(string cellReference, int rowDelta)
    {
        var (column, row) = ParseCellReference(cellReference);
        return GetCellReference(column, row + rowDelta);
    }

    private static void ShiftWorksheetDimension(Worksheet worksheet, int startRow, int rowDelta)
    {
        var dimension = worksheet.GetFirstChild<SheetDimension>();
        if (dimension?.Reference?.Value is not string reference)
        {
            return;
        }

        if (!TryParseRangeReference(reference, out var startCell, out var endCell))
        {
            return;
        }

        var (startColumn, rangeStartRow) = ParseCellReference(startCell);
        var (endColumn, rangeEndRow) = ParseCellReference(endCell);
        if (rangeStartRow >= startRow)
        {
            rangeStartRow += rowDelta;
        }

        if (rangeEndRow >= startRow)
        {
            rangeEndRow += rowDelta;
        }

        dimension.Reference = $"{GetCellReference(startColumn, rangeStartRow)}:{GetCellReference(endColumn, rangeEndRow)}";
    }

    private static void ExpandWorksheetDimensionToRow(Worksheet worksheet, int targetRow)
    {
        var dimension = worksheet.GetFirstChild<SheetDimension>();
        if (dimension?.Reference?.Value is not string reference)
        {
            return;
        }

        if (!TryParseRangeReference(reference, out var startCell, out var endCell))
        {
            return;
        }

        var (startColumn, startRow) = ParseCellReference(startCell);
        var (endColumn, endRow) = ParseCellReference(endCell);
        if (targetRow < startRow)
        {
            startRow = targetRow;
        }

        if (targetRow > endRow)
        {
            endRow = targetRow;
        }

        dimension.Reference = $"{GetCellReference(startColumn, startRow)}:{GetCellReference(endColumn, endRow)}";
    }

    private static void ShiftMergedRanges(Worksheet worksheet, int startRow, int rowDelta)
    {
        foreach (var mergeCell in worksheet.Descendants<MergeCell>())
        {
            if (mergeCell.Reference?.Value is not string reference || !TryParseRangeReference(reference, out var startCell, out var endCell))
            {
                continue;
            }

            var (startColumn, mergeStartRow) = ParseCellReference(startCell);
            var (endColumn, mergeEndRow) = ParseCellReference(endCell);
            if (mergeStartRow >= startRow)
            {
                mergeStartRow += rowDelta;
                mergeEndRow += rowDelta;
            }
            else if (mergeEndRow >= startRow)
            {
                mergeEndRow += rowDelta;
            }

            mergeCell.Reference = $"{GetCellReference(startColumn, mergeStartRow)}:{GetCellReference(endColumn, mergeEndRow)}";
        }
    }

    private static IEnumerable<(int Row, int StartColumn, int EndColumn)> GetHorizontalMergedRangesOnRows(Worksheet worksheet, int firstRow, int rowCount)
    {
        var lastRow = firstRow + rowCount - 1;
        foreach (var mergeCell in worksheet.Descendants<MergeCell>())
        {
            if (mergeCell.Reference?.Value is not string reference || !TryParseRangeReference(reference, out var startCell, out var endCell))
            {
                continue;
            }

            var (startColumn, mergeStartRow) = ParseCellReference(startCell);
            var (endColumn, mergeEndRow) = ParseCellReference(endCell);
            if (mergeStartRow == mergeEndRow && mergeStartRow >= firstRow && mergeStartRow <= lastRow)
            {
                yield return (mergeStartRow, startColumn, endColumn);
            }
        }
    }

    private static void DuplicateMergedRangesForGeneratedRows(
        Worksheet worksheet,
        IReadOnlyList<(int Row, int StartColumn, int EndColumn)> exemplarMergedRanges,
        int firstExampleRow,
        int existingRows,
        int targetRows)
    {
        if (exemplarMergedRanges.Count == 0)
        {
            return;
        }

        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells is null)
        {
            mergeCells = new MergeCells();
            worksheet.Append(mergeCells);
        }

        var existingReferences = mergeCells.Elements<MergeCell>()
            .Select(merge => merge.Reference?.Value)
            .Where(reference => !string.IsNullOrWhiteSpace(reference))
            .ToHashSet(StringComparer.Ordinal);

        for (var generatedRowIndex = firstExampleRow + existingRows; generatedRowIndex < firstExampleRow + targetRows; generatedRowIndex++)
        {
            var sourceRowIndex = firstExampleRow + ((generatedRowIndex - firstExampleRow) % existingRows);
            foreach (var merge in exemplarMergedRanges.Where(merge => merge.Row == sourceRowIndex))
            {
                var reference = $"{GetCellReference(merge.StartColumn, generatedRowIndex)}:{GetCellReference(merge.EndColumn, generatedRowIndex)}";
                if (existingReferences.Add(reference))
                {
                    mergeCells.AppendChild(new MergeCell { Reference = reference });
                }
            }
        }
    }

    private static bool TryParseRangeReference(string reference, out string startCell, out string endCell)
    {
        var parts = reference.Split(':', StringSplitOptions.TrimEntries);
        if (parts.Length == 1)
        {
            startCell = parts[0];
            endCell = parts[0];
            return true;
        }

        if (parts.Length == 2)
        {
            startCell = parts[0];
            endCell = parts[1];
            return true;
        }

        startCell = string.Empty;
        endCell = string.Empty;
        return false;
    }

    private static string TranslateFormulaRows(string formula, int rowDelta)
    {
        if (!TryTranslateFormulaRows(formula, rowDelta, out var translatedFormula, out _))
        {
            return formula;
        }

        return translatedFormula;
    }

    private static bool TryTranslateFormulaRows(string formula, int rowDelta, out string translatedFormula, out string? error)
    {
        string? formulaError = null;
        var result = FormulaCellReferencePattern.Replace(formula, match =>
        {
            if (ShouldSkipFormulaReferenceMatch(formula, match))
            {
                return match.Value;
            }

            var columnAbsolute = match.Groups[1].Value;
            var column = match.Groups[2].Value;
            var rowAbsolute = match.Groups[3].Value;
            var rowText = match.Groups[4].Value;
            if (rowAbsolute == "$" || !int.TryParse(rowText, out var row))
            {
                return match.Value;
            }

            var targetRow = row + rowDelta;
            if (targetRow < 1)
            {
                formulaError ??= $"formula translation would produce row < 1 from reference {match.Value}";
                return match.Value;
            }

            return $"{columnAbsolute}{column}{targetRow}";
        });

        translatedFormula = formulaError is null ? result : formula;
        error = formulaError;
        return formulaError is null;
    }

    private static void ShiftFormulasForInsertedRows(WorkbookPart workbookPart, string editedSheetName, int startRow, int rowDelta)
    {
        foreach (var (sheetName, worksheetPart) in GetWorksheetParts(workbookPart))
        {
            var changed = false;
            foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
            {
                if (cell.CellFormula?.Text is not string formula)
                {
                    continue;
                }

                var shiftedFormula = ShiftFormulaRowsForInsertion(formula, sheetName, editedSheetName, startRow, rowDelta);
                cell.CellFormula.Text = shiftedFormula;
                if (!string.Equals(shiftedFormula, formula, StringComparison.Ordinal))
                {
                    cell.CellValue = null;
                    changed = true;
                }
            }

            if (changed)
            {
                worksheetPart.Worksheet.Save();
            }
        }
    }

    private static string ShiftFormulaRowsForInsertion(string formula, string formulaSheetName, string editedSheetName, int startRow, int rowDelta)
    {
        return FormulaCellReferencePattern.Replace(formula, match =>
        {
            if (ShouldSkipFormulaReferenceMatch(formula, match))
            {
                return match.Value;
            }

            var qualifier = GetSheetQualifier(formula, match.Index);
            var targetsEditedSheet = qualifier is null
                ? string.Equals(formulaSheetName, editedSheetName, StringComparison.OrdinalIgnoreCase)
                : string.Equals(qualifier, editedSheetName, StringComparison.OrdinalIgnoreCase);
            if (!targetsEditedSheet)
            {
                return match.Value;
            }

            var columnAbsolute = match.Groups[1].Value;
            var column = match.Groups[2].Value;
            var rowAbsolute = match.Groups[3].Value;
            var rowText = match.Groups[4].Value;
            if (!int.TryParse(rowText, out var row) || row < startRow)
            {
                return match.Value;
            }

            return $"{columnAbsolute}{column}{rowAbsolute}{row + rowDelta}";
        });
    }

    private static bool ShouldSkipFormulaReferenceMatch(string formula, Match match)
    {
        return IsInsideQuotedSegment(formula, match.Index)
            || IsIdentifierOrFunctionNameMatch(formula, match)
            || IsUnquotedSheetNameMatch(formula, match);
    }

    private static bool IsInsideQuotedSegment(string formula, int index)
    {
        char? quote = null;
        for (var i = 0; i < index; i++)
        {
            if (formula[i] != '"' && formula[i] != '\'')
            {
                continue;
            }

            if (quote == formula[i] && i + 1 < formula.Length && formula[i + 1] == formula[i])
            {
                i++;
                continue;
            }

            if (quote == formula[i])
            {
                quote = null;
            }
            else if (quote is null)
            {
                quote = formula[i];
            }
        }

        return quote is not null;
    }

    private static bool IsIdentifierOrFunctionNameMatch(string formula, Match match)
    {
        var nextIndex = match.Index + match.Length;
        return nextIndex < formula.Length && (formula[nextIndex] == '(' || IsFormulaIdentifierCharacter(formula[nextIndex]));
    }

    private static bool IsUnquotedSheetNameMatch(string formula, Match match)
    {
        var nextIndex = match.Index + match.Length;
        return nextIndex < formula.Length && formula[nextIndex] == '!';
    }

    private static string? GetSheetQualifier(string formula, int referenceIndex)
    {
        var bangIndex = referenceIndex - 1;
        if (bangIndex < 1 || formula[bangIndex] != '!')
        {
            return null;
        }

        if (formula[bangIndex - 1] == '\'')
        {
            return GetQuotedSheetQualifier(formula, bangIndex - 1);
        }

        var start = bangIndex - 1;
        while (start >= 0 && IsFormulaIdentifierCharacter(formula[start]))
        {
            start--;
        }

        return formula[(start + 1)..bangIndex];
    }

    private static string? GetQuotedSheetQualifier(string formula, int closingQuoteIndex)
    {
        for (var i = closingQuoteIndex - 1; i >= 0; i--)
        {
            if (formula[i] != '\'')
            {
                continue;
            }

            if (i > 0 && formula[i - 1] == '\'')
            {
                i--;
                continue;
            }

            return formula[(i + 1)..closingQuoteIndex].Replace("''", "'", StringComparison.Ordinal);
        }

        return null;
    }

    private static bool IsFormulaIdentifierCharacter(char value)
    {
        return char.IsLetterOrDigit(value) || value == '_';
    }

    private static void SetCellValue(Cell cell, string value, WorkbookPart workbookPart, string? valueType)
    {
        var normalizedValueType = string.IsNullOrWhiteSpace(valueType) ? "auto" : valueType.Trim().ToLowerInvariant();
        if (normalizedValueType == "number")
        {
            if (TryGetNumericCellText(value, cell, workbookPart, allowTextFormat: true, out var numberText))
            {
                SetCellNumberValue(cell, numberText);
                return;
            }
        }
        else if (normalizedValueType == "auto")
        {
            if (TryGetNumericCellText(value, cell, workbookPart, allowTextFormat: false, out var numberText))
            {
                SetCellNumberValue(cell, numberText);
                return;
            }
        }

        SetCellStringValue(cell, value, workbookPart);
    }

    private static bool TryGetNumericCellText(string value, Cell cell, WorkbookPart workbookPart, bool allowTextFormat, out string numberText)
    {
        numberText = string.Empty;
        var text = value.Trim();
        if (text.Length == 0 || text.Contains('\n') || text.Contains('\r'))
        {
            return false;
        }

        if (!allowTextFormat && IsTextFormattedCell(cell, workbookPart))
        {
            return false;
        }

        var normalized = text.Replace(",", string.Empty, StringComparison.Ordinal);
        if (PercentTextPattern.IsMatch(normalized) && IsPercentFormattedCell(cell, workbookPart))
        {
            if (decimal.TryParse(normalized[..^1], NumberStyles.Number, CultureInfo.InvariantCulture, out var percent))
            {
                numberText = (percent / 100).ToString("G29", CultureInfo.InvariantCulture);
                return true;
            }
        }

        if (!NumericTextPattern.IsMatch(normalized) || HasUnsafeLeadingZero(normalized))
        {
            return false;
        }

        if (decimal.TryParse(normalized, NumberStyles.Number, CultureInfo.InvariantCulture, out var number))
        {
            numberText = number.ToString("G29", CultureInfo.InvariantCulture);
            return true;
        }

        return false;
    }

    private static bool HasUnsafeLeadingZero(string text)
    {
        var unsigned = text.TrimStart('+', '-');
        return unsigned.Length > 1 && unsigned[0] == '0' && unsigned[1] != '.';
    }

    private static void SetCellNumberValue(Cell cell, string numberText)
    {
        cell.CellFormula = null;
        cell.DataType = null;
        cell.InlineString = null;
        cell.CellValue = new CellValue(numberText);
    }

    private static void SetCellStringValue(Cell cell, string value, WorkbookPart workbookPart)
    {
        var sharedStringTablePart = workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringTablePart.SharedStringTable ??= new SharedStringTable();
        var sharedStringTable = sharedStringTablePart.SharedStringTable;

        var index = 0;
        var found = false;
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == value)
            {
                found = true;
                break;
            }
            index++;
        }

        if (!found)
        {
            sharedStringTable.AppendChild(new SharedStringItem(new Text(value)));
            sharedStringTable.Save();
        }

        cell.CellFormula = null;
        cell.InlineString = null;
        cell.DataType = CellValues.SharedString;
        cell.CellValue = new CellValue(index.ToString());
    }

    private static void ApplyCellBold(WorkbookPart workbookPart, Cell cell, bool bold)
    {
        var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet ??= new Stylesheet
        {
            Fonts = new Fonts(new Font()),
            Fills = new Fills(new Fill()),
            Borders = new Borders(new Border()),
            CellStyleFormats = new CellStyleFormats(new CellFormat()),
            CellFormats = new CellFormats(new CellFormat()),
        };

        var stylesheet = stylesPart.Stylesheet;
        stylesheet.Fonts ??= new Fonts(new Font());
        stylesheet.CellFormats ??= new CellFormats(new CellFormat());

        var sourceStyleIndex = cell.StyleIndex?.Value ?? 0U;
        var sourceFormat = stylesheet.CellFormats!.Elements<CellFormat>().ElementAtOrDefault((int)sourceStyleIndex) ?? stylesheet.CellFormats.Elements<CellFormat>().First();
        var sourceFontIndex = sourceFormat.FontId?.Value ?? 0U;
        var sourceFont = stylesheet.Fonts!.Elements<Font>().ElementAtOrDefault((int)sourceFontIndex) ?? stylesheet.Fonts.Elements<Font>().First();

        var targetFont = (Font)sourceFont.CloneNode(true);
        if (targetFont.Bold is null)
        {
            targetFont.Bold = new Bold();
        }
        targetFont.Bold.Val = bold;

        var fontIndex = (uint)stylesheet.Fonts!.Count();
        stylesheet.Fonts!.Append(targetFont);

        var targetFormat = (CellFormat)sourceFormat.CloneNode(true);
        targetFormat.FontId = fontIndex;
        var formatIndex = (uint)stylesheet.CellFormats!.Count();
        stylesheet.CellFormats!.Append(targetFormat);
        stylesPart.Stylesheet.Save();

        cell.StyleIndex = formatIndex;
    }

    private static bool IsTextFormattedCell(Cell cell, WorkbookPart workbookPart)
    {
        var formatCode = GetNumberFormatCode(cell, workbookPart);
        return string.Equals(formatCode, "@", StringComparison.Ordinal);
    }

    private static bool IsPercentFormattedCell(Cell cell, WorkbookPart workbookPart)
    {
        var formatCode = GetNumberFormatCode(cell, workbookPart);
        return formatCode?.Contains('%', StringComparison.Ordinal) == true;
    }

    private static string? GetNumberFormatCode(Cell cell, WorkbookPart workbookPart)
    {
        var styleIndex = cell.StyleIndex?.Value;
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (styleIndex is null || stylesPart?.Stylesheet.CellFormats is null)
        {
            return null;
        }

        var cellFormats = stylesPart.Stylesheet.CellFormats.Elements<CellFormat>().ToList();
        if (styleIndex.Value >= cellFormats.Count)
        {
            return null;
        }

        var numberFormatId = cellFormats[(int)styleIndex.Value].NumberFormatId?.Value;
        if (numberFormatId is null)
        {
            return null;
        }

        if (stylesPart.Stylesheet.NumberingFormats is not null)
        {
            var custom = stylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId.Value);
            if (custom?.FormatCode?.Value is string formatCode)
            {
                return formatCode;
            }
        }

        return numberFormatId.Value switch
        {
            9 or 10 => "0%",
            49 => "@",
            _ => null,
        };
    }

    private static (int Column, int Row) ParseCellReference(string cellReference)
    {
        var column = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var row = new string(cellReference.SkipWhile(char.IsLetter).ToArray());
        if (string.IsNullOrWhiteSpace(column) || !int.TryParse(row, out var rowIndex))
        {
            throw new InvalidOperationException($"Invalid cell reference: {cellReference}");
        }
        return (GetColumnIndex(column), rowIndex);
    }

    private static int GetColumnIndex(string columnName)
    {
        var index = 0;
        foreach (var ch in columnName.ToUpperInvariant())
        {
            index = index * 26 + (ch - 'A' + 1);
        }
        return index;
    }

    private static string GetCellReference(int column, int row)
    {
        var letters = new Stack<char>();
        while (column > 0)
        {
            column--;
            letters.Push((char)('A' + (column % 26)));
            column /= 26;
        }
        return $"{new string(letters.ToArray())}{row}";
    }
}
