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

    private static XlsxEditAppliedOperation InsertRowsOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || operation.SourceRow is null || operation.TargetRow is null || operation.Count is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, sourceRow, targetRow, and count are required");
        }

        if (operation.SourceRow.Value < 1 || operation.TargetRow.Value < 1 || operation.Count.Value < 1)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sourceRow, targetRow, and count must be positive");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, error!);
        }

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "worksheet has no sheet data");
        }

        var sourceRow = sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex?.Value == (uint)operation.SourceRow.Value);
        if (sourceRow is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, $"source row not found: {operation.SourceRow.Value}");
        }

        var targetRow = operation.TargetRow.Value;
        var count = operation.Count.Value;
        var copyValues = operation.CopyValues == true;
        ShiftRowsDown(sheetData, targetRow, count);
        ShiftMergeRangesDown(worksheetPart.Worksheet, targetRow, count);

        for (var offset = 0; offset < count; offset++)
        {
            var newRowIndex = targetRow + offset;
            var clonedRow = (Row)sourceRow.CloneNode(true);
            PrepareInsertedRow(clonedRow, operation.SourceRow.Value, newRowIndex, copyValues);
            InsertRow(sheetData, clonedRow);
        }

        worksheetPart.Worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Inserted {count} row(s) into {operation.Sheet} before row {targetRow}");
    }

    private static void ShiftRowsDown(SheetData sheetData, int startRow, int count)
    {
        foreach (var row in sheetData.Elements<Row>()
                     .Where(row => row.RowIndex?.Value >= (uint)startRow)
                     .OrderByDescending(row => row.RowIndex!.Value)
                     .ToList())
        {
            var oldRowIndex = (int)row.RowIndex!.Value;
            var newRowIndex = oldRowIndex + count;
            row.RowIndex = (uint)newRowIndex;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is not string reference)
                {
                    continue;
                }

                var (column, _) = ParseCellReference(reference);
                cell.CellReference = GetCellReference(column, newRowIndex);
            }
        }
    }

    private static void PrepareInsertedRow(Row row, int sourceRowIndex, int targetRowIndex, bool copyValues)
    {
        row.RowIndex = (uint)targetRowIndex;
        foreach (var cell in row.Elements<Cell>())
        {
            var sourceReference = cell.CellReference?.Value ?? GetCellReference(1, sourceRowIndex);
            var (column, _) = ParseCellReference(sourceReference);
            cell.CellReference = GetCellReference(column, targetRowIndex);

            if (cell.CellFormula is not null)
            {
                var translatedFormula = TranslateRelativeFormulaReferences(cell.CellFormula.Text ?? cell.CellFormula.InnerText, targetRowIndex - sourceRowIndex, 0);
                cell.CellFormula = new CellFormula(translatedFormula);
                cell.CellValue = null;
                continue;
            }

            if (!copyValues)
            {
                cell.CellValue = null;
                cell.DataType = null;
                cell.InlineString = null;
            }
        }
    }

    private static void ShiftMergeRangesDown(Worksheet worksheet, int startRow, int count)
    {
        var mergeCells = worksheet.Elements<MergeCells>().FirstOrDefault();
        if (mergeCells is null)
        {
            return;
        }

        foreach (var mergeCell in mergeCells.Elements<MergeCell>())
        {
            if (mergeCell.Reference?.Value is not string reference)
            {
                continue;
            }

            var shifted = ShiftRangeRows(reference, startRow, count);
            if (shifted is not null)
            {
                mergeCell.Reference = shifted;
            }
        }
    }

    private static string? ShiftRangeRows(string reference, int startRow, int count)
    {
        var parts = reference.Split(':', 2);
        if (parts.Length == 1)
        {
            return ShiftCellReferenceRow(parts[0], startRow, count);
        }

        var start = ShiftCellReferenceRow(parts[0], startRow, count);
        var end = ShiftCellReferenceRow(parts[1], startRow, count);
        return start is null || end is null ? null : $"{start}:{end}";
    }

    private static string? ShiftCellReferenceRow(string reference, int startRow, int count)
    {
        try
        {
            var (column, row) = ParseCellReference(reference);
            return GetCellReference(column, row >= startRow ? row + count : row);
        }
        catch (InvalidOperationException)
        {
            return null;
        }
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

    private static string TranslateRelativeFormulaReferences(string formula, int rowOffset, int columnOffset)
    {
        return Regex.Replace(formula, @"(?<![A-Za-z0-9_])(?<column>\$?[A-Z]{1,3})(?<row>\$?\d+)", match =>
        {
            var columnToken = match.Groups["column"].Value;
            var rowToken = match.Groups["row"].Value;
            var absoluteColumn = columnToken.StartsWith('$');
            var absoluteRow = rowToken.StartsWith('$');
            var columnName = absoluteColumn ? columnToken[1..] : columnToken;
            var rowText = absoluteRow ? rowToken[1..] : rowToken;
            var translatedColumn = absoluteColumn
                ? columnName
                : GetCellReference(GetColumnIndex(columnName) + columnOffset, 1).TakeWhile(char.IsLetter).Aggregate(string.Empty, (current, ch) => current + ch);
            var translatedRow = absoluteRow
                ? rowText
                : (int.Parse(rowText, CultureInfo.InvariantCulture) + rowOffset).ToString(CultureInfo.InvariantCulture);
            return $"{(absoluteColumn ? "$" : string.Empty)}{translatedColumn}{(absoluteRow ? "$" : string.Empty)}{translatedRow}";
        });
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
