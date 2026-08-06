using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

public static class Editor
{
    private const int MaximumWorksheetRow = 1_048_576;
    private const int MaximumWorksheetColumn = 16_384;
    private static readonly Regex NumericTextPattern = new(@"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$", RegexOptions.Compiled);
    private static readonly Regex PercentTextPattern = new(@"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)%$", RegexOptions.Compiled);
    private static readonly Regex FormulaCellReferencePattern = new(@"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)", RegexOptions.Compiled);
    private static readonly Regex PrintAreaRangePattern = new(@"^\$?(?<startColumn>[A-Za-z]{1,3})\$?(?<startRow>[1-9]\d*):\$?(?<endColumn>[A-Za-z]{1,3})\$?(?<endRow>[1-9]\d*)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex PrintTitleRowRangePattern = new(@"^\$[1-9]\d*:\$[1-9]\d*$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex PrintTitleColumnRangePattern = new(@"^\$[A-Za-z]{1,3}:\$[A-Za-z]{1,3}$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
        return result.AppliedOperations.All(operation => operation.Applied) ? 0 : 1;
    }

    public static XlsxEditResult Apply(string input, string output, IReadOnlyList<XlsxEditOperation> operations)
    {
        var preflight = operations.Select(ValidateWritableCoordinates).ToList();
        if (preflight.Any(error => error is not null))
        {
            return new XlsxEditResult(Path.GetFullPath(input), Path.GetFullPath(output), operations.Select((operation, index) =>
                new XlsxEditAppliedOperation(operation.Type, false, preflight[index] ?? "operation batch aborted by coordinate preflight")).ToList());
        }
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
            "setCellNumberFormat" => SetCellNumberFormatOperation(workbookPart, operation),
            "setPrintArea" => SetPrintAreaOperation(workbookPart, operation),
            "setPageSetup" => SetPageSetupOperation(workbookPart, operation),
            "setRowPageBreaks" => SetRowPageBreaksOperation(workbookPart, operation),
            "setColumnWidth" => SetColumnWidthOperation(workbookPart, operation),
            "setRichTextCellValue" => SetRichTextCellValueOperation(workbookPart, operation),
            "setRangeValues" => SetRangeValuesOperation(workbookPart, operation),
            "insertRows" => InsertRowsOperation(workbookPart, operation),
            "copyRow" => CopyRowOperation(workbookPart, operation),
            "expandSectionRows" => ExpandSectionRowsOperation(workbookPart, operation),
            _ => new XlsxEditAppliedOperation(operation.Type, false, $"Unknown operation type: {operation.Type}"),
        };
    }

    private static XlsxEditAppliedOperation SetRichTextCellValueOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || string.IsNullOrWhiteSpace(operation.Cell) || operation.Value is null || operation.Bold is null)
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, cell, value, and bold are required");
        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null) return new XlsxEditAppliedOperation(operation.Type, false, error!);
        var cell = GetOrCreateCell(worksheetPart, operation.Cell);
        cell.CellFormula = null; cell.CellValue = null; cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Run(new RunProperties(new Bold { Val = operation.Bold.Value }), new Text(operation.Value)));
        worksheetPart.Worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Updated rich text {operation.Sheet}!{operation.Cell}");
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
        if (operation.ShrinkToFit.HasValue || operation.WrapText.HasValue)
        {
            ApplyCellAlignment(workbookPart, cell, alignment =>
            {
                if (operation.ShrinkToFit.HasValue) alignment.ShrinkToFit = operation.ShrinkToFit.Value;
                if (operation.WrapText.HasValue) alignment.WrapText = operation.WrapText.Value;
            });
        }
        worksheetPart.Worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Updated {operation.Sheet}!{operation.Cell}");
    }

    private static XlsxEditAppliedOperation SetCellNumberFormatOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || string.IsNullOrWhiteSpace(operation.Cell))
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet and cell are required");
        }

        var hasExplicitFormat = !string.IsNullOrWhiteSpace(operation.NumberFormat);
        var hasSourceBinding = !string.IsNullOrWhiteSpace(operation.SourceSheet) || !string.IsNullOrWhiteSpace(operation.SourceCell);
        if (hasExplicitFormat == hasSourceBinding
            || hasSourceBinding && (string.IsNullOrWhiteSpace(operation.SourceSheet) || string.IsNullOrWhiteSpace(operation.SourceCell)))
        {
            return new XlsxEditAppliedOperation(operation.Type, false, "provide exactly one of numberFormat or sourceSheet/sourceCell");
        }

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null) return new XlsxEditAppliedOperation(operation.Type, false, error!);
        var targetCell = FindCell(worksheetPart, operation.Cell);
        if (targetCell is null)
        {
            return new XlsxEditAppliedOperation(operation.Type, false, $"Target cell not found: {operation.Sheet}!{operation.Cell}");
        }
        if (!TryGetCellFormat(workbookPart, targetCell, out var targetFormat, out error))
        {
            return new XlsxEditAppliedOperation(operation.Type, false, $"Target cell style invalid: {operation.Sheet}!{operation.Cell}: {error}");
        }

        uint numberFormatId;
        string formatDescription;
        if (hasExplicitFormat)
        {
            var formatCode = operation.NumberFormat!;
            if (!TryValidateNumberFormatCode(formatCode, out error))
            {
                return new XlsxEditAppliedOperation(operation.Type, false, error!);
            }
            numberFormatId = GetOrCreateNumberFormatId(workbookPart, formatCode);
            formatDescription = formatCode;
        }
        else
        {
            var sourceWorksheetPart = GetWorksheetPart(workbookPart, operation.SourceSheet!, out error);
            if (sourceWorksheetPart is null) return new XlsxEditAppliedOperation(operation.Type, false, error!);
            var sourceCell = FindCell(sourceWorksheetPart, operation.SourceCell!);
            if (sourceCell is null)
            {
                return new XlsxEditAppliedOperation(operation.Type, false, $"Source cell not found: {operation.SourceSheet}!{operation.SourceCell}");
            }
            if (!TryGetCellFormat(workbookPart, sourceCell, out var sourceFormat, out error)
                || !TryGetCellNumberFormatId(sourceFormat, workbookPart, out numberFormatId, out error))
            {
                return new XlsxEditAppliedOperation(operation.Type, false, $"Source cell style invalid: {operation.SourceSheet}!{operation.SourceCell}: {error}");
            }
            formatDescription = GetNumberFormatCode(sourceCell, workbookPart) ?? $"builtin:{numberFormatId}";
        }

        ApplyCellNumberFormat(workbookPart, targetCell, targetFormat, numberFormatId);
        worksheetPart.Worksheet.Save();
        return new XlsxEditAppliedOperation(
            operation.Type,
            true,
            $"Set number format {operation.Sheet}!{operation.Cell} to {formatDescription}",
            operation.Sheet,
            operation.Cell);
    }

    private static bool TryValidateNumberFormatCode(string formatCode, out string? error)
    {
        if (string.IsNullOrWhiteSpace(formatCode) || formatCode.Length > 255)
        {
            error = "numberFormat must contain between 1 and 255 non-whitespace characters";
            return false;
        }

        var quoted = false;
        var bracketed = false;
        var bracketHasContent = false;
        var sections = 1;
        for (var index = 0; index < formatCode.Length; index++)
        {
            var character = formatCode[index];
            if (char.IsControl(character))
            {
                error = "numberFormat must not contain control characters";
                return false;
            }
            if (character is '\\' or '_' or '*')
            {
                if (index + 1 >= formatCode.Length || char.IsControl(formatCode[index + 1]))
                {
                    error = "numberFormat escape, spacing, and repetition tokens require one printable following character";
                    return false;
                }
                index++;
                if (bracketed) bracketHasContent = true;
                continue;
            }
            if (character == '"')
            {
                if (bracketed)
                {
                    error = "numberFormat quoted literals are not allowed inside bracket expressions";
                    return false;
                }
                quoted = !quoted;
                continue;
            }
            if (quoted) continue;
            if (character == '[')
            {
                if (bracketed)
                {
                    error = "numberFormat bracket expressions must not be nested";
                    return false;
                }
                bracketed = true;
                bracketHasContent = false;
                continue;
            }
            if (character == ']')
            {
                if (!bracketed || !bracketHasContent)
                {
                    error = "numberFormat bracket expressions must be balanced and non-empty";
                    return false;
                }
                bracketed = false;
                continue;
            }
            if (bracketed)
            {
                bracketHasContent = true;
                continue;
            }
            if (character == ';' && ++sections > 4)
            {
                error = "numberFormat must contain at most four sections";
                return false;
            }
        }

        if (quoted || bracketed)
        {
            error = "numberFormat contains an unterminated quoted literal or bracket expression";
            return false;
        }
        error = null;
        return true;
    }

    private static XlsxEditAppliedOperation SetPrintAreaOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || string.IsNullOrWhiteSpace(operation.Range)
            || !TryParsePrintAreaRange(operation.Range, out var startCell, out var endCell))
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet and a valid A1 range are required");
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        var sheetIndex = sheets.FindIndex(sheet => string.Equals(sheet.Name?.Value, operation.Sheet, StringComparison.Ordinal));
        if (sheetIndex < 0) return new XlsxEditAppliedOperation(operation.Type, false, $"Worksheet not found: {operation.Sheet}");
        if (workbookPart.Workbook.DefinedNames is null)
        {
            var definedNames = new DefinedNames();
            OpenXmlElement anchor = workbookPart.Workbook.ExternalReferences
                ?? (OpenXmlElement?)workbookPart.Workbook.FunctionGroups
                ?? workbookPart.Workbook.Sheets
                ?? throw new InvalidOperationException("Workbook sheets are missing.");
            workbookPart.Workbook.InsertAfter(definedNames, anchor);
        }
        var workbookDefinedNames = workbookPart.Workbook.DefinedNames!;
        foreach (var existing in workbookDefinedNames.Elements<DefinedName>()
            .Where(name => name.Name?.Value == "_xlnm.Print_Area" && name.LocalSheetId?.Value == (uint)sheetIndex).ToList()) existing.Remove();
        static string absolute(string reference) => Regex.Replace(reference.ToUpperInvariant(), "^([A-Z]+)([0-9]+)$", "$$$1$$$2");
        var escapedSheet = operation.Sheet.Replace("'", "''", StringComparison.Ordinal);
        workbookDefinedNames.Append(new DefinedName($"'{escapedSheet}'!{absolute(startCell)}:{absolute(endCell)}")
        {
            Name = "_xlnm.Print_Area", LocalSheetId = (uint)sheetIndex,
        });
        workbookPart.Workbook.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Set print area {operation.Sheet}!{operation.Range}");
    }

    private static XlsxEditAppliedOperation SetColumnWidthOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet)
            || !TryParseWritableColumn(operation.Column, out var columnIndex)
            || operation.Width is null
            || !double.IsFinite(operation.Width.Value)
            || operation.Width.Value is <= 0 or > 255)
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet, a bounded column, and width in (0, 255] are required");

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null) return new XlsxEditAppliedOperation(operation.Type, false, error!);
        var worksheet = worksheetPart.Worksheet;
        var columns = worksheet.GetFirstChild<Columns>();
        if (columns is null)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>()
                ?? throw new InvalidOperationException("Worksheet sheet data not found.");
            columns = worksheet.InsertBefore(new Columns(), sheetData);
        }

        var existing = columns.Elements<Column>().ToList();
        var covering = existing.Where(column => column.Min?.Value <= (uint)columnIndex && column.Max?.Value >= (uint)columnIndex).ToList();
        if (covering.Count > 1)
            return new XlsxEditAppliedOperation(operation.Type, false, $"Column geometry is ambiguous: {operation.Sheet}!{operation.Column}");

        var rewritten = new List<Column>();
        foreach (var current in existing)
        {
            if (covering.Count == 0 || !ReferenceEquals(current, covering[0]))
            {
                rewritten.Add((Column)current.CloneNode(true));
                continue;
            }

            var min = current.Min?.Value ?? throw new InvalidOperationException("Column minimum is missing.");
            var max = current.Max?.Value ?? throw new InvalidOperationException("Column maximum is missing.");
            if (min < (uint)columnIndex)
            {
                var before = (Column)current.CloneNode(true);
                before.Min = min;
                before.Max = (uint)columnIndex - 1;
                rewritten.Add(before);
            }
            var target = (Column)current.CloneNode(true);
            target.Min = target.Max = (uint)columnIndex;
            target.Width = operation.Width.Value;
            target.CustomWidth = true;
            target.BestFit = false;
            rewritten.Add(target);
            if ((uint)columnIndex < max)
            {
                var after = (Column)current.CloneNode(true);
                after.Min = (uint)columnIndex + 1;
                after.Max = max;
                rewritten.Add(after);
            }
        }

        if (covering.Count == 0)
        {
            rewritten.Add(new Column
            {
                Min = (uint)columnIndex,
                Max = (uint)columnIndex,
                Width = operation.Width.Value,
                CustomWidth = true,
                BestFit = false,
            });
        }
        columns.RemoveAllChildren<Column>();
        foreach (var column in rewritten.OrderBy(column => column.Min?.Value ?? uint.MaxValue)) columns.Append(column);
        worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Set column width {operation.Sheet}!{operation.Column}:{operation.Column}");
    }

    private static XlsxEditAppliedOperation SetPageSetupOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet)
            || (operation.FitToPagesWide is null
                && operation.FitToPagesTall is null
                && operation.Orientation is null
                && operation.PaperSize is null
                && operation.RepeatRowsStart is null
                && operation.RepeatRowsEnd is null
                && operation.RepeatColsStart is null
                && operation.RepeatColsEnd is null)
            || operation.FitToPagesWide is < 1 or > 32767
            || operation.FitToPagesTall is < 1 or > 32767
            || !TryValidateRepeatRows(operation.RepeatRowsStart, operation.RepeatRowsEnd)
            || !TryValidateRepeatColumns(operation.RepeatColsStart, operation.RepeatColsEnd)
            || (operation.Orientation is not null
                && !string.Equals(operation.Orientation, "portrait", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(operation.Orientation, "landscape", StringComparison.OrdinalIgnoreCase))
            || !TryResolvePaperSize(operation.PaperSize, out var paperSizeCode))
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet and at least one valid page setup property are required");

        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null) return new XlsxEditAppliedOperation(operation.Type, false, error!);
        var worksheet = worksheetPart.Worksheet;
        if (operation.FitToPagesWide is not null || operation.FitToPagesTall is not null)
        {
            var sheetProperties = worksheet.GetFirstChild<SheetProperties>();
            if (sheetProperties is null)
            {
                sheetProperties = new SheetProperties();
                worksheet.PrependChild(sheetProperties);
            }
            var setupProperties = sheetProperties.GetFirstChild<PageSetupProperties>();
            if (setupProperties is null)
            {
                setupProperties = new PageSetupProperties();
                sheetProperties.Append(setupProperties);
            }
            setupProperties.FitToPage = true;
        }

        if (operation.FitToPagesWide is not null
            || operation.FitToPagesTall is not null
            || operation.Orientation is not null
            || operation.PaperSize is not null)
        {
            var pageSetup = worksheet.GetFirstChild<PageSetup>();
            if (pageSetup is null)
            {
                pageSetup = new PageSetup();
                var margins = worksheet.GetFirstChild<PageMargins>();
                if (margins is not null) worksheet.InsertAfter(pageSetup, margins);
                else worksheet.Append(pageSetup);
            }
            if (operation.FitToPagesWide is not null || operation.FitToPagesTall is not null)
            {
                pageSetup.FitToWidth = operation.FitToPagesWide is null ? 0u : (uint)operation.FitToPagesWide.Value;
                pageSetup.FitToHeight = operation.FitToPagesTall is null ? 0u : (uint)operation.FitToPagesTall.Value;
            }
            if (operation.Orientation is not null)
            {
                pageSetup.Orientation = string.Equals(operation.Orientation, "landscape", StringComparison.OrdinalIgnoreCase)
                    ? OrientationValues.Landscape
                    : OrientationValues.Portrait;
            }
            if (paperSizeCode is not null) pageSetup.PaperSize = paperSizeCode.Value;
        }
        if (operation.RepeatRowsStart is not null || operation.RepeatColsStart is not null)
            SetPrintTitles(
                workbookPart,
                operation.Sheet,
                operation.RepeatRowsStart,
                operation.RepeatRowsEnd,
                operation.RepeatColsStart,
                operation.RepeatColsEnd);
        worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Set page setup for {operation.Sheet}");
    }

    private static bool TryValidateRepeatRows(int? startRow, int? endRow)
        => (startRow is null && endRow is null)
            || (startRow is >= 1 and <= MaximumWorksheetRow
                && endRow is >= 1 and <= MaximumWorksheetRow
                && startRow <= endRow);

    private static bool TryValidateRepeatColumns(int? startColumn, int? endColumn)
        => (startColumn is null && endColumn is null)
            || (startColumn is >= 1 and <= MaximumWorksheetColumn
                && endColumn is >= 1 and <= MaximumWorksheetColumn
                && startColumn <= endColumn);

    private static void SetPrintTitles(
        WorkbookPart workbookPart,
        string sheetName,
        int? startRow,
        int? endRow,
        int? startColumn,
        int? endColumn)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        var sheetIndex = sheets.FindIndex(sheet => string.Equals(sheet.Name?.Value, sheetName, StringComparison.Ordinal));
        if (sheetIndex < 0) throw new InvalidOperationException($"Worksheet not found: {sheetName}");
        if (workbookPart.Workbook.DefinedNames is null)
        {
            var definedNames = new DefinedNames();
            OpenXmlElement anchor = workbookPart.Workbook.ExternalReferences
                ?? (OpenXmlElement?)workbookPart.Workbook.FunctionGroups
                ?? workbookPart.Workbook.Sheets
                ?? throw new InvalidOperationException("Workbook sheets are missing.");
            workbookPart.Workbook.InsertAfter(definedNames, anchor);
        }
        var workbookDefinedNames = workbookPart.Workbook.DefinedNames!;
        var existingTitles = workbookDefinedNames.Elements<DefinedName>()
            .Where(name => name.Name?.Value == "_xlnm.Print_Titles" && name.LocalSheetId?.Value == (uint)sheetIndex)
            .ToList();
        var preservedReferences = existingTitles
            .SelectMany(name => SplitDefinedNameReferences(name.Text))
            .Where(reference => !(startRow is not null && IsPrintTitleRowReference(reference, sheetName)))
            .Where(reference => !(startColumn is not null && IsPrintTitleColumnReference(reference, sheetName)))
            .ToList();
        foreach (var existing in existingTitles)
            existing.Remove();
        var escapedSheet = sheetName.Replace("'", "''", StringComparison.Ordinal);
        if (startColumn is not null && endColumn is not null)
            preservedReferences.Add($"'{escapedSheet}'!${GetColumnReference(startColumn.Value)}:${GetColumnReference(endColumn.Value)}");
        if (startRow is not null && endRow is not null)
            preservedReferences.Add($"'{escapedSheet}'!${startRow}:${endRow}");
        workbookDefinedNames.Append(new DefinedName(string.Join(",", preservedReferences))
        {
            Name = "_xlnm.Print_Titles",
            LocalSheetId = (uint)sheetIndex,
        });
        workbookPart.Workbook.Save();
    }

    private static IReadOnlyList<string> SplitDefinedNameReferences(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return [];
        var references = new List<string>();
        var start = 0;
        var quoted = false;
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] == '\'')
            {
                if (quoted && index + 1 < text.Length && text[index + 1] == '\'')
                {
                    index++;
                    continue;
                }
                quoted = !quoted;
            }
            else if (text[index] == ',' && !quoted)
            {
                var reference = text[start..index].Trim();
                if (reference.Length > 0) references.Add(reference);
                start = index + 1;
            }
        }
        var finalReference = text[start..].Trim();
        if (finalReference.Length > 0) references.Add(finalReference);
        return references;
    }

    private static bool IsPrintTitleRowReference(string reference, string sheetName)
        => IsPrintTitleReference(reference, sheetName, PrintTitleRowRangePattern);

    private static bool IsPrintTitleColumnReference(string reference, string sheetName)
        => IsPrintTitleReference(reference, sheetName, PrintTitleColumnRangePattern);

    private static bool IsPrintTitleReference(string reference, string sheetName, Regex rangePattern)
    {
        var separator = reference.LastIndexOf('!');
        if (separator <= 0) return false;
        var referenceSheet = reference[..separator].TrimStart('=').Trim();
        if (referenceSheet.Length >= 2 && referenceSheet[0] == '\'' && referenceSheet[^1] == '\'')
            referenceSheet = referenceSheet[1..^1].Replace("''", "'", StringComparison.Ordinal);
        return string.Equals(referenceSheet, sheetName, StringComparison.Ordinal)
            && rangePattern.IsMatch(reference[(separator + 1)..].Trim());
    }

    private static XlsxEditAppliedOperation SetRowPageBreaksOperation(WorkbookPart workbookPart, XlsxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.Sheet) || !TryValidateBreakBeforeRows(operation.BreakBeforeRows))
            return new XlsxEditAppliedOperation(operation.Type, false, "sheet and a strictly increasing list of rows in [2, 1048576] are required");
        var worksheetPart = GetWorksheetPart(workbookPart, operation.Sheet, out var error);
        if (worksheetPart is null) return new XlsxEditAppliedOperation(operation.Type, false, error!);
        var worksheet = worksheetPart.Worksheet;
        worksheet.GetFirstChild<RowBreaks>()?.Remove();
        var rowBreaks = new RowBreaks
        {
            Count = (uint)operation.BreakBeforeRows!.Count,
            ManualBreakCount = (uint)operation.BreakBeforeRows.Count,
        };
        foreach (var row in operation.BreakBeforeRows)
        {
            rowBreaks.Append(new Break
            {
                Id = (uint)(row - 1),
                Max = 16_383U,
                ManualPageBreak = true,
            });
        }
        worksheet.AddChild(rowBreaks, true);
        worksheet.Save();
        return new XlsxEditAppliedOperation(operation.Type, true, $"Set {operation.BreakBeforeRows.Count} row page breaks for {operation.Sheet}");
    }

    private static bool TryValidateBreakBeforeRows(IReadOnlyList<int>? rows)
        => rows is { Count: > 0 }
            && rows.All(row => row is >= 2 and <= MaximumWorksheetRow)
            && rows.Zip(rows.Skip(1), (left, right) => left < right).All(valid => valid);

    private static bool TryResolvePaperSize(string? paperSize, out uint? code)
    {
        code = paperSize?.ToLowerInvariant() switch
        {
            null => null,
            "letter" => 1u,
            "legal" => 5u,
            "a3" => 8u,
            "a4" => 9u,
            _ => uint.MaxValue,
        };
        if (code != uint.MaxValue) return true;
        code = null;
        return false;
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

    private static XlsxEditAppliedOperation InsertRowsOperation(WorkbookPart workbookPart, XlsxEditOperation operation, bool preserveMergedRanges = true, bool expandAdjacentPrintArea = false)
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
        MaterializeSharedFormulas(worksheet);
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
        ShiftPrintAreasForInsertedRows(workbookPart, operation.Sheet, startRow.Value, operation.Count.Value, expandAdjacentPrintArea);
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
        MaterializeSharedFormulas(worksheet);
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
        MaterializeSharedFormulas(worksheet);
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
        }, preserveMergedRanges, expandAdjacentPrintArea: true);
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

    private static Cell? FindCell(WorksheetPart worksheetPart, string cellReference)
    {
        var normalizedReference = cellReference.ToUpperInvariant();
        return worksheetPart.Worksheet.Descendants<Cell>()
            .FirstOrDefault(cell => string.Equals(cell.CellReference?.Value, normalizedReference, StringComparison.Ordinal));
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

    private static void ShiftPrintAreasForInsertedRows(WorkbookPart workbookPart, string sheetName, int startRow, int rowDelta, bool expandAdjacentPrintArea)
    {
        var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ?? [];
        var sheetIndex = sheets.FindIndex(sheet => string.Equals(sheet.Name?.Value, sheetName, StringComparison.Ordinal));
        if (sheetIndex < 0)
        {
            return;
        }

        var printAreas = workbookPart.Workbook.DefinedNames?.Elements<DefinedName>()
            .Where(name => name.Name?.Value == "_xlnm.Print_Area" && name.LocalSheetId?.Value == (uint)sheetIndex)
            .ToList() ?? [];
        foreach (var printArea in printAreas)
        {
            if (string.IsNullOrWhiteSpace(printArea.Text))
            {
                continue;
            }

            printArea.Text = PrintAreaRangePattern.Replace(printArea.Text, match =>
            {
                var startReference = match.Groups["start"].Value;
                var endReference = match.Groups["end"].Value;
                var (_, rangeStartRow) = ParseCellReference(startReference.Replace("$", string.Empty, StringComparison.Ordinal));
                var (_, rangeEndRow) = ParseCellReference(endReference.Replace("$", string.Empty, StringComparison.Ordinal));
                if (rangeStartRow >= startRow)
                {
                    rangeStartRow += rowDelta;
                    rangeEndRow += rowDelta;
                }
                else if (rangeEndRow >= startRow || (expandAdjacentPrintArea && rangeEndRow == startRow - 1))
                {
                    rangeEndRow += rowDelta;
                }
                else
                {
                    return match.Value;
                }

                return $"{ReplaceReferenceRow(startReference, rangeStartRow)}:{ReplaceReferenceRow(endReference, rangeEndRow)}";
            });
        }
    }

    private static string ReplaceReferenceRow(string reference, int row)
    {
        return Regex.Replace(reference, @"\d+$", row.ToString(CultureInfo.InvariantCulture));
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

    private static bool TryParsePrintAreaRange(string reference, out string startCell, out string endCell)
    {
        var match = PrintAreaRangePattern.Match(reference);
        if (!match.Success)
        {
            startCell = string.Empty;
            endCell = string.Empty;
            return false;
        }

        var startColumn = GetColumnIndex(match.Groups["startColumn"].Value);
        var endColumn = GetColumnIndex(match.Groups["endColumn"].Value);
        if (!int.TryParse(match.Groups["startRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var startRow)
            || !int.TryParse(match.Groups["endRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var endRow))
        {
            startCell = string.Empty;
            endCell = string.Empty;
            return false;
        }
        if (startColumn is < 1 or > 16384 || endColumn is < 1 or > 16384
            || startRow is < 1 or > 1048576 || endRow is < 1 or > 1048576
            || startColumn > endColumn || startRow > endRow)
        {
            startCell = string.Empty;
            endCell = string.Empty;
            return false;
        }

        startCell = $"{match.Groups["startColumn"].Value.ToUpperInvariant()}{startRow}";
        endCell = $"{match.Groups["endColumn"].Value.ToUpperInvariant()}{endRow}";
        return true;
    }

    private static string TranslateFormulaRows(string formula, int rowDelta)
    {
        if (!TryTranslateFormulaReferences(formula, rowDelta, columnDelta: 0, out var translatedFormula, out _))
        {
            return formula;
        }

        return translatedFormula;
    }

    private static bool TryTranslateFormulaRows(string formula, int rowDelta, out string translatedFormula, out string? error)
    {
        return TryTranslateFormulaReferences(formula, rowDelta, columnDelta: 0, out translatedFormula, out error);
    }

    private static bool TryTranslateFormulaReferences(string formula, int rowDelta, int columnDelta, out string translatedFormula, out string? error)
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
            var translatedColumn = column;
            if (columnAbsolute != "$")
            {
                var targetColumn = GetColumnIndex(column) + columnDelta;
                if (targetColumn < 1)
                {
                    formulaError ??= $"formula translation would produce column < 1 from reference {match.Value}";
                    return match.Value;
                }

                translatedColumn = PreserveColumnCase(GetColumnReference(targetColumn), column);
            }

            var translatedRow = rowText;
            if (rowAbsolute == "$" || !int.TryParse(rowText, out var row))
            {
                return $"{columnAbsolute}{translatedColumn}{rowAbsolute}{translatedRow}";
            }

            var targetRow = row + rowDelta;
            if (targetRow < 1)
            {
                formulaError ??= $"formula translation would produce row < 1 from reference {match.Value}";
                return match.Value;
            }

            translatedRow = targetRow.ToString(CultureInfo.InvariantCulture);
            return $"{columnAbsolute}{translatedColumn}{rowAbsolute}{translatedRow}";
        });

        translatedFormula = formulaError is null ? result : formula;
        error = formulaError;
        return formulaError is null;
    }

    private static void MaterializeSharedFormulas(Worksheet worksheet)
    {
        var sharedFormulaCells = worksheet.Descendants<Cell>()
            .Where(cell => cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared)
            .ToList();
        if (sharedFormulaCells.Count == 0)
        {
            return;
        }

        var masters = sharedFormulaCells
            .Where(cell => cell.CellFormula?.SharedIndex?.Value is not null
                && !string.IsNullOrWhiteSpace(cell.CellFormula.Text))
            .GroupBy(cell => cell.CellFormula!.SharedIndex!.Value)
            .ToDictionary(group => group.Key, group => group.First());

        foreach (var cell in sharedFormulaCells)
        {
            var formula = cell.CellFormula!;
            var sharedIndex = formula.SharedIndex?.Value;
            if (sharedIndex is null || !masters.TryGetValue(sharedIndex.Value, out var master))
            {
                continue;
            }

            var masterFormula = master.CellFormula?.Text;
            if (string.IsNullOrWhiteSpace(masterFormula)
                || master.CellReference?.Value is not string masterReference
                || cell.CellReference?.Value is not string cellReference)
            {
                continue;
            }

            var (masterColumn, masterRow) = ParseCellReference(masterReference);
            var (cellColumn, cellRow) = ParseCellReference(cellReference);
            if (TryTranslateFormulaReferences(
                    masterFormula,
                    cellRow - masterRow,
                    cellColumn - masterColumn,
                    out var materializedFormula,
                    out _))
            {
                formula.Text = materializedFormula;
            }

            formula.FormulaType = null;
            formula.Reference = null;
            formula.SharedIndex = null;
            cell.CellValue = null;
        }
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
        if (normalizedValueType == "date")
        {
            if (!DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var date)) throw new InvalidOperationException($"Invalid ISO date value: {value}");
            SetCellNumberValue(cell, date.ToOADate().ToString("G17", CultureInfo.InvariantCulture));
            return;
        }
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
        stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Elements<CellFormat>().Count();
        stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Elements<Font>().Count();
        stylesPart.Stylesheet.Save();

        cell.StyleIndex = formatIndex;
    }

    private static void ApplyCellAlignment(WorkbookPart workbookPart, Cell cell, Action<Alignment> mutate)
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
        stylesPart.Stylesheet.CellFormats ??= new CellFormats(new CellFormat());
        var formats = stylesPart.Stylesheet.CellFormats;
        var sourceStyleIndex = cell.StyleIndex?.Value ?? 0U;
        var sourceFormat = formats.Elements<CellFormat>().ElementAtOrDefault((int)sourceStyleIndex)
            ?? formats.Elements<CellFormat>().First();
        var targetFormat = (CellFormat)sourceFormat.CloneNode(true);
        var baseFormatIndex = sourceFormat.FormatId?.Value;
        var baseFormat = baseFormatIndex is not null
            ? stylesPart.Stylesheet.CellStyleFormats?.Elements<CellFormat>().ElementAtOrDefault((int)baseFormatIndex.Value)
            : null;
        targetFormat.Alignment = MaterializeAlignment(baseFormat?.Alignment, sourceFormat.ApplyAlignment?.Value == false ? null : sourceFormat.Alignment);
        mutate(targetFormat.Alignment);
        targetFormat.ApplyAlignment = true;
        var existingFormats = formats.Elements<CellFormat>().ToList();
        var equivalentIndex = existingFormats.FindIndex(format => format.OuterXml == targetFormat.OuterXml);
        uint formatIndex;
        if (equivalentIndex >= 0)
        {
            formatIndex = (uint)equivalentIndex;
        }
        else
        {
            formatIndex = (uint)existingFormats.Count;
            formats.Append(targetFormat);
        }
        formats.Count = (uint)formats.Elements<CellFormat>().Count();
        stylesPart.Stylesheet.Save();
        cell.StyleIndex = formatIndex;
    }

    private static void ApplyCellNumberFormat(WorkbookPart workbookPart, Cell cell, CellFormat sourceFormat, uint numberFormatId)
    {
        var stylesPart = EnsureStylesPart(workbookPart);
        var formats = stylesPart.Stylesheet.CellFormats!;
        var targetFormat = (CellFormat)sourceFormat.CloneNode(true);
        targetFormat.NumberFormatId = numberFormatId;
        targetFormat.ApplyNumberFormat = true;

        var existingFormats = formats.Elements<CellFormat>().ToList();
        var equivalentIndex = existingFormats.FindIndex(format => format.OuterXml == targetFormat.OuterXml);
        uint formatIndex;
        if (equivalentIndex >= 0)
        {
            formatIndex = (uint)equivalentIndex;
        }
        else
        {
            formatIndex = (uint)existingFormats.Count;
            formats.Append(targetFormat);
        }
        formats.Count = (uint)formats.Elements<CellFormat>().Count();
        stylesPart.Stylesheet.Save();
        cell.StyleIndex = formatIndex;
    }

    private static WorkbookStylesPart EnsureStylesPart(WorkbookPart workbookPart)
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
        stylesPart.Stylesheet.Fonts ??= new Fonts(new Font());
        stylesPart.Stylesheet.Fills ??= new Fills(new Fill());
        stylesPart.Stylesheet.Borders ??= new Borders(new Border());
        stylesPart.Stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
        stylesPart.Stylesheet.CellFormats ??= new CellFormats(new CellFormat());
        return stylesPart;
    }

    private static uint GetOrCreateNumberFormatId(WorkbookPart workbookPart, string formatCode)
    {
        var stylesPart = EnsureStylesPart(workbookPart);
        var stylesheet = stylesPart.Stylesheet;
        stylesheet.NumberingFormats ??= new NumberingFormats();
        var existing = stylesheet.NumberingFormats.Elements<NumberingFormat>()
            .FirstOrDefault(format => string.Equals(format.FormatCode?.Value, formatCode, StringComparison.Ordinal));
        if (existing?.NumberFormatId?.Value is uint existingId) return existingId;

        var usedIds = stylesheet.NumberingFormats.Elements<NumberingFormat>()
            .Where(format => format.NumberFormatId?.Value is not null)
            .Select(format => format.NumberFormatId!.Value)
            .Concat(stylesheet.CellFormats!.Elements<CellFormat>()
                .Where(format => format.NumberFormatId?.Value is not null)
                .Select(format => format.NumberFormatId!.Value))
            .ToHashSet();
        var numberFormatId = 164U;
        while (usedIds.Contains(numberFormatId)) numberFormatId++;
        stylesheet.NumberingFormats.Append(new NumberingFormat { NumberFormatId = numberFormatId, FormatCode = formatCode });
        stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Elements<NumberingFormat>().Count();
        stylesheet.Save();
        return numberFormatId;
    }

    private static bool TryGetCellFormat(WorkbookPart workbookPart, Cell cell, out CellFormat format, out string? error)
    {
        var stylesPart = EnsureStylesPart(workbookPart);
        var formats = stylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ToList();
        var styleIndex = cell.StyleIndex?.Value ?? 0U;
        if (styleIndex >= formats.Count)
        {
            format = null!; error = $"style index {styleIndex} is outside cellXfs count {formats.Count}"; return false;
        }
        format = formats[(int)styleIndex]; error = null; return true;
    }

    private static bool TryGetCellNumberFormatId(CellFormat format, WorkbookPart workbookPart, out uint numberFormatId, out string? error)
    {
        var baseFormats = workbookPart.WorkbookStylesPart?.Stylesheet.CellStyleFormats?.Elements<CellFormat>().ToList() ?? [];
        var baseFormatIndex = format.FormatId?.Value;
        if (baseFormatIndex is not null && baseFormatIndex.Value >= baseFormats.Count)
        {
            numberFormatId = 0U; error = $"base style index {baseFormatIndex.Value} is outside cellStyleXfs count {baseFormats.Count}"; return false;
        }
        var baseNumberFormatId = baseFormatIndex is not null ? baseFormats[(int)baseFormatIndex.Value].NumberFormatId?.Value : null;
        numberFormatId = format.ApplyNumberFormat?.Value switch
        {
            false => baseNumberFormatId ?? 0U,
            true => format.NumberFormatId?.Value ?? 0U,
            _ => format.NumberFormatId?.Value ?? baseNumberFormatId ?? 0U,
        };
        error = null; return true;
    }

    private static Alignment MaterializeAlignment(Alignment? inherited, Alignment? explicitAlignment)
    {
        var effective = inherited is null ? new Alignment() : (Alignment)inherited.CloneNode(true);
        if (explicitAlignment is null) return effective;
        foreach (var attribute in explicitAlignment.GetAttributes()) effective.SetAttribute(attribute);
        foreach (var child in explicitAlignment.ChildElements)
        {
            foreach (var existing in effective.ChildElements.Where(candidate => candidate.GetType() == child.GetType()).ToList()) existing.Remove();
            effective.Append(child.CloneNode(true));
        }
        return effective;
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
        var stylesPart = workbookPart.WorkbookStylesPart;
        if (stylesPart?.Stylesheet.CellFormats is null)
        {
            return null;
        }

        if (!TryGetCellFormat(workbookPart, cell, out var format, out _)
            || !TryGetCellNumberFormatId(format, workbookPart, out var numberFormatId, out _)) return null;

        if (stylesPart.Stylesheet.NumberingFormats is not null)
        {
            var custom = stylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId);
            if (custom?.FormatCode?.Value is string formatCode)
            {
                return formatCode;
            }
        }

        return numberFormatId switch
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

    private static string? ValidateWritableCoordinates(XlsxEditOperation operation)
    {
        if (operation.Type is "setCellValue" or "setRichTextCellValue")
            return TryParseWritableCell(operation.Cell, out _, out _) ? null : "cell must be a bounded A1 reference";
        if (operation.Type == "setCellNumberFormat")
            return TryParseWritableCell(operation.Cell, out _, out _)
                && (operation.SourceCell is null || TryParseWritableCell(operation.SourceCell, out _, out _))
                ? null
                : "target cell and optional sourceCell must be bounded A1 references";
        if (operation.Type == "setPrintArea")
            return !string.IsNullOrWhiteSpace(operation.Range) && TryParsePrintAreaRange(operation.Range, out _, out _) ? null : "range must be a bounded ordered A1 range";
        if (operation.Type == "setPageSetup")
            return !string.IsNullOrWhiteSpace(operation.Sheet)
                && (operation.FitToPagesWide is not null
                    || operation.FitToPagesTall is not null
                    || operation.Orientation is not null
                    || operation.PaperSize is not null
                    || operation.RepeatRowsStart is not null
                    || operation.RepeatRowsEnd is not null
                    || operation.RepeatColsStart is not null
                    || operation.RepeatColsEnd is not null)
                && (operation.FitToPagesWide is null or (>= 1 and <= 32767))
                && (operation.FitToPagesTall is null or (>= 1 and <= 32767))
                && TryValidateRepeatRows(operation.RepeatRowsStart, operation.RepeatRowsEnd)
                && TryValidateRepeatColumns(operation.RepeatColsStart, operation.RepeatColsEnd)
                && (operation.Orientation is null
                    || string.Equals(operation.Orientation, "portrait", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(operation.Orientation, "landscape", StringComparison.OrdinalIgnoreCase))
                && TryResolvePaperSize(operation.PaperSize, out _)
                ? null
                : "sheet and valid page setup properties are required";
        if (operation.Type == "setRowPageBreaks")
            return !string.IsNullOrWhiteSpace(operation.Sheet) && TryValidateBreakBeforeRows(operation.BreakBeforeRows)
                ? null
                : "sheet and a strictly increasing list of bounded rows are required";
        if (operation.Type == "setColumnWidth")
            return TryParseWritableColumn(operation.Column, out _)
                && operation.Width is not null
                && double.IsFinite(operation.Width.Value)
                && operation.Width.Value is > 0 and <= 255
                ? null
                : "column must be bounded and width must be in (0, 255]";
        if (operation.Type == "setRangeValues")
        {
            if (!TryParseWritableCell(operation.StartCell, out var column, out var row)) return "startCell must be a bounded A1 reference";
            var rowCount = operation.Values?.Count ?? 0;
            var columnCount = operation.Values?.Count > 0 ? operation.Values.Max(values => values?.Count ?? 0) : 0;
            if ((long)row + Math.Max(0, rowCount - 1) > 1048576L || (long)column + Math.Max(0, columnCount - 1) > 16384L) return "range exceeds worksheet bounds";
        }
        return null;
    }

    private static bool TryParseWritableCell(string? reference, out int column, out int row)
    {
        column = 0; row = 0;
        var match = Regex.Match(reference ?? string.Empty, "^(?<column>[A-Za-z]{1,3})(?<row>[1-9]\\d*)$", RegexOptions.CultureInvariant);
        if (!match.Success || !int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out row)) return false;
        column = GetColumnIndex(match.Groups["column"].Value);
        return column is >= 1 and <= 16384 && row is >= 1 and <= 1048576;
    }

    private static bool TryParseWritableColumn(string? name, out int column)
    {
        column = 0;
        if (!Regex.IsMatch(name ?? string.Empty, "^[A-Za-z]{1,3}$", RegexOptions.CultureInvariant)) return false;
        column = GetColumnIndex(name!);
        return column is >= 1 and <= 16384;
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

    private static string GetColumnReference(int column)
    {
        var letters = new Stack<char>();
        while (column > 0)
        {
            column--;
            letters.Push((char)('A' + (column % 26)));
            column /= 26;
        }

        return new string(letters.ToArray());
    }

    private static string GetCellReference(int column, int row)
    {
        return $"{GetColumnReference(column)}{row}";
    }

    private static string PreserveColumnCase(string columnReference, string originalColumn)
    {
        return originalColumn.All(char.IsLower)
            ? columnReference.ToLowerInvariant()
            : columnReference;
    }
}
