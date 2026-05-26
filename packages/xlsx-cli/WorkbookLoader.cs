using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Globalization;
using System.Text.RegularExpressions;
using NpoiCellType = NPOI.SS.UserModel.CellType;

namespace Dockit.Xlsx;

internal static class WorkbookLoader
{
    internal sealed record WorkbookData(
        IReadOnlyList<SheetDataModel> Sheets);

    internal sealed record SheetDataModel(
        string Name,
        IReadOnlyList<IReadOnlyList<string>> Rows,
        IReadOnlyList<IReadOnlyList<string>> FormattedRows,
        IReadOnlyList<CellDataModel> Cells,
        string? UsedRange,
        IReadOnlyList<string> MergedRanges,
        int FormulaCellCount);

    internal sealed record CellDataModel(
        string Reference,
        int Row,
        int Column,
        string Value,
        string FormattedValue,
        string? Formula);

    private sealed record SharedFormula(string Formula, int BaseRow, int BaseColumn);

    public static WorkbookData Load(string path, bool resolveMergedCells = false)
    {
        return IsLegacyXls(path)
            ? LoadLegacyXls(path, resolveMergedCells)
            : LoadOpenXmlWorkbook(path, resolveMergedCells);
    }

    public static bool IsLegacyXls(string path)
        => string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase);

    private static WorkbookData LoadOpenXmlWorkbook(string path, bool resolveMergedCells = false)
    {
        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        var stylesPart = workbookPart.WorkbookStylesPart;
        var sheets = new List<SheetDataModel>();

        foreach (var sheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = sheetPart.Worksheet;
            var sheetData = worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
            var rowsData = new List<IReadOnlyList<string>>();
            var formattedRowsData = new List<IReadOnlyList<string>>();
            var cellsData = new List<CellDataModel>();
            var formulaCellCount = 0;
            var sharedFormulas = new Dictionary<uint, SharedFormula>();

            var name = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .FirstOrDefault(s => s.Id == workbookPart.GetIdOfPart(sheetPart))?.Name?.Value ?? "Unknown";
            var usedRange = worksheet?.SheetDimension?.Reference?.Value;
            var mergedRanges = worksheet?.Elements<MergeCells>().FirstOrDefault()?
                .Elements<MergeCell>()
                .Select(cell => cell.Reference?.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Cast<string>()
                .ToList() ?? [];

            if (sheetData is not null)
            {
                var rawCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
                var formattedCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
                var maxRow = 1;
                var maxColumn = 1;

                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowData = new List<string>();
                    var formattedRowData = new List<string>();
                    var currentColumn = 1;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellReference = cell.CellReference?.Value;
                        var cellRow = (int)(row.RowIndex?.Value ?? (uint)(rowsData.Count + 1));
                        var cellColumn = currentColumn;
                        if (cellReference is not null)
                        {
                            var match = Regex.Match(cellReference, @"^[A-Z]+", RegexOptions.IgnoreCase);
                            if (match.Success)
                            {
                                cellColumn = GetColumnIndex(match.Value);
                                cellRow = ParseCellReference(cellReference).Row;
                                while (!resolveMergedCells && currentColumn < cellColumn)
                                {
                                    rowData.Add(string.Empty);
                                    formattedRowData.Add(string.Empty);
                                    currentColumn++;
                                }
                            }
                        }

                        var rawValue = GetOpenXmlRawCellValue(cell, sharedStringTable) ?? string.Empty;
                        var formattedValue = GetOpenXmlFormattedCellValue(cell, sharedStringTable, stylesPart) ?? string.Empty;
                        if (!resolveMergedCells)
                        {
                            rowData.Add(rawValue);
                            formattedRowData.Add(formattedValue);
                        }

                        if (cell.CellFormula is not null)
                        {
                            formulaCellCount++;
                        }

                        var formula = ResolveOpenXmlFormula(cell, cellRow, cellColumn, sharedFormulas);
                        var resolvedReference = cellReference ?? GetCellReference(cellColumn, cellRow);
                        cellsData.Add(new CellDataModel(
                            resolvedReference,
                            cellRow,
                            cellColumn,
                            rawValue,
                            formattedValue,
                            string.IsNullOrWhiteSpace(formula) ? null : formula));

                        if (resolveMergedCells && rawCells is not null && formattedCells is not null)
                        {
                            rawCells[resolvedReference] = rawValue;
                            formattedCells[resolvedReference] = formattedValue;
                            maxRow = Math.Max(maxRow, cellRow);
                            maxColumn = Math.Max(maxColumn, cellColumn);
                        }

                        currentColumn = cellColumn + 1;
                    }

                    if (!resolveMergedCells)
                    {
                        rowsData.Add(rowData);
                        formattedRowsData.Add(formattedRowData);
                    }
                }

                if (resolveMergedCells && rawCells is not null && formattedCells is not null)
                {
                    foreach (var rangeStr in mergedRanges)
                    {
                        var parts = rangeStr.Split(':');
                        if (parts.Length != 2)
                        {
                            continue;
                        }

                        var (startCol, startRow) = ParseCellReference(parts[0]);
                        var (endCol, endRow) = ParseCellReference(parts[1]);
                        var topRaw = rawCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                        var topFormatted = formattedCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                        maxRow = Math.Max(maxRow, endRow);
                        maxColumn = Math.Max(maxColumn, endCol);

                        for (var row = startRow; row <= endRow; row++)
                        {
                            for (var col = startCol; col <= endCol; col++)
                            {
                                var cellRef = GetCellReference(col, row);
                                rawCells[cellRef] = topRaw;
                                formattedCells[cellRef] = topFormatted;
                            }
                        }
                    }

                    for (var row = 1; row <= maxRow; row++)
                    {
                        var rowData = new List<string>(maxColumn);
                        var formattedRowData = new List<string>(maxColumn);
                        for (var column = 1; column <= maxColumn; column++)
                        {
                            var cellRef = GetCellReference(column, row);
                            rowData.Add(rawCells.GetValueOrDefault(cellRef) ?? string.Empty);
                            formattedRowData.Add(formattedCells.GetValueOrDefault(cellRef) ?? string.Empty);
                        }
                        rowsData.Add(rowData);
                        formattedRowsData.Add(formattedRowData);
                    }
                }
            }

            sheets.Add(new SheetDataModel(name, rowsData, formattedRowsData, cellsData, usedRange, mergedRanges, formulaCellCount));
        }

        return new WorkbookData(sheets);
    }

    private static WorkbookData LoadLegacyXls(string path, bool resolveMergedCells = false)
    {
        using var stream = File.OpenRead(path);
        var workbook = new HSSFWorkbook(stream);
        var sheets = new List<SheetDataModel>();

        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            var rows = new List<IReadOnlyList<string>>();
            var formattedRows = new List<IReadOnlyList<string>>();
            var cells = new List<CellDataModel>();
            var formatter = new DataFormatter(CultureInfo.InvariantCulture);
            var formulaCellCount = 0;
            var maxColumn = 0;
            var firstRowIndex = sheet.PhysicalNumberOfRows > 0 ? sheet.FirstRowNum : 0;
            var lastRowIndex = sheet.PhysicalNumberOfRows > 0 ? sheet.LastRowNum : 0;

            var mergedRanges = new List<string>();
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                mergedRanges.Add($"{ColumnIndexToName(region.FirstColumn + 1)}{region.FirstRow + 1}:{ColumnIndexToName(region.LastColumn + 1)}{region.LastRow + 1}");
            }

            var rawCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
            var formattedCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
            var maxRow = 0;

            for (var rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row is null)
                {
                    if (!resolveMergedCells)
                    {
                        rows.Add([]);
                        formattedRows.Add([]);
                    }
                    continue;
                }

                var rowValues = new List<string>();
                var formattedRowValues = new List<string>();
                var lastCellNum = Math.Max((int)row.LastCellNum, 0);
                maxColumn = Math.Max(maxColumn, lastCellNum);
                maxRow = Math.Max(maxRow, rowIndex + 1);

                for (var cellIndex = 0; cellIndex < lastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    string? formula = null;
                    if (cell?.CellType == NpoiCellType.Formula)
                    {
                        formulaCellCount++;
                        formula = cell.CellFormula;
                    }

                    var rawValue = GetLegacyRawCellValue(cell);
                    var formattedValue = GetLegacyFormattedCellValue(cell, formatter);
                    if (!resolveMergedCells)
                    {
                        rowValues.Add(rawValue);
                        formattedRowValues.Add(formattedValue);
                    }

                    if (cell is not null)
                    {
                        var cellRef = GetCellReference(cellIndex + 1, rowIndex + 1);
                        cells.Add(new CellDataModel(
                            cellRef,
                            rowIndex + 1,
                            cellIndex + 1,
                            rawValue,
                            formattedValue,
                            string.IsNullOrWhiteSpace(formula) ? null : formula));
                        if (resolveMergedCells && rawCells is not null && formattedCells is not null)
                        {
                            rawCells[cellRef] = rawValue;
                            formattedCells[cellRef] = formattedValue;
                        }
                    }
                }

                if (!resolveMergedCells)
                {
                    rows.Add(rowValues);
                    formattedRows.Add(formattedRowValues);
                }
            }

            if (resolveMergedCells && rawCells is not null && formattedCells is not null)
            {
                foreach (var rangeStr in mergedRanges)
                {
                    var parts = rangeStr.Split(':');
                    if (parts.Length != 2)
                    {
                        continue;
                    }

                    var (startCol, startRow) = ParseCellReference(parts[0]);
                    var (endCol, endRow) = ParseCellReference(parts[1]);
                    var topRaw = rawCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                    var topFormatted = formattedCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                    maxRow = Math.Max(maxRow, endRow);
                    maxColumn = Math.Max(maxColumn, endCol);

                    for (var row = startRow; row <= endRow; row++)
                    {
                        for (var column = startCol; column <= endCol; column++)
                        {
                            var cellRef = GetCellReference(column, row);
                            rawCells[cellRef] = topRaw;
                            formattedCells[cellRef] = topFormatted;
                        }
                    }
                }

                for (var row = 1; row <= maxRow; row++)
                {
                    var rowData = new List<string>(maxColumn);
                    var formattedRowData = new List<string>(maxColumn);
                    for (var column = 1; column <= maxColumn; column++)
                    {
                        var cellRef = GetCellReference(column, row);
                        rowData.Add(rawCells.GetValueOrDefault(cellRef) ?? string.Empty);
                        formattedRowData.Add(formattedCells.GetValueOrDefault(cellRef) ?? string.Empty);
                    }
                    rows.Add(rowData);
                    formattedRows.Add(formattedRowData);
                }
            }

            var usedRange = maxColumn > 0 && rows.Count > 0
                ? $"A1:{ColumnIndexToName(maxColumn)}{rows.Count}"
                : null;

            sheets.Add(new SheetDataModel(
                sheet.SheetName,
                rows,
                formattedRows,
                cells,
                usedRange,
                mergedRanges,
                formulaCellCount));
        }

        return new WorkbookData(sheets);
    }

    private static string? ResolveOpenXmlFormula(Cell cell, int row, int column, Dictionary<uint, SharedFormula> sharedFormulas)
    {
        var cellFormula = cell.CellFormula;
        if (cellFormula is null)
        {
            return null;
        }

        var formulaText = cellFormula.Text ?? cellFormula.InnerText;
        var sharedIndex = cellFormula.SharedIndex?.Value;
        if (sharedIndex is not null)
        {
            if (!string.IsNullOrWhiteSpace(formulaText))
            {
                sharedFormulas[sharedIndex.Value] = new SharedFormula(formulaText, row, column);
                return formulaText;
            }

            if (sharedFormulas.TryGetValue(sharedIndex.Value, out var sharedFormula))
            {
                return TranslateRelativeFormulaReferences(
                    sharedFormula.Formula,
                    row - sharedFormula.BaseRow,
                    column - sharedFormula.BaseColumn);
            }
        }

        return string.IsNullOrWhiteSpace(formulaText) ? null : formulaText;
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
                : ColumnIndexToName(GetColumnIndex(columnName) + columnOffset);
            var translatedRow = absoluteRow
                ? rowText
                : (int.Parse(rowText, CultureInfo.InvariantCulture) + rowOffset).ToString(CultureInfo.InvariantCulture);
            return $"{(absoluteColumn ? "$" : string.Empty)}{translatedColumn}{(absoluteRow ? "$" : string.Empty)}{translatedRow}";
        });
    }

    private static string? GetOpenXmlRawCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.InlineString is not null)
        {
            return cell.InlineString?.InnerText ?? cell.InnerText;
        }

        if (cell.CellValue == null)
        {
            return cell.InnerText;
        }

        var text = cell.CellValue.Text;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
        {
            if (int.TryParse(text, out var index))
            {
                return sharedStringTable.ElementAt(index).InnerText;
            }
        }

        if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
        {
            return text == "1" ? "TRUE" : "FALSE";
        }

        return text;
    }

    private static string? GetOpenXmlFormattedCellValue(Cell cell, SharedStringTable? sharedStringTable, WorkbookStylesPart? stylesPart)
    {
        var raw = GetOpenXmlRawCellValue(cell, sharedStringTable);
        if (raw is null)
        {
            return null;
        }

        if ((cell.DataType is null || cell.DataType.Value == CellValues.Number) &&
            double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var numericValue))
        {
            return FormatOpenXmlNumber(numericValue, cell.StyleIndex?.Value, stylesPart);
        }

        return raw;
    }

    private static string GetLegacyRawCellValue(ICell? cell)
    {
        if (cell is null)
        {
            return string.Empty;
        }

        return cell.CellType switch
        {
            NpoiCellType.String => cell.StringCellValue ?? string.Empty,
            NpoiCellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? string.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd}", cell.DateCellValue)
                : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            NpoiCellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            NpoiCellType.Formula => GetLegacyRawFormulaValue(cell),
            NpoiCellType.Blank => string.Empty,
            _ => cell.ToString() ?? string.Empty,
        };
    }

    private static string GetLegacyRawFormulaValue(ICell cell)
    {
        return cell.CachedFormulaResultType switch
        {
            NpoiCellType.String => cell.StringCellValue ?? string.Empty,
            NpoiCellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? string.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd}", cell.DateCellValue)
                : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            NpoiCellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            _ => cell.ToString() ?? string.Empty,
        };
    }

    private static string GetLegacyFormattedCellValue(ICell? cell, DataFormatter formatter)
        => cell is null ? string.Empty : formatter.FormatCellValue(cell);

    private static string FormatOpenXmlNumber(double value, uint? styleIndex, WorkbookStylesPart? stylesPart)
    {
        var formatCode = GetNumberFormatCode(styleIndex, stylesPart);
        if (string.IsNullOrWhiteSpace(formatCode) || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        if (IsDateFormat(formatCode))
        {
            try
            {
                return DateTime.FromOADate(value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch (ArgumentException)
            {
                return value.ToString("G15", CultureInfo.InvariantCulture);
            }
        }

        var decimals = DecimalPlacesFromFormat(formatCode);
        var scaledValue = formatCode.Contains('%', StringComparison.Ordinal) ? value * 100 : value;
        var formatted = scaledValue.ToString($"F{decimals}", CultureInfo.InvariantCulture);
        return formatCode.Contains('%', StringComparison.Ordinal) ? $"{formatted}%" : formatted;
    }

    private static string? GetNumberFormatCode(uint? styleIndex, WorkbookStylesPart? stylesPart)
    {
        if (styleIndex is null || stylesPart?.Stylesheet?.CellFormats is null)
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

        var custom = stylesPart.Stylesheet.NumberingFormats?
            .Elements<NumberingFormat>()
            .FirstOrDefault(format => format.NumberFormatId?.Value == numberFormatId.Value)
            ?.FormatCode?.Value;
        if (!string.IsNullOrWhiteSpace(custom))
        {
            return custom;
        }

        return BuiltInNumberFormat(numberFormatId.Value);
    }

    private static string? BuiltInNumberFormat(uint numberFormatId) => numberFormatId switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0;(#,##0)",
        38 => "#,##0;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };

    private static bool IsDateFormat(string formatCode)
    {
        var cleaned = StripQuotedAndEscapedFormatText(formatCode).ToLowerInvariant();
        return cleaned.Contains('y') || cleaned.Contains('d') || cleaned.Contains("m/");
    }

    private static int DecimalPlacesFromFormat(string formatCode)
    {
        var section = StripQuotedAndEscapedFormatText(formatCode).Split(';', 2)[0];
        var decimalIndex = section.IndexOf('.', StringComparison.Ordinal);
        if (decimalIndex < 0)
        {
            return 0;
        }

        var count = 0;
        for (var i = decimalIndex + 1; i < section.Length; i++)
        {
            if (section[i] is '0' or '#' or '?')
            {
                count++;
                continue;
            }
            break;
        }

        return count;
    }

    private static string StripQuotedAndEscapedFormatText(string formatCode)
    {
        var chars = new List<char>();
        var inQuote = false;
        for (var i = 0; i < formatCode.Length; i++)
        {
            var c = formatCode[i];
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote)
            {
                continue;
            }
            if (c == '\\' || c == '_' || c == '*')
            {
                i++;
                continue;
            }
            if (c == '[')
            {
                while (i < formatCode.Length && formatCode[i] != ']')
                {
                    i++;
                }
                continue;
            }
            chars.Add(c);
        }

        return new string(chars.ToArray());
    }

    private static (int Column, int Row) ParseCellReference(string reference)
    {
        var colStr = new string(reference.TakeWhile(char.IsLetter).ToArray());
        var rowStr = new string(reference.Skip(colStr.Length).ToArray());
        var col = GetColumnIndex(colStr);
        var row = int.Parse(rowStr, CultureInfo.InvariantCulture);
        return (col, row);
    }

    private static int GetColumnIndex(string columnName)
    {
        var sum = 0;
        foreach (var c in columnName.ToUpperInvariant())
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum;
    }

    private static string GetCellReference(int column, int row)
    {
        return $"{ColumnIndexToName(column)}{row}";
    }

    private static string ColumnIndexToName(int index)
    {
        if (index <= 0)
        {
            return "A";
        }

        var chars = new Stack<char>();
        var current = index;
        while (current > 0)
        {
            current--;
            chars.Push((char)('A' + (current % 26)));
            current /= 26;
        }

        return new string(chars.ToArray());
    }
}
