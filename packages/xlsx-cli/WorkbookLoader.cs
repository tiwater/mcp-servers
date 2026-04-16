using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
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
        string? UsedRange,
        IReadOnlyList<string> MergedRanges,
        int FormulaCellCount);

    public static WorkbookData Load(string path)
    {
        return IsLegacyXls(path)
            ? LoadLegacyXls(path)
            : LoadOpenXmlWorkbook(path);
    }

    public static bool IsLegacyXls(string path)
        => string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase);

    private static WorkbookData LoadOpenXmlWorkbook(string path)
    {
        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        var sheets = new List<SheetDataModel>();

        foreach (var sheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = sheetPart.Worksheet;
            var sheetData = worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
            var rowsData = new List<IReadOnlyList<string>>();
            var formulaCellCount = 0;

            if (sheetData is not null)
            {
                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowData = new List<string>();
                    var currentColumn = 1;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellReference = cell.CellReference?.Value;
                        if (cellReference is not null)
                        {
                            var match = Regex.Match(cellReference, @"^[A-Z]+");
                            if (match.Success)
                            {
                                var columnIndex = GetColumnIndex(match.Value);
                                while (currentColumn < columnIndex)
                                {
                                    rowData.Add(string.Empty);
                                    currentColumn++;
                                }
                            }
                        }

                        rowData.Add(GetOpenXmlCellValue(cell, sharedStringTable) ?? string.Empty);
                        if (cell.CellFormula is not null)
                        {
                            formulaCellCount++;
                        }

                        currentColumn++;
                    }

                    rowsData.Add(rowData);
                }
            }

            var name = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .FirstOrDefault(s => s.Id == workbookPart.GetIdOfPart(sheetPart))?.Name?.Value ?? "Unknown";
            var usedRange = worksheet?.SheetDimension?.Reference?.Value;
            var mergedRanges = worksheet?.Elements<MergeCells>().FirstOrDefault()?
                .Elements<MergeCell>()
                .Select(cell => cell.Reference?.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Cast<string>()
                .ToList() ?? [];

            sheets.Add(new SheetDataModel(name, rowsData, usedRange, mergedRanges, formulaCellCount));
        }

        return new WorkbookData(sheets);
    }

    private static WorkbookData LoadLegacyXls(string path)
    {
        using var stream = File.OpenRead(path);
        var workbook = new HSSFWorkbook(stream);
        var sheets = new List<SheetDataModel>();

        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            var rows = new List<IReadOnlyList<string>>();
            var formulaCellCount = 0;
            var maxColumn = 0;
            var firstRowIndex = sheet.PhysicalNumberOfRows > 0 ? sheet.FirstRowNum : 0;
            var lastRowIndex = sheet.PhysicalNumberOfRows > 0 ? sheet.LastRowNum : 0;

            for (var rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row is null)
                {
                    rows.Add([]);
                    continue;
                }

                var rowValues = new List<string>();
                var lastCellNum = Math.Max((int)row.LastCellNum, 0);
                if (lastCellNum > maxColumn)
                {
                    maxColumn = lastCellNum;
                }

                for (var cellIndex = 0; cellIndex < lastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    if (cell?.CellType == NpoiCellType.Formula)
                    {
                        formulaCellCount++;
                    }

                    rowValues.Add(GetLegacyCellValue(cell));
                }

                rows.Add(rowValues);
            }

            var usedRange = maxColumn > 0 && rows.Count > 0
                ? $"A1:{ColumnIndexToName(maxColumn)}{rows.Count}"
                : null;
            var mergedRanges = new List<string>();
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                mergedRanges.Add($"{ColumnIndexToName(region.FirstColumn + 1)}{region.FirstRow + 1}:{ColumnIndexToName(region.LastColumn + 1)}{region.LastRow + 1}");
            }

            sheets.Add(new SheetDataModel(
                sheet.SheetName,
                rows,
                usedRange,
                mergedRanges,
                formulaCellCount));
        }

        return new WorkbookData(sheets);
    }

    private static string? GetOpenXmlCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.CellValue == null)
        {
            return null;
        }

        var text = cell.CellValue.Text;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
        {
            if (int.TryParse(text, out var index))
            {
                return sharedStringTable.ElementAt(index).InnerText;
            }
        }

        return text;
    }

    private static string GetLegacyCellValue(ICell? cell)
    {
        if (cell is null)
        {
            return string.Empty;
        }

        return cell.CellType switch
        {
            NpoiCellType.String => cell.StringCellValue ?? string.Empty,
            NpoiCellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:yyyy-MM-dd}", cell.DateCellValue)
                : cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture),
            NpoiCellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            NpoiCellType.Formula => GetLegacyFormulaValue(cell),
            NpoiCellType.Blank => string.Empty,
            _ => cell.ToString() ?? string.Empty,
        };
    }

    private static string GetLegacyFormulaValue(ICell cell)
    {
        return cell.CachedFormulaResultType switch
        {
            NpoiCellType.String => cell.StringCellValue ?? string.Empty,
            NpoiCellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:yyyy-MM-dd}", cell.DateCellValue)
                : cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture),
            NpoiCellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            _ => cell.ToString() ?? string.Empty,
        };
    }

    private static int GetColumnIndex(string columnName)
    {
        var sum = 0;
        foreach (var c in columnName)
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum;
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
