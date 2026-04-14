using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

internal static class Inspector
{
    public static WorkbookReport Inspect(string path)
    {
        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart!;

        var sheets = new List<SheetReport>();

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

        foreach (var sheetPart in workbookPart.WorksheetParts)
        {
            var sheet = sheetPart.Worksheet;
            var sheetData = sheet?.Elements<SheetData>().FirstOrDefault();

            var rowCount = sheetData?.Elements<Row>().Count() ?? 0;
            var columnCount = 0;
            
            var placeholders = new HashSet<string>();
            var tablePlaceholders = new HashSet<string>();

            if (rowCount > 0 && sheetData != null)
            {
                foreach (var row in sheetData.Elements<Row>())
                {
                    var cellCount = row.Elements<Cell>().Count();
                    if (cellCount > columnCount)
                        columnCount = cellCount;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellValue = GetCellValue(cell, sharedStringTable);
                        if (cellValue != null && cellValue.StartsWith("{{") && cellValue.EndsWith("}}"))
                        {
                            if (cellValue.StartsWith("{{table:"))
                            {
                                tablePlaceholders.Add(cellValue[8..^2]);
                            }
                            else
                            {
                                placeholders.Add(cellValue[2..^2]);
                            }
                        }
                    }
                }
            }

            var sheetName = workbookPart.Workbook.Descendants<Sheet>()
                .First(s => s.Id == workbookPart.GetIdOfPart(sheetPart)).Name?.Value ?? "Unknown";

            sheets.Add(new SheetReport(sheetName, rowCount, columnCount, placeholders.ToList(), tablePlaceholders.ToList()));
        }

        return new WorkbookReport(path, sheets.Count, sheets);
    }

    private static string? GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.DataType?.Value == CellValues.SharedString && sharedStringTable != null)
        {
            if (int.TryParse(cell.InnerText, out var index))
            {
                return sharedStringTable.ElementAt(index).InnerText;
            }
        }

        return cell.InnerText;
    }
}
