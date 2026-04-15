using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

public static class Inspector
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
            var formulaCellCount = 0;
            var placeholders = new HashSet<string>();
            var tablePlaceholders = new HashSet<string>();
            var noteRows = new List<NoteRowReport>();

            if (rowCount > 0 && sheetData != null)
            {
                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowCells = row.Elements<Cell>().ToList();
                    var rowCellCount = rowCells
                        .Select(cell => cell.CellReference?.Value)
                        .Where(reference => !string.IsNullOrWhiteSpace(reference))
                        .Select(reference => ParseCellReference(reference!).Column)
                        .DefaultIfEmpty(0)
                        .Max();
                    if (rowCellCount > columnCount)
                    {
                        columnCount = rowCellCount;
                    }

                    var rowTexts = new List<string>();

                    foreach (var cell in rowCells)
                    {
                        var cellValue = GetCellValue(cell, sharedStringTable);
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            rowTexts.Add(cellValue.Trim());
                        }

                        if (cell.CellFormula is not null)
                        {
                            formulaCellCount++;
                        }

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

                    var combinedText = string.Join(" ", rowTexts);
                    if (!string.IsNullOrWhiteSpace(combinedText) &&
                        (combinedText.Contains("注意", StringComparison.Ordinal) || combinedText.Length >= 40))
                    {
                        noteRows.Add(new NoteRowReport((int)(row.RowIndex?.Value ?? 0), combinedText));
                    }
                }
            }

            var sheetName = workbookPart.Workbook.Descendants<Sheet>()
                .First(s => s.Id == workbookPart.GetIdOfPart(sheetPart)).Name?.Value ?? "Unknown";

            var usedRange = sheet?.SheetDimension?.Reference?.Value;
            var mergedRanges = sheet?.Elements<MergeCells>().FirstOrDefault()?
                .Elements<MergeCell>()
                .Select(cell => cell.Reference?.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Cast<string>()
                .ToList() ?? [];

            sheets.Add(new SheetReport(
                sheetName,
                rowCount,
                columnCount,
                placeholders.ToList(),
                tablePlaceholders.ToList(),
                usedRange,
                mergedRanges,
                formulaCellCount,
                noteRows));
        }

        return new WorkbookReport(path, sheets.Count, sheets);
    }

    private static (int Column, int Row) ParseCellReference(string cellReference)
    {
        var column = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var row = new string(cellReference.SkipWhile(char.IsLetter).ToArray());
        if (string.IsNullOrWhiteSpace(column) || !int.TryParse(row, out var rowIndex))
        {
            return (0, 0);
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
