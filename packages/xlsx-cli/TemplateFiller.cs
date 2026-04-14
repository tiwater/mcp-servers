using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace Dockit.Xlsx;

internal static class TemplateFiller
{
    public static void Fill(string templatePath, FillData data, string outputPath)
    {
        File.Copy(templatePath, outputPath, true);

        using var spreadsheet = SpreadsheetDocument.Open(outputPath, true);
        var workbookPart = spreadsheet.WorkbookPart!;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData == null)
                continue;

            // Fill single cell values
            if (data.CellValues.Count > 0)
            {
                FillCellValues(sheetData, data.CellValues, workbookPart);
            }

            // Fill table data
            if (data.TableData != null && data.TableData.Count > 0)
            {
                FillTableData(sheetData, data.TableData, workbookPart);
            }
        }

        spreadsheet.Save();
    }

    private static void FillCellValues(
        SheetData sheetData,
        Dictionary<string, string> cellValues,
        WorkbookPart workbookPart)
    {
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var cellReference = cell.CellReference?.Value;
                if (cellReference == null)
                    continue;

                var cellValue = GetCellValue(cell, sharedStringTable);

                // Check if this cell contains a placeholder like {{key}}
                if (cellValue != null && cellValue.StartsWith("{{") && cellValue.EndsWith("}}"))
                {
                    var key = cellValue[2..^2];

                    if (cellValues.TryGetValue(key, out var newValue))
                    {
                        SetCellValue(cell, newValue, workbookPart);
                    }
                }
            }
        }
    }

    private static void FillTableData(
        SheetData sheetData,
        Dictionary<string, List<List<string>>> tableData,
        WorkbookPart workbookPart)
    {
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;

        // Find tables by looking for {{table:key}} placeholders
        var tablesToFill = new List<(string TableKey, Cell AnchorCell, int StartRow)>();

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var cellValue = GetCellValue(cell, sharedStringTable);
                if (cellValue != null && cellValue.StartsWith("{{table:") && cellValue.EndsWith("}}"))
                {
                    var tableKey = cellValue[7..^2];
                    if (row.RowIndex != null)
                    {
                        tablesToFill.Add((tableKey, cell, (int)row.RowIndex.Value));
                    }
                }
            }
        }

        // Fill each table
        foreach (var (tableKey, anchorCell, startRow) in tablesToFill)
        {
            if (!tableData.TryGetValue(tableKey, out var rows))
                continue;

            var startColumn = GetColumnIndex(anchorCell.CellReference?.Value ?? "");

            // Write data rows
            for (var i = 0; i < rows.Count; i++)
            {
                var targetRow = startRow + i;
                var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)targetRow);

                if (row == null)
                {
                    row = new Row() { RowIndex = (uint)targetRow };
                    sheetData.Append(row);
                }

                for (var j = 0; j < rows[i].Count; j++)
                {
                    var cellReference = GetCellReference(startColumn + j, (int)targetRow);
                    var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

                    if (cell == null)
                    {
                        cell = new Cell() { CellReference = cellReference };
                        row.Append(cell);
                    }

                    SetCellValue(cell, rows[i][j], workbookPart);
                }
            }
        }
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

    private static void SetCellValue(Cell cell, string value, WorkbookPart workbookPart)
    {
        var sharedStringTablePart = workbookPart.SharedStringTablePart;

        if (sharedStringTablePart == null)
        {
            // Create shared string table if it doesn't exist
            sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
        }

        var sharedStringTable = sharedStringTablePart.SharedStringTable;

        // Check if string already exists
        var index = 0;
        var found = false;
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            if (item.Text?.Text == value)
            {
                found = true;
                break;
            }
            index++;
        }

        if (!found)
        {
            var item = new SharedStringItem(new Text(value));
            sharedStringTable.Append(item);
            sharedStringTable.Save();
        }

        cell.DataType = CellValues.SharedString;
        cell.CellValue = new CellValue(index.ToString());
    }

    private static int GetColumnIndex(string cellReference)
    {
        var column = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var index = 0;
        for (var i = 0; i < column.Length; i++)
        {
            index = index * 26 + (column[i] - 'A' + 1);
        }
        return index;
    }

    private static string GetCellReference(int column, int row)
    {
        var columnName = new StringBuilder();
        while (column > 0)
        {
            column--;
            columnName.Insert(0, (char)('A' + column % 26));
            column /= 26;
        }
        return $"{columnName}{row}";
    }
}
