using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Dockit.Xlsx;

public static class Extractor
{
    public static int RunExportJson(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("export-json requires <input.xlsx> [<output.json>]");
        }

        var input = Path.GetFullPath(args[0]);
        var output = args.Length > 1 ? Path.GetFullPath(args[1]) : null;

        using var spreadsheet = SpreadsheetDocument.Open(input, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");

        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        var result = new List<object>();

        foreach (var sheetPart in workbookPart.WorksheetParts)
        {
            var sheet = sheetPart.Worksheet;
            var sheetData = sheet?.Elements<SheetData>().FirstOrDefault();
            
            var sheetName = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Id == workbookPart.GetIdOfPart(sheetPart))?.Name?.Value ?? "Unknown";

            var rowsData = new List<List<string>>();

            if (sheetData != null)
            {
                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowData = new List<string>();
                    int currentColumn = 1;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        // Calculate column index to handle skipped empty cells
                        var cellReference = cell.CellReference?.Value;
                        if (cellReference != null)
                        {
                            var match = Regex.Match(cellReference, @"^[A-Z]+");
                            if (match.Success)
                            {
                                int columnIndex = GetColumnIndex(match.Value);
                                // Fill with empty string if cells were skipped
                                while (currentColumn < columnIndex)
                                {
                                    rowData.Add("");
                                    currentColumn++;
                                }
                            }
                        }

                        rowData.Add(GetCellValue(cell, sharedStringTable) ?? "");
                        currentColumn++;
                    }

                    // For rows that might be completely empty but exist, or trailing spaces.
                    // We just add what we have.
                    rowsData.Add(rowData);
                }
            }

            result.Add(new
            {
                Sheet = sheetName,
                Rows = rowsData
            });
        }

        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        if (output != null)
        {
            File.WriteAllText(output, json);
            Console.WriteLine(output);
        }
        else
        {
            Console.WriteLine(json);
        }

        return 0;
    }

    private static int GetColumnIndex(string columnName)
    {
        int sum = 0;
        foreach (char c in columnName)
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum;
    }

    private static string? GetCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.CellValue == null) return null;

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
}
