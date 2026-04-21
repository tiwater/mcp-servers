namespace Dockit.Xlsx;

public static class Inspector
{
    public static WorkbookReport Inspect(string path)
    {
        var workbook = WorkbookLoader.Load(path);
        var sheets = new List<SheetReport>();

        foreach (var sheet in workbook.Sheets)
        {
            var rowCount = sheet.Rows.Count;
            var columnCount = 0;
            var placeholders = new HashSet<string>();
            var tablePlaceholders = new HashSet<string>();

            if (rowCount > 0)
            {
                for (var rowIndex = 0; rowIndex < sheet.Rows.Count; rowIndex++)
                {
                    var row = sheet.Rows[rowIndex];
                    var rowCellCount = row.Count;
                    if (rowCellCount > columnCount)
                    {
                        columnCount = rowCellCount;
                    }

                    foreach (var cellValue in row)
                    {
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

            sheets.Add(new SheetReport(
                sheet.Name,
                rowCount,
                columnCount,
                placeholders.ToList(),
                tablePlaceholders.ToList(),
                sheet.UsedRange,
                sheet.MergedRanges.ToList(),
                sheet.FormulaCellCount));
        }

        return new WorkbookReport(path, sheets.Count, sheets);
    }
}
