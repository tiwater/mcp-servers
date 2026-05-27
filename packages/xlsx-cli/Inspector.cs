using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

public static class Inspector
{
    public static WorkbookReport Inspect(string path)
    {
        var workbook = WorkbookLoader.Load(path);
        var openXmlDetails = WorkbookLoader.IsLegacyXls(path)
            ? new Dictionary<string, SheetInspectionDetails>(StringComparer.Ordinal)
            : InspectOpenXmlDetails(path);
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
                sheet.FormulaCellCount,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.TextCells,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.FormulaCells,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.RowHeights,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.ColumnWidths));
        }

        return new WorkbookReport(path, sheets.Count, sheets);
    }

    private static Dictionary<string, SheetInspectionDetails> InspectOpenXmlDetails(string path)
    {
        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var details = new Dictionary<string, SheetInspectionDetails>(StringComparer.Ordinal);

        foreach (var sheet in workbookPart.Workbook.Descendants<Sheet>())
        {
            if (sheet.Id?.Value is null || workbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            var worksheet = worksheetPart.Worksheet;
            var textCells = new List<TextCellReport>();
            var formulaCells = new List<FormulaCellReport>();
            var rowHeights = new List<RowHeightReport>();
            var columnWidths = new List<ColumnWidthReport>();

            foreach (var column in worksheet.Elements<Columns>().SelectMany(columns => columns.Elements<Column>()))
            {
                if (column.Width?.Value is not double width)
                {
                    continue;
                }

                var min = column.Min?.Value ?? 1;
                var max = column.Max?.Value ?? min;
                for (var index = min; index <= max; index++)
                {
                    columnWidths.Add(new ColumnWidthReport(index, width));
                }
            }

            var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData is not null)
            {
                foreach (var row in sheetData.Elements<Row>())
                {
                    if (row.RowIndex?.Value is uint rowIndex && row.Height?.Value is double height)
                    {
                        rowHeights.Add(new RowHeightReport(rowIndex, height));
                    }

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var reference = cell.CellReference?.Value;
                        if (string.IsNullOrWhiteSpace(reference))
                        {
                            continue;
                        }

                        var visibleText = GetVisibleCellText(cell, sharedStrings);
                        if (!string.IsNullOrWhiteSpace(visibleText))
                        {
                            textCells.Add(new TextCellReport(reference, visibleText));
                        }

                        if (cell.CellFormula is not null)
                        {
                            formulaCells.Add(new FormulaCellReport(
                                reference,
                                cell.CellFormula.Text,
                                string.IsNullOrWhiteSpace(visibleText) ? null : visibleText));
                        }
                    }
                }
            }

            details[sheet.Name?.Value ?? "Unknown"] = new SheetInspectionDetails(
                textCells,
                formulaCells,
                rowHeights,
                columnWidths);
        }

        return details;
    }

    private static string? GetVisibleCellText(Cell cell, SharedStringTable? sharedStrings)
    {
        if (cell.InlineString is not null)
        {
            return cell.InlineString.InnerText;
        }

        var text = cell.CellValue?.Text;
        if (text is null)
        {
            return cell.InnerText;
        }

        if (cell.DataType?.Value == CellValues.SharedString && sharedStrings is not null && int.TryParse(text, out var index))
        {
            return sharedStrings.ElementAt(index).InnerText;
        }

        if (cell.DataType?.Value == CellValues.Boolean)
        {
            return text == "1" ? "TRUE" : "FALSE";
        }

        return text;
    }

    private sealed record SheetInspectionDetails(
        List<TextCellReport> TextCells,
        List<FormulaCellReport> FormulaCells,
        List<RowHeightReport> RowHeights,
        List<ColumnWidthReport> ColumnWidths);
}
