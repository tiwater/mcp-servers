using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

public static class Inspector
{
    public static System.Text.Json.JsonElement InspectEvidence(string path) =>
        System.Text.Json.JsonSerializer.SerializeToElement(new
        {
            workbook = Inspect(path),
            export = Extractor.Export(path)
        }, Json.Options);

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
                openXmlDetails.GetValueOrDefault(sheet.Name)?.TextCells ?? GetLoadedTextCells(sheet),
                openXmlDetails.GetValueOrDefault(sheet.Name)?.FormulaCells,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.RowHeights,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.ColumnWidths,
                openXmlDetails.GetValueOrDefault(sheet.Name)?.Cells));
        }

        return new WorkbookReport(path, sheets.Count, sheets);
    }

    private static List<TextCellReport> GetLoadedTextCells(WorkbookLoader.SheetDataModel sheet)
    {
        return sheet.Cells
            .Where(cell => !string.IsNullOrWhiteSpace(cell.Value))
            .Select(cell => new TextCellReport(cell.Reference, cell.Value, cell.RichTextRuns))
            .ToList();
    }

    private static Dictionary<string, SheetInspectionDetails> InspectOpenXmlDetails(string path)
    {
        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
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
            var cells = new List<CellEvidenceReport>();

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
                    if (row.RowIndex?.Value is uint rowIndex &&
                        row.Height?.Value is double height &&
                        row.CustomHeight?.Value == true)
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
                        cells.Add(new CellEvidenceReport(reference, visibleText ?? string.Empty, cell.CellFormula?.Text,
                            GetCellStyle(cell, stylesheet), OpenXmlRichText.GetCellRichTextRuns(cell, sharedStrings)));
                        if (!string.IsNullOrWhiteSpace(visibleText))
                        {
                            textCells.Add(new TextCellReport(
                                reference,
                                visibleText,
                                OpenXmlRichText.GetCellRichTextRuns(cell, sharedStrings)));
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
                columnWidths,
                cells);
        }

        return details;
    }

    private static CellStyleReport GetCellStyle(Cell cell, Stylesheet? stylesheet)
    {
        var styleIndex = cell.StyleIndex?.Value ?? 0U;
        var format = stylesheet?.CellFormats?.Elements<CellFormat>().ElementAtOrDefault((int)styleIndex);
        var numberFormatId = format?.NumberFormatId?.Value ?? 0U;
        var custom = stylesheet?.NumberingFormats?.Elements<NumberingFormat>().FirstOrDefault(item => item.NumberFormatId?.Value == numberFormatId)?.FormatCode?.Value;
        var code = custom ?? numberFormatId switch { 0 => "General", 1 => "0", 2 => "0.00", 9 => "0%", 10 => "0.00%", 14 => "m/d/yy", 49 => "@", _ => null };
        var alignment = format?.Alignment;
        return new CellStyleReport(styleIndex, numberFormatId, code, format?.FontId?.Value ?? 0U, format?.FillId?.Value ?? 0U,
            format?.BorderId?.Value ?? 0U, alignment?.Horizontal?.InnerText, alignment?.Vertical?.InnerText, alignment?.WrapText?.Value ?? false);
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
            return null;
        }

        if (cell.DataType?.Value == CellValues.SharedString && sharedStrings is not null && int.TryParse(text, out var index))
        {
            return sharedStrings.ElementAtOrDefault(index)?.InnerText;
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
        List<ColumnWidthReport> ColumnWidths,
        List<CellEvidenceReport> Cells);
}
