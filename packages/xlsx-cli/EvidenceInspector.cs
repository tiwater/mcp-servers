using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection;

namespace Dockit.Xlsx;

public static class EvidenceInspector
{
    public static object Inspect(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var styles = wb.WorkbookStylesPart?.Stylesheet;
        var formats = styles?.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        var customFormats = styles?.NumberingFormats?.Elements<NumberingFormat>()
            .Where(x => x.NumberFormatId?.Value is not null).ToDictionary(x => x.NumberFormatId!.Value, x => x.FormatCode?.Value) ?? [];
        var sheets = wb.Workbook.Descendants<Sheet>().Select((sheet, sheetIndex) =>
        {
            var part = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var ws = part.Worksheet;
            var cells = ws.Descendants<Cell>().Select(cell =>
            {
                var styleIndex = cell.StyleIndex?.Value ?? 0U;
                var format = styleIndex < formats.Count ? formats[(int)styleIndex] : null;
                var numId = format?.NumberFormatId?.Value;
                var alignment = format?.Alignment;
                var formula = cell.CellFormula;
                return new {
                    reference = cell.CellReference?.Value,
                    rawValue = cell.CellValue?.Text ?? cell.InlineString?.InnerText,
                    valueType = cell.DataType?.InnerText ?? "number",
                    styleIndex,
                    style = new {
                        fontId = format?.FontId?.Value, fillId = format?.FillId?.Value, borderId = format?.BorderId?.Value,
                        numberFormatId = numId, numberFormat = numId is null ? null : customFormats.GetValueOrDefault(numId.Value) ?? BuiltInFormat(numId.Value),
                        horizontalAlignment = alignment?.Horizontal?.InnerText, verticalAlignment = alignment?.Vertical?.InnerText,
                        wrapText = alignment?.WrapText?.Value, textRotation = alignment?.TextRotation?.Value
                    },
                    formula = formula is null ? null : new { text = formula.Text, type = formula.FormulaType?.InnerText, sharedIndex = formula.SharedIndex?.Value, reference = formula.Reference?.Value },
                };
            }).ToList();
            var view = ws.SheetViews?.Elements<SheetView>().FirstOrDefault();
            var setup = ws.GetFirstChild<PageSetup>();
            var margins = ws.GetFirstChild<PageMargins>();
            return new {
                name = sheet.Name?.Value, state = sheet.State?.InnerText, dimension = ws.SheetDimension?.Reference?.Value,
                mergedRanges = ws.Elements<MergeCells>().SelectMany(x => x.Elements<MergeCell>()).Select(x => x.Reference?.Value).Where(x => x is not null).ToList(),
                rowDimensions = ws.Descendants<Row>().Where(x => x.CustomHeight?.Value == true || x.Hidden?.Value == true).Select(x => new { row = x.RowIndex?.Value, height = x.Height?.Value, hidden = x.Hidden?.Value }).ToList(),
                columnDimensions = ws.Elements<Columns>().SelectMany(x => x.Elements<Column>()).Select(x => new { min = x.Min?.Value, max = x.Max?.Value, width = x.Width?.Value, hidden = x.Hidden?.Value }).ToList(),
                sheetView = view is null ? null : new { workbookViewId = view.WorkbookViewId?.Value, view = view.View?.InnerText, showGridLines = view.ShowGridLines?.Value, zoomScale = view.ZoomScale?.Value, topLeftCell = view.TopLeftCell?.Value },
                print = new { area = wb.Workbook.DefinedNames?.Elements<DefinedName>().FirstOrDefault(x => x.Name?.Value == "_xlnm.Print_Area" && x.LocalSheetId?.Value == (uint)sheetIndex)?.Text, orientation = setup?.Orientation?.InnerText, paperSize = setup?.PaperSize?.Value, scale = setup?.Scale?.Value, fitToWidth = setup?.FitToWidth?.Value, fitToHeight = setup?.FitToHeight?.Value, margins = margins is null ? null : new { left = margins.Left?.Value, right = margins.Right?.Value, top = margins.Top?.Value, bottom = margins.Bottom?.Value, header = margins.Header?.Value, footer = margins.Footer?.Value } },
                cells
            };
        }).ToList();
        return new { schema = "tiwater.xlsx.evidence/v1", toolVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(), file = Path.GetFullPath(path), dateSystem = wb.Workbook.WorkbookProperties?.Date1904?.Value == true ? "1904" : "1900", sheets };
    }

    private static string? BuiltInFormat(uint id) => id switch { 0 => "General", 1 => "0", 2 => "0.00", 9 => "0%", 10 => "0.00%", 14 => "m/d/yy", 22 => "m/d/yy h:mm", 49 => "@", _ => null };
}
