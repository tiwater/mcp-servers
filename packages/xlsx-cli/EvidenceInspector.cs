using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
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
            .Where(x => x.NumberFormatId?.Value is not null)
            .GroupBy(x => x.NumberFormatId!.Value)
            .ToDictionary(x => x.Key, x => x.Last().FormatCode?.Value) ?? [];
        var baseFormats = styles?.CellStyleFormats?.Elements<CellFormat>().ToList() ?? [];
        var formattedCellsBySheet = WorkbookLoader.Load(path).Sheets.ToDictionary(
            x => x.Name,
            x => x.Cells.ToDictionary(cell => cell.Reference, cell => cell.FormattedValue, StringComparer.OrdinalIgnoreCase),
            StringComparer.Ordinal);
        var uses1904Dates = wb.Workbook.WorkbookProperties?.Date1904?.Value == true;
        var sheets = wb.Workbook.Descendants<Sheet>().Select((sheet, sheetIndex) =>
        {
            var part = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var ws = part.Worksheet;
            var significantColumn = WorksheetEvidenceBounds.SignificantColumn(part);
            var formattedCells = formattedCellsBySheet.GetValueOrDefault(sheet.Name?.Value ?? string.Empty)
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var cells = ws.Descendants<Cell>().Where(cell => WorksheetEvidenceBounds.Include(cell, significantColumn)).Select(cell =>
            {
                var styleIndex = cell.StyleIndex?.Value ?? 0U;
                var format = styleIndex < formats.Count ? formats[(int)styleIndex] : null;
                var baseFormatIndex = format?.FormatId?.Value;
                var baseFormat = baseFormatIndex is not null && baseFormatIndex.Value < baseFormats.Count
                    ? baseFormats[(int)baseFormatIndex.Value]
                    : null;
                var numId = format?.NumberFormatId?.Value ?? baseFormat?.NumberFormatId?.Value ?? 0U;
                var formatCode = customFormats.GetValueOrDefault(numId) ?? BuiltInFormat(numId) ?? $"builtin:{numId}";
                var normalizedFormat = NormalizeNumberFormat(numId, formatCode, customFormats.ContainsKey(numId));
                var alignment = format?.Alignment;
                var baseAlignment = baseFormat?.Alignment;
                var formula = cell.CellFormula;
                var rawValue = cell.CellValue?.Text ?? cell.InlineString?.InnerText;
                var reference = cell.CellReference?.Value;
                var valueType = cell.DataType?.InnerText;
                var normalizedValue = NormalizeValue(rawValue, valueType, normalizedFormat.Kind, uses1904Dates);
                return new {
                    reference,
                    rawValue,
                    formattedValue = reference is null ? null : formattedCells.GetValueOrDefault(reference),
                    valueType = valueType ?? "number",
                    styleIndex,
                    style = new {
                        baseStyleIndex = baseFormatIndex,
                        fontId = format?.FontId?.Value ?? baseFormat?.FontId?.Value ?? 0U,
                        fillId = format?.FillId?.Value ?? baseFormat?.FillId?.Value ?? 0U,
                        borderId = format?.BorderId?.Value ?? baseFormat?.BorderId?.Value ?? 0U,
                        numberFormatId = numId,
                        numberFormat = formatCode,
                        numberFormatEvidence = normalizedFormat,
                        horizontalAlignment = alignment?.Horizontal?.InnerText ?? baseAlignment?.Horizontal?.InnerText,
                        verticalAlignment = alignment?.Vertical?.InnerText ?? baseAlignment?.Vertical?.InnerText,
                        wrapText = alignment?.WrapText?.Value ?? baseAlignment?.WrapText?.Value,
                        shrinkToFit = alignment?.ShrinkToFit?.Value ?? baseAlignment?.ShrinkToFit?.Value,
                        textRotation = alignment?.TextRotation?.Value ?? baseAlignment?.TextRotation?.Value
                    },
                    normalizedValue,
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
        return new { schema = "tiwater.xlsx.evidence/v1", toolVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(), file = Path.GetFullPath(path), dateSystem = uses1904Dates ? "1904" : "1900", sheets };
    }

    private static object NormalizeValue(string? rawValue, string? dataType, string formatKind, bool uses1904Dates)
    {
        if (rawValue is null) return new { kind = "blank", iso8601 = (string?)null };
        if ((formatKind is "date" or "time" or "datetime")
            && dataType is not "s" and not "inlineStr" and not "str"
            && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            var value = ExcelSerialDate(serial, uses1904Dates);
            value = new DateTime(
                ((value.Ticks + TimeSpan.TicksPerMillisecond / 2) / TimeSpan.TicksPerMillisecond) * TimeSpan.TicksPerMillisecond,
                value.Kind);
            var iso8601 = formatKind switch
            {
                "date" => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                "time" => value.ToString("HH:mm:ss.fffffff", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.'),
                _ => value.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.')
            };
            return new { kind = formatKind, iso8601 };
        }
        return new { kind = dataType ?? "number", iso8601 = (string?)null };
    }

    private static DateTime ExcelSerialDate(double serial, bool uses1904Dates)
    {
        if (!double.IsFinite(serial)) throw new InvalidDataException($"Invalid Excel date serial: {serial}");
        if (uses1904Dates) return new DateTime(1904, 1, 1).AddDays(serial);
        // Excel intentionally preserves Lotus 1-2-3's fictitious 1900-02-29 at serial 60.
        // Mapping serial 60 to the preceding real date keeps the evidence a valid ISO date.
        var wholeDays = Math.Floor(serial);
        var fraction = serial - wholeDays;
        var day = new DateTime(1899, 12, 31).AddDays(wholeDays >= 60 ? wholeDays - 1 : wholeDays);
        return day.AddDays(fraction);
    }

    private static NumberFormatEvidence NormalizeNumberFormat(uint id, string code, bool custom)
    {
        var normalizedCode = string.Concat(code.Where(c => !char.IsWhiteSpace(c))).ToLowerInvariant();
        var kind = ClassifyNumberFormat(id, code);
        return new NumberFormatEvidence(id, code, normalizedCode, custom ? "custom" : "builtIn", kind, kind is "date" or "time" or "datetime");
    }

    private static string ClassifyNumberFormat(uint id, string code)
    {
        if (id == 49 || code == "@") return "text";
        var semantic = StripLiteralsAndConditions(code).ToLowerInvariant();
        if (id == 0 || string.Equals(semantic, "general", StringComparison.Ordinal)) return "general";
        var hasDate = semantic.IndexOfAny(['y', 'd']) >= 0;
        var hasTime = semantic.Contains('h') || semantic.Contains('s') || semantic.Contains("am/pm", StringComparison.Ordinal);
        if (NPOI.SS.UserModel.DateUtil.IsADateFormat((int)id, code))
        {
            if (hasDate && hasTime) return "datetime";
            if (hasTime) return "time";
            return "date";
        }
        if (semantic.Contains('%')) return "percentage";
        if (semantic.Contains('e')) return "scientific";
        if (semantic.Contains('#') || semantic.Contains('0') || semantic.Contains('?')) return "number";
        return "general";
    }

    private static string StripLiteralsAndConditions(string code)
    {
        var result = new System.Text.StringBuilder(code.Length);
        var quoted = false;
        for (var i = 0; i < code.Length; i++)
        {
            var c = code[i];
            if (c == '"') { quoted = !quoted; continue; }
            if (quoted) continue;
            if (c == '\\' || c == '_' || c == '*') { i++; continue; }
            if (c == '[')
            {
                var end = code.IndexOf(']', i + 1);
                if (end < 0) break;
                var token = code[(i + 1)..end];
                if (token.All(x => x is 'h' or 'H' or 'm' or 'M' or 's' or 'S')) result.Append(token);
                i = end;
                continue;
            }
            result.Append(c);
        }
        return result.ToString();
    }

    private sealed record NumberFormatEvidence(uint Id, string Code, string NormalizedCode, string Source, string Kind, bool IsDateLike);

    private static string? BuiltInFormat(uint id) => id switch
    {
        0 => "General", 1 => "0", 2 => "0.00", 3 => "#,##0", 4 => "#,##0.00",
        9 => "0%", 10 => "0.00%", 11 => "0.00E+00", 12 => "# ?/?", 13 => "# ??/??",
        14 => "m/d/yy", 15 => "d-mmm-yy", 16 => "d-mmm", 17 => "mmm-yy",
        18 => "h:mm AM/PM", 19 => "h:mm:ss AM/PM", 20 => "h:mm", 21 => "h:mm:ss",
        22 => "m/d/yy h:mm", 37 => "#,##0 ;(#,##0)", 38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)", 40 => "#,##0.00;[Red](#,##0.00)", 45 => "mm:ss",
        46 => "[h]:mm:ss", 47 => "mmss.0", 48 => "##0.0E+0", 49 => "@", _ => null
    };
}
