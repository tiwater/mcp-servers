using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Globalization;
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
        IReadOnlyList<IReadOnlyList<string>> FormattedRows,
        IReadOnlyList<CellDataModel> Cells,
        string? UsedRange,
        IReadOnlyList<string> MergedRanges,
        int FormulaCellCount);

    internal sealed record CellDataModel(
        string Reference,
        int Row,
        int Column,
        string Value,
        string FormattedValue,
        string? Formula,
        IReadOnlyList<RichTextRunReport>? RichTextRuns = null,
        XlsxNormalizedValue? NormalizedValue = null);

    private sealed record SharedFormula(string Formula, int BaseRow, int BaseColumn);

    public static WorkbookData Load(
        string path,
        bool resolveMergedCells = false,
        bool includeNormalizedValues = false)
    {
        return IsLegacyXls(path)
            ? LoadLegacyXls(path, resolveMergedCells, includeNormalizedValues)
            : LoadOpenXmlWorkbook(path, resolveMergedCells, includeNormalizedValues);
    }

    public static bool IsLegacyXls(string path)
        => string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase);

    private static WorkbookData LoadOpenXmlWorkbook(
        string path,
        bool resolveMergedCells,
        bool includeNormalizedValues)
    {
        using var spreadsheet = SpreadsheetDocument.Open(path, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
        var stylesPart = workbookPart.WorkbookStylesPart;
        var uses1904Dates = workbookPart.Workbook.WorkbookProperties?.Date1904?.Value == true;
        var sheets = new List<SheetDataModel>();

        foreach (var sheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = sheetPart.Worksheet;
            var sheetData = worksheet?.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetData>().FirstOrDefault();
            var rowsData = new List<IReadOnlyList<string>>();
            var formattedRowsData = new List<IReadOnlyList<string>>();
            var cellsData = new List<CellDataModel>();
            var formulaCellCount = 0;
            var sharedFormulas = new Dictionary<uint, SharedFormula>();

            var name = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .FirstOrDefault(s => s.Id == workbookPart.GetIdOfPart(sheetPart))?.Name?.Value ?? "Unknown";
            var usedRange = worksheet?.SheetDimension?.Reference?.Value;
            var mergedRanges = worksheet?.Elements<MergeCells>().FirstOrDefault()?
                .Elements<MergeCell>()
                .Select(cell => cell.Reference?.Value)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Cast<string>()
                .ToList() ?? [];
            var significantColumn = WorksheetEvidenceBounds.SignificantColumn(sheetPart);

            if (sheetData is not null)
            {
                var rawCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
                var formattedCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
                var maxRow = 1;
                var maxColumn = 1;

                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowData = new List<string>();
                    var formattedRowData = new List<string>();
                    var currentColumn = 1;

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellReference = cell.CellReference?.Value;
                        var cellRow = (int)(row.RowIndex?.Value ?? (uint)(rowsData.Count + 1));
                        var cellColumn = currentColumn;
                        if (cellReference is not null)
                        {
                            var match = Regex.Match(cellReference, @"^[A-Z]+", RegexOptions.IgnoreCase);
                            if (match.Success)
                            {
                                cellColumn = GetColumnIndex(match.Value);
                                cellRow = ParseCellReference(cellReference).Row;
                            }
                        }

                        var rawValue = GetOpenXmlRawCellValue(cell, sharedStringTable) ?? string.Empty;
                        var formattedValue = GetOpenXmlFormattedCellValue(cell, sharedStringTable, stylesPart) ?? string.Empty;
                        if (!WorksheetEvidenceBounds.Include(cell, significantColumn)) continue;
                        while (!resolveMergedCells && currentColumn < cellColumn)
                        {
                            rowData.Add(string.Empty);
                            formattedRowData.Add(string.Empty);
                            currentColumn++;
                        }
                        var richTextRuns = OpenXmlRichText.GetCellRichTextRuns(cell, sharedStringTable, stylesPart?.Stylesheet);
                        if (!resolveMergedCells)
                        {
                            rowData.Add(rawValue);
                            formattedRowData.Add(formattedValue);
                        }

                        if (cell.CellFormula is not null)
                        {
                            formulaCellCount++;
                        }

                        var formula = ResolveOpenXmlFormula(cell, cellRow, cellColumn, sharedFormulas);
                        var resolvedReference = cellReference ?? GetCellReference(cellColumn, cellRow);
                        cellsData.Add(new CellDataModel(
                            resolvedReference,
                            cellRow,
                            cellColumn,
                            rawValue,
                            formattedValue,
                            string.IsNullOrWhiteSpace(formula) ? null : formula,
                            richTextRuns,
                            includeNormalizedValues
                                ? SpreadsheetValueNormalizer.FromOpenXmlCell(
                                    cell.CellValue?.Text ?? cell.InlineString?.InnerText,
                                    cell.DataType?.InnerText,
                                    stylesPart,
                                    cell.StyleIndex?.Value,
                                    uses1904Dates)
                                : null));

                        if (resolveMergedCells && rawCells is not null && formattedCells is not null)
                        {
                            rawCells[resolvedReference] = rawValue;
                            formattedCells[resolvedReference] = formattedValue;
                            maxRow = Math.Max(maxRow, cellRow);
                            maxColumn = Math.Max(maxColumn, cellColumn);
                        }

                        currentColumn = cellColumn + 1;
                    }

                    if (!resolveMergedCells)
                    {
                        rowsData.Add(rowData);
                        formattedRowsData.Add(formattedRowData);
                    }
                }

                if (resolveMergedCells && rawCells is not null && formattedCells is not null)
                {
                    foreach (var rangeStr in mergedRanges)
                    {
                        var parts = rangeStr.Split(':');
                        if (parts.Length != 2)
                        {
                            continue;
                        }

                        var (startCol, startRow) = ParseCellReference(parts[0]);
                        var (endCol, endRow) = ParseCellReference(parts[1]);
                        var topRaw = rawCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                        var topFormatted = formattedCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                        maxRow = Math.Max(maxRow, endRow);
                        maxColumn = Math.Max(maxColumn, endCol);

                        for (var row = startRow; row <= endRow; row++)
                        {
                            for (var col = startCol; col <= endCol; col++)
                            {
                                var cellRef = GetCellReference(col, row);
                                rawCells[cellRef] = topRaw;
                                formattedCells[cellRef] = topFormatted;
                            }
                        }
                    }

                    for (var row = 1; row <= maxRow; row++)
                    {
                        var rowData = new List<string>(maxColumn);
                        var formattedRowData = new List<string>(maxColumn);
                        for (var column = 1; column <= maxColumn; column++)
                        {
                            var cellRef = GetCellReference(column, row);
                            rowData.Add(rawCells.GetValueOrDefault(cellRef) ?? string.Empty);
                            formattedRowData.Add(formattedCells.GetValueOrDefault(cellRef) ?? string.Empty);
                        }
                        rowsData.Add(rowData);
                        formattedRowsData.Add(formattedRowData);
                    }
                }
            }

            sheets.Add(new SheetDataModel(name, rowsData, formattedRowsData, cellsData, usedRange, mergedRanges, formulaCellCount));
        }

        return new WorkbookData(sheets);
    }

    private static WorkbookData LoadLegacyXls(
        string path,
        bool resolveMergedCells,
        bool includeNormalizedValues)
    {
        using var stream = File.OpenRead(path);
        var workbook = new HSSFWorkbook(stream);
        var sheets = new List<SheetDataModel>();

        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            var rows = new List<IReadOnlyList<string>>();
            var formattedRows = new List<IReadOnlyList<string>>();
            var cells = new List<CellDataModel>();
            var formatter = new DataFormatter(CultureInfo.InvariantCulture);
            var formulaCellCount = 0;
            var maxColumn = 0;
            var firstRowIndex = sheet.PhysicalNumberOfRows > 0 ? sheet.FirstRowNum : 0;
            var lastRowIndex = sheet.PhysicalNumberOfRows > 0 ? sheet.LastRowNum : 0;

            var mergedRanges = new List<string>();
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var region = sheet.GetMergedRegion(i);
                mergedRanges.Add($"{ColumnIndexToName(region.FirstColumn + 1)}{region.FirstRow + 1}:{ColumnIndexToName(region.LastColumn + 1)}{region.LastRow + 1}");
            }

            var rawCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
            var formattedCells = resolveMergedCells ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) : null;
            var maxRow = 0;

            for (var rowIndex = firstRowIndex; rowIndex <= lastRowIndex; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row is null)
                {
                    if (!resolveMergedCells)
                    {
                        rows.Add([]);
                        formattedRows.Add([]);
                    }
                    continue;
                }

                var rowValues = new List<string>();
                var formattedRowValues = new List<string>();
                var lastCellNum = Math.Max((int)row.LastCellNum, 0);
                maxColumn = Math.Max(maxColumn, lastCellNum);
                maxRow = Math.Max(maxRow, rowIndex + 1);

                for (var cellIndex = 0; cellIndex < lastCellNum; cellIndex++)
                {
                    var cell = row.GetCell(cellIndex);
                    string? formula = null;
                    if (cell?.CellType == NpoiCellType.Formula)
                    {
                        formulaCellCount++;
                        formula = cell.CellFormula;
                    }

                    var rawValue = GetLegacyRawCellValue(cell);
                    var formattedValue = GetLegacyFormattedCellValue(cell, formatter);
                    var richTextRuns = GetLegacyRichTextRuns(cell, workbook);
                    if (!resolveMergedCells)
                    {
                        rowValues.Add(rawValue);
                        formattedRowValues.Add(formattedValue);
                    }

                    if (cell is not null)
                    {
                        var cellRef = GetCellReference(cellIndex + 1, rowIndex + 1);
                        cells.Add(new CellDataModel(
                            cellRef,
                            rowIndex + 1,
                            cellIndex + 1,
                            rawValue,
                            formattedValue,
                            string.IsNullOrWhiteSpace(formula) ? null : formula,
                            richTextRuns,
                            includeNormalizedValues ? LegacyNormalizedValue(cell) : null));
                        if (resolveMergedCells && rawCells is not null && formattedCells is not null)
                        {
                            rawCells[cellRef] = rawValue;
                            formattedCells[cellRef] = formattedValue;
                        }
                    }
                }

                if (!resolveMergedCells)
                {
                    rows.Add(rowValues);
                    formattedRows.Add(formattedRowValues);
                }
            }

            if (resolveMergedCells && rawCells is not null && formattedCells is not null)
            {
                foreach (var rangeStr in mergedRanges)
                {
                    var parts = rangeStr.Split(':');
                    if (parts.Length != 2)
                    {
                        continue;
                    }

                    var (startCol, startRow) = ParseCellReference(parts[0]);
                    var (endCol, endRow) = ParseCellReference(parts[1]);
                    var topRaw = rawCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                    var topFormatted = formattedCells.GetValueOrDefault(parts[0]) ?? string.Empty;
                    maxRow = Math.Max(maxRow, endRow);
                    maxColumn = Math.Max(maxColumn, endCol);

                    for (var row = startRow; row <= endRow; row++)
                    {
                        for (var column = startCol; column <= endCol; column++)
                        {
                            var cellRef = GetCellReference(column, row);
                            rawCells[cellRef] = topRaw;
                            formattedCells[cellRef] = topFormatted;
                        }
                    }
                }

                for (var row = 1; row <= maxRow; row++)
                {
                    var rowData = new List<string>(maxColumn);
                    var formattedRowData = new List<string>(maxColumn);
                    for (var column = 1; column <= maxColumn; column++)
                    {
                        var cellRef = GetCellReference(column, row);
                        rowData.Add(rawCells.GetValueOrDefault(cellRef) ?? string.Empty);
                        formattedRowData.Add(formattedCells.GetValueOrDefault(cellRef) ?? string.Empty);
                    }
                    rows.Add(rowData);
                    formattedRows.Add(formattedRowData);
                }
            }

            var usedRange = maxColumn > 0 && rows.Count > 0
                ? $"A1:{ColumnIndexToName(maxColumn)}{rows.Count}"
                : null;

            sheets.Add(new SheetDataModel(
                sheet.SheetName,
                rows,
                formattedRows,
                cells,
                usedRange,
                mergedRanges,
                formulaCellCount));
        }

        return new WorkbookData(sheets);
    }

    private static string? ResolveOpenXmlFormula(Cell cell, int row, int column, Dictionary<uint, SharedFormula> sharedFormulas)
    {
        var cellFormula = cell.CellFormula;
        if (cellFormula is null)
        {
            return null;
        }

        var formulaText = cellFormula.Text ?? cellFormula.InnerText;
        var sharedIndex = cellFormula.SharedIndex?.Value;
        if (sharedIndex is not null)
        {
            if (!string.IsNullOrWhiteSpace(formulaText))
            {
                sharedFormulas[sharedIndex.Value] = new SharedFormula(formulaText, row, column);
                return formulaText;
            }

            if (sharedFormulas.TryGetValue(sharedIndex.Value, out var sharedFormula))
            {
                return TranslateRelativeFormulaReferences(
                    sharedFormula.Formula,
                    row - sharedFormula.BaseRow,
                    column - sharedFormula.BaseColumn);
            }
        }

        return string.IsNullOrWhiteSpace(formulaText) ? null : formulaText;
    }

    private static string TranslateRelativeFormulaReferences(string formula, int rowOffset, int columnOffset)
    {
        return Regex.Replace(formula, @"(?<![A-Za-z0-9_])(?<column>\$?[A-Z]{1,3})(?<row>\$?\d+)", match =>
        {
            var columnToken = match.Groups["column"].Value;
            var rowToken = match.Groups["row"].Value;
            var absoluteColumn = columnToken.StartsWith('$');
            var absoluteRow = rowToken.StartsWith('$');
            var columnName = absoluteColumn ? columnToken[1..] : columnToken;
            var rowText = absoluteRow ? rowToken[1..] : rowToken;
            var translatedColumn = absoluteColumn
                ? columnName
                : ColumnIndexToName(GetColumnIndex(columnName) + columnOffset);
            var translatedRow = absoluteRow
                ? rowText
                : (int.Parse(rowText, CultureInfo.InvariantCulture) + rowOffset).ToString(CultureInfo.InvariantCulture);
            return $"{(absoluteColumn ? "$" : string.Empty)}{translatedColumn}{(absoluteRow ? "$" : string.Empty)}{translatedRow}";
        });
    }

    internal static string? GetOpenXmlRawCellValue(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.InlineString is not null)
        {
            return cell.InlineString?.InnerText ?? cell.InnerText;
        }

        if (cell.CellValue == null)
        {
            return cell.InnerText;
        }

        var text = cell.CellValue.Text;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
        {
            if (int.TryParse(text, out var index))
            {
                return sharedStringTable.ElementAt(index).InnerText;
            }
        }

        if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
        {
            return text == "1" ? "TRUE" : "FALSE";
        }

        return text;
    }

    internal static string? GetOpenXmlFormattedCellValue(Cell cell, SharedStringTable? sharedStringTable, WorkbookStylesPart? stylesPart)
    {
        var raw = GetOpenXmlRawCellValue(cell, sharedStringTable);
        if (raw is null)
        {
            return null;
        }

        if ((cell.DataType is null || cell.DataType.Value == CellValues.Number) &&
            double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var numericValue))
        {
            return FormatOpenXmlNumber(numericValue, cell.StyleIndex?.Value, stylesPart);
        }

        return raw;
    }

    private static string GetLegacyRawCellValue(ICell? cell)
    {
        if (cell is null)
        {
            return string.Empty;
        }

        return cell.CellType switch
        {
            NpoiCellType.String => cell.StringCellValue ?? string.Empty,
            NpoiCellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? string.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd}", cell.DateCellValue)
                : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            NpoiCellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            NpoiCellType.Formula => GetLegacyRawFormulaValue(cell),
            NpoiCellType.Blank => string.Empty,
            _ => cell.ToString() ?? string.Empty,
        };
    }

    private static string GetLegacyRawFormulaValue(ICell cell)
    {
        return cell.CachedFormulaResultType switch
        {
            NpoiCellType.String => cell.StringCellValue ?? string.Empty,
            NpoiCellType.Numeric => DateUtil.IsCellDateFormatted(cell)
                ? string.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd}", cell.DateCellValue)
                : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
            NpoiCellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
            _ => cell.ToString() ?? string.Empty,
        };
    }

    private static string GetLegacyFormattedCellValue(ICell? cell, DataFormatter formatter)
        => cell is null ? string.Empty : formatter.FormatCellValue(cell);

    private static XlsxNormalizedValue LegacyNormalizedValue(ICell cell)
    {
        var type = cell.CellType == NpoiCellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        if (type == NpoiCellType.Numeric && DateUtil.IsCellDateFormatted(cell))
            return SpreadsheetValueNormalizer.FromLegacyDate(
                cell.DateCellValue ?? throw new InvalidDataException("Legacy date cell has no date value."));
        return new XlsxNormalizedValue(type switch
        {
            NpoiCellType.String => "string",
            NpoiCellType.Numeric => "number",
            NpoiCellType.Boolean => "boolean",
            NpoiCellType.Error => "error",
            NpoiCellType.Blank => "blank",
            _ => "unknown",
        }, null);
    }

    private static IReadOnlyList<RichTextRunReport>? GetLegacyRichTextRuns(ICell? cell, HSSFWorkbook workbook)
    {
        if (cell is null || cell.CellType != NpoiCellType.String || cell.RichStringCellValue is not HSSFRichTextString richText)
        {
            return null;
        }

        var text = richText.String;
        if (string.IsNullOrEmpty(text) || richText.NumFormattingRuns == 0)
        {
            return null;
        }

        var cellFont = cell.CellStyle is not null ? workbook.GetFontAt(cell.CellStyle.FontIndex) : null;
        var runs = new List<RichTextRunReport>();
        var position = 0;
        for (var index = 0; index < richText.NumFormattingRuns; index++)
        {
            var start = richText.GetIndexOfFormattingRun(index);
            if (start > position)
            {
                runs.Add(BuildLegacyRichTextRun(text[position..start], cellFont, workbook));
            }

            var end = index + 1 < richText.NumFormattingRuns
                ? richText.GetIndexOfFormattingRun(index + 1)
                : text.Length;
            if (end <= start)
            {
                continue;
            }

            var fontIndex = richText.GetFontOfFormattingRun(index);
            var font = fontIndex >= 0 ? workbook.GetFontAt(fontIndex) : null;
            runs.Add(BuildLegacyRichTextRun(text[start..end], font, workbook));
            position = end;
        }

        if (position < text.Length)
        {
            runs.Add(BuildLegacyRichTextRun(text[position..], cellFont, workbook));
        }

        return runs.Count == 0 ? null : runs;
    }

    private static RichTextRunReport BuildLegacyRichTextRun(string text, IFont? font, HSSFWorkbook workbook)
        => new(
            text,
            font?.FontName,
            GetLegacyFontColor(font, workbook),
            GetLegacyUnderline(font),
            font?.IsBold ?? false,
            font?.IsItalic ?? false);

    private static string? GetLegacyFontColor(IFont? font, HSSFWorkbook workbook)
    {
        if (font is null)
        {
            return null;
        }

        var colorIndex = font.Color;
        if (colorIndex == (short)FontColor.Normal)
        {
            return null;
        }

        var color = workbook.GetCustomPalette().GetColor(colorIndex);
        if (color is null)
        {
            return colorIndex == 0 ? null : $"indexed:{colorIndex}";
        }

        var triplet = color.GetTriplet();
        return string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", triplet[0], triplet[1], triplet[2]);
    }

    private static string? GetLegacyUnderline(IFont? font)
    {
        if (font is null || font.Underline == FontUnderlineType.None)
        {
            return null;
        }

        return font.Underline.ToString().ToLowerInvariant();
    }

    private static string FormatOpenXmlNumber(double value, uint? styleIndex, WorkbookStylesPart? stylesPart)
    {
        if (!double.IsFinite(value))
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        var formatCode = GetNumberFormatCode(styleIndex, stylesPart);
        if (string.IsNullOrWhiteSpace(formatCode)
            || string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase)
            || string.Equals(formatCode, "@", StringComparison.Ordinal))
        {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        if (IsDateFormat(formatCode))
        {
            try
            {
                return DateTime.FromOADate(value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch (ArgumentException)
            {
                return value.ToString("G15", CultureInfo.InvariantCulture);
            }
        }

        var section = SelectNumericFormatSection(formatCode, value);
        if (string.IsNullOrWhiteSpace(section)
            || string.Equals(section, "General", StringComparison.OrdinalIgnoreCase)
            || string.Equals(section, "@", StringComparison.Ordinal))
        {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        if (TryFormatScientificNumber(value, section, out var scientificFormatted))
        {
            return scientificFormatted;
        }

        var decimals = DecimalPlacesFromSection(section);
        var isPercent = section.Contains('%', StringComparison.Ordinal);
        var formatted = FormatFixedAwayFromZero(value, decimals, isPercent);
        return isPercent ? $"{formatted}%" : formatted;
    }

    private static string SelectNumericFormatSection(string formatCode, double value)
    {
        var sections = StripQuotedAndEscapedFormatText(formatCode).Split(';');
        if (value < 0d && sections.Length >= 2 && !string.IsNullOrWhiteSpace(sections[1]))
        {
            return sections[1];
        }

        if (value == 0d && sections.Length >= 3 && !string.IsNullOrWhiteSpace(sections[2]))
        {
            return sections[2];
        }

        return sections[0];
    }

    private static string FormatFixedAwayFromZero(double value, int decimals, bool isPercent)
    {
        var g15 = value.ToString("G15", CultureInfo.InvariantCulture);
        if (decimal.TryParse(g15, NumberStyles.Float, CultureInfo.InvariantCulture, out var decimalValue))
        {
            var canScale = !isPercent || Math.Abs(decimalValue) <= decimal.MaxValue / 100m;
            if (canScale && decimals <= 28)
            {
                var scaled = isPercent ? decimalValue * 100m : decimalValue;
                var rounded = Math.Round(scaled, decimals, MidpointRounding.AwayFromZero);
                if (rounded == 0m)
                {
                    rounded = 0m;
                }

                return rounded.ToString($"F{decimals}", CultureInfo.InvariantCulture);
            }
        }

        var scaledDouble = isPercent ? value * 100d : value;
        var roundedDouble = decimals <= 15
            ? Math.Round(scaledDouble, decimals, MidpointRounding.AwayFromZero)
            : scaledDouble;
        if (roundedDouble == 0d)
        {
            roundedDouble = 0d;
        }

        return roundedDouble.ToString($"F{decimals}", CultureInfo.InvariantCulture);
    }

    private static bool TryFormatScientificNumber(double value, string section, out string formatted)
    {
        var match = Regex.Match(section, @"(?<mantissa>[0#?.,]+)(?<expMarker>[Ee])(?<expSign>[+-])(?<expDigits>[0#?]+)");
        if (!match.Success)
        {
            formatted = string.Empty;
            return false;
        }

        var mantissaPattern = match.Groups["mantissa"].Value;
        var expMarker = match.Groups["expMarker"].Value;
        var expSign = match.Groups["expSign"].Value;
        var expDigitsPattern = match.Groups["expDigits"].Value;
        var decimals = DecimalPlacesFromSection(mantissaPattern);
        var expMinDigits = Math.Max(1, expDigitsPattern.Count(c => c == '0'));

        var decimalIndex = mantissaPattern.IndexOf('.', StringComparison.Ordinal);
        var integerPart = decimalIndex >= 0 ? mantissaPattern[..decimalIndex] : mantissaPattern;
        var integerPlaceholders = integerPart.Count(c => c is '0' or '#' or '?');
        var exponentStep = integerPlaceholders > 1 && integerPart.Contains('#', StringComparison.Ordinal)
            ? integerPlaceholders
            : 1;

        if (value == 0d)
        {
            var zeroMantissa = 0m.ToString($"F{decimals}", CultureInfo.InvariantCulture);
            var zeroExpSign = expSign == "+" ? "+" : string.Empty;
            var zeroExpDigits = 0.ToString($"D{expMinDigits}", CultureInfo.InvariantCulture);
            formatted = $"{zeroMantissa}{expMarker}{zeroExpSign}{zeroExpDigits}";
            return true;
        }

        var isNegative = value < 0d;
        var e14 = Math.Abs(value).ToString("E14", CultureInfo.InvariantCulture);
        var parts = e14.Split('E');
        var mantissa = decimal.Parse(parts[0], CultureInfo.InvariantCulture);
        var exponent = int.Parse(parts[1], CultureInfo.InvariantCulture);

        if (exponentStep > 1)
        {
            var adjustedExponent = (int)Math.Floor((double)exponent / exponentStep) * exponentStep;
            var shift = exponent - adjustedExponent;
            mantissa *= Pow10Decimal(shift);
            exponent = adjustedExponent;
        }

        var roundDecimals = Math.Min(decimals, 28);
        var roundedMantissa = Math.Round(mantissa, roundDecimals, MidpointRounding.AwayFromZero);
        var carryThreshold = Pow10Decimal(exponentStep);
        if (roundedMantissa >= carryThreshold)
        {
            roundedMantissa /= carryThreshold;
            exponent += exponentStep;
        }

        var signText = isNegative ? "-" : string.Empty;
        var mantissaText = roundedMantissa.ToString($"F{decimals}", CultureInfo.InvariantCulture);
        var expSignText = exponent < 0 ? "-" : (expSign == "+" ? "+" : string.Empty);
        var expDigitsText = Math.Abs(exponent).ToString($"D{expMinDigits}", CultureInfo.InvariantCulture);
        formatted = $"{signText}{mantissaText}{expMarker}{expSignText}{expDigitsText}";
        return true;
    }

    private static decimal Pow10Decimal(int power)
    {
        var result = 1m;
        for (var i = 0; i < power; i++)
        {
            result *= 10m;
        }

        return result;
    }

    private static string? GetNumberFormatCode(uint? styleIndex, WorkbookStylesPart? stylesPart)
    {
        if (styleIndex is null || stylesPart?.Stylesheet?.CellFormats is null)
        {
            return null;
        }

        var stylesheet = stylesPart.Stylesheet;
        var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
        if (styleIndex.Value >= cellFormats.Count)
        {
            return null;
        }

        var cellFormat = cellFormats[(int)styleIndex.Value];
        var baseFormats = stylesheet.CellStyleFormats?.Elements<CellFormat>().ToList() ?? [];
        var baseFormatIndex = cellFormat.FormatId?.Value;
        var baseFormat = baseFormatIndex is not null && baseFormatIndex.Value < baseFormats.Count
            ? baseFormats[(int)baseFormatIndex.Value]
            : null;
        var numberFormatId = SpreadsheetValueNormalizer.EffectiveComponentId(
            cellFormat.NumberFormatId?.Value,
            cellFormat.ApplyNumberFormat?.Value,
            baseFormat?.NumberFormatId?.Value);

        var custom = stylesheet.NumberingFormats?
            .Elements<NumberingFormat>()
            .LastOrDefault(format => format.NumberFormatId?.Value == numberFormatId)
            ?.FormatCode?.Value;
        if (!string.IsNullOrWhiteSpace(custom))
        {
            return custom;
        }

        return BuiltInNumberFormat(numberFormatId);
    }

    private static string? BuiltInNumberFormat(uint numberFormatId) => numberFormatId switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0;(#,##0)",
        38 => "#,##0;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };

    private static bool IsDateFormat(string formatCode)
    {
        var cleaned = StripQuotedAndEscapedFormatText(formatCode).ToLowerInvariant();
        return cleaned.Contains('y') || cleaned.Contains('d') || cleaned.Contains("m/");
    }

    private static int DecimalPlacesFromSection(string section)
    {
        var decimalIndex = section.IndexOf('.', StringComparison.Ordinal);
        if (decimalIndex < 0)
        {
            return 0;
        }

        var count = 0;
        for (var i = decimalIndex + 1; i < section.Length; i++)
        {
            if (section[i] is '0' or '#' or '?')
            {
                count++;
                continue;
            }
            break;
        }

        return count;
    }

    private static string StripQuotedAndEscapedFormatText(string formatCode)
    {
        var chars = new List<char>();
        var inQuote = false;
        for (var i = 0; i < formatCode.Length; i++)
        {
            var c = formatCode[i];
            if (c == '"')
            {
                inQuote = !inQuote;
                continue;
            }
            if (inQuote)
            {
                continue;
            }
            if (c == '\\' || c == '_' || c == '*')
            {
                i++;
                continue;
            }
            if (c == '[')
            {
                while (i < formatCode.Length && formatCode[i] != ']')
                {
                    i++;
                }
                continue;
            }
            chars.Add(c);
        }

        return new string(chars.ToArray());
    }

    private static (int Column, int Row) ParseCellReference(string reference)
    {
        var colStr = new string(reference.TakeWhile(char.IsLetter).ToArray());
        var rowStr = new string(reference.Skip(colStr.Length).ToArray());
        var col = GetColumnIndex(colStr);
        var row = int.Parse(rowStr, CultureInfo.InvariantCulture);
        return (col, row);
    }

    private static int GetColumnIndex(string columnName)
    {
        var sum = 0;
        foreach (var c in columnName.ToUpperInvariant())
        {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum;
    }

    private static string GetCellReference(int column, int row)
    {
        return $"{ColumnIndexToName(column)}{row}";
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
