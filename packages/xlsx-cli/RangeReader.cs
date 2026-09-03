using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace Dockit.Xlsx;

public sealed record XlsxRangeReadReceipt(
    string Schema,
    long TotalCellCount,
    int ReturnedCellCount,
    long Remaining,
    long? NextOffset);

public sealed record XlsxRangeFormula(
    string Text,
    string? Type,
    uint? SharedIndex,
    string? Reference);

public sealed record XlsxRangeCell(
    string Reference,
    int Row,
    int Column,
    bool Physical,
    string? RawValue,
    string? FormattedValue,
    string? ValueType,
    XlsxNormalizedValue? NormalizedValue,
    XlsxRangeFormula? Formula,
    CellStyleReport? Style,
    IReadOnlyList<RichTextRunReport>? RichTextRuns,
    string? MergedRange,
    string? MergeOwner);

public sealed record XlsxRangeReadResult(
    string Schema,
    string ToolVersion,
    string File,
    string InputSha256,
    string Sheet,
    string Range,
    XlsxRangeReadReceipt Receipt,
    IReadOnlyList<XlsxRangeCell> Cells);

public static partial class RangeReader
{
    public const int MaximumPageCells = 256;

    public static XlsxRangeReadResult Read(
        string input,
        string sheetName,
        string range,
        long offset,
        int limit)
    {
        var fullPath = Path.GetFullPath(input);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("Workbook not found.", fullPath);
        if (!string.Equals(Path.GetExtension(fullPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("xlsx_read_range requires a current XLSX workbook; convert legacy XLS first.");
        if (string.IsNullOrWhiteSpace(sheetName)) throw new InvalidOperationException("sheet-is-required");
        if (offset < 0) throw new InvalidOperationException("offset-must-be-nonnegative");
        if (limit is < 1 or > MaximumPageCells)
            throw new InvalidOperationException($"limit-must-be-between-1-and-{MaximumPageCells}");

        var selectedRange = ParseRange(range);
        using var spreadsheet = SpreadsheetDocument.Open(fullPath, false);
        var workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .SingleOrDefault(candidate => string.Equals(candidate.Name?.Value, sheetName, StringComparison.Ordinal));
        if (sheet?.Id?.Value is not string relationshipId
            || workbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            throw new InvalidOperationException($"Worksheet not found: {sheetName}");

        var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
        var stylesPart = workbookPart.WorkbookStylesPart;
        var stylesheet = stylesPart?.Stylesheet;
        var uses1904Dates = workbookPart.Workbook.WorkbookProperties?.Date1904?.Value == true;
        var existingCells = worksheetPart.Worksheet.Descendants<Cell>()
            .Where(cell => !string.IsNullOrWhiteSpace(cell.CellReference?.Value))
            .ToDictionary(cell => cell.CellReference!.Value!.ToUpperInvariant(), StringComparer.OrdinalIgnoreCase);
        var mergedRanges = worksheetPart.Worksheet.Elements<MergeCells>()
            .SelectMany(container => container.Elements<MergeCell>())
            .Select(merge => merge.Reference?.Value)
            .Where(reference => !string.IsNullOrWhiteSpace(reference))
            .Select(reference => ParseRange(reference!))
            .ToList();

        var totalCellCount = checked((long)selectedRange.RowCount * selectedRange.ColumnCount);
        var startOffset = Math.Min(offset, totalCellCount);
        var returnedCellCount = (int)Math.Min(limit, totalCellCount - startOffset);
        var cells = new List<XlsxRangeCell>(returnedCellCount);
        for (var index = 0; index < returnedCellCount; index++)
        {
            var ordinal = startOffset + index;
            var row = checked(selectedRange.StartRow + (int)(ordinal / selectedRange.ColumnCount));
            var column = checked(selectedRange.StartColumn + (int)(ordinal % selectedRange.ColumnCount));
            var reference = CellReference(column, row);
            existingCells.TryGetValue(reference, out var cell);
            var merge = mergedRanges.FirstOrDefault(candidate => candidate.Contains(row, column));
            var rawValue = cell is null ? null : WorkbookLoader.GetOpenXmlRawCellValue(cell, sharedStrings);
            var formula = cell?.CellFormula;
            cells.Add(new XlsxRangeCell(
                reference,
                row,
                column,
                cell is not null,
                rawValue,
                cell is null ? null : WorkbookLoader.GetOpenXmlFormattedCellValue(cell, sharedStrings, stylesPart),
                cell?.DataType?.InnerText ?? (cell is null ? null : "number"),
                cell is null ? null : SpreadsheetValueNormalizer.FromOpenXmlCell(
                    cell.CellValue?.Text ?? cell.InlineString?.InnerText,
                    cell.DataType?.InnerText,
                    stylesPart,
                    cell.StyleIndex?.Value,
                    uses1904Dates),
                formula is null ? null : new XlsxRangeFormula(
                    formula.Text ?? string.Empty,
                    formula.FormulaType?.InnerText,
                    formula.SharedIndex?.Value,
                    formula.Reference?.Value),
                cell is null ? null : Inspector.GetCellStyle(cell, stylesheet),
                cell is null ? null : OpenXmlRichText.GetCellRichTextRuns(cell, sharedStrings),
                merge?.Canonical,
                merge is null ? null : CellReference(merge.StartColumn, merge.StartRow)));
        }

        var nextOffset = startOffset + returnedCellCount < totalCellCount
            ? startOffset + returnedCellCount
            : (long?)null;
        return new XlsxRangeReadResult(
            "tiwater.xlsx-range-page/v1",
            XlsxToolVersion.Current,
            fullPath,
            Sha256(fullPath),
            sheetName,
            selectedRange.Canonical,
            new XlsxRangeReadReceipt(
                "tiwater.xlsx-range-page-receipt/v1",
                totalCellCount,
                returnedCellCount,
                totalCellCount - startOffset - returnedCellCount,
                nextOffset),
            cells);
    }

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException("xlsx_read_range requires <request.json>");
        var request = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
            ?? throw new InvalidOperationException("xlsx-read-range-request-invalid");
        var input = RequiredString(request, "input");
        var sheet = RequiredString(request, "sheet");
        var range = RequiredString(request, "range");
        var offset = OptionalLong(request, "offset") ?? 0;
        var limit = OptionalInt(request, "limit")
            ?? throw new InvalidOperationException("limit-is-required");
        Console.WriteLine(JsonSerializer.Serialize(Read(input, sheet, range, offset, limit), Json.Options));
        return 0;
    }

    private static string RequiredString(JsonObject request, string property)
        => request[property] is JsonValue value
           && value.TryGetValue<string>(out var text)
           && !string.IsNullOrWhiteSpace(text)
            ? text
            : throw new InvalidOperationException($"{property}-is-required");

    private static int? OptionalInt(JsonObject request, string property)
        => request[property] is JsonValue value && value.TryGetValue<int>(out var number) ? number : null;

    private static long? OptionalLong(JsonObject request, string property)
        => request[property] is JsonValue value && value.TryGetValue<long>(out var number) ? number : null;

    private static RangeAddress ParseRange(string value)
    {
        var match = RangePattern().Match(value ?? string.Empty);
        if (!match.Success) throw new InvalidOperationException($"Invalid A1 range: {value}");
        var startColumn = ColumnNumber(match.Groups[1].Value);
        var startRow = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        var endColumn = match.Groups[3].Success ? ColumnNumber(match.Groups[3].Value) : startColumn;
        var endRow = match.Groups[4].Success
            ? int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture)
            : startRow;
        if (startColumn > endColumn || startRow > endRow)
            throw new InvalidOperationException($"A1 range must be top-left to bottom-right: {value}");
        if (endColumn > 16_384 || endRow > 1_048_576)
            throw new InvalidOperationException($"A1 range exceeds XLSX worksheet bounds: {value}");
        return new RangeAddress(startColumn, startRow, endColumn, endRow);
    }

    private static int ColumnNumber(string value)
    {
        var result = 0;
        foreach (var character in value.ToUpperInvariant())
            result = checked(result * 26 + character - 'A' + 1);
        return result;
    }

    private static string CellReference(int column, int row)
    {
        var name = string.Empty;
        while (column > 0)
        {
            column--;
            name = (char)('A' + column % 26) + name;
            column /= 26;
        }
        return $"{name}{row}";
    }

    private static string Sha256(string path)
    {
        using var stream = File.OpenRead(path);
        return System.Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private sealed record RangeAddress(int StartColumn, int StartRow, int EndColumn, int EndRow)
    {
        internal int ColumnCount => EndColumn - StartColumn + 1;
        internal int RowCount => EndRow - StartRow + 1;
        internal string Canonical => $"{CellReference(StartColumn, StartRow)}:{CellReference(EndColumn, EndRow)}";
        internal bool Contains(int row, int column) =>
            row >= StartRow && row <= EndRow && column >= StartColumn && column <= EndColumn;
    }

    [GeneratedRegex("^([A-Za-z]{1,3})([1-9][0-9]*)(?::([A-Za-z]{1,3})([1-9][0-9]*))?$", RegexOptions.CultureInvariant)]
    private static partial Regex RangePattern();
}
