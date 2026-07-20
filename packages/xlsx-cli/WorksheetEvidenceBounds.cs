using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace Dockit.Xlsx;

internal static partial class WorksheetEvidenceBounds
{
    public static int SignificantColumn(WorksheetPart worksheetPart)
    {
        var maximum = 1;
        foreach (var cell in worksheetPart.Worksheet.Descendants<Cell>())
        {
            if (!HasValueOrFormula(cell)) continue;
            maximum = Math.Max(maximum, Column(cell.CellReference?.Value));
        }
        foreach (var merge in worksheetPart.Worksheet.Elements<MergeCells>().SelectMany(value => value.Elements<MergeCell>()))
        {
            maximum = Math.Max(maximum, RangeEndColumn(merge.Reference?.Value));
        }
        foreach (var table in worksheetPart.TableDefinitionParts.Select(part => part.Table))
        {
            maximum = Math.Max(maximum, RangeEndColumn(table.Reference?.Value));
        }
        return maximum;
    }

    public static bool Include(Cell cell, int significantColumn) =>
        HasValueOrFormula(cell) || Column(cell.CellReference?.Value) <= significantColumn;

    public static int Column(string? reference)
    {
        var match = CellReference().Match(reference ?? string.Empty);
        if (!match.Success) return 0;
        var result = 0;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant()) result = checked(result * 26 + character - 'A' + 1);
        return result;
    }

    private static int RangeEndColumn(string? reference)
    {
        var end = (reference ?? string.Empty).Split(':').LastOrDefault();
        return Column(end?.Replace("$", string.Empty, StringComparison.Ordinal));
    }

    private static bool HasValueOrFormula(Cell cell) => cell.CellFormula is not null
        || cell.InlineString is not null
        || cell.CellValue?.Text is not null;

    [GeneratedRegex(@"^\$?([A-Z]+)\$?[0-9]+$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex CellReference();
}
