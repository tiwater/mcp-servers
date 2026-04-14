using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;

namespace Dockit.Xlsx;

public record WorkbookReport(
    string File,
    int SheetCount,
    List<SheetReport> Sheets
);

public record SheetReport(
    string Name,
    int RowCount,
    int ColumnCount,
    List<string> Placeholders,
    List<string> TablePlaceholders
);

public record FillData(
    Dictionary<string, string> CellValues,
    Dictionary<string, List<List<string>>>? TableData
);

internal static class Json
{
    public static JsonSerializerOptions Options => new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
    };
}
