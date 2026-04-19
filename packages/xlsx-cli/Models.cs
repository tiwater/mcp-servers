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
    List<string> TablePlaceholders,
    string? UsedRange = null,
    List<string>? MergedRanges = null,
    int FormulaCellCount = 0,
    List<NoteRowReport>? NoteRows = null
);

public record NoteRowReport(
    int RowIndex,
    string Text
);

public record FillData(
    Dictionary<string, string> CellValues,
    Dictionary<string, List<List<string>>>? TableData
);

public sealed record XlsxEditOperation(
    string Type,
    string? Sheet = null,
    string? Cell = null,
    string? Value = null,
    string? StartCell = null,
    IReadOnlyList<IReadOnlyList<string>>? Values = null
);

public sealed record XlsxEditDocument(
    IReadOnlyList<XlsxEditOperation> Operations
);

public sealed record XlsxEditAppliedOperation(
    string Type,
    bool Applied,
    string Detail
);

public sealed record XlsxEditResult(
    string Input,
    string Output,
    IReadOnlyList<XlsxEditAppliedOperation> AppliedOperations
);

internal static class Json
{
    public static JsonSerializerOptions Options => new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
    };
}
