using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

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
    int FormulaCellCount = 0
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
    string? ValueType = null,
    string? StartCell = null,
    IReadOnlyList<IReadOnlyList<string>>? Values = null,
    bool? Bold = null,
    int? SourceRow = null,
    int? TargetRow = null,
    int? Count = null,
    bool? CopyValues = null
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
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
    };
}
