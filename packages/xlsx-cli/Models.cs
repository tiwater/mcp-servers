using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace Dockit.Xlsx;

internal static class XlsxToolVersion
{
    internal static readonly string Current = typeof(XlsxToolVersion).Assembly
        .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
        .Cast<System.Reflection.AssemblyInformationalVersionAttribute>().Single().InformationalVersion.Split('+', 2)[0];
}

public record WorkbookReport(
    string File,
    int SheetCount,
    List<SheetReport> Sheets
);

public record RichTextRunReport(
    string Text,
    string? FontName = null,
    string? Color = null,
    string? Underline = null,
    bool Bold = false,
    bool Italic = false);

public record TextCellReport(string Reference, string Text, IReadOnlyList<RichTextRunReport>? RichTextRuns = null);

public record FormulaCellReport(string Reference, string Formula, string? CachedValue);

public record CellStyleReport(
    uint StyleIndex,
    uint NumberFormatId,
    string? NumberFormatCode,
    uint FontId,
    uint FillId,
    uint BorderId,
    string? HorizontalAlignment,
    string? VerticalAlignment,
    bool WrapText);

public record CellEvidenceReport(
    string Reference,
    string Value,
    string? Formula,
    CellStyleReport Style,
    IReadOnlyList<RichTextRunReport>? RichTextRuns = null);

public record RowHeightReport(uint Row, double Height);

public record ColumnWidthReport(uint Column, double Width);

public record SheetReport(
    string Name,
    int RowCount,
    int ColumnCount,
    List<string> Placeholders,
    List<string> TablePlaceholders,
    string? UsedRange = null,
    List<string>? MergedRanges = null,
    int FormulaCellCount = 0,
    List<TextCellReport>? TextCells = null,
    List<FormulaCellReport>? FormulaCells = null,
    List<RowHeightReport>? RowHeights = null,
    List<ColumnWidthReport>? ColumnWidths = null,
    List<CellEvidenceReport>? Cells = null
);

public record FillData(
    Dictionary<string, string> CellValues,
    Dictionary<string, List<List<string>>>? TableData
);

public sealed record XlsxEditOperation(
    string Type,
    string? Sheet = null,
    string? Cell = null,
    [property: JsonConverter(typeof(PrimitiveJsonValueConverter))] string? Value = null,
    string? ValueType = null,
    string? StartCell = null,
    IReadOnlyList<IReadOnlyList<string>>? Values = null,
    bool? Bold = null,
    bool? ShrinkToFit = null,
    bool? WrapText = null,
    int? StartRow = null,
    int? Count = null,
    int? SourceRow = null,
    int? TargetRow = null,
    bool? TranslateFormulas = null,
    string? AnchorText = null,
    int? ExampleRows = null,
    int? TargetRows = null,
    bool? PreserveStyle = null,
    bool? PreserveFormulas = null,
    bool? PreserveMergedRanges = null,
    string? Range = null,
    string? Column = null,
    double? Width = null,
    int? FitToPagesWide = null,
    int? FitToPagesTall = null,
    string? Orientation = null,
    string? PaperSize = null,
    int? RepeatRowsStart = null,
    int? RepeatRowsEnd = null);

internal sealed class PrimitiveJsonValueConverter : JsonConverter<string?>
{
    public override string? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) => reader.TokenType switch
    {
        JsonTokenType.Null => null,
        JsonTokenType.String => reader.GetString(),
        JsonTokenType.Number => reader.GetDecimal().ToString(CultureInfo.InvariantCulture),
        JsonTokenType.True => bool.TrueString.ToLowerInvariant(),
        JsonTokenType.False => bool.FalseString.ToLowerInvariant(),
        _ => throw new JsonException("XLSX edit value must be a JSON string, number, boolean, or null."),
    };

    public override void Write(Utf8JsonWriter writer, string? value, JsonSerializerOptions options)
    {
        if (value is null) writer.WriteNullValue();
        else writer.WriteStringValue(value);
    }
}

public sealed record XlsxEditDocument(
    IReadOnlyList<XlsxEditOperation> Operations
);

public sealed record XlsxEditAppliedOperation(
    string Type,
    bool Applied,
    string Detail,
    string? Sheet = null,
    string? ChangedRange = null,
    IReadOnlyList<string>? Warnings = null
);

public sealed record XlsxEditResult(
    string Input,
    string Output,
    IReadOnlyList<XlsxEditAppliedOperation> AppliedOperations
);

public sealed record XlsxValidationResult(
    string File,
    bool Valid,
    IReadOnlyList<string> Errors,
    IReadOnlyList<string> Warnings
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
