namespace Dockit.Convert;

public static class WorkbookRecalculator
{
    public sealed record RecalculationResult(string Backend, string? FallbackReason = null);

    public static RecalculationResult RecalculateXlsx(string input, string output)
    {
        var required = Environment.GetEnvironmentVariable("TIWATER_OFFICE_XLSX_BACKEND")?.Trim();
        if (!string.IsNullOrWhiteSpace(required) && !string.Equals(required, "wps-spreadsheet", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException($"Unsupported required XLSX backend: {required}");
        if (WpsSpreadsheetRecalculator.IsAvailable()) { WpsSpreadsheetRecalculator.Recalculate(input, output); return new("wps-spreadsheet"); }
        if (LimaWpsWriterPdfConverter.IsAvailable()) { LimaWpsWriterPdfConverter.RecalculateXlsx(input, output); return new("wps-spreadsheet"); }
        throw new InvalidOperationException("WPS Spreadsheet XLSX recalculation was required but neither local WPS RPC nor a configured Lima WPS runtime is available.");
    }
}
