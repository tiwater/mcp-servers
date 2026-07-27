namespace Dockit.Convert;

public static class WorkbookRecalculator
{
    public sealed record RecalculationResult(string Backend, string? FallbackReason = null);

    public static RecalculationResult RecalculateXlsx(string input, string output)
    {
        var required = Environment.GetEnvironmentVariable("TIWATER_OFFICE_XLSX_BACKEND")?.Trim();
        if (!string.IsNullOrWhiteSpace(required) && !string.Equals(required, "et", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException($"Unsupported required XLSX backend: {required}");
        if (EtRecalculator.IsAvailable()) { EtRecalculator.Recalculate(input, output); return new("et"); }
        if (LimaWpsPdfConverter.IsAvailable()) { LimaWpsPdfConverter.RecalculateXlsx(input, output); return new("et"); }
        throw new InvalidOperationException("ET XLSX recalculation was required but neither local WPS RPC nor a configured Lima WPS runtime is available.");
    }
}
