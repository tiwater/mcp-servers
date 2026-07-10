namespace Dockit.Convert;

public static class OfficePdfConverter
{
    public sealed record ConversionResult(string Backend, string? FallbackReason = null);

    public static string? FindSofficeBinary()
    {
        return OfficeConverter.FindSofficeBinary();
    }

    public static bool ShouldUseWps(string sourceFormat, bool wpsAvailable)
    {
        var normalized = sourceFormat.Trim().TrimStart('.').ToLowerInvariant();
        return wpsAvailable && normalized is "doc" or "docx" or "odt" or "rtf";
    }

    public static ConversionResult ConvertToPdf(string input, string output, string sourceFormat, string? sofficePath = null)
    {
        if (ShouldUseWps(sourceFormat, WpsWriterConverter.IsAvailable()))
        {
            try
            {
                WpsWriterConverter.ConvertToPdf(input, output);
                return new ConversionResult("wps-writer");
            }
            catch (Exception ex)
            {
                var fallbackReason = $"WPS Writer RPC conversion failed: {ex.Message}";
                try
                {
                    OfficeConverter.ConvertToPdf(input, output, sourceFormat, sofficePath);
                    return new ConversionResult("libreoffice", fallbackReason);
                }
                catch (Exception fallback)
                {
                    throw new InvalidOperationException($"{fallbackReason}; {fallback.Message}", fallback);
                }
            }
        }

        OfficeConverter.ConvertToPdf(input, output, sourceFormat, sofficePath);
        return new ConversionResult("libreoffice");
    }
}
