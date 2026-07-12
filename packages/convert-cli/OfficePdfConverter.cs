namespace Dockit.Convert;

public static class OfficePdfConverter
{
    public static string? FindSofficeBinary()
    {
        return OfficeConverter.FindSofficeBinary();
    }

    public static OfficePdfConversionResult ConvertToPdf(string input, string output, string sourceFormat, string? sofficePath = null)
    {
        return OfficeConverter.ConvertToPdf(input, output, sourceFormat, sofficePath);
    }
}
