namespace Dockit.Convert;

public sealed record DocumentFieldRefreshResult(string Backend);

public static class DocumentFieldRefresher
{
    public static DocumentFieldRefreshResult RefreshDocxFields(string input, string output)
    {
        if (!File.Exists(input))
            throw new InvalidOperationException($"Input file not found: {input}");
        if (!string.Equals(Path.GetExtension(input), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Document field refresh input must be a DOCX file.");
        if (!string.Equals(Path.GetExtension(output), ".docx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Document field refresh output must be a DOCX file.");
        if (string.Equals(Path.GetFullPath(input), Path.GetFullPath(output), StringComparison.Ordinal))
            throw new InvalidOperationException("Document field refresh requires distinct input and output paths.");

        if (WpsPdfConverter.IsAvailable())
        {
            WpsPdfConverter.RefreshDocxFields(input, output);
            return new DocumentFieldRefreshResult("wps");
        }
        if (LimaWpsPdfConverter.IsAvailable())
        {
            LimaWpsPdfConverter.RefreshDocxFields(input, output);
            return new DocumentFieldRefreshResult("wps");
        }

        throw new InvalidOperationException(
            "WPS Writer is required to refresh DOCX layout fields, but neither a local runtime nor a configured Lima WPS instance is available.");
    }
}
