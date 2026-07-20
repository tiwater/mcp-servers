using Dockit.Convert;
using Xunit;

namespace Dockit.Convert.Tests;

public class LimaWpsWriterPdfConverterTests
{
    [Fact]
    public void Lima_backend_starts_the_configured_instance_with_only_isolated_paths()
    {
        var staging = Path.Combine(Path.GetTempPath(), $"tiwater-lima-{Guid.NewGuid():N}");
        var input = Path.Combine(staging, "input.docx");
        var output = Path.Combine(staging, "output.pdf");

        var startInfo = LimaWpsWriterPdfConverter.CreateProcessStartInfo(
            "/usr/local/bin/limactl",
            "lucid-wps",
            input,
            output);

        Assert.Equal("/usr/local/bin/limactl", startInfo.FileName);
        Assert.Equal(new[] { "shell", "lucid-wps", "--", "bash", "-lc" }, startInfo.ArgumentList.Take(5));
        Assert.Contains(input, startInfo.ArgumentList[5]);
        Assert.Contains(output, startInfo.ArgumentList[5]);
        Assert.Contains("TIWATER_OFFICE_PDF_BACKEND=wps-writer", startInfo.ArgumentList[5]);
        Assert.DoesNotContain("soffice", startInfo.ArgumentList[5], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Lima_backend_preserves_the_writer_input_format()
    {
        var startInfo = LimaWpsWriterPdfConverter.CreateProcessStartInfo(
            "/usr/local/bin/limactl",
            "lucid-wps",
            "/tmp/lucid-wps-render/input.rtf",
            "/tmp/lucid-wps-render/output.pdf");

        Assert.Contains("rtf-to-pdf", startInfo.ArgumentList[5]);
    }

    [Fact]
    public void Lima_backend_runs_xls_conversion_with_required_wps_spreadsheet_identity()
    {
        var startInfo = LimaWpsWriterPdfConverter.CreateSpreadsheetConversionStartInfo(
            "/usr/local/bin/limactl",
            "lucid-wps",
            "/tmp/lucid-wps-render/input.xls",
            "/tmp/lucid-wps-render/output.xlsx");

        Assert.Equal("/usr/local/bin/limactl", startInfo.FileName);
        Assert.Equal(new[] { "shell", "lucid-wps", "--", "bash", "-lc" }, startInfo.ArgumentList.Take(5));
        Assert.Contains("TIWATER_OFFICE_XLSX_BACKEND=wps-spreadsheet", startInfo.ArgumentList[5]);
        Assert.Contains("xls-to-xlsx", startInfo.ArgumentList[5]);
        Assert.DoesNotContain("soffice", startInfo.ArgumentList[5], StringComparison.OrdinalIgnoreCase);
    }
}
