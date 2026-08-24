using Dockit.Convert;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace Dockit.Convert.Tests;

public class LimaWpsPdfConverterTests
{
    [Fact]
    public void Lima_backend_starts_the_configured_instance_with_only_isolated_paths()
    {
        var staging = Path.Combine(Path.GetTempPath(), $"tiwater-lima-{Guid.NewGuid():N}");
        var input = Path.Combine(staging, "input.docx");
        var output = Path.Combine(staging, "output.pdf");

        var startInfo = LimaWpsPdfConverter.CreateProcessStartInfo(
            "/usr/local/bin/limactl",
            "tiwater-office",
            input,
            output);

        Assert.Equal("/usr/local/bin/limactl", startInfo.FileName);
        Assert.Equal(new[] { "shell", "tiwater-office", "--", "bash", "-lc" }, startInfo.ArgumentList.Take(5));
        Assert.Contains(input, startInfo.ArgumentList[5]);
        Assert.Contains(output, startInfo.ArgumentList[5]);
        Assert.Contains("TIWATER_OFFICE_PDF_BACKEND=wps", startInfo.ArgumentList[5]);
        Assert.DoesNotContain("soffice", startInfo.ArgumentList[5], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Lima_backend_preserves_the_writer_input_format()
    {
        var startInfo = LimaWpsPdfConverter.CreateProcessStartInfo(
            "/usr/local/bin/limactl",
            "tiwater-office",
            "/tmp/tiwater-wps-render/input.rtf",
            "/tmp/tiwater-wps-render/output.pdf");

        Assert.Contains("rtf-to-pdf", startInfo.ArgumentList[5]);
    }

    [Fact]
    public void Lima_backend_runs_xls_conversion_with_required_wps_spreadsheet_identity()
    {
        var startInfo = LimaWpsPdfConverter.CreateSpreadsheetConversionStartInfo(
            "/usr/local/bin/limactl",
            "tiwater-office",
            "/tmp/tiwater-wps-render/input.xls",
            "/tmp/tiwater-wps-render/output.xlsx");

        Assert.Equal("/usr/local/bin/limactl", startInfo.FileName);
        Assert.Equal(new[] { "shell", "tiwater-office", "--", "bash", "-lc" }, startInfo.ArgumentList.Take(5));
        Assert.Contains("TIWATER_OFFICE_XLSX_BACKEND=et", startInfo.ArgumentList[5]);
        Assert.Contains("xls-to-xlsx", startInfo.ArgumentList[5]);
        Assert.DoesNotContain("soffice", startInfo.ArgumentList[5], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Lima_backend_runs_docx_field_refresh_with_native_wps_identity()
    {
        var startInfo = LimaWpsPdfConverter.CreateDocumentFieldRefreshStartInfo(
            "/usr/local/bin/limactl",
            "tiwater-office",
            "/tmp/tiwater-wps-render/input.docx",
            "/tmp/tiwater-wps-render/output.docx");

        Assert.Equal("/usr/local/bin/limactl", startInfo.FileName);
        Assert.Contains("refresh-docx-fields", startInfo.ArgumentList[5]);
        Assert.Contains("TIWATER_WPSRPC_PYTHON", startInfo.ArgumentList[5]);
        Assert.DoesNotContain("soffice", startInfo.ArgumentList[5], StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Lima_docx_field_refresh_evidence_binds_scope_and_exact_bytes()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docx-refresh-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var input = Path.Combine(root, "input.docx");
        var output = Path.Combine(root, "output.docx");
        File.WriteAllBytes(input, [0x50, 0x4b, 0x03, 0x04, 0x01]);
        File.WriteAllBytes(output, [0x50, 0x4b, 0x03, 0x04, 0x02]);
        var sha = (string file) => System.Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file))).ToLowerInvariant();
        var receipt = JsonSerializer.Serialize(new
        {
            schema = "tiwater.convert-refresh-docx-fields/v1", status = "ok", backend = "wps",
            source_format = "docx", target_format = "docx",
            refresh_scope = new[] { "table-of-contents", "table-of-figures" },
            input_sha256 = sha(input), output_sha256 = sha(output),
        });

        LimaWpsPdfConverter.ValidateDocumentFieldRefreshEvidence(receipt, input, output);
        File.AppendAllText(output, "changed");
        Assert.Throws<InvalidOperationException>(
            () => LimaWpsPdfConverter.ValidateDocumentFieldRefreshEvidence(receipt, input, output));
    }
}
