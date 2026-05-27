using Dockit.Xlsx;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace Dockit.Xlsx.Tests;

public class ValidationTests
{
    [Fact]
    public void Validate_accepts_generated_workbook_with_moved_merges_and_formulas()
    {
        var input = WorkbookFixtures.CreateAna14LikeWorkbookWithStyles();
        var output = Path.Combine(Path.GetTempPath(), $"xlsx-validate-generated-{Guid.NewGuid():N}.xlsx");

        Editor.Apply(input, output, [
            new XlsxEditOperation(
                "expandSectionRows",
                Sheet: "RP",
                AnchorText: "impurity peak area",
                ExampleRows: 2,
                TargetRows: 4)
        ]);

        var result = Validator.Validate(output);

        Assert.True(result.Valid, string.Join(Environment.NewLine, result.Errors));
        Assert.Empty(result.Errors);
    }

    [Fact]
    public void Validate_rejects_non_xlsx_text_file()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-invalid-text-{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(path, "this is not an xlsx package");

        var result = Validator.Validate(path);

        Assert.False(result.Valid);
        Assert.Contains(result.Errors, error => error.Contains("valid XLSX package", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_warns_when_error_output_reaches_cap()
    {
        var path = CreateWorkbookWithManyValidationErrors();

        var result = Validator.Validate(path);

        Assert.False(result.Valid);
        Assert.Equal(100, result.Errors.Count);
        Assert.Contains(result.Warnings, warning => warning.Contains("truncated", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Cli_validate_exits_one_and_emits_json_for_invalid_non_xlsx_file()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-cli-invalid-text-{Guid.NewGuid():N}.xlsx");
        File.WriteAllText(path, "this is not an xlsx package");

        var result = await RunXlsxCliAsync("validate", path);

        Assert.Equal(1, result.ExitCode);
        var validation = JsonSerializer.Deserialize<XlsxValidationResult>(
            result.Stdout,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        Assert.NotNull(validation);
        Assert.False(validation.Valid);
        Assert.Contains(validation.Errors, error => error.Contains("valid XLSX package", StringComparison.Ordinal));
    }

    private static string CreateWorkbookWithManyValidationErrors()
    {
        var path = Path.Combine(Path.GetTempPath(), $"xlsx-many-validation-errors-{Guid.NewGuid():N}.xlsx");
        using var spreadsheet = SpreadsheetDocument.Create(path, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());

        for (uint i = 1; i <= 120; i++)
        {
            sheets.AppendChild(new Sheet { SheetId = i });
        }

        workbookPart.Workbook.Save();
        return path;
    }

    private static async Task<(int ExitCode, string Stdout, string Stderr)> RunXlsxCliAsync(params string[] args)
    {
        var executable = Path.Combine(AppContext.BaseDirectory, OperatingSystem.IsWindows() ? "xlsx.exe" : "xlsx");
        var startInfo = new ProcessStartInfo
        {
            FileName = executable,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };

        foreach (var arg in args)
        {
            startInfo.ArgumentList.Add(arg);
        }

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start xlsx CLI.");
        var stdout = await process.StandardOutput.ReadToEndAsync();
        var stderr = await process.StandardError.ReadToEndAsync();
        await process.WaitForExitAsync();
        return (process.ExitCode, stdout, stderr);
    }
}
