using Dockit.Xlsx;
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
}
