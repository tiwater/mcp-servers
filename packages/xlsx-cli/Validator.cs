using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Dockit.Xlsx;

public static class Validator
{
    private const int MaxValidationErrors = 100;

    public static XlsxValidationResult Validate(string input)
    {
        var file = Path.GetFullPath(input);
        var errors = new List<string>();
        var warnings = new List<string>();

        if (!File.Exists(file))
        {
            errors.Add($"File not found: {file}");
            return new XlsxValidationResult(file, false, errors, warnings);
        }

        try
        {
            using var spreadsheet = SpreadsheetDocument.Open(file, false);
            if (spreadsheet.WorkbookPart is null)
            {
                errors.Add("Workbook part missing.");
                return new XlsxValidationResult(file, false, errors, warnings);
            }

            var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            {
                MaxNumberOfErrors = MaxValidationErrors
            };

            foreach (var error in validator.Validate(spreadsheet))
            {
                var xpath = error.Path?.XPath;
                errors.Add(string.IsNullOrWhiteSpace(xpath)
                    ? error.Description
                    : $"{xpath}: {error.Description}");
            }

            if (errors.Count >= MaxValidationErrors)
            {
                warnings.Add($"Validation returned {MaxValidationErrors} errors; validation output may be truncated.");
            }

            return new XlsxValidationResult(file, errors.Count == 0, errors, warnings);
        }
        catch (Exception ex)
        {
            errors.Add($"File is not a valid XLSX package: {ex.Message}");
            return new XlsxValidationResult(file, false, errors, warnings);
        }
    }
}
