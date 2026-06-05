using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Spreadsheet;

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

            ValidateSharedFormulas(spreadsheet, errors);

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

    private static void ValidateSharedFormulas(SpreadsheetDocument spreadsheet, List<string> errors)
    {
        foreach (var worksheetPart in spreadsheet.WorkbookPart!.WorksheetParts)
        {
            var sharedFormulaCells = worksheetPart.Worksheet
                .Descendants<Cell>()
                .Where(cell => cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Shared)
                .ToList();
            var masters = new Dictionary<uint, Cell>();

            foreach (var cell in sharedFormulaCells)
            {
                var formula = cell.CellFormula!;
                var sharedIndex = formula.SharedIndex?.Value;
                if (sharedIndex is null)
                {
                    errors.Add($"Shared formula at {CellReference(cell)} is missing shared index.");
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(formula.Text) || formula.Reference is not null)
                {
                    if (masters.TryGetValue(sharedIndex.Value, out var existing))
                    {
                        errors.Add($"Shared formula index {sharedIndex.Value} has duplicate masters at {CellReference(existing)} and {CellReference(cell)}.");
                    }
                    masters[sharedIndex.Value] = cell;

                    var reference = formula.Reference?.Value;
                    if (string.IsNullOrWhiteSpace(reference))
                    {
                        errors.Add($"Shared formula master at {CellReference(cell)} is missing ref range.");
                    }
                    else if (!RangeContains(reference, CellReference(cell)))
                    {
                        errors.Add($"Shared formula master at {CellReference(cell)} is outside ref range {reference}.");
                    }
                }
            }

            foreach (var cell in sharedFormulaCells)
            {
                var sharedIndex = cell.CellFormula!.SharedIndex?.Value;
                if (sharedIndex is not null && !masters.ContainsKey(sharedIndex.Value))
                {
                    errors.Add($"Shared formula follower at {CellReference(cell)} references missing master si={sharedIndex.Value}.");
                }
            }
        }
    }

    private static string CellReference(Cell cell) => cell.CellReference?.Value ?? "<unknown>";

    private static bool RangeContains(string rangeReference, string cellReference)
    {
        var parts = rangeReference.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var start = ParseCellReference(parts[0]);
        var end = ParseCellReference(parts.Length > 1 ? parts[1] : parts[0]);
        var cell = ParseCellReference(cellReference);

        return cell.Column >= Math.Min(start.Column, end.Column)
            && cell.Column <= Math.Max(start.Column, end.Column)
            && cell.Row >= Math.Min(start.Row, end.Row)
            && cell.Row <= Math.Max(start.Row, end.Row);
    }

    private static (int Column, int Row) ParseCellReference(string reference)
    {
        var column = 0;
        var row = 0;
        foreach (var ch in reference)
        {
            if (char.IsLetter(ch))
            {
                column = column * 26 + (char.ToUpperInvariant(ch) - 'A' + 1);
            }
            else if (char.IsDigit(ch))
            {
                row = row * 10 + (ch - '0');
            }
        }

        return (column, row);
    }
}
