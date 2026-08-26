using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Dockit.Pptx;

public static class Validator
{
    private const int MaxValidationErrors = 100;

    public static int Run(string[] args)
    {
        if (args.Length < 1) throw new InvalidOperationException("validate requires <input.pptx>");
        var result = Validate(args[0]);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.Pass ? 0 : 1;
    }

    public static PptxValidationResult Validate(string input)
    {
        var file = Path.GetFullPath(input);
        var errors = new List<PptxValidationIssue>();
        try
        {
            using var presentation = PresentationDocument.Open(file, false);
            if (presentation.PresentationPart is null)
            {
                errors.Add(new("Presentation part missing.", null, "presentation-part-missing"));
            }
            else
            {
                var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365)
                {
                    MaxNumberOfErrors = MaxValidationErrors,
                };
                errors.AddRange(validator.Validate(presentation).Select(error =>
                    new PptxValidationIssue(error.Description, error.Path?.XPath, error.Id)));
            }
        }
        catch (Exception error)
        {
            errors.Add(new($"File is not a valid PPTX package: {error.Message}", null, "invalid-pptx-package"));
        }

        return new PptxValidationResult(
            "tiwater.pptx.openxml-validation/v1",
            errors.Count == 0,
            file,
            errors.Count,
            errors.Take(MaxValidationErrors).ToList());
    }
}
