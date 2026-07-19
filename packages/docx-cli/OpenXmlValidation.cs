using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Dockit.Docx;

public static class OpenXmlValidation
{
    public static int Run(string[] args)
    {
        if (args.Length < 1) throw new InvalidOperationException("validate-openxml requires <input.docx>");
        var path = Path.GetFullPath(args[0]);
        using var document = WordprocessingDocument.Open(path, false);
        var errors = new OpenXmlValidator().Validate(document)
            .Take(100)
            .Select(error => new { error.Description, Path = error.Path?.XPath, error.Id })
            .ToList();
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            Schema = "tiwater.docx.openxml-validation/v1",
            Pass = errors.Count == 0,
            File = path,
            ErrorCount = errors.Count,
            Errors = errors,
        }, Json.Options));
        return errors.Count == 0 ? 0 : 1;
    }
}
