using System.Text.Json;

namespace Dockit.Xlsx;

public static class Extractor
{
    public static int RunExportJson(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("export-json requires <input.xlsx> [<output.json>]");
        }

        var input = Path.GetFullPath(args[0]);
        var output = args.Length > 1 ? Path.GetFullPath(args[1]) : null;

        var workbook = WorkbookLoader.Load(input);
        var result = new List<object>();

        foreach (var sheet in workbook.Sheets)
        {
            result.Add(new
            {
                Sheet = sheet.Name,
                Rows = sheet.Rows
            });
        }

        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        if (output != null)
        {
            File.WriteAllText(output, json);
            Console.WriteLine(output);
        }
        else
        {
            Console.WriteLine(json);
        }

        return 0;
    }
}
