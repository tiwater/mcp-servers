namespace Dockit.Xlsx;

public static class Extractor
{
    public static int RunExportJson(string[] args)
    {
        var resolveMergedCells = args.Contains("--resolve-merged-cells", StringComparer.OrdinalIgnoreCase) || args.Contains("-r", StringComparer.OrdinalIgnoreCase);
        var cleanArgs = args.Where(arg => !string.Equals(arg, "--resolve-merged-cells", StringComparison.OrdinalIgnoreCase) && !string.Equals(arg, "-r", StringComparison.OrdinalIgnoreCase)).ToArray();

        if (cleanArgs.Length < 1)
        {
            throw new InvalidOperationException("export-json requires <input.xlsx> [<output.json>]");
        }

        var input = Path.GetFullPath(cleanArgs[0]);
        var output = cleanArgs.Length > 1 ? Path.GetFullPath(cleanArgs[1]) : null;

        var workbook = WorkbookLoader.Load(input, resolveMergedCells);
        var result = new List<object>();

        foreach (var sheet in workbook.Sheets)
        {
            result.Add(new
            {
                Sheet = sheet.Name,
                Rows = sheet.Rows,
                FormattedRows = sheet.FormattedRows,
                Cells = sheet.Cells
            });
        }

        var json = System.Text.Json.JsonSerializer.Serialize(result, Json.Options);
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
