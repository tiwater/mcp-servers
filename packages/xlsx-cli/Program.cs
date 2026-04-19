using System.Text.Json;
using Dockit.Xlsx;

namespace Dockit.Xlsx.Cli;

internal static class Program
{
    public static Task<int> Main(string[] args) => Cli.RunAsync(args);
}

internal static class Cli
{
    public static Task<int> RunAsync(string[] args)
    {
        if (args.Length == 0)
        {
            PrintUsage();
            return Task.FromResult(1);
        }

        try
        {
            return args[0] switch
            {
                "inspect" => RunInspectAsync(args[1..]),
                "export-json" => Task.FromResult(Extractor.RunExportJson(args[1..])),
                "fill-template" => RunFillTemplateAsync(args[1..]),
                "edit" => Task.FromResult(Editor.RunEdit(args[1..])),
                _ => FailUnknown(args[0]),
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return Task.FromResult(1);
        }
    }

    private static Task<int> RunInspectAsync(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("inspect requires <input.xlsx>");
        }

        var input = args[0];
        var json = args.Skip(1).Contains("--json", StringComparer.Ordinal);
        var report = Inspector.Inspect(input);

        if (json)
        {
            WriteJson(report);
        }
        else
        {
            RenderInspect(report);
        }

        return Task.FromResult(0);
    }

    private static Task<int> RunFillTemplateAsync(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("fill-template requires <template.xlsx> <data.json> <output.xlsx>");
        }

        var template = args[0];
        var dataPath = args[1];
        var output = args[2];

        if (!File.Exists(dataPath))
        {
            throw new InvalidOperationException($"Data file not found: {dataPath}");
        }

        var jsonData = File.ReadAllText(dataPath);
        var fillData = JsonSerializer.Deserialize<FillData>(jsonData, Json.Options)
            ?? throw new InvalidOperationException("Failed to parse fill data");

        TemplateFiller.Fill(template, fillData, output);

        Console.WriteLine($"Filled template written to: {output}");
        return Task.FromResult(0);
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  inspect <input.xlsx> [--json]");
        Console.WriteLine("  export-json <input.xlsx> [<output.json>]");
        Console.WriteLine("  fill-template <template.xlsx> <data.json> <output.xlsx>");
        Console.WriteLine("  edit <input.xlsx> <operations.json> <output.xlsx>");
    }

    private static Task<int> FailUnknown(string command)
    {
        Console.Error.WriteLine($"Unknown command: {command}");
        PrintUsage();
        return Task.FromResult(1);
    }

    private static void WriteJson<T>(T value)
    {
        Console.WriteLine(JsonSerializer.Serialize(value, Json.Options));
    }

    private static void RenderInspect(WorkbookReport report)
    {
        Console.WriteLine($"File: {report.File}");
        Console.WriteLine($"Sheets: {report.SheetCount}");

        foreach (var sheet in report.Sheets)
        {
            Console.WriteLine($"  Sheet: {sheet.Name}");
            Console.WriteLine($"    Rows: {sheet.RowCount}");
            Console.WriteLine($"    Columns: {sheet.ColumnCount}");
            if (!string.IsNullOrWhiteSpace(sheet.UsedRange))
            {
                Console.WriteLine($"    Used Range: {sheet.UsedRange}");
            }
            Console.WriteLine($"    Formula Cells: {sheet.FormulaCellCount}");
        }
    }
}
