using System.Text.Json;
using Dockit.Pptx;

namespace Dockit.Pptx.Cli;

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
            throw new InvalidOperationException("inspect requires <input.pptx>");
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
            Console.WriteLine($"File: {report.File}");
            Console.WriteLine($"Slides: {report.SlideCount}");
            Console.WriteLine($"Placeholders: {string.Join(", ", report.Placeholders)}");
        }

        return Task.FromResult(0);
    }

    private static Task<int> RunFillTemplateAsync(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("fill-template requires <template.pptx> <data.json> <output.pptx>");
        }

        var template = args[0];
        var dataPath = args[1];
        var output = args[2];

        var result = TemplateFiller.Fill(template, dataPath, output);
        WriteJson(result);
        return Task.FromResult(0);
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  inspect <input.pptx> [--json]");
        Console.WriteLine("  export-json <input.pptx> [<output.json>]");
        Console.WriteLine("  fill-template <template.pptx> <data.json> <output.pptx>");
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
}
