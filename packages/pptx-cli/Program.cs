using System.Text.Json;
using Dockit.Pptx;

namespace Dockit.Pptx.Cli;

internal static class Program
{
    public static Task<int> Main(string[] args) => Cli.RunAsync(args);
}

internal static class Cli
{
    private static readonly string[] DiscoverableCommands =
    [
        "inspect",
        "export-json",
        "validate",
        "map-render-findings",
        "validate-render-finding-map",
        .. FixedCommandRunner.Commands,
    ];

    public static Task<int> RunAsync(string[] args)
    {
        if (args.Length == 1 && args[0] == "--list-tools")
        {
            WriteJson(new { schema = "tiwater.provider-tool-list/v1", commands = DiscoverableCommands });
            return Task.FromResult(0);
        }

        if (args.Length == 1 && args[0] is "--help" or "-h")
        {
            PrintUsage();
            return Task.FromResult(0);
        }

        if (args.Length == 2 && args[1] is "--help" or "-h" && PrintCommandUsage(args[0]))
            return Task.FromResult(0);

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
                "validate" => Task.FromResult(Validator.Run(args[1..])),
                "map-render-findings" => RunMapRenderFindingsAsync(args[1..]),
                "validate-render-finding-map" => RunValidateRenderFindingMapAsync(args[1..]),
                _ when FixedCommandRunner.IsCommand(args[0]) => Task.FromResult(FixedCommandRunner.Run(args[0], args[1..])),
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
            WriteJson(Inspector.InspectDetail(input));
        }
        else
        {
            Console.WriteLine($"File: {report.File}");
            Console.WriteLine($"Slides: {report.SlideCount}");
            Console.WriteLine($"Placeholders: {string.Join(", ", report.Placeholders)}");
        }

        return Task.FromResult(0);
    }

    private static Task<int> RunMapRenderFindingsAsync(string[] args)
    {
        if (args.Length != 4) throw new InvalidOperationException("map-render-findings requires <inspect.json> <render-manifest.json> <findings.json> <output.json>");
        var result = RenderedFindingMapper.MapFiles(args[0], args[1], args[2]);
        WriteNewJson(args[3], result); WriteJson(new { status = "ok", output = Path.GetFullPath(args[3]), findingCount = result.Findings.Count });
        return Task.FromResult(0);
    }

    private static Task<int> RunValidateRenderFindingMapAsync(string[] args)
    {
        if (args.Length != 5) throw new InvalidOperationException("validate-render-finding-map requires <inspect.json> <render-manifest.json> <findings.json> <map.json> <verdict.json>");
        var result = RenderedFindingValidator.ValidateFiles(args[0], args[1], args[2], args[3]);
        WriteNewJson(args[4], result); WriteJson(result);
        return Task.FromResult(result.Pass ? 0 : 1);
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  inspect <input.pptx> [--json]");
        Console.WriteLine("  export-json <input.pptx> [<output.json>]");
        Console.WriteLine("  validate <input.pptx>");
        Console.WriteLine("  map-render-findings <inspect.json> <render-manifest.json> <findings.json> <output.json>");
        Console.WriteLine("  validate-render-finding-map <inspect.json> <render-manifest.json> <findings.json> <map.json> <verdict.json>");
        foreach (var command in FixedCommandRunner.Commands)
            Console.WriteLine($"  {command} <request.json>");
    }

    private static bool PrintCommandUsage(string command)
    {
        if (FixedCommandRunner.IsCommand(command))
        {
            Console.WriteLine($"tiwater-pptx {command} <request.json>");
            return true;
        }

        return false;
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

    private static void WriteNewJson<T>(string path, T value)
    {
        var fullPath = Path.GetFullPath(path); Directory.CreateDirectory(Path.GetDirectoryName(fullPath) ?? ".");
        using var stream = new FileStream(fullPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        JsonSerializer.Serialize(stream, value, Json.Options);
        stream.WriteByte((byte)'\n');
    }
}
