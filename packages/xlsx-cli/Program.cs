using System.Text.Json;
using Dockit.Convert;
using Dockit.Xlsx;

namespace Dockit.Xlsx.Cli;

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
        "inventory-regions",
        "xlsx_read_range",
        "validate",
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

        if (args.Length == 0)
        {
            PrintUsage();
            return Task.FromResult(1);
        }

        if (args.Length == 2 && args[1] is "--help" or "-h" && PrintCommandUsage(args[0]))
        {
            return Task.FromResult(0);
        }

        try
        {
            return args[0] switch
            {
                "inspect" => RunInspectAsync(args[1..]),
                "export-json" => Task.FromResult(Extractor.RunExportJson(args[1..])),
                "inventory-regions" => Task.FromResult(RegionInventory.Run(args[1..])),
                "xlsx_read_range" => Task.FromResult(RangeReader.Run(args[1..])),
                "validate" => RunValidateAsync(args[1..]),
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
            throw new InvalidOperationException("inspect requires <input.xlsx>");
        }

        var input = args[0];
        var json = args.Skip(1).Contains("--json", StringComparer.Ordinal);
        if (json)
        {
            WriteJson(Inspector.InspectPublishedEvidence(input));
            return Task.FromResult(0);
        }
        var report = Inspector.Inspect(input);

        RenderInspect(report);

        return Task.FromResult(0);
    }

    private static Task<int> RunValidateAsync(string[] args)
    {
        if (args.Length < 1)
        {
            throw new InvalidOperationException("validate requires <input.xlsx>");
        }

        var result = Validator.Validate(args[0]);
        WriteJson(result);
        return Task.FromResult(result.Valid ? 0 : 1);
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  inspect <input.xlsx> [--json]");
        Console.WriteLine("  export-json <input.xlsx> [<output.json>]");
        Console.WriteLine("  inventory-regions <input.xlsx> [<output.json>] [--schema v1|v2]");
        Console.WriteLine("  xlsx_read_range <request.json>");
        Console.WriteLine("  validate <input.xlsx>");
        foreach (var command in FixedCommandRunner.Commands)
            Console.WriteLine($"  {command} <request.json>");
    }

    private static bool PrintCommandUsage(string command)
    {
        if (FixedCommandRunner.IsCommand(command))
        {
            Console.WriteLine($"tiwater-xlsx {command} <request.json>");
            return true;
        }

        return command switch
        {
            "inspect" => PrintUsageLine("tiwater-xlsx inspect <input.xlsx> [--json]"),
            "export-json" => PrintUsageLine("tiwater-xlsx export-json <input.xlsx> [<output.json>]"),
            "inventory-regions" => PrintUsageLine("tiwater-xlsx inventory-regions <input.xlsx> [<output.json>] [--schema v1|v2]"),
            "xlsx_read_range" => PrintUsageLine("tiwater-xlsx xlsx_read_range <request.json>"),
            "validate" => PrintUsageLine("tiwater-xlsx validate <input.xlsx>"),
            _ => false,
        };
    }

    private static bool PrintUsageLine(string usage)
    {
        Console.WriteLine(usage);
        return true;
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
