using System.Text.Json;
using Dockit.Convert;

namespace Dockit.Convert.Cli;

internal static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length < 3)
        {
            PrintUsage();
            return 1;
        }

        try
        {
            switch (args[0])
            {
                case "xls-to-xlsx":
                    WorkbookConverter.ConvertXlsToXlsx(args[1], args[2]);
                    Console.WriteLine(JsonSerializer.Serialize(new
                    {
                        status = "ok",
                        input = Path.GetFullPath(args[1]),
                        output = Path.GetFullPath(args[2]),
                        source_format = "xls",
                        target_format = "xlsx"
                    }));
                    return 0;
                default:
                    if (args[0].EndsWith("-to-pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        var sourceFormat = args[0][..^"-to-pdf".Length];
                        OfficePdfConverter.ConvertToPdf(args[1], args[2], sourceFormat);
                        Console.WriteLine(JsonSerializer.Serialize(new
                        {
                            status = "ok",
                            input = Path.GetFullPath(args[1]),
                            output = Path.GetFullPath(args[2]),
                            source_format = sourceFormat.ToLowerInvariant(),
                            target_format = "pdf"
                        }));
                        return 0;
                    }

                    Console.Error.WriteLine($"Unknown command: {args[0]}");
                    PrintUsage();
                    return 1;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  tiwater-convert xls-to-xlsx <input.xls> <output.xlsx>");
        Console.WriteLine("  tiwater-convert <docx|xlsx|pptx|doc|xls|ppt|odt|ods|odp|rtf>-to-pdf <input> <output.pdf>");
    }
}
