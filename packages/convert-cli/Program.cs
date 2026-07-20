using System.Text.Json;
using System.Reflection;
using System.Security.Cryptography;
using Dockit.Convert;

namespace Dockit.Convert.Cli;

internal static class Program
{
    private static readonly string ToolVersion =
        (typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "unknown")
        .Split('+', 2)[0];

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
                    var result = WorkbookConverter.ConvertXlsToXlsx(args[1], args[2]);
                    Console.WriteLine(JsonSerializer.Serialize(new
                    {
                        status = "ok",
                        input = Path.GetFullPath(args[1]),
                        input_sha256 = FileSha256(args[1]),
                        output = Path.GetFullPath(args[2]),
                        output_sha256 = FileSha256(args[2]),
                        source_format = "xls",
                        target_format = "xlsx",
                        version = ToolVersion,
                        backend = result.Backend,
                        fallback_reason = result.FallbackReason,
                    }));
                    return 0;
                default:
                    if (args[0].EndsWith("-to-pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        var sourceFormat = args[0][..^"-to-pdf".Length];
                        var pdfResult = OfficePdfConverter.ConvertToPdf(args[1], args[2], sourceFormat);
                        Console.WriteLine(JsonSerializer.Serialize(new
                        {
                            status = "ok",
                            input = Path.GetFullPath(args[1]),
                            input_sha256 = FileSha256(args[1]),
                            output = Path.GetFullPath(args[2]),
                            output_sha256 = FileSha256(args[2]),
                            source_format = sourceFormat.ToLowerInvariant(),
                            target_format = "pdf",
                            version = ToolVersion,
                            backend = pdfResult.Backend,
                            fallback_reason = pdfResult.FallbackReason,
                            page_count = pdfResult.NativeRenderProvenance?.PageCount,
                            native_render_provenance = pdfResult.NativeRenderProvenance,
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

    private static string FileSha256(string path)
    {
        using var stream = File.OpenRead(path);
        return System.Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }
}
