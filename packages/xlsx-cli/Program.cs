using System.Text.Json;
using Dockit.Convert;
using Dockit.Xlsx;
using Tiwater.FormatEvidence;

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
                "inspect-evidence" => Task.FromResult(FormatEvidenceCommand.RunProducer(args[1..], "tiwater-xlsx", XlsxToolVersion.Current, "xlsx", input => Inspector.InspectPublishedEvidence(input), WorkbookTargetObservations, WorkbookSourceFormats, ClassifyWorkbookEvidenceFailure)),
                "validate-inspect-evidence" => Task.FromResult(FormatEvidenceCommand.RunValidator(args[1..], "tiwater-xlsx", XlsxToolVersion.Current, "xlsx", input => Inspector.InspectPublishedEvidence(input), WorkbookTargetObservations, WorkbookSourceFormats, ClassifyWorkbookEvidenceFailure)),
                "inspect-evidence-v2" => Task.FromResult(FormatEvidenceCommand.RunProducerV2(args[1..], "tiwater-xlsx", XlsxToolVersion.Current, "xlsx", input => Inspector.InspectPublishedEvidence(input), WorkbookSourceFormats, ClassifyWorkbookEvidenceFailure, CandidateCapabilities)),
                "validate-inspect-evidence-v2" => Task.FromResult(FormatEvidenceCommand.RunValidatorV2(args[1..], "tiwater-xlsx", XlsxToolVersion.Current, "xlsx", input => Inspector.InspectPublishedEvidence(input), WorkbookSourceFormats, ClassifyWorkbookEvidenceFailure, CandidateCapabilities)),
                "derive-operation" => Task.FromResult(OperationDerivationCommand.RunProducer(args[1..], OperationContract())),
                "validate-derived-operation" => Task.FromResult(OperationDerivationCommand.RunValidator(args[1..], OperationContract())),
                "export-json" => Task.FromResult(Extractor.RunExportJson(args[1..])),
                "evidence" => RunEvidenceAsync(args[1..]),
                "fill-template" => RunFillTemplateAsync(args[1..]),
                "edit" => Task.FromResult(Editor.RunEdit(args[1..])),
                "validate" => RunValidateAsync(args[1..]),
                _ => FailUnknown(args[0]),
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.Message);
            return Task.FromResult(1);
        }
    }

    private static OperationDerivationCommand.Contract OperationContract() => new(
        "tiwater-xlsx",
        XlsxToolVersion.Current,
        "xlsx",
        "xlsx.edit",
        "1",
        "tiwater.xlsx-edit-v1.schema.json",
        "tiwater.xlsx-edit/v1",
        "current-artifact",
        "single-edit",
        true,
        value =>
        {
            var plan = JsonSerializer.Deserialize<XlsxEditDocument>(value.ToJsonString(), Json.Options)
                ?? throw new InvalidOperationException("XLSX derived operation could not be parsed");
            if (plan.Operations.Count != 1 || string.IsNullOrWhiteSpace(plan.Operations[0].Type))
                throw new InvalidOperationException("XLSX derived operation invalid");
        });

    private static IReadOnlyList<FormatEvidenceCommand.CandidateCapability> CandidateCapabilities(string pointer, IReadOnlySet<string> fields)
    {
        var kinds = new List<string>();
        if (fields.Contains("sheet") && fields.Contains("cell")) kinds.AddRange(["setCellValue", "setRichTextCellValue"]);
        if (fields.Contains("sheet") && fields.Contains("range")) kinds.Add("setPrintArea");
        if (fields.Contains("sheet") && fields.Contains("startCell")) kinds.Add("setRangeValues");
        if (fields.Contains("sheet") && fields.Contains("startRow")) kinds.Add("insertRows");
        if (fields.Contains("sheet") && fields.Contains("sourceRow") && fields.Contains("targetRow")) kinds.Add("copyRow");
        if (fields.Contains("sheet") && fields.Contains("anchorText")) kinds.Add("expandSectionRows");
        return kinds.Count == 0 ? [] : [new("xlsx.edit", "1", kinds)];
    }

    private static readonly IReadOnlySet<string> WorkbookSourceFormats = new HashSet<string>(StringComparer.Ordinal) { "xls", "xlsx" };

    internal static FormatEvidenceCommand.ErrorClassification? ClassifyWorkbookEvidenceFailure(Exception error)
        => error is AuthoritativeSpreadsheetRuntimeException
            ? new("inspect-evidence-runtime-unavailable", "runtime", true)
            : null;

    private static IReadOnlyList<FormatEvidenceCommand.AdditionalObservation> WorkbookTargetObservations(string path) =>
        WorkbookLoader.IsLegacyXls(path) ? [] :
        [
        new("workbook-target-1", "document.semantic-target", "structure", new
        {
            candidateId = "xlsx-workbook-root",
            semanticIdentity = new { format = "xlsx", scope = "workbook" },
            runtimeLocator = new { kind = "xlsx-workbook" },
            capabilities = new[] { "xlsx.edit" },
            resourceSet = new[] { new { resourceKey = "xlsx-workbook", access = "write" } },
            writeSet = new[] { new { resourceKey = "xlsx-workbook", writeKey = "workbook-cells" } }
        }, "/inspection/workbook")
        ];

    private static Task<int> RunEvidenceAsync(string[] args)
    {
        if (args.Length < 1) throw new InvalidOperationException("evidence requires <input.xlsx>");
        WriteJson(EvidenceInspector.Inspect(args[0]));
        return Task.FromResult(0);
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
        Console.WriteLine("  inspect-evidence --request <request.json> --output <evidence.json>");
        Console.WriteLine("  validate-inspect-evidence --request <request.json> --evidence <evidence.json> --output <verdict.json>");
        Console.WriteLine("  inspect-evidence-v2 --request <request.json> --output <evidence.json>");
        Console.WriteLine("  validate-inspect-evidence-v2 --request <request.json> --evidence <evidence.json> --output <verdict.json>");
        Console.WriteLine("  export-json <input.xlsx> [<output.json>]");
        Console.WriteLine("  evidence <input.xlsx>");
        Console.WriteLine("  fill-template <template.xlsx> <data.json> <output.xlsx>");
        Console.WriteLine("  edit <input.xlsx> <operations.json> <output.xlsx>");
        Console.WriteLine("  validate <input.xlsx>");
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
