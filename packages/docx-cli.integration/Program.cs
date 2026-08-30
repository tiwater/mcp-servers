using System.Diagnostics;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

if (args.Length != 1 || !File.Exists(args[0]))
    throw new InvalidOperationException("usage: docx-cli.integration <docx.dll>");

var cli = Path.GetFullPath(args[0]);
var root = Path.Combine(Path.GetTempPath(), "tiwater-docx-integration-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);

try
{
    var source = Path.Combine(root, "source.docx");
    var target = Path.Combine(root, "target.docx");
    CreateSource(source);
    CreateTarget(target);

    var sourceTable = FirstObject(Run("docx_list_objects", new
    {
        input = source, kinds = new[] { "table" }, scope = "/word/document.xml",
        limit = 10, output = Path.Combine(root, "source-tables.json")
    }));
    var sourceRows = Objects(Run("docx_list_objects", new
    {
        input = source, kinds = new[] { "row" }, scope = "/word/document.xml",
        parentRef = Ref(sourceTable), limit = 10, output = Path.Combine(root, "source-rows.json")
    }));
    var sourceRead = Run("docx_read_object", new
    {
        input = source, refs = sourceRows.Select(Ref).ToArray(), kinds = new[] { "cell", "paragraph" },
        output = Path.Combine(root, "source-read.json")
    });

    var targetState = ObserveTarget(target, "target");
    var invalidBoundary = RunExpectFailure("docx_replace_table_rows", Replacement(
        target, targetState, source, sourceTable, sourceRows, sourceRead,
        sourceFirst: 2, sourceLast: 4, targetFirst: 1, targetLast: targetState.Rows.Count - 1,
        Path.Combine(root, "invalid-boundary.docx"), Path.Combine(root, "invalid-boundary-receipt.json")));
    Require(invalidBoundary.Contains("source-row-range-starts-inside-vertical-merge", StringComparison.Ordinal),
        "source range beginning inside a vertical merge was not rejected");

    var first = Path.Combine(root, "first.docx");
    Run("docx_replace_table_rows", Replacement(
        target, targetState, source, sourceTable, sourceRows, sourceRead,
        sourceFirst: 1, sourceLast: 2, targetFirst: 1, targetLast: 2,
        first, Path.Combine(root, "first-receipt.json")));

    var stale = RunExpectFailure("docx_list_objects", new
    {
        input = first, kinds = new[] { "row" }, scope = "/word/document.xml",
        parentRef = Ref(targetState.Table), limit = 10, output = Path.Combine(root, "stale.json")
    });
    Require(stale.Contains("stale-parent-ref", StringComparison.Ordinal), "old parent ref was not rejected");

    var freshState = ObserveTarget(first, "first");
    Require(Ref(freshState.Table) != Ref(targetState.Table), "mutation did not produce fresh refs");

    var final = Path.Combine(root, "final.docx");
    Run("docx_replace_table_rows", Replacement(
        first, freshState, source, sourceTable, sourceRows, sourceRead,
        sourceFirst: 3, sourceLast: 4, targetFirst: 3, targetLast: freshState.Rows.Count - 1,
        final, Path.Combine(root, "final-receipt.json")));

    var finalState = ObserveTarget(final, "final");
    Require(finalState.Rows.Count == 5, $"expected 5 rows, found {finalState.Rows.Count}");
    ValidateFinal(final);
    RunInput("validate-openxml", final);

    Console.WriteLine("PASS docx observation -> mutation -> fresh observation -> mutation -> readback");
}
finally
{
    Directory.Delete(root, recursive: true);
}

object Replacement(
    string targetPath, TargetState targetState, string sourcePath, JsonElement sourceTable,
    IReadOnlyList<JsonElement> sourceRows, JsonElement sourceRead,
    int sourceFirst, int sourceLast, int targetFirst, int targetLast,
    string output, string receiptOutput)
{
    var sourceHeader = ChildObjects(Observation(sourceRead, 0));
    var selectedCells = Enumerable.Range(sourceFirst, sourceLast - sourceFirst + 1)
        .SelectMany(index => ChildObservations(Observation(sourceRead, index)))
        .Select(cell => new
        {
            sourceCellRef = Ref(cell.GetProperty("object")),
            sourceSelections = new[] { new { @ref = Ref(ChildObservations(cell)[0].GetProperty("object")) } }
        }).ToArray();

    return new
    {
        targetDocument = new { input = targetPath },
        tables = new[]
        {
            new
            {
                sourceDocument = new { input = sourcePath },
                sourceTableRef = Ref(sourceTable),
                sourceRows = new { firstRef = Ref(sourceRows[sourceFirst]), lastRef = Ref(sourceRows[sourceLast]) },
                targetTableRef = Ref(targetState.Table),
                targetRows = new { firstRef = Ref(targetState.Rows[targetFirst]), lastRef = Ref(targetState.Rows[targetLast]) },
                columns = sourceHeader.Zip(targetState.HeaderCells, (sourceCell, targetCell) => new
                {
                    sourceHeaderRef = Ref(sourceCell), targetHeaderRef = Ref(targetCell)
                }).ToArray(),
                sourceCellContents = selectedCells
            }
        },
        output,
        receiptOutput
    };
}

TargetState ObserveTarget(string input, string stem)
{
    var table = FirstObject(Run("docx_list_objects", new
    {
        input, kinds = new[] { "table" }, scope = "/word/document.xml",
        limit = 10, output = Path.Combine(root, stem + "-tables.json")
    }));
    var rows = Objects(Run("docx_list_objects", new
    {
        input, kinds = new[] { "row" }, scope = "/word/document.xml", parentRef = Ref(table),
        limit = 20, output = Path.Combine(root, stem + "-rows.json")
    }));
    var headerRead = Run("docx_read_object", new
    {
        input, refs = new[] { Ref(rows[0]) }, kinds = new[] { "cell", "paragraph" },
        output = Path.Combine(root, stem + "-header.json")
    });
    return new TargetState(table, rows, ChildObjects(Observation(headerRead, 0)));
}

JsonElement Run(string command, object request)
{
    var requestPath = Path.Combine(root, Guid.NewGuid().ToString("N") + ".json");
    File.WriteAllText(requestPath, JsonSerializer.Serialize(request));
    var result = Execute(command, requestPath);
    Require(result.ExitCode == 0, $"{command} failed: {result.Error}\n{result.Output}");
    return JsonDocument.Parse(result.Output).RootElement.Clone();
}

string RunExpectFailure(string command, object request)
{
    var requestPath = Path.Combine(root, Guid.NewGuid().ToString("N") + ".json");
    File.WriteAllText(requestPath, JsonSerializer.Serialize(request));
    var result = Execute(command, requestPath);
    Require(result.ExitCode != 0, $"{command} unexpectedly succeeded");
    return result.Error + result.Output;
}

void RunInput(string command, string input)
{
    var result = Execute(command, input);
    Require(result.ExitCode == 0, $"{command} failed: {result.Error}\n{result.Output}");
}

(int ExitCode, string Output, string Error) Execute(string command, string argument)
{
    var start = new ProcessStartInfo("dotnet")
    {
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        UseShellExecute = false
    };
    start.ArgumentList.Add(cli);
    start.ArgumentList.Add(command);
    start.ArgumentList.Add(argument);
    using var process = Process.Start(start) ?? throw new InvalidOperationException("failed to start docx cli");
    var output = process.StandardOutput.ReadToEnd();
    var error = process.StandardError.ReadToEnd();
    process.WaitForExit();
    return (process.ExitCode, output, error);
}

void CreateSource(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = BaseTable();
    table.Append(Header("字段", "内容"));
    table.Append(DataRow("甲", "甲内容", MergedCellValues.Restart));
    table.Append(DataRow("", "甲续行", MergedCellValues.Continue));
    table.Append(DataRow("乙", "乙内容", MergedCellValues.Restart));
    table.Append(DataRow("", "乙续行", MergedCellValues.Continue));
    main.Document = new Document(new Body(table));
    main.Document.Save();
}

void CreateTarget(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = BaseTable();
    table.Append(Header("字段", "内容"));
    for (var index = 0; index < 6; index++) table.Append(DataRow("", "", null, bilingual: false));
    main.Document = new Document(new Body(table));
    main.Document.Save();
}

Table BaseTable() => new(
    new TableProperties(new TableBorders(
        new TopBorder { Val = BorderValues.Single }, new LeftBorder { Val = BorderValues.Single },
        new BottomBorder { Val = BorderValues.Single }, new RightBorder { Val = BorderValues.Single },
        new InsideHorizontalBorder { Val = BorderValues.Single }, new InsideVerticalBorder { Val = BorderValues.Single })),
    new TableGrid(new GridColumn { Width = "3000" }, new GridColumn { Width = "6000" }));

TableRow Header(string left, string right) => new(Cell(left, null, false), Cell(right, null, false));

TableRow DataRow(string left, string right, MergedCellValues? merge, bool bilingual = true) =>
    new(Cell(left, merge, bilingual), Cell(right, null, bilingual));

TableCell Cell(string text, MergedCellValues? merge, bool bilingual)
{
    var properties = new TableCellProperties();
    if (merge is not null) properties.Append(new VerticalMerge { Val = merge.Value });
    var cell = new TableCell(properties, new Paragraph(new Run(new Text(text))));
    if (bilingual) cell.Append(new Paragraph(new Run(new Text("EN-" + (text.Length == 0 ? "continuation" : text)))));
    return cell;
}

void ValidateFinal(string path)
{
    using var document = WordprocessingDocument.Open(path, false);
    var rows = document.MainDocumentPart!.Document.Body!.Elements<Table>().Single().Elements<TableRow>().ToArray();
    var texts = rows.Skip(1).Select(row => row.InnerText).ToArray();
    Require(texts.SequenceEqual(new[] { "甲甲内容", "甲续行", "乙乙内容", "乙续行" }),
        "final text or Chinese-only selection is wrong: " + string.Join(" | ", texts));
    var merges = rows.Skip(1).Select(row => row.Elements<TableCell>().First()
        .TableCellProperties?.GetFirstChild<VerticalMerge>()?.Val?.Value).ToArray();
    Require(merges.SequenceEqual(new MergedCellValues?[]
        { MergedCellValues.Restart, MergedCellValues.Continue, MergedCellValues.Restart, MergedCellValues.Continue }),
        "vertical merge sequence is wrong");
}

static IReadOnlyList<JsonElement> Objects(JsonElement root) =>
    root.GetProperty("objects").EnumerateArray().Select(item => item.Clone()).ToArray();
static JsonElement FirstObject(JsonElement root) => Objects(root).Single();
static JsonElement Observation(JsonElement root, int index) => root.GetProperty("observations")[index];
static IReadOnlyList<JsonElement> ChildObservations(JsonElement observation) =>
    observation.GetProperty("children").EnumerateArray().Select(item => item.Clone()).ToArray();
static IReadOnlyList<JsonElement> ChildObjects(JsonElement observation) =>
    ChildObservations(observation).Select(item => item.GetProperty("object").Clone()).ToArray();
static string Ref(JsonElement value) => value.GetProperty("ref").GetString()!;
static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}

sealed record TargetState(JsonElement Table, IReadOnlyList<JsonElement> Rows, IReadOnlyList<JsonElement> HeaderCells);
