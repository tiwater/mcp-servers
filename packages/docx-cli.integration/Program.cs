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

    var sourceTableIndex = Run("docx_table_index", new
    {
        input = source, output = Path.Combine(root, "source-table-index.json")
    });
    var targetTableIndex = Run("docx_table_index", new
    {
        input = target, output = Path.Combine(root, "target-table-index.json")
    });
    Require(sourceTableIndex.GetProperty("tables").GetArrayLength() == 1,
        "source table index did not return its single table");
    Require(targetTableIndex.GetProperty("tables")[0].GetProperty("rowCount").GetInt32() == 7,
        "target table index row count is wrong");
    Require(targetTableIndex.GetProperty("tables")[0].GetProperty("columnCount").GetInt32() == 2,
        "target table index column count is wrong");

    var compactTable = Run("docx_read_table", new
    {
        input = source,
        table = sourceTableIndex.GetProperty("tables")[0].GetProperty("address"),
        output = Path.Combine(root, "source-table.json")
    });
    Require(compactTable.GetProperty("rowCount").GetInt32() == 5,
        "compact table row count is wrong");
    Require(compactTable.GetProperty("columnCount").GetInt32() == 2,
        "compact table column count is wrong");
    var compactRows = compactTable.GetProperty("rows");
    Require(compactRows[1].GetProperty("cells")[0].GetProperty("verticalMerge").GetString() == "restart",
        "compact table did not expose vertical merge restart");
    Require(compactRows[2].GetProperty("cells")[0].GetProperty("verticalMerge").GetString() == "continue",
        "compact table did not expose vertical merge continuation");
    var compactParagraphs = compactRows[1].GetProperty("cells")[0].GetProperty("paragraphs");
    Require(compactParagraphs.GetArrayLength() == 2 && compactParagraphs[0].GetProperty("text").GetString() == "甲",
        "compact table paragraph projection is wrong");
    Require(compactParagraphs[0].GetProperty("textNodes")[0].GetProperty("text").GetString() == "甲",
        "compact table text-node projection is wrong");
    Require(compactParagraphs[0].GetProperty("address").GetProperty("path").GetString()!.StartsWith('/'),
        "compact table paragraph address is not native");

    var nonTableRead = RunExpectFailure("docx_read_table", new
    {
        input = source,
        table = compactRows[0].GetProperty("address"),
        output = Path.Combine(root, "non-table.json")
    });
    Require(nonTableRead.Contains("table-reference-kind-invalid", StringComparison.Ordinal),
        "compact table accepted a non-table address");

    var sourceTable = FirstObject(Run("docx_list_objects", new
    {
        input = source, kinds = new[] { "table" }, scope = "/word/document.xml",
        limit = 10, output = Path.Combine(root, "source-tables.json")
    }));
    Require(Address(sourceTable).GetRawText()
            == sourceTableIndex.GetProperty("tables")[0].GetProperty("address").GetRawText(),
        "table index address differs from native object listing");
    var sourceRows = Objects(Run("docx_list_objects", new
    {
        input = source, kinds = new[] { "row" }, scope = "/word/document.xml",
        parent = Address(sourceTable), limit = 10, output = Path.Combine(root, "source-rows.json")
    }));
    var sourceRead = Run("docx_read_object", new
    {
        input = source, addresses = sourceRows.Select(Address).ToArray(), kinds = new[] { "cell", "paragraph" },
        output = Path.Combine(root, "source-read.json")
    });

    var targetState = ObserveTarget(target, "target");

    var modeOutput = Path.Combine(root, "paragraph-mode.docx");
    Run("docx_replace_table_rows", ParagraphModeReplacement(
        target, targetState, source, sourceTable, sourceRows, sourceRead,
        sourceFirst: 1, sourceLast: 2, targetFirst: 1,
        modeOutput, Path.Combine(root, "paragraph-mode-receipt.json")));
    using (var modeDocument = WordprocessingDocument.Open(modeOutput, false))
    {
        var modeRows = modeDocument.MainDocumentPart!.Document.Body!.Elements<Table>()
            .Single().Elements<TableRow>().ToArray();
        Require(modeRows.Length == 3, "omitted target last did not replace through table end");
        Require(modeRows[1].InnerText == "甲甲内容R2=1.000采用 HPLC 检测",
            "paired Latin prose was not removed or technical content was lost");
        Require(modeRows[2].InnerText == "English continuation甲续行",
            "all-Latin content without a Han pair was not preserved");
    }

    var incompatible = RunExpectFailure("docx_replace_table_rows", ParagraphModeReplacement(
        target, targetState, source, sourceTable, sourceRows, sourceRead,
        sourceFirst: 1, sourceLast: 2, targetFirst: 1,
        Path.Combine(root, "incompatible.docx"), Path.Combine(root, "incompatible-receipt.json"),
        includeExplicitSelections: true));
    Require(incompatible.Contains("source-cell-contents-and-source-paragraph-mode-are-mutually-exclusive",
            StringComparison.Ordinal),
        "paragraph mode and explicit source selections were not rejected together");

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

    var rowsFromStableTableAddress = Objects(Run("docx_list_objects", new
    {
        input = first, kinds = new[] { "row" }, scope = "/word/document.xml",
        parent = Address(targetState.Table), limit = 10, output = Path.Combine(root, "stable-table-address.json")
    }));
    Require(rowsFromStableTableAddress.Count > 0, "unchanged table address did not remain usable after row content replacement");

    var freshState = ObserveTarget(first, "first");
    Require(Address(freshState.Table).GetRawText() == Address(targetState.Table).GetRawText(), "unchanged table address changed after row content replacement");

    var beforeRejectedInPlace = File.ReadAllBytes(first);
    var rejectedInPlace = RunExpectFailure("docx_replace_table_rows", Replacement(
        first, freshState, source, sourceTable, sourceRows, sourceRead,
        sourceFirst: 2, sourceLast: 4, targetFirst: 1, targetLast: freshState.Rows.Count - 1,
        first, Path.Combine(root, "rejected-in-place-receipt.json")));
    Require(rejectedInPlace.Contains("source-row-range-starts-inside-vertical-merge", StringComparison.Ordinal),
        "invalid in-place mutation did not report its technical failure");
    Require(File.ReadAllBytes(first).SequenceEqual(beforeRejectedInPlace),
        "invalid in-place mutation changed the input document");

    var final = first;
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
            sourceCell = Address(cell.GetProperty("object")),
            sourceSelections = new[] { new { address = Address(ChildObservations(cell)[0].GetProperty("object")) } }
        }).ToArray();

    return new
    {
        input = targetPath,
        tables = new[]
        {
            new
            {
                sourceInput = sourcePath,
                sourceTable = Address(sourceTable),
                sourceRows = new { first = Address(sourceRows[sourceFirst]), last = Address(sourceRows[sourceLast]) },
                targetTable = Address(targetState.Table),
                targetRows = new { first = Address(targetState.Rows[targetFirst]), last = Address(targetState.Rows[targetLast]) },
                columns = sourceHeader.Zip(targetState.HeaderCells, (sourceCell, targetCell) => new
                {
                    sourceHeader = Address(sourceCell), targetHeader = Address(targetCell)
                }).ToArray(),
                sourceCellContents = selectedCells
            }
        },
        output,
        receiptOutput
    };
}

object ParagraphModeReplacement(
    string targetPath, TargetState targetState, string sourcePath, JsonElement sourceTable,
    IReadOnlyList<JsonElement> sourceRows, JsonElement sourceRead,
    int sourceFirst, int sourceLast, int targetFirst,
    string output, string receiptOutput, bool includeExplicitSelections = false)
{
    var sourceHeader = ChildObjects(Observation(sourceRead, 0));
    var selectedCells = Enumerable.Range(sourceFirst, sourceLast - sourceFirst + 1)
        .SelectMany(index => ChildObservations(Observation(sourceRead, index)))
        .Select(cell => new
        {
            sourceCell = Address(cell.GetProperty("object")),
            sourceSelections = new[] { new { address = Address(ChildObservations(cell)[0].GetProperty("object")) } }
        }).ToArray();
    return new
    {
        input = targetPath,
        tables = new[]
        {
            new
            {
                sourceInput = sourcePath,
                sourceTable = Address(sourceTable),
                sourceRows = new { first = Address(sourceRows[sourceFirst]), last = Address(sourceRows[sourceLast]) },
                targetTable = Address(targetState.Table),
                targetRows = new { first = Address(targetState.Rows[targetFirst]) },
                columns = sourceHeader.Zip(targetState.HeaderCells, (sourceCell, targetCell) => new
                {
                    sourceHeader = Address(sourceCell), targetHeader = Address(targetCell)
                }).ToArray(),
                sourceCellContents = includeExplicitSelections ? selectedCells : null,
                sourceParagraphMode = "omit-paired-latin-prose"
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
        input, kinds = new[] { "row" }, scope = "/word/document.xml", parent = Address(table),
        limit = 20, output = Path.Combine(root, stem + "-rows.json")
    }));
    var headerRead = Run("docx_read_object", new
    {
        input, addresses = new[] { Address(rows[0]) }, kinds = new[] { "cell", "paragraph" },
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
    if (bilingual)
    {
        cell.Append(new Paragraph(new Run(new Text(text.Length == 0 ? "English continuation" : "English translation"))));
        if (text == "甲内容")
        {
            cell.Append(new Paragraph(new Run(new Text("R2=1.000"))));
            cell.Append(new Paragraph(new Run(new Text("采用 HPLC 检测"))));
        }
    }
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
static JsonElement Address(JsonElement value) => value.GetProperty("address").Clone();
static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}

sealed record TargetState(JsonElement Table, IReadOnlyList<JsonElement> Rows, IReadOnlyList<JsonElement> HeaderCells);
