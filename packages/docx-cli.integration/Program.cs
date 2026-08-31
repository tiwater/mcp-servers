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
    var original = Path.Combine(root, "original.docx");
    CreateDocument(original);
    var initial = ReadTable(original, "initial");
    Require(initial.GetProperty("schema").GetString() == "tiwater.docx-table-read/v2", "table read schema is not v2");
    Require(initial.GetProperty("gridColumns").GetArrayLength() == 4, "grid columns were not exposed");
    var rows = initial.GetProperty("rows");
    Require(rows[0].GetProperty("repeatHeader").GetBoolean(), "repeat-header row property was not exposed");
    Require(rows[1].GetProperty("cantSplit").GetBoolean(), "no-split row property was not exposed");
    Require(rows[1].GetProperty("cells")[2].GetProperty("gridColumnStart").GetInt32() == 2,
        "cell grid start was not exposed");
    Require(rows[2].GetProperty("cells")[0].GetProperty("verticalMergeOwner").GetRawText()
            == rows[1].GetProperty("cells")[0].GetProperty("address").GetRawText(),
        "vertical continuation did not point to its restart owner");

    var batchColumns = Path.Combine(root, "batch-columns.docx");
    Run("docx_insert_table_columns", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                table = initial.GetProperty("address").Clone(),
                sourceColumn = initial.GetProperty("gridColumns")[0].GetProperty("address").Clone(),
                before = initial.GetProperty("gridColumns")[1].GetProperty("address").Clone()
            },
            new
            {
                table = initial.GetProperty("address").Clone(),
                sourceColumn = initial.GetProperty("gridColumns")[2].GetProperty("address").Clone(),
                before = initial.GetProperty("gridColumns")[3].GetProperty("address").Clone()
            }
        },
        output = batchColumns,
        receiptOutput = Path.Combine(root, "batch-columns-receipt.json")
    });
    var batchState = ReadTable(batchColumns, "batch-columns");
    var batchWidths = batchState.GetProperty("gridColumns").EnumerateArray()
        .Select(column => column.GetProperty("widthTwips").GetInt32()).ToArray();
    Require(batchWidths.SequenceEqual(new[] { 1200, 1200, 1200, 1800, 1800, 1800 }),
        "batch column insertion reused stale grid paths");

    var insertedInsideOtherSpan = Path.Combine(root, "inside-other-span.docx");
    Run("docx_insert_table_columns", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                table = initial.GetProperty("address").Clone(),
                sourceColumn = initial.GetProperty("gridColumns")[0].GetProperty("address").Clone(),
                before = initial.GetProperty("gridColumns")[3].GetProperty("address").Clone()
            }
        },
        output = insertedInsideOtherSpan,
        receiptOutput = Path.Combine(root, "inside-other-span-receipt.json")
    });
    var otherSpanState = ReadTable(insertedInsideOtherSpan, "inside-other-span");
    var otherSpanHeader = otherSpanState.GetProperty("rows")[0].GetProperty("cells");
    Require(otherSpanHeader[0].GetProperty("gridSpan").GetInt32() == 2
            && otherSpanHeader[1].GetProperty("gridSpan").GetInt32() == 3,
        "column insertion expanded the source header instead of the insertion-position header");

    var insertedColumns = Path.Combine(root, "inserted-columns.docx");
    Run("docx_insert_table_columns", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                table = initial.GetProperty("address").Clone(),
                sourceColumn = initial.GetProperty("gridColumns")[1].GetProperty("address").Clone(),
                before = initial.GetProperty("gridColumns")[2].GetProperty("address").Clone()
            }
        },
        output = insertedColumns,
        receiptOutput = Path.Combine(root, "inserted-columns-receipt.json")
    });
    var afterInsert = ReadTable(insertedColumns, "after-insert");
    Require(afterInsert.GetProperty("columnCount").GetInt32() == 5, "column insert did not expand the grid");
    Require(afterInsert.GetProperty("rows")[0].GetProperty("cells")[0].GetProperty("gridSpan").GetInt32() == 3,
        "column insert did not expand the spanning header");
    Require(afterInsert.GetProperty("rows")[1].GetProperty("cells").GetArrayLength() == 5,
        "column insert did not add a data cell");
    Require(string.Concat(afterInsert.GetProperty("rows")[1].GetProperty("cells")[2].GetProperty("paragraphs")
            .EnumerateArray().Select(item => item.GetProperty("text").GetString())).Length == 0,
        "inserted data cell copied business text");

    var deletedColumns = Path.Combine(root, "deleted-columns.docx");
    Run("docx_delete_table_columns", new
    {
        input = insertedColumns,
        changes = new[]
        {
            new
            {
                table = afterInsert.GetProperty("address").Clone(),
                columns = new[] { afterInsert.GetProperty("gridColumns")[2].GetProperty("address").Clone() }
            }
        },
        output = deletedColumns,
        receiptOutput = Path.Combine(root, "deleted-columns-receipt.json")
    });
    var afterDelete = ReadTable(deletedColumns, "after-delete");
    Require(afterDelete.GetProperty("columnCount").GetInt32() == 4, "column delete did not restore the grid");
    Require(afterDelete.GetProperty("rows")[0].GetProperty("cells")[0].GetProperty("gridSpan").GetInt32() == 2,
        "column delete did not shrink the spanning header");

    var insertedInsideMerge = Path.Combine(root, "inserted-inside-merge.docx");
    Run("docx_insert_objects", new
    {
        input = deletedColumns,
        changes = new[]
        {
            new
            {
                sourceInput = deletedColumns,
                sources = new[] { afterDelete.GetProperty("rows")[2].GetProperty("address").Clone() },
                targetParent = afterDelete.GetProperty("address").Clone(),
                before = afterDelete.GetProperty("rows")[2].GetProperty("address").Clone(),
                repeat = 2
            }
        },
        output = insertedInsideMerge,
        receiptOutput = Path.Combine(root, "inserted-inside-merge-receipt.json")
    });
    var insideMergeState = ReadTable(insertedInsideMerge, "inserted-inside-merge");
    Require(insideMergeState.GetProperty("rowCount").GetInt32() == 6,
        "row insertion inside a vertical merge did not add two rows");
    var mergeOwner = insideMergeState.GetProperty("rows")[1].GetProperty("cells")[0]
        .GetProperty("address").GetRawText();
    foreach (var rowIndex in new[] { 2, 3, 4 })
    {
        var cell = insideMergeState.GetProperty("rows")[rowIndex].GetProperty("cells")[0];
        Require(cell.GetProperty("verticalMerge").GetString() == "continue"
                && cell.GetProperty("verticalMergeOwner").GetRawText() == mergeOwner,
            "inserted row did not extend the existing vertical merge");
    }

    var incompatibleInsert = RunExpectFailure("docx_insert_objects", new
    {
        input = deletedColumns,
        changes = new[]
        {
            new
            {
                sourceInput = deletedColumns,
                sources = new[] { afterDelete.GetProperty("rows")[3].GetProperty("address").Clone() },
                targetParent = afterDelete.GetProperty("address").Clone(),
                before = afterDelete.GetProperty("rows")[2].GetProperty("address").Clone()
            }
        },
        output = Path.Combine(root, "incompatible-insert.docx"),
        receiptOutput = Path.Combine(root, "incompatible-insert-receipt.json")
    });
    Require(incompatibleInsert.Contains("row-insert-boundary-requires-compatible-vertical-merge", StringComparison.Ordinal),
        "row insertion accepted an incompatible vertical-merge boundary");

    var invalidDelete = RunExpectFailure("docx_delete_object", new
    {
        input = deletedColumns,
        changes = new[]
        {
            new { addresses = new[] { afterDelete.GetProperty("rows")[1].GetProperty("address").Clone() } }
        },
        output = Path.Combine(root, "invalid-delete.docx"),
        receiptOutput = Path.Combine(root, "invalid-delete-receipt.json")
    });
    Require(invalidDelete.Contains("row-selection-splits-vertical-merge", StringComparison.Ordinal),
        "row deletion accepted half of a vertical merge");

    var insertedRows = Path.Combine(root, "inserted-rows.docx");
    Run("docx_insert_objects", new
    {
        input = deletedColumns,
        changes = new[]
        {
            new
            {
                sourceInput = deletedColumns,
                sources = new[] { afterDelete.GetProperty("rows")[3].GetProperty("address").Clone() },
                targetParent = afterDelete.GetProperty("address").Clone(),
                repeat = 2
            }
        },
        output = insertedRows,
        receiptOutput = Path.Combine(root, "inserted-rows-receipt.json")
    });
    var afterRows = ReadTable(insertedRows, "after-rows");
    Require(afterRows.GetProperty("rowCount").GetInt32() == 6, "row insertion did not add two rows");

    var mergeInput = Path.Combine(root, "merge-input.docx");
    CreateFlatDocument(mergeInput);
    var mergeState = ReadTable(mergeInput, "merge-state");
    var mergeCells = mergeState.GetProperty("rows")[0].GetProperty("cells");
    var merged = Path.Combine(root, "merged.docx");
    Run("docx_merge_cells", new
    {
        input = mergeInput,
        changes = new[]
        {
            new { cells = new[] { Address(mergeCells[0]), Address(mergeCells[1]) } },
            new { cells = new[] { Address(mergeCells[2]), Address(mergeCells[3]) } }
        },
        output = merged,
        receiptOutput = Path.Combine(root, "merged-receipt.json")
    });
    var mergedState = ReadTable(merged, "merged-state");
    var finalCells = mergedState.GetProperty("rows")[0].GetProperty("cells");
    Require(finalCells.GetArrayLength() == 2
            && finalCells[0].GetProperty("gridSpan").GetInt32() == 2
            && finalCells[1].GetProperty("gridSpan").GetInt32() == 2,
        "batch merge used stale cell paths");

    RunInput("validate-openxml", insertedRows);
    RunInput("validate-openxml", insertedInsideMerge);
    RunInput("validate-openxml", batchColumns);
    RunInput("validate-openxml", insertedInsideOtherSpan);
    RunInput("validate-openxml", merged);
    Console.WriteLine("PASS table observation, row insertion, column edits, batch merges, and readback");
}
finally
{
    Directory.Delete(root, recursive: true);
}

JsonElement ReadTable(string input, string stem)
{
    var index = Run("docx_table_index", new
    {
        input,
        output = Path.Combine(root, stem + "-index.json")
    });
    return Run("docx_read_table", new
    {
        input,
        table = index.GetProperty("tables")[0].GetProperty("address").Clone(),
        output = Path.Combine(root, stem + "-table.json")
    });
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

void CreateDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = new Table(
        new TableProperties(),
        new TableGrid(
            new GridColumn { Width = "1200" },
            new GridColumn { Width = "1200" },
            new GridColumn { Width = "1800" },
            new GridColumn { Width = "1800" }));
    table.Append(new TableRow(
        new TableRowProperties(new TableHeader()),
        Cell("分组一", span: 2),
        Cell("分组二", span: 2)));
    table.Append(new TableRow(
        new TableRowProperties(new CantSplit()),
        Cell("甲", merge: MergedCellValues.Restart), Cell("甲一"), Cell("甲二"), Cell("甲三")));
    table.Append(new TableRow(
        Cell("", merge: MergedCellValues.Continue), Cell("乙一"), Cell("乙二"), Cell("乙三")));
    table.Append(new TableRow(Cell("独立"), Cell("丙一"), Cell("丙二"), Cell("丙三")));
    main.Document = new Document(new Body(table));
    main.Document.Save();
}

void CreateFlatDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = new Table(
        new TableProperties(),
        new TableGrid(
            new GridColumn { Width = "1000" }, new GridColumn { Width = "1000" },
            new GridColumn { Width = "1000" }, new GridColumn { Width = "1000" }),
        new TableRow(Cell("一"), Cell("二"), Cell("三"), Cell("四")));
    main.Document = new Document(new Body(table));
    main.Document.Save();
}

TableCell Cell(string text, int span = 1, MergedCellValues? merge = null)
{
    var properties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1200" });
    if (span > 1) properties.Append(new GridSpan { Val = span });
    if (merge is not null) properties.Append(new VerticalMerge { Val = merge.Value });
    return new TableCell(properties, new Paragraph(new Run(new Text(text))));
}

static JsonElement Address(JsonElement value) => value.GetProperty("address").Clone();
static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
