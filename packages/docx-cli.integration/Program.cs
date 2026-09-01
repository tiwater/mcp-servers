using System.Diagnostics;
using System.Security.Cryptography;
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
    var continuation = rows[2].GetProperty("cells")[0];
    var continuationObservation = Run("docx_read_object", new
    {
        input = original,
        addresses = new[] { continuation.GetProperty("address").Clone() },
        kinds = new[] { "paragraph" },
        output = Path.Combine(root, "continuation-observation.json")
    }).GetProperty("observations")[0].GetProperty("object");
    var observedOwner = continuationObservation.GetProperty("verticalMergeOwner");
    var expectedOwner = rows[1].GetProperty("cells")[0].GetProperty("address");
    Require(observedOwner.GetProperty("part").GetString() == expectedOwner.GetProperty("part").GetString()
            && observedOwner.GetProperty("path").GetString() == expectedOwner.GetProperty("path").GetString(),
        "narrow cell read did not expose the vertical-merge owner");
    Require(continuationObservation.GetProperty("logicalText").GetString()
            == rows[1].GetProperty("cells")[0].GetProperty("logicalText").GetString(),
        "narrow cell read did not resolve the restart cell text");

    var setBodyOutput = Path.Combine(root, "set-table-body.docx");
    var setBodyReceipt = Path.Combine(root, "set-table-body-receipt.json");
    var bodyColumns = initial.GetProperty("gridColumns").EnumerateArray()
        .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
        .ToArray();
    Run("docx_set_table_body", new
    {
        input = original,
        table = initial.GetProperty("address").Clone(),
        existingRows = new
        {
            first = rows[1].GetProperty("address").Clone(),
            last = rows[3].GetProperty("address").Clone(),
        },
        columns = bodyColumns,
        rows = new[]
        {
            new
            {
                prototypeRow = rows[1].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "精密度", verticalMerge = (string?)"restart" },
                    new { columns = new[] { "c1" }, text = "标准一", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果一", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "通过", verticalMerge = (string?)"restart" },
                }
            },
            new
            {
                prototypeRow = rows[3].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "", verticalMerge = (string?)"continue" },
                    new { columns = new[] { "c1" }, text = "标准二", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果二", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "", verticalMerge = (string?)"continue" },
                }
            }
        },
        output = setBodyOutput,
        receiptOutput = setBodyReceipt,
    });
    var setBodyState = ReadTable(setBodyOutput, "set-table-body");
    Require(setBodyState.GetProperty("rowCount").GetInt32() == 3,
        "set table body did not replace the complete data range");
    var setRows = setBodyState.GetProperty("rows");
    Require(CellAt(setBodyState, 1, 0).GetProperty("logicalText").GetString() == "精密度"
            && CellAt(setBodyState, 1, 1).GetProperty("logicalText").GetString() == "标准一"
            && CellAt(setBodyState, 1, 2).GetProperty("logicalText").GetString() == "结果一"
            && CellAt(setBodyState, 1, 3).GetProperty("logicalText").GetString() == "通过",
        "set table body shifted semantic columns");
    Require(CellAt(setBodyState, 2, 0).GetProperty("verticalMerge").GetString() == "continue"
            && CellAt(setBodyState, 2, 0).GetProperty("logicalText").GetString() == "精密度"
            && CellAt(setBodyState, 2, 1).GetProperty("logicalText").GetString() == "标准二"
            && CellAt(setBodyState, 2, 2).GetProperty("logicalText").GetString() == "结果二",
        "set table body lost the second semantic row");
    Require(CellAt(setBodyState, 1, 3).GetProperty("verticalMerge").GetString() == "restart"
            && CellAt(setBodyState, 2, 3).GetProperty("verticalMerge").GetString() == "continue"
            && CellAt(setBodyState, 2, 3).GetProperty("logicalText").GetString() == "通过",
        "set table body did not preserve the declared vertical group");

    var expandedBodyOutput = Path.Combine(root, "set-table-body-expanded.docx");
    Run("docx_set_table_body", new
    {
        input = original,
        table = initial.GetProperty("address").Clone(),
        existingRows = new
        {
            first = rows[1].GetProperty("address").Clone(),
            last = rows[3].GetProperty("address").Clone(),
        },
        columns = bodyColumns,
        rows = Enumerable.Range(1, 5).Select(index => new
        {
            prototypeRow = rows[3].GetProperty("address").Clone(),
            cells = new[]
            {
                new { columns = new[] { "c0", "c1" }, text = "横向" + index, verticalMerge = (string?)null },
                new { columns = new[] { "c2" }, text = "结果" + index, verticalMerge = (string?)null },
                new { columns = new[] { "c3" }, text = "结论" + index, verticalMerge = (string?)null },
            }
        }).ToArray(),
        output = expandedBodyOutput,
        receiptOutput = Path.Combine(root, "set-table-body-expanded-receipt.json"),
    });
    var expandedBodyState = ReadTable(expandedBodyOutput, "set-table-body-expanded");
    Require(expandedBodyState.GetProperty("rowCount").GetInt32() == 6
            && CellAt(expandedBodyState, 1, 0).GetProperty("gridSpan").GetInt32() == 2
            && CellAt(expandedBodyState, 5, 0).GetProperty("logicalText").GetString() == "横向5",
        "set table body did not expand rows with horizontal spans");

    var emptyBodyOutput = Path.Combine(root, "set-table-body-empty.docx");
    Run("docx_set_table_body", new
    {
        input = original,
        table = initial.GetProperty("address").Clone(),
        existingRows = new
        {
            first = rows[1].GetProperty("address").Clone(),
            last = rows[3].GetProperty("address").Clone(),
        },
        columns = bodyColumns,
        rows = Array.Empty<object>(),
        output = emptyBodyOutput,
        receiptOutput = Path.Combine(root, "set-table-body-empty-receipt.json"),
    });
    var emptyBodyState = ReadTable(emptyBodyOutput, "set-table-body-empty");
    Require(emptyBodyState.GetProperty("rowCount").GetInt32() == 1
            && CellAt(emptyBodyState, 0, 0).GetProperty("logicalText").GetString() == "分组一",
        "empty table body did not retain the row outside the replaced range");

    var emptyTableReceipt = Path.Combine(root, "set-table-body-empty-table-receipt.json");
    var emptyTable = RunExpectAtomicFailure("docx_set_table_body", original, emptyTableReceipt, new
    {
        input = original,
        table = initial.GetProperty("address").Clone(),
        existingRows = new
        {
            first = rows[0].GetProperty("address").Clone(),
            last = rows[3].GetProperty("address").Clone(),
        },
        columns = bodyColumns,
        rows = Array.Empty<object>(),
        output = original,
        receiptOutput = emptyTableReceipt,
    });
    Require(emptyTable.Contains("table-must-retain-at-least-one-row", StringComparison.Ordinal),
        "removing every table row did not fail explicitly");

    var failedBodyReceipt = Path.Combine(root, "set-table-body-failure-receipt.json");
    var failedBody = RunExpectAtomicFailure("docx_set_table_body", setBodyOutput, failedBodyReceipt, new
    {
        input = setBodyOutput,
        table = setBodyState.GetProperty("address").Clone(),
        existingRows = new
        {
            first = setRows[1].GetProperty("address").Clone(),
            last = setRows[2].GetProperty("address").Clone(),
        },
        columns = setBodyState.GetProperty("gridColumns").EnumerateArray()
            .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
            .ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = setRows[1].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "缺列", verticalMerge = (string?)null },
                    new { columns = new[] { "c1" }, text = "标准", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果", verticalMerge = (string?)null },
                }
            }
        },
        output = setBodyOutput,
        receiptOutput = failedBodyReceipt,
    });
    Require(failedBody.Contains("does-not-cover-table-grid", StringComparison.Ordinal),
        "incomplete table body did not fail explicitly");

    var orphanContinueReceipt = Path.Combine(root, "set-table-body-orphan-continue-receipt.json");
    var orphanContinue = RunExpectAtomicFailure("docx_set_table_body", setBodyOutput, orphanContinueReceipt, new
    {
        input = setBodyOutput,
        table = setBodyState.GetProperty("address").Clone(),
        existingRows = new
        {
            first = setRows[1].GetProperty("address").Clone(),
            last = setRows[2].GetProperty("address").Clone(),
        },
        columns = setBodyState.GetProperty("gridColumns").EnumerateArray()
            .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
            .ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = setRows[1].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "", verticalMerge = (string?)"continue" },
                    new { columns = new[] { "c1" }, text = "标准一", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果一", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "结论一", verticalMerge = (string?)null },
                }
            },
            new
            {
                prototypeRow = setRows[2].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "项目二", verticalMerge = (string?)null },
                    new { columns = new[] { "c1" }, text = "标准二", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果二", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "结论二", verticalMerge = (string?)null },
                }
            }
        },
        output = setBodyOutput,
        receiptOutput = orphanContinueReceipt,
    });
    Require(orphanContinue.Contains("verticalMerge-continue-without-previous", StringComparison.Ordinal),
        "orphan vertical continuation did not fail explicitly");

    var loneRestartReceipt = Path.Combine(root, "set-table-body-lone-restart-receipt.json");
    var loneRestart = RunExpectAtomicFailure("docx_set_table_body", setBodyOutput, loneRestartReceipt, new
    {
        input = setBodyOutput,
        table = setBodyState.GetProperty("address").Clone(),
        existingRows = new
        {
            first = setRows[1].GetProperty("address").Clone(),
            last = setRows[2].GetProperty("address").Clone(),
        },
        columns = setBodyState.GetProperty("gridColumns").EnumerateArray()
            .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
            .ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = setRows[1].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "项目一", verticalMerge = (string?)"restart" },
                    new { columns = new[] { "c1" }, text = "标准一", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果一", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "结论一", verticalMerge = (string?)null },
                }
            },
            new
            {
                prototypeRow = setRows[2].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "项目二", verticalMerge = (string?)null },
                    new { columns = new[] { "c1" }, text = "标准二", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果二", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "结论二", verticalMerge = (string?)null },
                }
            }
        },
        output = setBodyOutput,
        receiptOutput = loneRestartReceipt,
    });
    Require(loneRestart.Contains("verticalMerge-restart-without-continuation", StringComparison.Ordinal),
        "vertical restart without a continuation did not fail explicitly");

    var splitBodyReceipt = Path.Combine(root, "set-table-body-split-merge-receipt.json");
    var splitBody = RunExpectAtomicFailure("docx_set_table_body", original, splitBodyReceipt, new
    {
        input = original,
        table = initial.GetProperty("address").Clone(),
        existingRows = new
        {
            first = rows[1].GetProperty("address").Clone(),
            last = rows[1].GetProperty("address").Clone(),
        },
        columns = bodyColumns,
        rows = new[]
        {
            new
            {
                prototypeRow = rows[1].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "项目", verticalMerge = (string?)null },
                    new { columns = new[] { "c1" }, text = "标准", verticalMerge = (string?)null },
                    new { columns = new[] { "c2" }, text = "结果", verticalMerge = (string?)null },
                    new { columns = new[] { "c3" }, text = "结论", verticalMerge = (string?)null },
                }
            }
        },
        output = original,
        receiptOutput = splitBodyReceipt,
    });
    Require(splitBody.Contains("existingRows-split-vertical-merge", StringComparison.Ordinal),
        "table body range that split a vertical merge did not fail explicitly");

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
    RequireUniqueWordIdentities(insertedRows);
    RequireUniqueWordIdentities(insertedInsideMerge);

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

    foreach (var shape in new[] { "flat", "horizontal", "vertical", "mixed", "rectangle", "irregular", "multi-paragraph" })
        RunTableOperationMatrix(shape, path => CreateMatrixDocument(path, shape));
    RunTableOperationMatrix("nested", CreateNestedMatrixDocument, tableIndex: 1);
    RunNestedCellContentMatrix();
    RunTableFailureMatrix();

    Console.WriteLine("PASS table operation matrix across 8 table shapes");
}
finally
{
    Directory.Delete(root, recursive: true);
}

void RunTableOperationMatrix(string name, Action<string> create, int tableIndex = 0)
{
    var original = Path.Combine(root, $"matrix-{name}-original.docx");
    create(original);
    var baseline = ReadTableAt(original, $"matrix-{name}-baseline", tableIndex);
    RequireTableInvariants(baseline, $"{name}:baseline");
    var baselineSignature = TableSignature(baseline);
    RunExistingMergeRoundTrips(name, original, baseline, tableIndex);

    var written = Path.Combine(root, $"matrix-{name}-written.docx");
    var lastRow = baseline.GetProperty("rows").GetArrayLength() - 1;
    var writeTarget = CellAt(baseline, lastRow, 0).GetProperty("address").Clone();
    Run("docx_set_text", new
    {
        input = original,
        changes = new[] { new { target = writeTarget, text = $"write-{name}" } },
        output = written,
        receiptOutput = Path.Combine(root, $"matrix-{name}-written-receipt.json")
    });
    var writtenState = ReadTableAt(written, $"matrix-{name}-written", tableIndex);
    Require(CellText(CellAt(writtenState, lastRow, 0)) == $"write-{name}", $"{name}: cell write was not readable");
    RequireTableInvariants(writtenState, $"{name}:written");

    var rowStartSequence = Path.Combine(root, $"matrix-{name}-row-start.docx");
    Run("docx_insert_objects", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                sourceInput = original,
                sources = new[] { baseline.GetProperty("rows")[0].GetProperty("address").Clone() },
                targetParent = baseline.GetProperty("address").Clone(),
                before = baseline.GetProperty("rows")[0].GetProperty("address").Clone(),
                repeat = 1
            }
        },
        output = rowStartSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-start-insert-receipt.json")
    });
    var rowStarted = ReadTableAt(rowStartSequence, $"matrix-{name}-row-start-inserted", tableIndex);
    Require(rowStarted.GetProperty("rowCount").GetInt32() == baseline.GetProperty("rowCount").GetInt32() + 1,
        $"{name}: row insertion at table start count mismatch");
    RequireUniqueWordIdentities(rowStartSequence);
    Run("docx_delete_object", new
    {
        input = rowStartSequence,
        changes = new[] { new { addresses = new[] { rowStarted.GetProperty("rows")[0].GetProperty("address").Clone() } } },
        output = rowStartSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-start-delete-receipt.json")
    });
    var rowStartRoundTrip = ReadTableAt(rowStartSequence, $"matrix-{name}-row-start-roundtrip", tableIndex);
    Require(TableSignature(rowStartRoundTrip) == baselineSignature,
        $"{name}: insert/delete at table start changed the base table");
    RequireTableInvariants(rowStartRoundTrip, $"{name}:row-start-roundtrip");

    var rowSequence = Path.Combine(root, $"matrix-{name}-row-sequence.docx");
    Run("docx_insert_objects", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                sourceInput = original,
                sources = new[] { baseline.GetProperty("rows")[lastRow].GetProperty("address").Clone() },
                targetParent = baseline.GetProperty("address").Clone(),
                repeat = 1
            }
        },
        output = rowSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-insert-receipt.json")
    });
    var rowInserted = ReadTableAt(rowSequence, $"matrix-{name}-row-inserted", tableIndex);
    Require(rowInserted.GetProperty("rowCount").GetInt32() == baseline.GetProperty("rowCount").GetInt32() + 1,
        $"{name}: row insertion count mismatch");
    RequireUniqueWordIdentities(rowSequence);
    var insertedRowIndex = rowInserted.GetProperty("rows").GetArrayLength() - 1;
    Run("docx_set_text", new
    {
        input = rowSequence,
        changes = new[] { new { target = CellAt(rowInserted, insertedRowIndex, 0).GetProperty("address").Clone(), text = "row-chain" } },
        output = rowSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-write-receipt.json")
    });
    var rowWritten = ReadTableAt(rowSequence, $"matrix-{name}-row-written", tableIndex);
    var rowMergeCells = new[]
    {
        CellAt(rowWritten, insertedRowIndex, 1).GetProperty("address").Clone(),
        CellAt(rowWritten, insertedRowIndex, 2).GetProperty("address").Clone()
    };
    Run("docx_merge_cells", new
    {
        input = rowSequence,
        changes = new[] { new { cells = rowMergeCells } },
        output = rowSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-merge-receipt.json")
    });
    var rowMerged = ReadTableAt(rowSequence, $"matrix-{name}-row-merged", tableIndex);
    var rowMergeOwner = CellAt(rowMerged, insertedRowIndex, 1);
    Require(rowMergeOwner.GetProperty("gridSpan").GetInt32() == 2, $"{name}: inserted-row merge did not span two columns");
    Run("docx_split_cells", new
    {
        input = rowSequence,
        changes = new[] { new { cells = new[] { rowMergeOwner.GetProperty("address").Clone() } } },
        output = rowSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-split-receipt.json")
    });
    var rowSplit = ReadTableAt(rowSequence, $"matrix-{name}-row-split", tableIndex);
    Run("docx_delete_object", new
    {
        input = rowSequence,
        changes = new[] { new { addresses = new[] { rowSplit.GetProperty("rows")[insertedRowIndex].GetProperty("address").Clone() } } },
        output = rowSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-row-delete-receipt.json")
    });
    var rowRoundTrip = ReadTableAt(rowSequence, $"matrix-{name}-row-roundtrip", tableIndex);
    Require(TableSignature(rowRoundTrip) == baselineSignature, $"{name}: insert/write/merge/split/delete row sequence changed the base table");
    RequireTableInvariants(rowRoundTrip, $"{name}:row-roundtrip");

    var firstWorkspaceRow = baseline.GetProperty("rows").GetArrayLength() - 2;
    var secondWorkspaceRow = firstWorkspaceRow + 1;
    var verticalSequence = Path.Combine(root, $"matrix-{name}-vertical-sequence.docx");
    Run("docx_merge_cells", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                cells = new[]
                {
                    CellAt(baseline, firstWorkspaceRow, 3).GetProperty("address").Clone(),
                    CellAt(baseline, secondWorkspaceRow, 3).GetProperty("address").Clone()
                }
            }
        },
        output = verticalSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-vertical-merge-receipt.json")
    });
    var verticalMerged = ReadTableAt(verticalSequence, $"matrix-{name}-vertical-merged", tableIndex);
    var verticalOwner = CellAt(verticalMerged, firstWorkspaceRow, 3);
    var verticalContinuation = CellAt(verticalMerged, secondWorkspaceRow, 3);
    Require(verticalOwner.GetProperty("verticalMerge").GetString() == "restart"
            && verticalContinuation.GetProperty("verticalMerge").GetString() == "continue"
            && verticalContinuation.GetProperty("verticalMergeOwner").GetRawText()
               == verticalOwner.GetProperty("address").GetRawText(),
        $"{name}: vertical merge owner relationship is wrong");
    Require(CellText(verticalOwner) == "尾一\n尾二", $"{name}: vertical merge lost selected cell content");
    Run("docx_set_text", new
    {
        input = verticalSequence,
        changes = new[] { new { target = verticalOwner.GetProperty("address").Clone(), text = "vertical-chain" } },
        output = verticalSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-vertical-write-receipt.json")
    });
    var verticalWritten = ReadTableAt(verticalSequence, $"matrix-{name}-vertical-written", tableIndex);
    Run("docx_split_cells", new
    {
        input = verticalSequence,
        changes = new[] { new { cells = new[] { CellAt(verticalWritten, firstWorkspaceRow, 3).GetProperty("address").Clone() } } },
        output = verticalSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-vertical-split-receipt.json")
    });
    var verticalSplit = ReadTableAt(verticalSequence, $"matrix-{name}-vertical-split", tableIndex);
    Require(TableStructureSignature(verticalSplit) == TableStructureSignature(baseline),
        $"{name}: vertical merge/write/split sequence changed the base table structure");
    Require(CellText(CellAt(verticalSplit, firstWorkspaceRow, 3)) == "vertical-chain"
            && CellText(CellAt(verticalSplit, secondWorkspaceRow, 3)) == "",
        $"{name}: vertical split did not retain merged owner content");
    RequireTableInvariants(verticalSplit, $"{name}:vertical-split");

    var rectangleSequence = Path.Combine(root, $"matrix-{name}-rectangle-sequence.docx");
    var rectangleCells = new[]
    {
        CellAt(baseline, firstWorkspaceRow, 1).GetProperty("address").Clone(),
        CellAt(baseline, firstWorkspaceRow, 2).GetProperty("address").Clone(),
        CellAt(baseline, secondWorkspaceRow, 1).GetProperty("address").Clone(),
        CellAt(baseline, secondWorkspaceRow, 2).GetProperty("address").Clone()
    };
    Run("docx_merge_cells", new
    {
        input = original,
        changes = new[] { new { cells = rectangleCells } },
        output = rectangleSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-rectangle-merge-receipt.json")
    });
    var rectangleMerged = ReadTableAt(rectangleSequence, $"matrix-{name}-rectangle-merged", tableIndex);
    var rectangleOwner = CellAt(rectangleMerged, firstWorkspaceRow, 1);
    var rectangleContinuation = CellAt(rectangleMerged, secondWorkspaceRow, 1);
    Require(rectangleOwner.GetProperty("gridSpan").GetInt32() == 2
            && rectangleOwner.GetProperty("verticalMerge").GetString() == "restart"
            && rectangleContinuation.GetProperty("verticalMerge").GetString() == "continue",
        $"{name}: 2x2 merge did not create one rectangular owner");
    Run("docx_set_text", new
    {
        input = rectangleSequence,
        changes = new[] { new { target = rectangleOwner.GetProperty("address").Clone(), text = "rectangle-chain" } },
        output = rectangleSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-rectangle-write-receipt.json")
    });
    var rectangleWritten = ReadTableAt(rectangleSequence, $"matrix-{name}-rectangle-written", tableIndex);
    var currentContinuationRow = rectangleWritten.GetProperty("rows")[secondWorkspaceRow];
    Run("docx_insert_objects", new
    {
        input = rectangleSequence,
        changes = new[]
        {
            new
            {
                sourceInput = rectangleSequence,
                sources = new[] { currentContinuationRow.GetProperty("address").Clone() },
                targetParent = rectangleWritten.GetProperty("address").Clone(),
                before = currentContinuationRow.GetProperty("address").Clone(),
                repeat = 1
            }
        },
        output = rectangleSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-rectangle-insert-receipt.json")
    });
    var rectangleExtended = ReadTableAt(rectangleSequence, $"matrix-{name}-rectangle-extended", tableIndex);
    var ownerPath = CellAt(rectangleExtended, firstWorkspaceRow, 1).GetProperty("address").GetProperty("path").GetString();
    foreach (var rowIndex in new[] { firstWorkspaceRow + 1, firstWorkspaceRow + 2 })
    {
        var cell = CellAt(rectangleExtended, rowIndex, 1);
        Require(cell.GetProperty("verticalMerge").GetString() == "continue"
                && cell.GetProperty("verticalMergeOwner").GetProperty("path").GetString() == ownerPath,
            $"{name}: inserted row did not remain in the 2x2 merge group");
    }
    RequireUniqueWordIdentities(rectangleSequence);
    Run("docx_split_cells", new
    {
        input = rectangleSequence,
        changes = new[] { new { cells = new[] { CellAt(rectangleExtended, firstWorkspaceRow, 1).GetProperty("address").Clone() } } },
        output = rectangleSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-rectangle-split-receipt.json")
    });
    var rectangleSplit = ReadTableAt(rectangleSequence, $"matrix-{name}-rectangle-split", tableIndex);
    Run("docx_set_text", new
    {
        input = rectangleSequence,
        changes = new[] { new { target = CellAt(rectangleSplit, firstWorkspaceRow, 1).GetProperty("address").Clone(), text = "" } },
        output = rectangleSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-rectangle-clear-receipt.json")
    });
    var rectangleCleared = ReadTableAt(rectangleSequence, $"matrix-{name}-rectangle-cleared", tableIndex);
    Run("docx_delete_object", new
    {
        input = rectangleSequence,
        changes = new[] { new { addresses = new[] { rectangleCleared.GetProperty("rows")[firstWorkspaceRow + 1].GetProperty("address").Clone() } } },
        output = rectangleSequence,
        receiptOutput = Path.Combine(root, $"matrix-{name}-rectangle-delete-receipt.json")
    });
    var rectangleRoundTrip = ReadTableAt(rectangleSequence, $"matrix-{name}-rectangle-roundtrip", tableIndex);
    Require(TableSignature(rectangleRoundTrip) == baselineSignature,
        $"{name}: merge/write/insert/split/delete sequence changed the base table");
    RequireTableInvariants(rectangleRoundTrip, $"{name}:rectangle-roundtrip");

    var columnCount = baseline.GetProperty("columnCount").GetInt32();
    foreach (var (label, sourceColumn, insertionColumn) in new[]
             {
                 ("start", 0, 0),
                 ("middle", 0, 1),
                 ("end", columnCount - 1, columnCount)
             })
        RunColumnOperationSequence(name, label, original, baseline, baselineSignature, tableIndex,
            lastRow, sourceColumn, insertionColumn);

    foreach (var output in new[] { written, rowStartSequence, rowSequence, verticalSequence, rectangleSequence })
        RunInput("validate-openxml", output);
    Console.WriteLine($"PASS table matrix: {name}");
}

void RunExistingMergeRoundTrips(string shape, string original, JsonElement baseline, int tableIndex)
{
    var rows = baseline.GetProperty("rows");
    var regions = new List<(int Row, int Start, int Width, int Height)>();
    for (var rowIndex = 0; rowIndex < rows.GetArrayLength(); rowIndex++)
    foreach (var cell in rows[rowIndex].GetProperty("cells").EnumerateArray())
    {
        var merge = cell.GetProperty("verticalMerge");
        if (merge.ValueKind != JsonValueKind.Null && merge.GetString() == "continue") continue;
        var start = cell.GetProperty("gridColumnStart").GetInt32();
        var width = cell.GetProperty("gridSpan").GetInt32();
        var height = 1;
        if (merge.ValueKind != JsonValueKind.Null && merge.GetString() == "restart")
        {
            var ownerPath = cell.GetProperty("address").GetProperty("path").GetString();
            for (var next = rowIndex + 1; next < rows.GetArrayLength(); next++)
            {
                var continuation = rows[next].GetProperty("cells").EnumerateArray().FirstOrDefault(candidate =>
                    candidate.GetProperty("gridColumnStart").GetInt32() == start
                    && candidate.GetProperty("gridSpan").GetInt32() == width);
                if (continuation.ValueKind == JsonValueKind.Undefined
                    || continuation.GetProperty("verticalMerge").ValueKind == JsonValueKind.Null
                    || continuation.GetProperty("verticalMerge").GetString() != "continue"
                    || continuation.GetProperty("verticalMergeOwner").GetProperty("path").GetString() != ownerPath)
                    break;
                height++;
            }
        }
        if (width > 1 || height > 1) regions.Add((rowIndex, start, width, height));
    }

    for (var regionIndex = 0; regionIndex < regions.Count; regionIndex++)
    {
        var region = regions[regionIndex];
        var ownerWrite = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}-write.docx");
        Run("docx_set_text", new
        {
            input = original,
            changes = new[]
            {
                new
                {
                    target = CellAt(baseline, region.Row, region.Start).GetProperty("address").Clone(),
                    text = $"owner-{shape}-{regionIndex}"
                }
            },
            output = ownerWrite,
            receiptOutput = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}-write-receipt.json")
        });
        var ownerWritten = ReadTableAt(ownerWrite, $"matrix-{shape}-existing-merge-{regionIndex}-written", tableIndex);
        Require(TableStructureSignature(ownerWritten) == TableStructureSignature(baseline)
                && CellText(CellAt(ownerWritten, region.Row, region.Start)) == $"owner-{shape}-{regionIndex}",
            $"{shape}: writing an existing merge owner changed its structure");
        RunInput("validate-openxml", ownerWrite);

        var output = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}.docx");
        var originalOwnerText = CellText(CellAt(baseline, region.Row, region.Start));
        Run("docx_split_cells", new
        {
            input = original,
            changes = new[]
            {
                new { cells = new[] { CellAt(baseline, region.Row, region.Start).GetProperty("address").Clone() } }
            },
            output,
            receiptOutput = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}-split-receipt.json")
        });
        var split = ReadTableAt(output, $"matrix-{shape}-existing-merge-{regionIndex}-split", tableIndex);
        var cells = new List<JsonElement>();
        for (var row = region.Row; row < region.Row + region.Height; row++)
        for (var column = region.Start; column < region.Start + region.Width; column++)
        {
            var cell = CellAt(split, row, column);
            Require(cell.GetProperty("gridSpan").GetInt32() == 1
                    && cell.GetProperty("verticalMerge").ValueKind == JsonValueKind.Null,
                $"{shape}: split did not expose every cell in an existing merge");
            cells.Add(cell.GetProperty("address").Clone());
        }
        Run("docx_merge_cells", new
        {
            input = output,
            changes = new[] { new { cells = cells.ToArray() } },
            output,
            receiptOutput = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}-merge-receipt.json")
        });
        var restored = ReadTableAt(output, $"matrix-{shape}-existing-merge-{regionIndex}-restored", tableIndex);
        Require(TableStructureSignature(restored) == TableStructureSignature(baseline),
            $"{shape}: split/merge changed existing merge structure");
        Require(CellText(CellAt(restored, region.Row, region.Start)).TrimEnd('\n') == originalOwnerText.TrimEnd('\n'),
            $"{shape}: split/merge lost existing owner text");
        RequireTableInvariants(restored, $"{shape}:existing-merge-{regionIndex}");
        RunInput("validate-openxml", output);

        var deleted = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}-deleted.docx");
        Run("docx_delete_object", new
        {
            input = original,
            changes = new[]
            {
                new
                {
                    addresses = Enumerable.Range(region.Row, region.Height)
                        .Select(row => baseline.GetProperty("rows")[row].GetProperty("address").Clone()).ToArray()
                }
            },
            output = deleted,
            receiptOutput = Path.Combine(root, $"matrix-{shape}-existing-merge-{regionIndex}-delete-receipt.json")
        });
        var deletedState = ReadTableAt(deleted, $"matrix-{shape}-existing-merge-{regionIndex}-deleted", tableIndex);
        Require(deletedState.GetProperty("rowCount").GetInt32()
                == baseline.GetProperty("rowCount").GetInt32() - region.Height,
            $"{shape}: deleting a closed merged-row set changed the wrong row count");
        RequireTableInvariants(deletedState, $"{shape}:existing-merge-{regionIndex}-deleted");
        RunInput("validate-openxml", deleted);
    }
}

void RunColumnOperationSequence(
    string shape,
    string label,
    string original,
    JsonElement baseline,
    string baselineSignature,
    int tableIndex,
    int lastRow,
    int sourceColumn,
    int insertionColumn)
{
    var output = Path.Combine(root, $"matrix-{shape}-column-{label}.docx");
    var grid = baseline.GetProperty("gridColumns");
    var before = insertionColumn < grid.GetArrayLength()
        ? grid[insertionColumn].GetProperty("address").Clone()
        : (JsonElement?)null;
    Run("docx_insert_table_columns", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                table = baseline.GetProperty("address").Clone(),
                sourceColumn = grid[sourceColumn].GetProperty("address").Clone(),
                before
            }
        },
        output,
        receiptOutput = Path.Combine(root, $"matrix-{shape}-column-{label}-insert-receipt.json")
    });
    var inserted = ReadTableAt(output, $"matrix-{shape}-column-{label}-inserted", tableIndex);
    Require(inserted.GetProperty("columnCount").GetInt32() == baseline.GetProperty("columnCount").GetInt32() + 1,
        $"{shape}:{label}: column insertion count mismatch");
    Run("docx_set_text", new
    {
        input = output,
        changes = new[]
        {
            new
            {
                target = CellAt(inserted, lastRow, insertionColumn).GetProperty("address").Clone(),
                text = $"column-{label}"
            }
        },
        output,
        receiptOutput = Path.Combine(root, $"matrix-{shape}-column-{label}-write-receipt.json")
    });
    var written = ReadTableAt(output, $"matrix-{shape}-column-{label}-written", tableIndex);
    Run("docx_set_text", new
    {
        input = output,
        changes = new[]
        {
            new { target = CellAt(written, lastRow, insertionColumn).GetProperty("address").Clone(), text = "" }
        },
        output,
        receiptOutput = Path.Combine(root, $"matrix-{shape}-column-{label}-clear-receipt.json")
    });
    var cleared = ReadTableAt(output, $"matrix-{shape}-column-{label}-cleared", tableIndex);
    Run("docx_delete_table_columns", new
    {
        input = output,
        changes = new[]
        {
            new
            {
                table = cleared.GetProperty("address").Clone(),
                columns = new[] { cleared.GetProperty("gridColumns")[insertionColumn].GetProperty("address").Clone() }
            }
        },
        output,
        receiptOutput = Path.Combine(root, $"matrix-{shape}-column-{label}-delete-receipt.json")
    });
    var roundTrip = ReadTableAt(output, $"matrix-{shape}-column-{label}-roundtrip", tableIndex);
    Require(TableSignature(roundTrip) == baselineSignature,
        $"{shape}:{label}: insert/write/delete column sequence changed the base table");
    RequireTableInvariants(roundTrip, $"{shape}:column-{label}-roundtrip");
    RunInput("validate-openxml", output);
}

void RunNestedCellContentMatrix()
{
    var input = Path.Combine(root, "matrix-nested-cell.docx");
    CreateNestedMatrixDocument(input);
    var outer = ReadTableAt(input, "matrix-nested-cell-outer", 0);
    var inner = ReadTableAt(input, "matrix-nested-cell-inner", 1);
    var innerSignature = TableSignature(inner);
    var output = Path.Combine(root, "matrix-nested-cell-merged.docx");
    Run("docx_merge_cells", new
    {
        input,
        changes = new[]
        {
            new
            {
                cells = new[]
                {
                    CellAt(outer, 0, 0).GetProperty("address").Clone(),
                    CellAt(outer, 0, 1).GetProperty("address").Clone()
                }
            }
        },
        output,
        receiptOutput = Path.Combine(root, "matrix-nested-cell-merge-receipt.json")
    });
    var mergedOuter = ReadTableAt(output, "matrix-nested-cell-merged-outer", 0);
    var mergedInner = ReadTableAt(output, "matrix-nested-cell-merged-inner", 1);
    Require(mergedOuter.GetProperty("rows")[0].GetProperty("cells").GetArrayLength() == 1
            && CellAt(mergedOuter, 0, 0).GetProperty("gridSpan").GetInt32() == 2,
        "nested-cell merge did not create one owner");
    Require(TableSignature(mergedInner) == innerSignature, "nested table content was lost during outer-cell merge");
    Run("docx_split_cells", new
    {
        input = output,
        changes = new[]
        {
            new { cells = new[] { CellAt(mergedOuter, 0, 0).GetProperty("address").Clone() } }
        },
        output,
        receiptOutput = Path.Combine(root, "matrix-nested-cell-split-receipt.json")
    });
    var splitOuter = ReadTableAt(output, "matrix-nested-cell-split-outer", 0);
    var splitInner = ReadTableAt(output, "matrix-nested-cell-split-inner", 1);
    var index = Run("docx_table_index", new
    {
        input = output,
        output = Path.Combine(root, "matrix-nested-cell-final-index.json")
    });
    Require(splitOuter.GetProperty("rows")[0].GetProperty("cells").GetArrayLength() == 2,
        "nested-cell split did not restore two cells");
    Require(index.GetProperty("tables").GetArrayLength() == 2,
        "nested-cell split duplicated or deleted the nested table");
    Require(TableSignature(splitInner) == innerSignature,
        "nested table content changed during outer-cell split");
    RequireTableInvariants(splitOuter, "nested-cell:outer");
    RequireTableInvariants(splitInner, "nested-cell:inner");
    RunInput("validate-openxml", output);
    Console.WriteLine("PASS nested table cell-content preservation");
}

void RunTableFailureMatrix()
{
    var input = Path.Combine(root, "matrix-failure.docx");
    CreateMatrixDocument(input, "mixed");
    var state = ReadTable(input, "matrix-failure");
    var rows = state.GetProperty("rows");
    var firstWorkspaceRow = rows.GetArrayLength() - 2;
    var secondWorkspaceRow = firstWorkspaceRow + 1;

    var duplicateReceipt = Path.Combine(root, "matrix-failure-duplicate-write.json");
    var duplicateTarget = CellAt(state, firstWorkspaceRow, 0).GetProperty("address").Clone();
    var duplicateError = RunExpectAtomicFailure("docx_set_text", input, duplicateReceipt, new
    {
        input,
        changes = new[] { new { target = duplicateTarget, text = "一" }, new { target = duplicateTarget, text = "二" } },
        output = input,
        receiptOutput = duplicateReceipt
    });
    Require(duplicateError.Contains("target-address-duplicate", StringComparison.Ordinal),
        "duplicate set-text target did not fail explicitly");

    var nonRectangleReceipt = Path.Combine(root, "matrix-failure-non-rectangle.json");
    var nonRectangleError = RunExpectAtomicFailure("docx_merge_cells", input, nonRectangleReceipt, new
    {
        input,
        changes = new[]
        {
            new
            {
                cells = new[]
                {
                    CellAt(state, firstWorkspaceRow, 1).GetProperty("address").Clone(),
                    CellAt(state, firstWorkspaceRow, 2).GetProperty("address").Clone(),
                    CellAt(state, secondWorkspaceRow, 1).GetProperty("address").Clone()
                }
            }
        },
        output = input,
        receiptOutput = nonRectangleReceipt
    });
    Require(nonRectangleError.Contains("merge-cell-selection-must-be-one-closed-rectangle", StringComparison.Ordinal),
        "non-rectangular merge did not fail explicitly");

    var splitReceipt = Path.Combine(root, "matrix-failure-split-owner.json");
    var splitError = RunExpectAtomicFailure("docx_split_cells", input, splitReceipt, new
    {
        input,
        changes = new[] { new { cells = new[] { CellAt(state, firstWorkspaceRow, 0).GetProperty("address").Clone() } } },
        output = input,
        receiptOutput = splitReceipt
    });
    Require(splitError.Contains("split-cell-must-be-a-merge-owner", StringComparison.Ordinal),
        "split of an unmerged cell did not fail explicitly");

    RunInput("validate-openxml", input);
    Console.WriteLine("PASS table failure atomicity matrix");
}

JsonElement ReadTable(string input, string stem)
{
    return ReadTableAt(input, stem, 0);
}

JsonElement ReadTableAt(string input, string stem, int tableIndex)
{
    var index = Run("docx_table_index", new
    {
        input,
        output = Path.Combine(root, stem + "-index.json")
    });
    Require(index.GetProperty("tables").GetArrayLength() > tableIndex, $"{stem}: table index {tableIndex} missing");
    return Run("docx_read_table", new
    {
        input,
        table = index.GetProperty("tables")[tableIndex].GetProperty("address").Clone(),
        output = Path.Combine(root, stem + "-table.json")
    });
}

JsonElement Run(string command, object request)
{
    var requestPath = Path.Combine(root, Guid.NewGuid().ToString("N") + ".json");
    var requestJson = JsonSerializer.Serialize(request);
    File.WriteAllText(requestPath, requestJson);
    var result = Execute(command, requestPath);
    Require(result.ExitCode == 0, $"{command} failed: {result.Error}\n{result.Output}\nrequest={requestJson}");
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

string RunExpectAtomicFailure(string command, string input, string receipt, object request)
{
    var before = FileHash(input);
    var error = RunExpectFailure(command, request);
    Require(FileHash(input) == before, $"{command} changed its input after failure");
    Require(!File.Exists(receipt), $"{command} retained a success receipt after failure");
    return error;
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

string FileHash(string path)
{
    using var stream = File.OpenRead(path);
    return Convert.ToHexString(SHA256.HashData(stream));
}

JsonElement CellAt(JsonElement table, int rowIndex, int gridColumnStart)
{
    var match = table.GetProperty("rows")[rowIndex].GetProperty("cells").EnumerateArray()
        .FirstOrDefault(cell => cell.GetProperty("gridColumnStart").GetInt32() == gridColumnStart);
    Require(match.ValueKind != JsonValueKind.Undefined,
        $"cell missing: row={rowIndex}, gridColumnStart={gridColumnStart}");
    return match;
}

string CellText(JsonElement cell)
    => string.Join("\n", cell.GetProperty("paragraphs").EnumerateArray()
        .Select(paragraph => paragraph.GetProperty("text").GetString() ?? ""));

void RequireTableInvariants(JsonElement table, string name)
{
    var columnCount = table.GetProperty("columnCount").GetInt32();
    Require(columnCount == table.GetProperty("gridColumns").GetArrayLength(), $"{name}: grid column count mismatch");
    var seenAddresses = new HashSet<string>(StringComparer.Ordinal);
    var logicalTextByAddress = new Dictionary<string, string>(StringComparer.Ordinal);
    foreach (var row in table.GetProperty("rows").EnumerateArray())
    {
        var cursor = row.GetProperty("gridBefore").GetInt32();
        foreach (var cell in row.GetProperty("cells").EnumerateArray())
        {
            Require(cell.GetProperty("gridColumnStart").GetInt32() == cursor, $"{name}: cell grid has a gap or overlap");
            var span = cell.GetProperty("gridSpan").GetInt32();
            Require(span > 0, $"{name}: cell span is not positive");
            cursor += span;
            var address = cell.GetProperty("address").GetProperty("path").GetString() ?? "";
            Require(seenAddresses.Add(address), $"{name}: duplicate cell address");
            var merge = cell.GetProperty("verticalMerge");
            var owner = cell.GetProperty("verticalMergeOwner");
            var logicalText = cell.GetProperty("logicalText").GetString() ?? "";
            if (merge.ValueKind == JsonValueKind.Null)
            {
                Require(owner.ValueKind == JsonValueKind.Null, $"{name}: unmerged cell has an owner");
                Require(logicalText == CellText(cell), $"{name}: unmerged logical text differs from physical text");
            }
            else if (merge.GetString() == "restart")
            {
                Require(owner.GetProperty("path").GetString() == address, $"{name}: restart cell does not own itself");
                Require(logicalText == CellText(cell), $"{name}: restart logical text differs from physical text");
            }
            else
            {
                var ownerPath = owner.GetProperty("path").GetString() ?? "";
                Require(owner.ValueKind == JsonValueKind.Object
                        && seenAddresses.Contains(ownerPath),
                    $"{name}: continuation cell owner was not observed earlier");
                Require(logicalTextByAddress.TryGetValue(ownerPath, out var ownerText)
                        && logicalText == ownerText,
                    $"{name}: continuation logical text does not resolve its owner");
            }
            logicalTextByAddress[address] = logicalText;
        }
        cursor += row.GetProperty("gridAfter").GetInt32();
        Require(cursor == columnCount, $"{name}: row does not cover the declared table grid");
    }
}

string TableSignature(JsonElement table)
{
    var signature = new
    {
        columnCount = table.GetProperty("columnCount").GetInt32(),
        widths = table.GetProperty("gridColumns").EnumerateArray()
            .Select(column => column.GetProperty("widthTwips").GetInt32()).ToArray(),
        rows = table.GetProperty("rows").EnumerateArray().Select(row => new
        {
            repeatHeader = row.GetProperty("repeatHeader").GetBoolean(),
            cantSplit = row.GetProperty("cantSplit").GetBoolean(),
            gridBefore = row.GetProperty("gridBefore").GetInt32(),
            gridAfter = row.GetProperty("gridAfter").GetInt32(),
            cells = row.GetProperty("cells").EnumerateArray().Select(cell => new
            {
                start = cell.GetProperty("gridColumnStart").GetInt32(),
                span = cell.GetProperty("gridSpan").GetInt32(),
                merge = cell.GetProperty("verticalMerge").ValueKind == JsonValueKind.Null
                    ? null : cell.GetProperty("verticalMerge").GetString(),
                text = CellText(cell)
            }).ToArray()
        }).ToArray()
    };
    return JsonSerializer.Serialize(signature);
}

string TableStructureSignature(JsonElement table)
{
    var signature = new
    {
        columnCount = table.GetProperty("columnCount").GetInt32(),
        widths = table.GetProperty("gridColumns").EnumerateArray()
            .Select(column => column.GetProperty("widthTwips").GetInt32()).ToArray(),
        rows = table.GetProperty("rows").EnumerateArray().Select(row => new
        {
            repeatHeader = row.GetProperty("repeatHeader").GetBoolean(),
            cantSplit = row.GetProperty("cantSplit").GetBoolean(),
            gridBefore = row.GetProperty("gridBefore").GetInt32(),
            gridAfter = row.GetProperty("gridAfter").GetInt32(),
            cells = row.GetProperty("cells").EnumerateArray().Select(cell => new
            {
                start = cell.GetProperty("gridColumnStart").GetInt32(),
                span = cell.GetProperty("gridSpan").GetInt32(),
                merge = cell.GetProperty("verticalMerge").ValueKind == JsonValueKind.Null
                    ? null : cell.GetProperty("verticalMerge").GetString()
            }).ToArray()
        }).ToArray()
    };
    return JsonSerializer.Serialize(signature);
}

void CreateMatrixDocument(string path, string shape)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    main.Document = new Document(new Body(MatrixTable(shape)));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

void CreateNestedMatrixDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var inner = MatrixTable("mixed");
    var nestedCell = new TableCell(
        new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" }),
        new Paragraph(new Run(new Text("嵌套表前"))), inner, new Paragraph());
    var outer = new Table(
        new TableProperties(),
        new TableGrid(new GridColumn { Width = "2400" }, new GridColumn { Width = "2400" }),
        new TableRow(Cell("外层"), nestedCell));
    main.Document = new Document(new Body(outer));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

Table MatrixTable(string shape)
{
    var table = new Table(
        new TableProperties(),
        new TableGrid(
            new GridColumn { Width = "1200" }, new GridColumn { Width = "1200" },
            new GridColumn { Width = "1200" }, new GridColumn { Width = "1200" }));
    switch (shape)
    {
        case "flat":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("甲"), Cell("乙"), Cell("丙"), Cell("丁")));
            table.Append(new TableRow(new TableRowProperties(new CantSplit()), Cell("一"), Cell("二"), Cell("三"), Cell("四")));
            table.Append(DataRow("五", "六", "七", "八"));
            break;
        case "horizontal":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("左组", span: 2), Cell("右组", span: 2)));
            table.Append(new TableRow(new TableRowProperties(new CantSplit()), Cell("一"), Cell("二"), Cell("三"), Cell("四")));
            table.Append(DataRow("五", "六", "七", "八"));
            break;
        case "vertical":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("甲"), Cell("乙"), Cell("丙"), Cell("丁")));
            table.Append(new TableRow(new TableRowProperties(new CantSplit()),
                Cell("纵组", merge: MergedCellValues.Restart), Cell("二"), Cell("三"), Cell("四")));
            table.Append(new TableRow(Cell("", merge: MergedCellValues.Continue), Cell("六"), Cell("七"), Cell("八")));
            break;
        case "mixed":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("左组", span: 2), Cell("右组", span: 2)));
            table.Append(new TableRow(new TableRowProperties(new CantSplit()),
                Cell("纵组", merge: MergedCellValues.Restart), Cell("二"), Cell("三"), Cell("四")));
            table.Append(new TableRow(Cell("", merge: MergedCellValues.Continue), Cell("六"), Cell("七"), Cell("八")));
            break;
        case "rectangle":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("甲"), Cell("乙"), Cell("丙"), Cell("丁")));
            table.Append(new TableRow(new TableRowProperties(new CantSplit()),
                Cell("二维组", span: 2, merge: MergedCellValues.Restart), Cell("三"), Cell("四")));
            table.Append(new TableRow(Cell("", span: 2, merge: MergedCellValues.Continue), Cell("七"), Cell("八")));
            break;
        case "irregular":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("甲"), Cell("乙"), Cell("丙"), Cell("丁")));
            table.Append(new TableRow(
                new TableRowProperties(new CantSplit(), new GridBefore { Val = 1 }, new GridAfter { Val = 1 }),
                Cell("中一"), Cell("中二")));
            table.Append(DataRow("五", "六", "七", "八"));
            break;
        case "multi-paragraph":
            table.Append(new TableRow(new TableRowProperties(new TableHeader()), Cell("甲"), Cell("乙"), Cell("丙"), Cell("丁")));
            table.Append(new TableRow(new TableRowProperties(new CantSplit()),
                CellWithParagraphs("第一段", "第二段"), Cell(""), Cell("三"), Cell("四")));
            table.Append(DataRow("五", "六", "", "八"));
            break;
        default:
            throw new InvalidOperationException($"unknown table shape: {shape}");
    }
    table.Append(DataRow("工作一", "", "", "尾一"));
    table.Append(DataRow("工作二", "", "", "尾二"));
    return table;
}

TableRow DataRow(string first, string second, string third, string fourth)
    => new(Cell(first), Cell(second), Cell(third), Cell(fourth));

TableCell CellWithParagraphs(params string[] values)
{
    var properties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1200" });
    return new TableCell(new OpenXmlElement[] { properties }
        .Concat(values.Select(value => (OpenXmlElement)new Paragraph(new Run(new Text(value))))));
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
    AssignParagraphIdentities(main.Document);
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

void AssignParagraphIdentities(OpenXmlElement root)
{
    const string word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
    var next = 1;
    foreach (var element in root.Descendants().Where(element => element is Paragraph or TableRow))
    {
        element.SetAttribute(new OpenXmlAttribute("w14", "paraId", word2010, next.ToString("X8")));
        element.SetAttribute(new OpenXmlAttribute("w14", "textId", word2010, (next + 1000).ToString("X8")));
        next++;
    }
}

void RequireUniqueWordIdentities(string path)
{
    const string word2010 = "http://schemas.microsoft.com/office/word/2010/wordml";
    using var document = WordprocessingDocument.Open(path, false);
    var values = document.MainDocumentPart?.Document?.Descendants()
        .Select(element => element.GetAttributes().FirstOrDefault(attribute =>
            attribute.LocalName == "paraId" && attribute.NamespaceUri == word2010).Value)
        .Where(value => !string.IsNullOrWhiteSpace(value)).ToArray() ?? [];
    Require(values.Length == values.Distinct(StringComparer.Ordinal).Count(),
        "inserted objects retained duplicate Word identities");
}

static JsonElement Address(JsonElement value) => value.GetProperty("address").Clone();
static void Require(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
