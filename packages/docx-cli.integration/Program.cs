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
    var commented = Path.Combine(root, "commented.docx");
    var uncommented = Path.Combine(root, "uncommented.docx");
    var uncommentedReceipt = Path.Combine(root, "uncommented-receipt.json");
    CreateCommentedDocument(commented);
    var deleteComments = Run("docx_delete_comments", new
    {
        input = commented,
        output = uncommented,
        receiptOutput = uncommentedReceipt,
    });
    Require(deleteComments.GetProperty("summary").GetProperty("appliedCount").GetInt32() == 1,
        "delete comments did not report the removed comment");
    using (var uncommentedDocument = WordprocessingDocument.Open(uncommented, false))
    {
        var main = uncommentedDocument.MainDocumentPart!;
        Require(main.WordprocessingCommentsPart is null, "delete comments retained the comments part");
        Require(!main.Document.Descendants().Any(element =>
                element is CommentRangeStart or CommentRangeEnd or CommentReference),
            "delete comments retained a main-story comment marker");
        Require(main.Document.Body!.InnerText == "Visible content",
            "delete comments changed visible document content");
    }
    var noCommentOutput = Path.Combine(root, "uncommented-noop.docx");
    var noComment = Run("docx_delete_comments", new
    {
        input = uncommented,
        output = noCommentOutput,
        receiptOutput = Path.Combine(root, "uncommented-noop-receipt.json"),
    });
    Require(noComment.GetProperty("summary").GetProperty("appliedCount").GetInt32() == 0,
        "delete comments no-op reported a removed comment");
    using (var noCommentDocument = WordprocessingDocument.Open(noCommentOutput, false))
        Require(noCommentDocument.MainDocumentPart!.Document.Body!.InnerText == "Visible content",
            "delete comments no-op changed visible document content");
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
    Require(initial.GetProperty("tableWidth").GetProperty("type").GetString() == "dxa"
            && initial.GetProperty("tableWidth").GetProperty("value").GetString() == "6000",
        "table read did not expose the explicit native table width");

    var fontOutput = Path.Combine(root, "set-text-font.docx");
    var fontReceipt = Path.Combine(root, "set-text-font-receipt.json");
    Run("docx_set_text", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                target = rows[1].GetProperty("cells")[1].GetProperty("address").Clone(),
                text = "Latin 42\n下一行",
                fontName = "Unseen Latin Font",
            },
        },
        output = fontOutput,
        receiptOutput = fontReceipt,
    });
    using (var fontDocument = WordprocessingDocument.Open(fontOutput, false))
    {
        var fontRun = fontDocument.MainDocumentPart!.Document.Descendants<Run>()
            .Single(run => run.InnerText == "Latin 42下一行");
        Require(fontRun.RunProperties?.RunFonts?.Ascii?.Value == "Unseen Latin Font"
                && fontRun.RunProperties?.RunFonts?.HighAnsi?.Value == "Unseen Latin Font"
                && fontRun.RunProperties?.RunFonts?.ComplexScript?.Value == "Unseen Latin Font",
            "set text did not apply the selected Latin font family");
    }
    var multilineReadback = ReadTable(fontOutput, "set-text-multiline-output");
    Require(multilineReadback.GetProperty("rows")[1].GetProperty("cells")[1]
            .GetProperty("paragraphs")[0].GetProperty("text").GetString() == "Latin 42\n下一行",
        "table read did not preserve a native line break written by set text");
    var emptyFontReceipt = Path.Combine(root, "set-text-empty-font-receipt.json");
    var emptyFont = RunExpectAtomicFailure("docx_set_text", original, emptyFontReceipt, new
    {
        input = original,
        changes = new[] { new { target = rows[1].GetProperty("cells")[1].GetProperty("address").Clone(), text = "", fontName = "Unseen Latin Font" } },
        output = original,
        receiptOutput = emptyFontReceipt,
    });
    Require(emptyFont.Contains("font-name-requires-nonempty-text", StringComparison.Ordinal),
        "set text claimed a font postcondition for empty replacement text");

    var widthOutput = Path.Combine(root, "set-table-width.docx");
    var widthReceipt = Path.Combine(root, "set-table-width-receipt.json");
    Run("docx_set_table_width", new
    {
        input = original,
        changes = new[] { new { table = initial.GetProperty("address").Clone(), width = new { type = "pct", value = "4321" } } },
        output = widthOutput,
        receiptOutput = widthReceipt,
    });
    var widthReadback = ReadTable(widthOutput, "set-table-width-output");
    Require(widthReadback.GetProperty("tableWidth").GetProperty("type").GetString() == "pct"
            && widthReadback.GetProperty("tableWidth").GetProperty("value").GetString() == "4321",
        "set table width did not produce the requested native width");
    var autoWidthOutput = Path.Combine(root, "set-table-width-auto.docx");
    Run("docx_set_table_width", new
    {
        input = original,
        changes = new[] { new { table = initial.GetProperty("address").Clone(), width = new { type = "auto", value = "0" } } },
        output = autoWidthOutput,
        receiptOutput = Path.Combine(root, "set-table-width-auto-receipt.json"),
    });
    var autoWidthReadback = ReadTable(autoWidthOutput, "set-table-width-auto-output");
    Require(autoWidthReadback.GetProperty("tableWidth").GetProperty("type").GetString() == "auto"
            && autoWidthReadback.GetProperty("tableWidth").GetProperty("value").GetString() == "0",
        "set table width did not preserve the auto width variant");
    var invalidWidthReceipt = Path.Combine(root, "set-table-width-invalid-receipt.json");
    var invalidWidth = RunExpectAtomicFailure("docx_set_table_width", original, invalidWidthReceipt, new
    {
        input = original,
        changes = new[] { new { table = initial.GetProperty("address").Clone(), width = new { type = "auto", value = "1" } } },
        output = original,
        receiptOutput = invalidWidthReceipt,
    });
    Require(invalidWidth.Contains("table-width-invalid", StringComparison.Ordinal),
        "set table width accepted an unsupported width type");
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

    var paginationOutput = Path.Combine(root, "paragraph-pagination.docx");
    var paginationReceipt = Path.Combine(root, "paragraph-pagination-receipt.json");
    var captionParagraph = rows[0].GetProperty("cells")[0].GetProperty("paragraphs")[0]
        .GetProperty("address").Clone();
    var bodyParagraph = rows[1].GetProperty("cells")[1].GetProperty("paragraphs")[0]
        .GetProperty("address").Clone();
    var pagination = Run("docx_set_paragraph_pagination", new
    {
        input = original,
        changes = new object[]
        {
            new { paragraph = captionParagraph, keepWithNext = true, keepLinesTogether = true },
            new { paragraph = bodyParagraph, pageBreakBefore = false, preventWidowOrphanLines = false },
        },
        output = paginationOutput,
        receiptOutput = paginationReceipt,
    });
    Require(pagination.GetProperty("summary").GetProperty("appliedCount").GetInt32() == 2,
        "paragraph pagination did not apply every requested paragraph");
    using (var paginationDocument = WordprocessingDocument.Open(paginationOutput, false))
    {
        var paragraphs = paginationDocument.MainDocumentPart!.Document.Body!.Descendants<Paragraph>().ToArray();
        Require(paragraphs[0].ParagraphProperties?.KeepNext?.Val?.Value == true
                && paragraphs[0].ParagraphProperties?.KeepLines?.Val?.Value == true,
            "paragraph pagination did not set positive properties");
        var explicitFalseParagraph = paragraphs.Single(paragraph => paragraph.InnerText == "甲一");
        Require(explicitFalseParagraph.ParagraphProperties?.PageBreakBefore?.Val?.Value == false
                && explicitFalseParagraph.ParagraphProperties?.WidowControl?.Val?.Value == false,
            "paragraph pagination did not preserve explicit false properties");
    }
    var missingPaginationReceipt = Path.Combine(root, "paragraph-pagination-missing-receipt.json");
    var missingPagination = RunExpectAtomicFailure(
        "docx_set_paragraph_pagination", original, missingPaginationReceipt, new
        {
            input = original,
            changes = new[] { new { paragraph = captionParagraph } },
            output = original,
            receiptOutput = missingPaginationReceipt,
        });
    Require(missingPagination.Contains("paragraph-pagination-change-must-set-a-property", StringComparison.Ordinal),
        "paragraph pagination accepted a change without properties");

    RunVerticalTextAlignmentObservation();
    RunNativeInlineSelectionComposition();
    RunRichTargetContentReplacement();
    RunTextNodeSetText();
    RunBookmarkedParagraphInsertion();
    RunLegacyQualifiedTableLookInsertion();

    var replacementSource = Path.Combine(root, "content-replacement-source.docx");
    var replacementTarget = Path.Combine(root, "content-replacement-target.docx");
    CreateContentReplacementSourceDocument(replacementSource);
    CreateContentReplacementTargetDocument(replacementTarget);
    var replacementSourceState = ReadTable(replacementSource, "content-replacement-source");
    var replacementTargetState = ReadTable(replacementTarget, "content-replacement-target");
    var replacementSourceRows = replacementSourceState.GetProperty("rows");
    var replacementTargetRows = replacementTargetState.GetProperty("rows");
    var chineseItem = replacementSourceRows[1].GetProperty("cells")[0]
        .GetProperty("paragraphs")[0].GetProperty("address").Clone();
    var chineseResult = replacementSourceRows[1].GetProperty("cells")[1]
        .GetProperty("paragraphs")[0].GetProperty("address").Clone();
    var invalidSourceAddress = new
    {
        part = chineseResult.GetProperty("part").GetString(),
        path = replacementSourceRows[1].GetProperty("cells")[1].GetProperty("address")
            .GetProperty("path").GetString() + "/w:p[99]",
    };
    var invalidReplacementReceipt = Path.Combine(root, "content-replacement-invalid-receipt.json");
    var invalidReplacement = RunExpectAtomicFailure(
        "docx_replace_content_from_source", replacementTarget, invalidReplacementReceipt, new
        {
            input = replacementTarget,
            changes = new object[]
            {
                new
                {
                    target = CellAt(replacementTargetState, 1, 1).GetProperty("address").Clone(),
                    sourceInput = replacementSource,
                    sourceSelections = new[] { new { address = chineseResult } },
                },
                new
                {
                    target = CellAt(replacementTargetState, 1, 0).GetProperty("address").Clone(),
                    sourceInput = replacementSource,
                    sourceSelections = new[] { new { address = invalidSourceAddress } },
                },
            },
            output = replacementTarget,
            receiptOutput = invalidReplacementReceipt,
        });
    Require(invalidReplacement.Contains("object-address-not-found", StringComparison.Ordinal),
        "invalid source selection did not fail before content replacement");

    var replacementReceipt = Path.Combine(root, "content-replacement-receipt.json");
    Run("docx_replace_content_from_source", new
    {
        input = replacementTarget,
        changes = new object[]
        {
            new
            {
                target = CellAt(replacementTargetState, 1, 0).GetProperty("address").Clone(),
                sourceInput = replacementSource,
                sourceSelections = new[] { new { address = chineseItem } },
            },
            new
            {
                target = CellAt(replacementTargetState, 1, 1).GetProperty("address").Clone(),
                sourceInput = replacementSource,
                sourceSelections = new[] { new { address = chineseResult } },
            },
        },
        output = replacementTarget,
        receiptOutput = replacementReceipt,
    });
    var replacedContent = ReadTable(replacementTarget, "content-replacement-output");
    Require(CellAt(replacedContent, 1, 0).GetProperty("logicalText").GetString() == "中文项目"
            && CellAt(replacedContent, 1, 1).GetProperty("logicalText").GetString() == "中文结果",
        "selected source paragraphs were not retained as target content");
    Require(CellAt(replacedContent, 1, 0).GetProperty("verticalMerge").GetString() == "restart"
            && CellAt(replacedContent, 2, 0).GetProperty("verticalMerge").GetString() == "continue"
            && CellAt(replacedContent, 2, 0).GetProperty("logicalText").GetString() == "中文项目",
        "content replacement changed the target vertical merge");
    RunInput("validate-openxml", replacementTarget);

    var atomicTableTarget = Path.Combine(root, "set-table-native-content-target.docx");
    var atomicTableOutput = Path.Combine(root, "set-table-native-content-output.docx");
    CreateContentReplacementTargetDocument(atomicTableTarget);
    var atomicTargetState = ReadTable(atomicTableTarget, "set-table-native-content-target");
    var atomicTargetRows = atomicTargetState.GetProperty("rows");
    var atomicTargetColumns = atomicTargetState.GetProperty("gridColumns");
    Run("docx_set_table", new
    {
        input = atomicTableTarget,
        table = atomicTargetState.GetProperty("address").Clone(),
        existingRows = new
        {
            first = atomicTargetRows[1].GetProperty("address").Clone(),
            last = atomicTargetRows[2].GetProperty("address").Clone(),
        },
        columns = atomicTargetColumns.EnumerateArray()
            .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
            .ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = atomicTargetRows[1].GetProperty("address").Clone(),
                cantSplit = true,
                cells = new object[]
                {
                    new
                    {
                        columns = new[] { "c0" }, rowSpan = (int?)null, text = (string?)null,
                        sourceInput = replacementSource,
                        sourceSelections = new[] { new { address = chineseItem } },
                    },
                    new
                    {
                        columns = new[] { "c1" }, rowSpan = (int?)null, text = (string?)null,
                        sourceInput = replacementSource,
                        sourceSelections = new[] { new { address = chineseResult } },
                    },
                },
            },
        },
        output = atomicTableOutput,
        receiptOutput = Path.Combine(root, "set-table-native-content-receipt.json"),
    });
    var atomicTableState = ReadTable(atomicTableOutput, "set-table-native-content-output");
    Require(atomicTableState.GetProperty("rowCount").GetInt32() == 2
            && CellAt(atomicTableState, 1, 0).GetProperty("logicalText").GetString() == "中文项目"
            && CellAt(atomicTableState, 1, 1).GetProperty("logicalText").GetString() == "中文结果"
            && atomicTableState.GetProperty("rows")[1].GetProperty("cantSplit").GetBoolean(),
        "docx_set_table did not atomically apply shape and native source content");
    RunInput("validate-openxml", atomicTableOutput);

    var richRunTableOutput = Path.Combine(root, "set-table-rich-runs-output.docx");
    Run("docx_set_table", new
    {
        input = atomicTableTarget,
        table = atomicTargetState.GetProperty("address").Clone(),
        existingRows = new
        {
            first = atomicTargetRows[1].GetProperty("address").Clone(),
            last = atomicTargetRows[2].GetProperty("address").Clone(),
        },
        columns = atomicTargetColumns.EnumerateArray()
            .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
            .ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = atomicTargetRows[1].GetProperty("address").Clone(),
                cells = new object[]
                {
                    new
                    {
                        columns = new[] { "c0" }, rowSpan = (int?)null,
                        textRuns = new object[]
                        {
                            new { text = "Ω-42", color = "FF4A7BC8", underline = (object)"double" },
                            new { text = " / control", color = (string?)null, underline = (object)false },
                        },
                    },
                    new
                    {
                        columns = new[] { "c1" }, rowSpan = (int?)null,
                        textRuns = new[] { new { text = "single-run", color = (string?)null, underline = "single" } },
                    },
                },
            },
        },
        output = richRunTableOutput,
        receiptOutput = Path.Combine(root, "set-table-rich-runs-receipt.json"),
    });
    using (var richRunDocument = WordprocessingDocument.Open(richRunTableOutput, false))
    {
        var richCells = richRunDocument.MainDocumentPart!.Document.Descendants<Table>()
            .Single().Elements<TableRow>().ElementAt(1).Elements<TableCell>().ToArray();
        var firstRuns = richCells[0].Descendants<Run>().Where(run => run.InnerText.Length > 0).ToArray();
        Require(firstRuns.Length == 2
                && firstRuns[0].InnerText == "Ω-42"
                && firstRuns[0].RunProperties?.Color?.Val?.Value == "4A7BC8"
                && firstRuns[0].RunProperties?.Underline?.Val?.Value == UnderlineValues.Double
                && firstRuns[1].InnerText == " / control"
                && firstRuns[1].RunProperties?.Color is null
                && firstRuns[1].RunProperties?.Underline is null,
            "docx_set_table did not preserve the exact ordered rich text run evidence");
        var boundaryRun = richCells[1].Descendants<Run>().Single(run => run.InnerText.Length > 0);
        Require(boundaryRun.InnerText == "single-run"
                && boundaryRun.RunProperties?.Color is null
                && boundaryRun.RunProperties?.Underline?.Val?.Value == UnderlineValues.Single,
            "docx_set_table did not support a one-run rich text boundary");
    }
    RunInput("validate-openxml", richRunTableOutput);

    var invalidRichRunReceipt = Path.Combine(root, "set-table-rich-runs-invalid-receipt.json");
    var invalidRichRun = RunExpectAtomicFailure("docx_set_table", atomicTableTarget, invalidRichRunReceipt, new
    {
        input = atomicTableTarget,
        table = atomicTargetState.GetProperty("address").Clone(),
        existingRows = new
        {
            first = atomicTargetRows[1].GetProperty("address").Clone(),
            last = atomicTargetRows[2].GetProperty("address").Clone(),
        },
        columns = atomicTargetColumns.EnumerateArray()
            .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
            .ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = atomicTargetRows[1].GetProperty("address").Clone(),
                cells = new object[]
                {
                    new { columns = new[] { "c0" }, text = "conflicting", textRuns = new[] { new { text = "run", color = "123456" } } },
                    new { columns = new[] { "c1" }, text = "unchanged" },
                },
            },
        },
        output = atomicTableTarget,
        receiptOutput = invalidRichRunReceipt,
    });
    Require(invalidRichRun.Contains("requires-exactly-one-content-mode", StringComparison.Ordinal),
        "docx_set_table accepted both plain text and rich text runs for one cell");
    var emptyRichRunReceipt = Path.Combine(root, "set-table-rich-runs-empty-receipt.json");
    var emptyRichRun = RunExpectAtomicFailure("docx_set_table", atomicTableTarget, emptyRichRunReceipt, new
    {
        input = atomicTableTarget,
        table = atomicTargetState.GetProperty("address").Clone(),
        existingRows = new { first = atomicTargetRows[1].GetProperty("address").Clone(), last = atomicTargetRows[2].GetProperty("address").Clone() },
        columns = atomicTargetColumns.EnumerateArray().Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() }).ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = atomicTargetRows[1].GetProperty("address").Clone(),
                cells = new object[]
                {
                    new { columns = new[] { "c0" }, textRuns = Array.Empty<object>() },
                    new { columns = new[] { "c1" }, text = "unchanged" },
                },
            },
        },
        output = atomicTableTarget,
        receiptOutput = emptyRichRunReceipt,
    });
    Require(emptyRichRun.Contains("text-runs-empty", StringComparison.Ordinal),
        "docx_set_table accepted a rich text cell after its final run was deleted");
    var invalidRichColorReceipt = Path.Combine(root, "set-table-rich-runs-color-receipt.json");
    var invalidRichColor = RunExpectAtomicFailure("docx_set_table", atomicTableTarget, invalidRichColorReceipt, new
    {
        input = atomicTableTarget,
        table = atomicTargetState.GetProperty("address").Clone(),
        existingRows = new { first = atomicTargetRows[1].GetProperty("address").Clone(), last = atomicTargetRows[2].GetProperty("address").Clone() },
        columns = atomicTargetColumns.EnumerateArray().Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() }).ToArray(),
        rows = new[]
        {
            new
            {
                prototypeRow = atomicTargetRows[1].GetProperty("address").Clone(),
                cells = new object[]
                {
                    new { columns = new[] { "c0" }, textRuns = new[] { new { text = "invalid", color = "not-rgb" } } },
                    new { columns = new[] { "c1" }, text = "unchanged" },
                },
            },
        },
        output = atomicTableTarget,
        receiptOutput = invalidRichColorReceipt,
    });
    Require(invalidRichColor.Contains("color-invalid", StringComparison.Ordinal),
        "docx_set_table accepted an invalid rich text color");

    var validAddress = rows[1].GetProperty("address");
    var malformedAddress = new
    {
        part = validAddress.GetProperty("part").GetString(),
        path = validAddress.GetProperty("path").GetString() + "}]"
    };
    var malformedAddressError = RunExpectFailure("docx_read_object", new
    {
        input = original,
        addresses = new[] { malformedAddress },
        kinds = new[] { "paragraph" }
    });
    Require(malformedAddressError.Contains("addresses[0].path-is-invalid", StringComparison.Ordinal),
        "malformed native path did not fail at the address contract");

    var nestedOutput = Path.Combine(root, "new-artifacts", "documents", "result.docx");
    var nestedReceipt = Path.Combine(root, "new-artifacts", "receipts", "result.json");
    Run("docx_set_text", new
    {
        input = original,
        changes = new[]
        {
            new
            {
                target = rows[1].GetProperty("cells")[1].GetProperty("address").Clone(),
                text = "nested output parent"
            }
        },
        output = nestedOutput,
        receiptOutput = nestedReceipt
    });
    Require(File.Exists(nestedOutput) && File.Exists(nestedReceipt),
        "mutation did not create requested output parent directories");

    var setBodyOutput = Path.Combine(root, "set-table-body.docx");
    var setBodyReceipt = Path.Combine(root, "set-table-body-receipt.json");
    var bodyColumns = initial.GetProperty("gridColumns").EnumerateArray()
        .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
        .ToArray();
    Run("docx_set_table", new
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
                cantSplit = false,
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "精密度", rowSpan = (int?)2 },
                    new { columns = new[] { "c1" }, text = "标准一", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "甲二", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "通过", rowSpan = (int?)2 },
                }
            },
            new
            {
                prototypeRow = rows[3].GetProperty("address").Clone(),
                cantSplit = true,
                cells = new[]
                {
                    new { columns = new[] { "c1" }, text = "标准二", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "结果二", rowSpan = (int?)null },
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
            && CellAt(setBodyState, 1, 2).GetProperty("logicalText").GetString() == "甲二"
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
    Require(!setRows[1].GetProperty("cantSplit").GetBoolean()
            && setRows[2].GetProperty("cantSplit").GetBoolean(),
        "set table body did not apply explicit row pagination properties");
    using (var preservedFormatting = WordprocessingDocument.Open(setBodyOutput, false))
    {
        var matchingParagraph = preservedFormatting.MainDocumentPart!.Document
            .Descendants<Paragraph>().Single(paragraph => paragraph.InnerText == "甲二");
        Require(matchingParagraph.Elements<Run>().Any(run =>
                run.InnerText == "二"
                && run.RunProperties?.VerticalTextAlignment?.Val?.Value
                    == VerticalPositionValues.Superscript),
            "set table body discarded existing rich text when cell text was unchanged");
    }

    var expandedBodyOutput = Path.Combine(root, "set-table-body-expanded.docx");
    Run("docx_set_table", new
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
    Run("docx_set_table", new
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
    var emptyTable = RunExpectAtomicFailure("docx_set_table", original, emptyTableReceipt, new
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
    var failedBody = RunExpectAtomicFailure("docx_set_table", setBodyOutput, failedBodyReceipt, new
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

    var invalidRowSpanReceipt = Path.Combine(root, "set-table-body-invalid-row-span-receipt.json");
    var invalidRowSpan = RunExpectAtomicFailure("docx_set_table", setBodyOutput, invalidRowSpanReceipt, new
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
                    new { columns = new[] { "c0" }, text = "项目一", rowSpan = (int?)3 },
                    new { columns = new[] { "c1" }, text = "标准一", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "结果一", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "结论一", rowSpan = (int?)null },
                }
            },
            new
            {
                prototypeRow = setRows[2].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "项目二", rowSpan = (int?)null },
                    new { columns = new[] { "c1" }, text = "标准二", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "结果二", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "结论二", rowSpan = (int?)null },
                }
            }
        },
        output = setBodyOutput,
        receiptOutput = invalidRowSpanReceipt,
    });
    Require(invalidRowSpan.Contains("rowSpan-invalid", StringComparison.Ordinal),
        "row span beyond the final row did not fail explicitly");

    var overlappingRowSpanReceipt = Path.Combine(root, "set-table-body-overlapping-row-span-receipt.json");
    var overlappingRowSpan = RunExpectAtomicFailure("docx_set_table", setBodyOutput, overlappingRowSpanReceipt, new
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
                    new { columns = new[] { "c0" }, text = "项目一", rowSpan = (int?)2 },
                    new { columns = new[] { "c1" }, text = "标准一", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "结果一", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "结论一", rowSpan = (int?)null },
                }
            },
            new
            {
                prototypeRow = setRows[2].GetProperty("address").Clone(),
                cells = new[]
                {
                    new { columns = new[] { "c0" }, text = "另一个项目", rowSpan = (int?)null },
                    new { columns = new[] { "c1" }, text = "标准二", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "结果二", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "结论二", rowSpan = (int?)null },
                }
            }
        },
        output = setBodyOutput,
        receiptOutput = overlappingRowSpanReceipt,
    });
    Require(overlappingRowSpan.Contains("columns-overlap", StringComparison.Ordinal),
        "a cell overlapping an active row span did not fail explicitly");

    var splitBodyReceipt = Path.Combine(root, "set-table-body-split-merge-receipt.json");
    var splitBody = RunExpectAtomicFailure("docx_set_table", original, splitBodyReceipt, new
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

    var narrowSource = Path.Combine(root, "narrow-source.docx");
    CreateNarrowDocument(narrowSource);
    var narrowState = ReadTable(narrowSource, "narrow-source");
    var incompatibleGridReceipt = Path.Combine(root, "incompatible-grid-insert-receipt.json");
    var incompatibleGridInsert = RunExpectAtomicFailure("docx_insert_objects", deletedColumns, incompatibleGridReceipt, new
    {
        input = deletedColumns,
        changes = new[]
        {
            new
            {
                sourceInput = narrowSource,
                sources = new[] { narrowState.GetProperty("rows")[0].GetProperty("address").Clone() },
                targetParent = afterDelete.GetProperty("address").Clone()
            }
        },
        output = Path.Combine(root, "incompatible-grid-insert.docx"),
        receiptOutput = incompatibleGridReceipt
    });
    Require(incompatibleGridInsert.Contains("row-source-table-grid-incompatible-with-target", StringComparison.Ordinal),
        "row insertion accepted a source row from an incompatible table grid");

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

    RunTocStylePolicyMatrix();
    RunMergedHeaderSetBodyMatrix();
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

void RunTocStylePolicyMatrix()
{
    foreach (var (italic, indentCharacters) in new[] { (false, 2), (true, 1) })
    {
        var input = Path.Combine(root, $"toc-policy-{italic}-input.docx");
        var output = Path.Combine(root, $"toc-policy-{italic}-output.docx");
        CreateTocPolicyDocument(input);
        Run("docx_apply_toc_style_policy", new
        {
            input,
            italic,
            indentCharactersPerLevel = indentCharacters,
            output,
            receiptOutput = Path.Combine(root, $"toc-policy-{italic}-receipt.json")
        });
        using var document = WordprocessingDocument.Open(output, false);
        var entries = document.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
            .Where(paragraph => paragraph.Descendants<Text>().Any(text => text.Text.StartsWith("Entry", StringComparison.Ordinal)))
            .ToList();
        Require(entries.Count == 2, "TOC policy did not retain both result entries");
        for (var index = 0; index < entries.Count; index++)
        {
            var expectedIndent = index * 2 * indentCharacters * 100;
            Require(entries[index].ParagraphProperties?.GetFirstChild<Indentation>()?.LeftChars?.Value == expectedIndent,
                "TOC policy did not apply the actual entry indentation");
            Require(entries[index].ParagraphProperties?.Tabs?.Elements<TabStop>()
                    .Any(tab => tab.Val?.Value == TabStopValues.Right && tab.Leader?.Value == TabStopLeaderCharValues.Dot) == true,
                "TOC policy removed the template tab leader");
            Require(entries[index].Descendants<Run>().All(run =>
                    (run.RunProperties?.Italic?.Val?.Value ?? true) == italic
                    && (run.RunProperties?.ItalicComplexScript?.Val?.Value ?? true) == italic),
                "TOC policy did not apply the actual entry italic setting");
        }
        Require(entries[1].ParagraphProperties?.ParagraphStyleId?.Val?.Value == "7",
            "TOC policy replaced an undefined template style reference");
        RunInput("validate-openxml", output);
    }
}

void RunMergedHeaderSetBodyMatrix()
{
    var input = Path.Combine(root, "set-body-merged-header-input.docx");
    var output = Path.Combine(root, "set-body-merged-header-output.docx");
    CreateMergedHeaderSetBodyDocument(input);
    var before = ReadTable(input, "set-body-merged-header-before");
    var rows = before.GetProperty("rows");
    var columns = before.GetProperty("gridColumns").EnumerateArray()
        .Select((column, index) => new { id = "c" + index, gridColumn = column.GetProperty("address").Clone() })
        .ToArray();

    Run("docx_set_table", new
    {
        input,
        table = before.GetProperty("address").Clone(),
        existingRows = new
        {
            first = rows[2].GetProperty("address").Clone(),
            last = rows[3].GetProperty("address").Clone(),
        },
        columns,
        rows = new object[]
        {
            new
            {
                prototypeRow = rows[2].GetProperty("address").Clone(),
                cells = new object[]
                {
                    new { columns = new[] { "c0" }, text = "group", rowSpan = (int?)2 },
                    new { columns = new[] { "c1" }, text = "item-a", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "method-a", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "v3-a", rowSpan = (int?)null },
                    new { columns = new[] { "c4" }, text = "v4-a", rowSpan = (int?)null },
                    new { columns = new[] { "c5" }, text = "v5-a", rowSpan = (int?)null },
                    new { columns = new[] { "c6" }, text = "v6-a", rowSpan = (int?)null },
                    new { columns = new[] { "c7" }, text = "v7-a", rowSpan = (int?)null },
                    new { columns = new[] { "c8" }, text = "v8-a", rowSpan = (int?)null },
                }
            },
            new
            {
                prototypeRow = rows[3].GetProperty("address").Clone(),
                cells = new object[]
                {
                    new { columns = new[] { "c1" }, text = "item-b", rowSpan = (int?)null },
                    new { columns = new[] { "c2" }, text = "method-b", rowSpan = (int?)null },
                    new { columns = new[] { "c3" }, text = "v3-b", rowSpan = (int?)null },
                    new { columns = new[] { "c4" }, text = "v4-b", rowSpan = (int?)null },
                    new { columns = new[] { "c5" }, text = "v5-b", rowSpan = (int?)null },
                    new { columns = new[] { "c6" }, text = "v6-b", rowSpan = (int?)null },
                    new { columns = new[] { "c7" }, text = "v7-b", rowSpan = (int?)null },
                    new { columns = new[] { "c8" }, text = "v8-b", rowSpan = (int?)null },
                }
            }
        },
        output,
        receiptOutput = Path.Combine(root, "set-body-merged-header-receipt.json")
    });

    var after = ReadTable(output, "set-body-merged-header-after");
    Require(after.GetProperty("columnCount").GetInt32() == 9,
        "set table body changed the nine-column grid under a merged header");
    Require(CellAt(after, 0, 0).GetProperty("gridSpan").GetInt32() == 2
            && CellAt(after, 0, 3).GetProperty("gridSpan").GetInt32() == 6,
        "set table body changed the merged header geometry");
    Require(CellAt(after, 2, 0).GetProperty("logicalText").GetString() == "group"
            && CellAt(after, 2, 1).GetProperty("logicalText").GetString() == "item-a"
            && CellAt(after, 2, 2).GetProperty("logicalText").GetString() == "method-a"
            && CellAt(after, 2, 8).GetProperty("logicalText").GetString() == "v8-a",
        "set table body shifted a value under the merged header");
    Require(CellAt(after, 3, 0).GetProperty("verticalMerge").GetString() == "continue"
            && CellAt(after, 3, 0).GetProperty("logicalText").GetString() == "group"
            && CellAt(after, 3, 1).GetProperty("logicalText").GetString() == "item-b"
            && CellAt(after, 3, 2).GetProperty("logicalText").GetString() == "method-b"
            && CellAt(after, 3, 8).GetProperty("logicalText").GetString() == "v8-b",
        "set table body shifted the second row or lost its vertical merge owner");
    RequireTableInvariants(after, "set-body-merged-header");
    RunInput("validate-openxml", output);
    Console.WriteLine("PASS set table body under merged two-row header");
}

void RunVerticalTextAlignmentObservation()
{
    var input = Path.Combine(root, "vertical-text-observation.docx");
    CreateVerticalTextObservationDocument(input);
    var table = ReadTable(input, "vertical-text-observation-table");
    var paragraph = table.GetProperty("rows")[0].GetProperty("cells")[0]
        .GetProperty("paragraphs")[0];
    var textNodes = paragraph.GetProperty("textNodes");
    Require(textNodes[1].GetProperty("verticalTextAlignment").GetString() == "superscript"
            && textNodes[3].GetProperty("verticalTextAlignment").GetString() == "subscript",
        "table detail did not expose native superscript and subscript text");

    var observation = Run("docx_read_object", new
    {
        input,
        addresses = new[] { paragraph.GetProperty("address").Clone() },
        kinds = new[] { "run", "text" },
        output = Path.Combine(root, "vertical-text-observation-object.json")
    }).GetProperty("observations")[0];
    var verticalRuns = ObservationObjects(observation)
        .Where(item => item.GetProperty("kind").GetString() == "run"
            && item.GetProperty("verticalTextAlignment").ValueKind == JsonValueKind.String)
        .Select(item => item.GetProperty("verticalTextAlignment").GetString())
        .ToArray();
    Require(verticalRuns.SequenceEqual(new[] { "superscript", "subscript" }),
        "narrow object read did not expose native superscript and subscript runs");
    Console.WriteLine("PASS vertical text alignment observation");
}

void RunNativeInlineSelectionComposition()
{
    var source = Path.Combine(root, "native-inline-selection-source.docx");
    var target = Path.Combine(root, "native-inline-selection-target.docx");
    CreateNativeInlineSelectionSourceDocument(source);
    CreateNativeInlineSelectionTargetDocument(target);

    var sourceTable = ReadTable(source, "native-inline-selection-source");
    var targetTable = ReadTable(target, "native-inline-selection-target");
    var sourceCell = CellAt(sourceTable, 0, 0);
    var sourceBreakCell = CellAt(sourceTable, 1, 0);
    var sourceParagraphs = sourceCell.GetProperty("paragraphs");
    var sourceObjects = Run("docx_read_object", new
    {
        input = source,
        addresses = new[] { sourceCell.GetProperty("address").Clone() },
        kinds = new[] { "run", "text" },
        output = Path.Combine(root, "native-inline-selection-objects.json")
    }).GetProperty("observations")[0];
    var objects = ObservationObjects(sourceObjects).ToArray();
    const string expectedInline = "alpha xy middle pq omega";
    var firstParagraphPath = sourceParagraphs[0].GetProperty("address").GetProperty("path").GetString();
    var runAddresses = objects
        .Where(item => item.GetProperty("kind").GetString() == "run"
            && item.GetProperty("address").GetProperty("path").GetString()!.StartsWith(firstParagraphPath!, StringComparison.Ordinal))
        .Select(item => item.GetProperty("address").Clone())
        .ToArray();
    var textAddresses = objects
        .Where(item => item.GetProperty("kind").GetString() == "text"
            && item.GetProperty("address").GetProperty("path").GetString()!.StartsWith(firstParagraphPath!, StringComparison.Ordinal))
        .Select(item => item.GetProperty("address").Clone())
        .ToArray();

    Run("docx_replace_content_from_source", new
    {
        input = target,
        changes = new object[]
        {
            new
            {
                target = CellAt(targetTable, 0, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = runAddresses.Select(address => new { address }).ToArray(),
            },
            new
            {
                target = CellAt(targetTable, 1, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = textAddresses.Select(address => new { address }).ToArray(),
            },
            new
            {
                target = CellAt(targetTable, 2, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = new[] { new { address = textAddresses[1], range = new { start = 1, length = 1 } } },
            },
            new
            {
                target = CellAt(targetTable, 3, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = sourceParagraphs.EnumerateArray()
                    .Select(paragraph => new { address = paragraph.GetProperty("address").Clone() }).ToArray(),
            },
            new
            {
                target = CellAt(targetTable, 4, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = new[]
                {
                    new
                    {
                        address = sourceCell.GetProperty("address").Clone(),
                        range = new { start = 0, length = expectedInline.Length },
                    },
                },
            },
            new
            {
                target = CellAt(targetTable, 5, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = new[]
                {
                    new
                    {
                        address = sourceCell.GetProperty("address").Clone(),
                        range = new { start = 0, length = 0 },
                    },
                },
            },
            new
            {
                target = CellAt(targetTable, 6, 0).GetProperty("address").Clone(),
                sourceInput = source,
                sourceSelections = new[]
                {
                    new
                    {
                        address = sourceBreakCell.GetProperty("address").Clone(),
                        range = new { start = 0, length = "before\nafter".Length },
                    },
                },
            },
        },
        output = target,
        receiptOutput = Path.Combine(root, "native-inline-selection-receipt.json")
    });

    var result = ReadTable(target, "native-inline-selection-result");
    foreach (var rowIndex in new[] { 0, 1, 4 })
    {
        var cell = CellAt(result, rowIndex, 0);
        var paragraphs = cell.GetProperty("paragraphs");
        var textNodes = paragraphs[0].GetProperty("textNodes");
        Require(cell.GetProperty("logicalText").GetString() == expectedInline
                && paragraphs.GetArrayLength() == 1,
            $"inline selections did not remain in one source paragraph for target row {rowIndex}");
        Require(textNodes[1].GetProperty("verticalTextAlignment").GetString() == "superscript"
                && textNodes[3].GetProperty("verticalTextAlignment").GetString() == "subscript",
            $"inline selections lost vertical text alignment for target row {rowIndex}");
    }
    var rangeCell = CellAt(result, 2, 0);
    Require(rangeCell.GetProperty("logicalText").GetString() == "y"
            && rangeCell.GetProperty("paragraphs")[0].GetProperty("textNodes")[0]
                .GetProperty("verticalTextAlignment").GetString() == "superscript",
        "text range did not inherit its source run formatting");
    var multiParagraphCell = CellAt(result, 3, 0);
    Require(multiParagraphCell.GetProperty("paragraphs").GetArrayLength() == 2
            && multiParagraphCell.GetProperty("paragraphs")[0].GetProperty("text").GetString() == expectedInline
            && multiParagraphCell.GetProperty("paragraphs")[1].GetProperty("text").GetString() == "second paragraph",
        "selections from distinct source paragraphs lost their paragraph boundary");
    Require(CellAt(result, 5, 0).GetProperty("logicalText").GetString() == string.Empty,
        "zero-length native selection did not clear the target");
    Require(CellAt(result, 6, 0).GetProperty("logicalText").GetString() == "before\nafter",
        "cell range did not use the published line-break text offsets");
    RunInput("validate-openxml", target);
    Console.WriteLine("PASS native inline selection composition");
}

IEnumerable<JsonElement> ObservationObjects(JsonElement node)
{
    if (node.TryGetProperty("object", out var item)) yield return item;
    if (!node.TryGetProperty("children", out var children)) yield break;
    foreach (var child in children.EnumerateArray())
    foreach (var descendant in ObservationObjects(child))
        yield return descendant;
}

void CreateTocPolicyDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
    stylesPart.Styles = new Styles(
        HeadingStyle("Heading1", 0),
        HeadingStyle("Heading3", 2),
        new Style(new StyleName { Val = "Template TOC top" })
        {
            Type = StyleValues.Paragraph,
            StyleId = "TemplateTocTop",
            CustomStyle = true
        });
    stylesPart.Styles.Save();
    main.Document = new Document(new Body(
        Heading("Heading1", "_TocPolicyOne", "Heading one", "1"),
        Heading("Heading3", "_TocPolicyTwo", "Heading two", "2"),
        TocEntry("TemplateTocTop", "_TocPolicyOne", "Entry one", true),
        TocEntry("7", "_TocPolicyTwo", "Entry two", true)));
    main.Document.Save();
}

void RunBookmarkedParagraphInsertion()
{
    var input = Path.Combine(root, "bookmarked-paragraph-insertion.docx");
    using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(HeadingStyle("Heading1", 0));
        stylesPart.Styles.Save();
        main.Document = new Document(new Body(
            Heading("Heading1", "_TocInsertionSource", "Reusable heading", "14"),
            new Paragraph(new Run(new Text("Insertion boundary")))));
        AssignParagraphIdentities(main.Document);
        main.Document.Save();
    }

    var output = Path.Combine(root, "bookmarked-paragraph-insertion-output.docx");
    Run("docx_insert_objects", new
    {
        input,
        changes = new[]
        {
            new
            {
                sourceInput = input,
                sources = new[] { new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[1]" } },
                targetParent = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]" },
                before = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[2]" },
            }
        },
        output,
        receiptOutput = Path.Combine(root, "bookmarked-paragraph-insertion-receipt.json")
    });

    using var result = WordprocessingDocument.Open(output, false);
    var body = result.MainDocumentPart?.Document?.Body
        ?? throw new InvalidOperationException("bookmarked paragraph output body missing");
    Require(body.Elements<Paragraph>().Count(paragraph => paragraph.InnerText == "Reusable heading") == 2,
        "bookmarked paragraph was not inserted");
    Require(body.Descendants<BookmarkStart>().Count() == 1
            && body.Descendants<BookmarkEnd>().Count() == 1,
        "inserted paragraph retained source bookmark identity");
    RequireUniqueWordIdentities(output);
    RunInput("validate-openxml", output);
    Console.WriteLine("PASS bookmarked paragraph insertion");
}

void RunLegacyQualifiedTableLookInsertion()
{
    const string wordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var source = Path.Combine(root, "legacy-table-look-source.docx");
    using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        var look = new TableLook { Val = "04A0" };
        foreach (var (name, value) in new[]
        {
            ("firstRow", "1"), ("lastRow", "0"), ("firstColumn", "1"),
            ("lastColumn", "0"), ("noHBand", "0"), ("noVBand", "1"),
        })
            look.SetAttribute(new OpenXmlAttribute("w", name, wordprocessingNamespace, value));
        main.Document = new Document(new Body(
            new Table(
                new TableProperties(look),
                new TableGrid(new GridColumn()),
                new TableRow(new TableCell(new Paragraph(new Run(new Text("legacy table"))))))));
        main.Document.Save();
    }

    var target = Path.Combine(root, "legacy-table-look-target.docx");
    using (var document = WordprocessingDocument.Create(target, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(new Run(new Text("boundary")))));
        main.Document.Save();
    }

    var output = Path.Combine(root, "legacy-table-look-output.docx");
    Run("docx_insert_objects", new
    {
        input = target,
        changes = new[]
        {
            new
            {
                sourceInput = source,
                sources = new[] { new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:tbl[1]" } },
                targetParent = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]" },
                before = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[1]" },
            }
        },
        output,
        receiptOutput = Path.Combine(root, "legacy-table-look-receipt.json")
    });

    using (var document = WordprocessingDocument.Open(output, false))
    {
        var look = document.MainDocumentPart!.Document.Body!.GetFirstChild<Table>()!
            .GetFirstChild<TableProperties>()!.GetFirstChild<TableLook>()!;
        Require(look.Val?.Value == "04A0", "table look value changed while normalizing copied markup");
        Require(!look.GetAttributes().Any(attribute => attribute.NamespaceUri == wordprocessingNamespace
                && new[] { "firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand" }
                    .Contains(attribute.LocalName, StringComparer.Ordinal)),
            "redundant qualified table look flags survived object insertion");
    }
    RunInput("validate-openxml", output);
    Console.WriteLine("PASS legacy qualified table look insertion");
}

void RunRichTargetContentReplacement()
{
    var source = Path.Combine(root, "rich-target-content-source.docx");
    using (var document = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text("plain replacement"))),
            new Paragraph(
                new Run(new Text("rich replacement ")),
                new Run(
                    new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                    new Text("2")))));
        AssignParagraphIdentities(main.Document);
        main.Document.Save();
    }

    var target = Path.Combine(root, "rich-target-content-target.docx");
    using (var document = WordprocessingDocument.Create(target, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        var hyperlink = main.AddHyperlinkRelationship(new Uri("https://example.com/old"), true);
        main.Document = new Document(new Body(
            new Paragraph(
                new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
                new Run(new RunProperties(new Color { Val = "336699" }), new Text("plain old content"))),
            new Paragraph(
                new ParagraphProperties(new Indentation { Left = "720" }),
                new BookmarkStart { Name = "OldBookmark", Id = "41" },
                new Run(new RunProperties(new Color { Val = "993366" }), new Text("rich old content ")),
                new Hyperlink(new Run(new Text("old link"))) { Id = hyperlink.Id },
                new SimpleField(new Run(new Text("7"))) { Instruction = " PAGE " },
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" NUMPAGES ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("9")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new BookmarkEnd { Id = "41" })));
        AssignParagraphIdentities(main.Document);
        main.Document.Save();
    }

    Run("docx_replace_content_from_source", new
    {
        input = target,
        changes = new[]
        {
            new
            {
                target = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[1]" },
                sourceInput = source,
                sourceSelections = new[]
                {
                    new { address = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[1]" } }
                }
            },
            new
            {
                target = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[2]" },
                sourceInput = source,
                sourceSelections = new[]
                {
                    new { address = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[2]" } }
                }
            }
        },
        output = target,
        receiptOutput = Path.Combine(root, "rich-target-content-receipt.json")
    });

    using var result = WordprocessingDocument.Open(target, false);
    var paragraphs = result.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToArray();
    Require(paragraphs.Length == 2
            && paragraphs[0].InnerText == "plain replacement"
            && paragraphs[1].InnerText == "rich replacement 2",
        "plain and rich target paragraphs did not receive exact source content");
    Require(paragraphs[0].ParagraphProperties?.Justification?.Val?.Value == JustificationValues.Center
            && paragraphs[1].ParagraphProperties?.Indentation?.Left?.Value == "720",
        "content replacement did not preserve target paragraph formatting");
    Require(paragraphs[0].Descendants<Run>().Single().RunProperties?.Color?.Val?.Value == "336699"
            && paragraphs[1].Descendants<Run>().First().RunProperties?.Color?.Val?.Value == "993366"
            && paragraphs[1].Descendants<Run>().Last().RunProperties?.VerticalTextAlignment?.Val?.Value
                == VerticalPositionValues.Superscript,
        "content replacement did not preserve target run formatting and selected source semantics");
    Require(!paragraphs.SelectMany(paragraph => paragraph.Descendants<OpenXmlElement>()).Any(element =>
            element is BookmarkStart or BookmarkEnd or Hyperlink or SimpleField or FieldChar or FieldCode),
        "content replacement retained non-plain target content");
    RunInput("validate-openxml", target);
    Console.WriteLine("PASS rich target content replacement");
}

void RunTextNodeSetText()
{
    var input = Path.Combine(root, "set-text-node-input.docx");
    using (var document = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
    {
        var main = document.AddMainDocumentPart();
        main.Document = new Document(new Body(new Paragraph(
            new Run(new Text("OLD")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" REF _Ref1 ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("Table 1")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }))));
        main.Document.Save();
    }
    var detail = Run("docx_read_object", new
    {
        input,
        addresses = new[] { new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[1]" } },
        kinds = new[] { "text" },
    });
    Require(detail.GetProperty("observations")[0].GetProperty("children")[0].GetProperty("object")
            .GetProperty("text").GetString() == "OLD",
        "read object did not expose exact text-node content");
    var output = Path.Combine(root, "set-text-node-output.docx");
    Run("docx_set_text", new
    {
        input,
        changes = new[] { new { target = new { part = "/word/document.xml", path = "/w:document[1]/w:body[1]/w:p[1]/w:r[1]/w:t[1]" }, text = "NEW" } },
        output,
        receiptOutput = Path.Combine(root, "set-text-node-receipt.json"),
    });
    using var result = WordprocessingDocument.Open(output, false);
    var paragraph = result.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
    Require(paragraph.InnerText == "NEW REF _Ref1 Table 1"
            && paragraph.Descendants<FieldChar>().Count() == 3
            && paragraph.Descendants<FieldCode>().Single().Text == " REF _Ref1 ",
        "text-node update did not preserve the surrounding field structure");
    RunInput("validate-openxml", output);
    Console.WriteLine("PASS text-node set text");
}

void CreateMergedHeaderSetBodyDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = new Table(
        new TableProperties(),
        new TableGrid(Enumerable.Range(0, 9).Select(_ => new GridColumn { Width = "1000" })),
        new TableRow(
            new TableRowProperties(new TableHeader()),
            Cell("items", span: 2, merge: MergedCellValues.Restart),
            Cell("method", merge: MergedCellValues.Restart),
            Cell("validation", span: 6)),
        new TableRow(
            new TableRowProperties(new TableHeader()),
            Cell("", span: 2, merge: MergedCellValues.Continue),
            Cell("", merge: MergedCellValues.Continue),
            Cell("v3"), Cell("v4"), Cell("v5"), Cell("v6"), Cell("v7"), Cell("v8")),
        new TableRow(Cell("placeholder"), Cell("placeholder"), Cell("placeholder"),
            Cell(""), Cell(""), Cell(""), Cell(""), Cell(""), Cell("")),
        new TableRow(Cell("placeholder"), Cell("placeholder"), Cell("placeholder"),
            Cell(""), Cell(""), Cell(""), Cell(""), Cell(""), Cell("")));
    main.Document = new Document(new Body(table));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

void CreateVerticalTextObservationDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var paragraph = new Paragraph(
        new Run(new Text("R")),
        new Run(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new Text("2")),
        new Run(new Text(" and H")),
        new Run(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }),
            new Text("2")),
        new Run(new Text("O")));
    main.Document = new Document(new Body(new Table(
        new TableGrid(new GridColumn { Width = "4000" }),
        new TableRow(new TableCell(paragraph)))));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

void CreateNativeInlineSelectionSourceDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var first = new Paragraph(
        new Run(new Text("alpha ") { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
            new Text("xy")),
        new Run(new Text(" middle ") { Space = SpaceProcessingModeValues.Preserve }),
        new Run(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript }),
            new Text("pq")),
        new Run(new Text(" omega") { Space = SpaceProcessingModeValues.Preserve }));
    var second = new Paragraph(new Run(new Text("second paragraph")));
    var withBreak = new Paragraph(new Run(new Text("before"), new Break(), new Text("after")));
    main.Document = new Document(new Body(new Table(
        new TableProperties(),
        new TableGrid(new GridColumn { Width = "5000" }),
        new TableRow(new TableCell(first, second)),
        new TableRow(new TableCell(withBreak)))));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

void CreateNativeInlineSelectionTargetDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    main.Document = new Document(new Body(new Table(
        new TableProperties(),
        new TableGrid(new GridColumn { Width = "5000" }),
        new TableRow(Cell("run target")),
        new TableRow(Cell("text target")),
        new TableRow(Cell("range target")),
        new TableRow(Cell("paragraph target")),
        new TableRow(Cell("cell range target")),
        new TableRow(Cell("empty range target")),
        new TableRow(Cell("line-break range target")))));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

Style HeadingStyle(string id, int level)
    => new(new StyleName { Val = id }, new StyleParagraphProperties(new OutlineLevel { Val = level }))
    {
        Type = StyleValues.Paragraph,
        StyleId = id
    };

Paragraph Heading(string style, string bookmark, string text, string id)
    => new(
        new ParagraphProperties(new ParagraphStyleId { Val = style }),
        new BookmarkStart { Name = bookmark, Id = id },
        new Run(new Text(text)),
        new BookmarkEnd { Id = id });

Paragraph TocEntry(string style, string bookmark, string text, bool italic)
    => new(
        new ParagraphProperties(
            new ParagraphStyleId { Val = style },
            new Tabs(new TabStop { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 9000 }),
            new Indentation { FirstLineChars = 200 }),
        new Run(new FieldCode($" HYPERLINK \\l {bookmark} ")),
        new Run(new RunProperties(new Italic { Val = italic }, new ItalicComplexScript { Val = italic }), new Text(text)),
        new Run(new TabChar()),
        new Run(new Text("1")));

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
        new TableProperties(new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "6000" }),
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
        Cell("甲", merge: MergedCellValues.Restart), Cell("甲一"), CellWithSuperscript("甲", "二"), Cell("甲三")));
    table.Append(new TableRow(
        Cell("", merge: MergedCellValues.Continue), Cell("乙一"), Cell("乙二"), Cell("乙三")));
    table.Append(new TableRow(Cell("独立"), Cell("丙一"), Cell("丙二"), Cell("丙三")));
    main.Document = new Document(new Body(table));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

void CreateCommentedDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    main.Document = new Document(new Body(new Paragraph(
        new CommentRangeStart { Id = "0" },
        new Run(new Text("Visible content")),
        new CommentRangeEnd { Id = "0" },
        new Run(new CommentReference { Id = "0" }))));
    var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
    commentsPart.Comments = new Comments(new Comment(
        new Paragraph(new Run(new Text("Template instruction XXXX"))))
    {
        Id = "0",
        Author = "Author",
    });
    commentsPart.Comments.Save();
    main.Document.Save();
}

void CreateContentReplacementSourceDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = new Table(
        new TableProperties(),
        new TableGrid(new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }),
        new TableRow(Cell("项目"), Cell("结果")),
        new TableRow(
            CellWithParagraphs("中文项目", "English item"),
            CellWithParagraphs("中文结果", "English result")));
    main.Document = new Document(new Body(table));
    AssignParagraphIdentities(main.Document);
    main.Document.Save();
}

void CreateContentReplacementTargetDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = new Table(
        new TableProperties(),
        new TableGrid(new GridColumn { Width = "2000" }, new GridColumn { Width = "2000" }),
        new TableRow(Cell("项目"), Cell("结果")),
        new TableRow(CellWithSimpleField("占位", merge: MergedCellValues.Restart), Cell("占位结果")),
        new TableRow(Cell("", merge: MergedCellValues.Continue), Cell("保留")));
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

void CreateNarrowDocument(string path)
{
    using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
    var main = document.AddMainDocumentPart();
    var table = new Table(
        new TableProperties(),
        new TableGrid(
            new GridColumn { Width = "1000" }, new GridColumn { Width = "1000" },
            new GridColumn { Width = "1000" }),
        new TableRow(Cell("一"), Cell("二"), Cell("三")));
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

TableCell CellWithSimpleField(string text, int span = 1, MergedCellValues? merge = null)
{
    var properties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1200" });
    if (span > 1) properties.Append(new GridSpan { Val = span });
    if (merge is not null) properties.Append(new VerticalMerge { Val = merge.Value });
    return new TableCell(properties, new Paragraph(
        new SimpleField(new Run(new Text(text))) { Instruction = " PAGE " }));
}

TableCell CellWithSuperscript(string text, string superscript)
{
    var properties = new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "1200" });
    return new TableCell(properties, new Paragraph(
        new Run(new Text(text)),
        new Run(new RunProperties(
            new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }), new Text(superscript))));
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
