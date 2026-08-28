using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace Dockit.Docx;

public static partial class Editor
{
    public static int RunEdit(string[] args)
    {
        if (args.Length < 3)
        {
            throw new InvalidOperationException("edit requires <input.docx> <operations.json> <output.docx>");
        }

        var input = Path.GetFullPath(args[0]);
        var operationsPath = Path.GetFullPath(args[1]);
        var output = Path.GetFullPath(args[2]);
        var request = LoadOperations(operationsPath);
        var result = Apply(input, output, request.Operations);
        Console.WriteLine(JsonSerializer.Serialize(result, Json.Options));
        return result.AppliedOperations.All(operation => operation.Applied) ? 0 : 1;
    }

    public static DocxEditResult Apply(string input, string output, IReadOnlyList<DocxEditOperation> operations)
    {
        File.Copy(input, output, overwrite: true);
        using var doc = WordprocessingDocument.Open(output, true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var repeatHeaderFailures = ValidateRepeatHeaderBatch(doc, operations);
        if (repeatHeaderFailures is not null)
        {
            return new DocxEditResult(Path.GetFullPath(input), Path.GetFullPath(output), repeatHeaderFailures);
        }
        var applied = new List<DocxEditAppliedOperation>();

        foreach (var operation in operations)
        {
            applied.Add(ApplyOperation(doc, body, operation));
        }

        NormalizeGeneratedOpenXml(doc);
        mainPart.Document.Save();
        foreach (var headerPart in mainPart.HeaderParts)
        {
            headerPart.Header?.Save();
        }
        foreach (var footerPart in mainPart.FooterParts)
        {
            footerPart.Footer?.Save();
        }
        mainPart.DocumentSettingsPart?.Settings?.Save();
        return new DocxEditResult(Path.GetFullPath(input), Path.GetFullPath(output), applied);
    }

    private static DocxEditDocument LoadOperations(string path)
    {
        var json = File.ReadAllText(path);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new DocxEditDocument([]);
        }

        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind == JsonValueKind.Array)
        {
            var ops = JsonSerializer.Deserialize<List<DocxEditOperation>>(json, Json.Options) ?? [];
            return new DocxEditDocument(ops);
        }

        return JsonSerializer.Deserialize<DocxEditDocument>(json, Json.Options) ?? new DocxEditDocument([]);
    }

    private static DocxEditAppliedOperation ApplyOperation(WordprocessingDocument doc, Body body, DocxEditOperation operation)
    {
        return operation.Type switch
        {
            "replaceAnchoredText" => ReplaceAnchoredText(body, operation),
            "replaceParagraphText" => ReplaceParagraphText(body, operation),
            "replaceParagraphRunText" => ReplaceParagraphRunText(body, operation),
            "replaceBodyText" => ReplaceBodyText(body, operation),
            "deleteBodyParagraph" => DeleteBodyParagraph(body, operation),
            "deleteBodyDrawingBeforeParagraph" => DeleteBodyDrawingBeforeParagraph(body, operation),
            "insertBodyRange" => DocxObjectActions.InsertBodyRange(doc, operation),
            "replaceDrawingImage" => DocxObjectActions.ReplaceDrawingImage(doc, operation),
            "insertBodyImage" => DocxObjectActions.InsertBodyImage(doc, operation),
            "deleteBodyRange" => DeleteBodyRange(body, operation),
            "startSectionBeforeParagraph" => StartSectionBeforeParagraph(body, operation),
            "replaceAllHeaderParagraphText" => ReplaceAllHeaderParagraphText(doc, operation),
            "replaceHeaderParagraphText" => ReplaceHeaderParagraphText(doc, operation),
            "replaceHeaderParagraphRunText" => ReplaceHeaderParagraphRunText(doc, operation),
            "replaceFooterParagraphText" => ReplaceFooterParagraphText(doc, operation),
            "replaceFooterParagraphRunText" => ReplaceFooterParagraphRunText(doc, operation),
            "replaceHeaderText" => ReplaceHeaderText(doc, operation),
            "replaceTableCellText" => ReplaceTableCellText(body, operation),
            "replaceTableCellRunText" => ReplaceTableCellRunText(body, operation),
            "replaceHeaderTableCellText" => ReplacePartTableCellText(doc, operation, "header"),
            "replaceHeaderTableCellRunText" => ReplacePartTableCellRunText(doc, operation, "header"),
            "replaceFooterTableCellText" => ReplacePartTableCellText(doc, operation, "footer"),
            "replaceFooterTableCellRunText" => ReplacePartTableCellRunText(doc, operation, "footer"),
            "replaceTableCellRichText" => ReplaceTableCellRichText(body, operation),
            "replaceTable" => ReplaceTable(body, operation),
            "insertTableRows" => InsertTableRows(body, operation),
            "deleteTableRows" => DeleteTableRows(body, operation),
            "replaceTableRows" => ReplaceTableRows(body, operation),
            "insertTableColumns" => InsertTableColumns(body, operation),
            "setTableWidth" => SetTableWidth(body, operation),
            "setTableCellAlignment" => SetTableCellAlignment(body, operation),
            "setTableCellNoWrap" => SetTableCellNoWrap(body, operation),
            "setTableCellFontSize" => SetTableCellFontSize(body, operation),
            "applyDocumentFontPolicy" => ApplyDocumentFontPolicy(body, operation),
            "setTableRowHeight" => SetTableRowHeight(body, operation),
            "setTableRowCantSplit" => SetTableRowCantSplit(body, operation),
            "setTableRowRepeatAsHeader" => SetTableRowRepeatAsHeader(doc, operation),
            "setTableRowKeepNext" => SetTableRowKeepNext(body, operation),
            "setBodyParagraphKeepNext" => SetBodyParagraphKeepNext(body, operation),
            "setBodyParagraphKeepLines" => SetBodyParagraphKeepLines(body, operation),
            "applyTocStylePolicy" => ApplyTocStylePolicy(doc, operation),
            "setHeaderParagraphFontSize" => SetHeaderParagraphFontSize(doc, operation),
            "collapseTrailingEmptySection" => CollapseTrailingEmptySection(body, operation),
            "collapseTrailingEmptyBodyParagraphs" => CollapseTrailingEmptyBodyParagraphs(body, operation),
            "mergeTableCells" => MergeTableCells(body, operation),
            "unmergeTableRowHorizontalCells" => UnmergeTableRowHorizontalCells(body, operation),
            "unmergeTableColumnVerticalCells" => UnmergeTableColumnVerticalCells(body, operation),
            "deleteComment" => DeleteComments(doc, operation.CommentId is { Length: > 0 } id ? [id] : []),
            "deleteComments" => DeleteComments(doc, operation.CommentIds ?? []),
            "markFieldsDirty" => MarkFieldsDirty(doc),
            "sanitizeFields" => SanitizeFields(doc),
            "freezeFields" => FreezeFields(doc),
            _ => new DocxEditAppliedOperation(operation.Type, false, $"Unknown operation type: {operation.Type}"),
        };
    }
}
