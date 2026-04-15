using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class Editor
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
        return 0;
    }

    public static DocxEditResult Apply(string input, string output, IReadOnlyList<DocxEditOperation> operations)
    {
        File.Copy(input, output, overwrite: true);
        using var doc = WordprocessingDocument.Open(output, true);
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var applied = new List<DocxEditAppliedOperation>();

        foreach (var operation in operations)
        {
            applied.Add(ApplyOperation(doc, body, operation));
        }

        mainPart.Document.Save();
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
            "replaceTableCellText" => ReplaceTableCellText(body, operation),
            "deleteComment" => DeleteComments(doc, operation.CommentId is { Length: > 0 } id ? [id] : []),
            "deleteComments" => DeleteComments(doc, operation.CommentIds ?? []),
            "markFieldsDirty" => MarkFieldsDirty(doc),
            _ => new DocxEditAppliedOperation(operation.Type, false, $"Unknown operation type: {operation.Type}"),
        };
    }

    private static DocxEditAppliedOperation ReplaceAnchoredText(Body body, DocxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.CommentId) || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "commentId and text are required");
        }

        var targetParagraph = body.Descendants<Paragraph>()
            .FirstOrDefault(paragraph => paragraph.Descendants<CommentRangeStart>().Any(start => start.Id?.Value == operation.CommentId));
        if (targetParagraph is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Comment anchor {operation.CommentId} not found");
        }

        var replaced = ReplaceCommentRangeInParagraph(targetParagraph, operation.CommentId, operation.Text);
        if (!replaced)
        {
            ReplaceWholeParagraphText(targetParagraph, operation.Text);
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated comment anchor {operation.CommentId}");
    }

    private static DocxEditAppliedOperation ReplaceParagraphText(Body body, DocxEditOperation operation)
    {
        if (operation.ParagraphIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "paragraphIndex and text are required");
        }

        var paragraphs = body.Descendants<Paragraph>().ToList();
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range");
        }

        ReplaceWholeParagraphText(paragraphs[operation.ParagraphIndex.Value], operation.Text);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated paragraph {operation.ParagraphIndex}");
    }

    private static DocxEditAppliedOperation ReplaceTableCellText(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and text are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");
        }

        var cells = rows[operation.RowIndex.Value].Elements<TableCell>().ToList();
        if (operation.CellIndex.Value < 0 || operation.CellIndex.Value >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {operation.CellIndex} is out of range");
        }

        ReplaceTableCellText(cells[operation.CellIndex.Value], operation.Text);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}]");
    }

    private static DocxEditAppliedOperation DeleteComments(WordprocessingDocument doc, IReadOnlyList<string> commentIds)
    {
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var deleteAll = commentIds.Count == 0;
        var targets = deleteAll
            ? mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Select(comment => comment.Id?.Value).Where(id => !string.IsNullOrWhiteSpace(id)).Cast<string>().ToHashSet(StringComparer.Ordinal) ?? []
            : commentIds.Where(id => !string.IsNullOrWhiteSpace(id)).ToHashSet(StringComparer.Ordinal);

        foreach (var root in Inspector.GetRoots(doc))
        {
            root.Descendants<CommentRangeStart>().Where(node => node.Id?.Value is string id && targets.Contains(id)).ToList().ForEach(node => node.Remove());
            root.Descendants<CommentRangeEnd>().Where(node => node.Id?.Value is string id && targets.Contains(id)).ToList().ForEach(node => node.Remove());
            root.Descendants<CommentReference>().Where(node => node.Id?.Value is string id && targets.Contains(id)).ToList().ForEach(node => node.Remove());
        }

        var commentsPart = mainPart.WordprocessingCommentsPart;
        if (commentsPart?.Comments is not null)
        {
            commentsPart.Comments.Elements<Comment>().Where(comment => comment.Id?.Value is string id && targets.Contains(id)).ToList().ForEach(comment => comment.Remove());
            commentsPart.Comments.Save();
            if (!commentsPart.Comments.Elements<Comment>().Any())
            {
                mainPart.DeletePart(commentsPart);
                if (mainPart.WordprocessingCommentsExPart is not null)
                {
                    mainPart.DeletePart(mainPart.WordprocessingCommentsExPart);
                }
            }
        }

        return new DocxEditAppliedOperation("deleteComments", true, deleteAll ? "Deleted all comments" : $"Deleted {targets.Count} comments");
    }

    private static DocxEditAppliedOperation MarkFieldsDirty(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        settingsPart.Settings.RemoveAllChildren<UpdateFieldsOnOpen>();
        settingsPart.Settings.AppendChild(new UpdateFieldsOnOpen { Val = true });

        foreach (var field in Inspector.GetRoots(doc).SelectMany(root => root.Descendants<SimpleField>()))
        {
            field.Dirty = true;
        }

        return new DocxEditAppliedOperation("markFieldsDirty", true, "Marked fields dirty and enabled update on open");
    }

    private static bool ReplaceCommentRangeInParagraph(Paragraph paragraph, string commentId, string replacementText)
    {
        var children = paragraph.ChildElements.ToList();
        var startIndex = children.FindIndex(child => child is CommentRangeStart start && start.Id?.Value == commentId);
        var endIndex = children.FindIndex(child => child is CommentRangeEnd end && end.Id?.Value == commentId);
        if (startIndex < 0 || endIndex < 0 || endIndex <= startIndex)
        {
            return false;
        }

        var elementsBetween = children.Skip(startIndex + 1).Take(endIndex - startIndex - 1).ToList();
        var firstRun = elementsBetween.OfType<Run>().FirstOrDefault();
        foreach (var element in elementsBetween)
        {
            element.Remove();
        }

        paragraph.InsertBefore(CreateStyledRunLike(firstRun, replacementText), paragraph.ChildElements[endIndex - elementsBetween.Count]);
        return true;
    }

    private static void ReplaceWholeParagraphText(Paragraph paragraph, string replacementText)
    {
        var firstRun = paragraph.Descendants<Run>().FirstOrDefault();
        var texts = paragraph.Descendants<Text>().ToList();
        if (texts.Count > 0)
        {
            texts[0].Text = replacementText;
            foreach (var extra in texts.Skip(1))
            {
                extra.Text = string.Empty;
            }
            return;
        }

        paragraph.RemoveAllChildren<Run>();
        paragraph.Append(CreateStyledRunLike(firstRun, replacementText));
    }

    private static void ReplaceTableCellText(TableCell cell, string replacementText)
    {
        var firstRun = cell.Descendants<Run>().FirstOrDefault();
        cell.RemoveAllChildren<Paragraph>();
        cell.Append(new Paragraph(CreateStyledRunLike(firstRun, replacementText)));
    }

    private static Run CreateStyledRunLike(Run? templateRun, string text)
    {
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (templateRun?.RunProperties is not null)
        {
            run.RunProperties = (RunProperties)templateRun.RunProperties.CloneNode(true);
        }
        return run;
    }
}
