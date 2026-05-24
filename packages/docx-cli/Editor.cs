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
            "replaceAllHeaderParagraphText" => ReplaceAllHeaderParagraphText(doc, operation),
            "replaceHeaderParagraphText" => ReplaceHeaderParagraphText(doc, operation),
            "replaceHeaderText" => ReplaceHeaderText(doc, operation),
            "replaceTableCellText" => ReplaceTableCellText(body, operation),
            "replaceTable" => ReplaceTable(body, operation),
            "setTableWidth" => SetTableWidth(body, operation),
            "setTableCellAlignment" => SetTableCellAlignment(body, operation),
            "mergeTableCells" => MergeTableCells(body, operation),
            "fillTableSemantically" => FillTableSemantically(body, operation),
            "deleteComment" => DeleteComments(doc, operation.CommentId is { Length: > 0 } id ? [id] : []),
            "deleteComments" => DeleteComments(doc, operation.CommentIds ?? []),
            "markFieldsDirty" => MarkFieldsDirty(doc),
            "sanitizeFields" => SanitizeFields(doc),
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

        var paragraphs = body.Elements<Paragraph>().ToList();
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range");
        }

        ReplaceWholeParagraphText(paragraphs[operation.ParagraphIndex.Value], operation.Text);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated paragraph {operation.ParagraphIndex}");
    }

    private static DocxEditAppliedOperation ReplaceHeaderParagraphText(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (operation.HeaderIndex is null || operation.ParagraphIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "headerIndex, paragraphIndex, and text are required");
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var headers = mainPart.HeaderParts
            .Where(part => part.Header is not null)
            .OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal)
            .ToList();
        if (operation.HeaderIndex.Value < 0 || operation.HeaderIndex.Value >= headers.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"headerIndex {operation.HeaderIndex} is out of range");
        }

        var paragraphs = headers[operation.HeaderIndex.Value].Header!.Elements<Paragraph>().ToList();
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range for header {operation.HeaderIndex}");
        }

        ReplaceWholeParagraphText(paragraphs[operation.ParagraphIndex.Value], operation.Text);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated header[{operation.HeaderIndex}].paragraph[{operation.ParagraphIndex}]");
    }

    private static DocxEditAppliedOperation ReplaceAllHeaderParagraphText(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (operation.ParagraphIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "paragraphIndex and text are required");
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var updated = 0;
        foreach (var headerPart in mainPart.HeaderParts.Where(part => part.Header is not null))
        {
            var paragraphs = headerPart.Header!.Elements<Paragraph>().ToList();
            if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
            {
                continue;
            }

            ReplaceWholeParagraphText(paragraphs[operation.ParagraphIndex.Value], operation.Text);
            updated++;
        }

        if (updated == 0)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} was not found in any header");
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated paragraph {operation.ParagraphIndex} in {updated} header part(s)");
    }

    private static DocxEditAppliedOperation ReplaceHeaderText(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (string.IsNullOrEmpty(operation.FindText) || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "findText and text are required");
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var replaced = 0;
        foreach (var headerPart in mainPart.HeaderParts.Where(part => part.Header is not null))
        {
            foreach (var paragraph in headerPart.Header!.Descendants<Paragraph>())
            {
                if (ReplaceTextInParagraph(paragraph, operation.FindText, operation.Text))
                {
                    replaced++;
                }
            }
        }

        if (replaced == 0)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Header text not found: {operation.FindText}");
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Replaced header text in {replaced} paragraph(s)");
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

        ReplaceTableCellText(cells[operation.CellIndex.Value], operation.Text, operation.Alignment);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}]");
    }

    private static DocxEditAppliedOperation SetTableWidth(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex is required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var properties = tables[operation.TableIndex.Value].GetFirstChild<TableProperties>() ?? tables[operation.TableIndex.Value].PrependChild(new TableProperties());
        properties.RemoveAllChildren<TableWidth>();
        var widthType = string.Equals(operation.WidthType, "dxa", StringComparison.OrdinalIgnoreCase)
            ? TableWidthUnitValues.Dxa
            : TableWidthUnitValues.Pct;
        properties.PrependChild(new TableWidth
        {
            Width = string.IsNullOrWhiteSpace(operation.Width) ? "5000" : operation.Width,
            Type = widthType,
        });
        properties.RemoveAllChildren<TableLayout>();
        properties.AppendChild(new TableLayout { Type = TableLayoutValues.Autofit });
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}] width");
    }

    private static DocxEditAppliedOperation SetTableCellAlignment(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || string.IsNullOrWhiteSpace(operation.Alignment))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and alignment are required");
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

        ApplyCellAlignment(cells[operation.CellIndex.Value], operation.Alignment);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}] alignment");
    }

    private static DocxEditAppliedOperation ReplaceTable(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.Rows is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and rows are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var sourceTable = tables[operation.TableIndex.Value];
        var replacement = BuildReplacementTable(sourceTable, operation.Rows);
        sourceTable.InsertAfterSelf(replacement);
        sourceTable.Remove();
        return new DocxEditAppliedOperation(operation.Type, true, $"Replaced table[{operation.TableIndex}] with {operation.Rows.Count} row(s)");
    }

    private static Table BuildReplacementTable(Table sourceTable, IReadOnlyList<IReadOnlyList<DocxTableCellInput>> rows)
    {
        var table = new Table();
        var sourceProperties = sourceTable.GetFirstChild<TableProperties>();
        table.AppendChild(sourceProperties is null ? new TableProperties() : (TableProperties)sourceProperties.CloneNode(true));
        EnsureFullWidth(table.GetFirstChild<TableProperties>()!);

        var maxColumns = rows.Count == 0 ? 1 : rows.Max(row => row.Sum(cell => Math.Max(1, cell.GridSpan ?? 1)));
        var sourceGrid = sourceTable.GetFirstChild<TableGrid>();
        if (sourceGrid is not null)
        {
            var grid = (TableGrid)sourceGrid.CloneNode(true);
            while (grid.Elements<GridColumn>().Count() < maxColumns)
            {
                grid.AppendChild(new GridColumn { Width = "1200" });
            }
            table.AppendChild(grid);
        }
        else
        {
            var grid = new TableGrid();
            for (var i = 0; i < maxColumns; i++)
            {
                grid.AppendChild(new GridColumn { Width = "1200" });
            }
            table.AppendChild(grid);
        }

        var templateRows = sourceTable.Elements<TableRow>().ToList();
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var templateRow = templateRows.ElementAtOrDefault(Math.Min(rowIndex, Math.Max(0, templateRows.Count - 1)));
            var row = BuildReplacementRow(templateRow, rows[rowIndex], rowIndex == 0 || rows[rowIndex].Any(cell => cell.Header == true));
            table.AppendChild(row);
        }

        return table;
    }

    private static void EnsureFullWidth(TableProperties properties)
    {
        properties.RemoveAllChildren<TableWidth>();
        properties.PrependChild(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });
        properties.RemoveAllChildren<TableLayout>();
        properties.AppendChild(new TableLayout { Type = TableLayoutValues.Autofit });
    }

    private static TableRow BuildReplacementRow(TableRow? templateRow, IReadOnlyList<DocxTableCellInput> cells, bool isHeader)
    {
        var row = new TableRow();
        var templateProperties = templateRow?.GetFirstChild<TableRowProperties>();
        if (templateProperties is not null)
        {
            row.AppendChild((TableRowProperties)templateProperties.CloneNode(true));
        }
        if (isHeader)
        {
            var properties = row.GetFirstChild<TableRowProperties>() ?? row.PrependChild(new TableRowProperties());
            if (!properties.Elements<TableHeader>().Any())
            {
                properties.AppendChild(new TableHeader());
            }
        }

        var templateCells = templateRow?.Elements<TableCell>().ToList() ?? [];
        for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
        {
            var templateCell = templateCells.ElementAtOrDefault(Math.Min(cellIndex, Math.Max(0, templateCells.Count - 1)));
            row.AppendChild(BuildReplacementCell(templateCell, cells[cellIndex], isHeader));
        }

        return row;
    }

    private static TableCell BuildReplacementCell(TableCell? templateCell, DocxTableCellInput input, bool rowIsHeader)
    {
        var cell = new TableCell();
        var templateProperties = templateCell?.GetFirstChild<TableCellProperties>();
        if (templateProperties is not null)
        {
            cell.AppendChild((TableCellProperties)templateProperties.CloneNode(true));
        }
        else
        {
            cell.AppendChild(new TableCellProperties());
        }

        var properties = cell.GetFirstChild<TableCellProperties>()!;
        properties.RemoveAllChildren<TableCellWidth>();
        
        properties.RemoveAllChildren<GridSpan>();
        if (input.GridSpan is > 1)
        {
            properties.AppendChild(new GridSpan { Val = input.GridSpan.Value });
        }

        properties.RemoveAllChildren<VerticalMerge>();
        if (input.VMerge is { Length: > 0 } vMergeVal)
        {
            var vmVal = vMergeVal.ToLowerInvariant() == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue;
            properties.AppendChild(new VerticalMerge { Val = vmVal });
        }

        if (input.Shading is { Length: > 0 } hexColor)
        {
            properties.RemoveAllChildren<Shading>();
            properties.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = hexColor });
        }
        NormalizeTableCellProperties(properties);

        var paragraph = new Paragraph();
        var paragraphProperties = new ParagraphProperties();
        
        if (input.Alignment is { Length: > 0 } align)
        {
            var jcVal = align.ToLowerInvariant() switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                _ => JustificationValues.Left
            };
            paragraphProperties.AppendChild(new Justification { Val = jcVal });
        }
        paragraph.AppendChild(paragraphProperties);

        var run = new Run();
        if (input.Bold == true || rowIsHeader)
        {
            run.AppendChild(new RunProperties(new Bold()));
        }
        AppendTextWithLineBreaks(run, input.Text ?? string.Empty);
        paragraph.AppendChild(run);
        cell.AppendChild(paragraph);
        return cell;
    }

    private static void AppendTextWithLineBreaks(Run run, string text)
    {
        var lines = text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                run.AppendChild(new Break());
            }
            run.AppendChild(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
        }
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

    private static DocxEditAppliedOperation SanitizeFields(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        mainPart.DocumentSettingsPart?.Settings?.RemoveAllChildren<UpdateFieldsOnOpen>();

        foreach (var root in Inspector.GetRoots(doc))
        {
            foreach (var fieldChar in root.Descendants<FieldChar>().Where(fieldChar => fieldChar.Dirty != null))
            {
                fieldChar.Dirty = null;
            }
        }

        return new DocxEditAppliedOperation("sanitizeFields", true, "Sanitized field-update risks");
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

    private static bool ReplaceTextInParagraph(Paragraph paragraph, string findText, string replacementText)
    {
        var texts = paragraph.Descendants<Text>().ToList();
        if (texts.Count == 0)
        {
            return false;
        }

        var textSpans = new List<(Text Text, int Start, int End)>();
        var cursor = 0;
        foreach (var text in texts)
        {
            var value = text.Text ?? string.Empty;
            textSpans.Add((text, cursor, cursor + value.Length));
            cursor += value.Length;
        }

        var fullText = string.Concat(texts.Select(text => text.Text ?? string.Empty));
        var index = fullText.IndexOf(findText, StringComparison.Ordinal);
        if (index < 0)
        {
            return false;
        }

        var endIndex = index + findText.Length;
        var startSpanIndex = textSpans.FindIndex(span => index >= span.Start && index < span.End);
        var endSpanIndex = textSpans.FindIndex(span => endIndex > span.Start && endIndex <= span.End);
        if (startSpanIndex < 0 || endSpanIndex < 0)
        {
            return false;
        }

        var startSpan = textSpans[startSpanIndex];
        var endSpan = textSpans[endSpanIndex];
        var prefix = (startSpan.Text.Text ?? string.Empty)[..(index - startSpan.Start)];
        var suffix = (endSpan.Text.Text ?? string.Empty)[(endIndex - endSpan.Start)..];

        if (startSpanIndex == endSpanIndex)
        {
            startSpan.Text.Text = prefix + replacementText + suffix;
            return true;
        }

        startSpan.Text.Text = prefix + replacementText;
        for (var i = startSpanIndex + 1; i < endSpanIndex; i++)
        {
            textSpans[i].Text.Text = string.Empty;
        }
        endSpan.Text.Text = suffix;

        return true;
    }

    private static void ReplaceTableCellText(TableCell cell, string replacementText, string? alignment = null)
    {
        var firstRun = cell.Descendants<Run>().FirstOrDefault();
        cell.RemoveAllChildren<Paragraph>();
        var paragraph = new Paragraph(CreateStyledRunLike(firstRun, replacementText));
        if (!string.IsNullOrWhiteSpace(alignment))
        {
            ApplyParagraphAlignment(paragraph, alignment);
        }
        cell.Append(paragraph);
    }

    private static void ApplyCellAlignment(TableCell cell, string alignment)
    {
        foreach (var paragraph in cell.Elements<Paragraph>())
        {
            ApplyParagraphAlignment(paragraph, alignment);
        }
    }

    private static void ApplyParagraphAlignment(Paragraph paragraph, string alignment)
    {
        var properties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
        properties.RemoveAllChildren<Justification>();
        properties.AppendChild(new Justification
        {
            Val = alignment.ToLowerInvariant() switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                _ => JustificationValues.Left,
            },
        });
    }

    private static Run CreateStyledRunLike(Run? templateRun, string text)
    {
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (templateRun?.RunProperties is not null)
        {
            run.RunProperties = (RunProperties)templateRun.RunProperties.CloneNode(true);
            NormalizeRunProperties(run.RunProperties);
        }
        return run;
    }

    private static void NormalizeGeneratedOpenXml(WordprocessingDocument doc)
    {
        foreach (var root in Inspector.GetRoots(doc))
        {
            foreach (var properties in root.Descendants<TableProperties>())
            {
                NormalizeTableProperties(properties);
            }
            foreach (var properties in root.Descendants<TableCellProperties>())
            {
                NormalizeTableCellProperties(properties);
            }
            foreach (var properties in root.Descendants<RunProperties>())
            {
                NormalizeRunProperties(properties);
            }
        }
    }

    private static void NormalizeRunProperties(RunProperties properties)
        => SortChildrenByOpenXmlOrder(properties, RunPropertyOrder);

    private static void NormalizeTableCellProperties(TableCellProperties properties)
        => SortChildrenByOpenXmlOrder(properties, TableCellPropertyOrder);

    private static void NormalizeTableProperties(TableProperties properties)
        => SortChildrenByOpenXmlOrder(properties, TablePropertyOrder);

    private static void SortChildrenByOpenXmlOrder(OpenXmlCompositeElement parent, IReadOnlyDictionary<Type, int> order)
    {
        var children = parent.ChildElements.ToList();
        if (children.Count < 2)
        {
            return;
        }

        var sorted = children
            .Select((child, index) => new { Child = child, Index = index })
            .OrderBy(item => order.TryGetValue(item.Child.GetType(), out var childOrder) ? childOrder : int.MaxValue)
            .ThenBy(item => item.Index)
            .Select(item => item.Child.CloneNode(true))
            .ToList();
        parent.RemoveAllChildren();
        foreach (var child in sorted)
        {
            parent.AppendChild(child);
        }
    }

    private static readonly IReadOnlyDictionary<Type, int> RunPropertyOrder = new Dictionary<Type, int>
    {
        [typeof(RunStyle)] = 0,
        [typeof(RunFonts)] = 1,
        [typeof(Bold)] = 2,
        [typeof(BoldComplexScript)] = 3,
        [typeof(Italic)] = 4,
        [typeof(ItalicComplexScript)] = 5,
        [typeof(Caps)] = 6,
        [typeof(SmallCaps)] = 7,
        [typeof(Strike)] = 8,
        [typeof(DoubleStrike)] = 9,
        [typeof(Outline)] = 10,
        [typeof(Shadow)] = 11,
        [typeof(Emboss)] = 12,
        [typeof(Imprint)] = 13,
        [typeof(NoProof)] = 14,
        [typeof(SnapToGrid)] = 15,
        [typeof(Vanish)] = 16,
        [typeof(WebHidden)] = 17,
        [typeof(Color)] = 20,
        [typeof(Spacing)] = 21,
        [typeof(CharacterScale)] = 22,
        [typeof(Kern)] = 23,
        [typeof(Position)] = 24,
        [typeof(FontSize)] = 30,
        [typeof(FontSizeComplexScript)] = 31,
        [typeof(Highlight)] = 32,
        [typeof(Underline)] = 33,
        [typeof(TextEffect)] = 34,
        [typeof(Border)] = 35,
        [typeof(Shading)] = 36,
        [typeof(FitText)] = 37,
        [typeof(VerticalTextAlignment)] = 38,
        [typeof(RightToLeftText)] = 39,
        [typeof(Languages)] = 40,
    };

    private static readonly IReadOnlyDictionary<Type, int> TableCellPropertyOrder = new Dictionary<Type, int>
    {
        [typeof(ConditionalFormatStyle)] = 0,
        [typeof(TableCellWidth)] = 1,
        [typeof(GridSpan)] = 2,
        [typeof(HorizontalMerge)] = 3,
        [typeof(VerticalMerge)] = 4,
        [typeof(TableCellBorders)] = 5,
        [typeof(Shading)] = 6,
        [typeof(NoWrap)] = 7,
        [typeof(TableCellMargin)] = 8,
        [typeof(TextDirection)] = 9,
        [typeof(TableCellFitText)] = 10,
        [typeof(TableCellVerticalAlignment)] = 11,
        [typeof(HideMark)] = 12,
    };

    private static readonly IReadOnlyDictionary<Type, int> TablePropertyOrder = new Dictionary<Type, int>
    {
        [typeof(TableStyle)] = 0,
        [typeof(TablePositionProperties)] = 1,
        [typeof(TableOverlap)] = 2,
        [typeof(BiDiVisual)] = 3,
        [typeof(TableStyleRowBandSize)] = 4,
        [typeof(TableStyleColumnBandSize)] = 5,
        [typeof(TableWidth)] = 6,
        [typeof(TableJustification)] = 7,
        [typeof(TableCellSpacing)] = 8,
        [typeof(TableIndentation)] = 9,
        [typeof(TableBorders)] = 10,
        [typeof(Shading)] = 11,
        [typeof(TableLayout)] = 12,
        [typeof(TableCellMarginDefault)] = 13,
        [typeof(TableLook)] = 14,
        [typeof(TableCaption)] = 15,
        [typeof(TableDescription)] = 16,
    };

    private static DocxEditAppliedOperation MergeTableCells(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex is required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var rows = table.Elements<TableRow>().ToList();

        if (operation.RowIndex is not null)
        {
            var rowIndex = operation.RowIndex.Value;
            if (rowIndex < 0 || rowIndex >= rows.Count)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {rowIndex} is out of range");
            }

            var row = rows[rowIndex];
            var cells = row.Elements<TableCell>().ToList();
            var startCellIndex = operation.StartCellIndex ?? 0;
            var endCellIndex = operation.EndCellIndex ?? (cells.Count - 1);

            if (startCellIndex < 0 || endCellIndex >= cells.Count || endCellIndex <= startCellIndex)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"Invalid cell range {startCellIndex} to {endCellIndex}");
            }

            var selected = cells.Skip(startCellIndex).Take(endCellIndex - startCellIndex + 1).ToList();
            var totalSpan = selected.Sum(cell => {
                var span = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>();
                if (span?.Val?.Value is int val) return val;
                return 1;
            });

            var properties = selected[0].GetFirstChild<TableCellProperties>() ?? selected[0].PrependChild(new TableCellProperties());
            properties.RemoveAllChildren<GridSpan>();
            if (totalSpan > 1)
            {
                properties.AppendChild(new GridSpan { Val = totalSpan });
            }
            NormalizeTableCellProperties(properties);

            foreach (var cell in selected.Skip(1))
            {
                row.RemoveChild(cell);
            }

            return new DocxEditAppliedOperation(operation.Type, true, $"Merged table[{operation.TableIndex}].row[{rowIndex}].cells[{startCellIndex}..{endCellIndex}]");
        }
        else if (operation.CellIndex is not null)
        {
            var cellIndex = operation.CellIndex.Value;
            var startRowIndex = operation.StartRowIndex ?? 0;
            var endRowIndex = operation.EndRowIndex ?? (rows.Count - 1);

            if (startRowIndex < 0 || endRowIndex >= rows.Count || endRowIndex <= startRowIndex)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {startRowIndex} to {endRowIndex}");
            }

            for (var rIdx = startRowIndex; rIdx <= endRowIndex; rIdx++)
            {
                var rCells = rows[rIdx].Elements<TableCell>().ToList();
                if (cellIndex >= rCells.Count)
                {
                    return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {cellIndex} is out of range in row {rIdx}");
                }
            }

            for (var rIdx = startRowIndex; rIdx <= endRowIndex; rIdx++)
            {
                var cell = rows[rIdx].Elements<TableCell>().ElementAt(cellIndex);
                var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
                properties.RemoveAllChildren<VerticalMerge>();
                var mergeValue = rIdx == startRowIndex ? MergedCellValues.Restart : MergedCellValues.Continue;
                properties.AppendChild(new VerticalMerge { Val = mergeValue });
                NormalizeTableCellProperties(properties);
                if (rIdx != startRowIndex)
                {
                    cell.RemoveAllChildren<Paragraph>();
                    cell.AppendChild(new Paragraph());
                }
            }

            return new DocxEditAppliedOperation(operation.Type, true, $"Vertically merged table[{operation.TableIndex}].cell[{cellIndex}].rows[{startRowIndex}..{endRowIndex}]");
        }

        return new DocxEditAppliedOperation(operation.Type, false, "Either rowIndex (horizontal) or cellIndex (vertical) must be specified for merge");
    }

    private static DocxEditAppliedOperation FillTableSemantically(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.Cells is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and cells are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var gridMap = new TableGridMap(table);
        var appliedCount = 0;

        foreach (var rule in operation.Cells)
        {
            for (var r = 0; r < gridMap.RowCount; r++)
            {
                var rowContext = gridMap.GetRowContext(r);
                var rowMatches = rule.RowPatterns.All(p => rowContext.Contains(p, StringComparison.OrdinalIgnoreCase));
                if (!rowMatches)
                {
                    continue;
                }

                for (var col = 0; col < gridMap.ColumnCount; col++)
                {
                    var colContext = gridMap.GetColumnContext(col);
                    var colMatches = rule.ColPatterns.All(p => colContext.Contains(p, StringComparison.OrdinalIgnoreCase));
                    if (!colMatches)
                    {
                        continue;
                    }

                    var cell = gridMap.Grid[r, col];
                    if (cell != null)
                    {
                        ReplaceTableCellText(cell, rule.Text);
                        appliedCount++;
                    }
                }
            }
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Successfully applied semantic fills to {appliedCount} cell(s) in table[{operation.TableIndex}]");
    }
}
