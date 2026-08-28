using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace Dockit.Docx;

internal static partial class Editor
{
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
        settingsPart.Settings.AddChild(new UpdateFieldsOnOpen { Val = true }, true);

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

    private static DocxEditAppliedOperation FreezeFields(WordprocessingDocument doc)
    {
        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        mainPart.DocumentSettingsPart?.Settings?.RemoveAllChildren<UpdateFieldsOnOpen>();

        var frozenSimpleFields = 0;
        var frozenComplexFields = 0;

        foreach (var root in Inspector.GetRoots(doc))
        {
            foreach (var simpleField in root.Descendants<SimpleField>().ToList())
            {
                if (!ShouldFreezeFieldInstruction(simpleField.Instruction?.Value))
                {
                    continue;
                }

                var replacement = simpleField.ChildElements.Select(child => child.CloneNode(true)).ToList();
                foreach (var child in replacement)
                {
                    simpleField.InsertBeforeSelf(child);
                }

                simpleField.Remove();
                frozenSimpleFields++;
            }

            foreach (var paragraph in root.Descendants<Paragraph>().ToList())
            {
                frozenComplexFields += FreezeComplexFieldsInParagraph(paragraph);
            }
        }

        return new DocxEditAppliedOperation(
            "freezeFields",
            true,
            $"Froze {frozenSimpleFields} simple field(s) and {frozenComplexFields} complex field(s)");
    }

    private static int FreezeComplexFieldsInParagraph(Paragraph paragraph)
    {
        var frozen = 0;
        var index = 0;

        while (index < paragraph.ChildElements.Count)
        {
            var children = paragraph.ChildElements.ToList();
            var begin = children.FindIndex(index, IsFieldBeginRun);
            if (begin < 0)
            {
                break;
            }

            var depth = 0;
            var separate = -1;
            var end = -1;
            for (var cursor = begin; cursor < children.Count; cursor++)
            {
                if (children[cursor] is not Run run)
                {
                    continue;
                }

                var fieldCharType = run.GetFirstChild<FieldChar>()?.FieldCharType?.Value;
                if (fieldCharType == FieldCharValues.Begin)
                {
                    depth++;
                }
                else if (fieldCharType == FieldCharValues.Separate && depth == 1)
                {
                    separate = cursor;
                }
                else if (fieldCharType == FieldCharValues.End)
                {
                    depth--;
                    if (depth == 0)
                    {
                        end = cursor;
                        break;
                    }
                }
            }

            if (end < 0)
            {
                index = begin + 1;
                continue;
            }

            var instruction = string.Concat(children
                .Skip(begin + 1)
                .Take((separate >= 0 ? separate : end) - begin - 1)
                .OfType<Run>()
                .SelectMany(run => run.Elements<FieldCode>())
                .Select(code => code.Text));
            if (!ShouldFreezeFieldInstruction(instruction))
            {
                index = end + 1;
                continue;
            }

            var resultStart = separate >= 0 ? separate + 1 : end;
            var resultRuns = children
                .Skip(resultStart)
                .Take(end - resultStart)
                .Where(child => child is Run run && !IsFieldCodeRun(run))
                .Select(child => child.CloneNode(true))
                .ToList();

            foreach (var child in resultRuns)
            {
                paragraph.InsertBefore(child, children[begin]);
            }

            for (var cursor = begin; cursor <= end; cursor++)
            {
                children[cursor].Remove();
            }

            frozen++;
            index = begin + resultRuns.Count;
        }

        return frozen;
    }

    private static bool IsFieldBeginRun(OpenXmlElement element)
        => element is Run run && run.GetFirstChild<FieldChar>()?.FieldCharType?.Value == FieldCharValues.Begin;

    private static bool IsFieldCodeRun(Run run)
        => run.Elements<FieldChar>().Any() || run.Elements<FieldCode>().Any();

    private static bool ShouldFreezeFieldInstruction(string? instruction)
    {
        var trimmed = (instruction ?? string.Empty).TrimStart();
        return trimmed.StartsWith("REF ", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase);
    }
}
