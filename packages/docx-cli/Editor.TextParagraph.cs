using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace Dockit.Docx;

public static partial class Editor
{
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

    private static DocxEditAppliedOperation ReplaceParagraphRunText(Body body, DocxEditOperation operation)
    {
        if (operation.ParagraphIndex is null || operation.RunIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "paragraphIndex, runIndex, and text are required");
        }

        var paragraphs = body.Elements<Paragraph>().ToList();
        if (!TryReplaceRunText(paragraphs, operation.ParagraphIndex.Value, operation.RunIndex.Value, operation.Text, out var error))
        {
            return new DocxEditAppliedOperation(operation.Type, false, error);
        }
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated paragraph {operation.ParagraphIndex} run {operation.RunIndex}");
    }

    private static DocxEditAppliedOperation ReplaceBodyText(Body body, DocxEditOperation operation)
    {
        if (string.IsNullOrEmpty(operation.FindText) || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "findText and text are required");
        }

        var replaced = 0;
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            if (ReplaceTextInParagraph(paragraph, operation.FindText, operation.Text))
            {
                replaced++;
            }
        }

        if (replaced == 0)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Body text not found: {operation.FindText}");
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Replaced body text in {replaced} paragraph(s)");
    }

    private static DocxEditAppliedOperation DeleteBodyParagraph(Body body, DocxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.FindText))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "findText is required");
        }

        if (!TryResolveParagraphMatchMode(operation.MatchMode, out var matchMode, out var modeError))
        {
            return new DocxEditAppliedOperation(operation.Type, false, modeError);
        }

        var matches = body.Descendants<Paragraph>()
            .Where(paragraph => ParagraphMatches(paragraph, operation.FindText, matchMode, operation.ParagraphStyle))
            .ToList();
        if (matches.Count != 1)
        {
            return new DocxEditAppliedOperation(
                operation.Type,
                false,
                $"Expected exactly one body paragraph for {matchMode} selector '{operation.FindText}', found {matches.Count}");
        }

        matches[0].Remove();
        return new DocxEditAppliedOperation(operation.Type, true, $"Deleted body paragraph matching: {operation.FindText}");
    }

    private static DocxEditAppliedOperation DeleteBodyDrawingBeforeParagraph(Body body, DocxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.FindText))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "findText is required");
        }

        if (!TryResolveParagraphMatchMode(operation.MatchMode, out var matchMode, out var modeError))
        {
            return new DocxEditAppliedOperation(operation.Type, false, modeError);
        }

        var children = body.ChildElements.ToList();
        var matches = children
            .Select((child, index) => (child, index))
            .Where(candidate => candidate.child is Paragraph paragraph
                && ParagraphMatches(paragraph, operation.FindText, matchMode, operation.ParagraphStyle))
            .ToList();
        if (matches.Count != 1)
        {
            return new DocxEditAppliedOperation(
                operation.Type,
                false,
                $"Expected exactly one direct body paragraph for {matchMode} selector '{operation.FindText}', found {matches.Count}");
        }

        var anchorIndex = matches[0].index;
        if (anchorIndex == 0 || children[anchorIndex - 1] is not Paragraph drawingParagraph)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "preceding direct body element is not a paragraph");
        }

        var drawings = drawingParagraph.Descendants<Drawing>().ToList();
        if (drawings.Count != 1)
        {
            return new DocxEditAppliedOperation(
                operation.Type,
                false,
                $"Expected exactly one drawing in the preceding body paragraph, found {drawings.Count}");
        }

        if (!string.IsNullOrWhiteSpace(GetParagraphText(drawingParagraph)))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "preceding drawing paragraph also contains text");
        }

        drawingParagraph.Remove();
        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Deleted the single drawing-only body paragraph before: {operation.FindText}");
    }

    private static DocxEditAppliedOperation DeleteBodyRange(Body body, DocxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.FindText))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "findText is required");
        }

        if (!TryResolveParagraphMatchMode(operation.MatchMode, out var startMatchMode, out var startModeError))
        {
            return new DocxEditAppliedOperation(operation.Type, false, startModeError);
        }

        var deleteToBodyEnd = operation.DeleteToBodyEnd == true;
        if (deleteToBodyEnd == !string.IsNullOrWhiteSpace(operation.EndFindText))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "exactly one of endFindText or deleteToBodyEnd=true is required");
        }

        var children = body.ChildElements.ToList();
        var startMatches = children
            .Select((child, index) => (child, index))
            .Where(candidate => candidate.child is Paragraph paragraph
                && ParagraphMatches(paragraph, operation.FindText, startMatchMode, operation.ParagraphStyle))
            .ToList();
        if (startMatches.Count != 1)
        {
            return new DocxEditAppliedOperation(
                operation.Type,
                false,
                $"Expected exactly one direct body paragraph for {startMatchMode} selector '{operation.FindText}', found {startMatches.Count}");
        }

        var startIndex = startMatches[0].index;
        var endIndex = children.Count;
        if (!deleteToBodyEnd)
        {
            if (!TryResolveParagraphMatchMode(operation.EndMatchMode, out var endMatchMode, out var endModeError))
            {
                return new DocxEditAppliedOperation(operation.Type, false, endModeError);
            }

            var endMatches = children
                .Select((child, index) => (child, index))
                .Where(candidate => candidate.index > startIndex
                    && candidate.child is Paragraph paragraph
                    && ParagraphMatches(paragraph, operation.EndFindText!, endMatchMode, operation.EndParagraphStyle))
                .ToList();
            if (endMatches.Count != 1)
            {
                return new DocxEditAppliedOperation(
                    operation.Type,
                    false,
                    $"Expected exactly one following direct body paragraph for {endMatchMode} selector '{operation.EndFindText}', found {endMatches.Count}");
            }
            endIndex = endMatches[0].index;
        }
        else
        {
            var finalSectionProperties = children.FindLastIndex(child => child is SectionProperties);
            if (finalSectionProperties >= startIndex)
            {
                endIndex = finalSectionProperties;
            }
        }

        if (endIndex <= startIndex)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "body range is empty or reversed");
        }

        var selected = children.Skip(startIndex).Take(endIndex - startIndex).ToList();
        if (selected.Any(child => child is SectionProperties))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "body range cannot delete document-level section properties");
        }

        var removedPrecedingPageBreak = false;
        if (operation.RemovePrecedingPageBreak == true)
        {
            if (!deleteToBodyEnd)
            {
                return new DocxEditAppliedOperation(operation.Type, false, "removePrecedingPageBreak requires deleteToBodyEnd=true");
            }
            if (startIndex == 0 || children[startIndex - 1] is not Paragraph boundaryParagraph)
            {
                return new DocxEditAppliedOperation(operation.Type, false, "preceding body element is not a paragraph");
            }

            var pageBreaks = boundaryParagraph.Descendants<Break>()
                .Where(element => element.Type?.Value == BreakValues.Page)
                .ToList();
            if (pageBreaks.Count != 1)
            {
                return new DocxEditAppliedOperation(
                    operation.Type,
                    false,
                    $"Expected exactly one explicit page break in the preceding paragraph, found {pageBreaks.Count}");
            }

            pageBreaks[0].Remove();
            removedPrecedingPageBreak = true;
        }

        foreach (var child in selected)
        {
            child.Remove();
        }

        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Deleted {selected.Count} direct body element(s) beginning at paragraph: {operation.FindText}"
                + (removedPrecedingPageBreak ? " and removed the preceding explicit page break" : string.Empty));
    }

    private static bool TryResolveParagraphMatchMode(string? requested, out string mode, out string error)
    {
        mode = string.IsNullOrWhiteSpace(requested) ? "exact" : requested.Trim();
        if (mode is "exact" or "startsWith")
        {
            error = string.Empty;
            return true;
        }

        error = $"Unsupported paragraph matchMode: {mode}";
        return false;
    }

    private static bool ParagraphMatches(Paragraph paragraph, string expected, string mode, string? expectedStyle)
    {
        var actual = GetParagraphText(paragraph).Trim();
        var selector = expected.Trim();
        var textMatches = mode == "startsWith"
            ? actual.StartsWith(selector, StringComparison.Ordinal)
            : string.Equals(actual, selector, StringComparison.Ordinal);
        if (!textMatches || string.IsNullOrWhiteSpace(expectedStyle))
        {
            return textMatches;
        }

        var actualStyle = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return string.Equals(actualStyle, expectedStyle.Trim(), StringComparison.Ordinal);
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

    private static DocxEditAppliedOperation ReplaceHeaderParagraphRunText(WordprocessingDocument doc, DocxEditOperation operation)
        => ReplacePartParagraphRunText(doc, operation, "header");

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

    private static DocxEditAppliedOperation ReplaceFooterParagraphText(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (operation.FooterIndex is null || operation.ParagraphIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "footerIndex, paragraphIndex, and text are required");
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var footers = mainPart.FooterParts
            .Where(part => part.Footer is not null)
            .OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal)
            .ToList();
        if (operation.FooterIndex.Value < 0 || operation.FooterIndex.Value >= footers.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"footerIndex {operation.FooterIndex} is out of range");
        }

        var paragraphs = footers[operation.FooterIndex.Value].Footer!.Elements<Paragraph>().ToList();
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range for footer {operation.FooterIndex}");
        }

        ReplaceWholeParagraphText(paragraphs[operation.ParagraphIndex.Value], operation.Text);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated footer[{operation.FooterIndex}].paragraph[{operation.ParagraphIndex}]");
    }

    private static DocxEditAppliedOperation ReplaceFooterParagraphRunText(WordprocessingDocument doc, DocxEditOperation operation)
        => ReplacePartParagraphRunText(doc, operation, "footer");

    private static DocxEditAppliedOperation ReplacePartParagraphRunText(WordprocessingDocument doc, DocxEditOperation operation, string partKind)
    {
        if (operation.ParagraphIndex is null || operation.RunIndex is null || operation.Text is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "paragraphIndex, runIndex, and text are required");
        }

        var mainPart = doc.MainDocumentPart ?? throw new InvalidOperationException("Main document part not found.");
        var partIndex = partKind == "header" ? operation.HeaderIndex : operation.FooterIndex;
        if (partIndex is null) return new DocxEditAppliedOperation(operation.Type, false, $"{partKind}Index is required");
        var roots = partKind == "header"
            ? mainPart.HeaderParts.Where(part => part.Header is not null).OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => (OpenXmlPartRootElement)part.Header!).ToList()
            : mainPart.FooterParts.Where(part => part.Footer is not null).OrderBy(part => mainPart.GetIdOfPart(part), StringComparer.Ordinal).Select(part => (OpenXmlPartRootElement)part.Footer!).ToList();
        if (partIndex.Value < 0 || partIndex.Value >= roots.Count) return new DocxEditAppliedOperation(operation.Type, false, $"{partKind}Index {partIndex} is out of range");
        if (!TryReplaceRunText(roots[partIndex.Value].Elements<Paragraph>().ToList(), operation.ParagraphIndex.Value, operation.RunIndex.Value, operation.Text, out var error))
        {
            return new DocxEditAppliedOperation(operation.Type, false, error);
        }
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated {partKind}[{partIndex}].paragraph[{operation.ParagraphIndex}].run[{operation.RunIndex}]");
    }

    private static bool TryReplaceRunText(IReadOnlyList<Paragraph> paragraphs, int paragraphIndex, int runIndex, string text, out string error)
    {
        error = string.Empty;
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            error = $"paragraphIndex {paragraphIndex} is out of range";
            return false;
        }
        var runs = paragraphs[paragraphIndex].Descendants<Run>().ToList();
        if (runIndex < 0 || runIndex >= runs.Count)
        {
            error = $"runIndex {runIndex} is out of range for paragraph {paragraphIndex}";
            return false;
        }
        var texts = runs[runIndex].Descendants<Text>().ToList();
        if (texts.Count == 0)
        {
            error = $"runIndex {runIndex} in paragraph {paragraphIndex} has no text";
            return false;
        }
        texts[0].Text = text;
        foreach (var extra in texts.Skip(1)) extra.Text = string.Empty;
        return true;
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
        if (findText.Length == 0)
        {
            return false;
        }

        var replaced = false;
        var searchStart = 0;
        while (true)
        {
            var texts = paragraph.Descendants<Text>().ToList();
            if (texts.Count == 0)
            {
                return replaced;
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
            var index = fullText.IndexOf(findText, searchStart, StringComparison.Ordinal);
            if (index < 0)
            {
                return replaced;
            }

            var endIndex = index + findText.Length;
            var startSpanIndex = textSpans.FindIndex(span => index >= span.Start && index < span.End);
            var endSpanIndex = textSpans.FindIndex(span => endIndex > span.Start && endIndex <= span.End);
            if (startSpanIndex < 0 || endSpanIndex < 0)
            {
                return replaced;
            }

            var startSpan = textSpans[startSpanIndex];
            var endSpan = textSpans[endSpanIndex];
            var prefix = (startSpan.Text.Text ?? string.Empty)[..(index - startSpan.Start)];
            var suffix = (endSpan.Text.Text ?? string.Empty)[(endIndex - endSpan.Start)..];

            if (startSpanIndex == endSpanIndex)
            {
                startSpan.Text.Text = prefix + replacementText + suffix;
            }
            else
            {
                startSpan.Text.Text = prefix + replacementText;
                for (var i = startSpanIndex + 1; i < endSpanIndex; i++)
                {
                    textSpans[i].Text.Text = string.Empty;
                }
                endSpan.Text.Text = suffix;
            }

            replaced = true;
            searchStart = index + replacementText.Length;
        }
    }

    private static string GetParagraphText(Paragraph paragraph)
        => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
}

