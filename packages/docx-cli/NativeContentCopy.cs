using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeContentCopy
{
    public const string Command = "docx_replace_content_from_source";
    private const string MainStory = "/word/document.xml";
    private const string Word2010Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<CopyContentRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("replace-content-from-source-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = Describe(request.ReceiptOutput),
            output = Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static CopyContentReceipt Apply(
        CopyContentRequest request,
        IReadOnlyDictionary<int, IReadOnlySet<string>>? preservedBookmarkIdsByChange = null)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var targetPath = paths.Input;
        var outputPath = paths.Output;
        var receiptPath = paths.Receipt;

        var targetRefs = Observation.ResolveAddresses(
            targetPath,
            request.Changes.Select(change => change.Target).ToArray(),
            "changes.target");
        if (targetRefs.Any(item => item.Kind is not "paragraph" and not "cell" || item.StoryPart != MainStory))
            throw new InvalidOperationException("target-address-must-be-main-document-paragraph-or-cell");
        if (targetRefs.Select(item => item.Address).Distinct().Count() != targetRefs.Count)
            throw new InvalidOperationException("target-address-duplicate");

        var prepared = PrepareSources(request, targetRefs, preservedBookmarkIdsByChange);
        IReadOnlyDictionary<string, int> baselineIssues;
        using (var targetDocument = WordprocessingDocument.Open(targetPath, false))
        {
            var targets = targetRefs.Select(item =>
                Observation.ResolveNativePath(targetDocument, item.StoryPart, item.NativePath)).ToArray();
            NativeMutationSupport.RejectOverlappingTargets(targets);
            baselineIssues = ValidationIssueCounts(targetDocument);
            PreflightImports(targetDocument, prepared);
        }

        var temporaryPath = outputPath + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(targetPath, temporaryPath);
            using (var outputDocument = WordprocessingDocument.Open(temporaryPath, true))
            {
                ApplyChanges(outputDocument, prepared);
                outputDocument.MainDocumentPart?.Document?.Save();
                var issues = ValidationIssueCounts(outputDocument);
                var added = issues.FirstOrDefault(item => item.Value > baselineIssues.GetValueOrDefault(item.Key));
                if (added.Key is not null)
                    throw new InvalidOperationException($"output-added-openxml-validation-issues: {added.Key}");
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            var readback = ReadBack(outputPath, prepared);
            var receipt = new CopyContentReceipt(
                "tiwater.docx-replace-content-from-source-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                readback,
                outputPath);
            File.WriteAllText(receiptPath, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    private static IReadOnlyList<PreparedChange> PrepareSources(
        CopyContentRequest request,
        IReadOnlyList<ResolvedDocxAddress> targetRefs,
        IReadOnlyDictionary<int, IReadOnlySet<string>>? preservedBookmarkIdsByChange)
    {
        var result = new PreparedChange?[request.Changes.Count];
        var indexedChanges = request.Changes.Select((change, index) => new
        {
            Change = change,
            Index = index,
            SourcePath = Path.GetFullPath(change.SourceInput),
        });
        foreach (var group in indexedChanges.GroupBy(item => item.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            var items = group.ToArray();
            var selections = items.SelectMany(item => item.Change.SourceSelections.Select((selection, selectionIndex) => new
            {
                item.Index,
                SelectionIndex = selectionIndex,
                Selection = selection,
            })).ToArray();
            var sourceRefs = Observation.ResolveAddresses(
                group.Key,
                selections.Select(item => item.Selection.Address).ToArray(),
                "changes.sourceSelections.address");
            if (sourceRefs.Any(item => item.StoryPart != MainStory))
                throw new InvalidOperationException("source-address-must-be-main-document-object");
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceStyles = sourceDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            var sourceOffset = 0;
            foreach (var item in items)
            {
                var paragraphs = new List<PreparedParagraph>();
                Paragraph? activeInlineSourceParagraph = null;
                for (var selectionIndex = 0; selectionIndex < item.Change.SourceSelections.Count; selectionIndex++)
                {
                    var resolved = sourceRefs[sourceOffset++];
                    var element = Observation.ResolveNativePath(sourceDocument, resolved.StoryPart, resolved.NativePath);
                    try
                    {
                        AppendSelection(
                            paragraphs,
                            ref activeInlineSourceParagraph,
                            element,
                            sourceStyles,
                            item.Change.SourceSelections[selectionIndex],
                            preservedBookmarkIdsByChange?.GetValueOrDefault(item.Index));
                    }
                    catch (InvalidOperationException exception)
                    {
                        throw new InvalidOperationException(
                            $"changes[{item.Index}].sourceSelections[{selectionIndex}] " +
                            $"{resolved.StoryPart}#{resolved.NativePath}: {exception.Message}",
                            exception);
                    }
                }
                if (paragraphs.Count == 0) paragraphs.Add(new PreparedParagraph(new Paragraph(), false, null));
                result[item.Index] = new PreparedChange(targetRefs[item.Index], group.Key, paragraphs);
            }
        }
        return result.Select(item => item ?? throw new InvalidOperationException("source-change-not-prepared")).ToArray();
    }

    private static void PreflightImports(WordprocessingDocument targetDocument, IReadOnlyList<PreparedChange> changes)
    {
        var targetMain = targetDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        foreach (var group in changes.GroupBy(change => change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(change => change.Paragraphs).Select(item => item.Paragraph).Cast<OpenXmlElement>().ToArray();
            foreach (var relationshipId in DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal))
                if (!DocxObjectActions.CanCopyRelationship(sourceMain, relationshipId, out var error))
                    throw new InvalidOperationException(error);
            if (!DocxObjectActions.TryImportStyles(sourceMain, targetMain, roots, apply: false, out var styleError))
                throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, targetMain, roots, apply: false, out var numberingError))
                throw new InvalidOperationException(numberingError);
        }
    }

    private static void ApplyChanges(WordprocessingDocument outputDocument, IReadOnlyList<PreparedChange> changes)
    {
        var outputMain = outputDocument.MainDocumentPart ?? throw new InvalidOperationException("target-main-part-not-found");
        var outputBody = outputMain.Document?.Body ?? throw new InvalidOperationException("target-body-not-found");
        var targetStyles = outputMain.StyleDefinitionsPart?.Styles;
        var targets = changes.Select(change =>
            Observation.ResolveNativePath(outputDocument, change.Target.StoryPart, change.Target.NativePath)).ToArray();

        foreach (var group in changes.Select((change, index) => (change, index))
                     .GroupBy(item => item.change.SourcePath, StringComparer.OrdinalIgnoreCase))
        {
            using var sourceDocument = WordprocessingDocument.Open(group.Key, false);
            var sourceMain = sourceDocument.MainDocumentPart ?? throw new InvalidOperationException("source-main-part-not-found");
            var roots = group.SelectMany(item => item.change.Paragraphs).Select(p => p.Paragraph).Cast<OpenXmlElement>().ToArray();
            if (!DocxObjectActions.TryImportStyles(sourceMain, outputMain, roots, apply: true, out var styleError))
                throw new InvalidOperationException(styleError);
            if (!DocxObjectActions.TryImportNumbering(sourceMain, outputMain, roots, apply: true, out var numberingError))
                throw new InvalidOperationException(numberingError);
            var relationshipMap = DocxObjectActions.RelationshipIds(roots).Distinct(StringComparer.Ordinal)
                .ToDictionary(id => id, id => DocxObjectActions.CopyRelationship(sourceMain, outputMain, id), StringComparer.Ordinal);
            foreach (var item in group)
            {
                var paragraphs = item.change.Paragraphs
                    .Select(prepared => prepared with { Paragraph = (Paragraph)prepared.Paragraph.CloneNode(true) })
                    .ToArray();
                foreach (var paragraph in paragraphs) DocxObjectActions.RewriteRelationships(paragraph.Paragraph, relationshipMap);
                DocxObjectActions.RemapDrawingIds(outputBody, paragraphs.Select(p => p.Paragraph).Cast<OpenXmlElement>().ToArray());
                ReplaceTargetContent(targets[item.index], paragraphs, targetStyles);
            }
        }
    }

    private static void AppendSelection(
        List<PreparedParagraph> output,
        ref Paragraph? activeInlineSourceParagraph,
        OpenXmlElement element,
        Styles? sourceStyles,
        CopyContentSelection selection,
        IReadOnlySet<string>? preservedBookmarkIds)
    {
        if (selection.Range is not null)
        {
            if (selection.Range.Length == 0)
            {
                ValidateEmptyRange(element, selection.Range.Start);
                activeInlineSourceParagraph = null;
                return;
            }
            if (element is TableCell or Paragraph)
            {
                output.AddRange(ParagraphsForTextRange(
                    element,
                    sourceStyles,
                    selection.Range.Start,
                    selection.Range.Length));
                activeInlineSourceParagraph = null;
                return;
            }
            AppendInlineSelection(
                output,
                ref activeInlineSourceParagraph,
                SourceParagraph(element),
                sourceStyles,
                RunForTextRange(element, selection.Range.Start, selection.Range.Length));
            return;
        }

        switch (element)
        {
            case TableCell cell:
                output.AddRange(cell.Elements<Paragraph>().Select(paragraph =>
                    PrepareParagraph(paragraph, CloneParagraph(paragraph, preservedBookmarkIds), sourceStyles)));
                activeInlineSourceParagraph = null;
                return;
            case Paragraph paragraph:
                output.Add(PrepareParagraph(paragraph, CloneParagraph(paragraph, preservedBookmarkIds), sourceStyles));
                activeInlineSourceParagraph = null;
                return;
            case DocumentFormat.OpenXml.Wordprocessing.Run run:
                AppendInlineSelection(
                    output,
                    ref activeInlineSourceParagraph,
                    SourceParagraph(run),
                    sourceStyles,
                    (DocumentFormat.OpenXml.Wordprocessing.Run)run.CloneNode(true));
                return;
            case Text text:
                AppendInlineSelection(
                    output,
                    ref activeInlineSourceParagraph,
                    SourceParagraph(text),
                    sourceStyles,
                    RunForText(text));
                return;
            default:
                throw new InvalidOperationException($"source-ref-kind-not-supported-for-content-copy: {element.LocalName}");
        }
    }

    private static PreparedParagraph PrepareParagraph(Paragraph source, Paragraph cloned, Styles? sourceStyles)
    {
        var (isNumbered, level) = ResolveEffectiveNumbering(source, sourceStyles);
        return new PreparedParagraph(cloned, isNumbered, level);
    }

    private static void ValidateEmptyRange(OpenXmlElement element, int start)
    {
        var length = element switch
        {
            TableCell cell => cell.Elements<Paragraph>().Sum(paragraph => ScalarLength(NativeMutationSupport.PlainText(paragraph)))
                + Math.Max(0, cell.Elements<Paragraph>().Count() - 1),
            Paragraph paragraph => ScalarLength(NativeMutationSupport.PlainText(paragraph)),
            DocumentFormat.OpenXml.Wordprocessing.Run run => ScalarLength(run.InnerText),
            Text text => ScalarLength(text.Text),
            _ => throw new InvalidOperationException($"source-ref-kind-not-supported-for-content-copy: {element.LocalName}"),
        };
        if (start < 0 || start > length)
            throw new InvalidOperationException("text-range-out-of-bounds");
    }

    private static IReadOnlyList<PreparedParagraph> ParagraphsForTextRange(
        OpenXmlElement element,
        Styles? sourceStyles,
        int start,
        int length)
    {
        var paragraphs = element switch
        {
            TableCell cell => cell.Elements<Paragraph>().ToArray(),
            Paragraph paragraph => new[] { paragraph },
            _ => throw new InvalidOperationException("text-range-block-must-be-cell-or-paragraph"),
        };
        var lengths = paragraphs.Select(paragraph => ScalarLength(NativeMutationSupport.PlainText(paragraph))).ToArray();
        var totalLength = lengths.Sum() + Math.Max(0, paragraphs.Length - 1);
        if (start < 0 || length < 0 || start > totalLength - length)
            throw new InvalidOperationException("text-range-out-of-bounds");
        if (length == 0) return Array.Empty<PreparedParagraph>();

        var end = start + length;
        var cursor = 0;
        var selected = new List<PreparedParagraph>();
        for (var index = 0; index < paragraphs.Length; index++)
        {
            var paragraphEnd = cursor + lengths[index];
            var overlapStart = Math.Max(start, cursor);
            var overlapEnd = Math.Min(end, paragraphEnd);
            if (overlapStart < overlapEnd)
                selected.Add(PrepareParagraph(
                    paragraphs[index],
                    ParagraphForTextRange(
                        paragraphs[index],
                        overlapStart - cursor,
                        overlapEnd - overlapStart),
                    sourceStyles));
            cursor = paragraphEnd + 1;
        }
        if (selected.Count == 0)
            throw new InvalidOperationException("text-range-must-select-text");
        return selected;
    }

    private static Paragraph ParagraphForTextRange(Paragraph source, int start, int length)
    {
        var output = CloneParagraph(source);
        var cursor = 0;
        var end = start + length;
        foreach (var element in output.Descendants<OpenXmlElement>()
                     .Where(item => item is Text or Break or CarriageReturn or TabChar).ToArray())
        {
            var value = element is Text text ? text.Text : element is TabChar ? "\t" : "\n";
            var textLength = ScalarLength(value);
            var overlapStart = Math.Max(start, cursor);
            var overlapEnd = Math.Min(end, cursor + textLength);
            if (overlapStart >= overlapEnd)
            {
                element.Remove();
            }
            else if (element is Text selectedText)
            {
                selectedText.Text = SliceScalars(selectedText.Text, overlapStart - cursor, overlapEnd - overlapStart);
                selectedText.Space = SpaceProcessingModeValues.Preserve;
            }
            cursor += textLength;
        }
        foreach (var run in output.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
                     .Where(run => !run.Descendants<OpenXmlElement>()
                         .Any(item => item is Text or Break or CarriageReturn or TabChar)).ToArray())
            run.Remove();
        if (!output.Descendants<OpenXmlElement>().Any(item => item is Text or Break or CarriageReturn or TabChar))
            throw new InvalidOperationException("text-range-must-select-text");
        return output;
    }

    private static void AppendInlineSelection(
        List<PreparedParagraph> output,
        ref Paragraph? activeSourceParagraph,
        Paragraph sourceParagraph,
        Styles? sourceStyles,
        DocumentFormat.OpenXml.Wordprocessing.Run run)
    {
        if (!ReferenceEquals(activeSourceParagraph, sourceParagraph))
        {
            var (isNumbered, level) = ResolveEffectiveNumbering(sourceParagraph, sourceStyles);
            output.Add(new PreparedParagraph(new Paragraph(), isNumbered, level));
            activeSourceParagraph = sourceParagraph;
        }
        output[^1].Paragraph.Append(run);
    }

    private static Paragraph SourceParagraph(OpenXmlElement element)
        => element as Paragraph
            ?? element.Ancestors<Paragraph>().FirstOrDefault()
            ?? throw new InvalidOperationException("source-inline-ref-must-have-paragraph-parent");

    private static DocumentFormat.OpenXml.Wordprocessing.Run RunForText(Text text)
    {
        var sourceRun = text.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault()
            ?? throw new InvalidOperationException("source-text-ref-must-have-run-parent");
        var output = new DocumentFormat.OpenXml.Wordprocessing.Run();
        if (sourceRun.RunProperties is not null)
            output.RunProperties = (RunProperties)sourceRun.RunProperties.CloneNode(true);
        output.Append((Text)text.CloneNode(true));
        return output;
    }

    internal static Paragraph CloneParagraph(Paragraph paragraph, IReadOnlySet<string>? preservedBookmarkIds = null)
    {
        var clone = (Paragraph)paragraph.CloneNode(true);
        clone.RemoveAttribute("paraId", Word2010Namespace);
        clone.RemoveAttribute("textId", Word2010Namespace);
        clone.ParagraphProperties?.Remove();
        foreach (var bookmark in clone.Descendants<BookmarkStart>().Cast<OpenXmlElement>()
                     .Concat(clone.Descendants<BookmarkEnd>()).ToArray())
        {
            var id = bookmark switch
            {
                BookmarkStart start => start.Id?.Value,
                BookmarkEnd end => end.Id?.Value,
                _ => null,
            };
            if (id is null || preservedBookmarkIds?.Contains(id) != true)
                bookmark.Remove();
        }
        foreach (var cached in clone.Descendants<ProofError>().Cast<OpenXmlElement>()
                     .Concat(clone.Descendants<LastRenderedPageBreak>()).ToArray())
            cached.Remove();
        return clone;
    }

    private static DocumentFormat.OpenXml.Wordprocessing.Run RunForTextRange(
        OpenXmlElement element,
        int start,
        int length)
    {
        if (element is not DocumentFormat.OpenXml.Wordprocessing.Run and not Text)
            throw new InvalidOperationException("text-range-requires-run-or-text-ref");
        if (element is DocumentFormat.OpenXml.Wordprocessing.Run run && run.ChildElements.Any(child => child is not RunProperties and not Text))
            throw new InvalidOperationException("text-range-run-must-contain-only-text");
        var value = element is Text text ? text.Text : string.Concat(element.Descendants<Text>().Select(item => item.Text));
        var selected = SliceScalars(value, start, length);
        var sourceRun = element as DocumentFormat.OpenXml.Wordprocessing.Run
            ?? element.Ancestors<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault()
            ?? throw new InvalidOperationException("source-text-ref-must-have-run-parent");
        if (sourceRun.ChildElements.Any(child => child is not RunProperties and not Text))
            throw new InvalidOperationException("text-range-run-must-contain-only-text");
        var outputRun = (DocumentFormat.OpenXml.Wordprocessing.Run)sourceRun.CloneNode(true);
        outputRun.RemoveAllChildren<Text>();
        outputRun.Append(new Text(selected) { Space = SpaceProcessingModeValues.Preserve });
        return outputRun;
    }

    private static string SliceScalars(string value, int start, int length)
    {
        var runes = value.EnumerateRunes().ToArray();
        if (start < 0 || length < 0 || start > runes.Length - length)
            throw new InvalidOperationException("text-range-out-of-bounds");
        var builder = new StringBuilder();
        foreach (var rune in runes.Skip(start).Take(length)) builder.Append(rune.ToString());
        return builder.ToString();
    }

    private static int ScalarLength(string value) => value.EnumerateRunes().Count();

    internal static void ReplaceCellContent(
        TableCell target,
        IReadOnlyList<PreparedParagraph> sourceParagraphs,
        Styles? targetStyles = null)
    {
        var targetParagraphs = target.Elements<Paragraph>().ToArray();
        var targetNumbering = targetParagraphs
            .Select(paragraph => ResolveEffectiveNumbering(paragraph, targetStyles).IsNumbered)
            .ToArray();
        var firstTarget = targetParagraphs.FirstOrDefault();
        var firstNumberedTarget = targetParagraphs
            .Where((_, index) => targetNumbering[index])
            .FirstOrDefault();
        var firstUnnumberedTarget = targetParagraphs
            .Where((_, index) => !targetNumbering[index])
            .FirstOrDefault();
        var defaultRunProperties = firstTarget?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
            .FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

        foreach (var child in target.ChildElements.Where(child => child is not TableCellProperties).ToArray())
            child.Remove();

        for (var index = 0; index < sourceParagraphs.Count; index++)
        {
            var source = sourceParagraphs[index];
            var template = SelectTargetParagraphTemplate(
                targetParagraphs,
                targetNumbering,
                index,
                source.IsNumbered,
                firstTarget,
                firstNumberedTarget,
                firstUnnumberedTarget);
            var paragraphProperties = template?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
            if (!source.IsNumbered
                && template is not null
                && ResolveEffectiveNumbering(template, targetStyles).IsNumbered)
            {
                paragraphProperties = SuppressListNumbering(paragraphProperties, targetStyles);
            }
            var templateRunProperties = template?.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
                .FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties
                ?? defaultRunProperties;

            var paragraph = (Paragraph)source.Paragraph.CloneNode(true);
            paragraph.ParagraphProperties?.Remove();
            if (paragraphProperties is not null)
                paragraph.PrependChild(paragraphProperties);
            foreach (var run in paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                run.RunProperties = MergeRunProperties(templateRunProperties, run.RunProperties);
            target.Append(paragraph);
        }
    }

    private static Paragraph? SelectTargetParagraphTemplate(
        Paragraph[] targetParagraphs,
        bool[] targetNumbering,
        int index,
        bool sourceIsNumbered,
        Paragraph? firstTarget,
        Paragraph? firstNumberedTarget,
        Paragraph? firstUnnumberedTarget)
    {
        if (index < targetParagraphs.Length && targetNumbering[index] == sourceIsNumbered)
            return targetParagraphs[index];
        return sourceIsNumbered
            ? firstNumberedTarget ?? firstTarget
            : firstUnnumberedTarget ?? firstTarget;
    }

    private static ParagraphProperties? SuppressListNumbering(ParagraphProperties? properties, Styles? targetStyles)
    {
        var result = properties ?? new ParagraphProperties();
        result.NumberingProperties = null;
        if (StyleHasActiveNumbering(result.ParagraphStyleId?.Val?.Value, targetStyles))
            result.NumberingProperties = new NumberingProperties(new NumberingId { Val = 0 });
        if (result.Indentation?.Hanging is not null)
        {
            result.Indentation.Hanging = null;
            if (!result.Indentation.HasAttributes && !result.Indentation.HasChildren)
                result.Indentation = null;
        }
        return result.ChildElements.Count == 0 && !result.HasAttributes ? null : result;
    }

    internal static (bool IsNumbered, int? Level) ResolveEffectiveNumbering(Paragraph paragraph, Styles? styles)
    {
        var direct = paragraph.ParagraphProperties?.NumberingProperties;
        if (direct?.NumberingId?.Val?.Value is int directNumId)
        {
            if (directNumId == 0) return (false, null);
            return (true, direct.NumberingLevelReference?.Val?.Value ?? 0);
        }
        return ResolveStyleNumbering(paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, styles);
    }

    private static bool StyleHasActiveNumbering(string? styleId, Styles? styles)
        => ResolveStyleNumbering(styleId, styles).IsNumbered;

    private static (bool IsNumbered, int? Level) ResolveStyleNumbering(string? styleId, Styles? styles)
    {
        if (string.IsNullOrWhiteSpace(styleId) || styles is null) return (false, null);
        var styleMap = styles.Elements<Style>()
            .Where(style => !string.IsNullOrWhiteSpace(style.StyleId?.Value))
            .GroupBy(style => style.StyleId!.Value!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var currentId = styleId;
        while (!string.IsNullOrWhiteSpace(currentId) && visited.Add(currentId))
        {
            if (!styleMap.TryGetValue(currentId, out var style)) break;
            var numbering = style.StyleParagraphProperties?.NumberingProperties;
            if (numbering?.NumberingId?.Val?.Value is int styleNumId)
            {
                if (styleNumId == 0) return (false, null);
                return (true, numbering.NumberingLevelReference?.Val?.Value ?? 0);
            }
            currentId = style.BasedOn?.Val?.Value;
        }
        return (false, null);
    }

    private static void ReplaceTargetContent(
        OpenXmlElement target,
        IReadOnlyList<PreparedParagraph> sourceParagraphs,
        Styles? targetStyles)
    {
        if (target is TableCell cell)
        {
            ReplaceCellContent(cell, sourceParagraphs, targetStyles);
            return;
        }
        if (target is not Paragraph paragraph)
            throw new InvalidOperationException("target-ref-must-be-main-document-paragraph-or-cell");
        var targetRunProperties = paragraph.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>()
            .FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        foreach (var child in paragraph.ChildElements.Where(child => child is not ParagraphProperties).ToArray())
            child.Remove();
        var nextIndex = paragraph.ChildElements.Count;
        foreach (var source in sourceParagraphs)
        {
            foreach (var child in source.Paragraph.ChildElements.Where(child => child is not ParagraphProperties))
            {
                var clone = child.CloneNode(true);
                foreach (var run in clone.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                    run.RunProperties = MergeRunProperties(targetRunProperties, run.RunProperties);
                if (clone is DocumentFormat.OpenXml.Wordprocessing.Run directRun)
                    directRun.RunProperties = MergeRunProperties(targetRunProperties, directRun.RunProperties);
                paragraph.InsertAt(clone, nextIndex++);
            }
        }
    }

    private static RunProperties? MergeRunProperties(RunProperties? target, RunProperties? source)
    {
        var result = target?.CloneNode(true) as RunProperties ?? new RunProperties();
        if (source is not null)
        {
            foreach (var semantic in source.ChildElements.Where(element => element is Bold or Italic or Underline
                or Strike or DoubleStrike or Caps or SmallCaps or VerticalTextAlignment).ToArray())
            {
                SetRunProperty(result, semantic);
            }
        }
        return result.ChildElements.Count == 0 ? null : result;
    }

    private static void SetRunProperty(RunProperties properties, OpenXmlElement property)
    {
        switch (property)
        {
            case Bold value:
                properties.Bold = (Bold)value.CloneNode(true);
                break;
            case Italic value:
                properties.Italic = (Italic)value.CloneNode(true);
                break;
            case Underline value:
                properties.Underline = (Underline)value.CloneNode(true);
                break;
            case Strike value:
                properties.Strike = (Strike)value.CloneNode(true);
                break;
            case DoubleStrike value:
                properties.DoubleStrike = (DoubleStrike)value.CloneNode(true);
                break;
            case Caps value:
                properties.Caps = (Caps)value.CloneNode(true);
                break;
            case SmallCaps value:
                properties.SmallCaps = (SmallCaps)value.CloneNode(true);
                break;
            case VerticalTextAlignment value:
                properties.VerticalTextAlignment = (VerticalTextAlignment)value.CloneNode(true);
                break;
            default:
                throw new InvalidOperationException("run-property-not-supported-for-content-copy");
        }
    }

    private static IReadOnlyList<CopyContentReadback> ReadBack(
        string output,
        IReadOnlyList<PreparedChange> changes)
    {
        using var document = WordprocessingDocument.Open(output, false);
        return changes.Select(change =>
        {
            var target = Observation.ResolveNativePath(document, change.Target.StoryPart, change.Target.NativePath);
            var expected = string.Concat(change.Paragraphs.Select(paragraph => paragraph.Paragraph.InnerText));
            if (!StringComparer.Ordinal.Equals(target.InnerText, expected))
                throw new InvalidOperationException("output-readback-content-mismatch");
            return new CopyContentReadback(
                change.Target.Address,
                target.InnerText);
        }).ToArray();
    }

    private static IReadOnlyDictionary<string, int> ValidationIssueCounts(WordprocessingDocument document)
        => new OpenXmlValidator().Validate(document)
            .GroupBy(issue => $"{issue.Id}\0{issue.Description}", StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

    private static CopyContentArtifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new CopyContentArtifact(
            Path.GetFullPath(path),
            Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(),
            stream.Length);
    }

    internal sealed record PreparedParagraph(Paragraph Paragraph, bool IsNumbered, int? NumberingLevel);

    private sealed record PreparedChange(
        ResolvedDocxAddress Target,
        string SourcePath,
        IReadOnlyList<PreparedParagraph> Paragraphs);
}

public sealed record CopyContentSelection(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    CopyContentRange? Range = null);
public sealed record CopyContentRange(int Start, int Length);
public sealed record CopyContentChange(
    DocxObjectAddress Target,
    string SourceInput,
    IReadOnlyList<CopyContentSelection> SourceSelections);
public sealed record CopyContentRequest(
    string Input,
    IReadOnlyList<CopyContentChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record CopyContentArtifact(string Path, string Sha256, long Bytes);
public sealed record CopyContentReadback(DocxObjectAddress Address, string Text);
public sealed record CopyContentReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<CopyContentReadback> Changes,
    string Output);
