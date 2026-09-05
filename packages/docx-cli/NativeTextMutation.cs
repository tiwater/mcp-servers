using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class NativeTextMutation
{
    public const string Command = "docx_set_text";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetTextRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-text-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static SetTextReceipt Apply(SetTextRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        for (var index = 0; index < request.Changes.Count; index++)
        {
            var change = request.Changes[index];
            if (change.FontName is not null && (string.IsNullOrWhiteSpace(change.FontName) || change.Text.Length == 0))
                throw new InvalidOperationException($"font-name-requires-nonempty-text: changes[{index}]");
        }
        var addresses = request.Changes.Select(change => change.Target).ToArray();
        var duplicate = request.Changes
            .Select((change, index) => new { change.Target, Index = index })
            .GroupBy(item => item.Target)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null)
            throw new InvalidOperationException(
                $"target-address-duplicate: changes=[{string.Join(',', duplicate.Select(item => item.Index))}]");
        var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var resolved = Observation.ResolveAddresses(paths.Input, addresses, "changes.target");
        for (var index = 0; index < resolved.Count; index++)
            if (resolved[index].Kind is not "paragraph" and not "cell" and not "text")
                throw new InvalidOperationException(
                    $"target-must-be-paragraph-cell-or-text: changes[{index}].target; kind={resolved[index].Kind}");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            var targets = resolved.Select(item =>
                Observation.ResolveNativePath(input, item.StoryPart, item.NativePath)).ToArray();
            NativeMutationSupport.RejectOverlappingTargets(targets);
            for (var index = 0; index < targets.Length; index++)
            {
                if (targets[index] is Paragraph) continue;
                try
                {
                    NativeMutationSupport.RequirePlainTextContainer(targets[index]);
                }
                catch (InvalidOperationException error)
                {
                    throw new InvalidOperationException(
                        $"{error.Message}: changes[{index}].target", error);
                }
            }
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                for (var index = 0; index < resolved.Count; index++)
                {
                    var target = Observation.ResolveNativePath(output, resolved[index].StoryPart, resolved[index].NativePath);
                    SetText(target, request.Changes[index].Text, request.Changes[index].FontName);
                }
                SaveChangedStories(output, resolved.Select(item => item.StoryPart));
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);
            IReadOnlyList<SetTextReadback> readback;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                readback = resolved.Select(item =>
                {
                    var target = Observation.ResolveNativePath(output, item.StoryPart, item.NativePath);
                    var text = NativeMutationSupport.PlainText(target);
                    return new SetTextReadback(
                        item.Address,
                        item.Kind,
                        text,
                        ReadFontName(target));
                }).ToArray();
            }
            for (var index = 0; index < readback.Count; index++)
            {
                if (!StringComparer.Ordinal.Equals(readback[index].Text, request.Changes[index].Text))
                    throw new InvalidOperationException("output-readback-content-mismatch");
                if (request.Changes[index].FontName is not null
                    && !StringComparer.Ordinal.Equals(readback[index].FontName, request.Changes[index].FontName))
                    throw new InvalidOperationException("output-readback-font-mismatch");
            }
            var receipt = new SetTextReceipt(
                "tiwater.docx-set-text-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                readback,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    internal static void SetText(OpenXmlElement target, string text, string? fontName = null)
    {
        switch (target)
        {
            case Text textNode:
                textNode.Text = text;
                textNode.Space = SpaceProcessingModeValues.Preserve;
                if (fontName is not null)
                {
                    var textRun = textNode.Ancestors<Run>().First();
                    ApplyFontName(textRun.RunProperties ??= new RunProperties(), fontName);
                }
                break;
            case Paragraph paragraph:
                var runProperties = paragraph.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
                var paragraphBookmarkStarts = paragraph.Elements<BookmarkStart>().Select(bookmark => (BookmarkStart)bookmark.CloneNode(true)).ToArray();
                var paragraphBookmarkEnds = paragraph.Elements<BookmarkEnd>().Select(bookmark => (BookmarkEnd)bookmark.CloneNode(true)).ToArray();
                foreach (var child in paragraph.ChildElements.Where(child => child is not ParagraphProperties).ToArray()) child.Remove();
                foreach (var bookmark in paragraphBookmarkStarts) paragraph.Append(bookmark);
                if (fontName is not null) ApplyFontName(runProperties ??= new RunProperties(), fontName);
                var replacementRun = TextRun(runProperties, text);
                if (replacementRun is not null) paragraph.Append(replacementRun);
                foreach (var bookmark in paragraphBookmarkEnds) paragraph.Append(bookmark);
                break;
            case TableCell cell:
                var template = cell.Elements<Paragraph>().FirstOrDefault();
                var paragraphProperties = template?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
                var cellRunProperties = template?.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
                var bookmarkStarts = cell.Elements<Paragraph>()
                    .SelectMany(paragraph => paragraph.Elements<BookmarkStart>())
                    .Select(bookmark => (BookmarkStart)bookmark.CloneNode(true))
                    .ToArray();
                var bookmarkEnds = cell.Elements<Paragraph>()
                    .SelectMany(paragraph => paragraph.Elements<BookmarkEnd>())
                    .Select(bookmark => (BookmarkEnd)bookmark.CloneNode(true))
                    .ToArray();
                foreach (var child in cell.ChildElements.Where(child => child is not TableCellProperties).ToArray())
                    child.Remove();
                var replacement = new Paragraph();
                if (paragraphProperties is not null) replacement.Append((ParagraphProperties)paragraphProperties.CloneNode(true));
                foreach (var bookmark in bookmarkStarts) replacement.Append(bookmark);
                if (fontName is not null) ApplyFontName(cellRunProperties ??= new RunProperties(), fontName);
                AppendText(replacement, cellRunProperties, text);
                foreach (var bookmark in bookmarkEnds) replacement.Append(bookmark);
                cell.Append(replacement);
                break;
            default:
                throw new InvalidOperationException("target-ref-must-be-paragraph-or-cell");
        }
    }

    internal static void SetTextRuns(TableCell cell, IReadOnlyList<SetTableTextRun> textRuns)
    {
        var template = cell.Elements<Paragraph>().FirstOrDefault();
        var paragraphProperties = template?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var baseRunProperties = template?.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
        var bookmarkStarts = cell.Elements<Paragraph>()
            .SelectMany(paragraph => paragraph.Elements<BookmarkStart>())
            .Select(bookmark => (BookmarkStart)bookmark.CloneNode(true)).ToArray();
        var bookmarkEnds = cell.Elements<Paragraph>()
            .SelectMany(paragraph => paragraph.Elements<BookmarkEnd>())
            .Select(bookmark => (BookmarkEnd)bookmark.CloneNode(true)).ToArray();
        foreach (var child in cell.ChildElements.Where(child => child is not TableCellProperties).ToArray()) child.Remove();
        var paragraph = new Paragraph();
        if (paragraphProperties is not null) paragraph.Append(paragraphProperties);
        foreach (var bookmark in bookmarkStarts) paragraph.Append(bookmark);
        foreach (var textRun in textRuns)
        {
            var properties = baseRunProperties?.CloneNode(true) as RunProperties ?? new RunProperties();
            properties.RemoveAllChildren<Color>();
            properties.RemoveAllChildren<Underline>();
            var color = NativeSetTableMutation.ColorValue(textRun.Color);
            var underline = NativeSetTableMutation.UnderlineValue(textRun.Underline);
            if (color is not null) properties.Append(new Color { Val = color.ToUpperInvariant() });
            if (underline is not null)
                properties.Append(new Underline
                {
                    Val = underline == "double" ? UnderlineValues.Double : UnderlineValues.Single,
                });
            AppendText(paragraph, properties, textRun.Text);
        }
        foreach (var bookmark in bookmarkEnds) paragraph.Append(bookmark);
        cell.Append(paragraph);
    }

    private static void ApplyFontName(RunProperties properties, string? fontName)
    {
        if (fontName is null) return;
        if (string.IsNullOrWhiteSpace(fontName)) throw new InvalidOperationException("font-name-must-not-be-empty");
        var fonts = properties.RunFonts ?? new RunFonts();
        fonts.Ascii = fontName;
        fonts.HighAnsi = fontName;
        fonts.ComplexScript = fontName;
        if (fonts.Parent is null) properties.AddChild(fonts, true);
    }

    private static string? ReadFontName(OpenXmlElement target)
    {
        var runs = target is Text text
            ? text.Ancestors<Run>().Take(1).ToArray()
            : target.Descendants<Run>().Where(run => !string.IsNullOrEmpty(run.InnerText)).ToArray();
        if (runs.Length == 0) return null;
        var names = runs.Select(run =>
        {
            var fonts = run.RunProperties?.RunFonts;
            return fonts?.Ascii?.Value is { } ascii
                && StringComparer.Ordinal.Equals(fonts.HighAnsi?.Value, ascii)
                && StringComparer.Ordinal.Equals(fonts.ComplexScript?.Value, ascii) ? ascii : null;
        }).Distinct(StringComparer.Ordinal).ToArray();
        return names.Length == 1 ? names[0] : null;
    }

    private static void AppendText(OpenXmlCompositeElement parent, RunProperties? properties, string text)
    {
        var run = TextRun(properties, text);
        if (run is not null) parent.Append(run);
    }

    private static Run? TextRun(RunProperties? properties, string text)
    {
        if (text.Length == 0) return null;
        var run = new Run();
        if (properties is not null) run.Append((RunProperties)properties.CloneNode(true));
        var start = 0;
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character is not '\r' and not '\n' and not '\t') continue;
            if (index > start)
                run.Append(new Text(text[start..index]) { Space = SpaceProcessingModeValues.Preserve });
            if (character == '\t')
                run.Append(new TabChar());
            else
            {
                run.Append(new Break());
                if (character == '\r' && index + 1 < text.Length && text[index + 1] == '\n') index++;
            }
            start = index + 1;
        }
        if (start < text.Length)
            run.Append(new Text(text[start..]) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static void SaveChangedStories(WordprocessingDocument document, IEnumerable<string> storyParts)
    {
        var parts = storyParts.Distinct(StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
        if (main.Document is not null && parts.Contains(PartUri(main.Uri))) main.Document.Save();
        foreach (var part in main.HeaderParts.Where(part => part.Header is not null && parts.Contains(PartUri(part.Uri)))) part.Header!.Save();
        foreach (var part in main.FooterParts.Where(part => part.Footer is not null && parts.Contains(PartUri(part.Uri)))) part.Footer!.Save();
        if (main.FootnotesPart?.Footnotes is not null && parts.Contains(PartUri(main.FootnotesPart.Uri))) main.FootnotesPart.Footnotes.Save();
        if (main.EndnotesPart?.Endnotes is not null && parts.Contains(PartUri(main.EndnotesPart.Uri))) main.EndnotesPart.Endnotes.Save();
        if (main.WordprocessingCommentsPart?.Comments is not null && parts.Contains(PartUri(main.WordprocessingCommentsPart.Uri))) main.WordprocessingCommentsPart.Comments.Save();
    }

    private static string PartUri(Uri uri)
        => uri.OriginalString.StartsWith("/", StringComparison.Ordinal) ? uri.OriginalString : "/" + uri.OriginalString;
}

public sealed record SetTextChange(DocxObjectAddress Target, string Text, string? FontName = null);
public sealed record SetTextRequest(string Input, IReadOnlyList<SetTextChange> Changes, string Output, string ReceiptOutput);
public sealed record SetTextReadback(DocxObjectAddress Address, string Kind, string Text, string? FontName);
public sealed record SetTextReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<SetTextReadback> Changes,
    string Output);
