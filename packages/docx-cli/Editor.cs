using System.Text.Json;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

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

        var formatOperationTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            "setParagraphFormat",
            "setRunFormat",
            "setSectionFormat",
            "setHeaderFooterParagraphFormat"
        };
        if (operations.Any(operation => !formatOperationTypes.Contains(operation.Type)))
        {
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
        }
        else if (applied.Any(operation => operation.Applied && operation.Type is not "setHeaderFooterParagraphFormat"))
        {
            mainPart.Document.Save();
        }
        else if (!applied.Any(operation => operation.Applied))
        {
            doc.Dispose();
            File.Copy(input, output, overwrite: true);
        }
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
            "replaceBodyText" => ReplaceBodyText(body, operation),
            "startSectionBeforeParagraph" => StartSectionBeforeParagraph(body, operation),
            "replaceAllHeaderParagraphText" => ReplaceAllHeaderParagraphText(doc, operation),
            "replaceHeaderParagraphText" => ReplaceHeaderParagraphText(doc, operation),
            "replaceHeaderText" => ReplaceHeaderText(doc, operation),
            "replaceTableCellText" => ReplaceTableCellText(body, operation),
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
            "setTableRowHeight" => SetTableRowHeight(body, operation),
            "setTableRowCantSplit" => SetTableRowCantSplit(body, operation),
            "mergeTableCells" => MergeTableCells(body, operation),
            "unmergeTableRowHorizontalCells" => UnmergeTableRowHorizontalCells(body, operation),
            "unmergeTableColumnVerticalCells" => UnmergeTableColumnVerticalCells(body, operation),
            "fillTableSemantically" => FillTableSemantically(body, operation),
            "setParagraphFormat" => SetParagraphFormat(body, operation),
            "setRunFormat" => SetRunFormat(body, operation),
            "setSectionFormat" => SetSectionFormat(body, operation),
            "setHeaderFooterParagraphFormat" => SetHeaderFooterParagraphFormat(doc, body, operation),
            "deleteComment" => DeleteComments(doc, operation.CommentId is { Length: > 0 } id ? [id] : []),
            "deleteComments" => DeleteComments(doc, operation.CommentIds ?? []),
            "markFieldsDirty" => MarkFieldsDirty(doc),
            "sanitizeFields" => SanitizeFields(doc),
            "freezeFields" => FreezeFields(doc),
            _ => new DocxEditAppliedOperation(operation.Type, false, $"Unknown operation type: {operation.Type}"),
        };
    }

    private sealed record ResolvedFormatElement(OpenXmlElement Owner, OpenXmlElement? Properties, DocxResolvedFormatTarget Target);

    private static readonly IReadOnlySet<string> ParagraphFormatProperties = new HashSet<string>(StringComparer.Ordinal)
    {
        "alignment",
        "spacingBeforeTwips",
        "spacingAfterTwips",
        "lineSpacingTwips",
        "lineSpacingRule",
        "keepWithNext",
        "keepLines",
        "pageBreakBefore",
        "widowControl"
    };

    private static readonly IReadOnlySet<string> RunFormatProperties = new HashSet<string>(StringComparer.Ordinal)
    {
        "fontAscii",
        "fontHighAnsi",
        "fontEastAsia",
        "fontComplexScript",
        "fontSizeHalfPoints",
        "bold",
        "italic",
        "underline"
    };

    private static readonly IReadOnlySet<string> SectionFormatProperties = new HashSet<string>(StringComparer.Ordinal)
    {
        "orientation",
        "marginTopTwips",
        "marginRightTwips",
        "marginBottomTwips",
        "marginLeftTwips",
        "headerDistanceTwips",
        "footerDistanceTwips",
        "gutterTwips"
    };

    private static DocxEditAppliedOperation SetParagraphFormat(Body body, DocxEditOperation operation)
    {
        var error = ValidateFormatOperation(operation, "paragraph", ParagraphFormatProperties);
        if (error is not null)
        {
            return FormatFailure(operation, error);
        }

        var resolved = ResolveBodyParagraph(body, operation.FormatTarget!);
        if (resolved.Error is not null)
        {
            return FormatFailure(operation, resolved.Error);
        }

        return ApplyResolvedFormat(operation, resolved.Value!, properties =>
            ApplyParagraphProperties((Paragraph)resolved.Value!.Owner, properties));
    }

    private static DocxEditAppliedOperation SetRunFormat(Body body, DocxEditOperation operation)
    {
        var error = ValidateFormatOperation(operation, "run", RunFormatProperties);
        if (error is not null)
        {
            return FormatFailure(operation, error);
        }

        var target = operation.FormatTarget!;
        var paragraph = ResolveBodyParagraph(body, target);
        if (paragraph.Error is not null)
        {
            return FormatFailure(operation, paragraph.Error);
        }
        if (string.IsNullOrEmpty(target.RunText) || target.RunOccurrence is null || target.RunOccurrence < 0)
        {
            return FormatFailure(operation, "runText and non-negative runOccurrence are required semantic run selectors");
        }

        var runs = ((Paragraph)paragraph.Value!.Owner).Descendants<Run>()
            .Where(run => string.Equals(GetRunText(run), target.RunText, StringComparison.Ordinal))
            .ToList();
        if (target.RunOccurrence.Value >= runs.Count)
        {
            return FormatFailure(operation, $"Run target '{target.RunText}' occurrence {target.RunOccurrence} was not found");
        }

        var run = runs[target.RunOccurrence.Value];
        var allRuns = ((Paragraph)paragraph.Value.Owner).Descendants<Run>().ToList();
        var runIndex = allRuns.IndexOf(run);
        var exactId = $"{paragraph.Value.Target.ExactId}:r{runIndex}";
        var resolved = new ResolvedFormatElement(
            run,
            run.RunProperties,
            new DocxResolvedFormatTarget("run", exactId, paragraph.Value.Target.ParagraphId, runIndex));
        return ApplyResolvedFormat(operation, resolved, properties => ApplyRunProperties(run, properties));
    }

    private static DocxEditAppliedOperation SetSectionFormat(Body body, DocxEditOperation operation)
    {
        var error = ValidateFormatOperation(operation, "section", SectionFormatProperties);
        if (error is not null)
        {
            return FormatFailure(operation, error);
        }

        var resolved = ResolveSection(body, operation.FormatTarget!);
        if (resolved.Error is not null)
        {
            return FormatFailure(operation, resolved.Error);
        }

        return ApplyResolvedFormat(operation, resolved.Value!, properties =>
            ApplySectionProperties((SectionProperties)resolved.Value!.Owner, properties));
    }

    private static DocxEditAppliedOperation SetHeaderFooterParagraphFormat(
        WordprocessingDocument doc,
        Body body,
        DocxEditOperation operation)
    {
        var error = ValidateFormatOperation(operation, "headerFooterParagraph", ParagraphFormatProperties);
        if (error is not null)
        {
            return FormatFailure(operation, error);
        }

        var resolved = ResolveHeaderFooterParagraph(doc, body, operation.FormatTarget!);
        if (resolved.Error is not null)
        {
            return FormatFailure(operation, resolved.Error);
        }

        var result = ApplyResolvedFormat(operation, resolved.Value!, properties =>
            ApplyParagraphProperties((Paragraph)resolved.Value!.Owner, properties));
        if (result.Applied)
        {
            var paragraph = (Paragraph)resolved.Value!.Owner;
            paragraph.Ancestors<Header>().FirstOrDefault()?.Save();
            paragraph.Ancestors<Footer>().FirstOrDefault()?.Save();
        }
        return result;
    }

    private static string? ValidateFormatOperation(
        DocxEditOperation operation,
        string expectedKind,
        IReadOnlySet<string> allowlist)
    {
        if (operation.FormatTarget is null || !string.Equals(operation.FormatTarget.Kind, expectedKind, StringComparison.Ordinal))
        {
            return $"formatTarget.kind must be '{expectedKind}'";
        }
        if (string.IsNullOrWhiteSpace(operation.ExpectedCurrentFormatHash)
            || operation.ExpectedCurrentFormatHash.Length != 64
            || !operation.ExpectedCurrentFormatHash.All(Uri.IsHexDigit))
        {
            return "expectedCurrentFormatHash must be a 64-character SHA-256 hex digest";
        }
        if (operation.FormatProperties is null || operation.FormatProperties.Count == 0)
        {
            return "formatProperties must declare at least one property";
        }
        if (operation.Text is not null || operation.FindText is not null || operation.RichText is not null
            || operation.Rows is not null || operation.Cells is not null)
        {
            return "format operations must not contain replacement text payloads";
        }
        if (operation.CommentId is not null || operation.HeaderIndex is not null || operation.ParagraphIndex is not null
            || operation.TableIndex is not null || operation.RowIndex is not null || operation.CellIndex is not null
            || operation.CommentIds is not null || operation.StartCellIndex is not null || operation.EndCellIndex is not null
            || operation.StartRowIndex is not null || operation.EndRowIndex is not null || operation.TemplateRowIndex is not null
            || operation.ColumnIndex is not null || operation.ColumnCount is not null || operation.TemplateColumnIndex is not null
            || operation.Alignment is not null || operation.Width is not null || operation.WidthType is not null
            || operation.Orientation is not null || operation.FontSize is not null || operation.Height is not null
            || operation.HeightRule is not null || operation.NoWrap is not null || operation.CantSplit is not null)
        {
            return "format operations must not contain legacy physical selectors or undeclared edit fields";
        }

        var unknown = operation.FormatProperties.Keys.FirstOrDefault(key => !allowlist.Contains(key));
        if (unknown is not null)
        {
            return $"Unknown format property '{unknown}' for {operation.Type}";
        }

        return ValidateFormatPropertyValues(operation.Type, operation.FormatProperties);
    }

    private static string? ValidateFormatPropertyValues(string operationType, IReadOnlyDictionary<string, string> properties)
    {
        foreach (var (name, value) in properties)
        {
            switch (name)
            {
                case "alignment" when value is not ("left" or "center" or "right" or "both" or "distribute"):
                    return $"Invalid alignment '{value}'";
                case "spacingBeforeTwips":
                case "spacingAfterTwips":
                case "lineSpacingTwips":
                case "fontSizeHalfPoints":
                case "marginTopTwips":
                case "marginRightTwips":
                case "marginBottomTwips":
                case "marginLeftTwips":
                case "headerDistanceTwips":
                case "footerDistanceTwips":
                case "gutterTwips":
                    if (!uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var numeric) || numeric > 31680U)
                    {
                        return $"Invalid unsigned twip/half-point value '{value}' for {name}";
                    }
                    if (name == "fontSizeHalfPoints" && numeric == 0U)
                    {
                        return "fontSizeHalfPoints must be greater than zero";
                    }
                    break;
                case "keepWithNext":
                case "keepLines":
                case "pageBreakBefore":
                case "widowControl":
                case "bold":
                case "italic":
                    if (!bool.TryParse(value, out _))
                    {
                        return $"Invalid boolean value '{value}' for {name}";
                    }
                    break;
                case "lineSpacingRule" when value is not ("auto" or "atLeast" or "exact"):
                    return $"Invalid lineSpacingRule '{value}'";
                case "underline" when value is not ("none" or "single" or "double" or "words"):
                    return $"Invalid underline '{value}'";
                case "orientation" when value is not ("portrait" or "landscape"):
                    return $"Invalid orientation '{value}'";
                case "fontAscii":
                case "fontHighAnsi":
                case "fontEastAsia":
                case "fontComplexScript":
                    if (string.IsNullOrWhiteSpace(value) || value.Length > 255)
                    {
                        return $"Invalid font name for {name}";
                    }
                    break;
            }
        }

        return null;
    }

    private sealed record Resolution(ResolvedFormatElement? Value, string? Error);

    private static Resolution ResolveBodyParagraph(Body body, DocxFormatTarget target)
    {
        if (string.IsNullOrEmpty(target.ParagraphText) || target.ParagraphOccurrence is null || target.ParagraphOccurrence < 0)
        {
            return new Resolution(null, "paragraphText and non-negative paragraphOccurrence are required semantic selectors; physical ids alone are not accepted");
        }

        var allParagraphs = body.Descendants<Paragraph>().ToList();
        var matches = allParagraphs.Where(paragraph => string.Equals(GetNormalizedParagraphText(paragraph), target.ParagraphText, StringComparison.Ordinal)).ToList();
        if (target.ParagraphOccurrence.Value >= matches.Count)
        {
            return new Resolution(null, $"Paragraph target '{target.ParagraphText}' occurrence {target.ParagraphOccurrence} was not found");
        }

        var paragraph = matches[target.ParagraphOccurrence.Value];
        var paragraphId = $"body-p{allParagraphs.IndexOf(paragraph)}";
        if (target.ParagraphId is not null && !string.Equals(target.ParagraphId, paragraphId, StringComparison.Ordinal))
        {
            return new Resolution(null, $"Paragraph identity mismatch: resolved {paragraphId}, requested {target.ParagraphId}");
        }

        return new Resolution(
            new ResolvedFormatElement(
                paragraph,
                paragraph.ParagraphProperties,
                new DocxResolvedFormatTarget("paragraph", paragraphId, paragraphId)),
            null);
    }

    private static Resolution ResolveSection(Body body, DocxFormatTarget target)
    {
        if (string.IsNullOrEmpty(target.SectionId)
            || string.IsNullOrEmpty(target.ParagraphText)
            || target.ParagraphOccurrence is null
            || target.ParagraphOccurrence < 0)
        {
            return new Resolution(null, "sectionId, paragraphText, and non-negative paragraphOccurrence are required semantic section selectors");
        }
        if (!TryParseStableId(target.SectionId, "section", out var requestedSection))
        {
            return new Resolution(null, $"Invalid section identity '{target.SectionId}'");
        }

        var paragraphs = body.Elements<Paragraph>().ToList();
        var sections = body.Elements<Paragraph>()
            .Select(paragraph => (Properties: paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>(), Paragraph: paragraph))
            .Where(item => item.Properties is not null)
            .Select(item => (item.Properties!, (Paragraph?)item.Paragraph))
            .Concat(body.Elements<SectionProperties>().Select(properties => (properties, (Paragraph?)null)))
            .ToList();
        if (requestedSection < 0 || requestedSection >= sections.Count)
        {
            return new Resolution(null, $"Section target '{target.SectionId}' was not found");
        }

        var start = requestedSection == 0
            ? 0
            : sections[requestedSection - 1].Item2 is Paragraph previousEnd ? paragraphs.IndexOf(previousEnd) + 1 : paragraphs.Count;
        var end = sections[requestedSection].Item2 is Paragraph currentEnd ? paragraphs.IndexOf(currentEnd) : paragraphs.Count - 1;
        var matches = paragraphs.Skip(start).Take(Math.Max(0, end - start + 1))
            .Where(paragraph => string.Equals(GetNormalizedParagraphText(paragraph), target.ParagraphText, StringComparison.Ordinal))
            .ToList();
        if (target.ParagraphOccurrence.Value >= matches.Count)
        {
            return new Resolution(null, $"Semantic paragraph '{target.ParagraphText}' occurrence {target.ParagraphOccurrence} does not belong to {target.SectionId}");
        }

        var section = sections[requestedSection].Item1;
        return new Resolution(
            new ResolvedFormatElement(
                section,
                section,
                new DocxResolvedFormatTarget("section", target.SectionId, SectionId: target.SectionId)),
            null);
    }

    private static Resolution ResolveHeaderFooterParagraph(WordprocessingDocument doc, Body body, DocxFormatTarget target)
    {
        if (string.IsNullOrEmpty(target.SectionId) || string.IsNullOrEmpty(target.PartId)
            || string.IsNullOrEmpty(target.ParagraphId) || string.IsNullOrEmpty(target.ParagraphText)
            || target.ParagraphOccurrence is null || target.ParagraphOccurrence < 0)
        {
            return new Resolution(null, "sectionId, partId, paragraphId, paragraphText, and non-negative paragraphOccurrence are required header/footer selectors");
        }
        if (!TryParseStableId(target.SectionId, "section", out var requestedSection))
        {
            return new Resolution(null, $"Invalid section identity '{target.SectionId}'");
        }

        var mainPart = doc.MainDocumentPart!;
        var sectionProperties = body.Elements<Paragraph>()
            .Select(paragraph => paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>())
            .Where(properties => properties is not null)
            .Cast<SectionProperties>()
            .Concat(body.Elements<SectionProperties>())
            .ToList();
        if (requestedSection < 0 || requestedSection >= sectionProperties.Count)
        {
            return new Resolution(null, $"Section target '{target.SectionId}' was not found");
        }

        var headerIds = new Dictionary<HeaderPart, string>();
        var footerIds = new Dictionary<FooterPart, string>();
        var currentHeaders = new Dictionary<string, HeaderPart>(StringComparer.Ordinal);
        var currentFooters = new Dictionary<string, FooterPart>(StringComparer.Ordinal);
        for (var sectionIndex = 0; sectionIndex <= requestedSection; sectionIndex++)
        {
            foreach (var reference in sectionProperties[sectionIndex].Elements<HeaderReference>())
            {
                var type = reference.Type?.Value.ToString() ?? "default";
                var part = (HeaderPart)mainPart.GetPartById(reference.Id!.Value!);
                if (!headerIds.ContainsKey(part)) headerIds.Add(part, $"header-{headerIds.Count}");
                currentHeaders[type] = part;
            }
            foreach (var reference in sectionProperties[sectionIndex].Elements<FooterReference>())
            {
                var type = reference.Type?.Value.ToString() ?? "default";
                var part = (FooterPart)mainPart.GetPartById(reference.Id!.Value!);
                if (!footerIds.ContainsKey(part)) footerIds.Add(part, $"footer-{footerIds.Count}");
                currentFooters[type] = part;
            }
        }

        OpenXmlPartRootElement? root = null;
        if (target.PartId.StartsWith("header-", StringComparison.Ordinal))
        {
            var pair = currentHeaders.Values.Distinct().Select(part => (Part: part, Id: headerIds[part]))
                .SingleOrDefault(item => string.Equals(item.Id, target.PartId, StringComparison.Ordinal));
            root = pair.Part?.Header;
        }
        else if (target.PartId.StartsWith("footer-", StringComparison.Ordinal))
        {
            var pair = currentFooters.Values.Distinct().Select(part => (Part: part, Id: footerIds[part]))
                .SingleOrDefault(item => string.Equals(item.Id, target.PartId, StringComparison.Ordinal));
            root = pair.Part?.Footer;
        }
        if (root is null)
        {
            return new Resolution(null, $"Part {target.PartId} is not bound to {target.SectionId}");
        }

        var allParagraphs = root.Descendants<Paragraph>().ToList();
        var matches = allParagraphs.Where(paragraph => string.Equals(GetNormalizedParagraphText(paragraph), target.ParagraphText, StringComparison.Ordinal)).ToList();
        if (target.ParagraphOccurrence.Value >= matches.Count)
        {
            return new Resolution(null, $"Paragraph target '{target.ParagraphText}' occurrence {target.ParagraphOccurrence} was not found in {target.PartId}");
        }
        var paragraph = matches[target.ParagraphOccurrence.Value];
        var paragraphId = $"{target.PartId}-p{allParagraphs.IndexOf(paragraph)}";
        if (!string.Equals(target.ParagraphId, paragraphId, StringComparison.Ordinal))
        {
            return new Resolution(null, $"Paragraph identity mismatch: resolved {paragraphId}, requested {target.ParagraphId}");
        }

        return new Resolution(
            new ResolvedFormatElement(
                paragraph,
                paragraph.ParagraphProperties,
                new DocxResolvedFormatTarget("headerFooterParagraph", paragraphId, paragraphId, SectionId: target.SectionId, PartId: target.PartId)),
            null);
    }

    private static bool TryParseStableId(string value, string prefix, out int index)
    {
        index = -1;
        return value.StartsWith(prefix + "-", StringComparison.Ordinal)
            && value.Length > prefix.Length + 1
            && int.TryParse(value.AsSpan(prefix.Length + 1), NumberStyles.None, CultureInfo.InvariantCulture, out index);
    }

    private static DocxEditAppliedOperation ApplyResolvedFormat(
        DocxEditOperation operation,
        ResolvedFormatElement resolved,
        Action<IReadOnlyDictionary<string, string>> apply)
    {
        var priorHash = ComputeFormatHash(resolved.Properties);
        if (!string.Equals(priorHash, operation.ExpectedCurrentFormatHash, StringComparison.OrdinalIgnoreCase))
        {
            return FormatFailure(operation, $"Stale format hash for {resolved.Target.ExactId}", resolved.Target, priorHash);
        }

        apply(operation.FormatProperties!);
        OpenXmlElement? newProperties = resolved.Owner switch
        {
            Paragraph paragraph => paragraph.ParagraphProperties,
            Run run => run.RunProperties,
            SectionProperties section => section,
            _ => throw new InvalidOperationException($"Unsupported format owner {resolved.Owner.GetType().Name}")
        };
        var newHash = ComputeFormatHash(newProperties);
        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Updated format for {resolved.Target.ExactId}",
            resolved.Target,
            priorHash,
            newHash);
    }

    private static DocxEditAppliedOperation FormatFailure(
        DocxEditOperation operation,
        string detail,
        DocxResolvedFormatTarget? target = null,
        string? priorHash = null)
        => new(operation.Type, false, detail, target, priorHash);

    private static string ComputeFormatHash(OpenXmlElement? properties)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(properties?.OuterXml ?? "<absent/>")));

    private static void ApplyParagraphProperties(Paragraph paragraph, IReadOnlyDictionary<string, string> values)
    {
        var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
        foreach (var (name, value) in values)
        {
            switch (name)
            {
                case "alignment":
                    properties.Justification = new Justification { Val = ParseJustification(value) };
                    break;
                case "spacingBeforeTwips":
                    (properties.SpacingBetweenLines ??= new SpacingBetweenLines()).Before = value;
                    break;
                case "spacingAfterTwips":
                    (properties.SpacingBetweenLines ??= new SpacingBetweenLines()).After = value;
                    break;
                case "lineSpacingTwips":
                    (properties.SpacingBetweenLines ??= new SpacingBetweenLines()).Line = value;
                    break;
                case "lineSpacingRule":
                    (properties.SpacingBetweenLines ??= new SpacingBetweenLines()).LineRule = value switch
                    {
                        "auto" => LineSpacingRuleValues.Auto,
                        "atLeast" => LineSpacingRuleValues.AtLeast,
                        _ => LineSpacingRuleValues.Exact
                    };
                    break;
                case "keepWithNext": properties.KeepNext = new KeepNext { Val = bool.Parse(value) }; break;
                case "keepLines": properties.KeepLines = new KeepLines { Val = bool.Parse(value) }; break;
                case "pageBreakBefore": properties.PageBreakBefore = new PageBreakBefore { Val = bool.Parse(value) }; break;
                case "widowControl": properties.WidowControl = new WidowControl { Val = bool.Parse(value) }; break;
            }
        }
    }

    private static JustificationValues ParseJustification(string value) => value switch
    {
        "left" => JustificationValues.Left,
        "center" => JustificationValues.Center,
        "right" => JustificationValues.Right,
        "both" => JustificationValues.Both,
        _ => JustificationValues.Distribute
    };

    private static void ApplyRunProperties(Run run, IReadOnlyDictionary<string, string> values)
    {
        var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
        foreach (var (name, value) in values)
        {
            switch (name)
            {
                case "fontAscii": (properties.RunFonts ??= new RunFonts()).Ascii = value; break;
                case "fontHighAnsi": (properties.RunFonts ??= new RunFonts()).HighAnsi = value; break;
                case "fontEastAsia": (properties.RunFonts ??= new RunFonts()).EastAsia = value; break;
                case "fontComplexScript": (properties.RunFonts ??= new RunFonts()).ComplexScript = value; break;
                case "fontSizeHalfPoints": properties.FontSize = new FontSize { Val = value }; break;
                case "bold": properties.Bold = new Bold { Val = bool.Parse(value) }; break;
                case "italic": properties.Italic = new Italic { Val = bool.Parse(value) }; break;
                case "underline":
                    properties.Underline = new Underline
                    {
                        Val = value switch
                        {
                            "none" => UnderlineValues.None,
                            "double" => UnderlineValues.Double,
                            "words" => UnderlineValues.Words,
                            _ => UnderlineValues.Single
                        }
                    };
                    break;
            }
        }
    }

    private static void ApplySectionProperties(SectionProperties section, IReadOnlyDictionary<string, string> values)
    {
        var pageSize = section.GetFirstChild<PageSize>();
        var pageMargin = section.GetFirstChild<PageMargin>();
        foreach (var (name, value) in values)
        {
            var numeric = uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) ? parsed : 0U;
            switch (name)
            {
                case "orientation":
                    pageSize ??= EnsurePageSize(section);
                    var width = pageSize.Width?.Value ?? 11906U;
                    var height = pageSize.Height?.Value ?? 16838U;
                    var shortSide = Math.Min(width, height);
                    var longSide = Math.Max(width, height);
                    pageSize.Width = value == "landscape" ? longSide : shortSide;
                    pageSize.Height = value == "landscape" ? shortSide : longSide;
                    pageSize.Orient = value == "landscape" ? PageOrientationValues.Landscape : null;
                    break;
                case "marginTopTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Top = (int)numeric; break;
                case "marginRightTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Right = numeric; break;
                case "marginBottomTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Bottom = (int)numeric; break;
                case "marginLeftTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Left = numeric; break;
                case "headerDistanceTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Header = numeric; break;
                case "footerDistanceTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Footer = numeric; break;
                case "gutterTwips": (pageMargin ??= EnsurePageMargin(section, pageSize)).Gutter = numeric; break;
            }
        }
    }

    private static PageMargin EnsurePageMargin(SectionProperties section, PageSize? pageSize)
    {
        var margin = new PageMargin { Top = 1440, Right = 1440U, Bottom = 1440, Left = 1440U, Header = 720U, Footer = 720U, Gutter = 0U };
        var anchor = pageSize as OpenXmlElement ?? FindLastSectionChildBeforePageSize(section);
        return anchor is null ? section.PrependChild(margin) : section.InsertAfter(margin, anchor);
    }

    private static PageSize EnsurePageSize(SectionProperties section)
    {
        var pageSize = new PageSize { Width = 11906U, Height = 16838U };
        var anchor = FindLastSectionChildBeforePageSize(section);
        return anchor is null ? section.PrependChild(pageSize) : section.InsertAfter(pageSize, anchor);
    }

    private static OpenXmlElement? FindLastSectionChildBeforePageSize(SectionProperties section)
        => section.ChildElements.LastOrDefault(child => child is HeaderReference
            or FooterReference
            or FootnoteProperties
            or EndnoteProperties
            or SectionType);

    private static string GetRunText(Run run)
        => string.Concat(run.Descendants<Text>().Select(text => text.Text));

    private static string GetNormalizedParagraphText(Paragraph paragraph)
        => GetParagraphText(paragraph).Trim();

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

    private static DocxEditAppliedOperation StartSectionBeforeParagraph(Body body, DocxEditOperation operation)
    {
        if (string.IsNullOrWhiteSpace(operation.FindText) || string.IsNullOrWhiteSpace(operation.Orientation))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "findText and orientation are required");
        }

        var children = body.ChildElements.ToList();
        var target = children
            .OfType<Paragraph>()
            .FirstOrDefault(paragraph => GetParagraphText(paragraph).Contains(operation.FindText, StringComparison.Ordinal));
        if (target is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Paragraph not found: {operation.FindText}");
        }

        var targetIndex = children.IndexOf(target);
        if (targetIndex < 0)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Paragraph is not a direct body child: {operation.FindText}");
        }

        var nextSectionProperties = FindNextSectionProperties(children, targetIndex);
        if (nextSectionProperties is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"No following section properties found after paragraph: {operation.FindText}");
        }

        var breakParagraph = new Paragraph(new ParagraphProperties((SectionProperties)nextSectionProperties.CloneNode(true)));
        body.InsertBefore(breakParagraph, target);
        SetSectionOrientation(nextSectionProperties, operation.Orientation);

        return new DocxEditAppliedOperation(operation.Type, true, $"Started {operation.Orientation} section before paragraph containing: {operation.FindText}");
    }

    private static SectionProperties? FindNextSectionProperties(IReadOnlyList<OpenXmlElement> bodyChildren, int startIndex)
    {
        for (var index = startIndex; index < bodyChildren.Count; index++)
        {
            if (bodyChildren[index] is Paragraph paragraph)
            {
                var sectionProperties = paragraph.ParagraphProperties?.GetFirstChild<SectionProperties>();
                if (sectionProperties is not null)
                {
                    return sectionProperties;
                }
            }

            if (bodyChildren[index] is SectionProperties bodySectionProperties)
            {
                return bodySectionProperties;
            }
        }

        return null;
    }

    private static void SetSectionOrientation(SectionProperties sectionProperties, string orientation)
    {
        var pageSize = sectionProperties.GetFirstChild<PageSize>();
        if (pageSize is null)
        {
            pageSize = sectionProperties.PrependChild(new PageSize { Width = 11906, Height = 16838 });
        }

        var width = pageSize.Width?.Value ?? 11906U;
        var height = pageSize.Height?.Value ?? 16838U;
        var shortSide = Math.Min(width, height);
        var longSide = Math.Max(width, height);

        if (string.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase))
        {
            pageSize.Width = longSide;
            pageSize.Height = shortSide;
            pageSize.Orient = PageOrientationValues.Landscape;
            return;
        }

        if (string.Equals(orientation, "portrait", StringComparison.OrdinalIgnoreCase))
        {
            pageSize.Width = shortSide;
            pageSize.Height = longSide;
            pageSize.Orient = null;
            return;
        }

        throw new InvalidOperationException($"Unsupported section orientation: {orientation}");
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

        var fallbackRun = FindNearestTableRun(rows, operation.RowIndex.Value, operation.CellIndex.Value);
        ReplaceTableCellText(cells[operation.CellIndex.Value], operation.Text, operation.Alignment, fallbackRun);
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

    private static DocxEditAppliedOperation SetTableCellNoWrap(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and cellIndex are required");
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

        var properties = cells[operation.CellIndex.Value].GetFirstChild<TableCellProperties>()
            ?? cells[operation.CellIndex.Value].PrependChild(new TableCellProperties());
        properties.RemoveAllChildren<NoWrap>();
        var noWrap = operation.NoWrap is not false;
        if (noWrap)
        {
            properties.AppendChild(new NoWrap());
        }
        NormalizeTableCellProperties(properties);

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}] noWrap={noWrap}");
    }

    private static DocxEditAppliedOperation SetTableCellFontSize(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || string.IsNullOrWhiteSpace(operation.FontSize))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and fontSize are required");
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

        var normalizedSize = NormalizeFontSize(operation.FontSize);
        if (normalizedSize is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid fontSize: {operation.FontSize}");
        }
        foreach (var run in cells[operation.CellIndex.Value].Descendants<Run>())
        {
            var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
            properties.RemoveAllChildren<FontSize>();
            properties.RemoveAllChildren<FontSizeComplexScript>();
            properties.AppendChild(new FontSize { Val = normalizedSize });
            properties.AppendChild(new FontSizeComplexScript { Val = normalizedSize });
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}] font size");
    }

    private static DocxEditAppliedOperation SetTableRowHeight(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || string.IsNullOrWhiteSpace(operation.Height))
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and height are required");
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

        if (!uint.TryParse(operation.Height, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var height))
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid height: {operation.Height}");
        }

        var heightRule = ParseHeightRule(operation.HeightRule);
        if (heightRule is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid heightRule: {operation.HeightRule}");
        }

        var properties = rows[operation.RowIndex.Value].GetFirstChild<TableRowProperties>() ?? rows[operation.RowIndex.Value].PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<TableRowHeight>();
        properties.AppendChild(new TableRowHeight
        {
            Val = UInt32Value.FromUInt32(height),
            HeightType = heightRule.Value,
        });

        return new DocxEditAppliedOperation(operation.Type, true, $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}] height");
    }

    private static DocxEditAppliedOperation SetTableRowCantSplit(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CantSplit is null)
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and cantSplit are required");

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");

        var properties = rows[operation.RowIndex.Value].GetFirstChild<TableRowProperties>()
            ?? rows[operation.RowIndex.Value].PrependChild(new TableRowProperties());
        properties.RemoveAllChildren<CantSplit>();
        if (operation.CantSplit.Value) properties.AppendChild(new CantSplit());

        return new DocxEditAppliedOperation(operation.Type, true,
            $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}] cantSplit={operation.CantSplit.Value.ToString().ToLowerInvariant()}");
    }

    private static string? NormalizeFontSize(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var normalized = value.Trim();
        if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var pointValue = normalized[..^2].Trim();
            if (!decimal.TryParse(pointValue, System.Globalization.NumberStyles.AllowDecimalPoint, System.Globalization.CultureInfo.InvariantCulture, out var points) || points <= 0)
            {
                return null;
            }

            return decimal.Round(points * 2, 0, MidpointRounding.AwayFromZero).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        if (!uint.TryParse(normalized, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out var halfPoints) || halfPoints == 0)
        {
            return null;
        }

        return halfPoints.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static HeightRuleValues? ParseHeightRule(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return HeightRuleValues.AtLeast;
        }

        return value.Trim().ToLowerInvariant() switch
        {
            "auto" => HeightRuleValues.Auto,
            "atleast" or "at-least" or "at_least" => HeightRuleValues.AtLeast,
            "exact" => HeightRuleValues.Exact,
            _ => null,
        };
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

    private static DocxEditAppliedOperation InsertTableRows(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.Rows is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and rows are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var existingRows = table.Elements<TableRow>().ToList();
        var insertBeforeIndex = operation.RowIndex.Value;
        if (insertBeforeIndex < 0 || insertBeforeIndex > existingRows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {insertBeforeIndex} is out of range");
        }

        var templateRowResult = ResolveTemplateRow(existingRows, operation.TemplateRowIndex, insertBeforeIndex);
        if (!templateRowResult.Valid)
        {
            return new DocxEditAppliedOperation(operation.Type, false, templateRowResult.Error ?? "Invalid template row");
        }

        var templateRow = templateRowResult.Row;
        InsertBuiltRows(table, existingRows.ElementAtOrDefault(insertBeforeIndex), templateRow, operation.Rows);
        return new DocxEditAppliedOperation(operation.Type, true, $"Inserted {operation.Rows.Count} row(s) into table[{operation.TableIndex}] before row[{insertBeforeIndex}]");
    }

    private static DocxEditAppliedOperation ReplaceTableRows(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.StartRowIndex is null || operation.EndRowIndex is null || operation.Rows is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, startRowIndex, endRowIndex, and rows are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var existingRows = table.Elements<TableRow>().ToList();
        var start = operation.StartRowIndex.Value;
        var end = operation.EndRowIndex.Value;
        if (start < 0 || end >= existingRows.Count || end < start)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {start} to {end}");
        }

        var templateRowResult = ResolveTemplateRow(existingRows, operation.TemplateRowIndex, start);
        if (!templateRowResult.Valid)
        {
            return new DocxEditAppliedOperation(operation.Type, false, templateRowResult.Error ?? "Invalid template row");
        }

        var templateRow = templateRowResult.Row;
        var anchor = existingRows[start];
        var templateCandidates = existingRows.Skip(start).Take(end - start + 1).ToList();
        InsertBuiltRows(table, anchor, templateRow, operation.Rows, templateCandidates);
        foreach (var row in existingRows.Skip(start).Take(end - start + 1))
        {
            row.Remove();
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Replaced table[{operation.TableIndex}].rows[{start}..{end}] with {operation.Rows.Count} row(s)");
    }

    private static DocxEditAppliedOperation DeleteTableRows(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.StartRowIndex is null || operation.EndRowIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, startRowIndex, and endRowIndex are required");
        }

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var existingRows = table.Elements<TableRow>().ToList();
        var start = operation.StartRowIndex.Value;
        var end = operation.EndRowIndex.Value;
        if (start < 0 || end >= existingRows.Count || end < start)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {start} to {end}");
        }

        foreach (var row in existingRows.Skip(start).Take(end - start + 1))
        {
            row.Remove();
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Deleted table[{operation.TableIndex}].rows[{start}..{end}]");
    }

    private static (bool Valid, TableRow? Row, string? Error) ResolveTemplateRow(IReadOnlyList<TableRow> rows, int? templateRowIndex, int fallbackIndex)
    {
        if (rows.Count == 0)
        {
            return (true, null, null);
        }

        var index = templateRowIndex ?? Math.Clamp(fallbackIndex, 0, rows.Count - 1);
        if (index < 0 || index >= rows.Count)
        {
            return (false, null, $"templateRowIndex {index} is out of range");
        }

        return (true, rows[index], null);
    }

    private static void InsertBuiltRows(
        Table table,
        TableRow? beforeRow,
        TableRow? templateRow,
        IReadOnlyList<IReadOnlyList<DocxTableCellInput>> rows,
        IReadOnlyList<TableRow>? templateCandidates = null)
    {
        foreach (var rowInput in rows)
        {
            var rowTemplate = ResolveReplacementRowTemplate(rowInput, templateCandidates, templateRow);
            var row = BuildReplacementRow(rowTemplate, rowInput, rowInput.Any(cell => cell.Header == true));
            if (beforeRow is null)
            {
                table.AppendChild(row);
            }
            else
            {
                table.InsertBefore(row, beforeRow);
            }
        }
    }

    private static TableRow? ResolveReplacementRowTemplate(
        IReadOnlyList<DocxTableCellInput> rowInput,
        IReadOnlyList<TableRow>? candidates,
        TableRow? fallback)
    {
        if (candidates is not { Count: > 0 })
        {
            return fallback;
        }

        var inputPattern = rowInput.Select(cell => Math.Max(1, cell.GridSpan ?? 1)).ToArray();
        var exact = candidates.FirstOrDefault(row => row.Elements<TableCell>().Select(GetCellGridSpan).SequenceEqual(inputPattern));
        if (exact is not null)
        {
            return exact;
        }

        var sameCellCount = candidates.FirstOrDefault(row => row.Elements<TableCell>().Count() == rowInput.Count);
        return sameCellCount ?? fallback;
    }

    private static DocxEditAppliedOperation InsertTableColumns(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.ColumnIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and columnIndex are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var columnIndex = operation.ColumnIndex.Value;
        var columnCount = operation.ColumnCount ?? 1;
        if (columnIndex < 0 || columnCount < 1)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid columnIndex {columnIndex} or columnCount {columnCount}");
        }

        var table = tables[operation.TableIndex.Value];
        var existingGridWidth = GetTableGridWidth(table);
        if (columnIndex > existingGridWidth)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"columnIndex {columnIndex} is out of range for grid width {existingGridWidth}");
        }

        var templateColumnIndex = operation.TemplateColumnIndex ?? Math.Clamp(columnIndex == existingGridWidth ? columnIndex - 1 : columnIndex, 0, Math.Max(0, existingGridWidth - 1));
        if (existingGridWidth > 0 && (templateColumnIndex < 0 || templateColumnIndex >= existingGridWidth))
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"templateColumnIndex {templateColumnIndex} is out of range for grid width {existingGridWidth}");
        }

        InsertGridColumns(table, columnIndex, templateColumnIndex, columnCount, existingGridWidth);

        var rows = table.Elements<TableRow>().ToList();
        foreach (var row in rows)
        {
            InsertColumnsIntoRow(row, columnIndex, templateColumnIndex, columnCount);
        }

        return new DocxEditAppliedOperation(operation.Type, true, $"Inserted {columnCount} column(s) into table[{operation.TableIndex}] before grid column {columnIndex}");
    }

    private static int GetTableGridWidth(Table table)
    {
        var gridColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0;
        var rowWidth = table.Elements<TableRow>()
            .Select(row => row.Elements<TableCell>().Sum(GetCellGridSpan))
            .DefaultIfEmpty(0)
            .Max();
        return Math.Max(gridColumns, rowWidth);
    }

    private static void InsertGridColumns(Table table, int columnIndex, int templateColumnIndex, int columnCount, int existingGridWidth)
    {
        var properties = table.GetFirstChild<TableProperties>();
        if (properties is null)
        {
            properties = new TableProperties();
            table.PrependChild(properties);
        }

        var grid = table.GetFirstChild<TableGrid>();
        if (grid is null)
        {
            grid = new TableGrid();
            for (var i = 0; i < existingGridWidth; i++)
            {
                grid.AppendChild(new GridColumn { Width = "1200" });
            }

            table.InsertAfter(grid, properties);
        }

        var columns = grid.Elements<GridColumn>().ToList();
        var template = columns.ElementAtOrDefault(templateColumnIndex);
        for (var i = 0; i < columnCount; i++)
        {
            var inserted = template is null
                ? new GridColumn { Width = "1200" }
                : (GridColumn)template.CloneNode(true);
            var before = columns.ElementAtOrDefault(columnIndex + i);
            if (before is null)
            {
                grid.AppendChild(inserted);
            }
            else
            {
                grid.InsertBefore(inserted, before);
            }
            columns.Insert(Math.Min(columnIndex + i, columns.Count), inserted);
        }
    }

    private static void InsertColumnsIntoRow(TableRow row, int columnIndex, int templateColumnIndex, int columnCount)
    {
        var cells = row.Elements<TableCell>().ToList();
        var cursor = 0;
        TableCell? insertBefore = null;
        TableCell? templateCell = null;

        foreach (var cell in cells)
        {
            var span = GetCellGridSpan(cell);
            var start = cursor;
            var endExclusive = cursor + span;

            if (templateCell is null && templateColumnIndex >= start && templateColumnIndex < endExclusive)
            {
                templateCell = cell;
            }

            if (columnIndex > start && columnIndex < endExclusive)
            {
                SetCellGridSpan(cell, span + columnCount);
                return;
            }

            if (columnIndex <= start)
            {
                insertBefore = cell;
                break;
            }

            cursor = endExclusive;
        }

        templateCell ??= cells.LastOrDefault();
        for (var i = 0; i < columnCount; i++)
        {
            var cell = BuildInsertedColumnCell(templateCell);
            if (insertBefore is null)
            {
                row.AppendChild(cell);
            }
            else
            {
                row.InsertBefore(cell, insertBefore);
            }
        }
    }

    private static int GetCellGridSpan(TableCell cell)
        => Math.Max(1, cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1);

    private static void SetCellGridSpan(TableCell cell, int span)
    {
        var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
        properties.RemoveAllChildren<GridSpan>();
        if (span > 1)
        {
            properties.AppendChild(new GridSpan { Val = span });
        }
        NormalizeTableCellProperties(properties);
    }

    private static TableCell BuildInsertedColumnCell(TableCell? templateCell)
    {
        var cell = templateCell is null
            ? new TableCell(new TableCellProperties(), new Paragraph(new Run(new Text(string.Empty))))
            : (TableCell)templateCell.CloneNode(true);
        SetCellGridSpan(cell, 1);
        cell.GetFirstChild<TableCellProperties>()?.RemoveAllChildren<VerticalMerge>();
        ReplaceTableCellText(cell, string.Empty);
        return cell;
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
        if (input.GridSpan is not null)
        {
            properties.RemoveAllChildren<GridSpan>();
            if (input.GridSpan is > 1)
            {
                properties.AppendChild(new GridSpan { Val = input.GridSpan.Value });
            }
        }

        if (input.VMerge is not null)
        {
            properties.RemoveAllChildren<VerticalMerge>();
            if (input.VMerge is { Length: > 0 } vMergeVal)
            {
                var vmVal = vMergeVal.ToLowerInvariant() == "restart" ? MergedCellValues.Restart : MergedCellValues.Continue;
                properties.AppendChild(new VerticalMerge { Val = vmVal });
            }
        }

        if (input.Shading is { Length: > 0 } hexColor)
        {
            properties.RemoveAllChildren<Shading>();
            properties.AppendChild(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = hexColor });
        }
        NormalizeTableCellProperties(properties);

        var paragraph = CreateParagraphLike(templateCell?.Elements<Paragraph>().FirstOrDefault());
        var paragraphProperties = paragraph.GetFirstChild<ParagraphProperties>() ?? paragraph.PrependChild(new ParagraphProperties());
        
        if (input.Alignment is { Length: > 0 } align)
        {
            var jcVal = align.ToLowerInvariant() switch
            {
                "center" => JustificationValues.Center,
                "right" => JustificationValues.Right,
                _ => JustificationValues.Left
            };
            paragraphProperties.RemoveAllChildren<Justification>();
            paragraphProperties.AppendChild(new Justification { Val = jcVal });
        }

        if (input.RichText is { Count: > 0 } richText)
        {
            foreach (var segment in richText)
            {
                paragraph.AppendChild(CreateRichRunLike(templateCell?.Descendants<Run>().FirstOrDefault(), segment, rowIsHeader));
            }
        }
        else
        {
            var run = CreateStyledRunLike(templateCell?.Descendants<Run>().FirstOrDefault(), input.Text ?? string.Empty);
            if (input.Bold == true || rowIsHeader)
            {
                var runProperties = run.RunProperties ?? run.PrependChild(new RunProperties());
                runProperties.RemoveAllChildren<Bold>();
                runProperties.AppendChild(new Bold());
            }
            paragraph.AppendChild(run);
        }
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

    private static void ReplaceTableCellText(TableCell cell, string replacementText, string? alignment = null, Run? fallbackRun = null)
    {
        var ownRun = cell.Descendants<Run>().FirstOrDefault();
        var firstRun = ownRun ?? fallbackRun;
        var firstParagraph = cell.Elements<Paragraph>().FirstOrDefault();
        cell.RemoveAllChildren<Paragraph>();
        var paragraph = CreateParagraphLike(firstParagraph);
        paragraph.AppendChild(CreateStyledRunLike(firstRun, replacementText, preserveEmphasis: ownRun is not null));
        if (!string.IsNullOrWhiteSpace(alignment))
        {
            ApplyParagraphAlignment(paragraph, alignment);
        }
        cell.Append(paragraph);
    }

    private static DocxEditAppliedOperation ReplaceTableCellRichText(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null || operation.RichText is not { Count: > 0 })
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, cellIndex, and richText are required");
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

        var fallbackRun = FindNearestTableRun(rows, operation.RowIndex.Value, operation.CellIndex.Value);
        ReplaceTableCellRichText(cells[operation.CellIndex.Value], operation.RichText, operation.Alignment, fallbackRun);
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated rich text in table[{operation.TableIndex}].row[{operation.RowIndex}].cell[{operation.CellIndex}]");
    }

    private static void ReplaceTableCellRichText(TableCell cell, IReadOnlyList<DocxRichTextSegment> segments, string? alignment = null, Run? fallbackRun = null)
    {
        var ownRun = cell.Descendants<Run>().FirstOrDefault();
        var firstRun = ownRun ?? fallbackRun;
        var firstParagraph = cell.Elements<Paragraph>().FirstOrDefault();
        cell.RemoveAllChildren<Paragraph>();
        var paragraph = CreateParagraphLike(firstParagraph);
        foreach (var segment in segments)
        {
            paragraph.AppendChild(CreateRichRunLike(firstRun, segment, preserveEmphasis: ownRun is not null));
        }
        if (!string.IsNullOrWhiteSpace(alignment))
        {
            ApplyParagraphAlignment(paragraph, alignment);
        }
        cell.Append(paragraph);
    }

    private static Run? FindNearestTableRun(IReadOnlyList<TableRow> rows, int rowIndex, int cellIndex)
    {
        var exactRow = FindNearestRunInRow(rows[rowIndex], cellIndex);
        if (exactRow is not null)
        {
            return exactRow;
        }

        for (var offset = 1; offset < rows.Count; offset++)
        {
            var previous = rowIndex - offset;
            if (previous >= 0)
            {
                var run = FindNearestRunInRow(rows[previous], cellIndex);
                if (run is not null)
                {
                    return run;
                }
            }

            var next = rowIndex + offset;
            if (next < rows.Count)
            {
                var run = FindNearestRunInRow(rows[next], cellIndex);
                if (run is not null)
                {
                    return run;
                }
            }
        }

        return null;
    }

    private static Run? FindNearestRunInRow(TableRow row, int cellIndex)
    {
        var cells = row.Elements<TableCell>().ToList();
        if (cells.Count == 0)
        {
            return null;
        }

        if (cellIndex >= 0 && cellIndex < cells.Count)
        {
            var run = cells[cellIndex].Descendants<Run>().FirstOrDefault();
            if (run is not null)
            {
                return run;
            }
        }

        for (var offset = 1; offset < cells.Count; offset++)
        {
            var previous = cellIndex - offset;
            if (previous >= 0)
            {
                var run = cells[previous].Descendants<Run>().FirstOrDefault();
                if (run is not null)
                {
                    return run;
                }
            }

            var next = cellIndex + offset;
            if (next < cells.Count)
            {
                var run = cells[next].Descendants<Run>().FirstOrDefault();
                if (run is not null)
                {
                    return run;
                }
            }
        }

        return null;
    }

    private static Paragraph CreateParagraphLike(Paragraph? templateParagraph)
    {
        var paragraph = new Paragraph();
        var templateProperties = templateParagraph?.GetFirstChild<ParagraphProperties>();
        if (templateProperties is not null)
        {
            paragraph.AppendChild((ParagraphProperties)templateProperties.CloneNode(true));
        }
        return paragraph;
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

    private static Run CreateStyledRunLike(Run? templateRun, string text, bool preserveEmphasis = true)
    {
        var run = new Run();
        if (templateRun?.RunProperties is not null)
        {
            run.RunProperties = (RunProperties)templateRun.RunProperties.CloneNode(true);
            if (!preserveEmphasis)
            {
                RemoveEmphasis(run.RunProperties);
            }
            NormalizeRunProperties(run.RunProperties);
        }
        AppendTextWithLineBreaks(run, text);
        return run;
    }

    private static Run CreateRichRunLike(Run? templateRun, DocxRichTextSegment segment, bool forceBold = false, bool preserveEmphasis = true)
    {
        var run = new Run();
        if (templateRun?.RunProperties is not null)
        {
            run.RunProperties = (RunProperties)templateRun.RunProperties.CloneNode(true);
        }

        var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
        RemoveTextFill(properties);
        if (!preserveEmphasis)
        {
            RemoveEmphasis(properties);
        }

        if (forceBold || segment.Bold == true)
        {
            properties.RemoveAllChildren<Bold>();
            properties.AppendChild(new Bold());
        }
        else if (segment.Bold == false)
        {
            properties.RemoveAllChildren<Bold>();
        }

        if (!string.IsNullOrWhiteSpace(segment.Color))
        {
            properties.RemoveAllChildren<Color>();
            properties.AppendChild(new Color { Val = segment.Color });
        }

        if (!string.IsNullOrWhiteSpace(segment.FontName))
        {
            properties.RemoveAllChildren<RunFonts>();
            properties.PrependChild(new RunFonts
            {
                Ascii = segment.FontName,
                HighAnsi = segment.FontName,
            });
        }

        if (segment.Underline == true)
        {
            properties.RemoveAllChildren<Underline>();
            properties.AppendChild(new Underline { Val = UnderlineValues.Single });
        }
        else if (segment.Underline == false)
        {
            properties.RemoveAllChildren<Underline>();
        }

        NormalizeRunProperties(properties);
        AppendTextWithLineBreaks(run, segment.Text);
        return run;
    }

    private static void RemoveEmphasis(RunProperties properties)
    {
        properties.RemoveAllChildren<Bold>();
        properties.RemoveAllChildren<BoldComplexScript>();
        properties.RemoveAllChildren<Italic>();
        properties.RemoveAllChildren<ItalicComplexScript>();
    }

    private static void RemoveTextFill(RunProperties properties)
    {
        foreach (var textFill in properties.Elements<W14.FillTextEffect>().ToList())
        {
            textFill.Remove();
        }
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
                if (rIdx == startRowIndex && IsVerticalMergeContinuation(properties))
                {
                    var ownerCell = FindPreviousVerticalMergeOwner(rows, startRowIndex, cellIndex);
                    CopyMissingParagraphProperties(ownerCell, cell);
                }
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

    private static bool IsVerticalMergeContinuation(TableCellProperties properties)
    {
        var merge = properties.GetFirstChild<VerticalMerge>();
        return merge is not null && (merge.Val is null || merge.Val.Value == MergedCellValues.Continue);
    }

    private static TableCell? FindPreviousVerticalMergeOwner(IReadOnlyList<TableRow> rows, int startRowIndex, int cellIndex)
    {
        for (var rowIndex = startRowIndex - 1; rowIndex >= 0; rowIndex--)
        {
            var cells = rows[rowIndex].Elements<TableCell>().ToList();
            if (cellIndex >= cells.Count)
            {
                continue;
            }

            var cell = cells[cellIndex];
            var properties = cell.GetFirstChild<TableCellProperties>();
            if (!IsVerticalMergeContinuation(properties ?? new TableCellProperties()))
            {
                return cell;
            }
        }

        return null;
    }

    private static void CopyMissingParagraphProperties(TableCell? sourceCell, TableCell targetCell)
    {
        var sourceProperties = sourceCell?.Elements<Paragraph>().FirstOrDefault()?.GetFirstChild<ParagraphProperties>();
        if (sourceProperties is null)
        {
            return;
        }

        var targetParagraph = targetCell.Elements<Paragraph>().FirstOrDefault();
        if (targetParagraph is null)
        {
            targetParagraph = targetCell.AppendChild(new Paragraph());
        }

        var targetProperties = targetParagraph.GetFirstChild<ParagraphProperties>() ?? targetParagraph.PrependChild(new ParagraphProperties());
        foreach (var sourceProperty in sourceProperties.ChildElements)
        {
            if (targetProperties.ChildElements.Any(existing => existing.GetType() == sourceProperty.GetType()))
            {
                continue;
            }

            targetProperties.AppendChild(sourceProperty.CloneNode(true));
        }
    }

    private static DocxEditAppliedOperation UnmergeTableRowHorizontalCells(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.CellIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and cellIndex are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var rows = table.Elements<TableRow>().ToList();
        var rowIndex = operation.RowIndex.Value;
        if (rowIndex < 0 || rowIndex >= rows.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {rowIndex} is out of range");
        }

        var row = rows[rowIndex];
        var cells = row.Elements<TableCell>().ToList();
        var cellIndex = operation.CellIndex.Value;
        if (cellIndex < 0 || cellIndex >= cells.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {cellIndex} is out of range");
        }

        var cell = cells[cellIndex];
        var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
        var span = properties.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
        if (span <= 1)
        {
            return new DocxEditAppliedOperation(operation.Type, true, $"Cell table[{operation.TableIndex}].row[{rowIndex}].cell[{cellIndex}] is not horizontally merged");
        }

        var gridStart = cells.Take(cellIndex).Sum(GetCellGridSpan);
        var splitWidths = GetUnmergedCellWidths(table, row, gridStart, span);

        properties.RemoveAllChildren<GridSpan>();
        SetCellWidth(properties, splitWidths[0]);
        NormalizeTableCellProperties(properties);

        for (var i = 1; i < span; i++)
        {
            var newCell = (TableCell)cell.CloneNode(true);
            foreach (var child in newCell.ChildElements.Where(child => child is not TableCellProperties).ToList())
            {
                child.Remove();
            }
            newCell.AppendChild(new Paragraph());
            var newProperties = newCell.GetFirstChild<TableCellProperties>() ?? newCell.PrependChild(new TableCellProperties());
            newProperties.RemoveAllChildren<GridSpan>();
            SetCellWidth(newProperties, splitWidths[i]);
            NormalizeTableCellProperties(newProperties);
            row.InsertAfter(newCell, cell);
            cell = newCell;
        }

        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Unmerged horizontal cell in table[{operation.TableIndex}].row[{rowIndex}].cell[{cellIndex}], expanded {span} grid columns");
    }

    private static IReadOnlyList<string> GetUnmergedCellWidths(Table table, TableRow currentRow, int startColumn, int count)
    {
        var visibleReference = FindVisibleCellWidthsForGridRange(table, currentRow, startColumn, count);
        if (visibleReference is not null)
        {
            return visibleReference;
        }

        return GetTableGridWidths(table, startColumn, count);
    }

    private static IReadOnlyList<string>? FindVisibleCellWidthsForGridRange(Table table, TableRow currentRow, int startColumn, int count)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            if (ReferenceEquals(row, currentRow))
            {
                continue;
            }

            var cursor = 0;
            var widths = new List<string>();
            foreach (var cell in row.Elements<TableCell>())
            {
                var span = GetCellGridSpan(cell);
                if (cursor >= startColumn && cursor + span <= startColumn + count)
                {
                    if (span != 1)
                    {
                        widths.Clear();
                        break;
                    }

                    var width = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<TableCellWidth>()?.Width?.Value;
                    widths.Add(string.IsNullOrWhiteSpace(width) ? GetTableGridWidths(table, cursor, 1)[0] : width!);
                }
                else if (cursor < startColumn + count && cursor + span > startColumn)
                {
                    widths.Clear();
                    break;
                }

                cursor += span;
                if (cursor >= startColumn + count)
                {
                    break;
                }
            }

            if (widths.Count == count)
            {
                return widths;
            }
        }

        return null;
    }

    private static IReadOnlyList<string> GetTableGridWidths(Table table, int startColumn, int count)
    {
        var columns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList() ?? [];
        var widths = new List<string>();
        for (var offset = 0; offset < count; offset++)
        {
            var column = columns.ElementAtOrDefault(startColumn + offset);
            widths.Add(string.IsNullOrWhiteSpace(column?.Width) ? "1200" : column.Width!);
        }
        return widths;
    }

    private static void SetCellWidth(TableCellProperties properties, string width)
    {
        properties.RemoveAllChildren<TableCellWidth>();
        properties.PrependChild(new TableCellWidth { Width = width, Type = TableWidthUnitValues.Dxa });
    }

    private static DocxEditAppliedOperation UnmergeTableColumnVerticalCells(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.CellIndex is null)
        {
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex and cellIndex are required");
        }

        var tables = body.Descendants<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");
        }

        var table = tables[operation.TableIndex.Value];
        var rows = table.Elements<TableRow>().ToList();
        var cellIndex = operation.CellIndex.Value;
        var startRowIndex = operation.StartRowIndex ?? 0;
        var endRowIndex = operation.EndRowIndex ?? (rows.Count - 1);

        if (startRowIndex < 0 || endRowIndex >= rows.Count || endRowIndex < startRowIndex)
        {
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid row range {startRowIndex} to {endRowIndex}");
        }

        List<OpenXmlElement>? latestVisibleContent = null;
        var changed = 0;
        for (var rIdx = startRowIndex; rIdx <= endRowIndex; rIdx++)
        {
            var cells = rows[rIdx].Elements<TableCell>().ToList();
            if (cellIndex < 0 || cellIndex >= cells.Count)
            {
                return new DocxEditAppliedOperation(operation.Type, false, $"cellIndex {cellIndex} is out of range in row {rIdx}");
            }

            var cell = cells[cellIndex];
            var properties = cell.GetFirstChild<TableCellProperties>() ?? cell.PrependChild(new TableCellProperties());
            var verticalMerge = properties.GetFirstChild<VerticalMerge>();
            var isContinuation = verticalMerge is not null
                && (verticalMerge.Val is null || verticalMerge.Val.Value == MergedCellValues.Continue);

            if (!isContinuation)
            {
                latestVisibleContent = cell.Elements<Paragraph>()
                    .Select(p => p.CloneNode(true))
                    .ToList();
            }

            if (verticalMerge is not null)
            {
                properties.RemoveAllChildren<VerticalMerge>();
                NormalizeTableCellProperties(properties);
                changed++;
            }

            if (isContinuation && latestVisibleContent is { Count: > 0 })
            {
                cell.RemoveAllChildren<Paragraph>();
                foreach (var paragraph in latestVisibleContent)
                {
                    cell.Append(paragraph.CloneNode(true));
                }
            }
        }

        return new DocxEditAppliedOperation(
            operation.Type,
            true,
            $"Unmerged vertical cells in table[{operation.TableIndex}].cell[{cellIndex}].rows[{startRowIndex}..{endRowIndex}], removed {changed} vMerge marker(s)");
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

    private static string GetParagraphText(Paragraph paragraph)
        => string.Concat(paragraph.Descendants<Text>().Select(text => text.Text));
}
