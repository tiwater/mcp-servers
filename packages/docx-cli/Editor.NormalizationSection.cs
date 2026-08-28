using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace Dockit.Docx;

public static partial class Editor
{
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

    private static DocxEditAppliedOperation SetTableRowKeepNext(Body body, DocxEditOperation operation)
    {
        if (operation.TableIndex is null || operation.RowIndex is null || operation.KeepNext is null)
            return new DocxEditAppliedOperation(operation.Type, false, "tableIndex, rowIndex, and keepNext are required");

        var tables = body.Elements<Table>().ToList();
        if (operation.TableIndex.Value < 0 || operation.TableIndex.Value >= tables.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"tableIndex {operation.TableIndex} is out of range");

        var rows = tables[operation.TableIndex.Value].Elements<TableRow>().ToList();
        if (operation.RowIndex.Value < 0 || operation.RowIndex.Value >= rows.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"rowIndex {operation.RowIndex} is out of range");

        var paragraphs = rows[operation.RowIndex.Value].Elements<TableCell>()
            .SelectMany(cell => cell.Elements<Paragraph>()).ToList();
        if (paragraphs.Count == 0)
            return new DocxEditAppliedOperation(operation.Type, false, $"row {operation.RowIndex} has no paragraphs");

        foreach (var paragraph in paragraphs)
        {
            var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
            properties.RemoveAllChildren<KeepNext>();
            if (operation.KeepNext.Value) properties.AddChild(new KeepNext(), true);
        }

        return new DocxEditAppliedOperation(operation.Type, true,
            $"Updated table[{operation.TableIndex}].row[{operation.RowIndex}] keepNext={operation.KeepNext.Value.ToString().ToLowerInvariant()}");
    }

    private static DocxEditAppliedOperation SetBodyParagraphKeepNext(Body body, DocxEditOperation operation)
    {
        if (operation.ParagraphIndex is null || operation.KeepNext is null)
            return new DocxEditAppliedOperation(operation.Type, false, "paragraphIndex and keepNext are required");

        var paragraphs = body.Elements<Paragraph>().ToList();
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range");

        var paragraph = paragraphs[operation.ParagraphIndex.Value];
        var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
        properties.RemoveAllChildren<KeepNext>();
        if (operation.KeepNext.Value) properties.AddChild(new KeepNext(), true);

        return new DocxEditAppliedOperation(operation.Type, true,
            $"Updated body paragraph[{operation.ParagraphIndex}] keepNext={operation.KeepNext.Value.ToString().ToLowerInvariant()}");
    }

    private static DocxEditAppliedOperation SetBodyParagraphKeepLines(Body body, DocxEditOperation operation)
    {
        if (operation.ParagraphIndex is null || operation.KeepLines is null)
            return new DocxEditAppliedOperation(operation.Type, false, "paragraphIndex and keepLines are required");

        var paragraphs = body.Elements<Paragraph>().ToList();
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range");

        var paragraph = paragraphs[operation.ParagraphIndex.Value];
        var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
        properties.RemoveAllChildren<KeepLines>();
        if (operation.KeepLines.Value) properties.AddChild(new KeepLines(), true);

        return new DocxEditAppliedOperation(operation.Type, true,
            $"Updated body paragraph[{operation.ParagraphIndex}] keepLines={operation.KeepLines.Value.ToString().ToLowerInvariant()}");
    }

    private static DocxEditAppliedOperation ApplyTocStylePolicy(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (operation.Italic is null || operation.IndentCharactersPerLevel is null or < 0)
            return new DocxEditAppliedOperation(operation.Type, false, "italic and non-negative indentCharactersPerLevel are required");
        var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles is null) return new DocxEditAppliedOperation(operation.Type, false, "document styles are missing");

        var matched = 0;
        foreach (var style in styles.Elements<Style>().Where(style => style.Type?.Value == StyleValues.Paragraph))
        {
            var id = style.StyleId?.Value ?? string.Empty;
            var name = style.StyleName?.Val?.Value ?? string.Empty;
            var token = id.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) ? id[3..]
                : name.StartsWith("toc ", StringComparison.OrdinalIgnoreCase) ? name[4..] : string.Empty;
            if (!int.TryParse(token, out var level) || level < 1) continue;

            var paragraph = style.StyleParagraphProperties ?? style.AppendChild(new StyleParagraphProperties());
            paragraph.RemoveAllChildren<Indentation>();
            paragraph.AppendChild(new Indentation { LeftChars = (level - 1) * operation.IndentCharactersPerLevel.Value * 100 });

            var run = style.StyleRunProperties ?? style.AppendChild(new StyleRunProperties());
            run.RemoveAllChildren<Italic>();
            run.RemoveAllChildren<ItalicComplexScript>();
            run.AppendChild(new Italic { Val = operation.Italic.Value });
            run.AppendChild(new ItalicComplexScript { Val = operation.Italic.Value });
            matched += 1;
        }
        if (matched == 0) return new DocxEditAppliedOperation(operation.Type, false, "no built-in TOC paragraph styles were found");
        return new DocxEditAppliedOperation(operation.Type, true, $"Updated {matched} TOC paragraph styles");
    }

    private static DocxEditAppliedOperation SetHeaderParagraphFontSize(WordprocessingDocument doc, DocxEditOperation operation)
    {
        if (operation.HeaderIndex is null || operation.ParagraphIndex is null || string.IsNullOrWhiteSpace(operation.FontSize))
            return new DocxEditAppliedOperation(operation.Type, false, "headerIndex, paragraphIndex, and fontSize are required");
        if (!uint.TryParse(operation.FontSize, out var halfPoints) || halfPoints == 0)
            return new DocxEditAppliedOperation(operation.Type, false, $"Invalid fontSize: {operation.FontSize}");

        var headers = doc.MainDocumentPart?.HeaderParts.ToList() ?? [];
        if (operation.HeaderIndex.Value < 0 || operation.HeaderIndex.Value >= headers.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"headerIndex {operation.HeaderIndex} is out of range");

        var paragraphs = headers[operation.HeaderIndex.Value].Header?.Elements<Paragraph>().ToList() ?? [];
        if (operation.ParagraphIndex.Value < 0 || operation.ParagraphIndex.Value >= paragraphs.Count)
            return new DocxEditAppliedOperation(operation.Type, false, $"paragraphIndex {operation.ParagraphIndex} is out of range for header {operation.HeaderIndex}");

        var runs = paragraphs[operation.ParagraphIndex.Value].Descendants<Run>().ToList();
        if (runs.Count == 0)
            return new DocxEditAppliedOperation(operation.Type, false, $"header[{operation.HeaderIndex}].paragraph[{operation.ParagraphIndex}] has no runs");
        foreach (var run in runs)
        {
            var properties = run.RunProperties ?? run.PrependChild(new RunProperties());
            properties.RemoveAllChildren<FontSize>();
            properties.RemoveAllChildren<FontSizeComplexScript>();
            properties.AddChild(new FontSize { Val = operation.FontSize }, true);
            properties.AddChild(new FontSizeComplexScript { Val = operation.FontSize }, true);
        }

        return new DocxEditAppliedOperation(operation.Type, true,
            $"Updated header[{operation.HeaderIndex}].paragraph[{operation.ParagraphIndex}] fontSize={halfPoints}");
    }

    private static DocxEditAppliedOperation CollapseTrailingEmptySection(Body body, DocxEditOperation operation)
    {
        if (!TryFindTrailingEmptySection(body, out var precedingSection, out var finalSection, out var trailingEmptyParagraphs))
            return new DocxEditAppliedOperation(operation.Type, false, "no collapsible trailing empty section");

        var replacement = (SectionProperties)precedingSection.CloneNode(true);
        precedingSection.Remove();
        foreach (var paragraph in trailingEmptyParagraphs) paragraph.Remove();
        finalSection.Remove();
        body.AppendChild(replacement);

        return new DocxEditAppliedOperation(operation.Type, true, "Collapsed one trailing empty section");
    }

    internal static bool HasTrailingEmptySection(Body body)
        => TryFindTrailingEmptySection(body, out _, out _, out _);

    private static DocxEditAppliedOperation CollapseTrailingEmptyBodyParagraphs(Body body, DocxEditOperation operation)
    {
        if (operation.ExpectedCount is null or <= 0)
            return new DocxEditAppliedOperation(operation.Type, false, "expectedCount must be greater than zero");

        var trailingParagraphs = GetTrailingEmptyBodyParagraphs(body);
        if (trailingParagraphs.Count != operation.ExpectedCount.Value)
            return new DocxEditAppliedOperation(operation.Type, false,
                $"Expected {operation.ExpectedCount.Value} trailing empty body paragraph(s), found {trailingParagraphs.Count}");

        foreach (var paragraph in trailingParagraphs) paragraph.Remove();
        return new DocxEditAppliedOperation(operation.Type, true,
            $"Collapsed {trailingParagraphs.Count} trailing empty body paragraph(s)");
    }

    internal static IReadOnlyList<Paragraph> GetTrailingEmptyBodyParagraphs(Body body)
    {
        var children = body.ChildElements.ToList();
        if (children.LastOrDefault() is not SectionProperties) return [];

        var result = new List<Paragraph>();
        for (var index = children.Count - 2; index >= 0; index--)
        {
            if (children[index] is not Paragraph paragraph || !IsRemovableEmptyBodyParagraph(paragraph)) break;
            result.Add(paragraph);
        }
        result.Reverse();
        return result;
    }

    private static bool IsRemovableEmptyBodyParagraph(Paragraph paragraph)
    {
        if (!string.IsNullOrWhiteSpace(paragraph.InnerText)) return false;
        return !paragraph.Descendants().Any(element => element is
            Drawing or Break or TabChar or CarriageReturn or FieldChar or FieldCode
            or FootnoteReference or EndnoteReference or CommentReference
            or BookmarkStart or BookmarkEnd or Hyperlink);
    }

    private static bool TryFindTrailingEmptySection(
        Body body,
        out SectionProperties precedingSection,
        out SectionProperties finalSection,
        out List<Paragraph> trailingEmptyParagraphs)
    {
        precedingSection = null!;
        finalSection = null!;
        trailingEmptyParagraphs = [];
        var children = body.ChildElements.ToList();
        if (children.LastOrDefault() is not SectionProperties final) return false;

        var boundaryIndex = -1;
        SectionProperties? preceding = null;
        for (var index = children.Count - 2; index >= 0; index--)
        {
            if (children[index] is not Paragraph paragraph) continue;
            preceding = paragraph.ParagraphProperties?.SectionProperties;
            if (preceding is null) continue;
            boundaryIndex = index;
            break;
        }
        if (preceding is null) return false;

        foreach (var child in children.Skip(boundaryIndex + 1).Take(children.Count - boundaryIndex - 2))
        {
            if (child is BookmarkStart or BookmarkEnd) continue;
            if (child is not Paragraph paragraph
                || !string.IsNullOrWhiteSpace(paragraph.InnerText)
                || paragraph.Descendants<Drawing>().Any())
                return false;
            trailingEmptyParagraphs.Add(paragraph);
        }
        if (trailingEmptyParagraphs.Count == 0) return false;

        precedingSection = preceding;
        finalSection = final;
        return true;
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

    private static void NormalizeGeneratedOpenXml(WordprocessingDocument doc)
    {
        const string w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        var paragraphIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var root in Inspector.GetRoots(doc))
        {
            foreach (var paragraph in root.Descendants<Paragraph>())
            {
                var paragraphId = paragraph.GetAttributes()
                    .FirstOrDefault(attribute => attribute.LocalName == "paraId" && attribute.NamespaceUri == w14).Value;
                if (string.IsNullOrWhiteSpace(paragraphId) || paragraphIds.Add(paragraphId)) continue;
                paragraph.RemoveAttribute("paraId", w14);
                paragraph.RemoveAttribute("textId", w14);
            }
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
}
