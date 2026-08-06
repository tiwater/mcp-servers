using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Dockit.Xlsx;

public static class EvidenceInspector
{
    public static object Inspect(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart ?? throw new InvalidOperationException("Workbook not found.");
        var styles = wb.WorkbookStylesPart?.Stylesheet;
        var formats = styles?.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        var customFormats = styles?.NumberingFormats?.Elements<NumberingFormat>()
            .Where(x => x.NumberFormatId?.Value is not null)
            .GroupBy(x => x.NumberFormatId!.Value)
            .ToDictionary(x => x.Key, x => x.Last().FormatCode?.Value) ?? [];
        var baseFormats = styles?.CellStyleFormats?.Elements<CellFormat>().ToList() ?? [];
        var fonts = styles?.Fonts?.Elements<Font>().ToList() ?? [];
        var fills = styles?.Fills?.Elements<Fill>().ToList() ?? [];
        var borders = styles?.Borders?.Elements<Border>().ToList() ?? [];
        var formattedCellsBySheet = WorkbookLoader.Load(path).Sheets.ToDictionary(
            x => x.Name,
            x => x.Cells.ToDictionary(cell => cell.Reference, cell => cell.FormattedValue, StringComparer.OrdinalIgnoreCase),
            StringComparer.Ordinal);
        var uses1904Dates = wb.Workbook.WorkbookProperties?.Date1904?.Value == true;
        var sheets = wb.Workbook.Descendants<Sheet>().Select((sheet, sheetIndex) =>
        {
            var part = (WorksheetPart)wb.GetPartById(sheet.Id!);
            var ws = part.Worksheet;
            var significantColumn = WorksheetEvidenceBounds.SignificantColumn(part);
            var formattedCells = formattedCellsBySheet.GetValueOrDefault(sheet.Name?.Value ?? string.Empty)
                ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var cells = ws.Descendants<Cell>().Where(cell => WorksheetEvidenceBounds.Include(cell, significantColumn)).Select(cell =>
            {
                var styleIndex = cell.StyleIndex?.Value ?? 0U;
                var format = styleIndex < formats.Count ? formats[(int)styleIndex] : null;
                var baseFormatIndex = format?.FormatId?.Value;
                var baseFormat = baseFormatIndex is not null && baseFormatIndex.Value < baseFormats.Count
                    ? baseFormats[(int)baseFormatIndex.Value]
                    : null;
                var numId = SpreadsheetValueNormalizer.EffectiveComponentId(format?.NumberFormatId?.Value, format?.ApplyNumberFormat?.Value, baseFormat?.NumberFormatId?.Value);
                var fontId = SpreadsheetValueNormalizer.EffectiveComponentId(format?.FontId?.Value, format?.ApplyFont?.Value, baseFormat?.FontId?.Value);
                var fillId = SpreadsheetValueNormalizer.EffectiveComponentId(format?.FillId?.Value, format?.ApplyFill?.Value, baseFormat?.FillId?.Value);
                var borderId = SpreadsheetValueNormalizer.EffectiveComponentId(format?.BorderId?.Value, format?.ApplyBorder?.Value, baseFormat?.BorderId?.Value);
                var formatCode = customFormats.GetValueOrDefault(numId) ?? SpreadsheetValueNormalizer.BuiltInFormat(numId) ?? $"builtin:{numId}";
                var normalizedFormat = SpreadsheetValueNormalizer.NormalizeNumberFormat(numId, formatCode, customFormats.ContainsKey(numId));
                var alignment = EffectiveAlignment(format, baseFormat);
                var alignmentEvidence = ResolveAlignment(alignment);
                var protection = EffectiveProtection(format, baseFormat);
                var applyProtection = format?.ApplyProtection?.Value ?? false;
                var locked = protection?.Locked?.Value ?? true;
                var hidden = protection?.Hidden?.Value ?? false;
                var formula = cell.CellFormula;
                var rawValue = cell.CellValue?.Text ?? cell.InlineString?.InnerText;
                var reference = cell.CellReference?.Value;
                var valueType = cell.DataType?.InnerText;
                var normalizedValue = SpreadsheetValueNormalizer.Normalize(rawValue, valueType, normalizedFormat.Kind, uses1904Dates);
                return new {
                    reference,
                    rawValue,
                    formattedValue = reference is null ? null : formattedCells.GetValueOrDefault(reference),
                    valueType = valueType ?? "number",
                    styleIndex,
                    style = new {
                        baseStyleIndex = baseFormatIndex,
                        fontId,
                        fillId,
                        borderId,
                        bold = EffectiveBold(fonts, fontId),
                        fontFingerprint = ComponentFingerprint(fonts, fontId, "font"),
                        fillFingerprint = ComponentFingerprint(fills, fillId, "fill"),
                        borderFingerprint = ComponentFingerprint(borders, borderId, "border"),
                        protectionFingerprint = ProtectionFingerprint(locked, hidden),
                        numberFormatId = numId,
                        numberFormat = formatCode,
                        numberFormatEvidence = normalizedFormat,
                        numberFormatFingerprint = NumberFormatFingerprint(
                            NumberFormatFingerprintCode(numId, formatCode, customFormats.ContainsKey(numId)),
                            normalizedFormat.Kind),
                        alignmentFingerprint = AlignmentFingerprint(alignmentEvidence),
                        alignment = alignmentEvidence,
                        horizontalAlignment = alignment?.Horizontal?.InnerText,
                        verticalAlignment = alignment?.Vertical?.InnerText,
                        wrapText = alignment?.WrapText?.Value,
                        shrinkToFit = alignment?.ShrinkToFit?.Value,
                        textRotation = alignment?.TextRotation?.Value,
                        protection = new {
                            applyProtection,
                            locked,
                            hidden
                        }
                    },
                    normalizedValue,
                    formula = formula is null ? null : new { text = formula.Text, type = formula.FormulaType?.InnerText, sharedIndex = formula.SharedIndex?.Value, reference = formula.Reference?.Value },
                };
            }).ToList();
            var view = ws.SheetViews?.Elements<SheetView>().FirstOrDefault();
            var setup = ws.GetFirstChild<PageSetup>();
            var margins = ws.GetFirstChild<PageMargins>();
            var printArea = wb.Workbook.DefinedNames?.Elements<DefinedName>()
                .FirstOrDefault(x => x.Name?.Value == "_xlnm.Print_Area" && x.LocalSheetId?.Value == (uint)sheetIndex)?.Text;
            var printTitles = wb.Workbook.DefinedNames?.Elements<DefinedName>()
                .FirstOrDefault(x => x.Name?.Value == "_xlnm.Print_Titles" && x.LocalSheetId?.Value == (uint)sheetIndex)?.Text;
            var repeatRows = ExtractPrintTitleReference(printTitles, @"^\$[1-9]\d*:\$[1-9]\d*$");
            var repeatCols = ExtractPrintTitleReference(printTitles, @"^\$[A-Za-z]{1,3}:\$[A-Za-z]{1,3}$");
            var breakBeforeRows = ws.GetFirstChild<RowBreaks>()?.Elements<Break>()
                .Where(item => item.ManualPageBreak?.Value == true && item.Id?.Value is not null)
                .Select(item => item.Id!.Value + 1U)
                .OrderBy(row => row)
                .ToList() ?? [];
            return new {
                name = sheet.Name?.Value, state = sheet.State?.InnerText, dimension = ws.SheetDimension?.Reference?.Value,
                mergedRanges = ws.Elements<MergeCells>().SelectMany(x => x.Elements<MergeCell>()).Select(x => x.Reference?.Value).Where(x => x is not null).ToList(),
                rowDimensions = ws.Descendants<Row>().Where(x => x.CustomHeight?.Value == true || x.Hidden?.Value == true).Select(x => new { row = x.RowIndex?.Value, height = x.Height?.Value, hidden = x.Hidden?.Value }).ToList(),
                columnDimensions = ws.Elements<Columns>().SelectMany(x => x.Elements<Column>()).Select(x => new { min = x.Min?.Value, max = x.Max?.Value, width = x.Width?.Value, hidden = x.Hidden?.Value }).ToList(),
                sheetView = view is null ? null : new { workbookViewId = view.WorkbookViewId?.Value, view = view.View?.InnerText, showGridLines = view.ShowGridLines?.Value, zoomScale = view.ZoomScale?.Value, topLeftCell = view.TopLeftCell?.Value },
                print = new { area = printArea, normalizedArea = NormalizeDefinedNameText(printArea), repeatRows, normalizedRepeatRows = NormalizeDefinedNameText(repeatRows), repeatCols, normalizedRepeatCols = NormalizeDefinedNameText(repeatCols), breakBeforeRows, orientation = setup?.Orientation?.InnerText, paperSize = setup?.PaperSize?.Value, scale = setup?.Scale?.Value, fitToWidth = setup?.FitToWidth?.Value, fitToHeight = setup?.FitToHeight?.Value, margins = margins is null ? null : new { left = margins.Left?.Value, right = margins.Right?.Value, top = margins.Top?.Value, bottom = margins.Bottom?.Value, header = margins.Header?.Value, footer = margins.Footer?.Value } },
                cells
            };
        }).ToList();
        var sheetNames = sheets.Select(sheet => sheet.name ?? string.Empty).ToList();
        // v1 attests the validator-required name/scope/text/visibility semantics.
        // Macro/function metadata is intentionally outside this evidence contract.
        var definedNames = wb.Workbook.DefinedNames?.Elements<DefinedName>()
            .Select(name => new {
                name = name.Name?.Value ?? throw new InvalidDataException("Workbook defined name is missing its name."),
                localSheetName = LocalSheetName(name.LocalSheetId?.Value, sheetNames),
                text = name.Text ?? string.Empty,
                normalizedText = NormalizeDefinedNameText(name.Text ?? string.Empty),
                hidden = name.Hidden?.Value ?? false
            })
            .OrderBy(name => name.name, StringComparer.Ordinal)
            .ThenBy(name => name.localSheetName, StringComparer.Ordinal)
            .ThenBy(name => name.text, StringComparer.Ordinal)
            .ThenBy(name => name.hidden)
            .ToList() ?? [];
        return new { schema = "tiwater.xlsx.evidence/v1", toolVersion = XlsxToolVersion.Current, file = Path.GetFullPath(path), dateSystem = uses1904Dates ? "1904" : "1900", definedNames, sheets };
    }

    private static string? ExtractPrintTitleReference(string? text, string rangePattern)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var start = 0;
        var quoted = false;
        for (var index = 0; index <= text.Length; index++)
        {
            if (index < text.Length && text[index] == '\'')
            {
                if (quoted && index + 1 < text.Length && text[index + 1] == '\'')
                {
                    index++;
                    continue;
                }
                quoted = !quoted;
                continue;
            }
            if (index < text.Length && (text[index] != ',' || quoted)) continue;
            var reference = text[start..index].Trim();
            var separator = reference.LastIndexOf('!');
            if (separator > 0 && System.Text.RegularExpressions.Regex.IsMatch(
                reference[(separator + 1)..].Trim(),
                rangePattern,
                System.Text.RegularExpressions.RegexOptions.CultureInvariant))
                return reference;
            start = index + 1;
        }
        return null;
    }

    private static string? NormalizeDefinedNameText(string? text)
    {
        if (text is null) return null;
        var value = text.Trim();
        var prefix = value.StartsWith('=') ? "=" : string.Empty;
        var body = prefix.Length == 0 ? value : value[1..];
        string sheet;
        string reference;
        if (body.StartsWith('\''))
        {
            var decoded = new StringBuilder();
            var index = 1;
            for (; index < body.Length; index++)
            {
                if (body[index] != '\'') { decoded.Append(body[index]); continue; }
                if (index + 1 < body.Length && body[index + 1] == '\'') { decoded.Append('\''); index++; continue; }
                break;
            }
            if (index + 1 >= body.Length || body[index + 1] != '!') return value;
            sheet = decoded.ToString();
            reference = body[(index + 2)..];
        }
        else
        {
            var separator = body.IndexOf('!');
            if (separator <= 0) return value;
            sheet = body[..separator];
            if (sheet.Any(character => !(char.IsLetterOrDigit(character) || character is '_' or '.'))) return value;
            reference = body[(separator + 1)..];
        }
        if (sheet.Length == 0 || reference.Length == 0) return value;
        return $"{prefix}'{sheet.Replace("'", "''", StringComparison.Ordinal)}'!{reference}";
    }

    private static string? LocalSheetName(uint? localSheetId, IReadOnlyList<string> sheetNames)
    {
        if (localSheetId is null) return null;
        if (localSheetId.Value >= sheetNames.Count)
            throw new InvalidDataException($"Defined name localSheetId is out of range: {localSheetId.Value}");
        return sheetNames[(int)localSheetId.Value];
    }

    private static bool EffectiveBold(IReadOnlyList<Font> fonts, uint fontId)
    {
        if (fonts.Count == 0 && fontId == 0) return false;
        if (fontId >= fonts.Count) throw new InvalidDataException($"Workbook font id is out of range: {fontId}");
        var bold = fonts[(int)fontId].GetFirstChild<Bold>();
        return bold is not null && (bold.Val?.Value ?? true);
    }

    private static Protection? EffectiveProtection(CellFormat? format, CellFormat? baseFormat)
    {
        if (format?.ApplyProtection?.Value == false) return baseFormat?.Protection;
        if (format?.Protection is not null) return format.Protection;
        return format?.ApplyProtection?.Value == true ? null : baseFormat?.Protection;
    }

    private static Alignment? EffectiveAlignment(CellFormat? format, CellFormat? baseFormat)
    {
        if (format?.ApplyAlignment?.Value == false) return baseFormat?.Alignment;
        if (format?.Alignment is not null) return format.Alignment;
        return format?.ApplyAlignment?.Value == true ? null : baseFormat?.Alignment;
    }

    private static AlignmentEvidence ResolveAlignment(Alignment? alignment) => new(
        alignment?.Horizontal?.InnerText ?? "general",
        alignment?.Vertical?.InnerText ?? "bottom",
        alignment?.TextRotation?.Value ?? 0U,
        alignment?.WrapText?.Value ?? false,
        alignment?.ShrinkToFit?.Value ?? false,
        alignment?.Indent?.Value ?? 0U,
        alignment?.RelativeIndent?.Value ?? 0,
        alignment?.JustifyLastLine?.Value ?? false,
        alignment?.ReadingOrder?.Value ?? 0U,
        ParseBoolean(alignment?.MergeCell?.Value));

    private static bool ParseBoolean(string? value) => value switch
    {
        null or "0" or "false" => false,
        "1" or "true" => true,
        _ => throw new InvalidDataException($"Invalid OpenXML boolean value: {value}")
    };

    private static string ComponentFingerprint<T>(IReadOnlyList<T> components, uint id, string kind) where T : OpenXmlElement
    {
        if (components.Count == 0 && id == 0) return Sha256($"implicit:{kind}:default");
        if (id >= components.Count) throw new InvalidDataException($"Workbook {kind} id is out of range: {id}");
        var canonical = new StringBuilder();
        AppendCanonicalElement(canonical, components[(int)id]);
        return Sha256(canonical.ToString());
    }

    private static string ProtectionFingerprint(bool locked, bool hidden)
        => Sha256($"protection:locked={(locked ? 1 : 0)};hidden={(hidden ? 1 : 0)}");

    private static string NumberFormatFingerprint(string exactCode, string kind)
    {
        var canonical = new StringBuilder();
        AppendToken(canonical, "number-format");
        AppendToken(canonical, exactCode);
        AppendToken(canonical, kind);
        return Sha256(canonical.ToString());
    }

    private static string NumberFormatFingerprintCode(uint id, string exactCode, bool custom)
        => !custom && id is 14 or 58 ? "wps-locale-short-date" : exactCode;

    private static string AlignmentFingerprint(AlignmentEvidence alignment)
    {
        var canonical = new StringBuilder();
        AppendToken(canonical, "alignment");
        AppendToken(canonical, alignment.Horizontal);
        AppendToken(canonical, alignment.Vertical);
        AppendToken(canonical, alignment.TextRotation.ToString(CultureInfo.InvariantCulture));
        AppendToken(canonical, alignment.WrapText ? "1" : "0");
        AppendToken(canonical, alignment.ShrinkToFit ? "1" : "0");
        AppendToken(canonical, alignment.Indent.ToString(CultureInfo.InvariantCulture));
        AppendToken(canonical, alignment.RelativeIndent.ToString(CultureInfo.InvariantCulture));
        AppendToken(canonical, alignment.JustifyLastLine ? "1" : "0");
        AppendToken(canonical, alignment.ReadingOrder.ToString(CultureInfo.InvariantCulture));
        AppendToken(canonical, alignment.MergeCell ? "1" : "0");
        return Sha256(canonical.ToString());
    }

    private static void AppendCanonicalElement(StringBuilder output, OpenXmlElement element)
    {
        AppendToken(output, "element");
        AppendToken(output, element.NamespaceUri);
        AppendToken(output, element.LocalName);
        foreach (var attribute in element.GetAttributes()
            .OrderBy(attribute => attribute.NamespaceUri, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.LocalName, StringComparer.Ordinal))
        {
            AppendToken(output, "attribute");
            AppendToken(output, attribute.NamespaceUri);
            AppendToken(output, attribute.LocalName);
            AppendToken(output, NormalizeAttributeValue(attribute.Value ?? string.Empty));
        }
        if (!element.HasChildren && !string.IsNullOrEmpty(element.InnerText))
        {
            AppendToken(output, "text");
            AppendToken(output, element.InnerText);
        }
        foreach (var child in element.ChildElements) AppendCanonicalElement(output, child);
        AppendToken(output, "end");
    }

    private static string NormalizeAttributeValue(string value) => value switch
    {
        "true" => "1",
        "false" => "0",
        _ => value
    };

    private static void AppendToken(StringBuilder output, string? value)
    {
        value ??= string.Empty;
        output.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value);
    }

    private static string Sha256(string value)
        => System.Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    private sealed record AlignmentEvidence(
        string Horizontal,
        string Vertical,
        uint TextRotation,
        bool WrapText,
        bool ShrinkToFit,
        uint Indent,
        int RelativeIndent,
        bool JustifyLastLine,
        uint ReadingOrder,
        bool MergeCell);

}
