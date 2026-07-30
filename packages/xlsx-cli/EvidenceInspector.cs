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
                var numId = EffectiveComponentId(format?.NumberFormatId?.Value, format?.ApplyNumberFormat?.Value, baseFormat?.NumberFormatId?.Value);
                var fontId = EffectiveComponentId(format?.FontId?.Value, format?.ApplyFont?.Value, baseFormat?.FontId?.Value);
                var fillId = EffectiveComponentId(format?.FillId?.Value, format?.ApplyFill?.Value, baseFormat?.FillId?.Value);
                var borderId = EffectiveComponentId(format?.BorderId?.Value, format?.ApplyBorder?.Value, baseFormat?.BorderId?.Value);
                var formatCode = customFormats.GetValueOrDefault(numId) ?? BuiltInFormat(numId) ?? $"builtin:{numId}";
                var normalizedFormat = NormalizeNumberFormat(numId, formatCode, customFormats.ContainsKey(numId));
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
                var normalizedValue = NormalizeValue(rawValue, valueType, normalizedFormat.Kind, uses1904Dates);
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
            var repeatRows = wb.Workbook.DefinedNames?.Elements<DefinedName>()
                .FirstOrDefault(x => x.Name?.Value == "_xlnm.Print_Titles" && x.LocalSheetId?.Value == (uint)sheetIndex)?.Text;
            return new {
                name = sheet.Name?.Value, state = sheet.State?.InnerText, dimension = ws.SheetDimension?.Reference?.Value,
                mergedRanges = ws.Elements<MergeCells>().SelectMany(x => x.Elements<MergeCell>()).Select(x => x.Reference?.Value).Where(x => x is not null).ToList(),
                rowDimensions = ws.Descendants<Row>().Where(x => x.CustomHeight?.Value == true || x.Hidden?.Value == true).Select(x => new { row = x.RowIndex?.Value, height = x.Height?.Value, hidden = x.Hidden?.Value }).ToList(),
                columnDimensions = ws.Elements<Columns>().SelectMany(x => x.Elements<Column>()).Select(x => new { min = x.Min?.Value, max = x.Max?.Value, width = x.Width?.Value, hidden = x.Hidden?.Value }).ToList(),
                sheetView = view is null ? null : new { workbookViewId = view.WorkbookViewId?.Value, view = view.View?.InnerText, showGridLines = view.ShowGridLines?.Value, zoomScale = view.ZoomScale?.Value, topLeftCell = view.TopLeftCell?.Value },
                print = new { area = printArea, normalizedArea = NormalizeDefinedNameText(printArea), repeatRows, normalizedRepeatRows = NormalizeDefinedNameText(repeatRows), orientation = setup?.Orientation?.InnerText, paperSize = setup?.PaperSize?.Value, scale = setup?.Scale?.Value, fitToWidth = setup?.FitToWidth?.Value, fitToHeight = setup?.FitToHeight?.Value, margins = margins is null ? null : new { left = margins.Left?.Value, right = margins.Right?.Value, top = margins.Top?.Value, bottom = margins.Bottom?.Value, header = margins.Header?.Value, footer = margins.Footer?.Value } },
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

    private static uint EffectiveComponentId(uint? directId, bool? applyDirect, uint? baseId)
    {
        if (applyDirect == false) return baseId ?? 0U;
        if (directId is not null) return directId.Value;
        return applyDirect == true ? 0U : baseId ?? 0U;
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

    private static object NormalizeValue(string? rawValue, string? dataType, string formatKind, bool uses1904Dates)
    {
        if (rawValue is null) return new { kind = "blank", iso8601 = (string?)null };
        if ((formatKind is "date" or "time" or "datetime")
            && dataType is not "s" and not "inlineStr" and not "str"
            && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            var value = ExcelSerialDate(serial, uses1904Dates);
            value = new DateTime(
                ((value.Ticks + TimeSpan.TicksPerMillisecond / 2) / TimeSpan.TicksPerMillisecond) * TimeSpan.TicksPerMillisecond,
                value.Kind);
            var iso8601 = formatKind switch
            {
                "date" => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                "time" => value.ToString("HH:mm:ss.fffffff", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.'),
                _ => value.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff", CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.')
            };
            return new { kind = formatKind, iso8601 };
        }
        return new { kind = dataType ?? "number", iso8601 = (string?)null };
    }

    private static DateTime ExcelSerialDate(double serial, bool uses1904Dates)
    {
        if (!double.IsFinite(serial)) throw new InvalidDataException($"Invalid Excel date serial: {serial}");
        if (uses1904Dates) return new DateTime(1904, 1, 1).AddDays(serial);
        // Excel intentionally preserves Lotus 1-2-3's fictitious 1900-02-29 at serial 60.
        // Mapping serial 60 to the preceding real date keeps the evidence a valid ISO date.
        var wholeDays = Math.Floor(serial);
        var fraction = serial - wholeDays;
        var day = new DateTime(1899, 12, 31).AddDays(wholeDays >= 60 ? wholeDays - 1 : wholeDays);
        return day.AddDays(fraction);
    }

    private static NumberFormatEvidence NormalizeNumberFormat(uint id, string code, bool custom)
    {
        var normalizedCode = !custom && id is 14 or 58
            ? "wps-locale-short-date"
            : string.Concat(code.Where(c => !char.IsWhiteSpace(c))).ToLowerInvariant();
        var kind = ClassifyNumberFormat(id, code);
        return new NumberFormatEvidence(id, code, normalizedCode, custom ? "custom" : "builtIn", kind, kind is "date" or "time" or "datetime");
    }

    private static string ClassifyNumberFormat(uint id, string code)
    {
        if (id == 49 || code == "@") return "text";
        if (id is 14 or 58) return "date";
        var semantic = StripLiteralsAndConditions(code).ToLowerInvariant();
        if (id == 0 || string.Equals(semantic, "general", StringComparison.Ordinal)) return "general";
        var hasDate = semantic.IndexOfAny(['y', 'd']) >= 0;
        var hasTime = semantic.Contains('h') || semantic.Contains('s') || semantic.Contains("am/pm", StringComparison.Ordinal);
        if (NPOI.SS.UserModel.DateUtil.IsADateFormat((int)id, code))
        {
            if (hasDate && hasTime) return "datetime";
            if (hasTime) return "time";
            return "date";
        }
        if (semantic.Contains('%')) return "percentage";
        if (semantic.Contains('e')) return "scientific";
        if (semantic.Contains('#') || semantic.Contains('0') || semantic.Contains('?')) return "number";
        return "general";
    }

    private static string StripLiteralsAndConditions(string code)
    {
        var result = new System.Text.StringBuilder(code.Length);
        var quoted = false;
        for (var i = 0; i < code.Length; i++)
        {
            var c = code[i];
            if (c == '"') { quoted = !quoted; continue; }
            if (quoted) continue;
            if (c == '\\' || c == '_' || c == '*') { i++; continue; }
            if (c == '[')
            {
                var end = code.IndexOf(']', i + 1);
                if (end < 0) break;
                var token = code[(i + 1)..end];
                if (token.All(x => x is 'h' or 'H' or 'm' or 'M' or 's' or 'S')) result.Append(token);
                i = end;
                continue;
            }
            result.Append(c);
        }
        return result.ToString();
    }

    private sealed record NumberFormatEvidence(uint Id, string Code, string NormalizedCode, string Source, string Kind, bool IsDateLike);

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

    private static string? BuiltInFormat(uint id) => id switch
    {
        0 => "General", 1 => "0", 2 => "0.00", 3 => "#,##0", 4 => "#,##0.00",
        9 => "0%", 10 => "0.00%", 11 => "0.00E+00", 12 => "# ?/?", 13 => "# ??/??",
        14 => "m/d/yy", 15 => "d-mmm-yy", 16 => "d-mmm", 17 => "mmm-yy",
        18 => "h:mm AM/PM", 19 => "h:mm:ss AM/PM", 20 => "h:mm", 21 => "h:mm:ss",
        22 => "m/d/yy h:mm", 37 => "#,##0 ;(#,##0)", 38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)", 40 => "#,##0.00;[Red](#,##0.00)", 45 => "mm:ss",
        46 => "[h]:mm:ss", 47 => "mmss.0", 48 => "##0.0E+0", 49 => "@",
        58 => "builtin:58", _ => null
    };
}
