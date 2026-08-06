using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Dockit.Xlsx;

public sealed record XlsxNormalizedValue(string Kind, string? Iso8601);

internal sealed record SpreadsheetNumberFormatEvidence(
    uint Id,
    string Code,
    string NormalizedCode,
    string Source,
    string Kind,
    bool IsDateLike);

internal static class SpreadsheetValueNormalizer
{
    internal static XlsxNormalizedValue FromOpenXmlCell(
        string? rawValue,
        string? dataType,
        WorkbookStylesPart? stylesPart,
        uint? styleIndex,
        bool uses1904Dates)
        => Normalize(rawValue, dataType, ResolveNumberFormatKind(stylesPart, styleIndex), uses1904Dates);

    internal static XlsxNormalizedValue Normalize(
        string? rawValue,
        string? dataType,
        string formatKind,
        bool uses1904Dates)
    {
        if (rawValue is null) return new XlsxNormalizedValue("blank", null);
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
            return new XlsxNormalizedValue(formatKind, iso8601);
        }
        return new XlsxNormalizedValue(dataType ?? "number", null);
    }

    internal static XlsxNormalizedValue FromLegacyDate(DateTime value)
        => new("date", value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));

    private static string ResolveNumberFormatKind(WorkbookStylesPart? stylesPart, uint? styleIndex)
    {
        var styles = stylesPart?.Stylesheet;
        var formats = styles?.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        if (styleIndex is null || styleIndex.Value >= formats.Count) return "general";

        var format = formats[(int)styleIndex.Value];
        var baseFormats = styles?.CellStyleFormats?.Elements<CellFormat>().ToList() ?? [];
        var baseFormatIndex = format.FormatId?.Value;
        var baseFormat = baseFormatIndex is not null && baseFormatIndex.Value < baseFormats.Count
            ? baseFormats[(int)baseFormatIndex.Value]
            : null;
        var numberFormatId = EffectiveComponentId(
            format.NumberFormatId?.Value,
            format.ApplyNumberFormat?.Value,
            baseFormat?.NumberFormatId?.Value);
        var customFormat = styles?.NumberingFormats?.Elements<NumberingFormat>()
            .Where(item => item.NumberFormatId?.Value == numberFormatId)
            .Select(item => item.FormatCode?.Value)
            .LastOrDefault(item => !string.IsNullOrWhiteSpace(item));
        var formatCode = customFormat ?? BuiltInFormat(numberFormatId) ?? $"builtin:{numberFormatId}";
        return NormalizeNumberFormat(numberFormatId, formatCode, customFormat is not null).Kind;
    }

    internal static uint EffectiveComponentId(uint? directId, bool? applyDirect, uint? baseId)
    {
        if (applyDirect == false) return baseId ?? 0U;
        if (directId is not null) return directId.Value;
        return applyDirect == true ? 0U : baseId ?? 0U;
    }

    internal static SpreadsheetNumberFormatEvidence NormalizeNumberFormat(uint id, string code, bool custom)
    {
        var normalizedCode = !custom && id is 14 or 58
            ? "wps-locale-short-date"
            : string.Concat(code.Where(character => !char.IsWhiteSpace(character))).ToLowerInvariant();
        var kind = ClassifyNumberFormat(id, code);
        return new SpreadsheetNumberFormatEvidence(
            id,
            code,
            normalizedCode,
            custom ? "custom" : "builtIn",
            kind,
            kind is "date" or "time" or "datetime");
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
        for (var index = 0; index < code.Length; index++)
        {
            var character = code[index];
            if (character == '"') { quoted = !quoted; continue; }
            if (quoted) continue;
            if (character == '\\' || character == '_' || character == '*') { index++; continue; }
            if (character == '[')
            {
                var end = code.IndexOf(']', index + 1);
                if (end < 0) break;
                var token = code[(index + 1)..end];
                if (token.All(item => item is 'h' or 'H' or 'm' or 'M' or 's' or 'S')) result.Append(token);
                index = end;
                continue;
            }
            result.Append(character);
        }
        return result.ToString();
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

    internal static string? BuiltInFormat(uint id) => id switch
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
