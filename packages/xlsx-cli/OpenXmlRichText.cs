using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

internal static class OpenXmlRichText
{
    public static IReadOnlyList<RichTextRunReport>? GetCellRichTextRuns(
        Cell cell,
        SharedStringTable? sharedStringTable,
        Stylesheet? stylesheet = null)
    {
        var cellFont = ResolveCellFont(cell, stylesheet);

        if (cell.InlineString is not null)
        {
            return GetRichTextRuns(cell.InlineString.Elements<Run>(), cellFont);
        }

        var text = cell.CellValue?.Text;
        if (text is null ||
            cell.DataType?.Value != CellValues.SharedString ||
            sharedStringTable is null ||
            !int.TryParse(text, out var index))
        {
            return null;
        }

        var sharedString = sharedStringTable.ElementAtOrDefault(index);
        return sharedString is null ? null : GetRichTextRuns(sharedString.Elements<Run>(), cellFont);
    }

    private static Font? ResolveCellFont(Cell cell, Stylesheet? stylesheet)
    {
        if (stylesheet is null)
        {
            return null;
        }

        var fonts = stylesheet.Fonts?.Elements<Font>().ToList();
        if (fonts is null || fonts.Count == 0)
        {
            return null;
        }

        var styleIndex = cell.StyleIndex?.Value ?? 0U;
        var formats = stylesheet.CellFormats?.Elements<CellFormat>().ToList() ?? [];
        var format = styleIndex < formats.Count ? formats[(int)styleIndex] : null;
        var baseFormats = stylesheet.CellStyleFormats?.Elements<CellFormat>().ToList() ?? [];
        var baseFormatIndex = format?.FormatId?.Value;
        var baseFormat = baseFormatIndex is not null && baseFormatIndex.Value < baseFormats.Count
            ? baseFormats[(int)baseFormatIndex.Value]
            : null;
        var fontId = SpreadsheetValueNormalizer.EffectiveComponentId(
            format?.FontId?.Value,
            format?.ApplyFont?.Value,
            baseFormat?.FontId?.Value);
        return fontId < fonts.Count ? fonts[(int)fontId] : null;
    }

    private static IReadOnlyList<RichTextRunReport>? GetRichTextRuns(IEnumerable<Run> runs, Font? cellFont)
    {
        var reports = new List<RichTextRunReport>();
        foreach (var run in runs)
        {
            var text = string.Concat(run.Elements<Text>().Select(part => part.Text));
            if (string.IsNullOrEmpty(text))
            {
                text = run.InnerText;
            }

            if (string.IsNullOrEmpty(text))
            {
                continue;
            }

            var properties = run.RunProperties;
            if (properties is null)
            {
                var cellFontName = cellFont?.GetFirstChild<FontName>()?.Val?.Value;
                var cellColor = cellFont?.GetFirstChild<Color>();
                var cellUnderline = cellFont?.GetFirstChild<Underline>();
                var cellBold = cellFont?.GetFirstChild<Bold>();
                var cellItalic = cellFont?.GetFirstChild<Italic>();
                reports.Add(new RichTextRunReport(
                    text,
                    cellFontName,
                    GetColor(cellColor),
                    GetUnderline(cellUnderline),
                    cellBold is not null && (cellBold.Val?.Value ?? true),
                    cellItalic is not null && (cellItalic.Val?.Value ?? true)));
                continue;
            }

            var font = properties.GetFirstChild<RunFont>();
            var color = properties.GetFirstChild<Color>();
            var underline = properties.GetFirstChild<Underline>();
            var bold = properties.GetFirstChild<Bold>();
            var italic = properties.GetFirstChild<Italic>();
            reports.Add(new RichTextRunReport(
                text,
                font?.Val?.Value,
                GetColor(color),
                GetUnderline(underline),
                bold is not null && (bold.Val?.Value ?? true),
                italic is not null && (italic.Val?.Value ?? true)));
        }

        return reports.Count == 0 ? null : reports;
    }

    private static string? GetColor(Color? color)
    {
        if (color is null)
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(color.Rgb?.Value))
        {
            return color.Rgb.Value;
        }

        if (color.Indexed?.Value is not null)
        {
            return $"indexed:{color.Indexed.Value}";
        }

        if (color.Theme?.Value is not null)
        {
            return $"theme:{color.Theme.Value}";
        }

        return color.Auto?.Value == true ? "auto" : null;
    }

    private static string? GetUnderline(Underline? underline)
    {
        if (underline is null)
        {
            return null;
        }

        if (underline.Val?.Value == UnderlineValues.None)
        {
            return null;
        }

        return underline.Val?.InnerText ?? "single";
    }
}
