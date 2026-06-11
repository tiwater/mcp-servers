using DocumentFormat.OpenXml.Spreadsheet;

namespace Dockit.Xlsx;

internal static class OpenXmlRichText
{
    public static IReadOnlyList<RichTextRunReport>? GetCellRichTextRuns(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.InlineString is not null)
        {
            return GetRichTextRuns(cell.InlineString.Elements<Run>());
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
        return sharedString is null ? null : GetRichTextRuns(sharedString.Elements<Run>());
    }

    private static IReadOnlyList<RichTextRunReport>? GetRichTextRuns(IEnumerable<Run> runs)
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
            var font = properties?.GetFirstChild<RunFont>();
            var color = properties?.GetFirstChild<Color>();
            var underline = properties?.GetFirstChild<Underline>();
            var bold = properties?.GetFirstChild<Bold>();
            var italic = properties?.GetFirstChild<Italic>();
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

        return underline.Val?.InnerText ?? "single";
    }
}
