using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class TemplateTransformValidator
{
    private static readonly Regex WhitespacePattern = new(@"\s+", RegexOptions.Compiled);

    public static TemplateTransformValidationReport Validate(string sourceTemplate, string targetTemplate)
    {
        var sourcePath = Path.GetFullPath(sourceTemplate);
        var targetPath = Path.GetFullPath(targetTemplate);
        var sourceSlots = ExtractBodyFieldSlots(sourcePath);
        var targetSlots = ExtractBodyFieldSlots(targetPath);
        var warnings = new List<string>();
        var errors = new List<string>();

        if (sourceSlots.Count != targetSlots.Count)
        {
            errors.Add($"Body field slot count differs: {sourceSlots.Count} vs {targetSlots.Count}.");
        }

        var mismatches = new List<TemplateSlotMismatch>();
        foreach (var pair in sourceSlots.Zip(targetSlots, (sourceSlot, targetSlot) => (sourceSlot, targetSlot)))
        {
            if (!StringComparer.Ordinal.Equals(pair.sourceSlot.Text, pair.targetSlot.Text))
            {
                mismatches.Add(new TemplateSlotMismatch(pair.sourceSlot.Path, pair.sourceSlot.Text, pair.targetSlot.Text));
            }
        }

        if (mismatches.Count > 0)
        {
            errors.Add($"Found {mismatches.Count} mismatched body field slots.");
        }

        var sourceEmptySlots = sourceSlots.Count(slot => slot.IsEmptyInputSlot);
        var targetEmptySlots = targetSlots.Count(slot => slot.IsEmptyInputSlot);
        if (sourceEmptySlots != targetEmptySlots)
        {
            errors.Add($"Empty input slot count differs: {sourceEmptySlots} vs {targetEmptySlots}.");
        }

        var sourceInspection = Inspector.Inspect(sourcePath);
        var targetInspection = Inspector.Inspect(targetPath);
        if (sourceInspection.Content.HeaderPartCount != targetInspection.Content.HeaderPartCount ||
            sourceInspection.Content.FooterPartCount != targetInspection.Content.FooterPartCount)
        {
            warnings.Add("Header/footer part counts differ. Body field transform may still be valid, but document chrome changed.");
        }

        if (sourceInspection.Package.PartCount != targetInspection.Package.PartCount)
        {
            warnings.Add("Package part counts differ.");
        }

        if (sourceInspection.Content.TableCount != targetInspection.Content.TableCount)
        {
            warnings.Add($"Overall table count differs: {sourceInspection.Content.TableCount} vs {targetInspection.Content.TableCount}. This does not block transform if body slots still match.");
        }

        return new TemplateTransformValidationReport(
            SourceTemplate: sourcePath,
            TargetTemplate: targetPath,
            IsCompatible: errors.Count == 0,
            SourceBodyFieldSlotCount: sourceSlots.Count,
            TargetBodyFieldSlotCount: targetSlots.Count,
            SourceEmptyInputSlotCount: sourceEmptySlots,
            TargetEmptyInputSlotCount: targetEmptySlots,
            MismatchedBodySlots: mismatches.Take(50).ToList(),
            Errors: errors,
            Warnings: warnings);
    }

    private static List<TemplateFieldSlot> ExtractBodyFieldSlots(string input)
    {
        using var doc = WordprocessingDocument.Open(input, false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Document body not found.");
        var slots = new List<TemplateFieldSlot>();

        var tables = body.Elements<Table>().ToList();
        for (var tableIndex = 0; tableIndex < tables.Count; tableIndex++)
        {
            var rows = tables[tableIndex].Elements<TableRow>().ToList();
            for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var cells = rows[rowIndex].Elements<TableCell>().ToList();
                for (var cellIndex = 0; cellIndex < cells.Count; cellIndex++)
                {
                    var text = Normalize(string.Concat(cells[cellIndex].Descendants<Text>().Select(text => text.Text)));
                    slots.Add(new TemplateFieldSlot(
                        Scope: "body",
                        Path: $"table[{tableIndex}].row[{rowIndex}].cell[{cellIndex}]",
                        Text: text,
                        IsEmptyInputSlot: string.IsNullOrEmpty(text)));
                }
            }
        }

        return slots;
    }

    private static string Normalize(string text)
        => WhitespacePattern.Replace(text, " ").Trim();
}
