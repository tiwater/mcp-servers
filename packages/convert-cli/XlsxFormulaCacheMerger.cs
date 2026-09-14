using System.IO.Compression;
using System.Xml.Linq;

namespace Dockit.Convert;

internal static class XlsxFormulaCacheMerger
{
    private static readonly XNamespace Spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    internal static void Merge(string input, string recalculated, string output)
    {
        input = Path.GetFullPath(input);
        recalculated = Path.GetFullPath(recalculated);
        output = Path.GetFullPath(output);
        if (!File.Exists(input) || !File.Exists(recalculated))
            throw new InvalidOperationException("XLSX formula-cache merge requires both input and recalculated workbooks.");
        if (string.Equals(input, output, StringComparison.Ordinal)
            || string.Equals(recalculated, output, StringComparison.Ordinal))
            throw new InvalidOperationException("XLSX formula-cache merge requires a distinct output path.");

        var worksheetUpdates = BuildWorksheetUpdates(input, recalculated);
        var outputDirectory = Path.GetDirectoryName(output);
        if (!string.IsNullOrWhiteSpace(outputDirectory)) Directory.CreateDirectory(outputDirectory);
        var temporary = Path.Combine(outputDirectory ?? Path.GetTempPath(), $".{Path.GetFileName(output)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.Copy(input, temporary, overwrite: false);
            using (var archive = ZipFile.Open(temporary, ZipArchiveMode.Update))
            {
                foreach (var (entryName, content) in worksheetUpdates)
                {
                    var prior = archive.GetEntry(entryName)
                        ?? throw new InvalidOperationException($"Input XLSX worksheet part is missing during formula-cache merge: {entryName}");
                    var lastWriteTime = prior.LastWriteTime;
                    prior.Delete();
                    var replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                    replacement.LastWriteTime = lastWriteTime;
                    using var stream = replacement.Open();
                    stream.Write(content);
                }
            }
            File.Move(temporary, output, overwrite: true);
        }
        finally
        {
            try { if (File.Exists(temporary)) File.Delete(temporary); } catch { }
        }
    }

    private static Dictionary<string, byte[]> BuildWorksheetUpdates(string input, string recalculated)
    {
        using var sourceArchive = ZipFile.OpenRead(input);
        using var recalculatedArchive = ZipFile.OpenRead(recalculated);
        var sourceWorksheets = sourceArchive.Entries
            .Where(static entry => entry.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal)
                && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(static entry => entry.FullName, StringComparer.Ordinal);
        var recalculatedWorksheets = recalculatedArchive.Entries
            .Where(static entry => entry.FullName.StartsWith("xl/worksheets/", StringComparison.Ordinal)
                && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(static entry => entry.FullName, StringComparer.Ordinal);
        if (!sourceWorksheets.Keys.Order(StringComparer.Ordinal).SequenceEqual(recalculatedWorksheets.Keys.Order(StringComparer.Ordinal), StringComparer.Ordinal))
            throw new InvalidOperationException("ET recalculation changed the XLSX worksheet part inventory.");

        var updates = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var (entryName, sourceEntry) in sourceWorksheets)
        {
            var source = Load(sourceEntry);
            var refreshed = Load(recalculatedWorksheets[entryName]);
            var sourceFormulas = FormulaCells(source, entryName);
            var refreshedFormulas = FormulaCells(refreshed, entryName);
            if (!sourceFormulas.Keys.Order(StringComparer.Ordinal).SequenceEqual(refreshedFormulas.Keys.Order(StringComparer.Ordinal), StringComparer.Ordinal))
                throw new InvalidOperationException($"ET recalculation changed the formula cell inventory: {entryName}");

            foreach (var (reference, sourceCell) in sourceFormulas)
            {
                var refreshedCell = refreshedFormulas[reference];
                if (FormulaSignature(sourceCell.Element(Spreadsheet + "f")!) != FormulaSignature(refreshedCell.Element(Spreadsheet + "f")!))
                    throw new InvalidOperationException($"ET recalculation changed a formula: {entryName}!{reference}");
                ReplaceCachedValue(sourceCell, refreshedCell);
            }
            if (sourceFormulas.Count > 0) updates.Add(entryName, Serialize(source));
        }
        return updates;
    }

    private static XDocument Load(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static Dictionary<string, XElement> FormulaCells(XDocument document, string entryName)
    {
        var result = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in document.Descendants(Spreadsheet + "c").Where(static cell => cell.Element(Spreadsheet + "f") is not null))
        {
            var reference = ((string?)cell.Attribute("r"))?.Trim();
            if (string.IsNullOrWhiteSpace(reference) || !result.TryAdd(reference, cell))
                throw new InvalidOperationException($"XLSX formula cell identity is missing or duplicate: {entryName}");
        }
        return result;
    }

    private static string FormulaSignature(XElement formula)
        => $"{formula.Value}\n{string.Join("\n", formula.Attributes().OrderBy(static attribute => attribute.Name.ToString(), StringComparer.Ordinal).Select(static attribute => $"{attribute.Name}={attribute.Value}"))}";

    private static void ReplaceCachedValue(XElement sourceCell, XElement refreshedCell)
    {
        var refreshedType = refreshedCell.Attribute("t");
        sourceCell.Attribute("t")?.Remove();
        if (refreshedType is not null) sourceCell.Add(new XAttribute("t", refreshedType.Value));

        var prior = sourceCell.Element(Spreadsheet + "v");
        var refreshed = refreshedCell.Element(Spreadsheet + "v");
        if (refreshed is null)
        {
            prior?.Remove();
            return;
        }
        if (prior is not null) prior.ReplaceWith(new XElement(refreshed));
        else sourceCell.Element(Spreadsheet + "f")!.AddAfterSelf(new XElement(refreshed));
    }

    private static byte[] Serialize(XDocument document)
    {
        using var stream = new MemoryStream();
        document.Save(stream, SaveOptions.DisableFormatting);
        return stream.ToArray();
    }
}
