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
            var sourceSharedMasters = SharedFormulaMasters(source, entryName);
            var refreshedSharedMasters = SharedFormulaMasters(refreshed, entryName);
            if (!sourceFormulas.Keys.Order(StringComparer.Ordinal).SequenceEqual(refreshedFormulas.Keys.Order(StringComparer.Ordinal), StringComparer.Ordinal))
                throw new InvalidOperationException($"ET recalculation changed the formula cell inventory: {entryName}");

            foreach (var (reference, sourceCell) in sourceFormulas)
            {
                var refreshedCell = refreshedFormulas[reference];
                var sourceFormula = ResolveFormula(sourceCell, reference, sourceSharedMasters, entryName);
                var refreshedFormula = ResolveFormula(refreshedCell, reference, refreshedSharedMasters, entryName);
                if (FormulaSignature(sourceCell.Element(Spreadsheet + "f")!, sourceFormula) != FormulaSignature(refreshedCell.Element(Spreadsheet + "f")!, refreshedFormula))
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

    private sealed record SharedFormulaMaster(string Reference, string Formula);

    private static Dictionary<uint, SharedFormulaMaster> SharedFormulaMasters(XDocument document, string entryName)
    {
        var result = new Dictionary<uint, SharedFormulaMaster>();
        foreach (var cell in document.Descendants(Spreadsheet + "c"))
        {
            var formula = cell.Element(Spreadsheet + "f");
            if (formula is null || !string.Equals((string?)formula.Attribute("t"), "shared", StringComparison.OrdinalIgnoreCase)) continue;
            var text = formula.Value;
            if (string.IsNullOrWhiteSpace(text)) continue;
            var reference = ((string?)cell.Attribute("r"))?.Trim();
            var index = (string?)formula.Attribute("si");
            if (string.IsNullOrWhiteSpace(reference) || !uint.TryParse(index, out var sharedIndex))
                throw new InvalidOperationException($"ET recalculation produced an invalid shared formula master: {entryName}");
            if (!result.TryAdd(sharedIndex, new SharedFormulaMaster(reference, text)))
                throw new InvalidOperationException($"ET recalculation produced duplicate shared formula masters: {entryName} si={sharedIndex}");
        }
        return result;
    }

    private static string ResolveFormula(XElement cell, string reference, Dictionary<uint, SharedFormulaMaster> masters, string entryName)
    {
        var formula = cell.Element(Spreadsheet + "f")
            ?? throw new InvalidOperationException($"XLSX formula cell is missing during formula-cache merge: {entryName}!{reference}");
        var type = (string?)formula.Attribute("t");
        var text = formula.Value;
        if (!string.Equals(type, "shared", StringComparison.OrdinalIgnoreCase)) return text;
        var indexText = (string?)formula.Attribute("si");
        if (!uint.TryParse(indexText, out var sharedIndex) || !masters.TryGetValue(sharedIndex, out var master))
            throw new InvalidOperationException($"ET recalculation produced a shared formula follower without a valid master: {entryName}!{reference}");
        if (!string.IsNullOrWhiteSpace(text)) return text;
        var (masterColumn, masterRow) = ParseCellReference(master.Reference);
        var (column, row) = ParseCellReference(reference);
        return TranslateRelativeFormulaReferences(master.Formula, row - masterRow, column - masterColumn);
    }

    private static string FormulaSignature(XElement formula, string resolvedFormula)
        => $"{NormalizeFormula(resolvedFormula)}\n{string.Join("\n", formula.Attributes()
            .Where(attribute => !IsSharedFormulaAttribute(formula, attribute))
            .OrderBy(static attribute => attribute.Name.ToString(), StringComparer.Ordinal)
            .Select(static attribute => $"{attribute.Name}={attribute.Value}"))}";

    private static bool IsSharedFormulaAttribute(XElement formula, XAttribute attribute)
        => string.Equals((string?)formula.Attribute("t"), "shared", StringComparison.OrdinalIgnoreCase)
            && attribute.Name.LocalName is "t" or "si" or "ref";

    private static string TranslateRelativeFormulaReferences(string formula, int rowOffset, int columnOffset)
    {
        var result = new System.Text.StringBuilder(formula.Length);
        var segmentStart = 0;
        var inDoubleQuote = false;
        var inSingleQuote = false;
        for (var index = 0; index < formula.Length; index++)
        {
            var character = formula[index];
            if (character == '"' && !inSingleQuote)
            {
                result.Append(TranslateUnquotedFormulaSegment(formula[segmentStart..index], rowOffset, columnOffset));
                result.Append(character);
                if (inDoubleQuote && index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    result.Append(formula[++index]);
                }
                else inDoubleQuote = !inDoubleQuote;
                segmentStart = index + 1;
                continue;
            }
            if (character == '\'' && !inDoubleQuote)
            {
                result.Append(TranslateUnquotedFormulaSegment(formula[segmentStart..index], rowOffset, columnOffset));
                result.Append(character);
                while (++index < formula.Length)
                {
                    result.Append(formula[index]);
                    if (formula[index] != '\'') continue;
                    if (index + 1 < formula.Length && formula[index + 1] == '\'') result.Append(formula[++index]);
                    else break;
                }
                segmentStart = index + 1;
            }
        }
        result.Append(TranslateUnquotedFormulaSegment(formula[segmentStart..], rowOffset, columnOffset));
        return result.ToString();
    }

    private static string TranslateUnquotedFormulaSegment(string formula, int rowOffset, int columnOffset)
        => System.Text.RegularExpressions.Regex.Replace(
            formula,
            @"(?<![A-Za-z0-9_])(?<column>\$?[A-Za-z]{1,3})(?<row>\$?\d+)(?![A-Za-z0-9_])",
            match =>
            {
                var columnToken = match.Groups["column"].Value;
                var rowToken = match.Groups["row"].Value;
                var absoluteColumn = columnToken.StartsWith('$');
                var absoluteRow = rowToken.StartsWith('$');
                var columnName = absoluteColumn ? columnToken[1..] : columnToken;
                var rowText = absoluteRow ? rowToken[1..] : rowToken;
                var translatedColumn = absoluteColumn ? columnName : ColumnIndexToName(GetColumnIndex(columnName) + columnOffset);
                var translatedRow = absoluteRow ? rowText : (int.Parse(rowText, System.Globalization.CultureInfo.InvariantCulture) + rowOffset).ToString(System.Globalization.CultureInfo.InvariantCulture);
                return $"{(absoluteColumn ? "$" : string.Empty)}{translatedColumn}{(absoluteRow ? "$" : string.Empty)}{translatedRow}";
            });

    private static (int Column, int Row) ParseCellReference(string reference)
    {
        var columnText = new string(reference.TakeWhile(char.IsLetter).ToArray());
        var rowText = reference[columnText.Length..];
        if (columnText.Length == 0 || !int.TryParse(rowText, out var row) || row < 1)
            throw new InvalidOperationException($"Invalid shared formula cell reference: {reference}");
        return (GetColumnIndex(columnText), row);
    }

    private static int GetColumnIndex(string columnName)
    {
        var result = 0;
        foreach (var character in columnName.ToUpperInvariant()) result = result * 26 + character - 'A' + 1;
        return result;
    }

    private static string ColumnIndexToName(int column)
    {
        if (column < 1) throw new InvalidOperationException($"Invalid translated formula column: {column}");
        var result = new System.Text.StringBuilder();
        for (var current = column; current > 0; current = (current - 1) / 26) result.Insert(0, (char)('A' + (current - 1) % 26));
        return result.ToString();
    }

    // ET may canonicalize a redundant unary-plus sequence (for example `A+ +B`
    // or the compact `A++B`) while recalculating.  That rewrite is semantically
    // equivalent, unlike changing an operand, function, or reference.  Compare
    // only this narrow lexical normalization and retain all other formula text.
    private static string NormalizeFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return formula;
        var result = new System.Text.StringBuilder(formula.Length);
        var quoted = false;
        for (var index = 0; index < formula.Length; index++)
        {
            var character = formula[index];
            if (character == '"')
            {
                result.Append(character);
                if (quoted && index + 1 < formula.Length && formula[index + 1] == '"')
                    result.Append(formula[++index]);
                else
                    quoted = !quoted;
                continue;
            }
            if (character == '\'' && !quoted)
            {
                result.Append(character);
                while (++index < formula.Length)
                {
                    result.Append(formula[index]);
                    if (formula[index] != '\'') continue;
                    if (index + 1 < formula.Length && formula[index + 1] == '\'')
                        result.Append(formula[++index]);
                    else break;
                }
                continue;
            }
            if (!quoted && character == '+' && result.Length > 0 && result[^1] == '+') continue;
            result.Append(character);
        }
        return result.ToString();
    }

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
