using System.IO.Compression;
using System.Xml.Linq;

namespace Dockit.Convert;

internal static class DocxWpsRenderNormalizer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static string Prepare(string input, string temporaryRoot)
    {
        if (!string.Equals(Path.GetExtension(input), ".docx", StringComparison.OrdinalIgnoreCase))
            return input;

        var replacements = PageNumberStoryParts(input);
        if (replacements.Count == 0) return input;

        var prepared = Path.Combine(temporaryRoot, "render-input.docx");
        File.Copy(input, prepared, overwrite: false);
        using var archive = ZipFile.Open(prepared, ZipArchiveMode.Update);
        foreach (var (partName, document) in replacements)
        {
            var entry = archive.GetEntry(partName)
                ?? throw new InvalidOperationException($"DOCX package is missing {partName}.");
            entry.Delete();
            var replacement = archive.CreateEntry(partName, CompressionLevel.Optimal);
            using var stream = replacement.Open();
            document.Save(stream, SaveOptions.DisableFormatting);
        }
        return prepared;
    }

    private static IReadOnlyDictionary<string, XDocument> PageNumberStoryParts(string input)
    {
        using var archive = ZipFile.OpenRead(input);
        var result = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries.Where(entry =>
                     IsHeaderOrFooter(entry.FullName)).OrderBy(entry => entry.FullName, StringComparer.Ordinal))
        {
            using var stream = entry.Open();
            var document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            var changed = false;
            foreach (var control in document.Descendants(W + "sdt").Reverse().ToArray())
            {
                var gallery = (string?)control.Element(W + "sdtPr")
                    ?.Element(W + "docPartObj")?.Element(W + "docPartGallery")?.Attribute(W + "val");
                var content = control.Element(W + "sdtContent");
                if (content is null || gallery is null
                    || !gallery.StartsWith("Page Numbers", StringComparison.OrdinalIgnoreCase)
                    || !ContainsPageField(content)) continue;
                control.ReplaceWith(content.Nodes().ToArray());
                changed = true;
            }
            if (changed) result.Add(entry.FullName, document);
        }
        return result;
    }

    private static bool IsHeaderOrFooter(string partName)
    {
        if (!partName.StartsWith("word/", StringComparison.OrdinalIgnoreCase)
            || !partName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return false;
        var fileName = Path.GetFileNameWithoutExtension(partName);
        return fileName.StartsWith("header", StringComparison.OrdinalIgnoreCase)
            || fileName.StartsWith("footer", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsPageField(XElement content)
        => content.Descendants(W + "instrText").Select(element => element.Value.Trim())
            .Any(value => value.Equals("PAGE", StringComparison.OrdinalIgnoreCase)
                || value.Equals("NUMPAGES", StringComparison.OrdinalIgnoreCase)
                || value.Equals("SECTIONPAGES", StringComparison.OrdinalIgnoreCase));
}
