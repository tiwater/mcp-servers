using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

namespace Dockit.Docx;

public static class DocxPackageNormalizer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly IReadOnlyDictionary<string, string> PrefixToNamespace = new Dictionary<string, string>
    {
        ["w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ["m"] = "http://schemas.openxmlformats.org/officeDocument/2006/math",
        ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ["w14"] = "http://schemas.microsoft.com/office/word/2010/wordml",
        ["w15"] = "http://schemas.microsoft.com/office/word/2012/wordml",
        ["wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ["w16se"] = "http://schemas.microsoft.com/office/word/2015/wordml/symex",
        ["w16cid"] = "http://schemas.microsoft.com/office/word/2016/wordml/cid",
        ["w16"] = "http://schemas.microsoft.com/office/word/2018/wordml",
        ["w16cex"] = "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        ["w16sdtdh"] = "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
        ["w16sdtfl"] = "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
        ["w16du"] = "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    };

    private static readonly IReadOnlyDictionary<string, string> NamespaceToPrefix =
        PrefixToNamespace.GroupBy(item => item.Value).ToDictionary(group => group.Key, group => group.First().Key);

    private static readonly IReadOnlyDictionary<XName, int> RunPropertyOrder = new Dictionary<XName, int>
    {
        [W + "rStyle"] = 0,
        [W + "rFonts"] = 1,
        [W + "b"] = 2,
        [W + "bCs"] = 3,
        [W + "i"] = 4,
        [W + "iCs"] = 5,
        [W + "caps"] = 6,
        [W + "smallCaps"] = 7,
        [W + "strike"] = 8,
        [W + "dstrike"] = 9,
        [W + "outline"] = 10,
        [W + "shadow"] = 11,
        [W + "emboss"] = 12,
        [W + "imprint"] = 13,
        [W + "noProof"] = 14,
        [W + "snapToGrid"] = 15,
        [W + "vanish"] = 16,
        [W + "webHidden"] = 17,
        [W + "color"] = 20,
        [W + "spacing"] = 21,
        [W + "w"] = 22,
        [W + "kern"] = 23,
        [W + "position"] = 24,
        [W + "sz"] = 30,
        [W + "szCs"] = 31,
        [W + "highlight"] = 32,
        [W + "u"] = 33,
        [W + "effect"] = 34,
        [W + "bdr"] = 35,
        [W + "shd"] = 36,
        [W + "fitText"] = 37,
        [W + "vertAlign"] = 38,
        [W + "rtl"] = 39,
        [W + "lang"] = 40,
        [W + "eastAsianLayout"] = 41,
        [W + "specVanish"] = 42,
        [W + "oMath"] = 43,
    };

    private static readonly IReadOnlyDictionary<XName, int> TableCellPropertyOrder = new Dictionary<XName, int>
    {
        [W + "cnfStyle"] = 0,
        [W + "tcW"] = 1,
        [W + "gridSpan"] = 2,
        [W + "hMerge"] = 3,
        [W + "vMerge"] = 4,
        [W + "tcBorders"] = 5,
        [W + "shd"] = 6,
        [W + "noWrap"] = 7,
        [W + "tcMar"] = 8,
        [W + "textDirection"] = 9,
        [W + "tcFitText"] = 10,
        [W + "vAlign"] = 11,
        [W + "hideMark"] = 12,
    };

    private static readonly IReadOnlyDictionary<XName, int> TablePropertyOrder = new Dictionary<XName, int>
    {
        [W + "tblStyle"] = 0,
        [W + "tblpPr"] = 1,
        [W + "tblOverlap"] = 2,
        [W + "bidiVisual"] = 3,
        [W + "tblStyleRowBandSize"] = 4,
        [W + "tblStyleColBandSize"] = 5,
        [W + "tblW"] = 6,
        [W + "jc"] = 7,
        [W + "tblCellSpacing"] = 8,
        [W + "tblInd"] = 9,
        [W + "tblBorders"] = 10,
        [W + "shd"] = 11,
        [W + "tblLayout"] = 12,
        [W + "tblCellMar"] = 13,
        [W + "tblLook"] = 14,
        [W + "tblCaption"] = 15,
        [W + "tblDescription"] = 16,
    };

    public static int RunNormalize(string[] args)
    {
        if (args.Length < 2)
        {
            throw new InvalidOperationException("normalize-openxml requires <input.docx> <output.docx>");
        }

        Normalize(args[0], args[1]);
        Console.WriteLine(JsonSerializer.Serialize(new { input = Path.GetFullPath(args[0]), output = Path.GetFullPath(args[1]) }, Json.Options));
        return 0;
    }

    public static void Normalize(string input, string output)
    {
        var inputPath = Path.GetFullPath(input);
        var outputPath = Path.GetFullPath(output);
        if (!string.Equals(inputPath, outputPath, StringComparison.Ordinal))
        {
            File.Copy(inputPath, outputPath, overwrite: true);
        }

        using var archive = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        foreach (var entry in archive.Entries.Where(entry => entry.FullName.StartsWith("word/", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).ToList())
        {
            var xml = ReadEntry(entry);
            if (TryNormalizeXml(xml, out var normalized) && !string.Equals(xml, normalized, StringComparison.Ordinal))
            {
                var name = entry.FullName;
                entry.Delete();
                var replacement = archive.CreateEntry(name, CompressionLevel.Optimal);
                using var stream = replacement.Open();
                using var writer = new StreamWriter(stream);
                writer.Write(normalized);
            }
        }
    }

    private static bool TryNormalizeXml(string xml, out string normalized)
    {
        normalized = xml;
        XDocument document;
        try
        {
            document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        }
        catch
        {
            return false;
        }

        if (document.Root is null)
        {
            return false;
        }

        NormalizeNamespaces(document.Root);
        NormalizeChildOrder(document.Root);
        normalized = document.Declaration is null
            ? document.ToString(SaveOptions.DisableFormatting)
            : document.Declaration + document.ToString(SaveOptions.DisableFormatting);
        return true;
    }

    private static void NormalizeNamespaces(XElement root)
    {
        var usedNamespaces = root.DescendantsAndSelf()
            .Select(element => element.Name.NamespaceName)
            .Concat(root.DescendantsAndSelf().Attributes().Where(attribute => !attribute.IsNamespaceDeclaration).Select(attribute => attribute.Name.NamespaceName))
            .Where(ns => ns.Length > 0)
            .ToHashSet(StringComparer.Ordinal);

        var ignorable = root.Attribute(Mc + "Ignorable");
        if (ignorable is not null)
        {
            foreach (var prefix in ignorable.Value.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                if (PrefixToNamespace.TryGetValue(prefix, out var ns))
                {
                    usedNamespaces.Add(ns);
                }
            }
        }

        foreach (var element in root.DescendantsAndSelf())
        {
            foreach (var attribute in element.Attributes().Where(ShouldRemoveNamespaceAlias).ToList())
            {
                attribute.Remove();
            }
        }

        foreach (var ns in usedNamespaces)
        {
            if (NamespaceToPrefix.TryGetValue(ns, out var prefix))
            {
                root.SetAttributeValue(XNamespace.Xmlns + prefix, ns);
            }
        }
    }

    private static bool ShouldRemoveNamespaceAlias(XAttribute attribute)
    {
        if (!attribute.IsNamespaceDeclaration)
        {
            return false;
        }

        var prefix = attribute.Name.LocalName;
        return prefix.StartsWith("ns", StringComparison.Ordinal)
            && prefix.Skip(2).All(char.IsDigit)
            && NamespaceToPrefix.ContainsKey(attribute.Value);
    }

    private static void NormalizeChildOrder(XElement root)
    {
        foreach (var element in root.DescendantsAndSelf())
        {
            if (element.Name == W + "rPr")
            {
                SortChildren(element, RunPropertyOrder);
            }
            else if (element.Name == W + "tcPr")
            {
                SortChildren(element, TableCellPropertyOrder);
            }
            else if (element.Name == W + "tblPr")
            {
                SortChildren(element, TablePropertyOrder);
            }
        }
    }

    private static void SortChildren(XElement element, IReadOnlyDictionary<XName, int> order)
    {
        var nodes = element.Nodes().ToList();
        var sortable = nodes.OfType<XElement>().Select((child, index) => new { Child = child, Index = index }).ToList();
        if (sortable.Count < 2)
        {
            return;
        }

        var sorted = sortable
            .OrderBy(item => order.TryGetValue(item.Child.Name, out var childOrder) ? childOrder : int.MaxValue)
            .ThenBy(item => item.Index)
            .Select(item => new XElement(item.Child))
            .Cast<XNode>()
            .ToList();
        element.ReplaceNodes(sorted);
    }

    private static string ReadEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
