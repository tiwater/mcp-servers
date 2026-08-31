using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class Observation
{
    private const int DefaultLimit = 100;

    private static readonly IReadOnlySet<string> Kinds = new HashSet<string>(StringComparer.Ordinal)
    {
        "part", "paragraph", "table", "gridColumn", "row", "cell", "run", "text", "drawing"
    };
    private static readonly IReadOnlySet<string> ListKinds = new HashSet<string>(Kinds.Where(kind => kind != "part"), StringComparer.Ordinal);
    private static readonly IReadOnlySet<string> ReadKinds = new HashSet<string>(Kinds.Where(kind => kind != "part"), StringComparer.Ordinal);

    public static DocxObservationListResult List(
        string input,
        IReadOnlySet<string> kinds,
        string? scope,
        DocxObjectAddress? parent,
        int limit,
        int offset)
    {
        var snapshot = Snapshot.Open(input);
        if (kinds.Count == 0) throw new InvalidOperationException("kinds-is-required");
        foreach (var kind in kinds)
            if (!ListKinds.Contains(kind)) throw new InvalidOperationException($"unsupported-list-kind: {kind}");
        var objects = SelectObjects(snapshot, scope, parent)
            .Where(item => kinds.Contains(item.Kind))
            .ToList();
        var page = Page(objects, limit, offset);
        return new DocxObservationListResult(
            "tiwater.docx-observation-list/v1",
            Receipt("list", page.TotalCount, page.Items.Count, page.Remaining, page.NextOffset),
            page.Items.Select(item => ToObject(snapshot, item)).ToList());
    }

    public static DocxObservationFindResult Find(
        string input,
        string literal,
        string? kind,
        string? scope,
        DocxObjectAddress? parent,
        int limit,
        int offset)
    {
        if (string.IsNullOrEmpty(literal))
            throw new InvalidOperationException("find-literal-must-not-be-empty");

        var snapshot = Snapshot.Open(input);
        if (kind is not null) ValidateKind(kind);
        var matches = SelectObjects(snapshot, scope, parent)
            .Where(item => kind is null || item.Kind == kind)
            .Select(item => new { Item = item, Ranges = FindRanges(TechnicalText(item.Element), literal) })
            .Where(item => item.Ranges.Count > 0)
            .Select(item => new DocxObservationMatch(ToObject(snapshot, item.Item), item.Ranges))
            .ToList();
        var page = Page(matches, limit, offset);
        return new DocxObservationFindResult(
            "tiwater.docx-observation-find/v1",
            Receipt("find", page.TotalCount, page.Items.Count, page.Remaining, page.NextOffset),
            page.Items);
    }

    public static DocxObservationReadResult Read(
        string input,
        IReadOnlyList<DocxObjectAddress> addresses,
        IReadOnlySet<string> kinds)
    {
        var snapshot = Snapshot.Open(input);
        if (addresses.Count == 0) throw new InvalidOperationException("addresses-is-required");
        if (kinds.Count == 0) throw new InvalidOperationException("kinds-is-required");
        foreach (var kind in kinds)
            if (!ReadKinds.Contains(kind)) throw new InvalidOperationException($"unsupported-read-kind: {kind}");
        var selectedObjects = ResolveAddresses(snapshot, addresses, "addresses")
            .Select(item => new NativeObject(item.Address, item.Kind, item.Element))
            .ToList();

        DocxObservationDetail Detail(NativeObject selected)
        {
            var included = new Dictionary<OpenXmlElement, NativeObject>(ReferenceEqualityComparer.Instance);
            foreach (var item in snapshot.Objects.Where(item =>
                !ReferenceEquals(item.Element, selected.Element)
                && kinds.Contains(item.Kind)
                && item.Element.Ancestors().Any(ancestor => ReferenceEquals(ancestor, selected.Element))))
                included.Add(item.Element, item);
            DocxObservationNode Node(NativeObject item)
            {
                var children = included.Values.Where(candidate =>
                {
                    var parent = candidate.Element.Parent;
                    while (parent is not null && !ReferenceEquals(parent, item.Element))
                    {
                        if (included.ContainsKey(parent)) return false;
                        parent = parent.Parent;
                    }
                    return ReferenceEquals(parent, item.Element);
                }).Select(Node).ToList();
                return new DocxObservationNode(ToObject(snapshot, item), children);
            }
            return new DocxObservationDetail(
                ToObject(snapshot, selected),
                included.Values.Where(item =>
                {
                    var parent = item.Element.Parent;
                    while (parent is not null && !ReferenceEquals(parent, selected.Element))
                    {
                        if (included.ContainsKey(parent)) return false;
                        parent = parent.Parent;
                    }
                    return ReferenceEquals(parent, selected.Element);
                }).Select(Node).ToList());
        }

        var details = selectedObjects.Select(Detail).ToList();
        return new DocxObservationReadResult(
            "tiwater.docx-observation-read/v1",
            Receipt("read", details.Count, details.Count, 0, null),
            details);
    }

    public static DocxTableIndexResult TableIndex(string input)
    {
        var snapshot = Snapshot.Open(input);
        var tables = snapshot.Objects.Where(item => item.Kind == "table").Select(item =>
        {
            var identity = ToObject(snapshot, item);
            var context = TableContext(snapshot, item);
            var table = (Table)item.Element;
            var rows = table.Elements<TableRow>().ToList();
            var columnCount = Math.Max(
                table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0,
                rows.Select(TableRowWidth).DefaultIfEmpty(0).Max());
            return new DocxTableIndexEntry(
                identity.Address,
                identity.ParentAddress,
                rows.Count,
                columnCount,
                identity.TextPreview ?? string.Empty,
                identity.TextLength ?? 0,
                context.PrecedingParagraph,
                context.FollowingParagraph);
        }).ToList();
        return new DocxTableIndexResult("tiwater.docx-table-index/v1", tables);
    }

    public static DocxTableReadResult ReadTable(string input, DocxObjectAddress address)
    {
        var snapshot = Snapshot.Open(input);
        var selected = ResolveAddresses(snapshot, new[] { address }, "table").Single();
        if (selected.Kind != "table" || selected.Element is not Table table)
            throw new InvalidOperationException("table-reference-kind-invalid");
        var gridColumns = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>()
            .Select(column => new DocxTableReadGridColumn(
                Address(selected.StoryPart, NativePathFor(column)),
                int.TryParse(column.Width?.Value, out var width) ? width : null))
            .ToArray() ?? [];
        var nativeRows = table.Elements<TableRow>().ToArray();
        var projectedRows = nativeRows.Select(row =>
        {
            var cursor = RowOffset(row.TableRowProperties, "gridBefore");
            var cells = row.Elements<TableCell>().Select(cell =>
            {
                var span = Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                var projection = new ProjectedTableCell(cell, cursor, span, VerticalMergeValue(cell.TableCellProperties));
                cursor += span;
                return projection;
            }).ToArray();
            return cells;
        }).ToArray();
        ProjectedTableCell LogicalOwner(int rowIndex, ProjectedTableCell cell)
        {
            if (cell.VerticalMerge is null || cell.VerticalMerge == "restart") return cell;
            for (var previous = rowIndex - 1; previous >= 0; previous--)
            {
                var candidate = projectedRows[previous].SingleOrDefault(item =>
                    item.GridColumnStart == cell.GridColumnStart && item.GridSpan == cell.GridSpan);
                if (candidate is null || candidate.VerticalMerge is null) break;
                if (candidate.VerticalMerge == "restart") return candidate;
            }
            throw new InvalidOperationException("vertical-merge-owner-not-found");
        }
        var rows = nativeRows.Select((row, rowIndex) => new DocxTableReadRow(
            Address(selected.StoryPart, NativePathFor(row)),
            row.TableRowProperties?.GetFirstChild<TableHeader>() is not null,
            row.TableRowProperties?.GetFirstChild<CantSplit>() is not null,
            RowOffset(row.TableRowProperties, "gridBefore"),
            RowOffset(row.TableRowProperties, "gridAfter"),
            projectedRows[rowIndex].Select(cell =>
            {
                var logicalOwner = LogicalOwner(rowIndex, cell);
                return new DocxTableReadCell(
                    Address(selected.StoryPart, NativePathFor(cell.Cell)),
                    cell.GridColumnStart,
                    cell.GridSpan,
                    cell.VerticalMerge,
                    cell.VerticalMerge is null
                        ? null
                        : Address(selected.StoryPart, NativePathFor(logicalOwner.Cell)),
                    string.Join("\n", logicalOwner.Cell.Elements<Paragraph>()
                        .Select(paragraph => paragraph.InnerText)),
                    cell.Cell.Elements<Paragraph>().Select(paragraph => new DocxTableReadParagraph(
                    Address(selected.StoryPart, NativePathFor(paragraph)),
                    paragraph.InnerText,
                    paragraph.Descendants<Text>().Select(value => new DocxTableReadText(
                        Address(selected.StoryPart, NativePathFor(value)), value.Text)).ToArray()
                    )).ToArray());
            }).ToArray()
        )).ToArray();
        var columnCount = Math.Max(
            table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count() ?? 0,
            rows.Select(row => row.GridBefore + row.Cells.Sum(cell => cell.GridSpan) + row.GridAfter)
                .DefaultIfEmpty(0).Max());
        var context = TableContext(snapshot,
            snapshot.Objects.Single(item => item.Address == selected.Address));
        return new DocxTableReadResult("tiwater.docx-table-read/v2", address, rows.Length,
            columnCount, gridColumns, context.PrecedingParagraph, context.FollowingParagraph, rows);
    }

    private static DocxTableContext TableContext(Snapshot snapshot, NativeObject table)
    {
        var parentAddress = PublishedParentAddress(snapshot, table);
        if (parentAddress is null) return new DocxTableContext(null, null);
        var parent = snapshot.Objects.First(item => item.Address == parentAddress);
        var siblings = snapshot.Objects
            .Where(item => IsDirectPublishedChild(snapshot, item, parent))
            .ToList();
        var tableIndex = siblings.FindIndex(item => ReferenceEquals(item.Element, table.Element));
        if (tableIndex < 0) return new DocxTableContext(null, null);

        DocxTableContextParagraph? ParagraphAt(IEnumerable<NativeObject> candidates)
        {
            var paragraph = candidates.FirstOrDefault(item =>
                item.Kind == "paragraph" && !string.IsNullOrWhiteSpace(TechnicalText(item.Element)));
            if (paragraph is null) return null;
            var text = TechnicalText(paragraph.Element);
            return new DocxTableContextParagraph(paragraph.Address, Clip(text, 160), text.Length);
        }

        return new DocxTableContext(
            ParagraphAt(siblings.Take(tableIndex).Reverse()),
            ParagraphAt(siblings.Skip(tableIndex + 1)));
    }

    private static string? VerticalMergeValue(TableCellProperties? properties)
        => properties?.VerticalMerge is null
            ? null
            : properties.VerticalMerge.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "val").Value ?? "continue";

    private static int TableRowWidth(TableRow row)
        => RowOffset(row.TableRowProperties, "gridBefore")
            + row.Elements<TableCell>().Sum(cell => Math.Max(1, cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1))
            + RowOffset(row.TableRowProperties, "gridAfter");

    private static int RowOffset(OpenXmlElement? properties, string localName)
    {
        var value = properties?.ChildElements.FirstOrDefault(child => child.LocalName == localName)
            ?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return int.TryParse(value, out var result) ? result : 0;
    }

    internal static DocxObjectAddress Address(string storyPart, string nativePath)
        => new(storyPart, nativePath);

    internal static string NativePathFor(OpenXmlElement element) => Snapshot.NativePath(element);

    internal static IReadOnlyList<ResolvedDocxAddress> ResolveAddresses(
        string input,
        IReadOnlyList<DocxObjectAddress> addresses,
        string name = "addresses")
    {
        var snapshot = Snapshot.Open(input);
        return ResolveAddresses(snapshot, addresses, name);
    }

    private static IReadOnlyList<ResolvedDocxAddress> ResolveAddresses(
        Snapshot snapshot,
        IReadOnlyList<DocxObjectAddress> addresses,
        string name)
        => addresses.Select((address, index) =>
        {
            ValidateAddress(address, $"{name}[{index}]");
            var selected = snapshot.Objects.FirstOrDefault(item => item.Address == address)
                ?? throw new InvalidOperationException($"object-address-not-found: {name}[{index}]");
            return new ResolvedDocxAddress(selected.Address, selected.Kind, selected.Element);
        }).ToList();

    private static void ValidateAddress(DocxObjectAddress address, string name)
    {
        if (string.IsNullOrWhiteSpace(address.Part)) throw new InvalidOperationException($"{name}.part-is-required");
        if (string.IsNullOrWhiteSpace(address.Path) || !address.Path.StartsWith("/", StringComparison.Ordinal))
            throw new InvalidOperationException($"{name}.path-is-invalid");
    }

    internal static OpenXmlElement ResolveNativePath(
        WordprocessingDocument document,
        string storyPart,
        string nativePath)
    {
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
        var stories = new List<(string Part, OpenXmlPartRootElement Root)>();
        if (main.Document is not null) stories.Add((PartUri(main.Uri), main.Document));
        stories.AddRange(main.HeaderParts.Where(part => part.Header is not null).Select(part => (PartUri(part.Uri), (OpenXmlPartRootElement)part.Header!)));
        stories.AddRange(main.FooterParts.Where(part => part.Footer is not null).Select(part => (PartUri(part.Uri), (OpenXmlPartRootElement)part.Footer!)));
        if (main.FootnotesPart?.Footnotes is not null) stories.Add((PartUri(main.FootnotesPart.Uri), main.FootnotesPart.Footnotes));
        if (main.EndnotesPart?.Endnotes is not null) stories.Add((PartUri(main.EndnotesPart.Uri), main.EndnotesPart.Endnotes));
        if (main.WordprocessingCommentsPart?.Comments is not null) stories.Add((PartUri(main.WordprocessingCommentsPart.Uri), main.WordprocessingCommentsPart.Comments));
        var story = stories.SingleOrDefault(item => StringComparer.Ordinal.Equals(item.Part, storyPart));
        if (story.Root is null) throw new InvalidOperationException("object-story-part-not-found");

        var segments = nativePath.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0) throw new InvalidOperationException("object-native-path-invalid");
        OpenXmlElement current = story.Root;
        if (!SegmentMatches(current, segments[0], 1)) throw new InvalidOperationException("object-native-path-invalid");
        foreach (var segment in segments.Skip(1))
        {
            if (!TryParseSegment(segment, out var prefix, out var localName, out var siblingIndex))
                throw new InvalidOperationException("object-native-path-invalid");
            var namespaceUri = NamespaceForPrefix(prefix);
            var candidates = current.ChildElements
                .Where(item => item.NamespaceUri == namespaceUri && item.LocalName == localName)
                .ToList();
            if (siblingIndex < 1 || siblingIndex > candidates.Count)
                throw new InvalidOperationException("object-native-path-not-found");
            current = candidates[siblingIndex - 1];
        }
        return current;
    }

    private static string PartUri(Uri uri)
        => uri.OriginalString.StartsWith("/", StringComparison.Ordinal) ? uri.OriginalString : "/" + uri.OriginalString;

    private static bool SegmentMatches(OpenXmlElement element, string segment, int expectedIndex)
        => TryParseSegment(segment, out var prefix, out var localName, out var index)
            && index == expectedIndex
            && localName == element.LocalName
            && NamespaceForPrefix(prefix) == element.NamespaceUri;

    private static bool TryParseSegment(string segment, out string prefix, out string localName, out int index)
    {
        prefix = localName = string.Empty;
        index = 0;
        var colon = segment.IndexOf(':');
        var bracket = segment.LastIndexOf('[');
        if (colon <= 0 || bracket <= colon + 1 || !segment.EndsWith(']')) return false;
        prefix = segment[..colon];
        localName = segment[(colon + 1)..bracket];
        return int.TryParse(segment[(bracket + 1)..^1], out index) && index > 0;
    }

    private static string NamespaceForPrefix(string prefix)
        => prefix switch
        {
            "w" => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "a" => "http://schemas.openxmlformats.org/drawingml/2006/main",
            "wp" => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "pic" => "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            _ => throw new InvalidOperationException($"object-native-path-prefix-unsupported: {prefix}"),
        };

    public static int DefaultPageLimit => DefaultLimit;

    private static DocxObservationReceipt Receipt(
        string operation,
        int totalCount,
        int returnedCount,
        int remaining,
        int? nextOffset)
        => new(
            "tiwater.docx-observation-receipt/v1",
            operation,
            totalCount,
            returnedCount,
            remaining,
            nextOffset);

    private static PageResult<T> Page<T>(
        IReadOnlyList<T> items,
        int limit,
        int offset)
    {
        if (limit < 1)
            throw new InvalidOperationException("limit-must-be-positive");
        if (offset < 0 || offset > items.Count)
            throw new InvalidOperationException("offset-invalid");

        var pageItems = items.Skip(offset).Take(limit).ToList();
        var remaining = items.Count - offset - pageItems.Count;
        int? next = remaining == 0 ? null : offset + pageItems.Count;
        return new PageResult<T>(items.Count, pageItems, remaining, next);
    }

    private static void ValidateKind(string kind)
    {
        if (!Kinds.Contains(kind))
            throw new InvalidOperationException($"object-kind-unsupported: {kind}");
    }

    private static bool InScope(NativeObject item, string? scope)
        => scope is null || StringComparer.Ordinal.Equals(item.StoryPart, scope);

    private static IEnumerable<NativeObject> SelectObjects(
        Snapshot snapshot,
        string? scope,
        DocxObjectAddress? parentAddress)
    {
        if (parentAddress is null)
        {
            var roots = snapshot.Objects
                .Where(item => item.Kind == "part" && InScope(item, scope))
                .ToList();
            return snapshot.Objects.Where(item =>
                item.Kind != "part"
                && InScope(item, scope)
                && roots.Any(root => IsDirectPublishedChild(snapshot, item, root)));
        }
        ValidateAddress(parentAddress, "parent");
        var parent = snapshot.Objects.FirstOrDefault(item => item.Address == parentAddress)
            ?? throw new InvalidOperationException("parent-address-not-found");
        if (scope is not null && !StringComparer.Ordinal.Equals(parent.StoryPart, scope))
            throw new InvalidOperationException("parent-address-outside-scope");

        return snapshot.Objects.Where(item =>
            InScope(item, scope) && IsDirectPublishedChild(snapshot, item, parent));
    }

    private static bool IsDirectPublishedChild(
        Snapshot snapshot,
        NativeObject item,
        NativeObject parent)
    {
        if (ReferenceEquals(item.Element, parent.Element)) return false;
        var current = item.Element.Parent;
        while (current is not null)
        {
            if (ReferenceEquals(current, parent.Element)) return true;
            if (snapshot.PublishedElements.Contains(current)) return false;
            current = current.Parent;
        }
        return false;
    }

    private static IReadOnlyList<DocxTextMatch> FindRanges(string text, string literal)
    {
        var result = new List<DocxTextMatch>();
        var offset = 0;
        while (offset <= text.Length - literal.Length)
        {
            var match = text.IndexOf(literal, offset, StringComparison.Ordinal);
            if (match < 0) break;
            result.Add(new DocxTextMatch(match, literal.Length));
            offset = match + literal.Length;
        }
        return result;
    }

    private static DocxObservationObject ToObject(Snapshot snapshot, NativeObject item)
    {
        var text = item.Kind == "part" ? null : TechnicalText(item.Element);
        var cellProperties = (item.Element as TableCell)?.TableCellProperties;
        var verticalMerge = cellProperties?.VerticalMerge is null
            ? null
            : cellProperties.VerticalMerge.GetAttributes()
                .FirstOrDefault(attribute => attribute.LocalName == "val").Value ?? "continue";
        return new DocxObservationObject(
            item.Address,
            PublishedParentAddress(snapshot, item),
            item.Kind,
            item.Element.LocalName,
            text is null ? null : Clip(text, 160),
            text?.Length,
            item.Element.ChildElements.Count,
            cellProperties is null ? null : Math.Max(1, cellProperties.GridSpan?.Val?.Value ?? 1),
            verticalMerge);
    }

    private static DocxObjectAddress? PublishedParentAddress(Snapshot snapshot, NativeObject item)
    {
        var current = item.Element.Parent;
        while (current is not null)
        {
            if (snapshot.ObjectsByElement.TryGetValue(current, out var parent)) return parent.Address;
            current = current.Parent;
        }
        return null;
    }

    private static string Clip(string value, int maximum)
        => value.Length <= maximum ? value : value[..maximum] + "...";

    private static string TechnicalText(OpenXmlElement element)
        => element is Paragraph paragraph ? Inspector.GetParagraphText(paragraph) : element.InnerText;

    private sealed record PageResult<T>(int TotalCount, IReadOnlyList<T> Items, int Remaining, int? NextOffset);

    private sealed record NativeObject(
        DocxObjectAddress Address,
        string Kind,
        OpenXmlElement Element)
    {
        public string StoryPart => Address.Part;
        public string NativePath => Address.Path;
    }

    private sealed record Story(string Part, OpenXmlElement Root);

    private sealed class Snapshot
    {
        private Snapshot(IReadOnlyList<NativeObject> objects)
        {
            Objects = objects;
            ObjectsByElement = objects.ToDictionary<NativeObject, OpenXmlElement, NativeObject>(
                item => item.Element,
                item => item,
                ReferenceEqualityComparer.Instance);
            PublishedElements = ObjectsByElement.Keys.ToHashSet<OpenXmlElement>(ReferenceEqualityComparer.Instance);
        }

        public IReadOnlyList<NativeObject> Objects { get; }
        public IReadOnlyDictionary<OpenXmlElement, NativeObject> ObjectsByElement { get; }
        public IReadOnlySet<OpenXmlElement> PublishedElements { get; }

        public static Snapshot Open(string input)
        {
            var path = Path.GetFullPath(input);
            if (!File.Exists(path)) throw new FileNotFoundException("input-docx-not-found", path);

            using var document = WordprocessingDocument.Open(path, false);
            var objects = new List<NativeObject>();
            foreach (var story in Stories(document))
            {
                foreach (var element in Walk(story.Root))
                {
                    var kind = KindOf(element);
                    if (kind is null && !ReferenceEquals(element, story.Root)) continue;
                    kind ??= "part";
                    var nativePath = NativePath(element);
                    objects.Add(new NativeObject(
                        new DocxObjectAddress(story.Part, nativePath),
                        kind,
                        element));
                }
            }
            return new Snapshot(objects);
        }

        private static IEnumerable<Story> Stories(WordprocessingDocument document)
        {
            var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
            if (main.Document?.Body is not null) yield return new Story(PartUri(main.Uri), main.Document.Body);

            foreach (var part in main.HeaderParts.OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal))
                if (part.Header is not null) yield return new Story(PartUri(part.Uri), part.Header);
            foreach (var part in main.FooterParts.OrderBy(part => main.GetIdOfPart(part), StringComparer.Ordinal))
                if (part.Footer is not null) yield return new Story(PartUri(part.Uri), part.Footer);
            if (main.FootnotesPart?.Footnotes is not null)
                yield return new Story(PartUri(main.FootnotesPart.Uri), main.FootnotesPart.Footnotes);
            if (main.EndnotesPart?.Endnotes is not null)
                yield return new Story(PartUri(main.EndnotesPart.Uri), main.EndnotesPart.Endnotes);
            if (main.WordprocessingCommentsPart?.Comments is not null)
                yield return new Story(PartUri(main.WordprocessingCommentsPart.Uri), main.WordprocessingCommentsPart.Comments);
        }

        private static string PartUri(Uri uri)
        {
            var value = uri.OriginalString;
            return value.StartsWith("/", StringComparison.Ordinal) ? value : "/" + value;
        }

        private static IEnumerable<OpenXmlElement> Walk(OpenXmlElement element)
        {
            yield return element;
            foreach (var child in element.ChildElements)
                foreach (var descendant in Walk(child))
                    yield return descendant;
        }

        private static string? KindOf(OpenXmlElement element)
            => element switch
            {
                Paragraph => "paragraph",
                Table => "table",
                GridColumn => "gridColumn",
                TableRow => "row",
                TableCell => "cell",
                Run => "run",
                Text => "text",
                Drawing => "drawing",
                _ => null,
            };

        internal static string NativePath(OpenXmlElement element)
        {
            var segments = new Stack<string>();
            OpenXmlElement? current = element;
            while (current is not null)
            {
                var siblings = current.Parent?.ChildElements
                    .Where(candidate => candidate.NamespaceUri == current.NamespaceUri && candidate.LocalName == current.LocalName)
                    .ToList();
                var index = siblings is null ? 1 : siblings.IndexOf(current) + 1;
                segments.Push($"{Prefix(current.NamespaceUri)}:{current.LocalName}[{index}]");
                current = current.Parent;
            }
            return "/" + string.Join("/", segments);
        }

        private static string Prefix(string namespaceUri)
            => namespaceUri switch
            {
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main" => "w",
                "http://schemas.openxmlformats.org/drawingml/2006/main" => "a",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" => "wp",
                "http://schemas.openxmlformats.org/drawingml/2006/picture" => "pic",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships" => "r",
                _ => "ns",
            };
    }
}

public sealed record DocxObjectAddress(
    [property: JsonPropertyName("part")] string Part,
    [property: JsonPropertyName("path")] string Path);

public sealed record DocxObservationReceipt(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("operation")] string Operation,
    [property: JsonPropertyName("totalCount")] int TotalCount,
    [property: JsonPropertyName("returnedCount")] int ReturnedCount,
    [property: JsonPropertyName("remaining")] int Remaining,
    [property: JsonPropertyName("nextOffset")] int? NextOffset);

public sealed record DocxObservationObject(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("parentAddress")] DocxObjectAddress? ParentAddress,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("localName")] string LocalName,
    [property: JsonPropertyName("textPreview")] string? TextPreview,
    [property: JsonPropertyName("textLength")] int? TextLength,
    [property: JsonPropertyName("childCount")] int ChildCount,
    [property: JsonPropertyName("gridSpan")] int? GridSpan,
    [property: JsonPropertyName("verticalMerge")] string? VerticalMerge);

public sealed record DocxTextMatch(
    [property: JsonPropertyName("offset")] int Offset,
    [property: JsonPropertyName("length")] int Length);

public sealed record DocxObservationMatch(
    [property: JsonPropertyName("object")] DocxObservationObject Object,
    [property: JsonPropertyName("matches")] IReadOnlyList<DocxTextMatch> Matches);

public sealed record DocxObservationDetail(
    [property: JsonPropertyName("object")] DocxObservationObject Object,
    [property: JsonPropertyName("children")] IReadOnlyList<DocxObservationNode> Children);

public sealed record DocxObservationNode(
    [property: JsonPropertyName("object")] DocxObservationObject Object,
    [property: JsonPropertyName("children")] IReadOnlyList<DocxObservationNode> Children);

public sealed record DocxObservationListResult(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("receipt")] DocxObservationReceipt Receipt,
    [property: JsonPropertyName("objects")] IReadOnlyList<DocxObservationObject> Objects);

public sealed record DocxObservationFindResult(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("receipt")] DocxObservationReceipt Receipt,
    [property: JsonPropertyName("matches")] IReadOnlyList<DocxObservationMatch> Matches);

public sealed record DocxObservationReadResult(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("receipt")] DocxObservationReceipt Receipt,
    [property: JsonPropertyName("observations")] IReadOnlyList<DocxObservationDetail> Observations);

public sealed record DocxTableIndexEntry(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("parentAddress")] DocxObjectAddress? ParentAddress,
    [property: JsonPropertyName("rowCount")] int RowCount,
    [property: JsonPropertyName("columnCount")] int ColumnCount,
    [property: JsonPropertyName("textPreview")] string TextPreview,
    [property: JsonPropertyName("textLength")] int TextLength,
    [property: JsonPropertyName("precedingParagraph")] DocxTableContextParagraph? PrecedingParagraph,
    [property: JsonPropertyName("followingParagraph")] DocxTableContextParagraph? FollowingParagraph);

public sealed record DocxTableContextParagraph(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("textPreview")] string TextPreview,
    [property: JsonPropertyName("textLength")] int TextLength);

public sealed record DocxTableContext(
    DocxTableContextParagraph? PrecedingParagraph,
    DocxTableContextParagraph? FollowingParagraph);

public sealed record DocxTableIndexResult(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("tables")] IReadOnlyList<DocxTableIndexEntry> Tables);

public sealed record DocxTableReadText(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("text")] string Text);

public sealed record DocxTableReadParagraph(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("text")] string Text,
    [property: JsonPropertyName("textNodes")] IReadOnlyList<DocxTableReadText> TextNodes);

public sealed record DocxTableReadCell(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("gridColumnStart")] int GridColumnStart,
    [property: JsonPropertyName("gridSpan")] int GridSpan,
    [property: JsonPropertyName("verticalMerge")] string? VerticalMerge,
    [property: JsonPropertyName("verticalMergeOwner")] DocxObjectAddress? VerticalMergeOwner,
    [property: JsonPropertyName("logicalText")] string LogicalText,
    [property: JsonPropertyName("paragraphs")] IReadOnlyList<DocxTableReadParagraph> Paragraphs);

public sealed record DocxTableReadRow(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("repeatHeader")] bool RepeatHeader,
    [property: JsonPropertyName("cantSplit")] bool CantSplit,
    [property: JsonPropertyName("gridBefore")] int GridBefore,
    [property: JsonPropertyName("gridAfter")] int GridAfter,
    [property: JsonPropertyName("cells")] IReadOnlyList<DocxTableReadCell> Cells);

public sealed record DocxTableReadGridColumn(
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("widthTwips")] int? WidthTwips);

public sealed record DocxTableReadResult(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("address")] DocxObjectAddress Address,
    [property: JsonPropertyName("rowCount")] int RowCount,
    [property: JsonPropertyName("columnCount")] int ColumnCount,
    [property: JsonPropertyName("gridColumns")] IReadOnlyList<DocxTableReadGridColumn> GridColumns,
    [property: JsonPropertyName("precedingParagraph")] DocxTableContextParagraph? PrecedingParagraph,
    [property: JsonPropertyName("followingParagraph")] DocxTableContextParagraph? FollowingParagraph,
    [property: JsonPropertyName("rows")] IReadOnlyList<DocxTableReadRow> Rows);

internal sealed record ProjectedTableCell(
    TableCell Cell,
    int GridColumnStart,
    int GridSpan,
    string? VerticalMerge);

internal sealed record ResolvedDocxAddress(
    DocxObjectAddress Address,
    string Kind,
    OpenXmlElement Element)
{
    public string StoryPart => Address.Part;
    public string NativePath => Address.Path;
}
