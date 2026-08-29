using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Dockit.Docx;

public static class Observation
{
    private const string RevisionSchema = "tiwater.docx-revision/v1";
    private const string ObjectReferencePrefix = "dox1_";
    private const int DefaultLimit = 100;
    private const int MaximumLimit = 1000;

    private static readonly IReadOnlySet<string> Kinds = new HashSet<string>(StringComparer.Ordinal)
    {
        "part", "paragraph", "table", "gridColumn", "row", "cell", "run", "text", "drawing"
    };

    public static DocxObservationListResult List(
        string input,
        string kind,
        string? scope,
        string? parentReference,
        int limit,
        string? continuation)
    {
        var snapshot = Snapshot.Open(input);
        ValidateKind(kind);
        var objects = SelectObjects(snapshot, scope, parentReference)
            .Where(item => item.Kind == kind)
            .ToList();
        var selection = SelectionKey("list", kind, scope, parentReference, null);
        var page = Page(objects, snapshot.Revision, selection, limit, continuation);
        return new DocxObservationListResult(
            "tiwater.docx-observation-list/v1",
            Receipt("list", snapshot.Revision, page.TotalCount, page.Items.Count, page.Remaining, page.Continuation),
            page.Items.Select(ToObject).ToList());
    }

    public static DocxObservationFindResult Find(
        string input,
        string literal,
        string? kind,
        string? scope,
        string? parentReference,
        int limit,
        string? continuation)
    {
        if (string.IsNullOrEmpty(literal))
            throw new InvalidOperationException("find-literal-must-not-be-empty");

        var snapshot = Snapshot.Open(input);
        if (kind is not null) ValidateKind(kind);
        var matches = SelectObjects(snapshot, scope, parentReference)
            .Where(item => kind is null || item.Kind == kind)
            .Select(item => new { Item = item, Ranges = FindRanges(TechnicalText(item.Element), literal) })
            .Where(item => item.Ranges.Count > 0)
            .Select(item => new DocxObservationMatch(ToObject(item.Item), item.Ranges))
            .ToList();
        var selection = SelectionKey("find", kind, scope, parentReference, literal);
        var page = Page(matches, snapshot.Revision, selection, limit, continuation);
        return new DocxObservationFindResult(
            "tiwater.docx-observation-find/v1",
            Receipt("find", snapshot.Revision, page.TotalCount, page.Items.Count, page.Remaining, page.Continuation),
            page.Items);
    }

    public static DocxObservationReadResult Read(string input, string reference, string? expectedRevision)
    {
        var snapshot = Snapshot.Open(input);
        if (!IsObjectReference(reference))
            throw new InvalidOperationException("object-ref-invalid");
        if (expectedRevision is not null && !StringComparer.Ordinal.Equals(expectedRevision, snapshot.Revision.Id))
            throw new InvalidOperationException("stale-revision");

        var selected = snapshot.Objects.FirstOrDefault(item =>
            StringComparer.Ordinal.Equals(item.Reference, reference));
        if (selected is null)
            throw new InvalidOperationException("stale-object-ref");

        var detail = new DocxObservationDetail(
            ToObject(selected),
            selected.Element.NamespaceUri,
            selected.Element.OuterXml,
            selected.Element.ChildElements.Count,
            selected.Element.GetAttributes()
                .Select(attribute => new DocxOpenXmlAttribute(attribute.LocalName, attribute.NamespaceUri, attribute.Value ?? string.Empty))
                .ToList());
        return new DocxObservationReadResult(
            "tiwater.docx-observation-read/v1",
            Receipt("read", snapshot.Revision, 1, 1, 0, null),
            detail);
    }

    internal static DocxRevision CurrentRevision(string input)
        => Snapshot.Open(input).Revision;

    internal static string MakeReference(DocxRevision revision, string kind, string storyPart, string nativePath)
        => MakeObjectReference(revision, kind, storyPart, nativePath);

    internal static string NativePathFor(OpenXmlElement element) => Snapshot.NativePath(element);

    internal static IReadOnlyList<ResolvedDocxReference> ResolveReferences(
        string input,
        string expectedRevision,
        IReadOnlyList<string> references)
    {
        var snapshot = Snapshot.Open(input);
        if (!StringComparer.Ordinal.Equals(expectedRevision, snapshot.Revision.Id))
            throw new InvalidOperationException("stale-revision");
        return references.Select(reference =>
        {
            if (!IsObjectReference(reference)) throw new InvalidOperationException("object-ref-invalid");
            var selected = snapshot.Objects.FirstOrDefault(item => StringComparer.Ordinal.Equals(item.Reference, reference))
                ?? throw new InvalidOperationException("stale-object-ref");
            return new ResolvedDocxReference(selected.Reference, selected.Kind, selected.StoryPart, selected.NativePath);
        }).ToList();
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
    public static int MaximumPageLimit => MaximumLimit;

    private static DocxObservationReceipt Receipt(
        string operation,
        DocxRevision revision,
        int totalCount,
        int returnedCount,
        int remaining,
        string? continuation)
        => new(
            "tiwater.docx-observation-receipt/v1",
            operation,
            revision,
            totalCount,
            returnedCount,
            remaining,
            continuation);

    private static PageResult<T> Page<T>(
        IReadOnlyList<T> items,
        DocxRevision revision,
        string selection,
        int limit,
        string? continuation)
    {
        if (limit is < 1 or > MaximumLimit)
            throw new InvalidOperationException($"limit-must-be-between-1-and-{MaximumLimit}");

        var offset = 0;
        if (continuation is not null)
        {
            var cursor = DecodeCursor(continuation);
            if (!StringComparer.Ordinal.Equals(cursor.Revision, revision.Id)
                || !StringComparer.Ordinal.Equals(cursor.Selection, selection))
                throw new InvalidOperationException("continuation-does-not-match-current-query");
            offset = cursor.Offset;
        }

        if (offset < 0 || offset > items.Count)
            throw new InvalidOperationException("continuation-offset-invalid");

        var pageItems = items.Skip(offset).Take(limit).ToList();
        var remaining = items.Count - offset - pageItems.Count;
        var next = remaining == 0
            ? null
            : EncodeCursor(new Cursor(revision.Id, selection, offset + pageItems.Count));
        return new PageResult<T>(items.Count, pageItems, remaining, next);
    }

    private static string SelectionKey(
        string operation,
        string? kind,
        string? scope,
        string? parentReference,
        string? literal)
    {
        var value = string.Join(
            "\0",
            operation,
            kind ?? string.Empty,
            scope ?? string.Empty,
            parentReference ?? string.Empty,
            literal ?? string.Empty);
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    }

    private static string EncodeCursor(Cursor cursor)
    {
        var json = JsonSerializer.SerializeToUtf8Bytes(cursor, Json.Options);
        return Convert.ToBase64String(json).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    private static Cursor DecodeCursor(string value)
    {
        try
        {
            var padded = value.Replace('-', '+').Replace('_', '/');
            padded = padded.PadRight(padded.Length + (4 - padded.Length % 4) % 4, '=');
            var cursor = JsonSerializer.Deserialize<Cursor>(Convert.FromBase64String(padded), Json.Options);
            return cursor ?? throw new InvalidOperationException("continuation-invalid");
        }
        catch (FormatException)
        {
            throw new InvalidOperationException("continuation-invalid");
        }
        catch (JsonException)
        {
            throw new InvalidOperationException("continuation-invalid");
        }
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
        string? parentReference)
    {
        if (parentReference is null)
            return snapshot.Objects.Where(item => InScope(item, scope));
        if (!IsObjectReference(parentReference))
            throw new InvalidOperationException("parent-ref-invalid");

        var parent = snapshot.Objects.FirstOrDefault(item =>
            StringComparer.Ordinal.Equals(item.Reference, parentReference))
            ?? throw new InvalidOperationException("stale-parent-ref");
        if (scope is not null && !StringComparer.Ordinal.Equals(parent.StoryPart, scope))
            throw new InvalidOperationException("parent-ref-outside-scope");

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

    private static DocxObservationObject ToObject(NativeObject item)
    {
        var text = item.Kind == "part" ? null : TechnicalText(item.Element);
        return new DocxObservationObject(
            item.Reference,
            item.Kind,
            item.StoryPart,
            item.NativePath,
            item.Element.LocalName,
            text is null ? null : Clip(text, 160),
            text?.Length,
            item.Element.ChildElements.Count);
    }

    private static bool IsObjectReference(string value)
    {
        if (!value.StartsWith(ObjectReferencePrefix, StringComparison.Ordinal)
            || value.Length != ObjectReferencePrefix.Length + 64)
            return false;
        return value[ObjectReferencePrefix.Length..].All(character => Uri.IsHexDigit(character));
    }

    private static string MakeObjectReference(DocxRevision revision, string kind, string part, string path)
    {
        var value = string.Join("\0", "tiwater.docx-object/v1", revision.Id, kind, part, path);
        return ObjectReferencePrefix + Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
    }

    private static string Clip(string value, int maximum)
        => value.Length <= maximum ? value : value[..maximum] + "...";

    private static string TechnicalText(OpenXmlElement element)
        => element is Paragraph paragraph ? Inspector.GetParagraphText(paragraph) : element.InnerText;

    private sealed record Cursor(string Revision, string Selection, int Offset);
    private sealed record PageResult<T>(int TotalCount, IReadOnlyList<T> Items, int Remaining, string? Continuation);

    private sealed record NativeObject(
        string Reference,
        string Kind,
        string StoryPart,
        string NativePath,
        OpenXmlElement Element);

    private sealed record Story(string Part, OpenXmlPartRootElement Root);

    private sealed class Snapshot
    {
        private Snapshot(DocxRevision revision, IReadOnlyList<NativeObject> objects)
        {
            Revision = revision;
            Objects = objects;
            PublishedElements = objects
                .Select(item => item.Element)
                .ToHashSet<OpenXmlElement>(ReferenceEqualityComparer.Instance);
        }

        public DocxRevision Revision { get; }
        public IReadOnlyList<NativeObject> Objects { get; }
        public IReadOnlySet<OpenXmlElement> PublishedElements { get; }

        public static Snapshot Open(string input)
        {
            var path = Path.GetFullPath(input);
            if (!File.Exists(path)) throw new FileNotFoundException("input-docx-not-found", path);
            var inputSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
            var toolVersion = RuntimeIdentity.Version;
            var revisionDigest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
                string.Join("\0", inputSha256, "tiwater.docx.cli", toolVersion))));
            var revision = new DocxRevision(
                "docx-rev-v1-" + revisionDigest.ToLowerInvariant(),
                inputSha256,
                "tiwater.docx.cli",
                toolVersion);

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
                        MakeObjectReference(revision, kind, story.Part, nativePath),
                        kind,
                        story.Part,
                        nativePath,
                        element));
                }
            }
            return new Snapshot(revision, objects);
        }

        private static IEnumerable<Story> Stories(WordprocessingDocument document)
        {
            var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
            if (main.Document is not null) yield return new Story(PartUri(main.Uri), main.Document);

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

public sealed record DocxRevision(
    [property: JsonPropertyName("id")] string Id,
    [property: JsonPropertyName("inputSha256")] string InputSha256,
    [property: JsonPropertyName("provider")] string Provider,
    [property: JsonPropertyName("toolVersion")] string ToolVersion);

public sealed record DocxObservationReceipt(
    [property: JsonPropertyName("schema")] string Schema,
    [property: JsonPropertyName("operation")] string Operation,
    [property: JsonPropertyName("revision")] DocxRevision Revision,
    [property: JsonPropertyName("totalCount")] int TotalCount,
    [property: JsonPropertyName("returnedCount")] int ReturnedCount,
    [property: JsonPropertyName("remaining")] int Remaining,
    [property: JsonPropertyName("continuation")] string? Continuation);

public sealed record DocxObservationObject(
    [property: JsonPropertyName("ref")] string Reference,
    [property: JsonPropertyName("kind")] string Kind,
    [property: JsonPropertyName("storyPart")] string StoryPart,
    [property: JsonPropertyName("nativePath")] string NativePath,
    [property: JsonPropertyName("localName")] string LocalName,
    [property: JsonPropertyName("textPreview")] string? TextPreview,
    [property: JsonPropertyName("textLength")] int? TextLength,
    [property: JsonPropertyName("childCount")] int ChildCount);

public sealed record DocxTextMatch(
    [property: JsonPropertyName("offset")] int Offset,
    [property: JsonPropertyName("length")] int Length);

public sealed record DocxObservationMatch(
    [property: JsonPropertyName("object")] DocxObservationObject Object,
    [property: JsonPropertyName("matches")] IReadOnlyList<DocxTextMatch> Matches);

public sealed record DocxOpenXmlAttribute(
    [property: JsonPropertyName("localName")] string LocalName,
    [property: JsonPropertyName("namespaceUri")] string NamespaceUri,
    [property: JsonPropertyName("value")] string Value);

public sealed record DocxObservationDetail(
    [property: JsonPropertyName("object")] DocxObservationObject Object,
    [property: JsonPropertyName("namespaceUri")] string NamespaceUri,
    [property: JsonPropertyName("outerXml")] string OuterXml,
    [property: JsonPropertyName("childCount")] int ChildCount,
    [property: JsonPropertyName("attributes")] IReadOnlyList<DocxOpenXmlAttribute> Attributes);

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
    [property: JsonPropertyName("observation")] DocxObservationDetail Observation);

internal sealed record ResolvedDocxReference(
    string Reference,
    string Kind,
    string StoryPart,
    string NativePath);
