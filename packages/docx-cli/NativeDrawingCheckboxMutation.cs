using System.Buffers.Binary;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace Dockit.Docx;

public static class NativeDrawingCheckboxMutation
{
    public const string Command = "docx_set_drawing_checkbox_state";

    public static int Run(string[] args)
    {
        if (args.Length != 1) throw new InvalidOperationException($"{Command} requires <request.json>");
        var request = JsonSerializer.Deserialize<SetDrawingCheckboxStateRequest>(File.ReadAllText(args[0]), Json.Options)
            ?? throw new InvalidOperationException("set-drawing-checkbox-state-request-invalid");
        var receipt = Apply(request);
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            tool = Command,
            receipt = NativeMutationSupport.Describe(request.ReceiptOutput),
            output = NativeMutationSupport.Describe(receipt.Output),
            summary = new { pass = true, operationCount = request.Changes.Count, appliedCount = receipt.Changes.Count },
        }, Json.CamelCaseOptions));
        return 0;
    }

    public static SetDrawingCheckboxStateReceipt Apply(SetDrawingCheckboxStateRequest request)
    {
        if (request.Changes.Count == 0) throw new InvalidOperationException("changes-must-not-be-empty");
        var duplicate = request.Changes.GroupBy(change => change.Drawing).FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null) throw new InvalidOperationException("drawing-address-duplicate");

        using var paths = NativeMutationSupport.Paths(request.Input, request.Output, request.ReceiptOutput);
        var resolved = Observation.ResolveAddresses(
            paths.Input,
            request.Changes.Select(change => change.Drawing).ToArray(),
            "changes.drawing");
        if (resolved.Any(item => item.Kind != "drawing"))
            throw new InvalidOperationException("target-must-be-drawing");

        IReadOnlyDictionary<string, int> baseline;
        using (var input = WordprocessingDocument.Open(paths.Input, false))
        {
            baseline = NativeMutationSupport.ValidationIssueCounts(input);
            for (var index = 0; index < resolved.Count; index++)
            {
                var drawing = Observation.ResolveNativePath(input, resolved[index].StoryPart, resolved[index].NativePath) as Drawing
                    ?? throw new InvalidOperationException("target-must-be-drawing");
                _ = Prepare(input, resolved[index].StoryPart, drawing, $"changes[{index}].drawing");
            }
        }

        var temporaryPath = paths.Output + ".tmp-" + Guid.NewGuid().ToString("N");
        try
        {
            Tiwater.Office.WritableFileCopy.Copy(paths.Input, temporaryPath);
            using (var output = WordprocessingDocument.Open(temporaryPath, true))
            {
                for (var index = 0; index < resolved.Count; index++)
                {
                    var drawing = Observation.ResolveNativePath(output, resolved[index].StoryPart, resolved[index].NativePath) as Drawing
                        ?? throw new InvalidOperationException("output-drawing-not-found");
                    var prepared = Prepare(output, resolved[index].StoryPart, drawing, $"changes[{index}].drawing");
                    var bytes = prepared.Bitmap.WithState(request.Changes[index].Checked);
                    var replacement = AddImagePart(prepared.OwnerPart, prepared.ImagePart.ContentType);
                    using (var stream = new MemoryStream(bytes, writable: false)) replacement.FeedData(stream);
                    prepared.Blip.Embed = prepared.OwnerPart.GetIdOfPart(replacement);
                }
                SaveStories(output, resolved.Select(item => item.StoryPart));
                NativeMutationSupport.RejectAddedValidationIssues(output, baseline);
            }
            NativeMutationSupport.Commit(temporaryPath, paths);

            IReadOnlyList<DrawingCheckboxStateReadback> readback;
            using (var output = WordprocessingDocument.Open(paths.Output, false))
            {
                readback = resolved.Select((item, index) =>
                {
                    var drawing = Observation.ResolveNativePath(output, item.StoryPart, item.NativePath) as Drawing
                        ?? throw new InvalidOperationException("readback-drawing-not-found");
                    var prepared = Prepare(output, item.StoryPart, drawing, $"changes[{index}].drawing");
                    var state = prepared.Bitmap.Checked;
                    if (state != request.Changes[index].Checked)
                        throw new InvalidOperationException("output-readback-checkbox-state-mismatch");
                    return new DrawingCheckboxStateReadback(item.Address, state, prepared.Bitmap.Width, prepared.Bitmap.Height);
                }).ToArray();
            }
            var receipt = new SetDrawingCheckboxStateReceipt(
                "tiwater.docx-set-drawing-checkbox-state-receipt/v1",
                "tiwater.docx.cli",
                RuntimeIdentity.Version,
                readback,
                paths.Output);
            File.WriteAllText(paths.Receipt, JsonSerializer.Serialize(receipt, Json.CamelCaseOptions));
            return receipt;
        }
        catch
        {
            NativeMutationSupport.CleanupFailure(temporaryPath, paths);
            throw;
        }
    }

    internal static bool? ObserveCheckboxState(
        WordprocessingDocument document,
        string storyPart,
        Drawing drawing)
    {
        try { return Prepare(document, storyPart, drawing, "drawing").Bitmap.Checked; }
        catch (InvalidOperationException) { return null; }
    }

    private static PreparedDrawing Prepare(
        WordprocessingDocument document,
        string storyPart,
        Drawing drawing,
        string name)
    {
        var blips = drawing.Descendants<A.Blip>().ToArray();
        if (blips.Length != 1 || string.IsNullOrWhiteSpace(blips[0].Embed?.Value))
            throw new InvalidOperationException($"drawing-must-contain-one-embedded-picture: {name}");
        var owner = StoryPart(document, storyPart);
        ImagePart image;
        try { image = owner.GetPartById(blips[0].Embed!.Value!) as ImagePart
            ?? throw new InvalidOperationException($"drawing-image-relationship-invalid: {name}"); }
        catch (ArgumentOutOfRangeException error)
        {
            throw new InvalidOperationException($"drawing-image-relationship-invalid: {name}", error);
        }
        using var stream = image.GetStream(FileMode.Open, FileAccess.Read);
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        var bitmap = CheckboxDib.Open(memory.ToArray(), name);
        return new PreparedDrawing(owner, image, blips[0], bitmap);
    }

    private static OpenXmlPartContainer StoryPart(WordprocessingDocument document, string storyPart)
    {
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
        static string Uri(OpenXmlPart part) => part.Uri.OriginalString.StartsWith('/') ? part.Uri.OriginalString : "/" + part.Uri.OriginalString;
        IEnumerable<OpenXmlPart> parts = new OpenXmlPart[] { main }
            .Concat(main.HeaderParts)
            .Concat(main.FooterParts)
            .Concat(main.FootnotesPart is null ? [] : [main.FootnotesPart])
            .Concat(main.EndnotesPart is null ? [] : [main.EndnotesPart])
            .Concat(main.WordprocessingCommentsPart is null ? [] : [main.WordprocessingCommentsPart]);
        return parts.SingleOrDefault(part => StringComparer.Ordinal.Equals(Uri(part), storyPart))
            ?? throw new InvalidOperationException("object-story-part-not-found");
    }

    private static ImagePart AddImagePart(OpenXmlPartContainer owner, string contentType)
        => owner switch
        {
            MainDocumentPart part => part.AddImagePart(contentType),
            HeaderPart part => part.AddImagePart(contentType),
            FooterPart part => part.AddImagePart(contentType),
            FootnotesPart part => part.AddImagePart(contentType),
            EndnotesPart part => part.AddImagePart(contentType),
            WordprocessingCommentsPart part => part.AddImagePart(contentType),
            _ => throw new InvalidOperationException("drawing-story-part-does-not-support-images"),
        };

    private static void SaveStories(WordprocessingDocument document, IEnumerable<string> storyParts)
    {
        var selected = storyParts.Distinct(StringComparer.Ordinal).ToHashSet(StringComparer.Ordinal);
        var main = document.MainDocumentPart ?? throw new InvalidOperationException("main-document-part-not-found");
        static string Uri(OpenXmlPart part) => part.Uri.OriginalString.StartsWith('/') ? part.Uri.OriginalString : "/" + part.Uri.OriginalString;
        if (main.Document is not null && selected.Contains(Uri(main))) main.Document.Save();
        foreach (var part in main.HeaderParts.Where(part => part.Header is not null && selected.Contains(Uri(part)))) part.Header!.Save();
        foreach (var part in main.FooterParts.Where(part => part.Footer is not null && selected.Contains(Uri(part)))) part.Footer!.Save();
        if (main.FootnotesPart?.Footnotes is not null && selected.Contains(Uri(main.FootnotesPart))) main.FootnotesPart.Footnotes.Save();
        if (main.EndnotesPart?.Endnotes is not null && selected.Contains(Uri(main.EndnotesPart))) main.EndnotesPart.Endnotes.Save();
        if (main.WordprocessingCommentsPart?.Comments is not null && selected.Contains(Uri(main.WordprocessingCommentsPart))) main.WordprocessingCommentsPart.Comments.Save();
    }

    private sealed record PreparedDrawing(OpenXmlPartContainer OwnerPart, ImagePart ImagePart, A.Blip Blip, CheckboxDib Bitmap);

    private sealed class CheckboxDib
    {
        private readonly byte[] _bytes;
        private readonly int _pixelOffset;
        private readonly int _stride;
        private readonly bool _bottomUp;
        private readonly int _margin;

        private CheckboxDib(byte[] bytes, int pixelOffset, int width, int height, int stride, bool bottomUp, int margin, bool isChecked)
        {
            _bytes = bytes;
            _pixelOffset = pixelOffset;
            Width = width;
            Height = height;
            _stride = stride;
            _bottomUp = bottomUp;
            _margin = margin;
            Checked = isChecked;
        }

        public int Width { get; }
        public int Height { get; }
        public bool Checked { get; }

        public static CheckboxDib Open(byte[] bytes, string name)
        {
            string? rejected = null;
            for (var offset = 0; offset <= bytes.Length - 40; offset++)
            {
                if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4)) != 40) continue;
                var width = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset + 4, 4));
                var signedHeight = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset + 8, 4));
                var height = Math.Abs(signedHeight);
                var planes = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 12, 2));
                var bits = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset + 14, 2));
                var compression = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 16, 4));
                if (width is < 8 or > 128 || height is < 8 or > 128 || planes != 1 || bits != 24 || compression != 0) continue;
                if (width * 3 < height * 2 || height * 3 < width * 2) continue;
                var stride = ((width * 24 + 31) / 32) * 4;
                var pixelOffset = offset + 40;
                if (pixelOffset + stride * height > bytes.Length) continue;
                var margin = Math.Max(2, Math.Min(width, height) / 8);
                var candidate = new CheckboxDib(bytes, pixelOffset, width, height, stride, signedHeight > 0, margin, false);
                var interior = candidate.Pixels(margin, margin, width - margin, height - margin).ToArray();
                var edge = candidate.Pixels(0, 0, width, height)
                    .Where(pixel => pixel.X < margin || pixel.Y < margin || pixel.X >= width - margin || pixel.Y >= height - margin)
                    .ToArray();
                var whiteRatio = interior.Count(pixel => pixel.Brightness >= 245) / (double)interior.Length;
                var darkInterior = interior.Count(pixel => pixel.Brightness <= 100);
                var darkEdgeRatio = edge.Count(pixel => pixel.Brightness < 230) / (double)edge.Length;
                if (whiteRatio < 0.50 || darkEdgeRatio < 0.35)
                {
                    rejected = $"offset={offset}; size={width}x{height}; white={whiteRatio:F3}; edge={darkEdgeRatio:F3}";
                    continue;
                }
                return new CheckboxDib(bytes, pixelOffset, width, height, stride, signedHeight > 0, margin,
                    darkInterior >= Math.Max(3, Math.Min(width, height) / 3));
            }
            throw new InvalidOperationException($"drawing-is-not-supported-image-checkbox: {name}{(rejected is null ? string.Empty : $"; {rejected}")}");
        }

        public byte[] WithState(bool isChecked)
        {
            var result = (byte[])_bytes.Clone();
            for (var y = _margin; y < Height - _margin; y++)
            for (var x = _margin; x < Width - _margin; x++) Set(result, x, y, 255);
            if (isChecked)
            {
                var innerWidth = Width - 2 * _margin;
                var innerHeight = Height - 2 * _margin;
                var thickness = Math.Max(1, Math.Min(Width, Height) / 12);
                for (var step = 0; step <= innerWidth / 3; step++)
                    Paint(result, _margin + innerWidth / 7 + step, _margin + innerHeight / 2 + step * innerHeight / Math.Max(1, innerWidth), thickness);
                for (var step = 0; step <= innerWidth * 2 / 3; step++)
                    Paint(result, _margin + innerWidth * 4 / 9 + step, _margin + innerHeight * 5 / 7 - step * innerHeight / Math.Max(1, innerWidth), thickness);
            }
            return result;
        }

        private IEnumerable<Pixel> Pixels(int left, int top, int right, int bottom)
        {
            for (var y = top; y < bottom; y++)
            for (var x = left; x < right; x++)
            {
                var offset = Offset(x, y);
                yield return new Pixel(x, y, (_bytes[offset] + _bytes[offset + 1] + _bytes[offset + 2]) / 3);
            }
        }

        private void Paint(byte[] bytes, int x, int y, int thickness)
        {
            for (var dy = -thickness; dy <= thickness; dy++)
            for (var dx = -thickness; dx <= thickness; dx++)
                if (x + dx >= _margin && x + dx < Width - _margin && y + dy >= _margin && y + dy < Height - _margin)
                    Set(bytes, x + dx, y + dy, 0);
        }

        private void Set(byte[] bytes, int x, int y, byte value)
        {
            var offset = Offset(x, y);
            bytes[offset] = bytes[offset + 1] = bytes[offset + 2] = value;
        }

        private int Offset(int x, int y)
        {
            var row = _bottomUp ? Height - 1 - y : y;
            return _pixelOffset + row * _stride + x * 3;
        }

        private sealed record Pixel(int X, int Y, int Brightness);
    }
}

public sealed record SetDrawingCheckboxStateChange(DocxObjectAddress Drawing, bool Checked);
public sealed record SetDrawingCheckboxStateRequest(
    string Input,
    IReadOnlyList<SetDrawingCheckboxStateChange> Changes,
    string Output,
    string ReceiptOutput);
public sealed record DrawingCheckboxStateReadback(DocxObjectAddress Drawing, bool Checked, int PixelWidth, int PixelHeight);
public sealed record SetDrawingCheckboxStateReceipt(
    string Schema,
    string Provider,
    string ToolVersion,
    IReadOnlyList<DrawingCheckboxStateReadback> Changes,
    string Output);
