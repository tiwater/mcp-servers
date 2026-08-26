using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace Dockit.Pptx.Tests;

public sealed class PptxObjectEditorTests
{
    private static readonly byte[] FirstPng = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
    private static readonly byte[] SecondPng = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z2S0AAAAASUVORK5CYII=");

    [Fact]
    public void Published_object_edit_contracts_are_closed_fixed_action_batches()
    {
        foreach (var (file, id, fields) in new[]
        {
            ("tiwater.pptx-shape-geometry-v1.schema.json", "tiwater.pptx-shape-geometry/v1", new[] { "slideNumber", "shapeId", "x", "y", "cx", "cy" }),
            ("tiwater.pptx-picture-replacement-v1.schema.json", "tiwater.pptx-picture-replacement/v1", new[] { "slideNumber", "shapeId", "image" }),
        })
        {
            using var schema = JsonDocument.Parse(File.ReadAllText(RepositoryPath($"packages/pptx-cli/contracts/{file}")));
            Assert.Equal(id, schema.RootElement.GetProperty("$id").GetString());
            Assert.False(schema.RootElement.GetProperty("additionalProperties").GetBoolean());
            var changes = schema.RootElement.GetProperty("properties").GetProperty("changes");
            Assert.Equal(1, changes.GetProperty("minItems").GetInt32());
            var item = changes.GetProperty("items");
            Assert.False(item.GetProperty("additionalProperties").GetBoolean());
            Assert.Equal(fields, item.GetProperty("required").EnumerateArray().Select(value => value.GetString()));
            Assert.False(item.GetProperty("properties").TryGetProperty("type", out _));
        }

        foreach (var (file, id) in new[]
        {
            ("tiwater.pptx-shape-geometry-result-v1.schema.json", "tiwater.pptx-shape-geometry-result/v1"),
            ("tiwater.pptx-picture-replacement-result-v1.schema.json", "tiwater.pptx-picture-replacement-result/v1"),
        })
        {
            using var schema = JsonDocument.Parse(File.ReadAllText(RepositoryPath($"packages/pptx-cli/contracts/{file}")));
            Assert.Equal(id, schema.RootElement.GetProperty("$id").GetString());
            Assert.False(schema.RootElement.GetProperty("additionalProperties").GetBoolean());
            Assert.Equal(new[] { "input", "output", "operationCount", "appliedCount", "changes", "issues" },
                schema.RootElement.GetProperty("required").EnumerateArray().Select(value => value.GetString()));
            Assert.False(schema.RootElement.GetProperty("properties").GetProperty("changes").GetProperty("items").GetProperty("additionalProperties").GetBoolean());
            Assert.False(schema.RootElement.GetProperty("properties").GetProperty("issues").GetProperty("items").GetProperty("additionalProperties").GetBoolean());
        }
    }

    [Fact]
    public void SetShapeGeometry_updates_each_supported_top_level_kind_and_only_native_bounds()
    {
        var source = CreateFixture(sharedPictureRelationship: false);
        var beforeHash = Sha256(source);
        var output = Temporary("geometry");
        var changes = new[]
        {
            new ShapeGeometryChange(1, 2U, -101L, 102L, 103L, 104L),
            new ShapeGeometryChange(1, 3U, 201L, -202L, 203L, 204L),
            new ShapeGeometryChange(1, 4U, 301L, 302L, 303L, 304L),
            new ShapeGeometryChange(1, 5U, 401L, 402L, 403L, 404L),
            new ShapeGeometryChange(2, 2U, 501L, 502L, 503L, 504L),
        };

        var result = ShapeGeometryEditor.Apply(source, new ShapeGeometryPlan(changes), output);

        Assert.Empty(result.Issues);
        Assert.Equal(changes.Length, result.OperationCount);
        Assert.Equal(changes.Length, result.AppliedCount);
        Assert.Equal(beforeHash, Sha256(source));
        var detail = Inspector.InspectDetail(output);
        foreach (var change in changes)
            Assert.Equal(new TransformInfo(change.X, change.Y, change.Cx, change.Cy), detail.Slides[change.SlideNumber - 1].Shapes.Single(shape => shape.ShapeId == change.ShapeId).Transform);
        using var edited = PresentationDocument.Open(output, false);
        var firstSlide = edited.PresentationPart!.SlideParts.First().Slide;
        var shapeTransform = firstSlide.Descendants<P.Shape>().Single(shape => ShapeId(shape) == 2U).ShapeProperties!.Transform2D!;
        Assert.Equal(17, shapeTransform.Rotation?.Value);
        Assert.True(shapeTransform.HorizontalFlip?.Value);
        var group = firstSlide.Descendants<P.GroupShape>().Single(shape => ShapeId(shape) == 5U);
        Assert.Equal((11L, 12L, 13L, 14L), ChildBounds(group));
        Assert.Equal(Validator.Validate(source).Errors, Validator.Validate(output).Errors);
    }

    [Fact]
    public void SetShapeGeometry_preflights_missing_ambiguous_duplicate_and_untransformable_targets_atomically()
    {
        var source = CreateFixture(sharedPictureRelationship: false, ambiguousShapeId: true);
        var sourceHash = Sha256(source);
        foreach (var plan in new[]
        {
            new ShapeGeometryPlan([new(1, 999U, 1, 2, 3, 4)]),
            new ShapeGeometryPlan([new(1, 3U, 1, 2, 3, 4)]),
            new ShapeGeometryPlan([new(1, 6U, 1, 2, 3, 4)]),
            new ShapeGeometryPlan([new(1, 2U, 1, 2, 3, 4), new(1, 2U, 5, 6, 7, 8)]),
            new ShapeGeometryPlan([new(0, 2U, 1, 2, 3, 4)]),
            new ShapeGeometryPlan([new(1, 2U, 1, 2, 0, 4)]),
        })
        {
            var output = Temporary("geometry-rejected");
            var result = ShapeGeometryEditor.Apply(source, plan, output);
            Assert.NotEmpty(result.Issues);
            Assert.Equal(0, result.AppliedCount);
            Assert.False(File.Exists(output));
            Assert.Equal(sourceHash, Sha256(source));
        }
    }

    [Fact]
    public void SetShapeGeometry_is_exact_and_idempotent_across_boundary_values()
    {
        var source = CreateFixture(sharedPictureRelationship: false);
        var random = new Random(3126);
        for (var iteration = 0; iteration < 20; iteration++)
        {
            var expected = new TransformInfo(random.NextInt64(-2_000_000, 2_000_000), random.NextInt64(-2_000_000, 2_000_000), random.NextInt64(1, 4_000_000), random.NextInt64(1, 4_000_000));
            var first = Temporary("geometry-property");
            var second = Temporary("geometry-idempotent");
            var plan = new ShapeGeometryPlan([new(1, 2U, expected.X, expected.Y, expected.Cx, expected.Cy)]);
            Assert.Empty(ShapeGeometryEditor.Apply(source, plan, first).Issues);
            Assert.Empty(ShapeGeometryEditor.Apply(first, plan, second).Issues);
            Assert.Equal(expected, Inspector.InspectDetail(first).Slides[0].Shapes.Single(shape => shape.ShapeId == 2U).Transform);
            Assert.Equal(expected, Inspector.InspectDetail(second).Slides[0].Shapes.Single(shape => shape.ShapeId == 2U).Transform);
        }
    }

    [Fact]
    public void ReplacePictureImage_preserves_picture_identity_geometry_crop_and_shared_media_consumers()
    {
        var source = CreateFixture(sharedPictureRelationship: true);
        var replacement = TemporaryFile("replacement.png", SecondPng);
        var output = Temporary("picture");
        var before = Inspector.InspectDetail(source).Slides[0];
        var targetBefore = before.Shapes.Single(shape => shape.ShapeId == 3U);
        var peerBefore = before.Shapes.Single(shape => shape.ShapeId == 7U);

        var result = PictureImageEditor.Apply(source, new PictureImagePlan([new(1, 3U, replacement)]), output);

        Assert.Empty(result.Issues);
        Assert.Equal(1, result.AppliedCount);
        var after = Inspector.InspectDetail(output).Slides[0];
        var targetAfter = after.Shapes.Single(shape => shape.ShapeId == 3U);
        var peerAfter = after.Shapes.Single(shape => shape.ShapeId == 7U);
        Assert.Equal((targetBefore.ShapeId, targetBefore.Name, targetBefore.Transform), (targetAfter.ShapeId, targetAfter.Name, targetAfter.Transform));
        Assert.NotEqual(targetBefore.MediaSha256, targetAfter.MediaSha256);
        Assert.Equal(Sha256(SecondPng), targetAfter.MediaSha256);
        Assert.Equal(peerBefore.MediaSha256, peerAfter.MediaSha256);
        using var edited = PresentationDocument.Open(output, false);
        var picture = edited.PresentationPart!.SlideParts.First().Slide.Descendants<P.Picture>().Single(value => ShapeId(value) == 3U);
        Assert.Equal((1, 2, 3, 4), Crop(picture));
        Assert.Equal(Validator.Validate(source).Errors, Validator.Validate(output).Errors);
    }

    [Fact]
    public void ReplacePictureImage_rejects_wrong_media_missing_ambiguous_and_non_picture_targets_atomically()
    {
        var source = CreateFixture(sharedPictureRelationship: false);
        var png = TemporaryFile("replacement.png", SecondPng);
        var jpegNamedPng = TemporaryFile("replacement.jpg", SecondPng);
        var sourceHash = Sha256(source);
        foreach (var plan in new[]
        {
            new PictureImagePlan([new(1, 999U, png)]),
            new PictureImagePlan([new(1, 2U, png)]),
            new PictureImagePlan([new(1, 3U, png), new(1, 3U, png)]),
            new PictureImagePlan([new(1, 3U, jpegNamedPng)]),
            new PictureImagePlan([new(1, 3U, png), new(1, 8U, png)]),
        })
        {
            var output = Temporary("picture-rejected");
            var result = PictureImageEditor.Apply(source, plan, output);
            Assert.NotEmpty(result.Issues);
            Assert.Equal(0, result.AppliedCount);
            Assert.False(File.Exists(output));
            Assert.Equal(sourceHash, Sha256(source));
        }

        var ambiguous = CreateFixture(sharedPictureRelationship: false, ambiguousShapeId: true);
        var ambiguousOutput = Temporary("picture-ambiguous");
        var ambiguousResult = PictureImageEditor.Apply(ambiguous, new PictureImagePlan([new(1, 3U, png)]), ambiguousOutput);
        Assert.Equal("shape reference is ambiguous", Assert.Single(ambiguousResult.Issues).Message);
        Assert.False(File.Exists(ambiguousOutput));
    }

    [Fact]
    public async Task Fixed_cli_commands_accept_only_their_closed_plans_and_preserve_existing_command_compatibility()
    {
        var source = CreateFixture(sharedPictureRelationship: false);
        var geometryPlan = TemporaryFile("geometry.json", JsonSerializer.SerializeToUtf8Bytes(new { changes = new[] { new { slideNumber = 1, shapeId = 2, x = -1, y = 2, cx = 3, cy = 4 } } }));
        var geometryOutput = Temporary("geometry-cli");
        Assert.Equal(0, await Dockit.Pptx.Cli.Cli.RunAsync(["set-shape-geometry", source, geometryPlan, geometryOutput]));
        Assert.Equal(new TransformInfo(-1, 2, 3, 4), Inspector.InspectDetail(geometryOutput).Slides[0].Shapes.Single(shape => shape.ShapeId == 2U).Transform);

        var image = TemporaryFile("replacement.png", SecondPng);
        var picturePlan = TemporaryFile("picture.json", JsonSerializer.SerializeToUtf8Bytes(new { changes = new[] { new { slideNumber = 1, shapeId = 3, image } } }));
        Assert.Equal(0, await Dockit.Pptx.Cli.Cli.RunAsync(["replace-picture-image", source, picturePlan, Temporary("picture-cli")]));

        var injectedPlan = TemporaryFile("injected.json", JsonSerializer.SerializeToUtf8Bytes(new { changes = new[] { new { slideNumber = 1, shapeId = 2, x = 1, y = 2, cx = 3, cy = 4, type = "replace-picture-image" } } }));
        var rejectedOutput = Temporary("geometry-cli-rejected");
        Assert.Equal(1, await Dockit.Pptx.Cli.Cli.RunAsync(["set-shape-geometry", source, injectedPlan, rejectedOutput]));
        Assert.False(File.Exists(rejectedOutput));
    }

    [Fact]
    public void Existing_format_action_remains_run_only_and_does_not_duplicate_geometry()
    {
        var source = CreateFixture(sharedPictureRelationship: false);
        var before = Inspector.InspectDetail(source).Slides[0].Shapes.Single(shape => shape.ShapeId == 2U).Transform;
        var output = Temporary("format-compatibility");
        var result = FormatEditor.Apply(source, new FormatEditPlan([new(1, 2U, 0, "Arial", null, null, null, null)]), output);
        Assert.Empty(result.Issues);
        Assert.Equal(before, Inspector.InspectDetail(output).Slides[0].Shapes.Single(shape => shape.ShapeId == 2U).Transform);
    }

    private static string CreateFixture(bool sharedPictureRelationship, bool ambiguousShapeId = false)
    {
        var path = Temporary("object-fixture");
        using var document = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentation = document.AddPresentationPart();
        var master = presentation.AddNewPart<SlideMasterPart>("rIdMaster");
        var layout = master.AddNewPart<SlideLayoutPart>("rIdLayout");
        layout.SlideLayout = new P.SlideLayout(new P.CommonSlideData(ShapeTree()), new P.ColorMapOverride(new A.MasterColorMapping()));
        master.SlideMaster = new P.SlideMaster(new P.CommonSlideData(ShapeTree()), ColorMap(), new P.SlideLayoutIdList(new P.SlideLayoutId { Id = 2147483648U, RelationshipId = "rIdLayout" }), new P.TextStyles());
        layout.AddPart(master, "rIdMasterBack");
        var slide1 = presentation.AddNewPart<SlidePart>("rIdSlide1");
        slide1.AddPart(layout, "rIdLayoutForSlide");
        var image = slide1.AddImagePart(ImagePartType.Png, "rIdPicture");
        image.FeedData(new MemoryStream(FirstPng));
        var text = TextShape(2U, "content", 10, 20, 30, 40);
        text.ShapeProperties!.Transform2D!.Rotation = 17;
        text.ShapeProperties.Transform2D.HorizontalFlip = true;
        var picture = Picture(3U, "rIdPicture", 50, 60, 70, 80);
        var frame = new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(new P.NonVisualDrawingProperties { Id = 4U, Name = "frame" }, new P.NonVisualGraphicFrameDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(new A.Offset { X = 90, Y = 100 }, new A.Extents { Cx = 110, Cy = 120 }),
            new A.Graphic(new A.GraphicData { Uri = "urn:fixed:test" }));
        var group = new P.GroupShape(
            new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 5U, Name = "group" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(new A.Offset { X = 130, Y = 140 }, new A.Extents { Cx = 150, Cy = 160 }, new A.ChildOffset { X = 11, Y = 12 }, new A.ChildExtents { Cx = 13, Cy = 14 })),
            new P.ShapeTree());
        var noTransform = new P.Shape(new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = 6U, Name = "inherited" }, new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()), new P.ShapeProperties(), new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("inherited")))));
        var children = new List<OpenXmlElement> { text, picture, frame, group, noTransform };
        if (sharedPictureRelationship) children.Add(Picture(7U, "rIdPicture", 170, 180, 190, 200));
        if (ambiguousShapeId) children.Add(TextShape(3U, "duplicate id", 210, 220, 230, 240));
        children.Add(new P.ConnectionShape(new P.NonVisualConnectionShapeProperties(new P.NonVisualDrawingProperties { Id = 8U, Name = "unsupported" }, new P.NonVisualConnectorShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()), new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 1, Y = 2 }, new A.Extents { Cx = 3, Cy = 4 }))));
        slide1.Slide = new P.Slide(new P.CommonSlideData(ShapeTree(children.ToArray())), new P.ColorMapOverride(new A.MasterColorMapping()));
        var slide2 = presentation.AddNewPart<SlidePart>("rIdSlide2");
        slide2.AddPart(layout, "rIdLayoutForSlide");
        slide2.Slide = new P.Slide(new P.CommonSlideData(ShapeTree(TextShape(2U, "second slide", 1, 2, 3, 4))), new P.ColorMapOverride(new A.MasterColorMapping()));
        presentation.Presentation = new P.Presentation(
            new P.SlideMasterIdList(new P.SlideMasterId { Id = 2147483648U, RelationshipId = "rIdMaster" }),
            new P.SlideIdList(new P.SlideId { Id = 256U, RelationshipId = "rIdSlide1" }, new P.SlideId { Id = 257U, RelationshipId = "rIdSlide2" }),
            new P.SlideSize { Cx = 9_144_000, Cy = 6_858_000 }, new P.NotesSize { Cx = 6_858_000, Cy = 9_144_000 });
        slide1.Slide.Save(); slide2.Slide.Save(); layout.SlideLayout.Save(); master.SlideMaster.Save(); presentation.Presentation.Save();
        return path;
    }

    private static P.Shape TextShape(uint id, string text, long x, long y, long cx, long cy) => new(
        new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = id, Name = text }, new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
        new P.ShapeProperties(new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy })),
        new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(text)))));

    private static P.Picture Picture(uint id, string relationshipId, long x, long y, long cx, long cy) => new(
        new P.NonVisualPictureProperties(new P.NonVisualDrawingProperties { Id = id, Name = $"picture-{id}" }, new P.NonVisualPictureDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
        new P.BlipFill(new A.Blip { Embed = relationshipId }, new A.SourceRectangle { Left = 1, Top = 2, Right = 3, Bottom = 4 }, new A.Stretch(new A.FillRectangle())),
        new P.ShapeProperties(new A.Transform2D(new A.Offset { X = x, Y = y }, new A.Extents { Cx = cx, Cy = cy })));

    private static P.ShapeTree ShapeTree(params OpenXmlElement[] children)
    {
        var tree = new P.ShapeTree(new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties { Id = 1U, Name = "" }, new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()), new P.GroupShapeProperties(new A.TransformGroup()));
        tree.Append(children); return tree;
    }

    private static P.ColorMap ColorMap() => new() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink };
    private static uint ShapeId(OpenXmlElement value) => value.Descendants<P.NonVisualDrawingProperties>().First().Id!.Value;
    private static (long, long, long, long) ChildBounds(P.GroupShape group) { var t = group.GroupShapeProperties!.TransformGroup!; return (t.ChildOffset!.X!.Value, t.ChildOffset.Y!.Value, t.ChildExtents!.Cx!.Value, t.ChildExtents.Cy!.Value); }
    private static (int, int, int, int) Crop(P.Picture picture) { var r = picture.BlipFill!.SourceRectangle!; return (r.Left!.Value, r.Top!.Value, r.Right!.Value, r.Bottom!.Value); }
    private static string Temporary(string stem) => Path.Combine(Path.GetTempPath(), $"pptx-{stem}-{Guid.NewGuid():N}.pptx");
    private static string TemporaryFile(string name, byte[] bytes) { var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}-{name}"); File.WriteAllBytes(path, bytes); return path; }
    private static string Sha256(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static string RepositoryPath(string relativePath) { for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent) { var candidate = Path.Combine(directory.FullName, relativePath); if (File.Exists(candidate)) return candidate; } throw new FileNotFoundException(relativePath); }
}
