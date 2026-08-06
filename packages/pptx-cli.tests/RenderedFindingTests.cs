using System.Security.Cryptography;
using Dockit.Pptx;
using Xunit;

namespace Dockit.Pptx.Tests;

public class RenderedFindingTests
{
    [Fact]
    public void Maps_unseen_slide_text_overflow_to_one_text_shape_without_using_slide_identity()
    {
        var fixture = Fixture([
            Shape(41, "shape", "Unseen alpha narrative", new(1000, 1000, 2000, 1000)),
            Shape(77, "picture", "", new(6000, 1000, 2000, 1000))
        ]);
        var request = Request(fixture, new("finding-a", 1, fixture.Manifest.Pages[0].Sha256, "text-overflow", new(100, 100, 200, 100), "alpha narrative"));

        var result = RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, request);

        var binding = Assert.Single(result.Findings);
        Assert.Equal("unique", binding.Status);
        Assert.Equal("slide", binding.Target?.Scope);
        Assert.Equal(41U, binding.Target?.ShapeId);
        Assert.Equal("not-generated", binding.OperationDisposition);
        Assert.True(RenderedFindingValidator.Validate(fixture.Inspection, fixture.Manifest, request, result).Pass);
    }

    [Fact]
    public void Maps_master_footer_and_layout_picture_but_does_not_invent_geometry_corrections()
    {
        var fixture = Fixture(
            [Shape(2, "shape", "Body", new(1000, 1000, 1000, 1000))],
            layoutShapes: [Shape(12, "picture", "", new(6000, 1000, 1000, 1000))],
            masterShapes: [Shape(8, "shape", "Confidential footer", new(1000, 4500, 3000, 300))]);
        var footer = Request(fixture, new("footer", 1, fixture.Manifest.Pages[0].Sha256, "occlusion", new(100, 450, 300, 30), "Confidential"));
        var picture = Request(fixture, new("layout-image", 1, fixture.Manifest.Pages[0].Sha256, "edge-clipping", new(600, 100, 100, 100)));

        var footerBinding = Assert.Single(RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, footer).Findings);
        var pictureBinding = Assert.Single(RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, picture).Findings);

        Assert.Equal(("master", 8U, "not-generated"), (footerBinding.Target?.Scope, footerBinding.Target?.ShapeId, footerBinding.OperationDisposition));
        Assert.Equal(("layout", 12U, "not-generated"), (pictureBinding.Target?.Scope, pictureBinding.Target?.ShapeId, pictureBinding.OperationDisposition));
    }

    [Fact]
    public void Ambiguous_overlap_and_duplicate_text_fail_closed_without_selecting_a_shape()
    {
        var fixture = Fixture(
            [
                Shape(2, "groupShape", "Repeated label", new(1000, 1000, 2000, 1000)),
                Shape(3, "shape", "Repeated label", new(1500, 1000, 2000, 1000))
            ],
            layoutShapes: [Shape(9, "picture", "", new(1000, 1000, 2000, 1000))]);
        var withText = Request(fixture, new("duplicate", 1, fixture.Manifest.Pages[0].Sha256, "text-overflow", new(150, 100, 50, 100), "Repeated label"));
        var withoutText = Request(fixture, new("overlap", 1, fixture.Manifest.Pages[0].Sha256, "occlusion", new(150, 100, 50, 100)));

        var duplicate = Assert.Single(RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, withText).Findings);
        var overlap = Assert.Single(RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, withoutText).Findings);

        Assert.Equal("ambiguous", duplicate.Status); Assert.Null(duplicate.Target); Assert.Equal(2, duplicate.Candidates.Count);
        Assert.Equal("ambiguous", overlap.Status); Assert.Null(overlap.Target); Assert.Equal(3, overlap.Candidates.Count);
    }

    [Fact]
    public void Validator_rejects_forged_binding_and_mapper_rejects_raster_rebinding()
    {
        var fixture = Fixture([Shape(2, "shape", "Bound text", new(1000, 1000, 2000, 1000))]);
        var request = Request(fixture, new("bound", 1, fixture.Manifest.Pages[0].Sha256, "text-overflow", new(100, 100, 200, 100), "Bound"));
        var valid = RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, request);
        var original = Assert.Single(valid.Findings); var target = original.Target!;
        var forged = valid with { Findings = [original with { Target = target with { ShapeId = 999U } }] };

        Assert.False(RenderedFindingValidator.Validate(fixture.Inspection, fixture.Manifest, request, forged).Pass);
        var rebound = request with { Findings = [request.Findings[0] with { RasterSha256 = new string('f', 64) }] };
        Assert.Throws<InvalidOperationException>(() => RenderedFindingMapper.Map(fixture.Inspection, fixture.Manifest, rebound));
    }

    private static (PresentationDetailReport Inspection, RenderFindingManifest Manifest) Fixture(
        IReadOnlyList<ShapeDetail> slideShapes,
        IReadOnlyList<ShapeDetail>? layoutShapes = null,
        IReadOnlyList<ShapeDetail>? masterShapes = null)
    {
        var png = Path.Combine(Path.GetTempPath(), $"render-finding-{Guid.NewGuid():N}.png");
        var bytes = new byte[24]; new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }.CopyTo(bytes, 0); "IHDR"u8.CopyTo(bytes.AsSpan(12, 4));
        System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), 1000); System.Buffers.Binary.BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20, 4), 500); File.WriteAllBytes(png, bytes);
        var hash = Convert.ToHexStringLower(SHA256.HashData(bytes)); var artifact = new string('a', 64);
        var layout = new LayoutDetail("ppt/slideLayouts/slideLayout1.xml", "", null, new string('b', 64), layoutShapes ?? []);
        var master = new MasterDetail("ppt/slideMasters/slideMaster1.xml", "", new string('c', 64), null, null, masterShapes ?? [], [layout]);
        var slide = new SlideDetailReport(1, "ppt/slides/slide1.xml", master.Path, layout.Path, slideShapes);
        return (new PresentationDetailReport("current.pptx", 1, new SlideSizeInfo(10000, 5000), [master], [slide]), new RenderFindingManifest(new RenderFindingArtifact(artifact), [new RenderFindingPage(1, png, hash)]));
    }

    private static RenderFindingRequest Request((PresentationDetailReport Inspection, RenderFindingManifest Manifest) fixture, RenderFindingCandidate finding) => new("tiwater.pptx-render-findings/v1", fixture.Manifest.Artifact.Sha256, [finding]);

    private static ShapeDetail Shape(uint id, string kind, string text, TransformInfo transform) => new(id, $"shape-{id}", kind, 0, null, null, null, text, transform, [], [], null);
}
