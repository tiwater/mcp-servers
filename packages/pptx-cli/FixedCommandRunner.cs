using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Pptx;

public static class FixedCommandRunner
{
    private sealed record Artifact(string Path, string Sha256, long Bytes);

    private static readonly IReadOnlySet<string> CommandSet = new HashSet<string>(StringComparer.Ordinal)
    {
        "pptx_apply_template",
        "pptx_apply_format",
        "pptx_set_shape_geometry",
        "pptx_replace_picture_image",
    };

    public static IReadOnlyCollection<string> Commands => CommandSet.ToArray();

    public static bool IsCommand(string command) => CommandSet.Contains(command);

    public static int Run(string command, string[] args)
    {
        if (!IsCommand(command))
            throw new InvalidOperationException($"Unknown fixed PPTX command: {command}");
        if (args.Length != 1)
            throw new InvalidOperationException($"{command} requires <request.json>");

        string? output = null;
        string? receiptOutput = null;
        Artifact? inputArtifact = null;

        try
        {
            var request = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
                ?? throw new InvalidOperationException("fixed-pptx-request-invalid");
            var input = RequirePath(request, "input");
            output = RequirePath(request, "output");
            receiptOutput = RequirePath(request, "receiptOutput");
            RequireNewPath(output, "output");
            RequireNewPath(receiptOutput, "receiptOutput");
            if (PathsEqual(output, receiptOutput))
                throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
            if (PathsEqual(output, input))
                throw new InvalidOperationException("output-must-not-overwrite-input");

            inputArtifact = Describe(input);
            var result = command switch
            {
                "pptx_apply_template" => RunTemplate(request, input, output),
                "pptx_apply_format" => RunFormat(request, input, output),
                "pptx_set_shape_geometry" => RunShapeGeometry(request, input, output),
                "pptx_replace_picture_image" => RunPictureImage(request, input, output),
                _ => throw new InvalidOperationException($"Unknown fixed PPTX command: {command}"),
            };
            var pass = result.Pass && File.Exists(output);
            var outputArtifact = pass ? Describe(output) : null;
            if (!pass && File.Exists(output)) File.Delete(output);

            var receiptPayload = new
            {
                schema = "tiwater.office.fixed-edit-receipt/v2",
                tool = command,
                pass,
                input = inputArtifact,
                acceptedCall = request,
                output = outputArtifact,
                operationCount = result.OperationCount,
                appliedCount = result.AppliedCount,
                changes = result.Changes,
                issues = result.Issues,
            };
            var receipt = WriteJsonArtifact(receiptOutput, receiptPayload);
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                tool = command,
                receipt,
                output = outputArtifact,
                summary = new
                {
                    pass,
                    operationCount = result.OperationCount,
                    appliedCount = result.AppliedCount,
                },
            }, Json.Options));
            return pass ? 0 : 1;
        }
        catch (Exception error)
        {
            if (output is not null && File.Exists(output)) File.Delete(output);

            if (receiptOutput is not null && !File.Exists(receiptOutput))
            {
                try
                {
                    var receipt = WriteJsonArtifact(receiptOutput, new
                    {
                        schema = "tiwater.office.fixed-edit-receipt/v2",
                        tool = command,
                        pass = false,
                        input = inputArtifact,
                        output = (Artifact?)null,
                        error = error.Message,
                    });
                    Console.WriteLine(JsonSerializer.Serialize(new
                    {
                        tool = command,
                        receipt,
                        output = (Artifact?)null,
                        summary = new { pass = false, operationCount = 0, appliedCount = 0 },
                    }, Json.Options));
                    return 1;
                }
                catch
                {
                    // Keep the provider error when a failure receipt cannot be written.
                }
            }

            Console.Error.WriteLine(error.Message);
            return 1;
        }
    }

    private static FixedResult RunTemplate(JsonObject request, string input, string output)
    {
        var template = RequirePath(request, "template");
        var plan = new TemplateApplicationPlan(
            RequireString(request, "targetMasterPath"),
            DeserializeRequired<IReadOnlyList<SlideLayoutAssignment>>(request, "slides"),
            "preserve");
        var result = TemplateApplicator.Apply(input, template, plan, output);
        return new(
            result.Issues.Count == 0 && result.ChangedSlideCount == plan.Slides.Count,
            plan.Slides.Count,
            result.ChangedSlideCount,
            result.MaterializedLayoutShapes,
            result.Issues);
    }

    private static FixedResult RunFormat(JsonObject request, string input, string output)
    {
        var operations = DeserializeRequired<IReadOnlyList<FormatEditOperation>>(request, "changes");
        if (operations.Count == 0)
            throw new InvalidOperationException("changes-must-contain-at-least-one-item");
        var result = FormatEditor.Apply(input, new FormatEditPlan(operations), output);
        return new(result.Issues.Count == 0, operations.Count, result.ChangedCount, result.Changes, result.Issues);
    }

    private static FixedResult RunShapeGeometry(JsonObject request, string input, string output)
    {
        var changes = DeserializeRequired<IReadOnlyList<ShapeGeometryChange>>(request, "changes");
        if (changes.Count == 0)
            throw new InvalidOperationException("changes-must-contain-at-least-one-item");
        var result = ShapeGeometryEditor.Apply(input, new ShapeGeometryPlan(changes), output);
        return new(result.Issues.Count == 0 && result.AppliedCount == changes.Count,
            changes.Count, result.AppliedCount, result.Changes, result.Issues);
    }

    private static FixedResult RunPictureImage(JsonObject request, string input, string output)
    {
        var changes = DeserializeRequired<IReadOnlyList<PictureImageChange>>(request, "changes");
        if (changes.Count == 0)
            throw new InvalidOperationException("changes-must-contain-at-least-one-item");
        var result = PictureImageEditor.Apply(input, new PictureImagePlan(changes), output);
        return new(result.Issues.Count == 0 && result.AppliedCount == changes.Count,
            changes.Count, result.AppliedCount, result.Changes, result.Issues);
    }

    private sealed record FixedResult(
        bool Pass,
        int OperationCount,
        int AppliedCount,
        object? Changes,
        object? Issues);

    private static T DeserializeRequired<T>(JsonObject request, string property)
    {
        if (request[property] is null)
            throw new InvalidOperationException($"{property}-is-required");
        return request[property]!.Deserialize<T>(Json.Options)
            ?? throw new InvalidOperationException($"{property}-is-invalid");
    }

    private static string RequireString(JsonObject request, string property)
    {
        if (request[property] is not JsonValue value || !value.TryGetValue<string>(out var text) || string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException($"{property}-is-required");
        return text;
    }

    private static string RequirePath(JsonObject request, string property)
        => Path.GetFullPath(RequireString(request, property));

    private static void RequireNewPath(string path, string property)
    {
        if (File.Exists(path) || Directory.Exists(path))
            throw new InvalidOperationException($"{property}-already-exists");
        var directory = Path.GetDirectoryName(path);
        if (string.IsNullOrWhiteSpace(directory))
            throw new InvalidOperationException($"{property}-directory-not-found");
        Directory.CreateDirectory(directory);
    }

    private static bool PathsEqual(string left, string right)
        => StringComparer.OrdinalIgnoreCase.Equals(Path.GetFullPath(left), Path.GetFullPath(right));

    private static Artifact Describe(string path)
    {
        using var stream = File.OpenRead(path);
        return new(Path.GetFullPath(path), Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), stream.Length);
    }

    private static Artifact WriteJsonArtifact<T>(string path, T payload)
    {
        var bytes = JsonSerializer.SerializeToUtf8Bytes(payload, Json.Options);
        using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        {
            stream.Write(bytes);
            stream.WriteByte((byte)'\n');
        }
        return Describe(path);
    }
}
