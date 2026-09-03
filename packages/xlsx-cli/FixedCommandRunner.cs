using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Xlsx;

public static class FixedCommandRunner
{
    private sealed record Definition(string OperationType);
    private sealed record Artifact(string Path, string Sha256, long Bytes);

    private static readonly IReadOnlyDictionary<string, Definition> Definitions =
        new Dictionary<string, Definition>(StringComparer.Ordinal)
        {
            ["xlsx_set_cell_value"] = new("setCellValue"),
            ["xlsx_set_cell_number_format"] = new("setCellNumberFormat"),
            ["xlsx_set_rich_text_cell_value"] = new("setRichTextCellValue"),
            ["xlsx_set_range_values"] = new("setRangeValues"),
            ["xlsx_insert_rows"] = new("insertRows"),
            ["xlsx_delete_rows"] = new("deleteRows"),
            ["xlsx_copy_row"] = new("copyRow"),
            ["xlsx_expand_section_rows"] = new("expandSectionRows"),
            ["xlsx_set_print_area"] = new("setPrintArea"),
            ["xlsx_set_page_setup"] = new("setPageSetup"),
            ["xlsx_set_row_page_breaks"] = new("setRowPageBreaks"),
            ["xlsx_set_column_width"] = new("setColumnWidth"),
        };

    public static IReadOnlyCollection<string> Commands => Definitions.Keys.ToArray();

    public static bool IsCommand(string command) => Definitions.ContainsKey(command);

    public static int Run(string command, string[] args)
    {
        if (!Definitions.TryGetValue(command, out var definition))
            throw new InvalidOperationException($"Unknown fixed XLSX command: {command}");
        if (args.Length != 1)
            throw new InvalidOperationException($"{command} requires <request.json>");

        string? output = null;
        string? receiptOutput = null;
        Artifact? inputArtifact = null;
        var inPlace = false;

        try
        {
            var root = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
                ?? throw new InvalidOperationException("fixed-xlsx-request-invalid");
            var input = RequirePath(root, "input");
            output = RequirePath(root, "output");
            receiptOutput = RequirePath(root, "receiptOutput");
            inPlace = PathsEqual(output, input);
            if (!inPlace) RequireNewPath(output, "output");
            RequireNewPath(receiptOutput, "receiptOutput");
            if (PathsEqual(output, receiptOutput))
                throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");

            var changes = RequireChanges(root);
            if (changes.Count == 0)
                throw new InvalidOperationException("changes-must-contain-at-least-one-item");

            var operations = BuildOperations(definition, changes);
            inputArtifact = Describe(input);
            var editResult = Editor.Apply(input, output, operations);
            var applied = editResult.AppliedOperations.Select((operation, index) => new
            {
                index,
                applied = operation.Applied,
                detail = operation.Detail,
                errorCode = operation.ErrorCode,
            }).ToArray();
            var pass = editResult.AppliedOperations.Count == operations.Count
                && editResult.AppliedOperations.All(operation => operation.Applied)
                && File.Exists(output);
            var outputArtifact = pass ? Describe(output) : null;
            if (!pass && !inPlace && File.Exists(output)) File.Delete(output);

            var receiptPayload = new
            {
                schema = "tiwater.office.fixed-edit-receipt/v2",
                tool = command,
                pass,
                input = inputArtifact,
                acceptedCall = root,
                output = outputArtifact,
                operationCount = operations.Count,
                appliedOperations = applied,
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
                    operationCount = operations.Count,
                    appliedCount = applied.Count(operation => operation.applied),
                },
            }, Json.Options));
            return pass ? 0 : 1;
        }
        catch (Exception error)
        {
            if (!inPlace && output is not null && File.Exists(output)) File.Delete(output);

            if (receiptOutput is not null && !File.Exists(receiptOutput))
            {
                try
                {
                    var receiptPayload = new
                    {
                        schema = "tiwater.office.fixed-edit-receipt/v2",
                        tool = command,
                        pass = false,
                        input = inputArtifact,
                        output = (Artifact?)null,
                        error = error.Message,
                    };
                    var receipt = WriteJsonArtifact(receiptOutput, receiptPayload);
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
                    // Preserve the original provider error when a failure receipt cannot be written.
                }
            }

            Console.Error.WriteLine(error.Message);
            return 1;
        }
    }

    private static IReadOnlyList<XlsxEditOperation> BuildOperations(Definition definition, JsonArray changes)
        => changes.Select(change =>
        {
            var operation = change as JsonObject
                ?? throw new InvalidOperationException("change-must-be-an-object");
            var normalized = operation.DeepClone() as JsonObject
                ?? throw new InvalidOperationException("change-must-be-an-object");
            if (definition.OperationType is "setCellValue" or "setRangeValues")
                NormalizePrimitiveValues(normalized, definition.OperationType == "setRangeValues");
            normalized["type"] = definition.OperationType;
            return normalized.Deserialize<XlsxEditOperation>(Json.Options)
                ?? throw new InvalidOperationException("change-does-not-match-fixed-operation");
        }).ToArray();

    private static JsonArray RequireChanges(JsonObject root)
        => root["changes"] as JsonArray
            ?? throw new InvalidOperationException("changes-is-required-and-must-be-an-array");

    private static void NormalizePrimitiveValues(JsonObject change, bool range)
    {
        if (!range)
        {
            if (change.ContainsKey("value"))
                change["value"] = ScalarText(change["value"]);
            return;
        }

        if (change["values"] is not JsonArray rows)
            throw new InvalidOperationException("values-must-be-an-array");
        foreach (var row in rows)
        {
            if (row is not JsonArray cells)
                throw new InvalidOperationException("values-must-contain-arrays");
            for (var index = 0; index < cells.Count; index++)
                cells[index] = ScalarText(cells[index]);
        }
    }

    private static string ScalarText(JsonNode? value)
    {
        if (value is null) return string.Empty;
        if (value is JsonValue jsonValue && jsonValue.TryGetValue<string>(out var text)) return text;
        if (value is JsonValue && value.ToJsonString() is var raw)
            return raw == "null" ? string.Empty : raw;
        throw new InvalidOperationException("value-must-be-a-string-number-boolean-or-null");
    }

    private static string RequirePath(JsonObject root, string property)
    {
        if (root[property] is not JsonValue value || !value.TryGetValue<string>(out var path) || string.IsNullOrWhiteSpace(path))
            throw new InvalidOperationException($"{property}-is-required");
        return Path.GetFullPath(path);
    }

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
        return new Artifact(Path.GetFullPath(path), System.Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(), stream.Length);
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
