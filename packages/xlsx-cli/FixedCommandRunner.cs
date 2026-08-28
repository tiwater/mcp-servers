using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Xlsx;

public static class FixedCommandRunner
{
    private sealed record Definition(string OperationType, IReadOnlySet<string> RequiredFields, IReadOnlySet<string> AllowedFields);
    private sealed record Artifact(string Path, string Sha256, long Bytes);

    private static readonly IReadOnlyDictionary<string, Definition> Definitions =
        new Dictionary<string, Definition>(StringComparer.Ordinal)
        {
            ["xlsx_set_cell_value"] = DefinitionFor("setCellValue", ["sheet", "cell", "value"], ["sheet", "cell", "value", "valueType", "bold", "shrinkToFit", "wrapText"]),
            ["xlsx_set_cell_number_format"] = DefinitionFor("setCellNumberFormat", ["sheet", "cell", "numberFormat"], ["sheet", "cell", "numberFormat"]),
            ["xlsx_set_rich_text_cell_value"] = DefinitionFor("setRichTextCellValue", ["sheet", "cell", "value", "bold"], ["sheet", "cell", "value", "bold"]),
            ["xlsx_set_range_values"] = DefinitionFor("setRangeValues", ["sheet", "startCell", "values"], ["sheet", "startCell", "values", "valueType"]),
            ["xlsx_insert_rows"] = DefinitionFor("insertRows", ["sheet", "startRow", "count"], ["sheet", "startRow", "count", "preserveHorizontalMergedRanges", "expandAdjacentVerticalMergedRanges"]),
            ["xlsx_delete_rows"] = DefinitionFor("deleteRows", ["sheet", "startRow", "count"], ["sheet", "startRow", "count"]),
            ["xlsx_copy_row"] = DefinitionFor("copyRow", ["sheet", "sourceRow", "targetRow"], ["sheet", "sourceRow", "targetRow", "translateFormulas"]),
            ["xlsx_expand_section_rows"] = DefinitionFor("expandSectionRows", ["sheet", "anchorText", "exampleRows", "targetRows"], ["sheet", "anchorText", "exampleRows", "targetRows", "preserveStyle", "preserveFormulas", "preserveMergedRanges"]),
            ["xlsx_set_print_area"] = DefinitionFor("setPrintArea", ["sheet", "range"], ["sheet", "range"]),
            ["xlsx_set_page_setup"] = DefinitionFor("setPageSetup", ["sheet"], ["sheet", "fitToPagesWide", "fitToPagesTall", "orientation", "paperSize", "repeatRowsStart", "repeatRowsEnd", "repeatColsStart", "repeatColsEnd"]),
            ["xlsx_set_row_page_breaks"] = DefinitionFor("setRowPageBreaks", ["sheet", "breakBeforeRows"], ["sheet", "breakBeforeRows"]),
            ["xlsx_set_column_width"] = DefinitionFor("setColumnWidth", ["sheet", "column", "width"], ["sheet", "column", "width"]),
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
        var outputMayBeRemoved = false;

        try
        {
            var root = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
                ?? throw new InvalidOperationException("fixed-xlsx-request-invalid");
            RequireOnly(root, ["input", "output", "receiptOutput", "changes"]);
            var input = RequirePath(root, "input");
            output = RequirePath(root, "output");
            receiptOutput = RequirePath(root, "receiptOutput");
            RequireNewPath(output, "output");
            RequireNewPath(receiptOutput, "receiptOutput");
            if (PathsEqual(output, receiptOutput))
                throw new InvalidOperationException("output-and-receiptOutput-must-be-distinct");
            if (PathsEqual(output, input))
                throw new InvalidOperationException("output-must-not-overwrite-input");

            var changes = RequireChanges(root);
            if (changes.Count == 0)
                throw new InvalidOperationException("changes-must-contain-at-least-one-item");

            var operations = BuildOperations(definition, changes);
            inputArtifact = Describe(input);
            var editResult = Editor.Apply(input, output, operations);
            outputMayBeRemoved = File.Exists(output);
            var applied = editResult.AppliedOperations.Select((operation, index) => new
            {
                index,
                type = operation.Type,
                applied = operation.Applied,
                detail = operation.Detail,
                errorCode = operation.ErrorCode,
            }).ToArray();
            var pass = editResult.AppliedOperations.Count == operations.Count
                && editResult.AppliedOperations.All(operation => operation.Applied)
                && File.Exists(output);
            var outputArtifact = pass ? Describe(output) : null;
            if (!pass && File.Exists(output)) File.Delete(output);

            var receiptPayload = new
            {
                schema = "tiwater.xlsx.fixed-command-receipt/v1",
                provider = "tiwater-xlsx",
                tool = command,
                pass,
                input = inputArtifact,
                acceptedCall = root,
                output = outputArtifact,
                operationCount = operations.Count,
                appliedOperations = applied,
                providerResult = editResult,
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
            if (outputMayBeRemoved && output is not null && File.Exists(output)) File.Delete(output);

            if (receiptOutput is not null && !File.Exists(receiptOutput))
            {
                try
                {
                    var receiptPayload = new
                    {
                        schema = "tiwater.xlsx.fixed-command-receipt/v1",
                        provider = "tiwater-xlsx",
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

    private static Definition DefinitionFor(string operationType, string[] required, string[] allowed)
        => new(operationType, new HashSet<string>(required, StringComparer.Ordinal), new HashSet<string>(allowed, StringComparer.Ordinal));

    private static IReadOnlyList<XlsxEditOperation> BuildOperations(Definition definition, JsonArray changes)
        => changes.Select(change =>
        {
            var operation = change as JsonObject
                ?? throw new InvalidOperationException("change-must-be-an-object");
            RequireOnly(operation, definition.AllowedFields);
            foreach (var required in definition.RequiredFields)
            {
                if (!operation.ContainsKey(required) || operation[required] is null && required != "value")
                    throw new InvalidOperationException($"{required}-is-required");
            }

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

    private static void RequireOnly(JsonObject value, IEnumerable<string> allowed)
    {
        var allowedSet = allowed is IReadOnlySet<string> set
            ? set
            : new HashSet<string>(allowed, StringComparer.Ordinal);
        var unexpected = value.Select(property => property.Key).FirstOrDefault(key => !allowedSet.Contains(key));
        if (unexpected is not null)
            throw new InvalidOperationException($"unexpected-property: {unexpected}");
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
