using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Xlsx;

public static class AtomicOperationRunner
{
    private sealed record Artifact(string Path, string Sha256, long Bytes);

    private static readonly HashSet<string> SupportedOperationTypes = new(StringComparer.Ordinal)
    {
        "setCellValue",
        "setCellNumberFormat",
        "setRichTextCellValue",
        "setRangeValues",
        "insertRows",
        "deleteRows",
        "copyRow",
        "copyCellFormula",
        "expandSectionRows",
        "setPrintArea",
        "setPageSetup",
        "setRowPageBreaks",
        "setColumnWidth",
    };

    public static int Run(string[] args)
    {
        if (args.Length != 1)
            throw new InvalidOperationException("xlsx_apply_operations requires <request.json>");

        string? receiptOutput = null;
        Artifact? inputArtifact = null;
        Artifact? operationsArtifact = null;
        IDisposable? writeLease = null;
        try
        {
            var request = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
                ?? throw new InvalidOperationException("xlsx-atomic-operation-request-invalid");
            var input = RequirePath(request, "input");
            var operationsInput = RequirePath(request, "operationsInput");
            var output = RequirePath(request, "output");
            receiptOutput = RequirePath(request, "receiptOutput");
            writeLease = Tiwater.Office.OutputWriteLease.Acquire(output, receiptOutput);
            var inPlace = PathsEqual(input, output);
            if (!inPlace) RequireNewPath(output, "output");
            RequireNewPath(receiptOutput, "receiptOutput");
            if (PathsEqual(output, receiptOutput) || PathsEqual(operationsInput, output) || PathsEqual(operationsInput, receiptOutput))
                throw new InvalidOperationException("xlsx-atomic-operation-paths-must-be-distinct");

            var handoff = JsonNode.Parse(File.ReadAllText(operationsInput)) as JsonObject
                ?? throw new InvalidOperationException("xlsx-atomic-operation-handoff-invalid");
            if (handoff.Count != 1 || handoff["operations"] is not JsonArray operationNodes || operationNodes.Count == 0)
                throw new InvalidOperationException("xlsx-atomic-operation-handoff-must-contain-only-nonempty-operations");
            var operations = operationNodes.Select(node =>
            {
                var operation = node?.Deserialize<XlsxEditOperation>(Json.Options)
                    ?? throw new InvalidOperationException("xlsx-atomic-operation-invalid");
                if (!SupportedOperationTypes.Contains(operation.Type))
                    throw new InvalidOperationException($"xlsx-atomic-operation-type-unsupported:{operation.Type}");
                return operation;
            }).ToArray();

            inputArtifact = Describe(input);
            operationsArtifact = Describe(operationsInput);
            var editResult = Editor.Apply(input, output, operations);
            var applied = editResult.AppliedOperations.Select((operation, index) => new
            {
                index,
                type = operations[index].Type,
                applied = operation.Applied,
                detail = operation.Detail,
                errorCode = operation.ErrorCode,
            }).ToArray();
            var pass = editResult.AppliedOperations.Count == operations.Length
                && editResult.AppliedOperations.All(operation => operation.Applied)
                && File.Exists(output);
            var outputArtifact = pass ? Describe(output) : null;
            var receipt = WriteJsonArtifact(receiptOutput, new
            {
                schema = "tiwater.office.xlsx-atomic-operation-receipt/v1",
                tool = "xlsx_apply_operations",
                pass,
                input = inputArtifact,
                operationsInput = operationsArtifact,
                acceptedCall = request,
                output = outputArtifact,
                operationCount = operations.Length,
                appliedOperations = applied,
            });
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                tool = "xlsx_apply_operations",
                receipt,
                output = outputArtifact,
                summary = new
                {
                    pass,
                    operationCount = operations.Length,
                    appliedCount = applied.Count(operation => operation.applied),
                },
            }, Json.Options));
            return pass ? 0 : 1;
        }
        catch (Exception error)
        {
            if (receiptOutput is not null && !File.Exists(receiptOutput))
            {
                try
                {
                    var receipt = WriteJsonArtifact(receiptOutput, new
                    {
                        schema = "tiwater.office.xlsx-atomic-operation-receipt/v1",
                        tool = "xlsx_apply_operations",
                        pass = false,
                        input = inputArtifact,
                        operationsInput = operationsArtifact,
                        output = (Artifact?)null,
                        error = error.Message,
                    });
                    Console.WriteLine(JsonSerializer.Serialize(new
                    {
                        tool = "xlsx_apply_operations",
                        receipt,
                        output = (Artifact?)null,
                        summary = new { pass = false, operationCount = 0, appliedCount = 0 },
                    }, Json.Options));
                    return 1;
                }
                catch
                {
                }
            }
            Console.Error.WriteLine(error.Message);
            return 1;
        }
        finally
        {
            writeLease?.Dispose();
        }
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
