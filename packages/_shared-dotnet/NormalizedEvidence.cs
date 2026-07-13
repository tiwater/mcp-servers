using System.Text.Json;
using System.Text.Json.Nodes;

namespace Tiwater.RuntimeContracts;

public sealed record NormalizedEvidenceResult(
    JsonElement Payload,
    IReadOnlyList<EvidenceObject> Objects);

public static class NormalizedEvidence
{
    private static readonly HashSet<string> RootPathFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "file",
        "input",
        "output",
    };

    public static NormalizedEvidenceResult Build(JsonElement report)
    {
        var nodes = new JsonArray();
        var objects = new List<EvidenceObject>();
        Visit(report, "$", null, "document", 0, nodes, objects);
        var payload = new JsonObject
        {
            ["schemaVersion"] = RuntimeContractVersions.EvidenceEnvelope,
            ["nodes"] = nodes,
        };
        using var document = JsonDocument.Parse(payload.ToJsonString());
        return new NormalizedEvidenceResult(document.RootElement.Clone(), objects);
    }

    private static void Visit(
        JsonElement value,
        string runtimeNodeId,
        string? parentNodeId,
        string kind,
        int depth,
        JsonArray nodes,
        List<EvidenceObject> objects)
    {
        var valueType = ValueType(value);
        var node = new JsonObject
        {
            ["runtimeNodeId"] = runtimeNodeId,
            ["kind"] = kind,
            ["valueType"] = valueType,
            ["value"] = ScalarValue(value),
            ["locator"] = runtimeNodeId,
            ["derivedFrom"] = new JsonArray(),
            ["containedBy"] = parentNodeId,
        };
        nodes.Add(node);
        objects.Add(new EvidenceObject(
            runtimeNodeId,
            kind,
            parentNodeId is null,
            parentNodeId,
            new DerivedEvidenceIdentity("normalized-json-pointer-v1", [runtimeNodeId])));

        if (value.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in value.EnumerateObject().OrderBy(item => item.Name, StringComparer.Ordinal))
            {
                if (depth == 0 && RootPathFields.Contains(property.Name)) continue;
                var childId = $"{(runtimeNodeId == "$" ? string.Empty : runtimeNodeId)}/{Escape(property.Name)}";
                Visit(property.Value, childId, runtimeNodeId, property.Name, depth + 1, nodes, objects);
            }
        }
        else if (value.ValueKind == JsonValueKind.Array)
        {
            var itemKind = Singular(kind);
            var index = 0;
            foreach (var item in value.EnumerateArray())
            {
                Visit(item, $"{runtimeNodeId}/{index}", runtimeNodeId, itemKind, depth + 1, nodes, objects);
                index += 1;
            }
        }
    }

    private static JsonNode? ScalarValue(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => JsonValue.Create(value.GetString()),
        JsonValueKind.True => JsonValue.Create(true),
        JsonValueKind.False => JsonValue.Create(false),
        JsonValueKind.Null => null,
        JsonValueKind.Number when value.TryGetInt64(out var integer)
            && integer is >= -9_007_199_254_740_991 and <= 9_007_199_254_740_991 => JsonValue.Create(integer),
        JsonValueKind.Number => JsonValue.Create(value.GetRawText()),
        _ => null,
    };

    private static string ValueType(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Object => "object",
        JsonValueKind.Array => "array",
        JsonValueKind.String => "string",
        JsonValueKind.Number when value.TryGetInt64(out _) => "integer",
        JsonValueKind.Number => "decimal-string",
        JsonValueKind.True or JsonValueKind.False => "boolean",
        JsonValueKind.Null => "null",
        _ => throw new InvalidOperationException($"Unsupported evidence value kind: {value.ValueKind}"),
    };

    private static string Escape(string segment) => segment.Replace("~", "~0", StringComparison.Ordinal).Replace("/", "~1", StringComparison.Ordinal);

    private static string Singular(string value) => value switch
    {
        "children" => "child",
        _ when value.EndsWith("ies", StringComparison.Ordinal) => $"{value[..^3]}y",
        _ when value.EndsWith("s", StringComparison.Ordinal) && value.Length > 1 => value[..^1],
        _ => "item",
    };
}
