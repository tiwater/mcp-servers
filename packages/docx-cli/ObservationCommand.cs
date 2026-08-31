using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Docx;

public static class ObservationCommand
{
    private static readonly IReadOnlySet<string> CommandSet = new HashSet<string>(StringComparer.Ordinal)
    {
        "docx_list_objects",
        "docx_table_index",
        "docx_read_table",
        "docx_find_literal",
        "docx_read_object",
    };

    public static IReadOnlyCollection<string> Commands => CommandSet.ToArray();
    public static bool IsCommand(string command) => CommandSet.Contains(command);

    public static int Run(string command, string[] args)
    {
        if (!IsCommand(command)) throw new InvalidOperationException($"Unknown DOCX observation command: {command}");
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        var request = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
            ?? throw new InvalidOperationException("docx-observation-request-invalid");
        var input = RequireString(request, "input");
        object result = command switch
        {
            "docx_list_objects" => Observation.List(
                input,
                RequireStringArray(request, "kinds"),
                OptionalString(request, "scope"),
                OptionalAddress(request, "parent"),
                OptionalInt(request, "limit") ?? Observation.DefaultPageLimit,
                OptionalInt(request, "offset") ?? 0),
            "docx_table_index" => Observation.TableIndex(input),
            "docx_read_table" => Observation.ReadTable(input, RequireAddress(request, "table")),
            "docx_find_literal" => Observation.Find(
                input,
                RequireString(request, "literal"),
                OptionalString(request, "kind"),
                OptionalString(request, "scope"),
                OptionalAddress(request, "parent"),
                OptionalInt(request, "limit") ?? Observation.DefaultPageLimit,
                OptionalInt(request, "offset") ?? 0),
            "docx_read_object" => Observation.Read(
                input,
                RequireAddressList(request, "addresses"),
                RequireStringArray(request, "kinds")),
            _ => throw new InvalidOperationException($"Unknown DOCX observation command: {command}"),
        };
        Console.WriteLine(JsonSerializer.Serialize(result, Json.CamelCaseOptions));
        return 0;
    }

    private static string RequireString(JsonObject request, string property)
    {
        var value = OptionalString(request, property);
        return string.IsNullOrWhiteSpace(value)
            ? throw new InvalidOperationException($"{property}-is-required")
            : value;
    }

    private static string? OptionalString(JsonObject request, string property)
        => request[property] is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;

    private static int? OptionalInt(JsonObject request, string property)
        => request[property] is JsonValue value && value.TryGetValue<int>(out var number) ? number : null;

    private static IReadOnlySet<string> RequireStringArray(JsonObject request, string property)
    {
        if (request[property] is not JsonArray array || array.Count == 0)
            throw new InvalidOperationException($"{property}-is-required");
        var rawValues = array.Select(item => item?.GetValue<string>() ?? string.Empty).ToList();
        if (rawValues.Any(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException($"{property}-is-invalid");
        var values = rawValues.ToHashSet(StringComparer.Ordinal);
        if (rawValues.Count != values.Count)
            throw new InvalidOperationException($"{property}-contains-duplicates");
        return values;
    }

    private static DocxObjectAddress? OptionalAddress(JsonObject request, string property)
        => request[property] is null ? null : ReadAddress(request[property]!, property);

    private static DocxObjectAddress RequireAddress(JsonObject request, string property)
        => request[property] is null
            ? throw new InvalidOperationException($"{property}-is-required")
            : ReadAddress(request[property]!, property);

    private static DocxObjectAddress ReadAddress(JsonNode node, string name)
    {
        var address = node.Deserialize<DocxObjectAddress>(Json.Options)
            ?? throw new InvalidOperationException($"{name}-is-invalid");
        if (string.IsNullOrWhiteSpace(address.Part) || string.IsNullOrWhiteSpace(address.Path))
            throw new InvalidOperationException($"{name}-is-invalid");
        return address;
    }

    private static IReadOnlyList<DocxObjectAddress> RequireAddressList(JsonObject request, string property)
    {
        if (request[property] is not JsonArray array || array.Count == 0)
            throw new InvalidOperationException($"{property}-is-required");
        var values = array.Select((item, index) => item is null
            ? throw new InvalidOperationException($"{property}[{index}]-is-invalid")
            : ReadAddress(item, $"{property}[{index}]")).ToList();
        if (values.Count != values.Distinct().Count())
            throw new InvalidOperationException($"{property}-contains-duplicates");
        return values;
    }
}
