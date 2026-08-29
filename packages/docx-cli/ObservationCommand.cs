using System.Text.Json;
using System.Text.Json.Nodes;

namespace Dockit.Docx;

public static class ObservationCommand
{
    private static readonly IReadOnlySet<string> CommandSet = new HashSet<string>(StringComparer.Ordinal)
    {
        "docx_list_objects",
        "docx_find_literal",
        "docx_read_object",
        "docx_copy_table_range",
    };

    public static IReadOnlyCollection<string> Commands => CommandSet.ToArray();
    public static bool IsCommand(string command) => CommandSet.Contains(command);

    public static int Run(string command, string[] args)
    {
        if (!IsCommand(command)) throw new InvalidOperationException($"Unknown DOCX observation command: {command}");
        if (args.Length != 1) throw new InvalidOperationException($"{command} requires <request.json>");
        if (command == "docx_copy_table_range") return TableRangeCopy.Run(args);

        var request = JsonNode.Parse(File.ReadAllText(args[0])) as JsonObject
            ?? throw new InvalidOperationException("docx-observation-request-invalid");
        var input = RequireString(request, "input");
        object result = command switch
        {
            "docx_list_objects" => Observation.List(
                input,
                RequireString(request, "kind"),
                OptionalString(request, "scope"),
                OptionalString(request, "parentRef"),
                OptionalInt(request, "limit") ?? Observation.DefaultPageLimit,
                OptionalString(request, "continuation")),
            "docx_find_literal" => Observation.Find(
                input,
                RequireString(request, "literal"),
                OptionalString(request, "kind"),
                OptionalString(request, "scope"),
                OptionalString(request, "parentRef"),
                OptionalInt(request, "limit") ?? Observation.DefaultPageLimit,
                OptionalString(request, "continuation")),
            "docx_read_object" => Observation.Read(
                input,
                RequireString(request, "ref"),
                OptionalString(request, "revision")),
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
}
