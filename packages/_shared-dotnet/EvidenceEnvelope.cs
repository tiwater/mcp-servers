using System.Text.Encodings.Web;
using System.Text.Json;

namespace Tiwater.RuntimeContracts;

public static class EvidenceEnvelope
{
    private const long MaxSafeInteger = 9_007_199_254_740_991;

    public static ArtifactIdentity IdentifyCanonicalJson(JsonElement payload, SchemaIdentity schema)
    {
        var bytes = CanonicalJsonBytes(payload);
        return FileIdentity.IdentifyArtifactBytes(
            bytes,
            mediaType: "application/json",
            encoding: "canonical-json",
            schema);
    }

    public static byte[] CanonicalJsonBytes(JsonElement value)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            Indented = false,
        }))
        {
            WriteCanonical(writer, value);
        }
        var bytes = stream.ToArray();
        NormalizeUnicodeEscapeHex(bytes);
        return bytes;
    }

    private static void WriteCanonical(Utf8JsonWriter writer, JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                var properties = value.EnumerateObject().ToArray();
                if (properties.Select(property => property.Name).Distinct(StringComparer.Ordinal).Count() != properties.Length)
                {
                    throw new InvalidOperationException("Canonical JSON objects cannot contain duplicate keys.");
                }
                writer.WriteStartObject();
                foreach (var property in properties.OrderBy(property => property.Name, StringComparer.Ordinal))
                {
                    writer.WritePropertyName(property.Name);
                    WriteCanonical(writer, property.Value);
                }
                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (var item in value.EnumerateArray()) WriteCanonical(writer, item);
                writer.WriteEndArray();
                break;
            case JsonValueKind.String:
                writer.WriteStringValue(value.GetString());
                break;
            case JsonValueKind.Number:
                if (!value.TryGetDecimal(out var number)
                    || decimal.Truncate(number) != number
                    || number is < -MaxSafeInteger or > MaxSafeInteger)
                {
                    throw new InvalidOperationException("Canonical JSON v1 accepts cross-language safe integer numbers only; encode exact decimal values as strings.");
                }
                writer.WriteNumberValue((long)number);
                break;
            case JsonValueKind.True:
                writer.WriteBooleanValue(true);
                break;
            case JsonValueKind.False:
                writer.WriteBooleanValue(false);
                break;
            case JsonValueKind.Null:
                writer.WriteNullValue();
                break;
            default:
                throw new InvalidOperationException($"Unsupported JSON value kind: {value.ValueKind}");
        }
    }

    private static void NormalizeUnicodeEscapeHex(Span<byte> bytes)
    {
        var backslashRun = 0;
        for (var index = 0; index < bytes.Length; index += 1)
        {
            if (bytes[index] == (byte)'\\')
            {
                backslashRun += 1;
                continue;
            }

            if (bytes[index] == (byte)'u' && backslashRun % 2 == 1 && index + 4 < bytes.Length)
            {
                for (var digit = index + 1; digit <= index + 4; digit += 1)
                {
                    if (bytes[digit] is >= (byte)'A' and <= (byte)'F') bytes[digit] += 32;
                }
                index += 4;
            }

            backslashRun = 0;
        }
    }
}
