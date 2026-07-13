using System.Globalization;
using System.Text;
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
        var builder = new StringBuilder();
        WriteCanonical(builder, value);
        return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true)
            .GetBytes(builder.ToString());
    }

    private static void WriteCanonical(StringBuilder writer, JsonElement value)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                var properties = value.EnumerateObject().ToArray();
                if (properties.Select(property => property.Name).Distinct(StringComparer.Ordinal).Count() != properties.Length)
                {
                    throw new InvalidOperationException("Canonical JSON objects cannot contain duplicate keys.");
                }
                writer.Append('{');
                var propertyIndex = 0;
                foreach (var property in properties.OrderBy(property => property.Name, StringComparer.Ordinal))
                {
                    if (propertyIndex > 0) writer.Append(',');
                    WriteJsonString(writer, property.Name);
                    writer.Append(':');
                    WriteCanonical(writer, property.Value);
                    propertyIndex += 1;
                }
                writer.Append('}');
                break;
            case JsonValueKind.Array:
                writer.Append('[');
                var itemIndex = 0;
                foreach (var item in value.EnumerateArray())
                {
                    if (itemIndex > 0) writer.Append(',');
                    WriteCanonical(writer, item);
                    itemIndex += 1;
                }
                writer.Append(']');
                break;
            case JsonValueKind.String:
                WriteJsonString(writer, value.GetString()!);
                break;
            case JsonValueKind.Number:
                var rawNumber = value.GetRawText();
                if (rawNumber.Contains('.') || rawNumber.Contains('e') || rawNumber.Contains('E')
                    || !long.TryParse(rawNumber, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var integer)
                    || integer is < -MaxSafeInteger or > MaxSafeInteger)
                {
                    throw new InvalidOperationException("Canonical JSON v1 accepts lossless lexical safe integers only; encode exact decimal values as strings.");
                }
                writer.Append(integer.ToString(CultureInfo.InvariantCulture));
                break;
            case JsonValueKind.True:
                writer.Append("true");
                break;
            case JsonValueKind.False:
                writer.Append("false");
                break;
            case JsonValueKind.Null:
                writer.Append("null");
                break;
            default:
                throw new InvalidOperationException($"Unsupported JSON value kind: {value.ValueKind}");
        }
    }

    private static void WriteJsonString(StringBuilder writer, string value)
    {
        writer.Append('"');
        foreach (var character in value)
        {
            switch (character)
            {
                case '"': writer.Append("\\\""); break;
                case '\\': writer.Append("\\\\"); break;
                case '\b': writer.Append("\\b"); break;
                case '\f': writer.Append("\\f"); break;
                case '\n': writer.Append("\\n"); break;
                case '\r': writer.Append("\\r"); break;
                case '\t': writer.Append("\\t"); break;
                default:
                    if (character < ' ')
                    {
                        writer.Append("\\u");
                        writer.Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        writer.Append(character);
                    }
                    break;
            }
        }
        writer.Append('"');
    }
}
