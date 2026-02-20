using System;
using System.Collections.Generic;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace System.Web.Script.Serialization;

/// <summary>
/// Minimal shim for JavaScriptSerializer to keep legacy code intact on .NET 8.
/// Uses System.Text.Json and returns Dictionary/List primitives compatible with legacy payload handling.
/// </summary>
public class JavaScriptSerializer
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        NumberHandling = JsonNumberHandling.AllowReadingFromString,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public T? Deserialize<T>(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return default;
        }

        if (typeof(T) == typeof(Dictionary<string, object>))
        {
            JsonNode? root = JsonNode.Parse(input);
            if (root == null)
            {
                return default;
            }

            var converted = ConvertNode(root) as Dictionary<string, object>;
            if (converted == null)
            {
                return default;
            }

            return (T)(object)converted;
        }

        return JsonSerializer.Deserialize<T>(input, Options);
    }

    public string Serialize(object obj) => JsonSerializer.Serialize(obj, Options);

    private static object? ConvertNode(JsonNode? node)
    {
        if (node == null)
        {
            return null;
        }

        if (node is JsonValue value)
        {
            return ConvertValue(value);
        }

        if (node is JsonObject jsonObject)
        {
            var dict = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach ((string key, JsonNode? child) in jsonObject)
            {
                dict[key] = ConvertNode(child)!;
            }
            return dict;
        }

        if (node is JsonArray jsonArray)
        {
            var list = new List<object?>();
            foreach (JsonNode? child in jsonArray)
            {
                list.Add(ConvertNode(child));
            }
            return list;
        }

        return node.ToJsonString();
    }

    private static object? ConvertValue(JsonValue value)
    {
        if (value.TryGetValue(out JsonElement element))
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number when element.TryGetInt64(out long l) => l,
                JsonValueKind.Number when element.TryGetDouble(out double d) => d,
                JsonValueKind.Number => element.GetRawText(),
                JsonValueKind.True => true,
                JsonValueKind.False => false,
                JsonValueKind.Null or JsonValueKind.Undefined => null,
                _ => element.GetRawText()
            };
        }

        if (value.TryGetValue(out string? s)) return s;
        if (value.TryGetValue(out bool b)) return b;
        if (value.TryGetValue(out double n)) return n;

        return value.ToJsonString();
    }
}
