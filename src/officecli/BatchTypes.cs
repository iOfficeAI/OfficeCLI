// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeCli;

internal class LenientStringDictionaryConverter : JsonConverter<Dictionary<string, string>>
{
    public override Dictionary<string, string>? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType == JsonTokenType.Null) return null;
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Array form: ["key=value", ...]. This mirrors the single-command MCP
        // `props` argument and the CLI `--prop key=value` flag, so an agent that
        // learned props from `set`/`add` produces the same shape inside a batch
        // item. Before this, batch props was object-only and every array-form
        // batch failed with "Expected object for props" — observed as a 100%
        // batch-failure for models that (correctly) reused the single-command
        // props shape. Lenient split on the first '=' matches McpServer.ParseProps.
        if (reader.TokenType == JsonTokenType.StartArray)
        {
            while (reader.Read())
            {
                if (reader.TokenType == JsonTokenType.EndArray) return dict;
                if (reader.TokenType != JsonTokenType.String)
                    throw new JsonException("Expected \"key=value\" string in props array");
                var kv = reader.GetString()!;
                var eq = kv.IndexOf('=');
                if (eq > 0) dict[kv[..eq]] = kv[(eq + 1)..];  // skip malformed, as ParseProps does
            }
            throw new JsonException("Unexpected end of JSON");
        }
        if (reader.TokenType != JsonTokenType.StartObject)
            throw new JsonException("Expected object or [\"key=value\"] array for props");
        while (reader.Read())
        {
            if (reader.TokenType == JsonTokenType.EndObject) return dict;
            if (reader.TokenType != JsonTokenType.PropertyName)
                throw new JsonException("Expected property name");
            var key = reader.GetString()!;
            reader.Read();
            var value = reader.TokenType switch
            {
                JsonTokenType.String => reader.GetString()!,
                JsonTokenType.Number => reader.TryGetInt64(out var l) ? l.ToString() : reader.GetDouble().ToString(),
                JsonTokenType.True => "true",
                JsonTokenType.False => "false",
                JsonTokenType.Null => "",
                _ => throw new JsonException($"Unexpected token {reader.TokenType} for prop value '{key}'")
            };
            dict[key] = value;
        }
        throw new JsonException("Unexpected end of JSON");
    }

    public override void Write(Utf8JsonWriter writer, Dictionary<string, string> value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        foreach (var kv in value)
            writer.WriteString(kv.Key, kv.Value);
        writer.WriteEndObject();
    }
}

internal class BatchItemConverter : JsonConverter<BatchItem>
{
    private static readonly LenientStringDictionaryConverter PropsConverter = new();

    public override BatchItem? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType != JsonTokenType.StartObject)
            throw new JsonException("Expected StartObject for BatchItem");

        var item = new BatchItem();
        while (reader.Read())
        {
            if (reader.TokenType == JsonTokenType.EndObject) return item;
            if (reader.TokenType != JsonTokenType.PropertyName)
                throw new JsonException("Expected PropertyName");
            var prop = reader.GetString()!;
            reader.Read();
            switch (prop.ToLowerInvariant())
            {
                case "command":
                case "op":
                    item.Command = reader.GetString() ?? "";
                    break;
                case "path": item.Path = reader.GetString(); break;
                case "parent": item.Parent = reader.GetString(); break;
                case "type": item.Type = reader.GetString(); break;
                case "from": item.From = reader.GetString(); break;
                case "index": item.Index = reader.TokenType == JsonTokenType.Null ? null : reader.GetInt32(); break;
                case "after": item.After = reader.GetString(); break;
                case "before": item.Before = reader.GetString(); break;
                case "to": item.To = reader.GetString(); break;
                case "path2": item.Path2 = reader.GetString(); break;
                case "props": item.Props = PropsConverter.Read(ref reader, typeof(Dictionary<string, string>), options); break;
                case "selector": item.Selector = reader.GetString(); break;
                case "text": item.Text = reader.GetString(); break;
                case "mode": item.Mode = reader.GetString(); break;
                case "depth": item.Depth = reader.TokenType == JsonTokenType.Null ? null : reader.GetInt32(); break;
                case "part": item.Part = reader.GetString(); break;
                case "xpath": item.Xpath = reader.GetString(); break;
                case "action": item.Action = reader.GetString(); break;
                case "xml": item.Xml = reader.GetString(); break;
                case "dumpversion": item.DumpVersion = reader.TokenType == JsonTokenType.Null ? null : reader.GetInt32(); break;
                default: reader.Skip(); break;
            }
        }
        throw new JsonException("Unexpected end of JSON for BatchItem");
    }

    public override void Write(Utf8JsonWriter writer, BatchItem value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        if (!string.IsNullOrEmpty(value.Command)) writer.WriteString("command", value.Command);
        if (value.Path != null) writer.WriteString("path", value.Path);
        if (value.Parent != null) writer.WriteString("parent", value.Parent);
        if (value.Type != null) writer.WriteString("type", value.Type);
        if (value.From != null) writer.WriteString("from", value.From);
        if (value.Index.HasValue) writer.WriteNumber("index", value.Index.Value);
        if (value.After != null) writer.WriteString("after", value.After);
        if (value.Before != null) writer.WriteString("before", value.Before);
        if (value.To != null) writer.WriteString("to", value.To);
        if (value.Path2 != null) writer.WriteString("path2", value.Path2);
        if (value.Props != null) { writer.WritePropertyName("props"); PropsConverter.Write(writer, value.Props, options); }
        if (value.Selector != null) writer.WriteString("selector", value.Selector);
        if (value.Text != null) writer.WriteString("text", value.Text);
        if (value.Mode != null) writer.WriteString("mode", value.Mode);
        if (value.Depth.HasValue) writer.WriteNumber("depth", value.Depth.Value);
        if (value.Part != null) writer.WriteString("part", value.Part);
        if (value.Xpath != null) writer.WriteString("xpath", value.Xpath);
        if (value.Action != null) writer.WriteString("action", value.Action);
        if (value.Xml != null) writer.WriteString("xml", value.Xml);
        if (value.DumpVersion.HasValue) writer.WriteNumber("dumpVersion", value.DumpVersion.Value);
        writer.WriteEndObject();
    }
}

[JsonConverter(typeof(BatchItemConverter))]
public class BatchItem
{
    public string Command { get; set; } = "";
    public string? Path { get; set; }
    public string? Parent { get; set; }
    public string? Type { get; set; }
    public string? From { get; set; }
    public int? Index { get; set; }
    public string? After { get; set; }
    public string? Before { get; set; }
    public string? To { get; set; }
    // swap's second element. Canonical name across the single-command MCP tool
    // and the CLI (`swap path1 path2`); a batch swap that reused that name was
    // previously dropped (BatchItem had no path2), so swap only worked via the
    // off-name `to`. Accept both — see the swap case in ExecuteBatchItem.
    public string? Path2 { get; set; }
    public Dictionary<string, string>? Props { get; set; }
    public string? Selector { get; set; }
    public string? Text { get; set; }
    public string? Mode { get; set; }
    public int? Depth { get; set; }
    public string? Part { get; set; }
    public string? Xpath { get; set; }
    public string? Action { get; set; }
    public string? Xml { get; set; }
    // NEWLINE-SEMANTICS-V2: dumps are versioned via a leading
    // {"command":"meta","dumpVersion":2} item. v2 encodes soft line breaks
    // as '\v' in text props ('\n' means a paragraph boundary); dumps
    // WITHOUT a meta item are legacy v1, where '\n' meant a soft break —
    // BatchCompat rewrites those on replay so old dump files keep restoring
    // the exact original structure.
    public int? DumpVersion { get; set; }

    internal static readonly HashSet<string> KnownFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "command", "op", "path", "parent", "type", "from", "index", "after", "before", "to", "path2",
        "props", "selector", "text", "mode", "depth", "part", "xpath", "action", "xml", "dumpversion"
    };

    public ResidentRequest ToResidentRequest()
    {
        var req = new ResidentRequest { Command = Command };

        if (Path != null) req.Args["path"] = Path;
        if (Parent != null) req.Args["parent"] = Parent;
        if (Type != null) req.Args["type"] = Type;
        if (From != null) req.Args["from"] = From;
        if (Index.HasValue) req.Args["index"] = Index.Value.ToString();
        if (After != null) req.Args["after"] = After;
        if (Before != null) req.Args["before"] = Before;
        if (To != null) req.Args["to"] = To;
        if (Path2 != null) req.Args["path2"] = Path2;
        if (Selector != null) req.Args["selector"] = Selector;
        if (Text != null) req.Args["text"] = Text;
        if (Mode != null) req.Args["mode"] = Mode;
        if (Depth.HasValue) req.Args["depth"] = Depth.Value.ToString();
        if (Part != null) req.Args["part"] = Part;
        if (Xpath != null) req.Args["xpath"] = Xpath;
        if (Action != null) req.Args["action"] = Action;
        if (Xml != null) req.Args["xml"] = Xml;

        if (Props != null)
            req.Props = Props;

        return req;
    }
}

[JsonConverter(typeof(BatchResultConverter))]
public class BatchResult
{
    public int Index { get; set; }
    public bool Success { get; set; }
    public string? Output { get; set; }
    public string? Error { get; set; }
    /// <summary>
    /// Machine-readable error code for a failed item (same closed list as the
    /// envelope-level error.code). Null when the failure is unclassified —
    /// consumers fall back to the Error text. Purely additive field.
    /// </summary>
    public string? Code { get; set; }
    /// <summary>The original batch item, included when the command fails so the agent can inspect/retry.</summary>
    public BatchItem? Item { get; set; }
}

/// <summary>
/// Custom converter for BatchResult that writes Output as raw JSON (not double-encoded)
/// when the Output string is valid JSON.
/// </summary>
internal class BatchResultConverter : JsonConverter<BatchResult>
{
    public override BatchResult? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        using var doc = JsonDocument.ParseValue(ref reader);
        var root = doc.RootElement;
        var result = new BatchResult();
        if (root.TryGetProperty("index", out var idx)) result.Index = idx.GetInt32();
        if (root.TryGetProperty("success", out var suc)) result.Success = suc.GetBoolean();
        if (root.TryGetProperty("output", out var outp)) result.Output = outp.ValueKind == JsonValueKind.String ? outp.GetString() : outp.GetRawText();
        if (root.TryGetProperty("error", out var err)) result.Error = err.GetString();
        if (root.TryGetProperty("code", out var cod)) result.Code = cod.GetString();
        if (root.TryGetProperty("item", out var itm)) result.Item = JsonSerializer.Deserialize(itm.GetRawText(), BatchJsonContext.Default.BatchItem);
        return result;
    }

    public override void Write(Utf8JsonWriter writer, BatchResult value, JsonSerializerOptions options)
    {
        writer.WriteStartObject();
        writer.WriteNumber("index", value.Index);
        writer.WriteBoolean("success", value.Success);
        if (value.Output != null)
        {
            // If Output is valid JSON (object or array), write it as raw JSON to avoid double-encoding
            if (IsJsonObjectOrArray(value.Output))
            {
                writer.WritePropertyName("output");
                using var doc = JsonDocument.Parse(value.Output);
                doc.RootElement.WriteTo(writer);
            }
            else
            {
                writer.WriteString("output", value.Output);
            }
        }
        if (value.Error != null)
        {
            writer.WriteString("error", value.Error);
            if (value.Code != null)
                writer.WriteString("code", value.Code);
            if (value.Item != null)
            {
                writer.WritePropertyName("item");
                JsonSerializer.Serialize(writer, value.Item, BatchJsonContext.Default.BatchItem);
            }
        }
        writer.WriteEndObject();
    }

    private static bool IsJsonObjectOrArray(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return false;
        var trimmed = s.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed[0] != '{' && trimmed[0] != '[') return false;
        try
        {
            using var doc = JsonDocument.Parse(s);
            return doc.RootElement.ValueKind is JsonValueKind.Object or JsonValueKind.Array;
        }
        catch { return false; }
    }
}

[JsonSourceGenerationOptions]
[JsonSerializable(typeof(BatchItem))]
[JsonSerializable(typeof(List<BatchItem>))]
[JsonSerializable(typeof(List<BatchResult>))]
internal partial class BatchJsonContext : JsonSerializerContext;
