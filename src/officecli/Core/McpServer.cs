// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Minimal MCP (Model Context Protocol) server over stdio.
/// Implements JSON-RPC 2.0 with initialize, tools/list, and tools/call.
/// All JSON is hand-written via Utf8JsonWriter to avoid reflection (PublishTrimmed).
/// </summary>
public static class McpServer
{
    public static async Task RunAsync()
    {
        using var reader = new StreamReader(Console.OpenStandardInput());
        using var writer = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };

        while (true)
        {
            var line = await reader.ReadLineAsync();
            if (line == null) break;
            if (string.IsNullOrWhiteSpace(line)) continue;

            try
            {
                using var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                var method = root.TryGetProperty("method", out var m) ? m.GetString() : null;
                var id = root.TryGetProperty("id", out var idEl) ? idEl.Clone() : (JsonElement?)null;

                var response = method switch
                {
                    "initialize" => HandleInitialize(id),
                    "notifications/initialized" => null,
                    "tools/list" => HandleToolsList(id),
                    "tools/call" => HandleToolsCall(id, root),
                    "ping" => WriteJson(w => { w.WriteStartObject(); Rpc(w, id); w.WriteStartObject("result"); w.WriteEndObject(); w.WriteEndObject(); }),
                    _ => id.HasValue ? ErrorJson(id, -32601, $"Method not found: {method}") : null,
                };

                if (response != null)
                    await writer.WriteLineAsync(response);
            }
            catch (JsonException)
            {
                await writer.WriteLineAsync(ErrorJson(null, -32700, "Parse error"));
            }
            catch (Exception ex)
            {
                await writer.WriteLineAsync(ErrorJson(null, -32603, $"Internal error: {ex.Message}"));
            }
        }
    }

    // ==================== Handlers ====================

    private static string HandleInitialize(JsonElement? id) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("result");
        w.WriteString("protocolVersion", "2024-11-05");
        w.WriteStartObject("capabilities");
        w.WriteStartObject("tools"); w.WriteBoolean("listChanged", false); w.WriteEndObject();
        w.WriteEndObject();
        w.WriteStartObject("serverInfo"); w.WriteString("name", "officecli"); w.WriteString("version", "1.0.17"); w.WriteEndObject();
        w.WriteEndObject();
        w.WriteEndObject();
    });

    private static string HandleToolsList(JsonElement? id) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("result");
        w.WriteStartArray("tools");
        WriteToolDefinitions(w);
        w.WriteEndArray();
        w.WriteEndObject();
        w.WriteEndObject();
    });

    private static string HandleToolsCall(JsonElement? id, JsonElement root)
    {
        if (!root.TryGetProperty("params", out var p))
            return ErrorJson(id, -32602, "Missing params");
        var name = p.TryGetProperty("name", out var n) ? n.GetString() : null;
        var args = p.TryGetProperty("arguments", out var a) ? a : default;
        if (string.IsNullOrEmpty(name))
            return ErrorJson(id, -32602, "Missing tool name");

        try
        {
            var result = ExecuteTool(name, args);
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                w.WriteStartObject(); w.WriteString("type", "text"); w.WriteString("text", result); w.WriteEndObject();
                w.WriteEndArray();
                w.WriteBoolean("isError", false);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
        catch (Exception ex)
        {
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                w.WriteStartObject(); w.WriteString("type", "text"); w.WriteString("text", $"Error: {ex.Message}"); w.WriteEndObject();
                w.WriteEndArray();
                w.WriteBoolean("isError", true);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
    }

    // ==================== Tool Execution ====================

    private static string ExecuteTool(string name, JsonElement args)
    {
        string Arg(string key) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) ? v.GetString() ?? "" : "";
        int ArgInt(string key, int def) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : def;
        int? ArgIntOpt(string key) => args.ValueKind == JsonValueKind.Object && args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : null;
        string[] ArgStringArray(string key)
        {
            if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(key, out var v) || v.ValueKind != JsonValueKind.Array) return [];
            return v.EnumerateArray().Select(e => e.GetString() ?? "").ToArray();
        }

        switch (name)
        {
            case "create":
            {
                var file = Arg("file");
                BlankDocCreator.Create(file);
                return $"Created {file}";
            }
            case "view":
            {
                var file = Arg("file");
                var mode = Arg("mode");
                var start = ArgIntOpt("start");
                var end = ArgIntOpt("end");
                var maxLines = ArgIntOpt("max_lines");
                using var handler = DocumentHandlerFactory.Open(file);
                if (mode is "html" or "h" && handler is Handlers.PowerPointHandler pptH)
                    return pptH.ViewAsHtml(start, end);
                return mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(start, end, maxLines, null),
                    "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, null),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(null, null), OutputFormat.Json),
                    _ => throw new ArgumentException($"Unknown mode: {mode}")
                };
            }
            case "get":
            {
                var file = Arg("file");
                var path = Arg("path"); if (string.IsNullOrEmpty(path)) path = "/";
                var depth = ArgInt("depth", 1);
                using var handler = DocumentHandlerFactory.Open(file);
                var node = handler.Get(path, depth);
                return OutputFormatter.FormatNode(node, OutputFormat.Json);
            }
            case "query":
            {
                var file = Arg("file");
                var selector = Arg("selector");
                using var handler = DocumentHandlerFactory.Open(file);
                var filters = AttributeFilter.Parse(selector);
                var (results, _) = AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
                return OutputFormatter.FormatNodes(results, OutputFormat.Json);
            }
            case "set":
            {
                var file = Arg("file");
                var path = Arg("path");
                var props = ParseProps(ArgStringArray("props"));
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var unsupported = handler.Set(path, props);
                var applied = props.Where(kv => !unsupported.Contains(kv.Key)).ToList();
                var msg = applied.Count > 0
                    ? $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
                    : $"No properties applied to {path}";
                if (unsupported.Count > 0)
                    msg += $"\nUnsupported: {string.Join(", ", unsupported)}";
                return msg;
            }
            case "add":
            {
                var file = Arg("file");
                var parent = Arg("parent");
                var type = Arg("type");
                var index = ArgIntOpt("index");
                var props = ParseProps(ArgStringArray("props"));
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var resultPath = handler.Add(parent, type, index, props);
                return $"Added {type} at {resultPath}";
            }
            case "remove":
            {
                var file = Arg("file");
                var path = Arg("path");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                handler.Remove(path);
                return $"Removed {path}";
            }
            case "move":
            {
                var file = Arg("file");
                var path = Arg("path");
                var to = Arg("to"); if (string.IsNullOrEmpty(to)) to = null;
                var index = ArgIntOpt("index");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var resultPath = handler.Move(path, to, index);
                return $"Moved to {resultPath}";
            }
            case "validate":
            {
                var file = Arg("file");
                using var handler = DocumentHandlerFactory.Open(file);
                var errors = handler.Validate();
                if (errors.Count == 0) return "Validation passed: no errors found.";
                var lines = errors.Select(e => $"[{e.ErrorType}] {e.Description}" +
                    (e.Path != null ? $" (Path: {e.Path})" : ""));
                return $"Found {errors.Count} error(s):\n{string.Join("\n", lines)}";
            }
            case "batch":
            {
                var file = Arg("file");
                var commands = Arg("commands");
                var items = JsonSerializer.Deserialize<List<BatchItem>>(commands, BatchJsonContext.Default.ListBatchItem);
                if (items == null || items.Count == 0)
                    throw new ArgumentException("No commands found in input.");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var results = new List<BatchResult>();
                foreach (var item in items)
                {
                    try
                    {
                        var output = CommandBuilder.ExecuteBatchItem(handler, item, true);
                        results.Add(new BatchResult { Success = true, Output = output });
                    }
                    catch (Exception ex)
                    {
                        results.Add(new BatchResult { Success = false, Error = ex.Message });
                    }
                }
                return JsonSerializer.Serialize(results, BatchJsonContext.Default.ListBatchResult);
            }
            case "raw":
            {
                var file = Arg("file");
                var part = Arg("part"); if (string.IsNullOrEmpty(part)) part = "/document";
                using var handler = DocumentHandlerFactory.Open(file);
                return handler.Raw(part, null, null, null);
            }
            default:
                throw new ArgumentException($"Unknown tool: {name}");
        }
    }

    private static Dictionary<string, string> ParseProps(string[] propStrs)
    {
        var props = new Dictionary<string, string>();
        foreach (var p in propStrs)
        {
            var eq = p.IndexOf('=');
            if (eq > 0) props[p[..eq]] = p[(eq + 1)..];
        }
        return props;
    }

    // ==================== Tool Definitions ====================

    private static void WriteToolDefinitions(Utf8JsonWriter w)
    {
        WriteTool(w, "create", "Create a blank Office document (.docx, .xlsx, .pptx)",
            s => { s.WriteProp("file", "string", "Output file path"); },
            ["file"]);

        WriteTool(w, "view", "View document content (text, annotated, outline, stats, issues, html)",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteEnum("mode", "View mode", ["text", "annotated", "outline", "stats", "issues", "html"]);
                s.WriteProp("start", "number", "Start line number");
                s.WriteProp("end", "number", "End line number");
                s.WriteProp("max_lines", "number", "Maximum lines to output");
            }, ["file", "mode"]);

        WriteTool(w, "get", "Get a document node by DOM path with properties, text, format, and children",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WritePropDefault("path", "string", "DOM path (e.g. /body/p[1], /slide[1]/shape[2])", "/");
                s.WritePropDefault("depth", "number", "Depth of child nodes", "1");
            }, ["file"]);

        WriteTool(w, "query", "Query elements with CSS-like selectors (e.g. shape[fill=#FF0000])",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteProp("selector", "string", "CSS-like selector");
            }, ["file", "selector"]);

        WriteTool(w, "set", "Modify a document node's properties",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteProp("path", "string", "DOM path to the element");
                s.WriteArrayProp("props", "key=value pairs (e.g. bold=true, color=#FF0000)");
            }, ["file", "path", "props"]);

        WriteTool(w, "add", "Add a new element (slide, shape, paragraph, table, picture, chart, etc.)",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteProp("parent", "string", "Parent DOM path (e.g. /, /slide[1], /body)");
                s.WriteProp("type", "string", "Element type (slide, shape, paragraph, table, picture, chart, row, cell, run, etc.)");
                s.WriteArrayProp("props", "key=value pairs");
                s.WriteProp("index", "number", "Insert position (0-based)");
            }, ["file", "parent", "type"]);

        WriteTool(w, "remove", "Remove an element from the document",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteProp("path", "string", "DOM path of the element to remove");
            }, ["file", "path"]);

        WriteTool(w, "move", "Move an element to a new position or parent",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteProp("path", "string", "DOM path of the element to move");
                s.WriteProp("to", "string", "Target parent path");
                s.WriteProp("index", "number", "Insert position (0-based)");
            }, ["file", "path"]);

        WriteTool(w, "validate", "Validate document against OpenXML schema",
            s => { s.WriteProp("file", "string", "Document file path"); },
            ["file"]);

        WriteTool(w, "batch", "Execute multiple commands in one open/save cycle",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WriteProp("commands", "string", "JSON array of commands");
            }, ["file", "commands"]);

        WriteTool(w, "raw", "View raw XML of a document part",
            s => {
                s.WriteProp("file", "string", "Document file path");
                s.WritePropDefault("part", "string", "Part path (e.g. /document, /styles, /slide[1])", "/document");
            }, ["file"]);
    }

    private static void WriteTool(Utf8JsonWriter w, string name, string desc, Action<SchemaWriter> schema, string[] required)
    {
        w.WriteStartObject();
        w.WriteString("name", name);
        w.WriteString("description", desc);
        w.WriteStartObject("inputSchema");
        w.WriteString("type", "object");
        w.WriteStartObject("properties");
        schema(new SchemaWriter(w));
        w.WriteEndObject();
        w.WriteStartArray("required");
        foreach (var r in required) w.WriteStringValue(r);
        w.WriteEndArray();
        w.WriteEndObject();
        w.WriteEndObject();
    }

    private readonly ref struct SchemaWriter(Utf8JsonWriter w)
    {
        public void WriteProp(string name, string type, string desc)
        {
            w.WriteStartObject(name);
            w.WriteString("type", type);
            w.WriteString("description", desc);
            w.WriteEndObject();
        }
        public void WritePropDefault(string name, string type, string desc, string def)
        {
            w.WriteStartObject(name);
            w.WriteString("type", type);
            w.WriteString("description", desc);
            w.WriteString("default", def);
            w.WriteEndObject();
        }
        public void WriteEnum(string name, string desc, string[] values)
        {
            w.WriteStartObject(name);
            w.WriteString("type", "string");
            w.WriteString("description", desc);
            w.WriteStartArray("enum");
            foreach (var v in values) w.WriteStringValue(v);
            w.WriteEndArray();
            w.WriteEndObject();
        }
        public void WriteArrayProp(string name, string desc)
        {
            w.WriteStartObject(name);
            w.WriteString("type", "array");
            w.WriteStartObject("items"); w.WriteString("type", "string"); w.WriteEndObject();
            w.WriteString("description", desc);
            w.WriteEndObject();
        }
    }

    // ==================== JSON-RPC Helpers ====================

    private static string WriteJson(Action<Utf8JsonWriter> build)
    {
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms)) build(w);
        return Encoding.UTF8.GetString(ms.ToArray());
    }

    private static void Rpc(Utf8JsonWriter w, JsonElement? id)
    {
        w.WriteString("jsonrpc", "2.0");
        if (id.HasValue) { w.WritePropertyName("id"); id.Value.WriteTo(w); }
        else w.WriteNull("id");
    }

    private static string ErrorJson(JsonElement? id, int code, string message) => WriteJson(w =>
    {
        w.WriteStartObject();
        Rpc(w, id);
        w.WriteStartObject("error");
        w.WriteNumber("code", code);
        w.WriteString("message", message);
        w.WriteEndObject();
        w.WriteEndObject();
    });
}
