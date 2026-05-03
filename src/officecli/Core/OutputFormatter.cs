// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace OfficeCli.Core;

internal enum OutputFormat
{
    Text,
    Json
}

internal class ViewResult
{
    [JsonPropertyName("view")]
    public string View { get; set; } = "";
    [JsonPropertyName("content")]
    public string Content { get; set; } = "";
}

internal class NodesResult
{
    [JsonPropertyName("matches")]
    public int Matches { get; set; }
    [JsonPropertyName("results")]
    public List<DocumentNode> Results { get; set; } = new();
}

internal class IssuesResult
{
    [JsonPropertyName("count")]
    public int Count { get; set; }
    [JsonPropertyName("issues")]
    public List<DocumentIssue> Issues { get; set; } = new();
}

internal class ErrorResult
{
    [JsonPropertyName("error")]
    public string Error { get; set; } = "";
    [JsonPropertyName("code")]
    public string? Code { get; set; }
    [JsonPropertyName("suggestion")]
    public string? Suggestion { get; set; }
    [JsonPropertyName("help")]
    public string? Help { get; set; }
    [JsonPropertyName("validValues")]
    public string[]? ValidValues { get; set; }
    [JsonPropertyName("format")]
    public string? Format { get; set; }
    [JsonPropertyName("operation")]
    public string? Operation { get; set; }
    [JsonPropertyName("engine")]
    public string? Engine { get; set; }
    [JsonPropertyName("engineMode")]
    public string? EngineMode { get; set; }
    [JsonPropertyName("nextCommand")]
    public string? NextCommand { get; set; }
}

internal class CliWarning
{
    [JsonPropertyName("message")]
    public string Message { get; set; } = "";
    [JsonPropertyName("code")]
    public string? Code { get; set; }
    [JsonPropertyName("suggestion")]
    public string? Suggestion { get; set; }
}

/// <summary>
/// Thread-static context for capturing warnings during command execution in JSON mode.
/// </summary>
internal static class WarningContext
{
    [ThreadStatic]
    private static List<CliWarning>? _warnings;

    public static void Begin() => _warnings = new List<CliWarning>();

    public static void Add(string message, string? code = null, string? suggestion = null)
    {
        _warnings?.Add(new CliWarning { Message = message, Code = code, Suggestion = suggestion });
    }

    public static List<CliWarning>? End()
    {
        var result = _warnings;
        _warnings = null;
        return result?.Count > 0 ? result : null;
    }

    public static bool IsActive => _warnings != null;
}

[JsonSourceGenerationOptions(
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(ViewResult))]
[JsonSerializable(typeof(NodesResult))]
[JsonSerializable(typeof(IssuesResult))]
[JsonSerializable(typeof(ErrorResult))]
[JsonSerializable(typeof(CliWarning))]
[JsonSerializable(typeof(List<CliWarning>))]
[JsonSerializable(typeof(string[]))]
[JsonSerializable(typeof(DocumentNode))]
[JsonSerializable(typeof(List<DocumentNode>))]
[JsonSerializable(typeof(List<DocumentIssue>))]
[JsonSerializable(typeof(Dictionary<string, object?>))]
[JsonSerializable(typeof(List<Dictionary<string, object?>>))]
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(int))]
[JsonSerializable(typeof(long))]
[JsonSerializable(typeof(short))]
[JsonSerializable(typeof(uint))]
[JsonSerializable(typeof(double))]
[JsonSerializable(typeof(string))]
internal partial class AppJsonContext : JsonSerializerContext;

internal static class OutputFormatter
{
    public static readonly JsonSerializerOptions PublicJsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        TypeInfoResolver = AppJsonContext.Default
    };

    /// <summary>
    /// Wraps pre-serialized data JSON into a unified envelope with optional warnings.
    /// Output: { "success": true, "data": ..., "warnings": [...] }
    /// </summary>
    public static string WrapEnvelope(string dataJson, List<CliWarning>? warnings = null)
    {
        var envelope = new JsonObject { ["success"] = true };

        // Parse and embed data as-is (preserves original structure)
        try { envelope["data"] = JsonNode.Parse(dataJson); }
        catch { envelope["data"] = dataJson; } // fallback: plain string

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps a plain text result (like "Updated ..." or "Added ...") into an envelope.
    /// </summary>
    public static string WrapEnvelopeText(string message, List<CliWarning>? warnings = null, int? matched = null)
    {
        var envelope = new JsonObject
        {
            ["success"] = true,
            ["message"] = message
        };

        if (matched.HasValue)
            envelope["matched"] = matched.Value;

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    public static string WrapEnvelopeWithData(string message, DocumentNode data, List<CliWarning>? warnings = null, int? matched = null)
    {
        var envelope = new JsonObject
        {
            ["success"] = true,
            ["message"] = message,
            ["data"] = JsonSerializer.SerializeToNode(data, AppJsonContext.Default.DocumentNode)
        };

        if (matched.HasValue)
            envelope["matched"] = matched.Value;

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps a failed text result (e.g. all properties unsupported) into an envelope.
    /// Output: { "success": false, "message": "...", "warnings": [...] }
    /// </summary>
    public static string WrapEnvelopeError(string message, List<CliWarning>? warnings = null)
    {
        var envelope = new JsonObject
        {
            ["success"] = false,
            ["message"] = message
        };

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps an error into an envelope.
    /// Output: { "success": false, "error": { ... } }
    /// </summary>
    public static string WrapErrorEnvelope(Exception ex)
    {
        var errorResult = BuildErrorResult(ex);
        var envelope = new JsonObject
        {
            ["success"] = false,
            ["error"] = JsonSerializer.SerializeToNode(errorResult, AppJsonContext.Default.ErrorResult)
        };
        return envelope.ToJsonString(JsonOptions);
    }

    public static string FormatError(Exception ex)
    {
        return JsonSerializer.Serialize(BuildErrorResult(ex), AppJsonContext.Default.ErrorResult);
    }

    private static ErrorResult BuildErrorResult(Exception ex)
    {
        var result = new ErrorResult { Error = ex.Message };

        if (ex is CliException cli)
        {
            result.Code = cli.Code;
            result.Suggestion = cli.Suggestion;
            result.Help = cli.Help;
            result.ValidValues = cli.ValidValues;
        }
        else if (ex is OfficeCli.Handlers.Hwp.HwpEngineException hwp)
        {
            result.Code = hwp.Error.Code;
            result.Suggestion = hwp.Error.Suggestion;
            result.Help = "officecli help hwp";
            result.ValidValues = hwp.Error.ValidValues;
            result.Format = hwp.Error.Format;
            result.Operation = hwp.Error.Operation;
            result.Engine = hwp.Error.Engine;
            result.EngineMode = hwp.Error.EngineMode;
            result.NextCommand = BuildHwpNextCommand(hwp.Error.Code, hwp.Error.Operation);
        }
        else
        {
            EnrichFromMessage(result, ex);
        }

        return result;
    }

    private static string BuildHwpNextCommand(string? code, string? operation)
    {
        if (code is "bridge_not_enabled" or "bridge_missing")
            return "officecli hwp doctor --json";

        if (operation is not null)
            return "officecli capabilities --json";

        return "officecli help hwp";
    }

    private static void EnrichFromMessage(ErrorResult result, Exception ex)
    {
        var msg = ex.Message;

        // Pattern: "Slide 50 not found (total: 8)" → code=not_found, suggestion about valid range
        var notFoundMatch = System.Text.RegularExpressions.Regex.Match(msg, @"^(\w+)\s+(\d+)\s+not found \(total:\s*(\d+)\)");
        if (notFoundMatch.Success)
        {
            var elementType = notFoundMatch.Groups[1].Value;
            var total = int.Parse(notFoundMatch.Groups[3].Value);
            result.Code = "not_found";
            result.Suggestion = total == 0
                ? $"No {elementType} elements exist. Add one first."
                : $"Valid {elementType} index range: 1-{total}";
            return;
        }

        // Pattern: "Unknown part: X. Available: ..."
        var unknownPartMatch = System.Text.RegularExpressions.Regex.Match(msg, @"Unknown part: (.+?)\. Available: (.+)");
        if (unknownPartMatch.Success)
        {
            result.Code = "invalid_path";
            result.ValidValues = unknownPartMatch.Groups[2].Value.Split(", ");
            return;
        }

        // Pattern: "Unsupported file type: .xyz. Supported: ..."
        if (msg.Contains("Unsupported file type"))
        {
            result.Code = "unsupported_type";
            return;
        }

        // Pattern: "Invalid font size: ..." / "Invalid color value: ..." / "Invalid ... value"
        if (msg.StartsWith("Invalid "))
        {
            result.Code = "invalid_value";
            // Extract "Valid values: ..." if present
            var validMatch = System.Text.RegularExpressions.Regex.Match(msg, @"Valid values?:\s*(.+?)\.?$");
            if (validMatch.Success)
                result.ValidValues = validMatch.Groups[1].Value.Split(", ");
            return;
        }

        // Pattern: "UNSUPPORTED props: ..."
        if (msg.StartsWith("UNSUPPORTED props:"))
        {
            result.Code = "unsupported_property";
            result.Help = "officecli help <format>-set";
            return;
        }

        // Pattern: "'X' property is required for Y type"
        if (msg.Contains("property is required"))
        {
            result.Code = "missing_property";
            return;
        }

        // Pattern: "File not found: ..."
        if (ex is FileNotFoundException)
        {
            result.Code = "file_not_found";
            return;
        }
    }

    public static string FormatView(string view, string content, OutputFormat format)
    {
        return format switch
        {
            OutputFormat.Json => JsonSerializer.Serialize(new ViewResult { View = view, Content = content }, AppJsonContext.Default.ViewResult),
            _ => content
        };
    }

    public static string FormatNode(DocumentNode node, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(node, AppJsonContext.Default.DocumentNode);

        return FormatNodeAsText(node);
    }

    public static string FormatNodes(List<DocumentNode> nodes, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(new NodesResult { Matches = nodes.Count, Results = nodes }, AppJsonContext.Default.NodesResult);

        var sb = new StringBuilder();
        foreach (var node in nodes)
            sb.AppendLine(FormatNodeOneline(node));
        return sb.ToString().TrimEnd();
    }

    public static string FormatIssues(List<DocumentIssue> issues, OutputFormat format)
    {
        if (format == OutputFormat.Json)
            return JsonSerializer.Serialize(new IssuesResult { Count = issues.Count, Issues = issues }, AppJsonContext.Default.IssuesResult);

        var sb = new StringBuilder();
        sb.AppendLine($"Found {issues.Count} issue(s):");
        sb.AppendLine();

        var grouped = issues.GroupBy(i => i.Type);
        foreach (var group in grouped)
        {
            var typeName = group.Key switch
            {
                IssueType.Format => "Format Issues",
                IssueType.Content => "Content Issues",
                IssueType.Structure => "Structure Issues",
                _ => "Other"
            };
            sb.AppendLine($"{typeName} ({group.Count()}):");

            foreach (var issue in group)
            {
                var severity = issue.Severity switch
                {
                    IssueSeverity.Error => "ERROR",
                    IssueSeverity.Warning => "WARN",
                    _ => "INFO"
                };
                sb.AppendLine($"  [{issue.Id}] {issue.Path}: {issue.Message}");
                if (issue.Context != null)
                    sb.AppendLine($"       Context: \"{issue.Context}\"");
                if (issue.Suggestion != null)
                    sb.AppendLine($"       Suggestion: {issue.Suggestion}");
            }
            sb.AppendLine();
        }

        return sb.ToString().TrimEnd();
    }

    private static string FormatNodeAsText(DocumentNode node)
    {
        var sb = new StringBuilder();

        sb.AppendLine(FormatNodeOneline(node));

        foreach (var child in node.Children)
            sb.Append(FormatNodeAsText(child));

        return sb.ToString();
    }

    /// <summary>
    /// Single-line format: path (type) "text" children=N style=X key=val key=val ...
    /// Grep-friendly: every line is a complete, self-contained record.
    /// </summary>
    private static string FormatNodeOneline(DocumentNode node)
    {
        var sb = new StringBuilder();

        sb.Append($"{node.Path} ({node.Type})");
        if (node.Text != null) sb.Append($" \"{node.Text.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "").Replace("\n", "\\n")}\"");
        if (node.ChildCount > 0 && node.Children.Count == 0) sb.Append($" children={node.ChildCount}");
        if (node.Style != null) sb.Append($" style={node.Style}");

        foreach (var (key, val) in node.Format)
        {
            // style is already shown via node.Style; skip duplicate
            if (key == "style" && node.Style != null) continue;
            sb.Append($" {key}={FormatNodeValue(val)}");
        }

        return sb.ToString();
    }

    // Render a Format value for the one-line text output. Most values are
    // primitives whose ToString is already correct, but some readers store
    // structured values (e.g. paragraph `tabs` is a List<Dictionary>) and
    // those need explicit formatting — the default ToString prints
    // "System.Collections.Generic.List`1[...]" which is useless to users.
    private static string FormatNodeValue(object? val)
    {
        if (val == null) return "";
        if (val is string s) return s;
        // Lower-case bool to match the canonical-value convention
        // ("true"/"false"); .NET's default Boolean.ToString() returns
        // "True"/"False", which leaks PascalCase into Format readbacks
        // (header bold/italic, toc hyperlinks, validation flags, etc.).
        if (val is bool b) return b ? "true" : "false";
        if (val is System.Collections.IEnumerable e and not string)
        {
            var parts = new List<string>();
            foreach (var item in e)
            {
                if (item is System.Collections.IDictionary d)
                {
                    var kvs = new List<string>();
                    foreach (System.Collections.DictionaryEntry de in d)
                        kvs.Add($"{de.Key}={de.Value}");
                    parts.Add("{" + string.Join(",", kvs) + "}");
                }
                else
                {
                    parts.Add(item?.ToString() ?? "");
                }
            }
            return "[" + string.Join(",", parts) + "]";
        }
        return val.ToString() ?? "";
    }

}
