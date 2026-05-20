// Copyright 2025 OfficeCLI (officecli.ai)
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
// OOXML UInt16Value/ByteValue/UInt32Value/SByteValue/UInt64Value frequently
// land in DocumentNode.Format[] as boxed primitives (e.g. chart hole/skip/
// rotateX/style/firstSliceAngle). Without these JsonSerializable hooks the
// source-gen polymorphic writer throws JsonTypeInfo missing-metadata when
// `get --json` hits a node carrying any of them. See R43-6.
[JsonSerializable(typeof(ushort))]
[JsonSerializable(typeof(byte))]
[JsonSerializable(typeof(sbyte))]
[JsonSerializable(typeof(ulong))]
[JsonSerializable(typeof(float))]
[JsonSerializable(typeof(decimal))]
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
    /// Output: { "success": true|false, "data": ..., "warnings": [...] }
    ///
    /// CONTRACT: `success` reflects the *business* outcome of the command, not
    /// process liveness. Pass `success: false` when the command ran to
    /// completion but its judgment is "failed" (e.g. validate found schema
    /// errors, batch had a failed step). For *probe* commands like
    /// `view --mode issues`, success stays true even when issues are listed —
    /// listing issues is the command's normal output, not a failure verdict.
    /// See CLAUDE.md "JSON Envelope" for the per-command judgment table.
    /// </summary>
    public static string WrapEnvelope(string dataJson, List<CliWarning>? warnings = null, bool success = true)
    {
        var envelope = new JsonObject { ["success"] = success };

        // Parse and embed data as-is (preserves original structure)
        try { envelope["data"] = JsonNode.Parse(dataJson); }
        catch { envelope["data"] = dataJson; } // fallback: plain string

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    /// <summary>
    /// Wraps a plain text result (like "Updated ..." or "Added ...") into an envelope.
    /// See WrapEnvelope's CONTRACT note for `success` semantics.
    /// </summary>
    public static string WrapEnvelopeText(string message, List<CliWarning>? warnings = null, int? matched = null, bool success = true)
    {
        var envelope = new JsonObject
        {
            ["success"] = success,
            // BUG-R6-04: `add --json` previously emitted only `message`,
            // diverging from get/set/dump which surface a `data` field.
            // Keep `message` for backwards compatibility but also expose
            // it under `data` so a single parser (`.data`) works across
            // every command's --json output.
            ["data"] = message,
            ["message"] = message
        };

        if (matched.HasValue)
            envelope["matched"] = matched.Value;

        if (warnings is { Count: > 0 })
            envelope["warnings"] = JsonSerializer.SerializeToNode(warnings, AppJsonContext.Default.ListCliWarning);

        return envelope.ToJsonString(JsonOptions);
    }

    public static string WrapEnvelopeWithData(string message, DocumentNode data, List<CliWarning>? warnings = null, int? matched = null, bool success = true)
    {
        var envelope = new JsonObject
        {
            ["success"] = success,
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
        else
        {
            EnrichFromMessage(result, ex);
        }

        return result;
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

        // Pattern: "<ElementType> <N> not found" without the (total: …) tail —
        // e.g. "Paragraph 99 not found" raised by Add when the parent index
        // overshoots without the handler also reporting the total. Before
        // this the message fell through to the internal_error catch-all even
        // though semantically the same as the (total:…) variant.
        var notFoundShortMatch = System.Text.RegularExpressions.Regex.Match(msg, @"^(\w+)\s+(\d+)\s+not found$");
        if (notFoundShortMatch.Success)
        {
            result.Code = "not_found";
            return;
        }

        // Pattern: "Path not found: …" — generic path-resolve failure raised
        // by handlers when an absolute DOM path can't be walked. Surfacing
        // this as not_found instead of internal_error mirrors how every
        // other missing-element error is coded.
        if (msg.StartsWith("Path not found:", StringComparison.Ordinal))
        {
            result.Code = "not_found";
            return;
        }

        // Pattern: "Sheet not found: <name>" — xlsx-specific not_found surface.
        // Handlers throw this when a worksheet name doesn't resolve; classify
        // alongside other missing-element errors instead of internal_error.
        if (msg.StartsWith("Sheet not found:", StringComparison.Ordinal))
        {
            result.Code = "not_found";
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

        // Pattern: "Row <N> in cell reference '...' is out of valid range. …" /
        // "Column '<X>' in cell reference '...' is out of range. …" —
        // raised by ParseCellReference (ExcelHandler.Selector.cs) when a
        // cell address overshoots Excel's XFD1048576 ceiling. The Set
        // path already coerces row overflow into invalid_value via its
        // "Invalid row index N." text; the Add path runs through
        // ParseCellReference and previously fell through to
        // internal_error. Map both shapes to invalid_value so add/set
        // produce the same business code for the same overflow class.
        if (msg.StartsWith("Row ", StringComparison.Ordinal)
            && msg.Contains(" in cell reference ", StringComparison.Ordinal)
            && msg.Contains(" out of ", StringComparison.Ordinal))
        {
            result.Code = "invalid_value";
            return;
        }
        if (msg.StartsWith("Column '", StringComparison.Ordinal)
            && msg.Contains(" in cell reference ", StringComparison.Ordinal)
            && msg.Contains(" out of range", StringComparison.Ordinal))
        {
            result.Code = "invalid_value";
            return;
        }

        // Pattern: "Cell <ref> not found" — raised by RemoveCell when the
        // caller targets an empty/missing cell. Symmetric with the
        // existing "Path not found:" / "Sheet not found:" rules; without
        // it the message fell through to internal_error and agents had
        // no stable code to distinguish a missing-cell remove from a
        // genuine handler crash.
        if (System.Text.RegularExpressions.Regex.IsMatch(msg, @"^Cell\s+[A-Z]+\d+\s+not found"))
        {
            result.Code = "not_found";
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

        // Pattern: "Unknown <thing>: ..." — handlers throw this when a token
        // (chart type, geometry, anchor, …) doesn't match any known value.
        // Same semantic class as "Invalid <…>" — surface invalid_value.
        if (msg.StartsWith("Unknown ", StringComparison.Ordinal))
        {
            result.Code = "invalid_value";
            var validMatch = System.Text.RegularExpressions.Regex.Match(msg, @"Valid values?:\s*(.+?)\.?$");
            if (validMatch.Success)
                result.ValidValues = validMatch.Groups[1].Value.Split(", ");
            return;
        }

        // Pattern: "<Type> requires a '<prop>' property" — handler-side
        // pre-condition check that a creation/Set call is missing a required
        // property. Maps to missing_property like "X property is required".
        if (System.Text.RegularExpressions.Regex.IsMatch(msg, @"requires a '\w+' property"))
        {
            result.Code = "missing_property";
            return;
        }

        // Pattern: "<thing> already exists: <name>" — uniqueness violation
        // (duplicate sheet name, defined name, etc). Distinct from
        // invalid_value: the value is well-formed but collides with an
        // existing entity.
        if (msg.Contains("already exists", StringComparison.Ordinal))
        {
            result.Code = "duplicate_name";
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

        // Pattern: "Batch input must be a JSON array..."
        if (msg.StartsWith("Batch input must be"))
        {
            result.Code = "invalid_input";
            return;
        }

        // Pattern: System.Text.Json error like "'I' is an invalid start of a value..."
        if (ex is System.Text.Json.JsonException)
        {
            result.Code = "invalid_json";
            return;
        }

        // Pattern: "No shape found with @id=NNN" / "No <element> found with ..."
        if (System.Text.RegularExpressions.Regex.IsMatch(msg, @"^No \w+ found with "))
        {
            result.Code = "not_found";
            return;
        }

        // Pattern: System.Xml.XPath invalid expression — surfaces as
        // XPathException with message "Expression must evaluate to a node-set."
        // or similar parser-side text.
        if (ex is System.Xml.XPath.XPathException
            || msg.Contains("Expression must evaluate")
            || msg.Contains("invalid token")
            || msg.Contains("invalid XPath"))
        {
            result.Code = "invalid_xpath";
            return;
        }

        // Pattern: file-system IO denial / disk errors (UnauthorizedAccess,
        // DirectoryNotFound, generic IOException for path-level failures).
        if (ex is UnauthorizedAccessException || ex is DirectoryNotFoundException
            || ex is PathTooLongException)
        {
            result.Code = "io_error";
            return;
        }
        if (ex is IOException && !(ex is FileNotFoundException))
        {
            result.Code = "io_error";
            return;
        }

        // Final catch-all: every WrapEnvelopeError consumer expects a 'code'
        // field for stable error routing. Unhandled exceptions previously
        // produced { error: "..." } with no code, leaving agent callers to
        // string-match free-form messages. internal_error mirrors the
        // 'unknown business failure' bucket used by other envelope code paths.
        if (string.IsNullOrEmpty(result.Code))
            result.Code = "internal_error";
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
