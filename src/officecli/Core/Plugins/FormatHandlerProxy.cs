// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// <see cref="IDocumentHandler"/> implementation that delegates every call to a
/// running format-handler plugin via <see cref="FormatHandlerSession"/>. Per
/// docs/plugin-protocol.md §2.3, this is what wraps the plugin so existing
/// get/view/query pipelines work transparently on foreign formats.
///
/// Scope: read-path (ViewAs*, Get, Query, Validate) and mutation
/// (Set/Add/Remove/Move/CopyFrom/Raw/RawSet/AddPart/TryExtractBinary)
/// are all proxied. Plugins that don't implement a given verb should
/// reply with error code <c>unsupported_command</c> per docs/plugin-protocol.md §5.3.
/// </summary>
internal sealed class FormatHandlerProxy : IDocumentHandler
{
    private readonly FormatHandlerSession _session;

    public FormatHandlerProxy(FormatHandlerSession session) { _session = session; }

    // ----- Semantic layer (text views) -----------------------------------

    public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewString("text", startLine, endLine, maxLines, cols);

    public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewString("annotated", startLine, endLine, maxLines, cols);

    public string ViewAsOutline()
        => SendViewString("outline");

    public string ViewAsStats()
        => SendViewString("stats");

    public JsonNode ViewAsStatsJson() => SendViewJson("stats");
    public JsonNode ViewAsOutlineJson() => SendViewJson("outline");
    public JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
        => SendViewJson("text", startLine, endLine, maxLines, cols);

    public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null)
    {
        var args = new JsonObject { ["mode"] = "issues" };
        if (issueType != null) args["type"] = issueType;
        if (limit.HasValue) args["limit"] = limit.Value;
        var result = _session.Send("command", "view", args);
        if (result is null) return new List<DocumentIssue>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListDocumentIssue) ?? new List<DocumentIssue>();
    }

    // ----- Query layer --------------------------------------------------

    public DocumentNode Get(string path, int depth = 1)
    {
        var result = _session.Send("command", "get", new JsonObject
        {
            ["path"] = path,
            ["depth"] = depth,
        });
        if (result is null)
            return new DocumentNode { Path = path, Type = "error", Text = "Plugin returned null result." };
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.DocumentNode)
            ?? new DocumentNode { Path = path, Type = "error", Text = "Plugin result did not deserialize as DocumentNode." };
    }

    public List<DocumentNode> Query(string selector)
    {
        var result = _session.Send("command", "query", new JsonObject { ["selector"] = selector });
        if (result is null) return new List<DocumentNode>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListDocumentNode) ?? new List<DocumentNode>();
    }

    public List<ValidationError> Validate()
    {
        var result = _session.Send("command", "validate", new JsonObject());
        if (result is null) return new List<ValidationError>();
        return JsonSerializer.Deserialize(result.ToJsonString(), PluginJsonContext.Default.ListValidationError) ?? new List<ValidationError>();
    }

    // ----- Mutation layer ----------------------------------------------

    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var args = new JsonObject { ["path"] = path };
        var props = PropsToJson(properties);
        var result = _session.Send("command", "set", args, props);
        // Protocol §5.3 (set): result is `{"unsupported_properties":["k1",...]}`.
        return ParseUnsupportedProperties(result);
    }

    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        var args = new JsonObject
        {
            ["parent_path"] = parentPath,
            ["type"] = type,
        };
        if (position is not null) args["position"] = PositionToJson(position);
        var props = PropsToJson(properties);
        var result = _session.Send("command", "add", args, props);
        // Protocol §5.3 (add): result is `{"path":"...","unsupported_properties":[...]}`.
        return ParseAddPath(result);
    }

    public string? Remove(string path)
    {
        var result = _session.Send("command", "remove", new JsonObject { ["path"] = path });
        return result?.GetValue<string>();
    }

    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        var args = new JsonObject { ["source_path"] = sourcePath };
        if (targetParentPath is not null) args["target_parent_path"] = targetParentPath;
        if (position is not null) args["position"] = PositionToJson(position);
        var result = _session.Send("command", "move", args);
        return result?.GetValue<string>() ?? "";
    }

    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        var args = new JsonObject
        {
            ["source_path"] = sourcePath,
            ["target_parent_path"] = targetParentPath,
        };
        if (position is not null) args["position"] = PositionToJson(position);
        var result = _session.Send("command", "copy", args);
        return result?.GetValue<string>() ?? "";
    }

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var args = new JsonObject { ["part_path"] = partPath };
        if (startRow.HasValue) args["start_row"] = startRow.Value;
        if (endRow.HasValue) args["end_row"] = endRow.Value;
        if (cols != null && cols.Count > 0) args["cols"] = string.Join(",", cols);
        var result = _session.Send("command", "raw", args);
        return result?.GetValue<string>() ?? "";
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var args = new JsonObject
        {
            ["part_path"] = partPath,
            ["xpath"] = xpath,
            ["action"] = action,
        };
        if (xml is not null) args["xml"] = xml;
        _session.Send("command", "raw_set", args);
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var args = new JsonObject
        {
            ["parent_part_path"] = parentPartPath,
            ["part_type"] = partType,
        };
        var props = properties is not null ? PropsToJson(properties) : null;
        var result = _session.Send("command", "add_part", args, props)?.AsObject();
        if (result is null)
            throw new CliException("Format-handler add-part returned null.") { Code = "protocol_mismatch" };
        var relId = result["rel_id"]?.GetValue<string>() ?? "";
        var partPath = result["part_path"]?.GetValue<string>() ?? "";
        return (relId, partPath);
    }

    // ----- Format-specific view extensions ------------------------------
    //
    // These are NOT on IDocumentHandler — they're entry points used by main's
    // CommandBuilder.View when a built-in handler (Word/Excel/PPT) declines.
    // `view html` and `view forms` historically downcast to a concrete handler;
    // now `else if (handler is FormatHandlerProxy proxy) ...` provides the
    // plugin-side fallback. Each method maps onto the corresponding `view`
    // command with a mode key the plugin chooses how to render.

    /// <summary>
    /// Request SVG preview from the plugin (`view mode=svg`). Returns null
    /// if the plugin replies with <c>unsupported_command</c>.
    /// </summary>
    public string? ViewAsSvg(int? page = null)
    {
        try
        {
            var args = new JsonObject { ["mode"] = "svg" };
            if (page.HasValue) args["page"] = page.Value;
            var result = _session.Send("command", "view", args);
            return result?.GetValue<string>();
        }
        catch (CliException ex) when (ex.Code == "unsupported_command")
        {
            return null;
        }
    }

    /// <summary>
    /// Request HTML preview from the plugin (`view mode=html`). Returns null
    /// if the plugin replies with <c>unsupported_command</c> — caller may then
    /// raise its own "unsupported_type" CliException.
    /// </summary>
    public string? ViewAsHtml(int? page = null)
    {
        try
        {
            var args = new JsonObject { ["mode"] = "html" };
            if (page.HasValue) args["page"] = page.Value;
            var result = _session.Send("command", "view", args);
            return result?.GetValue<string>();
        }
        catch (CliException ex) when (ex.Code == "unsupported_command")
        {
            return null;
        }
    }

    /// <summary>
    /// Request forms JSON from the plugin (`view mode=forms-json`). Returns
    /// null if the plugin replies with <c>unsupported_command</c>.
    /// </summary>
    public JsonNode? ViewAsFormsJson(bool auto = true)
    {
        try
        {
            var args = new JsonObject
            {
                ["mode"] = "forms-json",
                ["auto"] = auto,
            };
            return _session.Send("command", "view", args);
        }
        catch (CliException ex) when (ex.Code == "unsupported_command")
        {
            return null;
        }
    }

    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        try
        {
            var result = _session.Send("command", "extract_binary", new JsonObject
            {
                ["path"] = path,
                ["dest_path"] = destPath,
            })?.AsObject();
            if (result is null) return false;
            var found = result["found"]?.GetValue<bool>() ?? false;
            if (!found) return false;
            contentType = result["content_type"]?.GetValue<string>();
            // Tolerant integer parse: some plugins / language runtimes encode
            // counts as JSON doubles (`42.0`) which System.Text.Json refuses
            // to deserialize as long via the strict GetValue<long> path.
            byteCount = ParseLongTolerant(result["byte_count"]);
            return true;
        }
        catch (CliException ex) when (ex.Code == "unsupported_command")
        {
            return false;
        }
    }

    public void Dispose() => _session.Dispose();

    // ----- Helpers ------------------------------------------------------

    private string SendViewString(string mode, int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var args = BuildViewArgs(mode, startLine, endLine, maxLines, cols);
        var result = _session.Send("command", "view", args);
        return result?.GetValue<string>() ?? "";
    }

    private JsonNode SendViewJson(string mode, int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null)
    {
        var args = BuildViewArgs(mode, startLine, endLine, maxLines, cols);
        args["format"] = "json";
        var result = _session.Send("command", "view", args);
        return result ?? new JsonObject();
    }

    private static JsonObject BuildViewArgs(string mode, int? startLine, int? endLine, int? maxLines, HashSet<string>? cols)
    {
        var args = new JsonObject { ["mode"] = mode };
        if (startLine.HasValue) args["start"] = startLine.Value;
        if (endLine.HasValue) args["end"] = endLine.Value;
        if (maxLines.HasValue) args["max-lines"] = maxLines.Value;
        if (cols != null && cols.Count > 0) args["cols"] = string.Join(",", cols);
        return args;
    }

    private static JsonObject PropsToJson(Dictionary<string, string> properties)
    {
        var obj = new JsonObject();
        foreach (var kv in properties)
            obj[kv.Key] = kv.Value;
        return obj;
    }

    private static JsonObject PositionToJson(InsertPosition pos)
    {
        var obj = new JsonObject();
        if (pos.Index.HasValue) obj["index"] = pos.Index.Value;
        if (pos.After != null) obj["after"] = pos.After;
        if (pos.Before != null) obj["before"] = pos.Before;
        return obj;
    }

    private static long ParseLongTolerant(JsonNode? node)
    {
        if (node is null) return 0;
        try
        {
            var v = node.AsValue();
            if (v.TryGetValue<long>(out var l)) return l;
            if (v.TryGetValue<double>(out var d)) return (long)d;
            if (v.TryGetValue<string>(out var s) && long.TryParse(s, out var sl)) return sl;
        }
        catch { }
        return 0;
    }

    /// <summary>
    /// Extract the <c>unsupported_properties</c> list from a <c>set</c> reply.
    /// Strict to the protocol §5.3 shape: <c>{"unsupported_properties":[...]}</c>.
    /// Bare-array replies are a plugin bug and fail loudly with
    /// <c>protocol_mismatch</c>, so plugin authors discover the drift
    /// immediately instead of having it silently absorbed by the host.
    /// </summary>
    private static List<string> ParseUnsupportedProperties(JsonNode? result)
    {
        if (result is null) return new List<string>();
        if (result is not JsonObject obj)
            throw new CliException(
                "Format-handler `set` reply must be a JSON object with `unsupported_properties` array (§5.3). " +
                $"Got: {result.GetType().Name}.")
            { Code = "protocol_mismatch" };
        if (obj["unsupported_properties"] is not JsonArray ups)
            return new List<string>();
        return ups
            .Where(n => n is not null)
            .Select(n => n!.GetValue<string>())
            .ToList();
    }

    /// <summary>
    /// Extract the new element's path from an <c>add</c> reply. Strict to the
    /// protocol §5.3 shape: <c>{"path":"...","unsupported_properties":[...]}</c>.
    /// Bare-string replies fail loudly with <c>protocol_mismatch</c>.
    /// </summary>
    private static string ParseAddPath(JsonNode? result)
    {
        if (result is null) return "";
        if (result is not JsonObject obj)
            throw new CliException(
                "Format-handler `add` reply must be a JSON object with `path` field (§5.3). " +
                $"Got: {result.GetType().Name}.")
            { Code = "protocol_mismatch" };
        return obj["path"]?.GetValue<string>() ?? "";
    }
}
