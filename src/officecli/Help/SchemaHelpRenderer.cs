// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;

namespace OfficeCli.Help;

/// <summary>
/// Renders a help schema JsonDocument into human-readable text or raw JSON.
/// </summary>
internal static class SchemaHelpRenderer
{
    internal static string RenderJson(JsonDocument doc)
    {
        // Use Utf8JsonWriter directly so the call is trim-safe (no reflection-
        // based serializer). JsonElement.WriteTo honors the writer's
        // WriteIndented setting.
        using var ms = new System.IO.MemoryStream();
        using (var writer = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
        {
            doc.RootElement.WriteTo(writer);
        }
        return System.Text.Encoding.UTF8.GetString(ms.ToArray());
    }

    /// <summary>
    /// Render a schema as human-readable text. When <paramref name="verbFilter"/>
    /// is one of add/set/get/query/remove, properties are filtered to those
    /// that declare <c>verbFilter: true</c>; the header carries a "(verb-view)"
    /// marker so callers can tell they are seeing a filtered page.
    /// </summary>
    internal static string RenderHuman(JsonDocument doc, string? verbFilter = null)
    {
        var sb = new StringBuilder();
        var root = doc.RootElement;

        var format = root.TryGetProperty("format", out var f) ? f.GetString() ?? "" : "";
        var element = root.TryGetProperty("element", out var e) ? e.GetString() ?? "" : "";
        var isContainer = root.TryGetProperty("container", out var c)
                          && c.ValueKind == JsonValueKind.True;

        var header = verbFilter == null
            ? $"{format} {element}"
            : $"{format} {verbFilter} {element}";
        sb.AppendLine(header);
        sb.AppendLine(new string('-', Math.Max(14, header.Length)));

        // When a verb filter is active, short-circuit if the element doesn't
        // support that verb at all — clearer than rendering an empty page.
        if (verbFilter != null
            && root.TryGetProperty("operations", out var opsEl)
            && (!opsEl.TryGetProperty(verbFilter, out var opVal)
                || opVal.ValueKind != JsonValueKind.True))
        {
            sb.AppendLine($"'{verbFilter}' is not supported on {format} {element}.");
            return sb.ToString().TrimEnd('\r', '\n');
        }

        if (isContainer)
            sb.AppendLine("Read-only container (never created or removed via CLI).");

        if (root.TryGetProperty("parent", out var parent))
        {
            var parentStr = parent.ValueKind switch
            {
                JsonValueKind.String => parent.GetString() ?? "",
                JsonValueKind.Array => string.Join(", ",
                    parent.EnumerateArray().Select(p => p.GetString() ?? "")),
                _ => "",
            };
            if (!string.IsNullOrEmpty(parentStr))
                sb.AppendLine($"Parent: {parentStr}");
        }

        if (root.TryGetProperty("paths", out var paths))
        {
            var pathList = new List<string>();
            if (paths.TryGetProperty("stable", out var stable))
                foreach (var p in stable.EnumerateArray())
                    if (p.GetString() is { } s) pathList.Add(s);
            if (paths.TryGetProperty("positional", out var pos))
                foreach (var p in pos.EnumerateArray())
                    if (p.GetString() is { } s) pathList.Add(s);
            if (pathList.Count > 0)
                sb.AppendLine($"Paths: {string.Join("  ", pathList)}");
        }

        if (root.TryGetProperty("addressing", out var addressing))
        {
            var form = addressing.TryGetProperty("pathForm", out var pf) ? pf.GetString() : null;
            if (!string.IsNullOrEmpty(form))
                sb.AppendLine($"Addressing: {form}");
        }

        if (root.TryGetProperty("operations", out var ops))
        {
            var active = new List<string>();
            foreach (var op in ops.EnumerateObject())
            {
                if (op.Value.ValueKind == JsonValueKind.True)
                    active.Add(op.Name);
            }
            if (active.Count > 0)
                sb.AppendLine($"Operations: {string.Join(" ", active)}");

            // Usage examples block: synthesize one CLI line per supported verb
            // from `paths.positional[0]` (fallback `paths.stable[0]`) + `element`.
            RenderUsageBlock(sb, root, element, isContainer, verbFilter);
        }

        if (root.TryGetProperty("properties", out var props)
            && props.ValueKind == JsonValueKind.Object)
        {
            sb.AppendLine();
            sb.AppendLine(verbFilter == null
                ? "Properties:"
                : $"Properties ({verbFilter}):");
            int shown = 0;
            foreach (var prop in props.EnumerateObject())
            {
                // When verb filter active, skip props that don't declare that verb.
                if (verbFilter != null)
                {
                    if (!prop.Value.TryGetProperty(verbFilter, out var pv)
                        || pv.ValueKind != JsonValueKind.True)
                        continue;
                }
                RenderProperty(sb, prop, isContainer);
                shown++;
            }
            if (verbFilter != null && shown == 0)
                sb.AppendLine($"  (no properties participate in '{verbFilter}' for this element)");
        }

        if (root.TryGetProperty("children", out var children)
            && children.ValueKind == JsonValueKind.Array
            && children.GetArrayLength() > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Children:");
            foreach (var child in children.EnumerateArray())
            {
                var el = child.TryGetProperty("element", out var ce) ? ce.GetString() : "?";
                var seg = child.TryGetProperty("pathSegment", out var ps) ? ps.GetString() : "?";
                var card = child.TryGetProperty("cardinality", out var cd) ? cd.GetString() : "?";
                sb.AppendLine($"  {el}  ({card})  /{seg}");
            }
        }

        if (root.TryGetProperty("note", out var note) && note.GetString() is { } noteStr)
        {
            sb.AppendLine();
            sb.AppendLine($"Note: {noteStr}");
        }

        if (root.TryGetProperty("examples", out var topExamples)
            && topExamples.ValueKind == JsonValueKind.Array
            && topExamples.GetArrayLength() > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Examples:");
            foreach (var ex in topExamples.EnumerateArray())
            {
                if (ex.ValueKind == JsonValueKind.String)
                {
                    if (ex.GetString() is { } s) sb.AppendLine($"  {s}");
                }
                else if (ex.ValueKind == JsonValueKind.Object)
                {
                    var title = ex.TryGetProperty("title", out var t) ? t.GetString() : null;
                    if (!string.IsNullOrEmpty(title)) sb.AppendLine($"  {title}:");
                    if (ex.TryGetProperty("commands", out var cmds) && cmds.ValueKind == JsonValueKind.Array)
                        foreach (var cmdElement in cmds.EnumerateArray())
                            if (cmdElement.GetString() is { } cs) sb.AppendLine($"    {cs}");
                    else if (ex.TryGetProperty("command", out var cmd) && cmd.GetString() is { } cmdStr)
                        sb.AppendLine($"    {cmdStr}");
                }
            }
        }

        return sb.ToString().TrimEnd('\r', '\n');
    }

    /// <summary>
    /// Emit a "Usage:" block with one CLI line per operation declared true
    /// in the schema. Parent path is derived from the first available
    /// positional/stable path by dropping its last segment.
    /// </summary>
    private static void RenderUsageBlock(
        StringBuilder sb, JsonElement root, string element,
        bool isContainer, string? verbFilter)
    {
        if (!root.TryGetProperty("operations", out var ops)) return;

        // Pick the first positional path, falling back to stable.
        string? firstPath = null;
        if (root.TryGetProperty("paths", out var paths))
        {
            if (paths.TryGetProperty("positional", out var pos)
                && pos.ValueKind == JsonValueKind.Array
                && pos.GetArrayLength() > 0)
            {
                firstPath = pos[0].GetString();
            }
            if (string.IsNullOrEmpty(firstPath)
                && paths.TryGetProperty("stable", out var stable)
                && stable.ValueKind == JsonValueKind.Array
                && stable.GetArrayLength() > 0)
            {
                firstPath = stable[0].GetString();
            }
        }
        if (string.IsNullOrEmpty(firstPath) || string.IsNullOrEmpty(element))
            return;

        var derivedParent = DeriveParentPath(firstPath!);
        var targetPath = firstPath!;

        // Prefer explicit `addParent` (string or array). When the element's
        // positional path describes the element's own location (e.g.
        // /comments/comment[N]) rather than a valid Add parent, schema authors
        // must declare addParent to keep the Usage line accurate.
        var addParents = new List<string>();
        if (root.TryGetProperty("addParent", out var apEl))
        {
            if (apEl.ValueKind == JsonValueKind.String && apEl.GetString() is { } aps)
                addParents.Add(aps);
            else if (apEl.ValueKind == JsonValueKind.Array)
                foreach (var p in apEl.EnumerateArray())
                    if (p.GetString() is { } ps) addParents.Add(ps);
        }
        if (addParents.Count == 0)
            addParents.Add(derivedParent);

        bool Has(string v) =>
            ops.TryGetProperty(v, out var ov) && ov.ValueKind == JsonValueKind.True;

        bool WantVerb(string v) => verbFilter == null || verbFilter == v;

        var lines = new List<string>();
        if (Has("add") && !isContainer && WantVerb("add"))
            foreach (var ap in addParents)
                lines.Add($"  officecli add <file> {ap} --type {element} [--prop key=val ...]");
        if (Has("set") && !isContainer && WantVerb("set"))
            lines.Add($"  officecli set <file> {targetPath} --prop key=val ...");
        if (Has("get") && WantVerb("get"))
            lines.Add($"  officecli get <file> {targetPath}");
        if (Has("query") && WantVerb("query"))
            lines.Add($"  officecli query <file> {element}");
        if (Has("remove") && !isContainer && WantVerb("remove"))
            lines.Add($"  officecli remove <file> {targetPath}");

        if (lines.Count == 0) return;

        sb.AppendLine();
        sb.AppendLine("Usage:");
        foreach (var line in lines) sb.AppendLine(line);
    }

    /// <summary>
    /// Drop the last segment of a path: "/body/p[N]" → "/body",
    /// "/slide[N]/shape[N]" → "/slide[N]", "/Sheet1/A1" → "/Sheet1".
    /// Single-segment paths are returned unchanged.
    /// </summary>
    private static string DeriveParentPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return path;
        var trimmed = path.TrimEnd('/');
        var lastSlash = trimmed.LastIndexOf('/');
        if (lastSlash < 0) return path;     // no slash at all — keep as-is
        if (lastSlash == 0) return "/";      // single absolute segment → root
        return trimmed.Substring(0, lastSlash);
    }

    private static void RenderProperty(StringBuilder sb, JsonProperty prop, bool isContainer)
    {
        var name = prop.Name;
        var body = prop.Value;

        var type = body.TryGetProperty("type", out var t) ? t.GetString() ?? "" : "";

        var opList = new List<string>();
        // For containers, skip Add/Set columns per spec.
        foreach (var op in new[] { "add", "set", "get" })
        {
            if (isContainer && op != "get") continue;
            if (body.TryGetProperty(op, out var val) && val.ValueKind == JsonValueKind.True)
                opList.Add(op);
        }
        var opsStr = opList.Count > 0 ? string.Join("/", opList) : "-";

        var aliasStr = "";
        if (body.TryGetProperty("aliases", out var aliases))
        {
            if (aliases.ValueKind == JsonValueKind.Array)
            {
                var list = aliases.EnumerateArray()
                    .Select(a => a.GetString())
                    .Where(a => !string.IsNullOrEmpty(a))
                    .ToList();
                if (list.Count > 0) aliasStr = $"   aliases: {string.Join(", ", list!)}";
            }
            else if (aliases.ValueKind == JsonValueKind.Object)
            {
                var list = aliases.EnumerateObject().Select(a => a.Name).ToList();
                if (list.Count > 0) aliasStr = $"   aliases: {string.Join(", ", list)}";
            }
        }

        sb.AppendLine($"  {name}   {type}   [{opsStr}]{aliasStr}");

        if (body.TryGetProperty("description", out var desc) && desc.GetString() is { } dstr)
            sb.AppendLine($"    description: {dstr}");

        if (body.TryGetProperty("values", out var values)
            && values.ValueKind == JsonValueKind.Array)
        {
            var vlist = values.EnumerateArray()
                .Select(v => v.GetString()).Where(v => !string.IsNullOrEmpty(v)).ToList();
            if (vlist.Count > 0)
                sb.AppendLine($"    values: {string.Join(", ", vlist!)}");
        }

        if (body.TryGetProperty("examples", out var examples)
            && examples.ValueKind == JsonValueKind.Array)
        {
            foreach (var ex in examples.EnumerateArray())
                if (ex.GetString() is { } exs)
                    sb.AppendLine($"    example: {exs}");
        }

        if (body.TryGetProperty("readback", out var rb) && rb.GetString() is { } rbstr)
            sb.AppendLine($"    readback: {rbstr}");
    }
}
