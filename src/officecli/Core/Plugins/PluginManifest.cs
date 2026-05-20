// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Serialization;
using OfficeCli.Core;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// The three plugin responsibilities defined in docs/plugin-protocol.md.
/// String values are the wire form used in plugin manifests.
/// </summary>
public enum PluginKind
{
    /// <summary>Foreign format → officecli commands (e.g. .doc → .docx via add/set).</summary>
    DumpReader,

    /// <summary>Native format → foreign output file (e.g. .docx → .pdf).</summary>
    Exporter,

    /// <summary>Plugin owns a foreign format end-to-end (e.g. .hwpx editing).</summary>
    FormatHandler,
}

public static class PluginKindExtensions
{
    public static string ToWireString(this PluginKind kind) => kind switch
    {
        PluginKind.DumpReader    => "dump-reader",
        PluginKind.Exporter      => "exporter",
        PluginKind.FormatHandler => "format-handler",
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    public static bool TryParseWire(string s, out PluginKind kind)
    {
        switch (s)
        {
            case "dump-reader":    kind = PluginKind.DumpReader;    return true;
            case "exporter":       kind = PluginKind.Exporter;      return true;
            case "format-handler": kind = PluginKind.FormatHandler; return true;
            default:               kind = default;                  return false;
        }
    }
}

/// <summary>
/// Manifest emitted by a plugin in response to `<plugin> --info`. Mirrors
/// the schema defined in docs/plugin-protocol.md §4.
/// </summary>
public sealed class PluginManifest
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("version")]
    public string Version { get; set; } = "";

    /// <summary>Protocol major version. Must be 1 for v1 plugins; the registry rejects mismatches.</summary>
    [JsonPropertyName("protocol")]
    public int Protocol { get; set; }

    /// <summary>Wire-form kind strings (`"dump-reader"`, etc.). Parsed via <see cref="PluginKindExtensions.TryParseWire"/>.</summary>
    [JsonPropertyName("kinds")]
    public List<string> Kinds { get; set; } = new();

    /// <summary>File extensions including the leading dot (e.g. `".doc"`).</summary>
    [JsonPropertyName("extensions")]
    public List<string> Extensions { get; set; } = new();

    /// <summary>
    /// Native format the plugin produces (dump-reader: the format the emitted
    /// batch is replayed into; exporter: the source-side native format).
    /// One of <c>"docx"</c>, <c>"xlsx"</c>, <c>"pptx"</c>. Required for
    /// dump-readers; the manifest reader applies a "docx" default when omitted
    /// so plugins authored before this field stay loadable.
    /// </summary>
    [JsonPropertyName("target")]
    public string? Target { get; set; }

    /// <summary>
    /// Declarative runtime tag, for diagnostics / `plugins list` display only.
    /// Main does not branch on this. One of: <c>dotnet</c>, <c>native</c>,
    /// <c>go</c>, <c>rust</c>, <c>python</c>, <c>other</c>.
    /// </summary>
    [JsonPropertyName("runtime")]
    public string? Runtime { get; set; }

    /// <summary>
    /// Idle-timeout budget per verb. Main's watchdog kills the plugin when no
    /// stdout/reply/heartbeat is observed for this many seconds. See §5.6.
    /// Required by the protocol; manifest reader applies a safe default
    /// (<see cref="PluginIdleTimeout.SafeDefault"/>) when the field is missing.
    /// </summary>
    [JsonPropertyName("idle_timeout_seconds")]
    public PluginIdleTimeout? IdleTimeoutSeconds { get; set; }

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("tier")]
    public string? Tier { get; set; }

    [JsonPropertyName("supports")]
    public List<string>? Supports { get; set; }

    [JsonPropertyName("limits")]
    public Dictionary<string, object>? Limits { get; set; }

    [JsonPropertyName("homepage")]
    public string? Homepage { get; set; }

    [JsonPropertyName("license")]
    public string? License { get; set; }

    [JsonPropertyName("vocabulary")]
    public PluginVocabulary? Vocabulary { get; set; }
}

/// <summary>
/// Idle-timeout block from manifest. <see cref="Default"/> applies to any verb
/// not listed in <see cref="Verbs"/>.
/// </summary>
public sealed class PluginIdleTimeout
{
    /// <summary>Fallback timeout in seconds. Must be positive (0 is disallowed in manifests).</summary>
    [JsonPropertyName("default")]
    public int Default { get; set; }

    /// <summary>Verb-specific overrides (e.g. <c>{"dump": 30, "save": 60}</c>).</summary>
    [JsonPropertyName("verbs")]
    public Dictionary<string, int>? Verbs { get; set; }

    /// <summary>Sentinel returned by the registry when a plugin omits the block.</summary>
    public static PluginIdleTimeout SafeDefault => new() { Default = 60 };

    /// <summary>Resolve the budget for a specific verb. Falls back to <see cref="Default"/>.</summary>
    public int For(string verb)
    {
        if (Verbs is not null && Verbs.TryGetValue(verb, out var v) && v > 0)
            return v;
        return Default > 0 ? Default : SafeDefault.Default;
    }
}

public static class PluginManifestExtensions
{
    /// <summary>
    /// Canonical target format name ("docx"/"xlsx"/"pptx"). Defaults to
    /// "docx" for plugins that omit the field. Throws if the manifest declares
    /// an unsupported target.
    /// </summary>
    public static string ResolveTargetFormat(this PluginManifest m)
    {
        var t = (m.Target ?? "docx").ToLowerInvariant();
        return t switch
        {
            "docx" or "xlsx" or "pptx" => t,
            _ => throw new InvalidOperationException(
                $"Plugin '{m.Name}' declares unsupported target '{m.Target}'. Expected one of: docx, xlsx, pptx."),
        };
    }

    /// <summary>
    /// File extension (with leading dot) for the plugin's target format.
    /// </summary>
    public static string ResolveTargetExtension(this PluginManifest m) =>
        "." + m.ResolveTargetFormat();

    /// <summary>
    /// Resolve the idle timeout for a verb, applying the safe default when the
    /// manifest is silent.
    /// </summary>
    public static int ResolveIdleTimeout(this PluginManifest m, string verb)
    {
        // Environment-variable escape hatch so a user hitting a hung plugin
        // can bypass the manifest budget without rebuilding the plugin:
        //
        //   OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS=0  → disable the watchdog
        //   OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS=N  → use N seconds for every verb
        //
        // This intentionally overrides per-verb manifest entries — the user
        // already knows the plugin is misbehaving; respecting the plugin's
        // declared limits at that point would defeat the purpose.
        var envOverride = Environment.GetEnvironmentVariable("OFFICECLI_PLUGIN_IDLE_TIMEOUT_SECONDS");
        if (!string.IsNullOrEmpty(envOverride) && int.TryParse(envOverride, out var envValue) && envValue >= 0)
            return envValue;
        return (m.IdleTimeoutSeconds ?? PluginIdleTimeout.SafeDefault).For(verb);
    }

    /// <summary>
    /// Inspect <paramref name="m"/> for soft-failures: a manifest can pass
    /// the hard protocol gate (§13) and still be missing recommended fields
    /// or carry values that will trip later at invocation time. Returns one
    /// human-readable line per finding; empty list = clean.
    ///
    /// Used by <c>plugins list</c> and <c>plugins lint</c> to surface drift
    /// at discovery time instead of at first command — plugin authors get
    /// feedback before users hit the failure path.
    /// </summary>
    public static List<string> Warnings(this PluginManifest m)
    {
        var warnings = new List<string>();

        if (m.IdleTimeoutSeconds is null || m.IdleTimeoutSeconds.Default <= 0)
            warnings.Add($"missing `idle_timeout_seconds.default` (host falls back to {PluginIdleTimeout.SafeDefault.Default}s); declare it explicitly per §4.1");

        if (m.Kinds is null || m.Kinds.Count == 0)
            warnings.Add("manifest declares no `kinds`; the plugin will never resolve for any verb");
        else
        {
            foreach (var k in m.Kinds)
                if (!PluginKindExtensions.TryParseWire(k, out _))
                    warnings.Add($"unknown kind '{k}'; expected one of: dump-reader / exporter / format-handler");
        }

        // dump-reader manifests should pin a target native format. The host
        // accepts a missing field as "docx" (see ResolveTargetFormat), but
        // warns when the field is present and unsupported.
        if (m.Kinds is not null && m.Kinds.Contains("dump-reader") && m.Target is not null)
        {
            var t = m.Target.ToLowerInvariant();
            if (t is not ("docx" or "xlsx" or "pptx"))
                warnings.Add($"`target` is '{m.Target}'; dump-reader must target one of: docx / xlsx / pptx");
        }

        if (m.Kinds is not null && m.Kinds.Contains("format-handler") && m.Vocabulary is null)
            warnings.Add("format-handler manifest is missing `vocabulary`; the runtime open-handshake snapshot still works but `plugins list` / `--help` will have nothing to show");

        return warnings;
    }
}

/// <summary>
/// Format-handler plugins declare the document model they expose via this
/// vocabulary. Used by main for autocomplete, command validation, and help.
/// Main does not interpret the semantics — it forwards commands using these names.
/// </summary>
public sealed class PluginVocabulary
{
    [JsonPropertyName("addable_types")]
    public List<string> AddableTypes { get; set; } = new();

    /// <summary>Map from type name (e.g. `"page"`) to the property names that type accepts.</summary>
    [JsonPropertyName("settable_props")]
    public Dictionary<string, List<string>> SettableProps { get; set; } = new();

    [JsonPropertyName("path_segments")]
    public List<string> PathSegments { get; set; } = new();
}

/// <summary>
/// Result of the open handshake (docs/plugin-protocol.md §5.3). Returned by
/// the plugin on session start; main caches it for the session's lifetime to
/// short-circuit unsupported commands and to resolve runtime vocabulary that
/// may differ from the manifest.
/// </summary>
public sealed class PluginSessionCapabilities
{
    [JsonPropertyName("capabilities")]
    public PluginCapabilities? Capabilities { get; set; }

    [JsonPropertyName("vocabulary")]
    public PluginVocabulary? Vocabulary { get; set; }
}

/// <summary>Capabilities subobject of <see cref="PluginSessionCapabilities"/>.</summary>
public sealed class PluginCapabilities
{
    /// <summary>Wire-form command verbs the plugin implements (e.g. <c>["get","set","save"]</c>).</summary>
    [JsonPropertyName("commands")]
    public List<string>? Commands { get; set; }

    /// <summary>Optional feature tags (e.g. <c>["save","extract-binary"]</c>).</summary>
    [JsonPropertyName("features")]
    public List<string>? Features { get; set; }
}

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(PluginManifest))]
[JsonSerializable(typeof(PluginVocabulary))]
[JsonSerializable(typeof(PluginIdleTimeout))]
[JsonSerializable(typeof(PluginSessionCapabilities))]
[JsonSerializable(typeof(PluginCapabilities))]
[JsonSerializable(typeof(List<string>))]
[JsonSerializable(typeof(DocumentNode))]
[JsonSerializable(typeof(List<DocumentNode>))]
[JsonSerializable(typeof(DocumentIssue))]
[JsonSerializable(typeof(List<DocumentIssue>))]
[JsonSerializable(typeof(ValidationError))]
[JsonSerializable(typeof(List<ValidationError>))]
internal partial class PluginJsonContext : JsonSerializerContext;
