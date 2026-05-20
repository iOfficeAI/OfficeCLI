// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Core.Plugins;
using OfficeCli.Help;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildPluginsCommand(Option<bool> jsonOption)
    {
        var pluginsCommand = new Command("plugins", "Manage and inspect installed plugins");
        pluginsCommand.Add(BuildPluginsListCommand(jsonOption));
        pluginsCommand.Add(BuildPluginsInfoCommand(jsonOption));
        pluginsCommand.Add(BuildPluginsLintCommand(jsonOption));
        return pluginsCommand;
    }

    private static Command BuildPluginsListCommand(Option<bool> jsonOption)
    {
        var cmd = new Command("list", "List plugins discoverable in the standard search paths");
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var plugins = PluginRegistry.EnumerateAll();

            if (json)
            {
                using var stream = new MemoryStream();
                using (var w = new Utf8JsonWriter(stream))
                {
                    w.WriteStartArray();
                    foreach (var p in plugins)
                    {
                        w.WriteStartObject();
                        w.WriteString("name", p.Manifest.Name);
                        w.WriteString("version", p.Manifest.Version);
                        w.WriteNumber("protocol", p.Manifest.Protocol);
                        w.WritePropertyName("kinds");
                        JsonSerializer.Serialize(w, p.Manifest.Kinds, PluginJsonContext.Default.ListString);
                        w.WritePropertyName("extensions");
                        JsonSerializer.Serialize(w, p.Manifest.Extensions, PluginJsonContext.Default.ListString);
                        if (p.Manifest.Tier is { } tier) w.WriteString("tier", tier);
                        w.WriteString("path", p.ExecutablePath);
                        // Per-plugin manifest warnings — emit the same data
                        // that the text-mode table prints in its trailing
                        // Warnings: section, so script/AI consumers of --json
                        // can react to drift without parsing stderr.
                        var pluginWarnings = p.Manifest.Warnings();
                        if (pluginWarnings.Count > 0)
                        {
                            w.WritePropertyName("warnings");
                            JsonSerializer.Serialize(w, pluginWarnings, PluginJsonContext.Default.ListString);
                        }
                        w.WriteEndObject();
                    }
                    w.WriteEndArray();
                }
                Console.WriteLine(OutputFormatter.WrapEnvelope(Encoding.UTF8.GetString(stream.ToArray())));
                return 0;
            }

            if (plugins.Count == 0)
            {
                Console.WriteLine("No plugins installed.");
                Console.WriteLine("");
                Console.WriteLine("Plugins extend officecli to support additional formats (.doc, .hwpx, .pdf export, ...).");
                Console.WriteLine("See: docs/plugin-protocol.md for installation paths.");
                return 0;
            }

            // Plain-text table.
            var rows = plugins
                .Select(p => new
                {
                    Name = p.Manifest.Name,
                    Version = p.Manifest.Version,
                    Kinds = string.Join(",", p.Manifest.Kinds),
                    Exts = string.Join(",", p.Manifest.Extensions),
                    Path = p.ExecutablePath,
                })
                .ToList();

            int wName = Math.Max(4, rows.Max(r => r.Name.Length));
            int wVer = Math.Max(7, rows.Max(r => r.Version.Length));
            int wKinds = Math.Max(5, rows.Max(r => r.Kinds.Length));
            int wExts = Math.Max(11, rows.Max(r => r.Exts.Length));

            Console.WriteLine($"{"NAME".PadRight(wName)}  {"VERSION".PadRight(wVer)}  {"KINDS".PadRight(wKinds)}  {"EXTENSIONS".PadRight(wExts)}  PATH");
            foreach (var r in rows)
                Console.WriteLine($"{r.Name.PadRight(wName)}  {r.Version.PadRight(wVer)}  {r.Kinds.PadRight(wKinds)}  {r.Exts.PadRight(wExts)}  {r.Path}");

            // Surface manifest-level warnings discovered during enumeration —
            // missing idle_timeout_seconds, unknown kinds, invalid target,
            // missing format-handler vocabulary. These do not fail the
            // listing, but they help plugin authors notice drift before
            // users hit it at invocation time.
            var withWarnings = plugins
                .Select(p => (p.Manifest.Name, Warnings: p.Manifest.Warnings()))
                .Where(t => t.Warnings.Count > 0)
                .ToList();
            if (withWarnings.Count > 0)
            {
                Console.WriteLine();
                Console.WriteLine("Warnings:");
                foreach (var (name, ws) in withWarnings)
                    foreach (var w in ws)
                        Console.WriteLine($"  [{name}] {w}");
            }

            return 0;
        }, json); });

        return cmd;
    }

    private static Command BuildPluginsInfoCommand(Option<bool> jsonOption)
    {
        var nameArg = new Argument<string>("name")
        {
            Description = "Plugin name or path to its executable",
        };

        var cmd = new Command("info", "Show the full manifest for a single plugin");
        cmd.Add(nameArg);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var target = result.GetValue(nameArg) ?? "";
            var resolved = ResolveByNameOrPath(target);
            if (resolved is null)
                throw new CliException($"Plugin not found: '{target}'")
                {
                    Code = "plugin_not_found",
                    Suggestion = "Run `officecli plugins list` to see installed plugins, or provide the absolute path to the plugin executable.",
                };

            // Re-read the manifest raw rather than re-serializing from our typed
            // class: this preserves any extra fields the plugin emits beyond
            // what PluginManifest knows about, so `plugins info` is faithful to
            // the plugin's actual --info output.
            using var p = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = resolved.ExecutablePath,
                    Arguments = "--info",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                }
            };
            p.Start();
            var rawManifest = p.StandardOutput.ReadToEnd();
            if (!p.WaitForExit(5000)) { try { p.Kill(true); } catch { } }

            if (json)
            {
                var envelope = new System.Text.Json.Nodes.JsonObject
                {
                    ["path"] = resolved.ExecutablePath,
                    ["manifest"] = System.Text.Json.Nodes.JsonNode.Parse(rawManifest),
                };
                Console.WriteLine(OutputFormatter.WrapEnvelope(envelope.ToJsonString()));
                return 0;
            }

            Console.WriteLine($"Path: {resolved.ExecutablePath}");
            Console.WriteLine();
            // Pretty-print the manifest JSON via Utf8JsonWriter (AOT-safe,
            // unlike JsonSerializer.Serialize(JsonElement) which trips IL2026).
            try
            {
                using var doc = JsonDocument.Parse(rawManifest);
                using var ms = new MemoryStream();
                using (var w = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
                    doc.RootElement.WriteTo(w);
                Console.WriteLine(Encoding.UTF8.GetString(ms.ToArray()));
            }
            catch
            {
                Console.WriteLine(rawManifest);
            }
            return 0;
        }, json); });

        return cmd;
    }

    private static Command BuildPluginsLintCommand(Option<bool> jsonOption)
    {
        var nameArg = new Argument<string>("name")
        {
            Description = "Plugin name or path to its executable",
        };
        var fixtureOption = new Option<string?>("--fixture")
        {
            Description = "Path to a source file the dump-reader will read. " +
                          "Falls back to $OFFICECLI_LINT_FIXTURE if unset.",
        };

        var cmd = new Command("lint",
            "Lint a dump-reader plugin against the main schema. " +
            "Runs `plugin dump <fixture>`, parses the BatchItem stream, and verifies " +
            "every --prop key on add commands is declared in the plugin's target-format " +
            "schemas/help/<target>/*.json. Exits 1 if any unknown prop is found.");
        cmd.Add(nameArg);
        cmd.Add(fixtureOption);
        cmd.Add(jsonOption);

        cmd.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var target = result.GetValue(nameArg) ?? "";
            var fixturePath = result.GetValue(fixtureOption);
            var resolved = ResolveByNameOrPath(target)
                ?? throw new CliException($"Plugin not found: '{target}'")
                {
                    Code = "plugin_not_found",
                    Suggestion = "Run `officecli plugins list` to see installed plugins, or provide the absolute path to the plugin executable.",
                };

            // Only dump-readers are lintable today — the lint contract is
            // "the plugin's emitted batch must use prop keys the schema
            // recognises". Exporters and format-handlers have their own
            // contracts (binary output, vocabulary block) that don't go
            // through the schema validator.
            if (!resolved.Manifest.Kinds.Contains("dump-reader"))
                throw new CliException(
                    $"Plugin '{resolved.Manifest.Name}' is not a dump-reader; lint only applies to dump-reader plugins.")
                {
                    Code = "unsupported_plugin_kind",
                    Suggestion = "Run `officecli plugins info <name>` to see what kinds the plugin advertises.",
                };

            // Resolve a fixture: explicit arg → env var. No bundled fallback —
            // a fixture is a per-plugin concern; the plugin author (or CI)
            // pins it via --fixture or $OFFICECLI_LINT_FIXTURE.
            var srcExt = resolved.Manifest.Extensions.FirstOrDefault() ?? "";
            fixturePath ??= Environment.GetEnvironmentVariable("OFFICECLI_LINT_FIXTURE");
            if (string.IsNullOrEmpty(fixturePath) || !File.Exists(fixturePath))
                throw new CliException(
                    $"No lint fixture available for plugin '{resolved.Manifest.Name}'" +
                    (string.IsNullOrEmpty(srcExt) ? "" : $" ({srcExt})") + ". " +
                    "Pass --fixture <path>, or set OFFICECLI_LINT_FIXTURE.")
                {
                    Code = "lint_fixture_missing",
                    Suggestion = "Provide a small foreign-format file the plugin can dump.",
                };

            // The plugin's declared target format determines which schema
            // tree to validate emitted props against (default: docx).
            var schemaFormat = resolved.Manifest.ResolveTargetFormat();

            // Run the plugin and stream JSONL.
            var items = new List<BatchItem>();
            var findings = new List<LintFinding>();

            void OnLine(string raw)
            {
                // Mirrors DumpReaderInvoker: strip per-line UTF-8 BOM so a
                // plugin that BOMs every JSONL line passes lint as well as
                // it passes invocation.
                var line = raw.TrimStart('﻿').Trim();
                if (line.Length == 0) return;
                if (line[0] == '[')
                    throw new CliException(
                        $"Plugin '{resolved.Manifest.Name}' emitted a JSON array; protocol v1 requires JSONL (one BatchItem per line).")
                    { Code = "corrupt_batch" };

                BatchItem? item;
                try { item = JsonSerializer.Deserialize(line, BatchJsonContext.Default.BatchItem); }
                catch (JsonException ex)
                {
                    throw new CliException(
                        $"Plugin '{resolved.Manifest.Name}' emitted invalid JSON at line #{items.Count}: {ex.Message}")
                    { Code = "plugin_contract_violation" };
                }
                if (item is null) return;
                items.Add(item);
            }

            var idle = resolved.Manifest.ResolveIdleTimeout("dump");
            var runResult = PluginProcess.Run(new PluginProcess.RunOptions
            {
                ExecutablePath = resolved.ExecutablePath,
                Arguments = new[] { "dump", fixturePath },
                IdleTimeoutSeconds = idle,
                OnStdoutLine = OnLine,
            });
            if (PluginProcess.LineCallbackError is CliException ce) throw ce;
            if (PluginProcess.LineCallbackError is not null) throw PluginProcess.LineCallbackError;
            if (runResult.IdleTimedOut)
                throw new CliException(
                    $"Plugin '{resolved.Manifest.Name}' produced no output for {idle}s — likely hung.")
                { Code = "plugin_idle_timeout" };
            if (runResult.ExitCode != 0)
                throw new CliException(
                    $"Plugin '{resolved.Manifest.Name}' dump failed (exit {runResult.ExitCode}) on fixture '{fixturePath}': {TruncateForLint(runResult.Stderr, 500)}")
                { Code = "plugin_failed" };

            // Validate both add and set props against the target-format
            // schema. BatchItem.Type is used verbatim as the schema element
            // name for add; set commands look up the element via the path's
            // leaf type when available, falling back to lenient validation
            // when the schema doesn't recognize the inferred element.
            for (int i = 0; i < items.Count; i++)
            {
                var it = items[i];
                if (it is null) continue;
                var verb = (it.Command ?? "").ToLowerInvariant();
                if (verb != "add" && verb != "set") continue;
                if (it.Props == null || it.Props.Count == 0) continue;

                // For add: explicit Type. For set: best-effort infer from the
                // last segment of the path (e.g. "/p[1]/r[2]" → "r"). This
                // matches main's own ValidateProperties contract.
                string typeName = verb == "add"
                    ? (it.Type ?? "")
                    : InferTypeFromPath(it.Path ?? "");
                if (string.IsNullOrEmpty(typeName)) continue;

                bool schemaExists;
                try { using var _ = SchemaHelpLoader.LoadSchema(schemaFormat, typeName); schemaExists = true; }
                catch { schemaExists = false; }

                if (!schemaExists)
                {
                    if (verb == "add")
                        findings.Add(new LintFinding(i, typeName, typeName, "<unknown_type>"));
                    // For set: silently skip unknown leaf types — the path
                    // might reference an aliased element the loader doesn't
                    // expose by name. Don't false-positive plugin authors.
                    continue;
                }

                var unknown = SchemaHelpLoader.ValidateProperties(
                    schemaFormat, typeName, verb, it.Props);
                foreach (var key in unknown)
                    findings.Add(new LintFinding(i, typeName, typeName, key));
            }

            if (json)
            {
                using var ms = new MemoryStream();
                using (var w = new Utf8JsonWriter(ms))
                {
                    w.WriteStartObject();
                    w.WriteString("plugin", resolved.Manifest.Name);
                    w.WriteString("path", resolved.ExecutablePath);
                    w.WriteString("fixture", fixturePath);
                    w.WriteNumber("batch_items", items.Count);
                    w.WriteNumber("unknown_prop_count", findings.Count);
                    w.WritePropertyName("unknown_props");
                    w.WriteStartArray();
                    foreach (var f in findings)
                    {
                        w.WriteStartObject();
                        w.WriteNumber("index", f.Index);
                        w.WriteString("type", f.Type);
                        w.WriteString("element", f.Element);
                        w.WriteString("prop", f.Prop);
                        w.WriteEndObject();
                    }
                    w.WriteEndArray();
                    w.WriteEndObject();
                }
                Console.WriteLine(OutputFormatter.WrapEnvelope(Encoding.UTF8.GetString(ms.ToArray())));
            }
            else
            {
                Console.WriteLine($"Plugin: {resolved.Manifest.Name}  ({resolved.ExecutablePath})");
                Console.WriteLine($"Fixture: {fixturePath}");
                Console.WriteLine($"Batch items: {items.Count}");
                if (findings.Count == 0)
                {
                    Console.WriteLine($"Result: OK — every emitted prop is declared in the {schemaFormat} schema.");
                }
                else
                {
                    Console.WriteLine($"Result: {findings.Count} unknown prop(s):");
                    foreach (var f in findings)
                        Console.WriteLine($"  [#{f.Index}] type={f.Type}  element={f.Element}  prop=\"{f.Prop}\"  (not declared in schemas/help/{schemaFormat}/{f.Element}.json)");
                }

                // Surface manifest-level soft warnings (§4 recommended fields,
                // unknown kinds, suspect target) so lint covers the same
                // ground as `plugins list` without users having to run both.
                var manifestWarnings = resolved.Manifest.Warnings();
                if (manifestWarnings.Count > 0)
                {
                    Console.WriteLine();
                    Console.WriteLine($"Manifest warnings:");
                    foreach (var w in manifestWarnings)
                        Console.WriteLine($"  - {w}");
                }
            }

            return findings.Count == 0 ? 0 : 1;
        }, json); });

        return cmd;
    }

    private sealed record LintFinding(int Index, string Type, string Element, string Prop);

    private static string TruncateForLint(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";

    /// <summary>
    /// Best-effort extraction of the leaf element type from a path like
    /// <c>"/body/p[1]/r[2]"</c> → <c>"r"</c>. Used by lint to look up the
    /// schema for set commands, which (unlike add) do not carry an explicit
    /// Type. Returns empty string for unrecognizable paths.
    /// </summary>
    private static string InferTypeFromPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return "";
        var slash = path.LastIndexOf('/');
        var leaf = slash < 0 ? path : path.Substring(slash + 1);
        var bracket = leaf.IndexOf('[');
        if (bracket > 0) leaf = leaf.Substring(0, bracket);
        return leaf;
    }

    private static ResolvedPlugin? ResolveByNameOrPath(string target)
    {
        // Path mode: absolute or relative path that exists.
        if (target.Contains(Path.DirectorySeparatorChar) || target.Contains(Path.AltDirectorySeparatorChar) || File.Exists(target))
        {
            var full = Path.GetFullPath(target);
            if (File.Exists(full) && PluginRegistry.TryReadManifest(full, out var m))
                return new ResolvedPlugin(full, m);
            return null;
        }

        // Name mode: search the full enumeration for a manifest whose name matches.
        var all = PluginRegistry.EnumerateAll();
        return all.FirstOrDefault(p =>
            string.Equals(p.Manifest.Name, target, StringComparison.OrdinalIgnoreCase));
    }
}
