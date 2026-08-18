// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

/// <summary>
/// Minimal MCP (Model Context Protocol) server over stdio.
/// Implements JSON-RPC 2.0 with initialize, tools/list, and tools/call.
/// All JSON is hand-written via Utf8JsonWriter to avoid reflection (PublishTrimmed).
/// </summary>
public static class McpServer
{
    private sealed record InFlightRequest(CancellationTokenSource Cancellation);

    public static async Task RunAsync()
    {
        using var reader = new StreamReader(Console.OpenStandardInput());
        using var writer = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };

        // Default this stdio process to NOT auto-spawn a resident. This opts
        // out of spawning only — it does NOT bypass an existing one: TryResident
        // still routes through a resident another officecli already holds for the
        // file (probe-then-TrySend in CommandBuilder.TryResident), so two writers
        // never fight over the file and no update is lost. The effect of the
        // opt-out:
        //   - no resident holds the file -> the command opens, applies, and
        //     eager-saves directly, so the mutation is on disk by the time the
        //     response returns;
        //   - a resident already holds the file -> the command routes through it
        //     and follows that resident's deferred flush (on disk at its
        //     save/close/idle), same as any other client of that resident.
        // Defaulting the opt-out on keeps a lone MCP session from leaving a
        // spawned resident (and its deferred-flush surprise) behind it. An
        // explicit user value (e.g. to opt INTO spawning residents) is respected.
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT")))
            Environment.SetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT", "1");

        // The MCP process always has stdin redirected (it IS the JSON-RPC
        // channel), so `batch --commands/--input` would emit the "stdin is also
        // redirected; stdin will be ignored" warning on EVERY call — noise that
        // also lands in the result text and breaks naive JSON parsing of the
        // batch envelope. Under MCP that warning describes the transport, not a
        // user mistake, so default its existing opt-out on (respect any explicit
        // value).
        if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT")))
            Environment.SetEnvironmentVariable("OFFICECLI_BATCH_ALLOW_STDIN_REDIRECT", "1");

        // MCP server is a long-lived stdio process. The normal
        // per-invocation auto-upgrade path (Program.cs:112) is
        // short-circuited for `officecli mcp` because CheckInBackground
        // is called AFTER the mcp branch in Program.cs — so without
        // this hook, an MCP instance started once and left running for
        // days/weeks would never see a new release.
        //
        // Run the upgrade path in the background: fire once at startup
        // (applies any pending .update from a previous run and kicks a
        // fresh check if >24h stale), then every hour. The hourly wake
        // is cheap because CheckInBackground is debounced by the same
        // 24h timestamp in ~/.officecli/config.json as the normal CLI
        // path, so 23 of 24 wakes no-op. The actual download / verify /
        // File.Move happens in a spawned subprocess whose stdio is
        // redirected (see UpdateChecker.SpawnRefreshProcess), so
        // nothing it does can corrupt our stdout JSON-RPC stream.
        using var upgradeCts = new CancellationTokenSource();
        var upgradeTask = RunPeriodicUpgradeCheckAsync(upgradeCts.Token);
        var inFlight = new ConcurrentDictionary<string, InFlightRequest>();
        var requestTasks = new ConcurrentDictionary<int, Task>();
        using var executionGate = new SemaphoreSlim(1, 1);
        using var writeGate = new SemaphoreSlim(1, 1);
        var nextTaskId = 0;

        try
        {
            while (true)
            {
                var line = await reader.ReadLineAsync();
                if (line == null) break;
                if (string.IsNullOrWhiteSpace(line)) continue;

                try
                {
                    using var doc = JsonDocument.Parse(line);
                    var root = doc.RootElement;
                    // The JSON-RPC root must be an Object (single request). Arrays
                    // are valid JSON-RPC 2.0 batch requests that we don't support;
                    // numbers/strings/bools/nulls are malformed entirely. Guard
                    // here before TryGetProperty, which throws on non-Object.
                    if (root.ValueKind != JsonValueKind.Object)
                    {
                        var msg = root.ValueKind == JsonValueKind.Array
                            ? "Invalid Request: batch requests are not supported"
                            : "Invalid Request: request must be a JSON object";
                        await WriteResponseAsync(writer, writeGate, ErrorJson(null, -32600, msg));
                        continue;
                    }
                    // Parse id BEFORE method so a malformed method ('method': 42)
                    // can still echo the original id back per JSON-RPC 2.0 §5.
                    JsonElement? id = root.TryGetProperty("id", out var idEl) ? idEl.Clone() : null;
                    // method must be a string per spec; non-string is an
                    // Invalid Request (-32600), not an internal error.
                    string? method = null;
                    if (root.TryGetProperty("method", out var m))
                    {
                        if (m.ValueKind != JsonValueKind.String)
                        {
                            await WriteResponseAsync(writer, writeGate, ErrorJson(id, -32600, "Invalid Request: 'method' must be a string"));
                            continue;
                        }
                        method = m.GetString();
                    }

                    if (method == "notifications/cancelled")
                    {
                        CancelRequest(root, inFlight);
                        continue;
                    }

                    var requestRoot = root.Clone();
                    var taskId = Interlocked.Increment(ref nextTaskId);
                    var task = ProcessRequestAsync(
                        method,
                        id,
                        requestRoot,
                        writer,
                        writeGate,
                        executionGate,
                        inFlight);
                    requestTasks[taskId] = task;
                    _ = task.ContinueWith(
                        completedTask => requestTasks.TryRemove(taskId, out _),
                        CancellationToken.None,
                        TaskContinuationOptions.ExecuteSynchronously,
                        TaskScheduler.Default);
                }
                catch (JsonException)
                {
                    await WriteResponseAsync(writer, writeGate, ErrorJson(null, -32700, "Parse error"));
                }
            }

            await Task.WhenAll(requestTasks.Values);
        }
        finally
        {
            upgradeCts.Cancel();
            try { await upgradeTask; } catch { }
        }
    }

    private static async Task ProcessRequestAsync(
        string? method,
        JsonElement? id,
        JsonElement root,
        StreamWriter writer,
        SemaphoreSlim writeGate,
        SemaphoreSlim executionGate,
        ConcurrentDictionary<string, InFlightRequest> inFlight)
    {
        var requestKey = id.HasValue ? id.Value.GetRawText() : null;
        using var cancellation = new CancellationTokenSource();
        if (requestKey != null && !inFlight.TryAdd(requestKey, new InFlightRequest(cancellation)))
        {
            await WriteResponseAsync(writer, writeGate, ErrorJson(id, -32600, "Invalid Request: duplicate in-flight request id"));
            return;
        }

        try
        {
            string? response;
            if (method == "tools/call")
            {
                await executionGate.WaitAsync(cancellation.Token);
                try
                {
                    cancellation.Token.ThrowIfCancellationRequested();
                    response = await Task.Run(() => HandleToolsCall(id, root), CancellationToken.None);
                }
                finally
                {
                    executionGate.Release();
                }
            }
            else
            {
                response = method switch
                {
                    "initialize" => HandleInitialize(id),
                    "notifications/initialized" => null,
                    "tools/list" => HandleToolsList(id),
                    "ping" => WriteJson(w => { w.WriteStartObject(); Rpc(w, id); w.WriteStartObject("result"); w.WriteEndObject(); w.WriteEndObject(); }),
                    _ => id.HasValue ? ErrorJson(id, -32601, $"Method not found: {OfficeCli.Help.SchemaHelpLoader.TruncateForError(method ?? "", 64)}") : null,
                };
            }

            // MCP cancellation is fire-and-forget. A cancelled request must not
            // race a late response onto stdout. Synchronous Office operations
            // are allowed to finish safely once entered; queued work is stopped.
            if (response != null && !cancellation.IsCancellationRequested)
                await WriteResponseAsync(writer, writeGate, response);
        }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested)
        {
            // Per MCP 2024-11-05, cancellation notifications have no response.
        }
        catch (Exception ex)
        {
            if (!cancellation.IsCancellationRequested)
                await WriteResponseAsync(writer, writeGate, ErrorJson(id, -32603, $"Internal error: {ex.Message}"));
        }
        finally
        {
            if (requestKey != null)
                inFlight.TryRemove(requestKey, out _);
        }
    }

    private static void CancelRequest(
        JsonElement root,
        ConcurrentDictionary<string, InFlightRequest> inFlight)
    {
        if (!root.TryGetProperty("params", out var parameters)
            || parameters.ValueKind != JsonValueKind.Object
            || !parameters.TryGetProperty("requestId", out var requestId))
            return;

        if (inFlight.TryGetValue(requestId.GetRawText(), out var request))
            request.Cancellation.Cancel();
    }

    private static async Task WriteResponseAsync(StreamWriter writer, SemaphoreSlim writeGate, string response)
    {
        await writeGate.WaitAsync();
        try
        {
            await writer.WriteLineAsync(response);
        }
        finally
        {
            writeGate.Release();
        }
    }

    private static async Task RunPeriodicUpgradeCheckAsync(CancellationToken token)
    {
        // Fire once at startup — no matter what state the config is in,
        // this applies any pending .update from a previous run and
        // (if stale) spawns a fresh download. Does not block the main
        // loop: this method runs on a background task.
        try { UpdateChecker.CheckInBackground(); } catch { }

        while (!token.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(TimeSpan.FromHours(1), token);
                UpdateChecker.CheckInBackground();
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch
            {
                // Never crash the MCP server over an update-check failure.
                // UpdateChecker already swallows exceptions internally, so
                // this is belt-and-braces for any future change that might
                // leak one through.
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
        var ver = Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0";
        w.WriteStartObject("serverInfo"); w.WriteString("name", "officecli"); w.WriteString("version", ver); w.WriteEndObject();
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
        // This server advertises exactly one tool. A misrouted call must not
        // silently execute (it would mutate files under a bogus tool name);
        // reject anything else the way an unknown tool should fail.
        if (name != "officecli")
            return ErrorJson(id, -32602, $"Unknown tool: {name}. This server exposes a single tool: officecli.");

        try
        {
            // Thin shell: the officecli tool takes a CLI command line in
            // `command` and runs it through the shared System.CommandLine root.
            var (contents, isError) = ExecuteCommandLine(args);
            return WriteJson(w =>
            {
                w.WriteStartObject();
                Rpc(w, id);
                w.WriteStartObject("result");
                w.WriteStartArray("content");
                foreach (var c in contents)
                {
                    w.WriteStartObject();
                    w.WriteString("type", c.Type);
                    if (c.Text != null) w.WriteString("text", c.Text);
                    if (c.Data != null) w.WriteString("data", c.Data);
                    if (c.MimeType != null) w.WriteString("mimeType", c.MimeType);
                    w.WriteEndObject();
                }
                w.WriteEndArray();
                w.WriteBoolean("isError", isError);
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
                // Only PRE-handler failures reach here now — argv extraction and
                // CLI parse/validation errors (the shared-grammar "free win"
                // messages). A handler that ran and exited non-zero (batch /
                // validate business verdicts) returns its stdout verbatim with
                // isError=true from ExecuteCommandLine, not via this path.
                var errText = $"Error: {ex.Message}";
                w.WriteStartObject(); w.WriteString("type", "text"); w.WriteString("text", errText); w.WriteEndObject();
                w.WriteEndArray();
                w.WriteBoolean("isError", true);
                w.WriteEndObject();
                w.WriteEndObject();
            });
        }
    }

    // ==================== Tool Execution ====================

    /// <summary>
    /// MCP content block. Most tool responses are a single text block; screenshot
    /// returns a text caption + an image block (base64 PNG). Fields not relevant
    /// to a given Type are left null and omitted on serialization.
    /// </summary>
    private sealed record McpContent(string Type, string? Text = null, string? Data = null, string? MimeType = null);

    // ==================== Thin command-line exec ====================
    // The MCP tool is a thin shell over the CLI: the caller passes the officecli
    // command line (a string, or a pre-split argv array) and it runs through the
    // SAME System.CommandLine root the CLI uses. No per-command marshalling here
    // means no argument can be silently dropped (every CLI flag works for free),
    // and the model writes exactly what the skills' CLI examples show.

    private static (IReadOnlyList<McpContent> Contents, bool IsError) ExecuteCommandLine(JsonElement args)
    {
        var argv = ExtractArgv(args);
        if (argv.Length == 0)
            throw new ArgumentException("Provide the officecli command line as `command`, e.g. "
                + "command=\"help\" or command=\"add deck.pptx /slide[1] --type shape --prop text=Hi\".");
        // load_skill / skills live in Program.cs early-dispatch, not in the
        // System.CommandLine root, so RunCliRaw can't reach them. Serve them
        // here from the same SkillInstaller the CLI uses.
        if (argv[0] is "load_skill" or "skill" or "skills")
            return (new[] { new McpContent("text", Text: HandleSkillCommand(argv)) }, false);
        if (IsScreenshot(argv))
            return (RunScreenshotArgv(argv), false);
        return SurfaceCliResult(RunCliRaw(argv));
    }

    private static string[] ExtractArgv(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty("command", out var c))
            return Array.Empty<string>();
        string[] argv;
        if (c.ValueKind == JsonValueKind.Array)
            // Preserve empty-string elements: the array form is exactly how a
            // caller delivers an intentionally-empty argument value (e.g.
            // `--prop text=` to clear text), so dropping "" would silently
            // diverge from the equivalent quoted "" in the string form. A
            // non-string element (a bare number/bool) is rendered to its JSON
            // text rather than throwing.
            argv = c.EnumerateArray()
                    .Select(e => e.ValueKind == JsonValueKind.String ? (e.GetString() ?? "") : e.GetRawText())
                    .ToArray();
        else if (c.ValueKind == JsonValueKind.String)
            argv = Tokenize(c.GetString() ?? "");
        else
            // A non-string, non-array `command` (number, bool, object, null) is a
            // client mistake — fall through to the empty-argv friendly guidance
            // rather than letting JsonElement.GetString() throw a raw .NET message.
            return Array.Empty<string>();
        // A model copying a skill example may include the leading binary name.
        if (argv.Length > 0 && (argv[0] == "officecli" || argv[0] == "officecli.exe"
            || argv[0].EndsWith("/officecli", StringComparison.Ordinal)))
            argv = argv[1..];
        return argv;
    }

    // Quote-aware tokenizer: splits on whitespace, honours single/double quotes
    // and backslash escapes inside double quotes. Never invokes a shell, so there
    // is no command-injection surface — tokens go straight to the in-process
    // System.CommandLine parser.
    private static string[] Tokenize(string s)
    {
        var tokens = new List<string>();
        var sb = new StringBuilder();
        bool inTok = false; char quote = '\0';
        for (int i = 0; i < s.Length; i++)
        {
            char ch = s[i];
            if (quote != '\0')
            {
                if (ch == quote) quote = '\0';
                // Inside double quotes a backslash only escapes the two chars
                // that would otherwise affect quoting itself (" and \), matching
                // bash double-quote semantics. ANY other backslash sequence is
                // preserved verbatim — crucially `\n` / `\t` stay two characters
                // so the downstream prop parser can turn them into a newline/tab.
                // (The old "escape the next char" rule swallowed the backslash,
                // turning text="A\nB" into the literal "AnB".)
                else if (ch == '\\' && quote == '"' && i + 1 < s.Length && (s[i + 1] == '"' || s[i + 1] == '\\'))
                    sb.Append(s[++i]);
                else sb.Append(ch);
                inTok = true;
            }
            else if (ch == '"' || ch == '\'') { quote = ch; inTok = true; }
            else if (char.IsWhiteSpace(ch)) { if (inTok) { tokens.Add(sb.ToString()); sb.Clear(); inTok = false; } }
            else { sb.Append(ch); inTok = true; }
        }
        if (inTok) tokens.Add(sb.ToString());
        return tokens.ToArray();
    }

    private static bool IsScreenshot(string[] argv)
        => argv.Length >= 2 && argv[0] == "view" && Array.IndexOf(argv, "screenshot") >= 0;

    // Mirror the CLI's load_skill early-dispatch (Program.cs): no name → catalog;
    // a name → that skill's SKILL.md; name + --path <rel> → one reference file.
    private static string HandleSkillCommand(string[] argv)
    {
        string? name = null, relPath = null;
        for (int i = 1; i < argv.Length; i++)
        {
            var a = argv[i];
            if (a == "--path" && i + 1 < argv.Length) { relPath = argv[++i]; continue; }
            if (a == "list") continue;   // `skills list`
            name ??= a;
        }
        if (string.IsNullOrEmpty(name))
            return OfficeCli.Core.SkillInstaller.BuildSkillCatalog();
        return string.IsNullOrEmpty(relPath)
            ? OfficeCli.Core.SkillInstaller.LoadSkillContent(name)
            : OfficeCli.Core.SkillInstaller.LoadSkillFile(name, relPath);
    }

    // screenshot delegates to the CLI (view <file> screenshot ... -o <tmp>) and
    // returns the rendered PNG inline as an image content block. Injects an -o
    // path when the caller didn't give one so we know which file to read back.
    private static IReadOnlyList<McpContent> RunScreenshotArgv(string[] argv)
    {
        var list = argv.ToList();
        int oi = list.FindIndex(a => a == "-o" || a == "--out");
        string outPath;
        string? autoTemp = null;
        if (oi >= 0 && oi + 1 < list.Count) outPath = list[oi + 1];
        else { outPath = Path.Combine(Path.GetTempPath(), $"officecli_mcp_shot_{Guid.NewGuid():N}.png"); autoTemp = outPath; list.Add("-o"); list.Add(outPath); }
        var r = RunCliRaw(list.ToArray());
        if (r.Exit != 0)
            throw new ArgumentException(StripErrPrefix(FirstNonEmpty(r.Stderr.Trim(), r.Stdout.Trim())));
        if (!File.Exists(outPath))
        {
            var m = System.Text.RegularExpressions.Regex.Match(r.Stdout, @"(\S+\.png)");
            if (m.Success && File.Exists(m.Groups[1].Value)) outPath = m.Groups[1].Value;
        }
        if (!File.Exists(outPath))
            return new[] { new McpContent("text", Text: r.Stdout.Trim().Length > 0 ? r.Stdout.Trim() : "Screenshot produced no image file.") };
        var b64 = Convert.ToBase64String(File.ReadAllBytes(outPath));
        // The PNG is returned inline as base64; an auto-injected temp file has no
        // further use, so don't leave it accumulating in the system temp dir. A
        // caller-supplied -o is theirs to keep.
        var caption = $"Screenshot saved to {outPath}";
        if (autoTemp != null && outPath == autoTemp)
        {
            try { File.Delete(autoTemp); } catch { /* best effort — inline data already captured */ }
            caption = "Screenshot rendered (returned inline).";
        }
        return new[]
        {
            new McpContent("text", Text: caption),
            new McpContent("image", Data: b64, MimeType: "image/png"),
        };
    }

    // ====================================================================
    // Shared-grammar dispatch (Phase 1 of routing MCP through the CLI's one
    // System.CommandLine root). Translating the MCP JSON into the CLI token
    // vector and parsing it with the SAME root the CLI uses means argument
    // validation, business logic, and the {success,data} envelope are shared
    // by construction — not re-marshalled (and re-bugged) by hand here.
    // ====================================================================
    private static RootCommand? _rootCommand;
    private static RootCommand RootCommand => _rootCommand ??= CommandBuilder.BuildRootCommand();

    private readonly record struct CliResult(int Exit, string Stdout, string Stderr);

    /// <summary>
    /// Parse+invoke argv through the shared CLI root, capturing stdout AND
    /// stderr and the exit code. argv is the CLI token vector, e.g.
    /// ["get", file, "/body", "--depth", "1", "--json"].
    ///
    /// Parse/validation failures are NOT short-circuited here. Letting Invoke
    /// run renders the SAME error + usage block a terminal user sees
    /// (System.CommandLine writes it to the captured stream and returns a
    /// non-zero exit WITHOUT running the handler), so the agent receives the
    /// full message — including the option list that points it at the right
    /// flag (e.g. batch's --commands/--input) — instead of a terse, usage-
    /// stripped one-liner. Surfacing only `pr.Errors` here used to drop that
    /// usage block, making MCP less informative than the bare CLI.
    /// </summary>
    private static CliResult RunCliRaw(string[] argv)
    {
        var pr = RootCommand.Parse(argv);
        var prevOut = Console.Out;
        var prevErr = Console.Error;
        var so = new System.IO.StringWriter();
        var se = new System.IO.StringWriter();
        int exit;
        try { Console.SetOut(so); Console.SetError(se); exit = pr.Invoke(); }
        finally { Console.SetOut(prevOut); Console.SetError(prevErr); }
        return new CliResult(exit, so.ToString(), se.ToString());
    }

    // Translate a CLI invocation's (exit, stdout, stderr) into MCP content.
    //
    // Always surface stdout AND stderr together when both are present, so the
    // caller never loses context — neither the "Added/Updated …" success line nor
    // an advisory caveat. A dangling-style add/set exits 0 with the warning on
    // stderr; both now report success WITH the warning visible (the warning was
    // previously dropped on the exit-0 path).
    //
    // Only a genuine failure raises: exit 1 (e.g. path not found, nothing applied),
    // or exit 2 with no stdout (goto/mark emit a usage error to stderr and write
    // no success line). An exit-2 add/set with a populated stdout is the CLI's
    // "applied with caveats" path — the element was added (envelope success:true)
    // and only an unsupported property was dropped; surfacing that as a hard error
    // makes agents re-issue an op that already landed. Scripts still see exit 2
    // from the CLI itself — fail-fast is preserved there, not on the MCP surface.
    private static (IReadOnlyList<McpContent> Contents, bool IsError) SurfaceCliResult(CliResult r)
    {
        var stdout = r.Stdout.TrimEnd('\n', '\r');
        var stderr = r.Stderr.Trim();
        var combined = stdout.Length > 0 && stderr.Length > 0 ? $"{stdout}\n{stderr}"
                     : stdout.Length > 0 ? stdout : stderr;
        // exit 2 with stdout = "applied with caveats" (element added, only an
        // unsupported property dropped) — the op landed, so it is NOT an error.
        bool appliedWithCaveats = r.Exit == 2 && stdout.Length > 0;
        bool isError = r.Exit != 0 && !appliedWithCaveats;
        // Surface the CLI output VERBATIM — exit mirrors envelope.success, so a
        // non-zero *business* verdict (batch with a failed step, validate
        // failure) still wrote its {success:false} envelope to stdout; prefixing
        // or munging it would break JSON parsing. A genuine process error has no
        // stdout and its handler-written stderr already carries its own "Error: "
        // prefix, so it too is surfaced as-is (no doubled prefix). The pass/fail
        // bit rides on isError, not on the text — exactly like a terminal user
        // reading stdout plus the exit code.
        var text = combined.Length == 0 ? (isError ? "Command failed." : "(ok)") : combined;
        return (new[] { new McpContent("text", Text: text) }, isError);
    }

    private static string FirstNonEmpty(params string[] xs) =>
        xs.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? "Command failed.";

    // The MCP catch block prepends "Error: "; the text-mode handlers already
    // wrote "Error: ..." to stderr. Strip one leading prefix to avoid doubling.
    private static string StripErrPrefix(string s) =>
        s.StartsWith("Error: ", StringComparison.Ordinal) ? s["Error: ".Length..] : s;

    // ==================== Tool Definitions ====================

    // MCP-specific guidance prepended to every help response. Cannot be derived
    // from schemas/help/*.json — it's about how to use the *tool*, not what the
    // *document model* exposes.
    private const string McpHelpStrategy = @"## Strategy
Use view (outline/stats/issues/annotated) to understand the document first, then get/query to inspect details, then set/add/remove to modify.
View modes: text, annotated, outline, stats, issues, html, svg (pptx only), screenshot, forms (docx only).
Before delivering, pass the delivery gate (see the tool description): validate clean, view issues clean, then a visual audit via view mode=screenshot when layout matters (slide decks most of all). Whether the visual audit is mandatory is format-specific — run `load_skill <pptx|word|excel>` for the authoritative per-format gate.
For 3+ mutations on the same file, use batch (one open/save cycle) instead of separate calls.
Get output keys can be used directly as Set input keys (round-trip safe).
Colors: FF0000, red, rgb(255,0,0), accent1. Sizes: 24pt. Positions: 2cm, 1in, 72pt, or raw EMU.
Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1.

";

    private const string ToolDescription = @"Create, read, and modify Office documents (.docx, .xlsx, .pptx) by running officecli command lines.

Pass an officecli command line in `command` (string or pre-split argv array); it runs through the same CLI you'd use in a terminal. Verbs: create, view (modes: text|annotated|outline|stats|issues|html|svg|screenshot|forms), get, query, set, add, remove, move, swap, validate, batch, raw, help, load_skill. Add --json to get/query/validate/view-issues for structured output. Examples (CLI syntax): create deck.pptx · add deck.pptx /slide[1] --type shape --prop ""text=Hi"" · set report.docx /body/p[1] --prop bold=true · view deck.pptx screenshot --page 2 · query book.xlsx ""cell[bold=true]"". Discover verbs/flags with `help`, and an element's schema with `help <format> <element>` (e.g. help pptx shape).

Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1. Props are key=value strings.

Delivery gate (before reporting a document finished — any failure = fix and re-check, do NOT deliver; validate passing is NOT delivery, 'looks like a real document' is):
1. Schema: `validate <file>` -> clean, no errors.
2. Content: `view <file> issues` -> no overflow/format/structure issues; and scan `view <file> text` for leftover placeholders (xxxx, lorem/ipsum, <TODO>, {{...}}, $VAR$, empty ()/[]).
3. Visual audit: `view <file> screenshot --page N` renders the page/slide and returns it as an image shown to you (or --grid auto for a whole-doc contact sheet). Judge it adversarially (assume problems exist) for overlap, text overflow, off-slide shapes, dark-on-dark, misalignment; fix positions/sizes (`set <file> <path> --prop x=.. --prop y=..`) and re-screenshot until right; if the screenshot can't render, say 'not visually verified'. Whether this audit is mandatory is format-specific (slide decks need it most — absolute-positioned shapes overlap invisibly to text modes), so run `load_skill pptx` (or word / excel) for the authoritative gate. The per-format SKILL.md, not this blurb, is the source of truth for what 'done' requires.
4. Flush to disk: end with `save <file>` — this guarantees your edits are written to disk before you hand the file off. Required final step, not optional (always safe — never errors or loses work; `close` also flushes if you want to end the session too).";

    private static void WriteToolDefinitions(Utf8JsonWriter w)
    {
        w.WriteStartObject();
        w.WriteString("name", "officecli");
        // Append a compact always-on skill-trigger summary so the agent is
        // prompted to load the right skill without the full ~1.2k of routing
        // descriptions resident in context. Detail stays lazy behind load_skill.
        w.WriteString("description", ToolDescription + "\n\n" + McpHelpStrategy + "\n"
            + OfficeCli.Core.SkillInstaller.BuildSkillTriggerSummary());
        w.WriteStartObject("inputSchema");
        w.WriteString("type", "object");
        w.WriteStartObject("properties");
        // Single param: the officecli command line, as a string or a pre-split
        // argv array. Everything else (verbs, flags, schemas) is discovered via
        // `help` and the loaded skills — no per-command schema to drift.
        w.WriteStartObject("command");
        w.WriteStartArray("type"); w.WriteStringValue("string"); w.WriteStringValue("array"); w.WriteEndArray();
        w.WriteStartObject("items"); w.WriteString("type", "string"); w.WriteEndObject();
        w.WriteString("description",
            "The officecli command line — either a single string (e.g. \"add deck.pptx /slide[1] --type shape --prop text=Hi\") "
            + "or a pre-split argv array of strings (use the array form when an argument contains spaces or quotes). A leading "
            + "'officecli' is optional. Examples: \"help\" lists commands; \"help pptx shape\" shows an element's schema; "
            + "\"view deck.pptx text\" reads it; \"view deck.pptx screenshot --page 2\" returns a rendered image; add --json to "
            + "get/query/validate for structured output. Run help first to learn the verbs and flags.");
        w.WriteEndObject();
        w.WriteEndObject(); // end properties
        w.WriteStartArray("required"); w.WriteStringValue("command"); w.WriteEndArray();
        w.WriteEndObject(); // end inputSchema
        w.WriteEndObject(); // end tool
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
