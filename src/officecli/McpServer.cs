// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Text;
using System.Text.Json;
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
    public static async Task RunAsync()
    {
        using var reader = new StreamReader(Console.OpenStandardInput());
        using var writer = new StreamWriter(Console.OpenStandardOutput()) { AutoFlush = true };

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

        try
        {
            while (true)
            {
                var line = await reader.ReadLineAsync();
                if (line == null) break;
                if (string.IsNullOrWhiteSpace(line)) continue;

                JsonElement? id = null;
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
                        await writer.WriteLineAsync(ErrorJson(null, -32600, msg));
                        continue;
                    }
                    // Parse id BEFORE method so a malformed method ('method': 42)
                    // can still echo the original id back per JSON-RPC 2.0 §5.
                    id = root.TryGetProperty("id", out var idEl) ? idEl.Clone() : null;
                    // method must be a string per spec; non-string is an
                    // Invalid Request (-32600), not an internal error.
                    string? method = null;
                    if (root.TryGetProperty("method", out var m))
                    {
                        if (m.ValueKind != JsonValueKind.String)
                        {
                            await writer.WriteLineAsync(ErrorJson(id, -32600, "Invalid Request: 'method' must be a string"));
                            continue;
                        }
                        method = m.GetString();
                    }

                    var response = method switch
                    {
                        "initialize" => HandleInitialize(id),
                        "notifications/initialized" => null,
                        "tools/list" => HandleToolsList(id),
                        "tools/call" => HandleToolsCall(id, root),
                        "ping" => WriteJson(w => { w.WriteStartObject(); Rpc(w, id); w.WriteStartObject("result"); w.WriteEndObject(); w.WriteEndObject(); }),
                        // CONSISTENCY(mcp-error): truncate caller-supplied value to prevent
                        // response amplification (echo arbitrary-length input back unchanged).
                        _ => id.HasValue ? ErrorJson(id, -32601, $"Method not found: {OfficeCli.Help.SchemaHelpLoader.TruncateForError(method ?? "", 64)}") : null,
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
                    await writer.WriteLineAsync(ErrorJson(id, -32603, $"Internal error: {ex.Message}"));
                }
            }
        }
        finally
        {
            upgradeCts.Cancel();
            try { await upgradeTask; } catch { }
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

        try
        {
            // Unified tool: route by "command" arg; legacy: route by tool name
            var toolName = name == "officecli" && args.ValueKind == JsonValueKind.Object && args.TryGetProperty("command", out var cmd)
                ? cmd.GetString() ?? name : name;
            var contents = ExecuteToolMulti(toolName, args);
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

    /// <summary>
    /// MCP content block. Most tool responses are a single text block; screenshot
    /// returns a text caption + an image block (base64 PNG). Fields not relevant
    /// to a given Type are left null and omitted on serialization.
    /// </summary>
    private sealed record McpContent(string Type, string? Text = null, string? Data = null, string? MimeType = null);

    /// <summary>
    /// Multi-modal wrapper around <see cref="ExecuteTool"/>. Special-cases
    /// view+screenshot (returns text caption + base64 PNG); everything else
    /// gets the legacy single-text path. Lets us add image responses without
    /// touching the ~50 string-returning case branches.
    /// </summary>
    private static IReadOnlyList<McpContent> ExecuteToolMulti(string name, JsonElement args)
    {
        if (name == "view" && args.ValueKind == JsonValueKind.Object
            && args.TryGetProperty("mode", out var m) && m.ValueKind == JsonValueKind.String)
        {
            var mode = m.GetString() ?? "";
            if (mode is "screenshot" or "p")
                return RunScreenshot(args);
        }
        return new[] { new McpContent("text", Text: ExecuteTool(name, args)) };
    }

    /// <summary>
    /// Render the document as HTML, headless-screenshot to PNG, return both a
    /// text caption (with the saved tmp PNG path, for agents with fs access)
    /// and the base64 PNG (for MCP-only agents). Mirrors the CLI's
    /// <c>view &lt;file&gt; screenshot</c> path; same backend probing
    /// (playwright → chrome → firefox) via <see cref="HtmlScreenshot"/>.
    /// </summary>
    private static IReadOnlyList<McpContent> RunScreenshot(JsonElement args)
    {
        string Arg(string key) => args.TryGetProperty(key, out var v) ? v.GetString() ?? "" : "";
        int? ArgIntOpt(string key) => args.TryGetProperty(key, out var v) && v.TryGetInt32(out var i) ? i : null;
        int ArgInt(string key, int def) => ArgIntOpt(key) ?? def;

        var file = Arg("file");
        if (string.IsNullOrEmpty(file)) throw new ArgumentException("file= required for screenshot");
        var start = ArgIntOpt("start");
        var end = ArgIntOpt("end");
        var width = ArgInt("screenshot_width", 1600);
        var height = ArgInt("screenshot_height", 1200);
        var grid = ArgInt("grid", 0);
        var renderMode = (Arg("render") is { Length: > 0 } rm ? rm : "auto").ToLowerInvariant();
        if (renderMode is not ("auto" or "native" or "html"))
            throw new ArgumentException($"Invalid render value: {renderMode}. Valid: auto, native, html");

        using var handler = DocumentHandlerFactory.Open(file);
        string? html = null;
        byte[]? directPng = null;
        if (handler is Handlers.PowerPointHandler ppt)
        {
            var pStart = start ?? 1;
            var pEnd = end ?? pStart;
            html = ppt.ViewAsHtml(pStart, pEnd, grid, width);
        }
        else if (handler is Handlers.ExcelHandler ex) html = ex.ViewAsHtml();
        else if (handler is Handlers.WordHandler wh)
        {
            // CONSISTENCY(screenshot-default-first-page): mirror CLI — screenshot
            // mode defaults to page 1 for docx so multi-page docs aren't silently
            // cropped by the viewport. Caller can pass start=N to override.
            var pageFilter = (start ?? 1).ToString();
            if (end is int e && e >= (start ?? 1)) pageFilter = $"{start ?? 1}-{e}";
            if (renderMode != "html" && OperatingSystem.IsWindows())
            {
                try { directPng = OfficeCli.Core.WordPdfBackend.Render(file, pageFilter); } catch { directPng = null; }
            }
            if (renderMode == "native" && directPng == null)
                throw new ArgumentException("render=native requires Windows with Microsoft Word installed.");
            if (directPng == null) html = wh.ViewAsHtml(pageFilter);
        }

        if (html == null && directPng == null)
            throw new ArgumentException("Screenshot mode is only supported for .pptx, .xlsx, and .docx files.");

        var stem = Path.GetFileNameWithoutExtension(file);
        var pngPath = Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{stem}_{Guid.NewGuid():N}.png");
        string backendName;
        if (directPng != null)
        {
            File.WriteAllBytes(pngPath, directPng);
            backendName = "word";
        }
        else
        {
            var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_preview_{stem}_{Guid.NewGuid():N}.html");
            File.WriteAllText(tmpHtml, html!);
            var r = OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, width, height);
            try { File.Delete(tmpHtml); } catch { /* ignore */ }
            if (!r.Ok)
                throw new InvalidOperationException(
                    "No headless browser available. Install Chrome/Edge/Chromium or Firefox, "
                    + "or `pip install playwright && playwright install chromium`."
                    + (r.Error != null ? $" Last error: {r.Error}" : ""));
            backendName = r.Backend;
        }

        var bytes = File.ReadAllBytes(pngPath);
        var b64 = Convert.ToBase64String(bytes);
        string pagesNote = "";
        if (handler is Handlers.PowerPointHandler pptp)
            pagesNote = $" Slides: {pptp.GetSlideCount()}.";
        var caption = $"Screenshot saved to {pngPath} ({bytes.Length} bytes, backend: {backendName}).{pagesNote}";
        return new[]
        {
            new McpContent("text", Text: caption),
            new McpContent("image", Data: b64, MimeType: "image/png"),
        };
    }

    private static string StatsWithOptionalPageCount(IDocumentHandler handler, JsonElement args, string file)
    {
        var stats = handler.ViewAsStats();
        var wantPages = args.ValueKind == JsonValueKind.Object
            && args.TryGetProperty("page_count", out var pcv)
            && (pcv.ValueKind == JsonValueKind.True || (pcv.ValueKind == JsonValueKind.String && pcv.GetString() == "true"));
        if (!wantPages || handler is not Handlers.WordHandler wh) return stats;
        int? pages = null;
        if (OperatingSystem.IsWindows())
        {
            try { pages = Core.WordPdfBackend.GetPageCount(file); } catch { pages = null; }
        }
        if (pages == null)
        {
            var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_pc_{Path.GetFileNameWithoutExtension(file)}_{Guid.NewGuid():N}.html");
            try
            {
                File.WriteAllText(tmpHtml, wh.ViewAsHtml(null));
                pages = Core.HtmlScreenshot.GetPageCountFromDom(tmpHtml);
            }
            finally { try { File.Delete(tmpHtml); } catch { } }
        }
        return pages.HasValue ? $"Pages: {pages}\n" + stats : stats;
    }

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
                if (mode is "html" or "h")
                {
                    if (handler is Handlers.PowerPointHandler pptH)
                        return pptH.ViewAsHtml(start, end);
                    if (handler is Handlers.ExcelHandler excelH)
                        return excelH.ViewAsHtml();
                    if (handler is Handlers.WordHandler wordH)
                        return wordH.ViewAsHtml();
                }
                if (mode is "svg" or "g" && handler is Handlers.PowerPointHandler pptSvg)
                    return pptSvg.ViewAsSvg(start ?? 1);
                return mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(start, end, maxLines, null),
                    "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, null),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => StatsWithOptionalPageCount(handler, args, file),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(null, null), OutputFormat.Json),
                    "forms" or "f" => handler switch {
                        Handlers.WordHandler wfh => wfh.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions),
                        Handlers.HwpxHandler hfh => hfh.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions),
                        _ => throw new ArgumentException("Forms view is only supported for .docx and .hwpx files."),
                    },
                    "tables" or "tbl" => handler is Handlers.HwpxHandler htm
                        ? htm.ViewAsTablesJson().ToJsonString(OutputFormatter.PublicJsonOptions)
                        : throw new ArgumentException("Tables view is only supported for .hwpx files."),
                    "markdown" or "md" => handler is Handlers.HwpxHandler hmm
                        ? hmm.ViewAsMarkdown()
                        : throw new ArgumentException("Markdown view is only supported for .hwpx files."),
                    "objects" or "obj" => handler is Handlers.HwpxHandler hom
                        ? hom.ViewAsObjectsJson().ToJsonString(OutputFormatter.PublicJsonOptions)
                        : throw new ArgumentException("Objects view is only supported for .hwpx files."),
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
                var textFilter = Arg("text");
                using var handler = DocumentHandlerFactory.Open(file);
                var filters = AttributeFilter.Parse(selector);
                if (handler is OfficeCli.Handlers.ExcelHandler
                    && selector.TrimStart().StartsWith("cell", StringComparison.OrdinalIgnoreCase))
                {
                    filters = AttributeFilter.NormalizeKeys(
                        filters, OfficeCli.Handlers.ExcelHandler.ResolveCellAttributeAlias);
                }
                var (results, _) = AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
                if (!string.IsNullOrEmpty(textFilter))
                    results = results.Where(n => n.Text != null && n.Text.Contains(textFilter, StringComparison.OrdinalIgnoreCase)).ToList();
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
                var after = Arg("after"); if (string.IsNullOrEmpty(after)) after = null;
                var before = Arg("before"); if (string.IsNullOrEmpty(before)) before = null;
                var position = index.HasValue ? InsertPosition.AtIndex(index.Value)
                    : after != null ? InsertPosition.AfterElement(after)
                    : before != null ? InsertPosition.BeforeElement(before)
                    : null;
                var props = ParseProps(ArgStringArray("props"));
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var resultPath = handler.Add(parent, type, position, props);
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
                var mvAfter = Arg("after"); if (string.IsNullOrEmpty(mvAfter)) mvAfter = null;
                var mvBefore = Arg("before"); if (string.IsNullOrEmpty(mvBefore)) mvBefore = null;
                var mvPosition = index.HasValue ? InsertPosition.AtIndex(index.Value)
                    : mvAfter != null ? InsertPosition.AfterElement(mvAfter)
                    : mvBefore != null ? InsertPosition.BeforeElement(mvBefore)
                    : null;
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var resultPath = handler.Move(path, to, mvPosition);
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
                var forceStr = Arg("force");
                var stopOnError = !string.Equals(forceStr, "true", StringComparison.OrdinalIgnoreCase);
                var items = JsonSerializer.Deserialize<List<BatchItem>>(commands, BatchJsonContext.Default.ListBatchItem);
                if (items == null || items.Count == 0)
                    throw new ArgumentException("No commands found in input.");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var results = new List<BatchResult>();
                for (int bi = 0; bi < items.Count; bi++)
                {
                    var item = items[bi];
                    try
                    {
                        var output = CommandBuilder.ExecuteBatchItem(handler, item, true);
                        results.Add(new BatchResult { Index = bi, Success = true, Output = output });
                    }
                    catch (Exception ex)
                    {
                        results.Add(new BatchResult { Index = bi, Success = false, Item = item, Error = ex.Message });
                        if (stopOnError) break;
                    }
                }
                var sw = new System.IO.StringWriter();
                CommandBuilder.PrintBatchResults(results, json: true, totalCount: items.Count, output: sw);
                return sw.ToString().Trim();
            }
            case "swap":
            {
                var file = Arg("file");
                var path = Arg("path");
                var path2 = Arg("path2");
                using var handler = DocumentHandlerFactory.Open(file, editable: true);
                var (p1, p2) = handler switch
                {
                    Handlers.PowerPointHandler ppt => ppt.Swap(path, path2),
                    Handlers.WordHandler word => word.Swap(path, path2),
                    Handlers.ExcelHandler excel => excel.Swap(path, path2),
                    _ => throw new InvalidOperationException("swap not supported for this document type")
                };
                return $"Swapped {p1} <-> {p2}";
            }
            case "raw":
            {
                var file = Arg("file");
                var part = Arg("part"); if (string.IsNullOrEmpty(part)) part = "/document";
                using var handler = DocumentHandlerFactory.Open(file);
                return handler.Raw(part, null, null, null);
            }
            case "help":
            {
                // Schema-driven help — single source of truth shared with the CLI's
                // `officecli help` command. The previous implementation was ~150 lines
                // of hardcoded markdown cheat sheets that drifted from schemas/help/*.json
                // (e.g. when chart aliases were added, this block was never updated).
                //
                // Shape (mirrors `officecli help <format> [<element>]`):
                //   {command:"help"}                          → list formats
                //   {command:"help", format:"docx"}           → list elements in that format
                //   {command:"help", format:"docx", type:"paragraph"} → full element schema
                //
                // The Strategy preamble is MCP-specific guidance that schemas don't (and
                // shouldn't) encode — kept inline as McpHelpStrategy.
                var format = Arg("format").ToLowerInvariant();
                var element = Arg("type"); // optional element to drill into

                if (string.IsNullOrEmpty(format))
                    return McpHelpStrategy
                        + "Supported formats: docx, xlsx, pptx.\n"
                        + "Call again with format=<docx|xlsx|pptx> to list elements; "
                        + "add type=<element> for full schema (properties, aliases, examples).";

                if (!OfficeCli.Help.SchemaHelpLoader.IsKnownFormat(format))
                {
                    // CONSISTENCY(mcp-error): truncate user-supplied value in error messages to prevent
                    // response amplification (caller echoes arbitrary-length input back unchanged).
                    var displayFormat = OfficeCli.Help.SchemaHelpLoader.TruncateForError(format, 64);
                    return $"Unknown format '{displayFormat}'. Supported: docx, xlsx, pptx.";
                }

                var canonical = OfficeCli.Help.SchemaHelpLoader.NormalizeFormat(format);
                var sb = new StringBuilder(McpHelpStrategy);

                if (string.IsNullOrEmpty(element))
                {
                    sb.Append("# ").Append(canonical.ToUpperInvariant()).AppendLine(" Elements");
                    sb.AppendLine();
                    foreach (var el in OfficeCli.Help.SchemaHelpLoader.ListElements(canonical))
                        sb.Append("- ").AppendLine(el);
                    sb.AppendLine();
                    var sampleElement = canonical switch { "docx" => "paragraph", "xlsx" => "cell", _ => "shape" };
                    sb.Append("Call again with type=<element> for the full schema. ");
                    sb.Append("Example: {\"command\":\"help\",\"format\":\"").Append(canonical)
                      .Append("\",\"type\":\"").Append(sampleElement).AppendLine("\"}");
                    return sb.ToString();
                }

                try
                {
                    using var doc = OfficeCli.Help.SchemaHelpLoader.LoadSchema(canonical, element);
                    sb.Append(OfficeCli.Help.SchemaHelpRenderer.RenderHuman(doc, null));
                    return sb.ToString();
                }
                catch (Exception ex)
                {
                    return $"{ex.Message}\n\nList available elements via: {{\"command\":\"help\",\"format\":\"{canonical}\"}}";
                }
            }
            case "load_skill":
            {
                // Return the embedded SKILL.md content for the named skill. Pure
                // read — no install side-effect. Identical semantics to the CLI
                // `officecli load_skill <name>` command (both share LoadSkillContent).
                // Agents that want disk-resident skills run `officecli skills install`
                // themselves.
                var skill = Arg("name");
                if (string.IsNullOrEmpty(skill))
                    throw new ArgumentException($"name= required. Available: {OfficeCli.Core.SkillInstaller.KnownSkillsList()}");
                try { return OfficeCli.Core.SkillInstaller.LoadSkillContent(skill); }
                catch (ArgumentException ex)
                {
                    // CONSISTENCY(mcp-error): error message already includes the
                    // truncated input via SkillInstaller; re-throw as-is so MCP
                    // returns a structured error to the caller.
                    throw new ArgumentException(ex.Message);
                }
            }
            default:
                // CONSISTENCY(mcp-error): truncate caller-supplied value to prevent
                // response amplification (echo arbitrary-length input back unchanged).
                throw new ArgumentException($"Unknown tool: {OfficeCli.Help.SchemaHelpLoader.TruncateForError(name, 64)}");
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

    // MCP-specific guidance prepended to every help response. Cannot be derived
    // from schemas/help/*.json — it's about how to use the *tool*, not what the
    // *document model* exposes.
    private const string McpHelpStrategy = @"## Strategy
Use view (outline/stats/issues/annotated) to understand the document first, then get/query to inspect details, then set/add/remove to modify.
View modes: text, annotated, outline, stats, issues, html, svg (pptx only), forms (docx only).
For 3+ mutations on the same file, use batch (one open/save cycle) instead of separate calls.
Get output keys can be used directly as Set input keys (round-trip safe).
Colors: FF0000, red, rgb(255,0,0), accent1. Sizes: 24pt. Positions: 2cm, 1in, 72pt, or raw EMU.
Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1.

";

    private const string ToolDescription = @"Create, read, and modify Office documents (.docx, .xlsx, .pptx).

Commands: create (file), view (file, mode: text|annotated|outline|stats|issues|html|svg|screenshot|forms), get (file, path, depth), query (file, selector), set (file, path, props[]), add (file, parent, type, props[], index/after/before), remove (file, path), move (file, path, to, index/after/before), swap (file, path, path2), validate (file), batch (file, commands), raw (file, part), help (format: docx|xlsx|pptx, optional type=<element> for full schema), load_skill (name: pptx|word|excel|morph-ppt|morph-ppt-3d|pitch-deck|academic-paper|data-dashboard|financial-model — returns the skill's SKILL.md guidance).

Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1. Props are key=value strings. Call help with format= to list elements, then help with format= and type= to drill into a specific element's schema (properties, aliases, examples).";

    private static void WriteToolDefinitions(Utf8JsonWriter w)
    {
        w.WriteStartObject();
        w.WriteString("name", "officecli");
        w.WriteString("description", ToolDescription);
        w.WriteStartObject("inputSchema");
        w.WriteString("type", "object");
        w.WriteStartObject("properties");
        // command
        w.WriteStartObject("command"); w.WriteString("type", "string");
        w.WriteStartArray("enum");
        foreach (var c in new[] { "create", "view", "get", "query", "set", "add", "remove", "move", "swap", "validate", "batch", "raw", "help", "load_skill" })
            w.WriteStringValue(c);
        w.WriteEndArray();
        w.WriteString("description", "Command to execute");
        w.WriteEndObject();
        // file
        w.WriteStartObject("file"); w.WriteString("type", "string"); w.WriteString("description", "Document file path"); w.WriteEndObject();
        // path
        w.WriteStartObject("path"); w.WriteString("type", "string"); w.WriteString("description", "DOM path (e.g. /slide[1]/shape[1], /Sheet1/A1, /body/p[1])"); w.WriteEndObject();
        // parent
        w.WriteStartObject("parent"); w.WriteString("type", "string"); w.WriteString("description", "Parent DOM path for add"); w.WriteEndObject();
        // type
        w.WriteStartObject("type"); w.WriteString("type", "string"); w.WriteString("description", "Element type for add (slide, shape, paragraph, run, table, picture, chart, etc.)"); w.WriteEndObject();
        // selector
        w.WriteStartObject("selector"); w.WriteString("type", "string"); w.WriteString("description", "CSS-like selector for query. Valid element types per handler: PPT — shape, textbox, title, picture, table, chart, placeholder, connector, group, zoom, ole, equation (NOT 'slide' — use 'slide[N]>shape' to scope); Excel — cell, sheet, row, column, table, chart, image; Word — paragraph, run, table, image, hyperlink, heading, list. Supports attribute filters ('shape[text=Hello]', 'paragraph[style=Normal] > run[font!=Arial]'), pseudo-selectors (:contains(...), :empty), and Excel cell aliases (bold, size → font.bold, font.size). Path-style selectors starting with '/' are rejected except '/slide[N]/...' scoping in PPT."); w.WriteEndObject();
        // text (query post-filter)
        w.WriteStartObject("text"); w.WriteString("type", "string"); w.WriteString("description", "Filter query results to elements whose text contains this substring (case-insensitive)"); w.WriteEndObject();
        // props
        w.WriteStartObject("props"); w.WriteString("type", "array");
        w.WriteStartObject("items"); w.WriteString("type", "string"); w.WriteEndObject();
        w.WriteString("description", "key=value pairs (e.g. bold=true, color=FF0000, text=Hello)"); w.WriteEndObject();
        // mode
        w.WriteStartObject("mode"); w.WriteString("type", "string"); w.WriteString("description", "View mode: text, annotated, outline, stats, issues, html, svg (pptx), forms (docx/hwpx)"); w.WriteEndObject();
        // depth
        w.WriteStartObject("depth"); w.WriteString("type", "number"); w.WriteString("description", "Child depth for get (default 1)"); w.WriteEndObject();
        // index
        w.WriteStartObject("index"); w.WriteString("type", "number"); w.WriteString("description", "Insert position (0-based) for add/move"); w.WriteEndObject();
        // to
        w.WriteStartObject("to"); w.WriteString("type", "string"); w.WriteString("description", "Target parent path for move"); w.WriteEndObject();
        // after, before, path2
        w.WriteStartObject("after"); w.WriteString("type", "string"); w.WriteString("description", "Insert after this sibling path (for add/move)"); w.WriteEndObject();
        w.WriteStartObject("before"); w.WriteString("type", "string"); w.WriteString("description", "Insert before this sibling path (for add/move)"); w.WriteEndObject();
        w.WriteStartObject("path2"); w.WriteString("type", "string"); w.WriteString("description", "Second path for swap"); w.WriteEndObject();
        // start, end, max_lines
        w.WriteStartObject("start"); w.WriteString("type", "number"); w.WriteString("description", "Start line for view"); w.WriteEndObject();
        w.WriteStartObject("end"); w.WriteString("type", "number"); w.WriteString("description", "End line for view"); w.WriteEndObject();
        w.WriteStartObject("max_lines"); w.WriteString("type", "number"); w.WriteString("description", "Max lines for view"); w.WriteEndObject();
        // commands
        w.WriteStartObject("commands"); w.WriteString("type", "string"); w.WriteString("description", "JSON array of batch commands"); w.WriteEndObject();
        // force
        w.WriteStartObject("force"); w.WriteString("type", "string"); w.WriteString("description", "Set to 'true' to continue batch on error (default: stop on first error)"); w.WriteEndObject();
        // part
        w.WriteStartObject("part"); w.WriteString("type", "string"); w.WriteString("description", "Part path for raw (e.g. /document, /styles, /slide[1])"); w.WriteEndObject();
        // format
        w.WriteStartObject("format"); w.WriteString("type", "string"); w.WriteString("description", "Document format for help: xlsx, pptx, docx"); w.WriteEndObject();
        // name (for load_skill)
        w.WriteStartObject("name"); w.WriteString("type", "string"); w.WriteString("description", "Skill name for load_skill: pptx, word, excel, morph-ppt, morph-ppt-3d, pitch-deck, academic-paper, data-dashboard, financial-model"); w.WriteEndObject();
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
