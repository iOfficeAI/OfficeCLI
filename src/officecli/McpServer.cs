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
            var result = ExecuteTool(toolName, args);
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
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(null, null), OutputFormat.Json),
                    "forms" or "f" => handler is Handlers.WordHandler wfh
                        ? wfh.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions)
                        : throw new ArgumentException("Forms view is only supported for .docx files."),
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
                var format = Arg("format").ToLowerInvariant();
                const string strategy = @"## Strategy
Use view (outline/stats/issues/annotated) to understand the document first, then get/query to inspect details, then set/add/remove to modify.
View modes: text, annotated, outline, stats, issues, html, svg (pptx only), forms (docx only).
For 3+ mutations on the same file, use batch (one open/save cycle) instead of separate calls.
Get output keys can be used directly as Set input keys (round-trip safe).
Colors: FF0000, red, rgb(255,0,0), accent1. Sizes: 24pt. Positions: 2cm, 1in, 72pt, or raw EMU.

";
                var reference = format switch
                {
                    "xlsx" => @"# XLSX Reference

## Add types
sheet, row, cell, col, run (rich text in cell), shape, chart, picture, comment, namedrange, table, validation, pivottable, autofilter, pagebreak, colbreak, rowbreak, sparkline
cf (conditional formatting): set type= to databar|colorscale|iconset|formula|topn|aboveaverage|duplicatevalues|uniquevalues|containstext|dateoccurring

## Cell properties (Set/Add)
value, formula, arrayformula, type (string|number|boolean), clear, link
bold, italic, strike, underline (true|single|double), superscript, subscript
font.color (#FF0000), font.size (14pt), font.name (Calibri), fill (#4472C4)
border.all (thin|medium|thick), border.left/right/top/bottom, border.color
alignment.horizontal (left|center|right), alignment.vertical, alignment.wrapText
numfmt (0%|#,##0.00|...), rotation (0-180), indent, shrinktofit
locked (true|false), formulahidden (true|false)

## Sheet properties (Set)
name, freeze (A2|B3|none), zoom (75-200), tabcolor (#FF0000|none)
autofilter (A1:F100|none), merge (A1:D1), protect (true|false), password
printarea ($A$1:$D$10|none), orientation (landscape|portrait), papersize (1=Letter|9=A4)
fittopage (1x2|true), header (&CPage &P), footer (&LConfidential), sort (A:asc,B:desc|none)

## Run properties (Set /Sheet/A1/run[N])
text, bold, italic, strike, underline, superscript, subscript, size, color, font

## Pivot table properties (Add/Set)
source/src (Sheet1!A1:D100), position/pos, rows, cols/columns, values (Name:Sum), filters
style (PivotStyleMedium9), layout (compact|outline|tabular), sort (asc|desc)
subtotals (on|off), rowGrandTotals, colGrandTotals
showRowHeaders, showColHeaders, showRowStripes, showColStripes, showLastColumn
aggregate (Sum|Count|Average|Max|Min|Product|CountNums|StdDev|Var)
showDataAs (normal|difference|percent|percent_of_row|percent_of_col|percent_of_total|running_total|rank_asc|rank_desc|index)
dataField{N}.showAs — per-field display transform override

## Chart properties (Add/Set)
charttype (bar|column|line|pie|doughnut|area|scatter|radar|...), title, dataRange/range, data
x, y, width, height (cm/in/pt/emu), categories, colors, series1..20 (.name, .values, .categories)
datalabels/labels, labelpos, labelfont, axistitle/vtitle, cattitle/htitle
axismin/min, axismax/max, majorunit, minorunit, axisnumfmt
gridlines, majorgridlines, minorgridlines, plotareafill/plotfill, chartareafill/chartfill
linewidth, linedash/dash, marker/markers, markersize, smooth, style/styleid/preset
gradient, gradients, transparency/opacity, trendline, secondaryaxis/secondary
referenceline, colorrule/conditionalcolor, view3d, camera, perspective
holesize, firstsliceangle (pie/doughnut), gapwidth/gap, overlap (bar/column)
legend (top|bottom|left|right|false), legendfont, title.font/size/color/bold
datatable, varycolors, scatterstyle, radarstyle

## Sparkline properties (Add/Set)
cell (F1), range/data (A1:E1), type (line|column|stacked), color, negativecolor
markers, highpoint, lowpoint, firstpoint, lastpoint, negative (boolean flags)
highmarkercolor, lowmarkercolor, firstmarkercolor, lastmarkercolor, markerscolor, lineweight

## CF properties
sqref/range, color (font), fill, bold, italic, strike, underline, border (thin|medium), numfmt
topn: rank, bottom (true), percent (true)
aboveaverage: below (true)
containstext: text
dateoccurring: period (today|yesterday|tomorrow|last7days|thisweek|lastweek|thismonth|lastmonth)

## Validation properties (Add)
sqref/ref, type (list|whole|decimal|date|time|textlength|custom), operator (between|equal|...)
formula1, formula2, allowBlank, showError, showInput, errorTitle, error, promptTitle, prompt

## Workbook settings (Set / or /workbook)
workbook.date1904, workbook.codeName, workbook.filterPrivacy
calc.mode (auto|manual), calc.iterate, calc.iterateCount, calc.refMode (A1|R1C1)
workbook.lockStructure, workbook.lockWindows

## Extended properties
extended.company, extended.manager, extended.template",

                    "pptx" => @"# PPTX Reference

## Add types
slide, shape, textbox, picture, chart, table, row, cell, col/column, paragraph, run
group, connector, animation, video, equation, notes, zoom, 3dmodel/model3d

## Shape properties (Set/Add)
text, bold, italic, underline, strike, superscript, subscript
color (#FF0000), font (Arial), size (24pt), align (left|center|right)
fill (#4472C4|gradient), outline (#000000), rotation (45)
x, y, width, height (in cm/in/pt/emu)
shadow, glow, reflection, softedge, effect3d
link (https://...), alt (alt text)

## Slide properties (Set)
layout, background, transition, notes

## Presentation settings (Set /)
firstSlideNum, rtl, compatMode
show.loop, show.narration, show.animation, show.useTimings
print.what (slides|notes|outline), print.colorMode, print.frameSlides
theme.color.accent1..6, theme.color.dk1/lt1/dk2/lt2/hlink/folHlink
theme.font.major.latin, theme.font.minor.latin

## Extended properties
extended.company, extended.manager, extended.template",

                    "docx" => @"# DOCX Reference

## Add types
paragraph, run, table, row, cell, picture, hyperlink, section
style, chart, equation, footnote, endnote, bookmark, comment
toc, pagebreak, columnbreak, header, footer, watermark, sdt
field (pagenum|numpages|date|author), formfield

## Run properties (Set/Add)
text, bold, italic, underline, strike, superscript, subscript
color (#FF0000), font (Arial), size (14pt), highlight
caps, smallcaps, vanish

## Paragraph properties (Set/Add)
alignment (left|center|right|justify)
spaceBefore (12pt), spaceAfter (6pt), lineSpacing (1.5x|18pt)
indent, hanging, firstline
pagebreakbefore (true|false)

## Section properties
pagewidth, pageheight, orientation (landscape|portrait)
margintop, marginbottom, marginleft, marginright (supports 2cm|1in|36pt|raw twips)

## Document settings (Set /)
docDefaults.font, docDefaults.fontSize, docDefaults.lineSpacing, docDefaults.color
docGrid.type (default|lines|linesAndChars|snapToCharacters), docGrid.linePitch
autoSpaceDE, autoSpaceDN, kinsoku, overflowPunct (true|false)
charSpacingControl (doNotCompress|compressPunctuation)
compatibility.preset (word2019|word2010|css-layout), compatibility.mode, compatibility.<flag>
embedFonts, mirrorMargins, gutterAtTop, bookFoldPrinting, evenAndOddHeaders
defaultTabStop (720|1.27cm), columns.count, columns.space, section.type

## Extended properties
extended.company, extended.manager, extended.template",

                    _ => null
                };
                if (reference == null)
                    return "Supported formats: xlsx, pptx, docx. Call help with one of these.";
                return strategy + reference;
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

    private const string ToolDescription = @"Create, read, and modify Office documents (.docx, .xlsx, .pptx).

Commands: create (file), view (file, mode: text|annotated|outline|stats|issues|html|svg|forms), get (file, path, depth), query (file, selector), set (file, path, props[]), add (file, parent, type, props[], index/after/before), remove (file, path), move (file, path, to, index/after/before), swap (file, path, path2), validate (file), batch (file, commands), raw (file, part), help (format: xlsx|pptx|docx).

Paths are 1-based: /slide[1]/shape[2], /body/p[3], /Sheet1/A1. Props are key=value strings. Call help for detailed property reference per format.";

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
        foreach (var c in new[] { "create", "view", "get", "query", "set", "add", "remove", "move", "swap", "validate", "batch", "raw", "help" })
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
        w.WriteStartObject("selector"); w.WriteString("type", "string"); w.WriteString("description", "CSS-like selector for query"); w.WriteEndObject();
        // props
        w.WriteStartObject("props"); w.WriteString("type", "array");
        w.WriteStartObject("items"); w.WriteString("type", "string"); w.WriteEndObject();
        w.WriteString("description", "key=value pairs (e.g. bold=true, color=FF0000, text=Hello)"); w.WriteEndObject();
        // mode
        w.WriteStartObject("mode"); w.WriteString("type", "string"); w.WriteString("description", "View mode: text, annotated, outline, stats, issues, html, svg (pptx), forms (docx)"); w.WriteEndObject();
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
