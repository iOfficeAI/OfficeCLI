// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    public static RootCommand BuildRootCommand()
    {
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON (AI-friendly)" };

        var rootCommand = new RootCommand("""
            officecli: AI-friendly CLI for Office documents (.docx, .xlsx, .pptx)

            Help navigation (start from the deepest level you know):
              officecli pptx set              All settable elements and their properties
              officecli pptx set shape        Shape properties in detail
              officecli pptx set shape.fill   Specific property format and examples

            Replace 'pptx' with 'docx' or 'xlsx'. Commands: view, get, query, set, add, raw.
            """);
        rootCommand.Add(jsonOption);

        // ==================== open command (start resident) ====================
        var openFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var openCommand = new Command("open", "Start a resident process to keep the document in memory for faster subsequent commands");
        openCommand.Add(openFileArg);
        openCommand.Add(jsonOption);

        openCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(openFileArg)!;
            var filePath = file.FullName;

            // If already running, reuse the existing resident
            if (ResidentClient.TryConnect(filePath, out _))
            {
                var msg = $"Opened {file.Name} (already running, do NOT call close)";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
                return 0;
            }

            if (OperatingSystem.IsWindows())
            {
                // Windows: run the resident server in-process on a background thread.
                // Forking a child process deadlocks on Windows due to .NET single-file
                // host + redirected-pipe interactions. In-process avoids this while
                // keeping the same named-pipe API.
                //
                // Readiness is detected via ManualResetEventSlim instead of connecting
                // back through the named pipe (same-process pipe I/O via
                // StreamReader/StreamWriter deadlocks on Windows).
                var server = new ResidentServer(filePath);
                var cts = new CancellationTokenSource();
                var serverTask = Task.Run(() => server.RunAsync(cts.Token));

                if (!server.WaitUntilReady(TimeSpan.FromSeconds(5)))
                {
                    if (serverTask.IsCompleted)
                    {
                        server.Dispose();
                        if (serverTask.IsFaulted)
                            throw new InvalidOperationException($"Resident server failed: {serverTask.Exception?.InnerException?.Message}");
                        throw new InvalidOperationException("Resident server exited unexpectedly.");
                    }
                    cts.Cancel();
                    server.Dispose();
                    throw new InvalidOperationException("Resident server failed to start.");
                }

                var msg2 = $"Opened {file.Name} (remember to call close when done)";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg2));
                else Console.WriteLine(msg2);
                // Block on the server task — keeps the process alive until
                // close is called via the named pipe.
                try { serverTask.GetAwaiter().GetResult(); } catch (OperationCanceledException) { }
                server.Dispose();
                return 0;
            }

            // Linux/macOS: fork a background process running the resident server.
            // The open command returns immediately, leaving the child alive.
            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath == null)
                throw new InvalidOperationException("Cannot determine executable path.");

            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = $"__resident-serve__ \"{filePath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            var process = Process.Start(startInfo);
            if (process == null)
                throw new InvalidOperationException("Failed to start resident process.");

            // Wait briefly for the server to start accepting connections
            for (int i = 0; i < 50; i++) // up to 5 seconds
            {
                Thread.Sleep(100);
                if (ResidentClient.TryConnect(filePath, out _))
                {
                    var msg2 = $"Opened {file.Name} (remember to call close when done)";
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg2));
                    else Console.WriteLine(msg2);
                    return 0;
                }
                if (process.HasExited)
                {
                    var stderr = process.StandardError.ReadToEnd();
                    throw new InvalidOperationException($"Resident process exited. {stderr}");
                }
            }

            throw new InvalidOperationException("Resident process started but not responding.");
        }, json); });

        rootCommand.Add(openCommand);

        // ==================== close command (stop resident) ====================
        var closeFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var closeCommand = new Command("close", "Stop the resident process for the document");
        closeCommand.Add(closeFileArg);
        closeCommand.Add(jsonOption);

        closeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(closeFileArg)!;
            if (ResidentClient.SendClose(file.FullName))
            {
                var msg = $"Resident closed for {file.Name}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
            }
            else
            {
                throw new InvalidOperationException($"No resident running for {file.Name}");
            }
            return 0;
        }, json); });

        rootCommand.Add(closeCommand);

        // ==================== __resident-serve__ (internal, hidden) ====================
        var serveFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var serveCommand = new Command("__resident-serve__", "Internal: run resident server (do not call directly)");
        serveCommand.Hidden = true;
        serveCommand.Add(serveFileArg);

        serveCommand.SetAction(result =>
        {
            var file = result.GetValue(serveFileArg)!;
            using var server = new ResidentServer(file.FullName);
            server.RunAsync().GetAwaiter().GetResult();
        });

        rootCommand.Add(serveCommand);

        // Register commands from partial files
        rootCommand.Add(BuildWatchCommand());
        rootCommand.Add(BuildUnwatchCommand());
        rootCommand.Add(BuildViewCommand(jsonOption));
        rootCommand.Add(BuildGetCommand(jsonOption));
        rootCommand.Add(BuildQueryCommand(jsonOption));
        rootCommand.Add(BuildSetCommand(jsonOption));
        rootCommand.Add(BuildAddCommand(jsonOption));
        rootCommand.Add(BuildRemoveCommand(jsonOption));
        rootCommand.Add(BuildMoveCommand(jsonOption));
        rootCommand.Add(BuildRawCommand(jsonOption));
        rootCommand.Add(BuildRawSetCommand(jsonOption));
        rootCommand.Add(BuildAddPartCommand(jsonOption));
        rootCommand.Add(BuildValidateCommand(jsonOption));
        rootCommand.Add(BuildCheckCommand(jsonOption));
        rootCommand.Add(BuildBatchCommand(jsonOption));
        rootCommand.Add(BuildImportCommand(jsonOption));
        rootCommand.Add(BuildCreateCommand(jsonOption));
        rootCommand.Add(BuildMergeCommand(jsonOption));

        HelpCommands.Register(rootCommand);

        return rootCommand;
    }

    // ==================== Helper: try forwarding to resident ====================
    internal static int? TryResident(string filePath, Action<ResidentRequest> configure, bool json = false)
    {
        var request = new ResidentRequest();
        configure(request);
        if (json) request.Json = true;

        var response = ResidentClient.TrySend(filePath, request);
        if (response == null)
            return null;

        if (json)
        {
            // JSON mode: resident already built the envelope, just pass through
            if (!string.IsNullOrEmpty(response.Stdout))
                Console.WriteLine(response.Stdout);
        }
        else
        {
            if (!string.IsNullOrEmpty(response.Stdout))
                Console.WriteLine(response.Stdout);
            if (!string.IsNullOrEmpty(response.Stderr))
                Console.Error.WriteLine(response.Stderr);
        }

        return response.ExitCode;
    }


    internal static int SafeRun(Func<int> action, bool json = false)
    {
        if (!OfficeCli.Core.CliLogger.Enabled)
        {
            try
            {
                return action();
            }
            catch (Exception ex)
            {
                WriteError(ex, json);
                return 1;
            }
        }

        // Logging enabled: capture stdout/stderr
        var stdoutWriter = new StringWriter();
        var stderrWriter = new StringWriter();
        var origOut = Console.Out;
        var origErr = Console.Error;
        Console.SetOut(new TeeWriter(origOut, stdoutWriter));
        Console.SetError(new TeeWriter(origErr, stderrWriter));
        try
        {
            var code = action();
            var stdout = stdoutWriter.ToString().TrimEnd('\r', '\n');
            OfficeCli.Core.CliLogger.LogOutput(stdout);
            return code;
        }
        catch (Exception ex)
        {
            WriteError(ex, json);
            var stderr = stderrWriter.ToString().TrimEnd('\r', '\n');
            OfficeCli.Core.CliLogger.LogError(stderr);
            return 1;
        }
        finally
        {
            Console.SetOut(origOut);
            Console.SetError(origErr);
        }
    }

    private static void WriteError(Exception ex, bool json)
    {
        if (json)
        {
            // JSON mode: structured error envelope to stdout so AI agents get it in the same stream
            WarningContext.End(); // discard any partial warnings
            Console.WriteLine(OutputFormatter.WrapErrorEnvelope(ex));
        }
        else
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    internal static string ExecuteBatchItem(OfficeCli.Core.IDocumentHandler handler, BatchItem item, bool json)
    {
        var format = json ? OfficeCli.Core.OutputFormat.Json : OfficeCli.Core.OutputFormat.Text;
        var props = item.Props ?? new Dictionary<string, string>();

        switch (item.Command.ToLowerInvariant())
        {
            case "get":
            {
                var path = item.Path ?? "/";
                var depth = item.Depth ?? 1;
                var node = handler.Get(path, depth);
                return OfficeCli.Core.OutputFormatter.FormatNode(node, format);
            }
            case "query":
            {
                var selector = item.Selector ?? "";
                var filters = OfficeCli.Core.AttributeFilter.Parse(selector);
                var (results, warnings) = OfficeCli.Core.AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
                if (item.Text is { } textFilter && !string.IsNullOrEmpty(textFilter))
                    results = results.Where(n => n.Text != null && n.Text.Contains(textFilter, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var w in warnings) Console.Error.WriteLine(w);
                return OfficeCli.Core.OutputFormatter.FormatNodes(results, format);
            }
            case "set":
            {
                if (string.IsNullOrEmpty(item.Path))
                    throw new ArgumentException("'set' command requires 'path' field. Example: {\"command\": \"set\", \"path\": \"/slide[1]\", \"props\": {\"bold\": \"true\"}}");
                var path = item.Path;
                var unsupported = handler.Set(path, props);
                var applied = props.Where(kv => !unsupported.Contains(kv.Key)).ToList();
                var parts = new List<string>();
                if (applied.Count > 0)
                    parts.Add($"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}");
                if (unsupported.Count > 0)
                    parts.Add(FormatUnsupported(unsupported));
                return string.Join("\n", parts);
            }
            case "add":
            {
                var parentPath = item.Parent ?? item.Path;
                if (string.IsNullOrEmpty(parentPath))
                    throw new ArgumentException("'add' command requires 'parent' field. Example: {\"command\": \"add\", \"parent\": \"/slide[1]\", \"type\": \"shape\", \"props\": {\"text\": \"Hello\"}}");
                if (string.IsNullOrEmpty(item.Type) && string.IsNullOrEmpty(item.From))
                    throw new ArgumentException("'add' command requires 'type' or 'from' field. Example: {\"command\": \"add\", \"parent\": \"/\", \"type\": \"slide\"}");
                if (!string.IsNullOrEmpty(item.From))
                {
                    var resultPath = handler.CopyFrom(item.From, parentPath, item.Index.HasValue ? InsertPosition.AtIndex(item.Index.Value) : null);
                    return $"Copied to {resultPath}";
                }
                else
                {
                    var type = item.Type ?? "";
                    var resultPath = handler.Add(parentPath, type, item.Index.HasValue ? InsertPosition.AtIndex(item.Index.Value) : null, props);
                    return $"Added {type} at {resultPath}";
                }
            }
            case "remove":
            {
                if (string.IsNullOrEmpty(item.Path))
                    throw new ArgumentException("'remove' command requires 'path' field. Example: {\"command\": \"remove\", \"path\": \"/slide[1]/shape[2]\"}");
                var path = item.Path;
                var warning = handler.Remove(path);
                var msg = $"Removed {path}";
                if (warning != null) msg += $"\n{warning}";
                return msg;
            }
            case "move":
            {
                var path = item.Path ?? "/";
                var resultPath = handler.Move(path, item.To, item.Index.HasValue ? InsertPosition.AtIndex(item.Index.Value) : null);
                return $"Moved to {resultPath}";
            }
            case "view":
            {
                var mode = item.Mode ?? "text";
                if (mode.ToLowerInvariant() is "html" or "h")
                {
                    if (handler is OfficeCli.Handlers.PowerPointHandler pptH)
                        return pptH.ViewAsHtml();
                    if (handler is OfficeCli.Handlers.ExcelHandler excelH)
                        return excelH.ViewAsHtml();
                    if (handler is OfficeCli.Handlers.WordHandler wordH)
                        return wordH.ViewAsHtml();
                }
                if (mode.ToLowerInvariant() is "svg" or "g" && handler is OfficeCli.Handlers.PowerPointHandler pptSvg)
                {
                    return pptSvg.ViewAsSvg(1);
                }
                return mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(null, null, null, null),
                    "annotated" or "a" => handler.ViewAsAnnotated(null, null, null, null),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OfficeCli.Core.OutputFormatter.FormatIssues(handler.ViewAsIssues(null, null), format),
                    _ => $"Unknown mode: {mode}"
                };
            }
            case "raw":
            {
                if (string.IsNullOrEmpty(item.Part))
                    throw new ArgumentException("'raw' command requires 'part' field. Example: {\"command\": \"raw\", \"part\": \"/document\"} (docx), {\"command\": \"raw\", \"part\": \"/presentation\"} (pptx), {\"command\": \"raw\", \"part\": \"/sheet[1]\"} (xlsx)");
                return handler.Raw(item.Part, null, null, null);
            }
            case "raw-set":
            {
                var partPath = item.Part ?? "/document";
                var xpath = item.Xpath ?? "";
                var action = item.Action ?? "";
                handler.RawSet(partPath, xpath, action, item.Xml);
                return $"raw-set {action} applied";
            }
            case "validate":
            {
                var errors = handler.Validate();
                if (errors.Count == 0) return "Validation passed: no errors found.";
                var lines = new List<string> { $"Found {errors.Count} validation error(s):" };
                foreach (var err in errors)
                {
                    lines.Add($"  [{err.ErrorType}] {err.Description}");
                    if (err.Path != null) lines.Add($"    Path: {err.Path}");
                    if (err.Part != null) lines.Add($"    Part: {err.Part}");
                }
                return string.Join("\n", lines);
            }
            default:
                if (string.IsNullOrEmpty(item.Command))
                    throw new InvalidOperationException(
                        "Batch item missing required 'command' field. " +
                        "Valid commands: get, query, set, add, remove, move, view, raw, validate. " +
                        "Example: {\"command\": \"set\", \"path\": \"/Sheet1/A1\", \"props\": {\"value\": \"hello\"}}");
                throw new InvalidOperationException($"Unknown command: '{item.Command}'. Valid commands: get, query, set, add, remove, move, view, raw, validate.");
        }
    }

    private static Dictionary<string, string> ParsePropsArray(string[]? props)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in props ?? Array.Empty<string>())
        {
            var eqIdx = prop.IndexOf('=');
            if (eqIdx > 0)
                dict[prop[..eqIdx]] = prop[(eqIdx + 1)..];
        }
        return dict;
    }

    internal static void PrintBatchResults(List<BatchResult> results, bool json, int totalCount = 0, TextWriter? output = null)
    {
        var @out = output ?? Console.Out;
        if (totalCount == 0) totalCount = results.Count;

        if (json)
        {
            var succeeded = results.Count(r => r.Success);
            var failed = results.Count - succeeded;
            var skipped = totalCount - results.Count;

            using var ms = new System.IO.MemoryStream();
            using (var writer = new System.Text.Json.Utf8JsonWriter(ms))
            {
                writer.WriteStartObject();
                writer.WritePropertyName("results");
                System.Text.Json.JsonSerializer.Serialize(writer, results, BatchJsonContext.Default.ListBatchResult);
                writer.WriteStartObject("summary");
                writer.WriteNumber("total", totalCount);
                writer.WriteNumber("executed", results.Count);
                writer.WriteNumber("succeeded", succeeded);
                writer.WriteNumber("failed", failed);
                writer.WriteNumber("skipped", skipped);
                writer.WriteEndObject();
                writer.WriteEndObject();
            }

            var fullBytes = ms.ToArray();
            if (fullBytes.Length <= 8192)
            {
                @out.WriteLine(System.Text.Encoding.UTF8.GetString(fullBytes));
            }
            else
            {
                // Spill full output to temp file
                var tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"officecli_batch_{Guid.NewGuid():N}.json");
                System.IO.File.WriteAllBytes(tempPath, fullBytes);

                // Write slim envelope
                using var slimMs = new System.IO.MemoryStream();
                using (var slimWriter = new System.Text.Json.Utf8JsonWriter(slimMs))
                {
                    slimWriter.WriteStartObject();
                    slimWriter.WriteString("outputFile", tempPath);
                    slimWriter.WriteNumber("outputSize", fullBytes.Length);
                    slimWriter.WriteStartArray("results");
                    foreach (var r in results)
                    {
                        slimWriter.WriteStartObject();
                        slimWriter.WriteNumber("index", r.Index);
                        slimWriter.WriteBoolean("success", r.Success);
                        if (r.Error != null)
                        {
                            slimWriter.WriteString("error", r.Error);
                            if (r.Item != null)
                            {
                                slimWriter.WritePropertyName("item");
                                System.Text.Json.JsonSerializer.Serialize(slimWriter, r.Item, BatchJsonContext.Default.BatchItem);
                            }
                        }
                        slimWriter.WriteEndObject();
                    }
                    slimWriter.WriteEndArray();
                    slimWriter.WriteStartObject("summary");
                    slimWriter.WriteNumber("total", totalCount);
                    slimWriter.WriteNumber("executed", results.Count);
                    slimWriter.WriteNumber("succeeded", succeeded);
                    slimWriter.WriteNumber("failed", failed);
                    slimWriter.WriteNumber("skipped", skipped);
                    slimWriter.WriteEndObject();
                    slimWriter.WriteEndObject();
                }
                @out.WriteLine(System.Text.Encoding.UTF8.GetString(slimMs.ToArray()));
            }
        }
        else
        {
            for (int i = 0; i < results.Count; i++)
            {
                var r = results[i];
                var prefix = $"[{i + 1}] ";
                if (r.Success)
                {
                    if (!string.IsNullOrEmpty(r.Output))
                        @out.WriteLine($"{prefix}{r.Output}");
                    else
                        @out.WriteLine($"{prefix}OK");
                }
                else
                {
                    @out.WriteLine($"{prefix}ERROR: {r.Error}");
                }
            }

            var succeeded = results.Count(r => r.Success);
            var failed = results.Count - succeeded;
            @out.WriteLine($"\nBatch complete: {succeeded} succeeded, {failed} failed, {results.Count} total");
        }
    }

    private static string FormatValidationErrors(List<ValidationError> errors)
    {
        var sb = new StringBuilder();
        sb.Append("{\"count\":").Append(errors.Count).Append(",\"errors\":[");
        for (int i = 0; i < errors.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var e = errors[i];
            sb.Append("{\"type\":\"").Append(EscapeJson(e.ErrorType)).Append('"');
            sb.Append(",\"description\":\"").Append(EscapeJson(e.Description)).Append('"');
            if (e.Path != null) sb.Append(",\"path\":\"").Append(EscapeJson(e.Path)).Append('"');
            if (e.Part != null) sb.Append(",\"part\":\"").Append(EscapeJson(e.Part)).Append('"');
            sb.Append('}');
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static string EscapeJson(string s) => s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");

    internal static List<CliWarning>? ReportNewErrorsAsWarnings(OfficeCli.Core.IDocumentHandler handler, HashSet<string> errorsBefore)
    {
        var errorsAfter = handler.Validate();
        var newErrors = errorsAfter.Where(e => !errorsBefore.Contains(e.Description)).ToList();
        if (newErrors.Count == 0) return null;
        return newErrors.Select(err => new CliWarning
        {
            Message = $"[{err.ErrorType}] {err.Description}" +
                (err.Path != null ? $" (Path: {err.Path})" : "") +
                (err.Part != null ? $" (Part: {err.Part})" : ""),
            Code = "validation_error"
        }).ToList();
    }

    internal static void ReportNewErrors(OfficeCli.Core.IDocumentHandler handler, HashSet<string> errorsBefore, List<CliWarning>? preComputed = null)
    {
        var warnings = preComputed ?? ReportNewErrorsAsWarnings(handler, errorsBefore);
        if (warnings is { Count: > 0 })
        {
            Console.WriteLine($"VALIDATION: {warnings.Count} new error(s) introduced:");
            foreach (var w in warnings)
                Console.WriteLine($"  {w.Message}");
        }
    }

    /// <summary>
    /// Detect bare key=value tokens and --key value flag patterns in unmatched arguments (user forgot --prop).
    /// Returns a list of "key=value" strings suitable for --prop suggestions.
    /// </summary>
    internal static List<string> DetectUnmatchedKeyValues(System.CommandLine.ParseResult parseResult)
    {
        var result = new List<string>();
        var tokens = parseResult.UnmatchedTokens;
        var knownPropsLower = new HashSet<string>(KnownProps.Select(p => p.ToLowerInvariant()));

        for (int i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];

            // Pattern 1: bare key=value (e.g. "text=Hello")
            if (System.Text.RegularExpressions.Regex.IsMatch(token, @"^[A-Za-z_.][A-Za-z0-9_.]*=.+$"))
            {
                result.Add(token);
                continue;
            }

            // Pattern 2: --key value (e.g. "--text Hello" or "--fill yellow")
            // Only match if the key (without --) is a known property name
            if (token.StartsWith("--") && token.Length > 2)
            {
                var key = token[2..];
                if (knownPropsLower.Contains(key.ToLowerInvariant()) && i + 1 < tokens.Count)
                {
                    var nextToken = tokens[i + 1];
                    // Don't consume the next token if it also looks like a flag
                    if (!nextToken.StartsWith("--"))
                    {
                        result.Add($"{key}={nextToken}");
                        i++; // skip the value token
                        continue;
                    }
                }
            }
        }
        return result;
    }

    internal static string FormatUnsupported(IEnumerable<string> unsupported)
    {
        var parts = new List<string>();
        foreach (var prop in unsupported)
        {
            var suggestion = SuggestProperty(prop);
            parts.Add(suggestion != null ? $"{prop} (did you mean: {suggestion}?)" : prop);
        }
        return $"UNSUPPORTED props: {string.Join(", ", parts)}. Use 'officecli help <format>-set' to see available properties, or use raw-set for direct XML manipulation.";
    }

    internal static readonly string[] KnownProps = new[]
    {
        "text", "bold", "italic", "underline", "strike", "font", "size", "color",
        "highlight", "alignment", "spacing", "indent", "shd", "border",
        "width", "height", "valign", "header", "formula", "value", "type",
        "fill", "src", "path", "title", "name", "style", "caps", "smallcaps",
        "lineSpacing", "listStyle", "start", "level", "cols", "rows",
        "gridspan", "vmerge", "nowrap", "padding", "margin",
        "orientation", "pageWidth", "pageHeight",
        "x", "y", "cx", "cy", "rotation", "opacity",
        "border.color", "border.width", "border.style",
        "font.color", "font.size", "font.name", "font.bold", "font.italic",
        "hyperlink", "link", "tooltip", "alt", "description",
        "font.strike", "font.underline", "tabColor", "shadow", "glow", "numberformat",
        // Chart properties
        "chartType", "title", "legend", "dataLabels", "labelPos", "labelFont",
        "axisFont", "axisTitle", "catTitle", "axisMin", "axisMax", "majorUnit", "minorUnit",
        "axisNumFmt", "axisVisible", "majorTickMark", "minorTickMark", "tickLabelPos",
        "axisPosition", "crosses", "crossesAt", "crossBetween", "axisOrientation", "logBase",
        "dispUnits", "labelOffset", "tickLabelSkip",
        "gridlines", "minorGridlines", "plotFill", "chartFill",
        "colors", "gradient", "gradients", "lineWidth", "lineDash",
        "marker", "markerSize", "transparency", "smooth", "showMarker",
        "scatterStyle", "radarStyle", "varyColors", "dispBlanksAs",
        "roundedCorners", "plotVisOnly", "trendline", "invertIfNeg", "explosion",
        "errBars", "gapWidth", "overlap", "secondaryAxis", "dataTable",
        "firstSliceAngle", "holeSize", "bubbleScale", "shape", "gapDepth",
        "dropLines", "hiLowLines", "upDownBars", "serLines",
        "plotArea.border", "chartArea.border", "legend.overlay",
        "plotArea.x", "plotArea.y", "plotArea.w", "plotArea.h",
        "title.x", "title.y", "title.w", "title.h",
        "legend.x", "legend.y", "legend.w", "legend.h",
        "datalabels.separator", "datalabels.numfmt", "leaderLines",
        "view3d", "categories", "data",
        "referenceLine", "refLine", "targetLine", "preset", "colorRule",
        "conditionalColor", "comboTypes", "axisLine",
    };

    internal static string? SuggestProperty(string input)
    {
        var (best, _, _) = SuggestPropertyWithDistance(input);
        return best;
    }

    /// <summary>
    /// Returns (bestMatch, distance, isUnique) where isUnique means no other candidate shares the same distance.
    /// </summary>
    internal static (string? Best, int Distance, bool IsUnique) SuggestPropertyWithDistance(string input)
    {
        // Strip help text suffix if present (e.g. "key (valid props: ...)")
        var rawInput = input.Contains(' ') ? input[..input.IndexOf(' ')] : input;
        var lower = rawInput.ToLowerInvariant();
        string? best = null;
        int bestDist = int.MaxValue;
        int bestCount = 0; // how many props share the best distance

        foreach (var prop in KnownProps)
        {
            var dist = LevenshteinDistance(lower, prop.ToLowerInvariant());
            if (dist > 0 && dist <= Math.Max(2, rawInput.Length / 3))
            {
                if (dist < bestDist)
                {
                    bestDist = dist;
                    best = prop;
                    bestCount = 1;
                }
                else if (dist == bestDist)
                {
                    bestCount++;
                }
            }
        }

        return best != null ? (best, bestDist, bestCount == 1) : (null, int.MaxValue, false);
    }

    internal static int LevenshteinDistance(string s, string t)
    {
        if (s.Length == 0) return t.Length;
        if (t.Length == 0) return s.Length;

        var d = new int[s.Length + 1, t.Length + 1];
        for (int i = 0; i <= s.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= t.Length; j++) d[0, j] = j;

        for (int i = 1; i <= s.Length; i++)
        {
            for (int j = 1; j <= t.Length; j++)
            {
                int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
            }
        }

        return d[s.Length, t.Length];
    }

    // ==================== PPT spatial info helpers ====================

    /// <summary>
    /// Check if a .docx file has document protection enforced.
    /// Returns 0 if no protection or if the path targets an editable element.
    /// Returns 1 with error output if the document is protected and the target is not an editable region.
    /// </summary>
    private static int CheckDocxProtection(string filePath, string path, bool json)
    {
        try
        {
            using var handler = DocumentHandlerFactory.Open(filePath, editable: false);
            var root = handler.Get("/");
            var protection = root.Format.TryGetValue("protection", out var pVal) ? pVal?.ToString() : "none";
            var enforced = root.Format.TryGetValue("protectionEnforced", out var eVal) && eVal is true;

            if (!enforced || protection == "none")
                return 0;

            // Allow writes to formfield and SDT paths (they handle their own editable check)
            if (path.StartsWith("/formfield[", StringComparison.OrdinalIgnoreCase))
                return 0;
            if (path.Contains("/sdt[", StringComparison.OrdinalIgnoreCase))
                return 0;

            // Document is protected — block the write
            var msg = $"Document is protected (mode: {protection}). " +
                      "Use Query(\"editable\") to find editable fields, or use --force to override protection.";
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg, new List<OfficeCli.Core.CliWarning>()));
            else
                Console.Error.WriteLine($"ERROR: {msg}");
            return 1;
        }
        catch
        {
            // If we can't read protection info, allow the write to proceed
            return 0;
        }
    }

    private static readonly HashSet<string> PositionKeys = new(StringComparer.OrdinalIgnoreCase)
        { "x", "left", "y", "top", "width", "w", "height", "h" };

    /// <summary>
    /// For PPT spatial elements, return coordinate string like "x: 0cm  y: 5cm  width: 33.87cm  height: 5cm".
    /// Returns null for non-spatial elements (slide, Word, Excel).
    /// </summary>
    private static string? GetPptSpatialLine(IDocumentHandler handler, string path)
    {
        if (handler is not OfficeCli.Handlers.PowerPointHandler) return null;
        try
        {
            var node = handler.Get(path);
            if (node == null) return null;
            // Only for spatial types (shape, textbox, picture, table, chart, connector, group, equation)
            if (node.Type is "slide" or "paragraph" or "run" or "cell" or "row") return null;
            if (!node.Format.ContainsKey("x") || !node.Format.ContainsKey("y")) return null;
            var x = node.Format.TryGetValue("x", out var xv) ? xv : "?";
            var y = node.Format.TryGetValue("y", out var yv) ? yv : "?";
            var w = node.Format.TryGetValue("width", out var wv) ? wv : "?";
            var h = node.Format.TryGetValue("height", out var hv) ? hv : "?";
            return $"x: {x}  y: {y}  width: {w}  height: {h}";
        }
        catch { return null; }
    }

    /// <summary>
    /// Check if the element at <paramref name="path"/> has the same (x,y) as any sibling.
    /// Returns list of overlapping sibling names, or empty.
    /// </summary>
    private static List<string> CheckPositionOverlap(IDocumentHandler handler, string path)
    {
        var overlaps = new List<string>();
        if (handler is not OfficeCli.Handlers.PowerPointHandler) return overlaps;
        try
        {
            var node = handler.Get(path);
            if (node == null || !node.Format.ContainsKey("x") || !node.Format.ContainsKey("y")) return overlaps;
            var myX = node.Format["x"]?.ToString();
            var myY = node.Format["y"]?.ToString();
            if (myX == null || myY == null) return overlaps;

            // Get parent (slide) to enumerate siblings
            var slidePathMatch = System.Text.RegularExpressions.Regex.Match(path, @"^(/slide\[\d+\])");
            if (!slidePathMatch.Success) return overlaps;
            var slidePath = slidePathMatch.Value;
            var slideNode = handler.Get(slidePath);
            if (slideNode == null) return overlaps;

            foreach (var child in slideNode.Children)
            {
                if (child.Path == path) continue;
                if (!child.Format.ContainsKey("x") || !child.Format.ContainsKey("y")) continue;
                var cx = child.Format["x"]?.ToString();
                var cy = child.Format["y"]?.ToString();
                if (cx == myX && cy == myY)
                {
                    // Skip false positive: both shapes at default (0,0) means neither was explicitly positioned
                    if (myX == "0cm" && myY == "0cm" && cx == "0cm" && cy == "0cm") continue;
                    var name = child.Format.TryGetValue("name", out var n) ? n?.ToString() : child.Path;
                    overlaps.Add(name ?? child.Path);
                }
            }
        }
        catch { /* ignore */ }
        return overlaps;
    }

    /// <summary>
    /// Check if a shape's text overflows its bounds using CJK-aware character measurement.
    /// Returns a warning message or null.
    /// </summary>
    private static string? CheckTextOverflow(IDocumentHandler handler, string path)
    {
        if (handler is not OfficeCli.Handlers.PowerPointHandler pptHandler) return null;
        try
        {
            return pptHandler.CheckShapeTextOverflow(path);
        }
        catch { return null; }
    }

    /// <summary>
    /// Notify watch server with pre-rendered HTML from the handler.
    /// Call this while the handler is still open (before Dispose).
    /// </summary>
    private static void NotifyWatch(IDocumentHandler handler, string filePath, string? changedPath)
    {
        if (handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            string? scrollTo = null;
            var sheetName = WatchMessage.ExtractSheetName(changedPath);
            if (sheetName != null)
            {
                var idx = excel.GetSheetIndex(sheetName);
                if (idx >= 0) scrollTo = $".sheet-content[data-sheet=\"{idx}\"]";
            }
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = excel.ViewAsHtml(), ScrollTo = scrollTo });
            return;
        }
        if (handler is OfficeCli.Handlers.WordHandler word)
        {
            var scrollTo = WatchMessage.ExtractWordScrollTarget(changedPath);
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = word.ViewAsHtml(), ScrollTo = scrollTo });
            return;
        }
        if (handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var slideNum = WatchMessage.ExtractSlideNum(changedPath);
        if (slideNum > 0)
        {
            var html = ppt.RenderSlideHtml(slideNum);
            if (html != null)
            {
                WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "replace", Slide = slideNum, Html = html, FullHtml = ppt.ViewAsHtml() });
                return;
            }
        }
        WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = ppt.ViewAsHtml() });
    }

    private static void NotifyWatchRoot(IDocumentHandler handler, string filePath, int oldSlideCount)
    {
        if (handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = excel.ViewAsHtml() });
            return;
        }
        if (handler is OfficeCli.Handlers.WordHandler word)
        {
            // Scroll to last page (new content is typically appended)
            var html = word.ViewAsHtml();
            var pageCount = System.Text.RegularExpressions.Regex.Matches(html, @"data-page=""\d+""").Count;
            var scrollTo = pageCount > 0 ? $".page[data-page=\"{pageCount}\"]" : null;
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = html, ScrollTo = scrollTo });
            return;
        }
        if (handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var newCount = ppt.GetSlideCount();
        if (newCount > oldSlideCount)
        {
            var html = ppt.RenderSlideHtml(newCount);
            if (html != null)
            {
                WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "add", Slide = newCount, Html = html, FullHtml = ppt.ViewAsHtml() });
                return;
            }
        }
        else if (newCount < oldSlideCount)
        {
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "remove", Slide = oldSlideCount, FullHtml = ppt.ViewAsHtml() });
            return;
        }
        WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = ppt.ViewAsHtml() });
    }

    /// <summary>
    /// TextWriter that writes to two targets simultaneously (tee pattern).
    /// </summary>
    private class TeeWriter : TextWriter
    {
        private readonly TextWriter _a;
        private readonly TextWriter _b;
        public TeeWriter(TextWriter a, TextWriter b) { _a = a; _b = b; }
        public override Encoding Encoding => _a.Encoding;
        public override void Write(char value) { _a.Write(value); _b.Write(value); }
        public override void Write(string? value) { _a.Write(value); _b.Write(value); }
        public override void WriteLine(string? value) { _a.WriteLine(value); _b.WriteLine(value); }
        public override void Flush() { _a.Flush(); _b.Flush(); }
    }
}
