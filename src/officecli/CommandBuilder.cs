// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli;

static class CommandBuilder
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
                return;
            }

            // Fork a background process running the resident server
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
                    var msg = $"Opened {file.Name} (remember to call close when done)";
                    if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                    else Console.WriteLine(msg);
                    return;
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
        }, json); });

        rootCommand.Add(closeCommand);

        // ==================== watch command ====================
        var watchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx)" };
        var watchPortOpt = new Option<int>("--port") { Description = "HTTP port for preview server" };
        watchPortOpt.DefaultValueFactory = _ => 18080;

        var watchCommand = new Command("watch", "Start a live preview server that auto-refreshes when the document changes");
        watchCommand.Add(watchFileArg);
        watchCommand.Add(watchPortOpt);

        watchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(watchFileArg)!;
            var port = result.GetValue(watchPortOpt);

            // Render initial HTML from existing file content
            string? initialHtml = null;
            if (file.Exists)
            {
                try
                {
                    using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
                    if (handler is OfficeCli.Handlers.PowerPointHandler ppt)
                        initialHtml = ppt.ViewAsHtml();
                }
                catch { /* ignore — will show waiting page */ }
            }

            using var cts = new CancellationTokenSource();
            Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

            using var watch = new WatchServer(file.FullName, port, initialHtml: initialHtml);
            watch.RunAsync(cts.Token).GetAwaiter().GetResult();
        }));

        rootCommand.Add(watchCommand);

        // ==================== unwatch command ====================
        var unwatchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx)" };
        var unwatchCommand = new Command("unwatch", "Stop the watch preview server for the document");
        unwatchCommand.Add(unwatchFileArg);

        unwatchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(unwatchFileArg)!;
            if (WatchNotifier.SendClose(file.FullName))
                Console.WriteLine($"Watch stopped for {file.Name}");
            else
                Console.Error.WriteLine($"No watch running for {file.Name}");
        }));

        rootCommand.Add(unwatchCommand);

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

        // ==================== view command ====================
        var viewFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.docx, .xlsx, .pptx)" };
        var viewModeArg = new Argument<string>("mode") { Description = "View mode: text, annotated, outline, stats, issues, html" };
        var startLineOpt = new Option<int?>("--start") { Description = "Start line/paragraph number" };
        var endLineOpt = new Option<int?>("--end") { Description = "End line/paragraph number" };
        var maxLinesOpt = new Option<int?>("--max-lines") { Description = "Maximum number of lines/rows/slides to output (truncates with total count)" };
        var issueTypeOpt = new Option<string?>("--type") { Description = "Issue type filter: format, content, structure" };
        var limitOpt = new Option<int?>("--limit") { Description = "Limit number of results" };

        var colsOpt = new Option<string?>("--cols") { Description = "Column filter, comma-separated (Excel only, e.g. A,B,C)" };
        var browserOpt = new Option<bool>("--browser") { Description = "Open HTML output in browser (html mode only)" };

        var viewCommand = new Command("view", "View document in different modes");
        viewCommand.Add(viewFileArg);
        viewCommand.Add(viewModeArg);
        viewCommand.Add(startLineOpt);
        viewCommand.Add(endLineOpt);
        viewCommand.Add(maxLinesOpt);
        viewCommand.Add(issueTypeOpt);
        viewCommand.Add(limitOpt);
        viewCommand.Add(colsOpt);
        viewCommand.Add(browserOpt);
        viewCommand.Add(jsonOption);

        viewCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(viewFileArg)!;
            var mode = result.GetValue(viewModeArg)!;
            var start = result.GetValue(startLineOpt);
            var end = result.GetValue(endLineOpt);
            var maxLines = result.GetValue(maxLinesOpt);
            var issueType = result.GetValue(issueTypeOpt);
            var limit = result.GetValue(limitOpt);
            var colsStr = result.GetValue(colsOpt);
            var browser = result.GetValue(browserOpt);

            // Try resident first
            if (TryResident(file.FullName, req =>
            {
                req.Command = "view";
                req.Json = json;
                req.Args["mode"] = mode;
                if (start.HasValue) req.Args["start"] = start.Value.ToString();
                if (end.HasValue) req.Args["end"] = end.Value.ToString();
                if (maxLines.HasValue) req.Args["max-lines"] = maxLines.Value.ToString();
                if (issueType != null) req.Args["type"] = issueType;
                if (limit.HasValue) req.Args["limit"] = limit.Value.ToString();
                if (colsStr != null) req.Args["cols"] = colsStr;
                if (browser) req.Args["browser"] = "true";
            }, json)) return;

            var format = json ? OutputFormat.Json : OutputFormat.Text;
            var cols = colsStr != null ? new HashSet<string>(colsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);

            if (mode.ToLowerInvariant() is "html" or "h")
            {
                if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                {
                    var html = pptHandler.ViewAsHtml(start, end);

                    if (browser)
                    {
                        // --browser: write to temp file and open in browser
                        var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}.html");
                        File.WriteAllText(htmlPath, html);
                        Console.WriteLine(htmlPath);
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* silently ignore if browser can't be opened */ }
                    }
                    else
                    {
                        // Default: output HTML to stdout
                        Console.Write(html);
                    }
                }
                else
                {
                    throw new OfficeCli.Core.CliException("HTML preview is only supported for .pptx files.")
                    {
                        Code = "unsupported_type",
                        Suggestion = "Use a .pptx file, or use mode 'text' or 'annotated' for other formats.",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues"]
                    };
                }
                return;
            }

            if (json)
            {
                // Structured JSON output — no Content string wrapping
                var modeKey = mode.ToLowerInvariant();
                if (modeKey is "stats" or "s")
                    Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsStatsJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
                else if (modeKey is "outline" or "o")
                    Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsOutlineJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
                else if (modeKey is "text" or "t")
                    Console.WriteLine(OutputFormatter.WrapEnvelope(handler.ViewAsTextJson(start, end, maxLines, cols).ToJsonString(OutputFormatter.PublicJsonOptions)));
                else if (modeKey is "annotated" or "a")
                    Console.WriteLine(OutputFormatter.WrapEnvelope(
                        OutputFormatter.FormatView(mode, handler.ViewAsAnnotated(start, end, maxLines, cols), OutputFormat.Json)));
                else if (modeKey is "issues" or "i")
                    Console.WriteLine(OutputFormatter.WrapEnvelope(
                        OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, limit), OutputFormat.Json)));
                else
                    throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html"]
                    };
            }
            else
            {
                var output = mode.ToLowerInvariant() switch
                {
                    "text" or "t" => handler.ViewAsText(start, end, maxLines, cols),
                    "annotated" or "a" => handler.ViewAsAnnotated(start, end, maxLines, cols),
                    "outline" or "o" => handler.ViewAsOutline(),
                    "stats" or "s" => handler.ViewAsStats(),
                    "issues" or "i" => OutputFormatter.FormatIssues(handler.ViewAsIssues(issueType, limit), OutputFormat.Text),
                    _ => throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html"]
                    }
                };
                Console.WriteLine(output);
            }
        }, json); });

        rootCommand.Add(viewCommand);

        // ==================== get command ====================
        var getFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var pathArg = new Argument<string>("path") { Description = "DOM path (e.g. /body/p[1])" };
        pathArg.DefaultValueFactory = _ => "/";
        var depthOpt = new Option<int>("--depth") { Description = "Depth of child nodes to include" };
        depthOpt.DefaultValueFactory = _ => 1;

        var getCommand = new Command("get", "Get a document node by path");
        getCommand.Add(getFileArg);
        getCommand.Add(pathArg);
        getCommand.Add(depthOpt);
        getCommand.Add(jsonOption);

        getCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(getFileArg)!;
            var path = result.GetValue(pathArg)!;
            var depth = result.GetValue(depthOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "get";
                req.Json = json;
                req.Args["path"] = path;
                req.Args["depth"] = depth.ToString();
            }, json)) return;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var node = handler.Get(path, depth);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNode(node, OutputFormat.Json)));
            else
                Console.WriteLine(OutputFormatter.FormatNode(node, OutputFormat.Text));
        }, json); });

        rootCommand.Add(getCommand);

        // ==================== query command ====================
        var queryFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var selectorArg = new Argument<string>("selector") { Description = "CSS-like selector (e.g. paragraph[style=Normal] > run[font!=Arial])" };

        var queryCommand = new Command("query", "Query document elements with CSS-like selectors");
        queryCommand.Add(queryFileArg);
        queryCommand.Add(selectorArg);
        queryCommand.Add(jsonOption);

        queryCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(queryFileArg)!;
            var selector = result.GetValue(selectorArg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "query";
                req.Json = json;
                req.Args["selector"] = selector;
            }, json)) return;

            var format = json ? OutputFormat.Json : OutputFormat.Text;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var filters = OfficeCli.Core.AttributeFilter.Parse(selector);
            var (results, warnings) = OfficeCli.Core.AttributeFilter.ApplyWithWarnings(handler.Query(selector), filters);
            if (json)
            {
                var cliWarnings = warnings.Select(w => new OfficeCli.Core.CliWarning { Message = w, Code = "filter_warning" }).ToList();
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNodes(results, OutputFormat.Json),
                    cliWarnings.Count > 0 ? cliWarnings : null));
            }
            else
            {
                foreach (var w in warnings) Console.Error.WriteLine(w);
                Console.WriteLine(OutputFormatter.FormatNodes(results, OutputFormat.Text));
            }
        }, json); });

        rootCommand.Add(queryCommand);

        // ==================== set command ====================
        var setFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var setPathArg = new Argument<string>("path") { Description = "DOM path to the element" };
        var propsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value)", AllowMultipleArgumentsPerToken = true };

        var setCommand = new Command("set", "Modify a document node's properties");
        setCommand.Add(setFileArg);
        setCommand.Add(setPathArg);
        setCommand.Add(propsOpt);
        setCommand.Add(jsonOption);

        setCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(setFileArg)!;
            var path = result.GetValue(setPathArg)!;
            var props = result.GetValue(propsOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "set";
                req.Args["path"] = path;
                req.Props = props;
            }, json)) { return; }

            var properties = new Dictionary<string, string>();
            foreach (var prop in props ?? Array.Empty<string>())
            {
                var eqIdx = prop.IndexOf('=');
                if (eqIdx > 0)
                {
                    properties[prop[..eqIdx]] = prop[(eqIdx + 1)..];
                }
            }

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var unsupported = handler.Set(path, properties);
            // unsupported entries may contain help text like "key (valid props: ...)" — extract raw keys
            var unsupportedKeys = unsupported.Select(u => u.Contains(' ') ? u[..u.IndexOf(' ')] : u).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var applied = properties.Where(kv => !unsupportedKeys.Contains(kv.Key)).ToList();
            var message = applied.Count > 0
                ? $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
                : $"No properties applied to {path}";
            if (json)
            {
                var warnings = unsupported.Count > 0
                    ? unsupported.Select(p =>
                    {
                        var suggestion = SuggestProperty(p);
                        return new OfficeCli.Core.CliWarning
                        {
                            Message = suggestion != null ? $"Unsupported property: {p} (did you mean: {suggestion}?)" : $"Unsupported property: {p}",
                            Code = "unsupported_property",
                            Suggestion = suggestion
                        };
                    }).ToList()
                    : null;
                bool allFailed = applied.Count == 0 && unsupported.Count > 0;
                Console.WriteLine(allFailed
                    ? OutputFormatter.WrapEnvelopeError(message, warnings)
                    : OutputFormatter.WrapEnvelopeText(message, warnings));
            }
            else
            {
                Console.WriteLine(message);
                if (unsupported.Count > 0)
                    Console.Error.WriteLine(FormatUnsupported(unsupported));
            }
            NotifyWatch(handler, file.FullName, path);
        }, json); });

        rootCommand.Add(setCommand);

        // ==================== add command ====================
        var addFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var addParentPathArg = new Argument<string>("parent") { Description = "Parent DOM path (e.g. /body, /Sheet1, /slide[1])" };
        var addTypeOpt = new Option<string>("--type") { Description = "Element type to add (e.g. paragraph, run, table, sheet, row, cell, slide, shape)" };
        var addFromOpt = new Option<string?>("--from") { Description = "Copy from an existing element path (e.g. /slide[1]/shape[2])" };
        var addIndexOpt = new Option<int?>("--index") { Description = "Insert position (0-based). If omitted, appends to end" };
        var addPropsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value)", AllowMultipleArgumentsPerToken = true };

        var addCommand = new Command("add", "Add a new element to the document");
        addCommand.Add(addFileArg);
        addCommand.Add(addParentPathArg);
        addCommand.Add(addTypeOpt);
        addCommand.Add(addFromOpt);
        addCommand.Add(addIndexOpt);
        addCommand.Add(addPropsOpt);
        addCommand.Add(jsonOption);

        addCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(addFileArg)!;
            var parentPath = result.GetValue(addParentPathArg)!;
            var type = result.GetValue(addTypeOpt);
            var from = result.GetValue(addFromOpt);
            var index = result.GetValue(addIndexOpt);
            var props = result.GetValue(addPropsOpt);

            if (string.IsNullOrEmpty(type) && string.IsNullOrEmpty(from))
            {
                throw new OfficeCli.Core.CliException("Either --type or --from must be specified.")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --type to specify element type, or --from to copy an existing element.",
                    Help = "officecli add <file> <parent> --type <type> --prop key=value"
                };
            }

            if (!string.IsNullOrEmpty(from))
            {
                // Copy from existing element
                if (TryResident(file.FullName, req =>
                {
                    req.Command = "add";
                    req.Args["parent"] = parentPath;
                    req.Args["from"] = from;
                    if (index.HasValue) req.Args["index"] = index.Value.ToString();
                }, json)) { return; }

                using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
                var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
                var resultPath = handler.CopyFrom(from, parentPath, index);
                var message = $"Copied to {resultPath}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
                else Console.WriteLine(message);
                if (parentPath == "/") NotifyWatchRoot(handler, file.FullName, oldCount);
                else NotifyWatch(handler, file.FullName, parentPath);
            }
            else
            {
                if (TryResident(file.FullName, req =>
                {
                    req.Command = "add";
                    req.Args["parent"] = parentPath;
                    req.Args["type"] = type!;
                    if (index.HasValue) req.Args["index"] = index.Value.ToString();
                    req.Props = props;
                }, json)) { return; }

                var properties = new Dictionary<string, string>();
                foreach (var prop in props ?? Array.Empty<string>())
                {
                    var eqIdx = prop.IndexOf('=');
                    if (eqIdx > 0)
                    {
                        properties[prop[..eqIdx]] = prop[(eqIdx + 1)..];
                    }
                }

                using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
                var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
                var resultPath = handler.Add(parentPath, type!, index, properties);
                var message = $"Added {type} at {resultPath}";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
                else Console.WriteLine(message);
                if (parentPath == "/") NotifyWatchRoot(handler, file.FullName, oldCount);
                else NotifyWatch(handler, file.FullName, parentPath);
            }
        }, json); });

        rootCommand.Add(addCommand);

        // ==================== remove command ====================
        var removeFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var removePathArg = new Argument<string>("path") { Description = "DOM path of the element to remove" };

        var removeCommand = new Command("remove", "Remove an element from the document");
        removeCommand.Add(removeFileArg);
        removeCommand.Add(removePathArg);
        removeCommand.Add(jsonOption);

        removeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(removeFileArg)!;
            var path = result.GetValue(removePathArg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "remove";
                req.Args["path"] = path;
            }, json)) { return; }

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var oldCount = (handler as OfficeCli.Handlers.PowerPointHandler)?.GetSlideCount() ?? 0;
            var warning = handler.Remove(path);
            var message = $"Removed {path}";
            if (warning != null) message += $"\n{warning}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            var slideNum = WatchMessage.ExtractSlideNum(path);
            if (slideNum > 0 && !path.Contains("/shape["))
                NotifyWatchRoot(handler, file.FullName, oldCount);
            else
                NotifyWatch(handler, file.FullName, path);
        }, json); });

        rootCommand.Add(removeCommand);

        // ==================== move command ====================
        var moveFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var movePathArg = new Argument<string>("path") { Description = "DOM path of the element to move" };
        var moveToOpt = new Option<string?>("--to") { Description = "Target parent path. If omitted, reorders within the current parent" };
        var moveIndexOpt = new Option<int?>("--index") { Description = "Insert position (0-based). If omitted, appends to end" };

        var moveCommand = new Command("move", "Move an element to a new position or parent");
        moveCommand.Add(moveFileArg);
        moveCommand.Add(movePathArg);
        moveCommand.Add(moveToOpt);
        moveCommand.Add(moveIndexOpt);
        moveCommand.Add(jsonOption);

        moveCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(moveFileArg)!;
            var path = result.GetValue(movePathArg)!;
            var to = result.GetValue(moveToOpt);
            var index = result.GetValue(moveIndexOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "move";
                req.Args["path"] = path;
                if (to != null) req.Args["to"] = to;
                if (index.HasValue) req.Args["index"] = index.Value.ToString();
            }, json)) { return; }

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var resultPath = handler.Move(path, to, index);
            var message = $"Moved to {resultPath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            NotifyWatch(handler, file.FullName, path);
        }, json); });

        rootCommand.Add(moveCommand);

        // ==================== raw command ====================
        var rawFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var rawPathArg = new Argument<string>("part") { Description = "Part path (e.g. /document, /styles, /header[0])" };
        rawPathArg.DefaultValueFactory = _ => "/document";

        var rawStartOpt = new Option<int?>("--start") { Description = "Start row number (Excel sheets only)" };
        var rawEndOpt = new Option<int?>("--end") { Description = "End row number (Excel sheets only)" };

        var rawColsOpt = new Option<string?>("--cols") { Description = "Column filter, comma-separated (Excel only, e.g. A,B,C)" };

        var rawCommand = new Command("raw", "View raw XML of a document part");
        rawCommand.Add(rawFileArg);
        rawCommand.Add(rawPathArg);
        rawCommand.Add(rawStartOpt);
        rawCommand.Add(rawEndOpt);
        rawCommand.Add(rawColsOpt);
        rawCommand.Add(jsonOption);

        rawCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(rawFileArg)!;
            var partPath = result.GetValue(rawPathArg)!;
            var startRow = result.GetValue(rawStartOpt);
            var endRow = result.GetValue(rawEndOpt);
            var rawColsStr = result.GetValue(rawColsOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "raw";
                req.Args["part"] = partPath;
                if (startRow.HasValue) req.Args["start"] = startRow.Value.ToString();
                if (endRow.HasValue) req.Args["end"] = endRow.Value.ToString();
                if (rawColsStr != null) req.Args["cols"] = rawColsStr;
            }, json)) return;

            var rawCols = rawColsStr != null ? new HashSet<string>(rawColsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var xml = handler.Raw(partPath, startRow, endRow, rawCols);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(xml));
            else Console.WriteLine(xml);
        }, json); });

        rootCommand.Add(rawCommand);

        // ==================== raw-set command ====================
        var rawSetFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var rawSetPartArg = new Argument<string>("part") { Description = "Part path (e.g. /document, /styles, /Sheet1, /slide[1])" };
        var rawSetXpathOpt = new Option<string>("--xpath") { Description = "XPath to target element(s)", Required = true };
        var rawSetActionOpt = new Option<string>("--action") { Description = "Action: append, prepend, insertbefore, insertafter, replace, remove, setattr", Required = true };
        var rawSetXmlOpt = new Option<string?>("--xml") { Description = "XML fragment or attr=value for setattr" };

        var rawSetCommand = new Command("raw-set", "Modify raw XML in a document part (universal fallback for any OpenXML operation)");
        rawSetCommand.Add(rawSetFileArg);
        rawSetCommand.Add(rawSetPartArg);
        rawSetCommand.Add(rawSetXpathOpt);
        rawSetCommand.Add(rawSetActionOpt);
        rawSetCommand.Add(rawSetXmlOpt);
        rawSetCommand.Add(jsonOption);

        rawSetCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(rawSetFileArg)!;
            var partPath = result.GetValue(rawSetPartArg)!;
            var xpath = result.GetValue(rawSetXpathOpt)!;
            var action = result.GetValue(rawSetActionOpt)!;
            var xml = result.GetValue(rawSetXmlOpt);

            if (TryResident(file.FullName, req =>
            {
                req.Command = "raw-set";
                req.Args["part"] = partPath;
                req.Args["xpath"] = xpath;
                req.Args["action"] = action;
                if (xml != null) req.Args["xml"] = xml;
            }, json)) { return; }

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var errorsBefore = handler.Validate().Select(e => e.Description).ToHashSet();
            handler.RawSet(partPath, xpath, action, xml);
            var warnings = ReportNewErrorsAsWarnings(handler, errorsBefore);
            var message = $"raw-set applied: {action} at {xpath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message, warnings));
            else
            {
                Console.WriteLine(message);
                ReportNewErrors(handler, errorsBefore, warnings);
            }
            NotifyWatch(handler, file.FullName, null);
        }, json); });

        rootCommand.Add(rawSetCommand);

        // ==================== add-part command ====================
        var addPartFileArg = new Argument<string>("file") { Description = "Document file path" };
        var addPartParentArg = new Argument<string>("parent") { Description = "Parent part path (e.g. / for document root, /Sheet1 for Excel sheet, /slide[0] for PPT slide)" };
        var addPartTypeOpt = new Option<string>("--type") { Description = "Part type to create (chart, header, footer)", Required = true };
        var addPartCommand = new Command("add-part", "Create a new document part and return its relationship ID for use with raw-set");
        addPartCommand.Add(addPartFileArg);
        addPartCommand.Add(addPartParentArg);
        addPartCommand.Add(addPartTypeOpt);
        addPartCommand.Add(jsonOption);

        addPartCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(addPartFileArg)!;
            var parent = result.GetValue(addPartParentArg)!;
            var type = result.GetValue(addPartTypeOpt)!;

            if (TryResident(file, req =>
            {
                req.Command = "add-part";
                req.Args["parent"] = parent;
                req.Args["type"] = type;
            }, json)) { return; }

            using var handler = DocumentHandlerFactory.Open(file, editable: true);
            var errorsBefore = handler.Validate().Select(e => e.Description).ToHashSet();
            var (relId, partPath) = handler.AddPart(parent, type);
            var warnings = ReportNewErrorsAsWarnings(handler, errorsBefore);
            var message = $"Created {type} part: relId={relId} path={partPath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message, warnings));
            else
            {
                Console.WriteLine(message);
                ReportNewErrors(handler, errorsBefore, warnings);
            }
            NotifyWatch(handler, file, null);
        }, json); });

        rootCommand.Add(addPartCommand);

        // ==================== validate command ====================
        var validateFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var validateCommand = new Command("validate", "Validate document against OpenXML schema");
        validateCommand.Add(validateFileArg);
        validateCommand.Add(jsonOption);
        validateCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(validateFileArg)!;

            if (TryResident(file.FullName, req =>
            {
                req.Command = "validate";
                req.Json = json;
            }, json)) return;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var errors = handler.Validate();
            if (json)
            {
                var validationJson = FormatValidationErrors(errors);
                Console.WriteLine(OutputFormatter.WrapEnvelope(validationJson));
            }
            else
            {
                if (errors.Count == 0)
                {
                    Console.WriteLine("Validation passed: no errors found.");
                }
                else
                {
                    Console.WriteLine($"Found {errors.Count} validation error(s):");
                    foreach (var err in errors)
                    {
                        Console.WriteLine($"  [{err.ErrorType}] {err.Description}");
                        if (err.Path != null) Console.WriteLine($"    Path: {err.Path}");
                        if (err.Part != null) Console.WriteLine($"    Part: {err.Part}");
                    }
                }
            }
        }, json); });
        rootCommand.Add(validateCommand);

        // ==================== batch command ====================
        var batchFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var batchInputOpt = new Option<FileInfo?>("--input") { Description = "JSON file containing batch commands. If omitted, reads from stdin" };
        var batchStopOnErrorOpt = new Option<bool>("--stop-on-error") { Description = "Stop execution on first error (default: continue all)" };
        var batchCommand = new Command("batch", "Execute multiple commands from a JSON array (one open/save cycle)");
        batchCommand.Add(batchFileArg);
        batchCommand.Add(batchInputOpt);
        batchCommand.Add(batchStopOnErrorOpt);
        batchCommand.Add(jsonOption);

        batchCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(batchFileArg)!;
            var inputFile = result.GetValue(batchInputOpt);
            var stopOnError = result.GetValue(batchStopOnErrorOpt);

            string jsonText;
            if (inputFile != null)
            {
                if (!inputFile.Exists)
                {
                    throw new FileNotFoundException($"Input file not found: {inputFile.FullName}");
                }
                jsonText = File.ReadAllText(inputFile.FullName);
            }
            else
            {
                // Read from stdin
                jsonText = Console.In.ReadToEnd();
            }

            var items = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(jsonText, BatchJsonContext.Default.ListBatchItem);
            if (items == null || items.Count == 0)
            {
                throw new ArgumentException("No commands found in input.");
            }

            // If a resident process is running, forward each command to it
            if (ResidentClient.TryConnect(file.FullName, out _))
            {
                var results = new List<BatchResult>();
                foreach (var item in items)
                {
                    var req = item.ToResidentRequest();
                    req.Json = json;
                    var response = ResidentClient.TrySend(file.FullName, req);
                    if (response == null)
                    {
                        results.Add(new BatchResult { Success = false, Error = "Failed to send to resident" });
                        if (stopOnError) break;
                        continue;
                    }
                    var success = string.IsNullOrEmpty(response.Stderr);
                    results.Add(new BatchResult { Success = success, Output = response.Stdout, Error = response.Stderr });
                    if (!success && stopOnError) break;
                }
                PrintBatchResults(results, json);
                if (results.Any(r => !r.Success))
                    throw new InvalidOperationException($"Batch completed with {results.Count(r => !r.Success)} error(s)");
                return;
            }

            // Non-resident: open file once, execute all commands, save once
            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var batchResults = new List<BatchResult>();
            foreach (var item in items)
            {
                try
                {
                    var output = ExecuteBatchItem(handler, item, json);
                    batchResults.Add(new BatchResult { Success = true, Output = output });
                }
                catch (Exception ex)
                {
                    batchResults.Add(new BatchResult { Success = false, Error = ex.Message });
                    if (stopOnError) break;
                }
            }
            PrintBatchResults(batchResults, json);
            if (batchResults.Any(r => r.Success))
                NotifyWatch(handler, file.FullName, null);
            if (batchResults.Any(r => !r.Success))
                throw new InvalidOperationException($"Batch completed with {batchResults.Count(r => !r.Success)} error(s)");
        }, json); });

        rootCommand.Add(batchCommand);

        // ==================== create command ====================
        var createFileArg = new Argument<string>("file") { Description = "Output file path (.docx, .xlsx, .pptx)" };
        var createTypeOpt = new Option<string>("--type") { Description = "Document type (docx, xlsx, pptx) — optional, inferred from file extension" };
        var createCommand = new Command("create", "Create a blank Office document");
        createCommand.Add(createFileArg);
        createCommand.Add(createTypeOpt);
        createCommand.Add(jsonOption);

        createCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(createFileArg)!;
            var type = result.GetValue(createTypeOpt);

            // If file has no extension but --type is provided, append it
            if (!string.IsNullOrEmpty(type) && string.IsNullOrEmpty(Path.GetExtension(file)))
            {
                var ext = type.StartsWith('.') ? type : "." + type;
                file += ext;
            }

            // Check if the file is held by a resident process
            var fullPath = Path.GetFullPath(file);
            if (ResidentClient.TryConnect(fullPath, out _))
            {
                throw new CliException($"{Path.GetFileName(file)} is currently opened by a resident process. Please run 'officecli close \"{file}\"' first.")
                {
                    Code = "file_locked",
                    Suggestion = $"Run: officecli close \"{file}\""
                };
            }

            OfficeCli.BlankDocCreator.Create(file);
            var fullCreatedPath = Path.GetFullPath(file);
            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelopeText($"Created: {fullCreatedPath}"));
            }
            else
            {
                Console.WriteLine($"Created: {file}");
                if (Path.GetExtension(file).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  totalSlides: 0");
                    Console.WriteLine($"  slideWidth: {Core.EmuConverter.FormatEmu(12192000)}");
                    Console.WriteLine($"  slideHeight: {Core.EmuConverter.FormatEmu(6858000)}");
                }
            }
        }, json); });

        rootCommand.Add(createCommand);

        HelpCommands.Register(rootCommand);

        return rootCommand;
    }

    // ==================== Helper: try forwarding to resident ====================
    internal static bool TryResident(string filePath, Action<ResidentRequest> configure, bool json = false)
    {
        var request = new ResidentRequest();
        configure(request);
        if (json) request.Json = true;

        var response = ResidentClient.TrySend(filePath, request);
        if (response == null)
            return false;

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

        return true;
    }

    internal static int SafeRun(Action action, bool json = false)
    {
        if (!OfficeCli.Core.CliLogger.Enabled)
        {
            try
            {
                action();
                return 0;
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
            action();
            var stdout = stdoutWriter.ToString().TrimEnd('\r', '\n');
            OfficeCli.Core.CliLogger.LogOutput(stdout);
            return 0;
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
                foreach (var w in warnings) Console.Error.WriteLine(w);
                return OfficeCli.Core.OutputFormatter.FormatNodes(results, format);
            }
            case "set":
            {
                var path = item.Path ?? "/";
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
                var parentPath = item.Parent ?? item.Path ?? "/";
                if (!string.IsNullOrEmpty(item.From))
                {
                    var resultPath = handler.CopyFrom(item.From, parentPath, item.Index);
                    return $"Copied to {resultPath}";
                }
                else
                {
                    var type = item.Type ?? "";
                    var resultPath = handler.Add(parentPath, type, item.Index, props);
                    return $"Added {type} at {resultPath}";
                }
            }
            case "remove":
            {
                var path = item.Path ?? "/";
                var warning = handler.Remove(path);
                var msg = $"Removed {path}";
                if (warning != null) msg += $"\n{warning}";
                return msg;
            }
            case "move":
            {
                var path = item.Path ?? "/";
                var resultPath = handler.Move(path, item.To, item.Index);
                return $"Moved to {resultPath}";
            }
            case "view":
            {
                var mode = item.Mode ?? "text";
                if (mode.ToLowerInvariant() is "html" or "h" && handler is OfficeCli.Handlers.PowerPointHandler pptH)
                {
                    return pptH.ViewAsHtml();
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
                var partPath = item.Part ?? "/document";
                return handler.Raw(partPath, null, null, null);
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
                throw new InvalidOperationException($"Unknown command: {item.Command}");
        }
    }

    internal static void PrintBatchResults(List<BatchResult> results, bool json)
    {
        if (json)
        {
            Console.WriteLine(System.Text.Json.JsonSerializer.Serialize(results, BatchJsonContext.Default.ListBatchResult));
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
                        Console.WriteLine($"{prefix}{r.Output}");
                    else
                        Console.WriteLine($"{prefix}OK");
                }
                else
                {
                    Console.Error.WriteLine($"{prefix}ERROR: {r.Error}");
                }
            }

            var succeeded = results.Count(r => r.Success);
            var failed = results.Count - succeeded;
            Console.WriteLine($"\nBatch complete: {succeeded} succeeded, {failed} failed, {results.Count} total");
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

    internal static string? SuggestProperty(string input)
    {
        var knownProps = new[]
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
            "font.strike", "font.underline", "tabColor", "shadow", "glow",
        };

        var lower = input.ToLowerInvariant();
        string? best = null;
        int bestDist = int.MaxValue;

        foreach (var prop in knownProps)
        {
            var dist = LevenshteinDistance(lower, prop.ToLowerInvariant());
            if (dist < bestDist && dist <= Math.Max(2, input.Length / 3))
            {
                bestDist = dist;
                best = prop;
            }
        }

        return best;
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

    /// <summary>
    /// Notify watch server with pre-rendered HTML from the handler.
    /// Call this while the handler is still open (before Dispose).
    /// </summary>
    private static void NotifyWatch(IDocumentHandler handler, string filePath, string? changedPath)
    {
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
