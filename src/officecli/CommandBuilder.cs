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
                return 0;
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

        // ==================== watch command ====================
        var watchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
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
                    else if (handler is OfficeCli.Handlers.ExcelHandler excel)
                        initialHtml = excel.ViewAsHtml();
                    else if (handler is OfficeCli.Handlers.WordHandler word)
                        initialHtml = word.ViewAsHtml();
                }
                catch { /* ignore — will show waiting page */ }
            }

            using var cts = new CancellationTokenSource();
            Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };

            using var watch = new WatchServer(file.FullName, port, initialHtml: initialHtml);
            watch.RunAsync(cts.Token).GetAwaiter().GetResult();
            return 0;
        }));

        rootCommand.Add(watchCommand);

        // ==================== unwatch command ====================
        var unwatchFileArg = new Argument<FileInfo>("file") { Description = "Office document path (.pptx, .xlsx, .docx)" };
        var unwatchCommand = new Command("unwatch", "Stop the watch preview server for the document");
        unwatchCommand.Add(unwatchFileArg);

        unwatchCommand.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(unwatchFileArg)!;
            if (WatchNotifier.SendClose(file.FullName))
                Console.WriteLine($"Watch stopped for {file.Name}");
            else
                Console.Error.WriteLine($"No watch running for {file.Name}");
            return 0;
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
        var viewModeArg = new Argument<string>("mode") { Description = "View mode: text, annotated, outline, stats, issues, html, svg, forms" };
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
            }, json) is {} rc) return rc;

            var format = json ? OutputFormat.Json : OutputFormat.Text;
            var cols = colsStr != null ? new HashSet<string>(colsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);

            if (mode.ToLowerInvariant() is "html" or "h")
            {
                string? html = null;
                if (handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                    html = pptHandler.ViewAsHtml(start, end);
                else if (handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                    html = excelHandler.ViewAsHtml();
                else if (handler is OfficeCli.Handlers.WordHandler wordHandler)
                    html = wordHandler.ViewAsHtml();

                if (html != null)
                {
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
                    throw new OfficeCli.Core.CliException("HTML preview is only supported for .pptx, .xlsx, and .docx files.")
                    {
                        Code = "unsupported_type",
                        Suggestion = "Use a .pptx, .xlsx, or .docx file, or use mode 'text' or 'annotated' for other formats.",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues"]
                    };
                }
                return 0;
            }

            if (mode.ToLowerInvariant() is "svg" or "g")
            {
                if (handler is OfficeCli.Handlers.PowerPointHandler pptSvgHandler)
                {
                    var slideNum = start ?? 1;
                    var svg = pptSvgHandler.ViewAsSvg(slideNum);

                    if (browser)
                    {
                        string outPath;
                        if (svg.Contains("data-formula"))
                        {
                            // Wrap SVG in HTML shell for KaTeX formula rendering
                            outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}.html");
                            var html = $"<!DOCTYPE html><html><head><meta charset='UTF-8'><link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css'><script defer src='https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js'></script><style>body{{margin:0;display:flex;justify-content:center;background:#f0f0f0}}</style></head><body>{svg}<script>window.addEventListener('load',function(){{document.querySelectorAll('[data-formula]').forEach(function(el){{try{{katex.render(el.getAttribute('data-formula'),el,{{throwOnError:false,displayMode:true}})}}catch(e){{}}}})}})</script></body></html>";
                            File.WriteAllText(outPath, html);
                        }
                        else
                        {
                            outPath = Path.Combine(Path.GetTempPath(), $"officecli_slide{slideNum}_{Path.GetFileNameWithoutExtension(file.Name)}_{DateTime.Now:HHmmss}.svg");
                            File.WriteAllText(outPath, svg);
                        }
                        Console.WriteLine(outPath);
                        try
                        {
                            var psi = new System.Diagnostics.ProcessStartInfo(outPath) { UseShellExecute = true };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch { /* silently ignore if browser can't be opened */ }
                    }
                    else
                    {
                        Console.Write(svg);
                    }
                }
                else
                {
                    throw new OfficeCli.Core.CliException("SVG preview is only supported for .pptx files.")
                    {
                        Code = "unsupported_type",
                        Suggestion = "Use a .pptx file, or use mode 'text' or 'annotated' for other formats.",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg"]
                    };
                }
                return 0;
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
                else if (modeKey is "forms" or "f")
                {
                    if (handler is OfficeCli.Handlers.WordHandler wordFormsHandler)
                        Console.WriteLine(OutputFormatter.WrapEnvelope(wordFormsHandler.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions)));
                    else
                        throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "forms"]
                        };
                }
                else
                    throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, forms")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "forms"]
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
                    "forms" or "f" => handler is OfficeCli.Handlers.WordHandler wfh
                        ? wfh.ViewAsForms()
                        : throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "forms"]
                        },
                    _ => throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, forms")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "forms"]
                    }
                };
                Console.WriteLine(output);
            }
            return 0;
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
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var node = handler.Get(path, depth);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelope(
                    OutputFormatter.FormatNode(node, OutputFormat.Json)));
            else
                Console.WriteLine(OutputFormatter.FormatNode(node, OutputFormat.Text));
            return 0;
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
            }, json) is {} rc) return rc;

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
            return 0;
        }, json); });

        rootCommand.Add(queryCommand);

        // ==================== set command ====================
        var forceOption = new Option<bool>("--force") { Description = "Force write even if document is protected" };
        var setFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var setPathArg = new Argument<string>("path") { Description = "DOM path to the element" };
        var propsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value)", AllowMultipleArgumentsPerToken = true };

        var setCommand = new Command("set", "Modify a document node's properties") { TreatUnmatchedTokensAsErrors = false };
        setCommand.Add(setFileArg);
        setCommand.Add(setPathArg);
        setCommand.Add(propsOpt);
        setCommand.Add(jsonOption);
        setCommand.Add(forceOption);

        setCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(setFileArg)!;
            var path = result.GetValue(setPathArg)!;
            var props = result.GetValue(propsOpt);
            var force = result.GetValue(forceOption);

            // Check document protection for .docx files
            if (!force && file.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                var protectionError = CheckDocxProtection(file.FullName, path, json);
                if (protectionError != 0) return protectionError;
            }

            // Detect bare key=value positional arguments (missing --prop)
            var unmatchedKvWarnings = DetectUnmatchedKeyValues(result);
            if (unmatchedKvWarnings.Count > 0)
            {
                if (json)
                {
                    var kvWarnings = unmatchedKvWarnings.Select(kv => new OfficeCli.Core.CliWarning
                    {
                        Message = $"Bare property '{kv}' ignored. Use --prop {kv}",
                        Code = "missing_prop_flag",
                        Suggestion = $"--prop {kv}"
                    }).ToList();
                    Console.WriteLine(OutputFormatter.WrapEnvelopeError(
                        $"Properties specified without --prop flag. Use: officecli set <file> <path> --prop {string.Join(" --prop ", unmatchedKvWarnings)}",
                        kvWarnings));
                }
                else
                {
                    foreach (var kv in unmatchedKvWarnings)
                        Console.Error.WriteLine($"WARNING: Bare property '{kv}' ignored. Did you mean: --prop {kv}");
                    Console.Error.WriteLine("Hint: Properties must be passed with --prop flag, e.g. officecli set <file> <path> --prop key=value");
                }
                if (props == null || props.Length == 0)
                    return 2;
            }

            if (TryResident(file.FullName, req =>
            {
                req.Command = "set";
                req.Args["path"] = path;
                req.Props = props;
            }, json) is {} rc) return rc;

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

            // Auto-correct: attempt to fix unsupported properties with Levenshtein distance == 1
            var autoCorrected = new List<(string Original, string Corrected, string Value)>();
            var stillUnsupported = new List<string>();
            foreach (var u in unsupported)
            {
                var rawKey = u.Contains(' ') ? u[..u.IndexOf(' ')] : u;
                if (properties.TryGetValue(rawKey, out var val))
                {
                    var (suggestion, dist, isUnique) = SuggestPropertyWithDistance(rawKey);
                    if (suggestion != null && dist == 1 && isUnique)
                    {
                        // Auto-correct: re-apply with corrected key
                        var correctedProps = new Dictionary<string, string> { [suggestion] = val };
                        var retryUnsupported = handler.Set(path, correctedProps);
                        if (retryUnsupported.Count == 0)
                        {
                            autoCorrected.Add((rawKey, suggestion, val));
                            continue;
                        }
                    }
                }
                stillUnsupported.Add(u);
            }

            // unsupported entries may contain help text like "key (valid props: ...)" — extract raw keys
            var unsupportedKeys = stillUnsupported.Select(u => u.Contains(' ') ? u[..u.IndexOf(' ')] : u).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var autoCorrectedKeys = autoCorrected.Select(ac => ac.Original).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var applied = properties.Where(kv => !unsupportedKeys.Contains(kv.Key) && !autoCorrectedKeys.Contains(kv.Key)).ToList();
            // Include auto-corrected props in applied list with the corrected key name
            foreach (var ac in autoCorrected)
                applied.Add(new KeyValuePair<string, string>(ac.Corrected, ac.Value));

            var message = applied.Count > 0
                ? $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
                : $"No properties applied to {path}";

            // Check if position-related props were changed → show coordinates + overlap warning
            var positionChanged = applied.Any(kv => PositionKeys.Contains(kv.Key));
            string? setSpatialLine = null;
            var setOverlaps = new List<string>();
            if (positionChanged)
            {
                setSpatialLine = GetPptSpatialLine(handler, path);
                if (setSpatialLine != null) setOverlaps = CheckPositionOverlap(handler, path);
            }

            if (json)
            {
                var allWarnings = new List<OfficeCli.Core.CliWarning>();
                foreach (var ac in autoCorrected)
                {
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"Auto-corrected '{ac.Original}' to '{ac.Corrected}'",
                        Code = "auto_corrected",
                        Suggestion = ac.Corrected
                    });
                }
                foreach (var p in stillUnsupported)
                {
                    var suggestion = SuggestProperty(p);
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = suggestion != null ? $"Unsupported property: {p} (did you mean: {suggestion}?)" : $"Unsupported property: {p}",
                        Code = "unsupported_property",
                        Suggestion = suggestion
                    });
                }
                if (setOverlaps.Count > 0)
                {
                    allWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"Same position as {string.Join(", ", setOverlaps)}",
                        Code = "position_overlap",
                        Suggestion = "Use different x/y values to avoid overlap"
                    });
                }
                var outputMsg = setSpatialLine != null ? $"{message}\n  {setSpatialLine}" : message;
                bool allFailed = applied.Count == 0 && (stillUnsupported.Count > 0 || unsupported.Count > 0);
                Console.WriteLine(allFailed
                    ? OutputFormatter.WrapEnvelopeError(outputMsg, allWarnings.Count > 0 ? allWarnings : null)
                    : OutputFormatter.WrapEnvelopeText(outputMsg, allWarnings.Count > 0 ? allWarnings : null));
            }
            else
            {
                foreach (var ac in autoCorrected)
                    Console.Error.WriteLine($"WARNING: Auto-corrected '{ac.Original}' to '{ac.Corrected}'");
                Console.WriteLine(message);
                if (setSpatialLine != null) Console.WriteLine($"  {setSpatialLine}");
                if (setOverlaps.Count > 0)
                    Console.Error.WriteLine($"  WARNING: Same position as {string.Join(", ", setOverlaps)}");
                if (stillUnsupported.Count > 0)
                    Console.Error.WriteLine(FormatUnsupported(stillUnsupported));
            }
            NotifyWatch(handler, file.FullName, path);

            if (stillUnsupported.Count > 0) return 2;
            return 0;
        }, json); });

        rootCommand.Add(setCommand);

        // ==================== add command ====================
        var addFileArg = new Argument<FileInfo>("file") { Description = "Office document path (required even with open/close mode)" };
        var addParentPathArg = new Argument<string>("parent") { Description = "Parent DOM path (e.g. /body, /Sheet1, /slide[1])" };
        var addTypeOpt = new Option<string>("--type") { Description = "Element type to add (e.g. paragraph, run, table, sheet, row, cell, slide, shape)" };
        var addFromOpt = new Option<string?>("--from") { Description = "Copy from an existing element path (e.g. /slide[1]/shape[2])" };
        var addIndexOpt = new Option<int?>("--index") { Description = "Insert position (0-based). If omitted, appends to end" };
        var addPropsOpt = new Option<string[]>("--prop") { Description = "Property to set (key=value)", AllowMultipleArgumentsPerToken = true };

        var addCommand = new Command("add", "Add a new element to the document") { TreatUnmatchedTokensAsErrors = false };
        addCommand.Add(addFileArg);
        addCommand.Add(addParentPathArg);
        addCommand.Add(addTypeOpt);
        addCommand.Add(addFromOpt);
        addCommand.Add(addIndexOpt);
        addCommand.Add(addPropsOpt);
        addCommand.Add(jsonOption);
        addCommand.Add(forceOption);

        addCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(addFileArg)!;
            var parentPath = result.GetValue(addParentPathArg)!;
            var type = result.GetValue(addTypeOpt);
            var from = result.GetValue(addFromOpt);
            var index = result.GetValue(addIndexOpt);
            var props = result.GetValue(addPropsOpt);
            var force = result.GetValue(forceOption);
            bool hadWarnings = false;

            // Check document protection for .docx files
            if (!force && file.Extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
            {
                var protectionError = CheckDocxProtection(file.FullName, parentPath, json);
                if (protectionError != 0) return protectionError;
            }

            // Detect bare key=value positional arguments (missing --prop)
            var unmatchedKvWarnings = DetectUnmatchedKeyValues(result);
            if (unmatchedKvWarnings.Count > 0)
            {
                hadWarnings = true;
                if (json)
                {
                    var kvWarnings = unmatchedKvWarnings.Select(kv => new OfficeCli.Core.CliWarning
                    {
                        Message = $"Bare property '{kv}' ignored. Use --prop {kv}",
                        Code = "missing_prop_flag",
                        Suggestion = $"--prop {kv}"
                    }).ToList();
                    Console.Error.WriteLine("WARNING: Properties specified without --prop flag.");
                }
                else
                {
                    foreach (var kv in unmatchedKvWarnings)
                        Console.Error.WriteLine($"WARNING: Bare property '{kv}' ignored. Did you mean: --prop {kv}");
                    Console.Error.WriteLine("Hint: Properties must be passed with --prop flag, e.g. officecli add <file> <parent> --type <type> --prop key=value");
                }
            }

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
                }, json) is {} rc) return rc != 0 ? rc : (hadWarnings ? 2 : 0);

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
                }, json) is {} rc) return rc != 0 ? rc : (hadWarnings ? 2 : 0);

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
                var spatialLine = GetPptSpatialLine(handler, resultPath);
                var overlapNames = spatialLine != null ? CheckPositionOverlap(handler, resultPath) : new();
                var addWarnings = new List<OfficeCli.Core.CliWarning>();
                if (overlapNames.Count > 0)
                {
                    addWarnings.Add(new OfficeCli.Core.CliWarning
                    {
                        Message = $"Same position as {string.Join(", ", overlapNames)}",
                        Code = "position_overlap",
                        Suggestion = "Use --prop x=... y=... to set distinct positions"
                    });
                }
                if (json)
                {
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(
                        spatialLine != null ? $"{message}\n  {spatialLine}" : message,
                        addWarnings.Count > 0 ? addWarnings : null));
                }
                else
                {
                    Console.WriteLine(message);
                    if (spatialLine != null) Console.WriteLine($"  {spatialLine}");
                    foreach (var w in addWarnings)
                        Console.Error.WriteLine($"  WARNING: {w.Message}");
                }
                if (parentPath == "/") NotifyWatchRoot(handler, file.FullName, oldCount);
                else NotifyWatch(handler, file.FullName, parentPath);
            }

            return hadWarnings ? 2 : 0;
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
            }, json) is {} rc) return rc;

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
            return 0;
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
            }, json) is {} rc) return rc;

            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: true);
            var resultPath = handler.Move(path, to, index);
            var message = $"Moved to {resultPath}";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(message));
            else Console.WriteLine(message);
            NotifyWatch(handler, file.FullName, path);
            return 0;
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
            }, json) is {} rc) return rc;

            var rawCols = rawColsStr != null ? new HashSet<string>(rawColsStr.Split(',').Select(c => c.Trim().ToUpperInvariant())) : null;

            using var handler = DocumentHandlerFactory.Open(file.FullName);
            var xml = handler.Raw(partPath, startRow, endRow, rawCols);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(xml));
            else Console.WriteLine(xml);
            return 0;
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
            }, json) is {} rc) return rc;

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
            return 0;
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
            }, json) is {} rc) return rc;

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
            return 0;
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
            }, json) is {} rc) return rc;

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
            return 0;
        }, json); });
        rootCommand.Add(validateCommand);

        // ==================== batch command ====================
        var batchFileArg = new Argument<FileInfo>("file") { Description = "Office document path" };
        var batchInputOpt = new Option<FileInfo?>("--input") { Description = "JSON file containing batch commands. If omitted, reads from stdin" };
        var batchCommandsOpt = new Option<string?>("--commands") { Description = "Inline JSON array of batch commands (alternative to --input or stdin)" };
        var batchStopOnErrorOpt = new Option<bool>("--stop-on-error") { Description = "Stop execution on first error (default: continue all)" };
        var batchCommand = new Command("batch", "Execute multiple commands from a JSON array (one open/save cycle)");
        batchCommand.Add(batchFileArg);
        batchCommand.Add(batchInputOpt);
        batchCommand.Add(batchCommandsOpt);
        batchCommand.Add(batchStopOnErrorOpt);
        batchCommand.Add(jsonOption);

        batchCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(batchFileArg)!;
            var inputFile = result.GetValue(batchInputOpt);
            var inlineCommands = result.GetValue(batchCommandsOpt);
            var stopOnError = result.GetValue(batchStopOnErrorOpt);

            string jsonText;
            if (inlineCommands != null)
            {
                jsonText = inlineCommands;
            }
            else if (inputFile != null)
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
                    var success = response.ExitCode == 0;
                    results.Add(new BatchResult { Success = success, Output = response.Stdout, Error = response.Stderr });
                    if (!success && stopOnError) break;
                }
                PrintBatchResults(results, json);
                if (results.Any(r => !r.Success))
                    throw new InvalidOperationException($"Batch completed with {results.Count(r => !r.Success)} error(s)");
                return 0;
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
            return 0;
        }, json); });

        rootCommand.Add(batchCommand);

        // ==================== import command ====================
        var importFileArg = new Argument<FileInfo>("file") { Description = "Target Excel file (.xlsx)" };
        var importParentPathArg = new Argument<string>("parent-path") { Description = "Sheet path (e.g. /Sheet1)" };
        var importSourceOpt = new Option<FileInfo?>("--file") { Description = "Source CSV/TSV file to import" };
        var importStdinOpt = new Option<bool>("--stdin") { Description = "Read CSV/TSV data from stdin" };
        var importFormatOpt = new Option<string?>("--format") { Description = "Data format: csv or tsv (default: inferred from file extension, or csv)" };
        var importHeaderOpt = new Option<bool>("--header") { Description = "First row is header: set AutoFilter and freeze pane" };
        var importStartCellOpt = new Option<string>("--start-cell") { Description = "Starting cell (default: A1)" };
        importStartCellOpt.DefaultValueFactory = _ => "A1";

        var importCommand = new Command("import", "Import CSV/TSV data into an Excel sheet");
        importCommand.Add(importFileArg);
        importCommand.Add(importParentPathArg);
        importCommand.Add(importSourceOpt);
        importCommand.Add(importStdinOpt);
        importCommand.Add(importFormatOpt);
        importCommand.Add(importHeaderOpt);
        importCommand.Add(importStartCellOpt);
        importCommand.Add(jsonOption);

        importCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(importFileArg)!;
            var parentPath = result.GetValue(importParentPathArg)!;
            var source = result.GetValue(importSourceOpt);
            var useStdin = result.GetValue(importStdinOpt);
            var format = result.GetValue(importFormatOpt);
            var header = result.GetValue(importHeaderOpt);
            var startCell = result.GetValue(importStartCellOpt)!;

            if (!file.Exists)
                throw new CliException($"File not found: {file.FullName}")
                {
                    Code = "file_not_found",
                    Suggestion = $"Create the file first: officecli create \"{file.FullName}\""
                };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext != ".xlsx")
                throw new CliException("Import is only supported for .xlsx files in V1")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .xlsx file"
                };

            // Read CSV content
            string csvContent;
            if (useStdin)
            {
                csvContent = Console.In.ReadToEnd();
            }
            else if (source != null)
            {
                if (!source.Exists)
                    throw new CliException($"Source file not found: {source.FullName}")
                    {
                        Code = "file_not_found"
                    };
                csvContent = File.ReadAllText(source.FullName, Encoding.UTF8);
            }
            else
            {
                throw new CliException("Either --file or --stdin must be specified")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --file <path> to specify a CSV/TSV file, or --stdin to read from standard input"
                };
            }

            // Determine delimiter: --format flag > file extension > default csv
            char delimiter = ',';
            if (!string.IsNullOrEmpty(format))
            {
                delimiter = format.ToLowerInvariant() switch
                {
                    "tsv" => '\t',
                    "csv" => ',',
                    _ => throw new CliException($"Unknown format: {format}. Use 'csv' or 'tsv'")
                    {
                        Code = "invalid_value",
                        ValidValues = ["csv", "tsv"]
                    }
                };
            }
            else if (source != null)
            {
                var sourceExt = Path.GetExtension(source.FullName).ToLowerInvariant();
                if (sourceExt == ".tsv" || sourceExt == ".tab")
                    delimiter = '\t';
            }

            using var handler = new OfficeCli.Handlers.ExcelHandler(file.FullName, editable: true);
            var msg = handler.Import(parentPath, csvContent, delimiter, header, startCell);
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
            else
                Console.WriteLine(msg);
            return 0;
        }, json); });

        rootCommand.Add(importCommand);

        // ==================== create command ====================
        var createFileArg = new Argument<string>("file") { Description = "Output file path (.docx, .xlsx, .pptx)" };
        var createTypeOpt = new Option<string>("--type") { Description = "Document type (docx, xlsx, pptx) — optional, inferred from file extension" };
        var createCommand = new Command("create", "Create a blank Office document");
        createCommand.Aliases.Add("new");
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
            return 0;
        }, json); });

        rootCommand.Add(createCommand);

        // ==================== merge command (template merge) ====================
        var mergeTemplateArg = new Argument<string>("template") { Description = "Template file path (.docx, .xlsx, .pptx) with {{key}} placeholders" };
        var mergeOutputArg = new Argument<string>("output") { Description = "Output file path" };
        var mergeDataOpt = new Option<string>("--data") { Description = "JSON data or path to .json file", Required = true };
        var mergeCommand = new Command("merge", "Merge template with JSON data, replacing {{key}} placeholders");
        mergeCommand.Add(mergeTemplateArg);
        mergeCommand.Add(mergeOutputArg);
        mergeCommand.Add(mergeDataOpt);
        mergeCommand.Add(jsonOption);

        mergeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var template = result.GetValue(mergeTemplateArg)!;
            var output = result.GetValue(mergeOutputArg)!;
            var dataArg = result.GetValue(mergeDataOpt)!;

            var data = Core.TemplateMerger.ParseMergeData(dataArg);
            var mergeResult = Core.TemplateMerger.Merge(template, output, data);

            if (json)
            {
                var jsonObj = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["output"] = Path.GetFullPath(output),
                    ["replacedKeys"] = mergeResult.UsedKeys.Count,
                    ["unresolvedPlaceholders"] = new System.Text.Json.Nodes.JsonArray(
                        mergeResult.UnresolvedPlaceholders.Select(p => (System.Text.Json.Nodes.JsonNode)p).ToArray())
                };
                Console.WriteLine(jsonObj.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = false }));
            }
            else
            {
                Console.WriteLine($"Merged: {output}");
                Console.WriteLine($"  Replaced keys: {mergeResult.UsedKeys.Count}");
                if (mergeResult.UnresolvedPlaceholders.Count > 0)
                {
                    Console.Error.WriteLine($"  Warning: {mergeResult.UnresolvedPlaceholders.Count} unresolved placeholder(s):");
                    foreach (var p in mergeResult.UnresolvedPlaceholders)
                        Console.Error.WriteLine($"    - {{{{{p}}}}}");
                }
            }
            return 0;
        }, json); });

        rootCommand.Add(mergeCommand);

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
                if (string.IsNullOrEmpty(item.Command))
                    throw new InvalidOperationException(
                        "Batch item missing required 'command' field. " +
                        "Valid commands: get, query, set, add, remove, move, view, raw, validate. " +
                        "Example: {\"command\": \"set\", \"path\": \"/Sheet1/A1\", \"props\": {\"value\": \"hello\"}}");
                throw new InvalidOperationException($"Unknown command: '{item.Command}'. Valid commands: get, query, set, add, remove, move, view, raw, validate.");
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

    /// <summary>
    /// Detect bare key=value tokens in unmatched arguments (user forgot --prop).
    /// </summary>
    internal static List<string> DetectUnmatchedKeyValues(System.CommandLine.ParseResult parseResult)
    {
        var result = new List<string>();
        foreach (var token in parseResult.UnmatchedTokens)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(token, @"^[A-Za-z_.][A-Za-z0-9_.]*=.+$"))
                result.Add(token);
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
        "font.strike", "font.underline", "tabColor", "shadow", "glow",
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
                    var name = child.Format.TryGetValue("name", out var n) ? n?.ToString() : child.Path;
                    overlaps.Add(name ?? child.Path);
                }
            }
        }
        catch { /* ignore */ }
        return overlaps;
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
