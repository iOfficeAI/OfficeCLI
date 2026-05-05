// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    public static RootCommand BuildRootCommand()
    {
        var jsonOption = new Option<bool>("--json") { Description = "Output as JSON (AI-friendly)" };

        var rootCommand = new RootCommand("""
            officecli: AI-friendly CLI for Office documents (.docx, .xlsx, .pptx, .hwpx, experimental .hwp)

            Run 'officecli help' for the schema-driven capability reference (formats, elements, properties).
            Run 'officecli help hwp' for experimental rhwp bridge setup, examples, and support boundaries.
            See the Commands section below for the full list of subcommands.
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

            // If already running, reuse the existing resident. This covers
            // two cases with the same code path:
            //   (a) user previously called `open` explicitly, or
            //   (b) `create` just auto-started a short-lived (60s) resident.
            // In either case we upgrade the idle timeout to the default 12min
            // via the __set-idle-timeout__ ping RPC. Failure is non-fatal —
            // the resident is still usable, it'll just exit on its original
            // schedule. `open` is idempotent, so repeated calls are safe.
            const int DefaultOpenIdleSeconds = 12 * 60;
            if (ResidentClient.TryConnect(filePath, out _))
            {
                ResidentClient.SendSetIdleTimeout(filePath, DefaultOpenIdleSeconds);
                var msg = $"Opened {file.Name} (reusing running resident, idle timeout set to 12min)";
                if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else Console.WriteLine(msg);
                return 0;
            }

            if (!TryStartResidentProcess(filePath, idleSeconds: null, out var startError))
                throw new InvalidOperationException(startError);

            var startedMsg = $"Opened {file.Name} (remember to call close when done)";
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelopeText(startedMsg));
            else Console.WriteLine(startedMsg);
            return 0;
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
            if (ResidentClient.SendCloseWithResponse(file.FullName, out var closeResp))
            {
                // BUG-BT-R26-2: resident may report a non-zero shutdown
                // (e.g. file vanished mid-session → data loss). Bubble
                // that up instead of pretending the close succeeded.
                if (closeResp != null && closeResp.ExitCode != 0)
                {
                    var err = !string.IsNullOrEmpty(closeResp.Stderr)
                        ? closeResp.Stderr
                        : $"Resident close reported error (exit {closeResp.ExitCode})";
                    throw new InvalidOperationException(err);
                }
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
        rootCommand.Add(BuildMarkCommand(jsonOption));
        rootCommand.Add(BuildUnmarkMarkCommand(jsonOption));
        rootCommand.Add(BuildGetMarksCommand(jsonOption));
        rootCommand.Add(BuildGotoCommand(jsonOption));
        rootCommand.Add(BuildViewCommand(jsonOption));
        rootCommand.Add(BuildGetCommand(jsonOption));
        rootCommand.Add(BuildQueryCommand(jsonOption));
        rootCommand.Add(BuildSetCommand(jsonOption));
        rootCommand.Add(BuildHwpHelpCommand(jsonOption));
        rootCommand.Add(BuildAddCommand(jsonOption));
        rootCommand.Add(BuildRemoveCommand(jsonOption));
        rootCommand.Add(BuildMoveCommand(jsonOption));
        rootCommand.Add(BuildSwapCommand(jsonOption));
        rootCommand.Add(BuildRawCommand(jsonOption));
        rootCommand.Add(BuildRawSetCommand(jsonOption));
        rootCommand.Add(BuildAddPartCommand(jsonOption));
        rootCommand.Add(BuildValidateCommand(jsonOption));
        rootCommand.Add(BuildBatchCommand(jsonOption));
        rootCommand.Add(BuildDumpCommand(jsonOption));
        rootCommand.Add(BuildImportCommand(jsonOption));
        rootCommand.Add(BuildCreateCommand(jsonOption));
        rootCommand.Add(BuildMergeCommand(jsonOption));
        rootCommand.Add(BuildCompareCommand(jsonOption));
        rootCommand.Add(BuildCapabilitiesCommand(jsonOption));
        rootCommand.Add(BuildSchemaCommand(jsonOption));

        foreach (var stub in BuildIntegrationStubCommands())
            rootCommand.Add(stub);

        rootCommand.Add(BuildHelpCommand(jsonOption, rootCommand));

        return rootCommand;
    }

    // ==================== Helper: fork a __resident-serve__ subprocess ====================
    //
    // Used by both `open` (explicit) and `create` (auto-start after
    // creating a blank file). Forks the current executable with the
    // internal __resident-serve__ verb and waits up to 5s for the ping
    // pipe to respond, so callers get a definitive success/fail answer.
    //
    // `idleSeconds` overrides the child's idle-exit timeout via the
    // OFFICECLI_RESIDENT_IDLE_SECONDS env var (1..86400). Passing null
    // inherits the server default (12 minutes). `create` passes 60 so
    // an auto-started resident that nobody follows up on exits quickly.
    //
    // Caller must first verify no resident is already running for this
    // file (e.g. via ResidentClient.TryConnect) — this helper always
    // starts a fresh child.
    internal static bool TryStartResidentProcess(string filePath, int? idleSeconds, out string? error)
    {
        error = null;
        var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (exePath == null)
        {
            error = "Cannot determine executable path.";
            return false;
        }

        // On Windows, .NET's UseShellExecute=false always calls CreateProcess
        // with bInheritHandles=TRUE (even without explicit redirects), which
        // leaks the caller's pipe handles into the resident child.  When the
        // caller's stdout is a pipe ($(), | cat, CI, SDK), the pipe never
        // gets EOF until the resident exits (~60s idle), blocking the caller.
        //
        // Fix: temporarily mark our own std handles as non-inheritable before
        // spawning, then restore.  This prevents the shell's pipe handles
        // from leaking into the resident while still allowing .NET's internal
        // handle plumbing to work.
        //
        // On macOS/Linux, posix_spawn inherits fds unless the child's
        // stdout/stderr are explicitly redirected.  RedirectStandardOutput /
        // RedirectStandardError = true makes .NET plumb a fresh pipe from
        // parent to child, so the caller's shell pipe (e.g. `| tail -1`,
        // $(...)) is NOT inherited and EOFs promptly when the client exits.
        // See ResidentStdoutInheritanceTests for the regression lock-in.
        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"__resident-serve__ \"{filePath}\"",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        if (idleSeconds.HasValue)
            startInfo.Environment["OFFICECLI_RESIDENT_IDLE_SECONDS"] = idleSeconds.Value.ToString();

        // Prevent the shell's pipe handles from leaking into the resident.
        bool isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        nint hStdOut = 0, hStdErr = 0, hStdIn = 0;
        if (isWindows)
        {
            hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
            hStdErr = GetStdHandle(STD_ERROR_HANDLE);
            hStdIn  = GetStdHandle(STD_INPUT_HANDLE);
            SetHandleInformation(hStdOut, HANDLE_FLAG_INHERIT, 0);
            SetHandleInformation(hStdErr, HANDLE_FLAG_INHERIT, 0);
            SetHandleInformation(hStdIn,  HANDLE_FLAG_INHERIT, 0);
        }

        Process? process;
        try { process = Process.Start(startInfo); }
        finally
        {
            if (isWindows)
            {
                SetHandleInformation(hStdOut, HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT);
                SetHandleInformation(hStdErr, HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT);
                SetHandleInformation(hStdIn,  HANDLE_FLAG_INHERIT, HANDLE_FLAG_INHERIT);
            }
        }

        if (process == null)
        {
            error = "Failed to start resident process.";
            return false;
        }

        // Wait briefly for the server to start accepting connections.
        for (int i = 0; i < 50; i++) // up to 5 seconds
        {
            Thread.Sleep(100);
            if (ResidentClient.TryConnect(filePath, out _))
            {
                process.Dispose();
                return true;
            }
            if (process.HasExited)
            {
                var stderr = process.StandardError.ReadToEnd();
                error = $"Resident process exited. {stderr}";
                process.Dispose();
                return false;
            }
        }

        error = "Resident process started but not responding.";
        process.Dispose();
        return false;
    }

    // ==================== Win32 P/Invoke for handle inheritance control ==========

    private const int STD_INPUT_HANDLE  = -10;
    private const int STD_OUTPUT_HANDLE = -11;
    private const int STD_ERROR_HANDLE  = -12;
    private const uint HANDLE_FLAG_INHERIT = 0x00000001;

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern nint GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetHandleInformation(nint hObject, uint dwMask, uint dwFlags);

    // ==================== Helper: try forwarding to resident ====================
    //
    // Two-step protocol (CONSISTENCY(resident-two-step): same shape as
    // CommandBuilder.Batch.cs's resident branch):
    //   1. Ping-pipe probe via TryConnect — fast (100ms) and isolated from the
    //      main command queue, so it stays responsive even under flood. Tells
    //      us definitively whether a resident owns this file.
    //   2. If yes, send the command on the main pipe with a generous connect
    //      timeout + a few retries. If the send STILL fails, surface a
    //      distinct "busy" error (exit code 3) instead of falling back to
    //      DocumentHandlerFactory.Open — the old silent fallback could race
    //      the live resident and lose writes.
    //   3. If no resident, return null so the caller opens the file directly.
    //
    // Exit code 3 is reserved for "resident is alive but couldn't deliver the
    // command" so callers can distinguish it from a command-level failure.
    private const int ResidentBusyExitCode = 3;
    private const int ResidentBusyConnectTimeoutMs = 30000;
    private const int ResidentBusyMaxRetries = 3;

    internal static int? TryResident(string filePath, Action<ResidentRequest> configure, bool json = false)
    {
        // Step 1: does a resident own this file? Probe via the -ping pipe,
        // which is never serialized behind main-pipe commands.
        if (!ResidentClient.TryConnect(filePath, out _))
        {
            // No resident running — auto-start one to avoid file-lock conflicts
            // when multiple commands hit the same file in parallel.
            // Opt-out: OFFICECLI_NO_AUTO_RESIDENT=1 disables auto-start (e.g.
            // sandbox environments where named pipes may not work reliably).
            var noAuto = Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_RESIDENT");
            if (noAuto == "1" || string.Equals(noAuto, "true", StringComparison.OrdinalIgnoreCase))
                return null;

            if (!TryStartResidentProcess(filePath, idleSeconds: 60, out _))
            {
                // Startup failed — maybe another process just started a resident
                // for the same file (parallel race). Re-probe before giving up.
                if (!ResidentClient.TryConnect(filePath, out _))
                    return null; // truly no resident → caller falls back to direct file access
            }
            // Intentionally no user-facing hint here. UX testing with an AI
            // agent showed a standalone "background process" hint on a random
            // mid-batch command (e.g. `get`) creates low-grade anxiety without
            // giving the caller a concrete action — auto-close in 60s already
            // handles the cleanup, and other officecli commands work normally
            // through the resident regardless. The `create` command keeps a
            // small inline suffix on its success line because it's contextual
            // to a freshly-created file, not a nag fired from anywhere.
        }

        var request = new ResidentRequest();
        configure(request);
        if (json) request.Json = true;

        // Step 2: resident is confirmed alive — wait for our turn in the main
        // pipe queue. Do NOT silently fall back on failure; letting a second
        // writer touch the file while the resident holds it in memory loses
        // data on the resident's eventual save.
        var response = ResidentClient.TrySend(
            filePath, request,
            maxRetries: ResidentBusyMaxRetries,
            connectTimeoutMs: ResidentBusyConnectTimeoutMs);

        if (response == null)
        {
            var fileName = Path.GetFileName(filePath);
            var msg = $"Resident for {fileName} is running but the command could not be delivered (main pipe busy or unresponsive). Retry, or run 'officecli close {fileName}' and try again.";
            if (json)
                Console.WriteLine(OutputFormatter.WrapEnvelopeError(msg));
            else
                Console.Error.WriteLine($"Error: {msg}");
            return ResidentBusyExitCode;
        }

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
        // CONSISTENCY(error-wrap): bare XmlException leaks ("Data at the root
        // level is invalid. Line 1, position 1.") when an OOXML part is
        // externally corrupted. Surface a friendlier message naming the
        // underlying cause so users know it's a malformed part, not a bug.
        var rendered = ex is System.Xml.XmlException xe
            ? new InvalidDataException(
                $"Malformed XML in document part: {xe.Message} " +
                $"(the file appears to have a corrupted OOXML part).", xe)
            : ex;
        if (json)
        {
            // JSON mode: structured error envelope to stdout so AI agents get it in the same stream
            WarningContext.End(); // discard any partial warnings
            Console.WriteLine(OutputFormatter.WrapErrorEnvelope(rendered));
        }
        else
        {
            Console.Error.WriteLine($"Error: {rendered.Message}");
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
                {
                    var msg = $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}";
                    if (props.ContainsKey("find"))
                    {
                        var matched = handler switch
                        {
                            OfficeCli.Handlers.WordHandler wh => wh.LastFindMatchCount,
                            OfficeCli.Handlers.PowerPointHandler ph => ph.LastFindMatchCount,
                            OfficeCli.Handlers.ExcelHandler eh => eh.LastFindMatchCount,
                            _ => 0
                        };
                        msg += $" ({matched} matched)";
                    }
                    parts.Add(msg);
                }
                if (unsupported.Count > 0)
                {
                    // /styles/<id> in Word: route through curated hints
                    // instead of the generic "use raw-set" message. raw-set
                    // is an escape hatch and pushing users there for missing
                    // curated coverage trains them out of the canonical
                    // vocabulary. See StyleUnsupportedHints.
                    if (handler is OfficeCli.Handlers.WordHandler
                        && path.StartsWith("/styles/", StringComparison.Ordinal))
                    {
                        var styleHint = OfficeCli.Core.StyleUnsupportedHints.Format(unsupported);
                        if (styleHint != null) parts.Add(styleHint);
                    }
                    else
                    {
                        string? batchScope = handler switch
                        {
                            OfficeCli.Handlers.ExcelHandler => "excel",
                            OfficeCli.Handlers.WordHandler => "word",
                            OfficeCli.Handlers.PowerPointHandler => "pptx",
                            _ => null,
                        };
                        parts.Add(FormatUnsupported(unsupported, batchScope));
                    }
                }
                return string.Join("\n", parts);
            }
            case "add":
            {
                var parentPath = item.Parent ?? item.Path;
                if (string.IsNullOrEmpty(parentPath))
                    throw new ArgumentException("'add' command requires 'parent' field. Example: {\"command\": \"add\", \"parent\": \"/slide[1]\", \"type\": \"shape\", \"props\": {\"text\": \"Hello\"}}");
                if (string.IsNullOrEmpty(item.Type) && string.IsNullOrEmpty(item.From))
                    throw new ArgumentException("'add' command requires 'type' or 'from' field. Example: {\"command\": \"add\", \"parent\": \"/\", \"type\": \"slide\"}");
                InsertPosition? pos = null;
                if (item.Index.HasValue) pos = InsertPosition.AtIndex(item.Index.Value);
                else if (!string.IsNullOrEmpty(item.After)) pos = InsertPosition.AfterElement(item.After);
                else if (!string.IsNullOrEmpty(item.Before)) pos = InsertPosition.BeforeElement(item.Before);

                if (!string.IsNullOrEmpty(item.From))
                {
                    var resultPath = handler.CopyFrom(item.From, parentPath, pos);
                    return $"Copied to {resultPath}";
                }
                else
                {
                    var type = item.Type ?? "";
                    var resultPath = handler.Add(parentPath, type, pos, props);
                    var addMsg = $"Added {type} at {resultPath}";

                    // Surface silent-drop props that the curated Add helper
                    // could not consume. AddStyle / AddParagraph / AddRun
                    // populate LastAddUnsupportedProps. Use the curated
                    // hint formatter (no raw-set recommendation) so users
                    // learn the right curated alternative instead of being
                    // pushed to the escape hatch. Scope label = result path
                    // truncated to the meaningful prefix (/styles,
                    // /body/p[N], /body/p[N]/r[N]).
                    if (handler is OfficeCli.Handlers.WordHandler addWh
                        && addWh.LastAddUnsupportedProps.Count > 0)
                    {
                        var scope = ScopeLabelForWordPath(resultPath);
                        var hint = OfficeCli.Core.StyleUnsupportedHints.Format(addWh.LastAddUnsupportedProps, scope);
                        if (hint != null) addMsg += "\nWARNING: " + hint;
                    }
                    return addMsg;
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
                InsertPosition? movePos = null;
                if (item.Index.HasValue) movePos = InsertPosition.AtIndex(item.Index.Value);
                else if (!string.IsNullOrEmpty(item.After)) movePos = InsertPosition.AfterElement(item.After);
                else if (!string.IsNullOrEmpty(item.Before)) movePos = InsertPosition.BeforeElement(item.Before);
                var resultPath = handler.Move(path, item.To, movePos);
                return $"Moved to {resultPath}";
            }
            case "swap":
            {
                if (string.IsNullOrEmpty(item.Path) || string.IsNullOrEmpty(item.To))
                    throw new ArgumentException("'swap' command requires 'path' and 'to' fields. Example: {\"command\": \"swap\", \"path\": \"/slide[1]\", \"to\": \"/slide[2]\"}");
                var (p1, p2) = handler switch
                {
                    OfficeCli.Handlers.PowerPointHandler ppt => ppt.Swap(item.Path, item.To),
                    OfficeCli.Handlers.WordHandler word => word.Swap(item.Path, item.To),
                    OfficeCli.Handlers.ExcelHandler excel => excel.Swap(item.Path, item.To),
                    OfficeCli.Handlers.HwpxHandler hwpx => hwpx.Swap(item.Path, item.To),
                    _ => throw new InvalidOperationException("swap not supported for this document type")
                };
                return $"Swapped {p1} <-> {p2}";
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
                throw new InvalidOperationException($"Unknown command: '{item.Command}'. Valid commands: get, query, set, add, remove, move, swap, view, raw, validate.");
        }
    }

    private static Dictionary<string, string> ParsePropsArray(string[]? props)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in props ?? Array.Empty<string>())
        {
            var eqIdx = prop.IndexOf('=');
            // BUG-R40-B12: previously `eqIdx > 0` silently dropped both
            // `--prop =value` (empty key, eqIdx==0) and `--prop key`
            // (no equals, eqIdx==-1). Surface the empty-key form as a
            // hard error so AI callers don't waste a turn wondering why
            // their property had no effect.
            if (eqIdx == 0)
                throw new ArgumentException(
                    $"Invalid --prop '{prop}': key is empty. Use key=value (e.g. --prop name=Title).");
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

            // Pattern 3 (BUG-BT-R6): common typos for the `--prop` option name.
            // `--props '{"k":"v"}'` is silently swallowed by System.CommandLine
            // because `--props` (with trailing s) is not a known option, so the
            // JSON value goes into UnmatchedTokens too. Catch the typo so the
            // existing warning machinery emits a clear hint instead of letting
            // the agent ship a shape with no text.
            if (token is "--props" or "-props" or "--prop=" && i + 1 < tokens.Count)
            {
                var nextToken = tokens[i + 1];
                if (!nextToken.StartsWith("--"))
                {
                    result.Add($"--prop {nextToken}");
                    i++;
                    continue;
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Reduce a Word handler result path to the meaningful scope label for
    /// UNSUPPORTED messages — "/styles", "/body/p[N]", "/body/p[N]/r[N]".
    /// Stops at the first segment that is not a known top-level Word
    /// container so unfamiliar paths fall back to the full path.
    /// </summary>
    private static string ScopeLabelForWordPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return "/";
        if (path.StartsWith("/styles/", StringComparison.Ordinal)) return "/styles";
        // Trim everything past the last bracketed-segment we recognize for
        // paragraph/run paths. Keep the path as-is for everything else.
        return path;
    }

    internal static string FormatUnsupported(IEnumerable<string> unsupported, string? scope = null)
    {
        var parts = new List<string>();
        foreach (var prop in unsupported)
        {
            var suggestion = SuggestPropertyScoped(prop, scope);
            parts.Add(suggestion != null ? $"{prop} (did you mean: {suggestion}?)" : prop);
        }
        return $"UNSUPPORTED props: {string.Join(", ", parts)}. Use 'officecli help <format>-set' to see available properties, or use raw-set for direct XML manipulation.";
    }

    /// <summary>
    /// Property keys that belong to PPTX shape/text semantics and should not
    /// be offered as suggestions when the caller is operating on an Excel
    /// document (R2-4). Keep the list conservative — only keys whose presence
    /// in an Excel error message would be clearly misleading.
    /// </summary>
    internal static readonly HashSet<string> PptxOnlyProps = new(StringComparer.OrdinalIgnoreCase)
    {
        "rotation", "opacity", "glow", "shadow",
        "firstSliceAngle", "holeSize", "bubbleScale", "explosion",
        "view3d", "varyColors",
    };

    /// <summary>
    /// Property keys exclusive to Word document-level concerns that should
    /// not bleed into Excel suggestions.
    /// </summary>
    internal static readonly HashSet<string> WordOnlyProps = new(StringComparer.OrdinalIgnoreCase)
    {
        "pageWidth", "pageHeight", "orientation",
    };

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
    /// Scoped variant: filters the suggestion pool against a target document
    /// format ("excel", "word", "pptx", or null for unscoped) to avoid
    /// cross-format leakage such as suggesting PPTX 'rotation' for an
    /// Excel pivot property (R2-4).
    /// </summary>
    internal static string? SuggestPropertyScoped(string input, string? scope)
    {
        var (best, _, _) = SuggestPropertyWithDistance(input, scope);
        return best;
    }

    /// <summary>
    /// Returns (bestMatch, distance, isUnique) where isUnique means no other candidate shares the same distance.
    /// </summary>
    internal static (string? Best, int Distance, bool IsUnique) SuggestPropertyWithDistance(string input, string? scope = null)
    {
        // Strip help text suffix if present (e.g. "key (valid props: ...)")
        var rawInput = input.Contains(' ') ? input[..input.IndexOf(' ')] : input;
        var lower = rawInput.ToLowerInvariant();
        string? best = null;
        int bestDist = int.MaxValue;
        int bestCount = 0; // how many props share the best distance

        HashSet<string>? exclude = null;
        switch (scope?.ToLowerInvariant())
        {
            case "excel":
                exclude = new HashSet<string>(PptxOnlyProps, StringComparer.OrdinalIgnoreCase);
                foreach (var w in WordOnlyProps) exclude.Add(w);
                break;
            case "word":
                exclude = PptxOnlyProps;
                break;
            case "pptx":
                exclude = WordOnlyProps;
                break;
        }

        foreach (var prop in KnownProps)
        {
            if (exclude != null && exclude.Contains(prop)) continue;
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
    internal static string? CheckTextOverflow(IDocumentHandler handler, string path)
    {
        try
        {
            return handler switch
            {
                OfficeCli.Handlers.PowerPointHandler ppt => ppt.CheckShapeTextOverflow(path),
                OfficeCli.Handlers.ExcelHandler xl => xl.CheckCellOverflow(path),
                _ => null
            };
        }
        catch { return null; }
    }

    /// <summary>
    /// Notify watch server with pre-rendered HTML from the handler.
    /// Call this while the handler is still open (before Dispose).
    /// </summary>
    private static void NotifyWatch(IDocumentHandler handler, string filePath, string? changedPath)
    {
        if (!WatchServer.IsWatching(filePath)) return;

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
        if (handler is OfficeCli.Handlers.HwpxHandler hwpx)
        {
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = hwpx.ViewAsHtml() });
            return;
        }
        if (handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var slideNum = WatchMessage.ExtractSlideNum(changedPath);
        if (slideNum > 0)
        {
            var html = ppt.RenderSlideHtml(slideNum);
            if (html != null)
            {
                // Slide-scoped replace: the watch server patches its cached _currentHtml in
                // place via PatchSlideInHtml; bundling a full ViewAsHtml() here is redundant
                // (and ResidentServer.NotifyWatchSlideChanged already omits it).
                WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "replace", Slide = slideNum, Html = html });
                return;
            }
        }
        WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = ppt.ViewAsHtml() });
    }

    private static void NotifyWatchRoot(IDocumentHandler handler, string filePath, int oldSlideCount)
    {
        if (!WatchServer.IsWatching(filePath)) return;

        if (handler is OfficeCli.Handlers.HwpxHandler hwpx)
        {
            WatchNotifier.NotifyIfWatching(filePath, new WatchMessage { Action = "full", FullHtml = hwpx.ViewAsHtml() });
            return;
        }
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
