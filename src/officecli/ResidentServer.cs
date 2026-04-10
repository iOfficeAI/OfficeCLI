// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

public class ResidentServer : IDisposable
{
    private readonly IDocumentHandler _handler;
    private readonly string _filePath;
    private readonly string _pipeName;
    // Shutdown uses TWO independent CTSs so the ping pipe can outlive the
    // handler dispose. This establishes the critical invariant that
    // TryResident relies on:
    //
    //   ping responds  ⇔  handler holds the file
    //
    // _mainCts gates the main command loop (accept + HandleClient). It is
    // cancelled FIRST during shutdown so no new commands start while we
    // are draining the in-flight one.
    //
    // _pingCts gates the ping responder and idle watchdog. It is cancelled
    // AFTER _handler.Dispose() completes, so any client that saw a live
    // ping is guaranteed to race against a still-locked file — and
    // therefore any subsequent fallback to direct file access will either
    // find the file released (ping gone) or get a retryable "busy" error.
    private CancellationTokenSource _mainCts = new();
    private CancellationTokenSource _pingCts = new();
    private readonly SemaphoreSlim _commandLock = new(1, 1);
    private readonly TimeSpan _idleTimeout = ResolveIdleTimeout();
    private CancellationTokenSource _idleCts = new();
    private bool _disposed;

    // Idle timeout is configurable via OFFICECLI_RESIDENT_IDLE_SECONDS so
    // tests can exercise the auto-shutdown path in seconds instead of
    // minutes, and so future "open file → auto-start resident" UX can
    // tune how aggressively the background process exits.
    // Valid range: 1s .. 24h. Anything outside falls back to the 12min
    // default. A value of "0" is rejected (would be an infinite-busy
    // spin on the watchdog task).
    private static TimeSpan ResolveIdleTimeout()
    {
        var raw = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_IDLE_SECONDS");
        if (!string.IsNullOrWhiteSpace(raw)
            && int.TryParse(raw, out var secs)
            && secs >= 1 && secs <= 86400)
        {
            return TimeSpan.FromSeconds(secs);
        }
        return TimeSpan.FromMinutes(12);
    }
    // Shared shutdown Task so __close__ and Dispose coordinate on a single
    // ordered teardown: drain in-flight command → dispose handler → ack client.
    private readonly object _shutdownLock = new();
    private Task? _shutdownTask;

    public string PipeName => _pipeName;

    public ResidentServer(string filePath, bool editable = true)
    {
        _filePath = Path.GetFullPath(filePath);
        _pipeName = GetPipeName(_filePath);
        _handler = DocumentHandlerFactory.Open(_filePath, editable);
    }

    public static string GetPipeName(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
            fullPath = fullPath.ToUpperInvariant();
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)))[..16];
        return $"officecli-{hash}";
    }

    public async Task RunAsync(CancellationToken externalToken = default)
    {
        // Main command loop is gated by _mainCts; ping responder and idle
        // watchdog are gated by _pingCts. The external token cancels both
        // via two linked CTSs so a caller's cancellation still shuts the
        // whole server down.
        using var mainLinked = CancellationTokenSource.CreateLinkedTokenSource(_mainCts.Token, externalToken);
        using var pingLinked = CancellationTokenSource.CreateLinkedTokenSource(_pingCts.Token, externalToken);
        var mainToken = mainLinked.Token;
        var pingToken = pingLinked.Token;

        // Hook graceful shutdown signals. Without this, a terminal HUP,
        // a cooperative `kill`, or a launcher's SIGTERM would terminate
        // the process before handler.Dispose() could flush the in-memory
        // tree to disk — the file lock would release but the user's
        // unsaved edits would be lost.
        //
        // PosixSignalRegistration runs our handler BEFORE the .NET
        // runtime begins its shutdown sequence, while the ThreadPool is
        // still fully healthy, so Task.Run continuations inside
        // DoShutdownAsync can complete reliably. On Unix it hooks
        // SIGTERM/SIGINT/SIGQUIT; on Windows it hooks the equivalent
        // console control events. Calling Cancel() on the context
        // suppresses the default abort so our shutdown can run to
        // completion.
        var signalRegs = new List<PosixSignalRegistration>();
        void HandleSignal(PosixSignalContext ctx)
        {
            ctx.Cancel = true;
            try { ShutdownAsync().Wait(TimeSpan.FromMinutes(10)); } catch { }
            Environment.Exit(0);
        }
        // SIGTERM and SIGINT work on every supported platform (Windows
        // maps SIGINT to Ctrl+C and SIGTERM to its equivalent console
        // control). SIGQUIT and SIGHUP are POSIX-only and throw
        // PlatformNotSupportedException on Windows — register each
        // individually so a Windows host still gets SIGTERM/SIGINT
        // coverage.
        void TryRegister(PosixSignal sig)
        {
            try { signalRegs.Add(PosixSignalRegistration.Create(sig, HandleSignal)); }
            catch (PlatformNotSupportedException) { /* skip on unsupported host */ }
        }
        TryRegister(PosixSignal.SIGTERM);
        TryRegister(PosixSignal.SIGINT);
        TryRegister(PosixSignal.SIGQUIT);
        TryRegister(PosixSignal.SIGHUP);

        // Also hook ProcessExit as a last-resort safety net for any exit
        // path that PosixSignalRegistration didn't cover (e.g.
        // Environment.Exit from other code).
        void OnProcessExit(object? s, EventArgs e)
        {
            try { ShutdownAsync().Wait(TimeSpan.FromMinutes(10)); } catch { }
        }
        AppDomain.CurrentDomain.ProcessExit += OnProcessExit;

        // Start ping responder on a dedicated pipe (never blocked by business commands)
        var pingTask = RunPingResponderAsync(pingToken);

        // Start idle watchdog
        var idleTask = RunIdleWatchdogAsync(pingToken);

        // Main command loop - accept connections concurrently, serialize
        // command execution. CONSISTENCY(pipe-precreate): same pre-create
        // pattern as RunPingResponderAsync (see BUG-FUZZER-R6-B-01). Creating
        // the next NamedPipeServerStream BEFORE handing off the accepted one
        // closes the window where no instance is listening — without this,
        // client bursts (e.g. 50 concurrent `officecli get`) race into the
        // gap and get ECONNREFUSED on macOS, which used to be silently hidden
        // by TryResident's fall-back path but now (correctly) surfaces as
        // "resident busy". Both instances coexist via MaxAllowedServerInstances
        // while the handler runs.
        NamedPipeServerStream NewMainServer() => new(_pipeName, PipeDirection.InOut,
            NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte, PipeOptions.Asynchronous);

        var currentMain = NewMainServer();
        try
        {
            while (!mainToken.IsCancellationRequested)
            {
                try
                {
                    await currentMain.WaitForConnectionAsync(mainToken);
                    // Hand over the accepted instance and immediately stand
                    // up a replacement so the pipe is never unlistened while
                    // the handler runs.
                    var accepted = currentMain;
                    currentMain = NewMainServer();
                    _ = HandleClientWithLockAsync(accepted, mainToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Resident error: {ex.Message}");
                    // currentMain is still the pre-created replacement; it is
                    // still valid for the next iteration's WaitForConnectionAsync.
                }
            }
        }
        finally
        {
            try { await currentMain.DisposeAsync(); } catch { }
        }

        // Main loop exited (via _mainCts cancel). The ping responder and
        // idle watchdog are still live under _pingCts; they will be
        // cancelled by DoShutdownAsync AFTER handler.Dispose() has
        // released the file lock. This keeps the ping-liveness invariant
        // intact even while the slow handler.Dispose() is running.
        try { await ShutdownAsync(); } catch { }

        try { await pingTask; } catch (OperationCanceledException) { }
        try { await idleTask; } catch (OperationCanceledException) { }

        AppDomain.CurrentDomain.ProcessExit -= OnProcessExit;
        foreach (var reg in signalRegs)
            try { reg.Dispose(); } catch { }
    }

    private void ResetIdleTimer()
    {
        // Cancel the old idle CTS to restart the delay; do not Dispose because
        // RunIdleWatchdogAsync may race between Volatile.Read and .Token access.
        var oldCts = Interlocked.Exchange(ref _idleCts, new CancellationTokenSource());
        oldCts.Cancel();
    }

    private async Task RunIdleWatchdogAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                // Snapshot the current idle CTS; ResetIdleTimer() swaps it to restart the wait
                var idleCts = Volatile.Read(ref _idleCts);
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(idleCts.Token, token);
                await Task.Delay(_idleTimeout, linked.Token);

                // Reached here = idle timeout elapsed without reset.
                // Kick off the ordered shutdown path instead of raw-
                // cancelling _mainCts / _pingCts, so the "ping liveness ⇔
                // file locked" invariant is preserved end-to-end: the
                // ping pipe stays alive until handler.Dispose() completes.
                Console.Error.WriteLine($"Resident idle for {_idleTimeout.TotalMinutes} minutes, closing.");
                _ = ShutdownAsync();
                break;
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                // _idleCts was cancelled (timer reset), loop and wait again
            }
        }
    }

    private async Task RunPingResponderAsync(CancellationToken token)
    {
        var pingPipeName = _pipeName + "-ping";

        // CONSISTENCY(pipe-precreate): pre-create the next server instance
        // BEFORE handing off the accepted one, so there is no window where
        // TryConnect can return false even though the resident is alive
        // (BUG-FUZZER-R6-B-01). Both instances live concurrently via
        // MaxAllowedServerInstances; the OS routes the next client to
        // whichever server is in WaitForConnectionAsync first.
        //
        // CONCURRENCY: the per-connection request handler runs
        // fire-and-forget so multiple ping probes can be serviced in
        // parallel. Without this, a burst of N concurrent
        // `ResidentClient.TryConnect` calls (e.g. from a fan-out of `set`
        // commands right after `open` returns) would serialize behind the
        // single accepted connection — and clients whose Connect(100ms)
        // expired during the wait would incorrectly conclude "no resident"
        // and fall back to direct file access, racing against the locked
        // file.
        NamedPipeServerStream NewServer() => new(pingPipeName, PipeDirection.InOut,
            NamedPipeServerStream.MaxAllowedServerInstances,
            PipeTransmissionMode.Byte, PipeOptions.Asynchronous);

        var current = NewServer();
        try
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    await current.WaitForConnectionAsync(token);

                    var accepted = current;
                    current = NewServer();

                    // Fire-and-forget the per-request handler so the loop
                    // can immediately go back to WaitForConnectionAsync on
                    // the replacement server. Exceptions are swallowed
                    // inside HandlePingRequestAsync.
                    _ = HandlePingRequestAsync(accepted, token);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Ping responder error: {ex.Message}");
                    // currentMain/current is already the replacement;
                    // loop continues.
                }
            }
        }
        finally
        {
            try { await current.DisposeAsync(); } catch { }
        }
    }

    private async Task HandlePingRequestAsync(NamedPipeServerStream accepted, CancellationToken token)
    {
        try
        {
            // Use raw byte I/O to dodge the StreamReader cancellation-
            // path deadlock on Windows named pipes under .NET 11 preview.
            var requestLine = await ReadLineFromPipeAsync(accepted, token);
            if (requestLine == null) return;

            var request = System.Text.Json.JsonSerializer.Deserialize<ResidentRequest>(
                requestLine, ResidentJsonContext.Default.ResidentRequest);
            if (request == null) return;

            if (request.Command == "__ping__")
            {
                var response = MakeResponse(0, _filePath, "");
                await WriteLineToPipeAsync(accepted, response, token);
                return;
            }

            if (request.Command == "__close__")
            {
                // Fully shut down the handler BEFORE acking, so the
                // client's subsequent file access races a guaranteed-
                // released file (see close-race commit for details).
                try { await ShutdownAsync(); }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Shutdown error during __close__: {ex.Message}");
                }

                var response = MakeResponse(0, "Closing resident.", "");
                // ShutdownAsync cancelled the ping token; write on a
                // fresh CTS so the client still gets the ack.
                using var writeCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                try { await WriteLineToPipeAsync(accepted, response, writeCts.Token); }
                catch { /* client may have disconnected; nothing to do */ }
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Ping handler error: {ex.Message}");
        }
        finally
        {
            try { await accepted.DisposeAsync(); } catch { }
        }
    }

    private async Task HandleClientWithLockAsync(NamedPipeServerStream server, CancellationToken token)
    {
        try
        {
            await _commandLock.WaitAsync(token);
            try
            {
                ResetIdleTimer();
                await HandleClientAsync(server, token);
            }
            finally
            {
                _commandLock.Release();
                ResetIdleTimer();
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Resident error: {ex.Message}");
        }
        finally
        {
            await server.DisposeAsync();
        }
    }

    private async Task HandleClientAsync(NamedPipeServerStream server, CancellationToken token)
    {
        var requestLine = await ReadLineFromPipeAsync(server, token);
        if (requestLine == null) return;

        var response = ProcessRequest(requestLine);
        await WriteLineToPipeAsync(server, response, token);
    }

    private string ProcessRequest(string requestLine)
    {
        ResidentRequest? request = null;
        try
        {
            request = System.Text.Json.JsonSerializer.Deserialize<ResidentRequest>(requestLine, ResidentJsonContext.Default.ResidentRequest);
            if (request == null)
                return MakeResponse(1, "", "Invalid request");

            // Capture stdout/stderr (safe: _commandLock serializes all commands)
            var stdoutWriter = new StringWriter();
            var stderrWriter = new StringWriter();
            var origOut = Console.Out;
            var origErr = Console.Error;
            Console.SetOut(stdoutWriter);
            Console.SetError(stderrWriter);

            try
            {
                ExecuteCommand(request);
            }
            finally
            {
                Console.SetOut(origOut);
                Console.SetError(origErr);
            }

            var stdout = stdoutWriter.ToString().TrimEnd('\r', '\n');
            var stderr = stderrWriter.ToString().TrimEnd('\r', '\n');

            if (request.Json)
            {
                // JSON mode: server builds the envelope so client just passes through
                var warnings = BuildWarnings(stderr);
                var isFailure = string.IsNullOrEmpty(stdout) && warnings is { Count: > 0 }
                    || stdout.StartsWith("No properties applied", StringComparison.Ordinal);
                var envelope = IsJson(stdout)
                    ? OutputFormatter.WrapEnvelope(stdout, warnings)
                    : isFailure
                        ? OutputFormatter.WrapEnvelopeError(stdout, warnings)
                        : OutputFormatter.WrapEnvelopeText(stdout, warnings);
                return MakeResponse(0, envelope, "");
            }

            return MakeResponse(0, stdout, stderr);
        }
        catch (Exception ex)
        {
            if (request?.Json == true)
            {
                // JSON mode: wrap error in envelope
                return MakeResponse(1, OutputFormatter.WrapErrorEnvelope(ex), "");
            }
            return MakeResponse(1, "", ex.Message);
        }
    }

    private static bool IsJson(string s)
    {
        var trimmed = s.AsSpan().TrimStart();
        return trimmed.Length > 0 && (trimmed[0] == '{' || trimmed[0] == '[');
    }

    private static List<CliWarning>? BuildWarnings(string stderr)
    {
        if (string.IsNullOrEmpty(stderr)) return null;
        var lines = stderr.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0) return null;
        return lines.Select(line =>
        {
            var warning = new CliWarning { Message = line.Trim() };
            if (line.Contains("UNSUPPORTED")) warning.Code = "unsupported_property";
            else if (line.Contains("VALIDATION")) warning.Code = "validation_error";
            else warning.Code = "warning";
            return warning;
        }).ToList();
    }

    private void ExecuteCommand(ResidentRequest request)
    {
        var format = request.Json ? OutputFormat.Json : OutputFormat.Text;

        switch (request.Command)
        {
            case "view":
                ExecuteView(request, format);
                break;
            case "get":
                ExecuteGet(request, format);
                break;
            case "query":
                ExecuteQuery(request, format);
                break;
            case "set":
                ExecuteSet(request);
                NotifyWatchSlideChanged(request.GetArg("path"));
                break;
            case "add":
            {
                var oldCount = GetPptSlideCount();
                ExecuteAdd(request);
                var parent = request.GetArg("parent");
                if (parent == "/")
                    NotifyWatchRootChanged(oldCount);
                else
                    NotifyWatchSlideChanged(parent);
                break;
            }
            case "remove":
            {
                var oldCount = GetPptSlideCount();
                var path = request.GetArg("path");
                ExecuteRemove(request);
                if (WatchMessage.ExtractSlideNum(path) > 0 && path != null && !path.Contains("/shape["))
                    NotifyWatchRootChanged(oldCount);
                else
                    NotifyWatchSlideChanged(path);
                break;
            }
            case "move":
                ExecuteMove(request);
                NotifyWatchSlideChanged(request.GetArg("path"));
                break;
            case "raw":
                ExecuteRaw(request);
                break;
            case "raw-set":
                ExecuteRawSet(request);
                NotifyWatchFullRefresh();
                break;
            case "add-part":
                ExecuteAddPart(request);
                NotifyWatchFullRefresh();
                break;
            case "validate":
                ExecuteValidate();
                break;
            case "batch":
                ExecuteBatch(request);
                break;
            default:
                // BUG-FUZZER-R6-A-06/07: previously this branch only wrote to
                // stderr and fell through, leaving the response with
                // ExitCode=0. Callers (and especially the AI agent piping the
                // CLI) had no way to detect that a typo / case-mangled verb
                // was actually rejected. Throw so ProcessRequest's exception
                // handler maps this to a proper non-zero ExitCode response.
                throw new InvalidOperationException($"Unknown command: {request.Command}");
        }
    }

    private void ExecuteBatch(ResidentRequest request)
    {
        var batchJson = request.GetArg("batchJson");
        var force = request.GetArg("force", "false")
            .Equals("true", StringComparison.OrdinalIgnoreCase);
        var stopOnError = !force;
        var json = request.Json;

        var items = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(
            batchJson, BatchJsonContext.Default.ListBatchItem) ?? new();

        var results = new List<BatchResult>();
        for (int bi = 0; bi < items.Count; bi++)
        {
            var item = items[bi];
            // Skip open/close commands inside batch — the resident already
            // holds the file open; issuing open/close would conflict.
            var cmd = (item.Command ?? "").ToLowerInvariant();
            if (cmd is "open" or "close")
            {
                results.Add(new BatchResult { Index = bi, Success = true, Output = $"Skipped '{cmd}' (resident mode)" });
                continue;
            }
            try
            {
                var output = CommandBuilder.ExecuteBatchItem(_handler, item, json);
                results.Add(new BatchResult { Index = bi, Success = true, Output = output });
            }
            catch (Exception ex)
            {
                results.Add(new BatchResult
                {
                    Index = bi, Success = false, Item = item,
                    Error = ex.Message
                });
                if (stopOnError) break;
            }
        }

        CommandBuilder.PrintBatchResults(results, json, items.Count);
    }

    // ==================== Watch notification helpers ====================

    private int GetPptSlideCount()
    {
        if (_handler is OfficeCli.Handlers.PowerPointHandler ppt)
            return ppt.GetSlideCount();
        return 0;
    }

    private void NotifyWatchSlideChanged(string? changedPath)
    {
        if (_handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            string? scrollTo = null;
            var sheetName = WatchMessage.ExtractSheetName(changedPath);
            if (sheetName != null)
            {
                var idx = excel.GetSheetIndex(sheetName);
                if (idx >= 0) scrollTo = $".sheet-content[data-sheet=\"{idx}\"]";
            }
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = excel.ViewAsHtml(), ScrollTo = scrollTo });
            return;
        }
        if (_handler is OfficeCli.Handlers.WordHandler word)
        {
            var scrollTo = WatchMessage.ExtractWordScrollTarget(changedPath);
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = word.ViewAsHtml(), ScrollTo = scrollTo });
            return;
        }
        if (_handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var slideNum = WatchMessage.ExtractSlideNum(changedPath);
        if (slideNum > 0)
        {
            var html = ppt.RenderSlideHtml(slideNum);
            if (html != null)
            {
                WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "replace", Slide = slideNum, Html = html });
                return;
            }
        }
        WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full" });
    }

    private void NotifyWatchRootChanged(int oldSlideCount)
    {
        if (_handler is OfficeCli.Handlers.WordHandler word)
        {
            var html = word.ViewAsHtml();
            var pageCount = System.Text.RegularExpressions.Regex.Matches(html, @"data-page=""\d+""").Count;
            var scrollTo = pageCount > 0 ? $".page[data-page=\"{pageCount}\"]" : null;
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = html, ScrollTo = scrollTo });
            return;
        }
        if (_handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = excel.ViewAsHtml() });
            return;
        }
        if (_handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var newCount = ppt.GetSlideCount();
        if (newCount > oldSlideCount)
        {
            var html = ppt.RenderSlideHtml(newCount);
            if (html != null)
            {
                WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "add", Slide = newCount, Html = html, FullHtml = ppt.ViewAsHtml() });
                return;
            }
        }
        else if (newCount < oldSlideCount)
        {
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "remove", Slide = oldSlideCount, FullHtml = ppt.ViewAsHtml() });
            return;
        }
        WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = ppt.ViewAsHtml() });
    }

    private void NotifyWatchFullRefresh()
    {
        string? fullHtml = null;
        if (_handler is OfficeCli.Handlers.PowerPointHandler ppt)
            fullHtml = ppt.ViewAsHtml();
        else if (_handler is OfficeCli.Handlers.ExcelHandler excel)
            fullHtml = excel.ViewAsHtml();
        else if (_handler is OfficeCli.Handlers.WordHandler word)
            fullHtml = word.ViewAsHtml();
        if (fullHtml != null)
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = fullHtml });
    }

    private void ExecuteView(ResidentRequest req, OutputFormat format)
    {
        var mode = req.GetArg("mode", "text")!;
        var start = req.GetIntArg("start");
        var end = req.GetIntArg("end");
        var maxLines = req.GetIntArg("max-lines");
        var issueType = req.GetArgOrNull("type");
        var limit = req.GetIntArg("limit");
        var cols = req.GetCols("cols");
        var pageFilter = req.GetArgOrNull("page");

        if (mode!.ToLowerInvariant() is "html" or "h")
        {
            string? html = null;
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
                html = pptHandler.ViewAsHtml(start, end);
            else if (_handler is OfficeCli.Handlers.ExcelHandler excelHandler)
                html = excelHandler.ViewAsHtml();
            else if (_handler is OfficeCli.Handlers.WordHandler wordHandler)
                html = wordHandler.ViewAsHtml(pageFilter);

            if (html != null)
            {
                if (req.Json)
                {
                    Console.Write(html);
                }
                else
                {
                    var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}.html");
                    File.WriteAllText(htmlPath, html);
                    Console.WriteLine(htmlPath);
                    try
                    {
                        var psi = new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi);
                    }
                    catch { /* silently ignore if browser can't be opened */ }
                }
            }
            else
            {
                Console.Error.WriteLine("HTML preview is only supported for .pptx, .xlsx, and .docx files.");
            }
            return;
        }

        if (mode!.ToLowerInvariant() is "svg" or "g")
        {
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptSvgHandler)
            {
                var slideNum = start ?? 1;
                var svg = pptSvgHandler.ViewAsSvg(slideNum);
                Console.Write(svg);
            }
            else
            {
                Console.Error.WriteLine("SVG preview is only supported for .pptx files.");
            }
            return;
        }

        if (req.Json)
        {
            var modeKey = mode!.ToLowerInvariant();
            if (modeKey is "stats" or "s")
                Console.WriteLine(_handler.ViewAsStatsJson().ToJsonString(OutputFormatter.PublicJsonOptions));
            else if (modeKey is "outline" or "o")
                Console.WriteLine(_handler.ViewAsOutlineJson().ToJsonString(OutputFormatter.PublicJsonOptions));
            else if (modeKey is "text" or "t")
                Console.WriteLine(_handler.ViewAsTextJson(start, end, maxLines, cols).ToJsonString(OutputFormatter.PublicJsonOptions));
            else if (modeKey is "annotated" or "a")
                Console.WriteLine(OutputFormatter.FormatView(mode, _handler.ViewAsAnnotated(start, end, maxLines, cols), format));
            else if (modeKey is "issues" or "i")
                Console.WriteLine(OutputFormatter.FormatIssues(_handler.ViewAsIssues(issueType, limit), format));
            else if (modeKey is "forms" or "f")
            {
                if (_handler is OfficeCli.Handlers.WordHandler wordFormsHandler)
                    Console.WriteLine(wordFormsHandler.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions));
                else
                    Console.Error.WriteLine("Forms view is only supported for .docx files.");
            }
            else
                Console.WriteLine($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, forms");
        }
        else
        {
            var output = mode!.ToLowerInvariant() switch
            {
                "text" or "t" => _handler.ViewAsText(start, end, maxLines, cols),
                "annotated" or "a" => _handler.ViewAsAnnotated(start, end, maxLines, cols),
                "outline" or "o" => _handler.ViewAsOutline(),
                "stats" or "s" => _handler.ViewAsStats(),
                "issues" or "i" => OutputFormatter.FormatIssues(_handler.ViewAsIssues(issueType, limit), format),
                "forms" or "f" => _handler is OfficeCli.Handlers.WordHandler wfh
                    ? wfh.ViewAsForms()
                    : "Forms view is only supported for .docx files.",
                _ => $"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, forms"
            };
            Console.WriteLine(output);
        }
    }

    private void ExecuteGet(ResidentRequest req, OutputFormat format)
    {
        var path = req.GetArg("path", "/");
        var depth = req.GetIntArg("depth") ?? 1;
        var node = _handler.Get(path, depth);
        Console.WriteLine(OutputFormatter.FormatNode(node, format));
    }

    private void ExecuteQuery(ResidentRequest req, OutputFormat format)
    {
        var selector = req.GetArg("selector", "");
        var filters = AttributeFilter.Parse(selector);
        var (results, warnings) = AttributeFilter.ApplyWithWarnings(_handler.Query(selector), filters);
        var textFilter = req.GetArgOrNull("text");
        if (!string.IsNullOrEmpty(textFilter))
            results = results.Where(n => n.Text != null && n.Text.Contains(textFilter, StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var w in warnings) Console.Error.WriteLine(w);
        Console.WriteLine(OutputFormatter.FormatNodes(results, format));
    }

    private void ExecuteSet(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var properties = req.GetProps();
        var unsupported = _handler.Set(path, properties);
        var applied = properties.Where(kv => !unsupported.Contains(kv.Key)).ToList();
        if (applied.Count > 0)
            Console.WriteLine($"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}");
        else if (unsupported.Count > 0)
            Console.WriteLine($"No properties applied to {path}");
        if (unsupported.Count > 0)
            Console.Error.WriteLine($"UNSUPPORTED props (use raw-set instead): {string.Join(", ", unsupported)}");
    }

    private void ExecuteAdd(ResidentRequest req)
    {
        var parentPath = req.GetArg("parent", "/body");
        var from = req.GetArgOrNull("from");
        var position = BuildInsertPosition(req);

        if (!string.IsNullOrEmpty(from))
        {
            var resultPath = _handler.CopyFrom(from, parentPath, position);
            Console.WriteLine($"Copied to {resultPath}");
        }
        else
        {
            var type = req.GetArg("type", "");
            var properties = req.GetProps();
            var resultPath = _handler.Add(parentPath, type, position, properties);
            Console.WriteLine($"Added {type} at {resultPath}");
        }
    }

    private void ExecuteRemove(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        _handler.Remove(path);
        Console.WriteLine($"Removed {path}");
    }

    private void ExecuteMove(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var to = req.GetArgOrNull("to");
        var resultPath = _handler.Move(path, to, BuildInsertPosition(req));
        Console.WriteLine($"Moved to {resultPath}");
    }

    private static InsertPosition? BuildInsertPosition(ResidentRequest req)
    {
        var index = req.GetIntArg("index");
        var after = req.GetArgOrNull("after");
        var before = req.GetArgOrNull("before");
        if (index.HasValue) return InsertPosition.AtIndex(index.Value);
        if (after != null) return InsertPosition.AfterElement(after);
        if (before != null) return InsertPosition.BeforeElement(before);
        return null;
    }

    private void ExecuteRaw(ResidentRequest req)
    {
        var partPath = req.GetArg("part", "/document");
        var startRow = req.GetIntArg("start");
        var endRow = req.GetIntArg("end");
        var cols = req.GetCols("cols");
        Console.WriteLine(_handler.Raw(partPath, startRow, endRow, cols));
    }

    private void ExecuteRawSet(ResidentRequest req)
    {
        var partPath = req.GetArg("part", "/document");
        var xpath = req.GetArg("xpath", "");
        var action = req.GetArg("action", "");
        var xml = req.GetArgOrNull("xml");

        var errorsBefore = _handler.Validate().Select(e => e.Description).ToHashSet();
        _handler.RawSet(partPath, xpath, action, xml);

        var errorsAfter = _handler.Validate();
        var newErrors = errorsAfter.Where(e => !errorsBefore.Contains(e.Description)).ToList();
        if (newErrors.Count > 0)
        {
            Console.WriteLine($"VALIDATION: {newErrors.Count} new error(s) introduced:");
            foreach (var err in newErrors)
            {
                Console.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    private void ExecuteAddPart(ResidentRequest req)
    {
        var parent = req.GetArg("parent", "/");
        var type = req.GetArg("type", "");
        var errorsBefore = _handler.Validate().Select(e => e.Description).ToHashSet();
        var (relId, partPath) = _handler.AddPart(parent, type);
        Console.WriteLine($"Created {type} part: relId={relId} path={partPath}");

        var errorsAfter = _handler.Validate();
        var newErrors = errorsAfter.Where(e => !errorsBefore.Contains(e.Description)).ToList();
        if (newErrors.Count > 0)
        {
            Console.WriteLine($"VALIDATION: {newErrors.Count} new error(s) introduced:");
            foreach (var err in newErrors)
            {
                Console.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    private void ExecuteValidate()
    {
        var errors = _handler.Validate();
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

    private static string MakeResponse(int exitCode, string stdout, string stderr)
    {
        var response = new ResidentResponse { ExitCode = exitCode, Stdout = stdout, Stderr = stderr };
        return System.Text.Json.JsonSerializer.Serialize(response, ResidentJsonContext.Default.ResidentResponse);
    }

    // ==================== Pipe I/O helpers ====================
    //
    // On Windows, StreamReader/StreamWriter deadlock on named pipes under .NET 11
    // preview.  Raw byte I/O avoids the issue.
    // On Linux/macOS, StreamReader/StreamWriter work fine and are faster.

    private const int MaxLineLength = 1_048_576; // 1 MB safety limit

    private static async Task<string?> ReadLineFromPipeAsync(Stream pipe, CancellationToken token)
    {
        if (!OperatingSystem.IsWindows())
        {
            using var reader = new StreamReader(pipe, Encoding.UTF8, leaveOpen: true);
            return await reader.ReadLineAsync(token);
        }
        var buffer = new byte[1];
        var lineBytes = new List<byte>(256);
        while (true)
        {
            var bytesRead = await pipe.ReadAsync(buffer.AsMemory(0, 1), token);
            if (bytesRead == 0) return lineBytes.Count > 0 ? Encoding.UTF8.GetString(lineBytes.ToArray()) : null;
            if (buffer[0] == (byte)'\n')
            {
                if (lineBytes.Count > 0 && lineBytes[^1] == (byte)'\r')
                    lineBytes.RemoveAt(lineBytes.Count - 1);
                return Encoding.UTF8.GetString(lineBytes.ToArray());
            }
            if (lineBytes.Count >= MaxLineLength)
                return null;
            lineBytes.Add(buffer[0]);
        }
    }

    private static async Task WriteLineToPipeAsync(Stream pipe, string line, CancellationToken token)
    {
        if (!OperatingSystem.IsWindows())
        {
            using var writer = new StreamWriter(pipe, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
            await writer.WriteLineAsync(line.AsMemory(), token);
            return;
        }
        var bytes = Encoding.UTF8.GetBytes(line + "\n");
        await pipe.WriteAsync(bytes, token);
        await pipe.FlushAsync(token);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Delegate to the shared shutdown task. If __close__ already drove
        // shutdown, this just awaits the cached Task (no-op). If not (e.g.
        // idle timeout, SIGTERM, crash cleanup), this runs the full ordered
        // teardown. Watchdog: if shutdown exceeds 10 min, force exit so the
        // process can never hang on a stuck handler dispose.
        try
        {
            if (!ShutdownAsync().Wait(TimeSpan.FromMinutes(10)))
            {
                Console.Error.WriteLine("Warning: shutdown timed out after 10 minutes, forcing exit.");
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: shutdown error: {ex.Message}");
        }

        try { _commandLock.Dispose(); } catch { }
        try { _mainCts.Dispose(); } catch { }
        try { _pingCts.Dispose(); } catch { }
        try { _idleCts.Dispose(); } catch { }
    }

    /// <summary>
    /// Idempotent, ordered resident shutdown. Safe to call from any thread
    /// and from every teardown entrypoint (__close__, idle watchdog,
    /// Dispose, ProcessExit, Ctrl+C) — all callers await the same cached
    /// <see cref="Task"/>.
    ///
    /// Ordering enforces the critical invariant
    /// <c>ping responds ⇔ handler holds the file</c>:
    ///
    ///   1. Cancel _mainCts → main command loop stops accepting NEW work.
    ///      Ping + idle are still live under _pingCts.
    ///   2. Kick the main pipe to unstick any in-flight WaitForConnectionAsync.
    ///   3. Drain _commandLock → the one in-flight command (if any) finishes.
    ///   4. Dispose the document handler → in-memory tree written to disk,
    ///      file lock released. This is the slow, load-bearing step.
    ///   5. Cancel _pingCts → ping responder and idle watchdog stop.
    ///   6. Kick the ping pipe to unstick its WaitForConnectionAsync.
    ///
    /// Between (1) and (4) the ping pipe is intentionally kept alive so
    /// clients can observe "resident still holds the file" and behave
    /// accordingly (return busy, retry, etc). Fallback paths that probe
    /// via <see cref="ResidentClient.TryConnect"/> therefore get a
    /// consistent answer: ping live ⇒ do NOT try to open the file
    /// directly, ping dead ⇒ safe to open.
    /// </summary>
    private Task ShutdownAsync()
    {
        lock (_shutdownLock)
        {
            return _shutdownTask ??= Task.Run(DoShutdownAsync);
        }
    }

    private async Task DoShutdownAsync()
    {
        // 1. Stop accepting new main-pipe commands. Ping responder and
        //    idle watchdog remain live under _pingCts.
        try { _mainCts.Cancel(); } catch (ObjectDisposedException) { }

        // 2. Kick the main pipe to wake WaitForConnectionAsync.
        KickPipe(_pipeName);

        // 3. Drain any currently-executing command. Typical command takes
        //    tens of ms (reads) up to a few hundred ms (writes); the 10 min
        //    bound matches the outer Dispose watchdog so a stuck command is
        //    caught by exactly one tier, not two.
        try
        {
            if (await _commandLock.WaitAsync(TimeSpan.FromMinutes(10)))
            {
                try { _commandLock.Release(); } catch (SemaphoreFullException) { }
            }
            else
            {
                Console.Error.WriteLine("Warning: timeout waiting for in-flight command to drain.");
            }
        }
        catch (ObjectDisposedException) { /* _commandLock already disposed */ }

        // 4. Dispose the handler. Slow (writes the OpenXML tree back to
        //    disk and closes the file handle). The ping pipe is still
        //    live right now, so any TryResident caller will correctly
        //    conclude "resident still owns the file".
        try { _handler.Dispose(); }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: handler dispose error: {ex.Message}");
        }

        // 5. NOW cancel ping + idle. Clients observing the ping pipe from
        //    this moment on will see it dead and can safely open the file
        //    directly.
        try { _pingCts.Cancel(); } catch (ObjectDisposedException) { }

        // 6. Kick ping pipe so RunPingResponderAsync unsticks.
        KickPipe(_pipeName + "-ping");
    }

    private static void KickPipe(string pipeName)
    {
        try
        {
            using var kick = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
            kick.Connect(500);
        }
        catch { }
    }
}

public class ResidentRequest
{
    public string Command { get; set; } = "";
    public Dictionary<string, string> Args { get; set; } = new();
    public Dictionary<string, string>? Props { get; set; }
    public bool Json { get; set; }

    public string GetArg(string key, string defaultValue = "")
    {
        return Args.TryGetValue(key, out var val) ? val : defaultValue;
    }

    public string? GetArgOrNull(string key)
    {
        return Args.TryGetValue(key, out var val) ? val : null;
    }

    public int? GetIntArg(string key)
    {
        if (Args.TryGetValue(key, out var val) && int.TryParse(val, out var n))
            return n;
        return null;
    }

    public HashSet<string>? GetCols(string key)
    {
        var val = GetArgOrNull(key);
        if (val == null) return null;
        return new HashSet<string>(val.Split(',').Select(c => c.Trim().ToUpperInvariant()));
    }

    public Dictionary<string, string> GetProps()
    {
        return Props ?? new Dictionary<string, string>();
    }
}

public class ResidentResponse
{
    public int ExitCode { get; set; }
    public string Stdout { get; set; } = "";
    public string Stderr { get; set; } = "";
}

[System.Text.Json.Serialization.JsonSourceGenerationOptions]
[System.Text.Json.Serialization.JsonSerializable(typeof(ResidentRequest))]
[System.Text.Json.Serialization.JsonSerializable(typeof(ResidentResponse))]
internal partial class ResidentJsonContext : System.Text.Json.Serialization.JsonSerializerContext;
