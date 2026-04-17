// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// CONSISTENCY(watch-isolation): 本文件不引用 OfficeCli.Handlers,不打开文件,不写盘。
// 见 CLAUDE.md "Watch Server Rules"。要放宽这条红线,
// grep "CONSISTENCY(watch-isolation)" 找全 watch 子系统所有文件项目级一起评审。

using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Pure SSE relay server. Never opens the document file.
/// Receives pre-rendered HTML from command processes via named pipe,
/// forwards to browsers via SSE.
/// </summary>
internal class WatchServer : IDisposable
{
    private readonly string _filePath;
    private readonly string _pipeName;
    private readonly int _port;
    private readonly TcpListener _tcpListener;
    private readonly List<NetworkStream> _sseClients = new();
    private readonly object _sseLock = new();
    private CancellationTokenSource _cts = new();
    private string _currentHtml = "";
    private int _version = 0;
    private bool _disposed;
    private DateTime _lastActivityTime = DateTime.UtcNow;
    private readonly TimeSpan _idleTimeout;

    // Shared shutdown Task so every teardown entrypoint — idle watchdog,
    // unwatch command, SIGTERM/SIGINT, Dispose — converges on a single
    // ordered sequence. Before this, idle/unwatch just called
    // _cts.Cancel() and hoped the async chain would unwind; but
    // TcpListener.AcceptTcpClientAsync on macOS under .NET 10 does NOT
    // reliably honour the cancellation token, so the main loop would
    // hang indefinitely in `await AcceptTcpClientAsync(token)` and the
    // process would ignore SIGINT for 15+ seconds (observed in
    // stress test) until something else kicked the TCP listener.
    private readonly object _shutdownLock = new();
    private Task? _shutdownTask;

    // Current selection — paths of elements selected in any connected browser.
    // Single shared list (last-write-wins): all browsers viewing the same file see
    // the same selection. CLI reads this via the named pipe "get-selection" command.
    //
    // CONSISTENCY(path-stability): selection 和 mark 共享同一套裸位置寻址契约,
    // 没有指纹/漂移检测。要升级到稳定 ID,grep "CONSISTENCY(path-stability)"
    // 找全所有 deferred 站点项目级一起改。见 CLAUDE.md "Design Principles"。
    private List<string> _currentSelection = new();
    private readonly object _selectionLock = new();

    // Current marks — advisory annotations attached to document paths. Live in
    // memory only. Server never opens the document and never inspects DOM —
    // marks are pure metadata; the browser computes match positions client-side.
    //
    // CONSISTENCY(path-stability): 元素删除/位置漂移的处理刻意和 selection 一致 ——
    // 裸位置寻址,无指纹,无漂移检测。stale 仅在 path 解析失败或 find 不命中时由
    // 客户端报告设置。见 CLAUDE.md "Design Principles" + "Watch Server Rules"。
    // 要修复成稳定 ID 路径,grep "CONSISTENCY(path-stability)" 找全所有 deferred 站点
    // (selection / mark / 未来其它 path 消费者)项目级一起改,不要在 mark 单点改。
    private readonly List<WatchMark> _currentMarks = new();
    private readonly object _marksLock = new();
    private int _marksVersion = 0;
    private int _nextMarkId = 1;

    private const string WaitingHtml = """
        <html><head><meta charset="utf-8"><title>Watching...</title>
        <style>body{font-family:system-ui;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f5f5f5;color:#666;}
        .msg{text-align:center;}</style></head>
        <body><div class="msg"><h2>Waiting for first update...</h2><p>Run an officecli command to see the preview.</p></div></body></html>
        """;

    // SSE script content loaded from embedded resources (watch-sse-core.js + watch-overlay.js).
    // Layer 1 (sse-core) handles SSE connection, DOM updates, word diff/patch, slide ops.
    // Layer 2 (overlay) handles selection, marks, rubber-band, CSS injection.
    // Coupling: Layer 1 calls window._watchReapplyHook() after DOM mutations;
    //           Layer 2 sets that hook to reapplyDecorations().
    private static readonly Lazy<string> _sseScriptBlock = new(() =>
    {
        var core = LoadWatchResource("Resources.watch-sse-core.js");
        var overlay = LoadWatchResource("Resources.watch-overlay.js");
        return $"<script>\n{core}\n</script>\n<script>\n{overlay}\n</script>";
    });

    // Test access: allows tests to verify SSE script content without reflection on a const field.
    internal static string SseScriptContent => _sseScriptBlock.Value;

    private static string LoadWatchResource(string name)
    {
        var assembly = typeof(WatchServer).Assembly;
        var fullName = $"OfficeCli.{name}";
        using var stream = assembly.GetManifestResourceStream(fullName);
        if (stream == null) return $"/* Resource not found: {fullName} */";
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    // Idle timeout is configurable via OFFICECLI_WATCH_IDLE_SECONDS so
    // tests can exercise the auto-shutdown path in seconds instead of
    // minutes. Callers that pass an explicit TimeSpan (tests that need
    // fixed values) bypass the env var. Valid range: 1s .. 24h.
    private static TimeSpan ResolveIdleTimeout()
    {
        var raw = Environment.GetEnvironmentVariable("OFFICECLI_WATCH_IDLE_SECONDS");
        if (!string.IsNullOrWhiteSpace(raw)
            && int.TryParse(raw, out var secs)
            && secs >= 1 && secs <= 86400)
        {
            return TimeSpan.FromSeconds(secs);
        }
        return TimeSpan.FromMinutes(5);
    }

    public WatchServer(string filePath, int port, TimeSpan? idleTimeout = null, string? initialHtml = null)
    {
        _filePath = Path.GetFullPath(filePath);
        _pipeName = GetWatchPipeName(_filePath);
        _port = port;
        _idleTimeout = idleTimeout ?? ResolveIdleTimeout();
        _tcpListener = new TcpListener(IPAddress.Loopback, _port);
        if (!string.IsNullOrEmpty(initialHtml))
            _currentHtml = initialHtml;
    }

    public static string GetWatchPipeName(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
            fullPath = fullPath.ToUpperInvariant();
        var hash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)))[..16];
        return $"officecli-watch-{hash}";
    }

    /// <summary>
    /// Check if another watch process is already running for this file.
    /// Returns the port number if running, or null if not.
    /// </summary>
    public static int? GetExistingWatchPort(string filePath)
    {
        try
        {
            int? result = null;
            var task = Task.Run(() =>
            {
                var pipeName = GetWatchPipeName(filePath);
                using var client = new System.IO.Pipes.NamedPipeClientStream(".", pipeName, System.IO.Pipes.PipeDirection.InOut);
                client.Connect(100);
                var noBom = new UTF8Encoding(false);
                using var writer = new StreamWriter(client, noBom, leaveOpen: true) { AutoFlush = true };
                writer.WriteLine("ping");
                writer.Flush();
                using var reader = new StreamReader(client, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                var response = reader.ReadLine();
                result = int.TryParse(response, out var port) ? port : 0;
            });
            return task.Wait(TimeSpan.FromSeconds(2)) ? result : null;
        }
        catch
        {
            return null; // not running
        }
    }

    public static bool IsWatching(string filePath)
    {
        return GetExistingWatchPort(filePath).HasValue;
    }

    public async Task RunAsync(CancellationToken externalToken = default)
    {
        // Prevent duplicate watch processes for the same file
        var existingPort = GetExistingWatchPort(_filePath);
        if (existingPort.HasValue)
        {
            var url = existingPort.Value > 0 ? $" at http://localhost:{existingPort.Value}" : "";
            throw new InvalidOperationException($"Another watch process is already running{url} for {_filePath}");
        }

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(_cts.Token, externalToken);
        var token = linkedCts.Token;

        _tcpListener.Start();
        Console.WriteLine($"Watch: http://localhost:{_port}");
        Console.WriteLine($"Watching: {_filePath}");
        Console.WriteLine("Press Ctrl+C to stop.");

        // Hook graceful shutdown signals. Cooperatively terminating a
        // watch process needs to (a) stop the TCP listener — the only
        // reliable way to kick AcceptTcpClientAsync on macOS, which
        // does NOT honour cancellation tokens on .NET 10 — and (b)
        // delete the $TMPDIR/CoreFxPipe_ socket file (.NET doesn't,
        // BUG-BT-003). Both steps happen inside StopAsync.
        //
        // Two signal paths cover the realistic user scenarios:
        //
        // 1. PosixSignalRegistration for SIGTERM / SIGHUP / SIGQUIT.
        //    These are the usual "kill this daemon" signals; they fire
        //    whether or not the process has a controlling TTY. Works
        //    reliably for `pkill officecli`, launcher kill, and
        //    terminal-close-while-backgrounded.
        //
        // 2. Console.CancelKeyPress for Ctrl+C (SIGINT). This fires
        //    when watch is running in the foreground of an interactive
        //    terminal — the realistic user scenario for "I pressed
        //    Ctrl+C to stop the watch I just started".
        //
        // Known limitation: sending SIGINT or SIGQUIT to a BACKGROUNDED
        // watch process (e.g. `officecli watch file & ; kill -INT %1`)
        // does not trigger either path because .NET's runtime gates
        // SIGINT/SIGQUIT handling on having a controlling TTY. This is
        // not a realistic daemon-termination pattern — callers who
        // need to stop a backgrounded watch should use `officecli
        // unwatch file` or SIGTERM, both of which work.
        var signalRegs = new List<PosixSignalRegistration>();
        void DoShutdownFromSignal()
        {
            try { StopAsync().Wait(TimeSpan.FromSeconds(10)); } catch { }
            Environment.Exit(0);
        }
        void HandleSignal(PosixSignalContext ctx)
        {
            ctx.Cancel = true;
            DoShutdownFromSignal();
        }
        void TryRegister(PosixSignal sig)
        {
            try { signalRegs.Add(PosixSignalRegistration.Create(sig, HandleSignal)); }
            catch (PlatformNotSupportedException) { /* host doesn't support this signal */ }
        }
        TryRegister(PosixSignal.SIGTERM);
        TryRegister(PosixSignal.SIGHUP);
        TryRegister(PosixSignal.SIGQUIT);

        ConsoleCancelEventHandler cancelHandler = (_, e) =>
        {
            e.Cancel = true;
            DoShutdownFromSignal();
        };
        Console.CancelKeyPress += cancelHandler;

        var pipeTask = RunPipeListenerAsync(token);
        var idleTask = RunIdleWatchdogAsync(token);

        while (!token.IsCancellationRequested)
        {
            try
            {
                var client = await _tcpListener.AcceptTcpClientAsync(token);
                _ = HandleClientAsync(client, token);
            }
            catch (OperationCanceledException) { break; }
            catch (SocketException) { break; }
            catch (ObjectDisposedException) { break; }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Watch HTTP error: {ex.Message}");
            }
        }

        // Main loop exited — drive the shared shutdown path. This cleans
        // up TCP listener, pipe listener, CoreFxPipe_ socket, and SSE
        // clients in order. Idempotent, so signal-driven and
        // cancellation-driven paths both converge here safely.
        try { await StopAsync(); } catch { }

        try { await pipeTask; } catch (OperationCanceledException) { }
        try { await idleTask; } catch (OperationCanceledException) { }

        foreach (var reg in signalRegs)
            try { reg.Dispose(); } catch { }
        Console.CancelKeyPress -= cancelHandler;
    }

    /// <summary>
    /// Idempotent, ordered shutdown. Every teardown path (idle watchdog,
    /// unwatch pipe command, SIGTERM/SIGINT/SIGHUP, Dispose) funnels
    /// through this method and awaits the same cached Task.
    ///
    /// Order:
    ///   1. Cancel _cts — idle watchdog and pipe listener exit their loops.
    ///   2. Call TcpListener.Stop() — only reliable way to unstick
    ///      AcceptTcpClientAsync on macOS under .NET 10.
    ///   3. Close all live SSE client streams so RunSseClientAsync
    ///      coroutines drop their references.
    ///   4. Kick the pipe listener via a local NamedPipeClientStream
    ///      connect so RunPipeListenerAsync unsticks on Windows (where
    ///      WaitForConnectionAsync doesn't honour cancellation).
    ///   5. On Unix, delete the stale $TMPDIR/CoreFxPipe_ socket file
    ///      (.NET doesn't clean it up — BUG-BT-003).
    /// </summary>
    public Task StopAsync()
    {
        lock (_shutdownLock)
        {
            return _shutdownTask ??= Task.Run(DoStopAsync);
        }
    }

    private async Task DoStopAsync()
    {
        // 1. Signal everything to stop.
        try { _cts.Cancel(); } catch (ObjectDisposedException) { }

        // 2. Stop the TCP listener. AcceptTcpClientAsync(token) on macOS
        //    under .NET 10 does not reliably respect cancellation; Stop()
        //    force-closes the underlying socket which makes the pending
        //    accept throw ObjectDisposedException and unwind the loop.
        try { _tcpListener.Stop(); } catch { }

        // 3. Close live SSE streams so the per-client coroutines unwind
        //    promptly. (They would eventually notice token cancellation,
        //    but a blocking write to a dead client can hang for seconds.)
        lock (_sseLock)
        {
            foreach (var s in _sseClients)
            {
                try { s.Close(); } catch { }
            }
            _sseClients.Clear();
        }

        // 4. Kick the pipe listener out of WaitForConnectionAsync.
        try
        {
            using var kick = new System.IO.Pipes.NamedPipeClientStream(
                ".", _pipeName, System.IO.Pipes.PipeDirection.InOut);
            kick.Connect(500);
        }
        catch { }

        // 5. Delete the stale CoreFxPipe_ socket on Unix. .NET does not
        //    do this on its own (BUG-BT-003 — fuzzer found 302 stale
        //    files). Run here in StopAsync rather than Dispose so it
        //    also works when the process exits via SIGTERM signal path.
        if (!OperatingSystem.IsWindows())
        {
            try
            {
                var sockPath = Path.Combine(Path.GetTempPath(), "CoreFxPipe_" + _pipeName);
                if (File.Exists(sockPath)) File.Delete(sockPath);
            }
            catch { /* best-effort cleanup */ }
        }

        // Small yield so any synchronous continuations scheduled on the
        // now-cancelled token get a chance to run before the caller
        // proceeds. Not strictly required for correctness.
        await Task.Yield();
    }

    private async Task RunIdleWatchdogAsync(CancellationToken token)
    {
        var checkInterval = TimeSpan.FromSeconds(Math.Min(30, Math.Max(1, _idleTimeout.TotalSeconds / 2)));
        while (!token.IsCancellationRequested)
        {
            await Task.Delay(checkInterval, token);
            int clientCount;
            lock (_sseLock) { clientCount = _sseClients.Count; }
            if (clientCount == 0 && DateTime.UtcNow - _lastActivityTime > _idleTimeout)
            {
                Console.WriteLine("Watch: idle timeout, shutting down.");
                // Go through the shared ordered shutdown path instead of
                // raw-cancelling _cts, so TcpListener.Stop() gets called
                // and the main loop doesn't hang waiting for an accept
                // that never completes.
                _ = StopAsync();
                break;
            }
        }
    }

    private async Task RunPipeListenerAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            var server = new System.IO.Pipes.NamedPipeServerStream(
                _pipeName, System.IO.Pipes.PipeDirection.InOut,
                System.IO.Pipes.NamedPipeServerStream.MaxAllowedServerInstances,
                System.IO.Pipes.PipeTransmissionMode.Byte,
                System.IO.Pipes.PipeOptions.Asynchronous);
            try
            {
                await server.WaitForConnectionAsync(token);
            }
            catch (OperationCanceledException) { await server.DisposeAsync(); break; }
            catch { await server.DisposeAsync(); continue; }

            // Handle the client on a background task and immediately loop back
            // to accept another connection. This avoids a tiny window where the
            // pipe is not listening between iterations and back-to-back CLI
            // calls (e.g. multiple mark adds in a tight test loop) get refused.
            _ = Task.Run(async () =>
            {
                using (server)
                {
                    try { await HandleSinglePipeClientAsync(server, token); }
                    catch { /* ignore individual client errors */ }
                }
            }, token);
        }
    }

    private async Task HandleSinglePipeClientAsync(System.IO.Pipes.NamedPipeServerStream server, CancellationToken token)
    {
            try
            {
                var noBom = new UTF8Encoding(false);
                using var reader = new StreamReader(server, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                using var writer = new StreamWriter(server, noBom, leaveOpen: true) { AutoFlush = true };

                var message = await reader.ReadLineAsync(token);
                _lastActivityTime = DateTime.UtcNow;

                if (message == "close")
                {
                    await writer.WriteLineAsync("ok".AsMemory(), token);
                    Console.WriteLine("Watch closed by remote command.");
                    // Go through shared shutdown — idempotent, ordered,
                    // also cleans up CoreFxPipe_ socket on Unix.
                    _ = StopAsync();
                    return;
                }
                else if (message == "ping")
                {
                    // Return port so callers can find the existing watch URL
                    await writer.WriteLineAsync(_port.ToString().AsMemory(), token);
                }
                else if (message == "get-selection")
                {
                    // Return current selection as a JSON array of paths.
                    // Empty selection → "[]". Never null.
                    string[] snapshot;
                    lock (_selectionLock) { snapshot = _currentSelection.ToArray(); }
                    var json = JsonSerializer.Serialize(snapshot, WatchSelectionJsonOptions.StringArrayInfo);
                    await writer.WriteLineAsync(json.AsMemory(), token);
                }
                else if (message == "get-marks")
                {
                    // Return {"version":N,"marks":[...]} so callers can do CAS-style
                    // detection. Empty marks → []. Never null.
                    // Uses Relaxed options so CJK content emits literal chars.
                    WatchMark[] snapshot;
                    int version;
                    lock (_marksLock)
                    {
                        snapshot = _currentMarks.ToArray();
                        version = _marksVersion;
                    }
                    var resp = new MarksResponse { Version = version, Marks = snapshot };
                    var payload = JsonSerializer.Serialize(resp, WatchMarkJsonOptions.MarksResponseInfo);
                    await writer.WriteLineAsync(payload.AsMemory(), token);
                }
                else if (message != null && message.StartsWith("mark ", StringComparison.Ordinal))
                {
                    // "mark <json>" — add a mark, return assigned id
                    var payload = message.Substring(5);
                    var resp = HandleMarkAdd(payload);
                    await writer.WriteLineAsync(resp.AsMemory(), token);
                }
                else if (message != null && message.StartsWith("unmark ", StringComparison.Ordinal))
                {
                    // "unmark <json>" — remove marks by path or all
                    var payload = message.Substring(7);
                    var resp = HandleMarkRemove(payload);
                    await writer.WriteLineAsync(resp.AsMemory(), token);
                }
                else if (message != null)
                {
                    await writer.WriteLineAsync("ok".AsMemory(), token);
                    // Try to parse as WatchMessage JSON
                    HandleWatchMessage(message);
                }
            }
            catch (OperationCanceledException) { return; }
            catch { /* ignore pipe errors */ }
    }

    private void HandleWatchMessage(string json)
    {
        try
        {
            var msg = JsonSerializer.Deserialize(json, WatchMessageJsonContext.Default.WatchMessage);
            if (msg == null) return;

            var oldHtml = _currentHtml;
            var baseVersion = _version;

            // Always update cached full HTML when provided (authoritative snapshot)
            if (!string.IsNullOrEmpty(msg.FullHtml))
            {
                _currentHtml = msg.FullHtml;
            }

            // Apply incremental patch when no full HTML was provided
            if (string.IsNullOrEmpty(msg.FullHtml))
            {
                if (msg.Action == "replace" && msg.Slide > 0 && msg.Html != null)
                    _currentHtml = PatchSlideInHtml(_currentHtml, msg.Slide, msg.Html);
                else if (msg.Action == "add" && msg.Html != null)
                    _currentHtml = AppendSlideToHtml(_currentHtml, msg.Html);
                else if (msg.Action == "remove" && msg.Slide > 0)
                    _currentHtml = RemoveSlideFromHtml(_currentHtml, msg.Slide);
            }

            _version++;

            // Reconcile all marks against the freshly updated snapshot. Flips
            // stale flags and refreshes matched_text when the underlying text
            // changed. CONSISTENCY(path-stability): same naive resolve used on
            // initial add, no fingerprint.
            ReconcileAllMarks();

            // Word: try block-level diff instead of full refresh
            if (msg.Action == "full" && !string.IsNullOrEmpty(msg.FullHtml)
                && !string.IsNullOrEmpty(oldHtml) && oldHtml.Contains("data-block=\"1\""))
            {
                var patches = ComputeWordPatches(oldHtml, msg.FullHtml);
                // Check if CSS styles changed
                var oldStyle = ExtractStyleBlock(oldHtml);
                var newStyle = ExtractStyleBlock(msg.FullHtml);
                var styleChanged = oldStyle != newStyle;

                if (patches != null || styleChanged)
                {
                    patches ??= new List<WordPatch>();
                    if (styleChanged)
                        patches.Insert(0, new WordPatch { Op = "style", Block = 0, Html = newStyle });
                    SendSseWordPatch(patches, _version, baseVersion, msg.ScrollTo);
                    return;
                }
            }

            // Excel: try row-level diff instead of full refresh
            if (msg.Action == "full" && !string.IsNullOrEmpty(msg.FullHtml)
                && !string.IsNullOrEmpty(oldHtml) && oldHtml.Contains("data-row=\""))
            {
                var excelPatches = ComputeExcelPatches(oldHtml, msg.FullHtml);
                var oldStyle = ExtractStyleBlock(oldHtml);
                var newStyle = ExtractStyleBlock(msg.FullHtml);
                var styleChanged = oldStyle != newStyle;

                if (excelPatches != null || styleChanged)
                {
                    excelPatches ??= new List<(string Op, string Row, string? Html)>();
                    if (styleChanged)
                        excelPatches.Insert(0, ("style", "", newStyle));
                    SendSseExcelPatch(excelPatches, _version, baseVersion, msg.ScrollTo);
                    return;
                }
            }

            // Forward to SSE clients (full or PPT incremental)
            SendSseEvent(msg.Action, msg.Slide, msg.Html, msg.ScrollTo, _version);
        }
        catch
        {
            // Legacy format or parse error — treat as full refresh signal
            _version++;
            SendSseEvent("full", 0, null, null, _version);
        }
    }

    // ==================== Marks ====================

    /// <summary>
    /// Add a new mark. Normalizes find: if regex flag (truthy via the find
    /// payload's "regex" field would be parsed by the CLI side; the server
    /// receives the canonical form already wrapped as r"..." or literal).
    /// However we ALSO accept the bare-find form here so that callers that
    /// don't pre-wrap still get correct behaviour. The CLI passes either
    /// the literal or a pre-wrapped r"..." string.
    /// </summary>
    internal string HandleMarkAdd(string json)
    {
        try
        {
            var req = JsonSerializer.Deserialize(json, WatchMarkJsonContext.Default.MarkRequest);
            if (req == null)
                return "{\"error\":\"invalid request\"}";

            // BUG-FUZZER-003/004: path hardening.
            //   1. Normalize: Trim() strips ASCII + Unicode whitespace from edges.
            //   2. Reject whitespace-only paths (IsNullOrWhiteSpace catches NBSP,
            //      U+3000 ideographic space, etc.).
            //   3. Require leading '/': zero-width space U+200B and BOM U+FEFF
            //      are not .NET whitespace but are never valid data-path prefixes,
            //      so a StartsWith('/') check also filters them out.
            //   4. Store the trimmed form so later `unmark --path /body/p[1]`
            //      matches what the user typed, not `" /body/p[1] "` with padding.
            // BUG-BT-R303: error messages must be actionable for AI agents — say
            // what the accepted format is, not just "invalid".
            var trimmedPath = req.Path?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(trimmedPath) || !trimmedPath.StartsWith("/"))
                return "{\"error\":\"invalid path: must start with '/' (e.g. /body/p[1] for Word, /slide[1]/shape[@id=N] for PowerPoint)\"}";

            // BUG-TESTER-002: validate color server-side. The browser sets
            // el.style.backgroundColor = mark.color verbatim, so an unsanitized
            // value injects CSS into every connected SSE client. Server is the
            // single trust boundary for both human-typed CLI and machine agents.
            // CONSISTENCY(mark-color-validation): one validator, both Add and
            // any future Set/update path must call IsValidMarkColor.
            //
            // BUG-FUZZER-001: Trim() before validation AND before storage, so
            // `"red\n"` doesn't end up stored as `"red\n"` after being accepted
            // (the validator trims for matching but used to leave the raw form
            // in the stored mark, causing a validator-vs-storage inconsistency).
            var trimmedColor = req.Color?.Trim();
            // BUG-A-R2-M01: accept bare hex (FF00FF, F0F) for consistency with the
            // rest of officecli's color parsers. The validator below requires the
            // canonical #-prefixed form, so promote 3/6/8-digit bare hex to that
            // form before validation. Anything else (named colors, rgb(...),
            // already-hashed hex) passes through unchanged.
            trimmedColor = NormalizeMarkColorInput(trimmedColor);
            // BUG-BT-R303: actionable error message — list the accepted formats
            // so AI agents can self-correct without reading the source.
            if (!string.IsNullOrEmpty(trimmedColor) && !IsValidMarkColor(trimmedColor))
                return "{\"error\":\"invalid color: accepted forms are #RGB / #RRGGBB / #RRGGBBAA hex (with or without # prefix), rgb(r,g,b), rgba(r,g,b,a), or named colors (red, blue, yellow, orange, green, purple, ...)\"}";

            var mark = new WatchMark
            {
                Path = trimmedPath,
                Find = req.Find,
                Color = string.IsNullOrEmpty(trimmedColor) ? "#ffeb3b" : trimmedColor,
                Note = req.Note,
                Tofix = req.Tofix,
                MatchedText = Array.Empty<string>(),
                Stale = false,
                CreatedAt = DateTime.UtcNow,
            };

            string assignedId;
            WatchMark[] snapshot;
            string htmlSnapshot;
            lock (_marksLock)
            {
                assignedId = _nextMarkId.ToString();
                _nextMarkId++;
                mark.Id = assignedId;
                // Snapshot _currentHtml under the lock so a concurrent
                // full-refresh can't race the resolve step.
                htmlSnapshot = _currentHtml;
                var resolved = ResolveMark(mark, htmlSnapshot);
                _currentMarks.Add(resolved);
                _marksVersion++;
                snapshot = _currentMarks.ToArray();
            }
            _lastActivityTime = DateTime.UtcNow;
            BroadcastMarkUpdate(snapshot);

            return JsonSerializer.Serialize(
                new MarkResponse { Id = assignedId },
                WatchMarkJsonContext.Default.MarkResponse);
        }
        catch
        {
            return "{\"error\":\"parse failed\"}";
        }
    }

    /// <summary>
    /// Remove marks. UnmarkRequest must have either Path set, or All=true,
    /// not both. Returns the number of marks removed.
    /// </summary>
    internal string HandleMarkRemove(string json)
    {
        try
        {
            var req = JsonSerializer.Deserialize(json, WatchMarkJsonContext.Default.UnmarkRequest);
            if (req == null) return "{\"removed\":0}";

            int removed = 0;
            WatchMark[] snapshot;
            lock (_marksLock)
            {
                if (req.All)
                {
                    removed = _currentMarks.Count;
                    _currentMarks.Clear();
                }
                else
                {
                    // BUG-FUZZER-003/004: Trim and require leading '/' for symmetry
                    // with HandleMarkAdd. Without Trim a `unmark --path " /p[1] "`
                    // would silently miss a mark added as `/p[1]` and vice versa.
                    var unmarkPath = req.Path?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(unmarkPath) && unmarkPath.StartsWith("/"))
                    {
                        removed = _currentMarks.RemoveAll(m =>
                            string.Equals(m.Path, unmarkPath, StringComparison.Ordinal));
                    }
                }
                if (removed > 0) _marksVersion++;
                snapshot = _currentMarks.ToArray();
            }
            _lastActivityTime = DateTime.UtcNow;
            if (removed > 0) BroadcastMarkUpdate(snapshot);

            return JsonSerializer.Serialize(
                new UnmarkResponse { Removed = removed },
                WatchMarkJsonContext.Default.UnmarkResponse);
        }
        catch
        {
            return "{\"removed\":0}";
        }
    }

    /// <summary>Test-only accessor for current marks snapshot.</summary>
    internal WatchMark[] GetMarksSnapshot()
    {
        lock (_marksLock) { return _currentMarks.ToArray(); }
    }

    /// <summary>Test-only accessor for the current marks version.</summary>
    internal int GetMarksVersion()
    {
        lock (_marksLock) { return _marksVersion; }
    }

    /// <summary>
    /// Test-only hook: install a full HTML snapshot synchronously and trigger
    /// mark reconciliation. Used by WatchMarkTests to verify ResolveMark without
    /// racing the pipe's "ack first, process later" ordering.
    /// </summary>
    internal void ApplyFullHtmlForTests(string html)
    {
        _currentHtml = html ?? "";
        _version++;
        ReconcileAllMarks();
    }

    // -------- Mark resolution (server-side reconcile) --------
    //
    // CONSISTENCY(path-stability): resolution uses naive positional
    // data-path lookup — no fingerprinting, no drift detection. If an
    // element is later removed or its find target no longer matches,
    // the mark is flipped to Stale=true with MatchedText=[]. Same
    // limitations as selection. grep "CONSISTENCY(path-stability)" for
    // all deferred sites that should move together if we ever switch
    // to stable IDs. See CLAUDE.md "Watch Server Rules".
    //
    // watch-isolation: this code runs pure-regex string-scraping on
    // the html snapshot already cached in _currentHtml. It does not
    // open the document, does not depend on OfficeCli.Handlers, and
    // does not reference any DOM parser. A real HTML parser would be
    // more correct but would introduce coupling; the MVP trades
    // precision for isolation and matches the browser-side
    // applyMarks() fallback behaviour.

    private static readonly System.Text.RegularExpressions.Regex _tagStripRx =
        new("<[^>]+>", System.Text.RegularExpressions.RegexOptions.Compiled);

    // BUG-TESTER-001: ResolveMark accepts arbitrary user regex via r"..." find
    // strings. A catastrophically backtracking pattern (e.g. r"(a+)+$") against
    // a long input would freeze the watch reconcile loop indefinitely. Bound
    // every user-supplied regex evaluation with this match timeout.
    private static readonly TimeSpan MarkRegexMatchTimeout = TimeSpan.FromMilliseconds(500);

    // BUG-TESTER-003: <script> and <style> bodies must be removed entirely
    // before tag-stripping, otherwise their inner text leaks into find matching
    // (e.g. find="secret" hits "<script>secret data</script>"). These regexes
    // strip the element including children, case-insensitive, dot-matches-newline.
    private static readonly System.Text.RegularExpressions.Regex _scriptBodyRx =
        new("<script\\b[^>]*>.*?</script\\s*>",
            System.Text.RegularExpressions.RegexOptions.Compiled
            | System.Text.RegularExpressions.RegexOptions.IgnoreCase
            | System.Text.RegularExpressions.RegexOptions.Singleline);
    private static readonly System.Text.RegularExpressions.Regex _styleBodyRx =
        new("<style\\b[^>]*>.*?</style\\s*>",
            System.Text.RegularExpressions.RegexOptions.Compiled
            | System.Text.RegularExpressions.RegexOptions.IgnoreCase
            | System.Text.RegularExpressions.RegexOptions.Singleline);

    // BUG-TESTER-002: server-side color whitelist for mark.color. Anything
    // accepted here gets written verbatim into el.style.backgroundColor on
    // every connected browser, so the validator must REJECT anything that
    // isn't unambiguously a color value. Three accepted shapes:
    //   1. #RGB / #RRGGBB / #RRGGBBAA hex
    //   2. rgb(r,g,b) / rgba(r,g,b,a) with numeric components
    //   3. one of the named colors in MarkNamedColors
    // CONSISTENCY(mark-color-validation): grep this tag if expanding the set.
    private static readonly System.Text.RegularExpressions.Regex _hexColorRx =
        new("^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$",
            System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _rgbFuncRx =
        new("^rgba?\\(\\s*\\d+(?:\\.\\d+)?\\s*,\\s*\\d+(?:\\.\\d+)?\\s*,\\s*\\d+(?:\\.\\d+)?(?:\\s*,\\s*\\d+(?:\\.\\d+)?)?\\s*\\)$",
            System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly HashSet<string> MarkNamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        "red", "green", "blue", "yellow", "orange", "purple", "pink", "cyan",
        "magenta", "brown", "black", "white", "gray", "grey", "lime", "teal",
        "navy", "olive", "maroon", "silver", "gold", "transparent",
    };

    // BUG-A-R2-M01 / BUG-TESTER-R302: Promote bare 3-, 6-, or 8-digit hex to
    // #-prefixed form so the validator and storage match the rest of officecli's
    // color convention. Returns the input unchanged for any other shape (named,
    // rgb(...), already #-prefixed, or null/empty). Idempotent.
    private static readonly System.Text.RegularExpressions.Regex _bareHex6Rx =
        new("^[0-9a-fA-F]{6}$", System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _bareHex3Rx =
        new("^[0-9a-fA-F]{3}$", System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _bareHex8Rx =
        new("^[0-9a-fA-F]{8}$", System.Text.RegularExpressions.RegexOptions.Compiled);
    internal static string? NormalizeMarkColorInput(string? color)
    {
        if (string.IsNullOrEmpty(color)) return color;
        if (color[0] == '#') return color;
        if (_bareHex6Rx.IsMatch(color))
            return "#" + color.ToUpperInvariant();
        if (_bareHex8Rx.IsMatch(color))
            return "#" + color.ToUpperInvariant();
        if (_bareHex3Rx.IsMatch(color))
        {
            var c = color.ToUpperInvariant();
            return $"#{c[0]}{c[0]}{c[1]}{c[1]}{c[2]}{c[2]}";
        }
        return color;
    }

    internal static bool IsValidMarkColor(string color)
    {
        if (string.IsNullOrWhiteSpace(color)) return false;
        var c = color.Trim();
        if (c.Length > 64) return false; // defensive bound
        if (MarkNamedColors.Contains(c)) return true;
        if (_hexColorRx.IsMatch(c)) return true;
        if (_rgbFuncRx.IsMatch(c)) return true;
        return false;
    }

    /// <summary>
    /// Locate the element with the given data-path in the cached HTML snapshot
    /// and return its inner HTML fragment (start tag + children + end tag).
    /// Uses bracket-depth counting of sibling tags to find the matching close.
    /// Returns null if the path is not present.
    /// </summary>
    private static string? FindDataPathInHtml(string html, string path)
    {
        if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(path)) return null;
        // Anchor the search on the data-path attribute. Path may contain [] so
        // we match it as a literal substring inside quotes.
        var marker = "data-path=\"" + path + "\"";
        var idx = html.IndexOf(marker, StringComparison.Ordinal);
        if (idx < 0) return null;
        // Walk back to the opening '<' of this element's start tag.
        var start = html.LastIndexOf('<', idx);
        if (start < 0) return null;
        // Find the end of the start tag.
        var startEnd = html.IndexOf('>', idx);
        if (startEnd < 0) return null;
        // Self-closing tag? (extremely unlikely for data-path targets but be safe)
        if (html[startEnd - 1] == '/')
            return html.Substring(start, startEnd - start + 1);
        // Extract the tag name so we can match its close.
        var tagEnd = start + 1;
        while (tagEnd < html.Length && !char.IsWhiteSpace(html[tagEnd]) && html[tagEnd] != '>')
            tagEnd++;
        var tag = html.Substring(start + 1, tagEnd - start - 1).ToLowerInvariant();
        var openToken = "<" + tag;
        var closeToken = "</" + tag;
        // Count nested open/close to find the matching close tag.
        var depth = 1;
        var cursor = startEnd + 1;
        while (cursor < html.Length && depth > 0)
        {
            var nextOpen = html.IndexOf(openToken, cursor, StringComparison.OrdinalIgnoreCase);
            var nextClose = html.IndexOf(closeToken, cursor, StringComparison.OrdinalIgnoreCase);
            if (nextClose < 0) return null;
            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                // Ensure the candidate open isn't actually part of a longer tag name
                var after = nextOpen + openToken.Length;
                if (after < html.Length && (html[after] == ' ' || html[after] == '>' || html[after] == '\t' || html[after] == '\n'))
                {
                    depth++;
                    cursor = after;
                    continue;
                }
                cursor = nextOpen + openToken.Length;
                continue;
            }
            depth--;
            cursor = nextClose + closeToken.Length;
            if (depth == 0)
            {
                // Advance past the close tag's '>'
                var gt = html.IndexOf('>', cursor);
                if (gt < 0) return null;
                return html.Substring(start, gt - start + 1);
            }
        }
        return null;
    }

    /// <summary>
    /// Extract plain text content from an HTML fragment: strip all tags, decode
    /// HTML entities, collapse whitespace minimally, and NFC-normalize. Pure
    /// regex — no DOM parser dependency.
    /// </summary>
    internal static string ExtractTextContent(string htmlFragment)
    {
        if (string.IsNullOrEmpty(htmlFragment)) return "";
        // BUG-TESTER-003: drop <script>...</script> and <style>...</style> bodies
        // BEFORE per-tag stripping. _tagStripRx only removes tags, so without
        // this step inner JS/CSS text leaks into find matching.
        var noScript = _scriptBodyRx.Replace(htmlFragment, "");
        var noStyle = _styleBodyRx.Replace(noScript, "");
        var stripped = _tagStripRx.Replace(noStyle, "");
        var decoded = System.Net.WebUtility.HtmlDecode(stripped);
        try { return decoded.Normalize(System.Text.NormalizationForm.FormC); }
        catch { return decoded; }
    }

    /// <summary>
    /// Resolve a mark against the current HTML snapshot: populate
    /// MatchedText and Stale based on whether the path still resolves
    /// and whether find still matches.
    ///
    /// Pure function: returns a new WatchMark, does not mutate the input.
    /// The caller is responsible for locking _marksLock if it's writing back
    /// into _currentMarks.
    /// </summary>
    internal static WatchMark ResolveMark(WatchMark mark, string currentHtml)
    {
        var resolved = new WatchMark
        {
            Id = mark.Id,
            Path = mark.Path,
            Find = mark.Find,
            Color = mark.Color,
            Note = mark.Note,
            Tofix = mark.Tofix,
            CreatedAt = mark.CreatedAt,
            // Defaults get overwritten below.
            MatchedText = Array.Empty<string>(),
            Stale = false,
        };

        if (string.IsNullOrEmpty(currentHtml))
        {
            // No snapshot yet (watch just started, first refresh not arrived) —
            // treat as "not resolvable yet" but don't flag stale: the CLI may
            // be adding marks before the first render. Stale stays false.
            return resolved;
        }

        var fragment = FindDataPathInHtml(currentHtml, mark.Path);
        if (fragment == null)
        {
            resolved.Stale = true;
            return resolved;
        }

        if (string.IsNullOrEmpty(mark.Find))
        {
            // Whole-element mark — no text matching needed.
            return resolved;
        }

        var text = ExtractTextContent(fragment);
        var find = mark.Find;

        // CONSISTENCY(find-regex): r"..." / r'...' raw-string prefix detection
        // matches WordHandler.Set.cs:60-61 and CommandBuilder.Mark.cs. Keep in
        // sync. grep "CONSISTENCY(find-regex)" for every project-wide site.
        bool isRegex = find.Length >= 3
            && find[0] == 'r'
            && (find[1] == '"' || find[1] == '\'')
            && find[^1] == find[1];

        if (isRegex)
        {
            var pattern = find.Substring(2, find.Length - 3);
            try
            {
                // BUG-TESTER-001: bound the match with MarkRegexMatchTimeout so a
                // catastrophic backtracker cannot freeze the reconcile loop.
                var matches = System.Text.RegularExpressions.Regex.Matches(
                    text, pattern,
                    System.Text.RegularExpressions.RegexOptions.None,
                    MarkRegexMatchTimeout);
                if (matches.Count == 0)
                {
                    resolved.Stale = true;
                    return resolved;
                }
                var list = new string[matches.Count];
                for (int i = 0; i < matches.Count; i++) list[i] = matches[i].Value;
                resolved.MatchedText = list;
                return resolved;
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException)
            {
                // Pattern took too long against this input → treat as stale with
                // empty matches. Future reconciles will retry against fresh HTML.
                resolved.Stale = true;
                resolved.MatchedText = Array.Empty<string>();
                return resolved;
            }
            catch
            {
                // Bad regex → treat as no match, stale.
                resolved.Stale = true;
                return resolved;
            }
        }
        else
        {
            var needle = find;
            try { needle = needle.Normalize(System.Text.NormalizationForm.FormC); } catch { }
            if (text.IndexOf(needle, StringComparison.Ordinal) < 0)
            {
                resolved.Stale = true;
                return resolved;
            }
            resolved.MatchedText = new[] { needle };
            return resolved;
        }
    }

    /// <summary>
    /// Re-run ResolveMark on every mark in the current list. Called when the
    /// cached HTML snapshot changes (document reload / full refresh). Updates
    /// each mark's MatchedText and Stale in place and bumps _marksVersion so
    /// clients that missed the change can detect it.
    /// </summary>
    private void ReconcileAllMarks()
    {
        WatchMark[] snapshot;
        lock (_marksLock)
        {
            if (_currentMarks.Count == 0) return;
            for (int i = 0; i < _currentMarks.Count; i++)
            {
                _currentMarks[i] = ResolveMark(_currentMarks[i], _currentHtml);
            }
            _marksVersion++;
            snapshot = _currentMarks.ToArray();
        }
        BroadcastMarkUpdate(snapshot);
    }

    /// <summary>Replace a single slide fragment in the full HTML by data-slide number.</summary>
    private static string PatchSlideInHtml(string html, int slideNum, string newFragment)
    {
        var (start, end) = FindSlideFragmentRange(html, slideNum);
        if (start < 0) return html;
        return string.Concat(html.AsSpan(0, start), newFragment, html.AsSpan(end));
    }

    /// <summary>Append a slide fragment before the last closing tag of the main container.</summary>
    private static string AppendSlideToHtml(string html, string fragment)
    {
        // Find the last </div> before </body> — that's the .main container's closing tag
        var bodyClose = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        if (bodyClose < 0) return html + fragment;
        // Find the </div> just before </body>
        var mainClose = html.LastIndexOf("</div>", bodyClose, StringComparison.OrdinalIgnoreCase);
        if (mainClose < 0) return html;
        return string.Concat(html.AsSpan(0, mainClose), fragment, "\n", html.AsSpan(mainClose));
    }

    /// <summary>Remove a slide fragment from the full HTML.</summary>
    private static string RemoveSlideFromHtml(string html, int slideNum)
    {
        var (start, end) = FindSlideFragmentRange(html, slideNum);
        if (start < 0) return html;
        return string.Concat(html.AsSpan(0, start), html.AsSpan(end));
    }

    /// <summary>Find the start/end character positions of a slide-container div in the HTML.</summary>
    private static (int Start, int End) FindSlideFragmentRange(string html, int slideNum)
    {
        var marker = $"data-slide=\"{slideNum}\"";
        var idx = html.IndexOf(marker, StringComparison.Ordinal);
        if (idx < 0) return (-1, -1);

        var start = html.LastIndexOf("<div ", idx, StringComparison.Ordinal);
        if (start < 0) return (-1, -1);

        // Find matching closing </div> by counting nesting
        var depth = 0;
        var pos = start;
        while (pos < html.Length)
        {
            var nextOpen = html.IndexOf("<div", pos, StringComparison.OrdinalIgnoreCase);
            var nextClose = html.IndexOf("</div>", pos, StringComparison.OrdinalIgnoreCase);

            if (nextClose < 0) break;

            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                depth++;
                pos = nextOpen + 4;
            }
            else
            {
                depth--;
                if (depth == 0)
                    return (start, nextClose + 6);
                pos = nextClose + 6;
            }
        }

        return (-1, -1);
    }

    /// <summary>Extract all &lt;style&gt; blocks from HTML head, concatenated.</summary>
    private static string? ExtractStyleBlock(string html)
    {
        var sb = new StringBuilder();
        var idx = 0;
        while (true)
        {
            var start = html.IndexOf("<style>", idx, StringComparison.OrdinalIgnoreCase);
            if (start < 0) start = html.IndexOf("<style ", idx, StringComparison.OrdinalIgnoreCase);
            if (start < 0) break;
            var end = html.IndexOf("</style>", start, StringComparison.OrdinalIgnoreCase);
            if (end < 0) break;
            end += 8; // include </style>
            sb.Append(html, start, end - start);
            idx = end;
        }
        return sb.Length > 0 ? sb.ToString() : null;
    }

    /// <summary>Split Word HTML into blocks keyed by block number. Returns dict of blockNum → content.</summary>
    private static Dictionary<int, string> SplitWordBlocks(string html)
    {
        var blocks = new Dictionary<int, string>();
        var beginRx = new System.Text.RegularExpressions.Regex(@"<span class=""wb"" data-block=""(\d+)"" style=""display:none""></span>");
        var matches = beginRx.Matches(html);
        for (int i = 0; i < matches.Count; i++)
        {
            var m = matches[i];
            var blockNum = int.Parse(m.Groups[1].Value);
            var contentStart = m.Index + m.Length;
            var endMarker = $"<span class=\"we\" data-block=\"{blockNum}\" style=\"display:none\"></span>";
            var endIdx = html.IndexOf(endMarker, contentStart, StringComparison.Ordinal);
            if (endIdx >= 0)
                blocks[blockNum] = html[contentStart..endIdx];
        }
        return blocks;
    }

    /// <summary>Compute block-level patches between old and new Word HTML. Returns null if diff is too large (fallback to full).</summary>
    internal static List<WordPatch>? ComputeWordPatches(string oldHtml, string newHtml)
    {
        // Only diff if both are Word documents with block markers
        if (string.IsNullOrEmpty(oldHtml) || string.IsNullOrEmpty(newHtml))
            return null;
        if (!oldHtml.Contains("data-block=\"1\"") || !newHtml.Contains("data-block=\"1\""))
            return null;

        var oldBlocks = SplitWordBlocks(oldHtml);
        var newBlocks = SplitWordBlocks(newHtml);

        if (oldBlocks.Count == 0 && newBlocks.Count == 0) return null;

        var patches = new List<WordPatch>();

        // Find max block number across both
        var maxBlock = 0;
        foreach (var k in oldBlocks.Keys) if (k > maxBlock) maxBlock = k;
        foreach (var k in newBlocks.Keys) if (k > maxBlock) maxBlock = k;

        for (int b = 1; b <= maxBlock; b++)
        {
            var inOld = oldBlocks.TryGetValue(b, out var oldContent);
            var inNew = newBlocks.TryGetValue(b, out var newContent);

            if (inOld && inNew)
            {
                if (oldContent != newContent)
                    patches.Add(new WordPatch { Op = "replace", Block = b, Html = newContent });
                // else: unchanged, skip
            }
            else if (!inOld && inNew)
            {
                patches.Add(new WordPatch { Op = "add", Block = b, Html = newContent });
            }
            else if (inOld && !inNew)
            {
                patches.Add(new WordPatch { Op = "remove", Block = b });
            }
        }

        if (patches.Count == 0) return null; // no changes

        // If more than 60% of blocks changed (and enough blocks to matter), fallback to full refresh
        var totalBlocks = Math.Max(oldBlocks.Count, newBlocks.Count);
        if (totalBlocks >= 5 && patches.Count > totalBlocks * 0.6)
            return null;

        return patches;
    }

    private void SendSseWordPatch(List<WordPatch> patches, int version, int baseVersion, string? scrollTo)
    {
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"word-patch\"");
        sb.Append(",\"version\":").Append(version);
        sb.Append(",\"baseVersion\":").Append(baseVersion);
        sb.Append(",\"patches\":[");
        for (int i = 0; i < patches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"op\":\"").Append(patches[i].Op).Append('"');
            sb.Append(",\"block\":").Append(patches[i].Block);
            if (patches[i].Html != null)
            {
                sb.Append(",\"html\":");
                AppendJsonString(sb, patches[i].Html!);
            }
            sb.Append('}');
        }
        sb.Append(']');
        if (scrollTo != null)
        {
            sb.Append(",\"scrollTo\":");
            AppendJsonString(sb, scrollTo);
        }
        sb.Append('}');
        BroadcastSse(sb.ToString());
    }

    // ==================== Excel Row-Level Diff ====================

    /// <summary>
    /// Signature of chart overlay positions — concatenation of all data-from-row/col
    /// values in document order. Different signature → chart was moved → need full refresh.
    /// </summary>
    private static string ChartOverlaySignature(string html)
    {
        var sb = new System.Text.StringBuilder();
        var rx = new System.Text.RegularExpressions.Regex(@"data-from-(?:row|col)=""(\d+)""");
        foreach (System.Text.RegularExpressions.Match m in rx.Matches(html))
            sb.Append(m.Value).Append(',');
        return sb.ToString();
    }

    /// <summary>Split Excel HTML into rows keyed by "sheetIdx-rowNum" from data-row attributes.</summary>
    private static Dictionary<string, string> SplitExcelRows(string html)
    {
        var rows = new Dictionary<string, string>();

        // Static mode: extract <tr data-row="sheetIdx-rowNum"> elements
        var rx = new System.Text.RegularExpressions.Regex(@"<tr\s[^>]*data-row=""([^""]+)""[^>]*>");
        var matches = rx.Matches(html);
        for (int i = 0; i < matches.Count; i++)
        {
            var m = matches[i];
            var key = m.Groups[1].Value;
            var contentStart = m.Index;
            var endTag = "</tr>";
            var endIdx = html.IndexOf(endTag, contentStart + m.Length, StringComparison.Ordinal);
            if (endIdx >= 0)
                rows[key] = html[contentStart..(endIdx + endTag.Length)];
        }

        // Virt mode: extract rows from <script type="application/json" id="virt-data-N">
        // Format: [{"r":R,"frozen":bool[,"h":H],"html":"<escaped inner html>"},...]
        var scriptRx = new System.Text.RegularExpressions.Regex(
            @"<script[^>]*id=""virt-data-(\d+)""[^>]*>([\s\S]*?)</script>");
        var rowRx = new System.Text.RegularExpressions.Regex(
            @"""r"":(\d+).*?""html"":""((?:[^""\\]|\\.)*)""");
        var heightRx = new System.Text.RegularExpressions.Regex(@"""h"":(\d+(?:\.\d+)?)");
        foreach (System.Text.RegularExpressions.Match scriptMatch in scriptRx.Matches(html))
        {
            var sheetIdx = scriptMatch.Groups[1].Value;
            var json = scriptMatch.Groups[2].Value;
            foreach (System.Text.RegularExpressions.Match rowMatch in rowRx.Matches(json))
            {
                var rowNum = rowMatch.Groups[1].Value;
                var key = $"{sheetIdx}-{rowNum}";
                if (rows.ContainsKey(key)) continue; // frozen row already captured from static <tr>
                var innerHtml = rowMatch.Groups[2].Value
                    .Replace("\\\"", "\"").Replace("\\\\", "\\")
                    .Replace("\\n", "\n").Replace("\\r", "\r").Replace("\\t", "\t");
                // Extract row height from metadata fields (the portion before "html":)
                var htmlFieldOffset = rowMatch.Value.IndexOf("\"html\":", StringComparison.Ordinal);
                var metaStr = htmlFieldOffset >= 0 ? rowMatch.Value.Substring(0, htmlFieldOffset) : "";
                var hm = heightRx.Match(metaStr);
                var heightStyle = hm.Success ? $" style=\"height:{hm.Groups[1].Value}pt\"" : "";
                rows[key] = $"<tr data-row=\"{key}\"{heightStyle}>{innerHtml}</tr>";
            }
        }

        return rows;
    }

    /// <summary>Compute row-level patches between old and new Excel HTML. Returns null if diff is too large (fallback to full).</summary>
    internal static List<(string Op, string Row, string? Html)>? ComputeExcelPatches(string oldHtml, string newHtml)
    {
        if (string.IsNullOrEmpty(oldHtml) || string.IsNullOrEmpty(newHtml))
            return null;
        // Two valid row-data signals:
        //  static: data-row="X..." where the value starts with an alphanumeric char (real keys
        //          are "N-M" or "word-N-M"; JS template literals have data-row="' + ... which
        //          starts with a single-quote, not alphanumeric).
        //  virt:   id="virt-data-N" on <script> data elements (numeric suffix, not "{n}" template
        //          used by the virt JS implementation script).
        static bool HasRowData(string h) =>
            System.Text.RegularExpressions.Regex.IsMatch(h, @"data-row=""[a-zA-Z0-9]") ||
            System.Text.RegularExpressions.Regex.IsMatch(h, @"id=""virt-data-\d+""");
        if (!HasRowData(oldHtml) || !HasRowData(newHtml))
            return null;

        // If chart overlay positions changed, fall back to full refresh.
        // excel-patch only patches <tr> rows; overlay divs are outside the table
        // and won't be updated by row-level patching.
        if (ChartOverlaySignature(oldHtml) != ChartOverlaySignature(newHtml))
            return null;

        var oldRows = SplitExcelRows(oldHtml);
        var newRows = SplitExcelRows(newHtml);

        if (oldRows.Count == 0 && newRows.Count == 0) return null;

        var patches = new List<(string Op, string Row, string? Html)>();

        // Check all keys from both old and new
        var allKeys = new HashSet<string>(oldRows.Keys);
        allKeys.UnionWith(newRows.Keys);

        foreach (var key in allKeys)
        {
            var inOld = oldRows.TryGetValue(key, out var oldContent);
            var inNew = newRows.TryGetValue(key, out var newContent);

            if (inOld && inNew)
            {
                if (oldContent != newContent)
                    patches.Add(("replace", key, newContent));
            }
            else if (!inOld && inNew)
            {
                patches.Add(("add", key, newContent));
            }
            else if (inOld && !inNew)
            {
                patches.Add(("remove", key, null));
            }
        }

        if (patches.Count == 0) return null;

        // If more than 60% of rows changed, fallback to full refresh
        var totalRows = Math.Max(oldRows.Count, newRows.Count);
        if (totalRows >= 5 && patches.Count > totalRows * 0.6)
            return null;

        return patches;
    }

    private void SendSseExcelPatch(List<(string Op, string Row, string? Html)> patches, int version, int baseVersion, string? scrollTo)
    {
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"excel-patch\"");
        sb.Append(",\"version\":").Append(version);
        sb.Append(",\"baseVersion\":").Append(baseVersion);
        sb.Append(",\"patches\":[");
        for (int i = 0; i < patches.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append("{\"op\":\"").Append(patches[i].Op).Append('"');
            sb.Append(",\"row\":\"").Append(patches[i].Row).Append('"');
            if (patches[i].Html != null)
            {
                sb.Append(",\"html\":");
                AppendJsonString(sb, patches[i].Html!);
            }
            sb.Append('}');
        }
        sb.Append(']');
        if (scrollTo != null)
        {
            sb.Append(",\"scrollTo\":");
            AppendJsonString(sb, scrollTo);
        }
        sb.Append('}');
        BroadcastSse(sb.ToString());
    }

    private void SendSseEvent(string action, int slideNum, string? html, string? scrollTo = null, int version = 0)
    {
        // Build JSON manually to avoid dependency
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"").Append(action).Append('"');
        sb.Append(",\"slide\":").Append(slideNum);
        sb.Append(",\"version\":").Append(version);
        if (html != null)
        {
            sb.Append(",\"html\":");
            AppendJsonString(sb, html);
        }
        if (scrollTo != null)
        {
            sb.Append(",\"scrollTo\":");
            AppendJsonString(sb, scrollTo);
        }
        sb.Append('}');

        BroadcastSse(sb.ToString());
    }

    private void BroadcastSse(string sseJson)
    {
        lock (_sseLock)
        {
            var dead = new List<NetworkStream>();
            foreach (var client in _sseClients)
            {
                try
                {
                    var data = Encoding.UTF8.GetBytes($"event: update\ndata: {sseJson}\n\n");
                    client.Write(data);
                    client.Flush();
                }
                catch
                {
                    dead.Add(client);
                }
            }
            foreach (var d in dead) _sseClients.Remove(d);
        }
    }

    private static void AppendJsonString(StringBuilder sb, string value)
    {
        sb.Append('"');
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (ch < 0x20)
                        sb.Append($"\\u{(int)ch:X4}");
                    else
                        sb.Append(ch);
                    break;
            }
        }
        sb.Append('"');
    }

    private async Task HandleClientAsync(TcpClient client, CancellationToken token)
    {
        try
        {
            var stream = client.GetStream();
            var (requestLine, headers, bodyPrefix) = await ReadHttpRequestHeaderAsync(stream, token);

            if (requestLine.Contains("GET /events"))
            {
                try
                {
                    await HandleSseAsync(stream, token);
                }
                finally
                {
                    client.Close();
                }
                return;
            }

            if (requestLine.StartsWith("POST /api/selection", StringComparison.Ordinal))
            {
                await HandlePostSelectionAsync(stream, headers, bodyPrefix, token);
                client.Close();
                return;
            }

            if (requestLine.StartsWith("POST /api/edit", StringComparison.Ordinal))
            {
                await HandlePostEditAsync(stream, headers, bodyPrefix, token);
                client.Close();
                return;
            }

            // BUG-TESTER-R503: GET/PUT/etc on /api/selection must return 405,
            // not fall through to the HTML preview. Without this, an API
            // client that uses the wrong verb gets back a 200 HTML page and
            // never realizes the request was malformed.
            if (requestLine.Contains(" /api/selection"))
            {
                var msg = Encoding.UTF8.GetBytes("Method Not Allowed: /api/selection only accepts POST");
                var hdr = Encoding.UTF8.GetBytes(
                    $"HTTP/1.1 405 Method Not Allowed\r\nAllow: POST\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {msg.Length}\r\nConnection: close\r\n\r\n");
                await stream.WriteAsync(hdr, token);
                await stream.WriteAsync(msg, token);
                client.Close();
                return;
            }

            // BUG-TESTER-R504: any other /api/... path is unknown and must
            // return 404. Without this, an agent that mistypes /api/marks
            // (we don't have a marks HTTP endpoint, only the pipe verb) gets
            // the HTML preview page back and silently misroutes.
            if (requestLine.Contains(" /api/"))
            {
                var msg = Encoding.UTF8.GetBytes("Not Found");
                var hdr = Encoding.UTF8.GetBytes(
                    $"HTTP/1.1 404 Not Found\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {msg.Length}\r\nConnection: close\r\n\r\n");
                await stream.WriteAsync(hdr, token);
                await stream.WriteAsync(msg, token);
                client.Close();
                return;
            }

            // Default: serve current HTML (GET / and everything else)
            var html = string.IsNullOrEmpty(_currentHtml)
                ? InjectSseScript(WaitingHtml)
                : InjectSseScript(_currentHtml);
            var bodyBytes = Encoding.UTF8.GetBytes(html);
            var header = Encoding.UTF8.GetBytes(
                $"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {bodyBytes.Length}\r\nConnection: close\r\n\r\n");
            await stream.WriteAsync(header, token);
            await stream.WriteAsync(bodyBytes, token);
            client.Close();
        }
        catch
        {
            try { client.Close(); } catch { }
        }
    }

    /// <summary>
    /// Read the HTTP request line and headers, plus any body bytes that arrived in the
    /// same TCP read. Returns (requestLine, headers, bodyPrefix). Caller is responsible
    /// for reading the rest of the body using Content-Length if needed.
    /// </summary>
    private static async Task<(string requestLine, Dictionary<string, string> headers, string bodyPrefix)>
        ReadHttpRequestHeaderAsync(NetworkStream stream, CancellationToken token)
    {
        var buffer = new byte[8192];
        var sb = new StringBuilder();
        int headerEnd = -1;
        while (headerEnd < 0)
        {
            var n = await stream.ReadAsync(buffer.AsMemory(), token);
            if (n == 0) break;
            sb.Append(Encoding.UTF8.GetString(buffer, 0, n));
            headerEnd = sb.ToString().IndexOf("\r\n\r\n", StringComparison.Ordinal);
            if (sb.Length > 32 * 1024) break; // safety cap
        }

        var raw = sb.ToString();
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (headerEnd < 0)
        {
            // No header terminator — treat the whole thing as a single line
            var firstLine = raw;
            var crlf = raw.IndexOf("\r\n", StringComparison.Ordinal);
            if (crlf >= 0) firstLine = raw[..crlf];
            return (firstLine, headers, "");
        }

        var headerSection = raw[..headerEnd];
        var bodyPrefix = raw[(headerEnd + 4)..];
        var lines = headerSection.Split("\r\n");
        var requestLine = lines.Length > 0 ? lines[0] : "";
        for (int i = 1; i < lines.Length; i++)
        {
            var colon = lines[i].IndexOf(':');
            if (colon > 0)
                headers[lines[i][..colon].Trim()] = lines[i][(colon + 1)..].Trim();
        }
        return (requestLine, headers, bodyPrefix);
    }

    // Maximum size of a POST /api/selection request body. 64 KB is plenty for tens
    // of thousands of selected paths and bounds memory + read time per request.
    private const int MaxSelectionBodyBytes = 64 * 1024;
    // Hard limit on how long we'll wait for the rest of a POST body to arrive.
    // Prevents slow-loris style stalls (Content-Length advertised, body never sent).
    private static readonly TimeSpan PostBodyReadTimeout = TimeSpan.FromSeconds(3);

    private async Task HandlePostSelectionAsync(NetworkStream stream, Dictionary<string, string> headers, string bodyPrefix, CancellationToken token)
    {
        int statusCode = 204;
        string statusText = "No Content";
        string body = bodyPrefix;

        try
        {
            // Reject runaway Content-Length up front (covers FUZZER-001 slow-loris).
            int contentLength = -1;
            if (headers.TryGetValue("Content-Length", out var clStr) && int.TryParse(clStr, out var parsedCl))
            {
                if (parsedCl < 0 || parsedCl > MaxSelectionBodyBytes)
                    throw new InvalidDataException("body too large");
                contentLength = parsedCl;
            }

            // If the bodyPrefix already exceeds Content-Length, trim it. Without this,
            // an attacker could smuggle extra bytes by sending a long body in the same
            // TCP segment as the headers (FUZZER-002).
            var prefixBytes = Encoding.UTF8.GetByteCount(body);
            if (contentLength >= 0 && prefixBytes > contentLength)
            {
                var prefBytes = Encoding.UTF8.GetBytes(body);
                body = Encoding.UTF8.GetString(prefBytes, 0, contentLength);
                prefixBytes = contentLength;
            }

            // Read any missing tail bytes, bounded by both size and time.
            if (contentLength > prefixBytes)
            {
                using var readCts = CancellationTokenSource.CreateLinkedTokenSource(token);
                readCts.CancelAfter(PostBodyReadTimeout);
                var sb = new StringBuilder(body, contentLength);
                int have = prefixBytes;
                var buf = new byte[8192];
                try
                {
                    while (have < contentLength)
                    {
                        var toRead = Math.Min(buf.Length, contentLength - have);
                        var n = await stream.ReadAsync(buf.AsMemory(0, toRead), readCts.Token);
                        if (n == 0) break;
                        sb.Append(Encoding.UTF8.GetString(buf, 0, n));
                        have += n;
                        if (have > MaxSelectionBodyBytes)
                            throw new InvalidDataException("body too large");
                    }
                }
                catch (OperationCanceledException) when (!token.IsCancellationRequested)
                {
                    throw new InvalidDataException("body read timed out");
                }
                body = sb.ToString();
            }

            // Expected JSON: {"paths": ["/slide[1]/shape[2]", ...]}
            var req = JsonSerializer.Deserialize(body, WatchSelectionJsonContext.Default.SelectionRequest);
            var rawSelection = req?.Paths ?? new List<string>();
            // BUG-TESTER-R501/R502 + BUG-FUZZER-R5-04: bring selection path
            // hardening up to parity with mark (Round 2/3 fixes). Each path is
            // Trim()-normalized; whitespace-only and paths not starting with
            // '/' are dropped; paths containing control characters (CR/LF/NUL
            // /etc) are dropped because they would corrupt the in-memory
            // representation and the SSE/pipe readback even though
            // AppendJsonString escapes them on the wire.
            // CONSISTENCY(path-stability): mirror of HandleMarkAdd's input
            // validation. If you change the path acceptance rules, change
            // both at once. grep CONSISTENCY(path-stability).
            var newSelection = new List<string>(rawSelection.Count);
            foreach (var raw in rawSelection)
            {
                if (string.IsNullOrEmpty(raw)) continue;
                var trimmed = raw.Trim();
                if (string.IsNullOrWhiteSpace(trimmed)) continue;
                if (!trimmed.StartsWith("/")) continue;
                var hasControl = false;
                for (int i = 0; i < trimmed.Length; i++)
                {
                    if (char.IsControl(trimmed[i])) { hasControl = true; break; }
                }
                if (hasControl) continue;
                newSelection.Add(trimmed);
            }

            lock (_selectionLock) { _currentSelection = newSelection; }
            _lastActivityTime = DateTime.UtcNow;

            // Broadcast to all SSE clients so other browsers can highlight in sync
            BroadcastSelectionUpdate(newSelection);
        }
        catch
        {
            statusCode = 400;
            statusText = "Bad Request";
        }

        var resp = Encoding.UTF8.GetBytes(
            $"HTTP/1.1 {statusCode} {statusText}\r\nContent-Length: 0\r\nConnection: close\r\nAccess-Control-Allow-Origin: *\r\n\r\n");
        await stream.WriteAsync(resp, token);
    }

    /// <summary>
    /// Handle POST /api/edit — spawn officecli set as a child process to modify the file.
    /// The set command will notify the watch server via named pipe, triggering an SSE refresh.
    /// WatchServer never opens the file directly (see CLAUDE.md "Watch Server Rules").
    /// </summary>
    private async Task HandlePostEditAsync(NetworkStream stream, Dictionary<string, string> headers, string bodyPrefix, CancellationToken token)
    {
        int statusCode = 204;
        string statusText = "No Content";
        try
        {
            // Read body (same pattern as selection handler)
            int contentLength = 0;
            if (headers.TryGetValue("Content-Length", out var clStr) && int.TryParse(clStr, out var cl))
                contentLength = cl;
            if (contentLength > MaxSelectionBodyBytes) throw new InvalidDataException("body too large");

            var body = bodyPrefix;
            if (contentLength > body.Length)
            {
                var sb = new StringBuilder(body);
                var buf = new byte[4096];
                int have = Encoding.UTF8.GetByteCount(body);
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(token);
                cts.CancelAfter(PostBodyReadTimeout);
                while (have < contentLength)
                {
                    var n = await stream.ReadAsync(buf, cts.Token);
                    if (n == 0) break;
                    sb.Append(Encoding.UTF8.GetString(buf, 0, n));
                    have += n;
                }
                body = sb.ToString();
            }

            // Parse: {"path": "...", "prop": "text", "value": "Hello"}
            // or:    {"path": "...", "props": {"x": "10pt", "y": "20pt"}}
            using var doc = System.Text.Json.JsonDocument.Parse(body);
            var root = doc.RootElement;
            var path = root.GetProperty("path").GetString() ?? "";

            // Spawn officecli set as child process
            var exe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName
                ?? (OperatingSystem.IsWindows() ? "officecli.exe" : "officecli");
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = exe,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            psi.ArgumentList.Add("set");
            psi.ArgumentList.Add(_filePath);
            psi.ArgumentList.Add(path);
            if (root.TryGetProperty("props", out var propsEl) && propsEl.ValueKind == System.Text.Json.JsonValueKind.Object)
            {
                foreach (var kv in propsEl.EnumerateObject())
                {
                    psi.ArgumentList.Add("--prop");
                    psi.ArgumentList.Add($"{kv.Name}={kv.Value.GetString() ?? ""}");
                }
            }
            else
            {
                var prop = root.GetProperty("prop").GetString() ?? "text";
                var value = root.GetProperty("value").GetString() ?? "";
                psi.ArgumentList.Add("--prop");
                psi.ArgumentList.Add($"{prop}={value}");
            }
            using var proc = System.Diagnostics.Process.Start(psi);
            if (proc != null)
            {
                await proc.WaitForExitAsync(token);
                // set command auto-notifies watch via named pipe → SSE refresh
            }
        }
        catch
        {
            statusCode = 400; statusText = "Bad Request";
        }
        var resp = Encoding.UTF8.GetBytes(
            $"HTTP/1.1 {statusCode} {statusText}\r\nContent-Length: 0\r\nConnection: close\r\nAccess-Control-Allow-Origin: *\r\n\r\n");
        await stream.WriteAsync(resp, token);
    }

    private void BroadcastSelectionUpdate(List<string> paths)
    {
        var sb = new StringBuilder();
        sb.Append("{\"action\":\"selection-update\",\"paths\":[");
        for (int i = 0; i < paths.Count; i++)
        {
            if (i > 0) sb.Append(',');
            AppendJsonString(sb, paths[i]);
        }
        sb.Append("]}");
        BroadcastSse(sb.ToString());
    }

    /// <summary>
    /// Wrap a WatchMark[] snapshot in a "mark-update" SSE envelope. Called
    /// after every mark add/remove, and during initial SSE client handshake.
    /// The version field is a monotonically-increasing counter that clients
    /// can use for CAS-style update detection.
    ///
    /// Uses the Relaxed encoder so CJK find/note/tofix bytes flow through
    /// as literal characters instead of \uXXXX escapes.
    /// </summary>
    private static string BuildMarkUpdateJson(WatchMark[] marks, int version)
    {
        var marksJson = JsonSerializer.Serialize(marks, WatchMarkJsonOptions.WatchMarkArrayInfo);
        return $"{{\"action\":\"mark-update\",\"version\":{version},\"marks\":{marksJson}}}";
    }

    private void BroadcastMarkUpdate(WatchMark[] marks)
    {
        int version;
        lock (_marksLock) { version = _marksVersion; }
        BroadcastSse(BuildMarkUpdateJson(marks, version));
    }

    private async Task HandleSseAsync(NetworkStream stream, CancellationToken token)
    {
        var header = Encoding.UTF8.GetBytes(
            "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream; charset=utf-8\r\nCache-Control: no-cache\r\nConnection: keep-alive\r\nAccess-Control-Allow-Origin: *\r\n\r\n");
        await stream.WriteAsync(header, token);

        _lastActivityTime = DateTime.UtcNow;

        // Send the current selection immediately so the new client can highlight
        // any elements that are already selected by other browsers viewing the same
        // file. CRITICAL: this write must happen BEFORE adding the stream to
        // _sseClients. Otherwise BroadcastSse (running on another thread under
        // _sseLock) could write to the same stream at the same time we are writing
        // the initial event here, and NetworkStream is not safe for concurrent writes
        // — interleaved bytes would corrupt SSE framing.
        try
        {
            string[] snapshot;
            lock (_selectionLock) { snapshot = _currentSelection.ToArray(); }
            var sb = new StringBuilder();
            sb.Append("{\"action\":\"selection-update\",\"paths\":[");
            for (int i = 0; i < snapshot.Length; i++)
            {
                if (i > 0) sb.Append(',');
                AppendJsonString(sb, snapshot[i]);
            }
            sb.Append("]}");
            var initEvt = Encoding.UTF8.GetBytes($"event: update\ndata: {sb}\n\n");
            await stream.WriteAsync(initEvt, token);

            // Also dump the current marks snapshot so a freshly connected browser
            // immediately sees any marks the CLI has already added. Mirrors the
            // selection init dump pattern above.
            WatchMark[] markSnapshot;
            int markVersion;
            lock (_marksLock)
            {
                markSnapshot = _currentMarks.ToArray();
                markVersion = _marksVersion;
            }
            var markJson = BuildMarkUpdateJson(markSnapshot, markVersion);
            var markInitEvt = Encoding.UTF8.GetBytes($"event: update\ndata: {markJson}\n\n");
            await stream.WriteAsync(markInitEvt, token);
        }
        catch { }

        // Now safe to register: any subsequent BroadcastSse will serialize against
        // future writes via _sseLock.
        lock (_sseLock) { _sseClients.Add(stream); }

        try
        {
            while (!token.IsCancellationRequested)
            {
                await Task.Delay(30000, token);
                var heartbeat = Encoding.UTF8.GetBytes(": heartbeat\n\n");
                await stream.WriteAsync(heartbeat, token);
            }
        }
        catch { }
        finally
        {
            lock (_sseLock) { _sseClients.Remove(stream); }
        }
    }

    private static string InjectSseScript(string html)
    {
        var script = _sseScriptBlock.Value;
        var idx = html.LastIndexOf("</body>", StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
            return html[..idx] + script + html[idx..];
        return html + script;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Delegate to shared shutdown. If RunAsync or a signal handler
        // already drove shutdown, this just awaits the cached Task.
        // Steps include TcpListener.Stop(), pipe kick, SSE cleanup, and
        // CoreFxPipe_ socket delete (BUG-BT-003).
        try { StopAsync().Wait(TimeSpan.FromSeconds(10)); }
        catch (Exception ex) { Console.Error.WriteLine($"Warning: watch shutdown error: {ex.Message}"); }

        try { _cts.Dispose(); } catch { }
    }
}
