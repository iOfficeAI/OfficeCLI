// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0
//
// CONSISTENCY(watch-isolation): this file does not reference OfficeCli.Handlers, does not open files,
// does not write to disk. See the project conventions "Watch Server Rules". To relax this red line,
// grep "CONSISTENCY(watch-isolation)" and review every file in the watch subsystem project-wide.

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
    // _filePath/_pipeName are mutable because POST /api/switch retargets the
    // server to a different document in place (port and SSE connections
    // survive). All writes happen under _switchLock; reads are lock-free
    // (reference assignment is atomic; a stale read is as benign as a request
    // racing the switch).
    private string _filePath;
    private string _pipeName;
    private int _port; // mutable: --port 0 binds an ephemeral port, resolved after Start()
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

    // Serializes POST /api/switch so two concurrent switches can't interleave
    // their pipe/marker/state swaps.
    private readonly SemaphoreSlim _switchLock = new(1, 1);

    // Current selection — paths of elements selected in any connected browser.
    // Single shared list (last-write-wins): all browsers viewing the same file see
    // the same selection. CLI reads this via the named pipe "get-selection" command.
    //
    // CONSISTENCY(path-stability): selection and mark share the same naive positional addressing
    // contract — no fingerprinting, no drift detection. To upgrade to stable IDs,
    // grep "CONSISTENCY(path-stability)" and update every deferred site project-wide in one pass.
    // See the project conventions "Design Principles".
    private List<string> _currentSelection = new();
    private readonly object _selectionLock = new();

    // Current marks — advisory annotations attached to document paths. Live in
    // memory only. Server never opens the document and never inspects DOM —
    // marks are pure metadata; the browser computes match positions client-side.
    //
    // CONSISTENCY(path-stability): element-deletion / position-drift handling deliberately matches
    // selection — naive positional addressing, no fingerprint, no drift detection. `stale` is only
    // set when the client reports a path-resolution failure or a `find` miss.
    // See the project conventions "Design Principles" + "Watch Server Rules".
    // To migrate to stable-ID paths, grep "CONSISTENCY(path-stability)" and update every deferred
    // site (selection / mark / any future path consumer) project-wide — never patch mark alone.
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
        // CONSISTENCY(path-identity): symlink-resolved so `watch test.pptx` and
        // `set /tmp/dir/test.pptx …` name the same pipe (see PathIdentity).
        var fullPath = PathIdentity.Canonical(filePath);
        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
            fullPath = fullPath.ToUpperInvariant();
        var hash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(fullPath)))[..16];
        return $"officecli-watch-{hash}";
    }

    /// <summary>
    /// Path of the on-disk marker that records {pid, port} for a running
    /// watch. Used by <see cref="GetExistingWatchPort"/> and
    /// <see cref="IsWatching"/> to answer "is anyone watching this file?"
    /// without a pipe round-trip. Same hash key as the pipe name — one
    /// file ↔ one pipe ↔ one marker.
    /// </summary>
    public static string GetWatchMarkerPath(string filePath)
    {
        return Path.Combine(Path.GetTempPath(), GetWatchPipeName(filePath) + ".port");
    }

    /// <summary>
    /// Check if another watch process is already running for this file.
    /// Returns the port number if running, or null if not.
    ///
    /// Implementation: reads the on-disk marker file
    /// ({pid}\n{port}\n{startTicksUtc}\n) and validates the pid is still
    /// alive AND is the same process incarnation that wrote the marker
    /// (start-time comparison — a bare pid check accepts an unrelated
    /// process that got the crashed writer's pid recycled to it, which
    /// made this report a dead watch as live forever). Legacy two-line
    /// markers ({pid}\n{port}\n) fall back to the bare pid check.
    /// Replaces the pre-1.0.51 pipe ping probe, which cost ~100ms and
    /// falsely reported "not watching" when the pipe server was
    /// momentarily busy with another connection.
    /// </summary>
    public static int? GetExistingWatchPort(string filePath)
    {
        var markerPath = GetWatchMarkerPath(filePath);
        try
        {
            var info = new FileInfo(markerPath);
            if (!info.Exists) return null;
            // The marker path is predictable ($TMPDIR/officecli-watch-<hash>.port),
            // so on a shared temp dir a local attacker can plant a symlink there.
            // Only a regular file we could have written is a trustworthy marker:
            // never read through (or delete) a symlink / reparse point (CWE-59).
            if ((info.Attributes & FileAttributes.ReparsePoint) != 0) return null;
            var lines = File.ReadAllLines(markerPath);
            if (lines.Length < 2) return null;
            if (!int.TryParse(lines[0], out var pid)) return null;
            if (!int.TryParse(lines[1], out var port)) return null;
            long? startTicksUtc = null;
            if (lines.Length >= 3 && long.TryParse(lines[2], out var ticks))
                startTicksUtc = ticks;
            if (!IsProcessAlive(pid, startTicksUtc))
            {
                // Stale marker — writer crashed or was killed without cleanup
                // (possibly with its pid since recycled to an unrelated
                // process). Best-effort remove so the caller can start a
                // fresh watch.
                try { File.Delete(markerPath); } catch { }
                return null;
            }
            return port;
        }
        catch
        {
            return null;
        }
    }

    public static bool IsWatching(string filePath)
    {
        return GetExistingWatchPort(filePath).HasValue;
    }

    private static bool IsProcessAlive(int pid, long? expectedStartTicksUtc)
    {
        try
        {
            using var p = System.Diagnostics.Process.GetProcessById(pid);
            if (p.HasExited) return false;
            if (expectedStartTicksUtc.HasValue)
            {
                try
                {
                    if (!StartTicksMatch(p.StartTime.ToUniversalTime().Ticks, expectedStartTicksUtc.Value))
                        return false; // pid recycled by an unrelated process
                }
                catch
                {
                    // StartTime unreadable — typically the pid was recycled to
                    // another user's process. Markers are 0600 files written by
                    // the same user, so an unverifiable identity means a
                    // recycled pid, not our watch: treat as dead.
                    return false;
                }
            }
            return true;
        }
        catch (ArgumentException) { return false; }
        catch (InvalidOperationException) { return false; }
    }

    // Process start-time identity check, tolerant of platform jitter. On Linux
    // Process.StartTime is derived from /proc/<pid>/stat jiffies + boot time,
    // and the same live process yields a slightly different tick count when
    // read by another process (~hundreds of microseconds) than the value it
    // recorded for itself — so an EXACT tick comparison wrongly reports a live
    // watch as dead, deleting its marker and breaking watch on Linux (regression
    // from the pid-only check). Windows/macOS read a precise kernel creation
    // timestamp that is stable across readers. A 2-second tolerance absorbs the
    // Linux jitter while still rejecting a recycled pid, whose unrelated process
    // started seconds-to-hours after the crashed writer.
    internal static readonly long StartTicksTolerance = TimeSpan.FromSeconds(2).Ticks;

    internal static bool StartTicksMatch(long actualTicks, long expectedTicks)
        => Math.Abs(actualTicks - expectedTicks) <= StartTicksTolerance;

    private void WriteMarker()
    {
        var markerPath = GetWatchMarkerPath(_filePath);
        try
        {
            // Refuse to follow a pre-planted symlink at the predictable marker
            // path: a local attacker who creates the marker as a symlink to a
            // victim-writable file would otherwise have us truncate that file
            // (CWE-59 symlink-follow). FileMode.CreateNew maps to O_CREAT|O_EXCL,
            // which by POSIX fails — without following — when the path already
            // exists, including when it is a symlink. A stale regular marker from
            // a dead writer was cleared by GetExistingWatchPort just above; if a
            // squatter still holds the name we simply skip the marker (IsWatching
            // then reports false — fail-safe, no clobber).
            // Third line: this process's start time (UTC ticks), so the
            // liveness check can tell "our writer" from "an unrelated process
            // that got the pid recycled to it after our writer crashed".
            using var self = System.Diagnostics.Process.GetCurrentProcess();
            var bytes = Encoding.UTF8.GetBytes(
                $"{self.Id}\n{_port}\n{self.StartTime.ToUniversalTime().Ticks}\n");
            var opts = new FileStreamOptions
            {
                Mode = FileMode.CreateNew,
                Access = FileAccess.Write,
                Share = FileShare.None,
            };
            if (!OperatingSystem.IsWindows())
                opts.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite; // 0600
            using var fs = new FileStream(markerPath, opts);
            fs.Write(bytes);
        }
        catch { /* best-effort; IsWatching just reports false if marker absent */ }
    }

    private void DeleteMarker()
    {
        try
        {
            var markerPath = GetWatchMarkerPath(_filePath);
            if (File.Exists(markerPath)) File.Delete(markerPath);
        }
        catch { /* best-effort cleanup */ }
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
        // --port 0 asks the OS for an ephemeral port; resolve the real one
        // before it reaches the marker file and the printed URL, otherwise
        // both would say 0 and per-file discovery (mark/goto/unwatch,
        // IsWatching) breaks.
        if (_port == 0)
            _port = ((IPEndPoint)_tcpListener.LocalEndpoint!).Port;
        WriteMarker();
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
        // SIGHUP: only treat as shutdown when we have a controlling TTY
        // (user closed the terminal hosting a foreground watch). For
        // non-interactive launchers (CI, agent schedulers using stdin=
        // /dev/null without setsid/nohup), the parent shell delivers a
        // spurious SIGHUP after eval; we must catch and IGNORE it,
        // because the kernel's default disposition for SIGHUP is
        // terminate — simply not registering would still kill us.
        bool sighupKills = !Console.IsInputRedirected;
        try
        {
            signalRegs.Add(PosixSignalRegistration.Create(PosixSignal.SIGHUP, ctx =>
            {
                ctx.Cancel = true;
                if (sighupKills) DoShutdownFromSignal();
                // else: swallow — headless watch survives stray SIGHUP.
            }));
        }
        catch (PlatformNotSupportedException) { /* host doesn't support */ }
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

        // 4b. Delete the on-disk watch marker so external IsWatching() probes
        //     immediately see "no watch running".
        DeleteMarker();

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
            // Re-read per iteration: /api/switch changes _pipeName and then
            // kicks the pending WaitForConnectionAsync with a dummy connect,
            // so the next iteration listens on the new document's pipe.
            var pipeName = _pipeName;
            var server = new System.IO.Pipes.NamedPipeServerStream(
                pipeName, System.IO.Pipes.PipeDirection.InOut,
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
                else if (message != null && message.StartsWith("scroll ", StringComparison.Ordinal))
                {
                    // "scroll <selector>" — validate the CSS selector against
                    // the cached HTML snapshot, broadcast on success, return
                    // "ok" or "err:<msg>". BUG-BT-R33-3: pure-positional
                    // existence check on the cached HTML so goto can fail
                    // exit=1 instead of silently exit=0 on missing anchors.
                    // CONSISTENCY(watch-isolation): no file open — only the
                    // already-cached HTML string is inspected.
                    var selector = message.Substring(7);
                    var found = SelectorExistsInHtml(_currentHtml, selector);
                    if (!found)
                    {
                        await writer.WriteLineAsync(("err:selector not found in current HTML: " + selector).AsMemory(), token);
                    }
                    else
                    {
                        await writer.WriteLineAsync("ok".AsMemory(), token);
                        SendSseEvent("scroll", 0, null, selector, _version);
                    }
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

            // Scroll-only event: broadcast a CSS selector to all SSE clients
            // without touching the cached HTML, version, or marks. Used by the
            // `goto` command to navigate already-running watch viewers.
            if (msg.Action == "scroll" && !string.IsNullOrEmpty(msg.ScrollTo))
            {
                SendSseEvent("scroll", 0, null, msg.ScrollTo, _version);
                return;
            }

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

            // Excel: try row-level diff instead of full refresh.
            // Skip when table chrome (colgroup/thead/table width) changed —
            // row patches can't express those changes, so fall through to
            // full-action so the browser rebuilds the whole body.
            if (msg.Action == "full" && !string.IsNullOrEmpty(msg.FullHtml)
                && !string.IsNullOrEmpty(oldHtml) && oldHtml.Contains("data-row=\"")
                && TableChromeSignature(oldHtml) == TableChromeSignature(msg.FullHtml))
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
    // to stable IDs. See the project conventions "Watch Server Rules".
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
    /// HTML-encode an attribute value mirroring how the renderer escapes
    /// data-path. Only the characters that change inside double-quoted
    /// attribute values matter (&, &lt;, &gt;, &quot;, &#39; / &apos;).
    /// </summary>
    private static string HtmlEncodeAttributeValue(string value)
    {
        // Order matters: replace '&' first so subsequent ampersand-introducing
        // entities aren't re-encoded.
        var sb = new StringBuilder(value.Length);
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&#39;"); break;
                default: sb.Append(ch); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Locate the element with the given data-path in the cached HTML snapshot
    /// and return its inner HTML fragment (start tag + children + end tag).
    /// Uses bracket-depth counting of sibling tags to find the matching close.
    /// Returns null if the path is not present.
    /// </summary>
    private static string? FindDataPathInHtml(string html, string path)
    {
        // CONSISTENCY(pptx-group-flatten): query may emit paths that point
        // inside a group (`/slide[1]/group[2]/shape[3]`), but HtmlPreview
        // currently only emits data-path on the outer group — see the
        // CONSISTENCY note in PowerPointHandler.HtmlPreview.Shapes.cs:~1040.
        // If the exact path isn't in the rendered HTML, walk up one segment
        // at a time and try again so a mark on a group-internal shape
        // resolves to the nearest ancestor that *is* rendered. Text-based
        // find/replace still runs against the ancestor's full text content,
        // so highlighting + find still work — only the visual outline drops
        // to the group level.
        var direct = FindDataPathInHtmlExact(html, path);
        if (direct != null) return direct;

        var current = path;
        while (true)
        {
            var lastSlash = current.LastIndexOf('/');
            if (lastSlash <= 0) return null;
            current = current.Substring(0, lastSlash);
            var hit = FindDataPathInHtmlExact(html, current);
            if (hit != null) return hit;
        }
    }

    private static string? FindDataPathInHtmlExact(string html, string path)
    {
        if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(path)) return null;
        // Anchor the search on the data-path attribute. Path may contain [] so
        // we match it as a literal substring inside quotes.
        // BUG-FIX(B9): the HTML emitter encodes attribute values, so a path
        // like /shape[@name="Foo"] is rendered as data-path="/shape[@name=&quot;Foo&quot;]".
        // Match against the encoded form so paths containing ", ', <, >, & don't
        // always come back stale.
        var encodedPath = HtmlEncodeAttributeValue(path);
        var marker = "data-path=\"" + encodedPath + "\"";
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
    /// Existence check for the small set of CSS selectors emitted by
    /// WatchNotifier.ExtractWordScrollTarget — `#anchor` (id=) or
    /// `[data-path="..."]`. Pure substring scan over the cached HTML;
    /// no DOM parser, mirrors FindDataPathInHtml's design.
    /// CONSISTENCY(watch-isolation): only the cached HTML is read.
    /// </summary>
    internal static bool SelectorExistsInHtml(string html, string selector)
    {
        if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(selector)) return false;

        // [data-path="..."] form
        var dpMatch = System.Text.RegularExpressions.Regex.Match(
            selector, @"^\[data-path=""(.+)""\]$");
        if (dpMatch.Success)
        {
            var path = dpMatch.Groups[1].Value;
            // CONSISTENCY(pptx-group-flatten): mirror FindDataPathInHtml's
            // ancestor fallback so a goto target inside a group still scrolls
            // to the nearest rendered ancestor instead of being rejected.
            return FindDataPathInHtml(html, path) != null;
        }

        // #anchor-id form
        if (selector.StartsWith("#"))
        {
            var id = selector.Substring(1);
            return html.IndexOf("id=\"" + id + "\"", StringComparison.Ordinal) >= 0
                || html.IndexOf("id='" + id + "'", StringComparison.Ordinal) >= 0;
        }

        // Unknown selector form — let it through (best-effort) so future
        // anchor styles aren't blocked.
        return true;
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
        // The sidebar also emits `<div class="thumb" data-slide="N">`, so matching
        // on `data-slide="N"` alone hits the thumb first and leaves the main
        // slide-container stale — user-visible as a white main view on every
        // incremental update. Pin to the slide-container class.
        var marker = $"class=\"slide-container\" data-slide=\"{slideNum}\"";
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

        // Section count change → fall back to full diff. Block <wb>/<we>
        // markers can straddle a section boundary (e.g. when a new section
        // is appended, the trailing block's <wb> sits in the prior section's
        // page-body and its <we> in the new section's page-body). Treating
        // that span as block content would inject structural markup
        // (</page-body></page></page-wrapper><page-wrapper data-section="N">…)
        // into the previous section's page-body, producing nested pages.
        var oldSecCount = System.Text.RegularExpressions.Regex.Matches(oldHtml, @"data-section=""\d+""").Count;
        var newSecCount = System.Text.RegularExpressions.Regex.Matches(newHtml, @"data-section=""\d+""").Count;
        if (oldSecCount != newSecCount) return null;

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

        // A block's <wb>…<we> markers can straddle a structural container, so its
        // captured content is structurally unbalanced — it opens a container it
        // never closes, or closes one it never opened. Known cases:
        //   • a paragraph with an inline <w:br type="page"/> — its span includes
        //     </page-body></page></page-wrapper><div class="page-wrapper">…<page-body>
        //     (page count is unchanged, so the section-count guard misses it);
        //   • a list — the <ol>/<ul> opens in the list block but the matching
        //     </ol>/</ul> closes inside the NEXT block's span;
        //   • multi-column / drop-cap wrappers split across blocks the same way.
        // Re-applying such a payload via innerHTML corrupts the live DOM (the
        // sibling-walk in wordPatchUpdate can't cross the container boundary):
        // an injected page-wrapper nests a page inside a page; an orphaned
        // </ol> wipes the list. Detect the straddle on the patch payload and
        // fall back to a full refresh, which rebuilds the structure correctly.
        foreach (var p in patches)
            if (WordPatchPayloadStraddlesStructure(p.Html))
                return null;

        // If more than 60% of blocks changed (and enough blocks to matter), fallback to full refresh
        var totalBlocks = Math.Max(oldBlocks.Count, newBlocks.Count);
        if (totalBlocks >= 5 && patches.Count > totalBlocks * 0.6)
            return null;

        return patches;
    }

    // Matches any HTML start/end tag: group1 = "/" for an end tag, group2 = tag
    // name, group3 = "/" for an explicit self-close (<x/>). Comments (<!-- -->)
    // and the XML/doctype declarations don't match — group2 requires a leading
    // ASCII letter. Attribute values never contain a raw '>' (the renderer
    // HTML-encodes them), so a greedy `[^>]*?` to the tag's own '>' is safe.
    private static readonly System.Text.RegularExpressions.Regex _htmlTagRx =
        new(@"<(/?)([a-zA-Z][a-zA-Z0-9:-]*)\b[^>]*?(/?)>",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    // Tags excluded from the balance count. Two groups, same reason — neither
    // can make a block straddle a structural boundary:
    //   • void elements — never carry children (<br>, <img>, <col> …);
    //   • inline elements — the renderer always opens AND closes them within a
    //     single run/paragraph render, so they are self-contained inside one
    //     block by construction. Skipping them also hardens the balance count
    //     against malformed inline markup buried in an attribute value (a raw
    //     '>' the real renderer would have encoded as &gt;).
    // Everything NOT in this set is treated as a potential block-level container
    // and counted — so a future block container the renderer starts emitting is
    // covered without editing this list. grep CONSISTENCY(word-patch-straddle).
    private static readonly HashSet<string> _inlineOrVoidHtmlTags = new(StringComparer.OrdinalIgnoreCase)
    {
        // void
        "area", "base", "br", "col", "embed", "hr", "img", "input",
        "link", "meta", "param", "source", "track", "wbr",
        // inline / phrasing
        "a", "abbr", "b", "bdi", "bdo", "cite", "code", "data", "dfn", "em",
        "font", "i", "kbd", "label", "mark", "q", "rp", "rt", "ruby", "s",
        "samp", "small", "span", "strong", "sub", "sup", "time", "u", "var",
    };

    /// <summary>
    /// True when a Word block-diff patch payload is unsafe to splice into the
    /// live DOM incrementally — i.e. the source block's &lt;wb&gt;/&lt;we&gt;
    /// markers straddle a structural element.
    ///
    /// Root invariant (not a list of known cases): the client splice
    /// (wordPatchUpdate) walks DOM *siblings* between the &lt;wb&gt; and
    /// &lt;we&gt; markers. That only works when both markers sit at the same DOM
    /// depth, which holds **iff** the captured payload is a well-balanced HTML
    /// fragment with no leading orphan-close. So we test exactly that, over
    /// EVERY element tag — no enumeration of containers (page-wrapper, ol/ul,
    /// multi-column / drop-cap div, table, …). Any present-or-future renderer
    /// shape that straddles a container is rejected, and the caller falls back
    /// to a full refresh. CONSISTENCY(word-patch-straddle).
    /// </summary>
    internal static bool WordPatchPayloadStraddlesStructure(string? html)
    {
        if (string.IsNullOrEmpty(html)) return false;

        var depth = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (System.Text.RegularExpressions.Match m in _htmlTagRx.Matches(html))
        {
            var tag = m.Groups[2].Value;
            if (m.Groups[3].Value == "/") continue;            // explicit self-close <x/>
            if (_inlineOrVoidHtmlTags.Contains(tag)) continue; // inline / void — never straddles

            if (m.Groups[1].Value == "/")
            {
                // A close whose matching open was never seen in this payload —
                // it lives in a sibling block (e.g. </ol> after a list block,
                // </page-wrapper> from a mid-paragraph page break). The markers
                // are at different DOM depths → unsafe.
                var d = depth.GetValueOrDefault(tag) - 1;
                if (d < 0) return true;
                depth[tag] = d;
            }
            else
            {
                depth[tag] = depth.GetValueOrDefault(tag) + 1;
            }
        }
        // Any element left open at the end straddles into the next block.
        foreach (var d in depth.Values) if (d != 0) return true;
        return false;
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

    /// <summary>
    /// Signature of Excel table chrome — concatenates each sheet's &lt;colgroup&gt;,
    /// &lt;thead&gt;, and the &lt;table&gt; open tag (which carries table width style).
    /// Row-level patches only swap &lt;tr&gt; nodes, so if this signature changes
    /// between old and new HTML (column added/removed, column width changed,
    /// thead style changed) the browser needs a full body refresh — otherwise
    /// new headers/widths stay stale until a manual reload.
    /// </summary>
    private static string TableChromeSignature(string html)
    {
        var sb = new System.Text.StringBuilder();
        foreach (System.Text.RegularExpressions.Match m in
            System.Text.RegularExpressions.Regex.Matches(
                html, @"<colgroup>.*?</colgroup>",
                System.Text.RegularExpressions.RegexOptions.Singleline))
            sb.Append(m.Value).Append('|');
        foreach (System.Text.RegularExpressions.Match m in
            System.Text.RegularExpressions.Regex.Matches(
                html, @"<thead>.*?</thead>",
                System.Text.RegularExpressions.RegexOptions.Singleline))
            sb.Append(m.Value).Append('|');
        foreach (System.Text.RegularExpressions.Match m in
            System.Text.RegularExpressions.Regex.Matches(html, @"<table[^>]*>"))
            sb.Append(m.Value).Append('|');
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

            // Anti-DNS-rebinding gate. A rebinding attack reaches this loopback
            // port but the request carries the attacker's domain in the Host
            // header (a header page JS cannot forge), so any request whose Host
            // is not a loopback name is rejected — including GET / and the SSE
            // stream, which would otherwise leak the whole document. Embedder-
            // agnostic: direct browser tabs and Electron <webview>s send a
            // localhost Host automatically; reverse proxies that forward a
            // non-loopback Host can allowlist it via OFFICECLI_WATCH_ALLOWED_HOSTS.
            if (!IsHostAllowed(headers))
            {
                await WriteForbiddenAsync(stream, ForbiddenHostMessage(headers), token);
                client.Close();
                return;
            }

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
                if (!IsOriginAllowed(headers))
                {
                    await WriteForbiddenAsync(stream, ForbiddenOriginMessage(headers), token);
                    client.Close();
                    return;
                }
                await HandlePostSelectionAsync(stream, headers, bodyPrefix, token);
                client.Close();
                return;
            }

            if (requestLine.StartsWith("POST /api/send", StringComparison.Ordinal))
            {
                if (!IsOriginAllowed(headers))
                {
                    await WriteForbiddenAsync(stream, ForbiddenOriginMessage(headers), token);
                    client.Close();
                    return;
                }
                await HandlePostSendAsync(stream, headers, bodyPrefix, WantsJson(requestLine), token);
                client.Close();
                return;
            }

            if (requestLine.StartsWith("POST /api/batch", StringComparison.Ordinal))
            {
                if (!IsOriginAllowed(headers))
                {
                    await WriteForbiddenAsync(stream, ForbiddenOriginMessage(headers), token);
                    client.Close();
                    return;
                }
                await HandlePostBatchAsync(stream, headers, bodyPrefix, WantsJson(requestLine), token);
                client.Close();
                return;
            }

            // GET /api/status — current document identity + version, so an SSE
            // client that reconnects (or an embedder like the editor) can
            // self-correct after missing a doc-switched event. Read-only; the
            // host gate above is the only guard, same as GET / and /events.
            if (requestLine.StartsWith("GET /api/status", StringComparison.Ordinal))
            {
                await HandleGetStatusAsync(stream, token);
                client.Close();
                return;
            }

            // POST /api/switch — retarget this server to another document in
            // place (port and SSE connections survive; clients get a
            // doc-switched event). See HandlePostSwitchAsync for semantics.
            if (requestLine.StartsWith("POST /api/switch", StringComparison.Ordinal))
            {
                if (!IsOriginAllowed(headers))
                {
                    await WriteForbiddenAsync(stream, ForbiddenOriginMessage(headers), token);
                    client.Close();
                    return;
                }
                await HandlePostSwitchAsync(stream, headers, bodyPrefix, token);
                client.Close();
                return;
            }

            // Wrong verb on the endpoints above → 405, mirroring the
            // /api/selection guard (BUG-TESTER-R503): an API client using the
            // wrong method must not silently receive the HTML preview page.
            if (requestLine.Contains(" /api/status"))
            {
                var msg405 = Encoding.UTF8.GetBytes("Method Not Allowed: /api/status only accepts GET");
                var hdr405 = Encoding.UTF8.GetBytes(
                    $"HTTP/1.1 405 Method Not Allowed\r\nAllow: GET\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {msg405.Length}\r\nConnection: close\r\n\r\n");
                await stream.WriteAsync(hdr405, token);
                await stream.WriteAsync(msg405, token);
                client.Close();
                return;
            }
            if (requestLine.Contains(" /api/switch"))
            {
                var msg405 = Encoding.UTF8.GetBytes("Method Not Allowed: /api/switch only accepts POST");
                var hdr405 = Encoding.UTF8.GetBytes(
                    $"HTTP/1.1 405 Method Not Allowed\r\nAllow: POST\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {msg405.Length}\r\nConnection: close\r\n\r\n");
                await stream.WriteAsync(hdr405, token);
                await stream.WriteAsync(msg405, token);
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

    // Loopback host names accepted in the Host/Origin headers. Seeded with the
    // standard loopback identities and extended (once, at first use) from
    // OFFICECLI_WATCH_ALLOWED_HOSTS for reverse-proxy setups that forward a
    // non-loopback Host upstream.
    private static readonly HashSet<string> _allowedHosts = BuildAllowedHosts();

    private static HashSet<string> BuildAllowedHosts()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "localhost", "127.0.0.1", "[::1]", "::1" };
        var extra = Environment.GetEnvironmentVariable("OFFICECLI_WATCH_ALLOWED_HOSTS");
        if (!string.IsNullOrWhiteSpace(extra))
            foreach (var h in extra.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                set.Add(h);
        return set;
    }

    /// <summary>Strip the optional <c>:port</c> from a Host header value, preserving bracketed IPv6.</summary>
    internal static string ExtractHostname(string hostHeader)
    {
        var v = hostHeader.Trim();
        if (v.StartsWith("[", StringComparison.Ordinal)) // [::1] or [::1]:port
        {
            var rb = v.IndexOf(']');
            return rb >= 0 ? v[..(rb + 1)] : v;
        }
        var colon = v.IndexOf(':');
        return colon >= 0 ? v[..colon] : v;
    }

    /// <summary>True if the request's Host header names an accepted loopback host (anti-rebinding).</summary>
    internal static bool IsHostAllowed(Dictionary<string, string> headers)
    {
        // HTTP/1.1 mandates Host and every browser sends it; a missing/blank
        // Host is treated as untrusted and rejected.
        if (!headers.TryGetValue("Host", out var host) || string.IsNullOrWhiteSpace(host))
            return false;
        return _allowedHosts.Contains(ExtractHostname(host));
    }

    /// <summary>
    /// True if a state-changing request's Origin is absent or names a loopback host.
    /// Absent Origin (server-side proxy hop, or a same-origin navigation that omits
    /// it) is allowed; a present cross-origin Origin is rejected (CSRF defense).
    /// </summary>
    internal static bool IsOriginAllowed(Dictionary<string, string> headers)
    {
        if (!headers.TryGetValue("Origin", out var origin) || string.IsNullOrWhiteSpace(origin))
            return true;
        if (Uri.TryCreate(origin.Trim(), UriKind.Absolute, out var u))
            return _allowedHosts.Contains(u.Host) || _allowedHosts.Contains($"[{u.Host}]");
        return false;
    }

    private static string ForbiddenHostMessage(Dictionary<string, string> headers)
    {
        headers.TryGetValue("Host", out var host);
        return $"403 Forbidden: request Host '{host ?? "(none)"}' is not a recognized loopback host.\n" +
               "The officecli watch preview only accepts Host: localhost / 127.0.0.1 (anti-DNS-rebinding).\n" +
               "If you reach it through a reverse proxy that forwards a different Host, set\n" +
               "OFFICECLI_WATCH_ALLOWED_HOSTS=<hostname>[,<hostname>...] before starting `officecli watch`.\n";
    }

    private static string ForbiddenOriginMessage(Dictionary<string, string> headers)
    {
        headers.TryGetValue("Origin", out var origin);
        return $"403 Forbidden: cross-origin request from Origin '{origin ?? "(none)"}' is not allowed for this endpoint.\n";
    }

    private static async Task WriteForbiddenAsync(NetworkStream stream, string message, CancellationToken token)
    {
        var msg = Encoding.UTF8.GetBytes(message);
        var hdr = Encoding.UTF8.GetBytes(
            $"HTTP/1.1 403 Forbidden\r\nContent-Type: text/plain; charset=utf-8\r\nContent-Length: {msg.Length}\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(hdr, token);
        await stream.WriteAsync(msg, token);
    }

    /// <summary>
    /// Read the HTTP request line and headers, plus any body bytes that arrived in the
    /// same TCP read. Returns (requestLine, headers, bodyPrefix). Caller is responsible
    /// for reading the rest of the body using Content-Length if needed.
    ///
    /// The body prefix comes back as raw BYTES, not text: a UTF-8 sequence can
    /// straddle the boundary between this read and the caller's, and decoding
    /// each half on its own turns both into U+FFFD. Decoding happens once, in
    /// ReadPostBodyAsync, over the assembled payload.
    /// </summary>
    private static async Task<(string requestLine, Dictionary<string, string> headers, byte[] bodyPrefix)>
        ReadHttpRequestHeaderAsync(NetworkStream stream, CancellationToken token)
    {
        var buffer = new byte[8192];
        var raw = new MemoryStream();
        int headerEnd = -1;
        while (headerEnd < 0)
        {
            var n = await stream.ReadAsync(buffer.AsMemory(), token);
            if (n == 0) break;
            raw.Write(buffer, 0, n);
            // Rescan from 3 bytes before the new data so a terminator split
            // across reads is still found. Bounded by the 32 KB cap below.
            var scanFrom = Math.Max(0, (int)raw.Length - n - 3);
            var found = raw.GetBuffer().AsSpan(scanFrom, (int)raw.Length - scanFrom).IndexOf("\r\n\r\n"u8);
            if (found >= 0) headerEnd = scanFrom + found;
            if (raw.Length > 32 * 1024) break; // safety cap
        }

        var rawBytes = raw.GetBuffer();
        var total = (int)raw.Length;
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (headerEnd < 0)
        {
            // No header terminator — treat the whole thing as a single line
            var all = Encoding.UTF8.GetString(rawBytes, 0, total);
            var firstLine = all;
            var crlf = all.IndexOf("\r\n", StringComparison.Ordinal);
            if (crlf >= 0) firstLine = all[..crlf];
            return (firstLine, headers, Array.Empty<byte>());
        }

        var headerSection = Encoding.UTF8.GetString(rawBytes, 0, headerEnd);
        var bodyPrefix = rawBytes.AsSpan(headerEnd + 4, total - (headerEnd + 4)).ToArray();
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

    /// <summary>
    /// Assemble a POST body from the bytes that arrived with the headers plus
    /// whatever Content-Length says is still on the wire, then decode the whole
    /// payload as UTF-8 exactly once.
    ///
    /// Decoding per socket read instead splits any multi-byte sequence that
    /// straddles a read boundary into two independent fragments, and each
    /// fragment decodes to U+FFFD — a large CJK payload came back from a
    /// 200 OK with characters silently replaced, and the replacements went
    /// straight into the document. Every /api POST shares this path so the
    /// bound checks stay identical across endpoints.
    ///
    /// Bounded by MaxSelectionBodyBytes (FUZZER-001 slow-loris) and
    /// PostBodyReadTimeout. A prefix overshooting Content-Length is trimmed:
    /// otherwise extra bytes could be smuggled in the header segment
    /// (FUZZER-002). A request with no Content-Length keeps the prefix as-is.
    /// </summary>
    private static async Task<string> ReadPostBodyAsync(
        NetworkStream stream, Dictionary<string, string> headers, byte[] bodyPrefix, CancellationToken token)
    {
        int contentLength = -1;
        if (headers.TryGetValue("Content-Length", out var clStr) && int.TryParse(clStr, out var parsedCl))
        {
            if (parsedCl < 0 || parsedCl > MaxSelectionBodyBytes)
                throw new InvalidDataException("body too large");
            contentLength = parsedCl;
        }

        if (contentLength < 0) return Encoding.UTF8.GetString(bodyPrefix);
        if (bodyPrefix.Length >= contentLength) return Encoding.UTF8.GetString(bodyPrefix, 0, contentLength);

        var body = new byte[contentLength];
        bodyPrefix.CopyTo(body, 0);
        int have = bodyPrefix.Length;
        using var readCts = CancellationTokenSource.CreateLinkedTokenSource(token);
        readCts.CancelAfter(PostBodyReadTimeout);
        try
        {
            while (have < contentLength)
            {
                var n = await stream.ReadAsync(body.AsMemory(have, contentLength - have), readCts.Token);
                if (n == 0) break;
                have += n;
            }
        }
        catch (OperationCanceledException) when (!token.IsCancellationRequested)
        {
            throw new InvalidDataException("body read timed out");
        }
        return Encoding.UTF8.GetString(body, 0, have);
    }

    private async Task HandlePostSelectionAsync(NetworkStream stream, Dictionary<string, string> headers, byte[] bodyPrefix, CancellationToken token)
    {
        int statusCode = 204;
        string statusText = "No Content";

        try
        {
            var body = await ReadPostBodyAsync(stream, headers, bodyPrefix, token);

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
            $"HTTP/1.1 {statusCode} {statusText}\r\nContent-Length: 0\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(resp, token);
    }

    /// <summary>
    /// Handle POST /api/send — spawn officecli set/add/remove as a child process
    /// to modify the file, mirroring the SDKs' send(item) (one batch-item, this
    /// call's own status is the result). That command auto-notifies the watch
    /// server via named pipe, triggering an SSE refresh. WatchServer never opens
    /// the file directly — widening this beyond `set` must stay a child-process
    /// spawn, not an in-process handler call, so this can never become a second
    /// writer racing a live resident.
    /// </summary>
    private async Task HandlePostSendAsync(NetworkStream stream, Dictionary<string, string> headers, byte[] bodyPrefix, bool json, CancellationToken token)
    {
        try
        {
            var body = await ReadPostBodyAsync(stream, headers, bodyPrefix, token);

            // Same batch-item vocabulary as CLI batch / the SDKs' send(item)
            // {"command": "set"|"add"|"remove", ...}.
            // Bare {"path", "props"} or legacy {"path", "prop", "value"} with
            // no "command" field default to "set" for pre-existing callers.
            using var doc = System.Text.Json.JsonDocument.Parse(body);
            var root = doc.RootElement;
            var command = root.TryGetProperty("command", out var cmdEl) ? cmdEl.GetString() ?? "set" : "set";

            var exe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName
                ?? (OperatingSystem.IsWindows() ? "officecli.exe" : "officecli");
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = exe,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            switch (command.ToLowerInvariant())
            {
                case "add":
                {
                    // Accept the canonical batch-item key "path" (used by /api/batch,
                    // the SDKs) as well
                    // as the legacy "parent". /api/send must accept the same item shape
                    // as /api/batch (see HandlePostBatchAsync doc) — otherwise `add`
                    // diverges per backend and a `path`-shaped item throws a raw
                    // KeyNotFoundException instead of running.
                    var parent = (root.TryGetProperty("path", out var addPathEl) ? addPathEl.GetString() : null)
                        ?? (root.TryGetProperty("parent", out var addParentEl) ? addParentEl.GetString() : null)
                        ?? "";
                    psi.ArgumentList.Add("add");
                    psi.ArgumentList.Add(_filePath);
                    psi.ArgumentList.Add(parent);
                    // --from clones an existing element (shape/slide); it is
                    // mutually exclusive with --type/--prop (see `add`), so when
                    // present it is the whole command.
                    if (root.TryGetProperty("from", out var fromEl) && fromEl.GetString() is { } from)
                    {
                        psi.ArgumentList.Add("--from");
                        psi.ArgumentList.Add(from);
                    }
                    else
                    {
                        if (root.TryGetProperty("type", out var typeEl) && typeEl.GetString() is { } type)
                        {
                            psi.ArgumentList.Add("--type");
                            psi.ArgumentList.Add(type);
                        }
                        AppendProps(psi, root);
                    }
                    // Position hints apply to both clone and typed add.
                    AppendPositionArgs(psi, root);
                    break;
                }
                case "remove":
                {
                    var path = root.GetProperty("path").GetString() ?? "";
                    psi.ArgumentList.Add("remove");
                    psi.ArgumentList.Add(_filePath);
                    psi.ArgumentList.Add(path);
                    break;
                }
                case "get":
                {
                    // Read-only: spawn `officecli get <path>` (served from the
                    // resident's current in-memory state). Used by the editor to
                    // read a property's prior value / capture an element before
                    // deletion for undo. Still a child process — no in-process
                    // document access, so the watch red line holds.
                    var path = root.GetProperty("path").GetString() ?? "";
                    psi.ArgumentList.Add("get");
                    psi.ArgumentList.Add(_filePath);
                    psi.ArgumentList.Add(path);
                    break;
                }
                case "move":
                {
                    var path = root.GetProperty("path").GetString() ?? "";
                    psi.ArgumentList.Add("move");
                    psi.ArgumentList.Add(_filePath);
                    psi.ArgumentList.Add(path);
                    if (root.TryGetProperty("to", out var toEl) && toEl.GetString() is { } to)
                    { psi.ArgumentList.Add("--to"); psi.ArgumentList.Add(to); }
                    AppendPositionArgs(psi, root);
                    break;
                }
                case "swap":
                {
                    var path1 = root.GetProperty("path").GetString() ?? "";
                    // Canonical second path is "path2"; accept legacy "to".
                    var path2 = root.TryGetProperty("path2", out var p2El) ? p2El.GetString() ?? ""
                        : root.TryGetProperty("to", out var toEl2) ? toEl2.GetString() ?? "" : "";
                    psi.ArgumentList.Add("swap");
                    psi.ArgumentList.Add(_filePath);
                    psi.ArgumentList.Add(path1);
                    psi.ArgumentList.Add(path2);
                    break;
                }
                case "set":
                default:
                {
                    var path = root.GetProperty("path").GetString() ?? "";
                    psi.ArgumentList.Add("set");
                    psi.ArgumentList.Add(_filePath);
                    psi.ArgumentList.Add(path);
                    if (root.TryGetProperty("props", out var propsEl) && propsEl.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        AppendProps(psi, root);
                    }
                    else
                    {
                        // Legacy shape: {"path", "prop", "value"} (single property, no "props" object).
                        var prop = root.GetProperty("prop").GetString() ?? "text";
                        var value = root.GetProperty("value").GetString() ?? "";
                        psi.ArgumentList.Add("--prop");
                        psi.ArgumentList.Add($"{prop}={PropValueForArgv(prop, value)}");
                    }
                    break;
                }
            }

            // --json is the CLI's opt-in for the structured envelope; omit it
            // for plain text. The flag only changes what officecli prints, i.e.
            // what ends up inside the comm envelope's `message`.
            if (json) psi.ArgumentList.Add("--json");

            string output = "";
            using var proc = System.Diagnostics.Process.Start(psi);
            if (proc != null)
            {
                // Drain stderr concurrently: it is redirected (so it MUST be
                // read) but never surfaced — a child that fills the ~64KB pipe
                // buffer with warnings would otherwise deadlock this request.
                var drainErr = proc.StandardError.ReadToEndAsync(token);
                output = await proc.StandardOutput.ReadToEndAsync(token);
                await proc.WaitForExitAsync(token);
                _ = await drainErr;
                // command auto-notifies watch via named pipe → SSE refresh
            }
            await WriteCommEnvelopeAsync(stream, true, output.TrimEnd('\n', '\r'), token);
        }
        catch (System.Exception ex)
        {
            await WriteCommEnvelopeAsync(stream, false, ex.Message, token);
        }
    }

    /// <summary>
    /// Handle POST /api/batch — spawn officecli batch as a child process,
    /// mirroring the SDKs' batch(items). The posted
    /// body is a JSON array of the same batch-item shape /api/send accepts;
    /// `officecli batch --commands` already takes that array verbatim, so
    /// unlike /api/send there is no per-command arg-building here — the
    /// whole body passes straight through. Same child-process-spawn
    /// constraint as /api/send — never touch the document in-process.
    /// </summary>
    private async Task HandlePostBatchAsync(NetworkStream stream, Dictionary<string, string> headers, byte[] bodyPrefix, bool json, CancellationToken token)
    {
        try
        {
            var body = await ReadPostBodyAsync(stream, headers, bodyPrefix, token);

            var exe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName
                ?? (OperatingSystem.IsWindows() ? "officecli.exe" : "officecli");
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = exe,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            psi.ArgumentList.Add("batch");
            psi.ArgumentList.Add(_filePath);
            psi.ArgumentList.Add("--commands");
            psi.ArgumentList.Add(body);
            // --json opts into the structured envelope; omit for plain text.
            if (json) psi.ArgumentList.Add("--json");

            string output = "";
            using var proc = System.Diagnostics.Process.Start(psi);
            if (proc != null)
            {
                // Drain stderr concurrently: it is redirected (so it MUST be
                // read) but never surfaced — a child that fills the ~64KB pipe
                // buffer with warnings would otherwise deadlock this request.
                var drainErr = proc.StandardError.ReadToEndAsync(token);
                output = await proc.StandardOutput.ReadToEndAsync(token);
                await proc.WaitForExitAsync(token);
                _ = await drainErr;
                // batch auto-notifies watch via named pipe → SSE refresh
            }
            await WriteCommEnvelopeAsync(stream, true, output.TrimEnd('\n', '\r'), token);
        }
        catch (System.Exception ex)
        {
            await WriteCommEnvelopeAsync(stream, false, ex.Message, token);
        }
    }

    /// <summary>Lowercase extension without the dot: pptx / docx / xlsx.</summary>
    private static string FormatOf(string filePath)
        => Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();

    /// <summary>{"file":..,"name":..,"fmt":..,"version":N} for /api/status,
    /// the /api/switch 200 body, and the doc-switched SSE event.</summary>
    private string BuildStatusJson()
    {
        var file = _filePath;
        var sb = new StringBuilder();
        sb.Append("{\"file\":");
        AppendJsonString(sb, file);
        sb.Append(",\"name\":");
        AppendJsonString(sb, Path.GetFileName(file));
        sb.Append(",\"fmt\":");
        AppendJsonString(sb, FormatOf(file));
        sb.Append(",\"version\":").Append(_version);
        sb.Append('}');
        return sb.ToString();
    }

    private static async Task WriteJsonResponseAsync(NetworkStream stream, int statusCode, string reason, string json, CancellationToken token)
    {
        var body = Encoding.UTF8.GetBytes(json);
        var header = Encoding.UTF8.GetBytes(
            $"HTTP/1.1 {statusCode} {reason}\r\nContent-Type: application/json; charset=utf-8\r\nContent-Length: {body.Length}\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(header, token);
        await stream.WriteAsync(body, token);
    }

    private async Task HandleGetStatusAsync(NetworkStream stream, CancellationToken token)
    {
        _lastActivityTime = DateTime.UtcNow;
        await WriteJsonResponseAsync(stream, 200, "OK", BuildStatusJson(), token);
    }

    /// <summary>
    /// Handle POST /api/switch {"file": "/abs/path.docx"} — retarget this
    /// server to another document IN PLACE: the port stays, live SSE
    /// connections stay, clients get a doc-switched event (built-in preview
    /// script reacts with location.reload; embedders reset their own state).
    ///
    /// Switch is atomic from the client's point of view: the new document's
    /// HTML is fully rendered BEFORE any state is swapped, so on any failure
    /// (bad request 400, missing file 404, already watched elsewhere 409,
    /// render failure 500) the current document keeps serving untouched and
    /// no connection is dropped. The 409 body includes the occupying watch's
    /// port so an embedder can reuse that server instead of erroring.
    ///
    /// Rendering follows the same red line as /api/send / /api/batch: spawn
    /// an officecli child process (`view file html --out tmp`) and take
    /// its baked HTML — resident-first routing lives inside the CLI
    /// (TryResident), and WatchServer itself never opens the document.
    /// </summary>
    private async Task HandlePostSwitchAsync(NetworkStream stream, Dictionary<string, string> headers, byte[] bodyPrefix, CancellationToken token)
    {
        try
        {
            var body = await ReadPostBodyAsync(stream, headers, bodyPrefix, token);

            string? requested;
            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(body);
                requested = doc.RootElement.TryGetProperty("file", out var fileEl) ? fileEl.GetString() : null;
            }
            catch (System.Text.Json.JsonException)
            {
                await WriteJsonResponseAsync(stream, 400, "Bad Request",
                    "{\"error\":\"body must be JSON: {\\\"file\\\": \\\"/abs/path\\\"}\"}", token);
                return;
            }
            if (string.IsNullOrWhiteSpace(requested))
            {
                await WriteJsonResponseAsync(stream, 400, "Bad Request",
                    "{\"error\":\"missing required field: file\"}", token);
                return;
            }

            var newPath = Path.GetFullPath(requested);
            var fmt = FormatOf(newPath);
            if (fmt is not ("pptx" or "docx" or "xlsx"))
            {
                await WriteJsonResponseAsync(stream, 400, "Bad Request",
                    "{\"error\":\"unsupported file type — expected .pptx, .docx or .xlsx\"}", token);
                return;
            }
            if (!File.Exists(newPath))
            {
                var sb404 = new StringBuilder("{\"error\":");
                AppendJsonString(sb404, $"file not found: {newPath}");
                sb404.Append('}');
                await WriteJsonResponseAsync(stream, 404, "Not Found", sb404.ToString(), token);
                return;
            }

            // Same-file no-op guard is intentionally NOT here: re-switching to
            // the current file is a supported "force refresh + reset marks"
            // gesture and takes the same path below (pipe/marker swap degrades
            // to a no-op because the names come out identical).

            // Target already held by ANOTHER live watch → 409 with its port so
            // the caller can reuse that server. Same single-consumer constraint
            // as the startup duplicate check in RunAsync: the per-file update
            // pipe cannot have two listeners. This is a pre-render fast-fail
            // only — the render below takes seconds, so the swap re-derives
            // pipe identity under _switchLock instead of trusting this value.
            if (GetWatchPipeName(newPath) != _pipeName)
            {
                var occupied = GetExistingWatchPort(newPath);
                if (occupied.HasValue)
                {
                    await WriteJsonResponseAsync(stream, 409, "Conflict",
                        $"{{\"error\":\"file is already watched by another process\",\"port\":{occupied.Value}}}", token);
                    return;
                }
            }

            // Render the new document BEFORE touching any state. Child process
            // + --out tempfile: stdout noise (e.g. skill-refresh notices) can
            // never corrupt the HTML, and stderr is drained to avoid pipe-full
            // deadlock, same as /api/send.
            var tmpOut = Path.Combine(Path.GetTempPath(), $"officecli_switch_{Guid.NewGuid():N}.html");
            string html;
            try
            {
                var exe = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName
                    ?? (OperatingSystem.IsWindows() ? "officecli.exe" : "officecli");
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = exe,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                    StandardOutputEncoding = System.Text.Encoding.UTF8,
                    StandardErrorEncoding = System.Text.Encoding.UTF8,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                psi.ArgumentList.Add("view");
                psi.ArgumentList.Add(newPath);
                psi.ArgumentList.Add("html"); // mode is positional: view <file> <mode>
                psi.ArgumentList.Add("--out");
                psi.ArgumentList.Add(tmpOut);

                string renderErr = "";
                int exitCode = -1;
                using (var proc = System.Diagnostics.Process.Start(psi))
                {
                    if (proc != null)
                    {
                        using var renderCts = CancellationTokenSource.CreateLinkedTokenSource(token);
                        renderCts.CancelAfter(TimeSpan.FromSeconds(120));
                        var drainErr = proc.StandardError.ReadToEndAsync(renderCts.Token);
                        var drainOut = proc.StandardOutput.ReadToEndAsync(renderCts.Token);
                        try
                        {
                            await proc.WaitForExitAsync(renderCts.Token);
                        }
                        catch (OperationCanceledException)
                        {
                            try { proc.Kill(entireProcessTree: true); } catch { }
                            throw new InvalidOperationException("render timed out after 120s");
                        }
                        renderErr = await drainErr;
                        _ = await drainOut;
                        exitCode = proc.ExitCode;
                    }
                }
                if (exitCode != 0 || !File.Exists(tmpOut))
                {
                    var reason = string.IsNullOrWhiteSpace(renderErr) ? $"render exited {exitCode}" : renderErr.Trim();
                    throw new InvalidOperationException(reason);
                }
                html = await File.ReadAllTextAsync(tmpOut, token);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                var sb500 = new StringBuilder("{\"error\":");
                AppendJsonString(sb500, $"failed to render {Path.GetFileName(newPath)}: {ex.Message}");
                sb500.Append('}');
                await WriteJsonResponseAsync(stream, 500, "Internal Server Error", sb500.ToString(), token);
                return;
            }
            finally
            {
                try { if (File.Exists(tmpOut)) File.Delete(tmpOut); } catch { }
            }

            // HTML in hand — swap state. Serialized so concurrent switches
            // can't interleave their pipe/marker swaps.
            await _switchLock.WaitAsync(token);
            try
            {
                // Re-derive pipe identity NOW, under the lock. The pre-render
                // check ran seconds ago; a concurrent switch may have
                // retargeted this server since. Acting on a stale latch split
                // identity from content: a same-file force-refresh that lost
                // the race skipped the _filePath/_pipeName update (gated) yet
                // still overwrote _currentHtml (unconditional), leaving status
                // reporting one document while GET / served another until the
                // next switch.
                var oldPipeName = _pipeName;
                var newPipeName = GetWatchPipeName(newPath);
                var pipeChanged = newPipeName != oldPipeName;
                if (pipeChanged)
                {
                    DeleteMarker(); // old file's marker — call before _filePath changes
                    _filePath = newPath;
                    _pipeName = newPipeName;
                    // Kick the pipe listener out of WaitForConnectionAsync on the
                    // old name so its next iteration listens on the new one
                    // (same dummy-connect trick as StopAsync — cancellation alone
                    // is not reliable cross-platform).
                    try
                    {
                        using var kick = new System.IO.Pipes.NamedPipeClientStream(
                            ".", oldPipeName, System.IO.Pipes.PipeDirection.InOut);
                        kick.Connect(500);
                    }
                    catch { }
                    // Old pipe's Unix socket file is now stale (BUG-BT-003).
                    if (!OperatingSystem.IsWindows())
                    {
                        try
                        {
                            var sockPath = Path.Combine(Path.GetTempPath(), "CoreFxPipe_" + oldPipeName);
                            if (File.Exists(sockPath)) File.Delete(sockPath);
                        }
                        catch { }
                    }
                    WriteMarker(); // new file's marker
                }
                else
                {
                    // Same pipe — still track the requested path verbatim.
                    // Pipe names case-fold the path on Windows/macOS, so equal
                    // names don't guarantee equal strings, and identity must
                    // always match the content installed below.
                    _filePath = newPath;
                }

                // Marks/selection are per-document positional state — never
                // carried across a switch (paths from the old document are
                // meaningless in the new one). Same-file re-switch also resets:
                // "switch always resets marks" is the documented contract.
                lock (_marksLock)
                {
                    _currentMarks.Clear();
                    _marksVersion++;
                }
                lock (_selectionLock) { _currentSelection.Clear(); }

                _version = 0;
                _currentHtml = html;
                _lastActivityTime = DateTime.UtcNow;
                Console.WriteLine($"Watching: {_filePath}");
            }
            finally
            {
                _switchLock.Release();
            }

            // Notify SSE clients: named event (not "update") so legacy update
            // handlers never mis-parse it. Built-in preview script reloads;
            // embedders reset marks/selection/scroll and re-fetch.
            BroadcastSseNamed("doc-switched", BuildStatusJson());

            await WriteJsonResponseAsync(stream, 200, "OK", BuildStatusJson(), token);
        }
        catch (System.Exception ex)
        {
            try
            {
                var sbErr = new StringBuilder("{\"error\":");
                AppendJsonString(sbErr, ex.Message);
                sbErr.Append('}');
                await WriteJsonResponseAsync(stream, 500, "Internal Server Error", sbErr.ToString(), token);
            }
            catch { /* client gone mid-error — nothing to do */ }
        }
    }

    /// <summary>
    /// Broadcast an SSE message under a custom event name. BroadcastSse
    /// hardcodes "event: update" (the DOM-swap channel); doc-switched must be
    /// a distinct event so pre-existing update listeners ignore it.
    /// </summary>
    private void BroadcastSseNamed(string eventName, string json)
    {
        lock (_sseLock)
        {
            var dead = new List<NetworkStream>();
            foreach (var client in _sseClients)
            {
                try
                {
                    var data = Encoding.UTF8.GetBytes($"event: {eventName}\ndata: {json}\n\n");
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

    /// <summary>
    /// Parse the <c>?json</c> query flag off the request line. Absent, or any
    /// value other than <c>0</c>/<c>false</c> =&gt; true (structured, the
    /// default); <c>?json=0</c> / <c>?json=false</c> =&gt; plain text. Mirrors
    /// the CLI's <c>--json</c> opt-in.
    /// </summary>
    private static bool WantsJson(string requestLine)
    {
        int q = requestLine.IndexOf('?');
        if (q < 0) return true;
        int sp = requestLine.IndexOf(' ', q);
        string query = sp < 0 ? requestLine.Substring(q + 1) : requestLine.Substring(q + 1, sp - q - 1);
        foreach (var pair in query.Split('&'))
        {
            int eq = pair.IndexOf('=');
            string k = eq < 0 ? pair : pair.Substring(0, eq);
            if (k == "json")
            {
                string v = eq < 0 ? "1" : pair.Substring(eq + 1);
                return !(v == "0" || v.Equals("false", System.StringComparison.OrdinalIgnoreCase));
            }
        }
        return true;
    }

    /// <summary>
    /// Write the /api/send + /api/batch response as a communication envelope:
    /// <c>{ "success": bool, "message"|"error": string }</c>. <c>success</c>
    /// reflects the transport/process layer only (did the request reach
    /// officecli and run without crashing) — NOT officecli's business verdict,
    /// which rides inside <c>message</c> (its own <c>--json</c> envelope, or
    /// plain text). Callers unwrap this: on success take <c>message</c>
    /// (officecli's raw stdout); on failure, <c>error</c>. Always HTTP 200 —
    /// the envelope's <c>success</c> is the status signal.
    /// </summary>
    private static async Task WriteCommEnvelopeAsync(NetworkStream stream, bool success, string content, CancellationToken token)
    {
        // Trim/AOT-safe JSON build via Utf8JsonWriter (no reflection) — mirrors
        // CommandBuilder.PrintBatchResults. `content` is escaped by WriteString.
        byte[] bodyBytes;
        using (var ms = new System.IO.MemoryStream())
        {
            using (var w = new System.Text.Json.Utf8JsonWriter(ms))
            {
                w.WriteStartObject();
                w.WriteBoolean("success", success);
                w.WriteString(success ? "message" : "error", content);
                w.WriteEndObject();
            }
            bodyBytes = ms.ToArray();
        }
        var header = Encoding.UTF8.GetBytes(
            $"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: {bodyBytes.Length}\r\nConnection: close\r\n\r\n");
        await stream.WriteAsync(header, token);
        await stream.WriteAsync(bodyBytes, token);
    }

    private static void AppendProps(System.Diagnostics.ProcessStartInfo psi, System.Text.Json.JsonElement root)
    {
        if (!root.TryGetProperty("props", out var propsEl) || propsEl.ValueKind != System.Text.Json.JsonValueKind.Object)
            return;
        foreach (var kv in propsEl.EnumerateObject())
        {
            psi.ArgumentList.Add("--prop");
            psi.ArgumentList.Add($"{kv.Name}={PropValueForArgv(kv.Name, kv.Value.GetString() ?? "")}");
        }
    }

    /// <summary>
    /// CONSISTENCY(text-escape-boundary): props reach this server as JSON, where
    /// a backslash is already the final character the caller wants. The child CLI
    /// resolves C-escapes in text-valued props, so handing the value over bare
    /// resolves it a second time — a shape whose text is the path C:\temp\new.docx
    /// came out as "C:" + TAB + "emp" + newline + "ew.docx". /api/batch passes the
    /// same JSON through untouched, so without this the two endpoints disagreed on
    /// what one item means. Double the backslashes for exactly the keys the child
    /// resolves, leaving paths, colors and numbers alone.
    /// </summary>
    private static string PropValueForArgv(string key, string value)
        => OfficeCli.CommandBuilder.KeyTakesCEscapes(key) ? TextEscape.Protect(value) : value;

    /// <summary>
    /// Append the shared insert-position hints (--index / --after / --before)
    /// that add and move accept. Order-neutral; officecli resolves precedence.
    /// </summary>
    private static void AppendPositionArgs(System.Diagnostics.ProcessStartInfo psi, System.Text.Json.JsonElement root)
    {
        if (root.TryGetProperty("index", out var idxEl) && idxEl.ValueKind == System.Text.Json.JsonValueKind.Number)
        { psi.ArgumentList.Add("--index"); psi.ArgumentList.Add(idxEl.GetInt32().ToString(System.Globalization.CultureInfo.InvariantCulture)); }
        if (root.TryGetProperty("after", out var afEl) && afEl.GetString() is { } af)
        { psi.ArgumentList.Add("--after"); psi.ArgumentList.Add(af); }
        if (root.TryGetProperty("before", out var beEl) && beEl.GetString() is { } be)
        { psi.ArgumentList.Add("--before"); psi.ArgumentList.Add(be); }
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
            "HTTP/1.1 200 OK\r\nContent-Type: text/event-stream; charset=utf-8\r\nCache-Control: no-cache\r\nConnection: keep-alive\r\n\r\n");
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
