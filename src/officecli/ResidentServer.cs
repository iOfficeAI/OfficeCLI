// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using OfficeCli.Core;
using OfficeCli.Handlers;
using OfficeCli.Help;

namespace OfficeCli;

public class ResidentServer : IDisposable
{
    private IDocumentHandler _handler;
    private readonly string _filePath;
    private readonly string _pipeName;
    // Mutable: PromoteToEditable() latches this to true on first write so
    // subsequent reopen sites (view-screenshot, page-count, refresh) use
    // the promoted mode rather than reverting to the ctor's editable=false.
    private bool _editable;
    // True when in-memory mutations have not yet been flushed to disk.
    // PromoteToEditable() (the shared prelude of every mutating command and
    // of batch) latches it true; ExecuteSave() and the idle-autosave clear it
    // after a successful _handler.Save(). The idle-autosave watchdog uses it
    // to skip flushing when nothing changed since the last save.
    private volatile bool _dirty;
    // Stderr captured during DocumentHandlerFactory.Open (i.e. while the
    // constructor was building _handler). At that point there's no
    // per-command Console.SetError scope, so warnings written by plugin
    // invokers (e.g. "dump-reader produced no commands") would otherwise
    // be swallowed by the un-drained resident-stderr pipe. Held until the
    // first HandleRequest, which folds it into that command's stderr
    // envelope and clears the buffer.
    private string? _startupStderr;
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
    // Idle timeout is mutable: `create` starts the resident with a short
    // 60s timeout, and a later `open` upgrades it to 12min via the
    // `__set-idle-timeout__` ping command. Stored as ticks so we can
    // do atomic Volatile reads/writes (TimeSpan is a multi-field struct
    // and can't be volatile'd directly).
    private long _idleTimeoutTicks = ResolveIdleTimeout().Ticks;
    private TimeSpan CurrentIdleTimeout => TimeSpan.FromTicks(Volatile.Read(ref _idleTimeoutTicks));
    private CancellationTokenSource _idleCts = new();
    // Two independent idle timers, both reset by command activity (ResetIdleTimer):
    //   - _idleCts (CurrentIdleTimeout, default 12min / 60s auto-start): on
    //     elapse the resident shuts down — flush + release the file ("close").
    //   - _autosaveCts (CurrentAutosaveInterval, adaptive 2s–10s by default):
    //     on elapse, if the DOM is dirty, flush to disk but KEEP the resident
    //     running ("save"). This makes edits visible to third-party readers
    //     shortly after going idle without paying the O(n) reopen cost a
    //     shutdown+restart would incur. Not started in off/each flush modes.
    // The shorter autosave fires first; if the session stays idle the longer
    // idle-shutdown follows. Autosave firing does NOT reset the shutdown timer,
    // so the resident still exits the configured time after the last command.
    private CancellationTokenSource _autosaveCts = new();
    private bool _disposed;

    // Safe stderr logging: the parent process may have redirected our stderr
    // to a pipe whose read-end closes when the parent exits, so any
    // Console.Error.WriteLine after that point throws IOException.  Swallow
    // it silently — these are best-effort diagnostics, not critical output.
    private static void LogStderr(string message)
    {
        try { Console.Error.WriteLine(message); } catch (IOException) { }
    }

    // Valid idle-timeout range: 1s .. 24h. Anything outside falls back to
    // the 12min default. A value of "0" is rejected (would be an infinite-
    // busy spin on the watchdog task). Shared between the startup env-var
    // path (OFFICECLI_RESIDENT_IDLE_SECONDS) and the runtime
    // __set-idle-timeout__ RPC so both observe identical bounds.
    private const int MinIdleSeconds = 1;
    private const int MaxIdleSeconds = 86400;
    private static readonly TimeSpan DefaultIdleTimeout = TimeSpan.FromMinutes(12);

    // Initial idle timeout: env var (OFFICECLI_RESIDENT_IDLE_SECONDS) takes
    // precedence, tests/CI use this to exercise short timeouts in seconds.
    // Future "open file → auto-start resident" UX can tune how aggressively
    // the background process exits by starting the child with this env var.
    private static TimeSpan ResolveIdleTimeout()
    {
        var raw = Environment.GetEnvironmentVariable("OFFICECLI_RESIDENT_IDLE_SECONDS");
        if (!string.IsNullOrWhiteSpace(raw)
            && int.TryParse(raw, out var secs)
            && secs >= MinIdleSeconds && secs <= MaxIdleSeconds)
        {
            return TimeSpan.FromSeconds(secs);
        }
        return DefaultIdleTimeout;
    }

    // Flush policy: when a dirty in-memory DOM is written to disk so a
    // third-party tool that opens the file directly (python-docx, openpyxl,
    // Word, a renderer) sees recent edits without an explicit `save`/`close`.
    // One knob, four modes (see Core/ResidentFlushPolicy.cs):
    //   each — flush before every mutation command returns (deterministic)
    //   auto — idle-debounced, interval adapts to measured save cost (default):
    //          clamp(4 × EMA(save duration), 2s, 10s). Typical documents hug
    //          the 2s floor; a data-heavy workbook whose save takes seconds
    //          backs off automatically so background saves never eat more
    //          than ~25% of wall-clock in a busy session.
    //   <N>  — idle-debounced, fixed N seconds (the pre-adaptive behavior)
    //   off  — flush only on save/close/shutdown
    // Read from OFFICECLI_RESIDENT_FLUSH; the legacy
    // OFFICECLI_RESIDENT_IDLE_SAVE_SECONDS (integer seconds / 0 / "off") is
    // honored when the new variable is unset. Paired with
    // OFFICECLI_RESIDENT_IDLE_SECONDS (idle→close); idle modes are bounded
    // the same way.
    private static readonly ResidentFlushMode FlushMode;
    private static readonly TimeSpan FixedFlushInterval;
    static ResidentServer()
    {
        foreach (var env in new[] { "OFFICECLI_RESIDENT_FLUSH", "OFFICECLI_RESIDENT_IDLE_SAVE_SECONDS" })
        {
            if (ResidentFlushPolicy.TryParse(Environment.GetEnvironmentVariable(env),
                    MinIdleSeconds, MaxIdleSeconds, out var mode, out var fixedInterval))
            {
                FlushMode = mode;
                FixedFlushInterval = fixedInterval;
                return;
            }
        }
        FlushMode = ResidentFlushMode.Auto;
    }

    // Adaptive-interval state (auto mode): EMA of measured save durations and
    // the derived debounce, both stored as atomically-readable primitives
    // (CONSISTENCY(volatile-ticks): same pattern as _idleTimeoutTicks — the
    // watchdog task reads them concurrently with command-thread writes).
    // -1 bits = no sample yet. Writers all run under _commandLock.
    private long _saveEmaMillis = -1;
    private long _adaptiveIntervalTicks = ResidentFlushPolicy.MinAdaptiveInterval.Ticks;

    private TimeSpan CurrentAutosaveInterval => FlushMode == ResidentFlushMode.Fixed
        ? FixedFlushInterval
        : TimeSpan.FromTicks(Volatile.Read(ref _adaptiveIntervalTicks));

    // Fold one measured save duration into the adaptive debounce. Called after
    // every successful _handler.Save() regardless of mode — the sample is
    // cheap to record and keeps the estimate warm across mode-irrelevant
    // saves (explicit `save`, each-mode flushes).
    private void RecordSaveDuration(TimeSpan elapsed)
    {
        var prev = Volatile.Read(ref _saveEmaMillis);
        var ema = ResidentFlushPolicy.NextEmaSeconds(
            prev < 0 ? -1 : prev / 1000.0, elapsed.TotalSeconds);
        Volatile.Write(ref _saveEmaMillis, (long)(ema * 1000));
        Volatile.Write(ref _adaptiveIntervalTicks, ResidentFlushPolicy.IntervalForEma(ema).Ticks);
    }

    // Runtime upgrade path for the idle timeout. Called from the ping
    // handler when a new `__set-idle-timeout__` request arrives. Returns
    // false if the seconds value is out of range. On success, the
    // watchdog loop is immediately kicked via ResetIdleTimer() so the
    // new value takes effect on the next iteration — otherwise the
    // in-flight Task.Delay would keep honouring the old duration.
    private bool TrySetIdleTimeout(int seconds)
    {
        if (seconds < MinIdleSeconds || seconds > MaxIdleSeconds)
            return false;
        Volatile.Write(ref _idleTimeoutTicks, TimeSpan.FromSeconds(seconds).Ticks);
        ResetIdleTimer();
        return true;
    }
    // Shared shutdown Task so __close__ and Dispose coordinate on a single
    // ordered teardown: drain in-flight command → dispose handler → ack client.
    private readonly object _shutdownLock = new();
    private Task? _shutdownTask;
    // BUG-BT-R26-2: silent data loss when the resident-held file is unlinked
    // out from under us. OpenXML SDK keeps writing to the orphaned inode on
    // Dispose, so the on-disk path stays missing and the user loses every
    // edit made during the resident session. We can't reliably resurrect
    // the data (the inode may have been replaced), but we MUST escalate so
    // close/set/get stop reporting bogus success. Set during DoShutdownAsync
    // and surfaced through the __close__ ack.
    private volatile bool _shutdownFileMissing;

    // BUG-INTERVIEW-EDIT-R10-B: separate from _shutdownFileMissing — set when
    // the dispose call succeeded but the original path no longer exists.
    // Surfaced as a stderr warning on the close ack; does NOT flip exit code
    // (rename is a legitimate, non-failure scenario indistinguishable here).
    private volatile bool _shutdownFileVanishedAfterDispose;

    public string PipeName => _pipeName;

    // BUG-RESIDENT-EDITABLE-DEFAULT: ctor now defaults to editable=false so
    // a resident that only ever serves read-only commands (view/get/query/
    // raw/validate/dump) never writes the file back on shutdown. Mutation
    // commands (set/add/remove/move/swap/refresh/raw-set/add-part/batch)
    // promote the handler to editable=true on first use via
    // PromoteToEditable(); the promotion is sticky for the resident's
    // lifetime, matching the pre-existing reopen pattern used by
    // view-screenshot/page-count/refresh.
    public ResidentServer(string filePath, bool editable = false)
    {
        _filePath = Path.GetFullPath(filePath);
        _pipeName = GetPipeName(_filePath);
        _editable = editable;

        // Capture Console.Error during handler open so any warnings emitted
        // by the dump-reader / format-handler open path (which run before
        // any per-command stderr scope exists) are routed to the first
        // command's reply envelope. Without this, plugin-side notices like
        // "dump-reader produced no commands" disappear into the resident's
        // own unread stderr pipe.
        var startupErrSink = new StringWriter();
        var origErr = Console.Error;
        Console.SetError(startupErrSink);
        try { _handler = DocumentHandlerFactory.Open(_filePath, editable); }
        finally { Console.SetError(origErr); }
        var captured = startupErrSink.ToString().TrimEnd('\r', '\n');
        if (captured.Length > 0) _startupStderr = captured;
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

        // Start idle-autosave watchdog (flush-to-disk without shutdown). Skipped
        // when auto-flush is disabled (off) or handled inline per command (each).
        var autosaveTask = FlushMode is ResidentFlushMode.Off or ResidentFlushMode.Each
            ? Task.CompletedTask
            : RunAutosaveWatchdogAsync(pingToken);

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
                    LogStderr($"Resident error: {ex.Message}");
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
        try { await autosaveTask; } catch (OperationCanceledException) { }

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
        // Command activity also restarts the idle-autosave countdown so autosave
        // only fires when the session is genuinely idle, never mid-workflow.
        var oldAuto = Interlocked.Exchange(ref _autosaveCts, new CancellationTokenSource());
        oldAuto.Cancel();
    }

    private async Task RunIdleWatchdogAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                // Snapshot the current idle CTS and timeout on each loop
                // iteration: ResetIdleTimer() swaps _idleCts to restart
                // the wait, and TrySetIdleTimeout() mutates
                // _idleTimeoutTicks and calls ResetIdleTimer() so the new
                // duration is observed here on the very next pass.
                var idleCts = Volatile.Read(ref _idleCts);
                var currentTimeout = CurrentIdleTimeout;
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(idleCts.Token, token);
                await Task.Delay(currentTimeout, linked.Token);

                // The clock says idle — but "time since the last command
                // boundary" is not the same as "nothing is running now". A single
                // command that runs longer than the idle window (e.g. a large
                // batch) produces no boundary event for its whole duration, so the
                // countdown armed when it started expires while it is still in
                // flight. Level-check the in-flight count before committing to
                // shutdown: a command waiting on or holding the command lock keeps
                // _pendingClients > 0 for its entire duration, so re-arm and wait
                // again instead of tearing the resident down mid-command. Shutting
                // down here would cancel the command's in-progress reply (the
                // client then reports "not delivered") while the exit-time flush
                // still writes the applied result to disk — a false failure over a
                // silently persisted change. When the command finishes it resets
                // the timer, so the next expiry counts a full window from
                // completion and shuts down cleanly only when truly idle.
                if (Volatile.Read(ref _pendingClients) > 0)
                    continue;

                // Reached here = idle timeout elapsed without reset.
                // Kick off the ordered shutdown path instead of raw-
                // cancelling _mainCts / _pingCts, so the "ping liveness ⇔
                // file locked" invariant is preserved end-to-end: the
                // ping pipe stays alive until handler.Dispose() completes.
                LogStderr($"Resident idle for {currentTimeout.TotalMinutes} minutes, closing.");
                _ = ShutdownAsync();
                break;
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                // _idleCts was cancelled (timer reset), loop and wait again
            }
        }
    }

    // Idle-autosave watchdog: mirrors RunIdleWatchdogAsync but, instead of
    // shutting the resident down, flushes a dirty DOM to disk and keeps running.
    // Disabled (never started) when the flush policy is off/each.
    private async Task RunAutosaveWatchdogAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            try
            {
                var autoCts = Volatile.Read(ref _autosaveCts);
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(autoCts.Token, token);
                // Re-read every iteration: in auto mode RecordSaveDuration
                // adjusts the interval after each measured save. An in-flight
                // delay keeps the old value until the next pass (same accepted
                // lag as TrySetIdleTimeout's in-flight Task.Delay).
                await Task.Delay(CurrentAutosaveInterval, linked.Token);
                // Reached here = idle for the autosave interval without a reset.
                await TryAutosaveAsync(token);
            }
            catch (OperationCanceledException) when (!token.IsCancellationRequested)
            {
                // _autosaveCts was cancelled (command activity reset the timer);
                // loop and wait again.
            }
        }
    }

    // Flush a dirty DOM to disk without stopping the resident. Serialized
    // against commands via _commandLock so it never races a mutation. The
    // under-lock _mainCts check makes it safe against shutdown: DoShutdownAsync
    // cancels _mainCts BEFORE it drains the command lock, so if shutdown has
    // begun we observe the cancellation here and skip the save (the shutdown
    // Dispose flushes instead).
    private async Task TryAutosaveAsync(CancellationToken token)
    {
        if (!_dirty || !_editable || _disposed) return;
        SemaphoreSlim? acquired = null;
        try
        {
            if (!await _commandLock.WaitAsync(TimeSpan.FromSeconds(5), token))
                return; // a command is running; its exit resets the timer, retry next round
            acquired = _commandLock;

            if (_disposed || _mainCts.IsCancellationRequested) return;
            if (!_dirty || !_editable) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            // Idle flush runs while nobody waits, so the xlsx formula-cache sweep
            // may take a generous budget — provided it yields the moment a client
            // connects and starts waiting on _commandLock (an aborted sweep falls
            // back to fullCalcOnLoad exactly like budget exhaustion, so this only
            // trades sweep completeness for command latency, never correctness).
            var xlsx = _handler as ExcelHandler;
            if (xlsx != null)
            {
                xlsx.SweepBudgetOverride = TimeSpan.FromSeconds(60);
                xlsx.SweepYieldRequested = () => Volatile.Read(ref _pendingClients) > 0;
            }
            try { _handler.Save(); }
            finally
            {
                if (xlsx != null) { xlsx.SweepBudgetOverride = null; xlsx.SweepYieldRequested = null; }
            }
            sw.Stop();
            RecordSaveDuration(sw.Elapsed);
            // A sweep interrupted by a command leaves formula caches partially
            // refreshed on disk. Keeping _dirty=true makes the next idle window
            // save again — that re-runs the sweep (the handler's Modified gate
            // stays open for the whole session), so the file converges to fully
            // swept once an idle window survives uninterrupted. Budget-exhausted
            // sweeps do NOT retry (fullCalcOnLoad already covers them).
            if (xlsx?.LastSweepTruncatedByYield == true)
            {
                LogStderr($"Autosaved {Path.GetFileName(_filePath)} (formula sweep yielded to a command; will re-sweep next idle window).");
            }
            else
            {
                _dirty = false;
                LogStderr($"Autosaved {Path.GetFileName(_filePath)} (idle flush, resident still running).");
            }
        }
        catch (OperationCanceledException) { }
        catch (ObjectDisposedException) { /* shutdown disposed _mainCts/_commandLock under us */ }
        catch (Exception ex)
        {
            LogStderr($"Autosave error: {ex.Message}");
        }
        finally
        {
            if (acquired != null) { try { acquired.Release(); } catch { } }
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
                    LogStderr($"Ping responder error: {ex.Message}");
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

            if (request.Command == "__set-idle-timeout__")
            {
                // Runtime upgrade path: `open` sends this when it finds
                // a resident that `create` auto-started with a short
                // (60s) timeout, so long editing sessions honour the
                // 12min `open` contract. Served on the ping pipe (not
                // the main pipe) so it bypasses _commandLock and stays
                // responsive even while the main pipe is busy. Safe
                // because it only mutates _idleTimeoutTicks (Volatile)
                // and nudges _idleCts — both of which are already
                // concurrency-safe for the watchdog loop.
                var secs = request.GetIntArg("seconds") ?? 0;
                if (TrySetIdleTimeout(secs))
                {
                    var ok = MakeResponse(0, $"{secs}", "");
                    await WriteLineToPipeAsync(accepted, ok, token);
                }
                else
                {
                    var err = MakeResponse(1, "",
                        $"Invalid idle timeout: {secs}s (must be {MinIdleSeconds}..{MaxIdleSeconds})");
                    await WriteLineToPipeAsync(accepted, err, token);
                }
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
                    LogStderr($"Shutdown error during __close__: {ex.Message}");
                }

                // BUG-BT-R26-2: report shutdown-time data-loss to the
                // client so the close command exits non-zero instead of
                // confirming a save that didn't land on disk.
                // BUG-INTERVIEW-EDIT-R10-B: also surface the softer
                // "file vanished" warning (rename-or-deleted, can't tell)
                // through the close ack stderr without flipping exit code.
                string response;
                if (_shutdownFileMissing)
                    response = MakeResponse(1, "",
                        $"save failed during shutdown — data may be lost: {_filePath}");
                else if (_shutdownFileVanishedAfterDispose)
                    response = MakeResponse(0, "Closing resident.",
                        $"WARNING: the backing file was missing at its original path during save: {_filePath}. " +
                        "It was deleted or moved externally; your changes were rebuilt at that path. " +
                        "If you renamed/moved it, a separate copy also exists at the new location.");
                else
                    response = MakeResponse(0, "Closing resident.", "");
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
            LogStderr($"Ping handler error: {ex.Message}");
        }
        finally
        {
            try { await accepted.DisposeAsync(); } catch { }
        }
    }

    // Number of accepted business clients not yet finished. Autosave's formula
    // sweep polls this (via the handler's SweepYieldRequested hook) so a command
    // arriving mid-sweep interrupts it instead of waiting out the sweep budget.
    private int _pendingClients;

    private async Task HandleClientWithLockAsync(NamedPipeServerStream server, CancellationToken token)
    {
        Interlocked.Increment(ref _pendingClients);
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
            LogStderr($"Resident error: {ex.Message}");
        }
        finally
        {
            Interlocked.Decrement(ref _pendingClients);
            await server.DisposeAsync();
        }
    }

    private async Task HandleClientAsync(NamedPipeServerStream server, CancellationToken token)
    {
        string? requestLine;
        try
        {
            requestLine = await ReadLineFromPipeAsync(server, token);
        }
        catch (Exception ex)
        {
            // BUG-R41-F1 (exception path): binary input with a UTF-16 BOM (0xFF 0xFE)
            // causes StreamReader on macOS to switch to UTF-16 mode and then throw
            // DecoderFallbackException or IOException when the byte stream is malformed
            // for that encoding. Previously the exception propagated to
            // HandleClientWithLockAsync's catch block which only logged to stderr —
            // the client received 0 bytes (unexpected EOF).
            // Now we surface a clean error JSON so the client always gets a response.
            try
            {
                var errResp = MakeResponse(1, "", $"Error: Request read failed ({ex.GetType().Name}: binary or encoding mismatch)");
                await WriteLineToPipeAsync(server, errResp, token);
            }
            catch { /* best-effort: if we can't write error, just close cleanly */ }
            return;
        }

        // BUG-R41-F1 (null path): ReadLineAsync returns null when the client closes
        // the connection before sending a newline (e.g. sends 0xFF 0xFE without a
        // valid UTF-16 newline). Send an error response instead of silent close.
        if (requestLine == null)
        {
            var errorResponse = MakeResponse(1, "", "Error: Empty or unreadable request (possible binary data or encoding mismatch)");
            try { await WriteLineToPipeAsync(server, errorResponse, token); } catch { }
            return;
        }

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
                return MakeResponse(1, "", "Error: Invalid request");

            // Capture stdout/stderr (safe: _commandLock serializes all commands)
            var stdoutWriter = new StringWriter();
            var stderrWriter = new StringWriter();
            var origOut = Console.Out;
            var origErr = Console.Error;
            Console.SetOut(stdoutWriter);
            Console.SetError(stderrWriter);

            // Replay any stderr captured during the constructor's
            // DocumentHandlerFactory.Open so plugin-side warnings (e.g.
            // "dump-reader produced no commands") reach the user on the
            // first command's reply. One-shot drain: subsequent commands
            // see an empty buffer.
            if (_startupStderr is not null)
            {
                Console.Error.WriteLine(_startupStderr);
                _startupStderr = null;
            }

            try
            {
                ExecuteCommand(request);
                // each mode: flush before the response is written so the
                // command's success implies the change is on disk — the
                // deterministic barrier for callers whose next step is an
                // external reader. A Save failure here throws and surfaces as
                // this command's error (exit != 0), never a silent stale disk.
                // Inside a batch the per-item DeferSave still applies; this
                // single flush at command granularity keeps batch O(n).
                if (FlushMode == ResidentFlushMode.Each && _dirty && _editable)
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    _handler.Save();
                    sw.Stop();
                    _dirty = false;
                    RecordSaveDuration(sw.Elapsed);
                }
            }
            finally
            {
                Console.SetOut(origOut);
                Console.SetError(origErr);
            }

            var stdout = stdoutWriter.ToString().TrimEnd('\r', '\n');
            var stderr = stderrWriter.ToString().TrimEnd('\r', '\n');

            // BUG-R40-B10: batch failure rows print to stdout, not stderr,
            // so the generic stderr/UNSUPPORTED inspection below sees a
            // clean stderr and returns exit 0 even when every batch item
            // failed. ExecuteBatch records the verdict in
            // _lastBatchHadFailure; promote it to a non-zero exit here.
            var isBatch = request.Command.Equals("batch", StringComparison.OrdinalIgnoreCase);
            var batchFailure = isBatch && _lastBatchHadFailure;

            // BUG-INTERVIEW-EDIT-R2: batch text-mode envelope is written by
            // the client via Console.Write (not WriteLine) to avoid double-
            // newlining single-command output. Without a trailing '\n' on
            // batch payloads, the shell prompt collides with the last line.
            // Re-append exactly one terminator for batch responses.
            if (isBatch && stdout.Length > 0) stdout += "\n";
            // R7-bt-3: validate must surface a non-zero exit code on schema
            // errors so callers (CI / shell scripts) can detect failure
            // without parsing the report text. ExecuteValidate writes the
            // report to stderr and records the count in
            // _lastValidateErrorCount.
            var isValidate = request.Command.Equals("validate", StringComparison.OrdinalIgnoreCase);
            var validateFailure = isValidate && _lastValidateErrorCount > 0;
            if (isValidate) _lastValidateErrorCount = 0;

            if (request.Json)
            {
                // JSON mode: server builds the envelope so client just passes through
                var warnings = BuildWarnings(stderr);
                var isFailure = string.IsNullOrEmpty(stdout) && warnings is { Count: > 0 }
                    || stdout.StartsWith("No properties applied", StringComparison.Ordinal);
                // JSON Envelope contract: propagate judgment-command verdicts
                // (batch / validate) into envelope.success so it agrees with
                // exit code on line below. Without this the resident path
                // emitted success=true while exit code went to 1 — exactly the
                // mismatch the non-resident path was already broken for.
                var businessSuccess = !(batchFailure || validateFailure);
                // If the sub-command already produced a top-level envelope
                // (json object with a `success` field — set/add/get/dump all
                // do this via WrapEnvelope*), forward it unchanged. Rewrapping
                // here used to nest the inner envelope under `data`, producing
                // a {success, data:{success, data:...}} double envelope. Only
                // merge warnings/businessSuccess into the existing envelope
                // for judgment commands whose verdict the resident layer adds.
                string envelope;
                if (IsAlreadyWrappedEnvelope(stdout))
                {
                    envelope = MergeIntoExistingEnvelope(stdout, warnings, batchFailure || validateFailure);
                }
                else
                {
                    envelope = IsJson(stdout)
                        ? OutputFormatter.WrapEnvelope(stdout, warnings, success: businessSuccess)
                        : isFailure
                            ? OutputFormatter.WrapEnvelopeError(stdout, warnings)
                            : OutputFormatter.WrapEnvelopeText(stdout, warnings, success: businessSuccess);
                }
                // BUG-R11-03: JSON-mode exit code must match text mode. Previously
                // hard-coded to 0, which silently swallowed every error type
                // (path-not-found, unsupported_property, failed open) for any
                // resident --json call. Map parity with text mode below:
                //   - envelope success:false                        -> 1
                //   - stderr contains UNSUPPORTED (unsupported_property) -> 2
                //   - otherwise                                      -> 0
                // Batch/validate verdict failures OUTRANK applied-with-caveats
                // markers, mirroring the non-resident batch path — an
                // atomically rolled-back batch whose only green item carried a
                // LaTeX warning must not report exit 2, which would claim
                // something was applied when nothing was. Single-command
                // marker precedence (all-unsupported set → 2) is unchanged.
                int jsonExitCode = 0;
                if (batchFailure || validateFailure)
                    jsonExitCode = 1;
                else if (stderr.Contains("UNSUPPORTED") || stderr.Contains(UnrecognizedLatexMarker))
                    jsonExitCode = 2;
                else if (!EnvelopeSuccess(envelope) || stderr.Contains("VALIDATION:"))
                    jsonExitCode = 1;
                return MakeResponse(jsonExitCode, envelope, "");
            }

            // BUG-DUMP12-01: surface stderr "VALIDATION:" token (emitted by
            // ExecuteRawSet / ExecuteAddPart when the SDK validator gains new
            // errors) as exit 1 so callers can detect rejected raw mutations.
            // Batch/validate verdict failures outrank applied-with-caveats
            // markers (mirrors the non-resident batch path); single-command
            // marker precedence is unchanged.
            int exitCode = (batchFailure || validateFailure) ? 1
                : ((stderr.Contains("UNSUPPORTED") || stderr.Contains(UnrecognizedLatexMarker)) ? 2
                : (stderr.Contains("VALIDATION:") ? 1 : 0));
            return MakeResponse(exitCode, stdout, stderr);
        }
        catch (Exception ex)
        {
            // CONSISTENCY(error-wrap): mirror CommandBuilder.WriteError —
            // surface a friendlier message when an OOXML part is externally
            // corrupted, instead of the raw "Data at the root level is
            // invalid. Line 1, position 1." XmlException leak (fuzz-3, fuzz-4).
            var rendered = ex is System.Xml.XmlException xe
                ? new InvalidDataException(
                    $"Malformed XML in document part: {xe.Message} " +
                    $"(the file appears to have a corrupted OOXML part).", xe)
                : ex;
            if (request?.Json == true)
            {
                // JSON mode: wrap error in envelope
                return MakeResponse(1, OutputFormatter.WrapErrorEnvelope(rendered), "");
            }
            // BUG-R11-02: prefix the stderr string with the canonical
            // "Error: " marker so resident-mode error output matches the
            // non-resident CLI path (WriteError in Program.cs). Without
            // this, clients diffing stderr across modes would mis-detect
            // failures.
            return MakeResponse(1, "", $"Error: {OfficeCli.Core.MsysPathHint.AugmentMessage(rendered.Message)}");
        }
    }

    private static bool IsJson(string s)
    {
        var trimmed = s.AsSpan().TrimStart();
        return trimmed.Length > 0 && (trimmed[0] == '{' || trimmed[0] == '[');
    }

    // BUG-R11-03 helper: inspect envelope JSON for the "success" field so
    // resident JSON-mode exit codes track the envelope's actual success flag
    // instead of always returning 0.
    private static bool EnvelopeSuccess(string envelopeJson)
    {
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(envelopeJson);
            if (doc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Object)
                return true;
            if (!doc.RootElement.TryGetProperty("success", out var s))
                return true;
            if (s.ValueKind == System.Text.Json.JsonValueKind.False) return false;
            if (s.ValueKind == System.Text.Json.JsonValueKind.True) return true;
            return true;
        }
        catch
        {
            return true; // malformed — don't synthesize a failure
        }
    }

    // Detect an envelope already minted by CommandBuilder.* (set/add/get/dump
    // call WrapEnvelope*). Rewrapping such payloads nested them under `data`,
    // producing {success, data:{success, ...}}. Heuristic: top-level JSON
    // object with a boolean `success` field — same contract WrapEnvelope* emits
    // and that EnvelopeSuccess already keys off.
    private static bool IsAlreadyWrappedEnvelope(string s)
    {
        if (!IsJson(s)) return false;
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(s);
            if (doc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Object) return false;
            if (!doc.RootElement.TryGetProperty("success", out var succ)) return false;
            return succ.ValueKind == System.Text.Json.JsonValueKind.True
                || succ.ValueKind == System.Text.Json.JsonValueKind.False;
        }
        catch { return false; }
    }

    // Forward the inner envelope unchanged but: (a) flip success=false if the
    // resident layer detected a batch/validate verdict failure the sub-command
    // didn't surface, (b) append our stderr-derived warnings to the existing
    // warnings array. No restructure of `data`/`message`/`error`.
    private static string MergeIntoExistingEnvelope(string envelopeJson, List<CliWarning>? extraWarnings, bool forceFailure)
    {
        try
        {
            var node = System.Text.Json.Nodes.JsonNode.Parse(envelopeJson);
            if (node is not System.Text.Json.Nodes.JsonObject obj) return envelopeJson;
            if (forceFailure) obj["success"] = false;
            if (extraWarnings is { Count: > 0 })
            {
                if (obj["warnings"] is System.Text.Json.Nodes.JsonArray existing)
                {
                    foreach (var w in extraWarnings)
                        existing.Add(System.Text.Json.JsonSerializer.SerializeToNode(w, AppJsonContext.Default.CliWarning));
                }
                else
                {
                    obj["warnings"] = System.Text.Json.JsonSerializer.SerializeToNode(extraWarnings, AppJsonContext.Default.ListCliWarning);
                }
            }
            return obj.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        }
        catch
        {
            return envelopeJson;
        }
    }

    private static List<CliWarning>? BuildWarnings(string stderr)
    {
        if (string.IsNullOrEmpty(stderr)) return null;
        var lines = stderr.Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0) return null;
        return lines.Select(line =>
        {
            var warning = new CliWarning { Message = line.Trim() };
            if (line.Contains(UnrecognizedLatexMarker))
            {
                warning.Code = "unrecognized_latex_command";
                // Strip any "WARNING:" prefix so the envelope message matches the
                // one-shot CLI's form ("unrecognized_latex_command: \foo").
                var idx = line.IndexOf(UnrecognizedLatexMarker, StringComparison.Ordinal);
                warning.Message = line.Substring(idx).Trim();
            }
            else if (line.Contains("UNSUPPORTED")) warning.Code = "unsupported_property";
            else if (line.Contains("VALIDATION")) warning.Code = "validation_error";
            // CONSISTENCY(dump-warning-code): mirror the
            // unsupported_element code emitted by CommandBuilder.Dump.cs so
            // resident-routed and direct dump callers see the same envelope.
            else if (line.StartsWith("warning: skipped ", StringComparison.Ordinal)) warning.Code = "unsupported_element";
            // CONSISTENCY(numfmt-warning): mirror the invalid_number_format
            // code the one-shot --json path emits via WarningContext (see
            // ExcelStyleManager.GetOrCreateNumFmt), so resident-routed and
            // direct callers see the same envelope. Strip the "Warning: "
            // prefix to match the one-shot message form.
            else if (line.StartsWith("Warning: number format ", StringComparison.Ordinal))
            {
                warning.Code = "invalid_number_format";
                warning.Message = line.Substring("Warning: ".Length).Trim();
            }
            else warning.Code = "warning";
            return warning;
        }).ToList();
    }

    // BUG-RESIDENT-EDITABLE-DEFAULT: dispose+reopen the handler in writable
    // mode the first time a mutation command runs. The naive Dispose +
    // DocumentHandlerFactory.Open round-trip mirrors the existing reopen
    // sites (view-screenshot, page-count, refresh). It runs inside the
    // _commandLock-serialized dispatcher so no caller can hold a stale
    // _handler reference. Once promoted, _editable stays true for the
    // resident's lifetime — matches the previous "always-editable" cost
    // model so a long mixed-mode session does not pay the reopen cost
    // repeatedly. The final Dispose during shutdown still flushes the
    // in-memory tree, which is now the correct behavior because the user
    // has actually mutated the document.
    private void PromoteToEditable()
    {
        if (!_editable)
        {
            _handler.Dispose();
            _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, editable: true);
            _editable = true;
        }
        // Resident keeps the document in memory and flushes to disk only on
        // save/close/idle-autosave, so per-mutation Document.Save() (an O(n)
        // re-serialize each) is pure waste — and over a long editing session it
        // compounds to O(n²). Defer it for the resident's whole life. save/close
        // and the idle-autosave go through _handler.Save() directly, bypassing
        // the flag, so they still flush. This is
        // the shared prelude of every mutation command, so it also re-applies
        // after a handler reopen (screenshot / page-count / refresh reset
        // _handler). The non-resident batch path sets DeferSave on its own.
        if (_handler is OfficeCli.Handlers.WordHandler wh) wh.DeferSave = true;
        // Mark the in-memory DOM as having unflushed changes. Cleared by the
        // next save/close/idle-autosave. Set here (the shared mutation prelude)
        // so single commands and batch alike are tracked.
        _dirty = true;
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
                PromoteToEditable();
                ExecuteSet(request);
                NotifyWatchSlideChanged(request.GetArg("path"));
                break;
            case "add":
            {
                PromoteToEditable();
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
                PromoteToEditable();
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
                PromoteToEditable();
                ExecuteMove(request);
                NotifyWatchSlideChanged(request.GetArg("path"));
                break;
            case "swap":
                PromoteToEditable();
                ExecuteSwap(request);
                NotifyWatchFullRefresh();
                break;
            case "refresh":
                PromoteToEditable();
                ExecuteRefresh(request);
                NotifyWatchFullRefresh();
                break;
            case "raw":
                ExecuteRaw(request);
                break;
            case "raw-set":
                PromoteToEditable();
                ExecuteRawSet(request);
                NotifyWatchFullRefresh();
                break;
            case "add-part":
                PromoteToEditable();
                ExecuteAddPart(request);
                NotifyWatchFullRefresh();
                break;
            case "validate":
                ExecuteValidate();
                break;
            case "save":
                ExecuteSave();
                break;
            case "batch":
                // Promotion happens INSIDE ExecuteBatch: the atomic flush
                // barrier must read the pre-batch _dirty state, and
                // PromoteToEditable latches _dirty=true — promoting here would
                // make every batch look dirty and pay a full serialize.
                ExecuteBatch(request);
                break;
            case "dump":
                ExecuteDump(request, format);
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

    // BUG-R40-B10: track whether the most recent ExecuteBatch saw any
    // failures so ProcessRequest can surface a non-zero exit code.
    // Without this, a batch where every item fails returned exit 0 because
    // the wrapper at the bottom of ProcessRequest only inspected stderr
    // (and batch failure rows are written to stdout).
    private bool _lastBatchHadFailure;

    // Batch verbs that never mutate the DOM. ExecuteBatch notifies any live
    // watch session after applying the items — parity with the per-verb
    // NotifyWatch* calls in ExecuteCommand (issue #169). A batch made up
    // entirely of these read-only verbs skips the full-refresh to avoid a
    // needless re-render + SSE push, skips the editable promotion, and skips
    // the atomic machinery. Fail-open contract and the canonical verb list
    // live in CommandBuilder.ReadOnlyBatchVerbs (shared with the non-resident
    // atomic-copy decision).
    private static HashSet<string> ReadOnlyBatchVerbs => CommandBuilder.ReadOnlyBatchVerbs;

    private void ExecuteBatch(ResidentRequest request)
    {
        _lastBatchHadFailure = false;
        var batchJson = request.GetArg("batchJson");
        // BUG-R4-BT2: stopOnError is now an explicit arg from the client.
        // For older clients that only send "force", fall back to the legacy
        // semantics (force=true ⇒ continue-on-error). Newer clients always
        // send stopOnError so the legacy fallback never fires.
        var force = request.GetArg("force", "false")
            .Equals("true", StringComparison.OrdinalIgnoreCase);
        var hasStopArg = request.Args?.ContainsKey("stopOnError") == true;
        var stopOnError = hasStopArg
            ? request.GetArg("stopOnError", "false").Equals("true", StringComparison.OrdinalIgnoreCase)
            : !force;
        var json = request.Json;

        var items = System.Text.Json.JsonSerializer.Deserialize<List<BatchItem>>(
            batchJson, BatchJsonContext.Default.ListBatchItem) ?? new();

        // BUG-R40-B11: parity with the non-resident path —
        // CommandBuilder.Batch.cs already rejects null entries, but
        // resident invocations bypass that check (the batchJson is
        // forwarded raw), so re-validate here.
        for (int ni = 0; ni < items.Count; ni++)
        {
            if (items[ni] == null)
                throw new ArgumentException(
                    $"batch item[{ni}] is null. Each entry must be a JSON object (e.g. {{\"command\":\"get\",\"path\":\"/\"}}).");
        }

        // Document-protection gate, evaluated against the resident's in-memory
        // DOM (not the on-disk file, which may lag uncommitted protection
        // changes — the resident flushes only on save/close/idle-autosave).
        // Mirrors the
        // non-resident batch path; honors --force. Throw so the command's
        // exception handler maps it to a non-zero exit with the message.
        if (!force && _handler is OfficeCli.Handlers.WordHandler)
        {
            var protBlock = CommandBuilder.GetBatchProtectionBlock(_handler, items);
            if (protBlock != null)
                throw new CliException(protBlock) { Code = "document_protected" };
        }

        var bestEffort = request.GetArg("bestEffort", "false")
            .Equals("true", StringComparison.OrdinalIgnoreCase);
        var hasMutating = items.Any(it => !ReadOnlyBatchVerbs.Contains(it.Command ?? ""));
        var atomic = !bestEffort && hasMutating;

        // Atomic flush barrier: make the on-disk file identical to the
        // pre-batch in-memory tree, so a failed batch can roll back by simply
        // discarding the poisoned DOM and reloading. Reads the TRUE pre-batch
        // _dirty state (PromoteToEditable below latches it), so a batch on an
        // already-flushed session pays nothing — the serialize below only
        // happens when a flush was owed anyway, just earlier than the idle
        // debounce would have run it. Ordered before promotion: mutations
        // cannot exist while !_editable, so the not-yet-promoted case needs no
        // barrier.
        if (atomic && _editable && _dirty)
        {
            // FLUSH=off promises "disk writes only on explicit save/close/
            // shutdown" — the barrier's implicit Save would break that
            // contract, and skipping it would make a rollback reload lose
            // the unflushed pre-batch edits. Fail closed with the two ways
            // out instead of silently picking either.
            if (FlushMode == ResidentFlushMode.Off)
                throw new CliException(
                    "atomic batch needs the pre-batch state on disk as its rollback point, " +
                    "but OFFICECLI_RESIDENT_FLUSH=off is holding unflushed changes in memory. " +
                    "Run 'save' first, or use 'batch --best-effort'.")
                { Code = "flush_policy_conflict", Suggestion = "officecli save <file> before the batch, or batch --best-effort" };
            var swBarrier = System.Diagnostics.Stopwatch.StartNew();
            _handler.Save();
            swBarrier.Stop();
            _dirty = false;
            RecordSaveDuration(swBarrier.Elapsed);
        }
        if (hasMutating) PromoteToEditable();

        // Defer per-mutation Document.Save() across the whole batch so N resident
        // mutations serialize once (at the next save/close) instead of N times —
        // the per-op Save was an O(N²) re-serialize of the growing part. Mirrors
        // the non-resident batch path. get/query inside the batch still read the
        // live in-memory DOM, so they observe every just-added element; the
        // resident flushes to disk only on `save`/`close`/idle-autosave, which
        // go through _handler.Save() directly (bypassing the deferred SaveDoc()).
        //
        // Unlike the dispose-based CLI/MCP surfaces (which leave DeferSave on and
        // let Dispose finalize), the resident handler is long-lived: it must
        // restore the previous DeferSave and run ReconcileGlobalIds (below)
        // explicitly. The replay loop itself is shared via ApplyBatchItems;
        // skipResidentOnlyCommands drops in-batch open/close that would conflict
        // with the already-open file.
        var deferHandler = _handler as OfficeCli.Handlers.WordHandler;
        var prevDefer = deferHandler?.DeferSave ?? false;
        if (deferHandler != null) deferHandler.DeferSave = true;
        // Staged docProps whole-part payloads live outside the flush barrier
        // (they land on disk only at a non-discard Dispose), so a rollback
        // must restore this pre-batch snapshot into the replacement handler —
        // otherwise confirmed pre-batch raw-set edits vanish with the
        // poisoned DOM. Taken before the batch so batch-staged entries roll
        // back too.
        var preBatchWholeParts = atomic ? deferHandler?.SnapshotPendingWholeParts() : null;
        List<BatchResult> results;
        // BUG-BT2: collect per-item unrecognized-LaTeX tokens across the whole
        // batch so the resident surfaces the same unrecognized_latex_command
        // warning + exit 2 the one-shot path does (the handler resets
        // LastUnrecognizedLatex per item, so a post-loop read would only see
        // the last item's tokens).
        var batchUnrecognizedLatex = new List<string>();
        try
        {
            results = CommandBuilder.ApplyBatchItems(_handler, items, stopOnError, json,
                skipResidentOnlyCommands: true, unrecognizedLatex: batchUnrecognizedLatex);
        }
        finally
        {
            if (deferHandler != null) deferHandler.DeferSave = prevDefer;
        }

        // Atomic rollback: any failed item discards the WHOLE batch. The
        // barrier above guaranteed disk == pre-batch state, and nothing
        // flushed mid-batch (DeferSave + _commandLock keeps the autosave
        // watchdog out), so rolling back is: drop the poisoned in-memory DOM
        // without serializing it (DiscardOnDispose) and reload the pre-batch
        // file. The reload pays one parse — only on the failure path.
        var anyFailed = results.Any(r => !r.Success);
        var rolledBack = false;
        if (atomic && anyFailed)
        {
            switch (_handler)
            {
                case OfficeCli.Handlers.WordHandler w: w.DiscardOnDispose = true; break;
                case OfficeCli.Handlers.ExcelHandler x: x.DiscardOnDispose = true; break;
                case OfficeCli.Handlers.PowerPointHandler p: p.DiscardOnDispose = true; break;
            }
            try { _handler.Dispose(); } catch { /* discard path */ }
            _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, editable: true);
            // Re-establish the resident's long-lived handler invariants
            // (mirrors PromoteToEditable's post-open state): _editable stays
            // latched, deferral re-applies, and memory now equals disk.
            if (_handler is OfficeCli.Handlers.WordHandler wh2)
            {
                wh2.DeferSave = true;
                wh2.AdoptPendingWholeParts(preBatchWholeParts);
            }
            _dirty = false;
            rolledBack = true;
        }

        // BUG-R7B(BUG2): reconcile document-wide ids (wp:docPr, paraId, sdt)
        // after the deferred batch. Each item ran under DeferSave so the
        // per-raw-set id passes were skipped; a raw-set of a header/footer part
        // can leave colliding wp:docPr ids in the in-memory tree until the next
        // save/close. Run the same document-scoped passes here so a `validate`
        // (or watch render) issued before the eventual save sees the same clean
        // state save/close would write. Mirrors the non-resident path, where
        // Dispose-time FinalizeDeferredIds already does this. Skipped after a
        // rollback — the batch's changes no longer exist and deferHandler
        // points at the disposed pre-rollback handler.
        if (!rolledBack) deferHandler?.ReconcileGlobalIds();

        // Judgment contract: batch is classified as a judgment command (root
        // the project conventions "Judgment: any batch step failed -> outer false"). The
        // verdict flips to failure as soon as ANY step is rejected. Keeps
        // envelope.success / exit code in lockstep with the non-resident
        // path.
        _lastBatchHadFailure = anyFailed;
        CommandBuilder.PrintBatchResults(results, json, items.Count, atomicRolledBack: rolledBack);
        // BUG-BT2: emit the collected unrecognized-LaTeX markers so the
        // dispatcher maps them to exit 2 and the envelope warning code, exactly
        // as the single-shot resident add/set path (EmitUnrecognizedLatex) does.
        foreach (var tok in batchUnrecognizedLatex)
            Console.Error.WriteLine($"  WARNING: {UnrecognizedLatexMarker} {tok}");

        // Notify any live watch session so the preview reflects the batch.
        // Single add/set/move/remove/etc. each call a NotifyWatch* helper in
        // ExecuteCommand; batch applied its items directly above and must do
        // the same, otherwise the watched DOM stays stale until the watch is
        // manually restarted (issue #169). A batch can span many
        // slides/sheets/pages with mixed verbs, so a targeted per-slide patch
        // isn't derivable — a full refresh mirrors swap / refresh / raw-set /
        // add-part. Skip only provably read-only batches to avoid a needless
        // re-render; unknown verbs fail open to notify. An atomic rollback
        // also skips: the document is byte-identical to what the preview
        // already shows, so pushing a frame would be a lie about a change
        // that never landed.
        if (hasMutating && !rolledBack)
            NotifyWatchFullRefresh();
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
        if (!WatchServer.IsWatching(_filePath)) return;

        if (_handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            string? scrollTo = null;
            var sheetName = WatchMessage.ExtractSheetName(changedPath);
            if (sheetName != null)
            {
                var idx = excel.GetSheetIndex(sheetName);
                if (idx >= 0) scrollTo = $".sheet-content[data-sheet=\"{idx}\"]";
            }
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = CommandBuilder.RenderViaRegistry(excel, "xlsx", new OfficeCli.Core.Rendering.RenderOptions()), ScrollTo = scrollTo });
            return;
        }
        if (_handler is OfficeCli.Handlers.WordHandler word)
        {
            var scrollTo = WatchMessage.ExtractWordScrollTarget(changedPath);
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = CommandBuilder.RenderViaRegistry(word, "docx", new OfficeCli.Core.Rendering.RenderOptions()), ScrollTo = scrollTo });
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
        if (!WatchServer.IsWatching(_filePath)) return;

        if (_handler is OfficeCli.Handlers.WordHandler word)
        {
            var html = CommandBuilder.RenderViaRegistry(word, "docx", new OfficeCli.Core.Rendering.RenderOptions())!;
            var pageCount = System.Text.RegularExpressions.Regex.Matches(html, @"data-page=""\d+""").Count;
            var scrollTo = pageCount > 0 ? $".page[data-page=\"{pageCount}\"]" : null;
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = html, ScrollTo = scrollTo });
            return;
        }
        if (_handler is OfficeCli.Handlers.ExcelHandler excel)
        {
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = CommandBuilder.RenderViaRegistry(excel, "xlsx", new OfficeCli.Core.Rendering.RenderOptions()) });
            return;
        }
        if (_handler is not OfficeCli.Handlers.PowerPointHandler ppt) return;
        var newCount = ppt.GetSlideCount();
        if (newCount > oldSlideCount)
        {
            var html = ppt.RenderSlideHtml(newCount);
            if (html != null)
            {
                WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "add", Slide = newCount, Html = html, FullHtml = CommandBuilder.RenderViaRegistry(ppt, "pptx", new OfficeCli.Core.Rendering.RenderOptions()) });
                return;
            }
        }
        else if (newCount < oldSlideCount)
        {
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "remove", Slide = oldSlideCount, FullHtml = CommandBuilder.RenderViaRegistry(ppt, "pptx", new OfficeCli.Core.Rendering.RenderOptions()) });
            return;
        }
        WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = CommandBuilder.RenderViaRegistry(ppt, "pptx", new OfficeCli.Core.Rendering.RenderOptions()) });
    }

    private void NotifyWatchFullRefresh()
    {
        if (!WatchServer.IsWatching(_filePath)) return;

        string? fullHtml = null;
        if (_handler is OfficeCli.Handlers.PowerPointHandler ppt)
            fullHtml = CommandBuilder.RenderViaRegistry(ppt, "pptx", new OfficeCli.Core.Rendering.RenderOptions());
        else if (_handler is OfficeCli.Handlers.ExcelHandler excel)
            fullHtml = CommandBuilder.RenderViaRegistry(excel, "xlsx", new OfficeCli.Core.Rendering.RenderOptions());
        else if (_handler is OfficeCli.Handlers.WordHandler word)
            fullHtml = CommandBuilder.RenderViaRegistry(word, "docx", new OfficeCli.Core.Rendering.RenderOptions());
        if (fullHtml != null)
            WatchNotifier.NotifyIfWatching(_filePath, new WatchMessage { Action = "full", FullHtml = fullHtml });
    }

    private void ExecuteView(ResidentRequest req, OutputFormat format)
    {
        var mode = req.GetArg("mode", "text")!;
        var start = req.GetIntArg("start");
        var end = req.GetIntArg("end");
        var maxLines = req.GetIntArg("max-lines");
        // Resident parity with CLI front-end: reject typos / unknown subtypes
        // here too so a `--type foo` call doesn't get a silent count:0 just
        // because the request came over the pipe instead of through argv.
        // Validate also trims and normalises empty/whitespace to null.
        var issueType = OfficeCli.Core.IssueSubtypes.Validate(req.GetArgOrNull("type"));
        var limit = req.GetIntArg("limit");
        var cols = req.GetCols("cols");
        var pageFilter = req.GetArgOrNull("page");

        if (mode!.ToLowerInvariant() is "html" or "h")
        {
            string? html = null;
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptHandler)
            {
                // BUG-R36-B7: honor --page on pptx html with strict bounds.
                var (pStart, pEnd) = ResolvePptHtmlPage(pageFilter, start, end, pptHandler);
                html = CommandBuilder.RenderViaRegistry(_handler, "pptx",
                    new OfficeCli.Core.Rendering.RenderOptions { StartPage = pStart, EndPage = pEnd });
            }
            else if (_handler is OfficeCli.Handlers.ExcelHandler)
                html = CommandBuilder.RenderViaRegistry(_handler, "xlsx", new OfficeCli.Core.Rendering.RenderOptions());
            else if (_handler is OfficeCli.Handlers.WordHandler)
                html = CommandBuilder.RenderViaRegistry(_handler, "docx",
                    new OfficeCli.Core.Rendering.RenderOptions { PageFilter = pageFilter });
            else if (_handler is OfficeCli.Core.Plugins.FormatHandlerProxy proxy)
                html = proxy.ViewAsHtml(int.TryParse(pageFilter, out var pp) ? pp : (int?)null);

            if (html != null)
            {
                // CONSISTENCY(view-html-stdout): mirror CommandBuilder.View.cs — default
                // mode writes HTML to stdout so `officecli view file html > out.html`
                // captures actual content. --out writes to the requested path; --browser
                // writes a temp file (or --out, if given) and opens it.
                var browser = req.GetArgOrNull("browser");
                var wantBrowser = browser != null && (browser == "true" || browser == "1");
                var outArg = req.GetArgOrNull("out");
                if (outArg != null || wantBrowser)
                {
                    // SECURITY: when falling back to a temp file, include a random token so
                    // the preview path is not predictable. Without it, a predictable path
                    // enables a symlink pre-placement attack that causes File.WriteAllText
                    // to clobber an arbitrary victim file. See CommandBuilder.View.cs for
                    // the same fix.
                    var htmlPath = outArg ?? Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                    File.WriteAllText(htmlPath, html);
                    Console.WriteLine(Path.GetFullPath(htmlPath));
                    if (wantBrowser)
                    {
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
                    Console.Write(html);
                }
            }
            else
            {
                Console.Error.WriteLine("HTML preview is only supported for .pptx, .xlsx, and .docx files.");
            }
            return;
        }

        if (mode!.ToLowerInvariant() is "screenshot" or "p")
        {
            string? html = null;
            byte[]? directPng = null;
            var gridCols = req.GetIntArg("grid") ?? 0;
            var renderMode = (req.GetArgOrNull("render") ?? "auto").ToLowerInvariant();
            // --range clips a data-path region out of the HTML preview — mirrors
            // CommandBuilder.View.cs: native/direct-PNG backends are bypassed.
            var rangeArg = req.GetArgOrNull("range");
            if (rangeArg != null) renderMode = "html";
            var sw = req.GetIntArg("screenshot-width") ?? 1600;
            var sh = req.GetIntArg("screenshot-height") ?? 1200;
            // CONSISTENCY(screenshot-default-first-page): mirror CommandBuilder.View.cs —
            // screenshot mode defaults to a single bounded visual unit (pptx → slide 1,
            // docx → page 1, xlsx → active sheet via CSS). Without this, multi-page docs
            // render the full HTML stacked vertically and get silently cropped by the
            // viewport height. Caller can opt into more via --page / --grid.
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptShotHandler)
            {
                var effectiveFilter = pageFilter;
                if (rangeArg != null && string.IsNullOrEmpty(effectiveFilter)
                    && System.Text.RegularExpressions.Regex.Match(rangeArg, @"^/slide\[(\d+)\]") is { Success: true } slideM)
                    effectiveFilter = slideM.Groups[1].Value;
                if (string.IsNullOrEmpty(effectiveFilter) && start is null && end is null && gridCols == 0)
                    effectiveFilter = "1";
                var (pStart, pEnd) = ResolvePptHtmlPage(effectiveFilter, start, end, pptShotHandler);

                // Native-first (mirrors docx --render auto/native/html): export the
                // slide(s) to PNG with the OS-native engine on Windows; grid is HTML-only.
                // Default export size is the slide's 96-DPI native pixels; a custom
                // --screenshot-width overrides it (aspect-matched height).
                var (nativeW, nativeH) = pptShotHandler.GetSlideNativePixels();
                int exportW = nativeW, exportH = nativeH;
                if (!(sw == 1600 && sh == 1200))
                {
                    exportW = sw;
                    exportH = sh == 1200 ? Math.Max(1, (int)Math.Round(sw * (double)nativeH / nativeW)) : sh;
                }
                // -1 = auto: pick columns from slide count + aspect (mirrors CLI).
                // Resolved BEFORE the native block's dispose so GetSlideCount is safe,
                // and used by both the native and HTML paths.
                int pptGridCols = gridCols < 0
                    ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns((pEnd ?? pptShotHandler.GetSlideCount()) - (pStart ?? 1) + 1, nativeW, nativeH)
                    : gridCols;
                if (renderMode != "html" && OperatingSystem.IsWindows())
                {
                    // A read-only handler holds only read access with FileShare.ReadWrite,
                    // so the app can open the file for read concurrently — no dispose
                    // needed. An editable handler holds a write handle that blocks the
                    // app, so release it first (Dispose flushes pending edits) and reopen.
                    // Grid cell size + slide count are resolved BEFORE any dispose so the
                    // count read doesn't touch a disposed handler.
                    var ps = pStart ?? 1;
                    int gGap = 12, gPad = 12, gCellW = 0, gCellH = 0, gEnd = 0;
                    if (pptGridCols > 0)
                    {
                        gCellW = Math.Max(1, (int)Math.Round((sw - 2 * gPad - (pptGridCols - 1) * gGap) / (double)pptGridCols));
                        gCellH = Math.Max(1, (int)Math.Round(gCellW * (double)nativeH / nativeW));
                        gEnd = pEnd ?? pptShotHandler.GetSlideCount();
                    }
                    if (_editable) _handler.Dispose();
                    try
                    {
                        directPng = pptGridCols > 0
                            ? OfficeCli.Core.PowerPointPngBackend.RenderGrid(_filePath, ps, gEnd, gCellW, gCellH, pptGridCols, gGap, gPad)
                            : OfficeCli.Core.PowerPointPngBackend.Render(_filePath, ps, pEnd ?? ps, exportW, exportH);
                    }
                    catch { directPng = null; }
                    if (_editable)
                    {
                        _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, _editable);
                        pptShotHandler = (OfficeCli.Handlers.PowerPointHandler)_handler;
                    }
                }
                if (renderMode == "native" && directPng == null)
                {
                    Console.Error.WriteLine("--render native requires Windows with Microsoft PowerPoint installed.");
                    return;
                }
                if (directPng == null)
                {
                    html = CommandBuilder.RenderViaRegistry(pptShotHandler, "pptx", new OfficeCli.Core.Rendering.RenderOptions
                    { StartPage = pStart, EndPage = pEnd, GridColumns = pptGridCols, ViewportPx = sw })!;
                    // Single slide: size the viewport to the slide's 96-DPI native
                    // pixels so the PNG is the slide, padding-free.
                    if (pStart == pEnd && gridCols == 0)
                    {
                        if (sw == 1600 && sh == 1200) { sw = nativeW; sh = nativeH; }
                        else if (sh == 1200) sh = Math.Max(1, (int)Math.Round(sw * (double)nativeH / nativeW));
                    }
                }
            }
            else if (_handler is OfficeCli.Handlers.ExcelHandler excelShotHandler)
                html = CommandBuilder.RenderViaRegistry(excelShotHandler, "xlsx", new OfficeCli.Core.Rendering.RenderOptions())!;
            else if (_handler is OfficeCli.Handlers.WordHandler wordShotGrid && gridCols != 0)
            {
                // Contact-sheet grid — mirrors CommandBuilder.View.cs's docx grid
                // branch (native-first on Windows, HTML fallback; incl. -1 = auto).
                const int gGap = 12, gPad = 12, gMaxDim = 1920, gScrollbar = 17;
                var (gNpW, gNpH) = wordShotGrid.GetPageNativePixels();
                int gPageCount = 1;
                var gTmp = Path.Combine(Path.GetTempPath(), $"officecli_gridcount_{Path.GetFileNameWithoutExtension(_filePath)}_{Guid.NewGuid():N}.html");
                try
                {
                    File.WriteAllText(gTmp, CommandBuilder.RenderViaRegistry(wordShotGrid, "docx", new OfficeCli.Core.Rendering.RenderOptions())!);
                    gPageCount = OfficeCli.Core.HtmlScreenshot.GetPageCountFromDom(gTmp) ?? 1;
                }
                catch { /* fall back to 1 row */ }
                finally { try { File.Delete(gTmp); } catch { /* ignore */ } }

                int gCols = gridCols < 0 ? OfficeCli.Core.HtmlScreenshot.AutoGridColumns(gPageCount, gNpW, gNpH) : gridCols;
                int gRows = Math.Max(1, (gPageCount + gCols - 1) / gCols);
                double gVpW = sw;
                double gCellW = Math.Max(1.0, (gVpW - gScrollbar - gPad * 2.0 - (gCols - 1) * gGap) / gCols);
                double gCellH = gCellW * gNpH / gNpW;
                double gVpH = gPad * 2 + gRows * gCellH + (gRows - 1) * gGap;
                double gOver = Math.Max(gVpW, gVpH) / gMaxDim;
                if (gOver > 1.0) { gVpW /= gOver; gCellW /= gOver; gCellH /= gOver; gVpH /= gOver; }

                // Native-first on Windows: release an editable write lock (blocks
                // Word) before rendering, then reopen — same dance as the single-page
                // branch below.
                if (renderMode != "html" && OperatingSystem.IsWindows())
                {
                    if (_editable) _handler.Dispose();
                    try { directPng = OfficeCli.Core.WordPdfBackend.RenderGrid(_filePath, $"1-{gPageCount}", (int)Math.Round(gCellW), (int)Math.Round(gCellH), gCols, gGap, gPad); }
                    catch { directPng = null; }
                    if (_editable)
                    {
                        _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, _editable);
                        wordShotGrid = (OfficeCli.Handlers.WordHandler)_handler;
                    }
                }
                if (renderMode == "native" && directPng == null)
                {
                    Console.Error.WriteLine("--render native requires Windows with Microsoft Word installed.");
                    return;
                }
                if (directPng == null)
                {
                    html = CommandBuilder.RenderViaRegistry(wordShotGrid, "docx", new OfficeCli.Core.Rendering.RenderOptions
                    { GridColumns = gCols, GridCellWidthPx = (int)Math.Round(gCellW) })!;
                    sw = Math.Max(1, (int)Math.Round(gVpW));
                    sh = Math.Max(1, (int)Math.Ceiling(gVpH));
                }
            }
            else if (_handler is OfficeCli.Handlers.WordHandler wordShotHandler)
            {
                var effectiveFilter = rangeArg != null
                    ? pageFilter
                    : (string.IsNullOrEmpty(pageFilter) ? "1" : pageFilter);
                if (renderMode != "html" && OperatingSystem.IsWindows())
                {
                    // See the pptx branch: only an editable handler must be released
                    // (its write handle blocks Word); a read-only handler coexists.
                    if (_editable) _handler.Dispose();
                    // effectiveFilter is only null under --range, which forces
                    // renderMode=html — this native branch is then unreachable.
                    try { directPng = OfficeCli.Core.WordPdfBackend.Render(_filePath, effectiveFilter!); } catch { directPng = null; }
                    if (_editable)
                    {
                        _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, _editable);
                        wordShotHandler = (OfficeCli.Handlers.WordHandler)_handler;
                    }
                }
                if (renderMode == "native" && directPng == null)
                {
                    Console.Error.WriteLine("--render native requires Windows with Microsoft Word installed.");
                    return;
                }
                if (directPng == null) html = CommandBuilder.RenderViaRegistry(wordShotHandler, "docx",
                    new OfficeCli.Core.Rendering.RenderOptions { PageFilter = effectiveFilter })!;
            }
            if (html == null && directPng == null)
            {
                Console.Error.WriteLine("Screenshot mode is only supported for .pptx, .xlsx, and .docx files.");
                return;
            }
            var pngPath = req.GetArgOrNull("out") ?? Path.Combine(Path.GetTempPath(), $"officecli_screenshot_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.png");
            if (directPng != null)
            {
                File.WriteAllBytes(pngPath, directPng);
            }
            else
            {
                var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
                File.WriteAllText(tmpHtml, html!);
                var rs = rangeArg != null
                    ? OfficeCli.Core.HtmlScreenshot.CaptureClipped(tmpHtml, pngPath,
                        OfficeCli.Core.HtmlScreenshot.ResolveClipDataPaths(rangeArg))
                    : OfficeCli.Core.HtmlScreenshot.Capture(tmpHtml, pngPath, sw, sh);
                try { File.Delete(tmpHtml); } catch { /* ignore */ }
                if (!rs.Ok)
                {
                    Console.Error.WriteLine("No headless browser available. Install Chrome/Edge/Chromium or Firefox, or `pip install playwright && playwright install chromium`."
                        + (rs.Error != null ? $" Last error: {rs.Error}" : ""));
                    return;
                }
            }
            Console.WriteLine(Path.GetFullPath(pngPath));
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptCnt)
                Console.Error.WriteLine($"[pages] total={pptCnt.GetSlideCount()}");
            if (req.GetArgOrNull("browser") == "true")
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo(pngPath) { UseShellExecute = true };
                    System.Diagnostics.Process.Start(psi);
                }
                catch { /* silently ignore */ }
            }
            return;
        }

        if (mode!.ToLowerInvariant() is "svg" or "g")
        {
            if (_handler is OfficeCli.Handlers.PowerPointHandler pptSvgHandler)
            {
                // CONSISTENCY(view-page): SVG mode honors --page like html mode; --page wins over --start.
                int slideNum = 1;
                if (!string.IsNullOrEmpty(pageFilter))
                {
                    var firstTok = pageFilter.Split(',')[0].Split('-')[0].Trim();
                    // CONSISTENCY(strict-page): mirror CommandBuilder.View.cs
                    // — reject non-positive --page values rather than
                    // silently rendering slide 1.
                    if (!int.TryParse(firstTok, out var p))
                        throw new ArgumentException(
                            $"Invalid --page value '{pageFilter}': expected a positive slide number.");
                    if (p <= 0)
                        throw new ArgumentException(
                            $"Invalid --page value '{pageFilter}': slide number must be >= 1.");
                    slideNum = p;
                }
                else if (start.HasValue && start.Value > 0)
                {
                    slideNum = start.Value;
                }
                var svg = CommandBuilder.RenderViaRegistry(pptSvgHandler, "pptx",
                    new OfficeCli.Core.Rendering.RenderOptions
                    { Output = OfficeCli.Core.Rendering.RenderOutputKind.Svg, StartPage = slideNum })!;
                Console.Write(svg);
            }
            else if (_handler is OfficeCli.Core.Plugins.FormatHandlerProxy svgProxy)
            {
                int? svgPage = null;
                if (!string.IsNullOrEmpty(pageFilter)
                    && int.TryParse(pageFilter.Split(',')[0].Split('-')[0].Trim(), out var sp))
                    svgPage = sp;
                var svg = svgProxy.ViewAsSvg(svgPage);
                if (svg is null)
                    Console.Error.WriteLine("SVG preview is not supported by the format-handler plugin.");
                else
                    Console.Write(svg);
            }
            else
            {
                Console.Error.WriteLine("SVG preview is only supported for .pptx files.");
            }
            return;
        }

        int? pageCountValue = null;
        if (req.GetArgOrNull("page-count") == "true" && (mode!.ToLowerInvariant() is "stats" or "s") && _handler is OfficeCli.Handlers.WordHandler whForCount)
        {
            if (OperatingSystem.IsWindows())
            {
                _handler.Dispose();
                try { pageCountValue = OfficeCli.Core.WordPdfBackend.GetPageCount(_filePath); } catch { pageCountValue = null; }
                _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, _editable);
                whForCount = (OfficeCli.Handlers.WordHandler)_handler;
            }
            if (pageCountValue == null)
            {
                var tmpHtml = Path.Combine(Path.GetTempPath(), $"officecli_pc_{Path.GetFileNameWithoutExtension(_filePath)}_{Guid.NewGuid():N}.html");
                try
                {
                    File.WriteAllText(tmpHtml, CommandBuilder.RenderViaRegistry(whForCount, "docx", new OfficeCli.Core.Rendering.RenderOptions())!);
                    pageCountValue = OfficeCli.Core.HtmlScreenshot.GetPageCountFromDom(tmpHtml);
                }
                finally { try { File.Delete(tmpHtml); } catch { } }
            }
            if (pageCountValue == null)
            {
                Console.Error.WriteLine("--page-count: failed to get page count (Word backend and HTML fallback both unavailable).");
                return;
            }
        }

        if (req.Json)
        {
            var modeKey = mode!.ToLowerInvariant();
            if (modeKey is "stats" or "s")
            {
                var statsJson = _handler.ViewAsStatsJson();
                if (pageCountValue.HasValue) statsJson["pages"] = pageCountValue.Value;
                Console.WriteLine(statsJson.ToJsonString(OutputFormatter.PublicJsonOptions));
            }
            else if (modeKey is "outline" or "o")
                Console.WriteLine(_handler.ViewAsOutlineJson().ToJsonString(OutputFormatter.PublicJsonOptions));
            else if (modeKey is "text" or "t")
                Console.WriteLine(_handler.ViewAsTextJson(start, end, maxLines, cols, req.GetArgOrNull("range")).ToJsonString(OutputFormatter.PublicJsonOptions));
            else if (modeKey is "annotated" or "a")
                Console.WriteLine(OutputFormatter.FormatView(mode, _handler.ViewAsAnnotated(start, end, maxLines, cols), format));
            else if (modeKey is "issues" or "i")
                Console.WriteLine(OutputFormatter.FormatIssues(_handler.ViewAsIssues(issueType, limit), format));
            else if (modeKey is "forms" or "f")
            {
                if (_handler is OfficeCli.Handlers.WordHandler wordFormsHandler)
                    Console.WriteLine(wordFormsHandler.ViewAsFormsJson().ToJsonString(OutputFormatter.PublicJsonOptions));
                else if (_handler is OfficeCli.Core.Plugins.FormatHandlerProxy formsProxy)
                {
                    var formsJson = formsProxy.ViewAsFormsJson();
                    if (formsJson is null)
                        Console.Error.WriteLine($"Forms view is not supported by the format-handler plugin.");
                    else
                        Console.WriteLine(formsJson.ToJsonString(OutputFormatter.PublicJsonOptions));
                }
                else
                    Console.Error.WriteLine("Forms view is only supported for .docx files.");
            }
            else
                // Unknown mode is a caller bug, not a successful view: throw
                // the same CliException CommandBuilder.View.cs raises in the
                // direct-mode path so the envelope reports success=false /
                // code=invalid_value instead of stdout-ing the error message
                // as the view payload.
                throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, forms")
                {
                    Code = "invalid_value",
                    ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                };
        }
        else
        {
            var modeKey = mode!.ToLowerInvariant();
            string output;
            switch (modeKey)
            {
                case "text" or "t":
                    output = _handler.ViewAsText(start, end, maxLines, cols, req.GetArgOrNull("range")); break;
                case "annotated" or "a":
                    output = _handler.ViewAsAnnotated(start, end, maxLines, cols); break;
                case "outline" or "o":
                    output = _handler.ViewAsOutline(); break;
                case "stats" or "s":
                    output = pageCountValue.HasValue
                        ? $"Pages: {pageCountValue}\n" + _handler.ViewAsStats()
                        : _handler.ViewAsStats(); break;
                case "issues" or "i":
                    output = OutputFormatter.FormatIssues(_handler.ViewAsIssues(issueType, limit), format); break;
                case "forms" or "f":
                    output = _handler switch
                    {
                        OfficeCli.Handlers.WordHandler wfh => wfh.ViewAsForms(),
                        OfficeCli.Core.Plugins.FormatHandlerProxy fp
                            => fp.ViewAsFormsJson()?.ToJsonString(OutputFormatter.PublicJsonOptions)
                               ?? throw new OfficeCli.Core.CliException("Forms view is not supported by the format-handler plugin.")
                                   { Code = "unsupported_type" },
                        _ => throw new OfficeCli.Core.CliException("Forms view is only supported for .docx files.")
                        {
                            Code = "unsupported_type",
                            ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                        }
                    };
                    break;
                default:
                    throw new OfficeCli.Core.CliException($"Unknown mode: {mode}. Available: text, annotated, outline, stats, issues, html, svg, screenshot, forms")
                    {
                        Code = "invalid_value",
                        ValidValues = ["text", "annotated", "outline", "stats", "issues", "html", "svg", "screenshot", "pdf", "forms"]
                    };
            }
            Console.WriteLine(output);
        }
    }

    private void ExecuteGet(ResidentRequest req, OutputFormat format)
    {
        var path = req.GetArg("path", "/");
        var depth = req.GetIntArg("depth") ?? 1;
        var node = _handler.Get(path, depth);

        // CONSISTENCY(get-not-found-exit): mirror CommandBuilder.GetQuery.cs.
        // Some handler Get paths surface "not found" via Type="error" rather
        // than throwing. Convert to a real exception so the resident response
        // exits non-zero, matching the direct-mode CLI behavior.
        if (string.Equals(node.Type, "error", StringComparison.Ordinal))
            throw new ArgumentException(node.Text ?? $"Path not found: {path}");

        // CONSISTENCY(get-save): mirror CommandBuilder.GetQuery.cs lines 59-74.
        // Direct-mode `get --save` extracts the binary payload backing an
        // ole/picture/media node to disk. Resident mode must honour the same
        // arg or it silently drops the extraction (BUG-R9-01).
        var savePath = req.GetArgOrNull("save");
        if (!string.IsNullOrEmpty(savePath))
        {
            if (!_handler.TryExtractBinary(path, savePath, out var contentType, out var byteCount))
                throw new InvalidOperationException(
                    $"Node at '{path}' has no binary payload to extract (only ole/picture/media/embedded nodes can be saved).");
            node.Format["savedTo"] = savePath;
            node.Format["savedBytes"] = byteCount;
            if (!string.IsNullOrEmpty(contentType))
                node.Format["savedContentType"] = contentType!;
        }

        // Unified envelope contract: mirror direct-mode CommandBuilder.GetQuery.cs
        // — single-path get JSON returns {matches, results: [node]}, so agents
        // and scripts use one jq path across get / get selected / query. Text
        // mode keeps the rich single-node rendering via FormatNode.
        if (format == OutputFormat.Json)
            Console.WriteLine(OutputFormatter.FormatNodes(new List<DocumentNode> { node }, format));
        else
            Console.WriteLine(OutputFormatter.FormatNode(node, format));
    }

    // BUG-DUMP-R6-01: dump used to bypass the resident and open its own
    // WordHandler, which collided with the resident's lock. Mirror the
    // direct-mode CommandBuilder.Dump.cs flow against `_handler`.
    private void ExecuteDump(ResidentRequest req, OutputFormat format)
    {
        var path = req.GetArg("path", "/");
        var dumpFormat = req.GetArg("format", "batch")!.ToLowerInvariant();
        var outPath = req.GetArgOrNull("out");

        if (dumpFormat != "batch")
            throw new CliException($"Unsupported --format: {dumpFormat}. Valid: batch")
                { Code = "invalid_format", ValidValues = ["batch"] };

        // CONSISTENCY(dump-text-clean-output): warnings reach stderr in --json
        // mode OR when --out diverts the JSON array to a file (stdout then
        // carries just the path, so the `dump 2>&1 | batch --input -` pipe
        // hazard doesn't apply). Mirrors CommandBuilder.Dump.cs's warnToStderr.
        var outIsFile = !string.IsNullOrEmpty(outPath) && outPath != "-";
        var warnToStderr = req.Json || outIsFile;

        // CONSISTENCY(dump-format-dispatch): mirrors docx vs pptx branch
        List<BatchItem> items;
        if (_handler is WordHandler word)
        {
            var (wItems, wWarnings) = WordBatchEmitter.EmitWordWithWarnings(word, path);
            items = wItems;
            // CONSISTENCY(dump-text-clean-output): see warnToStderr above —
            // stderr in --json mode or when --out diverts the JSON to a file.
            if (warnToStderr)
            {
                foreach (var w in wWarnings)
                    Console.Error.WriteLine($"warning: skipped {w.Element} at {w.Path}: {w.Reason}");
            }
        }
        else if (_handler is OfficeCli.Handlers.PowerPointHandler ppt)
        {
            var (pItems, pWarnings) = OfficeCli.Handlers.PptxBatchEmitter.EmitPptx(ppt, path);
            items = pItems;
            // CONSISTENCY(dump-text-clean-output): see docx branch above.
            if (warnToStderr)
            {
                foreach (var w in pWarnings)
                    Console.Error.WriteLine($"warning: skipped {w.Element} on {w.SlidePath}: {w.Reason}");
            }
        }
        else if (_handler is OfficeCli.Handlers.ExcelHandler xl)
        {
            var (xItems, xWarnings) = OfficeCli.Handlers.ExcelBatchEmitter.EmitExcel(xl, path);
            items = xItems;
            // CONSISTENCY(dump-text-clean-output): see docx branch above.
            if (warnToStderr)
            {
                foreach (var w in xWarnings)
                    Console.Error.WriteLine($"warning: skipped {w.Element} at {w.Path}: {w.Reason}");
            }
        }
        else
        {
            throw new CliException("dump currently supports .docx, .pptx and .xlsx only")
                { Code = "unsupported_format" };
        }
        var output = System.Text.Json.JsonSerializer.Serialize(items, BatchJsonContext.Default.ListBatchItem);

        if (outPath == "-") outPath = null;
        if (!string.IsNullOrEmpty(outPath))
        {
            // CONSISTENCY(trailing-newline): match CommandBuilder.Dump.cs —
            // stdout path ends with a newline (Console.WriteLine), file path
            // must too so consumers see the same payload either way.
            File.WriteAllText(outPath, output + "\n");
            if (req.Json)
            {
                var meta = new System.Text.Json.Nodes.JsonObject
                {
                    ["outputFile"] = outPath,
                    ["itemCount"] = items.Count
                };
                Console.WriteLine(meta.ToJsonString());
            }
            else
                Console.WriteLine(outPath);
        }
        else
        {
            // A dump streamed back to stdout rides the resident pipe as a single
            // message. Past the pipe ceiling it cannot be delivered, so instead of
            // letting the transport fail with an opaque error, steer the caller to
            // the on-disk path: with --out the resident writes the file directly
            // and only the path crosses the pipe. (Reuses the pipe's own ceiling so
            // the two limits never drift apart.)
            var bytes = Encoding.UTF8.GetByteCount(output);
            if (bytes > MaxMessageLength)
                throw new CliException(
                    $"dump produced {bytes / (1024 * 1024)} MB of output, too large to return over the resident pipe. " +
                    "Re-run with --out <file> to write it directly to disk (only the file path is returned).")
                    { Code = "output_too_large" };
            Console.WriteLine(output);
        }
    }

    private void ExecuteQuery(ResidentRequest req, OutputFormat format)
    {
        // `path` aliases `selector` (batch items routed here carry whichever
        // field the caller wrote); an empty selector would silently match
        // every node — mirror CommandBuilder's batch-query guard.
        var selector = req.GetArgOrNull("selector") ?? req.GetArgOrNull("path") ?? "";
        if (string.IsNullOrEmpty(selector))
            throw new ArgumentException("'query' requires a selector. Example: {\"command\": \"query\", \"selector\": \"row[Score>80]\"}");
        // CONSISTENCY(cell-selector-alias): mirror the direct-mode normalization +
        // boolean engine in CommandBuilder.GetQuery.cs — without alias normalization,
        // resident-mode Excel cell queries with short aliases (bold, size, ...)
        // silently drop every hit (BUG-R17-01). FilterSelector also runs the and/or
        // engine and keeps pure-AND selectors on the exact legacy path.
        Func<string, string>? keyResolver =
            _handler is ExcelHandler && ExcelHandler.SelectorTargetsCells(selector)
                ? ExcelHandler.ResolveCellAttributeAlias : null;
        var (results, warnings) = AttributeFilter.FilterSelector(selector, _handler.Query, keyResolver);
        var textFilter = req.GetArgOrNull("find");
        if (!string.IsNullOrEmpty(textFilter))
            results = results.Where(n => n.Text != null && AttributeFilter.MatchesTextFilter(n.Text, textFilter)).ToList();
        // --compact: same line renderer as direct mode (CommandBuilder owns the
        // format contract); short-circuits before JSON hydration.
        if (req.GetArgOrNull("compact") == "true")
        {
            foreach (var w2 in warnings) Console.Error.WriteLine(w2.Message);
            Console.WriteLine(CommandBuilder.FormatNodesCompact(_handler, results, req.GetArgOrNull("fields")));
            return;
        }
        // CONSISTENCY(query-json-children): hydrate Children from Get(path, depth=1)
        // for JSON output so consumers see the same shape as `get --json`. Mirrors
        // the post-processing in CommandBuilder.GetQuery.cs.
        if (format == OutputFormat.Json)
        {
            foreach (var n in results)
            {
                if (n.ChildCount > 0 && n.Children.Count == 0 && !string.IsNullOrEmpty(n.Path))
                {
                    try
                    {
                        var hydrated = _handler.Get(n.Path, depth: 1);
                        if (hydrated?.Children != null && hydrated.Children.Count > 0)
                            n.Children.AddRange(hydrated.Children);
                    }
                    catch { /* path may not be Get-resolvable; leave empty */ }
                }
            }
        }
        foreach (var w in warnings) Console.Error.WriteLine(w.Message);
        Console.WriteLine(OutputFormatter.FormatNodes(results, format));
    }

    private void ExecuteSet(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var properties = req.GetProps();

        // CONSISTENCY(selector-set): mirrored in CommandBuilder.Set.cs. A no-slash
        // path is a Query→Set-per-match selector (Sheet1!A1, Sheet1!row[工资>5000]),
        // the same engine batch uses — accepted so resident `set` matches get/query.
        // The handler selector branch throws on an empty match, so no silent no-op.
        // Agent-safety: reject a bare unscoped selector (mirrors CommandBuilder).
        OfficeCli.Core.MutationSelectorGuard.EnsureScoped(path, "set");
        // Shared core: apply + prop-autocorrect + categorise — one copy across
        // CLI / batch / MCP / resident (CommandBuilder.ApplySetWithCorrection).
        // This also gives the resident the distance-1 typo auto-correction it
        // previously lacked, so `set colot=color` behaves the same whether or
        // not a resident is alive. The resident's own envelope (find-count,
        // selector-count, watch, overflow, --json wrapping) stays below.
        // CONSISTENCY(applied-echo): mirrors CommandBuilder.Set.cs — pre/post
        // Format snapshots feed the " (applied: ...)" normalization echo.
        var beforeSnap = CommandBuilder.TryGetFormatSnapshot(_handler, path);
        var (applied, unsupported, autoCorrected) =
            CommandBuilder.ApplySetWithCorrection(_handler, path, properties);

        // CONSISTENCY(find-match-count): mirrored in CommandBuilder.Set.cs.
        // The resident path is hit whenever a resident process is open
        // (which `create` does by default), so both sites must surface
        // findMatchCount + zero_matches warning identically.
        int? findMatchCount = null;
        if (properties.ContainsKey("find"))
        {
            findMatchCount = _handler switch
            {
                WordHandler wh => wh.LastFindMatchCount,
                PowerPointHandler ph => ph.LastFindMatchCount,
                ExcelHandler eh => eh.LastFindMatchCount,
                _ => null
            };
        }

        // CONSISTENCY(selector-set): echo how many elements a multi-match selector
        // set touched, so a Sheet1!row[工资>5000]-style mutation makes its scope
        // visible. Single-target paths (count 1) stay quiet.
        int? selectorCount = null;
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
            selectorCount = _handler switch
            {
                WordHandler wh => wh.LastSelectorSetCount,
                PowerPointHandler ph => ph.LastSelectorSetCount,
                ExcelHandler eh => eh.LastSelectorSetCount,
                _ => null
            };

        // R4-bt-1: report the post-move resolvable path for equation mode
        // switches (oMathPara ⇄ oMath), consistent with the non-resident path.
        var reportPath = (_handler as WordHandler)?.LastSetNewPath ?? path;
        var appliedSuffix = CommandBuilder.BuildAppliedSuffix(applied,
            beforeSnap, CommandBuilder.TryGetFormatSnapshot(_handler, reportPath));
        var message = applied.Count > 0
            ? $"Updated {reportPath}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
              + appliedSuffix
              + (findMatchCount.HasValue ? $" ({findMatchCount.Value} matched)" : "")
              + (selectorCount > 1 ? $" ({selectorCount} elements matched)" : "")
            : (unsupported.Count > 0 ? $"No properties applied to {path}" : $"Updated {path}");

        var warnings = new List<OfficeCli.Core.CliWarning>();
        foreach (var ac in autoCorrected)
            warnings.Add(new OfficeCli.Core.CliWarning
            {
                Message = $"Auto-corrected '{ac.Original}' to '{ac.Corrected}'",
                Code = "auto_corrected",
            });
        if (findMatchCount is 0)
        {
            warnings.Add(new OfficeCli.Core.CliWarning
            {
                Message = $"find pattern matched 0 occurrences at {path} — original text may have been edited or the path is wrong",
                Code = "zero_matches",
                Suggestion = "verify the path still resolves and the find text is current"
            });
        }
        var overflow = CommandBuilder.CheckTextOverflow(_handler, path);
        if (overflow != null)
        {
            warnings.Add(new OfficeCli.Core.CliWarning
            {
                Message = overflow,
                Code = "text_overflow",
                Suggestion = "Increase shape height/width, reduce font size, or shorten text"
            });
        }
        if (_handler is WordHandler setWhWarn)
        {
            foreach (var w in setWhWarn.LastSetWarnings)
                warnings.Add(new OfficeCli.Core.CliWarning { Message = w, Code = "advisory" });
        }

        if (req.Json)
        {
            bool allFailed = applied.Count == 0 && unsupported.Count > 0;
            Console.WriteLine(allFailed
                ? OutputFormatter.WrapEnvelopeError(message, warnings.Count > 0 ? warnings : null)
                : OutputFormatter.WrapEnvelopeText(message, warnings.Count > 0 ? warnings : null, findMatchCount));
        }
        else
        {
            if (applied.Count > 0 || unsupported.Count > 0) Console.WriteLine(message);
            foreach (var ac in autoCorrected)
                Console.Error.WriteLine($"WARNING: Auto-corrected '{ac.Original}' to '{ac.Corrected}'");
            if (findMatchCount is 0)
                Console.Error.WriteLine($"WARNING: find pattern matched 0 occurrences at {path}");
            if (overflow != null)
                Console.Error.WriteLine($"  WARNING: {overflow}");
            if (_handler is WordHandler setWhWarnPlain)
            {
                foreach (var w in setWhWarnPlain.LastSetWarnings)
                    Console.Error.WriteLine($"  WARNING: {w}");
            }
        }

        // Unrecognized LaTeX from an equation Set (formula=). Same emission as
        // ExecuteAdd — distinctive stderr line drives exit 2 + envelope code.
        EmitUnrecognizedLatex(_handler);

        if (unsupported.Count > 0)
        {
            // /styles/<id> on Word: targeted curated hints, no raw-set push.
            // See StyleUnsupportedHints + matching branch in CommandBuilder.
            if (_handler is WordHandler
                && path.StartsWith("/styles/", StringComparison.Ordinal))
            {
                var styleHint = OfficeCli.Core.StyleUnsupportedHints.Format(unsupported);
                if (styleHint != null) Console.Error.WriteLine(styleHint);
            }
            else
            {
                Console.Error.WriteLine(FormatUnsupportedLine(unsupported));
            }
        }
    }

    // When the handler already attached a parenthetical hint naming the valid
    // alternative props (e.g. "fill (valid slide props: background, ...)"), the
    // generic "use raw-set instead" prefix misdirects the user away from the
    // real fix. Drop the prefix in that case and let the handler hint stand.
    // Stderr marker for unrecognized LaTeX commands/environments. Doubles as
    // the user-facing warning text (same string the one-shot CLI emits) AND
    // the token the dispatcher keys off (alongside UNSUPPORTED) to set exit 2
    // and that BuildWarnings maps to the unrecognized_latex_command envelope
    // code — keeping the resident path's UX identical to CommandBuilder.Add/Set.
    internal const string UnrecognizedLatexMarker = "unrecognized_latex_command:";

    private static void EmitUnrecognizedLatex(IDocumentHandler handler)
    {
        var tokens = handler switch
        {
            WordHandler wlx => wlx.LastUnrecognizedLatex,
            PowerPointHandler plx => plx.LastUnrecognizedLatex,
            _ => null,
        };
        if (tokens is not { Count: > 0 }) return;
        foreach (var tok in tokens)
            Console.Error.WriteLine($"  WARNING: {UnrecognizedLatexMarker} {tok}");
    }

    private static string FormatUnsupportedLine(List<string> unsupported, string? scope = null)
    {
        bool hasNamedAlternative = unsupported.Any(u => u.Contains("(valid ", StringComparison.Ordinal));
        // Attach per-key "did you mean" suggestions (mirrors the non-resident
        // CommandBuilder.FormatUnsupported path) so guidance like the 1-based
        // cell-key hint for a 0-based r0c0 reaches resident-mode users too.
        var parts = unsupported.Select(u =>
        {
            if (u.Contains('(')) return u; // handler already embedded a hint (valid props / no such column / …)
            var s = CommandBuilder.SuggestPropertyScoped(u, scope);
            return s != null ? $"{u} (did you mean: {s}?)" : u;
        });
        var body = $"UNSUPPORTED props: {string.Join(", ", parts)}";
        // When the handler already named the valid alternative, that specific
        // hint beats generic guidance. Otherwise point at help (prop discovery)
        // first and raw-set (escape hatch) second — mirrors the non-resident
        // CommandBuilder.FormatUnsupported message so both paths steer the user
        // to discover the real prop instead of reaching for raw XML.
        return hasNamedAlternative
            ? body
            : $"{body}. Run 'officecli help <format> <element>' to see valid props, or raw-set for raw XML.";
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

            // ARCHITECTURE(handler-as-truth): wrap user props in a tracking
            // dict; the handler reading a key counts as consumption. After
            // handler.Add returns, any unread input key is reported as
            // unsupported_property. CONSISTENCY: mirrors CommandBuilder.Add.
            var tracking = new OfficeCli.Core.TrackingPropertyDictionary(properties);
            var resultPath = _handler.Add(parentPath, type, position, tracking);
            Console.WriteLine($"Added {type} at {resultPath}");
            var overflow = CommandBuilder.CheckTextOverflow(_handler, resultPath);
            if (overflow != null)
                Console.Error.WriteLine($"  WARNING: {overflow}");

            var allUnsupported = tracking.UnusedKeys.ToList();
            if (_handler is WordHandler residWh)
            {
                allUnsupported.AddRange(residWh.LastAddUnsupportedProps);
                foreach (var w in residWh.LastAddWarnings)
                    Console.Error.WriteLine($"  WARNING: {w}");
            }

            // Unrecognized LaTeX commands/environments from an equation parse.
            // Emit a distinctive stderr line so the dispatcher maps it to
            // exit 2 and the envelope to the unrecognized_latex_command code,
            // mirroring the one-shot CLI path (CommandBuilder.Add).
            EmitUnrecognizedLatex(_handler);

            if (allUnsupported.Count > 0)
            {
                if (_handler is WordHandler)
                {
                    var scope = resultPath.StartsWith("/styles/", StringComparison.Ordinal) ? "/styles" : resultPath;
                    var hint = OfficeCli.Core.StyleUnsupportedHints.Format(allUnsupported, scope);
                    if (hint != null) Console.Error.WriteLine("WARNING: " + hint);
                }
                else
                {
                    string? scope = _handler switch
                    {
                        ExcelHandler => "excel",
                        WordHandler => "word",
                        PowerPointHandler => "pptx",
                        _ => null,
                    };
                    Console.Error.WriteLine(FormatUnsupportedLine(allUnsupported, scope));
                }
            }
        }
    }

    private void ExecuteRemove(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        // Agent-safety: reject a bare unscoped selector (mirrors CommandBuilder).
        OfficeCli.Core.MutationSelectorGuard.EnsureScoped(path, "remove");
        var shift = req.GetArgOrNull("shift");
        if (!string.IsNullOrEmpty(shift))
        {
            if (_handler is not OfficeCli.Handlers.ExcelHandler xl)
                throw new OfficeCli.Core.CliException(
                    "--shift is supported only for Excel cell paths (e.g. /Sheet1/B5).")
                { Code = "invalid_argument" };
            xl.RemoveCellWithShift(path, shift);
        }
        else
        {
            var props = req.GetProps();
            _handler.Remove(path, props.Count > 0 ? props : null);
        }
        Console.WriteLine($"Removed {path}");
    }

    private void ExecuteMove(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var to = req.GetArgOrNull("to");
        var props = req.GetProps();
        var resultPath = _handler.Move(path, to, BuildInsertPosition(req), props.Count > 0 ? props : null);
        Console.WriteLine($"Moved to {resultPath}");
    }

    private void ExecuteRefresh(ResidentRequest req)
    {
        if (_handler is not OfficeCli.Handlers.WordHandler)
        {
            Console.Error.WriteLine("refresh currently only supports .docx files.");
            return;
        }
        _handler.Dispose();
        bool ok = false;
        string backend = "";
        if (OperatingSystem.IsWindows())
        {
            try { ok = OfficeCli.Core.WordPdfBackend.RefreshFields(_filePath); } catch { }
            if (ok) backend = "word";
        }
        if (!ok)
        {
            try { ok = OfficeCli.Core.WordHtmlRefresh.RefreshViaHtml(_filePath); } catch { }
            if (ok) backend = "html";
        }
        _handler = OfficeCli.Handlers.DocumentHandlerFactory.Open(_filePath, _editable);
        if (!ok)
        {
            Console.Error.WriteLine("refresh failed (Word backend unavailable and HTML fallback failed).");
            return;
        }
        if (backend == "html")
            Console.Error.WriteLine("Note: HTML fallback used. TOC page numbers reflect officecli's HTML pagination.");
        Console.WriteLine($"Refreshed: {_filePath} (backend: {backend})");
    }

    private void ExecuteSwap(ResidentRequest req)
    {
        var path1 = req.GetArg("path", "/");
        var path2 = req.GetArg("to", "/");
        var (p1, p2) = _handler switch
        {
            OfficeCli.Handlers.PowerPointHandler ppt => ppt.Swap(path1, path2),
            OfficeCli.Handlers.WordHandler word => word.Swap(path1, path2),
            OfficeCli.Handlers.ExcelHandler excel => excel.Swap(path1, path2),
            _ => throw new InvalidOperationException("swap not supported for this document type")
        };
        Console.WriteLine($"Swapped {p1} <-> {p2}");
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
            // BUG-DUMP12-01: emit VALIDATION report to stderr (not stdout) so the
            // ProcessRequest exit-code logic — which checks stderr for failure
            // tokens — promotes the request to a non-zero exit code. Writing to
            // stdout also corrupted batch --json output (BUG-R5-01 rationale).
            Console.Error.WriteLine($"VALIDATION: {newErrors.Count} new error(s) introduced:");
            foreach (var err in newErrors)
            {
                Console.Error.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.Error.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.Error.WriteLine($"    Part: {err.Part}");
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
            // BUG-DUMP12-01: route VALIDATION report to stderr — see ExecuteRawSet
            // for rationale (mirrors CommandBuilder.ReportNewErrors and lets the
            // ProcessRequest exit-code logic promote the request to exit 1).
            Console.Error.WriteLine($"VALIDATION: {newErrors.Count} new error(s) introduced:");
            foreach (var err in newErrors)
            {
                Console.Error.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.Error.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.Error.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    // R7-bt-3 / R7-bt-4: validate exit code & stream destination.
    // Pre-fix the resident path printed everything (including failure
    // reports) to stdout and the wrapper at the bottom of ProcessRequest
    // returned exit 0 because no stderr / UNSUPPORTED token was emitted.
    // Track the validation outcome here so ProcessRequest can promote it
    // to a non-zero exit code, and write the failure report to stderr —
    // mirrors the standard convention for diagnostic / lint tools.
    private int _lastValidateErrorCount;

    private void ExecuteValidate()
    {
        var errors = _handler.Validate();
        _lastValidateErrorCount = errors.Count;
        if (errors.Count == 0)
        {
            Console.WriteLine("Validation passed: no errors found.");
        }
        else
        {
            Console.Error.WriteLine($"Found {errors.Count} validation error(s):");
            foreach (var err in errors)
            {
                Console.Error.WriteLine($"  [{err.ErrorType}] {err.Description}");
                if (err.Path != null) Console.Error.WriteLine($"    Path: {err.Path}");
                if (err.Part != null) Console.Error.WriteLine($"    Part: {err.Part}");
            }
        }
    }

    private void ExecuteSave()
    {
        // No mutations have happened yet — the resident is still in the
        // shared-read state, and disk already matches the in-memory tree.
        // Treat as a no-op rather than promoting to editable (which would
        // re-open the file and acquire a write lock for no benefit).
        if (!_editable)
        {
            Console.WriteLine($"No pending changes for {Path.GetFileName(_filePath)}");
            return;
        }
        // Clean-skip: disk already matches the in-memory tree (a prior
        // save/autosave/each-mode flush cleared _dirty and every mutation
        // re-latches it via PromoteToEditable). Skipping the O(n) re-serialize
        // makes a defensive `save` before external reads effectively free.
        // Same message and exit code as the editable no-op above so callers
        // scripting on `save` see an unchanged success contract.
        if (!_dirty)
        {
            Console.WriteLine($"No pending changes for {Path.GetFileName(_filePath)}");
            return;
        }
        var sw = System.Diagnostics.Stopwatch.StartNew();
        _handler.Save();
        sw.Stop();
        _dirty = false;
        RecordSaveDuration(sw.Elapsed);
        Console.WriteLine($"Saved {Path.GetFileName(_filePath)}");
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

    // Generous upper bound on a single pipe message (see ResidentClient for the
    // rationale): bounds memory against a runaway/garbage stream while never
    // silently truncating a real request, which would surface up the stack as a
    // bogus "command not delivered".
    private const int MaxMessageLength = 512 * 1024 * 1024; // 512 MB

    private static async Task<string?> ReadLineFromPipeAsync(Stream pipe, CancellationToken token)
    {
        if (!OperatingSystem.IsWindows())
        {
            // BUG-R41-F1: disable BOM detection so binary garbage with a UTF-16 BOM
            // (0xFF 0xFE) doesn't cause the StreamReader to switch to UTF-16 mode and
            // then get stuck or throw when the byte stream doesn't conform to UTF-16.
            // Without detectEncodingFromByteOrderMarks=false, 0xFF 0xFE + partial data
            // causes ReadLineAsync to return null (EOF) or throw, and our error-response
            // write then fails because the client has already disconnected — producing a
            // silent 0-byte response. With UTF-8 forced, 0xFF 0xFE is treated as
            // malformed UTF-8 (replaced by the substitution char) and returned as a
            // garbage string, which then fails JSON parsing with a proper error response.
            using var reader = new StreamReader(pipe, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
            return await reader.ReadLineAsync(token);
        }
        // Windows: read in whole chunks until the newline terminator. The request
        // is a single line of compact JSON (no embedded raw newlines) and nothing
        // follows it on a one-command-per-connection pipe, so the first '\n' ends
        // the message. The newline is located with a vectorized span scan and
        // chunks are appended in bulk — no per-byte work. A request that fits in
        // one read is decoded straight from the read buffer with zero intermediate
        // allocation. There is no length cap that silently drops an over-size
        // request; a generous ceiling only guards against a runaway stream.
        var chunk = new byte[65536];
        List<byte>? acc = null; // allocated only when a request spans multiple reads
        while (true)
        {
            var bytesRead = await pipe.ReadAsync(chunk.AsMemory(0, chunk.Length), token);
            if (bytesRead == 0)
                return acc is { Count: > 0 } ? DecodeLine(CollectionsMarshal.AsSpan(acc)) : null;

            var span = chunk.AsSpan(0, bytesRead);
            var nl = span.IndexOf((byte)'\n');
            if (nl >= 0)
            {
                var tail = span[..nl];
                if (acc is null || acc.Count == 0)
                    return DecodeLine(tail); // single-read fast path: no List, no copy
                acc.AddRange(tail);
                return DecodeLine(CollectionsMarshal.AsSpan(acc));
            }

            acc ??= new List<byte>(bytesRead * 2);
            acc.AddRange(span);
            if (acc.Count > MaxMessageLength)
                throw new IOException($"Resident pipe message exceeded {MaxMessageLength} bytes.");
        }
    }

    private static string DecodeLine(ReadOnlySpan<byte> line)
    {
        if (line.Length > 0 && line[^1] == (byte)'\r')
            line = line[..^1];
        return Encoding.UTF8.GetString(line);
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
                LogStderr("Warning: shutdown timed out after 10 minutes, forcing exit.");
                Environment.Exit(1);
            }
        }
        catch (Exception ex)
        {
            LogStderr($"Warning: shutdown error: {ex.Message}");
        }

        try { _commandLock.Dispose(); } catch { }
        try { _mainCts.Dispose(); } catch { }
        try { _pingCts.Dispose(); } catch { }
        try { _idleCts.Dispose(); } catch { }
        try { _autosaveCts.Dispose(); } catch { }
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
                LogStderr("Warning: timeout waiting for in-flight command to drain.");
            }
        }
        catch (ObjectDisposedException) { /* _commandLock already disposed */ }

        // 4. Dispose the handler. Slow (writes the OpenXML tree back to
        //    disk and closes the file handle). The ping pipe is still
        //    live right now, so any TryResident caller will correctly
        //    conclude "resident still owns the file".
        // Capture whether the backing file was already gone BEFORE the flush.
        // The atomic writer now recreates a missing target (File.Move fallback,
        // so a delete/rename mid-session no longer loses the edits), which means
        // the post-dispose File.Exists check below can no longer see the
        // vanish — the file is back. Sample it here instead so the close still
        // reports that the file had been removed and was rebuilt.
        bool backingMissingBeforeFlush = !File.Exists(_filePath);
        bool disposeFailed = false;
        try { _handler.Dispose(); }
        catch (Exception ex)
        {
            disposeFailed = true;
            LogStderr($"Warning: handler dispose error: {ex.Message}");
        }

        // BUG-BT-R26-2 / BUG-R43: detect data loss. The original probe used
        // File.Exists(_filePath) post-Dispose — but on macOS, renaming the
        // file via Finder/mv preserves the inode, so OpenXML still wrote our
        // data successfully (just under a different path). That triggered a
        // false-positive "data may be lost" close-time error.
        // Flag data loss only when Dispose itself failed (positive evidence
        // the save did not land). A vanished path alone (rename or unlink
        // post-save) is no longer treated as a loss — the bytes are durable
        // either way; an external rename is the user's intent.
        if (disposeFailed)
        {
            _shutdownFileMissing = true;
            LogStderr($"ERROR: save failed during shutdown — data may be lost: {_filePath}");
        }
        // BUG-INTERVIEW-EDIT-R10-B: the backing file was removed from its
        // original path during the session (rm/Trash, or an external
        // rename/mv). The atomic writer's move-fallback has now recreated it
        // at that path with the in-memory edits, so the bytes are NOT lost —
        // but the removal is still worth surfacing: the user deleted or moved
        // the file and we brought it back, which is surprising if unannounced,
        // and in the rename case a stale copy also exists at the new path.
        // stderr only — no _shutdownFileMissing flag, so the exit code stays 0.
        else if (backingMissingBeforeFlush)
        {
            _shutdownFileVanishedAfterDispose = true;
            LogStderr(
                $"WARNING: the backing file was missing at its original path during save: {_filePath}. " +
                "It was deleted or moved externally; your changes were rebuilt at that path. " +
                "If you renamed/moved it, a separate copy also exists at the new location.");
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

    /// <summary>
    /// BUG-R36-B7 helper. Mirror of CommandBuilder.View.ParsePptHtmlPage —
    /// validate --page against slide count and reject silent fallbacks.
    /// </summary>
    private static (int? start, int? end) ResolvePptHtmlPage(
        string? pageFilter, int? start, int? end,
        OfficeCli.Handlers.PowerPointHandler pptHandler)
    {
        if (string.IsNullOrEmpty(pageFilter)) return (start, end);
        var slideCount = pptHandler.Query("slide").Count;
        var firstTok = pageFilter.Split(',')[0].Trim();
        if (firstTok.Contains('-'))
        {
            var parts = firstTok.Split('-', 2);
            if (!int.TryParse(parts[0], out var ps) || !int.TryParse(parts[1], out var pe))
                throw new ArgumentException($"Invalid --page value '{pageFilter}': expected N or M-N or comma list.");
            if (ps <= 0 || pe <= 0)
                throw new ArgumentException($"Invalid --page value '{pageFilter}': slide number must be >= 1.");
            if (ps > slideCount)
                throw new ArgumentException($"--page {ps} out of range (total slides: {slideCount}).");
            return (ps, Math.Min(pe, slideCount));
        }
        if (!int.TryParse(firstTok, out var p))
            throw new ArgumentException($"Invalid --page value '{pageFilter}': expected a positive slide number.");
        if (p <= 0)
            throw new ArgumentException($"Invalid --page value '{pageFilter}': slide number must be >= 1.");
        if (p > slideCount)
            throw new ArgumentException($"--page {p} out of range (total slides: {slideCount}).");
        return (p, p);
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
