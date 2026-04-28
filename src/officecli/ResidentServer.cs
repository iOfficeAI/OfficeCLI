// Copyright 2025 OfficeCli (officecli.ai)
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
    // Idle timeout is mutable: `create` starts the resident with a short
    // 60s timeout, and a later `open` upgrades it to 12min via the
    // `__set-idle-timeout__` ping command. Stored as ticks so we can
    // do atomic Volatile reads/writes (TimeSpan is a multi-field struct
    // and can't be volatile'd directly).
    private long _idleTimeoutTicks = ResolveIdleTimeout().Ticks;
    private TimeSpan CurrentIdleTimeout => TimeSpan.FromTicks(Volatile.Read(ref _idleTimeoutTicks));
    private CancellationTokenSource _idleCts = new();
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
                // Snapshot the current idle CTS and timeout on each loop
                // iteration: ResetIdleTimer() swaps _idleCts to restart
                // the wait, and TrySetIdleTimeout() mutates
                // _idleTimeoutTicks and calls ResetIdleTimer() so the new
                // duration is observed here on the very next pass.
                var idleCts = Volatile.Read(ref _idleCts);
                var currentTimeout = CurrentIdleTimeout;
                using var linked = CancellationTokenSource.CreateLinkedTokenSource(idleCts.Token, token);
                await Task.Delay(currentTimeout, linked.Token);

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
                var response = _shutdownFileMissing
                    ? MakeResponse(1, "",
                        $"file vanished during resident session — data may be lost: {_filePath}")
                    : MakeResponse(0, "Closing resident.", "");
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
            LogStderr($"Resident error: {ex.Message}");
        }
        finally
        {
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

            // BUG-R40-B10: batch failure rows print to stdout, not stderr,
            // so the generic stderr/UNSUPPORTED inspection below sees a
            // clean stderr and returns exit 0 even when every batch item
            // failed. ExecuteBatch records the verdict in
            // _lastBatchHadFailure; promote it to a non-zero exit here.
            var isBatch = request.Command.Equals("batch", StringComparison.OrdinalIgnoreCase);
            var batchFailure = isBatch && _lastBatchHadFailure;

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
                // BUG-R11-03: JSON-mode exit code must match text mode. Previously
                // hard-coded to 0, which silently swallowed every error type
                // (path-not-found, unsupported_property, failed open) for any
                // resident --json call. Map parity with text mode below:
                //   - envelope success:false                        -> 1
                //   - stderr contains UNSUPPORTED (unsupported_property) -> 2
                //   - otherwise                                      -> 0
                int jsonExitCode = 0;
                if (stderr.Contains("UNSUPPORTED"))
                    jsonExitCode = 2;
                else if (!EnvelopeSuccess(envelope) || batchFailure)
                    jsonExitCode = 1;
                return MakeResponse(jsonExitCode, envelope, "");
            }

            int exitCode = stderr.Contains("UNSUPPORTED") ? 2 : (batchFailure ? 1 : 0);
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
            return MakeResponse(1, "", $"Error: {rendered.Message}");
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
            case "swap":
                ExecuteSwap(request);
                NotifyWatchFullRefresh();
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

    // BUG-R40-B10: track whether the most recent ExecuteBatch saw any
    // failures so ProcessRequest can surface a non-zero exit code.
    // Without this, a batch where every item fails returned exit 0 because
    // the wrapper at the bottom of ProcessRequest only inspected stderr
    // (and batch failure rows are written to stdout).
    private bool _lastBatchHadFailure;

    private void ExecuteBatch(ResidentRequest request)
    {
        _lastBatchHadFailure = false;
        var batchJson = request.GetArg("batchJson");
        var force = request.GetArg("force", "false")
            .Equals("true", StringComparison.OrdinalIgnoreCase);
        var stopOnError = !force;
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

        _lastBatchHadFailure = results.Any(r => !r.Success);
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
        if (!WatchServer.IsWatching(_filePath)) return;

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
        if (!WatchServer.IsWatching(_filePath)) return;

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
            {
                // BUG-R36-B7: honor --page on pptx html with strict bounds.
                var (pStart, pEnd) = ResolvePptHtmlPage(pageFilter, start, end, pptHandler);
                html = pptHandler.ViewAsHtml(pStart, pEnd);
            }
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
                    // SECURITY: include a random token so the preview path is not predictable.
                    // Without it, a predictable path enables a symlink pre-placement attack that
                    // causes File.WriteAllText to clobber an arbitrary victim file. See
                    // CommandBuilder.View.cs for the same fix.
                    var htmlPath = Path.Combine(Path.GetTempPath(), $"officecli_preview_{Path.GetFileNameWithoutExtension(_filePath)}_{DateTime.Now:HHmmss}_{Guid.NewGuid():N}.html");
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

        Console.WriteLine(OutputFormatter.FormatNode(node, format));
    }

    private void ExecuteQuery(ResidentRequest req, OutputFormat format)
    {
        var selector = req.GetArg("selector", "");
        var filters = AttributeFilter.Parse(selector);
        // CONSISTENCY(cell-selector-alias): mirror the direct-mode normalization in
        // CommandBuilder.GetQuery.cs — without this, resident-mode Excel cell queries
        // with short aliases (bold, size, ...) silently drop every hit (BUG-R17-01).
        if (_handler is ExcelHandler
            && selector.TrimStart().StartsWith("cell", StringComparison.OrdinalIgnoreCase))
        {
            filters = AttributeFilter.NormalizeKeys(filters, ExcelHandler.ResolveCellAttributeAlias);
        }
        var (results, warnings) = AttributeFilter.ApplyWithWarnings(_handler.Query(selector), filters);
        var textFilter = req.GetArgOrNull("text");
        if (!string.IsNullOrEmpty(textFilter))
            results = results.Where(n => n.Text != null && n.Text.Contains(textFilter, StringComparison.OrdinalIgnoreCase)).ToList();
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
        foreach (var w in warnings) Console.Error.WriteLine(w);
        Console.WriteLine(OutputFormatter.FormatNodes(results, format));
    }

    private void ExecuteSet(ResidentRequest req)
    {
        var path = req.GetArg("path", "/");
        var properties = req.GetProps();
        var unsupported = _handler.Set(path, properties);
        var unsupportedKeys = unsupported
            .Select(u => u.Contains(' ') ? u[..u.IndexOf(' ')] : u)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var applied = properties.Where(kv => !unsupportedKeys.Contains(kv.Key)).ToList();

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

        var message = applied.Count > 0
            ? $"Updated {path}: {string.Join(", ", applied.Select(kv => $"{kv.Key}={kv.Value}"))}"
              + (findMatchCount.HasValue ? $" ({findMatchCount.Value} matched)" : "")
            : (unsupported.Count > 0 ? $"No properties applied to {path}" : $"Updated {path}");

        var warnings = new List<OfficeCli.Core.CliWarning>();
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
            if (findMatchCount is 0)
                Console.Error.WriteLine($"WARNING: find pattern matched 0 occurrences at {path}");
            if (overflow != null)
                Console.Error.WriteLine($"  WARNING: {overflow}");
        }

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
                Console.Error.WriteLine($"UNSUPPORTED props (use raw-set instead): {string.Join(", ", unsupported)}");
            }
        }
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

            // BUG(add-lies): schema-level pre-check so bogus --prop keys
            // don't get silently swallowed by handler.Add. The UNSUPPORTED
            // stderr line is how ProcessRequest (above) escalates exit
            // code to 2 and sets the envelope warning code to
            // "unsupported_property" — so emitting it here is enough to
            // get parity with CLI-inline Add.
            // CONSISTENCY(schema-prop-validation): mirrors the call site in
            // CommandBuilder.Add.cs.
            var fmt = SchemaHelpLoader.FormatForExtension(Path.GetExtension(_filePath));
            var schemaUnsupported = fmt != null
                ? SchemaHelpLoader.ValidateProperties(fmt, type, "add", properties)
                : Array.Empty<string>();

            // CONSISTENCY(no-double-unsupported): see CommandBuilder.Add.cs —
            // strip schema-flagged keys before handler.Add so pivot / other
            // helpers don't re-warn with different phrasing.
            if (schemaUnsupported.Count > 0)
            {
                foreach (var u in schemaUnsupported)
                    properties.Remove(u);
            }

            var resultPath = _handler.Add(parentPath, type, position, properties);
            Console.WriteLine($"Added {type} at {resultPath}");
            var overflow = CommandBuilder.CheckTextOverflow(_handler, resultPath);
            if (overflow != null)
                Console.Error.WriteLine($"  WARNING: {overflow}");

            // Combine schema-level unsupported (caught before handler.Add) and
            // handler-level silent-drop (e.g. AddStyle props that pass schema
            // validation via the `font.` prefix but the curated AddStyle has
            // no slot for, like `font.eastAsia`). Both surface to the user.
            var allUnsupported = new List<string>(schemaUnsupported);
            if (_handler is WordHandler residWh)
                allUnsupported.AddRange(residWh.LastAddUnsupportedProps);

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
                    Console.Error.WriteLine($"UNSUPPORTED props (use raw-set instead): {string.Join(", ", allUnsupported)}");
                }
            }
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
        try { _handler.Dispose(); }
        catch (Exception ex)
        {
            LogStderr($"Warning: handler dispose error: {ex.Message}");
        }

        // BUG-BT-R26-2: detect file-deleted-mid-session. If the user (or an
        // external process) unlinked the file while the resident held it
        // open, OpenXML wrote into a now-orphaned inode and the on-disk
        // path no longer points at our data. Loud-fail so the close
        // command's exit code + stderr message reflect the loss instead
        // of pretending the save succeeded.
        try
        {
            if (!File.Exists(_filePath))
            {
                _shutdownFileMissing = true;
                LogStderr($"ERROR: file vanished during resident session — data may be lost: {_filePath}");
            }
        }
        catch { /* best-effort probe; don't mask the real shutdown */ }

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
