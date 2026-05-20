// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Owns a running format-handler plugin process and the stdin/stdout
/// channel used to talk to it. Per docs/plugin-protocol.md §2.3 / §5.3 / §6.1.
///
/// Lifecycle (matches the §6.7 state machine):
///   spawning → ready (after open-handshake) → busy (per Send) → broken
///   (on IO failure, idle timeout, or malformed reply) → closed (on Dispose).
///
/// Transport: plain stdin/stdout (one JSON object per line, UTF-8 no BOM).
/// stderr carries diagnostics plus heartbeat lines (§5.6). The choice of
/// stdin/stdout over named pipes makes every language a first-class plugin
/// runtime — no <c>NamedPipeClient</c> or <c>UnixStream</c> wrapper to
/// learn — and matches the framing dump-reader / exporter already use.
///
/// The session is single-threaded by way of <see cref="_ioLock"/>: each
/// public Send call serializes on the lock and completes a full
/// request/response round-trip before releasing.
/// </summary>
internal sealed class FormatHandlerSession : IDisposable
{
    private readonly string _filePath;
    private readonly ResolvedPlugin _plugin;
    private Process? _proc;
    private StreamReader? _stdoutReader;
    private StreamWriter? _stdinWriter;
    private Task? _stderrPump;
    private bool _disposed;
    private bool _broken;
    private long _lastActivityTicks;
    private PluginSessionCapabilities? _sessionCaps;
    private readonly object _ioLock = new();

    public ResolvedPlugin Plugin => _plugin;

    /// <summary>
    /// Capabilities and vocabulary the plugin reported during the open
    /// handshake. Null until <see cref="Start"/> completes.
    /// </summary>
    public PluginSessionCapabilities? Capabilities => _sessionCaps;

    /// <summary>
    /// True once an IO failure or idle timeout has poisoned the session.
    /// Further Send calls fail fast with <c>plugin_stream_closed</c>.
    /// </summary>
    public bool IsBroken => _broken;

    public FormatHandlerSession(string filePath, ResolvedPlugin plugin)
    {
        _filePath = Path.GetFullPath(filePath);
        _plugin = plugin;
    }

    public void Start(bool editable)
    {
        var utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var psi = new ProcessStartInfo
        {
            FileName = _plugin.ExecutablePath,
            ArgumentList = { "open", _filePath },
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            // Force UTF-8 no-BOM on all three streams. Windows defaults to
            // Console.InputEncoding/OutputEncoding which can be GBK/CP1252
            // depending on locale — wire format must be locale-independent.
            StandardInputEncoding = utf8NoBom,
            StandardOutputEncoding = utf8NoBom,
            StandardErrorEncoding = utf8NoBom,
        };
        var selfPath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(selfPath))
            psi.Environment["OFFICECLI_BIN"] = selfPath;

        _proc = Process.Start(psi)
            ?? throw new CliException($"Failed to start format-handler plugin '{_plugin.Manifest.Name}'.")
                { Code = "plugin_spawn_failed" };

        // Wrap stdin with an explicit UTF-8 no-BOM writer on the base
        // stream. Process.StandardInput's default StreamWriter buffers
        // independently and (on some runtimes) ignores AutoFlush — going
        // direct to BaseStream avoids the surprise.
        _stdinWriter = new StreamWriter(_proc.StandardInput.BaseStream, utf8NoBom, bufferSize: 8192, leaveOpen: true)
        {
            AutoFlush = true,
            NewLine = "\n",
        };
        _stdoutReader = _proc.StandardOutput;
        Volatile.Write(ref _lastActivityTicks, DateTime.UtcNow.Ticks);

        // Background stderr pump: heartbeat lines (`{"heartbeat":true}`)
        // reset the activity timer; everything else is diagnostic noise we
        // drain to keep the OS pipe buffer from filling and blocking the
        // plugin. We intentionally do not surface the diagnostic text here
        // — long-lived sessions can produce a lot of it, and the canonical
        // channel for plugin-side errors is the `error` envelope on
        // stdout, not freeform stderr.
        var stderr = _proc.StandardError;
        _stderrPump = Task.Run(() =>
        {
            try
            {
                string? line;
                while ((line = stderr.ReadLine()) is not null)
                {
                    if (PluginProcess.IsHeartbeat(line))
                        Volatile.Write(ref _lastActivityTicks, DateTime.UtcNow.Ticks);
                    // Non-heartbeat lines: drained, not surfaced.
                }
            }
            catch { /* stream closed on shutdown */ }
        });

        // Open handshake. Plugin must reply with `ok` carrying capabilities
        // + vocabulary before we surface the session to callers.
        var openArgs = new JsonObject
        {
            ["path"] = _filePath,
            ["editable"] = editable,
        };
        var reply = SendRaw("open", null, openArgs, props: null,
            idleTimeoutSec: _plugin.Manifest.ResolveIdleTimeout("open"));
        try
        {
            _sessionCaps = reply is null
                ? null
                : JsonSerializer.Deserialize(reply.ToJsonString(), PluginJsonContext.Default.PluginSessionCapabilities);
        }
        catch (JsonException)
        {
            _sessionCaps = null;
        }
    }

    /// <summary>
    /// Send a request envelope and synchronously wait for the matching reply.
    /// Throws <see cref="CliException"/> on protocol error, IO failure, or
    /// plugin-reported error responses.
    /// </summary>
    public JsonNode? Send(string msgType, string? command, JsonObject? args = null, JsonObject? props = null)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(FormatHandlerSession));
        if (_broken)
            throw new CliException(
                $"Format-handler session for '{_plugin.Manifest.Name}' is no longer usable (stream was closed earlier).")
            { Code = "plugin_stream_closed" };

        // Capability gate: short-circuit verbs the plugin already declared it
        // does not support, avoiding a wasted round-trip and ambiguous errors.
        if (command is not null && _sessionCaps?.Capabilities?.Commands is { Count: > 0 } cmds
            && !cmds.Contains(command))
        {
            throw new CliException(
                $"Format-handler plugin '{_plugin.Manifest.Name}' does not implement command '{command}'.")
            { Code = "unsupported_command" };
        }

        var verbForTimeout = command ?? msgType;
        var idle = _plugin.Manifest.ResolveIdleTimeout(verbForTimeout);
        return SendRaw(msgType, command, args, props, idle);
    }

    private JsonNode? SendRaw(string msgType, string? command, JsonObject? args, JsonObject? props, int idleTimeoutSec)
    {
        if (_stdinWriter is null || _stdoutReader is null)
            throw new InvalidOperationException("Session not started.");

        var request = new JsonObject
        {
            ["protocol"] = 1,
            ["msg_type"] = msgType,
        };
        if (command is not null) request["command"] = command;
        if (args is not null) request["args"] = args;
        if (props is not null) request["props"] = props;

        lock (_ioLock)
        {
            try
            {
                _stdinWriter.WriteLine(request.ToJsonString());
                Volatile.Write(ref _lastActivityTicks, DateTime.UtcNow.Ticks);

                // Read the reply with an idle-timeout watchdog. The budget is
                // "no activity for idleTimeoutSec seconds" — any stderr
                // heartbeat the plugin emits in the meantime resets the
                // timer (handled by the stderr pump), so long opaque work
                // can keep itself alive without being kicked.
                string? line;
                if (idleTimeoutSec > 0)
                {
                    line = ReadReplyWithIdleWatchdog(idleTimeoutSec, command ?? msgType);
                }
                else
                {
                    line = _stdoutReader.ReadLine();
                }

                if (line is null)
                {
                    _broken = true;
                    throw new CliException(
                        $"Format-handler plugin '{_plugin.Manifest.Name}' closed stdout unexpectedly (no reply to {msgType}/{command ?? ""}).")
                    { Code = "plugin_stream_closed" };
                }

                // Any protocol-shape failure poisons the session: §6.7 lists
                // "malformed reply" as a broken-state trigger, so we mark
                // _broken before throwing so the next Send fast-fails instead
                // of trying to write into a session whose protocol invariants
                // are gone. We also catch JsonException explicitly: the raw
                // System.Text.Json message ("'d' is an invalid start of a
                // value...") is opaque to users — wrap it in a clear
                // `protocol_mismatch` envelope that names the plugin and
                // shows a preview of what it actually wrote.
                JsonObject? reply;
                try
                {
                    reply = JsonNode.Parse(line)?.AsObject();
                }
                catch (JsonException ex)
                {
                    _broken = true;
                    throw new CliException(
                        $"Format-handler plugin '{_plugin.Manifest.Name}' wrote non-JSON to stdout (a JSONL envelope was expected): {ex.Message}. " +
                        $"First chars: \"{Truncate(line, 80)}\". This is a plugin bug — diagnostic output must go to stderr or --log-file, not stdout.")
                    { Code = "protocol_mismatch" };
                }

                if (reply is null)
                {
                    _broken = true;
                    throw new CliException(
                        $"Format-handler plugin '{_plugin.Manifest.Name}' reply is not a JSON object. First chars: \"{Truncate(line, 80)}\".")
                    { Code = "protocol_mismatch" };
                }

                var replyType = reply["msg_type"]?.GetValue<string>() ?? "";
                if (replyType == "ok")
                    return reply["result"];
                if (replyType == "error")
                {
                    var err = reply["error"]?.AsObject();
                    var code = err?["code"]?.GetValue<string>() ?? "plugin_error";
                    var msg = err?["message"]?.GetValue<string>() ?? "(no message)";
                    throw new CliException(
                        $"Format-handler plugin '{_plugin.Manifest.Name}' reported error on {command ?? msgType}: {msg}")
                    { Code = code };
                }
                _broken = true;
                throw new CliException(
                    $"Format-handler plugin '{_plugin.Manifest.Name}' replied with unknown msg_type '{replyType}'.")
                { Code = "protocol_mismatch" };
            }
            catch (IOException ex)
            {
                _broken = true;
                throw new CliException(
                    $"Format-handler plugin '{_plugin.Manifest.Name}' stdin/stdout I/O failed: {ex.Message}", ex)
                { Code = "plugin_stream_closed" };
            }
        }
    }

    /// <summary>
    /// Read one line from stdout, polling at short intervals so a stderr
    /// heartbeat that arrives mid-read can extend the budget. The naive
    /// <c>ReadLineAsync(ct)</c> approach uses a fixed cancellation
    /// deadline, which would kill a plugin that is heart-beating but
    /// slow to reply.
    /// </summary>
    private string? ReadReplyWithIdleWatchdog(int idleTimeoutSec, string verbForError)
    {
        var budgetTicks = TimeSpan.FromSeconds(idleTimeoutSec).Ticks;
        var readTask = Task.Run(() => _stdoutReader!.ReadLine());

        while (!readTask.IsCompleted)
        {
            // Poll at one-quarter the budget (250ms floor) so even short
            // timeouts fire reasonably close to the configured deadline.
            var pollMs = Math.Max(250, idleTimeoutSec * 1000 / 4);
            if (readTask.Wait(pollMs)) break;
            var since = DateTime.UtcNow.Ticks - Volatile.Read(ref _lastActivityTicks);
            if (since > budgetTicks)
            {
                _broken = true;
                TryKill();
                throw new CliException(
                    $"Format-handler plugin '{_plugin.Manifest.Name}' produced no activity for {idleTimeoutSec}s (command={verbForError}).")
                {
                    Code = "plugin_idle_timeout",
                    Suggestion = $"Raise `idle_timeout_seconds.verbs.{verbForError}` in the plugin's manifest, " +
                                 "emit periodic `{\"heartbeat\":true}` on stderr during long jobs, or pass --timeout 0 to disable.",
                };
            }
        }

        // readTask completed — propagate exceptions, then take the result.
        var line = readTask.GetAwaiter().GetResult();
        if (line is not null)
            Volatile.Write(ref _lastActivityTicks, DateTime.UtcNow.Ticks);
        return line;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Ask the plugin to shut down cleanly. If broken or unresponsive,
        // fall through to a hard kill below. We deliberately do not block on
        // the close reply for long — broken sessions should not delay exit.
        try
        {
            if (!_broken && _stdinWriter is not null && _stdoutReader is not null && _proc is { HasExited: false })
            {
                var close = new JsonObject
                {
                    ["protocol"] = 1,
                    ["msg_type"] = "close",
                };
                _stdinWriter.WriteLine(close.ToJsonString());

                // Closing stdin gives the plugin a clean EOF signal — the
                // canonical "we're done" handshake on stdin/stdout
                // transport. Plugins that ignore the close envelope still
                // see read() return 0 and exit.
                try { _proc!.StandardInput.Close(); } catch { }

                try
                {
                    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(2));
                    _ = _stdoutReader.ReadLineAsync(cts.Token).AsTask().GetAwaiter().GetResult();
                }
                catch { /* ack drain best-effort */ }
            }
            else if (_proc is { HasExited: false })
            {
                try { _proc.StandardInput.Close(); } catch { }
            }
        }
        catch { /* shutting down; ignore */ }

        try { _stdinWriter?.Dispose(); } catch { }
        try { _stdoutReader?.Dispose(); } catch { }

        if (_proc is not null)
        {
            try
            {
                if (!_proc.WaitForExit(2000))
                    TryKill();
            }
            catch { TryKill(); }
            try { _proc.Dispose(); } catch { }
        }

        try { _stderrPump?.Wait(500); } catch { }
    }

    private void TryKill()
    {
        try { _proc?.Kill(entireProcessTree: true); } catch { }
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
