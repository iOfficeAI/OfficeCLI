// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Text;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Shared subprocess-driver for short-lived plugin invocations
/// (dump-reader, exporter) — and for the spawn side of format-handler
/// sessions. Implements the §5.6 idle-timeout watchdog: any byte on stdout,
/// or a heartbeat line on stderr matching <c>{"heartbeat":true}</c>, resets
/// the activity timer; once the gap exceeds the budget, the process tree
/// is killed and the caller sees <c>plugin_idle_timeout</c>.
///
/// Wall-clock time is intentionally not bounded — a 4 GB .doc that takes
/// 20 minutes to dump but is constantly producing output is fine.
/// </summary>
public static class PluginProcess
{
    public sealed record RunResult(
        int ExitCode,
        string Stderr,
        bool IdleTimedOut);

    public sealed class RunOptions
    {
        public required string ExecutablePath { get; init; }
        public required IEnumerable<string> Arguments { get; init; }

        /// <summary>Idle timeout in seconds. 0 disables the watchdog entirely.</summary>
        public int IdleTimeoutSeconds { get; init; } = 60;

        /// <summary>Extra environment variables. <c>OFFICECLI_BIN</c> is added automatically.</summary>
        public Dictionary<string, string>? ExtraEnv { get; init; }

        /// <summary>
        /// Per-line stdout callback. If null, stdout is drained silently. Lines
        /// are delivered without the trailing newline. Callback exceptions are
        /// captured into <see cref="LineCallbackError"/> and stop the run.
        /// </summary>
        public Action<string>? OnStdoutLine { get; init; }

        /// <summary>
        /// Optional sink for stderr lines that are not heartbeats. If null,
        /// stderr is collected into <see cref="RunResult.Stderr"/> for the
        /// caller to surface on failure.
        /// </summary>
        public Action<string>? OnStderrLine { get; init; }
    }

    /// <summary>
    /// Most recent exception thrown by an <see cref="RunOptions.OnStdoutLine"/>
    /// callback in the current call, or null if the callback succeeded. Stored
    /// per-call (the AsyncLocal scope is per <see cref="Run"/> invocation).
    /// </summary>
    public static Exception? LineCallbackError { get; private set; }

    public static RunResult Run(RunOptions opts)
    {
        LineCallbackError = null;

        var psi = new ProcessStartInfo
        {
            FileName = opts.ExecutablePath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        foreach (var a in opts.Arguments) psi.ArgumentList.Add(a);

        var selfPath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(selfPath))
            psi.Environment["OFFICECLI_BIN"] = selfPath;
        if (opts.ExtraEnv is not null)
            foreach (var kv in opts.ExtraEnv)
                psi.Environment[kv.Key] = kv.Value;

        using var p = new Process { StartInfo = psi };
        if (!p.Start())
            return new RunResult(-1, "Failed to start plugin process.", false);

        // Activity timestamp shared by both reader tasks and the watchdog.
        // Wall-clock ticks (DateTime.UtcNow.Ticks, 100ns resolution) instead
        // of Stopwatch.GetTimestamp: Stopwatch is monotonic and on some
        // platform / hardware combinations keeps ticking through system
        // suspend, on others it pauses — behavior depends on QPC / TSC
        // properties. Wall-clock unambiguously advances during suspend, so
        // a laptop waking up after an hour mid-plugin gets an honest
        // "idle for an hour" reading and we kill the (likely-stale) plugin
        // instead of letting it run on dead network sockets / file handles.
        long lastActivityTicks = DateTime.UtcNow.Ticks;
        var stderrCollector = new StringBuilder();
        var stderrLock = new object();

        var stdoutTask = Task.Run(() => ReadStdout(p, opts, ref lastActivityTicks));
        var stderrTask = Task.Run(() => ReadStderr(p, opts, stderrCollector, stderrLock, ref lastActivityTicks));

        bool idleTimedOut = false;
        if (opts.IdleTimeoutSeconds > 0)
        {
            var budgetTicks = TimeSpan.FromSeconds(opts.IdleTimeoutSeconds).Ticks;
            var pollIntervalMs = Math.Max(250, opts.IdleTimeoutSeconds * 1000 / 4);

            while (!p.HasExited)
            {
                if (p.WaitForExit(pollIntervalMs)) break;
                var since = DateTime.UtcNow.Ticks - Volatile.Read(ref lastActivityTicks);
                if (since > budgetTicks)
                {
                    idleTimedOut = true;
                    try { p.Kill(entireProcessTree: true); } catch { }
                    break;
                }
            }
        }

        // Drain the reader tasks. WaitForExit guarantees the streams close.
        try { p.WaitForExit(); } catch { }
        try { stdoutTask.Wait(2000); } catch { }
        try { stderrTask.Wait(2000); } catch { }

        string stderr;
        lock (stderrLock) stderr = stderrCollector.ToString();

        return new RunResult(p.ExitCode, stderr, idleTimedOut);
    }

    private static void ReadStdout(Process p, RunOptions opts, ref long activityTicks)
    {
        var reader = p.StandardOutput;
        try
        {
            string? line;
            while ((line = reader.ReadLine()) is not null)
            {
                Volatile.Write(ref activityTicks, DateTime.UtcNow.Ticks);
                if (opts.OnStdoutLine is null) continue;
                try { opts.OnStdoutLine(line); }
                catch (Exception ex)
                {
                    LineCallbackError = ex;
                    try { p.Kill(entireProcessTree: true); } catch { }
                    return;
                }
            }
        }
        catch { /* stream closed / process killed */ }
    }

    private static void ReadStderr(Process p, RunOptions opts, StringBuilder collector, object collectorLock, ref long activityTicks)
    {
        var reader = p.StandardError;
        try
        {
            string? line;
            while ((line = reader.ReadLine()) is not null)
            {
                // Heartbeat lines reset the watchdog but are NOT surfaced to
                // the caller — they're plumbing, not diagnostics. Match
                // tolerantly: any JSON object that has a truthy
                // "heartbeat" field.
                if (IsHeartbeat(line))
                {
                    Volatile.Write(ref activityTicks, DateTime.UtcNow.Ticks);
                    continue;
                }

                // Any non-heartbeat stderr output also counts as activity.
                Volatile.Write(ref activityTicks, DateTime.UtcNow.Ticks);
                if (opts.OnStderrLine is not null)
                {
                    try { opts.OnStderrLine(line); } catch { /* ignore sink errors */ }
                }
                else
                {
                    lock (collectorLock)
                    {
                        collector.AppendLine(line);
                        // Cap collected stderr at 16 KB to bound memory if a
                        // plugin spams diagnostics. We keep the head — usually
                        // the first error line is the most useful.
                        if (collector.Length > 16 * 1024)
                            collector.Length = 16 * 1024;
                    }
                }
            }
        }
        catch { /* stream closed / process killed */ }
    }

    /// <summary>
    /// True if <paramref name="line"/> is a §5.6 heartbeat envelope
    /// (<c>{"heartbeat":true,...}</c>). Exposed so long-running drivers
    /// (FormatHandlerSession) can apply the same activity semantics on
    /// stderr without duplicating the parse.
    /// </summary>
    internal static bool IsHeartbeat(string line)
    {
        // Cheap pre-filter to avoid JSON parse cost on every diagnostic line:
        // every heartbeat envelope starts with `{` and contains "heartbeat".
        if (line.Length < 14) return false;
        if (line[0] != '{') return false;
        if (line.IndexOf("heartbeat", StringComparison.Ordinal) < 0) return false;
        try
        {
            using var doc = System.Text.Json.JsonDocument.Parse(line);
            if (doc.RootElement.ValueKind != System.Text.Json.JsonValueKind.Object) return false;
            if (!doc.RootElement.TryGetProperty("heartbeat", out var hb)) return false;
            return hb.ValueKind == System.Text.Json.JsonValueKind.True;
        }
        catch { return false; }
    }
}
