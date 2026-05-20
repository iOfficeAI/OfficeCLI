// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using OfficeCli.Handlers;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Runs a dump-reader plugin per docs/plugin-protocol.md §5.1. The plugin
/// reads a foreign source file (e.g. .doc) and **streams** BatchItem objects
/// as JSONL (one JSON object per line) on stdout. Main opens a fresh native
/// scratch file (extension from the plugin's manifest <c>target</c> field —
/// docx/xlsx/pptx), replays each item as it arrives, and returns the
/// populated path.
///
/// Streaming has two benefits over a buffered JSON-array transport: per-line
/// activity feeds the idle watchdog (§5.6), and main's memory does not scale
/// with batch size on multi-million-paragraph .doc files.
///
/// The conversion is one-shot: edits to the returned file are not propagated
/// back to the source file.
/// </summary>
public static class DumpReaderInvoker
{
    public sealed record DumpResult(string ConvertedPath, ResolvedPlugin Plugin);

    /// <summary>
    /// Resolve a dump-reader plugin for <paramref name="sourceExt"/>, invoke it
    /// against <paramref name="sourceFullPath"/>, and replay the resulting
    /// JSONL stream into a fresh native file. Throws CliException on
    /// resolution or invocation failure; otherwise the result references a
    /// temp file the caller must dispose (or leave for OS tmp cleanup).
    /// </summary>
    public static DumpResult Run(string sourceFullPath, string sourceExt)
    {
        var plugin = PluginRegistry.FindFor(PluginKind.DumpReader, sourceExt)
            ?? throw new CliException($"No dump-reader plugin found for {sourceExt}.")
            {
                Code = "dump_reader_not_found",
                Suggestion = "Install a dump-reader plugin (`officecli plugins list` to see installed; docs/plugin-protocol.md for paths).",
            };

        var targetExt = plugin.Manifest.ResolveTargetExtension();
        var tmpOut = Path.Combine(Path.GetTempPath(),
            $"officecli-dumpread-{Guid.NewGuid():N}{targetExt}");
        // minimal: true gives a bare-skeleton native file (no default styles,
        // theme, or docDefaults for docx; equivalent skeleton for xlsx/pptx).
        // The plugin's batch is expected to define everything it references —
        // round-trip dumps from `officecli dump` do exactly that.
        BlankDocCreator.Create(tmpOut, locale: null, minimal: true);

        int itemIndex = 0;
        Exception? replayError = null;

        // v6.4: open the handler AFTER the plugin process finishes streaming.
        // The previous design called ExecuteBatchItem inside the PluginProcess
        // stdout reader task — i.e. on a background Task.Run thread. The
        // OpenXml SDK's WordprocessingDocument (System.IO.Packaging.Package
        // underneath) is not thread-safe: a heavy add-stream (5000+ items
        // touching styles/numbering/header/footer/textbox) creates and
        // re-opens many Update-mode zip parts from that background thread,
        // and the SDK's internal package state intermittently throws
        // "Entries cannot be opened multiple times in Update mode" — usually
        // at Dispose-time save when the package tries to commit all pending
        // parts. The `officecli batch` path doesn't hit this because items
        // are deserialized up front and replayed synchronously on the calling
        // thread. Mirror that here: buffer all JSONL lines first, then open
        // the handler and replay on this thread.
        var bufferedLines = new List<string>();

        try
        {
            void OnLine(string raw)
            {
                // Strip a per-line UTF-8 BOM (U+FEFF). Some Windows JSON
                // serializers emit BOM on every line of JSONL output, which
                // is technically RFC 8259 noncompliant but easy to absorb at
                // the host. Trim handles trailing whitespace and CR from a
                // CRLF-on-Windows plugin.
                var line = raw.TrimStart('﻿').Trim();
                if (line.Length == 0) return;

                // Reject legacy top-level JSON arrays explicitly. Plugins that
                // emitted `[...]` under the old protocol now fail with a clear
                // error instead of being parsed as a malformed BatchItem.
                if (line[0] == '[')
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' emitted a JSON array; protocol v1 requires JSONL (one BatchItem per line).")
                    { Code = "corrupt_batch" };

                // Buffer raw line; defer JSON parse + replay to the
                // main-thread loop after plugin exit. JSON parse errors
                // surface there with the same item-index semantics.
                bufferedLines.Add(line);
            }

            var idle = plugin.Manifest.ResolveIdleTimeout("dump");
            var result = PluginProcess.Run(new PluginProcess.RunOptions
            {
                ExecutablePath = plugin.ExecutablePath,
                Arguments = new[] { "dump", sourceFullPath },
                IdleTimeoutSeconds = idle,
                OnStdoutLine = OnLine,
            });

            // Bubble up the per-line callback error first — its message is more
            // actionable than the generic non-zero exit that follows.
            if (PluginProcess.LineCallbackError is CliException ce)
                throw ce;
            if (PluginProcess.LineCallbackError is not null)
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' replay aborted: {PluginProcess.LineCallbackError.Message}",
                    PluginProcess.LineCallbackError)
                { Code = "plugin_command_failed" };

            if (result.IdleTimedOut)
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' produced no output for {idle}s — likely hung.")
                {
                    Code = "plugin_idle_timeout",
                    Suggestion = $"Override with --timeout 0 or set a longer `idle_timeout_seconds.verbs.dump` in the plugin's manifest.",
                };

            if (result.ExitCode != 0)
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' failed (exit {result.ExitCode}): {Truncate(result.Stderr, 500)}")
                {
                    Code = result.ExitCode switch
                    {
                        2 => "corrupt_input",
                        3 => "unsupported_feature",
                        4 => "license_expired",
                        5 => "protocol_mismatch",
                        6 => "plugin_idle_timeout",
                        _ => "plugin_failed",
                    },
                };

            // v6.4: now that the plugin has exited and all JSONL is buffered,
            // open the handler on this thread and replay synchronously. See
            // the rationale comment at the bufferedLines declaration above
            // (OpenXml SDK package state not thread-safe under heavy
            // multi-part Update-mode mutation).
            using (var handler = DocumentHandlerFactory.Open(tmpOut, editable: true))
            {
                foreach (var line in bufferedLines)
                {
                    BatchItem? item;
                    try
                    {
                        item = JsonSerializer.Deserialize(line, BatchJsonContext.Default.BatchItem);
                    }
                    catch (JsonException ex)
                    {
                        throw new CliException(
                            $"Dump-reader plugin '{plugin.Manifest.Name}' emitted invalid JSON at item #{itemIndex}: {ex.Message}")
                        { Code = "plugin_contract_violation" };
                    }
                    if (item is null)
                        throw new CliException(
                            $"Dump-reader plugin '{plugin.Manifest.Name}' emitted null at item #{itemIndex}.")
                        { Code = "plugin_contract_violation" };

                    try
                    {
                        CommandBuilder.ExecuteBatchItem(handler, item, json: false);
                    }
                    catch (Exception ex)
                    {
                        throw new CliException(
                            $"Dump-reader plugin '{plugin.Manifest.Name}' command #{itemIndex} ({item.Command}) failed while replaying: {ex.Message}", ex)
                        { Code = "plugin_command_failed" };
                    }

                    itemIndex++;
                }
            }

            // Empty output + exit 0 is ambiguous: the .doc might genuinely be
            // blank, or the plugin might have silently skipped content it
            // does not yet know how to translate. Surface a warning so users
            // do not discover the empty conversion only by opening the
            // result. Console.Error is captured by callers — ResidentServer
            // wraps the handler-open call in a temporary Console.SetError
            // scope so this line reaches the first command's reply envelope.
            if (itemIndex == 0)
                Console.Error.WriteLine(
                    $"[warning] dump-reader plugin '{plugin.Manifest.Name}' produced no commands for {Path.GetFileName(sourceFullPath)}. " +
                    $"The generated {targetExt} will be blank — this is usually a plugin gap, not a source-file property.");
        }
        catch (Exception ex)
        {
            replayError = ex;
            try { File.Delete(tmpOut); } catch { }
            throw;
        }
        finally
        {
            // If we threw, the tmp file is already cleaned up above. If we
            // succeeded, the caller takes ownership. Either way, leave nothing
            // dangling on disk.
            if (replayError is null && !File.Exists(tmpOut))
                throw new CliException(
                    $"Dump-reader plugin '{plugin.Manifest.Name}' replay produced no output file.")
                { Code = "plugin_contract_violation" };
        }

        return new DumpResult(tmpOut, plugin);
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
