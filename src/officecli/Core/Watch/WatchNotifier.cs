// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// CONSISTENCY(watch-isolation): 本文件不引用 OfficeCli.Handlers,不打开文件,不写盘。
// 见 CLAUDE.md "Watch Server Rules"。要放宽这条红线,
// grep "CONSISTENCY(watch-isolation)" 找全 watch 子系统所有文件项目级一起评审。

using System.IO.Pipes;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Sends refresh notifications (with rendered HTML) to a running watch process.
/// Non-blocking, fire-and-forget. Silently does nothing if no watch is running.
/// All pipe I/O is bounded by a timeout to prevent hangs.
/// </summary>
internal static class WatchNotifier
{
    private static readonly TimeSpan PipeTimeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Notify watch with a pre-built message.
    /// The watch server never opens the file — all rendering is done by the caller.
    /// </summary>
    public static void NotifyIfWatching(string filePath, WatchMessage message)
    {
        try
        {
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(100); // fast fail if no watch

                var json = JsonSerializer.Serialize(message, WatchMessageJsonContext.Default.WatchMessage);

                // Write first, then read. Creating StreamReader before writing
                // causes a deadlock: StreamReader's constructor probes for BOM by
                // reading from the pipe, but the server is waiting for our write.
                using var writer = new StreamWriter(client, new UTF8Encoding(false), leaveOpen: true) { AutoFlush = true };
                writer.WriteLine(json);

                using var reader = new StreamReader(client, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                reader.ReadLine(); // wait for ack
            }, PipeTimeout);
        }
        catch
        {
            // No watch process running, or timed out — silently ignore
        }
    }

    /// <summary>
    /// Query the running watch process for the current selection.
    /// Returns:
    ///   null  → no watch running for this file (or pipe failure)
    ///   []    → watch is running but nothing is selected
    ///   [...] → list of currently-selected element paths
    /// </summary>
    public static string[]? QuerySelection(string filePath)
    {
        try
        {
            string[]? result = null;
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(200);

                var noBom = new UTF8Encoding(false);
                using var writer = new StreamWriter(client, noBom, leaveOpen: true) { AutoFlush = true };
                writer.WriteLine("get-selection");
                writer.Flush();

                using var reader = new StreamReader(client, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                var json = reader.ReadLine();
                if (json == null) { result = Array.Empty<string>(); return; }
                result = JsonSerializer.Deserialize(json, WatchSelectionJsonContext.Default.StringArray)
                         ?? Array.Empty<string>();
            }, PipeTimeout);
            return result;
        }
        catch
        {
            return null; // no watch running, or timed out
        }
    }

    // ==================== Marks ====================

    /// <summary>
    /// Add a mark to the running watch process. Returns the assigned id, or
    /// null if no watch is running. Throws if the request payload is rejected.
    ///
    /// The find string should be passed as-is. The CLI must wrap with r"..."
    /// when regex=true (mirroring WordHandler.Set's vocabulary).
    /// </summary>
    public static string? AddMark(string filePath, MarkRequest request)
    {
        // BUG-BT-001: distinguish "no watch running" from "watch rejected the
        // request". Pipe failures → return null so CLI prints "start watch first".
        // Server-side reject (Error field) → throw MarkRejectedException so CLI
        // surfaces the real error instead of silently treating empty id as success.
        string? result = null;
        string? error = null;
        try
        {
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(200);

                var noBom = new UTF8Encoding(false);
                using var writer = new StreamWriter(client, noBom, leaveOpen: true) { AutoFlush = true };
                var payload = JsonSerializer.Serialize(request, WatchMarkJsonContext.Default.MarkRequest);
                writer.WriteLine("mark " + payload);
                writer.Flush();

                using var reader = new StreamReader(client, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                var responseLine = reader.ReadLine();
                if (string.IsNullOrEmpty(responseLine)) { result = null; return; }
                var resp = JsonSerializer.Deserialize(responseLine, WatchMarkJsonContext.Default.MarkResponse);
                // BUG-FUZZER-R3-M01: use IsNullOrWhiteSpace for symmetry with the
                // server-side path/color validation. A whitespace-only error string
                // would otherwise spuriously throw MarkRejectedException.
                if (!string.IsNullOrWhiteSpace(resp?.Error)) { error = resp!.Error; return; }
                result = string.IsNullOrEmpty(resp?.Id) ? null : resp.Id;
            }, PipeTimeout);
        }
        catch
        {
            return null; // no watch running, or pipe failure
        }
        if (error != null) throw new MarkRejectedException(error);
        return result;
    }

    /// <summary>
    /// Remove marks from the running watch process. Returns count removed,
    /// or null if no watch is running.
    /// </summary>
    public static int? RemoveMarks(string filePath, UnmarkRequest request)
    {
        try
        {
            int? result = null;
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(200);

                var noBom = new UTF8Encoding(false);
                using var writer = new StreamWriter(client, noBom, leaveOpen: true) { AutoFlush = true };
                var payload = JsonSerializer.Serialize(request, WatchMarkJsonContext.Default.UnmarkRequest);
                writer.WriteLine("unmark " + payload);
                writer.Flush();

                using var reader = new StreamReader(client, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                var responseLine = reader.ReadLine();
                if (string.IsNullOrEmpty(responseLine)) { result = 0; return; }
                var resp = JsonSerializer.Deserialize(responseLine, WatchMarkJsonContext.Default.UnmarkResponse);
                result = resp?.Removed ?? 0;
            }, PipeTimeout);
            return result;
        }
        catch
        {
            return null; // no watch running
        }
    }

    /// <summary>
    /// Query all marks currently held by the watch process. Returns null if
    /// no watch is running, an empty array if the watch is running but no
    /// marks have been added, or the full list of marks otherwise.
    ///
    /// Thin wrapper over <see cref="QueryMarksFull"/> for callers that only
    /// care about the array. Use QueryMarksFull if you need the version.
    /// </summary>
    public static WatchMark[]? QueryMarks(string filePath)
    {
        var full = QueryMarksFull(filePath);
        return full?.Marks;
    }

    /// <summary>
    /// Query marks + monotonic version. Returns null if no watch is running.
    /// The version field lets callers CAS-style detect whether marks changed
    /// between two reads; the CLI's get-marks --json output surfaces this
    /// directly so AI consumers can cache without re-parsing.
    /// </summary>
    public static MarksResponse? QueryMarksFull(string filePath)
    {
        try
        {
            MarksResponse? result = null;
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(200);

                var noBom = new UTF8Encoding(false);
                using var writer = new StreamWriter(client, noBom, leaveOpen: true) { AutoFlush = true };
                writer.WriteLine("get-marks");
                writer.Flush();

                using var reader = new StreamReader(client, noBom, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                var json = reader.ReadLine();
                if (json == null) { result = new MarksResponse(); return; }
                result = JsonSerializer.Deserialize(json, WatchMarkJsonContext.Default.MarksResponse)
                         ?? new MarksResponse();
            }, PipeTimeout);
            return result;
        }
        catch
        {
            return null; // no watch running
        }
    }

    /// <summary>
    /// Send a close command to a running watch process.
    /// Returns true if the watch was successfully closed.
    /// </summary>
    public static bool SendClose(string filePath)
    {
        try
        {
            bool result = false;
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(200);

                // Write first, then read — same ordering as NotifyIfWatching
                // to avoid BOM-detection deadlock on the pipe.
                using var writer = new StreamWriter(client, new UTF8Encoding(false), leaveOpen: true) { AutoFlush = true };
                writer.WriteLine("close");
                writer.Flush();

                using var reader = new StreamReader(client, new UTF8Encoding(false), detectEncodingFromByteOrderMarks: false, leaveOpen: true);
                reader.ReadLine();
                result = true;
            }, PipeTimeout);
            return result;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Run an action on a background thread with a timeout.
    /// Prevents the calling thread from hanging if the pipe server dies mid-conversation.
    /// </summary>
    private static void RunWithTimeout(Action action, TimeSpan timeout)
    {
        var task = Task.Run(action);
        if (!task.Wait(timeout))
            throw new TimeoutException("Pipe communication timed out");
        task.GetAwaiter().GetResult(); // propagate exceptions
    }
}

/// <summary>
/// Message sent from command processes to the watch server via named pipe.
/// </summary>
internal class WatchMessage
{
    /// <summary>"replace", "add", "remove", or "full"</summary>
    public string Action { get; set; } = "full";

    /// <summary>Slide number (0 for full refresh)</summary>
    public int Slide { get; set; }

    /// <summary>Single slide HTML fragment (for replace/add)</summary>
    public string? Html { get; set; }

    /// <summary>Full HTML of the entire presentation (for caching by watch server)</summary>
    public string? FullHtml { get; set; }

    /// <summary>CSS selector for the element to scroll to after full refresh (Word/Excel)</summary>
    public string? ScrollTo { get; set; }

    /// <summary>Incremental version number for ordering and gap detection.</summary>
    public int Version { get; set; }

    /// <summary>Version the client must have before applying these patches.</summary>
    public int BaseVersion { get; set; }

    /// <summary>Word block-level patches (for action="word-patch").</summary>
    public List<WordPatch>? Patches { get; set; }

    public static int ExtractSlideNum(string? path)
    {
        if (string.IsNullOrEmpty(path)) return 0;
        var match = System.Text.RegularExpressions.Regex.Match(path, @"/slide\[(\d+)\]");
        if (match.Success && int.TryParse(match.Groups[1].Value, out var num))
            return num;
        return 0;
    }

    /// <summary>Extract a CSS selector scroll target from a Word document path like /p[5] or /table[2].</summary>
    public static string? ExtractWordScrollTarget(string? path)
    {
        if (string.IsNullOrEmpty(path)) return null;
        var match = System.Text.RegularExpressions.Regex.Match(path, @"/(p|paragraph|table)\[(\d+)\]");
        if (!match.Success) return null;
        var type = match.Groups[1].Value;
        if (type == "paragraph") type = "p";
        return $"#w-{type}-{match.Groups[2].Value}";
    }

    /// <summary>Extract sheet name from an Excel document path like /Sheet1/A1 or Sheet1!A1.</summary>
    public static string? ExtractSheetName(string? path)
    {
        if (string.IsNullOrEmpty(path)) return null;
        // Match /SheetName/... or SheetName!...
        var match = System.Text.RegularExpressions.Regex.Match(path, @"^/?([^/!]+)[/!]");
        return match.Success ? match.Groups[1].Value : null;
    }
}

/// <summary>A single block-level change for Word incremental updates.</summary>
internal class WordPatch
{
    /// <summary>"replace", "add", or "remove"</summary>
    public string Op { get; set; } = "";

    /// <summary>Block number (matches <!--wB:N--> marker)</summary>
    public int Block { get; set; }

    /// <summary>New HTML content (null for remove)</summary>
    public string? Html { get; set; }
}

[System.Text.Json.Serialization.JsonSerializable(typeof(WatchMessage))]
[System.Text.Json.Serialization.JsonSerializable(typeof(WordPatch))]
internal partial class WatchMessageJsonContext : System.Text.Json.Serialization.JsonSerializerContext { }

/// <summary>
/// Request body for POST /api/selection — list of currently selected element paths.
/// </summary>
internal class SelectionRequest
{
    [System.Text.Json.Serialization.JsonPropertyName("paths")]
    public List<string>? Paths { get; set; }
}

[System.Text.Json.Serialization.JsonSerializable(typeof(SelectionRequest))]
[System.Text.Json.Serialization.JsonSerializable(typeof(string[]))]
internal partial class WatchSelectionJsonContext : System.Text.Json.Serialization.JsonSerializerContext { }

/// <summary>
/// Selection-side mirror of <see cref="WatchMarkJsonOptions"/>: same
/// UnsafeRelaxedJsonEscaping relaxation. Selection paths are usually ASCII
/// today but future path schemes may carry CJK or symbols (e.g. path
/// predicates referencing element text), so keep the two sides in sync.
/// </summary>
internal static class WatchSelectionJsonOptions
{
    public static readonly System.Text.Json.JsonSerializerOptions Relaxed =
        new(WatchSelectionJsonContext.Default.Options)
        {
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        };

    public static readonly System.Text.Json.Serialization.Metadata.JsonTypeInfo<string[]> StringArrayInfo =
        (System.Text.Json.Serialization.Metadata.JsonTypeInfo<string[]>)Relaxed.GetTypeInfo(typeof(string[]));
}
