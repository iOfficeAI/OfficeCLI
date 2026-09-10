// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Security.Cryptography;
using System.Text.Json;
using OfficeCli;

namespace OfficeCli.Core;

/// <summary>
/// Crash-resume ledger for non-resident batches (`batch --resume`).
///
/// The journal is an append-only JSONL file next to the target document
/// (".{stem}.batch-journal-{timestamp}.jsonl"): a header line, one line with
/// the full item payload list (so --resume needs no source re-supply), then
/// one line per finished item — each appended with a flush so a kill -9
/// leaves a readable prefix. A torn trailing line is ignored on read. JSON is
/// written via Utf8JsonWriter / source-gen context and read via JsonDocument
/// — no reflection, safe under publish trimming.
///
/// Lifecycle invariant: the journal is DELETED when a batch run completes
/// normally (committed, rolled back, or best-effort finished-with-failures —
/// the caller saw the verdict either way). A journal that still exists
/// therefore means exactly one thing: the process died mid-run.
///
/// Resume semantics (implemented in CommandBuilder.Batch):
///   - atomic-mode interruption: the original file was never touched (the
///     temp copy is orphan-swept), so resume re-runs EVERY item; the
///     sourceHash guard refuses when the file changed since the interrupted
///     run started (external edits would be stomped) unless forced.
///   - best-effort interruption: ok items are already applied in place, so
///     resume re-runs only failed/skipped/never-started items as a fresh
///     atomic batch. No hash guard — the statuses describe what ran; the
///     resume run itself carries the normal atomic protection.
/// </summary>
internal static class BatchJournal
{
    /// <summary>Journal path for a target document: same directory, dotted prefix.</summary>
    public static string PathFor(string targetPath)
    {
        var stem = System.IO.Path.GetFileNameWithoutExtension(targetPath);
        if (stem.Length > 40) stem = stem[..40];
        return System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(targetPath) ?? ".",
            $".{stem}.batch-journal-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N")[..6]}.jsonl");
    }

    /// <summary>Newest active (interrupted) journal for a target document, or null.</summary>
    public static string? FindLatest(string targetPath)
    {
        var dir = System.IO.Path.GetDirectoryName(targetPath) ?? ".";
        var stem = System.IO.Path.GetFileNameWithoutExtension(targetPath);
        if (stem.Length > 40) stem = stem[..40];
        try
        {
            var journals = System.IO.Directory.GetFiles(dir, $".{stem}.batch-journal-*.jsonl");
            return journals.Length == 0
                ? null
                : journals.OrderByDescending(System.IO.File.GetLastWriteTimeUtc).First();
        }
        catch { return null; }
    }

    /// <summary>SHA-256 of the target at batch start. Empty when the file
    /// cannot be read (a live resident holds it exclusively — the resident
    /// batch route): the journal then records no hash and the resume guard
    /// is skipped for that journal.</summary>
    public static string ComputeSourceHash(string targetPath)
    {
        try
        {
            using var stream = System.IO.File.OpenRead(targetPath);
            return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return "";
        }
    }

    public static BatchJournalWriter Start(string journalPath, string targetPath, string sourceHash, int total, string mode) =>
        new(journalPath, writer =>
        {
            writer.WriteStartObject();
            writer.WriteString("type", "header");
            writer.WriteNumber("v", 1);
            writer.WriteString("startedAt", DateTime.UtcNow.ToString("o"));
            writer.WriteString("file", targetPath);
            writer.WriteString("sourceHash", sourceHash);
            writer.WriteNumber("total", total);
            writer.WriteString("mode", mode);
            writer.WriteEndObject();
        });

    /// <summary>
    /// Fold a journal file into its recorded state. Torn/malformed lines
    /// (kill -9 mid-append) are skipped; a journal without a header yields
    /// null (nothing recoverable).
    /// </summary>
    public static BatchJournalState? Read(string journalPath)
    {
        if (!System.IO.File.Exists(journalPath)) return null;
        BatchJournalState? state = null;
        foreach (var raw in System.IO.File.ReadLines(journalPath))
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;
            JsonDocument doc;
            try { doc = JsonDocument.Parse(raw); }
            catch { continue; /* torn trailing line */ }
            using (doc)
            {
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object) continue;
                if (!root.TryGetProperty("type", out var typeEl)) continue;
                switch (typeEl.GetString())
                {
                    case "header":
                        state = new BatchJournalState
                        {
                            Path = journalPath,
                            SourceHash = root.TryGetProperty("sourceHash", out var hashEl) ? hashEl.GetString() ?? "" : "",
                            Total = root.TryGetProperty("total", out var totalEl) ? totalEl.GetInt32() : 0,
                            Mode = root.TryGetProperty("mode", out var modeEl) ? modeEl.GetString() ?? "atomic" : "atomic",
                        };
                        break;
                    case "items" when state != null && root.TryGetProperty("items", out var itemsEl):
                        try
                        {
                            state.AllItems = JsonSerializer.Deserialize(
                                itemsEl.GetRawText(), OfficeCli.BatchJsonContext.Default.ListBatchItem) ?? [];
                        }
                        catch { /* torn items line — resume reports nothing to re-run */ }
                        break;
                    case "item" when state != null:
                        if (root.TryGetProperty("index", out var idxEl))
                            state.Items[idxEl.GetInt32()] =
                                root.TryGetProperty("status", out var stEl) ? stEl.GetString() ?? "pending" : "pending";
                        break;
                }
            }
        }
        return state;
    }
}

/// <summary>Append-only journal writer; every line is flushed immediately (kill-safe).</summary>
internal sealed class BatchJournalWriter : IDisposable
{
    private readonly string _path;
    private readonly System.IO.StreamWriter _writer;

    internal BatchJournalWriter(string path, Action<Utf8JsonWriter>? writeHeader)
    {
        _path = path;
        _writer = new System.IO.StreamWriter(path, append: writeHeader == null) { AutoFlush = true };
        if (writeHeader != null) RawLine(writeHeader);
    }

    /// <summary>Append to an existing journal (the resident side of a
    /// CLI-created journal: the CLI wrote header+items, the resident appends
    /// item lines and completes).</summary>
    public static BatchJournalWriter Attach(string journalPath) => new(journalPath, writeHeader: null);

    /// <summary>Second journal line: the full item payload list for --resume.</summary>
    public void WriteItems(List<BatchItem> items) =>
        RawLine(writer =>
        {
            writer.WriteStartObject();
            writer.WriteString("type", "items");
            writer.WritePropertyName("items");
            JsonSerializer.Serialize(writer, items, BatchJsonContext.Default.ListBatchItem);
            writer.WriteEndObject();
        });

    /// <summary>Append one item-completion line.</summary>
    public void ItemDone(int index, string status, string? code, string? errorPreview) =>
        RawLine(writer =>
        {
            writer.WriteStartObject();
            writer.WriteString("type", "item");
            writer.WriteNumber("index", index);
            writer.WriteString("status", status);
            if (code != null) writer.WriteString("code", code);
            if (errorPreview != null)
            {
                if (errorPreview.Length > 160) errorPreview = errorPreview[..160];
                writer.WriteString("errorPreview", errorPreview);
            }
            writer.WriteEndObject();
        });

    private void RawLine(Action<Utf8JsonWriter> write)
    {
        using var buffer = new System.IO.MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            write(writer);
            writer.Flush();
        }
        _writer.WriteLine(System.Text.Encoding.UTF8.GetString(buffer.ToArray()));
    }

    /// <summary>Delete the journal — the run completed and the caller saw its verdict.</summary>
    public void Complete()
    {
        Dispose();
        try { System.IO.File.Delete(_path); } catch { /* best-effort */ }
    }

    public void Dispose() => _writer.Dispose();
}

internal class BatchJournalState
{
    public string Path { get; set; } = "";
    public string SourceHash { get; set; } = "";
    public int Total { get; set; }
    public string Mode { get; set; } = "atomic";
    /// <summary>The original batch items, recorded up front for --resume.</summary>
    public List<BatchItem> AllItems { get; set; } = [];
    /// <summary>index → recorded status ("ok" / "failed" / "skipped").</summary>
    public Dictionary<int, string> Items { get; } = new();
}
