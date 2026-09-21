// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Security.Cryptography;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Best-effort crash marker for resident mutations that have not reached disk.
/// This does not replay edits; it prevents a later process from falsely
/// reporting that the stale on-disk document is already saved.
/// </summary>
internal static class ResidentRecoveryMarker
{
    private const string Warning =
        "A previous resident ended while it may have held unflushed in-memory changes. " +
        "Those changes cannot be recovered and may have been lost; the file was reopened " +
        "from its last saved state.";

    internal static string WarningMessage(string filePath)
        => $"WARNING: {Warning} File: {Path.GetFileName(filePath)}. " +
           "For short-lived or externally managed agent processes, set " +
           "OFFICECLI_RESIDENT_FLUSH=each.";

    internal static CliException CreateLossException(string filePath)
        => new(WarningMessage(filePath))
        {
            Code = "resident_unflushed_changes_lost",
            Suggestion = "Repeat the lost edit if needed, then use OFFICECLI_RESIDENT_FLUSH=each."
        };

    internal static bool TryMark(string filePath, out string? error)
    {
        error = null;
        string? tempPath = null;
        try
        {
            var path = MarkerPath(filePath);
            var dir = Path.GetDirectoryName(path)!;
            Directory.CreateDirectory(dir);
            TryRestrictDirectory(dir);

            tempPath = $"{path}.{Environment.ProcessId}.{Guid.NewGuid():N}.tmp";
            using var owner = System.Diagnostics.Process.GetCurrentProcess();
            var started = owner.StartTime.ToUniversalTime().Ticks;
            var payload = $"v2\t{Environment.ProcessId}\t{started}\n";
            using (var stream = new FileStream(
                tempPath, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                bufferSize: 4096, FileOptions.WriteThrough))
            {
                var bytes = Encoding.UTF8.GetBytes(payload);
                stream.Write(bytes);
                stream.Flush(flushToDisk: true);
            }
            TryRestrictFile(tempPath);
            File.Move(tempPath, path, overwrite: true);
            tempPath = null;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
        finally
        {
            if (tempPath != null)
                try { File.Delete(tempPath); } catch { }
        }
    }

    internal static void Clear(string filePath)
    {
        try { File.Delete(MarkerPath(filePath)); } catch { }
    }

    // A failed pipe probe is not proof of exit. Hold the same singleton lock
    // as __resident-serve__ across inspection/deletion, excluding both an old
    // owner and a new session that could otherwise replace the marker.
    internal static bool TryConsumeAfterExit(string filePath)
    {
        FileStream residentLock;
        try
        {
            residentLock = new FileStream(ResidentServer.GetLockPath(filePath),
                FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None,
                bufferSize: 1, FileOptions.DeleteOnClose);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw CreateUnknownStateException(filePath);
        }
        using (residentLock)
            return TryConsume(filePath);
    }

    // Caller must hold the resident singleton lock. Startup already holds it
    // before constructing ResidentServer; save/close acquire it above.
    internal static bool TryConsume(string filePath)
    {
        var path = MarkerPath(filePath);
        string payload;
        try { payload = File.ReadAllText(path); }
        catch (FileNotFoundException) { return false; }
        catch (DirectoryNotFoundException) { return false; }
        catch { throw CreateUnknownStateException(filePath); }
        // Markers outlive a temp directory / IPC namespace. A live recorded
        // PID may own a resident under a different TMPDIR, so the lock alone
        // is not sufficient. Legacy markers have only a PID; v2 also checks
        // process start time so a recycled PID is not mistaken for the writer.
        // Never infer loss from an unreadable process or malformed marker.
        if (!WriterHasExited(payload)) throw CreateUnknownStateException(filePath);
        try { File.Delete(path); } catch { /* repeat the warning next time */ }
        return true;
    }

    private static bool WriterHasExited(string payload)
    {
        var fields = payload.TrimEnd('\r', '\n').Split('\t');
        if (fields.Length != 3 || (fields[0] != "v1" && fields[0] != "v2")
            || !int.TryParse(fields[1], out var pid) || pid <= 0)
            return false;
        long started = 0;
        if (fields[0] == "v2" && (!long.TryParse(fields[2], out started)
            || started <= 0 || started > DateTime.MaxValue.Ticks))
            return false;
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(pid);
            return process.HasExited || (fields[0] == "v2"
                && process.StartTime.ToUniversalTime().Ticks != started);
        }
        catch (ArgumentException) { return true; } // No process with this PID.
        catch { return false; } // Unknown is not dead.
    }

    internal static string ReadOnlyWarningMessage(string filePath)
        => $"WARNING: Resident recovery state for {Path.GetFileName(filePath)} could not be verified. " +
           "This session is read-only and shows the last saved file, not any pending edits " +
           "in another resident. The recovery marker is retained; no loss is confirmed. " +
           "Close this session and resolve the previous resident or its recovery marker before reopening to edit.";

    internal static CliException CreateReadOnlyException(string filePath)
        => new(ReadOnlyWarningMessage(filePath))
        {
            Code = "resident_state_unknown",
            Suggestion = "Read-only commands and close remain available. Preserve the recovery marker " +
                         "until the previous writer's state is resolved; do not repeat possibly pending edits."
        };

    private static CliException CreateUnknownStateException(string filePath)
        => new($"Resident state for {Path.GetFileName(filePath)} could not be verified. " +
               "The resident may still hold unsaved changes; no recovery marker was consumed.")
        {
            Code = "resident_state_unknown",
            Suggestion = "Retry save or close when the resident responds. Do not repeat edits based on this error."
        };

    private static string MarkerPath(string filePath)
    {
        var canonical = PathIdentity.Canonical(filePath);
        if (OperatingSystem.IsWindows() || OperatingSystem.IsMacOS())
            canonical = canonical.ToUpperInvariant();
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical)))[..24];
        var root = Path.Combine(UpdateChecker.ConfigDir, "resident-recovery");
        return Path.Combine(root, $"{hash}.dirty");
    }

    private static void TryRestrictDirectory(string path)
    {
        if (OperatingSystem.IsWindows()) return;
        try
        {
            File.SetUnixFileMode(path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
        catch { }
    }

    private static void TryRestrictFile(string path)
    {
        if (OperatingSystem.IsWindows()) return;
        try { File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.UserWrite); }
        catch { }
    }
}
