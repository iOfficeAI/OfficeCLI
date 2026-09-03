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
            var payload = $"v1\t{Environment.ProcessId}\t{DateTimeOffset.UtcNow:O}\n";
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

    internal static bool TryConsume(string filePath)
    {
        var path = MarkerPath(filePath);
        if (!File.Exists(path)) return false;
        try { File.Delete(path); } catch { /* repeat the warning next time */ }
        return true;
    }

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
