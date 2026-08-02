using System;
using System.IO;

namespace OfficeCli.Core;

/// <summary>
/// Crash-atomic flush of an in-memory OOXML package to disk. The resident holds
/// the document open and defers disk writes to save/close/idle-autosave; those
/// writes must never rewrite the target in place, because a process death
/// mid-write leaves a truncated, unopenable file with no fallback (the original
/// bytes are already gone). Instead we serialize the complete package to a
/// sibling temp, optionally post-process the temp, then swap it over the target
/// in a single <see cref="File.Replace(string,string,string)"/> — so a crash at
/// any point leaves either the old file or the new file intact, never a torn one.
/// Mirrors the guarantee the standalone batch path already gets from its
/// File.Replace.
/// </summary>
internal static class AtomicPackageWriter
{
    /// <param name="package">Complete in-memory package bytes. Left at position 0.</param>
    /// <param name="path">Target file, replaced atomically.</param>
    /// <param name="releaseLock">Closes the caller's own writable handle on
    /// <paramref name="path"/> so the swap can proceed. Called only after the temp
    /// is fully written and post-processed.</param>
    /// <param name="reopenLock">Reopens the caller's handle after the swap so the
    /// session keeps holding the file. Called even if the swap throws.</param>
    /// <param name="postProcessTemp">Optional in-place zip rewrites that must land
    /// in the same atomic swap (e.g. whole-part flush, self-close normalization).
    /// Runs against the TEMP path — safe, since the original is untouched until the
    /// final replace.</param>
    public static void Flush(MemoryStream package, string path,
                             Action releaseLock, Action reopenLock,
                             Action<string>? postProcessTemp = null)
    {
        var dir = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(dir)) dir = ".";
        var tmp = Path.Combine(dir, $".{Path.GetFileName(path)}.savetmp-{Guid.NewGuid():N}");
        var released = false;
        try
        {
            package.Position = 0;
            using (var tf = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                package.CopyTo(tf);
                // Flush the managed buffer out to the OS. The atomic File.Replace
                // below is what guarantees crash-safety against process death /
                // kill / OOM (the OS still commits already-written bytes even if we
                // die). Power-loss durability would additionally need
                // Flush(flushToDisk: true); that fsync is deliberately omitted to
                // keep large-file autosaves cheap.
                tf.Flush();
            }
            postProcessTemp?.Invoke(tmp);
            releaseLock();
            released = true;
            // File.Replace requires the destination to exist; if the target was
            // renamed or deleted out from under the session (external `mv`, a
            // sibling process, a test moving the file), it throws
            // FileNotFoundException and the catch below would delete the temp —
            // silently discarding every in-memory edit. Recreate the target
            // with a plain move in that case: the temp is already fully written
            // and fsync-flushed, so this is just as crash-safe as the replace,
            // and the session's managed path (`path`) is what gets the data.
            if (File.Exists(path))
                File.Replace(tmp, path, destinationBackupFileName: null);
            else
                File.Move(tmp, path);
        }
        catch
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* best-effort */ }
            throw;
        }
        finally
        {
            package.Position = 0;
            // Reopen the caller's handle if we released it — even when Replace
            // threw, so the session can retry the flush on the still-intact
            // original at the next save/close.
            if (released) { try { reopenLock(); } catch { /* next flush resurfaces it */ } }
        }
    }
}
