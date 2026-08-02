// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// On Unix, .NET named pipes are Unix domain sockets created as
/// $TMPDIR/CoreFxPipe_&lt;pipeName&gt;. The kernel caps the socket path
/// (sun_path) at 104 bytes on macOS/BSD and 108 on Linux. Agent sandboxes
/// commonly export a nested TMPDIR (e.g. /var/folders/.../T/.ctx-XXXX)
/// long enough that the resident/watch pipes exceed the cap — the resident
/// then either dies with ArgumentOutOfRangeException or, worse, binds its
/// main pipe but not the longer "-ping" pipe, leaving an orphan server the
/// client reports as "started but not responding" (issue #263).
///
/// Fix: once at process startup, if the worst-case socket path would not
/// fit, repoint TMPDIR at a short per-user directory. Every pipe endpoint
/// (resident client/server, watch server/notifier) and the .lock/.port
/// files key off Path.GetTempPath(), and spawned children inherit the
/// variable, so all parties converge on the same directory by running the
/// same deterministic rule.
/// </summary>
internal static class PipeTempDirGuard
{
    // Longest socket file name we ever create:
    //   "CoreFxPipe_" (11) + "officecli-watch-<16 hex>" (32) = 43.
    // Resident pipes are shorter: "officecli-<16 hex>" (26) / +"-ping" (31).
    private const int MaxSocketFileName = 43;

    /// <summary>
    /// Pure decision: returns the short fallback directory to repoint TMPDIR
    /// at, or null when the current temp path already fits every socket we
    /// create. Split out so the length rule is unit-testable without
    /// mutating process-wide environment state.
    /// </summary>
    internal static string? ComputeFallbackDir(string tempPath, bool isLinux, uint uid)
    {
        // sun_path cap: 108 bytes on Linux, 104 on macOS/BSD. Measured in
        // bytes (kernel view) rather than chars — strictly more conservative
        // than .NET's own char-count check, so we can never normalize later
        // than .NET would fail.
        int sunPathMax = isLinux ? 108 : 104;

        // GetTempPath() ends with '/'; sockets are created directly inside it.
        if (System.Text.Encoding.UTF8.GetByteCount(tempPath) + MaxSocketFileName <= sunPathMax)
            return null;

        return $"/tmp/officecli-{uid}";
    }

    public static void EnsurePipePathFits()
    {
        if (OperatingSystem.IsWindows())
            return; // \\.\pipe\ namespace, no path-length constraint

        var fallback = ComputeFallbackDir(Path.GetTempPath(), OperatingSystem.IsLinux(), GetUid());
        if (fallback == null)
            return;
        try
        {
            Directory.CreateDirectory(fallback);
            // 0700 — and an EPERM here means the name was squatted by another
            // user, in which case we must not share sockets with them: bail
            // and keep the original TMPDIR (same failure surface as before).
            File.SetUnixFileMode(fallback,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
        catch
        {
            return;
        }

        Environment.SetEnvironmentVariable("TMPDIR", fallback);
    }

    [System.Runtime.InteropServices.DllImport("libc", EntryPoint = "getuid")]
    private static extern uint GetUid();
}
