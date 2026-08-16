// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Simple file logger. Enabled via: officecli config log true
/// Logs to ~/.officecli/officecli.log (max 1 MB, auto-trimmed)
/// </summary>
internal static class CliLogger
{
    private static readonly string LogPath = Path.Combine(UpdateChecker.LegacyStateDir, "officecli.log");
    private const long MaxLogSize = 1024 * 1024; // 1 MB

    internal static bool Enabled
    {
        get
        {
            try { return UpdateChecker.LoadConfig().Log; }
            catch { return false; }
        }
    }

    internal static void LogCommand(string[] args)
    {
        if (!Enabled || args.Length == 0) return;
        // Skip internal commands
        if (args[0].StartsWith("__") && args[0].EndsWith("__")) return;
        Write($"> officecli {string.Join(" ", args)}");
    }

    internal static void Clear()
    {
        try { File.Delete(LogPath); }
        catch { }
    }

    internal static void LogOutput(string output)
    {
        if (!Enabled || string.IsNullOrEmpty(output)) return;
        Write(output);
    }

    internal static void LogError(string error)
    {
        if (!Enabled || string.IsNullOrEmpty(error)) return;
        Write($"[ERROR] {error}");
    }

    private static void Write(string message)
    {
        try
        {
            Directory.CreateDirectory(UpdateChecker.LegacyStateDir);

            var escaped = message.ReplaceLineEndings("\\n");
            TrimIfNeeded();
            File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {escaped}\n");
        }
        catch
        {
            // Logging should never break the CLI
        }
    }

    private static void TrimIfNeeded()
    {
        var info = new FileInfo(LogPath);
        if (!info.Exists || info.Length <= MaxLogSize) return;

        // Keep the last half of the file
        var text = File.ReadAllText(LogPath);
        var half = text.Length / 2;
        var start = text.IndexOf('\n', half);
        if (start < 0 || start >= text.Length - 1) return;
        File.WriteAllText(LogPath, text[(start + 1)..]);
    }
}
