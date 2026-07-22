// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;

namespace OfficeCli.Core;

/// <summary>
/// Installs officecli binary, skills, and MCP (for tools without skill support).
/// Usage:
///   officecli install [target]  — install binary + skills + fallback MCP
/// </summary>
internal static class Installer
{
    private static readonly string BinDir = OperatingSystem.IsWindows()
        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeCli")
        : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin");

    private static readonly string TargetPath = Path.Combine(BinDir,
        OperatingSystem.IsWindows() ? "officecli.exe" : "officecli");

    /// <summary>Canonical install location of the officecli binary
    /// (<c>~/.local/bin/officecli</c> on Unix, <c>%LOCALAPPDATA%\OfficeCli</c>
    /// on Windows). External registrations (MCP, etc.) should record this path
    /// rather than <see cref="Environment.ProcessPath"/> so the command survives
    /// upgrades — self-install overwrites this file in place.</summary>
    internal static string InstalledBinaryPath
    {
        get
        {
            // If we're running from a directory already on PATH, that location IS
            // the install — we don't duplicate into the canonical dir (see
            // InstallBinary). Record the reachable path so external registrations
            // (MCP, etc.) point at the file upgrades overwrite in place.
            var src = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(src) && IsDirOnPath(Path.GetDirectoryName(src)))
                return src;
            return TargetPath;
        }
    }

    /// <summary>
    /// MCP targets and the skill aliases that overlap with them.
    /// If any of the skill aliases were installed, skip MCP for that target.
    /// </summary>
    private static readonly (string McpTarget, string DetectDir, string[] SkillAliases)[] McpTargets =
    [
        ("claude", ".claude",                          ["claude", "claude-code"]),
        ("cursor", ".cursor",                          ["cursor"]),
        ("vscode", ".vscode",                          []),   // no skill equivalent
        ("lms",    ".cache/lm-studio",                 []),   // no skill equivalent
    ];

    public static int Run(string[] args)
    {
        InstallBinary();

        var target = args.Length >= 1 ? args[0] : "all";

        // Skip the skill phase when the target is MCP-only (vscode, lms).
        // SkillInstaller has no equivalent agent for these and would otherwise
        // print a misleading 'Unknown target' to stderr before InstallMcpFallback
        // succeeds. The skill/MCP target namespaces are deliberately allowed to
        // diverge — McpTargets with empty SkillAliases is the source of truth
        // for "no skill phase needed".
        var isMcpOnly = McpTargets.Any(t =>
            t.SkillAliases.Length == 0 &&
            t.McpTarget.Equals(target, StringComparison.OrdinalIgnoreCase));
        var skilledTools = isMcpOnly
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : SkillInstaller.Install(target);

        // Install MCP for tools that didn't get a skill
        var mcpInstalled = InstallMcpFallback(skilledTools, target);

        // Exit 1 when a specific target was named but neither skills nor MCP
        // recognized it. 'all' (default) is always success because there's
        // nothing to mistype. Without this, `officecli install bogus` would
        // exit 0 after only printing 'Unknown target' to stderr — automation
        // can't distinguish a typo from a successful install.
        var isAll = target.Equals("all", StringComparison.OrdinalIgnoreCase);
        if (!isAll && skilledTools.Count == 0 && !mcpInstalled)
            return 1;
        return 0;
    }

    private static bool InstallMcpFallback(HashSet<string> skilledTools, string target)
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var isAll = target.Equals("all", StringComparison.OrdinalIgnoreCase);
        var anyInstalled = false;

        foreach (var (mcpTarget, detectDir, skillAliases) in McpTargets)
        {
            // If targeting a specific tool, only process matching MCP target
            if (!isAll && !mcpTarget.Equals(target, StringComparison.OrdinalIgnoreCase))
                continue;

            // Skip if skill was already installed for this tool
            if (skillAliases.Any(a => skilledTools.Contains(a)))
                continue;

            // Only install if the tool's directory exists
            if (Directory.Exists(Path.Combine(home, detectDir)))
            {
                if (McpInstaller.Install(mcpTarget))
                    anyInstalled = true;
            }
        }

        return anyInstalled;
    }

    internal static bool InstallBinary(bool quiet = false)
    {
        var src = Environment.ProcessPath;
        if (string.IsNullOrEmpty(src))
            return false;

        // Already at target location — record version and skip the copy
        var pathComparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (string.Equals(Path.GetFullPath(src), Path.GetFullPath(TargetPath), pathComparison))
        {
            RecordInstalledVersion();
            return false;
        }

        // Already reachable on PATH from where it's running — a user-managed
        // install (added to PATH manually, or by install.ps1 into a pre-existing
        // on-PATH dir). Don't drop a second copy into the canonical dir; treat the
        // on-PATH location as the install. Mirrors install.ps1, which upgrades an
        // existing on-PATH copy in place rather than relocating it. Config and
        // plugins live in ~/.officecli regardless, so nothing else moves.
        if (IsDirOnPath(Path.GetDirectoryName(src)))
        {
            RecordInstalledVersion();
            return false;
        }

        // Skip binary copy when managed by a package manager (Homebrew, etc.)
        if (src.Contains("/Caskroom/") || src.Contains("/Cellar/"))
        {
            if (!quiet)
                Console.WriteLine("Skipping binary install: managed by Homebrew.");
            RecordInstalledVersion();
            return false;
        }

        // Skip if not a self-contained published binary (e.g. running via dotnet run)
        // Self-contained single-file binaries are typically >5MB; framework-dependent builds are <1MB
        var srcInfo = new FileInfo(src);
        if (srcInfo.Length < 5 * 1024 * 1024)
        {
            if (!quiet)
            {
                Console.WriteLine($"Skipping binary install: not a published self-contained binary.");
                Console.WriteLine($"  Run: dotnet publish -c Release -r <rid> --self-contained -p:PublishSingleFile=true");
            }
            return false;
        }

        Directory.CreateDirectory(BinDir);

        // Stage beside the target, then swap by rename — never copy ONTO the
        // live path. File.Copy(overwrite) truncates the destination inode in
        // place, so every process currently executing that binary (MCP
        // servers, residents, watch) has its text pages pulled out from under
        // it and dies (observed downstream: ~10 resident officecli processes
        // killed, exit 137, when a new build was cp'd over a shared install).
        // A rename leaves the old inode intact for as long as those processes
        // hold it open, and the new binary is visible atomically to the next
        // exec. Mirrors UpdateChecker.TryApplyPendingUpdate's swap, including
        // the move-aside step (Windows cannot rename ONTO a running image,
        // but can rename the running image away).
        var stagePath = TargetPath + ".new-" + Guid.NewGuid().ToString("N")[..8];
        File.Copy(src, stagePath, overwrite: true);

        // Set the mode on the staged file so the binary is never momentarily
        // present at TargetPath without its executable bit.
        if (!OperatingSystem.IsWindows())
        {
            try
            {
                File.SetUnixFileMode(stagePath,
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |
                    UnixFileMode.GroupRead | UnixFileMode.GroupExecute |
                    UnixFileMode.OtherRead | UnixFileMode.OtherExecute);
            }
            catch { /* best effort */ }
        }

        var asidePath = TargetPath + ".old";
        var movedAside = false;
        if (File.Exists(TargetPath))
        {
            try { File.Delete(asidePath); } catch { /* best effort */ }
            try { File.Move(TargetPath, asidePath, overwrite: true); movedAside = true; }
            catch { /* target may not be movable; the swap below still tries */ }
        }
        try
        {
            File.Move(stagePath, TargetPath, overwrite: true);
        }
        catch
        {
            if (movedAside)
                try { File.Move(asidePath, TargetPath, overwrite: true); } catch { /* best effort */ }
            try { File.Delete(stagePath); } catch { /* best effort */ }
            throw;
        }
        try { File.Delete(asidePath); } catch { /* best effort */ }

        RecordInstalledVersion();

        if (quiet)
            Console.Error.WriteLine($"note: officecli self-installed to {TargetPath}");
        else
            Console.WriteLine($"Installed binary to {TargetPath}");

        EnsurePath(quiet);
        return true;
    }

    private static void RecordInstalledVersion()
    {
        try
        {
            var current = UpdateChecker.GetCurrentVersionPublic();
            if (string.IsNullOrEmpty(current)) return;
            var config = UpdateChecker.LoadConfig();
            if (config.InstalledBinaryVersion == current) return;
            config.InstalledBinaryVersion = current;
            UpdateChecker.SaveConfig(config);
        }
        catch { /* best effort */ }
    }

    /// <summary>
    /// Auto-install hook called on every officecli invocation.
    /// - Target missing → full install (binary + skills + MCP fallback).
    /// - Target older than current → binary-only upgrade.
    /// - Otherwise → no-op (cheap path: one File.Exists + one config read).
    /// Never throws, never blocks the main command.
    /// </summary>
    internal static void MaybeAutoInstall(string[] args)
    {
        try
        {
            // Opt-out
            if (Environment.GetEnvironmentVariable("OFFICECLI_NO_AUTO_INSTALL") == "1")
                return;

            // Only trigger on bare `officecli` invocation (exploratory / discovery call).
            // Real work commands (view, set, add, create, ...) are left alone to keep
            // zero side-effects and zero overhead on the hot path.
            if (args.Length != 0)
                return;

            var src = Environment.ProcessPath;
            if (string.IsNullOrEmpty(src)) return;

            // Already reachable — running from the canonical dir, or from any
            // directory that's already on PATH (a user-managed install). Either
            // way officecli is invokable as a command, so don't bootstrap a
            // duplicate copy into the canonical dir. (RecordInstalledVersion is
            // handled by explicit `install`.)
            var pathComparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (string.Equals(Path.GetFullPath(src), Path.GetFullPath(TargetPath), pathComparison)
                || IsDirOnPath(Path.GetDirectoryName(src)))
                return;

            // Dev-build filter: framework-dependent / dotnet run binaries are <5MB
            FileInfo srcInfo;
            try { srcInfo = new FileInfo(src); }
            catch { return; }
            if (srcInfo.Length < 5 * 1024 * 1024) return;

            var currentVer = UpdateChecker.GetCurrentVersionPublic();
            if (string.IsNullOrEmpty(currentVer)) return;

            if (!File.Exists(TargetPath))
            {
                // Fresh install — full Run() (binary + skills + MCP fallback)
                Console.Error.WriteLine($"note: officecli not installed yet, running first-time install...");
                Run([]);
                return;
            }

            // Upgrade case — compare current vs config-recorded version
            var config = UpdateChecker.LoadConfig();
            var installedVer = config.InstalledBinaryVersion;
            if (string.IsNullOrEmpty(installedVer))
            {
                // Config field missing (older install) — fall back to subprocess once.
                installedVer = ReadVersionFromBinary(TargetPath);
                if (!string.IsNullOrEmpty(installedVer))
                {
                    config.InstalledBinaryVersion = installedVer;
                    try { UpdateChecker.SaveConfig(config); } catch { }
                }
            }

            if (string.IsNullOrEmpty(installedVer)) return;
            if (!UpdateChecker.IsNewerPublic(currentVer, installedVer)) return;

            // Strict upgrade — binary only, leave skills/MCP alone
            InstallBinary(quiet: true);
        }
        catch { /* never block the user's command */ }
    }

    private static string? ReadVersionFromBinary(string path)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = path,
                Arguments = "--version",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
            };
            using var proc = Process.Start(psi);
            if (proc == null) return null;
            if (!proc.WaitForExit(2000))
            {
                try { proc.Kill(); } catch { }
                return null;
            }
            var output = (proc.StandardOutput.ReadToEnd() + " " + proc.StandardError.ReadToEnd()).Trim();
            // Match first x.y.z token
            var match = System.Text.RegularExpressions.Regex.Match(output, @"\d+\.\d+\.\d+");
            return match.Success ? match.Value : null;
        }
        catch { return null; }
    }

    private static bool IsInPath() => IsDirOnPath(BinDir);

    /// <summary>True if <paramref name="dir"/> is one of the PATH entries
    /// (case-insensitive on Windows), i.e. a binary living there is invokable by
    /// bare command name.</summary>
    private static bool IsDirOnPath(string? dir)
    {
        if (string.IsNullOrEmpty(dir)) return false;
        string full;
        try { full = Path.GetFullPath(dir); }
        catch { return false; }
        var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        return pathEnv.Split(Path.PathSeparator).Any(p =>
        {
            if (string.IsNullOrWhiteSpace(p)) return false;
            try { return Path.GetFullPath(p).Equals(full, comparison); }
            catch { return false; }
        });
    }

    private static void EnsurePath(bool quiet = false)
    {
        if (IsInPath())
            return;

        var exportLine = $"export PATH=\"{BinDir}:$PATH\"";
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        // Determine shell profile to update
        string profilePath;
        if (OperatingSystem.IsWindows())
        {
            // Windows: add to user PATH via registry (same as install.ps1)
            var currentPath = Environment.GetEnvironmentVariable("Path", EnvironmentVariableTarget.User) ?? "";
            if (!currentPath.Split(Path.PathSeparator).Contains(BinDir, StringComparer.OrdinalIgnoreCase))
            {
                var newPath = string.IsNullOrEmpty(currentPath) ? BinDir : $"{currentPath}{Path.PathSeparator}{BinDir}";
                Environment.SetEnvironmentVariable("Path", newPath, EnvironmentVariableTarget.User);
                if (!quiet)
                {
                    Console.WriteLine($"  Added {BinDir} to PATH.");
                    Console.WriteLine($"  Restart your terminal to apply changes.");
                }
            }
            return;
        }

        var shell = Environment.GetEnvironmentVariable("SHELL") ?? "";
        if (shell.EndsWith("/zsh"))
            profilePath = Path.Combine(home, ".zshrc");
        else if (shell.EndsWith("/bash"))
            profilePath = Path.Combine(home, ".bashrc");
        else if (shell.EndsWith("/fish"))
        {
            // fish uses a different syntax
            var fishConfig = Path.Combine(home, ".config", "fish", "config.fish");
            var fishLine = $"fish_add_path {BinDir}";
            AppendIfMissing(fishConfig, fishLine, BinDir);
            return;
        }
        else
        {
            // Unknown shell — try .profile as fallback
            profilePath = Path.Combine(home, ".profile");
        }

        AppendIfMissing(profilePath, exportLine, BinDir);
    }

    private static void AppendIfMissing(string profilePath, string line, string marker)
    {
        // Check if already present in the file
        if (File.Exists(profilePath))
        {
            var content = File.ReadAllText(profilePath);
            if (content.Contains(marker))
                return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(profilePath)!);
        File.AppendAllText(profilePath, $"\n# Added by officecli\n{line}\n");
        Console.WriteLine($"  Added {marker} to PATH in {profilePath}");
        Console.WriteLine($"  Run: source {profilePath}  (or open a new terminal)");
    }
}
