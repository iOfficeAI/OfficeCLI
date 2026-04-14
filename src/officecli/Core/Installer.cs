// Copyright 2025 OfficeCli (officecli.ai)
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
        var skilledTools = SkillInstaller.Install(target);

        // Install MCP for tools that didn't get a skill
        InstallMcpFallback(skilledTools, target);

        return 0;
    }

    private static void InstallMcpFallback(HashSet<string> skilledTools, string target)
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var isAll = target.Equals("all", StringComparison.OrdinalIgnoreCase);

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
                McpInstaller.Install(mcpTarget);
        }
    }

    internal static bool InstallBinary(bool quiet = false)
    {
        var src = Environment.ProcessPath;
        if (string.IsNullOrEmpty(src))
            return false;

        // Already at target location — record version and skip the copy
        if (string.Equals(Path.GetFullPath(src), Path.GetFullPath(TargetPath), StringComparison.Ordinal))
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
        File.Copy(src, TargetPath, overwrite: true);

        // Preserve executable permission on Unix
        if (!OperatingSystem.IsWindows())
        {
            try
            {
                File.SetUnixFileMode(TargetPath,
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |
                    UnixFileMode.GroupRead | UnixFileMode.GroupExecute |
                    UnixFileMode.OtherRead | UnixFileMode.OtherExecute);
            }
            catch { /* best effort */ }
        }

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

            // Already running from target — nothing to do (RecordInstalledVersion is handled by explicit `install`)
            if (string.Equals(Path.GetFullPath(src), Path.GetFullPath(TargetPath), StringComparison.Ordinal))
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

    private static bool IsInPath()
    {
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        return pathEnv.Split(Path.PathSeparator).Any(p =>
        {
            try { return Path.GetFullPath(p).Equals(Path.GetFullPath(BinDir), StringComparison.OrdinalIgnoreCase); }
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
            // Windows: add to user PATH via registry
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
