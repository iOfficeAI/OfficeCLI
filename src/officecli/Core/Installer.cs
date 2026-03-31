// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Installs officecli binary, skills, and MCP (for tools without skill support).
/// Usage:
///   officecli install [target]  — install binary + skills + fallback MCP
/// </summary>
public static class Installer
{
    private static readonly string BinDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".local", "bin");

    private static readonly string TargetPath = Path.Combine(BinDir, "officecli");

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
        ("minimax", ".minimax",                        ["minimax", "minimax-cli"]),
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

    private static void InstallBinary()
    {
        var src = Environment.ProcessPath;
        if (string.IsNullOrEmpty(src))
            return;

        // Already at target location — skip
        if (string.Equals(Path.GetFullPath(src), Path.GetFullPath(TargetPath), StringComparison.Ordinal))
            return;

        // Skip if not a self-contained published binary (e.g. running via dotnet run)
        // Self-contained single-file binaries are typically >5MB; framework-dependent builds are <1MB
        var srcInfo = new FileInfo(src);
        if (srcInfo.Length < 5 * 1024 * 1024)
        {
            Console.WriteLine($"Skipping binary install: not a published self-contained binary.");
            Console.WriteLine($"  Run: dotnet publish -c Release -r <rid> --self-contained -p:PublishSingleFile=true");
            return;
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

        Console.WriteLine($"Installed binary to {TargetPath}");

        EnsurePath();
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

    private static void EnsurePath()
    {
        if (IsInPath())
            return;

        var exportLine = $"export PATH=\"{BinDir}:$PATH\"";
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        // Determine shell profile to update
        string profilePath;
        if (OperatingSystem.IsWindows())
        {
            // Windows: just advise, don't auto-modify registry
            Console.WriteLine($"  Add {BinDir} to your system PATH.");
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
