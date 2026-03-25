// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Core;

/// <summary>
/// Installs officecli skills (SKILL.md) into AI client skill directories.
/// </summary>
public static class SkillInstaller
{
    private static readonly (string[] Aliases, string DisplayName, string DetectDir, string SkillPath)[] Tools =
    [
        (["claude", "claude-code"],       "Claude Code",    ".claude",              Path.Combine(".claude", "skills", "officecli", "SKILL.md")),
        (["copilot", "github-copilot"],   "GitHub Copilot", ".copilot",             Path.Combine(".copilot", "skills", "officecli", "SKILL.md")),
        (["codex", "openai-codex"],       "Codex CLI",      ".agents",              Path.Combine(".agents", "skills", "officecli", "SKILL.md")),
        (["cursor"],                      "Cursor",         ".cursor",              Path.Combine(".cursor", "skills", "officecli", "SKILL.md")),
        (["windsurf"],                    "Windsurf",       ".windsurf",            Path.Combine(".windsurf", "skills", "officecli", "SKILL.md")),
        (["minimax", "minimax-cli"],      "MiniMax CLI",    ".minimax",             Path.Combine(".minimax", "skills", "officecli", "SKILL.md")),
        (["openclaw"],                    "OpenClaw",       ".openclaw",            Path.Combine(".openclaw", "skills", "officecli", "SKILL.md")),
        (["nanobot"],                     "NanoBot",        Path.Combine(".nanobot", "workspace"),   Path.Combine(".nanobot", "workspace", "skills", "officecli", "SKILL.md")),
        (["zeroclaw"],                    "ZeroClaw",       Path.Combine(".zeroclaw", "workspace"),  Path.Combine(".zeroclaw", "workspace", "skills", "officecli", "SKILL.md")),
    ];

    public static void Install(string target)
    {
        var key = target.ToLowerInvariant();

        if (key == "all")
        {
            var found = false;
            foreach (var tool in Tools)
            {
                if (Directory.Exists(Path.Combine(Home, tool.DetectDir)))
                {
                    found = true;
                    InstallTo(tool.DisplayName, Path.Combine(Home, tool.SkillPath));
                }
            }
            if (!found)
                Console.WriteLine("  No supported AI tools detected.");
            return;
        }

        foreach (var tool in Tools)
        {
            if (tool.Aliases.Contains(key))
            {
                InstallTo(tool.DisplayName, Path.Combine(Home, tool.SkillPath));
                return;
            }
        }

        Console.Error.WriteLine($"Unknown target: {target}");
        Console.Error.WriteLine("Supported: claude, copilot, codex, cursor, windsurf, minimax, openclaw, nanobot, zeroclaw, all");
    }

    private static void InstallTo(string displayName, string targetPath)
    {
        var content = LoadEmbeddedResource("OfficeCli.Resources.skill-officecli.md");
        if (content == null)
        {
            Console.Error.WriteLine($"  {displayName}: embedded resource not found");
            return;
        }

        if (File.Exists(targetPath) && File.ReadAllText(targetPath) == content)
        {
            Console.WriteLine($"  {displayName}: already up to date ({targetPath})");
            return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
        File.WriteAllText(targetPath, content);
        Console.WriteLine($"  {displayName}: installed ({targetPath})");
    }

    private static string Home => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    private static string? LoadEmbeddedResource(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
