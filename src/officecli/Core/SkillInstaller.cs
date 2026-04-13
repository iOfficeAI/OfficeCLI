// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Installs officecli skills into AI client skill directories.
/// - officecli skills install            → base SKILL.md to all detected agents
/// - officecli skills install morph-ppt  → specific skill to all detected agents
/// - officecli skills install claude     → base SKILL.md to specific agent (legacy)
/// </summary>
internal static class SkillInstaller
{
    private static readonly (string[] Aliases, string DisplayName, string DetectDir, string SkillDir)[] Tools =
    [
        (["claude", "claude-code"],       "Claude Code",    ".claude",              Path.Combine(".claude", "skills")),
        (["copilot", "github-copilot"],   "GitHub Copilot", ".copilot",             Path.Combine(".copilot", "skills")),
        (["codex", "openai-codex"],       "Codex CLI",      ".agents",              Path.Combine(".agents", "skills")),
        (["cursor"],                      "Cursor",         ".cursor",              Path.Combine(".cursor", "skills")),
        (["windsurf"],                    "Windsurf",       ".windsurf",            Path.Combine(".windsurf", "skills")),
        (["minimax", "minimax-cli"],      "MiniMax CLI",    ".minimax",             Path.Combine(".minimax", "skills")),
        (["opencode"],                    "OpenCode",       ".opencode",            Path.Combine(".opencode", "skills")),
        (["openclaw"],                    "OpenClaw",       ".openclaw",            Path.Combine(".openclaw", "skills")),
        (["nanobot"],                     "NanoBot",        Path.Combine(".nanobot", "workspace"),   Path.Combine(".nanobot", "workspace", "skills")),
        (["zeroclaw"],                    "ZeroClaw",       Path.Combine(".zeroclaw", "workspace"),  Path.Combine(".zeroclaw", "workspace", "skills")),
    ];

    // Guide name → skill folder name mapping
    private static readonly Dictionary<string, string> SkillMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["hwpx"]            = "officecli-hwpx",
        ["pptx"]            = "officecli-pptx",
        ["word"]            = "officecli-docx",
        ["excel"]           = "officecli-xlsx",
        ["morph-ppt"]       = "morph-ppt",
        ["pitch-deck"]      = "officecli-pitch-deck",
        ["academic-paper"]  = "officecli-academic-paper",
        ["data-dashboard"]  = "officecli-data-dashboard",
        ["financial-model"] = "officecli-financial-model",
    };

    /// <summary>
    /// List all available skills with install status and description.
    /// </summary>
    public static void ListSkills()
    {
        Console.WriteLine();
        Console.WriteLine("Available skills:");
        Console.WriteLine();

        // Collect all agent skill dirs to check install status
        var agentSkillDirs = new List<string>();
        foreach (var tool in Tools)
        {
            if (Directory.Exists(Path.Combine(Home, tool.DetectDir)))
                agentSkillDirs.Add(Path.Combine(Home, tool.SkillDir));
        }

        // Find max skill name length for alignment
        var maxLen = SkillMap.Keys.Max(k => k.Length);

        foreach (var (skillName, folder) in SkillMap)
        {
            // Check if installed in any agent
            var installed = agentSkillDirs.Any(dir =>
                File.Exists(Path.Combine(dir, folder, "SKILL.md")));

            var status = installed ? "[installed]" : "[not installed]";

            // Parse description from embedded SKILL.md
            var description = GetSkillDescription(folder);

            var padding = new string(' ', maxLen - skillName.Length);
            Console.WriteLine($"  {skillName}{padding}  {status,-15}  {description}");
        }

        Console.WriteLine();
        Console.WriteLine("Install: officecli skills install <name>");
        Console.WriteLine();
    }

    /// <summary>
    /// Parse description from the embedded SKILL.md front-matter for a given skill folder.
    /// </summary>
    private static string GetSkillDescription(string folder)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var resourceName = $"skills/{folder}/SKILL.md";

        if (resourceName == null) return "";

        var content = LoadEmbeddedResource(resourceName);
        if (content == null) return "";

        // Parse YAML front-matter: find description field
        if (!content.StartsWith("---")) return "";

        var endIdx = content.IndexOf("---", 3);
        if (endIdx < 0) return "";

        var frontMatter = content[3..endIdx];
        foreach (var line in frontMatter.Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("description:", StringComparison.OrdinalIgnoreCase))
            {
                var desc = trimmed["description:".Length..].Trim().Trim('"');
                // Truncate long descriptions for display
                if (desc.Length > 60)
                    desc = desc[..57] + "...";
                return desc;
            }
        }

        return "";
    }

    /// <summary>
    /// Main entry point. Handles all skills sub-commands.
    /// </summary>
    public static HashSet<string> Install(string target)
    {
        var key = target.ToLowerInvariant();

        // "install" with no further args → base SKILL.md to all detected agents
        if (key == "install")
            return InstallBaseToAll();

        // Check if it's a known skill name → install that skill to all detected agents
        if (SkillMap.ContainsKey(key))
            return InstallSkillToAll(key);

        // Check if second arg after "install" was passed via Program.cs
        // "all" → base SKILL.md to all detected agents
        if (key == "all")
            return InstallBaseToAll();

        // Otherwise treat as agent target name (legacy: officecli skills claude)
        return InstallBaseToAgent(key);
    }

    /// <summary>
    /// Install a specific skill by name to all detected agents.
    /// Called as: officecli skills install morph-ppt
    /// </summary>
    public static HashSet<string> InstallSkill(string skillName)
    {
        return InstallSkillToAll(skillName);
    }

    // ─── Base SKILL.md installation ───────────────────────────

    private static HashSet<string> InstallBaseToAll()
    {
        var installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var found = false;

        foreach (var tool in Tools)
        {
            if (Directory.Exists(Path.Combine(Home, tool.DetectDir)))
            {
                found = true;
                var targetPath = Path.Combine(Home, tool.SkillDir, "officecli", "SKILL.md");
                InstallBaseFile(tool.DisplayName, targetPath);
                foreach (var alias in tool.Aliases)
                    installed.Add(alias);
            }
        }

        if (!found)
            Console.WriteLine("  No supported AI tools detected.");

        return installed;
    }

    private static HashSet<string> InstallBaseToAgent(string agentKey)
    {
        var installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var tool in Tools)
        {
            if (tool.Aliases.Contains(agentKey))
            {
                var targetPath = Path.Combine(Home, tool.SkillDir, "officecli", "SKILL.md");
                InstallBaseFile(tool.DisplayName, targetPath);
                foreach (var alias in tool.Aliases)
                    installed.Add(alias);
                return installed;
            }
        }

        Console.Error.WriteLine($"Unknown target: {agentKey}");
        Console.Error.WriteLine("Supported: claude, copilot, codex, cursor, windsurf, minimax, opencode, openclaw, nanobot, zeroclaw, all");
        Console.Error.WriteLine($"Or a skill name: {string.Join(", ", SkillMap.Keys.OrderBy(k => k))}");
        return installed;
    }

    private static void InstallBaseFile(string displayName, string targetPath)
    {
        var content = LoadEmbeddedResource("OfficeCli.Resources.skill-officecli.md");
        if (content == null)
        {
            Console.Error.WriteLine($"  {displayName}: embedded resource not found");
            return;
        }

        if (File.Exists(targetPath) && File.ReadAllText(targetPath) == content)
        {
            Console.WriteLine($"  {displayName}: officecli already up to date");
            return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
        File.WriteAllText(targetPath, content);
        Console.WriteLine($"  {displayName}: officecli installed ({targetPath})");
    }

    // ─── Specific skill installation ───────────────────────────

    private static HashSet<string> InstallSkillToAll(string skillName)
    {
        var installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (!SkillMap.TryGetValue(skillName, out var folder))
        {
            Console.Error.WriteLine($"Unknown skill: {skillName}");
            Console.Error.WriteLine($"Available: {string.Join(", ", SkillMap.Keys.OrderBy(k => k))}");
            return installed;
        }

        // Find all embedded files for this skill
        var files = GetEmbeddedSkillFiles(folder);
        if (files.Count == 0)
        {
            Console.Error.WriteLine($"  No embedded files found for skill '{skillName}'");
            return installed;
        }

        var found = false;
        foreach (var tool in Tools)
        {
            if (Directory.Exists(Path.Combine(Home, tool.DetectDir)))
            {
                found = true;
                var skillDir = Path.Combine(Home, tool.SkillDir, folder);
                var updated = InstallSkillFiles(tool.DisplayName, skillDir, files);
                if (updated)
                {
                    foreach (var alias in tool.Aliases)
                        installed.Add(alias);
                }
            }
        }

        if (!found)
            Console.WriteLine("  No supported AI tools detected.");

        return installed;
    }

    /// <summary>Install all files for a skill into a target directory.</summary>
    private static bool InstallSkillFiles(string displayName, string targetDir, Dictionary<string, string> files)
    {
        var anyUpdated = false;

        foreach (var (fileName, content) in files)
        {
            var targetPath = Path.Combine(targetDir, fileName);
            // Only rewrite markdown files, leave scripts/other files as-is
            var rewritten = fileName.EndsWith(".md", StringComparison.OrdinalIgnoreCase)
                ? RewriteFileReferences(content, fileName)
                : content;

            if (File.Exists(targetPath) && File.ReadAllText(targetPath) == rewritten)
                continue;

            Directory.CreateDirectory(Path.GetDirectoryName(targetPath)!);
            File.WriteAllText(targetPath, rewritten);
            anyUpdated = true;
        }

        if (anyUpdated)
            Console.WriteLine($"  {displayName}: {Path.GetFileName(targetDir)} installed ({targetDir})");
        else
            Console.WriteLine($"  {displayName}: {Path.GetFileName(targetDir)} already up to date");

        return anyUpdated;
    }

    // ─── Embedded resource helpers ───────────────────────────

    private static Dictionary<string, string> GetEmbeddedSkillFiles(string folder)
    {
        var assembly = Assembly.GetExecutingAssembly();
        // LogicalName format: "skills/{folder}/path/to/file.ext"
        var prefix = $"skills/{folder}/";
        var files = new Dictionary<string, string>();

        foreach (var name in assembly.GetManifestResourceNames())
        {
            if (!name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                continue;

            // Preserve relative path: "SKILL.md", "reference/morph-helpers.sh", etc.
            var relativePath = name[prefix.Length..];
            var content = LoadEmbeddedResource(name);
            if (content != null)
                files[relativePath] = content;
        }

        return files;
    }

    /// <summary>
    /// Rewrite cross-skill file references at install time.
    /// Local creating.md/editing.md refs stay as-is (installed alongside).
    /// Cross-skill refs (../other-skill/file.md) → officecli skills install command.
    /// </summary>
    private static string RewriteFileReferences(string content, string currentFile)
    {
        var folderToSkill = SkillMap.ToDictionary(kv => kv.Value, kv => kv.Key, StringComparer.OrdinalIgnoreCase);

        // Cross-skill markdown links: [text](../officecli-pptx/creating.md) → install command
        content = System.Text.RegularExpressions.Regex.Replace(content,
            @"\[([^\]]*?)\]\(\.\./([^/]+)/(creating|editing|SKILL)\.md([^)]*)\)",
            m =>
            {
                var folder = m.Groups[2].Value;
                var file = m.Groups[3].Value;
                var skill = folderToSkill.GetValueOrDefault(folder, folder);
                return $"`officecli skills install {skill}` then read {file}.md";
            });

        // "officecli-xxx (editing.md)" pattern
        content = System.Text.RegularExpressions.Regex.Replace(content,
            @"officecli-(\w+)\s*\((creating|editing)\.md\)",
            m =>
            {
                var suffix = m.Groups[1].Value;
                var file = m.Groups[2].Value;
                var folder2 = "officecli-" + suffix;
                var skill = folderToSkill.GetValueOrDefault(folder2, suffix);
                return $"`officecli skills install {skill}` ({file}.md)";
            });

        return content;
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
