// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace OfficeCli.Core;

/// <summary>
/// One-shot AionUI integration: installs the officecli-track-changes skill
/// and registers the "Word 修订助手" assistant so users get revision
/// capabilities immediately after running `officecli setup-aionui`.
/// </summary>
internal static class AionuiInstaller
{
    private const string SkillFolder = "officecli-track-changes";
    private const string SkillResource = "skills/officecli-track-changes/SKILL.md";
    private const string ConfigFileName = "aionui-config.txt";
    private const string AssistantIdPrefix = "custom-officecli-revision-";

    /// <summary>
    /// Main entry point. Called from Program.cs early-dispatch.
    /// </summary>
    public static int Run(string[] args)
    {
        var verbose = args.Contains("--verbose") || args.Contains("-v");
        var dryRun = args.Contains("--dry-run");
        var force = args.Contains("--force");

        var configDir = DetectConfigDir();
        if (configDir == null)
        {
            Console.Error.WriteLine("AionUI config directory not found. Is AionUI installed?");
            Console.Error.WriteLine("Supported paths:");
            Console.Error.WriteLine("  macOS:   ~/Library/Application Support/AionUI/config/");
            Console.Error.WriteLine("  Linux:   ~/.config/AionUI/config/");
            Console.Error.WriteLine("  Windows: %APPDATA%/AionUI/config/");
            return 1;
        }

        Console.WriteLine($"AionUI detected at: {configDir}");
        if (dryRun) Console.WriteLine("[dry-run] No changes will be made.");
        Console.WriteLine();

        var ok = true;

        // 1. Install skill file
        ok &= InstallSkill(configDir, verbose, dryRun);

        // 2. Register assistant in config
        ok &= RegisterAssistant(configDir, verbose, dryRun, force);

        if (dryRun)
        {
            Console.WriteLine();
            Console.WriteLine("[dry-run] Run without --dry-run to apply changes.");
        }
        else if (ok)
        {
            Console.WriteLine();
            Console.WriteLine("Done! Restart AionUI to see \"Word 修订助手\" in the assistant list.");
        }

        return ok ? 0 : 1;
    }

    // ─── Detection ───────────────────────────────────────────

    private static string? DetectConfigDir()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        // Ordered by platform likelihood — first match wins
        var candidates = new List<string>();

        if (OperatingSystem.IsMacOS())
        {
            candidates.Add(Path.Combine(home, "Library", "Application Support", "AionUI", "config"));
        }
        else if (OperatingSystem.IsLinux())
        {
            candidates.Add(Path.Combine(home, ".config", "AionUI", "config"));
        }
        else if (OperatingSystem.IsWindows())
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            candidates.Add(Path.Combine(appData, "AionUI", "config"));
        }

        // Cross-platform fallback: try all known paths
        candidates.Add(Path.Combine(home, "Library", "Application Support", "AionUI", "config"));
        candidates.Add(Path.Combine(home, ".config", "AionUI", "config"));

        foreach (var dir in candidates)
        {
            if (Directory.Exists(dir))
                return dir;
        }

        return null;
    }

    // ─── Skill installation ──────────────────────────────────

    private static bool InstallSkill(string configDir, bool verbose, bool dryRun)
    {
        var skillsDir = Path.Combine(configDir, "skills");
        var targetDir = Path.Combine(skillsDir, SkillFolder);
        var targetFile = Path.Combine(targetDir, "SKILL.md");

        var content = LoadEmbeddedSkill();
        if (content == null)
        {
            Console.Error.WriteLine($"  ✗ Embedded skill resource not found: {SkillResource}");
            return false;
        }

        if (File.Exists(targetFile) && File.ReadAllText(targetFile) == content)
        {
            Console.WriteLine($"  ✓ Skill already up to date: {SkillFolder}");
            return true;
        }

        if (dryRun)
        {
            Console.WriteLine($"  [dry-run] Would install skill: {targetFile}");
            return true;
        }

        try
        {
            Directory.CreateDirectory(targetDir);
            File.WriteAllText(targetFile, content);
            Console.WriteLine($"  ✓ Skill installed: {targetFile}");
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  ✗ Failed to install skill: {ex.Message}");
            if (verbose) Console.Error.WriteLine(ex);
            return false;
        }
    }

    private static string? LoadEmbeddedSkill()
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(SkillResource);
        if (stream == null) return null;
        using var reader = new StreamReader(stream, Encoding.UTF8);
        return reader.ReadToEnd();
    }

    // ─── Assistant registration ──────────────────────────────

    private static bool RegisterAssistant(string configDir, bool verbose, bool dryRun, bool force)
    {
        var configPath = Path.Combine(configDir, ConfigFileName);
        if (!File.Exists(configPath))
        {
            Console.Error.WriteLine($"  ✗ Config file not found: {configPath}");
            return false;
        }

        try
        {
            // Read and decode config
            var raw = File.ReadAllText(configPath).Trim();
            var json = DecodeConfig(raw);
            var root = JsonNode.Parse(json);
            if (root == null)
            {
                Console.Error.WriteLine("  ✗ Failed to parse AionUI config JSON");
                return false;
            }

            var assistants = root["assistants"]?.AsArray();
            if (assistants == null)
            {
                // Create assistants array if missing
                assistants = new JsonArray();
                root["assistants"] = assistants;
            }

            // Check if already registered — use _setupBy marker for deterministic match
            // (more robust than name-only matching which can fail on encoding edge cases)
            var existing = assistants.FirstOrDefault(a =>
                a?["_setupBy"]?.GetValue<string>() == "officecli");
            if (existing == null)
            {
                // Fallback: name-based check for manually created assistants
                existing = assistants.FirstOrDefault(a =>
                    a?["name"]?.ToString() == "Word 修订助手");
            }
            if (existing != null && !force)
            {
                Console.WriteLine($"  ✓ Assistant already registered: Word 修订助手 (id: {existing["id"]})");
                return true;
            }

            if (existing != null && force)
            {
                assistants.Remove(existing);
                if (verbose) Console.WriteLine("  --force: removed existing assistant entry");
            }

            // Build assistant entry
            var entry = new JsonObject
            {
                ["id"] = AssistantIdPrefix + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
                ["name"] = "Word 修订助手",
                ["description"] = "专注 Word 文档修订（Track Changes）的助手，支持查找替换修订、格式修订、接受/拒绝修订。",
                ["avatar"] = "📝",
                ["isPreset"] = true,
                ["isBuiltin"] = false,
                ["presetAgentType"] = "opencode",
                ["enabled"] = true,
                ["enabledSkills"] = new JsonArray("officecli-docx", "officecli-track-changes"),
                ["customSkillNames"] = new JsonArray(),
                ["_setupBy"] = "officecli",
            };

            if (dryRun)
            {
                Console.WriteLine($"  [dry-run] Would register assistant: Word 修订助手 (id: {entry["id"]})");
                return true;
            }

            // Backup original config
            var backupPath = configPath + ".bak";
            File.Copy(configPath, backupPath, overwrite: true);
            if (verbose) Console.WriteLine($"  Backed up config to: {backupPath}");

            // Add assistant and re-encode
            assistants.Add((JsonNode)entry);
            var newJson = root.ToJsonString(new JsonSerializerOptions { WriteIndented = false });
            var encoded = EncodeConfig(newJson);
            File.WriteAllText(configPath, encoded);

            Console.WriteLine($"  ✓ Assistant registered: Word 修订助手 (id: {entry["id"]})");
            return true;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"  ✗ Failed to register assistant: {ex.Message}");
            if (verbose) Console.Error.WriteLine(ex);
            return false;
        }
    }

    // ─── Config encoding ─────────────────────────────────────

    /// <summary>
    /// Decode AionUI config: base64url → URL-decode → JSON string.
    /// </summary>
    private static string DecodeConfig(string raw)
    {
        var bytes = Base64UrlDecode(raw);
        var urlEncoded = Encoding.UTF8.GetString(bytes);
        return Uri.UnescapeDataString(urlEncoded);
    }

    /// <summary>
    /// Encode AionUI config: JSON string → URL-encode → base64url.
    /// </summary>
    private static string EncodeConfig(string json)
    {
        var urlEncoded = Uri.EscapeDataString(json);
        var bytes = Encoding.UTF8.GetBytes(urlEncoded);
        return Base64UrlEncode(bytes);
    }

    private static byte[] Base64UrlDecode(string input)
    {
        var base64 = input
            .Replace('-', '+')
            .Replace('_', '/');
        switch (base64.Length % 4)
        {
            case 2: base64 += "=="; break;
            case 3: base64 += "="; break;
        }
        return Convert.FromBase64String(base64);
    }

    private static string Base64UrlEncode(byte[] bytes)
    {
        return Convert.ToBase64String(bytes)
            .Replace('+', '-')
            .Replace('/', '_')
            .TrimEnd('=');
    }
}
