// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    // Umbrella skill folder name. Embedded via the `skills/**/*` glob in
    // officecli.csproj — same logical-name shape as every sub-skill, no
    // special-case resource path. Kept out of SkillMap on purpose so
    // `officecli skills list` and `load_skill` only surface sub-skills.
    private const string UmbrellaFolder = "officecli";
    private static string UmbrellaResource => $"skills/{UmbrellaFolder}/SKILL.md";

    private static readonly (string[] Aliases, string DisplayName, string DetectDir, string SkillDir)[] Tools =
    [
        (["claude", "claude-code"],       "Claude Code",    ".claude",              Path.Combine(".claude", "skills")),
        (["copilot", "github-copilot"],   "GitHub Copilot", ".copilot",             Path.Combine(".copilot", "skills")),
        (["codex", "openai-codex"],       "Codex CLI",      ".agents",              Path.Combine(".agents", "skills")),
        (["cursor"],                      "Cursor",         ".cursor",              Path.Combine(".cursor", "skills")),
        // Pi (earendil-works/pi) implements the Agent Skills standard; its
        // global skill locations are ~/.pi/agent/skills/ and ~/.agents/skills/.
        // Target the ~/.pi one — the ~/.agents fallback is already covered by
        // the Codex CLI row when that directory exists.
        (["pi", "pi-agent"],              "Pi",             ".pi",                  Path.Combine(".pi", "agent", "skills")),
        (["windsurf"],                    "Windsurf",       ".windsurf",            Path.Combine(".windsurf", "skills")),
        (["minimax", "minimax-cli"],      "MiniMax CLI",    ".minimax",             Path.Combine(".minimax", "skills")),
        (["opencode"],                    "OpenCode",       Path.Combine(".config", "opencode"), Path.Combine(".config", "opencode", "skills")),
        (["hermes", "hermes-agent"],      "Hermes Agent",   ".hermes",              Path.Combine(".hermes", "skills")),
        (["openclaw"],                    "OpenClaw",       ".openclaw",            Path.Combine(".openclaw", "skills")),
        (["nanobot"],                     "NanoBot",        Path.Combine(".nanobot", "workspace"),   Path.Combine(".nanobot", "workspace", "skills")),
        (["zeroclaw"],                    "ZeroClaw",       Path.Combine(".zeroclaw", "workspace"),  Path.Combine(".zeroclaw", "workspace", "skills")),
    ];

    // Guide name → skill folder name mapping
    private static readonly Dictionary<string, string> SkillMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["pptx"]            = "officecli-pptx",
        ["word"]            = "officecli-docx",
        ["excel"]           = "officecli-xlsx",
        ["morph-ppt"]       = "morph-ppt",
        ["morph-ppt-3d"]    = "morph-ppt-3d",
        ["pitch-deck"]      = "officecli-pitch-deck",
        ["academic-paper"]  = "officecli-academic-paper",
        ["data-dashboard"]  = "officecli-data-dashboard",
        ["financial-model"] = "officecli-financial-model",
        ["word-form"]       = "officecli-word-form",
    };

    // One-line trigger per skill — a compact, always-on discovery lure injected
    // into the MCP tool description. Full routing guidance stays lazy (load_skill
    // with no name returns the catalog; name=X returns the SKILL.md). This is the
    // "push minimal trigger, pull the detail" half of the discovery design: it
    // costs a fraction of the full descriptions but is enough to prompt the
    // agent to load the right skill. Keys must track SkillMap (a missing entry
    // degrades to the bare name via BuildSkillTriggerSummary; asserted in tests).
    private static readonly Dictionary<string, string> SkillTriggers = new(StringComparer.OrdinalIgnoreCase)
    {
        ["pptx"]            = "slide decks / presentations",
        ["word"]            = "Word docs, reports, letters, memos",
        ["excel"]           = "spreadsheets, financial models, dashboards",
        ["word-form"]       = "fillable forms, content controls, protected docs",
        ["morph-ppt"]       = "cross-slide Morph animation / continuous motion",
        ["morph-ppt-3d"]    = "3D Morph decks (GLB models, camera)",
        ["pitch-deck"]      = "fundraising / investor decks (seed, Series A/B/C)",
        ["academic-paper"]  = "academic papers / research reports",
        ["data-dashboard"]  = "data dashboards",
        ["financial-model"] = "financial models / projections",
    };

    /// <summary>
    /// Compact one-line-per-skill trigger summary for the MCP tool description.
    /// Always-on but small; the agent reads it to decide which skill to
    /// <c>load_skill</c>. Order follows SkillMap; a skill without a curated
    /// trigger degrades to its name so the list never silently drops a skill.
    /// </summary>
    public static string BuildSkillTriggerSummary()
    {
        var parts = SkillMap.Keys.Select(name =>
            SkillTriggers.TryGetValue(name, out var t) ? $"{name} → {t}" : name);
        // Directive, not merely informational: an informational phrasing ("for
        // the full guide") was empirically ignored even by capable models, which
        // jumped straight to create/add and guessed the schema. The imperative
        // FIRST … BEFORE … is what actually triggers a skill load + help-first.
        return "IMPORTANT — before you create/add/set/remove on any Office file, FIRST run the command "
            + "`load_skill <X>` for that file type (it loads the build guide and tells you to "
            + "consult `help` for the schema). Pick X by need: "
            + string.Join(" · ", parts)
            + ". (run `load_skill` with no name to list all skills.)";
    }

    // Bundled skill assets that cannot ride the MCP/CLI text channel intact
    // (read as text they would corrupt). Listed in the reference manifest but
    // served only via `officecli skills install`.
    private static readonly HashSet<string> BinarySkillExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".pptx", ".docx", ".xlsx", ".png", ".jpg", ".jpeg", ".gif", ".webp", ".glb", ".pdf", ".zip", ".ico",
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
        var desc = GetFullSkillDescription(folder);
        return desc.Length > 60 ? desc[..57] + "..." : desc;  // truncate for aligned console listing
    }

    /// <summary>
    /// Full (untruncated) front-matter <c>description</c> of a skill's SKILL.md.
    /// For officecli skills this field carries the routing guidance ("Use when…
    /// / Trigger on… / Do NOT trigger for…"), so it is exactly what an agent
    /// needs to pick a skill. Empty string when absent.
    /// </summary>
    private static string GetFullSkillDescription(string folder)
    {
        var content = LoadEmbeddedResource($"skills/{folder}/SKILL.md");
        if (content == null || !content.StartsWith("---")) return "";
        var endIdx = content.IndexOf("---", 3);
        if (endIdx < 0) return "";
        foreach (var line in content[3..endIdx].Split('\n'))
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("description:", StringComparison.OrdinalIgnoreCase))
                return trimmed["description:".Length..].Trim().Trim('"');
        }
        return "";
    }

    /// <summary>
    /// Agent-facing skill catalog: every skill's name plus its full routing
    /// description, with usage pointers. Returned by <c>load_skill</c> with no
    /// name (CLI and MCP) so an agent can discover which skill applies before
    /// drilling into its SKILL.md. Pay-on-demand — not injected into every
    /// session's tool description.
    /// </summary>
    public static string BuildSkillCatalog()
    {
        var sb = new StringBuilder();
        sb.Append("# officecli skills\n\n");
        sb.Append("Workflow guides for building documents. Match the triggers below, then:\n");
        sb.Append("- `load_skill <name>` — the skill's full SKILL.md + a manifest of its bundled reference files\n");
        sb.Append("- `load_skill <name> --path <relpath>` — one bundled reference file\n\n");
        foreach (var (name, folder) in SkillMap)
        {
            var desc = GetFullSkillDescription(folder);
            sb.Append($"## {name}\n{(desc.Length > 0 ? desc : "(no description)")}\n\n");
        }
        return sb.ToString().TrimEnd() + "\n";
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

        // Check if second arg after "install" was passed via Program.cs
        // "all" → base SKILL.md to all detected agents
        if (key == "all")
            return InstallBaseToAll();

        // Otherwise treat as agent target name (legacy: officecli skills claude).
        // The previous `officecli skills <skill>` shorthand for "install that
        // skill to all agents" was removed — use the explicit `skills install
        // <name>` form, or `load_skill <name>` if you only want the content.
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

    /// <summary>All known skill aliases, sorted, comma-joined for error messages.</summary>
    public static string KnownSkillsList() => string.Join(", ", SkillMap.Keys.OrderBy(k => k));

    /// <summary>
    /// Return the embedded SKILL.md content for <paramref name="skillName"/> with
    /// no side-effects and no stdout writes. Throws <see cref="ArgumentException"/>
    /// on unknown skill or missing embedded resource. Used by both the CLI
    /// `officecli load_skill &lt;name&gt;` command and the MCP `load_skill` tool —
    /// shared so the two surfaces have identical semantics.
    /// </summary>
    public static string LoadSkillContent(string skillName)
    {
        if (!SkillMap.TryGetValue(skillName, out var folder))
            throw new ArgumentException($"Unknown skill: {skillName}. Available: {KnownSkillsList()}");
        var content = LoadEmbeddedResource($"skills/{folder}/SKILL.md");
        if (content == null)
            throw new ArgumentException($"Embedded SKILL.md not found for '{skillName}'");
        // A SKILL.md is an entry point that defers detail to bundled reference
        // files (reference/*.md, helper scripts, style libraries). Append a
        // manifest so a text-channel caller (MCP / CLI, no skill install) knows
        // those files exist and how to fetch them — otherwise every
        // "see reference/foo" pointer in the body is a dead link.
        return StripSetupSection(content) + BuildReferenceManifest(skillName);
    }

    /// <summary>
    /// Relative paths of every embedded file for a skill except SKILL.md,
    /// sorted. Uses resource names only (no content read) so binary assets are
    /// listed without being mangled.
    /// </summary>
    public static IReadOnlyList<string> ListSkillFiles(string skillName)
    {
        if (!SkillMap.TryGetValue(skillName, out var folder))
            throw new ArgumentException($"Unknown skill: {skillName}. Available: {KnownSkillsList()}");
        var prefix = $"skills/{folder}/";
        return Assembly.GetExecutingAssembly().GetManifestResourceNames()
            .Where(n => n.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Select(n => n[prefix.Length..])
            .Where(rel => !rel.Equals("SKILL.md", StringComparison.OrdinalIgnoreCase))
            .OrderBy(rel => rel, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// Return the text content of one bundled reference file inside a skill
    /// (e.g. "reference/decision-rules.md"). Shared by the CLI
    /// `load_skill &lt;name&gt; --path &lt;rel&gt;` command and the MCP
    /// `load_skill` tool's path= argument. Throws on unknown skill, path
    /// traversal, a binary asset (cannot ride the text channel), or a missing
    /// file.
    /// </summary>
    public static string LoadSkillFile(string skillName, string relativePath)
    {
        if (!SkillMap.TryGetValue(skillName, out var folder))
            throw new ArgumentException($"Unknown skill: {skillName}. Available: {KnownSkillsList()}");
        var rel = (relativePath ?? "").Replace('\\', '/').TrimStart('/');
        if (rel.Length == 0)
            throw new ArgumentException("path is empty — pass a relative skill file, e.g. reference/decision-rules.md");
        // Contain to the skill folder: reject traversal and current-dir segments.
        if (rel.Split('/').Any(seg => seg is ".." or "."))
            throw new ArgumentException($"Invalid skill file path: {relativePath}");
        if (BinarySkillExtensions.Contains(Path.GetExtension(rel)))
            throw new ArgumentException(
                $"'{rel}' is a binary asset and cannot be served over the text channel. " +
                $"Install the skill to get it on disk: officecli skills install {skillName}");
        var content = LoadEmbeddedResource($"skills/{folder}/{rel}");
        if (content == null)
            throw new ArgumentException(
                $"Skill file not found: {rel}. List available files via the manifest at the end of: " +
                $"officecli load_skill {skillName}");
        return content;
    }

    /// <summary>
    /// Build the "Reference files" manifest appended to a SKILL.md. Deep trees
    /// (≥ 3 path segments, e.g. a 52-directory style library) collapse to one
    /// line per second-level directory so the manifest stays compact; shallow
    /// files (reference/foo.md) are listed individually. Empty string when the
    /// skill bundles nothing beyond SKILL.md.
    /// </summary>
    private static string BuildReferenceManifest(string skillName)
    {
        var files = ListSkillFiles(skillName);
        if (files.Count == 0) return "";
        var shallow = new List<string>();
        var deepGroups = new SortedDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in files)
        {
            var segs = f.Split('/');
            // List shallow files, plus any INDEX.md at any depth — an INDEX is
            // the documented entry into a collapsed tree (e.g. the style
            // library), so the agent shouldn't have to guess its path.
            if (segs.Length <= 2 || segs[^1].Equals("INDEX.md", StringComparison.OrdinalIgnoreCase))
                shallow.Add(f);
            else
            {
                var key = segs[0] + "/" + segs[1] + "/";
                deepGroups[key] = deepGroups.GetValueOrDefault(key) + 1;
            }
        }
        var sb = new StringBuilder();
        sb.Append("\n\n## Reference files (bundled with this skill)\n\n");
        sb.Append("This skill defers detail to the files below. The body's `reference/…` ");
        sb.Append("pointers refer to these. Fetch one with:\n");
        sb.Append($"- `load_skill {skillName} --path <relpath>`\n");
        sb.Append($"- or install the whole tree to disk: `officecli skills install {skillName}`\n\n");
        foreach (var f in shallow) sb.Append($"- `{f}`\n");
        foreach (var (g, n) in deepGroups)
            sb.Append($"- `{g}` — {n} files (binary assets need `skills install`; browse an `INDEX.md` here if present)\n");
        return sb.ToString();
    }

    /// <summary>
    /// Drop the `## Setup` section from a SKILL.md before handing it to an
    /// agent. Whoever just invoked load_skill obviously already has officecli
    /// installed, so the curl-install instructions in that section are pure
    /// noise eating the agent's context. The original on-disk/embedded file
    /// keeps the section intact for humans browsing the repo on GitHub.
    /// Boundary: from a line starting with "## Setup" up to (not including)
    /// the next line starting with "## ".
    /// </summary>
    private static string StripSetupSection(string content)
    {
        var lines = content.Split('\n');
        var sb = new StringBuilder(content.Length);
        var inSetup = false;
        foreach (var line in lines)
        {
            if (!inSetup && line.StartsWith("## Setup", StringComparison.Ordinal))
            {
                inSetup = true;
                continue;
            }
            if (inSetup && line.StartsWith("## ", StringComparison.Ordinal))
                inSetup = false;
            if (!inSetup) sb.Append(line).Append('\n');
        }
        // Split+rejoin may introduce a trailing newline; preserve original behavior.
        var result = sb.ToString();
        if (!content.EndsWith("\n", StringComparison.Ordinal) && result.EndsWith("\n", StringComparison.Ordinal))
            result = result[..^1];
        return result;
    }

    /// <summary>
    /// Install a specific skill by name to a single agent target.
    /// Accepts either order: (skill, agent) or (agent, skill) — skill names and
    /// agent aliases don't overlap so the order is auto-detected.
    /// Called as: officecli skills install morph-ppt hermes  /  officecli skills install hermes morph-ppt
    /// Skips agent detection — installs even if the agent's home dir is missing,
    /// matching the legacy `officecli skills &lt;agent&gt;` behavior.
    /// </summary>
    public static HashSet<string> InstallSkillToAgentTarget(string firstArg, string secondArg)
    {
        var installed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Auto-detect token order
        string? skillName = null;
        string? agentKey = null;
        if (SkillMap.ContainsKey(firstArg))
        {
            skillName = firstArg;
            agentKey = secondArg;
        }
        else if (SkillMap.ContainsKey(secondArg))
        {
            skillName = secondArg;
            agentKey = firstArg;
        }

        if (skillName is null)
        {
            Console.Error.WriteLine($"Unknown skill in: {firstArg} {secondArg}");
            Console.Error.WriteLine($"Available skills: {string.Join(", ", SkillMap.Keys.OrderBy(k => k))}");
            return installed;
        }

        var key = agentKey!.ToLowerInvariant();
        var folder = SkillMap[skillName];

        var tool = Tools.FirstOrDefault(t => t.Aliases.Contains(key));
        if (tool.Aliases is null)
        {
            Console.Error.WriteLine($"Unknown agent: {agentKey}");
            Console.Error.WriteLine("Supported: claude, copilot, codex, cursor, pi, windsurf, minimax, opencode, openclaw, nanobot, zeroclaw, hermes");
            return installed;
        }

        var files = GetEmbeddedSkillFiles(folder);
        if (files.Count == 0)
        {
            Console.Error.WriteLine($"  No embedded files found for skill '{skillName}'");
            return installed;
        }

        var skillDir = Path.Combine(Home, tool.SkillDir, folder);
        InstallSkillFiles(tool.DisplayName, skillDir, files);
        foreach (var alias in tool.Aliases)
            installed.Add(alias);

        return installed;
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
                var targetPath = Path.Combine(Home, tool.SkillDir, UmbrellaFolder, "SKILL.md");
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
                var targetPath = Path.Combine(Home, tool.SkillDir, UmbrellaFolder, "SKILL.md");
                InstallBaseFile(tool.DisplayName, targetPath);
                foreach (var alias in tool.Aliases)
                    installed.Add(alias);
                return installed;
            }
        }

        Console.Error.WriteLine($"Unknown target: {agentKey}");
        Console.Error.WriteLine("Supported agents: claude, copilot, codex, cursor, windsurf, minimax, opencode, openclaw, nanobot, zeroclaw, hermes, all");
        if (SkillMap.ContainsKey(agentKey))
        {
            Console.Error.WriteLine();
            Console.Error.WriteLine($"'{agentKey}' is a skill name, not an agent. Did you mean:");
            Console.Error.WriteLine($"  officecli skills install {agentKey}    (install to disk)");
            Console.Error.WriteLine($"  officecli load_skill {agentKey}        (print SKILL.md to stdout)");
        }
        return installed;
    }

    private static void InstallBaseFile(string displayName, string targetPath)
    {
        var content = LoadEmbeddedResource(UmbrellaResource);
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

        SafeCreateDirectory(Path.GetDirectoryName(targetPath)!);
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
                InstallSkillFiles(tool.DisplayName, skillDir, files);
                // CONSISTENCY(install-success): always add aliases when the
                // agent dir exists, matching InstallBaseToAll's semantics.
                // The exit code derived from this set is "install succeeded
                // for these agents", not "files were rewritten" — idempotent
                // re-install of an up-to-date skill must still report success.
                foreach (var alias in tool.Aliases)
                    installed.Add(alias);
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

            SafeCreateDirectory(Path.GetDirectoryName(targetPath)!);
            File.WriteAllText(targetPath, rewritten);
            anyUpdated = true;
        }

        if (anyUpdated)
            Console.WriteLine($"  {displayName}: {Path.GetFileName(targetDir)} installed ({targetDir})");
        else
            Console.WriteLine($"  {displayName}: {Path.GetFileName(targetDir)} already up to date");

        return anyUpdated;
    }

    // ─── Auto-refresh after binary upgrade ───────────────────

    /// <summary>
    /// Re-install only the skill files that are *already present* in detected
    /// agent directories. Called by UpdateChecker after a binary upgrade so
    /// installed skills stay in sync with the new binary's embedded copies.
    ///
    /// Conservative on purpose:
    ///   - Only refreshes skills the user previously installed (presence of
    ///     SKILL.md per skill folder).
    ///   - Never adds new agents or new sub-skills.
    ///   - Silent unless something actually changed (one summary line on stderr).
    ///   - Identical-content writes are skipped (existing diff-and-write path).
    /// </summary>
    internal static int RefreshInstalled()
    {
        var changedFiles = 0;
        var changedTargets = new List<string>();

        foreach (var tool in Tools)
        {
            // Per-tool isolation: a permission/IO error in one agent's skill
            // dir must not abort the refresh for other agents. Each tool's
            // base SKILL.md and each of its sub-skills are wrapped
            // individually so partial progress is preserved.
            if (!Directory.Exists(Path.Combine(Home, tool.DetectDir))) continue;
            var skillsDir = Path.Combine(Home, tool.SkillDir);
            if (!Directory.Exists(skillsDir)) continue;

            // Base SKILL.md
            try
            {
                var basePath = Path.Combine(skillsDir, UmbrellaFolder, "SKILL.md");
                if (File.Exists(basePath))
                {
                    var content = LoadEmbeddedResource(UmbrellaResource);
                    if (content != null && File.ReadAllText(basePath) != content)
                    {
                        File.WriteAllText(basePath, content);
                        changedFiles++;
                        changedTargets.Add($"{tool.DisplayName}/officecli");
                    }
                }
            }
            catch { /* per-agent failure is non-fatal — keep going */ }

            // Sub-skills present in this agent's skill directory
            foreach (var folder in SkillMap.Values)
            {
                try
                {
                    var subSkillFile = Path.Combine(skillsDir, folder, "SKILL.md");
                    if (!File.Exists(subSkillFile)) continue;

                    var files = GetEmbeddedSkillFiles(folder);
                    if (files.Count == 0) continue;

                    var targetDir = Path.Combine(skillsDir, folder);
                    var n = RewriteSkillFilesQuiet(targetDir, files);
                    if (n > 0)
                    {
                        changedFiles += n;
                        changedTargets.Add($"{tool.DisplayName}/{folder}");
                    }
                }
                catch { /* per-skill failure is non-fatal */ }
            }
        }

        if (changedFiles > 0)
            Console.Error.WriteLine($"officecli: refreshed {changedFiles} skill file(s) after upgrade ({string.Join(", ", changedTargets)})");

        return changedFiles;
    }

    /// <summary>Quiet variant of <see cref="InstallSkillFiles"/>: returns the
    /// number of files rewritten, prints nothing per file. Used by
    /// <see cref="RefreshInstalled"/>.</summary>
    private static int RewriteSkillFilesQuiet(string targetDir, Dictionary<string, string> files)
    {
        var n = 0;
        foreach (var (fileName, content) in files)
        {
            var targetPath = Path.Combine(targetDir, fileName);
            var rewritten = fileName.EndsWith(".md", StringComparison.OrdinalIgnoreCase)
                ? RewriteFileReferences(content, fileName)
                : content;

            if (File.Exists(targetPath) && File.ReadAllText(targetPath) == rewritten)
                continue;

            SafeCreateDirectory(Path.GetDirectoryName(targetPath)!);
            File.WriteAllText(targetPath, rewritten);
            n++;
        }
        return n;
    }

    // ─── Directory helpers ───────────────────────────────────

    /// <summary>
    /// Like Directory.CreateDirectory but handles dangling symlinks:
    /// if the path exists as a symlink whose target is missing, remove it first.
    /// </summary>
    private static void SafeCreateDirectory(string dir)
    {
        // CONSISTENCY(skill-install): dangling symlink guard — Directory.CreateDirectory
        // throws IOException when a path component is a dangling symlink; detect and remove it.
        // Use FileAttributes.ReparsePoint to detect symlinks regardless of whether target exists.
        if (!Directory.Exists(dir))
        {
            try
            {
                var attrs = File.GetAttributes(dir);
                if (attrs.HasFlag(FileAttributes.ReparsePoint))
                {
                    // Dangling symlink (or symlink to non-dir) — remove it so CreateDirectory can proceed
                    File.Delete(dir);
                }
            }
            catch (FileNotFoundException) { /* fine, doesn't exist at all */ }
            catch (DirectoryNotFoundException) { /* fine, parent also missing */ }
        }
        Directory.CreateDirectory(dir);
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
