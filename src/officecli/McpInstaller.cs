// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Text.Json;

namespace OfficeCli;

/// <summary>
/// Registers officecli as an MCP server in various AI clients.
/// </summary>
public static class McpInstaller
{
    // Path to record as the MCP server command. Must stay valid across
    // upgrades, so resolve to a STABLE location in priority order:
    //   1. The canonical self-install path (~/.local/bin/officecli) — self-
    //      install overwrites that file in place, so the path never changes.
    //   2. `officecli` as found on PATH. For a package-manager install this is
    //      the stable wrapper/symlink (e.g. /opt/homebrew/bin/officecli), which
    //      `brew upgrade` repoints without changing the path. We must NOT use
    //      Environment.ProcessPath here: it resolves the symlink to the
    //      versioned target (…/Cellar/officecli/1.0.106/…) which rots on upgrade.
    //   3. The running binary — last resort for a download/dev build that has
    //      not been installed anywhere on PATH yet.
    private static string OfficecliPath
    {
        get
        {
            var exe = OperatingSystem.IsWindows() ? "officecli.exe" : "officecli";

            var installed = Core.Installer.InstalledBinaryPath;
            if (File.Exists(installed))
                return installed;

            var onPath = ResolveOnPath(exe);
            if (onPath != null)
                return onPath;

            return Environment.ProcessPath ?? exe;
        }
    }

    /// <summary>First <paramref name="exe"/> found across PATH entries, or null.
    /// Returns the PATH-relative location verbatim (a symlink is NOT resolved),
    /// mirroring `which` — so package-manager wrappers stay version-stable.</summary>
    private static string? ResolveOnPath(string exe)
    {
        var pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrEmpty(pathEnv)) return null;
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrEmpty(dir)) continue;
            string candidate;
            try { candidate = Path.Combine(dir, exe); } catch { continue; }
            if (File.Exists(candidate)) return candidate;
        }
        return null;
    }

    /// <summary>Returns true if the target was recognized; false on unknown
    /// target (so the CLI can surface a non-zero exit code).</summary>
    public static bool Install(string target)
    {
        switch (target.ToLowerInvariant())
        {
            case "lms" or "lmstudio" or "lm-studio":
                InstallLmStudio();
                return true;
            case "claude" or "claude-code":
                InstallClaude();
                return true;
            case "cursor":
                InstallCursor();
                return true;
            case "vscode" or "copilot":
                InstallVsCode();
                return true;
            case "list":
                ListStatus();
                return true;
            case "uninstall":
                // Usage hint accompanies a non-zero exit (return false) — keep
                // it on stderr, matching the default branch below and
                // WriteEarlyDispatchUsage. Otherwise scripts that capture stdout
                // see the error text mixed into normal output.
                Console.Error.WriteLine("Usage: officecli mcp uninstall <target>");
                Console.Error.WriteLine("Targets: lms, claude, cursor, vscode");
                return false;
            default:
                Console.Error.WriteLine($"Unknown target: {target}");
                Console.Error.WriteLine("Supported: lms (LM Studio), claude (Claude Code), cursor, vscode (Copilot)");
                Console.Error.WriteLine("Use 'officecli mcp list' to see current status.");
                return false;
        }
    }

    /// <summary>Returns true if the target was recognized; false on unknown
    /// target.</summary>
    public static bool Uninstall(string target)
    {
        switch (target.ToLowerInvariant())
        {
            case "lms" or "lmstudio" or "lm-studio":
                UninstallLmStudio();
                return true;
            case "claude" or "claude-code":
                UninstallClaude();
                return true;
            case "cursor":
                UninstallJson("cursor", GetCursorMcpPath(), "mcpServers");
                return true;
            case "vscode" or "copilot":
                UninstallJson("vscode", GetVsCodeMcpPath(), "mcpServers");
                return true;
            default:
                Console.Error.WriteLine($"Unknown target: {target}");
                return false;
        }
    }

    // ==================== LM Studio ====================

    private static void InstallLmStudio()
    {
        var pluginDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".cache", "lm-studio", "extensions", "plugins", "mcp", "officecli");

        Directory.CreateDirectory(pluginDir);

        File.WriteAllText(Path.Combine(pluginDir, "manifest.json"),
            """{"type":"plugin","runner":"mcpBridge","owner":"mcp","name":"officecli"}""" + "\n");

        File.WriteAllText(Path.Combine(pluginDir, "mcp-bridge-config.json"),
            $$"""{"command":"{{EscapeJson(OfficecliPath)}}","args":["mcp"]}""" + "\n");

        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        File.WriteAllText(Path.Combine(pluginDir, "install-state.json"),
            $$"""{"by":"mcp-bridge-v1","at":{{now}}}""" + "\n");

        Console.WriteLine($"Registered officecli MCP in LM Studio.");
        Console.WriteLine($"  Plugin dir: {pluginDir}");
        Console.WriteLine("  Restart LM Studio to activate.");
    }

    private static void UninstallLmStudio()
    {
        var pluginDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".cache", "lm-studio", "extensions", "plugins", "mcp", "officecli");
        if (Directory.Exists(pluginDir))
        {
            Directory.Delete(pluginDir, true);
            Console.WriteLine("Removed officecli MCP from LM Studio. Restart to apply.");
        }
        else
        {
            Console.WriteLine("officecli MCP not found in LM Studio.");
        }
    }

    // ==================== Claude Code ====================

    // Claude Code reads user-scoped MCP servers from ~/.claude.json (top-level
    // "mcpServers"), NOT from ~/.claude/settings.json — settings.json has no
    // mcpServers key and Claude Code silently ignores it (the server never
    // appears in `claude mcp list`). This is the same file `claude mcp add -s
    // user` writes to, which is why the status check reads it directly.
    private static string GetClaudeConfigPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".claude.json");

    // Prefer the official `claude` CLI: it owns ~/.claude.json's format and
    // handles concurrent writes from a running Claude Code instance — unlike
    // the static config files of the other targets, ~/.claude.json is live
    // state we don't own. Fall back to a direct write only when `claude` isn't
    // on PATH, so registration still works without the CLI installed.
    private static void InstallClaude()
    {
        // Idempotent re-register: drop any stale entry first so re-running picks
        // up the current binary path (`claude mcp add` errors if the name
        // exists). A false return means `claude` isn't on PATH → fall back.
        if (!TryClaudeCli(["mcp", "remove", "-s", "user", "officecli"], out _, out _, out _))
        {
            InstallJson("Claude Code", GetClaudeConfigPath(), "mcpServers");
            return;
        }

        if (TryClaudeCli(["mcp", "add", "-s", "user", "officecli", "--", OfficecliPath, "mcp"],
                out var stdout, out var stderr, out var code) && code == 0)
        {
            Console.WriteLine("Registered officecli MCP in Claude Code.");
            Console.WriteLine("  Via: claude mcp add -s user (config: ~/.claude.json)");
            return;
        }

        var msg = (string.IsNullOrWhiteSpace(stderr) ? stdout : stderr).Trim();
        if (msg.Length > 0)
            Console.Error.WriteLine($"  claude CLI present but `mcp add` failed: {msg}");
        Console.Error.WriteLine("  Falling back to direct ~/.claude.json write.");
        InstallJson("Claude Code", GetClaudeConfigPath(), "mcpServers");
    }

    private static void UninstallClaude()
    {
        // `claude mcp remove` returns non-zero when no such server exists; in
        // that case still sweep ~/.claude.json directly, covering entries an
        // older officecli wrote without the CLI.
        if (TryClaudeCli(["mcp", "remove", "-s", "user", "officecli"], out _, out _, out var code)
            && code == 0)
        {
            Console.WriteLine("Removed officecli MCP from Claude Code.");
            return;
        }
        UninstallJson("Claude Code", GetClaudeConfigPath(), "mcpServers");
    }

    /// <summary>Runs `claude` with the given args. Returns false only when the
    /// process could not be started (claude not on PATH); true otherwise, with
    /// the captured streams and exit code.</summary>
    private static bool TryClaudeCli(string[] args, out string stdout, out string stderr, out int exitCode)
    {
        stdout = ""; stderr = ""; exitCode = -1;
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "claude",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
                UseShellExecute = false,
            };
            foreach (var a in args) psi.ArgumentList.Add(a);
            using var p = Process.Start(psi);
            if (p == null) return false;
            // Async-drain both streams and bound the wait: the serial
            // ReadToEnd pair deadlocked when the child CLI interleaved large
            // stderr output with stdout, and the unbounded WaitForExit hung
            // officecli for as long as the child lived.
            var outTask = p.StandardOutput.ReadToEndAsync();
            var errTask = p.StandardError.ReadToEndAsync();
            if (!p.WaitForExit(30_000))
            {
                try { p.Kill(true); } catch { }
                return false;
            }
            stdout = outTask.Result;
            stderr = errTask.Result;
            exitCode = p.ExitCode;
            return true;
        }
        catch
        {
            return false; // claude not found / not executable
        }
    }

    // ==================== Cursor ====================

    private static string GetCursorMcpPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".cursor", "mcp.json");

    private static void InstallCursor() =>
        InstallJson("Cursor", GetCursorMcpPath(), "mcpServers");

    // ==================== VS Code ====================

    private static string GetVsCodeMcpPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".vscode", "mcp.json");

    private static void InstallVsCode() =>
        InstallJson("VS Code Copilot", GetVsCodeMcpPath(), "mcpServers");

    // ==================== Generic JSON installer ====================

    private static void InstallJson(string clientName, string configPath, string serversKey)
    {
        var dir = Path.GetDirectoryName(configPath);
        if (dir != null) Directory.CreateDirectory(dir);

        var root = new Dictionary<string, object>();
        if (File.Exists(configPath))
        {
            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
                foreach (var prop in doc.RootElement.EnumerateObject())
                    root[prop.Name] = prop.Value.Clone();
            }
            catch { /* start fresh if parse fails */ }
        }

        // Build the mcpServers section
        var servers = new Dictionary<string, object>();
        if (root.TryGetValue(serversKey, out var existingServers) && existingServers is JsonElement el && el.ValueKind == JsonValueKind.Object)
        {
            foreach (var prop in el.EnumerateObject())
            {
                if (prop.Name != "officecli")
                    servers[prop.Name] = prop.Value;
            }
        }

        servers["officecli"] = new McpServerEntry { Command = OfficecliPath, Args = ["mcp"] };
        root[serversKey] = servers;

        // Write with proper formatting using Utf8JsonWriter
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
        {
            w.WriteStartObject();
            foreach (var kv in root)
            {
                w.WritePropertyName(kv.Key);
                if (kv.Value is JsonElement je)
                    je.WriteTo(w);
                else if (kv.Value is Dictionary<string, object> dict)
                    WriteServersDict(w, dict);
                else
                    w.WriteNullValue();
            }
            w.WriteEndObject();
        }

        File.WriteAllText(configPath, System.Text.Encoding.UTF8.GetString(ms.ToArray()) + "\n");

        Console.WriteLine($"Registered officecli MCP in {clientName}.");
        Console.WriteLine($"  Config: {configPath}");
    }

    private static void WriteServersDict(Utf8JsonWriter w, Dictionary<string, object> dict)
    {
        w.WriteStartObject();
        foreach (var kv in dict)
        {
            w.WritePropertyName(kv.Key);
            if (kv.Value is JsonElement je)
                je.WriteTo(w);
            else if (kv.Value is McpServerEntry entry)
            {
                w.WriteStartObject();
                w.WriteString("command", entry.Command);
                w.WriteStartArray("args");
                foreach (var a in entry.Args) w.WriteStringValue(a);
                w.WriteEndArray();
                w.WriteEndObject();
            }
        }
        w.WriteEndObject();
    }

    private static void UninstallJson(string clientName, string configPath, string serversKey)
    {
        if (!File.Exists(configPath))
        {
            Console.WriteLine($"officecli MCP not found in {clientName}.");
            return;
        }

        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
            using var ms = new MemoryStream();
            using (var w = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
            {
                w.WriteStartObject();
                foreach (var prop in doc.RootElement.EnumerateObject())
                {
                    if (prop.Name == serversKey && prop.Value.ValueKind == JsonValueKind.Object)
                    {
                        // Drop the serversKey entirely if officecli was the only
                        // server — avoid leaving an empty "mcpServers": {} residue.
                        var remaining = prop.Value.EnumerateObject()
                            .Where(s => s.Name != "officecli").ToList();
                        if (remaining.Count == 0)
                            continue;
                        w.WriteStartObject(serversKey);
                        foreach (var server in remaining)
                        {
                            w.WritePropertyName(server.Name);
                            server.Value.WriteTo(w);
                        }
                        w.WriteEndObject();
                    }
                    else
                    {
                        w.WritePropertyName(prop.Name);
                        prop.Value.WriteTo(w);
                    }
                }
                w.WriteEndObject();
            }
            File.WriteAllText(configPath, System.Text.Encoding.UTF8.GetString(ms.ToArray()) + "\n");
            Console.WriteLine($"Removed officecli MCP from {clientName}.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Failed to update {configPath}: {ex.Message}");
        }
    }

    // ==================== Status ====================

    private static void ListStatus()
    {
        Console.WriteLine("officecli MCP registration status:");
        Console.WriteLine();

        CheckStatus("LM Studio", Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".cache", "lm-studio", "extensions", "plugins", "mcp", "officecli", "manifest.json"));
        CheckJsonStatus("Claude Code", GetClaudeConfigPath());
        CheckJsonStatus("Cursor", GetCursorMcpPath());
        CheckJsonStatus("VS Code", GetVsCodeMcpPath());

        Console.WriteLine();
        Console.WriteLine("Commands:");
        Console.WriteLine("  officecli mcp <target>              Register (lms, claude, cursor, vscode)");
        Console.WriteLine("  officecli mcp uninstall <target>    Unregister");
    }

    private static void CheckStatus(string name, string path)
    {
        var exists = File.Exists(path);
        Console.WriteLine($"  {(exists ? "✓" : "✗")} {name,-15} {(exists ? "registered" : "not registered")}");
    }

    private static void CheckJsonStatus(string name, string path)
    {
        var registered = false;
        if (File.Exists(path))
        {
            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(path));
                registered = doc.RootElement.TryGetProperty("mcpServers", out var servers)
                    && servers.TryGetProperty("officecli", out _);
            }
            catch { }
        }
        Console.WriteLine($"  {(registered ? "✓" : "✗")} {name,-15} {(registered ? "registered" : "not registered")}");
    }

    private static string EscapeJson(string s) => s.Replace("\\", "\\\\").Replace("\"", "\\\"");

    private class McpServerEntry
    {
        public string Command { get; set; } = "";
        public string[] Args { get; set; } = [];
    }
}
