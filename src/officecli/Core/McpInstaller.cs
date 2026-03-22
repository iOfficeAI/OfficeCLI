// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Registers officecli as an MCP server in various AI clients.
/// </summary>
public static class McpInstaller
{
    private static string OfficecliPath =>
        Environment.ProcessPath ?? "officecli";

    public static void Install(string target)
    {
        switch (target.ToLowerInvariant())
        {
            case "lms" or "lmstudio" or "lm-studio":
                InstallLmStudio();
                break;
            case "claude" or "claude-code":
                InstallClaude();
                break;
            case "cursor":
                InstallCursor();
                break;
            case "vscode" or "copilot":
                InstallVsCode();
                break;
            case "list":
                ListStatus();
                break;
            case "uninstall":
                Console.WriteLine("Usage: officecli mcp uninstall <target>");
                Console.WriteLine("Targets: lms, claude, cursor, vscode");
                break;
            default:
                Console.Error.WriteLine($"Unknown target: {target}");
                Console.Error.WriteLine("Supported: lms (LM Studio), claude (Claude Code), cursor, vscode (Copilot)");
                Console.Error.WriteLine("Use 'officecli mcp list' to see current status.");
                break;
        }
    }

    public static void Uninstall(string target)
    {
        switch (target.ToLowerInvariant())
        {
            case "lms" or "lmstudio" or "lm-studio":
                UninstallLmStudio();
                break;
            case "claude" or "claude-code":
                UninstallJson("claude", GetClaudeSettingsPath(), "mcpServers");
                break;
            case "cursor":
                UninstallJson("cursor", GetCursorMcpPath(), "mcpServers");
                break;
            case "vscode" or "copilot":
                UninstallJson("vscode", GetVsCodeMcpPath(), "mcpServers");
                break;
            default:
                Console.Error.WriteLine($"Unknown target: {target}");
                break;
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
            $$"""{"command":"{{EscapeJson(OfficecliPath)}}","args":["mcp-serve"]}""" + "\n");

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

    private static string GetClaudeSettingsPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".claude", "settings.json");

    private static void InstallClaude() =>
        InstallJson("Claude Code", GetClaudeSettingsPath(), "mcpServers");

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
                var existing = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
                    File.ReadAllText(configPath));
                if (existing != null)
                {
                    foreach (var kv in existing)
                        root[kv.Key] = kv.Value;
                }
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

        servers["officecli"] = new McpServerEntry { Command = OfficecliPath, Args = ["mcp-serve"] };
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
                        w.WriteStartObject(serversKey);
                        foreach (var server in prop.Value.EnumerateObject())
                        {
                            if (server.Name != "officecli")
                            {
                                w.WritePropertyName(server.Name);
                                server.Value.WriteTo(w);
                            }
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
        CheckJsonStatus("Claude Code", GetClaudeSettingsPath());
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
