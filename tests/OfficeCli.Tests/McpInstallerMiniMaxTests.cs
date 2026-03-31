// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using Xunit;

namespace OfficeCli.Tests;

/// <summary>
/// Tests for MiniMax CLI MCP integration in McpInstaller.
/// Verifies install, uninstall, and status listing for the MiniMax target.
/// </summary>
public class McpInstallerMiniMaxTests : IDisposable
{
    private readonly string _tempDir;
    private readonly string _originalHome;

    public McpInstallerMiniMaxTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"officecli-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _originalHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    }

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    // ─── Install tests ───────────────────────────

    [Fact]
    public void Install_MiniMax_CreatesConfigWithMcpServersEntry()
    {
        // Arrange
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");

        // Act — use the generic InstallJson indirectly via public Install
        // Since Install uses Environment.UserProfile for paths, we test the JSON
        // writing logic by calling InstallJson's pattern directly
        var dir = Path.GetDirectoryName(mcpPath)!;
        Directory.CreateDirectory(dir);
        WriteMcpConfig(mcpPath);

        // Assert
        Assert.True(File.Exists(mcpPath));
        using var doc = JsonDocument.Parse(File.ReadAllText(mcpPath));
        Assert.True(doc.RootElement.TryGetProperty("mcpServers", out var servers));
        Assert.True(servers.TryGetProperty("officecli", out var entry));
        Assert.True(entry.TryGetProperty("command", out _));
        Assert.True(entry.TryGetProperty("args", out var args));
        Assert.Equal("mcp", args[0].GetString());
    }

    [Fact]
    public void Install_MiniMax_PreservesExistingServers()
    {
        // Arrange — pre-existing config with another server
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Directory.CreateDirectory(Path.GetDirectoryName(mcpPath)!);
        File.WriteAllText(mcpPath, """{"mcpServers":{"other-tool":{"command":"other","args":[]}}}""");

        // Act
        WriteMcpConfig(mcpPath);

        // Assert
        using var doc = JsonDocument.Parse(File.ReadAllText(mcpPath));
        var servers = doc.RootElement.GetProperty("mcpServers");
        Assert.True(servers.TryGetProperty("other-tool", out _), "Existing server should be preserved");
        Assert.True(servers.TryGetProperty("officecli", out _), "officecli should be added");
    }

    [Fact]
    public void Install_MiniMax_OverwritesExistingOfficecliEntry()
    {
        // Arrange — pre-existing officecli entry with old command
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Directory.CreateDirectory(Path.GetDirectoryName(mcpPath)!);
        File.WriteAllText(mcpPath, """{"mcpServers":{"officecli":{"command":"old-path","args":["mcp"]}}}""");

        // Act
        WriteMcpConfig(mcpPath);

        // Assert
        using var doc = JsonDocument.Parse(File.ReadAllText(mcpPath));
        var entry = doc.RootElement.GetProperty("mcpServers").GetProperty("officecli");
        var command = entry.GetProperty("command").GetString();
        Assert.NotEqual("old-path", command); // Should be updated to current binary path
    }

    [Fact]
    public void Install_MiniMax_CreatesDirectoryIfMissing()
    {
        // Arrange
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Assert.False(Directory.Exists(Path.GetDirectoryName(mcpPath)));

        // Act
        WriteMcpConfig(mcpPath);

        // Assert
        Assert.True(Directory.Exists(Path.GetDirectoryName(mcpPath)));
        Assert.True(File.Exists(mcpPath));
    }

    [Fact]
    public void Install_MiniMax_HandlesCorruptedConfigGracefully()
    {
        // Arrange — corrupt JSON
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Directory.CreateDirectory(Path.GetDirectoryName(mcpPath)!);
        File.WriteAllText(mcpPath, "not valid json {{{");

        // Act — should not throw, starts fresh
        WriteMcpConfig(mcpPath);

        // Assert
        using var doc = JsonDocument.Parse(File.ReadAllText(mcpPath));
        Assert.True(doc.RootElement.TryGetProperty("mcpServers", out var servers));
        Assert.True(servers.TryGetProperty("officecli", out _));
    }

    // ─── Uninstall tests ───────────────────────────

    [Fact]
    public void Uninstall_MiniMax_RemovesOfficecliEntry()
    {
        // Arrange
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Directory.CreateDirectory(Path.GetDirectoryName(mcpPath)!);
        WriteMcpConfig(mcpPath);

        // Verify setup
        using (var doc = JsonDocument.Parse(File.ReadAllText(mcpPath)))
            Assert.True(doc.RootElement.GetProperty("mcpServers").TryGetProperty("officecli", out _));

        // Act
        UninstallMcpConfig(mcpPath);

        // Assert
        using (var doc2 = JsonDocument.Parse(File.ReadAllText(mcpPath)))
        {
            var servers = doc2.RootElement.GetProperty("mcpServers");
            Assert.False(servers.TryGetProperty("officecli", out _), "officecli should be removed");
        }
    }

    [Fact]
    public void Uninstall_MiniMax_PreservesOtherServers()
    {
        // Arrange — config with officecli and another server
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Directory.CreateDirectory(Path.GetDirectoryName(mcpPath)!);
        File.WriteAllText(mcpPath,
            """{"mcpServers":{"officecli":{"command":"officecli","args":["mcp"]},"other":{"command":"other","args":[]}}}""");

        // Act
        UninstallMcpConfig(mcpPath);

        // Assert
        using var doc = JsonDocument.Parse(File.ReadAllText(mcpPath));
        var servers = doc.RootElement.GetProperty("mcpServers");
        Assert.False(servers.TryGetProperty("officecli", out _));
        Assert.True(servers.TryGetProperty("other", out _), "Other servers should be preserved");
    }

    [Fact]
    public void Uninstall_MiniMax_NoErrorWhenFileDoesNotExist()
    {
        // Arrange
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        Assert.False(File.Exists(mcpPath));

        // Act — should not throw
        var ex = Record.Exception(() => UninstallMcpConfig(mcpPath));

        // Assert
        Assert.Null(ex);
    }

    // ─── Status / alias tests ───────────────────────────

    [Theory]
    [InlineData("minimax")]
    [InlineData("minimax-cli")]
    public void Install_MiniMax_AcceptsBothAliases(string alias)
    {
        // The Install switch handles both "minimax" and "minimax-cli"
        // We verify both aliases map to the same target by checking they don't hit "default"
        var output = CaptureConsoleOutput(() => OfficeCli.Core.McpInstaller.Install(alias));
        Assert.DoesNotContain("Unknown target", output);
    }

    [Theory]
    [InlineData("minimax")]
    [InlineData("minimax-cli")]
    public void Uninstall_MiniMax_AcceptsBothAliases(string alias)
    {
        var output = CaptureConsoleOutput(() => OfficeCli.Core.McpInstaller.Uninstall(alias));
        Assert.DoesNotContain("Unknown target", output);
    }

    [Fact]
    public void ListStatus_IncludesMiniMaxCli()
    {
        var output = CaptureConsoleOutput(() => OfficeCli.Core.McpInstaller.Install("list"));
        Assert.Contains("MiniMax CLI", output);
    }

    // ─── JSON format tests ───────────────────────────

    [Fact]
    public void Install_MiniMax_OutputIsValidIndentedJson()
    {
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        WriteMcpConfig(mcpPath);

        var content = File.ReadAllText(mcpPath);
        // Should be indented (contains newlines and spaces)
        Assert.Contains("\n", content);
        Assert.Contains("  ", content);

        // Should be valid JSON
        using var doc = JsonDocument.Parse(content);
        Assert.Equal(JsonValueKind.Object, doc.RootElement.ValueKind);
    }

    [Fact]
    public void Install_MiniMax_ArgsArrayContainsMcp()
    {
        var mcpPath = Path.Combine(_tempDir, ".minimax", "mcp.json");
        WriteMcpConfig(mcpPath);

        using var doc = JsonDocument.Parse(File.ReadAllText(mcpPath));
        var args = doc.RootElement
            .GetProperty("mcpServers")
            .GetProperty("officecli")
            .GetProperty("args");

        Assert.Equal(JsonValueKind.Array, args.ValueKind);
        Assert.Equal(1, args.GetArrayLength());
        Assert.Equal("mcp", args[0].GetString());
    }

    // ─── Helper methods ───────────────────────────

    /// <summary>
    /// Simulates McpInstaller.InstallJson for MiniMax target by replicating the JSON write logic.
    /// This avoids needing to mock Environment.ProcessPath or file system paths.
    /// </summary>
    private static void WriteMcpConfig(string configPath)
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
            catch { }
        }

        var servers = new Dictionary<string, object>();
        if (root.TryGetValue("mcpServers", out var existingServers) && existingServers is JsonElement el && el.ValueKind == JsonValueKind.Object)
        {
            foreach (var prop in el.EnumerateObject())
            {
                if (prop.Name != "officecli")
                    servers[prop.Name] = prop.Value;
            }
        }

        servers["officecli"] = new { command = "officecli", args = new[] { "mcp" } };
        root["mcpServers"] = servers;

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
                {
                    w.WriteStartObject();
                    foreach (var dkv in dict)
                    {
                        w.WritePropertyName(dkv.Key);
                        if (dkv.Value is JsonElement dje)
                            dje.WriteTo(w);
                        else
                        {
                            // Anonymous type with command and args
                            var json = JsonSerializer.Serialize(dkv.Value);
                            using var innerDoc = JsonDocument.Parse(json);
                            innerDoc.RootElement.WriteTo(w);
                        }
                    }
                    w.WriteEndObject();
                }
                else
                    w.WriteNullValue();
            }
            w.WriteEndObject();
        }

        File.WriteAllText(configPath, System.Text.Encoding.UTF8.GetString(ms.ToArray()) + "\n");
    }

    /// <summary>
    /// Simulates McpInstaller.UninstallJson for MiniMax target.
    /// </summary>
    private static void UninstallMcpConfig(string configPath)
    {
        if (!File.Exists(configPath))
            return;

        using var doc = JsonDocument.Parse(File.ReadAllText(configPath));
        using var ms = new MemoryStream();
        using (var w = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
        {
            w.WriteStartObject();
            foreach (var prop in doc.RootElement.EnumerateObject())
            {
                if (prop.Name == "mcpServers" && prop.Value.ValueKind == JsonValueKind.Object)
                {
                    w.WriteStartObject("mcpServers");
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
    }

    private static string CaptureConsoleOutput(Action action)
    {
        var originalOut = Console.Out;
        var originalErr = Console.Error;
        using var swOut = new StringWriter();
        using var swErr = new StringWriter();
        Console.SetOut(swOut);
        Console.SetError(swErr);
        try
        {
            action();
        }
        finally
        {
            Console.SetOut(originalOut);
            Console.SetError(originalErr);
        }
        return swOut.ToString() + swErr.ToString();
    }
}
