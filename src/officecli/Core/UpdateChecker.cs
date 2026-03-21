// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Daily auto-update against GitHub releases.
/// - Config stored in ~/.officecli/config.json
/// - Checks at most once per day
/// - Zero performance impact: spawns background process to check and upgrade
/// - Silently skips if config dir is not writable
///
/// Also handles the __update-check__ internal command (called by the spawned background process).
/// </summary>
internal static class UpdateChecker
{
    private static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".officecli");
    private static readonly string ConfigPath = Path.Combine(ConfigDir, "config.json");
    private const string GitHubRepo = "iOfficeAI/OfficeCLI";
    private const string PrimaryBase = "https://officecli.ai";
    private const string FallbackBase = "https://github.com/iOfficeAI/OfficeCLI";
    private const int CheckIntervalHours = 24;

    /// <summary>
    /// Called on every officecli invocation. Spawns background upgrade if stale.
    /// Never blocks, never throws.
    /// </summary>
    internal static void CheckInBackground()
    {
        try
        {
            Directory.CreateDirectory(ConfigDir);
        }
        catch { return; }

        // Apply pending update from previous background check (.update file)
        ApplyPendingUpdate();

        var config = LoadConfig();

        // Respect autoUpdate setting
        if (!config.AutoUpdate) return;

        // If stale, spawn a background process to refresh (fire and forget)
        if (!config.LastUpdateCheck.HasValue ||
            (DateTime.UtcNow - config.LastUpdateCheck.Value).TotalHours >= CheckIntervalHours)
        {
            // Update timestamp immediately to prevent concurrent spawns
            config.LastUpdateCheck = DateTime.UtcNow;
            try { SaveConfig(config); } catch { }
            SpawnRefreshProcess();
        }
    }

    /// <summary>
    /// Internal command: checks for new version and auto-upgrades if available.
    /// Called by the spawned background process.
    /// </summary>
    internal static void RunRefresh()
    {
        try
        {
            var config = LoadConfig();
            var currentVersion = GetCurrentVersion();
            if (currentVersion == null) return;

            // Get latest version from redirect URL (no API, no rate limit)
            // Try primary (officecli.ai) first, fallback to GitHub
            using var handler = new HttpClientHandler { AllowAutoRedirect = false };
            using var client = new HttpClient(handler);
            client.DefaultRequestHeaders.Add("User-Agent", "OfficeCLI-UpdateChecker");
            client.Timeout = TimeSpan.FromSeconds(10);

            string? latestVersion = null;
            string resolvedBase = FallbackBase;
            foreach (var baseUrl in new[] { PrimaryBase, FallbackBase })
            {
                try
                {
                    var response = client.GetAsync($"{baseUrl}/releases/latest")
                        .GetAwaiter().GetResult();
                    var location = response.Headers.Location?.ToString();
                    if (string.IsNullOrEmpty(location)) continue;

                    var versionMatch = Regex.Match(location, @"/tag/v?(\d+\.\d+\.\d+)");
                    if (versionMatch.Success)
                    {
                        latestVersion = versionMatch.Groups[1].Value;
                        resolvedBase = baseUrl;
                        break;
                    }
                }
                catch { continue; }
            }
            if (latestVersion == null) return;

            config.LastUpdateCheck = DateTime.UtcNow;
            config.LatestVersion = latestVersion;
            SaveConfig(config);

            // Only download if newer
            if (!IsNewer(latestVersion, currentVersion)) return;

            var assetName = GetAssetName();
            if (assetName == null) return;

            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath == null) return;

            // Download binary (use the same base URL that returned the version)
            using var downloadClient = new HttpClient();
            downloadClient.DefaultRequestHeaders.Add("User-Agent", "OfficeCLI-UpdateChecker");
            downloadClient.Timeout = TimeSpan.FromMinutes(5);

            var downloadUrl = $"{resolvedBase}/releases/latest/download/{assetName}";
            var tempPath = exePath + ".update";
            using (var stream = downloadClient.GetStreamAsync(downloadUrl).GetAwaiter().GetResult())
            using (var fileStream = File.Create(tempPath))
            {
                stream.CopyTo(fileStream);
            }

            // Verify downloaded binary can start
            if (!OperatingSystem.IsWindows())
                Process.Start("chmod", $"+x \"{tempPath}\"")?.WaitForExit(3000);

            var verify = Process.Start(new ProcessStartInfo
            {
                FileName = tempPath,
                Arguments = "--version",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Environment = { ["OFFICECLI_SKIP_UPDATE"] = "1" }
            });
            var exited = verify?.WaitForExit(5000) ?? false;
            if (!exited || verify!.ExitCode != 0)
            {
                if (!exited) try { verify!.Kill(); } catch { }
                try { File.Delete(tempPath); } catch { }
                return;
            }

            if (OperatingSystem.IsWindows())
            {
                // Windows: can't replace running exe, leave .update for next startup
            }
            else
            {
                // Unix: replace in-place (safe even while running)
                var oldPath = exePath + ".old";
                try { File.Delete(oldPath); } catch { }
                File.Move(exePath, oldPath);
                File.Move(tempPath, exePath);
                try { File.Delete(oldPath); } catch { }
            }
        }
        catch
        {
            // Update timestamp even on failure to avoid retrying every command
            try
            {
                var config = LoadConfig();
                config.LastUpdateCheck = DateTime.UtcNow;
                SaveConfig(config);
            }
            catch { }
        }
    }

    /// <summary>
    /// Apply a pending update (.update file) from a previous background check.
    /// </summary>
    private static void ApplyPendingUpdate()
    {
        try
        {
            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath == null) return;

            var updatePath = exePath + ".update";
            if (!File.Exists(updatePath)) return;

            var oldPath = exePath + ".old";
            try { File.Delete(oldPath); } catch { }
            File.Move(exePath, oldPath);
            File.Move(updatePath, exePath);
            try { File.Delete(oldPath); } catch { }
        }
        catch { }
    }

    private static string? GetAssetName()
    {
        if (OperatingSystem.IsMacOS())
            return RuntimeInformation.ProcessArchitecture == Architecture.Arm64
                ? "officecli-mac-arm64" : "officecli-mac-x64";
        if (OperatingSystem.IsLinux())
            return RuntimeInformation.ProcessArchitecture == Architecture.Arm64
                ? "officecli-linux-arm64" : "officecli-linux-x64";
        if (OperatingSystem.IsWindows())
            return RuntimeInformation.ProcessArchitecture == Architecture.Arm64
                ? "officecli-win-arm64.exe" : "officecli-win-x64.exe";
        return null;
    }

    private static void SpawnRefreshProcess()
    {
        try
        {
            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath == null) return;

            var startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "__update-check__",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            var process = Process.Start(startInfo);
            // Don't wait — let it run independently
            process?.Dispose();
        }
        catch { }
    }

    /// <summary>
    /// Handle 'officecli config key [value]' command.
    /// </summary>
    internal static void HandleConfigCommand(string[] args)
    {
        var key = args[0].ToLowerInvariant();
        var config = LoadConfig();

        if (args.Length == 1)
        {
            // Read
            var value = key switch
            {
                "autoupdate" => config.AutoUpdate.ToString().ToLowerInvariant(),
                _ => null
            };
            if (value != null)
                Console.WriteLine(value);
            else
                Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: autoUpdate");
            return;
        }

        // Write
        var newValue = args[1];
        switch (key)
        {
            case "autoupdate":
                config.AutoUpdate = ParseHelpers.IsTruthy(newValue);
                break;
            default:
                Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: autoUpdate");
                return;
        }

        try
        {
            Directory.CreateDirectory(ConfigDir);
            SaveConfig(config);
            Console.WriteLine($"{args[0]} = {newValue}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error saving config: {ex.Message}");
        }
    }

    private static string? GetCurrentVersion()
    {
        var version = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (version == null) return null;
        var match = Regex.Match(version, @"^(\d+\.\d+\.\d+)");
        return match.Success ? match.Groups[1].Value : version;
    }

    private static bool IsNewer(string latest, string current)
    {
        var lp = latest.Split('.').Select(int.Parse).ToArray();
        var cp = current.Split('.').Select(int.Parse).ToArray();
        for (int i = 0; i < Math.Min(lp.Length, cp.Length); i++)
        {
            if (lp[i] > cp[i]) return true;
            if (lp[i] < cp[i]) return false;
        }
        return false;
    }

    private static UpdateConfig LoadConfig()
    {
        if (!File.Exists(ConfigPath)) return new UpdateConfig();
        try
        {
            var json = File.ReadAllText(ConfigPath);
            return JsonSerializer.Deserialize(json, UpdateConfigContext.Default.UpdateConfig) ?? new UpdateConfig();
        }
        catch { return new UpdateConfig(); }
    }

    private static void SaveConfig(UpdateConfig config)
    {
        var json = JsonSerializer.Serialize(config, UpdateConfigContext.Default.UpdateConfig);
        File.WriteAllText(ConfigPath, json);
    }
}

internal class UpdateConfig
{
    public DateTime? LastUpdateCheck { get; set; }
    public string? LatestVersion { get; set; }
    public bool AutoUpdate { get; set; } = true;
}

[JsonSerializable(typeof(UpdateConfig))]
[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
internal partial class UpdateConfigContext : JsonSerializerContext;
