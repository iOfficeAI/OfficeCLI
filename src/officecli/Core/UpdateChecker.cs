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
    internal static readonly string ConfigDir = Path.Combine(
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
            var finalPath = exePath + ".update";
            // Stage download to .partial so a crashed/killed download never leaves
            // a truncated PE at the canonical .update path that ApplyPendingUpdate would apply.
            var partialPath = exePath + ".update.partial";
            try { File.Delete(partialPath); } catch { }
            using (var stream = downloadClient.GetStreamAsync(downloadUrl).GetAwaiter().GetResult())
            using (var fileStream = File.Create(partialPath))
            {
                stream.CopyTo(fileStream);
            }

            // Verify downloaded binary can start
            if (!OperatingSystem.IsWindows())
                Process.Start("chmod", $"+x \"{partialPath}\"")?.WaitForExit(3000);

            var verify = Process.Start(new ProcessStartInfo
            {
                FileName = partialPath,
                Arguments = "--version",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Environment = { ["OFFICECLI_SKIP_UPDATE"] = "1" }
            });
            if (verify == null)
            {
                try { File.Delete(partialPath); } catch { }
                return;
            }
            var exited = verify.WaitForExit(5000);
            if (!exited || verify.ExitCode != 0)
            {
                if (!exited) try { verify.Kill(); } catch { }
                try { File.Delete(partialPath); } catch { }
                return;
            }

            // Atomically promote .partial -> .update only after verification.
            try { File.Delete(finalPath); } catch { }
            try
            {
                File.Move(partialPath, finalPath, overwrite: true);
            }
            catch
            {
                try { File.Delete(partialPath); } catch { }
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
                File.Move(exePath, oldPath, overwrite: true);
                try
                {
                    File.Move(finalPath, exePath, overwrite: true);
                }
                catch
                {
                    // Rollback: restore original if new file failed to move
                    try { File.Move(oldPath, exePath, overwrite: true); } catch { }
                    return;
                }
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
        var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
        if (exePath == null) return;
        TryApplyPendingUpdate(exePath);
    }

    /// <summary>
    /// Test seam: applies a pending <c>{exePath}.update</c> by swapping it into place.
    /// Note: only the canonical <c>.update</c> file is applied — a stale
    /// <c>.update.partial</c> from an interrupted download is intentionally ignored.
    /// </summary>
    internal static bool TryApplyPendingUpdate(string exePath)
    {
        try
        {
            var updatePath = exePath + ".update";
            if (!File.Exists(updatePath)) return false;

            var oldPath = exePath + ".old";
            try { File.Delete(oldPath); } catch { }
            File.Move(exePath, oldPath, overwrite: true);
            try
            {
                File.Move(updatePath, exePath, overwrite: true);
            }
            catch
            {
                // Rollback: restore original
                try { File.Move(oldPath, exePath, overwrite: true); } catch { }
                return false;
            }
            try { File.Delete(oldPath); } catch { }
            return true;
        }
        catch { return false; }
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
                CreateNoWindow = true,
                // Redirect child stdio away from the parent's console. Without
                // these flags the child inherits the parent's stdout/stderr,
                // which is a problem in two concrete scenarios:
                //   (a) the parent is an MCP server — its stdout carries the
                //       JSON-RPC protocol stream, and any byte the update-
                //       check writes there would corrupt the protocol and
                //       disconnect the MCP client;
                //   (b) the parent is an interactive shell command that exits
                //       before the child finishes — the child's "downloaded
                //       v1.2.3" or error messages would then surface on the
                //       user's terminal at a seemingly random later moment.
                // We redirect to pipes and never Read them; the pipes are
                // closed when the child exits. This cannot break the upgrade
                // itself: RunRefresh() only writes to stdout/stderr for
                // debugging/never (it's silent-on-success, silent-on-failure
                // by design), and the download / verify / File.Move chain
                // doesn't touch the console stream at all.
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                RedirectStandardInput = true
            };

            var process = Process.Start(startInfo);
            if (process == null) return;
            // Close our end of stdin immediately so the child sees EOF if it
            // ever tries to read (defensive — RunRefresh doesn't read stdin).
            try { process.StandardInput.Close(); } catch { }
            // Don't wait, don't Read the redirected streams. When the child
            // exits the OS closes its side of the pipes; the .NET runtime's
            // SIGCHLD reaper waits on it so it never becomes a zombie even
            // though we never call WaitForExit.
            process.Dispose();
        }
        catch { }
    }

    /// <summary>
    /// Handle 'officecli config key [value]' command.
    /// </summary>
    internal static void HandleConfigCommand(string[] args)
    {
        const string available = "autoUpdate, log, log clear";
        var key = args[0].ToLowerInvariant();
        var config = LoadConfig();

        // officecli config log clear
        if (key == "log" && args.Length == 2 && args[1].ToLowerInvariant() == "clear")
        {
            CliLogger.Clear();
            Console.WriteLine("Log cleared.");
            return;
        }

        if (args.Length == 1)
        {
            // Read
            var value = key switch
            {
                "autoupdate" => config.AutoUpdate.ToString().ToLowerInvariant(),
                "log" => config.Log.ToString().ToLowerInvariant(),
                _ => null
            };
            if (value != null)
                Console.WriteLine(value);
            else
                Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: {available}");
            return;
        }

        // Write
        var newValue = args[1];
        switch (key)
        {
            case "autoupdate":
                config.AutoUpdate = ParseHelpers.IsTruthy(newValue);
                break;
            case "log":
                config.Log = ParseHelpers.IsTruthy(newValue);
                break;
            default:
                Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: {available}");
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
        return lp.Length > cp.Length;
    }

    internal static AppConfig LoadConfig()
    {
        if (!File.Exists(ConfigPath)) return new AppConfig();
        try
        {
            var json = File.ReadAllText(ConfigPath);
            return JsonSerializer.Deserialize(json, AppConfigContext.Default.AppConfig) ?? new AppConfig();
        }
        catch { return new AppConfig(); }
    }

    internal static void SaveConfig(AppConfig config)
    {
        Directory.CreateDirectory(ConfigDir);
        var json = JsonSerializer.Serialize(config, AppConfigContext.Default.AppConfig);
        File.WriteAllText(ConfigPath, json);
    }

    internal static string? GetCurrentVersionPublic() => GetCurrentVersion();

    internal static bool IsNewerPublic(string latest, string current) => IsNewer(latest, current);
}

internal class AppConfig
{
    public DateTime? LastUpdateCheck { get; set; }
    public string? LatestVersion { get; set; }
    public bool AutoUpdate { get; set; } = true;
    public bool Log { get; set; }
    public string? InstalledBinaryVersion { get; set; }
}

[JsonSerializable(typeof(AppConfig))]
[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
internal partial class AppConfigContext : JsonSerializerContext;
