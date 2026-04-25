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

        // Apply pending update from previous background check (.update file).
        // After this returns, the current process image is still the OLD binary;
        // the NEW binary is on disk and will run on the *next* invocation.
        ApplyPendingUpdate();

        var config = LoadConfig();

        // Skill auto-refresh: if the running binary's version differs from the
        // last version that performed a refresh, push embedded skills from THIS
        // binary's resources into already-installed agent dirs. Runs once per
        // version transition (after upgrade, or on first install). Doing this
        // here — not in ApplyPendingUpdate — ensures we always copy the
        // resources of the binary actually executing, not the previous one.
        var currentVersion = GetCurrentVersion();
        if (currentVersion != null && config.LastSkillRefreshVersion != currentVersion)
        {
            try { SkillInstaller.RefreshInstalled(); } catch { /* best effort */ }
            config.LastSkillRefreshVersion = currentVersion;
            try { SaveConfig(config); } catch { /* best effort */ }
        }

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

            // Verify downloaded binary: magic bytes + smoke test
            if (!IsNativeBinary(partialPath))
            {
                try { File.Delete(partialPath); } catch { }
                return;
            }
            if (!OperatingSystem.IsWindows())
                TryChmodExecutable(partialPath);

            if (!RunVersionVerify(partialPath))
            {
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
        // Skill refresh used to live here, but ApplyPendingUpdate runs in the
        // OLD process image, so embedded resources read here are stale. The
        // refresh now happens later in CheckInBackground via a version-mismatch
        // check, which ensures the *new* binary writes its own resources on
        // its first run.
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

            // Defensive verification before swap. RunRefresh's download path
            // already runs --version on the .partial file before promoting
            // it to .update, so the canonical update flow has already been
            // verified. But .update can also be created out-of-band — by
            // failed cleanup, racing tools, accidental copies, or local user
            // mistake — and the swap would otherwise overwrite the live
            // binary with whatever is sitting there. Rerun the same check
            // here so any non-canonical .update is rejected and deleted
            // before it can corrupt the binary.

            // Step 1: cheap size sanity check. A self-contained .NET
            // single-file binary is multiple MB even when trimmed; anything
            // below 1MB is empty/text/truncated by definition.
            const long MinValidBinarySize = 1_000_000; // 1 MB
            var info = new FileInfo(updatePath);
            if (info.Length < MinValidBinarySize)
            {
                try { File.Delete(updatePath); } catch { }
                return false;
            }

            // Step 1b: native binary magic-byte check. Shell scripts, Python scripts,
            // and other interpreter-driven files (even if >1MB and exit 0) must be
            // rejected. See IsNativeBinary() for rationale.
            if (!IsNativeBinary(updatePath))
            {
                try { File.Delete(updatePath); } catch { }
                return false;
            }

            // Step 2: ensure the file is executable (Unix). Externally-
            // placed .update files often lack +x — without this, the swap
            // succeeds but the next exec fails with EACCES, bricking the
            // installed binary.
            if (!OperatingSystem.IsWindows())
                TryChmodExecutable(updatePath);

            // Step 3: smoke test — see RunVersionVerify for rationale (shebang
            // bypass, stdout regex, async pipe drain). On verify failure the
            // bad .update file is removed and the live binary is left intact.
            if (!RunVersionVerify(updatePath))
            {
                try { File.Delete(updatePath); } catch { }
                return false;
            }

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
    /// <summary>Returns 0 on success, 1 on unknown key (so callers can
    /// surface a non-zero exit code).</summary>
    internal static int HandleConfigCommand(string[] args)
    {
        const string available = "autoUpdate, log, log clear";
        var key = args[0].ToLowerInvariant();
        var config = LoadConfig();

        // officecli config log clear
        if (key == "log" && args.Length == 2 && args[1].ToLowerInvariant() == "clear")
        {
            CliLogger.Clear();
            Console.WriteLine("Log cleared.");
            return 0;
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
            {
                Console.WriteLine(value);
                return 0;
            }
            Console.Error.WriteLine($"Unknown config key: {args[0]}. Available: {available}");
            return 1;
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
                return 1;
        }

        try
        {
            Directory.CreateDirectory(ConfigDir);
            SaveConfig(config);
            Console.WriteLine($"{args[0]} = {newValue}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error saving config: {ex.Message}");
            return 1;
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

    /// <summary>
    /// Returns true if the file at <paramref name="path"/> starts with a native-binary
    /// magic-byte sequence for the current platform (Mach-O, ELF, or PE).
    /// Scripts and text files are rejected even if they happen to be >1 MB and exit 0,
    /// because on Unix the shebang exec causes .NET WaitForExit to return near-instantly
    /// (the kernel execs the interpreter process; the original pid exits), bypassing the
    /// 5-second timeout guard.
    /// </summary>
    private static bool IsNativeBinary(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var magic = new byte[4];
            if (fs.Read(magic, 0, 4) < 4) return false;
            if (OperatingSystem.IsMacOS())
                return
                    (magic[0] == 0xCF && magic[1] == 0xFA && magic[2] == 0xED && magic[3] == 0xFE) || // MH_MAGIC_64 LE (arm64/x64)
                    (magic[0] == 0xFE && magic[1] == 0xED && magic[2] == 0xFA && magic[3] == 0xCF) || // MH_MAGIC_64 BE
                    (magic[0] == 0xCA && magic[1] == 0xFE && magic[2] == 0xBA && magic[3] == 0xBE);   // FAT binary
            if (OperatingSystem.IsLinux())
                return magic[0] == 0x7F && magic[1] == 'E' && magic[2] == 'L' && magic[3] == 'F';
            if (OperatingSystem.IsWindows())
                return magic[0] == 'M' && magic[1] == 'Z';
            return true; // unknown platform — skip check
        }
        catch { return false; }
    }

    /// <summary>
    /// Make <paramref name="path"/> executable on Unix. No-op on Windows.
    /// Uses File.SetUnixFileMode (.NET 6+) instead of spawning chmod, so
    /// it's faster, has no shell-quoting concerns, and matches the
    /// approach already used in Installer.InstallBinary.
    /// </summary>
    private static void TryChmodExecutable(string path)
    {
        if (OperatingSystem.IsWindows()) return;
        try
        {
            File.SetUnixFileMode(path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute |
                UnixFileMode.GroupRead | UnixFileMode.GroupExecute |
                UnixFileMode.OtherRead | UnixFileMode.OtherExecute);
        }
        catch { /* best effort — verify will catch any resulting EACCES */ }
    }

    /// <summary>
    /// Run <c><paramref name="exePath"/> --version</c> in a sandboxed child
    /// process and return true iff it exits 0 within 5s AND stdout matches
    /// a semver string.
    ///
    /// Three subtleties this guards against:
    /// 1. <b>Shebang bypass</b>: scripts (#!/bin/sh) cause .NET WaitForExit
    ///    to return near-instantly because the kernel execs the interpreter
    ///    and the original pid exits. ExitCode=0 alone isn't enough — we
    ///    require the version regex to match.
    /// 2. <b>PipeBufferFull deadlock</b>: stdout AND stderr are redirected,
    ///    so both pipes need draining. A synchronous ReadToEnd on stdout
    ///    plus ignored stderr can deadlock if the child writes 64KB+ to
    ///    stderr before exiting. BeginOutput/ErrorReadLine pumps both
    ///    asynchronously without blocking.
    /// 3. <b>Recursion</b>: OFFICECLI_SKIP_UPDATE prevents the child's
    ///    own CheckInBackground from re-entering this code path.
    /// </summary>
    private static bool RunVersionVerify(string exePath)
    {
        try
        {
            using var verify = Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "--version",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Environment = { ["OFFICECLI_SKIP_UPDATE"] = "1" }
            });
            if (verify == null) return false;

            var stdout = new System.Text.StringBuilder();
            verify.OutputDataReceived += (_, e) => { if (e.Data != null) stdout.AppendLine(e.Data); };
            verify.ErrorDataReceived  += (_, _) => { /* drained, discarded */ };
            verify.BeginOutputReadLine();
            verify.BeginErrorReadLine();

            var exited = verify.WaitForExit(5000);
            if (!exited)
            {
                try { verify.Kill(); } catch { }
                return false;
            }
            // Ensure async readers have flushed before inspecting stdout.
            verify.WaitForExit();
            return verify.ExitCode == 0
                && Regex.IsMatch(stdout.ToString().Trim(), @"^\d+\.\d+\.\d+");
        }
        catch
        {
            return false;
        }
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
    /// <summary>Version that last successfully refreshed installed skill files.
    /// When this differs from the running binary's version, CheckInBackground
    /// triggers SkillInstaller.RefreshInstalled to push the new binary's
    /// embedded skills into already-installed agent dirs. This is the correct
    /// time to run the refresh — ApplyPendingUpdate fires it from the OLD
    /// process image, which would copy stale resources.</summary>
    public string? LastSkillRefreshVersion { get; set; }
}

[JsonSerializable(typeof(AppConfig))]
[JsonSourceGenerationOptions(WriteIndented = true, PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
internal partial class AppConfigContext : JsonSerializerContext;
