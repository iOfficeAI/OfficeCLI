// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
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
    // Resolved per-call rather than cached so tests can override $HOME between
    // cases without restarting the process. Production behavior is unchanged —
    // $HOME never moves under a running officecli invocation.
    internal static string ConfigDir => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".officecli");
    private static string ConfigPath => Path.Combine(ConfigDir, "config.json");
    private const string GitHubRepo = "iOfficeAI/OfficeCLI";
    // PrimaryBase is the project-controlled mirror (Cloudflare-fronted nginx on
    // a VPS that periodically syncs github releases). FallbackBase is the
    // upstream of last resort. Order matters: the mirror is exercised on every
    // daily check so issues surface fast, and CF edge caching makes it the
    // fastest path for most users; github is the safety net when CF or the
    // mirror is unreachable.
    private const string PrimaryBase = "https://d.officecli.ai";
    private const string FallbackBase = "https://github.com/iOfficeAI/OfficeCLI";
    private const int CheckIntervalHours = 24;

    /// <summary>
    /// Called on every officecli invocation. Spawns background upgrade if stale.
    /// Never blocks, never throws.
    /// </summary>
    internal static void CheckInBackground()
    {
        // Best-effort: SaveConfig falls back to $TMPDIR inside containers when
        // home is read-only, so a CreateDirectory failure here is not fatal.
        try { Directory.CreateDirectory(ConfigDir); } catch { /* continue */ }

        var config = LoadConfig();

        // Apply pending update from previous background check (.update file).
        // After this returns, the current process image is still the OLD binary;
        // the NEW binary is on disk and will run on the *next* invocation.
        //
        // GATED ON autoUpdate: the pending file sits NEXT TO THE BINARY
        // (shared across users) while the opt-out lives in per-HOME
        // config.json — so a refresh spawned under some other HOME (default
        // = enabled) could stage a .update that a later invocation applied
        // even though THIS user had autoUpdate off, silently swapping a
        // version-pinned binary. When the setting is off, the staged file is
        // left in place, just never applied by THIS user — deleting it would
        // only make an enabled sibling process re-download it next cycle.
        if (config.AutoUpdate)
            ApplyPendingUpdate();

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
            // No persisted timestamp → next invocation thinks stale → respawns.
            // Bail rather than burn an HTTP roundtrip we'll be repeating forever.
            if (!SaveConfig(config)) return;
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

            // Piggyback the diagram render's mermaid.js cache refresh on this daily
            // background pass (we're already once-per-24h and already reaching the
            // mirror). Only revalidates an existing cache; independent of the binary
            // update below, so it runs even for package-managed (Homebrew) installs.
            try { Diagram.MermaidImageRenderer.RefreshCacheIfPresent(); } catch { /* best effort */ }

            var currentVersion = GetCurrentVersion();
            if (currentVersion == null) return;

            // Get latest version by following the full redirect chain and
            // parsing the version out of the *final* URL (no API, no rate limit).
            //
            // Why follow the whole chain instead of reading the first Location:
            // PrimaryBase 302s to /releases/tag/vX.Y.Z (its own URL, served
            // from a mirror-maintained version file) and FallbackBase 302s
            // through github's 2-hop release chain. Either way the *final*
            // URL after redirects carries /tag/vX.Y.Z, so reading the first
            // Location wouldn't work for the github fallback path.
            using var handler = new HttpClientHandler { AllowAutoRedirect = true };
            using var client = new HttpClient(handler);
            // UA carries running version so the mirror can produce a version
            // distribution from access logs without any extra telemetry.
            AddUserAgent(client, currentVersion);
            client.Timeout = TimeSpan.FromSeconds(10);

            string? latestVersion = null;
            string resolvedBase = FallbackBase;
            foreach (var baseUrl in new[] { PrimaryBase, FallbackBase })
            {
                try
                {
                    // HEAD avoids downloading the release page body; we only need
                    // the final URL after redirects.
                    using var req = new HttpRequestMessage(HttpMethod.Head, $"{baseUrl}/releases/latest");
                    var response = client.SendAsync(req).GetAwaiter().GetResult();
                    var finalUrl = response.RequestMessage?.RequestUri?.ToString();
                    if (string.IsNullOrEmpty(finalUrl)) continue;

                    var versionMatch = Regex.Match(finalUrl, @"/tag/v?(\d+\.\d+\.\d+)");
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

            // Package-managed installs (Homebrew) must never self-replace their
            // binary: Homebrew owns the file and upgrades happen via `brew
            // upgrade`. Per Homebrew's Acceptable Formulae policy, software
            // self-update "should be disabled" for formulae. The latest version
            // is recorded above — so the daily check's telemetry and the version
            // readout stay intact — we simply stop before the download/swap.
            if (IsPackageManaged()) return;

            // Only download if newer
            if (!IsNewer(latestVersion, currentVersion)) return;

            var assetName = GetAssetName();
            if (assetName == null) return;

            var exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName;
            if (exePath == null) return;

            // Download binary using the version-pinned URL. This URL is
            // immutable: the same URL always points at the same bytes, which
            // lets CF cache it on the edge essentially forever, and removes
            // the race where /releases/latest/download/X could return v1.0.95
            // bytes while /releases/latest already reports v1.0.96. Works
            // identically on the github fallback (canonical github URL form).
            using var downloadClient = new HttpClient();
            AddUserAgent(downloadClient, currentVersion);
            downloadClient.Timeout = TimeSpan.FromMinutes(5);

            var downloadUrl = $"{resolvedBase}/releases/download/v{latestVersion}/{assetName}";
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

            // Integrity gate: verify the SHA256 of the downloaded asset against
            // the release's SHA256SUMS BEFORE running it. This is the same file
            // and verification install.sh performs; the self-updater previously
            // skipped it and trusted magic-bytes + `--version` alone, which means
            // an unauthenticated binary was executed during RunVersionVerify.
            // The SHA256SUMS lives at the same version-pinned (immutable) path as
            // the asset, so a corrupted/truncated CDN response or partial poison
            // is rejected. (A fully compromised mirror can still serve a matching
            // pair — defeating that needs a signature against a pinned key, which
            // is a separate, larger change.) Missing/unparseable/mismatched →
            // hard-fail: delete the partial and retry on the next daily check.
            // Every current release ships SHA256SUMS, so this never blocks a
            // legitimate update.
            var checksumsUrl = $"{resolvedBase}/releases/download/v{latestVersion}/SHA256SUMS";
            if (!VerifyChecksum(partialPath, assetName, checksumsUrl, downloadClient))
            {
                try { File.Delete(partialPath); } catch { }
                return;
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
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
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
        foreach (var path in ConfigPathCandidates())
        {
            if (!File.Exists(path)) continue;
            try
            {
                var json = File.ReadAllText(path);
                return JsonSerializer.Deserialize(json, AppConfigContext.Default.AppConfig)
                       ?? new AppConfig();
            }
            catch { continue; }
        }
        return new AppConfig();
    }

    /// <summary>Persist <paramref name="config"/>. Returns false when no
    /// candidate path is writable (e.g. read-only home outside a container,
    /// or read-only home + read-only tmp inside one). Callers should treat
    /// false as "we can't throttle next invocation" and skip work that would
    /// otherwise repeat on every command.</summary>
    internal static bool SaveConfig(AppConfig config)
    {
        var json = JsonSerializer.Serialize(config, AppConfigContext.Default.AppConfig);
        foreach (var path in ConfigPathCandidates())
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.WriteAllText(path, json);
                return true;
            }
            catch { continue; }
        }
        return false;
    }

    /// <summary>Order: $HOME first; $TMPDIR appended only when running in a
    /// container. The /tmp fallback exists for read-only-rootfs containers
    /// (docker --read-only, K8s readOnlyRootFilesystem, Lambda, Cloud Run)
    /// where ~/.officecli can't be written but /tmp is a tmpfs. Iteration
    /// stops at the first success — non-container hosts never touch /tmp,
    /// and containers with writable home never touch /tmp either.</summary>
    private static IEnumerable<string> ConfigPathCandidates()
    {
        yield return ConfigPath;
        if (IsInContainer())
            yield return Path.Combine(Path.GetTempPath(), "officecli-config.json");
    }

    /// <summary>True when running inside docker/k8s/podman or a major
    /// serverless runtime. Used both to enable the /tmp config fallback and
    /// to tag the UpdateChecker User-Agent with <c>(container)</c>.</summary>
    private static bool IsInContainer()
    {
        if (Environment.GetEnvironmentVariable("KUBERNETES_SERVICE_HOST") != null) return true;
        if (Environment.GetEnvironmentVariable("AWS_LAMBDA_FUNCTION_NAME") != null) return true;
        if (Environment.GetEnvironmentVariable("K_SERVICE") != null) return true;          // Cloud Run
        if (Environment.GetEnvironmentVariable("FUNCTION_TARGET") != null) return true;    // GCP Functions
        if (OperatingSystem.IsLinux())
        {
            if (File.Exists("/.dockerenv"))    return true;
            if (File.Exists("/.containerenv")) return true;  // Podman
        }
        return false;
    }

    /// <summary>UA shape: <c>OfficeCLI/{ver}</c>, optionally followed by
    /// <c>(container)</c>. Server-side log analytics use the OfficeCLI/
    /// prefix to identify update-check traffic and the comment to split
    /// container vs bare-metal device counts.</summary>
    private static void AddUserAgent(HttpClient client, string currentVersion)
    {
        var ua = $"OfficeCLI/{currentVersion}";
        if (IsInContainer()) ua += " (container)";
        client.DefaultRequestHeaders.Add("User-Agent", ua);
    }

    /// <summary>
    /// Verify the SHA256 of the downloaded asset against the release's
    /// <c>SHA256SUMS</c> manifest. Returns true only when the manifest is
    /// fetched, the asset's line is present, and the hash matches exactly.
    /// Any failure (network, missing line, mismatch) returns false so the
    /// caller discards the download — never executes an unverified binary.
    /// Mirrors the verification install.sh performs at install time.
    /// </summary>
    private static bool VerifyChecksum(string filePath, string assetName, string checksumsUrl, HttpClient client)
    {
        try
        {
            string manifest;
            using (var resp = client.GetAsync(checksumsUrl).GetAwaiter().GetResult())
            {
                if (!resp.IsSuccessStatusCode) return false;
                manifest = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();
            }
            return MatchChecksumManifest(manifest, assetName, filePath);
        }
        catch { return false; }
    }

    /// <summary>
    /// Pure (network-free) core of <see cref="VerifyChecksum"/>: given the text
    /// of a <c>SHA256SUMS</c> manifest, return true iff it lists
    /// <paramref name="assetName"/> and that hash matches the SHA256 of the file
    /// at <paramref name="filePath"/>. Split out so the matching logic is unit
    /// testable without an HTTP server.
    /// </summary>
    internal static bool MatchChecksumManifest(string manifest, string assetName, string filePath)
    {
        try
        {
            // SHA256SUMS lines are "<hex>  <filename>". Match the filename
            // column exactly (not a substring) so e.g. "officecli-win-arm64"
            // can't be satisfied by the "officecli-win-arm64.exe" line.
            string? expected = null;
            foreach (var line in manifest.Split('\n'))
            {
                var parts = line.Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2 && string.Equals(parts[1], assetName, StringComparison.Ordinal))
                {
                    expected = parts[0];
                    break;
                }
            }
            if (string.IsNullOrEmpty(expected)) return false;

            string actual;
            using (var fs = File.OpenRead(filePath))
                actual = Convert.ToHexString(SHA256.HashData(fs));

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
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
                // CONSISTENCY(child-stream-encoding): see BlankDocCreator.
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8,
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

    /// <summary>
    /// True when the running binary is owned by a package manager (Homebrew),
    /// which performs its own upgrades. Such installs must never self-replace
    /// their binary — RunRefresh records the latest version for telemetry but
    /// stops before downloading. Mirrors Homebrew's Acceptable Formulae policy
    /// ("self-update functionality should be disabled" for formulae).
    ///
    /// Resolves symlinks first: Homebrew exposes the formula binary as
    /// <c>/opt/homebrew/bin/officecli</c> → a symlink into <c>.../Cellar/...</c>,
    /// and Environment.ProcessPath may report the symlink rather than the
    /// resolved Cellar/Caskroom target. Both raw and resolved paths are checked
    /// so detection works regardless of which one the OS hands back. The
    /// <c>/Cellar/</c> + <c>/Caskroom/</c> substrings are prefix-agnostic, so
    /// they also catch Intel (<c>/usr/local/Cellar</c>) and Linuxbrew
    /// (<c>/home/linuxbrew/.linuxbrew/Cellar</c>) layouts.
    /// </summary>
    internal static bool IsPackageManaged()
        => IsPackageManagedPath(Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName);

    internal static bool IsPackageManagedPath(string? exePath)
    {
        if (string.IsNullOrEmpty(exePath)) return false;
        var candidates = new List<string> { exePath };
        try
        {
            var resolved = File.ResolveLinkTarget(exePath, returnFinalTarget: true)?.FullName;
            if (!string.IsNullOrEmpty(resolved)) candidates.Add(resolved);
        }
        catch { /* not a link or inaccessible — fall back to the raw path */ }
        return candidates.Any(p =>
            p.Contains("/Cellar/", StringComparison.Ordinal) ||
            p.Contains("/Caskroom/", StringComparison.Ordinal));
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
