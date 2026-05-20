// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Text.Json;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// A plugin executable resolved on disk together with its parsed manifest.
/// </summary>
public sealed record ResolvedPlugin(string ExecutablePath, PluginManifest Manifest);

/// <summary>
/// Locates plugin executables and reads their manifests. Implements the
/// 4-path discovery rules in docs/plugin-protocol.md §3.
///
/// Lookup is cached for the process lifetime. Negative results are cached too,
/// so a missing plugin is not re-probed on every operation.
/// </summary>
public static class PluginRegistry
{
    private const int InfoTimeoutMs = 5000;

    private static readonly Dictionary<(PluginKind kind, string ext), ResolvedPlugin?> _cache
        = new();
    private static readonly object _cacheLock = new();

    /// <summary>
    /// Resolve a plugin for the given kind + file extension. Returns null if no
    /// plugin is installed at any discovery path, or if the plugin failed
    /// <c>--info</c> probing.
    /// </summary>
    /// <param name="kind">Plugin kind we're looking for.</param>
    /// <param name="ext">File extension with leading dot, lowercase (e.g. ".doc").</param>
    public static ResolvedPlugin? FindFor(PluginKind kind, string ext)
    {
        ext = NormalizeExt(ext);
        var key = (kind, ext);

        lock (_cacheLock)
            if (_cache.TryGetValue(key, out var hit))
                return hit;

        var resolved = ResolveUncached(kind, ext);

        lock (_cacheLock)
            _cache[key] = resolved;
        return resolved;
    }

    /// <summary>
    /// Clear the resolution cache. Useful for `officecli plugins install` to
    /// force re-discovery without restarting the process. Not thread-safe with
    /// concurrent <see cref="FindFor"/> calls; callers must quiesce first.
    /// </summary>
    public static void InvalidateCache()
    {
        lock (_cacheLock) _cache.Clear();
    }

    /// <summary>
    /// Enumerate every plugin discoverable on this machine (across all kinds /
    /// extensions). Used by `officecli plugins list`. Each result reports its
    /// own kinds/extensions from the manifest.
    /// </summary>
    public static IReadOnlyList<ResolvedPlugin> EnumerateAll()
    {
        var seen = new Dictionary<string, ResolvedPlugin>(StringComparer.OrdinalIgnoreCase);
        foreach (var dir in CandidateDirectories())
        {
            if (!Directory.Exists(dir)) continue;
            foreach (var exe in EnumerateExecutablesUnder(dir))
            {
                if (seen.ContainsKey(exe)) continue;
                if (TryReadManifest(exe, out var m))
                    seen[exe] = new ResolvedPlugin(exe, m);
            }
        }
        return seen.Values.ToList();
    }

    // ---------------------------------------------------------------------
    // Discovery
    // ---------------------------------------------------------------------

    private static ResolvedPlugin? ResolveUncached(PluginKind kind, string ext)
    {
        foreach (var candidate in CandidatePaths(kind, ext))
        {
            if (!File.Exists(candidate)) continue;
            if (!TryReadManifest(candidate, out var m)) continue;
            if (!ManifestMatches(m, kind, ext)) continue;
            return new ResolvedPlugin(candidate, m);
        }
        return null;
    }

    /// <summary>
    /// The 4 discovery paths in priority order. Yields candidate executable
    /// paths; the caller is responsible for File.Exists checks.
    /// </summary>
    private static IEnumerable<string> CandidatePaths(PluginKind kind, string ext)
    {
        var kindWire = kind.ToWireString();
        var extBare = ext.TrimStart('.');

        // 1. Environment variable: $OFFICECLI_PLUGIN_<KIND>_<EXT>
        var envName = $"OFFICECLI_PLUGIN_{kindWire.ToUpperInvariant().Replace('-', '_')}_{extBare.ToUpperInvariant()}";
        var envValue = Environment.GetEnvironmentVariable(envName);
        if (!string.IsNullOrWhiteSpace(envValue))
            yield return envValue;

        // 2. User plugins directory
        var userHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        foreach (var name in PluginExeNames())
            yield return Path.Combine(userHome, ".officecli", "plugins", kindWire, extBare, name);

        // 3. Bundled directory (next to the main executable)
        var appDir = AppContext.BaseDirectory;
        foreach (var name in PluginExeNames())
            yield return Path.Combine(appDir, "plugins", kindWire, extBare, name);

        // 4. PATH lookup
        foreach (var pathExe in PathCandidates(kindWire, extBare))
            yield return pathExe;
    }

    /// <summary>
    /// All convention directories considered for full-machine enumeration. Used
    /// by <see cref="EnumerateAll"/> to discover everything regardless of kind.
    /// </summary>
    private static IEnumerable<string> CandidateDirectories()
    {
        var userHome = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        yield return Path.Combine(userHome, ".officecli", "plugins");
        yield return Path.Combine(AppContext.BaseDirectory, "plugins");
    }

    private static IEnumerable<string> PluginExeNames()
    {
        if (OperatingSystem.IsWindows())
        {
            yield return "plugin.exe";
            yield return "plugin";
        }
        else
        {
            yield return "plugin";
        }
    }

    /// <summary>
    /// PATH lookup for binaries named `officecli-<kind>-<ext>(.exe)` or, as a
    /// fallback, `officecli-<ext>(.exe)`. The latter is convenient for plugins
    /// that only implement one kind for one extension.
    /// </summary>
    private static IEnumerable<string> PathCandidates(string kindWire, string extBare)
    {
        var pathDirs = (Environment.GetEnvironmentVariable("PATH") ?? "")
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);

        var nameVariants = new[]
        {
            $"officecli-{kindWire}-{extBare}",
            $"officecli-{extBare}",
        };

        foreach (var dir in pathDirs)
        {
            foreach (var stem in nameVariants)
            {
                if (OperatingSystem.IsWindows())
                {
                    yield return Path.Combine(dir, stem + ".exe");
                    yield return Path.Combine(dir, stem);
                }
                else
                {
                    yield return Path.Combine(dir, stem);
                }
            }
        }
    }

    private static IEnumerable<string> EnumerateExecutablesUnder(string root)
    {
        // Two-level layout: <root>/<kind>/<ext>/plugin(.exe)
        IEnumerable<string> kindDirs;
        try { kindDirs = Directory.EnumerateDirectories(root); }
        catch { yield break; }

        foreach (var kindDir in kindDirs)
        {
            IEnumerable<string> extDirs;
            try { extDirs = Directory.EnumerateDirectories(kindDir); }
            catch { continue; }

            foreach (var extDir in extDirs)
            {
                foreach (var name in PluginExeNames())
                {
                    var candidate = Path.Combine(extDir, name);
                    if (File.Exists(candidate)) yield return candidate;
                }
            }
        }
    }

    // ---------------------------------------------------------------------
    // Manifest invocation
    // ---------------------------------------------------------------------

    /// <summary>
    /// Supported plugin protocol major version. The registry rejects any
    /// manifest whose <c>protocol</c> differs from this value (per §13).
    /// </summary>
    public const int SupportedProtocolVersion = 1;

    /// <summary>
    /// Run <c>plugin --info</c> and parse the resulting JSON. Returns false
    /// (and swallows the exception) if the plugin times out, exits non-zero,
    /// emits malformed JSON, or declares an incompatible protocol version.
    /// Callers should treat false the same way they treat "plugin not found".
    /// </summary>
    public static bool TryReadManifest(string executablePath, out PluginManifest manifest)
    {
        manifest = new PluginManifest();
        try
        {
            using var p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = executablePath,
                    Arguments = "--info",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                }
            };
            if (!p.Start()) return false;

            // Start the async stdout/stderr reads BEFORE WaitForExit. Synchronous
            // read-after-wait deadlocks when manifest output exceeds the pipe
            // buffer (rare for manifests, but happens when plugins emit verbose
            // diagnostics on stderr alongside --info).
            var stdoutTask = p.StandardOutput.ReadToEndAsync();
            var stderrTask = p.StandardError.ReadToEndAsync();
            if (!p.WaitForExit(InfoTimeoutMs))
            {
                try { p.Kill(entireProcessTree: true); } catch { }
                return false;
            }
            if (p.ExitCode != 0) return false;

            var stdout = stdoutTask.Result;
            _ = stderrTask.Result;
            if (string.IsNullOrWhiteSpace(stdout)) return false;

            var parsed = JsonSerializer.Deserialize(stdout, PluginJsonContext.Default.PluginManifest);
            if (parsed is null) return false;

            // Protocol gate. Mismatch is fatal — we will not load a plugin that
            // implements a different major version, since the wire format may
            // differ in ways main does not understand. Surface a one-line
            // warning so users debugging "plugin not found" can see that an
            // installed plugin was rejected for version reasons — silent
            // rejection would leave them guessing.
            if (parsed.Protocol != SupportedProtocolVersion)
            {
                Console.Error.WriteLine(
                    $"[warning] plugin at {executablePath} declares protocol={parsed.Protocol} " +
                    $"but main supports protocol={SupportedProtocolVersion}; plugin will not load.");
                return false;
            }

            manifest = parsed;
            return true;
        }
        catch
        {
            return false;
        }
    }

    // ---------------------------------------------------------------------
    // Helpers
    // ---------------------------------------------------------------------

    private static bool ManifestMatches(PluginManifest m, PluginKind kind, string ext)
    {
        var kindWire = kind.ToWireString();
        if (!m.Kinds.Contains(kindWire)) return false;
        if (!m.Extensions.Any(e => string.Equals(NormalizeExt(e), ext, StringComparison.OrdinalIgnoreCase)))
            return false;
        return true;
    }

    private static string NormalizeExt(string ext)
    {
        if (string.IsNullOrEmpty(ext)) return ext;
        if (!ext.StartsWith('.')) ext = "." + ext;
        return ext.ToLowerInvariant();
    }
}
