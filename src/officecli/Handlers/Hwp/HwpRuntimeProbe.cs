// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;

namespace OfficeCli.Handlers.Hwp;

internal sealed record HwpRuntimeProbeResult(
    bool EngineRequested,
    string? BridgePath,
    string? ApiPath,
    string? RhwpPath)
{
    public bool BridgeAvailable => BridgePath != null;
    public bool ApiAvailable => ApiPath != null;
    public bool RhwpAvailable => RhwpPath != null;
    public bool ReadRenderAvailable => BridgeAvailable && (ApiAvailable || RhwpAvailable);
    public bool MutationAvailable => BridgeAvailable && ApiAvailable;
    public bool CreateBlankAvailable => ApiAvailable;
}

internal static class HwpRuntimeProbe
{
    private const string BridgeExecutableName = "rhwp-officecli-bridge";
    private const string ApiExecutableName = "rhwp-field-bridge";
    private const string RhwpExecutableName = "rhwp";

    public static HwpRuntimeProbeResult Probe()
        => new(
            HwpEngineSelector.IsExperimentalBridgeEnabled(),
            DiscoverBridgePath(),
            DiscoverApiPath(),
            DiscoverRhwpPath());

    public static string? DiscoverBridgePath()
        => DiscoverExecutable(
            Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BRIDGE_PATH"),
            CandidateNames(BridgeExecutableName, includeDll: true));

    public static string? DiscoverApiPath()
        => DiscoverExecutable(
            Environment.GetEnvironmentVariable("OFFICECLI_RHWP_API_BIN"),
            CandidateNames(ApiExecutableName, includeDll: false));

    public static string? DiscoverRhwpPath()
        => DiscoverExecutable(
            Environment.GetEnvironmentVariable("OFFICECLI_RHWP_BIN"),
            CandidateNames(RhwpExecutableName, includeDll: false));

    private static string[] CandidateNames(string baseName, bool includeDll)
    {
        var names = new List<string> { baseName };
        if (OperatingSystem.IsWindows())
            names.Add(baseName + ".exe");
        if (includeDll)
            names.Add(baseName + ".dll");
        return names.ToArray();
    }

    private static string? DiscoverExecutable(string? explicitPath, string[] names)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath) && File.Exists(explicitPath))
            return explicitPath;

        foreach (var dir in CandidateDirectories())
        {
            foreach (var name in names)
            {
                var candidate = Path.Combine(dir, name);
                if (File.Exists(candidate)) return candidate;
            }
        }

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            foreach (var name in names)
            {
                var candidate = Path.Combine(dir, name);
                if (File.Exists(candidate)) return candidate;
            }
        }

        return null;
    }

    private static IEnumerable<string> CandidateDirectories()
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var dir in new[]
        {
            AppContext.BaseDirectory,
            Path.GetDirectoryName(Environment.ProcessPath ?? ""),
            Path.GetDirectoryName(Process.GetCurrentProcess().MainModule?.FileName ?? ""),
            Directory.GetCurrentDirectory()
        })
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            var full = Path.GetFullPath(dir);
            if (seen.Add(full)) yield return full;
        }
    }
}
