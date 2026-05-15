// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;

namespace OfficeCli.Handlers.Hwp;

internal sealed record HwpRuntimeProbeResult(
    bool EngineRequested,
    string? BridgePath,
    string? ApiPath,
    string? RhwpPath,
    IReadOnlySet<string> ApiCommands)
{
    public bool BridgeAvailable => BridgePath != null;
    public bool ApiAvailable => ApiPath != null;
    public bool RhwpAvailable => RhwpPath != null;
    public bool ReadRenderAvailable => BridgeAvailable && (ApiAvailable || RhwpAvailable);
    public bool MutationAvailable => BridgeAvailable && ApiAvailable;
    public bool CreateBlankAvailable => ApiAvailable;
    public bool ListFieldsAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("list-fields");
    public bool ReadFieldAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("get-field");
    public bool FillFieldAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("set-field");
    public bool ReplaceTextAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("replace-text");
    public bool InsertTextAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("insert-text");
    public bool RenderPngAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("render-png");
    public bool ExportPdfAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("export-pdf");
    public bool ExportMarkdownAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("export-markdown");
    public bool ThumbnailAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("thumbnail");
    public bool DocumentInfoAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("document-info");
    public bool DiagnosticsAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("diagnostics");
    public bool DumpControlsAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("dump-controls");
    public bool DumpPagesAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("dump-pages");
    public bool ReadTableCellAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("get-cell-text");
    public bool ScanCellsAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("scan-cells");
    public bool SetTableCellAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("set-cell-text");
    public bool ConvertToEditableAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("convert-to-editable");
    public bool NativeOpAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("native-op");
    public bool SaveAsHwpAvailable => BridgeAvailable && ApiAvailable && ApiCommands.Contains("save-as-hwp");
}

internal static class HwpRuntimeProbe
{
    private const string BridgeExecutableName = "rhwp-officecli-bridge";
    private const string ApiExecutableName = "rhwp-field-bridge";
    private const string RhwpExecutableName = "rhwp";
    private static readonly string[] KnownApiCommands =
    [
        "create-blank",
        "read-text",
        "render-svg",
        "render-png",
        "export-pdf",
        "export-markdown",
        "document-info",
        "diagnostics",
        "dump-controls",
        "dump-pages",
        "thumbnail",
        "list-fields",
        "get-field",
        "set-field",
        "replace-text",
        "insert-text",
        "get-cell-text",
        "scan-cells",
        "set-cell-text",
        "convert-to-editable",
        "native-op",
        "save-as-hwp"
    ];

    public static HwpRuntimeProbeResult Probe()
    {
        var apiPath = DiscoverApiPath();
        return new(
            HwpEngineSelector.IsExperimentalBridgeEnabled(),
            DiscoverBridgePath(),
            apiPath,
            DiscoverRhwpPath(),
            DiscoverApiCommands(apiPath));
    }

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
        if (!string.IsNullOrWhiteSpace(explicitPath))
            return File.Exists(explicitPath) ? explicitPath : null;

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

    private static IReadOnlySet<string> DiscoverApiCommands(string? apiPath)
    {
        var commands = new HashSet<string>(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(apiPath) || !File.Exists(apiPath))
            return commands;

        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = apiPath,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                }
            };
            process.StartInfo.ArgumentList.Add("--help");
            process.Start();
            if (!process.WaitForExit(2_000))
            {
                try { process.Kill(); } catch { }
                return commands;
            }
            var stdout = process.StandardOutput.ReadToEnd();

            foreach (var command in KnownApiCommands)
            {
                if (stdout.Contains(command, StringComparison.Ordinal))
                    commands.Add(command);
            }
        }
        catch
        {
            return commands;
        }

        return commands;
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
