// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using OfficeCli.Core;

namespace OfficeCli.Handlers.Hwp;

public static class HwpBlankCreator
{
    private const int CreateTimeoutMs = 30_000;

    public static void Create(string path)
    {
        var apiPath = HwpRuntimeProbe.DiscoverApiPath();
        if (apiPath == null)
            throw new CliException("Binary .hwp blank creation requires the rhwp-field-bridge sidecar, but it was not found.")
            {
                Code = "hwp_create_dependency_missing",
                Suggestion = "Run ./dev-install.sh or set OFFICECLI_RHWP_API_BIN to rhwp-field-bridge.",
                Help = "officecli hwp doctor --json"
            };

        var fullPath = Path.GetFullPath(path);
        var dir = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);

        var result = RunCreate(apiPath, fullPath);
        if (result.ExitCode != 0)
            throw new CliException($"rhwp-field-bridge create-blank failed: {result.Stderr.Trim()}")
            {
                Code = "hwp_create_failed",
                Suggestion = result.Stdout.Trim(),
                Help = "officecli hwp doctor --json"
            };

        if (!File.Exists(fullPath) || new FileInfo(fullPath).Length == 0)
            throw new CliException("rhwp-field-bridge create-blank completed but did not create a non-empty .hwp file.")
            {
                Code = "hwp_create_output_missing",
                Help = "officecli hwp doctor --json"
            };
    }

    private static ProcessResult RunCreate(string apiPath, string outputPath)
    {
        var psi = new ProcessStartInfo
        {
            FileName = apiPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        psi.ArgumentList.Add("create-blank");
        psi.ArgumentList.Add("--output");
        psi.ArgumentList.Add(outputPath);
        psi.ArgumentList.Add("--json");

        using var process = Process.Start(psi)
            ?? throw new CliException("Failed to start rhwp-field-bridge.")
            {
                Code = "hwp_create_start_failed",
                Help = "officecli hwp doctor --json"
            };

        if (!process.WaitForExit(CreateTimeoutMs))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            throw new CliException($"rhwp-field-bridge create-blank timed out after {CreateTimeoutMs}ms.")
            {
                Code = "hwp_create_timeout",
                Help = "officecli hwp doctor --json"
            };
        }

        return new ProcessResult(
            process.ExitCode,
            process.StandardOutput.ReadToEnd(),
            process.StandardError.ReadToEnd());
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
