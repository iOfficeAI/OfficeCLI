// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed class SafeSaveRunner
{
    public async Task<SafeSaveTransaction> RunAsync(
        SafeSaveOptions options,
        Func<string, Task> writeTempAsync,
        Func<string, Task<SafeSaveValidationResult>> validateAsync,
        CancellationToken cancellationToken)
    {
        if (options.InPlace)
            return NotReady(options);

        var outputPath = Path.GetFullPath(options.OutputPath);
        var inputPath = Path.GetFullPath(options.InputPath);
        if (PathsReferToSameLocation(inputPath, outputPath))
            return Failed(options, null, "same-path-output", "Output path equals input path. Use --in-place after safe-save support is ready.");

        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (string.IsNullOrWhiteSpace(outputDirectory))
            outputDirectory = Directory.GetCurrentDirectory();
        Directory.CreateDirectory(outputDirectory);

        var extension = Path.GetExtension(outputPath);
        var tempPath = Path.Combine(
            outputDirectory,
            $".{Path.GetFileNameWithoutExtension(outputPath)}.officecli-{DateTimeOffset.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid():N}{extension}");

        try
        {
            await writeTempAsync(tempPath).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            var checks = new List<SafeSaveCheck>
            {
                BuildTempWriteCheck(tempPath)
            };
            var validation = await validateAsync(tempPath).ConfigureAwait(false);
            checks.AddRange(validation.Checks);

            var missing = FindMissingRequiredChecks(options.Policy, checks);
            if (missing.Count > 0)
            {
                checks.Add(new SafeSaveCheck(
                    "required-checks",
                    false,
                    "error",
                    $"Missing or failed required safe-save check(s): {string.Join(", ", missing)}"));
                TryDelete(tempPath);
                return Transaction(options, tempPath, false, false, checks, validation, ["safe-save required checks failed"]);
            }

            File.Move(tempPath, outputPath, overwrite: true);
            return Transaction(options, tempPath, true, true, checks, validation, []);
        }
        catch
        {
            TryDelete(tempPath);
            throw;
        }
    }

    private static SafeSaveTransaction NotReady(SafeSaveOptions options) => Transaction(
        options,
        null,
        false,
        false,
        [
            new SafeSaveCheck(
                "in-place-not-ready",
                false,
                "error",
                "In-place safe save requires backup, manifest, and atomic replace slice.")
        ],
        null,
        ["in-place safe save requires backup, manifest, and atomic replace slice"]);

    private static SafeSaveTransaction Failed(
        SafeSaveOptions options,
        string? tempPath,
        string checkName,
        string message) => Transaction(
        options,
        tempPath,
        false,
        false,
        [new SafeSaveCheck(checkName, false, "error", message)],
        null,
        [message]);

    private static SafeSaveTransaction Transaction(
        SafeSaveOptions options,
        string? tempPath,
        bool ok,
        bool verified,
        IReadOnlyList<SafeSaveCheck> checks,
        SafeSaveValidationResult? validation,
        IReadOnlyList<string> warnings) => new(
        SchemaVersion: 1,
        Ok: ok,
        Format: options.Format,
        Operation: options.Operation,
        Mode: options.InPlace ? "in-place" : "output",
        InputPath: Path.GetFullPath(options.InputPath),
        OutputPath: Path.GetFullPath(options.OutputPath),
        TempPath: tempPath,
        BackupPath: null,
        ManifestPath: null,
        Verified: verified,
        Checks: checks,
        SemanticDelta: validation?.SemanticDelta,
        VisualDelta: validation?.VisualDelta,
        PackageIntegrity: validation?.PackageIntegrity,
        Warnings: warnings);

    private static SafeSaveCheck BuildTempWriteCheck(string tempPath)
    {
        var file = new FileInfo(tempPath);
        var ok = file.Exists && file.Length > 0;
        return new SafeSaveCheck(
            "temp-write",
            ok,
            ok ? "info" : "error",
            ok ? null : "Temporary output was not created or is empty.",
            new Dictionary<string, object?>
            {
                ["tempPath"] = tempPath,
                ["sizeBytes"] = file.Exists ? file.Length : 0
            });
    }

    private static List<string> FindMissingRequiredChecks(
        SafeSavePolicy policy,
        IReadOnlyList<SafeSaveCheck> checks)
    {
        var okChecks = checks
            .Where(check => check.Ok)
            .Select(check => check.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return policy.RequiredChecks
            .Where(required => !okChecks.Contains(required))
            .ToList();
    }

    private static bool PathsReferToSameLocation(string firstPath, string secondPath)
    {
        var comparison = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        return string.Equals(firstPath, secondPath, comparison);
    }

    private static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { /* best effort cleanup */ }
    }
}
