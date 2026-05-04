// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed class SafeSaveRunner
{
    private readonly ISafeSaveManifestWriter _manifestWriter;
    private readonly Action<string, string, bool> _replaceFile;

    public SafeSaveRunner()
        : this(new SafeSaveManifestWriter(), File.Move)
    {
    }

    internal SafeSaveRunner(ISafeSaveManifestWriter manifestWriter)
        : this(manifestWriter, File.Move)
    {
    }

    internal SafeSaveRunner(ISafeSaveManifestWriter manifestWriter, Action<string, string, bool> replaceFile)
    {
        _manifestWriter = manifestWriter;
        _replaceFile = replaceFile;
    }

    public async Task<SafeSaveTransaction> RunAsync(
        SafeSaveOptions options,
        Func<string, Task> writeTempAsync,
        Func<string, Task<SafeSaveValidationResult>> validateAsync,
        CancellationToken cancellationToken)
    {
        var timestamp = DateTimeOffset.UtcNow;
        var manifestPath = _manifestWriter.BuildManifestPath(options, timestamp);
        if (options.InPlace)
            return await RunInPlaceAsync(
                options,
                writeTempAsync,
                validateAsync,
                cancellationToken,
                timestamp,
                manifestPath).ConfigureAwait(false);

        return await RunOutputAsync(
            options,
            writeTempAsync,
            validateAsync,
            cancellationToken,
            manifestPath).ConfigureAwait(false);
    }

    private async Task<SafeSaveTransaction> RunOutputAsync(
        SafeSaveOptions options,
        Func<string, Task> writeTempAsync,
        Func<string, Task<SafeSaveValidationResult>> validateAsync,
        CancellationToken cancellationToken,
        string manifestPath)
    {
        var outputPath = Path.GetFullPath(options.OutputPath);
        var inputPath = Path.GetFullPath(options.InputPath);
        if (PathsReferToSameLocation(inputPath, outputPath))
            return FinalizeTransaction(
                options,
                null,
                null,
                manifestPath,
                false,
                false,
                [new SafeSaveCheck(
                    "same-path-output",
                    false,
                    "error",
                    "Output path equals input path. Use --in-place with --backup --verify.")],
                null,
                ["Output path equals input path. Use --in-place with --backup --verify."]);

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
                var failed = FinalizeTransaction(
                    options,
                    tempPath,
                    null,
                    manifestPath,
                    false,
                    false,
                    checks,
                    validation,
                    ["safe-save required checks failed"]);
                TryDelete(tempPath);
                return failed;
            }

            _replaceFile(tempPath, outputPath, true);
            return FinalizeTransaction(options, tempPath, null, manifestPath, true, true, checks, validation, []);
        }
        catch
        {
            TryDelete(tempPath);
            throw;
        }
    }

    private async Task<SafeSaveTransaction> RunInPlaceAsync(
        SafeSaveOptions options,
        Func<string, Task> writeTempAsync,
        Func<string, Task<SafeSaveValidationResult>> validateAsync,
        CancellationToken cancellationToken,
        DateTimeOffset timestamp,
        string manifestPath)
    {
        var readinessChecks = BuildInPlaceReadinessChecks(options);
        if (readinessChecks.Count > 0)
            return FinalizeTransaction(
                options,
                null,
                null,
                manifestPath,
                false,
                false,
                readinessChecks,
                null,
                ["in-place safe save requires --backup and --verify"]);

        var inputPath = Path.GetFullPath(options.InputPath);
        var outputDirectory = Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
        var extension = Path.GetExtension(inputPath);
        var tempPath = Path.Combine(
            outputDirectory,
            $".{Path.GetFileNameWithoutExtension(inputPath)}.officecli-{SafeSaveManifestWriter.FormatTimestamp(timestamp)}-{Guid.NewGuid():N}{extension}");
        var backupPath = SafeSaveBackup.BuildBackupPath(inputPath, timestamp);

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
                var failed = FinalizeTransaction(
                    options,
                    tempPath,
                    null,
                    manifestPath,
                    false,
                    false,
                    checks,
                    validation,
                    ["safe-save required checks failed"]);
                TryDelete(tempPath);
                return failed;
            }

            var backupCheck = SafeSaveBackup.Create(inputPath, backupPath);
            checks.Add(backupCheck);
            if (!backupCheck.Ok)
            {
                var failed = FinalizeTransaction(
                    options,
                    tempPath,
                    null,
                    manifestPath,
                    false,
                    false,
                    checks,
                    validation,
                    ["backup creation failed; source was not replaced"]);
                TryDelete(tempPath);
                return failed;
            }

            var manifestProbe = ProbeManifestWrite(manifestPath);
            if (!manifestProbe.Ok)
            {
                checks.Add(manifestProbe);
                TryDelete(tempPath);
                return Transaction(
                    options,
                    tempPath,
                    backupPath,
                    manifestPath,
                    false,
                    false,
                    checks,
                    validation,
                    ["manifest write failed; source was not replaced"]);
            }
            checks.Add(manifestProbe);

            try
            {
                _replaceFile(tempPath, inputPath, true);
                checks.Add(new SafeSaveCheck(
                    "atomic-replace",
                    true,
                    "info",
                    null,
                    new Dictionary<string, object?> { ["targetPath"] = inputPath }));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                checks.Add(new SafeSaveCheck(
                    "atomic-replace",
                    false,
                    "error",
                    $"Could not replace source file: {ex.Message}",
                    new Dictionary<string, object?> { ["targetPath"] = inputPath }));
                var failed = FinalizeTransaction(
                    options,
                    tempPath,
                    backupPath,
                    manifestPath,
                    false,
                    false,
                    checks,
                    validation,
                    ["atomic replace failed; source was not marked as replaced"]);
                TryDelete(tempPath);
                return failed;
            }

            return FinalizeTransaction(
                options,
                tempPath,
                backupPath,
                manifestPath,
                true,
                true,
                checks,
                validation,
                []);
        }
        catch
        {
            TryDelete(tempPath);
            throw;
        }
    }

    private SafeSaveTransaction FinalizeTransaction(
        SafeSaveOptions options,
        string? tempPath,
        string? backupPath,
        string manifestPath,
        bool ok,
        bool verified,
        IReadOnlyList<SafeSaveCheck> checks,
        SafeSaveValidationResult? validation,
        IReadOnlyList<string> warnings)
    {
        var finalChecks = checks.ToList();
        finalChecks.Add(new SafeSaveCheck(
            "manifest-write",
            true,
            "info",
            null,
            new Dictionary<string, object?> { ["manifestPath"] = manifestPath }));
        var transaction = Transaction(
            options,
            tempPath,
            backupPath,
            manifestPath,
            ok,
            verified,
            finalChecks,
            validation,
            warnings);
        try
        {
            _manifestWriter.Write(transaction);
            return transaction;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            finalChecks.RemoveAt(finalChecks.Count - 1);
            finalChecks.Add(new SafeSaveCheck(
                "manifest-write",
                false,
                "error",
                $"Could not write safe-save manifest: {ex.Message}",
                new Dictionary<string, object?> { ["manifestPath"] = manifestPath }));
            var finalWarnings = warnings.Concat(["manifest write failed"]).ToArray();
            return Transaction(
                options,
                tempPath,
                backupPath,
                manifestPath,
                false,
                false,
                finalChecks,
                validation,
                finalWarnings);
        }
    }

    private SafeSaveCheck ProbeManifestWrite(string manifestPath)
    {
        try
        {
            _manifestWriter.Probe(manifestPath);
            return new SafeSaveCheck(
                "manifest-probe",
                true,
                "info",
                null,
                new Dictionary<string, object?> { ["manifestPath"] = manifestPath });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new SafeSaveCheck(
                "manifest-write",
                false,
                "error",
                $"Could not prepare safe-save manifest: {ex.Message}",
                new Dictionary<string, object?> { ["manifestPath"] = manifestPath });
        }
    }

    private static SafeSaveTransaction Transaction(
        SafeSaveOptions options,
        string? tempPath,
        string? backupPath,
        string? manifestPath,
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
            BackupPath: backupPath,
            ManifestPath: manifestPath,
            Verified: verified,
            Checks: checks,
            SemanticDelta: validation?.SemanticDelta,
            VisualDelta: validation?.VisualDelta,
            PackageIntegrity: validation?.PackageIntegrity,
            Warnings: warnings);

    private static List<SafeSaveCheck> BuildInPlaceReadinessChecks(SafeSaveOptions options)
    {
        var checks = new List<SafeSaveCheck>();
        if (!options.Backup)
            checks.Add(new SafeSaveCheck(
                "in-place-requires-backup",
                false,
                "error",
                "In-place safe save requires --backup."));
        if (!options.Verify)
            checks.Add(new SafeSaveCheck(
                "in-place-requires-verify",
                false,
                "error",
                "In-place safe save requires --verify."));
        return checks;
    }

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
