// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal static class SafeSaveBackup
{
    public static string BuildBackupPath(string inputPath, DateTimeOffset timestamp)
        => $"{Path.GetFullPath(inputPath)}.bak-{SafeSaveManifestWriter.FormatTimestamp(timestamp)}";

    public static SafeSaveCheck Create(string inputPath, string backupPath)
    {
        try
        {
            var directory = Path.GetDirectoryName(backupPath);
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);
            File.Copy(inputPath, backupPath, overwrite: false);
            return new SafeSaveCheck(
                "backup-created",
                true,
                "info",
                null,
                new Dictionary<string, object?>
                {
                    ["backupPath"] = backupPath,
                    ["sizeBytes"] = new FileInfo(backupPath).Length
                });
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new SafeSaveCheck(
                "backup-created",
                false,
                "error",
                $"Could not create backup before in-place replace: {ex.Message}",
                new Dictionary<string, object?> { ["backupPath"] = backupPath });
        }
    }
}
