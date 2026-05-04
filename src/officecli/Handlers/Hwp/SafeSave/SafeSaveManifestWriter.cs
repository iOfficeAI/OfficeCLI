// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal interface ISafeSaveManifestWriter
{
    string BuildManifestPath(SafeSaveOptions options, DateTimeOffset timestamp);
    void Write(SafeSaveTransaction transaction);
}

internal sealed class SafeSaveManifestWriter : ISafeSaveManifestWriter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    public string BuildManifestPath(SafeSaveOptions options, DateTimeOffset timestamp)
    {
        var inputPath = Path.GetFullPath(options.InputPath);
        var outputPath = Path.GetFullPath(options.OutputPath);
        if (options.InPlace)
            return $"{inputPath}.officecli-transaction-{FormatTimestamp(timestamp)}.json";
        return $"{outputPath}.officecli-transaction.json";
    }

    public void Write(SafeSaveTransaction transaction)
    {
        var manifestPath = transaction.ManifestPath
            ?? throw new InvalidOperationException("Safe-save manifest path is missing.");
        var directory = Path.GetDirectoryName(manifestPath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);
        var json = SafeSaveJsonMapper.ToJson(transaction).ToJsonString(JsonOptions);
        File.WriteAllText(manifestPath, json);
    }

    internal static string FormatTimestamp(DateTimeOffset timestamp)
        => timestamp.UtcDateTime.ToString("yyyyMMddHHmmss");
}
