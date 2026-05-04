// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json.Nodes;

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal static class SafeSaveJsonMapper
{
    public static JsonObject ToJson(SafeSaveTransaction transaction)
    {
        var checks = new JsonArray();
        foreach (var check in transaction.Checks)
            checks.Add((JsonNode?)ToJson(check));

        return new JsonObject
        {
            ["schemaVersion"] = transaction.SchemaVersion,
            ["ok"] = transaction.Ok,
            ["format"] = transaction.Format,
            ["operation"] = transaction.Operation,
            ["mode"] = transaction.Mode,
            ["inputPath"] = transaction.InputPath,
            ["outputPath"] = transaction.OutputPath,
            ["tempPath"] = transaction.TempPath,
            ["backupPath"] = transaction.BackupPath,
            ["manifestPath"] = transaction.ManifestPath,
            ["verified"] = transaction.Verified,
            ["checks"] = checks,
            ["warnings"] = ToJsonArray(transaction.Warnings)
        };
    }

    private static JsonObject ToJson(SafeSaveCheck check)
    {
        var obj = new JsonObject
        {
            ["name"] = check.Name,
            ["ok"] = check.Ok,
            ["severity"] = check.Severity
        };
        if (check.Message != null) obj["message"] = check.Message;
        if (check.Details != null)
        {
            var details = new JsonObject();
            foreach (var item in check.Details)
                details[item.Key] = item.Value?.ToString();
            obj["details"] = details;
        }
        return obj;
    }

    private static JsonArray ToJsonArray(IEnumerable<string> values)
    {
        var array = new JsonArray();
        foreach (var value in values)
            array.Add((JsonNode?)JsonValue.Create(value));
        return array;
    }
}
