// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed record SafeSaveTransaction(
    int SchemaVersion,
    bool Ok,
    string Format,
    string Operation,
    string Mode,
    string InputPath,
    string OutputPath,
    string? TempPath,
    string? BackupPath,
    string? ManifestPath,
    bool Verified,
    IReadOnlyList<SafeSaveCheck> Checks,
    IReadOnlyDictionary<string, object?>? SemanticDelta,
    IReadOnlyDictionary<string, object?>? VisualDelta,
    IReadOnlyDictionary<string, object?>? PackageIntegrity,
    IReadOnlyList<string> Warnings
);
