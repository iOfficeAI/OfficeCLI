// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed record SafeSaveValidationResult(
    IReadOnlyList<SafeSaveCheck> Checks,
    IReadOnlyDictionary<string, object?>? SemanticDelta = null,
    IReadOnlyDictionary<string, object?>? VisualDelta = null,
    IReadOnlyDictionary<string, object?>? PackageIntegrity = null
)
{
    public static SafeSaveValidationResult FromChecks(IReadOnlyList<SafeSaveCheck> checks)
        => new(checks);
}
