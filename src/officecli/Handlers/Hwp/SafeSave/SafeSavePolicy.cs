// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed record SafeSavePolicy(
    IReadOnlySet<string> RequiredChecks,
    bool BackupRequired,
    bool TransactionRequired,
    bool ValidationRequired
)
{
    public static SafeSavePolicy OutputMode(params string[] requiredChecks) => new(
        new HashSet<string>(requiredChecks, StringComparer.OrdinalIgnoreCase),
        BackupRequired: false,
        TransactionRequired: true,
        ValidationRequired: requiredChecks.Length > 0);
}
