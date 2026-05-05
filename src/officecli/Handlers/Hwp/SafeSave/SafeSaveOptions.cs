// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed record SafeSaveOptions(
    string InputPath,
    string OutputPath,
    bool InPlace,
    bool Backup,
    bool Verify,
    string Operation,
    string Format,
    SafeSavePolicy Policy
);
