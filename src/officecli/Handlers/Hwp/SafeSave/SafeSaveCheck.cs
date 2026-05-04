// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Handlers.Hwp.SafeSave;

internal sealed record SafeSaveCheck(
    string Name,
    bool Ok,
    string Severity,
    string? Message = null,
    IReadOnlyDictionary<string, object?>? Details = null
);
