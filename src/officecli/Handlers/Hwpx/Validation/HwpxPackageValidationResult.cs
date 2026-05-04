// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Handlers.Hwp.SafeSave;

namespace OfficeCli.Handlers.Hwpx.Validation;

internal sealed record HwpxPackageValidationResult(
    IReadOnlyList<SafeSaveCheck> Checks,
    IReadOnlyDictionary<string, object?> PackageIntegrity
);
