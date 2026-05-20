// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Canonical normalization for free-form enum-like input from CLI users
/// (legend positions, fill types, conditional-format directions, chart types, …).
/// Lower-cases and strips punctuation so <c>top-right</c>, <c>top_right</c>,
/// <c>TOP_RIGHT</c>, and <c>topRight</c> all hash to the same token.
/// </summary>
internal static class SchemaKeyNormalizer
{
    /// <summary>
    /// Lower-case and strip <c>-</c>, <c>_</c>, and spaces.
    /// Null input is treated as empty.
    /// </summary>
    internal static string Normalize(string? value)
    {
        return (value ?? string.Empty)
            .ToLowerInvariant()
            .Replace("-", string.Empty)
            .Replace("_", string.Empty)
            .Replace(" ", string.Empty);
    }
}
