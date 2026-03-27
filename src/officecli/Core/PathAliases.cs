// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Maps human-friendly path segment names to their OpenXML local names.
/// Allows paths like /body/paragraph[1] in addition to /body/p[1].
/// </summary>
public static class PathAliases
{
    private static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        // Word
        ["paragraph"] = "p",
        ["run"] = "r",
        ["table"] = "tbl",
        ["row"] = "tr",
        ["cell"] = "tc",
        ["hyperlink"] = "hyperlink",
        // PowerPoint
        ["slide"] = "slide",
        ["shape"] = "shape",
        ["textbox"] = "textbox",
        ["picture"] = "picture",
    };

    /// <summary>
    /// Resolve a path segment name to its canonical OpenXML local name.
    /// Returns the original name if no alias is defined.
    /// </summary>
    public static string Resolve(string name)
        => Aliases.TryGetValue(name, out var canonical) ? canonical : name;
}
