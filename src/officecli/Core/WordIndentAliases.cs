// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Core;

/// <summary>
/// <c>w:ind</c> carries each horizontal indent under two spellings:
/// transitional <c>w:left</c>/<c>w:right</c> and ISO-strict
/// <c>w:start</c>/<c>w:end</c>. Word reads either, but an element holding
/// both says two different things and a later normalizing save keeps one
/// side at random — the "indent randomly disappears" symptom. officecli
/// writes the transitional spelling everywhere; this folds any strict alias
/// into it so an element never mixes the two.
/// </summary>
internal static class WordIndentAliases
{
    /// <summary>
    /// Fold <c>w:start</c> into <c>w:left</c> and <c>w:end</c> into
    /// <c>w:right</c>. A transitional value already present wins; the alias
    /// is dropped either way. Returns true when anything changed.
    /// </summary>
    public static bool Normalize(Indentation ind)
    {
        var changed = false;
        if (ind.Start != null)
        {
            if (ind.Left == null) ind.Left = ind.Start.Value;
            ind.Start = null;
            changed = true;
        }
        if (ind.End != null)
        {
            if (ind.Right == null) ind.Right = ind.End.Value;
            ind.End = null;
            changed = true;
        }
        return changed;
    }
}
