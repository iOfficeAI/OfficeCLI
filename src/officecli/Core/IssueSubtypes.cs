// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Central catalogue of <c>view issues --type</c> accepted values. Single
/// source of truth so the CLI front-end (CommandBuilder.View) and the
/// resident server (ResidentServer.ExecuteView) reject typos identically
/// and the cross-handler protocol documentation cannot drift from what the
/// validator actually accepts.
/// </summary>
public static class IssueSubtypes
{
    public const string FormulaNotEvaluated = "formula_not_evaluated";
    public const string FormulaCacheStale = "formula_cache_stale";
    public const string FormulaRefMissingSheet = "formula_ref_missing_sheet";
    public const string FormulaEvalError = "formula_eval_error";
    public const string FieldNotEvaluated = "field_not_evaluated";
    public const string FieldCacheStale = "field_cache_stale";
    public const string SlideFieldNotEvaluated = "slide_field_not_evaluated";
    public const string ChartSeriesRefMissingSheet = "chart_series_ref_missing_sheet";
    public const string ChartCacheStale = "chart_cache_stale";
    public const string DefinedNameBroken = "definedname_broken";
    public const string DefinedNameTargetMissing = "definedname_target_missing";
    public const string BrokenPartRef = "broken_part_ref";

    /// <summary>Broad IssueType bucket names — the canonical surface shown
    /// in error messages and help. Single-letter aliases (<see cref="BucketAliases"/>)
    /// are accepted by Validate but kept out of the user-facing list so the
    /// canonical-vs-alias distinction is visible.</summary>
    public static readonly string[] BucketNames =
        new[] { "format", "content", "structure" };

    /// <summary>Single-letter aliases accepted in addition to the canonical
    /// bucket names. Kept separate from <see cref="BucketNames"/> so error
    /// listings don't expose them as first-class values.</summary>
    public static readonly string[] BucketAliases =
        new[] { "f", "c", "s" };

    /// <summary>Combined accepted bucket inputs (canonical + aliases).</summary>
    public static readonly string[] ValidBuckets =
        BucketNames.Concat(BucketAliases).ToArray();

    /// <summary>Every subtype the <c>view issues</c> filter accepts by name.</summary>
    public static readonly string[] ValidSubtypes = new[]
    {
        FormulaNotEvaluated, FormulaCacheStale, FormulaRefMissingSheet, FormulaEvalError,
        FieldNotEvaluated, FieldCacheStale,
        SlideFieldNotEvaluated,
        ChartSeriesRefMissingSheet, ChartCacheStale,
        DefinedNameBroken, DefinedNameTargetMissing,
        BrokenPartRef,
    };

    /// <summary>Subtypes that are scanned by default and surface under
    /// <c>--type content</c>. Opt-in subtypes (currently only
    /// <see cref="ChartCacheStale"/>) require an exact-name request.</summary>
    public static readonly string[] OptInSubtypes = new[] { ChartCacheStale };

    /// <summary>One-line summary suitable for the CLI <c>--type</c> help
    /// text. Generated from <see cref="ValidSubtypes"/> so the help cannot
    /// drift from the validator.</summary>
    public static string TypeHelpDescription()
    {
        var defaults = ValidSubtypes.Where(s => !OptInSubtypes.Contains(s));
        return "Issue type filter. Broad buckets: "
            + string.Join(", ", BucketNames)
            + " (alias " + string.Join(", ", BucketAliases) + "). "
            + "Subtypes (Content bucket, returned by default and via --type content): "
            + string.Join(", ", defaults) + ". "
            + "Opt-in only (request by exact name; not included in --type content): "
            + string.Join(", ", OptInSubtypes) + ". "
            + "Subtypes are format-specific — formula_* / chart_* / definedname_* apply to xlsx, "
            + "field_* to docx, slide_field_* / broken_part_ref to pptx; requesting a subtype that does not apply to "
            + "the queried file returns count=0 (not an error). "
            + "All values are case-insensitive and surrounding whitespace is trimmed.";
    }

    /// <summary>
    /// Validate a user-supplied <c>--type</c> argument and return the
    /// canonicalised form. Null, empty, and whitespace-only inputs are
    /// normalised to null (treated as "no filter"). Surrounding whitespace
    /// is trimmed so values copied from shells with extra spaces still
    /// match. Recognised buckets and subtypes (case-insensitive) pass
    /// through unchanged. Anything else raises <see cref="CliException"/>
    /// with the full valid list — turning silent typos into a clear
    /// failure on both the CLI front-end and the resident-server fan-out.
    /// </summary>
    public static string? Validate(string? issueType)
    {
        if (string.IsNullOrWhiteSpace(issueType)) return null;
        var trimmed = issueType.Trim();
        var canonical = trimmed.ToLowerInvariant();
        foreach (var v in ValidBuckets) if (v == canonical) return trimmed;
        foreach (var v in ValidSubtypes) if (v == canonical) return trimmed;
        var all = ValidBuckets.Concat(ValidSubtypes).ToArray();
        throw new CliException(
            $"Invalid --type value: '{issueType}'. Valid buckets: {string.Join(", ", BucketNames)} (alias {string.Join(", ", BucketAliases)}). Valid subtypes: {string.Join(", ", ValidSubtypes)}.")
        { Code = "invalid_issue_type", ValidValues = all };
    }
}
