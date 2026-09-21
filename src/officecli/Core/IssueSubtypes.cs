// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    /// <summary>xlsx-only: a visible numeric/date cell with an explicit,
    /// width-stable number format cannot fit its formatted value in the visible
    /// column budget. General and unresolved formats are intentionally skipped
    /// because Excel can adapt their display to the available width. Format
    /// bucket, Warning.</summary>
    public const string NumericOverflow = "numeric_overflow";
    /// <summary>xlsx-only: a visible General-formatted numeric cell whose value
    /// needs more than 11 significant digits. Excel's General display caps
    /// there and switches to scientific notation REGARDLESS of column width,
    /// so the delivered document shows a rounded number with no visual cue and
    /// widening the column does not fix it — the remedy is an explicit number
    /// format. Disjoint from <see cref="NumericOverflow"/>, which covers the
    /// opposite case (an explicit format that cannot fit the column).
    /// Format bucket, Warning.</summary>
    public const string GeneralPrecisionLoss = "general_precision_loss";
    /// <summary>pptx-only: notesSlide raw-set passthrough references an
    /// rId (<c>r:embed</c> / <c>r:link</c>) the dump pass cannot reproduce
    /// on the replay target (e.g. a non-image rel attached to a NotesSlidePart
    /// — embedded media, OLE, etc.). The raw-set still emits, but PowerPoint
    /// shows the referenced object as a broken placeholder on open. Emitted
    /// as an UnsupportedWarning during dump; the surfaced site is the slide
    /// owning the notes (<c>/slide[N]/notes</c>).</summary>
    public const string NotesUnresolvedRid = "notes_unresolved_rid";
    /// <summary>pptx-only: a shape with its own opaque dark solid fill carries
    /// opaque dark text (fill brightness &lt; 30%, run brightness &lt; 80%) — the
    /// text is unreadable when projected. Declared-model only: the shape's
    /// explicit fill is compared against its explicit run colors, so the
    /// backdrop is unambiguous (no z-order guesswork). Scheme/inherited colors,
    /// translucent runs, and colors carrying lumMod/shade transforms are
    /// skipped to keep false positives near zero. Format bucket, Warning.</summary>
    public const string LowContrast = "low_contrast";

    /// <summary>pptx-only: text fill (textFill/textgradient) uses an advanced
    /// feature (path gradient, multiple stops, blip/image fill) that the HTML/
    /// SVG preview cannot render accurately — the preview will show a solid
    /// single color as an approximation. The OOXML document itself carries the
    /// full gradient/image fill so PowerPoint will render it correctly; the
    /// warning only flags that the CLI preview will not match what PowerPoint
    /// shows. Format bucket, Warning.</summary>
    public const string TextFillRendererApproximated = "text_fill_renderer_approximated";

    /// <summary>pptx-only: text warp (textWarp) is only marked by a class in the
    /// HTML preview; the preview cannot warp individual glyphs like PowerPoint
    /// does, so the visual approximation will not match what PowerPoint shows.
    /// Format bucket, Info.</summary>
    public const string TextWarpRendererApproximated = "text_warp_renderer_approximated";

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
        SlideFieldNotEvaluated, NotesUnresolvedRid, LowContrast,
        TextFillRendererApproximated, TextWarpRendererApproximated,
        ChartSeriesRefMissingSheet, ChartCacheStale,
        DefinedNameBroken, DefinedNameTargetMissing,
        BrokenPartRef, NumericOverflow, GeneralPrecisionLoss,
    };

    /// <summary>Subtypes that require an exact-name request rather than being
    /// scanned by default or via their broad issue bucket.</summary>
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
            + "Subtypes (returned by default and via their matching broad bucket): "
            + string.Join(", ", defaults) + ". "
            + "Opt-in only (request by exact name; not included in --type content): "
            + string.Join(", ", OptInSubtypes) + ". "
            + "Subtypes are format-specific — formula_* / chart_* / definedname_* / numeric_overflow apply to xlsx, "
            + "field_* to docx, slide_field_* / notes_unresolved_rid / broken_part_ref / low_contrast to pptx; requesting a subtype that does not apply to "
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
