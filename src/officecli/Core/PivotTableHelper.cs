// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Helper for building and reading pivot tables.
/// Manages PivotTableCacheDefinitionPart (workbook-level) and PivotTablePart (worksheet-level).
/// </summary>
internal static partial class PivotTableHelper
{
    // Sentinel used to represent Excel error cells (DataType=Error, e.g. #DIV/0!)
    // in the string[] columnData arrays passed between ReadSourceData and BuildCacheField.
    // This value never appears in normal cell text (U+0001 prefix makes it XML-illegal
    // for ordinary strings, so SanitizeXmlText would have stripped it). BuildCacheField
    // emits ErrorItem instead of StringItem when it sees this sentinel.
    internal const string ErrorCellSentinel = "\x01#ERROR";

    // ==================== XML text sanitization (R2-2) ====================
    //
    // XML 1.0 only permits a narrow set of character code points in element
    // content: Tab (U+0009), LF (U+000A), CR (U+000D), and anything in
    // [U+0020..U+D7FF] ∪ [U+E000..U+FFFD] ∪ [U+10000..U+10FFFF]. Everything
    // else — including the NUL byte — causes XmlWriter to throw
    // ArgumentException at save time, which tore down PivotCacheDefinition.Save
    // whenever a source cell contained a stray U+0000 (see FuzzPivotRound2Tests
    // Add_Pivot_NulCharInRowValue_ShouldNotThrow).
    //
    // Sanitization is applied ONLY to strings that get embedded in the pivot
    // cache (sharedItems <s v="..."/> and fieldGroup <groupItems>). The
    // original cell values in the source sheet are untouched — we just want
    // the cache write to succeed. Unpaired surrogates are also stripped so we
    // don't turn one invalid form into another.
    internal static string SanitizeXmlText(string? s)
    {
        if (string.IsNullOrEmpty(s)) return s ?? string.Empty;
        System.Text.StringBuilder? sb = null;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            bool ok;
            if (c == '\t' || c == '\n' || c == '\r') ok = true;
            else if (c < 0x20) ok = false;
            else if (c == 0xFFFE || c == 0xFFFF) ok = false;
            else if (char.IsHighSurrogate(c))
            {
                if (i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                {
                    if (sb != null) { sb.Append(c); sb.Append(s[i + 1]); }
                    i++;
                    continue;
                }
                ok = false;
            }
            else if (char.IsLowSurrogate(c)) ok = false; // unpaired trailing surrogate
            else ok = true;

            if (ok)
            {
                sb?.Append(c);
            }
            else
            {
                if (sb == null)
                {
                    sb = new System.Text.StringBuilder(s.Length);
                    sb.Append(s, 0, i);
                }
                // Drop the invalid code unit entirely.
            }
        }
        return sb?.ToString() ?? s;
    }

    // ==================== Pivot property key canonicalization ====================
    //
    // R12-2 / R12-3: pivot property keys arrive from three sources
    // (CLI --prop, batch JSON, programmatic Dictionary) with varying case
    // and legacy singular/plural spellings. Normalize them all through one
    // helper so every downstream lookup site sees the same canonical key.
    //
    // Canonical keys (matches the Get readback and the ParseFieldList sites):
    //   source, src, name, position, pos, rows, cols, filters, values,
    //   aggregate, showdataas, topn, style, sort, grandtotals,
    //   rowgrandtotals, colgrandtotals
    //
    // Aliases that normalize TO a canonical key:
    //   row, rowfield, rowfields             → rows
    //   col, column, columns, colfield,
    //   colfields, columnfield, columnfields → cols
    //   filter, filterfield, filterfields    → filters
    //   value, valuefield, valuefields       → values
    //   columngrandtotals                    → colgrandtotals
    //
    // CONSISTENCY(compatibility-aliases): matches CLAUDE.md rule that Add/Set
    // may accept legacy aliases so old scripts (e.g. Round 3's rowFields key)
    // keep round-tripping. Get continues to emit only the canonical form.
    private static readonly Dictionary<string, string> _pivotKeyAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            // rows aliases
            ["row"]          = "rows",
            ["rowfield"]     = "rows",
            ["rowfields"]    = "rows",
            // cols aliases
            ["col"]          = "cols",
            ["column"]       = "cols",
            ["columns"]      = "cols",
            ["colfield"]     = "cols",
            ["colfields"]    = "cols",
            ["columnfield"]  = "cols",
            ["columnfields"] = "cols",
            // filters aliases
            ["filter"]       = "filters",
            ["filterfield"]  = "filters",
            ["filterfields"] = "filters",
            // values aliases
            ["value"]        = "values",
            ["valuefield"]   = "values",
            ["valuefields"]  = "values",
            // grand totals
            ["columngrandtotals"] = "colgrandtotals",
            // <pivotTableStyleInfo> col/column spelling aliases: the
            // OOXML attribute names use "column" but we prefer "col" as
            // the canonical CLI key to match the existing `cols=` axis
            // key. Add-path warning suppression relies on this rewrite.
            ["showcolumnstripes"] = "showcolstripes",
            ["showcolumnheaders"] = "showcolheaders",
            // repeatItemLabels aliases
            ["repeatitemlabels"]  = "repeatlabels",
            ["repeatalllabels"]   = "repeatlabels",
            ["filldownlabels"]    = "repeatlabels",
            // blankRows aliases
            ["insertblankrow"]    = "blankrows",
            ["insertblankrows"]   = "blankrows",
            ["blankrow"]          = "blankrows",
            ["blankline"]         = "blankrows",
            ["blanklines"]        = "blankrows",
        };

    /// <summary>
    /// Map a pivot property key to its canonical form. Returns the lower-cased
    /// key if no alias applies. Used by both CreatePivotTable (Add) and
    /// SetPivotTableProperties (Set) so every downstream `properties["rows"]`
    /// lookup binds to user input written as `row` / `rowFields` / `ROWS`.
    /// </summary>
    private static string NormalizePivotPropKey(string key)
    {
        if (string.IsNullOrEmpty(key)) return key;
        var lower = key.ToLowerInvariant();
        return _pivotKeyAliases.TryGetValue(lower, out var canonical) ? canonical : lower;
    }

    /// <summary>
    /// Validate a user-supplied pivot table name and return the trimmed value.
    /// Throws ArgumentException for empty, whitespace-only, control-character,
    /// or over-255-character names. Does NOT check workbook-level uniqueness
    /// (that is the caller's responsibility).
    /// R16-2: extracted from CreatePivotTable so SetPivotTableProperties can
    /// reuse the same validation — previously Set accepted empty/whitespace
    /// names without any check.
    /// </summary>
    private static string ValidatePivotName(string name)
    {
        // Empty string is rejected — a blank name is always an error.
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("pivot name must not be empty");
        var trimmed = name.Trim();
        // Whitespace-only names are rejected — R8-4.
        if (trimmed.Length == 0)
            throw new ArgumentException("pivot name must not be whitespace-only");
        // ASCII control characters are rejected — R8-5.
        foreach (var ch in trimmed)
        {
            if (ch < 0x20 || ch == 0x7F)
                throw new ArgumentException("pivot name contains invalid control characters");
        }
        // 255-character limit — R11-4.
        if (trimmed.Length > 255)
            throw new ArgumentException("pivot name exceeds 255-character limit");
        return trimmed;
    }

    /// <summary>
    /// Canonical key set recognized by the pivot Add / Set pipeline. Any
    /// property whose NORMALIZED key is not in this set is reported as
    /// UNSUPPORTED (Add: stderr warning; Set: returned unsupported list).
    /// Must stay in sync with the switch in SetPivotTableProperties and
    /// every properties lookup in CreatePivotTable.
    /// </summary>
    private static readonly HashSet<string> _knownPivotKeys =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "source", "src", "name", "position", "pos", "style",
            "rows", "cols", "filters", "values",
            "aggregate", "showdataas", "topn",
            "sort", "layout", "repeatlabels", "blankrows", "grandtotalcaption",
            "grandtotals", "rowgrandtotals", "colgrandtotals",
            "subtotals", "defaultsubtotal",
            // <pivotTableStyleInfo> bool toggles (see ApplyPivotStyleInfoProps).
            // Canonical keys only; col/column aliases are handled by the switch
            // in SetPivotTableProperties and the helper's case labels.
            "showrowstripes", "showcolstripes",
            "showrowheaders", "showcolheaders",
            "showlastcolumn",
        };

    /// <summary>
    /// Return the subset of the caller's pivot-property keys that are not
    /// known to the pipeline after alias normalization. Used by Add to
    /// emit an UNSUPPORTED stderr warning (R12-1) and shared by Set to
    /// merge into its existing unsupported return list. Keys are echoed
    /// in their ORIGINAL spelling (Unicode, case) so the user sees exactly
    /// what they typed — matches the 'unsupported echoes caller key' rule
    /// followed by the Set default case.
    /// </summary>
    private static List<string> CollectUnknownPivotKeys(Dictionary<string, string> properties)
    {
        var unknown = new List<string>();
        if (properties == null) return unknown;
        foreach (var key in properties.Keys)
        {
            if (string.IsNullOrEmpty(key)) continue;
            var canonical = NormalizePivotPropKey(key);
            if (!_knownPivotKeys.Contains(canonical))
                unknown.Add(key);
        }
        return unknown;
    }

    /// <summary>
    /// Emit an UNSUPPORTED props warning to stderr for the Add pivot path.
    /// Set already surfaces unknown keys through its return list; Add has
    /// no such channel, so we write directly. Format mirrors
    /// CommandBuilder.FormatUnsupported so JSON envelope parsing (see
    /// OutputFormatter.cs line 273) picks up the same prefix.
    /// </summary>
    private static void WarnUnknownPivotProperties(List<string> unknownKeys)
    {
        if (unknownKeys == null || unknownKeys.Count == 0) return;
        Console.Error.WriteLine(
            $"UNSUPPORTED props: {string.Join(", ", unknownKeys)}. "
            + "Use 'officecli help excel-set' to see available pivot properties.");
    }

    /// <summary>
    /// Normalize a user-supplied pivot properties dict into a new dict whose
    /// alias keys are rewritten to their canonical form. Keys that are
    /// already canonical and keys that don't match any known alias are
    /// preserved VERBATIM so the downstream unsupported-list reports the
    /// original spelling (matches the CLI contract that Set return values
    /// echo the caller's key). Collisions between an alias and an already-
    /// present canonical key are resolved first-seen-wins.
    /// </summary>
    private static Dictionary<string, string> NormalizePivotProperties(
        Dictionary<string, string> properties)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (properties == null) return result;
        foreach (var (rawKey, value) in properties)
        {
            // Only rewrite keys that the alias table knows about; everything
            // else (canonical keys, typos, non-ASCII) passes through with
            // the original spelling so error messages can echo it.
            var lower = rawKey?.ToLowerInvariant() ?? string.Empty;
            var outKey = _pivotKeyAliases.TryGetValue(lower, out var canonical)
                ? canonical
                : rawKey!;
            if (!result.ContainsKey(outKey))
                result[outKey] = value;
        }
        return result;
    }

    // ==================== Axis sort options ====================
    //
    // Axis labels on every level are sorted through a single comparer that
    // CreatePivotTable / SetPivotTableProperties publishes into _axisSortMode
    // for the duration of the operation. Every sort site below reads
    // ActiveAxisComparer / ActiveAxisDescending rather than hard-coding
    // StringComparer.Ordinal.
    //
    // Why ThreadStatic instead of a parameter: the sort opts have to reach
    // ~15 deeply-nested call sites (cache builders, pivotField items writers,
    // per-level index maps, 5 specialized renderers). Threading a parameter
    // through all of them would balloon 15+ signatures with pass-through
    // boilerplate. The CLI is single-threaded per pivot operation, so
    // ThreadStatic is safe and dramatically less invasive.
    //
    // Supported modes:
    //   "asc"         — StringComparer.Ordinal ascending (DEFAULT, preserves
    //                   byte-level regression baselines)
    //   "desc"        — StringComparer.Ordinal descending
    //   "locale"      — zh-CN culture ascending (pinyin). Hard-coded to
    //                   zh-CN rather than StringComparer.CurrentCulture:
    //                   on non-Chinese process locales (e.g. en-US on CI or
    //                   most developer machines) CurrentCulture silently
    //                   degrades to Ordinal for CJK strings, making locale
    //                   indistinguishable from asc. Pinyin is the primary
    //                   use case this mode exists for; honoring it regardless
    //                   of process locale is worth the lost generality.
    //   "locale-desc" — zh-CN culture descending
    [ThreadStatic] private static string? _axisSortMode;

    private static readonly IComparer<string> ZhCnComparer =
        StringComparer.Create(System.Globalization.CultureInfo.GetCultureInfo("zh-CN"), ignoreCase: false);

    private static IComparer<string> ActiveAxisComparer => _axisSortMode switch
    {
        "locale" or "locale-desc" => ZhCnComparer,
        _ => StringComparer.Ordinal
    };

    private static bool ActiveAxisDescending => _axisSortMode switch
    {
        "desc" or "locale-desc" => true,
        _ => false
    };

    /// <summary>
    /// Set axis sort mode from the pivot properties and return a token that
    /// restores the previous value on Dispose. Usage:
    ///   using (PushAxisSortMode(properties)) { ... build pivot ... }
    /// </summary>
    private static readonly HashSet<string> _validSortModes = new(StringComparer.OrdinalIgnoreCase)
    {
        "asc", "desc", "locale", "locale-desc"
    };

    private static IDisposable PushAxisSortMode(Dictionary<string, string> properties)
    {
        var prev = _axisSortMode;
        if (properties.TryGetValue("sort", out var mode) && !string.IsNullOrWhiteSpace(mode))
        {
            var normalized = mode.Trim().ToLowerInvariant();
            // CONSISTENCY(strict-enums): unknown sort tokens are rejected
            // up front. Empty / whitespace fall through to the default
            // (no-op) so users can clear the sort by passing an empty
            // value without seeing an error.
            if (!_validSortModes.Contains(normalized))
                throw new ArgumentException(
                    $"invalid sort: '{mode}'. Valid: asc, desc, locale, locale-desc");
            _axisSortMode = normalized;
        }
        return new SortModeScope(prev);
    }

    private sealed class SortModeScope : IDisposable
    {
        private readonly string? _prev;
        public SortModeScope(string? prev) { _prev = prev; }
        public void Dispose() { _axisSortMode = _prev; }
    }

    // ==================== Grand totals options ====================
    //
    // CONSISTENCY(thread-static-pivot-opts): reuses the same ThreadStatic
    // pattern as _axisSortMode above. Grand totals need to reach the same
    // ~15 nested sites (item builders, geometry, all 6 renderers, definition
    // builder), and threading parameters would explode signature churn.
    //
    // OOXML semantics (ECMA-376 § 18.10.1.73 on pivotTableDefinition), EMPIRICALLY
    // VERIFIED against an Excel-authored pivot the user created via
    // "Grand Totals → On for Rows Only" in the UI (test-samples/grand_totals_demo_Fix.xlsx):
    //   rowGrandTotals  — BOTTOM grand total ROW (one row at the bottom of the
    //                     pivot containing the per-col grand totals). Excel UI's
    //                     "On for Rows Only" enables this and writes colGrandTotals=0.
    //   colGrandTotals  — RIGHTMOST grand total COLUMN (one column at the right
    //                     of the pivot containing the per-row grand totals). Excel UI's
    //                     "On for Columns Only" enables this and writes rowGrandTotals=0.
    //
    // ⚠️  WARNING — HISTORICAL BUG: the initial implementation of this feature had
    // the mapping BACKWARDS (assumed rowGrandTotals = right column). The ThreadStatic
    // names below are kept stable to minimize churn, but their meaning was REDEFINED
    // during bug fix commit: `_rowGrandTotals` is the CLI-level flag whose true/false
    // maps to "render right column yes/no" (= OOXML colGrandTotals), and
    // `_colGrandTotals` maps to "render bottom row yes/no" (= OOXML rowGrandTotals).
    // The renderer / geometry / item builders use `ActiveRowGrandTotals` /
    // `ActiveColGrandTotals` to mean "right col visible" / "bottom row visible"
    // respectively. The attribute writer / reader / parser swap the names when
    // talking to OOXML so the final XML and visual match Excel UI.
    //
    // Both default to true. We only write the attribute when the user
    // explicitly opts out (matches how real Excel + LibreOffice serialize).
    [ThreadStatic] private static bool? _rowGrandTotals;
    [ThreadStatic] private static bool? _colGrandTotals;

    // ActiveRowGrandTotals: "render the right grand-total column" (= OOXML colGrandTotals)
    // ActiveColGrandTotals: "render the bottom grand-total row"   (= OOXML rowGrandTotals)
    private static bool ActiveRowGrandTotals => _rowGrandTotals ?? true;
    private static bool ActiveColGrandTotals => _colGrandTotals ?? true;

    /// <summary>
    /// Parse grand-totals properties into the thread-static scope. Supports:
    ///   grandTotals=both|none|rows|cols|on|off|true|false
    ///   rowGrandTotals=true|false   (overrides grandTotals for the row-grand axis)
    ///   colGrandTotals=true|false   (overrides grandTotals for the col-grand axis)
    /// Returns a scope that restores the previous values on Dispose.
    /// </summary>
    private static IDisposable PushGrandTotalsOptions(Dictionary<string, string> properties)
    {
        var prevRow = _rowGrandTotals;
        var prevCol = _colGrandTotals;

        // Master 'grandTotals' key (friendly), matching Excel UI semantics:
        //   'rows' = Excel's "On for Rows Only" = BOTTOM row visible, right col hidden
        //   'cols' = Excel's "On for Columns Only" = RIGHT col visible, bottom row hidden
        // Internally: _rowGrandTotals = "render right col", _colGrandTotals = "render bottom row"
        // (see comment at the ThreadStatic declaration above).
        if (properties.TryGetValue("grandTotals", out var gt)
            || properties.TryGetValue("grandtotals", out gt))
        {
            switch ((gt ?? "").Trim().ToLowerInvariant())
            {
                case "both": case "on": case "true": case "1": case "yes":
                    _rowGrandTotals = true; _colGrandTotals = true; break;
                case "none": case "off": case "false": case "0": case "no":
                    _rowGrandTotals = false; _colGrandTotals = false; break;
                case "rows": case "row":
                    // "On for Rows Only" = only bottom row, no right col.
                    _rowGrandTotals = false; _colGrandTotals = true; break;
                case "cols": case "col": case "columns":
                    // "On for Columns Only" = only right col, no bottom row.
                    _rowGrandTotals = true; _colGrandTotals = false; break;
            }
        }

        // Fine-grained bool keys mirror OOXML attribute names (ECMA-376):
        //   rowGrandTotals=... → bottom row toggle (internal: _colGrandTotals)
        //   colGrandTotals=... → right col toggle  (internal: _rowGrandTotals)
        // Parsed AFTER the master key so they override it when both are supplied.
        if (TryParseBoolProp(properties, "rowGrandTotals", out var rgt))
            _colGrandTotals = rgt;
        if (TryParseBoolProp(properties, "colGrandTotals", out var cgt)
            || TryParseBoolProp(properties, "columnGrandTotals", out cgt))
            _rowGrandTotals = cgt;

        return new GrandTotalsScope(prevRow, prevCol);
    }

    private static bool TryParseBoolProp(Dictionary<string, string> properties, string key, out bool value)
    {
        value = false;
        if (!properties.TryGetValue(key, out var raw)
            && !properties.TryGetValue(key.ToLowerInvariant(), out raw))
            return false;
        switch ((raw ?? "").Trim().ToLowerInvariant())
        {
            case "true": case "1": case "yes": case "on": value = true; return true;
            case "false": case "0": case "no": case "off": value = false; return true;
            default: return false;
        }
    }

    private sealed class GrandTotalsScope : IDisposable
    {
        private readonly bool? _prevRow;
        private readonly bool? _prevCol;
        public GrandTotalsScope(bool? prevRow, bool? prevCol) { _prevRow = prevRow; _prevCol = prevCol; }
        public void Dispose() { _rowGrandTotals = _prevRow; _colGrandTotals = _prevCol; }
    }

    // ==================== Subtotals options ====================
    //
    // CONSISTENCY(thread-static-pivot-opts): same ThreadStatic precedent as
    // sort + grand totals. Subtotals (the outer-level group subtotal rows
    // and columns that appear between groups in 2+ row/col-field pivots)
    // need to reach item builders, geometry, and every multi-dim renderer.
    //
    // OOXML semantics (ECMA-376 § 18.10.1.69 on pivotField):
    //   defaultSubtotal (default true) — whether this pivot field's axis
    //                    emits an outer-level subtotal sentinel
    //                    (<item t="default"/> in pivotField.items).
    //
    // v1b scope: only on/off. subtotalTop (position = top vs bottom of
    // group) is deferred — our renderers always emit subtotals at the top
    // of each group, and switching position would require reordering the
    // sheetData write loop. Tracked as v1c.
    [ThreadStatic] private static bool? _defaultSubtotal;

    private static bool ActiveDefaultSubtotal => _defaultSubtotal ?? true;

    /// <summary>
    /// Parse subtotals properties into the thread-static scope. Supports:
    ///   subtotals=on|off|true|false|show|hide|yes|no|1|0
    ///   defaultSubtotal=true|false   (OOXML-level alias)
    /// Returns a scope that restores the previous value on Dispose.
    /// </summary>
    private static IDisposable PushSubtotalsOptions(Dictionary<string, string> properties)
    {
        var prev = _defaultSubtotal;

        if (properties.TryGetValue("subtotals", out var s)
            || properties.TryGetValue("Subtotals", out s))
        {
            switch ((s ?? "").Trim().ToLowerInvariant())
            {
                case "on": case "true": case "1": case "yes": case "show":
                    _defaultSubtotal = true; break;
                case "off": case "false": case "0": case "no": case "hide": case "none":
                    _defaultSubtotal = false; break;
            }
        }

        if (TryParseBoolProp(properties, "defaultSubtotal", out var ds))
            _defaultSubtotal = ds;

        return new SubtotalsScope(prev);
    }

    private sealed class SubtotalsScope : IDisposable
    {
        private readonly bool? _prev;
        public SubtotalsScope(bool? prev) { _prev = prev; }
        public void Dispose() { _defaultSubtotal = _prev; }
    }

    // ==================== Layout mode options ====================
    //
    // CONSISTENCY(thread-static-pivot-opts): same ThreadStatic precedent as
    // sort + grand totals + subtotals. Layout mode (compact/outline/tabular)
    // affects geometry (rowLabelCols), definition attributes, PivotField
    // attributes, and renderer column placement. Threading a parameter
    // through all 15+ call sites would be excessively invasive.
    //
    // Supported modes:
    //   "compact"  — (DEFAULT) all row fields share one column with indentation
    //   "outline"  — each row field gets its own column, labels on same row as data
    //   "tabular"  — each row field gets its own column, labels on separate row from data
    [ThreadStatic] private static string? _layoutMode;

    private static string ActiveLayoutMode => _layoutMode ?? "compact";

    /// <summary>
    /// Parse layout property into the thread-static scope. Supports:
    ///   layout=compact|outline|tabular
    /// Returns a scope that restores the previous value on Dispose.
    /// </summary>
    private static readonly HashSet<string> _validLayoutModes = new(StringComparer.OrdinalIgnoreCase)
    {
        "compact", "outline", "tabular"
    };

    private static IDisposable PushLayoutMode(Dictionary<string, string> properties)
    {
        var prev = _layoutMode;
        if (properties.TryGetValue("layout", out var mode) && !string.IsNullOrWhiteSpace(mode))
        {
            var normalized = mode.Trim().ToLowerInvariant();
            if (!_validLayoutModes.Contains(normalized))
                throw new ArgumentException(
                    $"invalid layout: '{mode}'. Valid: compact, outline, tabular");
            _layoutMode = normalized;
        }
        return new LayoutModeScope(prev);
    }

    private sealed class LayoutModeScope : IDisposable
    {
        private readonly string? _prev;
        public LayoutModeScope(string? prev) { _prev = prev; }
        public void Dispose() { _layoutMode = _prev; }
    }

    // CONSISTENCY(thread-static-pivot-opts): repeatItemLabels — "Repeat All
    // Item Labels" in Excel's Report Layout menu. When true, outer row axis
    // labels are repeated on every leaf row instead of appearing only once
    // at the top of each group. OOXML: fillDownLabelsDefault on x14:pivotTableDefinition.
    [ThreadStatic] private static bool? _repeatItemLabels;

    private static bool ActiveRepeatItemLabels => _repeatItemLabels ?? false;

    private static IDisposable PushRepeatItemLabels(Dictionary<string, string> properties)
    {
        var prev = _repeatItemLabels;
        if (properties.TryGetValue("repeatlabels", out var val) && !string.IsNullOrWhiteSpace(val))
            _repeatItemLabels = ParseHelpers.IsTruthy(val);
        return new RepeatItemLabelsScope(prev);
    }

    private sealed class RepeatItemLabelsScope : IDisposable
    {
        private readonly bool? _prev;
        public RepeatItemLabelsScope(bool? prev) { _prev = prev; }
        public void Dispose() { _repeatItemLabels = _prev; }
    }

    // CONSISTENCY(thread-static-pivot-opts): insertBlankRow — "Insert Blank
    // Line After Each Item" in Excel's Report Layout menu. When true, an
    // empty row is inserted after each outer group (after subtotal in tabular,
    // after last leaf in compact/outline). OOXML: insertBlankRow on pivotField.
    [ThreadStatic] private static bool? _insertBlankRow;

    private static bool ActiveInsertBlankRow => _insertBlankRow ?? false;

    private static IDisposable PushInsertBlankRow(Dictionary<string, string> properties)
    {
        var prev = _insertBlankRow;
        if (properties.TryGetValue("blankrows", out var val) && !string.IsNullOrWhiteSpace(val))
            _insertBlankRow = ParseHelpers.IsTruthy(val);
        return new InsertBlankRowScope(prev);
    }

    private sealed class InsertBlankRowScope : IDisposable
    {
        private readonly bool? _prev;
        public InsertBlankRowScope(bool? prev) { _prev = prev; }
        public void Dispose() { _insertBlankRow = _prev; }
    }

    // CONSISTENCY(thread-static-pivot-opts): grandTotalCaption — user-specified
    // label for the grand total row/column. Defaults to "Grand Total".
    [ThreadStatic] private static string? _grandTotalCaption;

    private static string ActiveGrandTotalCaption => _grandTotalCaption ?? "Grand Total";

    private static IDisposable PushGrandTotalCaption(Dictionary<string, string> properties)
    {
        var prev = _grandTotalCaption;
        if (properties.TryGetValue("grandtotalcaption", out var val) && !string.IsNullOrWhiteSpace(val))
            _grandTotalCaption = val.Trim();
        return new GrandTotalCaptionScope(prev);
    }

    private sealed class GrandTotalCaptionScope : IDisposable
    {
        private readonly string? _prev;
        public GrandTotalCaptionScope(string? prev) { _prev = prev; }
        public void Dispose() { _grandTotalCaption = _prev; }
    }

    /// <summary>
    /// Apply axis ordering (ascending/descending) to an OrderBy clause using
    /// the currently-active sort mode. All axis sort sites use this helper.
    /// </summary>
    private static IOrderedEnumerable<T> OrderByAxis<T>(this IEnumerable<T> source, Func<T, string> keySelector)
    {
        return ActiveAxisDescending
            ? source.OrderByDescending(keySelector, ActiveAxisComparer)
            : source.OrderBy(keySelector, ActiveAxisComparer);
    }

    // ==================== Top-N filter ====================
    //
    // Applies a Top-N filter to the source data BEFORE the cache / renderer
    // see it. Semantics (V1):
    //   * Ranks values of the OUTERMOST row field by the FIRST value field's
    //     aggregate (using that value field's func: sum/avg/count/...).
    //   * Keeps the top N keys by that aggregate (descending — "top = largest").
    //   * Drops source rows whose outer-row-field value is not in the kept set.
    //
    // Why filter source rows instead of emitting <top10>/<autoShow> OOXML:
    // the renderer writes pivot cells directly into sheetData as a static
    // snapshot. There is no Excel-side recompute step for an OOXML-level
    // filter to honour, so filtering the source is what keeps cache,
    // rendered cells, and grand totals in lock-step.
    //
    // Interaction with `sort`: independent. `topN` picks the set by VALUE
    // (largest aggregates), `sort` arranges the kept set by LABEL
    // (asc/desc/locale). Both compose cleanly.
    //
    // Known limitations (tracked for v2 expansion):
    //   * Outermost row field only — col-axis and inner-level Top-N are not
    //     supported.
    //   * Always "top" (largest). "bottom" / worst-N is not supported.
    //   * Ranks by the FIRST value field when multiple values exist.
    //   * Set operation does NOT re-apply Top-N (cache is already built at
    //     that point). Users must remove + re-add the pivot to re-filter.
    //
    // No-op cases (silently skipped — mirrors how `sort` handles degenerate
    // inputs):
    //   * topN <= 0
    //   * rows empty (nothing to rank on)
    //   * values empty (nothing to rank by)
    //   * topN >= distinct outer keys (keeps everything)
    private static void ApplyTopNFilter(
        List<string[]> columnData,
        List<int> rowFields,
        List<(int idx, string func, string showAs, string name)> valueFields,
        int topN)
    {
        if (topN <= 0 || rowFields.Count == 0 || valueFields.Count == 0 || columnData.Count == 0)
            return;

        var outerFieldIdx = rowFields[0];
        var valueFieldIdx = valueFields[0].idx;
        var valueFunc = valueFields[0].func;
        if (outerFieldIdx < 0 || outerFieldIdx >= columnData.Count) return;
        if (valueFieldIdx < 0 || valueFieldIdx >= columnData.Count) return;

        var outerCol = columnData[outerFieldIdx];
        var valueCol = columnData[valueFieldIdx];
        var rowCount = outerCol.Length;
        if (rowCount == 0) return;

        // Aggregate per outer-key using the first value field's function.
        var buckets = new Dictionary<string, List<double>>(StringComparer.Ordinal);
        for (int r = 0; r < rowCount; r++)
        {
            var key = outerCol[r];
            if (string.IsNullOrEmpty(key)) continue;
            if (r >= valueCol.Length) continue;
            if (!double.TryParse(valueCol[r], System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var v))
                continue;
            if (!buckets.TryGetValue(key, out var list))
            {
                list = new List<double>();
                buckets[key] = list;
            }
            list.Add(v);
        }

        if (buckets.Count <= topN) return; // keeps everything — no-op

        // Rank keys by aggregate descending; stable tie-break by ordinal label
        // so the kept set is deterministic across runs.
        var kept = buckets
            .Select(kv => (key: kv.Key, agg: ReducePivotValues(kv.Value, valueFunc)))
            .OrderByDescending(t => t.agg)
            .ThenBy(t => t.key, StringComparer.Ordinal)
            .Take(topN)
            .Select(t => t.key)
            .ToHashSet(StringComparer.Ordinal);

        // Build keep-mask over source rows.
        var keep = new bool[rowCount];
        int keepCount = 0;
        for (int r = 0; r < rowCount; r++)
        {
            var k = outerCol[r];
            if (!string.IsNullOrEmpty(k) && kept.Contains(k))
            {
                keep[r] = true;
                keepCount++;
            }
        }

        if (keepCount == rowCount) return; // nothing to drop

        // Apply mask to every column in place.
        for (int c = 0; c < columnData.Count; c++)
        {
            var src = columnData[c];
            var dst = new string[keepCount];
            int w = 0;
            for (int r = 0; r < rowCount && r < src.Length; r++)
            {
                if (keep[r]) dst[w++] = src[r];
            }
            columnData[c] = dst;
        }
    }

    /// <summary>
    /// Create a pivot table on the target worksheet.
    /// </summary>
    /// <param name="workbookPart">The workbook part</param>
    /// <param name="targetSheet">Worksheet where the pivot table will be placed</param>
    /// <param name="sourceSheet">Worksheet containing the source data</param>
    /// <param name="sourceSheetName">Name of the source worksheet</param>
    /// <param name="sourceRef">Source data range (e.g. "A1:D100")</param>
    /// <param name="position">Top-left cell for the pivot table (e.g. "F1")</param>
    /// <param name="properties">Configuration: rows, cols, values, filters, style, name</param>
    /// <returns>The 1-based index of the created pivot table</returns>
    internal static int CreatePivotTable(
        WorkbookPart workbookPart,
        WorksheetPart targetSheet,
        WorksheetPart sourceSheet,
        string sourceSheetName,
        string sourceRef,
        string position,
        Dictionary<string, string> properties)
    {
        // R12-1: detect unknown pivot property keys (including non-ASCII
        // like '源'/'行名') BEFORE normalization so the warning echoes the
        // original spelling. Previously these keys were silently dropped
        // and users saw an empty pivot with no diagnostic.
        WarnUnknownPivotProperties(CollectUnknownPivotKeys(properties));

        // R12-2 / R12-3: normalize alias keys (row→rows, rowFields→rows,
        // columngrandtotals→colgrandtotals, etc.) so every downstream
        // lookup below reads from the canonical dict. `row=Cat` then
        // binds to the same code path as `rows=Cat`.
        properties = NormalizePivotProperties(properties);

        // Publish the axis sort mode (asc/desc/locale/locale-desc) so every
        // sort site below — cache builder, pivotField items writer, per-level
        // index maps, specialized renderers — reads the same comparer.
        using var _sortScope = PushAxisSortMode(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern — grand totals
        // options reach item builders, geometry, and every renderer via
        // ActiveRowGrandTotals/ActiveColGrandTotals.
        using var _gtScope = PushGrandTotalsOptions(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for subtotals.
        using var _subScope = PushSubtotalsOptions(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for layout mode.
        using var _layoutScope = PushLayoutMode(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for repeatItemLabels.
        using var _repeatScope = PushRepeatItemLabels(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for insertBlankRow.
        using var _blankRowScope = PushInsertBlankRow(properties);
        // CONSISTENCY(thread-static-pivot-opts): same pattern for grandTotalCaption.
        using var _captionScope = PushGrandTotalCaption(properties);

        // 1. Read source data to build cache
        var (headers, columnData, columnStyleIds) = ReadSourceData(sourceSheet, sourceRef);
        if (headers.Length == 0)
            throw new ArgumentException("Source range has no data");
        // CONSISTENCY(empty-pivot-source): a header row with zero data rows
        // (e.g. A1:D1) silently produces an empty pivot whose cache has no
        // records — Excel opens it but renders nothing. Reject it with the
        // same family of ArgumentException as the no-headers case so callers
        // get a single, predictable error path. Bt#8 / fuzzer baseline.
        if (columnData.Count == 0 || columnData[0].Length == 0)
            throw new ArgumentException("Source range has no data rows");

        // 1b. Date auto-grouping preprocessing. Scans rows/cols/filters props
        // for `fieldName:grouping` syntax (e.g. `rows='日期:month,城市'`) and
        // creates a new virtual column per grouped field containing the
        // bucketed labels. The raw field spec is rewritten to reference the
        // new virtual column so ParseFieldList below sees a clean name.
        //
        // Supported groupings:
        //   :year    → "2024"
        //   :quarter → "2024-Q1"
        //   :month   → "2024-01"
        //   :day     → "2024-01-05"
        //
        // Compose multiple groupings for hierarchical date layouts:
        // `rows='日期:year,日期:quarter'` → 2-level year-then-quarter.
        //
        // Returns a list of DateGroupSpec describing each derived field so
        // BuildCacheDefinition can emit the native <fieldGroup> + <rangePr> +
        // <groupItems> XML that Excel requires to accept the pivot as a
        // real date-grouped table (without it, Excel detects a "fieldGroup
        // shape mismatch" and refuses to render the inner hierarchy levels).
        List<DateGroupSpec> dateGroups;
        (headers, columnData, dateGroups) = ApplyDateGrouping(headers, columnData, properties);

        // 2. Parse field assignments from properties
        var rowFields = ParseFieldList(properties, "rows", headers);
        var colFields = ParseFieldList(properties, "cols", headers);
        var filterFields = ParseFieldList(properties, "filters", headers);
        var valueFields = ParseValueFields(properties, "values", headers);

        // CONSISTENCY(aggregate-override / showdataas): parity with Set —
        // the sibling `aggregate=` / `showdataas=` properties are positional
        // comma-lists applied to the parsed value-field list so users can
        // write `values=Sales showdataas=percent_of_row` and have it take
        // effect at Add time, not only when re-specified via Set. R8-1.
        {
            string[]? aggOverrideAdd = null;
            string[]? showOverrideAdd = null;
            if (properties.TryGetValue("aggregate", out var aggSpecAdd) && !string.IsNullOrEmpty(aggSpecAdd))
                aggOverrideAdd = aggSpecAdd.Split(',').Select(s => s.Trim().ToLowerInvariant()).ToArray();
            if (properties.TryGetValue("showdataas", out var showSpecAdd) && !string.IsNullOrEmpty(showSpecAdd))
                showOverrideAdd = showSpecAdd.Split(',').Select(s => s.Trim().ToLowerInvariant()).ToArray();
            if (aggOverrideAdd != null || showOverrideAdd != null)
            {
                for (int i = 0; i < valueFields.Count; i++)
                {
                    var (idx, func, showAs, name) = valueFields[i];
                    if (aggOverrideAdd != null && i < aggOverrideAdd.Length && !string.IsNullOrEmpty(aggOverrideAdd[i]))
                        func = aggOverrideAdd[i];
                    if (showOverrideAdd != null && i < showOverrideAdd.Length && !string.IsNullOrEmpty(showOverrideAdd[i]))
                    {
                        // Validate via ParseShowDataAs — throws on unknown/unsupported tokens,
                        // matching the Set path and CONSISTENCY(strict-enums).
                        ParseShowDataAs(showOverrideAdd[i]);
                        showAs = showOverrideAdd[i];
                    }
                    valueFields[i] = (idx, func, showAs, name);
                }
            }
        }

        // Auto-assign: if no values specified, use the first numeric column
        if (valueFields.Count == 0)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                if (!rowFields.Contains(i) && !colFields.Contains(i) && !filterFields.Contains(i)
                    && columnData[i].All(v => double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _)))
                {
                    valueFields.Add((i, "sum", "normal", $"Sum of {headers[i]}"));
                    break;
                }
            }
        }

        // 2b. Apply Top-N filter to the source rows (ranked by the first value
        // field's aggregate on the outermost row field). Runs BEFORE cache
        // build so the cache, rendered cells, and grand totals all reflect
        // the filtered subset. See ApplyTopNFilter for semantics & limits.
        if ((properties.TryGetValue("topN", out var topNStr)
             || properties.TryGetValue("topn", out topNStr))
            && int.TryParse(topNStr, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var topN))
        {
            ApplyTopNFilter(columnData, rowFields, valueFields, topN);
        }

        // 3. Generate unique cache ID
        uint cacheId = 0;
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Workbook is missing");
        var pivotCaches = workbook.GetFirstChild<PivotCaches>();
        if (pivotCaches != null)
            cacheId = pivotCaches.Elements<PivotCache>().Select(pc => pc.CacheId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;

        // 3b. Collect all existing pivot names in the workbook so we can
        // reject duplicates (user-supplied) or auto-increment past collisions
        // (default name). Excel auto-renames on open to avoid the clash, but
        // the file as written with a duplicate is confusing and breaks any
        // downstream consumer keying pivots by name. R6-1.
        var existingPivotNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var wsp in workbookPart.WorksheetParts)
        {
            foreach (var ptp in wsp.PivotTableParts)
            {
                var existingName = ptp.PivotTableDefinition?.Name?.Value;
                if (!string.IsNullOrEmpty(existingName))
                    existingPivotNames.Add(existingName);
            }
        }

        // 4. Create PivotTableCacheDefinitionPart at workbook level
        var cachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var cacheRelId = workbookPart.GetIdOfPart(cachePart);

        // Build cache definition + per-field shared-item index maps. The maps are
        // needed to write pivotCacheRecords below: each non-numeric field value is
        // referenced as <x v="N"/> where N is the value's position in sharedItems.
        //
        // Axis fields (row/col/filter) ALWAYS go through the string/indexed
        // path even if their values parse as numeric. Otherwise the pivotField
        // items list (which AppendFieldItems builds by index) and the cache
        // records (which would emit <n v="..."/>) disagree on what "index 0"
        // means, and Excel refuses to render the row/col hierarchy. Date
        // grouping's "year" bucket (values like "2024"/"2025") was the
        // triggering case — the fix is to mark axis fields here.
        var axisFieldSet = new HashSet<int>();
        foreach (var r in rowFields) axisFieldSet.Add(r);
        foreach (var c in colFields) axisFieldSet.Add(c);
        foreach (var f in filterFields) axisFieldSet.Add(f);
        var (cacheDef, fieldNumeric, fieldValueIndex) =
            BuildCacheDefinition(sourceSheetName, sourceRef, headers, columnData, axisFieldSet, dateGroups);
        cachePart.PivotCacheDefinition = cacheDef;
        cachePart.PivotCacheDefinition.Save();

        // 4b. Create PivotTableCacheRecordsPart and write one record per source row.
        // Without records, Excel rejects the file with "PivotTable report is invalid"
        // because saveData defaults to true. Writing real records also makes the file
        // self-contained for non-refreshing consumers (POI, third-party parsers).
        var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        // Derived date-group fields (databaseField="0") must be excluded from
        // pivotCacheRecords — Excel computes them from the base field's
        // <fieldGroup> definition on the fly. Pass their indices so the
        // record writer skips them.
        var derivedFieldSet = dateGroups.Count > 0
            ? new HashSet<int>(dateGroups.Select(g => g.DerivedFieldIdx))
            : null;
        recordsPart.PivotCacheRecords = BuildCacheRecords(columnData, fieldNumeric, fieldValueIndex, derivedFieldSet);
        recordsPart.PivotCacheRecords.Save();

        // The pivotCacheDefinition element MUST carry an r:id attribute pointing to the
        // records part — Excel uses it to find records, not the package _rels alone.
        // LibreOffice writes this in xepivotxml.cxx:280 (FSNS(XML_r, XML_id)). Without
        // this attribute the file looks structurally complete but Excel rejects it.
        cacheDef.Id = cachePart.GetIdOfPart(recordsPart);
        cachePart.PivotCacheDefinition.Save();

        // Register in workbook's PivotCaches
        if (pivotCaches == null)
        {
            pivotCaches = new PivotCaches();
            // OOXML schema requires pivotCaches AFTER calcPr/oleSize/
            // customWorkbookViews and BEFORE smartTagPr/fileRecoveryPr/extLst.
            // AppendChild puts it after fileRecoveryPr, violating schema order
            // and causing Excel to report "problem with some content".
            var insertBefore = workbook.GetFirstChild<WebPublishing>()
                ?? workbook.GetFirstChild<FileRecoveryProperties>()
                ?? (OpenXmlElement?)workbook.GetFirstChild<WebPublishObjects>();
            if (insertBefore != null)
                workbook.InsertBefore(pivotCaches, insertBefore);
            else
                workbook.AppendChild(pivotCaches);
        }
        pivotCaches.AppendChild(new PivotCache { CacheId = cacheId, Id = cacheRelId });
        workbook.Save();

        // 5. Create PivotTablePart at worksheet level
        var pivotPart = targetSheet.AddNewPart<PivotTablePart>();
        // Link pivot table to cache definition
        pivotPart.AddPart(cachePart);

        string pivotName;
        if (properties.TryGetValue("name", out var explicitName) && !string.IsNullOrEmpty(explicitName))
        {
            // R8-4 / R8-5 / R11-4 / R16-2: delegate all name validation to
            // ValidatePivotName so Add and Set share identical rules.
            explicitName = ValidatePivotName(explicitName);
            // R6-1: user-supplied name must be unique within the workbook.
            // Throw ArgumentException rather than silently allowing the
            // collision (Excel would auto-rename on open, but the on-disk
            // file would still carry two pivots with the same name).
            if (existingPivotNames.Contains(explicitName))
                throw new ArgumentException($"Pivot name '{explicitName}' already exists in workbook");
            pivotName = explicitName;
        }
        else
        {
            // R6-1: auto-generated default names must also avoid collisions
            // (two pivots on different sheets otherwise both pick
            // PivotTable{cacheId+1} with the same cacheId path).
            pivotName = $"PivotTable{cacheId + 1}";
            int bump = 1;
            while (existingPivotNames.Contains(pivotName))
            {
                bump++;
                pivotName = $"PivotTable{cacheId + bump}";
            }
        }
        var style = properties.GetValueOrDefault("style", "PivotStyleLight16");

        // Resolve per-column numFmtId from the source StyleIndex so we can stamp
        // it onto DataField elements below. Excel uses DataField.NumberFormatId
        // as the PRIMARY display driver for pivot values — the cell-level
        // StyleIndex alone is not enough; without this, Excel renders pivot
        // values as plain General-format numbers even though the rendered cells
        // carry the correct style.
        var columnNumFmtIds = ResolveColumnNumFmtIds(workbookPart, columnStyleIds);

        // Page filters occupy rows ABOVE the pivot body. Ensure position leaves
        // enough headroom for filterCount filter rows + 1 blank separator row.
        if (filterFields.Count > 0)
        {
            var (posCol, posRow) = ParseCellRef(position);
            int minBodyRow = filterFields.Count + 2; // 1-based
            if (posRow < minBodyRow)
                position = $"{posCol}{minBodyRow}";
        }

        var pivotDef = BuildPivotTableDefinition(
            pivotName, cacheId, position, headers, columnData,
            rowFields, colFields, filterFields, valueFields, style, columnNumFmtIds, dateGroups);
        // Overlay user-supplied <pivotTableStyleInfo> bool attributes
        // (showRowStripes, showColStripes, showRowHeaders, showColHeaders,
        // showLastColumn) onto the style info element BuildPivotTableDefinition
        // just created with defaults. Shared helper with the Set path so
        // Add and Set accept the same vocabulary / validation.
        ApplyPivotStyleInfoProps(EnsurePivotTableStyle(pivotDef), properties);
        pivotPart.PivotTableDefinition = pivotDef;
        pivotPart.PivotTableDefinition.Save();

        // 6. RENDER the pivot output into the target sheet's <sheetData>.
        //
        // This is the critical step that distinguishes a "valid pivot file Excel
        // accepts" from a "pivot file Excel actually displays". Excel does NOT
        // recompute pivots from cache on open — it reads the rendered cells
        // directly from sheetData, exactly like any other range. We verified this
        // by inspecting an Excel-authored sample (excel_authored.xlsx → sheet2.xml):
        // every aggregated cell is a literal <c><v>200</v></c> element.
        //
        // Without this step the pivot opens as an empty drop-down skeleton — the
        // structure is valid but there is nothing to display. POI / Open XML SDK
        // suffer from exactly the same limitation; this is the lift that turns
        // officecli into a real pivot writer rather than a definition-only one.
        //
        // For unsupported configurations (multiple row/col fields, multiple data
        // fields, page filters), the renderer falls back to writing nothing, which
        // gives Excel an empty sheetData and the same skeleton-only behavior.
        // Those configs are tracked as a v2 expansion.
        RenderPivotIntoSheet(
            targetSheet, position, headers, columnData,
            rowFields, colFields, valueFields, filterFields, columnStyleIds);

        // After rendering, collapse any duplicate <row r="N"> elements the
        // renderer may have appended if this sheet already had pivot-rendered
        // rows (second pivot in same sheet → shared row indices). OOXML
        // requires unique row elements per index; Excel rejects the file with
        // "problem with some content" otherwise.
        var targetSheetData = targetSheet.Worksheet?.GetFirstChild<SheetData>();
        if (targetSheetData != null)
            DedupeSheetDataRows(targetSheetData);

        // Return 1-based index
        return targetSheet.PivotTableParts.ToList().IndexOf(pivotPart) + 1;
    }

    // ==================== Axis Tree (general N-level row/col abstraction) ====================
    //
    // For N≥3 row or col fields the existing specialized renderers (1×1, 2×1,
    // 1×2, 2×2 with K data variants) cannot be extended without an N² explosion
    // in case count. The AxisTree abstraction below replaces them with a single
    // recursive tree representation:
    //
    //   - The root has one child per unique value of the FIRST (outermost) field
    //   - Each level-L node has one child per unique value of the (L+1)-th field
    //     that appears in the source data PAIRED WITH the parent's path
    //   - Leaves are at depth N (i.e. path length = N field values)
    //
    // Example for rows=[地区, 城市, 区]:
    //   root
    //   ├── 华东
    //   │   ├── 上海
    //   │   │   ├── 浦东
    //   │   │   └── 徐汇
    //   │   └── 杭州
    //   │       └── 西湖
    //   └── 华北
    //       └── 北京
    //           ├── 朝阳
    //           └── 海淀
    //
    // Walk order produces (in display sequence): outer subtotals at internal
    // nodes + leaf rows at leaves + grand total at the very end. For 2D pivots
    // both row and col axes use independent AxisTrees and the renderer walks
    // them in lockstep.
    //
    // This abstraction is currently used ONLY for N≥3 cases via the dispatch in
    // RenderPivotIntoSheet. The 8 existing N≤2 cases continue to use their
    // specialized renderers (regression-tested via test-samples/pivot_baselines).

    /// <summary>
    /// One node in the axis tree. Represents either an internal node (subtotal
    /// row/col) or a leaf node (specific data row/col). Children are sorted in
    /// ordinal display order to keep rowItems/colItems indices consistent with
    /// the corresponding pivotField items list.
    /// </summary>
    private sealed class AxisNode
    {
        /// <summary>The label for this node (e.g. "华东"). Empty string for the root.</summary>
        public string Label { get; }
        /// <summary>0 = root, 1 = outermost field, 2 = next inner, ..., N = leaf level.</summary>
        public int Depth { get; }
        /// <summary>Path from root: [outerVal, ..., this.Label]. Length == Depth.</summary>
        public string[] Path { get; }
        /// <summary>Child nodes in ordinal display order. Empty for leaves.</summary>
        public List<AxisNode> Children { get; } = new();

        public AxisNode(string label, int depth, string[] path)
        {
            Label = label;
            Depth = depth;
            Path = path;
        }

        public bool IsLeaf => Children.Count == 0;
    }

    /// <summary>
    /// Build an AxisTree from columnData given the field indices for an axis.
    /// Only paths that actually appear in the source data are included — Excel
    /// does not enumerate empty cartesian intersections at any level.
    /// </summary>
    private static AxisNode BuildAxisTree(List<int> fieldIndices, List<string[]> columnData)
    {
        var root = new AxisNode(string.Empty, 0, Array.Empty<string>());
        if (fieldIndices.Count == 0 || columnData.Count == 0)
            return root;

        var rowCount = columnData[fieldIndices[0]].Length;
        // For each source row, walk down the tree, creating child nodes as needed.
        for (int r = 0; r < rowCount; r++)
        {
            var current = root;
            var validPath = true;
            var path = new string[fieldIndices.Count];

            for (int level = 0; level < fieldIndices.Count; level++)
            {
                var fieldIdx = fieldIndices[level];
                if (fieldIdx < 0 || fieldIdx >= columnData.Count) { validPath = false; break; }
                var values = columnData[fieldIdx];
                if (r >= values.Length) { validPath = false; break; }
                var v = values[r];
                if (string.IsNullOrEmpty(v)) { validPath = false; break; }
                path[level] = v;

                // Find or create child for this value at this level.
                var child = current.Children.FirstOrDefault(c => c.Label == v);
                if (child == null)
                {
                    var childPath = new string[level + 1];
                    Array.Copy(path, childPath, level + 1);
                    child = new AxisNode(v, level + 1, childPath);
                    current.Children.Add(child);
                }
                current = child;
            }

            // Drop the row entirely if any field had an empty value — matches the
            // "skip rows with missing values" semantics of the specialized renderers.
            _ = validPath;
        }

        // Sort children at every level using the same StringComparer.Ordinal that
        // BuildOuterInnerGroups and AppendFieldItems use, so the rowItems indices
        // line up with the pivotField items list.
        SortAxisTreeRecursive(root);
        return root;
    }

    private static void SortAxisTreeRecursive(AxisNode node)
    {
        var cmp = ActiveAxisComparer;
        var sign = ActiveAxisDescending ? -1 : 1;
        node.Children.Sort((a, b) => sign * cmp.Compare(a.Label, b.Label));
        foreach (var c in node.Children) SortAxisTreeRecursive(c);
    }

    /// <summary>
    /// Walk the tree in display order, yielding each node alongside whether it's
    /// a subtotal (internal) or a leaf, plus its absolute display row/col index
    /// (relative to the start of the data area).
    ///
    /// Display order for row axis is "pre-order": for each internal node, emit
    /// the subtotal row first, then recurse into children. The order matches
    /// what BuildMultiRowItems already produces for N=2 and what Excel writes
    /// for N≥3 in compact mode.
    ///
    /// For col axis it's the same plus an additional subtotal column AFTER the
    /// children of each internal node — Excel writes the col subtotal column
    /// to the right of the inner cols, not to the left like the row subtotal.
    /// </summary>
    private static IEnumerable<(AxisNode node, bool isLeaf, bool isSubtotal)> WalkAxisTree(
        AxisNode root, bool isCol)
    {
        // Skip the synthetic root, walk its children in order.
        foreach (var child in root.Children)
            foreach (var entry in WalkAxisTreeRecursive(child, isCol))
                yield return entry;
    }

    private static IEnumerable<(AxisNode node, bool isLeaf, bool isSubtotal)> WalkAxisTreeRecursive(
        AxisNode node, bool isCol)
    {
        if (node.IsLeaf)
        {
            yield return (node, true, false);
            yield break;
        }

        // Row axis subtotal position depends on layout:
        //   compact/outline: subtotal BEFORE children (subtotalTop, default)
        //   tabular: subtotal AFTER children (matches Excel-authored tabular pivots)
        // Col axis convention: subtotal col always AFTER children
        //                     (matches multi_col_authored.xlsx ground truth).
        bool subtotalAfter = isCol || ActiveLayoutMode == "tabular";
        if (!subtotalAfter)
            yield return (node, false, true);

        foreach (var child in node.Children)
            foreach (var entry in WalkAxisTreeRecursive(child, isCol))
                yield return entry;

        if (subtotalAfter)
            yield return (node, false, true);
    }

    /// <summary>Count all internal nodes (subtotal positions) in a tree.</summary>
    private static int CountSubtotalNodes(AxisNode root)
    {
        int count = 0;
        void Recurse(AxisNode n)
        {
            if (!n.IsLeaf && n.Depth > 0) count++;
            foreach (var c in n.Children) Recurse(c);
        }
        Recurse(root);
        return count;
    }

    /// <summary>Count all leaf nodes in a tree.</summary>
    private static int CountLeafNodes(AxisNode root)
    {
        int count = 0;
        void Recurse(AxisNode n)
        {
            if (n.IsLeaf && n.Depth > 0) count++;
            else foreach (var c in n.Children) Recurse(c);
        }
        Recurse(root);
        return count;
    }

    // ==================== Geometry & Cache Readback Helpers ====================

    /// <summary>Computed pivot table extent — anchor + bounding range + key offsets.</summary>
    private readonly struct PivotGeometry
    {
        public PivotGeometry(int anchorCol, int anchorRow, int width, int height, int rowLabelCols, string rangeRef)
        {
            AnchorCol = anchorCol;
            AnchorRow = anchorRow;
            Width = width;
            Height = height;
            RowLabelCols = rowLabelCols;
            RangeRef = rangeRef;
        }
        public int AnchorCol { get; }
        public int AnchorRow { get; }
        public int Width { get; }
        public int Height { get; }
        public int RowLabelCols { get; }
        public string RangeRef { get; }
    }

    /// <summary>
    /// Compute the bounding range and row-label column count for a pivot at the
    /// given anchor with the given field assignments. Used by both initial creation
    /// (BuildPivotTableDefinition) and post-Set rebuild (RebuildFieldAreas) so the
    /// two paths agree on layout.
    ///
    /// Layout assumes the standard compact/outline mode with:
    ///   width  = max(1, rowFieldCount)                    // row labels
    ///          + max(1, colUnique) * max(1, valueCount)    // data cells
    ///          + (colFieldCount > 0 ? 1 : 0)               // grand total column
    ///   height = (colFieldCount > 0 ? 2 : 1)               // header rows
    ///          + max(1, rowUnique)                          // data rows
    ///          + 1                                          // grand total row
    /// Page filter rows are excluded from the range per ECMA-376.
    /// </summary>
    private static PivotGeometry ComputePivotGeometry(
        string position, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields)
    {
        int dataFieldCount = Math.Max(1, valueFields.Count);
        // Compact: all row fields share one column. Outline/Tabular: one column per row field.
        int rowLabelCols = ActiveLayoutMode == "compact"
            ? 1
            : Math.Max(1, rowFieldIndices.Count);

        // CONSISTENCY(subtotals-opts): when subtotals=off, the per-group outer
        // subtotal row (2+ row fields) and outer subtotal column (2+ col fields)
        // are not rendered — shrink the geometry accordingly so location and
        // sheetData stay consistent.
        bool emitSubtotals = ActiveDefaultSubtotal;

        int valueCols, totalCols, dataRowCount, headerRows;

        // N≥3 on either axis, OR any axis is empty (0×*, 2×0): use AxisTree
        // for both width and height counts. The tree handles empty axes
        // naturally (zero leaves, zero subtotals).
        // N≤2 with both axes non-empty: keep the existing specialized formulas
        // (regression-tested via pivot_baselines).
        if (rowFieldIndices.Count >= 3 || colFieldIndices.Count >= 3
            || rowFieldIndices.Count == 0
            || (rowFieldIndices.Count == 2 && colFieldIndices.Count == 0))
        {
            var rowTree = BuildAxisTree(rowFieldIndices, columnData);
            var colTree = BuildAxisTree(colFieldIndices, columnData);

            // Display row count = subtotal positions + leaf positions
            // (the grand total row is added separately below). When subtotals
            // are off, only leaf rows contribute — unless compact mode where
            // parent group headers still appear as label-only rows.
            bool compactLabelRows = !emitSubtotals && ActiveLayoutMode == "compact"
                && rowFieldIndices.Count >= 2;
            int rowSubtotals = (emitSubtotals || compactLabelRows)
                ? CountSubtotalNodes(rowTree) : 0;
            int rowLeaves = CountLeafNodes(rowTree);
            dataRowCount = rowSubtotals + rowLeaves;

            int colSubtotals = emitSubtotals ? CountSubtotalNodes(colTree) : 0;
            int colLeaves = CountLeafNodes(colTree);
            // Per col position: K cells. Plus K grand totals.
            // When there are no col fields, colLeaves=0 but we still need K
            // value columns (one per data field).
            int colPositionCount = colSubtotals + colLeaves;
            valueCols = Math.Max(1, colPositionCount) * dataFieldCount;
            totalCols = colFieldIndices.Count > 0 ? dataFieldCount : 0;

            // Header rows:
            //   colN == 0 && K == 1: single header row with row label caption
            //              + data field name.
            //   colN == 0 && K >  1: TWO header rows — R0 carries the "Values"
            //              axis caption at col B (Excel injects a synthetic
            //              col field for multi-data pivots, and dataCaption
            //              appears at this row), R1 carries the row-label
            //              caption at col A plus the K data field names
            //              across cols B..B+K-1. Verified against Excel-
            //              authored pivot files (ref="A3:F36",
            //              firstHeaderRow=1, firstDataRow=2).
            //   colN >= 1: 1 caption + N_col field-label rows + optional dfRow
            //              when K>1.
            if (colFieldIndices.Count == 0)
                headerRows = dataFieldCount > 1 ? 2 : 1;
            else
                headerRows = 1 + colFieldIndices.Count + (dataFieldCount > 1 ? 1 : 0);
        }
        else if (colFieldIndices.Count >= 2)
        {
            var groups = BuildOuterInnerGroups(
                colFieldIndices[0], colFieldIndices[1], columnData);
            // Each outer group contributes inners.Count leaf cols + 1 subtotal col.
            // When subtotals=off, drop the per-group subtotal col.
            valueCols = groups.Sum(g => (g.inners.Count + (emitSubtotals ? 1 : 0)) * dataFieldCount);
            totalCols = dataFieldCount;

            if (rowFieldIndices.Count >= 2)
            {
                var rowGroups = BuildOuterInnerGroups(
                    rowFieldIndices[0], rowFieldIndices[1], columnData);
                // Each outer group contributes g.inners.Count leaf rows + 1 subtotal row.
                dataRowCount = rowGroups.Sum(g => (emitSubtotals ? 1 : 0) + g.inners.Count);
            }
            else
            {
                dataRowCount = Math.Max(1, ProductOfUniqueValues(rowFieldIndices, columnData));
            }
            headerRows = dataFieldCount > 1 ? 4 : 3;
        }
        else
        {
            int colUnique = ProductOfUniqueValues(colFieldIndices, columnData);
            valueCols = Math.Max(1, colUnique) * dataFieldCount;
            totalCols = colFieldIndices.Count > 0 ? dataFieldCount : 0;

            if (rowFieldIndices.Count >= 2)
            {
                var rowGroups = BuildOuterInnerGroups(
                    rowFieldIndices[0], rowFieldIndices[1], columnData);
                dataRowCount = rowGroups.Sum(g => (emitSubtotals ? 1 : 0) + g.inners.Count);
            }
            else
            {
                dataRowCount = Math.Max(1, ProductOfUniqueValues(rowFieldIndices, columnData));
            }

            if (colFieldIndices.Count > 0)
                headerRows = dataFieldCount > 1 ? 3 : 2;
            else
                // No col fields: renderer always writes 2 header rows (caption + col-label),
                // plus an extra data-field name row when there are multiple value fields.
                headerRows = dataFieldCount > 1 ? 3 : 2;
        }

        // Grand-totals toggles:
        //   rowGrandTotals=false → no rightmost grand-total COLUMN → drop totalCols
        //   colGrandTotals=false → no bottom grand-total ROW → drop the +1 in height
        if (!ActiveRowGrandTotals) totalCols = 0;
        int grandRowHeight = ActiveColGrandTotals ? 1 : 0;

        // insertBlankRow: one blank row after each outer group's subtotal/last leaf.
        int blankRowCount = 0;
        if (ActiveInsertBlankRow && rowFieldIndices.Count >= 2)
        {
            int outerGroups = rowFieldIndices[0] < columnData.Count
                ? columnData[rowFieldIndices[0]].Where(v => !string.IsNullOrEmpty(v)).Distinct().Count()
                : 0;
            blankRowCount = outerGroups;
        }

        int width = rowLabelCols + valueCols + totalCols;
        int height = headerRows + dataRowCount + blankRowCount + grandRowHeight;

        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var endColIdx = anchorColIdx + width - 1;
        var endRow = anchorRow + height - 1;
        var rangeRef = $"{position}:{IndexToCol(endColIdx)}{endRow}";

        return new PivotGeometry(anchorColIdx, anchorRow, width, height, rowLabelCols, rangeRef);
    }

    /// <summary>
    /// Build the &lt;location&gt; element with offsets that match what the
    /// renderer will actually write to sheetData. Shared by BuildPivotTableDefinition
    /// (initial creation) and RebuildFieldAreas (post-Set rebuild) so the two
    /// paths stay in sync.
    ///
    /// For the (N row × 0 col × K data) shape, Excel's canonical layout is a
    /// SINGLE header row at the top of the range, so firstHeaderRow=0 and
    /// firstDataRow=1 (verified against Excel-authored pivot in test_encrypted.xlsx:
    /// 4 row × 0 col × 5 data × 1 filter ⇒ ref="A3:F42", firstHeaderRow=0,
    /// firstDataRow=1, firstDataCol=1). For pivots with col fields, keep the
    /// previous convention (firstHeaderRow=1 = second row of the range, offset
    /// by the existing baselines under tests/pivot_baselines/).
    /// </summary>
    private static Location BuildLocation(
        PivotGeometry geom,
        List<int> rowFieldIndices,
        List<int> colFieldIndices,
        List<(int idx, string func, string showAs, string name)> valueFields,
        int filterCount)
    {
        uint firstHeaderRow;
        uint firstDataRow;
        if (colFieldIndices.Count == 0)
        {
            // colN==0 && K==1: single header row at the top.
            //   compact/outline: firstHeaderRow=0, firstDataRow=1
            //   tabular: firstHeaderRow=1, firstDataRow=1 (header and first
            //            data row share the same row — verified against
            //            Excel-authored tabular pivot)
            // colN==0 && K>1: two header rows — "Values" axis caption at R0
            //   and row-field caption + data field names at R1
            //   (firstHeaderRow=1, firstDataRow=2).
            if (valueFields.Count > 1)
            {
                firstHeaderRow = 1u;
                firstDataRow = 2u;
            }
            else if (ActiveLayoutMode == "tabular")
            {
                firstHeaderRow = 1u;
                firstDataRow = 1u;
            }
            else
            {
                firstHeaderRow = 0u;
                firstDataRow = 1u;
            }
        }
        else
        {
            firstHeaderRow = 1u;
            firstDataRow = (colFieldIndices.Count >= 2 && valueFields.Count > 1) ? 4u
                         : ((valueFields.Count > 1 || colFieldIndices.Count >= 2) ? 3u : 2u);
        }

        var location = new Location
        {
            Reference = geom.RangeRef,
            FirstHeaderRow = firstHeaderRow,
            FirstDataRow = firstDataRow,
            FirstDataColumn = (uint)geom.RowLabelCols
        };

        // rowPageCount / colPageCount: number of rows / columns the page filter
        // area occupies ABOVE the location range. Without these attributes,
        // Excel guesses filter-dropdown placement and ends up drawing the
        // dropdown one row below the actual filter cell (verified in the
        // regenerated encrypted_replica.xlsx). Excel-authored files
        // consistently emit both as 1 when the pivot has any page filter
        // (all filters stacked vertically on the outer row axis).
        //
        // Open XML SDK 3.x does not model these in the typed Location class,
        // so set them as raw unknown attributes. The serializer writes
        // unknown attributes without schema validation. Empty namespace URI
        // means unprefixed, inheriting the element's default namespace
        // (spreadsheetml main).
        if (filterCount > 0)
        {
            location.SetAttribute(new OpenXmlAttribute("rowPageCount", "", filterCount.ToString()));
            location.SetAttribute(new OpenXmlAttribute("colPageCount", "", "1"));
        }

        return location;
    }

    /// <summary>
    /// Reconstruct the per-field columnData from the cache definition + records.
    /// Used by RebuildFieldAreas after Set: the source sheet may not be readily
    /// reachable, but the cache holds the original values (string fields via
    /// sharedItems index, numeric fields directly in &lt;n v=...&gt;). This makes
    /// the rebuild self-contained on the cache part alone.
    /// </summary>
    private static (string[] headers, List<string[]> columnData) ReadColumnDataFromCache(
        PivotCacheDefinition cacheDef, PivotCacheRecords? records)
    {
        var cacheFields = cacheDef.GetFirstChild<CacheFields>();
        if (cacheFields == null) return (Array.Empty<string>(), new List<string[]>());

        var fieldList = cacheFields.Elements<CacheField>().ToList();
        var headers = fieldList.Select(cf => cf.Name?.Value ?? "").ToArray();
        var fieldCount = fieldList.Count;

        // Pre-resolve each field's sharedItems string lookup table (index → text).
        // Numeric fields without enumerated items leave the table empty; their
        // values come straight from <n v=...> in the records below.
        var perFieldStrings = new List<List<string>>(fieldCount);
        for (int f = 0; f < fieldCount; f++)
        {
            var items = fieldList[f].GetFirstChild<SharedItems>();
            var list = new List<string>();
            if (items != null)
            {
                foreach (var child in items.ChildElements)
                {
                    list.Add(child switch
                    {
                        StringItem s => s.Val?.Value ?? string.Empty,
                        NumberItem n => n.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                        DateTimeItem d => d.Val?.Value.ToString("yyyy-MM-dd") ?? string.Empty,
                        BooleanItem b => b.Val?.Value == true ? "true" : "false",
                        _ => string.Empty
                    });
                }
            }
            perFieldStrings.Add(list);
        }

        var recordList = records?.Elements<PivotCacheRecord>().ToList() ?? new List<PivotCacheRecord>();
        var columnData = new List<string[]>(fieldCount);
        for (int f = 0; f < fieldCount; f++)
            columnData.Add(new string[recordList.Count]);

        for (int r = 0; r < recordList.Count; r++)
        {
            var record = recordList[r];
            var children = record.ChildElements.ToList();
            for (int f = 0; f < fieldCount && f < children.Count; f++)
            {
                columnData[f][r] = children[f] switch
                {
                    FieldItem fi when fi.Val?.Value is uint idx
                        && idx < perFieldStrings[f].Count
                        => perFieldStrings[f][(int)idx],
                    NumberItem n => n.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                    StringItem s => s.Val?.Value ?? string.Empty,
                    DateTimeItem d => d.Val?.Value.ToString("yyyy-MM-dd") ?? string.Empty,
                    BooleanItem b => b.Val?.Value == true ? "true" : "false",
                    _ => string.Empty
                };
            }
        }

        return (headers, columnData);
    }

    /// <summary>
    /// Remove every cell in sheetData that falls inside the given pivot range.
    /// Called before re-rendering so stale cells from the previous pivot layout
    /// (e.g. row totals from a wider configuration) do not leak through.
    /// Also called by ExcelHandler.Remove to clean up rendered cells when a pivot is deleted.
    /// </summary>
    internal static void ClearPivotRangeCells(SheetData sheetData, string rangeRef)
    {
        var parts = rangeRef.Split(':');
        if (parts.Length != 2) return;
        var (startCol, startRow) = ParseCellRef(parts[0]);
        var (endCol, endRow) = ParseCellRef(parts[1]);
        var startColIdx = ColToIndex(startCol);
        var endColIdx = ColToIndex(endCol);

        var rowsToRemove = new List<Row>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var rIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rIdx < startRow || rIdx > endRow) continue;

            var cellsToRemove = row.Elements<Cell>()
                .Where(c =>
                {
                    var cref = c.CellReference?.Value ?? "";
                    var (cc, _) = ParseCellRef(cref);
                    var ci = ColToIndex(cc);
                    return ci >= startColIdx && ci <= endColIdx;
                })
                .ToList();
            foreach (var c in cellsToRemove) c.Remove();

            // If the row is now empty AND was entirely inside the pivot, drop it
            // entirely so we don't leave stray <row r="N"/> elements behind.
            if (!row.Elements<Cell>().Any())
                rowsToRemove.Add(row);
        }
        foreach (var r in rowsToRemove) r.Remove();
    }

    /// <summary>
    /// Merge duplicate &lt;row&gt; elements in sheetData into one element per
    /// RowIndex, consolidating all Cell children into the winner in column
    /// order. Also sorts the resulting rows by RowIndex.
    ///
    /// Why: OOXML schema requires each &lt;row r="N"&gt; to be unique within
    /// &lt;sheetData&gt;. When a second pivot is added to a sheet that already
    /// has pivot-rendered rows (e.g. a second pivot at J1 alongside an E1
    /// pivot in the same sheet), the per-renderer "new Row { RowIndex=N };
    /// sheetData.AppendChild(row)" pattern creates duplicates for any row
    /// index the two pivots share. Excel rejects the file with "We found a
    /// problem with some content" at open.
    ///
    /// Call this at the tail of any render path that may have appended rows.
    /// </summary>
    private static void DedupeSheetDataRows(SheetData sheetData)
    {
        // Group by RowIndex. Rows without RowIndex are left alone.
        var byIdx = new Dictionary<uint, List<Row>>();
        foreach (var row in sheetData.Elements<Row>().ToList())
        {
            var idx = row.RowIndex?.Value;
            if (idx == null) continue;
            if (!byIdx.TryGetValue(idx.Value, out var list))
            {
                list = new List<Row>();
                byIdx[idx.Value] = list;
            }
            list.Add(row);
        }

        foreach (var (idx, list) in byIdx)
        {
            if (list.Count <= 1) continue;
            // Merge: keep the first row element, move all cells from the rest
            // into it, then remove the empty duplicates.
            var winner = list[0];
            for (int i = 1; i < list.Count; i++)
            {
                foreach (var cell in list[i].Elements<Cell>().ToList())
                {
                    cell.Remove();
                    winner.AppendChild(cell);
                }
                list[i].Remove();
            }
            // Sort cells by column index for Excel-friendly ordering.
            var sorted = winner.Elements<Cell>()
                .OrderBy(c => ColToIndex((c.CellReference?.Value ?? "A1")
                    .TrimEnd('0','1','2','3','4','5','6','7','8','9')))
                .ToList();
            foreach (var c in sorted) { c.Remove(); winner.AppendChild(c); }
        }

        // Sort rows themselves by RowIndex to keep sheetData ordered.
        var orderedRows = sheetData.Elements<Row>()
            .OrderBy(r => r.RowIndex?.Value ?? 0)
            .ToList();
        foreach (var r in orderedRows) { r.Remove(); sheetData.AppendChild(r); }
    }

    /// <summary>
    /// Re-materialize pivot table cells for all pivots in the given worksheet.
    /// Called before HTML rendering so that existing Excel files whose sheetData
    /// contains stale/minimal pivot cache get properly expanded with hierarchical
    /// row labels and aggregated values.
    /// </summary>
    internal static void RefreshPivotCellsForView(WorksheetPart worksheetPart)
    {
        var pivotParts = worksheetPart.PivotTableParts.ToList();
        if (pivotParts.Count == 0) return;

        foreach (var pivotPart in pivotParts)
        {
            var pivotDef = pivotPart.PivotTableDefinition;
            if (pivotDef == null) continue;

            var cachePart = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault();
            if (cachePart?.PivotCacheDefinition == null) continue;

            var cacheFields = cachePart.PivotCacheDefinition.GetFirstChild<CacheFields>();
            if (cacheFields == null) continue;

            // Read field assignments from the existing definition
            var rowFieldIndices = ReadCurrentFieldIndices(
                pivotDef.RowFields?.Elements<Field>(), f => f.Index?.Value ?? -1);
            var colFieldIndices = ReadCurrentFieldIndices(
                pivotDef.ColumnFields?.Elements<Field>(), f => f.Index?.Value ?? -1);
            var filterFieldIndices = ReadCurrentFieldIndices(
                pivotDef.PageFields?.Elements<PageField>(), f => f.Field?.Value ?? -1);
            var valueFields = ReadCurrentDataFields(pivotDef.DataFields);

            if (valueFields.Count == 0) continue;

            // Read cache data
            var (cacheHeaders, cacheColumnData) = ReadColumnDataFromCache(
                cachePart.PivotCacheDefinition,
                cachePart.GetPartsOfType<PivotTableCacheRecordsPart>().FirstOrDefault()?.PivotCacheRecords);
            if (cacheColumnData.Count == 0) continue;

            // Detect layout mode from existing definition
            string? layoutMode = null;
            if (pivotDef.Compact?.Value == false)
            {
                var firstAxisField = pivotDef.PivotFields?.Elements<PivotField>()
                    .FirstOrDefault(pf => pf.Axis != null);
                if (firstAxisField?.Outline?.Value == false)
                    layoutMode = "tabular";
                else
                    layoutMode = "outline";
            }

            // Detect grand totals from definition (OOXML mapping is swapped)
            bool? rowGT = pivotDef.ColumnGrandTotals?.Value == false ? false : null;
            bool? colGT = pivotDef.RowGrandTotals?.Value == false ? false : null;

            // Detect subtotals
            bool? defaultSubtotal = null;
            if (pivotDef.PivotFields != null)
            {
                foreach (var pf in pivotDef.PivotFields.Elements<PivotField>())
                {
                    if (pf.DefaultSubtotal?.Value == false)
                    {
                        defaultSubtotal = false;
                        break;
                    }
                }
            }

            // Push thread-static options for the render pass
            var prevLayout = _layoutMode;
            var prevRowGT = _rowGrandTotals;
            var prevColGT = _colGrandTotals;
            var prevSubtotal = _defaultSubtotal;
            try
            {
                _layoutMode = layoutMode;
                _rowGrandTotals = rowGT;
                _colGrandTotals = colGT;
                _defaultSubtotal = defaultSubtotal;

                // Determine anchor position from the existing Location
                var locationRef = pivotDef.Location?.Reference?.Value;
                var anchorRef = locationRef?.Split(':')[0] ?? "A1";

                // Clear old cells and re-render
                var ws = worksheetPart.Worksheet;
                var sheetData = ws?.GetFirstChild<SheetData>();
                if (ws != null && sheetData != null && locationRef != null)
                {
                    ClearPivotRangeCells(sheetData, locationRef);

                    // Try to get source column styles for number formatting
                    uint?[]? sourceColumnStyleIds = null;
                    try
                    {
                        var wbPart = worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
                        var wsSource = cachePart.PivotCacheDefinition.CacheSource?.WorksheetSource;
                        if (wbPart != null && wsSource?.Sheet?.Value is string srcSheetName
                            && wsSource.Reference?.Value is string srcRef)
                        {
                            var sheetRef = wbPart.Workbook?.Sheets?.Elements<Sheet>()
                                .FirstOrDefault(s => s.Name?.Value == srcSheetName);
                            if (sheetRef?.Id?.Value is string relId
                                && wbPart.GetPartById(relId) is WorksheetPart srcWsPart)
                            {
                                var (_, _, ids) = ReadSourceData(srcWsPart, srcRef);
                                sourceColumnStyleIds = ids;
                            }
                        }
                    }
                    catch { /* best-effort */ }

                    RenderPivotIntoSheet(
                        worksheetPart, anchorRef, cacheHeaders, cacheColumnData,
                        rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices,
                        sourceColumnStyleIds);

                    DedupeSheetDataRows(sheetData);
                }
            }
            finally
            {
                _layoutMode = prevLayout;
                _rowGrandTotals = prevRowGT;
                _colGrandTotals = prevColGT;
                _defaultSubtotal = prevSubtotal;
            }
        }
    }
}
