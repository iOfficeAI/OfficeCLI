// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

internal static partial class PivotTableHelper
{

    // ==================== Parse Helpers ====================

    private static List<int> ParseFieldListWithWarning(Dictionary<string, string> props, string key, string[] headers)
    {
        var result = ParseFieldList(props, key, headers);
        if (result.Count == 0 && props.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value))
        {
            var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
            Console.Error.WriteLine($"WARNING: No matching fields for {key}={value}. Available: {available}");
        }
        return result;
    }

    private static List<(int idx, string func, string showAs, string name)> ParseValueFieldsWithWarning(
        Dictionary<string, string> props, string key, string[] headers)
    {
        var result = ParseValueFields(props, key, headers);
        if (result.Count == 0 && props.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value))
        {
            var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
            Console.Error.WriteLine($"WARNING: No matching fields for {key}={value}. Available: {available}");
        }
        return result;
    }

    // R4-2: Unicode field names may reach us in different normalization forms
    // (e.g. source header in NFD "e\u0301" vs user input in NFC "\u00E9"). An
    // ordinal compare would fail on semantically equivalent strings and report
    // the field as missing. Normalize both sides to NFC before lookup so
    // composed and decomposed spellings bind to the same header. We only
    // normalize for matching — stored header text is left unchanged.
    private static bool FieldNameMatches(string? header, string candidate)
    {
        if (header == null) return false;
        // Trim surrounding whitespace on both sides so header cells with
        // incidental leading/trailing spaces (a common paste-from-Excel
        // artefact) still resolve against clean user input. NFC normalisation
        // from Round 4 R4-2 is preserved. CONSISTENCY(pivot-field-matching).
        return header.Trim().Normalize(NormalizationForm.FormC)
            .Equals(candidate.Trim().Normalize(NormalizationForm.FormC), StringComparison.OrdinalIgnoreCase);
    }

    private static List<int> ParseFieldList(Dictionary<string, string> props, string key, string[] headers)
    {
        if (!props.TryGetValue(key, out var value) || string.IsNullOrEmpty(value))
            return new List<int>();

        var result = new List<int>();
        // CONSISTENCY(field-area-dedup): dedup within the same axis (rows/cols/filters).
        // A field index must appear at most once per axis; repeated tokens keep the first
        // occurrence and skip subsequent ones, matching cross-axis dedup semantics.
        var seen = new HashSet<int>();
        foreach (var f in value.Split(','))
        {
            var name = f.Trim();
            if (string.IsNullOrEmpty(name)) continue;

            // CONSISTENCY(field-name-validation): a numeric token is treated
            // as a column index (out-of-range still silently dropped — that
            // is the legacy contract used by tests with index hints). A
            // non-numeric token MUST resolve to an existing header, else we
            // throw with the available header list so users can fix typos
            // immediately instead of seeing an empty / wrong pivot.
            if (int.TryParse(name, out var idx))
            {
                if (idx >= 0 && idx < headers.Length && seen.Add(idx)) result.Add(idx);
                continue;
            }
            int found = -1;
            for (int i = 0; i < headers.Length; i++)
                if (FieldNameMatches(headers[i], name)) { found = i; break; }
            // CONSISTENCY(date-grouping-passthrough): unrecognized grouping
            // suffixes (e.g. "Date:hours") survive ApplyDateGrouping as
            // literals. Strip the suffix and re-resolve so the bare field
            // name still binds — matches the existing best-effort fuzz
            // contract that says invalid grouping must not crash.
            if (found < 0)
            {
                var colon = name.IndexOf(':');
                if (colon > 0)
                {
                    var bare = name.Substring(0, colon);
                    for (int i = 0; i < headers.Length; i++)
                        if (FieldNameMatches(headers[i], bare)) { found = i; break; }
                }
            }
            if (found < 0)
            {
                var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
                throw new ArgumentException($"field '{name}' not found in source headers: {available}");
            }
            if (seen.Add(found)) result.Add(found);
        }
        return result;
    }

    private static List<(int idx, string func, string showAs, string name)> ParseValueFields(
        Dictionary<string, string> props, string key, string[] headers)
    {
        if (!props.TryGetValue(key, out var value) || string.IsNullOrEmpty(value))
            return new List<(int, string, string, string)>();

        // CONSISTENCY(aggregate-override): the optional sibling 'aggregate'
        // property is a comma-list aligned positionally with 'values'. It
        // overrides the per-field func parsed from the colon-suffix syntax.
        // This lets users write `values=Sales,Sales aggregate=sum,count`
        // instead of `values=Sales:sum,Sales:count` — both forms are
        // equivalent. Per-spec colon syntax still wins for any slot the
        // aggregate list does not cover (shorter list ⇒ remaining slots
        // keep their parsed func).
        string[]? aggregateOverrides = null;
        if (props.TryGetValue("aggregate", out var aggSpec) && !string.IsNullOrEmpty(aggSpec))
            aggregateOverrides = aggSpec.Split(',').Select(s => s.Trim().ToLowerInvariant()).ToArray();

        var result = new List<(int idx, string func, string showAs, string name)>();
        var specs = value.Split(',');
        for (int specIndex = 0; specIndex < specs.Length; specIndex++)
        {
            var spec = specs[specIndex];
            // Format: "FieldName" | "FieldName:func" | "FieldName:func:showAs"
            //   default func    = sum
            //   default showAs  = normal
            // showAs accepts: normal | percent_of_total | percent_of_row |
            //                 percent_of_col | running_total | (+ camelCase aliases)
            // R11-2: Parse right-to-left so field names containing literal
            // colons (e.g. "A:B:sum" → field "A:B", func "sum") work without
            // requiring users to escape. Strategy:
            //   1. Split into all colon segments.
            //   2. Peek the rightmost segment: if it's a known showAs token,
            //      consume it as showAs, then peek again for func.
            //   3. Otherwise, if the rightmost segment is a known aggregate
            //      function, consume it as func.
            //   4. Anything not consumed (joined back with ':') is the field
            //      name, preserving any embedded colons.
            // The 1-segment case ("Sales") and 2-segment case ("Sales:sum") and
            // 3-segment case ("Sales:sum:percent_of_total") all keep working
            // because trailing tokens are still recognized — only the field
            // name parsing changes.
            var parts = spec.Trim().Split(':');
            string fieldName;
            string func = "sum";
            string showAs = "normal";
            // R34-3: optional custom display name. When non-null, overrides
            // the auto-generated "Sum of <Header>" displayName below. Valid
            // forms (right-to-left, all backwards-compatible):
            //   Field:Func:ShowAs:Name           ← 4-seg, both known tokens
            //   Field:Func:Name                  ← 3-seg, last is non-token
            //   Field:Func=name=Name             ← (not supported here)
            // The 1/2/3-seg cases with known trailing tokens are unchanged.
            string? customName = null;
            // R34-3: an explicit name= segment unambiguously marks the
            // custom DataField.Name slot, sidestepping the ambiguity that
            // makes a bare 3rd unknown token impossible to distinguish
            // from a typo in showAs (which existing strict-enum tests rely
            // on rejecting). Strip it before the walker runs so the
            // remaining 1/2/3-seg cases parse exactly as before.
            //   Sales:Sum:name=TotalSales
            //   Sales:Sum:percent_of_total:name=SalesShare
            for (int p = parts.Length - 1; p >= 1; p--)
            {
                var trimmed = parts[p].Trim();
                if (trimmed.StartsWith("name=", StringComparison.OrdinalIgnoreCase))
                {
                    customName = trimmed.Substring("name=".Length).Trim();
                    var next = new string[parts.Length - 1];
                    Array.Copy(parts, 0, next, 0, p);
                    if (p < parts.Length - 1)
                        Array.Copy(parts, p + 1, next, p, parts.Length - p - 1);
                    parts = next;
                    break;
                }
            }
            if (parts.Length == 1)
            {
                fieldName = parts[0].Trim();
            }
            else
            {
                int consumed = 0;
                var last = parts[parts.Length - 1].Trim().ToLowerInvariant();
                // R34-3: 4-segment Field:Func:ShowAs:Name form. The 4th
                // slot is treated as a custom DataField.Name only when
                // slot 3 is a recognized showAs token AND slot 2 is a
                // recognized aggregate — i.e. unambiguously past the
                // walker's known-token zone. Bare 3-segment unknowns
                // ("Sales:sum:bogus") deliberately keep flowing to the
                // strict "invalid showDataAs" rejection so typos still
                // surface (CONSISTENCY(strict-enums)).
                if (customName == null
                    && parts.Length >= 4
                    && !IsKnownShowAsToken(last)
                    && !IsKnownAggregateToken(last))
                {
                    var slot3 = parts[parts.Length - 2].Trim().ToLowerInvariant();
                    var slot2 = parts[parts.Length - 3].Trim().ToLowerInvariant();
                    if (IsKnownShowAsToken(slot3) && IsKnownAggregateToken(slot2))
                    {
                        customName = parts[parts.Length - 1].Trim();
                        Array.Resize(ref parts, parts.Length - 1);
                        last = parts[parts.Length - 1].Trim().ToLowerInvariant();
                    }
                }
                if (parts.Length >= 2 && IsKnownShowAsToken(last))
                {
                    showAs = last;
                    consumed = 1;
                    if (parts.Length - consumed >= 2)
                    {
                        var prev = parts[parts.Length - 1 - consumed].Trim().ToLowerInvariant();
                        if (IsKnownAggregateToken(prev))
                        {
                            func = prev;
                            consumed = 2;
                        }
                    }
                }
                else if (IsKnownAggregateToken(last))
                {
                    func = last;
                    consumed = 1;
                }
                else
                {
                    // Unknown trailing token: fall back to legacy left-to-right
                    // semantics so existing error messages (invalid showDataAs /
                    // unknown aggregate) still surface from ParseShowDataAs /
                    // ParseSubtotal downstream.
                    fieldName = parts[0].Trim();
                    func = parts.Length > 1 ? parts[1].Trim().ToLowerInvariant() : "sum";
                    showAs = parts.Length > 2 ? parts[2].Trim().ToLowerInvariant() : "normal";
                    goto afterParse;
                }
                var nameParts = parts.Take(parts.Length - consumed).ToList();
                // Drop trailing empty segments — the legacy "Sales::percent_of_total"
                // form (empty func slot, default "sum") leaves a "" between the
                // field name and the consumed showAs token. Right-to-left parsing
                // would otherwise concatenate "Sales:" as the field name and fail
                // header lookup. The empty func will be defaulted to "sum" below.
                while (nameParts.Count > 1 && string.IsNullOrEmpty(nameParts[nameParts.Count - 1]))
                    nameParts.RemoveAt(nameParts.Count - 1);
                fieldName = string.Join(":", nameParts).Trim();
                // Edge: "sum" alone with no field name (e.g. spec was ":sum")
                // → fall through to the same "field not found" error path.
            }
            afterParse:;

            // CONSISTENCY(pivot-roundtrip / R9-2): Get readback emits dataField{N}
            // as "{displayName}:{func}:{fieldIdx}" where displayName has the form
            // "Sum of Sales" and the third slot is a numeric cacheField index
            // (NOT a showAs token). Accept this shape so the output of Get can
            // be fed straight back into Set values=... without translation.
            // Disambiguation: only switch into round-trip mode when parts[0]
            // starts with a known English aggregate display prefix
            // ("Sum of ", "Count of ", ...). Otherwise the third slot stays
            // a showAs token, preserving the existing "Sales:sum:42" → invalid
            // showDataAs throw contract.
            var displayPrefixes = new[]
            {
                "Sum of ", "Count of ", "Average of ", "Max of ", "Min of ",
                "Product of ", "Count Numbers of ", "StdDev of ", "StdDevp of ",
                "Var of ", "Varp of ", "Std Dev of ", "Std Dev p of "
            };
            bool isGetReadbackShape = false;
            foreach (var p in displayPrefixes)
            {
                if (fieldName.StartsWith(p, StringComparison.OrdinalIgnoreCase))
                {
                    fieldName = fieldName.Substring(p.Length).Trim();
                    isGetReadbackShape = true;
                    break;
                }
            }
            int? roundTripFieldIdx = null;
            if (isGetReadbackShape && parts.Length > 2 && int.TryParse(parts[2].Trim(), out var rtIdx))
            {
                // Get readback packs cacheField index in slot 3; reset showAs
                // to canonical default (the sibling dataField{N}.showAs key
                // carries showDataAs round-trip).
                roundTripFieldIdx = rtIdx;
                showAs = "normal";
            }

            // Empty func slot ("Sales:" or "Sales::percent_of_total") is a
            // common user mistake from optional-segment trailing colons. Treat
            // as the documented default ("sum") rather than crashing on
            // func[0] below. This keeps the showAs slot positionally addressable.
            if (string.IsNullOrEmpty(func)) func = "sum";

            // CONSISTENCY(aggregate-override): if aggregate=<list> was passed
            // and has an entry at this position, it wins over the colon form.
            if (aggregateOverrides != null && specIndex < aggregateOverrides.Length
                && !string.IsNullOrEmpty(aggregateOverrides[specIndex]))
                func = aggregateOverrides[specIndex];

            int fieldIdx = -1;
            // CONSISTENCY(pivot-roundtrip / R9-2): when the Get readback shape
            // gave us an explicit numeric cacheField index, prefer it over the
            // (possibly stripped) display name. This makes Set values=GetOutput
            // robust even if the source headers were renamed between Get and
            // Set, and removes any ambiguity from the prefix-strip heuristic.
            if (roundTripFieldIdx.HasValue)
            {
                if (roundTripFieldIdx.Value < 0 || roundTripFieldIdx.Value >= headers.Length)
                    throw new ArgumentException(
                        $"field index {roundTripFieldIdx.Value} out of range (0..{headers.Length - 1})");
                fieldIdx = roundTripFieldIdx.Value;
            }
            else if (int.TryParse(fieldName, out var idx))
            {
                // CONSISTENCY(strict-enums / R8-6): a numeric token is a
                // column index. Out-of-range indices used to silently drop
                // the value-field, producing an empty pivot with no error.
                // Reject up front with the available-index range so users
                // catch the typo immediately (mirrors the throw used for
                // unknown field names).
                if (idx < 0 || idx >= headers.Length)
                    throw new ArgumentException(
                        $"field index {idx} out of range (0..{headers.Length - 1})");
                fieldIdx = idx;
            }
            else
            {
                for (int i = 0; i < headers.Length; i++)
                    if (FieldNameMatches(headers[i], fieldName)) { fieldIdx = i; break; }
                // CONSISTENCY(field-name-validation): non-numeric token must
                // resolve. Same throw shape as ParseFieldList.
                if (fieldIdx < 0)
                {
                    var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
                    throw new ArgumentException($"field '{fieldName}' not found in source headers: {available}");
                }
            }

            if (fieldIdx >= 0 && fieldIdx < headers.Length)
            {
                // R34-3: a user-supplied 4th (or 3rd-when-no-showAs) segment
                // becomes the DataField.Name (the column header rendered in
                // the pivot output). Falls back to "{Func} of {Header}" when
                // absent — matches Excel's default and preserves the
                // round-trip shape the existing prefix-strip relies on.
                var displayName = !string.IsNullOrEmpty(customName)
                    ? customName!
                    : $"{char.ToUpper(func[0])}{func[1..]} of {headers[fieldIdx]}";
                result.Add((fieldIdx, func, showAs, displayName));
            }
        }
        return result;
    }

    /// <summary>
    /// Map a user-facing showAs string to the OOXML ShowDataAsValues enum.
    /// Returns null for "normal" (no-op; DataField element omits the attribute).
    /// Accepts both snake_case and camelCase forms so users don't get punished
    /// by the convention split between CLI params (snake) and XML schema (camel).
    /// </summary>
    /// <summary>
    /// Inverse of ParseShowDataAs: map a stored OOXML ShowDataAsValues enum
    /// back to the canonical snake_case token used in CLI input/output.
    /// Used by ReadPivotTableProperties to surface dataField{N}.showAs in
    /// Get readback. Defaults to "normal" for unmapped enum values so the
    /// caller can suppress them via the Normal short-circuit.
    /// </summary>
    // CONSISTENCY(enum-innertext): switch over EnumValue<T>.InnerText (the
    // OOXML attribute literal), not over C# enum-value equality. OpenXML SDK
    // v3 exposes ShowDataAsValues.Percent AND ShowDataAsValues.PercentOfTotal
    // as distinct values; XML "percent" deserializes to .Percent, and
    // EnumValue<T>.ToString() yields garbage like "showdataasvalues { }"
    // (same class of bug as LineSpacingRuleValues.Auto.ToString() documented
    // in CLAUDE.md "Known API Quirks"). Reading InnerText sidesteps both
    // traps — no silent enum-fall-through, no SDK ToString() footguns.
    private static string ShowDataAsToCanonicalToken(EnumValue<ShowDataAsValues>? showDataAs)
    {
        var raw = showDataAs?.InnerText ?? "";
        return raw switch
        {
            "" or "normal" => "normal",
            // OOXML has two distinct ShowDataAs enum values ("percent" and
            // "percentOfTotal") that share the same canonical snake_case
            // output — matching ParseShowDataAs which already accepts both
            // input aliases for .PercentOfTotal. Keep the longer-form
            // canonical so pre-existing round-trip assertions (which expect
            // "percent_of_total") stay green.
            "percent" or "percentOfTotal" => "percent_of_total",
            "percentOfRow" => "percent_of_row",
            "percentOfCol" => "percent_of_col",
            "runTotal" => "running_total",
            "difference" => "difference",
            "percentDiff" => "percent_diff",
            "index" => "index",
            _ => raw,
        };
    }

    /// <summary>
    /// True if the showAs token is any of the percent_* family
    /// (percent_of_total / _row / _col + camelCase / "percent" aliases).
    /// Used to force DataField.NumberFormatId to built-in 10 ("0.00%") so
    /// computed fractions display as percentages instead of bare decimals.
    /// </summary>
    private static bool IsPercentShowAs(string showAs)
    {
        return showAs.ToLowerInvariant() switch
        {
            "percent_of_total" or "percentoftotal" or "percent" => true,
            "percent_of_row" or "percentofrow" => true,
            "percent_of_col" or "percent_of_column" or "percentofcol" or "percentofcolumn" => true,
            _ => false,
        };
    }

    private static ShowDataAsValues? ParseShowDataAs(string showAs)
    {
        return showAs.ToLowerInvariant() switch
        {
            "" or "normal" => null,
            "percent_of_total" or "percentoftotal" or "percent" => ShowDataAsValues.PercentOfTotal,
            "percent_of_row" or "percentofrow" => ShowDataAsValues.PercentOfRaw,
            "percent_of_col" or "percent_of_column" or "percentofcol" or "percentofcolumn" => ShowDataAsValues.PercentOfColumn,
            "running_total" or "runningtotal" or "runtotal" => ShowDataAsValues.RunTotal,
            // CONSISTENCY(strict-enums): difference / percent_diff / index are
            // accepted by the OOXML ShowDataAsValues enum, but ApplyShowDataAs1x1
            // has no matrix transformation for them, so rendered cells would
            // silently equal the raw aggregate. Reject up front until a proper
            // renderer exists, mirroring the invalid-sort / invalid-aggregate
            // policy from Round 1.
            "difference" or "diff" or "percent_diff" or "percentdiff" or "index" =>
                throw new ArgumentException(
                    $"showDataAs '{showAs}' is not yet supported by the renderer " +
                    "(would silently return raw aggregate). Supported: normal, " +
                    "percent_of_total, percent_of_row, percent_of_col, running_total."),
            // CONSISTENCY(strict-enums): unknown showAs tokens are rejected
            // up front so users see typos at Add/Set time, not on render.
            _ => throw new ArgumentException(
                $"invalid showDataAs: '{showAs}'. Valid: normal, percent_of_total, percent_of_row, " +
                "percent_of_col, running_total"),
        };
    }

    // R11-2: Right-to-left value-spec parser support. Token recognizers
    // mirror the cases ParseSubtotal / ParseShowDataAs accept (lowercase
    // canonical only — we lowercase the token before calling). Keep these
    // in sync if new aggregates / showAs tokens are added downstream.
    private static bool IsKnownAggregateToken(string token) => token switch
    {
        "sum" or "count" or "countnums" or "countnum" or "average" or "avg" or
        "max" or "min" or "product" or "stddev" or "std" or "stddevp" or "stdp" or
        "var" or "variance" or "varp" => true,
        _ => false,
    };

    private static bool IsKnownShowAsToken(string token) => token switch
    {
        "normal" or
        "percent_of_total" or "percentoftotal" or "percent" or
        "percent_of_row" or "percentofrow" or
        "percent_of_col" or "percent_of_column" or "percentofcol" or "percentofcolumn" or
        "running_total" or "runningtotal" or "runtotal" => true,
        _ => false,
    };

    /// <summary>
    /// R15-5: canonical English display prefix for the auto-generated
    /// DataField name ("Sum of Sales", "Count of Sales", ...). Matches the
    /// displayPrefixes table used by the values-spec round-trip parser.
    /// </summary>
    private static string AggregateDisplayName(string func) => func.ToLowerInvariant() switch
    {
        "sum" => "Sum",
        "count" => "Count",
        "countnums" or "countnum" => "Count Numbers",
        "average" or "avg" => "Average",
        "max" => "Max",
        "min" => "Min",
        "product" => "Product",
        "stddev" or "std" => "StdDev",
        "stddevp" or "stdp" => "StdDevp",
        "var" or "variance" => "Var",
        "varp" => "Varp",
        _ => "Sum",
    };

    /// <summary>
    /// R15-5: true when the current DataField name still matches the auto-
    /// generated "<AggDisplay> of <sourceHeader>" form, so a Set aggregate
    /// call is safe to rewrite it. Any name that does not end in " of
    /// <sourceHeader>" is treated as user-provided and left alone.
    /// </summary>
    private static bool LooksLikeAutoDataFieldName(string name, string sourceHeader)
    {
        if (string.IsNullOrEmpty(name)) return true;
        var suffix = " of " + sourceHeader;
        if (!name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return false;
        var prefix = name.Substring(0, name.Length - suffix.Length);
        return prefix is "Sum" or "Count" or "Count Numbers" or "Average" or "Max"
            or "Min" or "Product" or "StdDev" or "StdDevp" or "Var" or "Varp"
            or "Std Dev" or "Std Dev p";
    }

    private static DataConsolidateFunctionValues ParseSubtotal(string func)
    {
        return func.ToLowerInvariant() switch
        {
            "sum" => DataConsolidateFunctionValues.Sum,
            "count" => DataConsolidateFunctionValues.Count,
            "countnums" or "countnum" => DataConsolidateFunctionValues.CountNumbers,
            "average" or "avg" => DataConsolidateFunctionValues.Average,
            "max" => DataConsolidateFunctionValues.Maximum,
            "min" => DataConsolidateFunctionValues.Minimum,
            "product" => DataConsolidateFunctionValues.Product,
            "stddev" or "std" => DataConsolidateFunctionValues.StandardDeviation,
            "stddevp" or "stdp" => DataConsolidateFunctionValues.StandardDeviationP,
            "var" or "variance" => DataConsolidateFunctionValues.Variance,
            "varp" => DataConsolidateFunctionValues.VarianceP,
            // CONSISTENCY(strict-enums): mirror ParseShowDataAs / ParseFieldList —
            // unknown tokens throw at Add/Set time so typos surface immediately
            // instead of silently falling back to sum and producing the wrong
            // numbers on render (Bug #3).
            _ => throw new ArgumentException(
                $"invalid aggregate: '{func}'. Valid: sum, count, countNums, average/avg, " +
                "max, min, product, stdDev/std, stdDevp/stdp, var/variance, varP"),
        };
    }

    /// <summary>
    /// Aggregate a bag of numeric values using the given subtotal function.
    /// Matches LibreOffice's ScDPAggData semantics (sc/source/core/data/dptabres.cxx):
    ///   sum / product / min / max / count : trivial
    ///   countNums : count of numeric entries (identical to count here because
    ///     the caller only places parsed numerics into the bag)
    ///   average : arithmetic mean
    ///   stdDev  : sample std-dev  (sqrt(Σ(x-μ)²/(n-1))), requires n≥2
    ///   stdDevp : population std-dev (sqrt(Σ(x-μ)²/n)), requires n≥1
    ///   var     : sample variance (Σ(x-μ)²/(n-1)), requires n≥2
    ///   varp    : population variance (Σ(x-μ)²/n), requires n≥1
    /// Returns 0 for empty input and for stdDev/var when n&lt;2, matching the
    /// existing 0-on-empty convention that the rest of the renderer assumes.
    /// </summary>
    private static double ReducePivotValues(IEnumerable<double> values, string func)
    {
        var arr = values as double[] ?? values.ToArray();
        if (arr.Length == 0) return 0;
        switch (func.ToLowerInvariant())
        {
            case "sum": return arr.Sum();
            case "count": return arr.Length;
            case "countnums":
            case "countnum": return arr.Length;
            case "average":
            case "avg": return arr.Average();
            case "min": return arr.Min();
            case "max": return arr.Max();
            case "product":
                double p = 1;
                foreach (var v in arr) p *= v;
                return p;
            case "stddev":
            case "std":
            {
                if (arr.Length < 2) return 0;
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return Math.Sqrt(sq / (arr.Length - 1));
            }
            case "stddevp":
            case "stdp":
            {
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return Math.Sqrt(sq / arr.Length);
            }
            case "var":
            case "variance":
            {
                if (arr.Length < 2) return 0;
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return sq / (arr.Length - 1);
            }
            case "varp":
            {
                var mean = arr.Average();
                var sq = arr.Sum(x => (x - mean) * (x - mean));
                return sq / arr.Length;
            }
            default: return arr.Sum();
        }
    }

    /// <summary>
    /// Apply a showDataAs transform to a 1×1×K pivot matrix for data field d.
    /// Used by RenderPivotIntoSheet (the 1 row × 1 col × K data inline
    /// renderer). Other renderers share the same normalization by value
    /// type but not by matrix layout, so each renderer post-processes its
    /// own buckets after aggregation.
    ///
    /// Supported modes:
    ///   normal            — no-op
    ///   percent_of_total  — divide everything by grandTotals[d]
    ///   percent_of_row    — divide each (r,c) by rowTotals[r] (the whole row shares the divisor)
    ///   percent_of_col    — divide each (r,c) by colTotals[c]
    ///   running_total     — in-row cumulative sum across cols, left→right;
    ///                       rowTotals/grandTotals unchanged (cumulative ends at row total)
    /// Unknown modes are silently treated as "normal" so new modes added to
    /// ParseShowDataAs don't explode old renderers.
    /// </summary>
    private static void ApplyShowDataAs1x1(
        string mode, double?[,,] matrix, double[,] rowTotals, double[,] colTotals,
        double[] grandTotals, int rowCount, int colCount, int d)
    {
        switch (mode.ToLowerInvariant())
        {
            case "" or "normal":
                return;

            case "percent_of_total" or "percentoftotal" or "percent":
            {
                var gt = grandTotals[d];
                if (gt == 0) return;
                for (int r = 0; r < rowCount; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        if (matrix[r, c, d].HasValue)
                            matrix[r, c, d] = matrix[r, c, d]!.Value / gt;
                    }
                    rowTotals[r, d] = rowTotals[r, d] / gt;
                }
                for (int c = 0; c < colCount; c++)
                    colTotals[c, d] = colTotals[c, d] / gt;
                grandTotals[d] = 1.0;
                return;
            }

            case "percent_of_row" or "percentofrow":
            {
                for (int r = 0; r < rowCount; r++)
                {
                    var rt = rowTotals[r, d];
                    if (rt == 0) continue;
                    for (int c = 0; c < colCount; c++)
                    {
                        if (matrix[r, c, d].HasValue)
                            matrix[r, c, d] = matrix[r, c, d]!.Value / rt;
                    }
                    rowTotals[r, d] = 1.0;
                }
                // Col totals and grand lose their direct interpretation under
                // "percent of row" (they're sums of ratios across heterogeneous
                // row bases). Excel renders them as the sum of the per-row
                // ratios across the column, which equals colSum / grandTotal
                // only if all rows share the same total. Mirror that here:
                // recompute as "percent of total" for the col and grand cells
                // so the displayed numbers sum to 100% across each row but
                // col totals reflect "this col's share of the grand total".
                var grand = grandTotals[d];
                if (grand != 0)
                {
                    for (int c = 0; c < colCount; c++)
                        colTotals[c, d] = colTotals[c, d] / grand;
                    grandTotals[d] = 1.0;
                }
                return;
            }

            case "percent_of_col" or "percent_of_column" or "percentofcol" or "percentofcolumn":
            {
                for (int c = 0; c < colCount; c++)
                {
                    var ct = colTotals[c, d];
                    if (ct == 0) continue;
                    for (int r = 0; r < rowCount; r++)
                    {
                        if (matrix[r, c, d].HasValue)
                            matrix[r, c, d] = matrix[r, c, d]!.Value / ct;
                    }
                    colTotals[c, d] = 1.0;
                }
                var grand = grandTotals[d];
                if (grand != 0)
                {
                    for (int r = 0; r < rowCount; r++)
                        rowTotals[r, d] = rowTotals[r, d] / grand;
                    grandTotals[d] = 1.0;
                }
                return;
            }

            case "running_total" or "runningtotal" or "runtotal":
            {
                // In-row cumulative sum across cols, left→right. Cells with
                // null values count as 0 in the running sum but remain null
                // in the output so Excel shows blank instead of the previous
                // cumulative value (matches Excel's "(blank)" behavior).
                for (int r = 0; r < rowCount; r++)
                {
                    double running = 0;
                    for (int c = 0; c < colCount; c++)
                    {
                        if (matrix[r, c, d].HasValue)
                        {
                            running += matrix[r, c, d]!.Value;
                            matrix[r, c, d] = running;
                        }
                    }
                }
                // Row / col / grand totals are left as-is: running total's
                // final-column value already equals the row total, and col /
                // grand totals don't have a natural running interpretation
                // across rows in Excel's semantics.
                return;
            }

            default:
                return;
        }
    }

    private static (string col, int row) ParseCellRef(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        var col = cellRef[..i].ToUpperInvariant();
        var row = int.TryParse(cellRef[i..], out var r) ? r : 1;
        return (col, row);
    }

    private static int ColToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    private static string IndexToCol(int index)
    {
        // Inverse of ColToIndex (1-based: A=1, Z=26, AA=27, ...)
        var sb = new System.Text.StringBuilder();
        while (index > 0)
        {
            int rem = (index - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Multiply the cardinality (distinct non-empty values) of each field in the
    /// given index list. Used to size the pivot table's rendered area for the
    /// Location.ref range. Returns 1 when the list is empty (so layout math stays
    /// safe in pivots that have only column fields, only row fields, etc.).
    /// </summary>
    private static int ProductOfUniqueValues(List<int> fieldIndices, List<string[]> columnData)
    {
        if (fieldIndices.Count == 0) return 1;
        int product = 1;
        foreach (var idx in fieldIndices)
        {
            if (idx < 0 || idx >= columnData.Count) continue;
            var unique = columnData[idx].Where(v => !string.IsNullOrEmpty(v)).Distinct().Count();
            product *= Math.Max(1, unique);
        }
        return product;
    }
}
