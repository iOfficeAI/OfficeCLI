// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;

namespace OfficeCli.Core;

/// <summary>
/// Prefixes Excel 2016+ dynamic-array and "modern" function names with
/// <c>_xlfn.</c> when emitting OOXML. Excel refuses to resolve bare
/// post-2016 function names (e.g. <c>SEQUENCE(5)</c> → <c>#NAME?</c>)
/// unless the XML formula uses the namespaced form (<c>_xlfn.SEQUENCE(5)</c>).
/// Excel strips the prefix back out when displaying the formula to the user,
/// so the round-trip is transparent.
///
/// Also handles <c>_xlfn._xlws.</c> (worksheet-only namespace) for FILTER
/// and <c>_xlfn.ANCHORARRAY</c> for spilled-range references (<c>A1#</c> stays
/// user-facing; the XML serialization is a separate concern handled by Excel).
/// </summary>
public static class ModernFunctionQualifier
{
    // Functions that need just _xlfn.
    // Source: MS-XLSX / Excel 2016+ dynamic-array + modern function catalogue.
    private static readonly HashSet<string> XlfnFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "SEQUENCE", "SORT", "SORTBY", "UNIQUE",
        "XLOOKUP", "XMATCH",
        "LET", "LAMBDA", "REDUCE", "ISOMITTED",
        "MAP", "BYROW", "BYCOL", "SCAN", "MAKEARRAY",
        "IFS", "SWITCH",
        "MAXIFS", "MINIFS",
        "CONCAT", "TEXTJOIN",
        "STOCKHISTORY",
        "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
        "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE",
        "ISFORMULA", "SHEET", "SHEETS",
        // Engineering functions added in Excel 2013 (2007 IM* / DELTA / GESTEP
        // are bare).
        "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN",
        "BITAND", "BITOR", "BITXOR", "BITLSHIFT", "BITRSHIFT",
        // Financial functions added in Excel 2013 (other financials are bare).
        "PDURATION", "RRI",
        // Statistical / engineering distribution functions added in Excel 2010+
        // (legacy NORMDIST/GAMMADIST/CHIDIST/POISSON/ERF/… are bare).
        "NORM.DIST", "NORM.S.DIST", "NORM.INV", "NORM.S.INV",
        "GAMMA", "GAMMALN.PRECISE", "GAMMA.DIST", "GAMMA.INV",
        "CHISQ.DIST", "CHISQ.DIST.RT", "CHISQ.INV", "CHISQ.INV.RT",
        "POISSON.DIST", "EXPON.DIST", "CONFIDENCE.NORM",
        "ERF.PRECISE", "ERFC.PRECISE", "GAUSS", "PHI",
        "BETA.DIST", "BETA.INV", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T",
        "F.DIST", "F.DIST.RT", "F.INV", "F.INV.RT",
        "BINOM.DIST", "BINOM.INV", "NEGBINOM.DIST",
        "WEIBULL.DIST", "LOGNORM.DIST", "LOGNORM.INV", "HYPGEOM.DIST",
        "T.TEST", "CHISQ.TEST", "F.TEST", "Z.TEST",
        "SKEW.P", "COVARIANCE.P", "COVARIANCE.S", "QUARTILE.INC", "QUARTILE.EXC",
        "PERCENTILE.EXC", "FORECAST.LINEAR", "PERMUTATIONA",
        "TAKE", "DROP",
        "CHOOSECOLS", "CHOOSEROWS",
        "ARRAYTOTEXT", "VALUETOTEXT",
        "TOCOL", "TOROW",
        "WRAPCOLS", "WRAPROWS",
        "EXPAND",
        "HSTACK", "VSTACK",
        "ANCHORARRAY",
    };

    // Functions that need _xlfn._xlws. (dynamic-array, worksheet-only)
    private static readonly HashSet<string> XlwsFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "FILTER",
    };

    // Subset of modern functions that produce spilling dynamic-array results.
    // Their CellFormula MUST carry t="array" + ref="<cellRef>" or Excel 365
    // rejects the file (0x800A03EC). LET/LAMBDA can return arrays but are
    // commonly used for scalars too — include them when the spill metadata
    // is harmless to non-spilling outputs (Excel honors t="array" with ref=
    // pointing at the single anchor cell even when the formula resolves to
    // a single value). Source: the modern dynamic-array function catalogue.
    private static readonly HashSet<string> DynamicArrayFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "RANDARRAY",
        "XLOOKUP", "XMATCH",
        "LET", "LAMBDA",
        "MAP", "BYROW", "BYCOL", "SCAN", "MAKEARRAY",
        "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS",
        "TOCOL", "TOROW", "WRAPCOLS", "WRAPROWS", "EXPAND",
        "HSTACK", "VSTACK", "TRANSPOSE",
        "LINEST", "LOGEST", "TREND", "GROWTH",
        "TEXTSPLIT",
        "ANCHORARRAY",
    };

    /// <summary>
    /// True when the formula calls any function whose result must be written
    /// with <c>t="array"</c> + <c>ref="&lt;cellRef&gt;"</c> on the cell-level
    /// CellFormula element — the spill metadata Excel 365 requires for
    /// dynamic-array functions. Operates on the raw formula (no leading '=').
    /// Quoted string literals are skipped so a SORT call inside text doesn't
    /// trigger a false positive.
    /// </summary>
    public static bool IsDynamicArrayFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return false;
        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];
            if (c == '"')
            {
                i++;
                while (i < formula.Length)
                {
                    if (formula[i] == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"') { i += 2; continue; }
                        i++; break;
                    }
                    i++;
                }
                continue;
            }
            if (IsIdentStart(c) && (i == 0 || !IsIdentPrev(formula[i - 1])))
            {
                int start = i;
                while (i < formula.Length && IsIdentCont(formula[i])) i++;
                int j = i;
                while (j < formula.Length && formula[j] == ' ') j++;
                if (j < formula.Length && formula[j] == '(')
                {
                    var name = formula.Substring(start, i - start);
                    if (DynamicArrayFunctions.Contains(name)) return true;
                }
                continue;
            }
            i++;
        }
        return false;
    }

    // Match a bare function name (identifier followed by '('), not preceded by
    // a '.' or alphanumeric (so _xlfn.SEQUENCE and MYSEQUENCE are skipped),
    // and not inside a quoted string literal.
    private static readonly Regex FunctionCallRegex = new(
        @"(?<![A-Za-z0-9_\.])([A-Za-z_][A-Za-z0-9_]*)\s*\(",
        RegexOptions.Compiled);

    // Collect every LET / LAMBDA parameter name in the formula. For LET the
    // parameter is each odd-numbered argument except the last (name1, value1, …,
    // calc); for LAMBDA it is every argument except the last (p1, …, pN, body).
    private static HashSet<string> CollectLambdaParams(string formula)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];
            if (c == '"')   // skip string literals
            {
                i++;
                while (i < formula.Length)
                {
                    if (formula[i] == '"') { if (i + 1 < formula.Length && formula[i + 1] == '"') { i += 2; continue; } i++; break; }
                    i++;
                }
                continue;
            }
            if (IsIdentStart(c) && (i == 0 || !IsIdentPrev(formula[i - 1])))
            {
                int s = i;
                while (i < formula.Length && IsIdentCont(formula[i])) i++;
                string id = formula.Substring(s, i - s);
                int j = i;
                while (j < formula.Length && formula[j] == ' ') j++;
                bool isLet = id.Equals("LET", StringComparison.OrdinalIgnoreCase);
                bool isLambda = id.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase);
                if ((isLet || isLambda) && j < formula.Length && formula[j] == '(')
                {
                    var argSpans = TopLevelArgs(formula, j);
                    for (int k = 0; k < argSpans.Count - 1; k++)
                        if (isLambda || k % 2 == 0)
                        {
                            var nm = argSpans[k].Trim();
                            if (nm.Length > 0 && IsIdentStart(nm[0]) && nm.All(ch => IsIdentCont(ch)))
                                names.Add(nm);
                        }
                }
                continue;
            }
            i++;
        }
        return names;
    }

    // Split the parenthesised argument list starting at <paren> ('(') into the
    // top-level argument substrings (respecting nested parens and strings).
    private static List<string> TopLevelArgs(string formula, int paren)
    {
        var args = new List<string>();
        int depth = 0, i = paren, start = paren + 1;
        for (; i < formula.Length; i++)
        {
            char c = formula[i];
            if (c == '"') { i++; while (i < formula.Length && formula[i] != '"') i++; continue; }
            if (c == '(') depth++;
            else if (c == ')') { depth--; if (depth == 0) { args.Add(formula[start..i]); break; } }
            else if (c == ',' && depth == 1) { args.Add(formula[start..i]); start = i + 1; }
        }
        return args;
    }

    /// <summary>
    /// Scans a formula for an inline array constant <c>{…}</c> element that is
    /// not a literal (number, quoted string, TRUE/FALSE, error) — a function
    /// call, cell reference, name, or nested array. Returns the offending
    /// element text, or null when every array constant is well-formed. String
    /// literals are skipped so their content never counts as array syntax.
    /// </summary>
    internal static string? FindNonLiteralArrayConstant(string formula)
    {
        if (string.IsNullOrEmpty(formula) || !formula.Contains('{')) return null;
        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];
            if (c == '"') { i = SkipStringLiteral(formula, i); continue; }
            if (c == '{')
            {
                int j = i + 1, paren = 0;
                var elem = new System.Text.StringBuilder();
                while (j < formula.Length && formula[j] != '}')
                {
                    char d = formula[j];
                    if (d == '"') { int e = SkipStringLiteral(formula, j); elem.Append(formula, j, e - j); j = e; continue; }
                    if (d == '{') return "{ }";                              // nested array
                    if (d == '(') paren++;
                    else if (d == ')') paren--;
                    if ((d == ',' || d == ';') && paren == 0)                // element boundary (not inside a call)
                    {
                        if (!IsArrayLiteral(elem.ToString())) return elem.ToString().Trim();
                        elem.Clear(); j++; continue;
                    }
                    elem.Append(d); j++;
                }
                if (j >= formula.Length || !IsArrayLiteral(elem.ToString())) return elem.ToString().Trim();
                i = j + 1;
                continue;
            }
            i++;
        }
        return null;
    }

    // Index just past the closing quote of the string literal starting at s[i]
    // ('"' at i), honoring the "" escape for an embedded quote.
    private static int SkipStringLiteral(string s, int i)
    {
        i++;
        while (i < s.Length)
        {
            if (s[i] == '"')
            {
                if (i + 1 < s.Length && s[i + 1] == '"') { i += 2; continue; }
                return i + 1;
            }
            i++;
        }
        return i;
    }

    private static bool IsArrayLiteral(string e)
    {
        e = e.Trim();
        if (e.Length == 0) return false;
        if (e.Length >= 2 && e[0] == '"' && e[^1] == '"') return true;         // "string"
        if (e.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
            e.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return true; // boolean
        if (e[0] == '#') return true;                                          // #error
        return double.TryParse(e, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture, out _);         // number
    }

    /// <summary>
    /// Returns the formula with Excel 2016+ modern function names qualified
    /// with <c>_xlfn.</c> / <c>_xlfn._xlws.</c> as required by OOXML. Leaves
    /// already-qualified names, older functions, quoted string literals, and
    /// non-function identifiers untouched.
    /// </summary>
    public static string Qualify(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return formula;

        // Excel inline array constants {…} may hold only literal values.
        // A function call, cell reference, name, or nested array inside {…} makes
        // Excel report the whole workbook as corrupt (it prompts to repair on
        // open). Reject such a formula at write time so officecli never emits a
        // file Excel refuses to open.
        var badElem = FindNonLiteralArrayConstant(formula);
        if (badElem != null)
            throw new ArgumentException(
                $"Invalid array constant: '{badElem}'. Inline arrays {{...}} may contain only literal " +
                "values (numbers, \"text\", TRUE/FALSE, errors) — not functions, cell references, or " +
                "nested arrays. Reference a cell range instead (e.g. A1:A3).");

        // LET / LAMBDA parameter names are stored in the _xlpm. namespace; without
        // that prefix Excel reports the workbook as corrupt. Collect every such
        // name up front so all of its occurrences (declaration and uses) get the
        // prefix below.
        var lambdaParams = CollectLambdaParams(formula);

        // Walk the string and only rewrite identifiers outside quoted strings.
        // Excel formula strings are bounded by '"' with '""' as an escape.
        var sb = new System.Text.StringBuilder(formula.Length + 32);
        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];
            if (c == '"')
            {
                // Copy the entire string literal verbatim.
                sb.Append(c);
                i++;
                while (i < formula.Length)
                {
                    sb.Append(formula[i]);
                    if (formula[i] == '"')
                    {
                        // escaped "" → consume both, stay in string
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            sb.Append('"');
                            i += 2;
                            continue;
                        }
                        i++;
                        break;
                    }
                    i++;
                }
                continue;
            }

            // Outside a string: scan for an identifier-call.
            // Use regex-on-substring is awkward; instead detect manually.
            if (IsIdentStart(c) && (i == 0 || !IsIdentPrev(formula[i - 1])))
            {
                int start = i;
                while (i < formula.Length && IsIdentCont(formula[i])) i++;
                var name = formula.Substring(start, i - start);
                // A LET/LAMBDA parameter (incl. when used to call a stored lambda,
                // e.g. sq(7)) takes the _xlpm. prefix and shadows function names.
                if (lambdaParams.Contains(name))
                {
                    sb.Append("_xlpm.").Append(name);
                    continue;
                }
                // Skip whitespace then check for '('
                int j = i;
                while (j < formula.Length && formula[j] == ' ') j++;
                if (j < formula.Length && formula[j] == '(')
                {
                    if (XlwsFunctions.Contains(name))
                        sb.Append("_xlfn._xlws.").Append(name);
                    else if (XlfnFunctions.Contains(name))
                        sb.Append("_xlfn.").Append(name);
                    else
                        sb.Append(name);
                }
                else
                {
                    sb.Append(name);
                }
                continue;
            }

            sb.Append(c);
            i++;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Inverse of <see cref="Qualify"/> for readback: strips the
    /// <c>_xlfn.</c> / <c>_xlfn._xlws.</c> prefix so users see canonical
    /// function names instead of the OOXML-internal namespaced form.
    /// </summary>
    public static string Unqualify(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return formula;
        // Longer prefix first so we don't leave _xlws. stragglers.
        var s = formula.Replace("_xlfn._xlws.", "", StringComparison.Ordinal);
        s = s.Replace("_xlfn.", "", StringComparison.Ordinal);
        // LET / LAMBDA parameter namespace — strip so the evaluator and the user
        // see bare names (e.g. _xlpm.x → x).
        s = s.Replace("_xlpm.", "", StringComparison.Ordinal);
        return s;
    }

    /// <summary>
    /// Auto-quote unquoted sheet-name references in a formula when the sheet
    /// name needs single-quotes per Excel rules — i.e. starts with a digit,
    /// or contains a space, or contains any of <c>[ ] : / \ ? *</c> /
    /// punctuation. Already-quoted (e.g. <c>'1stQ'!A1</c>) refs are kept as-is.
    /// String literals are skipped.
    /// </summary>
    public static string AutoQuoteSheetRefs(string formula)
    {
        if (string.IsNullOrEmpty(formula) || !formula.Contains('!')) return formula;

        var sb = new System.Text.StringBuilder(formula.Length + 16);
        int i = 0;
        while (i < formula.Length)
        {
            char c = formula[i];
            // Skip string literals verbatim
            if (c == '"')
            {
                sb.Append(c);
                i++;
                while (i < formula.Length)
                {
                    sb.Append(formula[i]);
                    if (formula[i] == '"')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '"')
                        {
                            sb.Append('"'); i += 2; continue;
                        }
                        i++;
                        break;
                    }
                    i++;
                }
                continue;
            }
            // Skip already-quoted sheet refs verbatim
            if (c == '\'')
            {
                sb.Append(c);
                i++;
                while (i < formula.Length)
                {
                    sb.Append(formula[i]);
                    if (formula[i] == '\'')
                    {
                        if (i + 1 < formula.Length && formula[i + 1] == '\'')
                        {
                            sb.Append('\''); i += 2; continue;
                        }
                        i++;
                        break;
                    }
                    i++;
                }
                continue;
            }

            // Detect bare sheet-name token followed by '!'. A sheet name token
            // here is a maximal run of [A-Za-z0-9_.] possibly preceded only by
            // a non-identifier char.
            if ((char.IsLetterOrDigit(c) || c == '_') &&
                (i == 0 || !IsIdentPrev(formula[i - 1])))
            {
                int start = i;
                // Greedy scan: include identifier chars and embedded spaces, as long
                // as the run ultimately terminates at '!'. A bare sheet name with a
                // space (e.g. `My Sheet!A1`) must be quoted as a whole, not split
                // across the space.
                int j = i;
                while (j < formula.Length && (char.IsLetterOrDigit(formula[j]) || formula[j] == '_' || formula[j] == '.' || formula[j] == ' '))
                    j++;
                // Trim trailing spaces from the candidate name; they can't be part of
                // a sheet ref unless followed by more name chars then '!'.
                int end = j;
                while (end > start && formula[end - 1] == ' ') end--;
                if (end < formula.Length && formula[end] == '!' && end > start)
                {
                    var name = formula.Substring(start, end - start);
                    if (SheetNameNeedsQuoting(name))
                        sb.Append('\'').Append(name).Append('\'');
                    else
                        sb.Append(name);
                    i = end;
                    continue;
                }
                // No '!' terminator: only consume the leading non-space identifier
                // run (preserve old behavior for plain tokens / function calls).
                int k = i;
                while (k < formula.Length && (char.IsLetterOrDigit(formula[k]) || formula[k] == '_' || formula[k] == '.'))
                    k++;
                sb.Append(formula, start, k - start);
                i = k;
                continue;
            }

            sb.Append(c);
            i++;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Quote a worksheet name for use in a formula or defined-name reference
    /// (the part before <c>!</c>) when Excel requires it — i.e. the name starts
    /// with a digit or contains any character outside <c>[A-Za-z0-9_]</c> (space,
    /// hyphen, punctuation, …). Embedded single quotes are doubled per the OOXML
    /// rule. Names that don't need quoting are returned unchanged. Without this,
    /// a defined name like <c>2-Print!$A$1</c> is invalid and Excel refuses to
    /// open the entire workbook (0x800A03EC).
    /// </summary>
    public static string QuoteSheetNameForRef(string name)
    {
        if (string.IsNullOrEmpty(name)) return name;
        if (!SheetNameNeedsQuoting(name)) return name;
        return "'" + name.Replace("'", "''") + "'";
    }

    private static bool SheetNameNeedsQuoting(string name)
    {
        if (string.IsNullOrEmpty(name)) return false;
        // Starts with digit
        if (char.IsDigit(name[0])) return true;
        // Punctuation/special chars: space, [ ] : / \ ? *, plus '.','-','+',etc.
        foreach (var ch in name)
        {
            if (char.IsLetterOrDigit(ch) || ch == '_') continue;
            return true;
        }
        return false;
    }

    private static bool IsIdentStart(char c) => char.IsLetter(c) || c == '_';
    private static bool IsIdentCont(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.';
    // Prev char that would mean we're in the middle of an existing identifier
    // (incl. already-qualified `_xlfn.NAME`).
    private static bool IsIdentPrev(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.';
}
