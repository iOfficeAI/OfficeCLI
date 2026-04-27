// Copyright 2025 OfficeCli (officecli.ai)
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
        "LET", "LAMBDA",
        "IFS", "SWITCH",
        "MAXIFS", "MINIFS",
        "CONCAT", "TEXTJOIN",
        "STOCKHISTORY",
        "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT",
        "TAKE", "DROP",
        "CHOOSECOLS", "CHOOSEROWS",
        "ARRAYTOTEXT", "VALUETOTEXT",
        "TOCOL", "TOROW",
        "WRAPCOLS", "WRAPROWS",
        "EXPAND",
        "ANCHORARRAY",
    };

    // Functions that need _xlfn._xlws. (dynamic-array, worksheet-only)
    private static readonly HashSet<string> XlwsFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "FILTER",
    };

    // Match a bare function name (identifier followed by '('), not preceded by
    // a '.' or alphanumeric (so _xlfn.SEQUENCE and MYSEQUENCE are skipped),
    // and not inside a quoted string literal.
    private static readonly Regex FunctionCallRegex = new(
        @"(?<![A-Za-z0-9_\.])([A-Za-z_][A-Za-z0-9_]*)\s*\(",
        RegexOptions.Compiled);

    /// <summary>
    /// Returns the formula with Excel 2016+ modern function names qualified
    /// with <c>_xlfn.</c> / <c>_xlfn._xlws.</c> as required by OOXML. Leaves
    /// already-qualified names, older functions, quoted string literals, and
    /// non-function identifiers untouched.
    /// </summary>
    public static string Qualify(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return formula;

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
                // Skip whitespace then check for '('
                int j = i;
                while (j < formula.Length && formula[j] == ' ') j++;
                if (j < formula.Length && formula[j] == '(')
                {
                    var name = formula.Substring(start, i - start);
                    if (XlwsFunctions.Contains(name))
                        sb.Append("_xlfn._xlws.").Append(name);
                    else if (XlfnFunctions.Contains(name))
                        sb.Append("_xlfn.").Append(name);
                    else
                        sb.Append(name);
                }
                else
                {
                    sb.Append(formula, start, i - start);
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
