// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // A1 sqref shape: one or more space-separated single-cell or range
    // references (relative or absolute, with optional sheet-bare row/col
    // letters). Reject everything else up front so a conditional-formatting
    // rule cannot land an `INVALID!REF` literal in the sheet — Excel
    // refuses to open the file in that state.
    private static readonly System.Text.RegularExpressions.Regex SqrefShape =
        new(@"^\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?(\s+\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$",
            System.Text.RegularExpressions.RegexOptions.Compiled
            | System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    // Whole-column (A:A, B:XFD) and whole-row (1:1, 2:10) tokens are legal
    // sqref members — dump reads them from real files, so add/replay must
    // accept them too (a column-wide CF rule could not be round-tripped).
    private static readonly System.Text.RegularExpressions.Regex SqrefWholeToken =
        new(@"^(\$?[A-Z]+:\$?[A-Z]+|\$?[0-9]+:\$?[0-9]+)$",
            System.Text.RegularExpressions.RegexOptions.Compiled
            | System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    internal static string ValidateSqref(string value, string field)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"Invalid {field} '{value}': empty A1 range.");
        var trimmed = value.Trim();
        var ok = trimmed
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .All(tok => SqrefShape.IsMatch(tok) || SqrefWholeToken.IsMatch(tok));
        if (!ok)
            throw new ArgumentException(
                $"Invalid {field} '{value}': expected an A1 reference (e.g. 'A1', 'A1:D10', 'A:A', '1:3', 'A1 B2:C5').");
        // Shape-valid tokens can still point outside Excel's grid: sqref="A0"
        // passed here, saved fine, and real Excel refused the whole file
        // (0x800A03EC) — the same out-of-grid family the drawing-anchor parser
        // rejects. Bounds-check every cell/row/column component.
        foreach (var tok in trimmed.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            foreach (System.Text.RegularExpressions.Match cm in
                System.Text.RegularExpressions.Regex.Matches(tok, @"\$?([A-Z]+)?\$?([0-9]+)?",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            {
                if (cm.Length == 0) continue;
                if (cm.Groups[1].Success && cm.Groups[1].Value.Length > 0)
                {
                    var colIdx = ColumnNameToIndex(cm.Groups[1].Value.ToUpperInvariant());
                    if (colIdx < 1 || colIdx > 16384)
                        throw new ArgumentException(
                            $"Invalid {field} '{value}': column '{cm.Groups[1].Value}' is outside Excel's grid (A..XFD).");
                }
                if (cm.Groups[2].Success && cm.Groups[2].Value.Length > 0)
                {
                    if (!long.TryParse(cm.Groups[2].Value, out var rowNum) || rowNum < 1 || rowNum > 1048576)
                        throw new ArgumentException(
                            $"Invalid {field} '{value}': row '{cm.Groups[2].Value}' is outside Excel's grid (1..1048576).");
                }
            }
        }
        // Canonicalize inverted tokens (F5:D3 → D3:F5, per axis) — merge
        // rejects them and table normalizes them, but CF/DV wrote them
        // verbatim, leaving a non-canonical sqref whose behavior in real
        // Excel is undefined. Same convention as the drawing-anchor and
        // table-range normalization.
        var normTokens = trimmed.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(tok =>
            {
                var cm = System.Text.RegularExpressions.Regex.Match(tok,
                    @"^(\$?)([A-Z]+)(\$?)([0-9]+):(\$?)([A-Z]+)(\$?)([0-9]+)$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (cm.Success)
                {
                    var c1 = ColumnNameToIndex(cm.Groups[2].Value.ToUpperInvariant());
                    var c2 = ColumnNameToIndex(cm.Groups[6].Value.ToUpperInvariant());
                    var r1 = long.Parse(cm.Groups[4].Value);
                    var r2 = long.Parse(cm.Groups[8].Value);
                    var colA = c1 <= c2 ? cm.Groups[2].Value : cm.Groups[6].Value;
                    var colB = c1 <= c2 ? cm.Groups[6].Value : cm.Groups[2].Value;
                    var rowA = Math.Min(r1, r2);
                    var rowB = Math.Max(r1, r2);
                    return $"{colA}{rowA}:{colB}{rowB}";
                }
                var wm = System.Text.RegularExpressions.Regex.Match(tok,
                    @"^\$?([A-Z]+)\$?:\$?([A-Z]+)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (wm.Success
                    && ColumnNameToIndex(wm.Groups[1].Value.ToUpperInvariant())
                        > ColumnNameToIndex(wm.Groups[2].Value.ToUpperInvariant()))
                    return $"{wm.Groups[2].Value}:{wm.Groups[1].Value}";
                var rm = System.Text.RegularExpressions.Regex.Match(tok, @"^\$?([0-9]+)\$?:\$?([0-9]+)$");
                if (rm.Success && long.Parse(rm.Groups[1].Value) > long.Parse(rm.Groups[2].Value))
                    return $"{rm.Groups[2].Value}:{rm.Groups[1].Value}";
                return tok;
            });
        return string.Join(" ", normTokens);
    }

    /// <summary>
    /// Scan a formula text for plain A1-style cell references and validate
    /// each one against Excel's row/column limits (1-1048576, A-XFD). Skips
    /// quoted strings, sheet-qualified refs (delegated to RejectCrossWorkbookFormula
    /// + sheet existence checks), function names, and structured table refs.
    /// Throws ArgumentException on the first out-of-range reference. (B15)
    /// </summary>
    /// <summary>
    /// Recognise an Excel error sentinel in a cell display value. Covers the
    /// seven standard ECMA-376 codes (#NULL!, #DIV/0!, #VALUE!, #REF!,
    /// #NAME?, #NUM!, #N/A) and modern additions (#SPILL!, #CALC!, #FIELD!,
    /// #BLOCKED!, #CONNECT!, #GETTING_DATA, #UNKNOWN!, …) via structural
    /// shape match. Used by every consumer that wants to bucket formula
    /// errors: view issues subtype routing, view stats counter, view
    /// outline warning. Centralised so the three readers cannot drift on
    /// which codes count as errors. <paramref name="value"/> is the cell's
    /// display value (cachedValue text or evaluator result); the synthetic
    /// <c>#OCLI_NOTEVAL!</c> sentinel is excluded so unevaluated formulas
    /// route to their own subtype.
    /// </summary>
    internal static bool IsExcelErrorValue(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        if (value == "#OCLI_NOTEVAL!") return false;
        if (!value.StartsWith('#') || value.Length < 2) return false;
        // Every Excel error sentinel starts with `#` followed by an
        // uppercase letter. Covers the seven ECMA-376 codes, the modern
        // additions (#SPILL!, #CALC!, #FIELD!, #BLOCKED!, #CONNECT!,
        // #UNKNOWN!), and async-fetch sentinels (#GETTING_DATA) which
        // lack the trailing `!`. Intentionally lenient — there is no
        // OOXML BNF for the error namespace and new error codes have
        // been added over time. The trade-off is that `#FOO` would also match
        // here; the alternative (closed-set whitelist) would break the
        // moment a new error code lands.
        char c = value[1];
        return c >= 'A' && c <= 'Z';
    }

    /// <summary>
    /// Cell-aware overload. Recognises an Excel error in either the cell's
    /// display value (delegates to the string overload) or via the explicit
    /// <c>t="e"</c> data type. The DataType check covers writers that tag
    /// the cell type but leave the cached <c>&lt;v&gt;</c> in some unusual
    /// form (or empty). Shared by <c>view stats</c> and <c>view outline</c>
    /// for the <c>errorCells</c> count; <c>view issues</c> calls the same
    /// helper but additionally requires <c>cell.CellFormula != null</c>
    /// because the <c>formula_eval_error</c> subtype is, by definition,
    /// scoped to formula cells (a literal <c>#VALUE!</c> typed by a user
    /// into a non-formula cell is counted in stats/outline but is not a
    /// formula evaluation failure and has no matching subtype).
    /// </summary>
    internal static bool IsExcelErrorValue(Cell cell, string? displayValue)
    {
        if (IsExcelErrorValue(displayValue)) return true;
        return cell.DataType?.Value == CellValues.Error
            && displayValue != "#OCLI_NOTEVAL!";
    }

    /// <summary>
    /// Reject R1C1-style references (R2C3, RC[1], R[-1]C) in a formula string.
    /// The OOXML &lt;f&gt; element is A1-only; writing R1C1 verbatim makes real
    /// Excel refuse the file (0x800A03EC) while schema validation stays green.
    /// Shared by cell formulas and conditional-formatting formulas. Does NOT
    /// do grid-bounds checking (out-of-range A1 refs are tolerated by Excel in
    /// CF formulas — only cell-formula validation adds the bounds check).
    /// </summary>
    internal static void ValidateNoR1C1Reference(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return;
        var stripped = StripFormulaStringLiterals(formula.TrimStart('='));
        // Only unambiguous forms are rejected: bracketed offsets, or
        // R<digits>C<digits> (never a legal A1 token or name). "RC1"/"RC"
        // stay accepted — RC is a real A1 column / legal name.
        if (System.Text.RegularExpressions.Regex.IsMatch(stripped,
                @"(?<![A-Za-z0-9_$])(R\[-?\d+\]C(\[-?\d+\]|\d+)?|R\d*C\[-?\d+\]|R\d+C\d+)(?![A-Za-z0-9_])"))
            throw new ArgumentException(
                "Formula contains an R1C1-style reference (e.g. R2C3, RC[1]). OOXML stores formulas in A1 notation only — rewrite the reference in A1 form (e.g. C2, $B$3).");
    }

    // Blank out "..." string literals so cell-like substrings inside them
    // don't trigger reference validation.
    private static string StripFormulaStringLiterals(string trimmed)
    {
        var sb = new System.Text.StringBuilder(trimmed.Length);
        bool inStr = false;
        for (int i = 0; i < trimmed.Length; i++)
        {
            char c = trimmed[i];
            if (c == '"')
            {
                inStr = !inStr;
                sb.Append(' ');
                continue;
            }
            sb.Append(inStr ? ' ' : c);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Reject a formula in which any single function call has more than 255
    /// arguments — Excel's hard per-function limit. Over it, the file is
    /// schema-valid but real Excel refuses to open it (0x800A03EC). Counts
    /// top-level commas per parenthesis nesting level (a comma is attributed to
    /// its innermost enclosing "()"); string literals are skipped, and "{}"
    /// array-constant commas are not function args. 255 args (254 commas) is
    /// fine; 256 args (255 commas) throws.
    /// </summary>
    internal static void ValidateFormulaArgCount(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return;
        var f = formula.TrimStart('=');
        var commaStack = new Stack<int>();
        bool inStr = false;
        int arrayDepth = 0;
        for (int i = 0; i < f.Length; i++)
        {
            char c = f[i];
            if (c == '"')
            {
                if (inStr && i + 1 < f.Length && f[i + 1] == '"') { i++; continue; }
                inStr = !inStr;
                continue;
            }
            if (inStr) continue;
            switch (c)
            {
                case '{': arrayDepth++; break;
                case '}': if (arrayDepth > 0) arrayDepth--; break;
                case '(': commaStack.Push(0); break;
                case ')':
                    if (commaStack.Count > 0)
                    {
                        var commas = commaStack.Pop();
                        // args = commas + 1; reject 256+ args (>=255 commas).
                        if (commas >= 255)
                            throw new ArgumentException(
                                $"Formula has a function call with {commas + 1} arguments; Excel's limit is 255 per function. "
                                + "Split the call or reference a range instead.");
                    }
                    break;
                case ',':
                    if (arrayDepth == 0 && commaStack.Count > 0)
                        commaStack.Push(commaStack.Pop() + 1);
                    break;
            }
        }
    }

    // Excel's hard ceiling on the character length of a formula / defined-name
    // refersTo / conditional-format expression. Content beyond this is silently
    // accepted, persisted, and makes real Excel refuse the file (0x800A03EC).
    internal const int MaxFormulaLength = 8192;

    /// <summary>
    /// Reject a formula whose length exceeds Excel's 8192-character ceiling.
    /// Shared by cell formulas, defined-name refs, and conditional-format
    /// expressions so the limit is enforced identically everywhere.
    /// </summary>
    internal static void ValidateFormulaLength(string? formula, string context = "formula")
    {
        if (formula == null) return;
        var content = formula.TrimStart('=');
        if (content.Length > MaxFormulaLength)
            throw new ArgumentException(
                $"{context} is {content.Length} characters; Excel's limit is {MaxFormulaLength} per formula. " +
                "A longer expression makes Excel refuse to open the file.");
    }

    internal static void ValidateFormulaCellRefs(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return;
        var trimmed = formula.TrimStart('=');
        var stripped = StripFormulaStringLiterals(trimmed);

        // Formula-length ceiling (8192) — checked before the ref scan so an
        // oversized expression fails with a clear message instead of Excel's
        // 0x800A03EC after the fact.
        ValidateFormulaLength(formula);
        // R1C1-style references make real Excel refuse the file.
        ValidateNoR1C1Reference(formula);
        // Excel caps a function call at 255 arguments; a 256-arg call passes
        // schema validation but makes real Excel refuse the file (0x800A03EC).
        ValidateFormulaArgCount(formula);
        // Match A1-style refs: optional $ + 1-3 letters + optional $ + 1-8 digits.
        // (Excel's row ceiling 1048576 is 7-digit, but 8-digit numbers like
        // A10000000 must still be caught so they're rejected with the clean
        // "out-of-range" error rather than slipping through validation.)
        // Avoid matching inside an identifier (e.g. "FOO1") via a leading
        // boundary that requires either start-of-string or a non-letter.
        var rx = new System.Text.RegularExpressions.Regex(
            @"(?<![A-Za-z_])\$?([A-Za-z]{1,3})\$?([0-9]{1,8})\b");
        foreach (System.Text.RegularExpressions.Match m in rx.Matches(stripped))
        {
            var col = m.Groups[1].Value.ToUpperInvariant();
            if (!long.TryParse(m.Groups[2].Value, out var row)) continue;
            // Column index check: ColumnNameToIndex would throw on overflow,
            // but we want a clean validation message. Compute manually.
            int colIdx = 0;
            foreach (var ch in col) colIdx = colIdx * 26 + (ch - 'A' + 1);
            if (colIdx < 1 || colIdx > 16384 || row < 1 || row > 1048576)
            {
                throw new ArgumentException(
                    $"Formula contains out-of-range cell reference '{m.Value}'. " +
                    "Excel limits: rows 1-1048576, columns A-XFD.");
            }
        }
    }

    internal static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Invalid sheet name: name cannot be empty or whitespace.");
        if (name.Length > 31)
            throw new ArgumentException(
                $"Invalid sheet name '{name}': length {name.Length} exceeds Excel's 31-char limit.");
        var forbidden = new[] { '\\', '/', '?', '*', ':', '[', ']' };
        var hit = name.IndexOfAny(forbidden);
        if (hit >= 0)
            throw new ArgumentException(
                $"Invalid sheet name '{name}': contains forbidden character '{name[hit]}'. Excel rejects any of: \\ / ? * : [ ]");
        if (name.StartsWith('\'') || name.EndsWith('\''))
            throw new ArgumentException(
                $"Invalid sheet name '{name}': cannot start or end with an apostrophe (').");
        if (name.Equals("History", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "Invalid sheet name 'History': reserved by Excel for the change-history sheet.");
    }

    /// <summary>
    /// R35-3: cross-workbook cell formulas like "=[Other.xlsx]Sheet1!A1" or
    /// "=[1]Sheet1!A1" need an externalLinks part to resolve. Without one,
    /// Excel opens the file but the formula shows #REF!. Reject up-front
    /// rather than silently persist a broken formula.
    /// CONSISTENCY(cross-workbook-ref): mirrors the namedrange refersTo
    /// guard in ExcelHandler.Add.Tables.cs (R27-1).
    /// </summary>
    internal static void RejectCrossWorkbookFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return;
        var trimmed = formula.TrimStart('=', ' ', '\t');
        // CONSISTENCY(cross-workbook-vs-structured-ref): the older `^\[` guard
        // also matched OOXML structured table references like `[@Price]` and
        // `[Price]*[Qty]`, falsely rejecting valid Excel-365 formulas. Real
        // cross-workbook refs have one of two shapes:
        //   - numeric workbook index:  `[1]Sheet1!A1`        → `[<digits>]`
        //   - filename + extension:    `[Other.xlsx]Sheet!A1` → `[<name>.xls(x|m|b)?]`
        // Both forms are followed by a sheet reference (`Sheet!...`), but the
        // bracket payload alone is enough to disambiguate from `[@Col]` /
        // `[Col]` structured refs (which contain `@`, alphabetics without an
        // extension, or `:`).
        if (System.Text.RegularExpressions.Regex.IsMatch(trimmed,
                @"^\[(\d+|[^\]]*\.xls[xbm]?)\]",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            throw new ArgumentException(
                $"Cross-workbook references like '{formula}' require an externalLinks part which officecli doesn't expose; use raw-set for this case");
    }

    // Normalize user-supplied data-validation formula values so Excel accepts
    // them. `type=list` auto-quotes bare lists. `type=time` accepts HH:MM /
    // HH:MM:SS and converts to the Excel time serial fraction. `type=date`
    // accepts YYYY-MM-DD and converts to the Excel date serial. `type=custom`
    // strips a leading '=' since OOXML `<x:formula1>` expects the formula body
    // without one.
    internal static string NormalizeValidationFormula(string value, DataValidationValues? type)
    {
        if (string.IsNullOrEmpty(value)) return value;
        if (type == DataValidationValues.List)
        {
            // list: wrap bare "a,b,c" in quotes; leave cell/range refs and
            // already-quoted literals alone. V1: a leading `=` signals a
            // formula-ref (e.g. `=VOpts`, `=$Z$1:$Z$5`) — strip the `=`
            // (OOXML `<x:formula1>` expects the body without one) and
            // pass through without quoting.
            if (value.StartsWith("="))
                return value.Substring(1);
            if (value.StartsWith("\"") || value.Contains("!") || value.Contains(":"))
                return value;
            if (value.Contains(','))
                return $"\"{value}\"";
            // A bare unquoted value is valid as a list source if it's a legal
            // defined-name reference (name token) OR a formula (contains '(',
            // e.g. INDIRECT(B2)/OFFSET(...)). Anything else — e.g. "hello
            // world" — is neither a name, a ref, nor a quoted literal, so Excel
            // refuses the file (0x800A03EC); treat it as a literal single-item
            // list and quote it. Formulas must NOT be quoted (that would turn
            // them into a literal string and also block ref shifting).
            if (!value.Contains('(')
                && !System.Text.RegularExpressions.Regex.IsMatch(value, @"^[A-Za-z_\\][A-Za-z0-9_.]*$"))
                return $"\"{value}\"";
            return value;
        }
        if (type == DataValidationValues.Time)
        {
            var m = System.Text.RegularExpressions.Regex.Match(value.Trim(), @"^(\d{1,2}):(\d{2})(?::(\d{2}))?$");
            if (m.Success)
            {
                var h = int.Parse(m.Groups[1].Value);
                var mn = int.Parse(m.Groups[2].Value);
                var s = m.Groups[3].Success ? int.Parse(m.Groups[3].Value) : 0;
                var frac = (h * 3600 + mn * 60 + s) / 86400.0;
                return frac.ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        if (type == DataValidationValues.Date)
        {
            if (System.DateTime.TryParseExact(value.Trim(), "yyyy-MM-dd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var dt))
            {
                // Excel date serial: days since 1899-12-30 (accounts for the
                // 1900 leap bug baseline).
                var epoch = new System.DateTime(1899, 12, 30);
                return ((int)(dt - epoch).TotalDays).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        if (type == DataValidationValues.Custom)
        {
            if (value.StartsWith("="))
                return value.Substring(1);
        }
        // For non-list numeric/date/text types, formula1/formula2 must be a
        // number, date/time (handled above), cell/range ref, or a formula — a
        // bare value containing whitespace (e.g. "hello world") is invalid
        // OOXML formula syntax and makes real Excel refuse the file
        // (0x800A03EC). Reject it up front. (Quoted literals and refs pass.)
        if (type != DataValidationValues.Custom
            && value.Any(char.IsWhiteSpace)
            && !value.StartsWith("\"") && !value.StartsWith("=")
            && !value.Contains('!') && !value.Contains('('))
            throw new ArgumentException(
                $"validation formula '{value}' is not valid for this validation type: " +
                "expected a number, date, cell reference, or formula (a bare value with spaces is not valid formula syntax).");
        return value;
    }

    // CONSISTENCY(merge-overlap): centralize the "insert one MergeCell"
    // policy. Excel rejects overlapping <mergeCell> entries with a
    // "found a problem" repair dialog, but the OOXML SDK happily
    // appends them. Mirrors the T4 overlap-throws pattern used by
    // tables and AutoFilter+table.
    // - Exact-match ref: no-op (idempotent re-Add stays consistent
    //   with prior dedup behavior).
    // - Geometric overlap with a non-identical range: throw.
    // - Otherwise: append.
    private static readonly System.Text.RegularExpressions.Regex SingleMergeRefPattern =
        new(@"^[A-Z]+[0-9]+(:[A-Z]+[0-9]+)?$",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    // CONSISTENCY(merge-comma): callers should run this BEFORE creating an
    // empty <mergeCells> container, so a rejected ref doesn't leave a
    // schema-invalid empty container in the saved file.
    private static void ValidateMergeRefLiteral(string newRangeRef)
    {
        var refUpper = newRangeRef.ToUpperInvariant();
        if (refUpper.Contains(','))
            throw new ArgumentException(
                $"Invalid merge ref '{newRangeRef}': path is a single-target locator (no comma). " +
                $"Move ranges to a prop value, e.g. `set ... '/Sheet1' --prop merge={newRangeRef}`.");
        if (!SingleMergeRefPattern.IsMatch(refUpper))
            throw new ArgumentException(
                $"Invalid merge ref '{newRangeRef}': must be a single A1 cell (e.g. 'B2') or A1:B2 range (e.g. 'B4:E4').");
        // CONSISTENCY(merge-orientation): the ref must read top-left to
        // bottom-right. Z1:A1 / A10:A1 / B2:A1 (any reversed orientation)
        // were silently accepted; Excel itself only writes the canonical
        // form, so callers passing a reversed pair almost certainly typo'd.
        // Reject with a hint to swap, mirroring the orientation guard the
        // sheetShift normalizer applies after the fact (ExcelHandler.Set.cs
        // L1918) and matching how other range-bearing props (validation,
        // table, autofilter) demand canonical orientation up front.
        var colonIdx = refUpper.IndexOf(':');
        if (colonIdx > 0)
        {
            var lhs = refUpper.Substring(0, colonIdx);
            var rhs = refUpper.Substring(colonIdx + 1);
            try
            {
                var (lCol, lRow) = ParseCellReference(lhs);
                var (rCol, rRow) = ParseCellReference(rhs);
                int lColIdx = ColumnNameToIndex(lCol);
                int rColIdx = ColumnNameToIndex(rCol);
                if (lColIdx > rColIdx || lRow > rRow)
                {
                    throw new ArgumentException(
                        $"Invalid merge ref '{newRangeRef}': range must read top-left to bottom-right. " +
                        $"Pass the canonical orientation (e.g. 'A1:B2', not 'B2:A1').");
                }
            }
            catch (ArgumentException) { throw; }
            catch { /* parse failure already handled by SingleMergeRefPattern above */ }
        }
    }

    /// <summary>
    /// Scan a formula body for Sheet-qualified refs (bare `Sheet1!A1`
    /// or quoted `'My Data'!A1`) and return true if any referenced sheet
    /// name does not exist in the current workbook. Used to suppress the
    /// evaluator-based cachedValue fallback when cross-sheet refs point at
    /// a removed sheet — Real Excel shows `#REF!` there; we should not
    /// invent a "0".
    /// </summary>
    private bool FormulaReferencesMissingSheet(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return false;
        var wb = _doc.WorkbookPart?.Workbook;
        if (wb == null) return false;
        var names = new HashSet<string>(
            wb.Descendants<Sheet>().Select(s => s.Name?.Value ?? "").Where(n => n.Length > 0),
            StringComparer.OrdinalIgnoreCase);

        // Strip out double-quoted string literals first — formulas can carry
        // arbitrary text in `"..."` (e.g. `="World!"` or
        // `=INDIRECT("Foo!B2")`) and an unguarded scan would mis-flag the
        // contents as a sheet reference. OOXML escapes inner double quotes
        // by doubling (`""`), so the literal-content pattern matches either
        // `""` or any non-`"` character.
        var scan = System.Text.RegularExpressions.Regex.Replace(
            formula, @"""(""""|[^""])*""", "\"\"");

        // Quoted form: '...'! — inner single quotes escaped as ''
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(scan, @"'((?:[^']|'')+)'!"))
        {
            var name = m.Groups[1].Value.Replace("''", "'");
            if (!names.Contains(name)) return true;
        }
        // Bare form: Name! — letters/digits/underscore/period (Excel allows these
        // unquoted). '#' in the lookbehind excludes the #REF! ERROR literal
        // (e.g. `=#REF!*2` after a row/col delete): that is a cell-level error
        // marker, not a reference to a sheet named "REF" — classifying it as a
        // missing sheet sent users hunting for a sheet that never existed.
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(scan, @"(?<![A-Za-z0-9_'.#])([A-Za-z_][A-Za-z0-9_.]*)!"))
        {
            if (!names.Contains(m.Groups[1].Value)) return true;
        }
        return false;
    }

    // R13-1: Excel rejects cell values longer than 32767 chars (2^15 - 1) with
    // 0x800A03EC on save/open. Reject at write time with a clear error rather
    // than silently writing a file Excel will refuse to open.
    internal const int MaxCellTextLength = 32767;

    internal static void EnsureCellValueLength(string? value, string? cellRef = null)
    {
        if (value == null) return;
        if (value.Length > MaxCellTextLength)
        {
            var where = string.IsNullOrEmpty(cellRef) ? "" : $" at {cellRef}";
            throw new ArgumentException(
                $"Cell value{where} exceeds Excel's {MaxCellTextLength}-character limit (got {value.Length})");
        }
        // XML-illegal control chars / lone surrogates: without this the value
        // enters the in-memory DOM fine and only fails at close-time save —
        // "save failed during shutdown", leaving sheetData empty on disk
        // (total data loss). Same guard Word text paths already use.
        OfficeCli.Core.ParseHelpers.ValidateXmlText(value,
            string.IsNullOrEmpty(cellRef) ? "cell value" : $"cell value at {cellRef}");
    }

    // Numeric literal forms Excel's <v> parser accepts. double.TryParse is
    // far more lenient (leading '+', thousands separators, padding, currency
    // NumberStyles) — writing such text verbatim into an untyped (numeric)
    // <v> makes real Excel refuse the file (0x800A03EC) even though
    // schema validation stays green.
    private static readonly System.Text.RegularExpressions.Regex CanonicalNumericLiteral =
        new(@"^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$", System.Text.RegularExpressions.RegexOptions.Compiled);

    internal static bool IsCanonicalNumericText(string text) => CanonicalNumericLiteral.IsMatch(text);

    /// <summary>Shape-check a sparkline data range ("A1:E1" or
    /// "Sheet1!A1:E1", whole rows/cols allowed). Arbitrary strings written
    /// into &lt;xne:f&gt; make real Excel refuse the file while schema
    /// validation stays green.</summary>
    internal static void ValidateSparklineRange(string range)
    {
        var r = (range ?? "").Trim();
        var bang = r.LastIndexOf('!');
        if (bang >= 0) r = r[(bang + 1)..];
        r = r.Replace("$", "");
        var ok = r.Length > 0 && r.Split(':').All(tok =>
            System.Text.RegularExpressions.Regex.IsMatch(tok.Trim(), @"^([A-Za-z]{1,3}\d+|[A-Za-z]{1,3}|\d+)$"));
        if (!ok)
            throw new ArgumentException(
                $"Invalid sparkline range '{range}'. Expected an A1 range like A1:E1 (optionally sheet-qualified).");
        // Grid-bounds check, same family as ValidateSqref: shape-valid B0 or
        // A99999999 landed verbatim in <xne:f>. Excel tolerates rather than
        // rejects these, but the sparkline is silently broken — tighten to
        // the severity every other cell-ref entry point now enforces.
        foreach (var tok in r.Split(':'))
        {
            var tm = System.Text.RegularExpressions.Regex.Match(tok.Trim(),
                @"^([A-Za-z]{1,3})?(\d+)?$");
            if (tm.Groups[1].Success && tm.Groups[1].Value.Length > 0)
            {
                var colIdx = ColumnNameToIndex(tm.Groups[1].Value.ToUpperInvariant());
                if (colIdx < 1 || colIdx > 16384)
                    throw new ArgumentException(
                        $"Invalid sparkline range '{range}': column '{tm.Groups[1].Value}' is outside Excel's grid (A..XFD).");
            }
            if (tm.Groups[2].Success && tm.Groups[2].Value.Length > 0
                && (!long.TryParse(tm.Groups[2].Value, out var rowN) || rowN < 1 || rowN > 1048576))
                throw new ArgumentException(
                    $"Invalid sparkline range '{range}': row '{tm.Groups[2].Value}' is outside Excel's grid (1..1048576).");
        }
    }

    /// <summary>Sanity-check a defined-name body. Full formula validation is
    /// out of scope, but a sheet-qualified reference must name an existing
    /// sheet and carry a plausible range/name after the '!' — garbage like
    /// "乱码!!!" written verbatim makes real Excel refuse the file.</summary>
    internal void ValidateDefinedNameRef(string refText)
    {
        // Defined-name bodies are full formulas — validating them properly
        // is out of scope (functions, unions, cross-part brackets, escaped
        // apostrophes are all legal). Reject only the empirically fatal
        // patterns that pass schema validation but make real Excel refuse
        // the file: doubled/trailing '!' ("乱码!!!") and stray '#' outside
        // the known error literals ("乱码###").
        // Formula-length ceiling (8192) applies to defined-name bodies too.
        ValidateFormulaLength(refText, "defined-name ref");
        var body = (refText ?? "").TrimStart('=').Trim();
        if (body.Length == 0) return;
        if (body.Contains('"')) return; // string literals — leave to Excel
        // Strip the known error literals first: "#REF!" legitimately ends
        // with '!' and must not trip the dangling-bang check below.
        var probe = System.Text.RegularExpressions.Regex.Replace(body,
            @"#(REF!|N/A|NAME\?|DIV/0!|VALUE!|NULL!|NUM!|SPILL!|CALC!|GETTING_DATA)",
            "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (probe.Contains('#'))
            throw new ArgumentException(
                $"Defined name ref '{refText}' contains '#' outside a known error literal — not valid formula text.");
        if (probe.Contains("!!") || probe.EndsWith("!", StringComparison.Ordinal))
            throw new ArgumentException(
                $"Defined name ref '{refText}' has a dangling '!' — a sheet qualifier must be followed by a range (e.g. Sheet1!$A$1:$B$5).");
    }

    /// <summary>Text to store in a numeric cell's &lt;v&gt;: the literal digits
    /// when already canonical (preserves >15-significant-digit values that
    /// double cannot represent), else the parsed double re-serialized.</summary>
    internal static string NormalizeNumericCellText(string text, double parsed)
        => CanonicalNumericLiteral.IsMatch(text)
            ? text
            : parsed.ToString(System.Globalization.CultureInfo.InvariantCulture);
}
