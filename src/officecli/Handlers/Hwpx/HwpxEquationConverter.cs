// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Equation conversion logic based on LibreOffice hwpeq.cxx (MPL 2.0).
// Original: https://github.com/LibreOffice/core/blob/master/hwpfilter/source/hwpeq.cxx
// NOT derived from H2Orestart ConvEquation.java (GPLv3) — GPL infection risk.

using System.Text;
using System.Text.RegularExpressions;

namespace OfficeCli.Handlers;

/// <summary>
/// Convert Hancom equation script to StarMath and LaTeX formats.
/// Hancom uses a proprietary scripting language similar to StarMath but with
/// case differences and some structural variations.
/// </summary>
public static class HwpxEquationConverter
{
    // CJK-safe word boundary: matches keyword not surrounded by alphanumeric chars.
    // Standard \b fails on CJK boundaries.
    private static string WB(string keyword)
        => $@"(?<![a-zA-Z0-9]){Regex.Escape(keyword)}(?![a-zA-Z0-9])";

    // ==================== Hancom → StarMath ====================

    /// <summary>
    /// Keyword mapping: Hancom script → StarMath.
    /// Most Hancom keywords are identical to StarMath; only differences are mapped.
    /// Order matters: longer/more specific patterns first to avoid partial matches.
    /// </summary>
    private static readonly (string Pattern, string Replacement)[] HwpToStarMathMap =
    {
        // === Structural Commands (case normalization) ===
        (WB("SQRT"),     "sqrt"),
        (WB("PILE"),     "alignc"),
        (WB("LPILE"),    "alignl"),
        (WB("RPILE"),    "alignr"),
        (WB("LSUB"),     "lsub"),
        (WB("LSUP"),     "lsup"),
        (WB("TIMES"),    "times"),
        (WB("PROD"),     "prod"),

        // === Integral variants → StarMath (normalize to base form) ===
        // StarMath only has: int, iint, iiint, lint (contour)
        (WB("OTINT"),  "iiint"),    // triple contour → triple (approx)
        (WB("ODINT"),  "iint"),     // double contour → double (approx)
        (WB("TINT"),   "iiint"),    // triple integral
        (WB("DINT"),   "iint"),     // double integral
        (WB("OINT"),   "lint"),     // contour integral
        (WB("INT"),    "int"),

        // === Set operators ===
        (WB("SMALLUNION"),  "union"),
        (WB("smallunion"),  "union"),
        (WB("UNION"),       "union"),
        (WB("CAP"),         "union"),
        (WB("SMALLINTER"),  "intersection"),
        (WB("smallinter"),  "intersection"),
        (WB("INTER"),       "intersection"),

        // === Bracket case normalization ===
        (WB("LEFT"),    "left"),
        (WB("RIGHT"),   "right"),
        (WB("MATRIX"),  "matrix"),
        (WB("BMATRIX"), "bmatrix"),
        (WB("DMATRIX"), "dmatrix"),
        (WB("PMATRIX"), "pmatrix"),
        (WB("CASES"),   "cases"),

        // === Special symbols ===
        (WB("ALEPH"),     "aleph"),
        (WB("HBAR"),      "hbar"),
        (WB("IMAG"),      "im"),
        (WB("WP"),        "wp"),
        (WB("ANGSTROM"),  "{circle A}"),
        (WB("IMATH"),     "{italic i}"),
        (WB("JMATH"),     "{italic j}"),
        (WB("ELL"),       "{italic l}"),
        (WB("LITER"),     "{italic l}"),
        (WB("OHM"),       "%OMEGA"),

        // === Operators ===
        (WB("OPLUS"),   "oplus"),
        (WB("OMINUS"),  "ominus"),
        (WB("OTIMES"),  "otimes"),
        (WB("ODOT"),    "odot"),
        (WB("OSLASH"),  "odivide"),
        (WB("ODIV"),    "odivide"),
        (WB("VEE"),     "or"),
        (WB("LOR"),     "or"),
        (WB("WEDGE"),   "and"),

        // === Set relations ===
        (WB("SUBSET"),    "subset"),
        (WB("SUPSET"),    "supset"),
        (WB("SUPERSET"),  "supset"),
        (WB("SUBSETEQ"),  "subseteq"),
        (WB("SUPSETEQ"),  "supseteq"),
        (WB("IN"),        "in"),
        (WB("OWNS"),      "owns"),
        (WB("LEQ"),       "<="),
        (WB("GEQ"),       ">="),
        (WB("PREC"),      "prec"),
        (WB("SUCC"),      "succ"),

        // === Arithmetic / Logic ===
        (WB("PLUSMINUS"),  "plusminus"),
        (WB("MINUSPLUS"), "minusplus"),
        (WB("DIVIDE"),    "div"),
        (WB("divide"),    "div"),
        (WB("CIRC"),      "circ"),
        (WB("EMPTYSET"),  "emptyset"),
        (WB("EXIST"),     "exists"),
        (WB("SIM"),       "sim"),
        (WB("APPROX"),    "approx"),
        (WB("SIMEQ"),     "simeq"),
        (WB("EQUIV"),     "equiv"),
        (WB("FORALL"),    "forall"),
        (WB("PARTIAL"),   "partial"),
        (WB("INF"),       "infinity"),
        (WB("inf"),       "infinity"),

        // === Arrows ===
        (WB("LRARROW"),  "dlrarrow"),   // double left-right
        (WB("LARROW"),   "dlarrow"),    // double left (uppercase = double)
        (WB("RARROW"),   "drarrow"),    // double right
        (WB("lrarrow"),  "lrarrow"),    // single left-right (lowercase = single)
        (WB("larrow"),   "leftarrow"),
        (WB("rarrow"),   "rightarrow"),
        (WB("uarrow"),   "uparrow"),
        (WB("darrow"),   "downarrow"),
        (WB("VERT"),     "parallel"),
        (WB("vert"),     "divides"),

        // === Dots ===
        (WB("cdots"),  "dotsaxis"),
        (WB("LDOTS"),  "dotslow"),
        (WB("ldots"),  "dotslow"),
        (WB("VDOTS"),  "dotsvert"),
        (WB("DDOTS"),  "dotsdown"),

        // === Decorations (case normalization) ===
        (WB("ACUTE"),     "acute"),
        (WB("GRAVE"),     "grave"),
        (WB("TILDE"),     "tilde"),
        (WB("OVERLINE"),  "overline"),
        (WB("under"),     "underline"),

        // === Miscellaneous ===
        (WB("TRIANGLED"), "nabla"),
        (WB("SANGLE"),   "%angle"),
        (WB("BOT"),      "ortho"),
        (WB("hund"),     "%perthousand"),
    };

    /// <summary>
    /// Convert Hancom equation script to StarMath format.
    /// Most keywords are identical; this handles case differences and structural variations.
    /// </summary>
    public static string ToStarMath(string hwpScript)
    {
        if (string.IsNullOrWhiteSpace(hwpScript)) return hwpScript;

        var result = hwpScript;

        // Step 1: Keyword replacements
        foreach (var (pattern, replacement) in HwpToStarMathMap)
        {
            result = Regex.Replace(result, pattern, replacement);
        }

        // Step 2: BIGG (large divider)
        if (result.Contains("bigg", StringComparison.OrdinalIgnoreCase))
        {
            result = Regex.Replace(result, @"(?i)(bigg)\s*/\s*(.*)", "wideslash {$2}");
            result = Regex.Replace(result, @"(?i)(bigg)\s*\\\s*(.*)", "widebslash {$2}");
        }

        // Step 3: OVER — ensure bare operands are wrapped in braces
        // "a over b" → "{a over b}" (StarMath expects braces around fraction)
        if (result.Contains("over", StringComparison.OrdinalIgnoreCase))
        {
            result = Regex.Replace(result,
                @"([^\s\}]+)\s+(?i:over)\s+([^\{\s]+)",
                "{$1 over $2}");
        }

        // Step 4: MATRIX variants — convert # (row sep) and & (col sep) to StarMath format
        if (result.Contains("matrix", StringComparison.OrdinalIgnoreCase))
        {
            result = ConvertMatrix(result);
        }

        // Step 5: Decorations — hat/check/tilde expand to wide* for multi-char
        foreach (var deco in new[] { "hat", "check", "tilde" })
        {
            if (result.Contains(deco, StringComparison.OrdinalIgnoreCase))
            {
                var m = Regex.Match(result, $@"(?<![a-zA-Z]){deco}\s*(\{{[^}}]+\}})");
                if (m.Success && m.Groups[1].Value.Length > 3) // {ab} = 4 chars, single char = {a} = 3
                    result = Regex.Replace(result,
                        $@"(?<![a-zA-Z]){deco}\s*(\{{[^}}]+\}})",
                        $"wide{deco} $1");
            }
        }

        // Step 6: ATOP → binom
        if (result.Contains("atop", StringComparison.OrdinalIgnoreCase))
        {
            result = Regex.Replace(result,
                @"(\{[^}]+\}|[^\s]+)\s+(?i:atop)\s+(\{[^}]+\}|[^\s]+)",
                "binom $1 $2");
        }

        return result;
    }

    /// <summary>Convert MATRIX variants: # → ## (row sep), &amp; → # (col sep).</summary>
    private static string ConvertMatrix(string input)
    {
        var matrixPattern = @"(?i)(bmatrix|dmatrix|pmatrix|matrix)\s*\{((?:[^{}]|\{[^{}]*\})+)\}";
        return Regex.Replace(input, matrixPattern, m =>
        {
            var type = m.Groups[1].Value.ToLowerInvariant();
            var body = m.Groups[2].Value;
            var converted = body.Replace("#", "##").Replace("&", "#");

            return type switch
            {
                "bmatrix" => $"left [ matrix{{ {converted}}} right ]",
                "dmatrix" => $"left lline matrix{{ {converted}}} right rline",
                "pmatrix" => $"left ( matrix{{ {converted}}} right )",
                _         => $"matrix{{ {converted}}}",
            };
        });
    }

    // ==================== Hancom → LaTeX ====================

    /// <summary>
    /// Keyword mapping: Hancom script → LaTeX.
    /// Longer patterns first to avoid partial matches.
    /// </summary>
    private static readonly (string Pattern, string Replacement)[] HwpToLatexMap =
    {
        // === Greek Uppercase ===
        (WB("Alpha"),    @"\Alpha"),
        (WB("Beta"),     @"\Beta"),
        (WB("Gamma"),    @"\Gamma"),
        (WB("Delta"),    @"\Delta"),
        (WB("Epsilon"),  @"\Epsilon"),
        (WB("Zeta"),     @"\Zeta"),
        (WB("Eta"),      @"\Eta"),
        (WB("Theta"),    @"\Theta"),
        (WB("Iota"),     @"\Iota"),
        (WB("Kappa"),    @"\Kappa"),
        (WB("Lambda"),   @"\Lambda"),
        (WB("Mu"),       @"\Mu"),
        (WB("Nu"),       @"\Nu"),
        (WB("Xi"),       @"\Xi"),
        (WB("Omicron"),  @"\Omicron"),
        (WB("Pi"),       @"\Pi"),
        (WB("Rho"),      @"\Rho"),
        (WB("Sigma"),    @"\Sigma"),
        (WB("SIGMA"),    @"\Sigma"),
        (WB("Tau"),      @"\Tau"),
        (WB("Upsilon"),  @"\Upsilon"),
        (WB("Phi"),      @"\Phi"),
        (WB("Chi"),      @"\Chi"),
        (WB("Psi"),      @"\Psi"),
        (WB("Omega"),    @"\Omega"),

        // === Greek Lowercase ===
        (WB("alpha"),    @"\alpha"),
        (WB("beta"),     @"\beta"),
        (WB("gamma"),    @"\gamma"),
        (WB("delta"),    @"\delta"),
        (WB("epsilon"),  @"\epsilon"),
        (WB("varepsilon"), @"\varepsilon"),
        (WB("zeta"),     @"\zeta"),
        (WB("eta"),      @"\eta"),
        (WB("theta"),    @"\theta"),
        (WB("vartheta"), @"\vartheta"),
        (WB("iota"),     @"\iota"),
        (WB("kappa"),    @"\kappa"),
        (WB("lambda"),   @"\lambda"),
        (WB("mu"),       @"\mu"),
        (WB("nu"),       @"\nu"),
        (WB("xi"),       @"\xi"),
        (WB("omicron"),  @"\omicron"),
        (WB("pi"),       @"\pi"),
        (WB("varpi"),    @"\varpi"),
        (WB("rho"),      @"\rho"),
        (WB("sigma"),    @"\sigma"),
        (WB("varsigma"), @"\varsigma"),
        (WB("tau"),      @"\tau"),
        (WB("upsilon"),  @"\upsilon"),
        (WB("phi"),      @"\phi"),
        (WB("varphi"),   @"\varphi"),
        (WB("chi"),      @"\chi"),
        (WB("psi"),      @"\psi"),
        (WB("omega"),    @"\omega"),

        // === Integral variants (longer first!) ===
        (WB("OTINT"),  @"\oiiint"),
        (WB("ODINT"),  @"\oiint"),
        (WB("TINT"),   @"\iiint"),
        (WB("iiint"),  @"\iiint"),
        (WB("DINT"),   @"\iint"),
        (WB("iint"),   @"\iint"),
        (WB("OINT"),   @"\oint"),
        (WB("oint"),   @"\oint"),
        (WB("INT"),    @"\int"),
        (WB("int"),    @"\int"),

        // === Functions & Large Operators ===
        (WB("SQRT"),   @"\sqrt"),
        (WB("sqrt"),   @"\sqrt"),
        (WB("sum"),    @"\sum"),
        (WB("SUM"),    @"\sum"),
        (WB("prod"),   @"\prod"),
        (WB("PROD"),   @"\prod"),
        (WB("lim"),    @"\lim"),
        (WB("Lim"),    @"\lim"),
        (WB("INF"),    @"\infty"),
        (WB("inf"),    @"\infty"),
        (WB("PARTIAL"),@"\partial"),
        (WB("partial"),@"\partial"),

        // === Subscript/superscript keywords ===
        (WB("from"),   "_"),
        (WB("to"),     "^"),
        (WB("sub"),    "_"),
        (WB("sup"),    "^"),

        // === Operators ===
        (WB("TIMES"),     @"\times"),
        (WB("times"),     @"\times"),
        (WB("DIVIDE"),    @"\div"),
        (WB("divide"),    @"\div"),
        (WB("PLUSMINUS"),  @"\pm"),
        (WB("MINUSPLUS"), @"\mp"),
        (WB("CIRC"),      @"\circ"),
        (WB("OPLUS"),     @"\oplus"),
        (WB("OMINUS"),    @"\ominus"),
        (WB("OTIMES"),    @"\otimes"),
        (WB("ODOT"),      @"\odot"),

        // === Set operators ===
        (WB("SMALLUNION"),  @"\bigcup"),
        (WB("smallunion"),  @"\bigcup"),
        (WB("UNION"),       @"\bigcup"),
        (WB("CAP"),         @"\bigcup"),
        (WB("SMALLINTER"),  @"\bigcap"),
        (WB("smallinter"),  @"\bigcap"),
        (WB("INTER"),       @"\bigcap"),

        // === Relations ===
        (WB("SUBSET"),    @"\subset"),
        (WB("SUPSET"),    @"\supset"),
        (WB("SUPERSET"),  @"\supset"),
        (WB("SUBSETEQ"),  @"\subseteq"),
        (WB("SUPSETEQ"),  @"\supseteq"),
        (WB("IN"),        @"\in"),
        (WB("OWNS"),      @"\ni"),
        (WB("LEQ"),       @"\leq"),
        (WB("GEQ"),       @"\geq"),
        (WB("PREC"),      @"\prec"),
        (WB("SUCC"),      @"\succ"),
        (WB("SIM"),       @"\sim"),
        (WB("APPROX"),    @"\approx"),
        (WB("SIMEQ"),     @"\simeq"),
        (WB("EQUIV"),     @"\equiv"),
        (WB("FORALL"),    @"\forall"),
        (WB("forall"),    @"\forall"),
        (WB("EXIST"),     @"\exists"),
        (WB("EMPTYSET"),  @"\emptyset"),

        // === Arrows (longer first!) ===
        (WB("LRARROW"),  @"\Leftrightarrow"),
        (WB("lrarrow"),  @"\leftrightarrow"),
        (WB("LARROW"),   @"\Leftarrow"),
        (WB("larrow"),   @"\leftarrow"),
        (WB("RARROW"),   @"\Rightarrow"),
        (WB("rarrow"),   @"\rightarrow"),
        (WB("uarrow"),   @"\uparrow"),
        (WB("darrow"),   @"\downarrow"),

        // === Dots ===
        (WB("cdots"),  @"\cdots"),
        (WB("LDOTS"),  @"\ldots"),
        (WB("ldots"),  @"\ldots"),
        (WB("VDOTS"),  @"\vdots"),
        (WB("DDOTS"),  @"\ddots"),

        // === Decorations ===
        (WB("hat"),       @"\widehat"),
        (WB("tilde"),     @"\widetilde"),
        (WB("bar"),       @"\overline"),
        (WB("overline"),  @"\overline"),
        (WB("OVERLINE"),  @"\overline"),
        (WB("vec"),       @"\vec"),
        (WB("dot"),       @"\dot"),
        (WB("ddot"),      @"\ddot"),
        (WB("acute"),     @"\acute"),
        (WB("ACUTE"),     @"\acute"),
        (WB("grave"),     @"\grave"),
        (WB("GRAVE"),     @"\grave"),
        (WB("check"),     @"\check"),
        (WB("breve"),     @"\breve"),
        (WB("under"),     @"\underline"),
        (WB("underline"), @"\underline"),

        // === Brackets (case normalization) ===
        (WB("LEFT"),   @"\left"),
        (WB("RIGHT"),  @"\right"),

        // === Miscellaneous ===
        (WB("ALEPH"),     @"\aleph"),
        (WB("HBAR"),      @"\hbar"),
        (WB("TRIANGLED"), @"\nabla"),
        (WB("nabla"),     @"\nabla"),
        (WB("VERT"),      @"\parallel"),
        (WB("vert"),      @"\mid"),
        (WB("BOT"),       @"\perp"),

        // === Line separator ===
        // # in Hancom = line break, maps to \\ in LaTeX
        // Handled separately in ConvertStructure
    };

    /// <summary>
    /// Convert Hancom equation script to LaTeX format.
    /// V1: keyword substitution + simple over→\frac pattern matching.
    /// V2 (future): recursive descent parser ported from hwpeq.cxx.
    /// </summary>
    public static string ToLatex(string hwpScript)
    {
        if (string.IsNullOrWhiteSpace(hwpScript)) return hwpScript;

        var result = hwpScript;

        // Phase 1: Convert "over" to \frac (simple single-level braces)
        // {a over b} → \frac{a}{b}
        result = ConvertOverToFrac(result);

        // Phase 2: Keyword substitutions
        foreach (var (pattern, replacement) in HwpToLatexMap)
        {
            result = Regex.Replace(result, pattern, replacement);
        }

        // Phase 3: MATRIX → LaTeX environments
        result = ConvertMatrixToLatex(result);

        // Phase 4: CASES → LaTeX cases environment
        if (result.Contains("cases", StringComparison.OrdinalIgnoreCase))
        {
            result = Regex.Replace(result,
                @"(?i)cases\s*\{((?:[^{}]|\{[^{}]*\})+)\}",
                @"\begin{cases}$1\end{cases}");
        }

        // Phase 5: ATOP → \binom
        if (result.Contains("atop", StringComparison.OrdinalIgnoreCase))
        {
            result = Regex.Replace(result,
                @"\{([^{}]+)\s+(?i:atop)\s+([^{}]+)\}",
                @"\binom{$1}{$2}");
        }

        // Phase 6: Line separator # → \\
        result = result.Replace(" # ", @" \\ ");
        // Standalone # at line boundaries
        result = Regex.Replace(result, @"(?<=[})])\s*#\s*", @" \\ ");

        // Phase 7: COLOR
        if (result.Contains("color", StringComparison.OrdinalIgnoreCase))
        {
            result = Regex.Replace(result,
                @"(?i)color\s*\{\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\}",
                @"\textcolor[RGB]{$1,$2,$3}");
        }

        return result;
    }

    /// <summary>
    /// Convert "{a over b}" to "\frac{a}{b}".
    /// V1: handles single-level braces only. Nested braces preserved as-is.
    /// </summary>
    private static string ConvertOverToFrac(string input)
    {
        // Pattern: {numerator over denominator} where numerator/denominator have no nested braces
        return Regex.Replace(input,
            @"\{([^{}]+?)\s+over\s+([^{}]+?)\}",
            @"\frac{$1}{$2}");
    }

    /// <summary>Convert MATRIX/BMATRIX/DMATRIX/PMATRIX to LaTeX environments.</summary>
    private static string ConvertMatrixToLatex(string input)
    {
        var matrixPattern = @"(?i)(bmatrix|dmatrix|pmatrix|matrix)\s*\{((?:[^{}]|\{[^{}]*\})+)\}";
        return Regex.Replace(input, matrixPattern, m =>
        {
            var type = m.Groups[1].Value.ToLowerInvariant();
            var body = m.Groups[2].Value;
            // Hancom: & = column sep, # = row sep
            // LaTeX: & = column sep, \\ = row sep
            var converted = body.Replace("#", @" \\ ").Trim();

            var env = type switch
            {
                "bmatrix" => "bmatrix",
                "dmatrix" => "vmatrix",
                "pmatrix" => "pmatrix",
                _         => "matrix",
            };
            return $@"\begin{{{env}}}{converted}\end{{{env}}}";
        });
    }
}
