// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Core;

/// <summary>
/// Bidirectional converter between LaTeX-subset formula syntax and Office Math (OMML).
///
/// Supported LaTeX syntax:
///   _{}        subscript       H_{2}O
///   ^{}        superscript     x^{2}
///   \frac{}{}  fraction        \frac{a}{b}
///   \sqrt{}    square root     \sqrt{x}
///   \sqrt[n]{} nth root        \sqrt[3]{x}
///   \sum       summation       \sum_{i=1}^{n}
///   \int       integral        \int_{0}^{1}
///   \prod      product         \prod_{i=1}^{n}
///   \left( \right)  auto-sized delimiters  \left(\frac{a}{b}\right)
///   \begin{pmatrix} a & b \\ c & d \end{pmatrix}   matrix (pmatrix/bmatrix/vmatrix/matrix)
///   \overset{}{} upper annotation   \overset{\triangle}{\rightarrow}
///   \underset{}{} lower annotation   \underset{k}{\rightarrow}
///   \text{}     text mode (upright)   \text{if } x > 0
///   \overline{} overline              \overline{AB}
///   \underline{} underline            \underline{x}
///   \hat{} \bar{} \vec{} \dot{} \ddot{} \tilde{}  accent marks
///   \lim \sin \cos \tan \log \ln \exp \min \max    function names (upright)
///   \binom{}{} binomial coefficient   \binom{n}{k}
///   \cases     piecewise function     \begin{cases} x & x>0 \\ -x & x\leq 0 \end{cases}
///   \pm \times \cdot \rightarrow \leftarrow \uparrow \downarrow \triangle
///   \alpha \beta \gamma \delta \pi \theta \sigma \omega \lambda \mu \epsilon
///   Single-char shorthand: H_2 x^2 (braces optional for single char)
/// </summary>
public static class FormulaParser
{
    // ==================== LaTeX → OMML ====================

    private const string KatexDocsHint = "See https://katex.org/docs/supported.html for supported syntax.";

    public static OpenXmlElement Parse(string latex)
    {
        try
        {
            // Preprocess: fix double-escaped backslashes (common AI/JSON over-escaping)
            // \\frac → \frac, \\sqrt → \sqrt, etc. (only when \\ is directly followed by a letter)
            latex = FixDoubleEscapedCommands(latex);
            // Preprocess: convert {a \over b} to \frac{a}{b}
            latex = RewriteOver(latex);
            var tokens = Tokenize(latex);
            var pos = 0;
            var nodes = ParseGroup(tokens, ref pos, false);
            return WrapInOfficeMath(nodes);
        }
        catch (Exception ex)
        {
            throw new FormulaParseException(
                $"Failed to parse formula: {ex.Message} {KatexDocsHint}", ex);
        }
    }

    /// <summary>
    /// Fix double-escaped backslashes from AI/JSON over-escaping.
    /// Converts \\cmd → \cmd when \\ is directly followed by a letter sequence.
    /// Safe because \\letter is not valid LaTeX (line break immediately followed by
    /// a bare word has no mathematical meaning). Legitimate usage like \\ \frac always
    /// has a space between the line break and the next command.
    /// </summary>
    private static string FixDoubleEscapedCommands(string latex)
    {
        // Replace \\ followed directly by a letter with \ (single pass, left to right)
        var sb = new System.Text.StringBuilder(latex.Length);
        int i = 0;
        while (i < latex.Length)
        {
            if (i + 2 < latex.Length && latex[i] == '\\' && latex[i + 1] == '\\' && char.IsLetter(latex[i + 2]))
            {
                // Collapse \\ to \ before the command
                sb.Append('\\');
                i += 2; // skip both backslashes, the letter will be consumed in the next iteration
            }
            else
            {
                sb.Append(latex[i]);
                i++;
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Rewrite LaTeX old-style {numerator \over denominator} to \frac{numerator}{denominator}.
    /// Handles nested braces correctly.
    /// </summary>
    private static string RewriteOver(string latex)
    {
        while (true)
        {
            var idx = latex.IndexOf("\\over");
            if (idx < 0) break;

            // Find the opening brace that contains \over
            int braceStart = -1;
            int depth = 0;
            for (int i = idx - 1; i >= 0; i--)
            {
                if (latex[i] == '}') depth++;
                else if (latex[i] == '{')
                {
                    if (depth == 0) { braceStart = i; break; }
                    depth--;
                }
            }

            // Find the closing brace
            int braceEnd = -1;
            depth = 0;
            for (int i = idx + 5; i < latex.Length; i++)
            {
                if (latex[i] == '{') depth++;
                else if (latex[i] == '}')
                {
                    if (depth == 0) { braceEnd = i; break; }
                    depth--;
                }
            }

            if (braceStart < 0 || braceEnd < 0)
                break; // malformed, skip

            var num = latex.Substring(braceStart + 1, idx - braceStart - 1).Trim();
            var den = latex.Substring(idx + 5, braceEnd - idx - 5).Trim();
            latex = latex.Substring(0, braceStart) + $"\\frac{{{num}}}{{{den}}}" + latex.Substring(braceEnd + 1);
        }
        return latex;
    }

    public static OpenXmlElement ParseAsDisplayParagraph(string latex)
    {
        var math = Parse(latex);
        return new M.Paragraph(new M.OfficeMath(math.ChildElements.Select(e => e.CloneNode(true)).ToArray()));
    }

    // ==================== OMML → LaTeX ====================

    public static string ToLatex(OpenXmlElement element)
    {
        return ToLatexByName(element);
    }

    private static string ToLatexByName(OpenXmlElement element)
    {
        var name = element.LocalName;

        switch (name)
        {
            case "oMathPara":
                return string.Concat(element.ChildElements.Select(ToLatexByName));

            case "oMath":
                return string.Concat(element.ChildElements.Select(ToLatexByName));

            case "r":
            {
                var tElem = element.ChildElements.FirstOrDefault(e => e.LocalName == "t");
                var text = tElem?.InnerText ?? "";
                // Check for math style in run properties (mathbf, mathrm, etc.)
                var rPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "rPr");
                if (rPr != null)
                {
                    var sty = rPr.ChildElements.FirstOrDefault(e => e.LocalName == "sty");
                    var styVal = sty?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value;
                    var hasNor = rPr.ChildElements.Any(e => e.LocalName == "nor");
                    if (hasNor)
                        return $"\\text{{{EscapeLatex(text)}}}";
                    if (styVal == "b")
                        return $"\\mathbf{{{EscapeLatex(text)}}}";
                    if (styVal == "bi")
                        return $"\\boldsymbol{{{EscapeLatex(text)}}}";
                    if (styVal == "p")
                        return $"\\mathrm{{{EscapeLatex(text)}}}";
                }
                return EscapeLatex(text);
            }

            case "sSub":
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var subText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "sub"));
                return NeedsBraces(subText) ? $"{baseText}_{{{subText}}}" : $"{baseText}_{subText}";
            }

            case "sSup":
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var supText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "sup"));
                return NeedsBraces(supText) ? $"{baseText}^{{{supText}}}" : $"{baseText}^{supText}";
            }

            case "sSubSup":
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var subText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "sub"));
                var supText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "sup"));
                var subPart = NeedsBraces(subText) ? $"_{{{subText}}}" : $"_{subText}";
                var supPart = NeedsBraces(supText) ? $"^{{{supText}}}" : $"^{supText}";
                return $"{baseText}{subPart}{supPart}";
            }

            case "f": // fraction
            {
                var num = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "num"));
                var den = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "den"));
                return $"\\frac{{{num}}}{{{den}}}";
            }

            case "rad": // radical
            {
                var deg = element.ChildElements.FirstOrDefault(e => e.LocalName == "deg");
                var baseElem = element.ChildElements.FirstOrDefault(e => e.LocalName == "e");
                var baseText = ArgToLatex(baseElem);
                // Check if degree is hidden or empty
                var radPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "radPr");
                var hideDeg = radPr?.ChildElements.FirstOrDefault(e => e.LocalName == "degHide");
                var isHidden = hideDeg != null && (hideDeg.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value == "1"
                    || hideDeg.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value == "true");
                var degText = isHidden ? "" : ArgToLatex(deg);
                if (string.IsNullOrEmpty(degText))
                    return $"\\sqrt{{{baseText}}}";
                return $"\\sqrt[{degText}]{{{baseText}}}";
            }

            case "nary":
            {
                var naryPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "naryPr");
                var chrElem = naryPr?.ChildElements.FirstOrDefault(e => e.LocalName == "chr");
                var chr = chrElem?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? "∑";
                var cmd = NaryCharToCommand(chr);
                var subText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "sub"));
                var supText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "sup"));
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var result = cmd;
                if (!string.IsNullOrEmpty(subText))
                    result += NeedsBraces(subText) ? $"_{{{subText}}}" : $"_{subText}";
                if (!string.IsNullOrEmpty(supText))
                    result += NeedsBraces(supText) ? $"^{{{supText}}}" : $"^{supText}";
                if (!string.IsNullOrEmpty(baseText))
                    result += $" {baseText}";
                return result;
            }

            case "d": // delimiter
            {
                var dPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "dPr");
                var begChr = dPr?.ChildElements.FirstOrDefault(e => e.LocalName == "begChr");
                var endChr = dPr?.ChildElements.FirstOrDefault(e => e.LocalName == "endChr");
                var begin = begChr?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? "(";
                var end = endChr?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? ")";
                // Check if delimiter wraps a matrix — emit \begin{pmatrix} etc.
                var bases = element.ChildElements.Where(e => e.LocalName == "e").ToList();
                if (bases.Count == 1)
                {
                    var inner = bases[0].ChildElements.FirstOrDefault(e => e.LocalName == "m");
                    if (inner != null)
                    {
                        var envName = (begin, end) switch
                        {
                            ("(", ")") => "pmatrix",
                            ("[", "]") => "bmatrix",
                            ("{", "}") => "Bmatrix",
                            ("|", "|") => "vmatrix",
                            _ => null
                        };
                        var matrixContent = ToLatexByName(inner);
                        if (envName != null)
                            return $"\\begin{{{envName}}}{matrixContent}\\end{{{envName}}}";
                        return $"\\left{begin}\\begin{{matrix}}{matrixContent}\\end{{matrix}}\\right{end}";
                    }
                }
                var content = string.Concat(bases.Select(ArgToLatex));
                return $"{begin}{content}{end}";
            }

            case "limUpp": // upper limit (overset)
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var limText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "lim"));
                return $"\\overset{{{limText}}}{{{baseText}}}";
            }

            case "limLow": // lower limit (underset)
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var limText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "lim"));
                return $"\\underset{{{limText}}}{{{baseText}}}";
            }

            case "bar": // overline/underline
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var barPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "barPr");
                var posElem = barPr?.ChildElements.FirstOrDefault(e => e.LocalName == "pos");
                var posVal = posElem?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value;
                return posVal == "bot" ? $"\\underline{{{baseText}}}" : $"\\overline{{{baseText}}}";
            }

            case "acc": // accent
            {
                var baseText = ArgToLatex(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var accPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "accPr");
                var chrElem = accPr?.ChildElements.FirstOrDefault(e => e.LocalName == "chr");
                var chr = chrElem?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? "\u0302";
                var cmd = chr switch
                {
                    "\u0302" => "hat",
                    "\u0304" => "bar",
                    "\u20D7" => "vec",
                    "\u0307" => "dot",
                    "\u0308" => "ddot",
                    "\u0303" => "tilde",
                    _ => "hat"
                };
                return $"\\{cmd}{{{baseText}}}";
            }

            case "m": // matrix
            {
                var matrixRows = element.ChildElements.Where(e => e.LocalName == "mr").ToList();
                var rowStrings = matrixRows.Select(mr =>
                    string.Join(" & ", mr.ChildElements.Where(e => e.LocalName == "e").Select(ArgToLatex)));
                // Detect delimiter wrapping from parent
                return string.Join(" \\\\ ", rowStrings);
            }

            default:
                // Recurse into unknown containers
                return string.Concat(element.ChildElements.Select(ToLatexByName));
        }
    }

    private static bool NeedsBraces(string text) => text.Length != 1;

    /// <summary>
    /// Convert OMML to readable Unicode text (for view text display).
    /// Uses Unicode subscript/superscript characters where possible.
    /// </summary>
    public static string ToReadableText(OpenXmlElement element)
    {
        var name = element.LocalName;

        switch (name)
        {
            case "oMathPara":
                return string.Concat(element.ChildElements.Select(ToReadableText));

            case "oMath":
                return string.Concat(element.ChildElements.Select(ToReadableText));

            case "r":
            {
                var tElem = element.ChildElements.FirstOrDefault(e => e.LocalName == "t");
                return tElem?.InnerText ?? "";
            }

            case "sSub":
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var subText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "sub"));
                return baseText + ToUnicodeSubscript(subText);
            }

            case "sSup":
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var supText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "sup"));
                return baseText + ToUnicodeSuperscript(supText);
            }

            case "sSubSup":
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var subText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "sub"));
                var supText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "sup"));
                return baseText + ToUnicodeSubscript(subText) + ToUnicodeSuperscript(supText);
            }

            case "f": // fraction
            {
                var num = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "num"));
                var den = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "den"));
                return $"({num})/({den})";
            }

            case "rad": // radical
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                return $"√({baseText})";
            }

            case "nary":
            {
                var naryPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "naryPr");
                var chrElem = naryPr?.ChildElements.FirstOrDefault(e => e.LocalName == "chr");
                var chr = chrElem?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? "∑";
                var subText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "sub"));
                var supText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "sup"));
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var result = chr;
                if (!string.IsNullOrEmpty(subText)) result += ToUnicodeSubscript(subText);
                if (!string.IsNullOrEmpty(supText)) result += ToUnicodeSuperscript(supText);
                result += $" {baseText}";
                return result;
            }

            case "d": // delimiter
            {
                var dPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "dPr");
                var begChr = dPr?.ChildElements.FirstOrDefault(e => e.LocalName == "begChr");
                var endChr = dPr?.ChildElements.FirstOrDefault(e => e.LocalName == "endChr");
                var begin = begChr?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? "(";
                var end = endChr?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? ")";
                var content = string.Concat(element.ChildElements
                    .Where(e => e.LocalName == "e")
                    .Select(ArgToReadable));
                return $"{begin}{content}{end}";
            }

            case "limUpp": // upper limit (overset)
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var limText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "lim"));
                return $"{baseText}({limText})";
            }

            case "limLow": // lower limit (underset)
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var limText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "lim"));
                return $"{baseText}({limText})";
            }

            case "bar": // overline/underline
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                return baseText;
            }

            case "acc": // accent
            {
                var baseText = ArgToReadable(element.ChildElements.FirstOrDefault(e => e.LocalName == "e"));
                var accPr = element.ChildElements.FirstOrDefault(e => e.LocalName == "accPr");
                var chrElem = accPr?.ChildElements.FirstOrDefault(e => e.LocalName == "chr");
                var chr = chrElem?.GetAttribute("val", "http://schemas.openxmlformats.org/officeDocument/2006/math").Value ?? "";
                return baseText + chr;
            }

            case "m": // matrix
            {
                var matrixRows = element.ChildElements.Where(e => e.LocalName == "mr").ToList();
                var rowStrings = matrixRows.Select(mr =>
                    string.Join(", ", mr.ChildElements.Where(e => e.LocalName == "e").Select(ArgToReadable)));
                return "[" + string.Join("; ", rowStrings) + "]";
            }

            default:
                return string.Concat(element.ChildElements.Select(ToReadableText));
        }
    }

    // ==================== Tokenizer ====================

    private enum TokenType { Text, Sub, Sup, LBrace, RBrace, LBracket, RBracket, Command, ColSep, RowSep }

    private record Token(TokenType Type, string Value);

    private static List<Token> Tokenize(string input)
    {
        var tokens = new List<Token>();
        int i = 0;

        while (i < input.Length)
        {
            char c = input[i];

            switch (c)
            {
                case '_':
                    tokens.Add(new Token(TokenType.Sub, "_"));
                    i++;
                    break;
                case '^':
                    tokens.Add(new Token(TokenType.Sup, "^"));
                    i++;
                    break;
                case '{':
                    tokens.Add(new Token(TokenType.LBrace, "{"));
                    i++;
                    break;
                case '}':
                    tokens.Add(new Token(TokenType.RBrace, "}"));
                    i++;
                    break;
                case '[':
                    tokens.Add(new Token(TokenType.LBracket, "["));
                    i++;
                    break;
                case ']':
                    tokens.Add(new Token(TokenType.RBracket, "]"));
                    i++;
                    break;
                case '&':
                    tokens.Add(new Token(TokenType.ColSep, "&"));
                    i++;
                    break;
                case '\\':
                    i++;
                    // \\ → row separator
                    if (i < input.Length && input[i] == '\\')
                    {
                        tokens.Add(new Token(TokenType.RowSep, "\\\\"));
                        i++;
                        break;
                    }
                    // Escaped special chars: \{ \} \| → literal text
                    if (i < input.Length && (input[i] == '{' || input[i] == '}' || input[i] == '|'))
                    {
                        tokens.Add(new Token(TokenType.Text, input[i].ToString()));
                        i++;
                        break;
                    }
                    var cmd = "";
                    while (i < input.Length && char.IsLetter(input[i]))
                    {
                        cmd += input[i];
                        i++;
                    }
                    if (cmd.Length == 0)
                    {
                        // \<non-letter> like \, \; \: \! → spacing commands
                        if (i < input.Length)
                        {
                            var spaceChar = input[i] switch
                            {
                                ',' => "\u2009", // thin space
                                ';' => "\u2005", // medium space
                                ':' => "\u2005", // medium space
                                '!' => "",        // negative thin space (ignore)
                                _ => input[i].ToString()
                            };
                            if (spaceChar.Length > 0)
                                tokens.Add(new Token(TokenType.Text, spaceChar));
                            i++;
                        }
                    }
                    else
                    {
                        tokens.Add(new Token(TokenType.Command, cmd));
                    }
                    break;
                default:
                    // Collect consecutive text characters
                    var text = "";
                    while (i < input.Length && !IsSpecialChar(input[i]))
                    {
                        text += input[i];
                        i++;
                    }
                    if (text.Length > 0)
                        tokens.Add(new Token(TokenType.Text, text));
                    break;
            }
        }

        return tokens;
    }

    private static bool IsSpecialChar(char c) => c is '_' or '^' or '{' or '}' or '[' or ']' or '\\' or '&';

    // ==================== Parser ====================

    private static List<OpenXmlElement> ParseGroup(List<Token> tokens, ref int pos, bool insideBraces)
    {
        var elements = new List<OpenXmlElement>();

        while (pos < tokens.Count)
        {
            var token = tokens[pos];

            if (token.Type == TokenType.RBrace)
            {
                if (insideBraces) { pos++; break; }
                pos++;
                continue;
            }

            if (token.Type == TokenType.Text)
            {
                pos++;
                OpenXmlElement textElement = MakeMathRun(token.Value);
                // Check if next token is sub or sup
                textElement = TryAttachScript(tokens, ref pos, textElement);
                elements.Add(textElement);
            }
            else if (token.Type == TokenType.LBrace)
            {
                pos++;
                var inner = ParseGroup(tokens, ref pos, true);
                var grouped = WrapInOfficeMath(inner);
                // Check if next is sub/sup
                var result = TryAttachScript(tokens, ref pos, grouped);
                elements.Add(result);
            }
            else if (token.Type == TokenType.Command)
            {
                pos++;
                var cmdElement = ParseCommand(token.Value, tokens, ref pos);
                cmdElement = TryAttachScript(tokens, ref pos, cmdElement);
                elements.Add(cmdElement);
            }
            else if (token.Type == TokenType.Sub || token.Type == TokenType.Sup)
            {
                // Sub/sup without preceding element — use empty base
                var emptyRun = MakeMathRun("");
                var scripted = TryAttachScript(tokens, ref pos, emptyRun);
                elements.Add(scripted);
            }
            else if (token.Type == TokenType.LBracket || token.Type == TokenType.RBracket)
            {
                pos++;
                var bracketText = token.Type == TokenType.LBracket ? "[" : "]";
                OpenXmlElement bracketElement = MakeMathRun(bracketText);
                bracketElement = TryAttachScript(tokens, ref pos, bracketElement);
                elements.Add(bracketElement);
            }
            else
            {
                pos++;
            }
        }

        return elements;
    }

    private static OpenXmlElement TryAttachScript(List<Token> tokens, ref int pos, OpenXmlElement baseElement)
    {
        while (pos < tokens.Count)
        {
            if (tokens[pos].Type == TokenType.Sub)
            {
                pos++;
                var subContent = ParseSingleArg(tokens, ref pos);

                // Check if followed by superscript → SubSuperscript
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sup)
                {
                    pos++;
                    var supContent = ParseSingleArg(tokens, ref pos);
                    baseElement = new M.SubSuperscript(
                        new M.Base(ExtractChildren(baseElement)),
                        new M.SubArgument(ExtractChildren(subContent)),
                        new M.SuperArgument(ExtractChildren(supContent))
                    );
                }
                else
                {
                    baseElement = new M.Subscript(
                        new M.Base(ExtractChildren(baseElement)),
                        new M.SubArgument(ExtractChildren(subContent))
                    );
                }
            }
            else if (tokens[pos].Type == TokenType.Sup)
            {
                pos++;
                var supContent = ParseSingleArg(tokens, ref pos);

                // Check if followed by subscript → SubSuperscript
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sub)
                {
                    pos++;
                    var subContent = ParseSingleArg(tokens, ref pos);
                    baseElement = new M.SubSuperscript(
                        new M.Base(ExtractChildren(baseElement)),
                        new M.SubArgument(ExtractChildren(subContent)),
                        new M.SuperArgument(ExtractChildren(supContent))
                    );
                }
                else
                {
                    baseElement = new M.Superscript(
                        new M.Base(ExtractChildren(baseElement)),
                        new M.SuperArgument(ExtractChildren(supContent))
                    );
                }
            }
            else
            {
                break;
            }
        }

        return baseElement;
    }

    private static OpenXmlElement ParseSingleArg(List<Token> tokens, ref int pos)
    {
        if (pos >= tokens.Count) return MakeMathRun("");

        if (tokens[pos].Type == TokenType.LBrace)
        {
            pos++;
            var inner = ParseGroup(tokens, ref pos, true);
            return inner.Count == 1 ? inner[0] : WrapInOfficeMath(inner);
        }

        if (tokens[pos].Type == TokenType.Command)
        {
            pos++;
            return ParseCommand(tokens[pos - 1].Value, tokens, ref pos);
        }

        if (tokens[pos].Type == TokenType.Text)
        {
            // Single character for shorthand: H_2 takes just "2", but "2O" should take just "2"
            var text = tokens[pos].Value;
            pos++;
            if (text.Length == 1)
                return MakeMathRun(text);
            // For multi-char text in a subscript/superscript arg without braces, take only first char
            // Put the rest back as a new text token
            if (text.Length > 1)
            {
                tokens.Insert(pos, new Token(TokenType.Text, text[1..]));
            }
            return MakeMathRun(text[..1]);
        }

        pos++;
        return MakeMathRun("");
    }

    private static OpenXmlElement ParseCommand(string cmd, List<Token> tokens, ref int pos)
    {
        // Symbol commands
        var symbol = CommandToSymbol(cmd);
        if (symbol != null)
            return MakeMathRun(symbol);

        switch (cmd)
        {
            case "frac":
            {
                var num = ParseBracedArg(tokens, ref pos);
                var den = ParseBracedArg(tokens, ref pos);
                return new M.Fraction(
                    new M.Numerator(ExtractChildren(num)),
                    new M.Denominator(ExtractChildren(den))
                );
            }
            case "sqrt":
            {
                // Check for optional [degree]
                OpenXmlElement? degree = null;
                if (pos < tokens.Count && tokens[pos].Type == TokenType.LBracket)
                {
                    pos++; // skip [
                    var degTokens = new List<Token>();
                    while (pos < tokens.Count && tokens[pos].Type != TokenType.RBracket)
                    {
                        degTokens.Add(tokens[pos]);
                        pos++;
                    }
                    if (pos < tokens.Count) pos++; // skip ]
                    int degPos = 0;
                    var degElements = ParseGroup(degTokens, ref degPos, false);
                    degree = degElements.Count == 1 ? degElements[0] : WrapInOfficeMath(degElements);
                }
                var content = ParseBracedArg(tokens, ref pos);

                var radical = new M.Radical(
                    new M.RadicalProperties(),
                    new M.Degree(degree != null ? ExtractChildren(degree) : Array.Empty<OpenXmlElement>()),
                    new M.Base(ExtractChildren(content))
                );

                // For square root (no degree), hide the degree
                if (degree == null)
                {
                    radical.RadicalProperties!.AppendChild(new M.HideDegree { Val = M.BooleanValues.True });
                }

                return radical;
            }
            case "matrix":
            {
                // \matrix{a&b\\c&d} — shorthand syntax (no \begin/\end)
                if (pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
                {
                    pos++; // skip {
                    // Temporarily collect tokens until matching }
                    var matrixTokens = new List<Token>();
                    int braceDepth = 1;
                    while (pos < tokens.Count && braceDepth > 0)
                    {
                        if (tokens[pos].Type == TokenType.LBrace) braceDepth++;
                        else if (tokens[pos].Type == TokenType.RBrace) { braceDepth--; if (braceDepth == 0) { pos++; break; } }
                        matrixTokens.Add(tokens[pos]);
                        pos++;
                    }
                    // Insert into tokens stream and parse as matrix
                    int mpos = 0;
                    // Reuse the matrix parser by appending a fake \end token
                    matrixTokens.Add(new Token(TokenType.Command, "end"));
                    matrixTokens.Add(new Token(TokenType.LBrace, "{"));
                    matrixTokens.Add(new Token(TokenType.Text, "matrix"));
                    matrixTokens.Add(new Token(TokenType.RBrace, "}"));
                    return ParseMatrix("matrix", matrixTokens, ref mpos);
                }
                return MakeMathRun("matrix");
            }
            case "begin":
            {
                // Read environment name from {name}
                var envName = "";
                if (pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
                {
                    pos++;
                    while (pos < tokens.Count && tokens[pos].Type != TokenType.RBrace)
                    {
                        envName += tokens[pos].Value;
                        pos++;
                    }
                    if (pos < tokens.Count) pos++; // skip }
                }

                if (envName is "matrix" or "pmatrix" or "bmatrix" or "Bmatrix" or "vmatrix" or "cases"
                    or "array")
                {
                    // For array, skip optional column spec like {cc}
                    if (envName == "array" && pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
                    {
                        pos++; // skip {
                        while (pos < tokens.Count && tokens[pos].Type != TokenType.RBrace) pos++;
                        if (pos < tokens.Count) pos++; // skip }
                    }
                    return ParseMatrix(envName, tokens, ref pos);
                }
                if (envName is "align" or "align*" or "aligned" or "gathered" or "eqnarray"
                    or "eqnarray*" or "split")
                {
                    // Multi-line equation environments → m:eqArr (equation array)
                    // These use \\ for row breaks and & for alignment points
                    // Reuse matrix parser which already handles \\ and &
                    var matrixEl = ParseMatrix(envName, tokens, ref pos);
                    // ParseMatrix wraps in a delimiter for cases/pmatrix/etc.
                    // For align/gathered, we want the raw m:m (matrix) without delimiters
                    if (matrixEl is M.Delimiter delim)
                    {
                        // Extract the matrix from inside the delimiter
                        var innerBase = delim.GetFirstChild<M.Base>();
                        var innerMatrix = innerBase?.GetFirstChild<M.Matrix>();
                        if (innerMatrix != null)
                            return innerMatrix.CloneNode(true);
                    }
                    return matrixEl;
                }

                // Unknown environment, render as text
                return MakeMathRun($"\\begin{{{envName}}}");
            }
            case "end":
            {
                // Skip \end{name} — should be consumed by matrix parser
                if (pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
                {
                    pos++;
                    while (pos < tokens.Count && tokens[pos].Type != TokenType.RBrace) pos++;
                    if (pos < tokens.Count) pos++;
                }
                return MakeMathRun("");
            }
            case "left":
            {
                // Get opening delimiter character from next token
                var openChar = "(";
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Command)
                {
                    // Handle \left\langle, \left\lfloor, \left\lceil, \left\lvert, \left\|
                    var delimCmd = tokens[pos].Value;
                    var mapped = delimCmd switch
                    {
                        "langle" => "\u27E8",
                        "lceil" => "\u2308",
                        "lfloor" => "\u230A",
                        "lvert" => "|",
                        "lVert" => "\u2016",
                        "|" => "\u2016",
                        _ => null
                    };
                    if (mapped != null) { openChar = mapped; pos++; }
                }
                else if (pos < tokens.Count && tokens[pos].Type == TokenType.Text)
                {
                    openChar = tokens[pos].Value[..1];
                    if (tokens[pos].Value.Length > 1)
                        tokens[pos] = new Token(TokenType.Text, tokens[pos].Value[1..]);
                    else
                        pos++;
                }
                else if (pos < tokens.Count && tokens[pos].Type == TokenType.LBracket)
                {
                    openChar = "[";
                    pos++;
                }

                // Parse content until \right
                var content = new List<OpenXmlElement>();
                var closeChar = openChar switch { "(" => ")", "[" => "]", "{" => "}", "|" => "|", "\u27E8" => "\u27E9", "\u2308" => "\u2309", "\u230A" => "\u230B", "\u2016" => "\u2016", _ => ")" };
                while (pos < tokens.Count)
                {
                    if (tokens[pos].Type == TokenType.Command && tokens[pos].Value == "right")
                    {
                        pos++;
                        // Get closing delimiter character — capture the actual delimiter
                        if (pos < tokens.Count && tokens[pos].Type == TokenType.Command)
                        {
                            // Handle \right\rangle, \right\rfloor, \right\rceil, etc.
                            var rDelimCmd = tokens[pos].Value;
                            var rMapped = rDelimCmd switch
                            {
                                "rangle" => "\u27E9",
                                "rceil" => "\u2309",
                                "rfloor" => "\u230B",
                                "rvert" => "|",
                                "rVert" => "\u2016",
                                "|" => "\u2016",
                                _ => null
                            };
                            if (rMapped != null) { closeChar = rMapped; pos++; }
                        }
                        else if (pos < tokens.Count && tokens[pos].Type == TokenType.Text)
                        {
                            closeChar = tokens[pos].Value[..1];
                            if (tokens[pos].Value.Length > 1)
                                tokens[pos] = new Token(TokenType.Text, tokens[pos].Value[1..]);
                            else
                                pos++;
                        }
                        else if (pos < tokens.Count && tokens[pos].Type == TokenType.RBracket)
                        {
                            closeChar = "]";
                            pos++;
                        }
                        break;
                    }

                    // Reuse main parsing logic for each element
                    if (tokens[pos].Type == TokenType.Text)
                    {
                        var textEl = MakeMathRun(tokens[pos].Value);
                        pos++;
                        textEl = (M.Run)TryAttachScript(tokens, ref pos, textEl);
                        content.Add(textEl);
                    }
                    else if (tokens[pos].Type == TokenType.LBrace)
                    {
                        pos++;
                        var inner = ParseGroup(tokens, ref pos, true);
                        var grouped = WrapInOfficeMath(inner);
                        grouped = TryAttachScript(tokens, ref pos, grouped);
                        content.Add(grouped);
                    }
                    else if (tokens[pos].Type == TokenType.Command)
                    {
                        pos++;
                        var cmdEl = ParseCommand(tokens[pos - 1].Value, tokens, ref pos);
                        cmdEl = TryAttachScript(tokens, ref pos, cmdEl);
                        content.Add(cmdEl);
                    }
                    else if (tokens[pos].Type == TokenType.Sub || tokens[pos].Type == TokenType.Sup)
                    {
                        var emptyRun = MakeMathRun("");
                        var scripted = TryAttachScript(tokens, ref pos, emptyRun);
                        content.Add(scripted);
                    }
                    else if (tokens[pos].Type == TokenType.LBracket || tokens[pos].Type == TokenType.RBracket)
                    {
                        var bracketText = tokens[pos].Type == TokenType.LBracket ? "[" : "]";
                        var bracketRun = MakeMathRun(bracketText);
                        pos++;
                        bracketRun = (M.Run)TryAttachScript(tokens, ref pos, bracketRun);
                        content.Add(bracketRun);
                    }
                    else
                    {
                        pos++;
                    }
                }

                var dPr = new M.DelimiterProperties();
                if (openChar != "(")
                    dPr.AppendChild(new M.BeginChar { Val = openChar });
                if (closeChar != ")")
                    dPr.AppendChild(new M.EndChar { Val = closeChar });

                var delimiter = new M.Delimiter(dPr);
                var arg = new M.Base(content.Select(e => e.CloneNode(true)).ToArray());
                delimiter.AppendChild(arg);
                return delimiter;
            }
            case "right":
            {
                // Orphan \right — shouldn't happen if paired with \left, just skip
                return MakeMathRun("");
            }
            case "overset":
            {
                var above = ParseBracedArg(tokens, ref pos);
                var baseArg = ParseBracedArg(tokens, ref pos);
                return new M.LimitUpper(
                    new M.LimitUpperProperties(),
                    new M.Base(ExtractChildren(baseArg)),
                    new M.Limit(ExtractChildren(above))
                );
            }
            case "underset":
            {
                var below = ParseBracedArg(tokens, ref pos);
                var baseArg = ParseBracedArg(tokens, ref pos);
                return new M.LimitLower(
                    new M.LimitLowerProperties(),
                    new M.Base(ExtractChildren(baseArg)),
                    new M.Limit(ExtractChildren(below))
                );
            }
            case "text":
            {
                // \text{...} → M.Run with normal text properties (upright, not math italic)
                var content = ParseBracedArg(tokens, ref pos);
                var text = ExtractText(content);
                var run = new M.Run(
                    new M.RunProperties(new M.NormalText()),
                    new M.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                );
                return run;
            }
            case "overline":
            {
                var arg = ParseBracedArg(tokens, ref pos);
                return new M.Bar(
                    new M.BarProperties(new M.Position { Val = M.VerticalJustificationValues.Top }),
                    new M.Base(ExtractChildren(arg))
                );
            }
            case "underline":
            {
                var arg = ParseBracedArg(tokens, ref pos);
                return new M.Bar(
                    new M.BarProperties(new M.Position { Val = M.VerticalJustificationValues.Bottom }),
                    new M.Base(ExtractChildren(arg))
                );
            }
            case "hat" or "bar" or "vec" or "dot" or "ddot" or "tilde":
            {
                var accentChar = cmd switch
                {
                    "hat" => "\u0302",   // combining circumflex
                    "bar" => "\u0304",   // combining macron
                    "vec" => "\u20D7",   // combining right arrow above
                    "dot" => "\u0307",   // combining dot above
                    "ddot" => "\u0308",  // combining diaeresis
                    "tilde" => "\u0303", // combining tilde
                    _ => "\u0302"
                };
                var arg = ParseBracedArg(tokens, ref pos);
                return new M.Accent(
                    new M.AccentProperties(new M.AccentChar { Val = accentChar }),
                    new M.Base(ExtractChildren(arg))
                );
            }
            case "lim" or "sin" or "cos" or "tan" or "log" or "ln" or "exp" or "min" or "max"
                or "sup" or "inf" or "det" or "gcd" or "dim" or "ker" or "hom" or "deg"
                or "arg" or "sec" or "csc" or "cot" or "sinh" or "cosh" or "tanh":
            {
                // Function names: render upright (non-italic) using M.NormalText
                var funcRun = new M.Run(
                    new M.RunProperties(new M.NormalText()),
                    new M.Text(cmd) { Space = SpaceProcessingModeValues.Preserve }
                );

                // For \lim, check for sub/sup to create nary-like limLow structure
                if (cmd == "lim" && pos < tokens.Count && tokens[pos].Type == TokenType.Sub)
                {
                    pos++;
                    var subArg = ParseSingleArg(tokens, ref pos);
                    return new M.LimitLower(
                        new M.LimitLowerProperties(),
                        new M.Base(funcRun),
                        new M.Limit(ExtractChildren(subArg))
                    );
                }
                return funcRun;
            }
            case "binom":
            {
                var top = ParseBracedArg(tokens, ref pos);
                var bottom = ParseBracedArg(tokens, ref pos);
                // Binomial = parenthesized fraction with no bar
                var frac = new M.Fraction(
                    new M.FractionProperties(new M.FractionType { Val = M.FractionTypeValues.NoBar }),
                    new M.Numerator(ExtractChildren(top)),
                    new M.Denominator(ExtractChildren(bottom))
                );
                var delimiter = new M.Delimiter(new M.DelimiterProperties());
                delimiter.AppendChild(new M.Base(frac));
                return delimiter;
            }
            case "mathbf" or "mathrm" or "mathit" or "mathbb" or "mathcal" or "boldsymbol":
            {
                var arg = ParseBracedArg(tokens, ref pos);
                var text = ExtractText(arg);
                var style = cmd switch
                {
                    "mathbf" => M.StyleValues.Bold,
                    "boldsymbol" => M.StyleValues.BoldItalic,
                    "mathrm" => M.StyleValues.Plain,
                    "mathit" => M.StyleValues.Italic,
                    _ => M.StyleValues.Plain
                };
                if (cmd is "mathbb" or "mathcal")
                {
                    // Double-struck and calligraphic: use NormalText + special Unicode if available,
                    // otherwise render as styled text with script style
                    var rPr = new M.RunProperties(new M.NormalText());
                    if (cmd == "mathcal")
                        rPr = new M.RunProperties(new M.Style { Val = M.StyleValues.Plain }, new M.Script());
                    return new M.Run(
                        rPr,
                        new M.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                    );
                }
                return new M.Run(
                    new M.RunProperties(new M.Style { Val = style }),
                    new M.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                );
            }
            case "sum" or "int" or "iint" or "iiint" or "prod" or "coprod" or "bigcup" or "bigcap":
            {
                var naryChar = cmd switch
                {
                    "sum" => "∑",
                    "int" => "∫",
                    "iint" => "∬",
                    "iiint" => "∭",
                    "prod" => "∏",
                    "coprod" => "∐",
                    "bigcup" => "⋃",
                    "bigcap" => "⋂",
                    _ => "∑"
                };
                var naryProps = new M.NaryProperties(new M.AccentChar { Val = naryChar });

                // Parse optional sub and sup limits (they come as _{}^{} after the command)
                OpenXmlElement? subArg = null;
                OpenXmlElement? supArg = null;

                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sub)
                {
                    pos++;
                    subArg = ParseSingleArg(tokens, ref pos);
                }
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sup)
                {
                    pos++;
                    supArg = ParseSingleArg(tokens, ref pos);
                }

                // Hide sub/sup limits when not provided to avoid empty boxes
                if (subArg == null)
                    naryProps.AppendChild(new M.HideSubArgument { Val = M.BooleanValues.True });
                if (supArg == null)
                    naryProps.AppendChild(new M.HideSuperArgument { Val = M.BooleanValues.True });

                // Parse the base expression (next arg or next element)
                OpenXmlElement baseArg;
                if (pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
                {
                    baseArg = ParseBracedArg(tokens, ref pos);
                }
                else if (pos < tokens.Count && (tokens[pos].Type == TokenType.Text || tokens[pos].Type == TokenType.Command))
                {
                    baseArg = ParseSingleArg(tokens, ref pos);
                }
                else
                {
                    baseArg = MakeMathRun("");
                }

                return new M.Nary(
                    naryProps,
                    new M.SubArgument(subArg != null ? ExtractChildren(subArg) : Array.Empty<OpenXmlElement>()),
                    new M.SuperArgument(supArg != null ? ExtractChildren(supArg) : Array.Empty<OpenXmlElement>()),
                    new M.Base(ExtractChildren(baseArg))
                );
            }
            case "cancel":
            case "bcancel":
            case "xcancel":
            case "cancelto":
            {
                // Cancel/strikethrough: use m:borderBox with m:strikeH
                var cancelArg = ParseBracedArg(tokens, ref pos);
                var bbPr = new M.BorderBoxProperties();
                if (cmd is "bcancel")
                    bbPr.AppendChild(new M.StrikeBottomLeftToTopRight { Val = M.BooleanValues.True });
                else if (cmd is "xcancel")
                {
                    bbPr.AppendChild(new M.StrikeHorizontal { Val = M.BooleanValues.True });
                    bbPr.AppendChild(new M.StrikeBottomLeftToTopRight { Val = M.BooleanValues.True });
                }
                else
                    bbPr.AppendChild(new M.StrikeHorizontal { Val = M.BooleanValues.True });
                return new M.BorderBox(bbPr, new M.Base(ExtractChildren(cancelArg)));
            }
            case "boxed":
            {
                // \boxed{expr} → m:borderBox (all four sides)
                var arg = ParseBracedArg(tokens, ref pos);
                return new M.BorderBox(
                    new M.BorderBoxProperties(),
                    new M.Base(ExtractChildren(arg))
                );
            }
            case "underbrace":
            {
                // \underbrace{expr}_{label} → m:groupChr with ⏟ below
                var arg = ParseBracedArg(tokens, ref pos);
                var groupChr = new M.GroupChar(
                    new M.GroupCharProperties(
                        new M.AccentChar { Val = "\u23DF" },
                        new M.Position { Val = M.VerticalJustificationValues.Bottom }
                    ),
                    new M.Base(ExtractChildren(arg))
                );
                // Check for subscript label
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sub)
                {
                    pos++;
                    var label = ParseSingleArg(tokens, ref pos);
                    return new M.LimitLower(
                        new M.LimitLowerProperties(),
                        new M.Base(groupChr),
                        new M.Limit(ExtractChildren(label))
                    );
                }
                return groupChr;
            }
            case "overbrace":
            {
                // \overbrace{expr}^{label} → m:groupChr with ⏞ above
                var arg = ParseBracedArg(tokens, ref pos);
                var groupChr = new M.GroupChar(
                    new M.GroupCharProperties(
                        new M.AccentChar { Val = "\u23DE" },
                        new M.Position { Val = M.VerticalJustificationValues.Top }
                    ),
                    new M.Base(ExtractChildren(arg))
                );
                // Check for superscript label
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sup)
                {
                    pos++;
                    var label = ParseSingleArg(tokens, ref pos);
                    return new M.LimitUpper(
                        new M.LimitUpperProperties(),
                        new M.Base(groupChr),
                        new M.Limit(ExtractChildren(label))
                    );
                }
                return groupChr;
            }
            case "color":
            {
                // \color{red}{expr} → m:r with w:color run property
                var colorArg = ParseBracedArg(tokens, ref pos);
                var colorName = ExtractText(colorArg);
                var contentArg = ParseBracedArg(tokens, ref pos);
                var contentText = ExtractText(contentArg);
                var colorHex = NamedColorToHex(colorName);
                var run = new M.Run(
                    new M.Text(contentText) { Space = SpaceProcessingModeValues.Preserve }
                );
                // Insert w:rPr with color before the m:t
                var wrPr = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.Color { Val = colorHex }
                );
                run.InsertAt(wrPr, 0);
                return run;
            }
            case "textcolor":
            {
                // \textcolor{red}{expr} — alias for \color
                var colorArg = ParseBracedArg(tokens, ref pos);
                var colorName = ExtractText(colorArg);
                var contentArg = ParseBracedArg(tokens, ref pos);
                var contentText = ExtractText(contentArg);
                var colorHex = NamedColorToHex(colorName);
                var run = new M.Run(
                    new M.Text(contentText) { Space = SpaceProcessingModeValues.Preserve }
                );
                var wrPr = new DocumentFormat.OpenXml.Wordprocessing.RunProperties(
                    new DocumentFormat.OpenXml.Wordprocessing.Color { Val = colorHex }
                );
                run.InsertAt(wrPr, 0);
                return run;
            }
            case "pmod":
            {
                // \pmod{n} → (mod n) with upright "mod"
                var arg = ParseBracedArg(tokens, ref pos);
                var modRun = new M.Run(
                    new M.RunProperties(new M.NormalText()),
                    new M.Text("mod") { Space = SpaceProcessingModeValues.Preserve }
                );
                var spaceRun = MakeMathRun("\u2003");
                var dPr = new M.DelimiterProperties();
                // Parentheses are default, no need to set begin/end
                var delimiter = new M.Delimiter(dPr);
                delimiter.AppendChild(new M.Base(modRun, spaceRun, ExtractChildren(arg)[0].CloneNode(true)));
                return delimiter;
            }
            case "bmod":
            {
                // \bmod → upright "mod" (binary operator form)
                return new M.Run(
                    new M.RunProperties(new M.NormalText()),
                    new M.Text("\u2003mod\u2003") { Space = SpaceProcessingModeValues.Preserve }
                );
            }
            case "arcsin" or "arccos" or "arctan" or "arccot" or "arcsec" or "arccsc":
            {
                // Arc-trig functions: render upright like \sin, \cos, etc.
                var funcRun = new M.Run(
                    new M.RunProperties(new M.NormalText()),
                    new M.Text(cmd) { Space = SpaceProcessingModeValues.Preserve }
                );
                return funcRun;
            }
            case "operatorname":
            {
                // \operatorname{name} → upright function name
                var arg = ParseBracedArg(tokens, ref pos);
                var opText = ExtractText(arg);
                var funcRun = new M.Run(
                    new M.RunProperties(new M.NormalText()),
                    new M.Text(opText) { Space = SpaceProcessingModeValues.Preserve }
                );
                // Check for subscript limits (like \lim)
                if (pos < tokens.Count && tokens[pos].Type == TokenType.Sub)
                {
                    pos++;
                    var subArg = ParseSingleArg(tokens, ref pos);
                    return new M.LimitLower(
                        new M.LimitLowerProperties(),
                        new M.Base(funcRun),
                        new M.Limit(ExtractChildren(subArg))
                    );
                }
                return funcRun;
            }

            default:
                // Unknown command: render as text with backslash
                return MakeMathRun($"\\{cmd}");
        }
    }

    private static OpenXmlElement ParseBracedArg(List<Token> tokens, ref int pos)
    {
        if (pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
        {
            pos++;
            var inner = ParseGroup(tokens, ref pos, true);
            return inner.Count == 1 ? inner[0] : WrapInOfficeMath(inner);
        }
        return ParseSingleArg(tokens, ref pos);
    }

    private static OpenXmlElement ParseMatrix(string envName, List<Token> tokens, ref int pos)
    {
        var rows = new List<List<List<OpenXmlElement>>>();
        var currentRow = new List<List<OpenXmlElement>>();
        var currentCell = new List<OpenXmlElement>();

        while (pos < tokens.Count)
        {
            // Check for \end{envName}
            if (tokens[pos].Type == TokenType.Command && tokens[pos].Value == "end")
            {
                pos++;
                // Skip {envName}
                if (pos < tokens.Count && tokens[pos].Type == TokenType.LBrace)
                {
                    pos++;
                    while (pos < tokens.Count && tokens[pos].Type != TokenType.RBrace) pos++;
                    if (pos < tokens.Count) pos++;
                }
                break;
            }

            if (tokens[pos].Type == TokenType.RowSep)
            {
                pos++;
                currentRow.Add(currentCell);
                rows.Add(currentRow);
                currentRow = new List<List<OpenXmlElement>>();
                currentCell = new List<OpenXmlElement>();
                continue;
            }

            if (tokens[pos].Type == TokenType.ColSep)
            {
                pos++;
                currentRow.Add(currentCell);
                currentCell = new List<OpenXmlElement>();
                continue;
            }

            // Parse element into current cell (same logic as ParseGroup)
            if (tokens[pos].Type == TokenType.Text)
            {
                var el = MakeMathRun(tokens[pos].Value);
                pos++;
                var result = TryAttachScript(tokens, ref pos, el);
                currentCell.Add(result);
            }
            else if (tokens[pos].Type == TokenType.LBrace)
            {
                pos++;
                var inner = ParseGroup(tokens, ref pos, true);
                var grouped = WrapInOfficeMath(inner);
                grouped = TryAttachScript(tokens, ref pos, grouped);
                currentCell.Add(grouped);
            }
            else if (tokens[pos].Type == TokenType.Command)
            {
                pos++;
                var cmdEl = ParseCommand(tokens[pos - 1].Value, tokens, ref pos);
                cmdEl = TryAttachScript(tokens, ref pos, cmdEl);
                currentCell.Add(cmdEl);
            }
            else if (tokens[pos].Type == TokenType.Sub || tokens[pos].Type == TokenType.Sup)
            {
                var emptyRun = MakeMathRun("");
                var scripted = TryAttachScript(tokens, ref pos, emptyRun);
                currentCell.Add(scripted);
            }
            else
            {
                pos++;
            }
        }

        // Add last cell/row
        if (currentCell.Count > 0 || currentRow.Count > 0)
        {
            currentRow.Add(currentCell);
            rows.Add(currentRow);
        }

        // Build OMML Matrix
        var matrix = new M.Matrix(new M.MatrixProperties());
        foreach (var row in rows)
        {
            var mr = new M.MatrixRow();
            foreach (var cell in row)
            {
                var baseEl = new M.Base(cell.Select(e => e.CloneNode(true)).ToArray());
                mr.AppendChild(baseEl);
            }
            matrix.AppendChild(mr);
        }

        // Wrap with delimiter based on environment
        if (envName == "matrix")
            return matrix;

        var (beginChar, endChar) = envName switch
        {
            "pmatrix" => ("(", ")"),
            "bmatrix" => ("[", "]"),
            "Bmatrix" => ("{", "}"),
            "vmatrix" => ("|", "|"),
            "cases" => ("{", ""),
            _ => ("(", ")")
        };

        var dPr = new M.DelimiterProperties();
        if (beginChar != "(")
            dPr.AppendChild(new M.BeginChar { Val = beginChar });
        if (endChar != ")")
            dPr.AppendChild(new M.EndChar { Val = endChar });

        // For cases: left-align cells
        if (envName == "cases")
        {
            // Set column justification to left for the matrix
            var mPr = matrix.ChildElements.FirstOrDefault(e => e.LocalName == "mPr") as M.MatrixProperties;
            if (mPr != null)
            {
                var colCount = rows.Max(r => r.Count);
                var mcs = new M.MatrixColumns();
                for (int ci = 0; ci < colCount; ci++)
                {
                    mcs.AppendChild(new M.MatrixColumn(
                        new M.MatrixColumnJustification { Val = M.HorizontalAlignmentValues.Left }
                    ));
                }
                mPr.AppendChild(mcs);
            }
        }

        var delimiter = new M.Delimiter(dPr);
        delimiter.AppendChild(new M.Base(matrix));
        return delimiter;
    }

    // ==================== Helpers ====================

    private static M.Run MakeMathRun(string text)
    {
        return new M.Run(new M.Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static OpenXmlElement WrapInOfficeMath(List<OpenXmlElement> elements)
    {
        if (elements.Count == 1) return elements[0];
        var math = new M.OfficeMath();
        foreach (var e in elements)
            math.AppendChild(e.CloneNode(true));
        return math;
    }

    private static OpenXmlElement[] ExtractChildren(OpenXmlElement element)
    {
        if (element is M.OfficeMath math)
            return math.ChildElements.Select(e => e.CloneNode(true)).ToArray();
        return new[] { element.CloneNode(true) };
    }

    private static string NamedColorToHex(string color)
    {
        // Strip # prefix if present, return 6-digit hex
        color = color.Trim().TrimStart('#');
        if (color.Length == 6 && color.All(c => "0123456789ABCDEFabcdef".Contains(c)))
            return color.ToUpperInvariant();
        return color.ToLowerInvariant() switch
        {
            "red" => "FF0000",
            "blue" => "0000FF",
            "green" => "008000",
            "black" => "000000",
            "white" => "FFFFFF",
            "orange" => "FF8C00",
            "purple" => "800080",
            "brown" => "8B4513",
            "gray" or "grey" => "808080",
            "cyan" => "00FFFF",
            "magenta" => "FF00FF",
            "yellow" => "FFD700",
            "darkred" => "8B0000",
            "darkblue" => "00008B",
            "darkgreen" => "006400",
            "lightblue" => "ADD8E6",
            "lightgreen" => "90EE90",
            "pink" => "FFC0CB",
            "teal" => "008080",
            "navy" => "000080",
            "maroon" => "800000",
            "olive" => "808000",
            _ => "000000"
        };
    }

    private static string ExtractText(OpenXmlElement element)
    {
        if (element is M.Run run)
            return run.ChildElements.FirstOrDefault(e => e.LocalName == "t")?.InnerText ?? "";
        if (element is M.OfficeMath oMath)
            return string.Concat(oMath.ChildElements.Select(ExtractText));
        return element.InnerText;
    }

    private static string ArgToLatex(OpenXmlElement? arg)
    {
        if (arg == null) return "";
        return string.Concat(arg.ChildElements.Select(ToLatex));
    }

    private static string ArgToReadable(OpenXmlElement? arg)
    {
        if (arg == null) return "";
        return string.Concat(arg.ChildElements.Select(ToReadableText));
    }

    private static string EscapeLatex(string text)
    {
        // Reverse-map special Unicode symbols back to LaTeX commands
        foreach (var (symbol, cmd) in SymbolToCommandMap)
        {
            text = text.Replace(symbol, cmd);
        }
        return text;
    }

    private static string NaryCharToCommand(string chr) => chr switch
    {
        "∑" => "\\sum",
        "∫" => "\\int",
        "∬" => "\\iint",
        "∭" => "\\iiint",
        "∏" => "\\prod",
        "∐" => "\\coprod",
        "⋃" => "\\bigcup",
        "⋂" => "\\bigcap",
        _ => chr
    };

    private static string? CommandToSymbol(string cmd) => cmd switch
    {
        // Arrows
        "rightarrow" => "→",
        "leftarrow" => "←",
        "uparrow" => "↑",
        "downarrow" => "↓",
        "Rightarrow" => "⇒",
        "Leftarrow" => "⇐",
        "leftrightarrow" => "↔",
        "Leftrightarrow" => "⇔",
        "rightleftharpoons" => "⇌",
        // Operators
        "pm" => "±",
        "mp" => "∓",
        "times" => "×",
        "div" => "÷",
        "cdot" => "·",
        "ast" => "∗",
        "star" => "⋆",
        "circ" => "∘",
        "oplus" => "⊕",
        "ominus" => "⊖",
        "otimes" => "⊗",
        "odot" => "⊙",
        "bullet" => "∙",
        // Relations
        "leq" or "le" => "≤",
        "geq" or "ge" => "≥",
        "neq" or "ne" => "≠",
        "approx" => "≈",
        "equiv" => "≡",
        "sim" => "∼",
        "subset" => "⊂",
        "supset" => "⊃",
        "subseteq" => "⊆",
        "supseteq" => "⊇",
        "in" => "∈",
        "notin" => "∉",
        "forall" => "∀",
        "exists" => "∃",
        "nabla" => "∇",
        "partial" => "∂",
        "infty" => "∞",
        "triangle" => "△",
        "prime" => "′",
        "hbar" => "ℏ",
        "cdots" => "⋯",
        "ldots" => "…",
        "vdots" => "⋮",
        "ddots" => "⋱",
        // Delimiters (when used standalone, not with \left/\right)
        "langle" => "\u27E8",     // ⟨ mathematical left angle bracket
        "rangle" => "\u27E9",     // ⟩ mathematical right angle bracket
        "lceil" => "\u2308",      // ⌈ left ceiling
        "rceil" => "\u2309",      // ⌉ right ceiling
        "lfloor" => "\u230A",     // ⌊ left floor
        "rfloor" => "\u230B",     // ⌋ right floor
        "lvert" => "|",
        "rvert" => "|",
        "lVert" => "\u2016",      // ‖ double vertical line
        "rVert" => "\u2016",
        "vert" => "|",
        "Vert" => "\u2016",
        // Set notation
        "emptyset" => "∅",
        "varnothing" => "∅",
        "setminus" => "∖",
        "complement" => "∁",
        "cap" => "∩",
        "cup" => "∪",
        // Spacing
        "quad" => "\u2003",    // em space
        "qquad" => "\u2003\u2003", // double em space
        "," => "\u2009",       // thin space
        ";" => "\u2005",       // medium mathematical space
        "!" => "",             // negative thin space (approximate with nothing)
        // Greek lowercase
        "alpha" => "α",
        "beta" => "β",
        "gamma" => "γ",
        "delta" => "δ",
        "epsilon" => "ε",
        "zeta" => "ζ",
        "eta" => "η",
        "theta" => "θ",
        "iota" => "ι",
        "kappa" => "κ",
        "lambda" => "λ",
        "mu" => "μ",
        "nu" => "ν",
        "xi" => "ξ",
        "pi" => "π",
        "rho" => "ρ",
        "sigma" => "σ",
        "tau" => "τ",
        "upsilon" => "υ",
        "phi" => "φ",
        "chi" => "χ",
        "psi" => "ψ",
        "omega" => "ω",
        // Greek uppercase
        "Gamma" => "Γ",
        "Delta" => "Δ",
        "Theta" => "Θ",
        "Lambda" => "Λ",
        "Xi" => "Ξ",
        "Pi" => "Π",
        "Sigma" => "Σ",
        "Phi" => "Φ",
        "Psi" => "Ψ",
        "Omega" => "Ω",
        _ => null
    };

    private static readonly (string Symbol, string Command)[] SymbolToCommandMap = new[]
    {
        ("→", "\\rightarrow "), ("←", "\\leftarrow "), ("↑", "\\uparrow "), ("↓", "\\downarrow "),
        ("⇒", "\\Rightarrow "), ("⇐", "\\Leftarrow "),
        ("±", "\\pm "), ("×", "\\times "), ("÷", "\\div "), ("·", "\\cdot "),
        ("≤", "\\leq "), ("≥", "\\geq "), ("≠", "\\neq "), ("≈", "\\approx "), ("≡", "\\equiv "),
        ("∈", "\\in "), ("∀", "\\forall "), ("∃", "\\exists "), ("∞", "\\infty "), ("△", "\\triangle "),
        ("′", "\\prime "), ("ℏ", "\\hbar "), ("⇌", "\\rightleftharpoons "),
        ("α", "\\alpha "), ("β", "\\beta "), ("γ", "\\gamma "), ("δ", "\\delta "),
        ("ε", "\\epsilon "), ("θ", "\\theta "), ("λ", "\\lambda "), ("μ", "\\mu "),
        ("π", "\\pi "), ("σ", "\\sigma "), ("φ", "\\phi "), ("ω", "\\omega "),
        ("Σ", "\\Sigma "), ("Π", "\\Pi "), ("Δ", "\\Delta "), ("Ω", "\\Omega "),
    };

    // ==================== Unicode subscript/superscript ====================

    private static string ToUnicodeSubscript(string text)
    {
        return string.Concat(text.Select(c => c switch
        {
            '0' => '₀', '1' => '₁', '2' => '₂', '3' => '₃', '4' => '₄',
            '5' => '₅', '6' => '₆', '7' => '₇', '8' => '₈', '9' => '₉',
            '+' => '₊', '-' => '₋', '=' => '₌', '(' => '₍', ')' => '₎',
            'a' => 'ₐ', 'e' => 'ₑ', 'i' => 'ᵢ', 'n' => 'ₙ', 'o' => 'ₒ',
            'r' => 'ᵣ', 'x' => 'ₓ',
            _ => c
        }));
    }

    private static string ToUnicodeSuperscript(string text)
    {
        return string.Concat(text.Select(c => c switch
        {
            '0' => '⁰', '1' => '¹', '2' => '²', '3' => '³', '4' => '⁴',
            '5' => '⁵', '6' => '⁶', '7' => '⁷', '8' => '⁸', '9' => '⁹',
            '+' => '⁺', '-' => '⁻', '=' => '⁼', '(' => '⁽', ')' => '⁾',
            'n' => 'ⁿ', 'i' => 'ⁱ',
            _ => c
        }));
    }
}

/// <summary>
/// Exception thrown when FormulaParser fails to parse a LaTeX formula.
/// </summary>
public class FormulaParseException : Exception
{
    public FormulaParseException(string message, Exception innerException)
        : base(message, innerException) { }
}
