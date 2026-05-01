// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Selector ====================

    private record SelectorPart(string? Element, Dictionary<string, string> Attributes, string? ContainsText, SelectorPart? ChildSelector);

    private static SelectorPart ParseSelector(string selector)
    {
        // Support: element[attr=value] > child[attr=value]
        // Split on '>' but skip '>' inside [...] brackets (e.g. [size>=14pt])
        var childParts = SplitChildCombinator(selector);

        SelectorPart? childSelector = null;
        if (childParts.Length > 1)
        {
            childSelector = ParseSingleSelector(childParts[1]);
        }

        var main = ParseSingleSelector(childParts[0]);
        return main with { ChildSelector = childSelector };
    }

    /// <summary>
    /// Split selector on '>' child combinator, but skip '>' inside [...] brackets.
    /// "paragraph[size>=14pt] > run[bold=true]" → ["paragraph[size>=14pt]", "run[bold=true]"]
    /// </summary>
    private static string[] SplitChildCombinator(string selector)
    {
        int depth = 0;
        for (int i = 0; i < selector.Length; i++)
        {
            switch (selector[i])
            {
                case '[': depth++; break;
                case ']': depth--; break;
                case '>' when depth == 0:
                    // Found a top-level '>' combinator
                    return new[]
                    {
                        selector[..i].Trim(),
                        selector[(i + 1)..].Trim()
                    };
            }
        }
        return new[] { selector };
    }

    private static SelectorPart ParseSingleSelector(string selector)
    {
        var attrs = new Dictionary<string, string>();
        string? element = null;
        string? containsText = null;

        // Extract element name (before any [ or : modifier)
        var firstMod = selector.Length;
        var bracketIdx = selector.IndexOf('[');
        if (bracketIdx >= 0 && bracketIdx < firstMod) firstMod = bracketIdx;
        var colonIdx = selector.IndexOf(':');
        if (colonIdx >= 0 && colonIdx < firstMod) firstMod = colonIdx;

        element = selector[..firstMod].Trim();
        // CONSISTENCY(selector-case): element names are case-insensitive
        // ("OLE" == "ole" == "Ole"). Attribute values stay case-sensitive.
        element = element.ToLowerInvariant();
        if (string.IsNullOrEmpty(element)) element = null;

        // Parse [attr=value] attributes
        if (bracketIdx >= 0)
        {
            var attrPart = selector[bracketIdx..];
            // Verify brackets are balanced. CLI layer rejects unclosed brackets
            // before reaching here, but direct API callers can pass malformed
            // selectors — surface a clean error rather than silently returning
            // empty results.
            int openCount = 0, closeCount = 0;
            foreach (var ch in attrPart)
            {
                if (ch == '[') openCount++;
                else if (ch == ']') closeCount++;
            }
            if (openCount != closeCount)
                throw new ArgumentException($"Malformed selector: unclosed bracket in '{selector}'");

            foreach (System.Text.RegularExpressions.Match m in
                System.Text.RegularExpressions.Regex.Matches(attrPart, @"\[(\w+)(\\?!?=)([^\]]+)\]"))
            {
                var key = m.Groups[1].Value;
                var op = m.Groups[2].Value.Replace("\\", "");
                var val = m.Groups[3].Value.Trim('\'', '"');
                attrs[key] = (op == "!=" ? "!" : "") + val;
            }
        }

        // Parse :contains("text") pseudo-selector
        if (selector.Contains(":contains("))
        {
            var idx = selector.IndexOf(":contains(");
            var endIdx = selector.IndexOf(')', idx + 10);
            if (endIdx >= 0)
                containsText = selector[(idx + 10)..endIdx].Trim('\'', '"');
        }

        // Parse :empty pseudo-selector
        if (selector.Contains(":empty"))
        {
            attrs["__empty"] = "true";
        }

        // Parse :no-alt pseudo-selector
        if (selector.Contains(":no-alt"))
        {
            attrs["__no-alt"] = "true";
        }

        return new SelectorPart(element, attrs, containsText, null);
    }

    private bool MatchesSelector(Paragraph para, SelectorPart selector, int lineNum)
    {
        // If selector targets runs (has child selector), only match parent paragraph
        if (selector.ChildSelector != null)
        {
            // Check paragraph-level attributes
            if (selector.Element != null && selector.Element != "p" && selector.Element != "paragraph")
                return false;
            return MatchesParagraphAttrs(para, selector.Attributes);
        }

        if (selector.Element != null && selector.Element != "p" && selector.Element != "paragraph")
            return false;

        if (!MatchesParagraphAttrs(para, selector.Attributes))
            return false;

        if (selector.Attributes.ContainsKey("__empty"))
        {
            return string.IsNullOrWhiteSpace(GetParagraphText(para));
        }

        if (selector.ContainsText != null)
        {
            return GetParagraphText(para).Contains(selector.ContainsText);
        }

        return true;
    }

    private bool MatchesParagraphAttrs(Paragraph para, Dictionary<string, string> attrs)
    {
        // Cache first text-bearing run for run-level property checks
        Run? firstRun = null;
        bool firstRunResolved = false;

        foreach (var (key, rawVal) in attrs)
        {
            if (key == "__empty") continue;
            // BUG-R34-03: `text` and `type` are not paragraph XML attributes — they are
            // node-level metadata populated post-construction (DocumentNode.Text / .Type).
            // Pre-filter cannot resolve them, so falling through to GenericXmlQuery
            // returned null and silently zero-filtered the result. Skip these keys here
            // and let the CLI-level AttributeFilter post-filter handle them against the
            // populated DocumentNode (which already has .Text / .Type).
            // CONSISTENCY(query-pre-vs-post-filter): mirrors how `~=` is intentionally
            // not parsed by the Word selector regex so AttributeFilter handles it.
            if (key.Equals("text", StringComparison.OrdinalIgnoreCase) ||
                key.Equals("type", StringComparison.OrdinalIgnoreCase))
                continue;
            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            string? actual = key.ToLowerInvariant() switch
            {
                // CONSISTENCY(style-dual-key): `style` is lenient — matches
                // either styleId (`H5`) or display name (`H正文`). For
                // unambiguous queries use `styleId=` or `styleName=` below.
                "style" => GetStyleName(para),
                "styleid" => para.ParagraphProperties?.ParagraphStyleId?.Val?.Value,
                "stylename" => GetStyleName(para),
                "align" or "alignment" => para.ParagraphProperties?.Justification?.Val != null
                    ? para.ParagraphProperties.Justification.Val.InnerText : null,
                "firstlineindent" => para.ParagraphProperties?.Indentation?.FirstLine?.Value,
                "numId" or "numid" => para.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.HasValue == true
                    ? para.ParagraphProperties.NumberingProperties.NumberingId.Val.Value.ToString() : null,
                "numLevel" or "numlevel" or "ilvl" => para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.HasValue == true
                    ? para.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val.Value.ToString() : null,
                "liststyle" => GetParagraphListStyle(para),
                // R9-bt-1: pPr <w:bidi/> resolves to canonical 'direction' on
                // Get; selectors must accept the same key. Returns "rtl" /
                // "ltr" / null mirroring how Navigation emits it.
                "direction" or "dir" or "bidi" => para.ParagraphProperties?.BiDi switch
                {
                    null => null,
                    var b => TryReadOnOff(b.Val) is { } on ? (on ? "rtl" : "ltr") : null,
                },
                // R11-bt-5: `rtl` alias — mirrors paragraph-level direction in
                // boolean form so users can write paragraph[rtl=true] without
                // remembering whether bidi/direction is the canonical key.
                // rtl=true ⇔ BiDi present and truthy.
                // rtl=false ⇔ BiDi absent OR explicit val=0 (LTR is the
                // implicit default in OOXML, so absent w:bidi == ltr).
                "rtl" => para.ParagraphProperties?.BiDi switch
                {
                    null => "false",
                    var b => TryReadOnOff(b.Val) is { } on ? (on ? "true" : "false") : "true",
                },
                // Run-level properties: check first text-bearing run (same approach as Get readback)
                "bold" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved)?.RunProperties?.Bold != null ? "true" : "false",
                "italic" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved)?.RunProperties?.Italic != null ? "true" : "false",
                "font" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved) is { } fr1 ? GetRunFont(fr1) : null,
                "size" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved) is { } fr2 ? GetRunFontSize(fr2) : null,
                "color" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved)?.RunProperties?.Color?.Val?.Value is { } cv
                    ? ParseHelpers.FormatHexColor(cv) : null,
                "underline" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved)?.RunProperties?.Underline?.Val?.InnerText,
                "strike" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved)?.RunProperties?.Strike != null ? "true" : "false",
                "highlight" => GetFirstRunForSelector(para, ref firstRun, ref firstRunResolved)?.RunProperties?.Highlight?.Val?.InnerText,
                _ => GenericXmlQuery.GetAttributeValue(para, key)
                     ?? (para.ParagraphProperties != null ? GenericXmlQuery.GetAttributeValue(para.ParagraphProperties, key) : null)
            };

            // For style, also match against styleId (e.g., "Heading1" vs display name "heading 1")
            bool matches;
            if (key.Equals("style", StringComparison.OrdinalIgnoreCase))
            {
                var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase)
                       || string.Equals(styleId, val, StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            }
            if (negate ? matches : !matches) return false;
        }
        return true;
    }

    private static Run? GetFirstRunForSelector(Paragraph para, ref Run? cached, ref bool resolved)
    {
        if (!resolved)
        {
            cached = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
            resolved = true;
        }
        return cached;
    }

    private static bool MatchesRunSelector(Run run, Paragraph parent, SelectorPart selector)
    {
        // CONSISTENCY(run-special-content): query elements ptab / fieldChar /
        // instrText / tab / break each select runs whose primary inline
        // payload is the matching structural element. Mirrors Get's type
        // upgrade (WordHandler.Navigation.cs run branch) and AttributeFilter's
        // dual-key matching — the canonical name written by Get is the same
        // name accepted here on Query so users don't have to translate
        // between OOXML local-names (br/fldChar) and DOM types (break/fieldChar).
        var elementLower = selector.Element?.ToLowerInvariant();
        if (elementLower != null
            && elementLower != "r" && elementLower != "run"
            && elementLower != "ptab" && elementLower != "positionaltab"
            && elementLower != "fieldchar" && elementLower != "fldchar"
            && elementLower != "instrtext"
            && elementLower != "tab"
            && elementLower != "break" && elementLower != "br")
            return false;

        // Type filter: when element names a specialized run kind, the run's
        // actual content must match. Otherwise the run-walk would return
        // every paragraph child indiscriminately.
        if (elementLower is "ptab" or "positionaltab")
        {
            if (run.GetFirstChild<PositionalTab>() == null) return false;
        }
        else if (elementLower is "fieldchar" or "fldchar")
        {
            if (run.GetFirstChild<FieldChar>() == null) return false;
        }
        else if (elementLower == "instrtext")
        {
            if (run.GetFirstChild<FieldCode>() == null) return false;
        }
        else if (elementLower == "tab")
        {
            // Only match runs whose primary content is a tab (no <w:t>); a
            // run with text + tab still surfaces as type=run, not type=tab.
            if (run.GetFirstChild<TabChar>() == null) return false;
            if (run.GetFirstChild<Text>() != null) return false;
        }
        else if (elementLower is "break" or "br")
        {
            if (run.GetFirstChild<Break>() == null) return false;
            if (run.GetFirstChild<Text>() != null) return false;
        }

        foreach (var (key, rawVal) in selector.Attributes)
        {
            // CONSISTENCY(query-pre-vs-post-filter): see MatchesParagraphAttrs above.
            // `text` / `type` are not XML attributes — let AttributeFilter post-filter
            // resolve them against DocumentNode.Text / .Type.
            if (key.Equals("text", StringComparison.OrdinalIgnoreCase) ||
                key.Equals("type", StringComparison.OrdinalIgnoreCase))
                continue;
            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            string? actual = key.ToLowerInvariant() switch
            {
                "font" => GetRunFont(run),
                "size" => GetRunFontSize(run),
                "bold" => run.RunProperties?.Bold != null ? "true" : "false",
                "italic" => run.RunProperties?.Italic != null ? "true" : "false",
                // CONSISTENCY(run-special-content): structural inline-element
                // attributes mirror what Get exposes in node.Format.
                "align" or "alignment" => run.GetFirstChild<PositionalTab>()?.Alignment?.HasValue == true
                    ? run.GetFirstChild<PositionalTab>()!.Alignment!.InnerText : null,
                "relativeto" => run.GetFirstChild<PositionalTab>()?.RelativeTo?.HasValue == true
                    ? run.GetFirstChild<PositionalTab>()!.RelativeTo!.InnerText : null,
                "leader" => run.GetFirstChild<PositionalTab>()?.Leader?.HasValue == true
                    ? run.GetFirstChild<PositionalTab>()!.Leader!.InnerText : null,
                // R11-bt-5: `rtl` selector mirrors run rPr/rtl boolean.
                // Get returns node.Format["rtl"]=true|false; the selector
                // must accept the same key. Absent rtl element ⇒ null
                // (so rtl=false matches only runs with explicit w:rtl val=0).
                "rtl" or "direction" or "dir" or "bidi" => run.RunProperties?.RightToLeftText switch
                {
                    null => null,
                    var r => TryReadOnOff(r.Val) is { } on
                        ? (key.Equals("rtl", StringComparison.OrdinalIgnoreCase)
                            ? (on ? "true" : "false")
                            : (on ? "rtl" : "ltr"))
                        : (key.Equals("rtl", StringComparison.OrdinalIgnoreCase) ? "true" : "rtl"),
                },
                "fieldchartype" => run.GetFirstChild<FieldChar>()?.FieldCharType?.HasValue == true
                    ? run.GetFirstChild<FieldChar>()!.FieldCharType!.InnerText : null,
                "instr" => run.GetFirstChild<FieldCode>()?.Text,
                "breaktype" => run.GetFirstChild<Break>()?.Type?.HasValue == true
                    ? run.GetFirstChild<Break>()!.Type!.InnerText : null,
                _ => GenericXmlQuery.GetAttributeValue(run, key)
                     ?? (run.RunProperties != null ? GenericXmlQuery.GetAttributeValue(run.RunProperties, key) : null)
            };

            // CONSISTENCY(color-input): align selector input with Add/Set — accept
            // `#FF0000`, `FF0000`, or named colors. OOXML stores hex without `#`.
            if (key.Equals("color", StringComparison.OrdinalIgnoreCase))
            {
                actual = NormalizeColorForCompare(actual);
                val = NormalizeColorForCompare(val) ?? val;
            }

            bool matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            if (negate ? matches : !matches) return false;
        }

        if (selector.ContainsText != null)
        {
            return GetRunText(run).Contains(selector.ContainsText);
        }

        return true;
    }

    private static string? NormalizeColorForCompare(string? raw)
    {
        if (string.IsNullOrEmpty(raw)) return raw;
        var s = raw.Trim();
        if (s.StartsWith("#")) s = s[1..];
        return s.ToUpperInvariant();
    }

    private string GetHeaderRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var headerPart = _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault(idx - 1);
        return headerPart?.Header?.OuterXml ?? $"(header[{idx}] not found)";
    }

    private string GetFooterRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var footerPart = _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault(idx - 1);
        return footerPart?.Footer?.OuterXml ?? $"(footer[{idx}] not found)";
    }

    private string GetChartRawXml(string partPath)
    {
        var idx = 0;
        var bracketIdx = partPath.IndexOf('[');
        if (bracketIdx >= 0)
            int.TryParse(partPath[(bracketIdx + 1)..^0].TrimEnd(']'), out idx);

        var chartPart = _doc.MainDocumentPart?.ChartParts.ElementAtOrDefault(idx - 1);
        return chartPart?.ChartSpace?.OuterXml ?? $"(chart[{idx}] not found)";
    }
}
