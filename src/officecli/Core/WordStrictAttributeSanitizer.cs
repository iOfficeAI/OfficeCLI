// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

namespace OfficeCli.Core;

// Real-world docx files from legacy editors (WPS, older Word, third-party tools)
// sometimes carry attribute values that violate the OOXML schema — e.g.
// `<w:b w:val="yes"/>` or `<w:jc w:val="bogus"/>`. Native Word is lenient,
// but DocumentFormat.OpenXml throws FormatException the moment any reader
// accesses `.Val.Value` on the typed property. Since the crash is lazy, it
// surfaces unpredictably deep inside rendering code (HtmlPreview.Css,
// styling, etc.) rather than at open time.
//
// This sanitizer walks raw XML attributes (no typed conversion) right after
// Open, repairs or strips the offending values, and lets every downstream
// reader operate normally. Corresponds to KNOWN_ISSUES §9.
internal static class WordStrictAttributeSanitizer
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly HashSet<string> OnOffValid = new(StringComparer.OrdinalIgnoreCase)
        { "true", "false", "on", "off", "0", "1" };

    // Elements whose `w:val` attribute is an OnOff. Invalid values → strip val
    // (the element's mere presence means "true", matching Word's behavior).
    private static readonly HashSet<string> OnOffElements = new(StringComparer.Ordinal)
    {
        "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
        "vanish", "specVanish", "webHidden", "noProof",
        "emboss", "imprint", "outline", "shadow", "snapToGrid",
        "contextualSpacing", "kinsoku", "overflowPunct", "topLinePunct",
        "autoSpaceDE", "autoSpaceDN", "wordWrap",
        "suppressAutoHyphens", "suppressLineNumbers", "suppressOverlap",
        "widowControl", "keepNext", "keepLines", "pageBreakBefore",
        "hidden", "cantSplit", "tblHeader",
        "bookFoldPrinting", "bookFoldRevPrinting",
        "evenAndOddHeaders", "titlePg",
    };

    // Elements whose `w:val` is an enum. Invalid values → strip the whole
    // element (default behavior of the parent kicks in).
    private static readonly Dictionary<string, HashSet<string>> EnumElements = new(StringComparer.Ordinal)
    {
        ["jc"] = new(StringComparer.Ordinal)
        {
            "left", "center", "right", "both", "start", "end",
            "distribute", "mediumKashida", "lowKashida", "highKashida",
            "thaiDistribute", "numTab",
        },
        ["vAlign"] = new(StringComparer.Ordinal) { "top", "center", "bottom", "both" },
        ["textDirection"] = new(StringComparer.Ordinal)
            { "lrTb", "tbRl", "btLr", "lrTbV", "tbRlV", "tbLrV", "rl", "lr" },
    };

    public static void Sanitize(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main == null) return;

        // PERF(sanitize-fast-path): each entry probes the part's RAW XML —
        // read straight from the backing stream, BEFORE the SDK materializes
        // the DOM — and only falls through to the full typed walk when the
        // probe actually finds an invalid attribute value (rare; real files
        // are clean). The probe is a pure filter: a false positive (pattern
        // matched inside a comment or innocent text) merely triggers the
        // original full walk, so behavior is unchanged; a false negative is
        // impossible because the DOM is parsed from exactly these bytes, so
        // any real offending attribute exists in the raw text.
        TrySanitize(() => main.Document, main);
        TrySanitize(() => main.StyleDefinitionsPart?.Styles, main.StyleDefinitionsPart);
        TrySanitize(() => main.NumberingDefinitionsPart?.Numbering, main.NumberingDefinitionsPart);
        TrySanitize(() => main.FootnotesPart?.Footnotes, main.FootnotesPart);
        TrySanitize(() => main.EndnotesPart?.Endnotes, main.EndnotesPart);
        TrySanitize(() => main.DocumentSettingsPart?.Settings, main.DocumentSettingsPart);
        foreach (var h in main.HeaderParts) TrySanitize(() => h.Header, h);
        foreach (var f in main.FooterParts) TrySanitize(() => f.Footer, f);
    }

    private static void TrySanitize(Func<OpenXmlPartRootElement?> getRoot, OpenXmlPart? part)
    {
        if (part != null && !PartNeedsSanitize(part)) return;
        OpenXmlPartRootElement? root;
        try { root = getRoot(); }
        catch { return; }
        if (root != null) SanitizePart(root);
    }

    private static bool PartNeedsSanitize(OpenXmlPart part)
    {
        // One alternation regex for the shared OnOff whitelist, one per enum
        // element name (each has its own valid set). Tag can be `<w:b/>`
        // (attribute absent — untouched by the DOM pass anyway), or carry
        // w:val in any attribute position.
        try
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream);
            var raw = reader.ReadToEnd();
            foreach (Match m in OnOffBadValProbe.Matches(raw))
                if (!OnOffValid.Contains(m.Groups[1].Value))
                    return true;
            foreach (var kv in EnumElements)
            {
                var probe = s_enumProbes[kv.Key];
                foreach (Match m in probe.Matches(raw))
                    if (!kv.Value.Contains(m.Groups[1].Value))
                        return true;
            }
            return false;
        }
        catch
        {
            return true; // can't probe → do the safe full walk
        }
    }

    private static readonly Regex OnOffBadValProbe = NamesRegex(string.Join("|", OnOffElements));
    private static readonly Dictionary<string, Regex> s_enumProbes =
        EnumElements.ToDictionary(kv => kv.Key, kv => NamesRegex(kv.Key), StringComparer.Ordinal);

    private static Regex NamesRegex(string names)
    {
        // Any `<w:name>` tag carrying a w:val attribute (any attribute order,
        // covers self-closing and plain forms). Group 1 = the val literal.
        return new Regex(
            $"<w:(?:{string.Join('|', names)})" +
            "(?:\\s[^>]*?)?\\s+[^>]*?\\bw:val=\"([^\"]*)\"",
            RegexOptions.Compiled | RegexOptions.ExplicitCapture);
    }

    private static void SanitizePart(OpenXmlPartRootElement root)
    {
        var nodes = root.Descendants<OpenXmlElement>().ToList();
        var toRemove = new List<OpenXmlElement>();

        foreach (var elem in nodes)
        {
            if (elem.NamespaceUri != W) continue;
            var name = elem.LocalName;

            if (OnOffElements.Contains(name))
            {
                var raw = ReadValAttribute(elem);
                if (raw != null && !OnOffValid.Contains(raw))
                {
                    // Strip val — bare element = true, matching Word's
                    // lenient handling of `<w:b w:val="yes"/>`.
                    elem.RemoveAttribute("val", W);
                }
            }
            else if (EnumElements.TryGetValue(name, out var valid))
            {
                var raw = ReadValAttribute(elem);
                if (raw != null && !valid.Contains(raw))
                {
                    toRemove.Add(elem);
                }
            }
        }

        foreach (var elem in toRemove)
        {
            elem.Parent?.RemoveChild(elem);
        }
    }

    private static string? ReadValAttribute(OpenXmlElement elem)
    {
        foreach (var a in elem.GetAttributes())
        {
            if (a.LocalName == "val" && (a.NamespaceUri == W || a.NamespaceUri == ""))
                return a.Value;
        }
        return null;
    }
}
