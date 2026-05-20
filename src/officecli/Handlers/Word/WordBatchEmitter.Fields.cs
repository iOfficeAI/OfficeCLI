// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    private static List<DocumentNode> CollapseFieldChains(List<DocumentNode> children)
    {
        var result = new List<DocumentNode>();
        for (int i = 0; i < children.Count; i++)
        {
            var c = children[i];
            bool isBegin = c.Type == "fieldChar"
                && c.Format.TryGetValue("fieldCharType", out var fct)
                && string.Equals(fct?.ToString(), "begin", StringComparison.OrdinalIgnoreCase);
            if (!isBegin)
            {
                result.Add(c);
                continue;
            }

            // Walk forward to find instruction text and end marker.
            string instruction = "";
            string display = "";
            bool sawSeparate = false;
            int end = -1;
            for (int j = i + 1; j < children.Count; j++)
            {
                var k = children[j];
                if (k.Type == "instrText")
                {
                    if (k.Format.TryGetValue("instruction", out var iv) && iv != null)
                        instruction += iv.ToString();
                    else if (!string.IsNullOrEmpty(k.Text))
                        instruction += k.Text;
                }
                else if (k.Type == "fieldChar"
                    && k.Format.TryGetValue("fieldCharType", out var ft))
                {
                    var ftStr = ft?.ToString();
                    if (string.Equals(ftStr, "separate", StringComparison.OrdinalIgnoreCase))
                        sawSeparate = true;
                    else if (string.Equals(ftStr, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        end = j;
                        break;
                    }
                }
                else if (k.Type == "run" || k.Type == "r")
                {
                    // Cached display segments after fldChar(separate). Concatenate
                    // their text — formatting on the display run is dropped (the
                    // field renders fresh on replay).
                    if (!string.IsNullOrEmpty(k.Text)) display += k.Text;
                }
            }
            if (end < 0)
            {
                // Malformed (no end marker) — fall back to passing through.
                result.Add(c);
                continue;
            }
            var synth = new DocumentNode
            {
                Path = c.Path,
                Type = "field",
                Text = display,
                Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["instruction"] = instruction.Trim()
                }
            };
            // Source field has no <w:fldChar w:fldCharType="separate"/> — it's
            // the begin+instr+end shape (Word recomputes the result on open).
            // Flag this so EmitField on the field branch can pass `text=""`
            // explicitly to AddField, which short-circuits AddField's default
            // placeholder ("1" for PAGE etc.) and emits the same separator-
            // less shape. Without this flag, the second dump surfaces a
            // phantom `text="1"` key that the source never had.
            if (!sawSeparate)
                synth.Format["_noFieldSeparator"] = true;
            // BUG-DUMP18-02: propagate hyperlink-scope hint from the begin
            // run so the field-emit branch can target the hyperlink parent
            // on replay.
            if (c.Format.TryGetValue("_hyperlinkParent", out var hlp) && hlp != null)
                synth.Format["_hyperlinkParent"] = hlp;
            result.Add(synth);
            i = end;
        }
        return result;
    }

    // Build the prop bag AddField consumes from a parsed field instruction.
    // Returns null when the instruction is empty or its first token is not a
    // known field code; the caller falls back to a plain-text run for the
    // cached display value so the paragraph still renders.
    private static Dictionary<string, string>? BuildFieldAddProps(string instruction, string display)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;
        var trimmed = instruction.Trim();
        // First whitespace-separated token is the field code.
        var firstSpace = trimmed.IndexOfAny(new[] { ' ', '\t' });
        var code = (firstSpace < 0 ? trimmed : trimmed[..firstSpace]).ToUpperInvariant();
        var rest = firstSpace < 0 ? "" : trimmed[(firstSpace + 1)..].Trim();

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["fieldType"] = code
        };
        switch (code)
        {
            case "PAGE":
            case "NUMPAGES":
            case "AUTHOR":
            case "TITLE":
            case "SUBJECT":
            case "FILENAME":
            case "SECTION":
            case "SECTIONPAGES":
                break;
            case "DATE":
            case "TIME":
            case "CREATEDATE":
            case "SAVEDATE":
            case "PRINTDATE":
            {
                // Preserve the `\@ "MMMM d, yyyy"` format switch so dump
                // round-trips Word's locale-formatted date fields. Without
                // this, BuildFieldAddProps dropped `rest` and replay
                // produced a bare DATE field rendered in the default
                // locale (BUG-X6-3). AddField consumes the value via
                // --prop format=…
                var fmtMatch = System.Text.RegularExpressions.Regex.Match(
                    rest ?? "", "\\\\@\\s+\"([^\"]+)\"");
                if (fmtMatch.Success)
                    props["format"] = fmtMatch.Groups[1].Value;
                break;
            }
            case "REF":
            case "PAGEREF":
            case "NOTEREF":
            {
                // First arg is the bookmark name (may be quoted).
                var name = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(name)) return null;
                props["bookmarkName"] = name;
                break;
            }
            case "SEQ":
            {
                var ident = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(ident)) return null;
                props["identifier"] = ident;
                // BUG-DUMP17-01: preserve trailing switches (\* ARABIC, \r N,
                // \n, \c, \h, \s …). Without this, dump→batch round-trips
                // strip every SEQ formatting switch and replay produces a
                // bare " SEQ Figure ".
                var seqSw = ExtractTrailingSwitches(rest, ident);
                if (!string.IsNullOrEmpty(seqSw)) props["switches"] = seqSw;
                break;
            }
            case "MERGEFIELD":
            {
                var name = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(name)) return null;
                props["fieldName"] = name;
                // BUG-DUMP17-02: preserve trailing switches (\* MERGEFORMAT,
                // \b, \f, \v …). Same shape as the SEQ case above.
                var mfSw = ExtractTrailingSwitches(rest, name);
                if (!string.IsNullOrEmpty(mfSw)) props["switches"] = mfSw;
                break;
            }
            case "HYPERLINK":
            {
                // BUG-DUMP15-02: HYPERLINK may carry any combination of a base
                // URL, `\l "anchor"`, and `\o "tooltip"`. The previous code
                // checked `\l` first and returned only the anchor, dropping
                // the URL entirely; `\o` was never parsed. Parse all three
                // independently so dump→batch round-trips preserve them.
                // The first non-switch token (if any) is the base URL.
                var restStr = rest ?? "";
                if (!System.Text.RegularExpressions.Regex.IsMatch(restStr.TrimStart(), @"^\\"))
                {
                    var url = ExtractFirstArg(restStr);
                    if (!string.IsNullOrEmpty(url)) props["url"] = url;
                }
                var anchorMatch = System.Text.RegularExpressions.Regex.Match(restStr, "\\\\l\\s+\"([^\"]+)\"");
                if (anchorMatch.Success) props["anchor"] = anchorMatch.Groups[1].Value;
                var tooltipMatch = System.Text.RegularExpressions.Regex.Match(restStr, "\\\\o\\s+\"([^\"]+)\"");
                if (tooltipMatch.Success) props["tooltip"] = tooltipMatch.Groups[1].Value;
                if (!props.ContainsKey("url") && !props.ContainsKey("anchor"))
                    return null;
                break;
            }
            default:
                // BUG-DUMP7-05: AddField's switch has no case for `=`,
                // numeric expression fields like `= PAGE - 1`, or any other
                // unrecognised code. Emitting fieldType=<code> would make
                // replay throw `Unknown field type '<code>'`. Drop the
                // unhelpful fieldType and pass the full trimmed instruction
                // through `instr` instead — AddField's raw-instruction
                // fallback rebuilds the chain verbatim. Drops `fieldType`
                // entirely so the caller doesn't reject the row up-front.
                props.Remove("fieldType");
                props["instr"] = trimmed;
                break;
        }
        if (!string.IsNullOrEmpty(display))
            props["text"] = display;
        return props;
    }

    private static string ExtractFirstArg(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var t = s.TrimStart();
        if (t.StartsWith('"'))
        {
            var end = t.IndexOf('"', 1);
            return end > 0 ? t[1..end] : "";
        }
        var spc = t.IndexOfAny(new[] { ' ', '\t' });
        return spc < 0 ? t : t[..spc];
    }

    // Return the portion of `s` that follows the first arg (which
    // ExtractFirstArg already returned), trimmed. Used by SEQ /
    // MERGEFIELD field parsing to preserve trailing switches like
    // `\* ARABIC \r N` or `\* MERGEFORMAT` so AddField can replay them
    // verbatim. BUG-DUMP17-01 / BUG-DUMP17-02.
    private static string ExtractTrailingSwitches(string? s, string firstArg)
    {
        if (string.IsNullOrEmpty(s) || string.IsNullOrEmpty(firstArg)) return "";
        var t = s.TrimStart();
        int consumed;
        if (t.StartsWith('"'))
        {
            var end = t.IndexOf('"', 1);
            if (end < 0) return "";
            consumed = end + 1;
        }
        else
        {
            consumed = firstArg.Length;
        }
        return consumed >= t.Length ? "" : t[consumed..].Trim();
    }

    // Parse a TOC field instruction (` TOC \o "1-3" \h \u \z `) into the
    // prop bag AddToc accepts. AddToc emits the canonical instruction so
    // round-tripping the parsed props back through it lands at the same
    // OOXML even when the source instruction had extra whitespace or
    // switch ordering.
    private static Dictionary<string, string> ParseTocInstruction(string instruction)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var lvl = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\o\\s+\"([^\"]+)\"");
        if (lvl.Success) props["levels"] = lvl.Groups[1].Value;
        // \h = hyperlinks (default true on AddToc, but emit explicitly for clarity)
        props["hyperlinks"] = System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\h\\b")
            ? "true" : "false";
        // \z suppresses page numbers; absence means pageNumbers=true
        props["pageNumbers"] = System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\z\\b")
            ? "false" : "true";
        // BUG-X5-03: \t = custom-style→level mapping ("Style;level,..."),
        // \b = bookmark scope. Capture the quoted argument so AddToc can
        // round-trip them; otherwise custom TOC switches were silently
        // dropped on dump.
        var ct = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\t\\s+\"([^\"]+)\"");
        if (ct.Success) props["customStyles"] = ct.Groups[1].Value;
        var cb = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\b\\s+\"([^\"]+)\"");
        if (cb.Success) props["bookmark"] = cb.Groups[1].Value;
        return props;
    }
}
