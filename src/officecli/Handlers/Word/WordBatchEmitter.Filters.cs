// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    // Format keys that must NOT be emitted: derived (computed by Get, not
    // user-set), unstable (regenerate on save), or coordinate-system
    // (paths that only make sense in the source document).
    private static readonly HashSet<string> SkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "basedOn.path",
        "paraId", "textId", "rsidR", "rsidRDefault", "rsidRPr", "rsidP", "rsidTr",
        // Paragraph Get emits `style`, `styleId`, and `styleName` — all three
        // carry the same value (style id, repeated). AddParagraph only
        // consumes `style`; emitting the other two would either re-process
        // the same value (no-op) or, if Add ever grows divergent semantics
        // for them, cause double-application. Drop the aliases so the
        // dump bag stays minimal.
        "styleId", "styleName",
        // BUG-DUMP18-02: internal hyperlink-scope hint stamped on runs (and
        // propagated to synthetic field nodes) by Navigation. Consumed by the
        // field-emit branch only; never replayed as a Set/Add property.
        "_hyperlinkParent",
        // BUG-DUMP26-01: Navigation stamps this flag when numId/numLevel come
        // from ResolveNumPrFromStyle (paragraph inherits numbering through its
        // style). EmitParagraph consumes the flag to drop the inherited
        // numId/numLevel/numFmt/listStyle/start before they ride on `add p`.
        // Drop the flag itself from any emitted prop bag.
        "numInherited",
        // Document-internal relationship id (rId4 / X5c0e4d…). Assigned fresh
        // by every Add* path when it creates a new part-relationship, so the
        // value is unstable across replays even when the document is byte-
        // identical otherwise. Pictures, charts, OLE, hyperlinks all emit
        // relId on Get for diagnostics but it must not ride on `add`/`set`.
        "relId",
        // BUG-019: lineSpacing alone cannot distinguish AtLeast from Exact —
        // SpacingConverter.FormatWordLineSpacing serializes both as "Npt".
        // Set/AddParagraph now accept `lineRule` explicitly so it must flow
        // through dump for AtLeast spacing to round-trip without silent
        // downgrade to Exact (which clips tall glyphs).
    };

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // CONSISTENCY(border-fold): Get emits `pbdr.bottom: single`,
        // `pbdr.bottom.sz: 6`, `pbdr.bottom.color: #FF0000`, `pbdr.bottom.space: 1`
        // as separate keys (mirrors `border.*` on Excel). Set accepts a single
        // colon-encoded value `pbdr.bottom=single:6:#FF0000:1`. Without folding,
        // the 2-segment key applies an empty-style border and the 3-segment
        // subkeys hit unsupported (BUG BT-6: Title/Intense Quote lose bottom
        // border on round-trip). Fold the 4 keys into one before validation.
        var pbdrFold = new Dictionary<string, (string? style, string? sz, string? color, string? space)>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (val == null) continue;
            if (!key.StartsWith("pbdr.", StringComparison.OrdinalIgnoreCase)) continue;
            var parts = key.Split('.');
            if (parts.Length < 2) continue;
            var side = $"{parts[0]}.{parts[1]}"; // pbdr.bottom
            pbdrFold.TryGetValue(side, out var cur);
            var sval = val.ToString() ?? "";
            if (parts.Length == 2) cur.style = sval;
            else if (parts.Length == 3)
            {
                switch (parts[2].ToLowerInvariant())
                {
                    case "sz": cur.sz = sval; break;
                    case "color": cur.color = sval; break;
                    case "space": cur.space = sval; break;
                }
            }
            pbdrFold[side] = cur;
        }

        // BUG-X7-04: same fold for table `border.*` keys. Get emits
        // `border.top: single`, `border.top.sz: 12`, `border.top.color: #000000`
        // separately; Set accepts only the colon-encoded form
        // `border.top=single;12;#000000;1`. Without folding, dump strips the
        // 3-segment subkeys (see the explicit "drop them here" comment below)
        // and round-trip silently downgrades real borders to default thin
        // single. Fold sz/color/space into the 2-segment key.
        // BUG-X2-P1-5: Add path now seeds all 6 default borders and overlays
        // user props on top, so a partial spec (e.g. only border.top +
        // border.bottom) replays as 6 single-borders, not 2. Detect a
        // partial spec here and prepend an explicit `border=none` wipe so
        // genuine three-line / banner-line tables round-trip with the same
        // visible result. CONSISTENCY(border-default-overlay).
        var borderFold = new Dictionary<string, (string? style, string? sz, string? color, string? space)>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (val == null) continue;
            if (!key.StartsWith("border.", StringComparison.OrdinalIgnoreCase)) continue;
            var parts = key.Split('.');
            if (parts.Length < 2) continue;
            var side = $"{parts[0]}.{parts[1]}"; // border.top
            borderFold.TryGetValue(side, out var cur);
            var sval = val.ToString() ?? "";
            if (parts.Length == 2) cur.style = sval;
            else if (parts.Length == 3)
            {
                switch (parts[2].ToLowerInvariant())
                {
                    case "sz": cur.sz = sval; break;
                    case "color": cur.color = sval; break;
                    case "space": cur.space = sval; break;
                }
            }
            borderFold[side] = cur;
        }

        // CONSISTENCY(shading-fold): Get surfaces paragraph/run shading as
        // shading.val + shading.fill + shading.color sub-keys (per OOXML
        // attribute decomposition). AddText/AddParagraph accept only a
        // single semicolon-encoded `shading=VAL;FILL[;COLOR]` value. Without
        // folding, the sub-keys hit UNSUPPORTED on `add p` replay and the
        // shading was lost. Fold into a single `shading` key.
        string? shadingFolded = null;
        bool shadingPresent = false;
        {
            string? sVal = null, sFill = null, sColor = null;
            foreach (var (k, v) in raw)
            {
                if (v == null) continue;
                if (string.Equals(k, "shading.val", StringComparison.OrdinalIgnoreCase)) sVal = v.ToString();
                else if (string.Equals(k, "shading.fill", StringComparison.OrdinalIgnoreCase)) sFill = v.ToString();
                else if (string.Equals(k, "shading.color", StringComparison.OrdinalIgnoreCase)) sColor = v.ToString();
            }
            // shading.val="clear" with no fill/color is OOXML's "no shading"
            // form (<w:shd w:val="clear" w:fill="auto"/>). Emitting bare
            // "clear" without semicolons makes the Set/Add color parser
            // treat the whole value as a color name and reject it. Skip
            // the shading emit in this case — semantically identical to
            // the schema default (no shading).
            bool shadingIsEffectivelyNone = sVal != null
                && string.Equals(sVal, "clear", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrEmpty(sFill)
                && string.IsNullOrEmpty(sColor);
            // shadingPresent gates the drop-subkeys loop below. Set true in
            // both the real-shading case and the effectively-none case so
            // the raw `shading.val=clear` etc. don't leak through as
            // UNSUPPORTED top-level props on Add. Only the real-shading
            // case populates shadingFolded; effectively-none emits nothing.
            if (sVal != null || sFill != null || sColor != null)
                shadingPresent = true;
            if (!shadingIsEffectivelyNone && shadingPresent)
            {
                // AddText format: VAL;FILL[;COLOR]. Default val to "clear" when
                // only fill is present (mirrors AddText's single-arg path).
                var val = string.IsNullOrEmpty(sVal) ? "clear" : sVal;
                if (!string.IsNullOrEmpty(sColor))
                    shadingFolded = $"{val};{sFill ?? ""};{sColor}";
                else if (!string.IsNullOrEmpty(sFill))
                    shadingFolded = $"{val};{sFill}";
                else
                    shadingFolded = val;
            }
        }

        // CONSISTENCY(padding-fold): Get surfaces default cell margin as
        // `padding.top/bottom/left/right` on the table node (per-side OOXML
        // attribute decomposition). AddTable accepts only a single `padding`
        // scalar applied uniformly to all four sides. Without folding, every
        // table with non-default cell margin emitted four UNSUPPORTED
        // padding.* keys on `add table`. Fold into a single `padding` when
        // all four sides are equal; otherwise drop (per-side asymmetric
        // padding is a follow-up — AddTable can't express it today).
        string? paddingFolded = null;
        bool paddingFoldable = false;
        {
            string? top = null, bot = null, left = null, right = null;
            foreach (var (k, v) in raw)
            {
                if (v == null) continue;
                if (string.Equals(k, "padding.top", StringComparison.OrdinalIgnoreCase)) top = v.ToString();
                else if (string.Equals(k, "padding.bottom", StringComparison.OrdinalIgnoreCase)) bot = v.ToString();
                else if (string.Equals(k, "padding.left", StringComparison.OrdinalIgnoreCase)) left = v.ToString();
                else if (string.Equals(k, "padding.right", StringComparison.OrdinalIgnoreCase)) right = v.ToString();
            }
            if (top != null && top == bot && top == left && top == right)
            {
                paddingFolded = top;
                paddingFoldable = true;
            }
            // BUG-DUMP5-05: when sides differ we leave paddingFoldable=false
            // so the per-side `padding.top/bottom/left/right` keys flow
            // through the main loop unmodified. `Set tc` consumes per-side
            // padding directly (see WordHandler.Set.Element.cs); only
            // AddTable lacks per-side support, but tables only carry uniform
            // default cell margins on Add — asymmetric tcMar surfaces solely
            // from per-cell `set tc` rows where per-side keys round-trip
            // cleanly. Previously this branch dropped them entirely as
            // UNSUPPORTED, silently losing every asymmetric per-cell margin.
        }

        // <w:spacing w:line="0" w:lineRule="atLeast"> in the source means
        // "no minimum line height" — Word treats it as auto. Get surfaces
        // it as lineSpacing="0pt", but SpacingConverter rejects 0 on the
        // Set/Add path (w:line=0 is undefined OOXML; Word silently single-
        // spaces). Round-trip would fail with "Line spacing must be greater
        // than 0". Drop the zero-value pair on emit so the replayed
        // paragraph/style inherits the carrier's default — same visible
        // result as the source's "no minimum" semantics.
        bool dropLineSpacingZero = false;
        if (raw.TryGetValue("lineSpacing", out var lsVal) && lsVal is string lsStr)
        {
            var t = lsStr.Trim();
            if (t == "0" || t == "0pt" || t == "0.0pt" || t == "0x" || t == "0%")
                dropLineSpacingZero = true;
        }

        foreach (var (key, val) in raw)
        {
            if (SkipKeys.Contains(key)) continue;
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            if (key.EndsWith(".cs.source", StringComparison.OrdinalIgnoreCase)) continue;

            // lineSpacing="0pt" companion drop — see fold comment above the loop.
            if (dropLineSpacingZero &&
                (string.Equals(key, "lineSpacing", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(key, "lineRule", StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            // padding.* fold: drop sub-keys; emit single `padding` if uniform.
            if (paddingFoldable && key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // shading.* fold: drop sub-keys; emit single `shading` below.
            if (shadingPresent && key.StartsWith("shading.", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            // pbdr fold: skip subkeys, rewrite the bare side key into colon form.
            if (key.StartsWith("pbdr.", StringComparison.OrdinalIgnoreCase))
            {
                var parts = key.Split('.');
                if (parts.Length >= 3) continue; // subkey already folded
                var side = $"{parts[0]}.{parts[1]}";
                if (pbdrFold.TryGetValue(side, out var folded) && folded.style != null)
                {
                    // ParseBorderValue format: STYLE[;SIZE[;COLOR[;SPACE]]] — empties
                    // for missing intermediates so positional parts stay aligned.
                    var sz = folded.sz ?? "";
                    var col = folded.color ?? "";
                    var sp = folded.space ?? "";
                    var v = folded.style!;
                    if (folded.sz != null || folded.color != null || folded.space != null)
                        v += ";" + sz;
                    if (folded.color != null || folded.space != null)
                        v += ";" + col;
                    if (folded.space != null)
                        v += ";" + sp;
                    result[key] = v;
                }
                continue;
            }

            // BUG-X7-04: fold border.* like pbdr.*. Skip the 3-segment subkeys
            // (folded into the 2-segment side key below) and rewrite the bare
            // side key into the colon-encoded form Set's ParseBorderValue
            // expects.
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase))
            {
                var bparts = key.Split('.');
                if (bparts.Length >= 3) continue; // subkey already folded
                var bside = $"{bparts[0]}.{bparts[1]}";
                if (borderFold.TryGetValue(bside, out var folded) && folded.style != null)
                {
                    var sz = folded.sz ?? "";
                    var col = folded.color ?? "";
                    var sp = folded.space ?? "";
                    var v = folded.style!;
                    if (folded.sz != null || folded.color != null || folded.space != null)
                        v += ";" + sz;
                    if (folded.color != null || folded.space != null)
                        v += ";" + col;
                    if (folded.space != null)
                        v += ";" + sp;
                    result[key] = v;
                }
                continue;
            }

            // tabs is a List<Dict>, not a flat scalar. Both Add and Set ingest
            // tab stops via the dedicated `add ... --type tab` command (one
            // row per stop), not as a paragraph/style scalar prop. Skipping
            // here avoids serializing the .NET list type name into the prop
            // string (BUG-X2-01); paragraph emitters layer per-stop add rows
            // separately.
            if (string.Equals(key, "tabs", StringComparison.OrdinalIgnoreCase)) continue;

            if (val == null) continue;
            string s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }
        if (paddingFolded != null && !result.ContainsKey("padding"))
            result["padding"] = paddingFolded;
        if (shadingFolded != null && !result.ContainsKey("shading"))
            result["shading"] = shadingFolded;
        return result;
    }
}
