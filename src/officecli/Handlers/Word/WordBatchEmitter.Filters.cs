// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
        // Read-side convenience derived from a mermaid picture's alt-text
        // (`alt=mermaid:<source>`). The alt already carries + round-trips the
        // source, so re-emitting `mermaid=<source>` too would double the payload
        // and add an inert prop AddPicture ignores. Storage is alt; drop the mirror.
        "mermaid",
        // Comment resolved-state (done) + reply-parent (parentId) are readback
        // keys backed by word/commentsExtended.xml, which the dump round-trips
        // verbatim via a raw `/commentsExtended replace`. Emitting them as typed
        // `add comment` props too would double-apply and break the dump
        // fixed-point — the raw replace is the single source of truth.
        "done", "parentId", "resolved",
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
        // BUG-DUMP-BMSPAN: internal flag set by BookmarkStartToNode marking a
        // content-wrapping bookmark. Consumed by TryEmitBookmarkRun (translated
        // to open=true) / EmitBody only; never replayed verbatim as a property.
        "_spanOpen",
        // BUG-R12A(BUG1): synthetic flag set by CoalesceHyperlinkRuns to route a
        // multi-run / formatted hyperlink group through structured emit. Consumed
        // by EmitPlainOrHyperlinkRun only; never replayed as an Add/Set property.
        "_hlStructured",
        // BUG-DUMP-PGNUM: internal flag set by RunToNode when a run contains
        // <w:pgNum/>. Consumed by TryEmitPgNumRun only (routes the run to a
        // verbatim raw-set passthrough); never replayed as an Add/Set property.
        "_hasPgNum",
        // BUG-DUMP-DATEFIELD: internal flag set by RunToNode when a run contains
        // a date-component placeholder (<w:dayLong/> etc.). Consumed by
        // TryEmitDateFieldRun only (routes the run to a verbatim raw-set
        // passthrough); never replayed as an Add/Set property.
        "_hasDateField",
        // BUG-DUMP-R47-2: internal flag set by RunToNode when a run contains
        // <w:softHyphen/>/<w:noBreakHyphen/>. Consumed by TryEmitHyphenRun; the
        // /body path raw-sets the verbatim run, the header/footer/cell path now
        // emits the run text (glyph-degraded) through EmitPlainOrHyperlinkRun —
        // which runs FilterEmittableProps, so the marker must be stripped here.
        "_hasHyphen",
        // BUG-DUMP-R35-2: internal flag set by Navigation on a run synthesized
        // from inside a <w:smartTag>/<w:customXml> wrapper. Consumed by
        // EmitPlainOrHyperlinkRun (drives the deterministic "wrapper flattened"
        // warning); the inner run text/formatting is preserved but the wrapper
        // element is dropped. Never replayed as an Add/Set property.
        "_wrapperFlattened",
        // BUG-DUMP26-01: Navigation stamps this flag when numId/numLevel come
        // from ResolveNumPrFromStyle (paragraph inherits numbering through its
        // style). EmitParagraph consumes the flag to drop the inherited
        // numId/numLevel/numFmt/listStyle/start before they ride on `add p`.
        // Drop the flag itself from any emitted prop bag.
        "numInherited",
        // BUG-DUMP-R26-2: internal flags set by CollapseFieldChains when a
        // field's cached result has multiple distinctly-formatted runs. Consumed
        // by TryEmitFieldRun (routes the field to a verbatim raw-set chain);
        // never replayed as an Add/Set property.
        "_richFieldResult", "_fieldSlicePaths",
        // BUG-DUMP-H78: internal flag forcing the field-slice raw-set to use the
        // contiguous sibling-range extractor (captures a <w:del> wrapper inside a
        // live field result). Consumed by TryEmitFieldRun; never replayed.
        "_fieldSliceForceRange",
        // BUG-DUMP-R26-7: flag set when a field cached result wraps a hyperlink
        // (external rel) the typed path can't preserve — drives a deterministic
        // warning in TryEmitFieldRun. Never replayed as an Add/Set property.
        "_fieldResultHasExternalRel",
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

    // Shared allowlist for forwarding a note/comment FIRST paragraph's direct
    // paragraph-level formatting onto its `add footnote|endnote|comment` op.
    // Both EmitNoteReference (footnote/endnote) and the comment emit used to
    // carry byte-identical copies of this switch; they diverged only on numPr:
    // notes rebuild a list item via AddFootnote/AddEndnote, so they forward
    // numId/numLevel (allowNumPr=true), while AddComment has no numPr rebuild
    // path, so comments keep them out (allowNumPr=false). Callers still own the
    // numInherited guard and the !props.ContainsKey dedupe.
    // BUG-DUMP-NOTE-PBDR / -PPR-SWEEP / -NUMPR consolidated here.
    private static bool IsForwardableNoteFirstParaKey(string k, bool allowNumPr)
    {
        if (k.StartsWith("markRPr.", StringComparison.OrdinalIgnoreCase)
            || k.StartsWith("pbdr.", StringComparison.OrdinalIgnoreCase))
            return true;
        switch (k)
        {
            case "shading": case "shd":
            case "lineSpacing": case "lineRule": case "spaceBefore": case "spaceAfter":
            case "spaceBeforeLines": case "spaceAfterLines": case "alignment": case "align":
            case "direction": case "leftIndent": case "rightIndent": case "firstLine":
            case "indent": case "firstLineIndent": case "hangingIndent":
            case "hanging": case "contextualSpacing": case "spaceBeforeAuto": case "spaceAfterAuto":
            case "keepNext": case "keepLines": case "pageBreakBefore": case "widowControl":
            case "suppressLineNumbers": case "suppressAutoHyphens": case "suppressOverlap":
            case "kinsoku": case "wordWrap": case "overflowPunct": case "topLinePunct":
            case "autoSpaceDE": case "autoSpaceDN": case "adjustRightInd": case "snapToGrid":
            case "mirrorIndents": case "textAlignment": case "outlineLvl": case "textboxTightWrap":
                return true;
            // notes-only: AddFootnote/AddEndnote rebuild a direct <w:numPr>; a
            // comment's apply path has no equivalent, so it stays opt-in.
            case "numId": case "numLevel":
                return allowNumPr;
            default:
                return false;
        }
    }

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // CONSISTENCY(border-fold): Get emits `pbdr.bottom: single`,
        // `pbdr.bottom.sz: 6`, `pbdr.bottom.color: #FF0000`, `pbdr.bottom.space: 1`
        // as separate keys (mirrors `border.*` on Excel). Set accepts a single
        // semicolon-encoded value `pbdr.bottom=single;6;#FF0000;1`. Without folding,
        // the 2-segment key applies an empty-style border and the 3-segment
        // subkeys hit unsupported (BUG BT-6: Title/Intense Quote lose bottom
        // border on round-trip). Fold the 4 keys into one before validation.
        // BUG-DUMP-R36-1: fold tuple carries shadow/frame so the compound
        // string can append them as segments 5/6 (STYLE;SIZE;COLOR;SPACE;SHADOW;FRAME).
        var pbdrFold = new Dictionary<string, BorderFold>(
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
                    case "shadow": cur.shadow = sval; break;
                    case "frame": cur.frame = sval; break;
                    // BUG-DUMP-R41-2: theme linkage sub-keys (ReadBorder emits
                    // .themeColor / .themeShade / .themeTint).
                    case "themecolor": cur.themeColor = sval; break;
                    case "themeshade": cur.themeShade = sval; break;
                    case "themetint": cur.themeTint = sval; break;
                }
            }
            pbdrFold[side] = cur;
        }

        // BUG-X7-04: same fold for table `border.*` keys. Get emits
        // `border.top: single`, `border.top.sz: 12`, `border.top.color: #000000`
        // separately; Set accepts only the semicolon-encoded form
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
        var borderFold = new Dictionary<string, BorderFold>(
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
                    case "shadow": cur.shadow = sval; break;
                    case "frame": cur.frame = sval; break;
                    // BUG-DUMP-R41-2: theme linkage sub-keys (ReadBorder emits
                    // .themeColor / .themeShade / .themeTint).
                    case "themecolor": cur.themeColor = sval; break;
                    case "themeshade": cur.themeShade = sval; break;
                    case "themetint": cur.themeTint = sval; break;
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
            // BUG-DUMP-R41-4: theme-linkage attrs surfaced by ReadShadingTheme.
            string? sThemeFill = null, sThemeFillShade = null, sThemeFillTint = null;
            string? sThemeColor = null, sThemeShade = null, sThemeTint = null;
            foreach (var (k, v) in raw)
            {
                if (v == null) continue;
                if (string.Equals(k, "shading.val", StringComparison.OrdinalIgnoreCase)) sVal = v.ToString();
                else if (string.Equals(k, "shading.fill", StringComparison.OrdinalIgnoreCase)) sFill = v.ToString();
                else if (string.Equals(k, "shading.color", StringComparison.OrdinalIgnoreCase)) sColor = v.ToString();
                else if (string.Equals(k, "shading.themeFill", StringComparison.OrdinalIgnoreCase)) sThemeFill = v.ToString();
                else if (string.Equals(k, "shading.themeFillShade", StringComparison.OrdinalIgnoreCase)) sThemeFillShade = v.ToString();
                else if (string.Equals(k, "shading.themeFillTint", StringComparison.OrdinalIgnoreCase)) sThemeFillTint = v.ToString();
                else if (string.Equals(k, "shading.themeColor", StringComparison.OrdinalIgnoreCase)) sThemeColor = v.ToString();
                else if (string.Equals(k, "shading.themeShade", StringComparison.OrdinalIgnoreCase)) sThemeShade = v.ToString();
                else if (string.Equals(k, "shading.themeTint", StringComparison.OrdinalIgnoreCase)) sThemeTint = v.ToString();
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
            bool anyTheme = sThemeFill != null || sThemeFillShade != null || sThemeFillTint != null
                || sThemeColor != null || sThemeShade != null || sThemeTint != null;
            if (sVal != null || sFill != null || sColor != null || anyTheme)
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
                // BUG-DUMP-R41-4: append theme-linkage as backward-compatible
                // `key=val` tail segments. ParseShadingValue (Set side) strips
                // any `=`-bearing segment via ExtractThemeTail, so a non-themed
                // shading keeps the exact legacy VAL;FILL[;COLOR] shape.
                if (sThemeFill != null) shadingFolded += $";themeFill={sThemeFill}";
                if (sThemeFillShade != null) shadingFolded += $";themeFillShade={sThemeFillShade}";
                if (sThemeFillTint != null) shadingFolded += $";themeFillTint={sThemeFillTint}";
                if (sThemeColor != null) shadingFolded += $";themeColor={sThemeColor}";
                if (sThemeShade != null) shadingFolded += $";themeShade={sThemeShade}";
                if (sThemeTint != null) shadingFolded += $";themeTint={sThemeTint}";
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
        // lineSpacing="0pt" (w:line=0) now round-trips: with
        // lineRule=atLeast it means "no minimum line height" and dropping it
        // re-rendered those paragraphs at the style default height. Only the
        // degenerate multiplier forms (0x/0%) are still dropped — those have
        // no defined rendering.
        bool dropLineSpacingZero = false;
        if (raw.TryGetValue("lineSpacing", out var lsVal) && lsVal is string lsStr)
        {
            var t = lsStr.Trim();
            if (t == "0x" || t == "0%")
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

            // pbdr fold: skip subkeys, rewrite the bare side key into semicolon form.
            if (key.StartsWith("pbdr.", StringComparison.OrdinalIgnoreCase))
            {
                var parts = key.Split('.');
                if (parts.Length >= 3) continue; // subkey already folded
                var side = $"{parts[0]}.{parts[1]}";
                if (pbdrFold.TryGetValue(side, out var folded) && folded.style != null)
                {
                    result[key] = FoldBorderValue(folded);
                }
                continue;
            }

            // BUG-X7-04: fold border.* like pbdr.*. Skip the 3-segment subkeys
            // (folded into the 2-segment side key below) and rewrite the bare
            // side key into the semicolon-encoded form Set's ParseBorderValue
            // expects.
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase))
            {
                var bparts = key.Split('.');
                if (bparts.Length >= 3) continue; // subkey already folded
                var bside = $"{bparts[0]}.{bparts[1]}";
                if (borderFold.TryGetValue(bside, out var folded) && folded.style != null)
                {
                    result[key] = FoldBorderValue(folded);
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

            // A schema-valid but out-of-window font size. <w:sz>/<w:szCs>
            // (CT_HpsMeasure) permit val=0 and very large values, and Word
            // opens such documents; Get surfaces them as e.g. "0pt"/"30000pt".
            // ParseFontSize, however, caps the Add/Set path at [0.5pt, 4000pt]
            // (below 0.5 rounds to a zero half-point; above 4000 overflows the
            // int32 the pptx/word writers cast to). On a dump→batch round-trip
            // the raw value would reach ParseFontSize and throw — and because
            // `add p` is atomic, the WHOLE paragraph (text included) is dropped,
            // so a valid document silently loses content. Clamp the size to the
            // window so the replay stays valid and the run's text survives.
            // (Like the lineSpacing="0pt" drop above this keeps replay
            // parseable; we clamp rather than drop because sz=0 means "zero
            // size", not "inherit default" — clamping preserves the extreme
            // intent, dropping would reset the run to the style's size.)
            if ((string.Equals(key, "size", StringComparison.OrdinalIgnoreCase)
                 || string.Equals(key, "size.cs", StringComparison.OrdinalIgnoreCase))
                && val is string szRaw
                && TryClampFontSizeForEmit(szRaw, out var clampedSize))
            {
                result[key] = clampedSize;
                continue;
            }

            // BUG-DUMPR2-01: a zero-width gridCol/cell (<w:gridCol w:w="0"/> /
            // <w:tcW w:w="0"/>) is legal OOXML — Word emits it for a collapsed
            // column — but the Add/Set width guards reject 0 to catch the
            // layout-corrupting typo case. To keep the round-trip from tripping
            // its own guard (which would drop the whole table and its text),
            // clamp an emitted zero column/cell width up to 1 twip: ~1/1440",
            // visually identical to 0. Cell/gridCol zero-width emits as the
            // explicit "0dxa" form; a table's auto width is bare "0" (type=auto)
            // and is left untouched. Mirrors the out-of-window font-size clamp.
            if (string.Equals(key, "colWidths", StringComparison.OrdinalIgnoreCase)
                && val is string cwRaw && cwRaw.Contains('0'))
            {
                var clamped = string.Join(",", cwRaw.Split(',').Select(part =>
                {
                    var t = part.Trim();
                    var num = t.EndsWith("dxa", StringComparison.OrdinalIgnoreCase) ? t[..^3] : t;
                    return int.TryParse(num, out var n) && n <= 0 ? "1dxa" : t;
                }));
                result[key] = clamped;
                continue;
            }
            if (string.Equals(key, "width", StringComparison.OrdinalIgnoreCase)
                && val is string wRaw && wRaw.Trim() == "0dxa")
            {
                result[key] = "1dxa";
                continue;
            }

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

    // Returns true and sets `clamped` only when `raw` parses to a font size
    // OUTSIDE ParseFontSize's accepted [0.5pt, 4000pt] window; in-window values
    // return false so the original string (including its exact fractional form,
    // e.g. "10.5pt") flows through untouched. Unparseable values also return
    // false — let the normal Add/Set path surface a precise error rather than
    // masking it here. See the call site for why round-trip needs this.
    private static bool TryClampFontSizeForEmit(string raw, out string clamped)
    {
        clamped = "";
        var t = raw.Trim();
        if (t.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            t = t[..^2].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var pt)
            || double.IsNaN(pt) || double.IsInfinity(pt))
            return false;
        if (pt >= 0.5 && pt <= 4000) return false;
        var bounded = pt < 0.5 ? 0.5 : 4000.0;
        clamped = bounded.ToString(System.Globalization.CultureInfo.InvariantCulture) + "pt";
        return true;
    }

    // BUG-DUMP-R36-1: fold a captured border tuple into ParseBorderValue's
    // positional form STYLE[;SIZE[;COLOR[;SPACE[;SHADOW[;FRAME]]]]]. A trailing
    // segment is only emitted when it (or a later segment) is present, so plain
    // borders keep the legacy 4-field (or shorter) shape and never gain a
    // spurious shadow="false"/frame="false" on replay. Empty intermediates keep
    // positional alignment.
    // BUG-DUMP-R41-2: theme linkage (themeColor/themeShade/themeTint) is
    // appended as backward-compatible `key=val` tail segments AFTER the
    // positional fields. The Set-side ExtractThemeTail harvests any `=`-bearing
    // segment, so a value with no theme keys keeps the exact legacy positional
    // shape and a plain border round-trips byte-identically.
    private static string FoldBorderValue(BorderFold f)
    {
        bool hasTheme = f.themeColor != null || f.themeShade != null || f.themeTint != null;
        bool hasSz = f.sz != null, hasCol = f.color != null, hasSp = f.space != null,
             hasSh = f.shadow != null, hasFr = f.frame != null;
        var v = f.style!;
        if (hasSz || hasCol || hasSp || hasSh || hasFr) v += ";" + (f.sz ?? "");
        if (hasCol || hasSp || hasSh || hasFr) v += ";" + (f.color ?? "");
        if (hasSp || hasSh || hasFr) v += ";" + (f.space ?? "");
        if (hasSh || hasFr) v += ";" + (f.shadow ?? "");
        if (hasFr) v += ";" + (f.frame ?? "");
        if (f.themeColor != null) v += ";themeColor=" + f.themeColor;
        if (f.themeShade != null) v += ";themeShade=" + f.themeShade;
        if (f.themeTint != null) v += ";themeTint=" + f.themeTint;
        return v;
    }

    // BUG-DUMP-R41-2: border-side fold tuple. Promoted from an inline tuple
    // type so the theme-linkage slots (themeColor/themeShade/themeTint) can be
    // added without re-spelling the 9-field tuple at every use site.
    private struct BorderFold
    {
        public string? style, sz, color, space, shadow, frame, themeColor, themeShade, themeTint;
    }
}
