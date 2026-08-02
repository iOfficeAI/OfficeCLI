// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddParagraph(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // NEWLINE-SEMANTICS-V2: at this block-level surface '\n' is a
        // PARAGRAPH boundary (matching pptx and the Google Docs API); the
        // soft line break <w:br/> is written as '\v'. A text containing
        // '\n' therefore creates one paragraph per line, every line
        // inheriting the same style/format props — Word's Enter key, N
        // times. The properties dictionary is mutated in place per line and
        // restored afterwards (never copied) so TrackingPropertyDictionary
        // keeps recording handler reads for unsupported_property detection.
        if (properties.TryGetValue("text", out var rawParaText) && rawParaText != null)
        {
            var normalized = rawParaText.Replace("\r\n", "\n").Replace("\r", "\n");
            if (normalized.IndexOf('\n') >= 0)
            {
                var lines = normalized.Split('\n');
                string firstPath = "";
                try
                {
                    for (int li = 0; li < lines.Length; li++)
                    {
                        properties["text"] = lines[li];
                        var p = AddParagraph(parent, parentPath, index.HasValue ? index + li : null, properties);
                        if (li == 0) firstPath = p;
                    }
                }
                finally
                {
                    properties["text"] = rawParaText;
                }
                return firstPath;
            }
        }
        // See RejectBareRevisionKey: the bare `revision=` literal was retired
        // when creation/action split into revision.type / revision.action.
        RejectBareRevisionKey(properties);
        string resultPath;
        var para = new Paragraph();
        AssignParaId(para);
        var pProps = new ParagraphProperties();

        // CONSISTENCY(style-dual-key): mirror SetParagraph and AddStyle —
        // accept canonical readback aliases (styleId, styleName) so a
        // get→add clone of a paragraph round-trips its style intact.
        // styleName resolves the display name through the styles part;
        // falls back to verbatim if no match (lenient-input pattern).
        if (properties.TryGetValue("style", out var style)
            || properties.TryGetValue("styleId", out style)
            || properties.TryGetValue("styleid", out style))
        {
            // CONSISTENCY(style-warn): mirror SetParagraph (Set.cs:642) —
            // warn (advisory, non-fatal) when the style id is not defined
            // in the styles part; still store the ref (lenient-input).
            if (!StyleIdExists(style))
                LastAddWarnings.Add($"style '{style}' not found in styles part — will be referenced as-is");
            pProps.ParagraphStyleId = new ParagraphStyleId { Val = style };
        }
        else if (properties.TryGetValue("styleName", out var styleName)
            || properties.TryGetValue("stylename", out styleName))
        {
            // Resolve display name through styles part. Fall back to verbatim
            // only when the value is a plausible styleId (no spaces — OOXML
            // styleId disallows spaces). Spaced display names that fail to
            // resolve are skipped + warned rather than stored as invalid id.
            var resolved = ResolveStyleIdFromName(styleName);
            if (resolved != null)
            {
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = resolved };
            }
            else if (!styleName.Contains(' '))
            {
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = styleName };
            }
            else
            {
                LastAddWarnings.Add($"styleName '{styleName}' not found in styles part and contains spaces — skipped (OOXML styleId disallows spaces)");
            }
        }
        if (properties.TryGetValue("align", out var alignment) || properties.TryGetValue("alignment", out alignment) || properties.TryGetValue("jc", out alignment))
            pProps.Justification = new Justification { Val = ParseJustification(alignment) };
        // textAlignment (ST_TextAlignment) is the vertical baseline alignment.
        // Curate here so invalid values like "justified" throw upfront rather
        // than slipping through TypedAttributeFallback as schema-invalid XML.
        // Mirrors the Set case (Set.cs:1110).
        if (properties.TryGetValue("textAlignment", out var txtAlign)
            || properties.TryGetValue("textalignment", out txtAlign))
        {
            pProps.TextAlignment = new TextAlignment
            {
                Val = txtAlign.ToLowerInvariant() switch
                {
                    "auto"     => VerticalTextAlignmentValues.Auto,
                    "top"      => VerticalTextAlignmentValues.Top,
                    "center"   => VerticalTextAlignmentValues.Center,
                    "baseline" => VerticalTextAlignmentValues.Baseline,
                    "bottom"   => VerticalTextAlignmentValues.Bottom,
                    _ => throw new ArgumentException($"Invalid 'textAlignment' value: '{txtAlign}'. Valid: auto, top, center, baseline, bottom."),
                },
            };
        }
        // textboxTightWrap (ST_TextboxTightWrap): controls how text wraps
        // around the floating textbox. Curate so invalid input throws
        // upfront (the generic TypedAttributeFallback would silently store
        // a bogus string as an extension attribute).
        if (properties.TryGetValue("textboxTightWrap", out var tbtw)
            || properties.TryGetValue("textboxtightwrap", out tbtw))
        {
            pProps.TextBoxTightWrap = new TextBoxTightWrap
            {
                Val = tbtw.ToLowerInvariant() switch
                {
                    "none"             => TextBoxTightWrapValues.None,
                    "alllines"         => TextBoxTightWrapValues.AllLines,
                    "firstandlastline" => TextBoxTightWrapValues.FirstAndLastLine,
                    "firstlineonly"    => TextBoxTightWrapValues.FirstLineOnly,
                    "lastlineonly"     => TextBoxTightWrapValues.LastLineOnly,
                    _ => throw new ArgumentException($"Invalid 'textboxTightWrap' value: '{tbtw}'. Valid: none, allLines, firstAndLastLine, firstLineOnly, lastLineOnly."),
                },
            };
        }
        // Reading direction (Arabic / Hebrew). 'rtl' enables <w:bidi/> AND
        // writes <w:rtl/> on the paragraph mark (so any later runs added
        // via Set inherit the run-level direction without a separate flag).
        // CONSISTENCY(rtl-cascade): mirrors SetElementParagraph — direction
        // is a paragraph-scope shorthand for "this paragraph is fully RTL".
        // BUG-DUMP-MARKRPR-RTL-OVERSTAMP: when the dump forwards the whole ¶-mark
        // <w:rPr> verbatim (markRPr.xml), that subtree is the authoritative mark —
        // it already carries the source's exact mark-rtl state (present or absent).
        // The direction=rtl / rtl=true convenience cascade below must NOT also
        // inject <w:rtl/> into the mark, or a bidi paragraph whose source mark had
        // no <w:rtl/> (only <w:lang w:bidi=…/>) gains a spurious mark rtl that
        // changes the empty paragraph's line metrics and reflows the page. pPr
        // <w:bidi/> is still set (it is the paragraph direction, separate from the
        // mark and not carried in markRPr.xml). Mirrors the markRPrVerbatimApplied
        // guard on the dotted markRPr.* keys further down.
        bool hasVerbatimMarkRPr = properties.ContainsKey("markRPr.xml")
            || properties.ContainsKey("markrpr.xml");
        bool? paraRtl = null;
        if (properties.TryGetValue("direction", out var dirRaw)
            || properties.TryGetValue("dir", out dirRaw)
            || properties.TryGetValue("bidi", out dirRaw))
        {
            paraRtl = ParseDirectionRtl(dirRaw);
            if (paraRtl.Value)
            {
                // direction/dir/bidi sets ONLY the paragraph-direction flag
                // (pPr <w:bidi/>) — it must NOT inject <w:rtl/> into the ¶-mark
                // rPr. The mark glyph's rtl is an independent property: a bidi
                // paragraph whose source mark legitimately lacks <w:rtl/> (only
                // <w:lang w:bidi=…/>) otherwise gained a spurious mark rtl on
                // dump→batch replay, changing the empty paragraph's line metrics
                // and reflowing the page below it. The dump always carries the
                // mark's true rtl state explicitly (a dotted markRPr.rtl key or
                // the verbatim markRPr.xml subtree), so the mark round-trips
                // faithfully without this coupling. `rtl=true` below stays the
                // explicit "make the mark rtl too" request.
                pProps.BiDi = new BiDi();
            }
            else
            {
                // Clear semantics: direction=ltr removes any prior bidi marker.
                // R19-fuzz-1/2 + R20-fuzz-11: if ANY inherited source carries
                // bidi=true (style chain, enclosing section, docDefaults, or
                // numbering lvl), simply clearing pPr.bidi re-inherits RTL —
                // the user's explicit ltr override would silently disappear.
                // Emit <w:bidi w:val="0"/> to cancel. Style-chain check happens
                // here (no parent context needed); section / docDefaults /
                // numbering checks are deferred until after the paragraph is
                // inserted into the tree (see post-insert HasInheritedBidi
                // pass below). Mirrors paragraph Set/ApplyDirectionCascade.
                pProps.RemoveAllChildren<BiDi>();
                // CONSISTENCY(bidi-explicit-false-roundtrip): Navigation emits
                // `direction=ltr` ONLY when source pPr had an explicit
                // <w:bidi w:val="0"/>. Always stamp the explicit override on
                // replay so dump→batch preserves the source's literal pPr
                // shape — not just the subset where style-chain inheritance
                // would otherwise re-enable RTL.
                pProps.BiDi = new BiDi { Val = new DocumentFormat.OpenXml.OnOffValue(false) };
                var markRPr = pProps.ParagraphMarkRunProperties;
                markRPr?.RemoveAllChildren<RightToLeftText>();
            }
        }
        // CONSISTENCY(rtl-cascade): `rtl=true` on a paragraph add should
        // mirror direction=rtl — write <w:bidi/> on pPr AND <w:rtl/> on
        // the paragraph mark so the paragraph is fully RTL (not just any
        // text run). Without this, `add p --prop rtl=true` left the
        // paragraph LTR and only flagged individual runs.
        if (paraRtl == null && properties.TryGetValue("rtl", out var paraRtlRaw) && IsTruthy(paraRtlRaw))
        {
            paraRtl = true;
            pProps.BiDi = new BiDi();
            if (!hasVerbatimMarkRPr)
            {
                var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                ApplyRunFormatting(markRPr, "rtl", "true");
            }
        }
        // Complex-script run flags (bCs/iCs/szCs) hoisted above the text
        // block so an `add p --prop bold.cs=true` without explicit text
        // still records the flag on the paragraph mark rPr — matches how
        // bare bold round-trips via the generic TypedAttributeFallback
        // path. Without this, schema-strict round-trip tests for
        // bold.cs/italic.cs/size.cs lose the flag (no run carrier exists
        // when text is absent, and TypedAttributeFallback can't synthesise
        // <w:bCs/> / <w:iCs/> / <w:szCs/> child elements from a key).
        if ((properties.TryGetValue("bold.cs", out var paraBoldCs)
                || properties.TryGetValue("font.bold.cs", out paraBoldCs)))
        {
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "bold.cs", paraBoldCs);
        }
        if ((properties.TryGetValue("italic.cs", out var paraItalicCs)
                || properties.TryGetValue("font.italic.cs", out paraItalicCs)))
        {
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "italic.cs", paraItalicCs);
        }
        if (properties.TryGetValue("size.cs", out var paraSizeCs)
            || properties.TryGetValue("font.size.cs", out paraSizeCs))
        {
            var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(markRPr, "size.cs", paraSizeCs);
        }
        // BUG-R7-07: when the paragraph has no `text` prop, no run is created
        // — yet style-overriding run-level props (size, italic=false,
        // bold=false, color, font.* …) must still ride on the paragraph mark
        // rPr so they survive the next dump. Without this hoist, dump→batch
        // round-trip silently drops the override and the style's defaults
        // re-emerge (e.g. `style=TOC2 size=11pt` → 12pt because TOC2's
        // base size is 12pt). Mirrors the size.cs/italic.cs/bold.cs hoist
        // above. Only applied when there is no text run carrier.
        // BUG-DUMP-R44-3: the plain `shading`/`shd` key is always a
        // paragraph-level pPr/shd (whole-line banner) — never hoisted to the
        // ¶-mark rPr (that's what the explicit `markRPr.shading` key is for).
        // This flag therefore stays false; kept only so the pPr-level shading
        // guard below reads as an intentional "not consumed by a mark hoist".
        bool shadingHoistedToMarkRPr = false;
        if (!properties.ContainsKey("text"))
        {
            ParagraphMarkRunProperties? noTextMarkRPr = null;
            ParagraphMarkRunProperties EnsureNoTextMarkRPr() =>
                noTextMarkRPr ??= (pProps.ParagraphMarkRunProperties
                    ?? pProps.AppendChild(new ParagraphMarkRunProperties()));
            if (properties.TryGetValue("size", out var ntSize)
                || properties.TryGetValue("font.size", out ntSize)
                || properties.TryGetValue("fontsize", out ntSize))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "size", ntSize);
            // BUG-R7-07 / F-7: explicit `false` must produce <w:b w:val="false"/>
            // (resp. <w:i w:val="false"/>) so it overrides a style that sets
            // bold/italic=true. ApplyRunFormatting on its own removes the
            // element entirely on a falsy value — that contract is preserved
            // for the Set-after-create call sites (existing R25/R26 tests
            // depend on it). Only the Add path needs the explicit-override
            // semantics, so emit the val=false form directly here.
            if (properties.TryGetValue("bold", out var ntBold)
                || properties.TryGetValue("font.bold", out ntBold))
            {
                var rp = EnsureNoTextMarkRPr();
                rp.RemoveAllChildren<Bold>();
                if (IsTruthy(ntBold))
                    InsertRunPropInSchemaOrder(rp, new Bold());
                else if (IsExplicitFalseAddOverride(ntBold))
                    InsertRunPropInSchemaOrder(rp, new Bold { Val = OnOffValue.FromBoolean(false) });
            }
            if (properties.TryGetValue("italic", out var ntItalic)
                || properties.TryGetValue("font.italic", out ntItalic))
            {
                var rp = EnsureNoTextMarkRPr();
                rp.RemoveAllChildren<Italic>();
                if (IsTruthy(ntItalic))
                    InsertRunPropInSchemaOrder(rp, new Italic());
                else if (IsExplicitFalseAddOverride(ntItalic))
                    InsertRunPropInSchemaOrder(rp, new Italic { Val = OnOffValue.FromBoolean(false) });
            }
            if (properties.TryGetValue("color", out var ntColor)
                || properties.TryGetValue("font.color", out ntColor))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "color", ntColor);
            if (properties.TryGetValue("highlight", out var ntHighlight))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "highlight", ntHighlight);
            // BUG-DUMP-R44-3: the plain `shading`/`shd` key is a PARAGRAPH-level
            // property (a DIRECT <w:pPr><w:shd> that paints the whole-line
            // banner background) — it must NOT be hoisted onto the ¶-mark rPr
            // even on a no-text paragraph (e.g. an empty full-width banner bar).
            // The prior R27-1 hoist here pushed a genuine paragraph shading DOWN
            // into <w:pPr><w:rPr><w:shd>, which only colors the invisible pilcrow
            // and made the banner bar disappear on round-trip. A genuine ¶-mark
            // character shading is emitted by Navigation under the EXPLICIT
            // `markRPr.shading` key (read from <w:pPr><w:rPr><w:shd/>), which
            // routes through the markRPr.* branch below — so mark-only shading is
            // still preserved without conflating it with the paragraph banner.
            // Therefore: let the plain `shading`/`shd` key fall through to the
            // pPr-level pProps.Shading handler (~line 388); never set
            // shadingHoistedToMarkRPr here.
            if (properties.TryGetValue("underline", out var ntUl)
                || properties.TryGetValue("font.underline", out ntUl))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "underline", ntUl);
            if (properties.TryGetValue("strike", out var ntStrike)
                || properties.TryGetValue("font.strike", out ntStrike)
                || properties.TryGetValue("strikethrough", out ntStrike)
                || properties.TryGetValue("font.strikethrough", out ntStrike))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "strike", ntStrike);
            // BUG-DUMP-R62-MARKVANISH: <w:vanish/> on a no-text paragraph's ¶-mark
            // hides the pilcrow, collapsing the empty paragraph to zero height —
            // the mechanism Word uses for the hidden spacer between two adjacent
            // tables (keeps them flush). The run path applies vanish (~line 851)
            // but a no-text paragraph never enters it, so the hidden spacer
            // re-rendered visible on replay and opened a gap that reflowed the
            // document. Route it onto the ¶ mark rPr like the other toggles.
            // (vanish is the only character toggle that changes the empty
            // paragraph's height, so it's the one that matters here.)
            if (properties.TryGetValue("vanish", out var ntVanish)
                || properties.TryGetValue("hidden", out ntVanish))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "vanish", ntVanish);
            // Glyph props the ¶-mark readback now surfaces as bare keys on a
            // run-less paragraph (caps family, text effects, vertAlign, bdr,
            // specVanish, position) — route each straight onto the mark rPr
            // via ApplyRunFormatting so dump→batch round-trips them. caps/
            // smallCaps change the empty paragraph's line height; specVanish
            // is Word's style-separator marker.
            foreach (var ntGlyphKey in new[] { "caps", "smallcaps", "dstrike", "outline",
                "shadow", "emboss", "imprint", "superscript", "subscript", "vertalign",
                "bdr", "specVanish", "position" })
            {
                if (properties.TryGetValue(ntGlyphKey, out var ntGlyphVal))
                    ApplyRunFormatting(EnsureNoTextMarkRPr(), ntGlyphKey, ntGlyphVal);
            }
            if (properties.TryGetValue("font", out var ntFont)
                || properties.TryGetValue("font.name", out ntFont))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "font", ntFont);
            if (properties.TryGetValue("font.latin", out var ntFontLatin))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "font.latin", ntFontLatin);
            if (properties.TryGetValue("font.ea", out var ntFontEa)
                || properties.TryGetValue("font.eastasia", out ntFontEa)
                || properties.TryGetValue("font.eastasian", out ntFontEa))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "font.ea", ntFontEa);
            if (properties.TryGetValue("font.cs", out var ntFontCs)
                || properties.TryGetValue("font.complexscript", out ntFontCs)
                || properties.TryGetValue("font.complex", out ntFontCs))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "font.cs", ntFontCs);
            // BUG-DUMP-R31-2: the no-text ¶-mark hoist applied every rFonts slot
            // (latin/ea/cs/themes) EXCEPT the <w:hint> font-slot selector, so an
            // empty paragraph whose source mark rPr carried w:hint="eastAsia"
            // lost it on the dump → batch round-trip. Route font.hint onto the
            // mark rPr the same way the other slots go (ApplyRunFormatting's
            // font.hint case writes RunFonts.Hint).
            if (properties.TryGetValue("font.hint", out var ntFontHint))
                ApplyRunFormatting(EnsureNoTextMarkRPr(), "font.hint", ntFontHint);
            // BUG-DUMP33-02a: theme-font slots on no-text paragraph hoist.
            // Mirrors the text-run path (font.asciiTheme / font.hAnsiTheme /
            // font.eaTheme / font.csTheme) so `add p --prop font.eaTheme=...`
            // writes RunFonts.*Theme on the paragraph mark rPr instead of
            // falling to TypedAttributeFallback (which can't bind
            // dotted-theme keys onto the typed RunFonts element).
            string? ntAsciiTheme = null, ntHAnsiTheme = null, ntEaTheme = null, ntCsTheme = null;
            if (properties.TryGetValue("font.asciiTheme", out var ntAT) || properties.TryGetValue("font.asciitheme", out ntAT))
                ntAsciiTheme = ntAT;
            if (properties.TryGetValue("font.hAnsiTheme", out var ntHAT) || properties.TryGetValue("font.hansitheme", out ntHAT))
                ntHAnsiTheme = ntHAT;
            if (properties.TryGetValue("font.eaTheme", out var ntEAT) || properties.TryGetValue("font.eatheme", out ntEAT) || properties.TryGetValue("font.eastasiatheme", out ntEAT))
                ntEaTheme = ntEAT;
            if (properties.TryGetValue("font.csTheme", out var ntCST) || properties.TryGetValue("font.cstheme", out ntCST))
                ntCsTheme = ntCST;
            if (ntAsciiTheme != null || ntHAnsiTheme != null || ntEaTheme != null || ntCsTheme != null)
            {
                var rp = EnsureNoTextMarkRPr();
                var rf = rp.GetFirstChild<RunFonts>();
                if (rf == null)
                {
                    rf = new RunFonts();
                    InsertRunPropInSchemaOrder(rp, rf);
                }
                if (ntAsciiTheme != null)
                    rf.AsciiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(ntAsciiTheme));
                if (ntHAnsiTheme != null)
                    rf.HighAnsiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(ntHAnsiTheme));
                if (ntEaTheme != null)
                    rf.EastAsiaTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(ntEaTheme));
                if (ntCsTheme != null)
                    rf.ComplexScriptTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(ntCsTheme));
            }
        }
        if (properties.TryGetValue("firstlineindent", out var indent) || properties.TryGetValue("firstLineIndent", out indent))
        {
            // Lenient input: accept "2cm", "0.5in", "18pt", or bare twips (backward compat).
            // OOXML w:firstLine is ST_SignedTwipsMeasure — negatives are legal hanging
            // indents (Set already uses ParseWordSpacingSigned). CONSISTENCY(add-set-symmetry).
            var indentTwips = SpacingConverter.ParseWordSpacingSigned(indent);
            if (Math.Abs(indentTwips) > 31680)
                throw new OverflowException($"First line indent value out of range (|v| <= 31680 twips): {indent}");
            pProps.Indentation = new Indentation
            {
                FirstLine = indentTwips.ToString()  // signed twips, consistent with Set and Get
            };
        }
        if (properties.TryGetValue("spacebefore", out var sb4) || properties.TryGetValue("spaceBefore", out sb4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.Before = SpacingConverter.ParseWordSpacing(sb4).ToString();
        }
        if (properties.TryGetValue("spaceafter", out var sa4) || properties.TryGetValue("spaceAfter", out sa4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.After = SpacingConverter.ParseWordSpacing(sa4).ToString();
        }
        // BUG-DUMP-R24-5: <w:spacing w:beforeLines/w:afterLines> — ½-line
        // (font-relative, 1/100 of a line) spacing Word PRECEDES the fixed
        // before/after twips with. The dump already captures these on `add p`
        // (spaceBeforeLines/spaceAfterLines); without applying them here the
        // direct paragraph-spacing path strips them (the style path already
        // honoured the same attrs), reflowing the doc and inserting a spurious
        // blank page. Values are raw 1/100-line integers (no unit conversion).
        if (properties.TryGetValue("spacebeforelines", out var sbl4) || properties.TryGetValue("spaceBeforeLines", out sbl4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.BeforeLines = ParseHelpers.SafeParseInt(sbl4, "spaceBeforeLines");
        }
        if (properties.TryGetValue("spaceafterlines", out var sal4) || properties.TryGetValue("spaceAfterLines", out sal4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.AfterLines = ParseHelpers.SafeParseInt(sal4, "spaceAfterLines");
        }
        // BUG-DUMP-R44-4: auto-spacing on/off toggles (w:beforeAutospacing /
        // w:afterAutospacing). Round-trips the bool toggle the readback emits as
        // spaceBeforeAuto / spaceAfterAuto.
        if (properties.TryGetValue("spacebeforeauto", out var sba4) || properties.TryGetValue("spaceBeforeAuto", out sba4)
            || properties.TryGetValue("beforeautospacing", out sba4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.BeforeAutoSpacing = OnOffValue.FromBoolean(IsTruthy(sba4));
        }
        if (properties.TryGetValue("spaceafterauto", out var saa4) || properties.TryGetValue("spaceAfterAuto", out saa4)
            || properties.TryGetValue("afterautospacing", out saa4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.AfterAutoSpacing = OnOffValue.FromBoolean(IsTruthy(saa4));
        }
        if (properties.TryGetValue("linespacing", out var ls4) || properties.TryGetValue("lineSpacing", out ls4))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(ls4);
            spacing.Line = twips.ToString();
            spacing.LineRule = isMultiplier ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
        }
        // BUG-019: lineSpacing alone cannot distinguish AtLeast from Exact —
        // both serialize as "Npt" via SpacingConverter. Accept an explicit
        // `lineRule` prop (auto/exact/atLeast) so dump→batch round-trips
        // preserve the rule. Without this, AtLeast spacing silently
        // downgraded to Exact, producing glyph clipping on tall content.
        if (properties.TryGetValue("lineRule", out var pLineRule) || properties.TryGetValue("linerule", out pLineRule))
        {
            var spacing = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
            spacing.LineRule = ParseLineRule(pLineRule);
        }
        // Numbering properties. Parallel branches so `ilvl` alone still
        // emits <w:ilvl> (matching `set --prop ilvl=N` behaviour); both
        // inputs are range-checked so schema-invalid values never reach XML.
        if (properties.TryGetValue("numid", out var numId)
            || properties.TryGetValue("numId", out numId)
            || properties.TryGetValue("listId", out numId)
            || properties.TryGetValue("listid", out numId))
        {
            var numIdVal = ParseHelpers.SafeParseInt(numId, "numid");
            // numId=-1 is the OOXML negation marker (override inherited numbering
            // back to "no list"); treat it like 0 (skip existence check).
            if (numIdVal < -1)
                throw new ArgumentException($"numId must be >= -1 (got {numIdVal}).");
            // numId=0 is OOXML's way of saying "remove numbering" (no-list sentinel).
            // Positive numIds must reference an existing <w:num> to avoid silent dangling
            // references — Word renders such paragraphs without any list marker.
            if (numIdVal > 0)
            {
                var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                var numExists = numbering?.Elements<NumberingInstance>()
                    .Any(n => n.NumberID?.Value == numIdVal) ?? false;
                if (!numExists)
                    throw new ArgumentException(
                        $"numId={numIdVal} not found in /numbering. " +
                        "Create the num first (add /numbering --type num), or use numId=0 to remove numbering.");
            }
            var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
            numPr.NumberingId = new NumberingId { Val = numIdVal };
        }
        // Accept both "numlevel" and "ilvl" (the OOXML name); works with or
        // without numId to stay in sync with `set --prop ilvl=N`.
        if (properties.TryGetValue("numlevel", out var numLevel)
            || properties.TryGetValue("ilvl", out numLevel)
            || properties.TryGetValue("listLevel", out numLevel)
            || properties.TryGetValue("listlevel", out numLevel))
        {
            var ilvlVal = ParseHelpers.SafeParseInt(numLevel, "ilvl");
            // BUG-R4B(BUG2): clamp ilvl > 8 (Word tolerates it) instead of
            // rejecting, so dump→replay never drops the paragraph. Mirror the
            // Set path (WordHandler.Set.cs numLevel case) and the HtmlPreview
            // clamp.
            if (ilvlVal < 0 || ilvlVal > 8)
            {
                var clamped = Math.Clamp(ilvlVal, 0, 8);
                LastAddWarnings.Add($"ilvl {ilvlVal} out of OOXML range 0..8 — clamped to {clamped}");
                ilvlVal = clamped;
            }
            var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
            numPr.NumberingLevelReference = new NumberingLevelReference { Val = ilvlVal };
        }
        if (properties.TryGetValue("tabs", out var pTabsVal) || properties.TryGetValue("tabstops", out pTabsVal))
        {
            ApplyTabsShorthand(pProps, pTabsVal);
        }
        if (!shadingHoistedToMarkRPr
            && (properties.TryGetValue("shd", out var pShdVal) || properties.TryGetValue("shading", out pShdVal) || properties.TryGetValue("fill", out pShdVal)))
        {
            // BUG-DUMP-R41-4: route through the shared ParseShadingValue so the
            // theme-linkage key=val tail (themeFill/themeColor/…) round-trips;
            // it preserves the same VAL;FILL;COLOR positional semantics this
            // block hand-rolled before.
            pProps.Shading = ParseShadingValue(pShdVal);
        }
        if (properties.TryGetValue("leftindent", out var addLI) || properties.TryGetValue("leftIndent", out addLI) || properties.TryGetValue("indentleft", out addLI) || properties.TryGetValue("indent", out addLI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            // CONSISTENCY(lenient-spacing): route through SpacingConverter so indent accepts
            // "2cm"/"0.5in"/"24pt"/bare twips — parity with spaceBefore/spaceAfter/lineSpacing.
            // BUG-DUMP-NEGIND: w:ind/@w:left is ST_SignedTwipsMeasure — see
            // SpacingConverter.ParseWordSpacingSigned. Real docs (gov.cn TOC
            // overhangs) carry negative indents.
            ind.Left = SpacingConverter.ParseWordSpacingSigned(addLI).ToString();
        }
        if (properties.TryGetValue("rightindent", out var addRI) || properties.TryGetValue("rightIndent", out addRI) || properties.TryGetValue("indentright", out addRI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            // CONSISTENCY(lenient-spacing): see leftindent above.
            // BUG-DUMP-NEGIND: signed (see leftIndent above).
            ind.Right = SpacingConverter.ParseWordSpacingSigned(addRI).ToString();
        }
        if (properties.TryGetValue("hangingindent", out var addHI) || properties.TryGetValue("hangingIndent", out addHI) || properties.TryGetValue("hanging", out addHI))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            // CONSISTENCY(lenient-spacing): see leftindent above.
            ind.Hanging = SpacingConverter.ParseWordSpacing(addHI).ToString();
            ind.FirstLine = null;
        }
        // firstlineindent already handled above (line ~66-74) with × 480 conversion
        // BUG-R5-F3: Get already exposes char-based indent values that
        // CJK Word documents emit heavily (firstLineChars, leftChars,
        // rightChars, hangingChars — w:ind/@w:firstLineChars etc., units
        // of 1/100 of a Chinese-character width). Add ignored them, so
        // dump→replay produced 750+ UNSUPPORTED warnings on Chinese docs
        // and lost the chars-based indent silently. Accept them on Add.
        if (properties.TryGetValue("firstLineChars", out var addFLC) || properties.TryGetValue("firstlinechars", out addFLC))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.FirstLineChars = ParseHelpers.SafeParseInt(addFLC, "firstLineChars");
        }
        if (properties.TryGetValue("leftChars", out var addLC) || properties.TryGetValue("leftchars", out addLC))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.LeftChars = ParseHelpers.SafeParseInt(addLC, "leftChars");
        }
        if (properties.TryGetValue("rightChars", out var addRC) || properties.TryGetValue("rightchars", out addRC))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.RightChars = ParseHelpers.SafeParseInt(addRC, "rightChars");
        }
        if (properties.TryGetValue("hangingChars", out var addHC) || properties.TryGetValue("hangingchars", out addHC))
        {
            var ind = pProps.Indentation ?? (pProps.Indentation = new Indentation());
            ind.HangingChars = ParseHelpers.SafeParseInt(addHC, "hangingChars");
        }
        // v6.4: paragraph frame (<w:framePr/>). doc2 emits framePr.w /
        // framePr.h / framePr.x / framePr.y / framePr.hSpace / framePr.vSpace
        // (twips) plus framePr.wrap / framePr.hAnchor / framePr.vAnchor
        // (docx enum keywords). Each is optional — we only attach a
        // FrameProperties child when at least one frame-* prop was set.
        // SDK API: Width/X/Y/HorizontalSpace/VerticalSpace are StringValue,
        // Height is UInt32Value, HorizontalPosition/VerticalPosition carry
        // the anchor enums.
        FrameProperties? frameProps = null;
        FrameProperties EnsureFramePr() => frameProps ??= new FrameProperties();
        // CONSISTENCY(length-units): framePr size/position slots accept pt/cm/in
        // (bare = twips) via SpacingConverter, matching tc padding and every other
        // length slot. Convert to twips first, then enforce the ST_(Signed)TwipsMeasure
        // ±31680 bound on the resolved twips.
        int FrameTwips(string raw, string prop, bool signed)
        {
            int tw;
            try
            {
                tw = signed
                    ? OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(raw)
                    : (int)OfficeCli.Core.SpacingConverter.ParseWordSpacing(raw);
            }
            catch (ArgumentException)
            {
                // Re-wrap so the message names the slot ("auto" and other
                // non-length tokens are rejected here, not silently written).
                throw new ArgumentException($"Invalid '{prop}' value: '{raw}'. Must resolve to a twips length (bare number, or a pt/cm/in value).");
            }
            int lo = signed ? -31680 : 0;
            if (tw < lo || tw > 31680)
                throw new ArgumentException($"Invalid '{prop}' value: '{raw}'. Must resolve to {lo}..31680 twips (bare number, or a pt/cm/in length).");
            return tw;
        }
        if (properties.TryGetValue("framePr.w", out var fpW) || properties.TryGetValue("framepr.w", out fpW))
            EnsureFramePr().Width = FrameTwips(fpW, "framePr.w", signed: false).ToString();
        if (properties.TryGetValue("framePr.h", out var fpH) || properties.TryGetValue("framepr.h", out fpH))
            EnsureFramePr().Height = (uint)FrameTwips(fpH, "framePr.h", signed: false);
        if (properties.TryGetValue("framePr.x", out var fpX) || properties.TryGetValue("framepr.x", out fpX))
            EnsureFramePr().X = FrameTwips(fpX, "framePr.x", signed: true).ToString();
        if (properties.TryGetValue("framePr.y", out var fpY) || properties.TryGetValue("framepr.y", out fpY))
            EnsureFramePr().Y = FrameTwips(fpY, "framePr.y", signed: true).ToString();
        if (properties.TryGetValue("framePr.hSpace", out var fpHS) || properties.TryGetValue("framepr.hspace", out fpHS))
            EnsureFramePr().HorizontalSpace = FrameTwips(fpHS, "framePr.hSpace", signed: false).ToString();
        if (properties.TryGetValue("framePr.vSpace", out var fpVS) || properties.TryGetValue("framepr.vspace", out fpVS))
            EnsureFramePr().VerticalSpace = FrameTwips(fpVS, "framePr.vSpace", signed: false).ToString();
        if (properties.TryGetValue("framePr.wrap", out var fpWrap) || properties.TryGetValue("framepr.wrap", out fpWrap))
        {
            EnsureFramePr().Wrap = fpWrap.ToLowerInvariant() switch
            {
                "auto"      => TextWrappingValues.Auto,
                "around"    => TextWrappingValues.Around,
                "none"      => TextWrappingValues.None,
                "notbeside" => TextWrappingValues.NotBeside,
                "through"   => TextWrappingValues.Through,
                _ => throw new ArgumentException($"Invalid 'framePr.wrap' value: '{fpWrap}'. Valid values: auto, around, none, notBeside, through."),
            };
        }
        if (properties.TryGetValue("framePr.hAnchor", out var fpHA) || properties.TryGetValue("framepr.hanchor", out fpHA))
        {
            EnsureFramePr().HorizontalPosition = fpHA.ToLowerInvariant() switch
            {
                "page"   => HorizontalAnchorValues.Page,
                "margin" => HorizontalAnchorValues.Margin,
                "text"   => HorizontalAnchorValues.Text,
                _ => throw new ArgumentException($"Invalid 'framePr.hAnchor' value: '{fpHA}'. Valid values: page, margin, text."),
            };
        }
        if (properties.TryGetValue("framePr.vAnchor", out var fpVA) || properties.TryGetValue("framepr.vanchor", out fpVA))
        {
            EnsureFramePr().VerticalPosition = fpVA.ToLowerInvariant() switch
            {
                "page"   => VerticalAnchorValues.Page,
                "margin" => VerticalAnchorValues.Margin,
                "text"   => VerticalAnchorValues.Text,
                _ => throw new ArgumentException($"Invalid 'framePr.vAnchor' value: '{fpVA}'. Valid values: page, margin, text."),
            };
        }
        // OOXML ST_XAlign / ST_YAlign are enums; SDK FrameProperties.XAlign and
        // YAlign are EnumValue-typed but TypedAttributeFallback in Set still
        // writes the raw string via the SDK accessor. Curate here so invalid
        // values surface as ArgumentException instead of silently producing
        // schema-invalid XML.
        if (properties.TryGetValue("framePr.xAlign", out var fpXA) || properties.TryGetValue("framepr.xalign", out fpXA))
        {
            EnsureFramePr().XAlign = fpXA.ToLowerInvariant() switch
            {
                "left"    => HorizontalAlignmentValues.Left,
                "center"  => HorizontalAlignmentValues.Center,
                "right"   => HorizontalAlignmentValues.Right,
                "inside"  => HorizontalAlignmentValues.Inside,
                "outside" => HorizontalAlignmentValues.Outside,
                _ => throw new ArgumentException($"Invalid 'framePr.xAlign' value: '{fpXA}'. Valid values: left, center, right, inside, outside."),
            };
        }
        if (properties.TryGetValue("framePr.dropCap", out var fpDC) || properties.TryGetValue("framepr.dropcap", out fpDC))
        {
            EnsureFramePr().DropCap = fpDC.ToLowerInvariant() switch
            {
                "none"   => DropCapLocationValues.None,
                "drop"   => DropCapLocationValues.Drop,
                "margin" => DropCapLocationValues.Margin,
                _ => throw new ArgumentException($"Invalid 'framePr.dropCap' value: '{fpDC}'. Valid values: none, drop, margin."),
            };
        }
        if (properties.TryGetValue("framePr.yAlign", out var fpYA) || properties.TryGetValue("framepr.yalign", out fpYA))
        {
            EnsureFramePr().YAlign = fpYA.ToLowerInvariant() switch
            {
                "inline"  => VerticalAlignmentValues.Inline,
                "top"     => VerticalAlignmentValues.Top,
                "center"  => VerticalAlignmentValues.Center,
                "bottom"  => VerticalAlignmentValues.Bottom,
                "inside"  => VerticalAlignmentValues.Inside,
                "outside" => VerticalAlignmentValues.Outside,
                _ => throw new ArgumentException($"Invalid 'framePr.yAlign' value: '{fpYA}'. Valid values: inline, top, center, bottom, inside, outside."),
            };
        }
        // framePr.lines is the drop-cap line span. Word UI exposes 1..10;
        // anything outside renders unpredictably and round-trips poorly.
        // The generic TypedAttributeFallback below didn't range-check either
        // bound, so 0 / 50 silently slipped through.
        if (properties.TryGetValue("framePr.lines", out var fpLines) || properties.TryGetValue("framepr.lines", out fpLines))
        {
            if (!int.TryParse(fpLines, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var fpLinesInt)
                || fpLinesInt < 1 || fpLinesInt > 10)
                throw new ArgumentException($"Invalid 'framePr.lines' value: '{fpLines}'. Must be an integer 1..10 (drop-cap line span).");
            EnsureFramePr().Lines = fpLinesInt;
        }
        if (frameProps != null)
            pProps.FrameProperties = frameProps;

        // keepNext / keepLines / pageBreakBefore are <w:onOff>-typed: the
        // bare element means "true", and an explicit <w:keepNext w:val="0"/>
        // means "false" (and OVERRIDES a true inherited from a paragraph
        // style — common pattern in heading-style paragraphs that want to
        // disable the style's default keep-with-next). Write both forms.
        if (properties.TryGetValue("keepnext", out var addKN) || properties.TryGetValue("keepNext", out addKN))
            pProps.KeepNext = IsTruthy(addKN)
                ? new KeepNext()
                : new KeepNext { Val = OnOffValue.FromBoolean(false) };
        if (properties.TryGetValue("keeplines", out var addKL)
            || properties.TryGetValue("keeptogether", out addKL)
            || properties.TryGetValue("keepLines", out addKL)
            || properties.TryGetValue("keepTogether", out addKL))
            pProps.KeepLines = IsTruthy(addKL)
                ? new KeepLines()
                : new KeepLines { Val = OnOffValue.FromBoolean(false) };
        if (properties.TryGetValue("pagebreakbefore", out var addPBB) || properties.TryGetValue("pageBreakBefore", out addPBB))
            pProps.PageBreakBefore = IsTruthy(addPBB)
                ? new PageBreakBefore()
                : new PageBreakBefore { Val = OnOffValue.FromBoolean(false) };
        // fuzz-2: paragraph-context `break=newPage` alias → pageBreakBefore=true.
        // Mirrors Set-side handling in WordHandler.Set.cs (case "break").
        if (properties.TryGetValue("break", out var addBrk))
        {
            bool pbb = addBrk?.ToLowerInvariant() switch
            {
                "newpage" or "page" or "nextpage" or "pagebreak" => true,
                "none" or "" or null => false,
                _ => IsTruthy(addBrk)
            };
            if (pbb) pProps.PageBreakBefore = new PageBreakBefore();
        }
        if (properties.TryGetValue("widowcontrol", out var addWC) || properties.TryGetValue("widowControl", out addWC))
        {
            if (IsTruthy(addWC))
                pProps.WidowControl = new WidowControl();
            else
                pProps.WidowControl = new WidowControl { Val = false };
        }
        // CONSISTENCY(add-set-symmetry): snapToGrid is valid on BOTH pPr and rPr.
        // A bare `snapToGrid` key on `add p` is the PARAGRAPH-level property (the
        // dump emits run-level snapToGrid as a separate `add r` op). Route it to
        // pPr here — mirrors widowControl/wordWrap above — so it does NOT fall to
        // the bare-run fallback below (which, since ApplyRunFormatting gained a
        // snapToGrid case, would otherwise stamp it onto the content run and
        // change a paragraph-level grid opt-out into a run-level one). Both
        // true/false write an explicit element; the OFF form is the meaningful
        // one on a doc with a docGrid. snapToGrid is in bareConsumed so the loop
        // skips it after this.
        if (properties.TryGetValue("snaptogrid", out var addSnap) || properties.TryGetValue("snapToGrid", out addSnap))
        {
            pProps.SnapToGrid = IsTruthy(addSnap)
                ? new SnapToGrid()
                : new SnapToGrid { Val = OnOffValue.FromBoolean(false) };
        }
        // CONSISTENCY(add-set-symmetry): Set accepts wordWrap via the toggle
        // fallback in WordHandler.Set.cs; Add mirrors it so callers can build
        // CJK right-aligned paragraphs (which need wordWrap=false to preserve
        // trailing whitespace on right-aligned lines) in one call.
        if (properties.TryGetValue("wordwrap", out var addWW) || properties.TryGetValue("wordWrap", out addWW))
        {
            pProps.WordWrap = IsTruthy(addWW)
                ? new WordWrap()
                : new WordWrap { Val = false };
        }
        // CONSISTENCY(add-set-symmetry): Set supports contextualSpacing (WordHandler.Set.cs:529);
        // Add must accept the same prop so the "Add then Get" lifecycle test pattern works
        // without falling back to a separate Set call. Both true and false write an
        // explicit element — `false` is meaningful when a parent style sets
        // contextualSpacing=true, since omitting the element would inherit the
        // style's `true`. Setting `Val=false` explicitly overrides.
        if (properties.TryGetValue("contextualspacing", out var addCS) || properties.TryGetValue("contextualSpacing", out addCS))
            pProps.ContextualSpacing = IsTruthy(addCS)
                ? new ContextualSpacing()
                : new ContextualSpacing { Val = false };
        // CONSISTENCY(add-set-symmetry): Set supports outlineLvl via the
        // schema fallback (TrySetParagraphProp + TypedAttributeFallback);
        // Add must accept the same canonical key so dump round-trip stays
        // lossless — the dump emitter pulls outlineLvl from paragraph Get
        // readback (WordHandler.Navigation.cs:1265-1266) and surfaces it as
        // an Add prop. BUG-R4-BT4.
        if (properties.TryGetValue("outlineLvl", out var addOLvl)
            || properties.TryGetValue("outlinelvl", out addOLvl)
            || properties.TryGetValue("outlineLevel", out addOLvl)
            || properties.TryGetValue("outlinelevel", out addOLvl))
        {
            // OOXML w:outlineLvl/@w:val is ST_DecimalNumber 0..9. Reject
            // out-of-range upfront — silent-drop produced files where the
            // user-specified outline level disappeared without warning, and
            // even an unparseable int slipped through.
            if (!int.TryParse(addOLvl, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var olvl)
                || olvl < 0 || olvl > 9)
                throw new ArgumentException($"Invalid 'outlineLvl' value: '{addOLvl}'. Must be 0-9 (OOXML MaxInclusive=9).");
            pProps.OutlineLevel = new OutlineLevel { Val = olvl };
        }
        // CONSISTENCY(add-set-symmetry): paragraph rStyle binds the paragraph
        // mark's run style. Run Add already supports rStyle; paragraph dump
        // emit echoes it back from Get (mark rPr.rStyle) and the value
        // applies to all runs the paragraph carries via its mark inheritance.
        // BUG-R4-BT4. Stored in ParagraphMarkRunProperties so the run-style
        // sticks to the paragraph mark itself (not just any subsequently
        // added run).
        if (properties.TryGetValue("rStyle", out var addPRStyle) || properties.TryGetValue("rstyle", out addPRStyle))
        {
            var pmrp = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            pmrp.RemoveAllChildren<RunStyle>();
            pmrp.PrependChild(new RunStyle { Val = addPRStyle });
        }
        // CONSISTENCY(add-set-symmetry): Set accepts border.top/bottom/left/right/between/bar
        // (and bare "border"/"border.all"); Add must accept the same vocabulary so the
        // Add → Get → verify lifecycle works without a follow-up Set call.
        // 3-segment keys (pbdr.top.sz / pbdr.top.color / pbdr.top.space)
        // surface in Get readback but Set's TrySetParagraphProp switch
        // doesn't model them either — calling ApplyParagraphBorders with a
        // 3-segment key drives ParseBorderValue with the sub-attribute
        // value (e.g. "4"), which throws "Invalid border style: '4'".
        // Skip them here to keep Add/Set symmetry (BUG-R2-02 / BT-2).
        var appliedBorderKeys = new List<string>();
        foreach (var (pk, pv) in properties)
        {
            if ((pk.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase)
                 || pk.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                && pk.Count(ch => ch == '.') < 2)
            {
                ApplyParagraphBorders(pProps, pk, pv);
                appliedBorderKeys.Add(pk);
            }
        }
        // This loop reads border keys by iterating `properties` (static type
        // Dictionary<string,string>), which bypasses the TrackingPropertyDictionary
        // enumerator override — so per-side keys (border.top, pbdr.left, …) would
        // be flagged unsupported_property even though they apply fine. Mark them
        // consumed via the dictionary's sanctioned MarkAllConsumed API.
        (properties as OfficeCli.Core.TrackingPropertyDictionary)?.MarkAllConsumed(appliedBorderKeys);
        if (properties.TryGetValue("liststyle", out var listStyle) || properties.TryGetValue("listStyle", out listStyle))
        {
            para.AppendChild(pProps);
            int? startVal = null;
            if (properties.TryGetValue("start", out var sv))
                startVal = ParseHelpers.SafeParseInt(sv, "start");
            int? levelVal = null;
            if (properties.TryGetValue("listLevel", out var ll) || properties.TryGetValue("listlevel", out ll) || properties.TryGetValue("level", out ll) || properties.TryGetValue("numlevel", out ll))
            {
                levelVal = ParseHelpers.SafeParseInt(ll, "listLevel");
                // OOXML ST_DecimalNumber ilvl is bound to 0..8 (ECMA-376
                // §17.9.3) — Word silently drops out-of-range values, so
                // reject up-front to keep round-trip lossless.
                if (levelVal < 0 || levelVal > 8)
                    throw new ArgumentException($"listLevel must be in range 0..8 (got {levelVal}).");
            }
            ApplyListStyle(para, listStyle, startVal, levelVal, containerHint: parent);
            // pProps already appended, skip the append below
            goto paragraphPropsApplied;
        }

        para.AppendChild(pProps);
        paragraphPropsApplied:

        if (properties.TryGetValue("text", out var text))
        {
            var run = new Run();
            var rProps = new RunProperties();
            // Per-script font slots (font.latin / font.ea / font.cs) write
            // to ascii+hAnsi / eastAsia / cs respectively. Bare 'font'
            // populates ascii+hAnsi+eastAsia for backward compatibility.
            // Build a single RunFonts so per-slot values compose cleanly
            // when the user supplies more than one (e.g. font.latin=Calibri
            // + font.cs=Arabic Typesetting on the same run).
            string? rfAscii = null, rfHAnsi = null, rfEa = null, rfCs = null;
            if (properties.TryGetValue("font", out var font) || properties.TryGetValue("font.name", out font))
            {
                rfAscii = font; rfHAnsi = font; rfEa = font;
            }
            if (properties.TryGetValue("font.latin", out var fLatin))
            {
                rfAscii = fLatin; rfHAnsi = fLatin;
            }
            if (properties.TryGetValue("font.ea", out var fEa)
                || properties.TryGetValue("font.eastasia", out fEa)
                || properties.TryGetValue("font.eastasian", out fEa))
            {
                rfEa = fEa;
            }
            if (properties.TryGetValue("font.cs", out var fCs)
                || properties.TryGetValue("font.complexscript", out fCs)
                || properties.TryGetValue("font.complex", out fCs))
            {
                rfCs = fCs;
            }
            // BUG-DUMP14-03: theme-font slot support — bind a run to a theme
            // major/minor font (rFonts/@*Theme) instead of a literal face.
            string? rfAsciiTheme = null, rfHAnsiTheme = null, rfEaTheme = null, rfCsTheme = null;
            if (properties.TryGetValue("font.asciiTheme", out var fAT) || properties.TryGetValue("font.asciitheme", out fAT))
                rfAsciiTheme = fAT;
            if (properties.TryGetValue("font.hAnsiTheme", out var fHAT) || properties.TryGetValue("font.hansitheme", out fHAT))
                rfHAnsiTheme = fHAT;
            if (properties.TryGetValue("font.eaTheme", out var fEAT) || properties.TryGetValue("font.eatheme", out fEAT) || properties.TryGetValue("font.eastasiatheme", out fEAT))
                rfEaTheme = fEAT;
            if (properties.TryGetValue("font.csTheme", out var fCST) || properties.TryGetValue("font.cstheme", out fCST))
                rfCsTheme = fCST;
            // BUG-DUMP-R31-2: <w:rFonts w:hint> font-slot selector on the
            // implicit text run. A single-run paragraph collapses to `add p
            // {font.hint=eastAsia text=…}`; without binding the hint to the
            // implicit run's RunFonts the round-trip dropped it (the dump now
            // captures it on the paragraph node via the firstRun hoist, but the
            // text-run path never applied it). w:hint selects which font slot
            // renders boundary CJK glyphs.
            string? rfHint = properties.TryGetValue("font.hint", out var fHint) ? fHint : null;
            if (rfAscii != null || rfHAnsi != null || rfEa != null || rfCs != null
                || rfAsciiTheme != null || rfHAnsiTheme != null || rfEaTheme != null || rfCsTheme != null
                || rfHint != null)
            {
                var rFonts = new RunFonts();
                if (rfAscii != null) rFonts.Ascii = rfAscii;
                if (rfHAnsi != null) rFonts.HighAnsi = rfHAnsi;
                if (rfEa != null) rFonts.EastAsia = rfEa;
                if (rfCs != null) rFonts.ComplexScript = rfCs;
                if (rfAsciiTheme != null)
                    rFonts.AsciiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(rfAsciiTheme));
                if (rfHAnsiTheme != null)
                    rFonts.HighAnsiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(rfHAnsiTheme));
                if (rfEaTheme != null)
                    rFonts.EastAsiaTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(rfEaTheme));
                if (rfCsTheme != null)
                    rFonts.ComplexScriptTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(rfCsTheme));
                if (rfHint != null)
                {
                    var hintLower = rfHint.Trim().ToLowerInvariant();
                    rFonts.Hint = hintLower switch
                    {
                        "eastasia" => FontTypeHintValues.EastAsia,
                        "cs" => FontTypeHintValues.ComplexScript,
                        "default" => FontTypeHintValues.Default,
                        _ => rFonts.Hint
                    };
                }
                rProps.AppendChild(rFonts);
            }
            // BUG-R6-03 / F-3: rStyle binds the paragraph mark above (so the
            // style sticks to the paragraph) but the implicit text run
            // rendered alongside `text=…` previously inherited Normal —
            // every dump→batch round-trip silently dropped run-style
            // formatting from headings (`add p text=… rStyle=Strong`).
            // Apply rStyle to the implicit run rPr too so the visible text
            // picks up the character style in addition to the mark.
            if (properties.TryGetValue("rStyle", out var pRunRStyle)
                || properties.TryGetValue("rstyle", out pRunRStyle))
            {
                rProps.RunStyle = new RunStyle { Val = pRunRStyle };
            }
            if (properties.TryGetValue("size", out var size) || properties.TryGetValue("font.size", out size) || properties.TryGetValue("fontsize", out size))
            {
                rProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(size) * 2, MidpointRounding.AwayFromZero)).ToString() });
            }
            // CONSISTENCY(toggle-explicit-false): match the no-text branch
            // (BUG-R7-07) — explicit `false` must emit <w:b w:val="false"/>
            // so a run can override a style-asserted toggle. IsTruthy alone
            // would silently drop the override and the run would re-inherit
            // bold/italic from the style chain (e.g. non-bold span inside
            // Heading1, non-italic citation inside Quote).
            if (properties.TryGetValue("bold", out var bold) || properties.TryGetValue("font.bold", out bold))
            {
                if (IsTruthy(bold)) rProps.Bold = new Bold();
                else if (IsExplicitFalseAddOverride(bold))
                    rProps.Bold = new Bold { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("bold.cs", out var boldCs)
                || properties.TryGetValue("font.bold.cs", out boldCs))
            {
                // Mirror the r-level path (BUG-DUMP-BCS-FALSE): an explicit
                // OFF must land on the generated run too, not only on the
                // paragraph-mark rPr — an on-only gate dropped the override
                // and the run re-inherited the style's bold.
                if (IsTruthy(boldCs)) rProps.BoldComplexScript = new BoldComplexScript();
                else if (IsExplicitFalseAddOverride(boldCs))
                    rProps.BoldComplexScript = new BoldComplexScript { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("italic", out var pItalic) || properties.TryGetValue("font.italic", out pItalic))
            {
                if (IsTruthy(pItalic)) rProps.Italic = new Italic();
                else if (IsExplicitFalseAddOverride(pItalic))
                    rProps.Italic = new Italic { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("italic.cs", out var italicCs)
                || properties.TryGetValue("font.italic.cs", out italicCs))
            {
                if (IsTruthy(italicCs)) rProps.ItalicComplexScript = new ItalicComplexScript();
                else if (IsExplicitFalseAddOverride(italicCs))
                    rProps.ItalicComplexScript = new ItalicComplexScript { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("size.cs", out var sizeCs)
                || properties.TryGetValue("font.size.cs", out sizeCs))
            {
                rProps.FontSizeComplexScript = new FontSizeComplexScript
                {
                    Val = ((int)Math.Round(ParseFontSize(sizeCs) * 2, MidpointRounding.AwayFromZero)).ToString()
                };
            }
            if (properties.TryGetValue("color", out var pColor) || properties.TryGetValue("font.color", out pColor))
            {
                // CONSISTENCY(theme-color): Add paragraph color must accept
                // scheme color names (accent1, dark2, hyperlink, …) the same
                // way ApplyRunFormatting (Set path) does — otherwise
                // Add(.., {color=accent1}) would call SanitizeHex on the
                // scheme name and produce garbage hex.
                // CONSISTENCY(color-auto): bare "auto" is a legal Color val
                // (Word's "automatic" text color); short-circuit before the
                // scheme branch since "auto" is not a ThemeColorValues enum.
                // BUG-DUMP-R44-1: split the ';themeColor=…' tail so an inline-text
                // run color carrying an explicit hex + theme linkage keeps both.
                var (pColorPos, pColorTheme) = ExtractThemeTail(pColor);
                if (string.Equals(pColorPos, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    rProps.Color = new Color { Val = "auto" };
                }
                else if (pColorPos.Length == 0 && pColorTheme.Count > 0)
                {
                    rProps.Color = new Color { Val = "auto" };
                }
                else
                {
                    var pSchemeName = OfficeCli.Core.ParseHelpers.NormalizeSchemeColorName(pColorPos);
                    if (pSchemeName != null)
                        rProps.Color = new Color { Val = "auto", ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(pSchemeName)) };
                    else
                        rProps.Color = new Color { Val = SanitizeHex(pColorPos) };
                }
                ApplyColorTheme(rProps.Color, pColorTheme);
            }
            if (properties.TryGetValue("underline", out var pUnderline) || properties.TryGetValue("font.underline", out pUnderline))
            {
                var ulVal = NormalizeUnderlineValue(pUnderline);
                rProps.Underline = new Underline { Val = new UnderlineValues(ulVal) };
            }
            // CONSISTENCY(toggle-explicit-false): see bold/italic above.
            if (properties.TryGetValue("strike", out var pStrike)
                    || properties.TryGetValue("strikethrough", out pStrike)
                    || properties.TryGetValue("font.strike", out pStrike)
                    || properties.TryGetValue("font.strikethrough", out pStrike))
            {
                if (IsTruthy(pStrike)) rProps.Strike = new Strike();
                else if (IsExplicitFalseAddOverride(pStrike))
                    rProps.Strike = new Strike { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("highlight", out var pHighlight))
                rProps.Highlight = new Highlight { Val = ParseHighlightColor(pHighlight) };
            if (properties.TryGetValue("caps", out var pCaps)
                    || properties.TryGetValue("allcaps", out pCaps)
                    || properties.TryGetValue("allCaps", out pCaps))
            {
                if (IsTruthy(pCaps)) rProps.Caps = new Caps();
                else if (IsExplicitFalseAddOverride(pCaps))
                    rProps.Caps = new Caps { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("smallcaps", out var pSmallCaps) || properties.TryGetValue("smallCaps", out pSmallCaps))
            {
                if (IsTruthy(pSmallCaps)) rProps.SmallCaps = new SmallCaps();
                else if (IsExplicitFalseAddOverride(pSmallCaps))
                    rProps.SmallCaps = new SmallCaps { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("dstrike", out var pDstrike)
                || properties.TryGetValue("doublestrike", out pDstrike)
                || properties.TryGetValue("doubleStrike", out pDstrike))
            {
                if (IsTruthy(pDstrike)) rProps.DoubleStrike = new DoubleStrike();
                else if (IsExplicitFalseAddOverride(pDstrike))
                    rProps.DoubleStrike = new DoubleStrike { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("vanish", out var pVanish))
            {
                if (IsTruthy(pVanish)) rProps.Vanish = new Vanish();
                else if (IsExplicitFalseAddOverride(pVanish))
                    rProps.Vanish = new Vanish { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("outline", out var pOutline))
            {
                if (IsTruthy(pOutline)) rProps.Outline = new Outline();
                else if (IsExplicitFalseAddOverride(pOutline))
                    rProps.Outline = new Outline { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("shadow", out var pShadow))
            {
                if (IsTruthy(pShadow)) rProps.Shadow = new Shadow();
                else if (IsExplicitFalseAddOverride(pShadow))
                    rProps.Shadow = new Shadow { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("emboss", out var pEmboss))
            {
                if (IsTruthy(pEmboss)) rProps.Emboss = new Emboss();
                else if (IsExplicitFalseAddOverride(pEmboss))
                    rProps.Emboss = new Emboss { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("imprint", out var pImprint))
            {
                if (IsTruthy(pImprint)) rProps.Imprint = new Imprint();
                else if (IsExplicitFalseAddOverride(pImprint))
                    rProps.Imprint = new Imprint { Val = OnOffValue.FromBoolean(false) };
            }
            if (properties.TryGetValue("noproof", out var pNoProof))
            {
                if (IsTruthy(pNoProof)) rProps.NoProof = new NoProof();
                else if (IsExplicitFalseAddOverride(pNoProof))
                    rProps.NoProof = new NoProof { Val = OnOffValue.FromBoolean(false) };
            }
            // Run-level rtl: explicit `rtl=true` OR cascaded from paragraph
            // direction=rtl above. Skipping the cascade would leave Latin
            // character order inside an RTL paragraph (broken Arabic).
            // Routes through ApplyRunFormatting so schema order matches
            // direct Set path. See WordHandler.I18n.cs.
            if ((properties.TryGetValue("rtl", out var pRtl) && IsTruthy(pRtl))
                || paraRtl == true)
                ApplyRunFormatting(rProps, "rtl", "true");
            if (properties.TryGetValue("vertAlign", out var pVertAlign) || properties.TryGetValue("vertalign", out pVertAlign))
            {
                rProps.VerticalTextAlignment = new VerticalTextAlignment
                {
                    Val = pVertAlign.ToLowerInvariant() switch
                    {
                        "superscript" or "super" => VerticalPositionValues.Superscript,
                        "subscript" or "sub" => VerticalPositionValues.Subscript,
                        "baseline" => VerticalPositionValues.Baseline,
                        _ => throw new ArgumentException($"Invalid 'vertAlign' value: '{pVertAlign}'. Valid values: superscript, subscript, baseline."),
                    }
                };
            }
            if (properties.TryGetValue("superscript", out var pSup) && IsTruthy(pSup))
                rProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
            if (properties.TryGetValue("subscript", out var pSub) && IsTruthy(pSub))
                rProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
            if (properties.TryGetValue("charspacing", out var pCharSp) || properties.TryGetValue("charSpacing", out pCharSp)
                || properties.TryGetValue("letterspacing", out pCharSp) || properties.TryGetValue("letterSpacing", out pCharSp))
            {
                int pCsTwips = pCharSp.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? (int)Math.Round(ParseHelpers.SafeParseDouble(pCharSp[..^2], "charspacing") * 20, MidpointRounding.AwayFromZero)
                    : (int)Math.Round(ParseHelpers.SafeParseDouble(pCharSp, "charspacing"), MidpointRounding.AwayFromZero);
                rProps.Spacing = new Spacing { Val = pCsTwips };
            }
            // BUG-DUMP22-03: paragraph-level shading lives in pPr (written
            // above ~line 262/289). Do NOT also stamp it onto the inline
            // run's rPr — that produces a spurious <w:rPr><w:shd/></w:rPr>
            // duplicate that round-trips out as a separate run-level shading
            // command on dump replay.

            run.AppendChild(rProps);
            // w14 text effects (textFill/textOutline/w14glow/w14shadow/
            // w14reflection) are run-level; route them to the implicit text run
            // so `add paragraph --prop textFill=...` works like `--prop bold=...`.
            ApplyW14Effects(run, properties);
            AppendTextWithPageFields(para, run, rProps, text);
        }

        // Dotted-key fallback: any "element.attr=value" prop the hand-rolled
        // blocks above did not consume goes through the same generic helper
        // wired into Set. Pre-existing dotted prefixes already handled
        // upstream (pbdr.*) are skipped to avoid double application.
        // Anything still unconsumed is recorded as silent-drop so the CLI
        // layer can surface a WARNING. CONSISTENCY(add-set-symmetry).
        var rPropsForFallback = para.Descendants<RunProperties>().FirstOrDefault();
        // Set of bare (no-dot) keys that the curated text/run block above has
        // already consumed. Anything else bare is run-level (lang, bidi,
        // kern, …) and must reach ApplyRunFormatting / TypedAttributeFallback
        // — otherwise paragraph-add silently drops them while run-level Set /
        // Add accept them, breaking add/set symmetry.
        // CONSISTENCY(add-set-symmetry).
        var bareConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "type", "text", "html", "anchor", "anchorId", "anchorid",
            "style", "styleid", "stylename",
            "align", "alignment", "jc", "textAlignment", "textalignment",
            "textboxTightWrap", "textboxtightwrap",
            "direction", "dir", "bidi",
            "firstlineindent", "leftindent", "indentleft", "indent",
            // BUG-R5-F3: chars-based indent variants consumed above.
            "firstlinechars", "firstLineChars",
            "leftchars", "leftChars",
            "rightchars", "rightChars",
            "hangingchars", "hangingChars",
            "rightindent", "indentright", "hangingindent", "hanging",
            "spacebefore", "spaceafter", "linespacing", "lineSpacing", "linerule", "lineRule",
            // BUG-DUMP-R24-5: ½-line spacing attrs consumed by the
            // spaceBeforeLines/spaceAfterLines blocks above (mirror Set).
            "spacebeforelines", "spaceBeforeLines", "spaceafterlines", "spaceAfterLines",
            "keepnext", "keepwithnext", "keeplines", "keeptogether",
            "pagebreakbefore", "break",
            "widowcontrol", "widowControl",
            "snaptogrid", "snapToGrid",
            "numid", "numId", "ilvl", "numlevel", "numLevel",
            "liststyle", "listStyle", "start", "level", "listLevel", "listlevel",
            "outlinelevel", "outlineLevel",
            "outlinelvl", "outlineLvl",
            "rstyle", "rStyle",
            "tabs", "tabstops",
            "border", "borders", "shd", "shading", "fill",
            "font", "size", "fontsize", "fontSize", "bold", "italic", "color", "highlight",
            "underline", "strike", "strikethrough", "doublestrike", "dstrike",
            "vanish", "outline", "shadow", "emboss", "imprint", "noproof",
            // w14 text effects applied to the implicit run via ApplyW14Effects.
            "textfill", "textFill", "textoutline", "textOutline",
            "w14glow", "w14shadow", "w14reflection",
            "rtl", "vertAlign", "vertalign", "superscript", "subscript",
            "charspacing", "charSpacing", "letterspacing", "letterSpacing",
            "caps", "smallcaps",
            "boldcs", "italiccs", "sizecs",
            // BUG-DUMP23-01: bdr was previously listed here, which made the
            // fallback `continue` at line 765 skip it entirely (no curated
            // handler exists in the rProps block above either). Removed so
            // bdr falls through to ApplyRunFormatting like kern does.
            // kern was historically here too, "to prevent double-routing
            // through TypedAttributeFallback" — but the continue at the bare-
            // key fallback gate also skipped ApplyRunFormatting itself, so
            // kern was silently dropped on `add p kern=36` even though it
            // round-trips fine on `set r[N] kern=36`. Removed so kern reaches
            // ApplyRunFormatting on the bare-key fallback path below.
            // v5.9: paragraph-level format-revision marker keys consumed
            // by the pPrChange block at the end of AddParagraph. The bare
            // `revision` key is intentionally absent — creation uses
            // `revision.type=<kind>`, action uses `revision.action=accept|reject`;
            // a bare `revision` literal is an error (no silent allowlist).
            "revision.type",
            "revision.author",
            "revision.date",
            "revision.id",
        };
        // BUG-DUMP-MARKRPR-VERBATIM (class fix): if the dump emitted the WHOLE
        // ¶-mark <w:rPr> verbatim (markRPr.xml), apply it as the authoritative
        // mark rPr ONCE here and skip every per-property markRPr.* dotted key in
        // the loop below (they're redundant with the verbatim subtree and would
        // double-apply). This closes the hardcoded-allowlist class: every mark
        // rPr child — including ones no dotted key covers (w:em, w:effect, w:w,
        // w14:* OpenType extensions) — round-trips. markRPrVerbatimApplied also
        // signals the RTL cascade (ApplyDirectionCascade) NOT to re-stamp the
        // mark's <w:rtl/>, since the verbatim subtree already carries the source's
        // exact mark-rtl state (fixes the COP-13/Dari mark-rtl over-stamp).
        bool markRPrVerbatimApplied = false;
        if ((properties.TryGetValue("markRPr.xml", out var markRPrXml)
                || properties.TryGetValue("markrpr.xml", out markRPrXml))
            && !string.IsNullOrEmpty(markRPrXml) && markRPrXml.StartsWith("<"))
        {
            try
            {
                var pmRprVerbatim = new ParagraphMarkRunProperties(markRPrXml);
                pProps.RemoveAllChildren<ParagraphMarkRunProperties>();
                // CT_PPr schema order: ParagraphMarkRunProperties precedes
                // sectPr / pPrChange. Insert before the first of those, else
                // append (mirrors EnsureParagraphMarkRunPropertiesInSchemaOrder).
                OpenXmlElement? pmSuccessor = null;
                foreach (var child in pProps.ChildElements)
                {
                    if (child is SectionProperties || child is ParagraphPropertiesChange)
                    {
                        pmSuccessor = child;
                        break;
                    }
                }
                if (pmSuccessor != null)
                    pmSuccessor.InsertBeforeSelf(pmRprVerbatim);
                else
                    pProps.AppendChild(pmRprVerbatim);
                // Note: the rebuilt <w:rPr> carries a redundant xmlns:w decl (the
                // standalone fragment needed it to parse; the SDK keeps it on the
                // in-tree element). It is valid OOXML and idempotent (the next
                // dump re-emits the same OuterXml), just a cosmetically-verbose
                // but equivalent serialization of the same <w:rPr>.
                markRPrVerbatimApplied = true;
            }
            catch { /* malformed fragment — fall back to dotted keys below */ }
        }
        foreach (var (key, value) in properties)
        {
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            // Keys consumed by ApplyRunFormatting / TypedAttributeFallback /
            // GenericXmlQuery below leak as false unsupported without this.
            properties.ContainsKey(key);
            // BUG-DUMP9-02: paragraph-mark-only run formatting written under
            // the markRPr.* namespace. Mirrors SetElementParagraph; targets
            // ParagraphMarkRunProperties exclusively (does NOT propagate to
            // existing runs the way bare bold/color do).
            if (key.StartsWith("markRPr.", StringComparison.OrdinalIgnoreCase)
                || key.StartsWith("markrpr.", StringComparison.OrdinalIgnoreCase))
            {
                // Verbatim subtree already applied — ignore the redundant dotted
                // keys (and the markRPr.xml key itself) to avoid double-apply.
                if (markRPrVerbatimApplied) continue;
                var sub = key.Substring("markRPr.".Length);
                var pmRpr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                    ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                // BUG-DUMP33-02b: explicit-false markRPr.bold / markRPr.italic
                // must emit <w:b w:val="false"/> (resp. <w:i w:val="false"/>)
                // so the paragraph mark overrides a style that asserts
                // bold/italic. ApplyRunFormatting on its own removes the
                // element entirely on falsy input — same gap as the no-text
                // hoist block, fixed there with the IsExplicitFalseAddOverride
                // path. Mirror that here for round-trip parity.
                var subLower = sub.ToLowerInvariant();
                if (subLower == "bold" || subLower == "font.bold")
                {
                    pmRpr.RemoveAllChildren<Bold>();
                    if (IsTruthy(value))
                        InsertRunPropInSchemaOrder(pmRpr, new Bold());
                    else if (IsExplicitFalseAddOverride(value))
                        InsertRunPropInSchemaOrder(pmRpr, new Bold { Val = OnOffValue.FromBoolean(false) });
                    continue;
                }
                if (subLower == "italic" || subLower == "font.italic")
                {
                    pmRpr.RemoveAllChildren<Italic>();
                    if (IsTruthy(value))
                        InsertRunPropInSchemaOrder(pmRpr, new Italic());
                    else if (IsExplicitFalseAddOverride(value))
                        InsertRunPropInSchemaOrder(pmRpr, new Italic { Val = OnOffValue.FromBoolean(false) });
                    continue;
                }
                ApplyRunFormatting(pmRpr, sub, value);
                continue;
            }
            if (key.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase)) continue;
            if (!key.Contains('.') && bareConsumed.Contains(key)) continue;
            // revision.author / revision.date / revision.id — consumed by
            // AddParagraph's pPrChange block at end-of-function.
            if (key.StartsWith("revision.", StringComparison.OrdinalIgnoreCase))
                continue;
            // paraMarkDel.* / paraMarkIns.* — consumed by AddParagraph's
            // paraMarkDel / paraMarkIns blocks at end-of-function (BUG-DUMP-R44-6:
            // paraMarkIns is now a mark-only stamp here, no longer rewritten to a
            // bare revision.author content-insertion by the emitter).
            if (key.StartsWith("paraMarkDel.", StringComparison.OrdinalIgnoreCase)
                || key.StartsWith("paramarkdel.", StringComparison.OrdinalIgnoreCase)
                || key.StartsWith("paraMarkIns.", StringComparison.OrdinalIgnoreCase)
                || key.StartsWith("paramarkins.", StringComparison.OrdinalIgnoreCase)
                // BUG-DUMP-R49-1: numPrIns.* — consumed at end-of-function by
                // the numPrIns block (mirrors paraMarkIns.* handling). Skip here
                // so they don't hit UNSUPPORTED via the dotted-fallback paths.
                || key.StartsWith("numPrIns.", StringComparison.OrdinalIgnoreCase)
                || key.StartsWith("numprins.", StringComparison.OrdinalIgnoreCase))
                continue;
            if (!key.Contains('.'))
            {
                // Bare run-level key (lang, bidi, kern, …) — try
                // ApplyRunFormatting on the existing run rPr first, then on
                // the paragraph mark rPr (so it survives even with no text
                // run). Falls through to TypedAttributeFallback below.
                if (rPropsForFallback != null
                    && ApplyRunFormatting(rPropsForFallback, key, value)) continue;
                var bareMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                    ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                if (ApplyRunFormatting(bareMarkRPr, key, value)) continue;
                if (bareMarkRPr.ChildElements.Count == 0) bareMarkRPr.Remove();
            }
            // CONSISTENCY(font-dotted-alias): same skip-list as run-add.
            switch (key.ToLowerInvariant())
            {
                case "font.name":
                case "font.size":
                case "font.bold":
                case "font.italic":
                case "font.color":
                case "font.underline":
                case "font.strike":
                case "font.strikethrough":
                // Per-script font slots and CS toggles are already consumed
                // by the curated text/run block above; skip the typed-attr
                // fallback so they are not re-flagged as UNSUPPORTED.
                case "font.latin":
                case "font.ea":
                case "font.eastasia":
                case "font.eastasian":
                case "font.cs":
                case "font.complexscript":
                case "font.complex":
                // BUG-DUMP33-02a: theme-font slots — consumed by the no-text
                // hoist block (or the text-bearing run-creation block when a
                // run exists). TypedAttributeFallback can't bind these
                // dotted keys onto RunFonts so they would surface as
                // UNSUPPORTED on plain `add p`.
                case "font.asciitheme":
                case "font.hansitheme":
                case "font.eatheme":
                case "font.eastasiatheme":
                case "font.cstheme":
                // CS run flags (<w:bCs/> / <w:iCs/> / <w:szCs/>) — the
                // hoisted block at line 57-74 writes them to the paragraph
                // mark rPr; the dotted-fallback below would re-flag them
                // here because TypedAttributeFallback can't resolve the
                // dotted-name into the OpenXml element type.
                case "bold.cs":
                case "italic.cs":
                case "size.cs":
                case "font.bold.cs":
                case "font.italic.cs":
                case "font.size.cs":
                case "boldcs":
                case "italiccs":
                case "sizecs":
                    continue;
            }
            // CONSISTENCY(add-set-symmetry / bcp47-validation): route lang.*
            // through ApplyRunFormatting (Set's path) so the validator runs
            // on Add too. Target the existing run rPr if present, else the
            // paragraph mark rPr.
            switch (key.ToLowerInvariant())
            {
                case "lang.latin":
                case "lang.val":
                case "lang.ea":
                case "lang.eastasia":
                case "lang.eastasian":
                case "lang.cs":
                case "lang.complexscript":
                case "lang.bidi":
                {
                    if (rPropsForFallback != null
                        && ApplyRunFormatting(rPropsForFallback, key, value)) continue;
                    var langMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                        ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    if (ApplyRunFormatting(langMarkRPr, key, value)) continue;
                    break;
                }
            }
            // Paragraph border keys (border.* / pbdr.*, e.g. border.top) are
            // applied up-front by the ApplyParagraphBorders pass above; skip the
            // run/mark dotted fallback so they aren't re-flagged unsupported
            // (they target pPr/pBdr, not a run rPr child).
            if ((key.StartsWith("pbdr", StringComparison.OrdinalIgnoreCase)
                 || key.StartsWith("border", StringComparison.OrdinalIgnoreCase))
                && key.Count(ch => ch == '.') < 2)
                continue;
            if (Core.TypedAttributeFallback.TrySet(pProps, key, value)) continue;
            if (rPropsForFallback != null
                && Core.TypedAttributeFallback.TrySet(rPropsForFallback, key, value)) continue;
            // No text run on this paragraph yet; route run-level attrs to
            // the paragraph mark rPr (where they apply to the paragraph
            // mark glyph + inherited by future runs).
            var paraMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            if (Core.TypedAttributeFallback.TrySet(paraMarkRPr, key, value)) continue;
            if (paraMarkRPr.ChildElements.Count == 0) paraMarkRPr.Remove();
            // BUG-R5-04 / BUG-R5-05: bare-key val-leaves (textboxTightWrap,
            // divId, …) had no fallback path on Add — only TypedAttributeFallback,
            // which requires dotted keys. dump→batch round-trip emits these
            // as bare keys on `add p`, so they were silently dropped. Try
            // TryCreateTypedChild on pPr first (paragraph-scope leaves like
            // textboxTightWrap, divId), then on the run rPr / paragraph-mark
            // rPr for run-scope leaves (webHidden — BUG-R5-06: dump misplaces
            // it onto the paragraph, but accepting it on either container
            // here lets dump→replay succeed without losing the property).
            if (!key.Contains('.'))
            {
                if (Core.GenericXmlQuery.TryCreateTypedChild(pProps, key, value)) continue;
                if (rPropsForFallback != null
                    && Core.GenericXmlQuery.TryCreateTypedChild(rPropsForFallback, key, value)) continue;
                var fallbackMarkRPr = pProps.GetFirstChild<ParagraphMarkRunProperties>()
                    ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                if (Core.GenericXmlQuery.TryCreateTypedChild(fallbackMarkRPr, key, value)) continue;
                if (fallbackMarkRPr.ChildElements.Count == 0) fallbackMarkRPr.Remove();
            }
            LastAddUnsupportedProps.Add(key);
        }

        // Use ChildElements for index lookup so that tables and sectPr
        // siblings do not shift the effective insertion position. This
        // matches ResolveAnchorPosition, which computes anchor indices
        // against ChildElements.
        // PERF: only materialise the child list when an explicit index is given.
        // The append path (the batch-replay hot case — thousands of paragraphs)
        // does not need it, and ToList()'ing all children on every append made
        // body building O(N²).
        if (index.HasValue)
        {
            var allChildren = parent.ChildElements.ToList();
            if (index.Value < allChildren.Count)
            {
                var refElement = allChildren[index.Value];
                parent.InsertBefore(para, refElement);
                var paraPosIdx = PathIndex.FromArrayIndex(parent.Elements<Paragraph>().ToList().IndexOf(para));
                resultPath = $"{parentPath}/{BuildParaPathSegment(para, paraPosIdx)}";
                // Positional insert shifts which paragraph is last and the count;
                // drop the append-monotonic body cache.
                if (parent is Body) InvalidateBodyParaCache();
            }
            else
            {
                resultPath = $"{parentPath}/{BuildParaPathSegment(para, AppendBodyParaFast(parent, para))}";
            }
        }
        else
        {
            resultPath = $"{parentPath}/{BuildParaPathSegment(para, AppendBodyParaFast(parent, para))}";
        }
        // R20-fuzz-11: post-insert evaluation of inherited RTL for direction=ltr.
        // Only the style-chain layer can be evaluated before insertion; the
        // enclosing section, docDefaults, and numbering lvl all need the
        // paragraph to be parented. Mirror the Set path's HasInheritedBidi
        // helper and emit <w:bidi w:val="0"/> when any layer would otherwise
        // re-inherit RTL.
        if (paraRtl == false && pProps.GetFirstChild<BiDi>() == null && HasInheritedBidi(para))
        {
            pProps.BiDi = new BiDi { Val = new DocumentFormat.OpenXml.OnOffValue(false) };
        }

        // Paragraph-level `revision.type=format` → <w:pPrChange>.
        // Mirrors the run-side rPrChange path in AddRun. .doc carries
        // sprmPPropRMark (0xC63F); we stamp the marker with optional
        // author/date/id and leave the inner pPr empty (no recoverable
        // prior-property snapshot at v1).
        string? pTcKind = null;
        properties.TryGetValue("revision.type", out pTcKind);
        if (pTcKind?.Trim().ToLowerInvariant() == "format")
        {
            string? pTcAuthor = null;
            string? pTcDate = null;
            string? pTcId = null;
            properties.TryGetValue("revision.author", out pTcAuthor);
            properties.TryGetValue("revision.date", out pTcDate);
            properties.TryGetValue("revision.id", out pTcId);
            var pprChange = new ParagraphPropertiesChange();
            // BUG-DUMP-PPRCHANGE-AUTHOR: w:author is a REQUIRED attribute on
            // CT_TrackChange (pPrChange) — omitting it makes the file schema-invalid
            // and Word repairs-on-open. A source pPrChange authored with an EMPTY
            // name (w:author="") is common; Navigation's readback skips an empty
            // author so revision.author never arrives here, and the old
            // `if (!IsNullOrEmpty)` guard then dropped the attribute entirely.
            // Always stamp author (empty string when unknown) so the marker stays
            // schema-valid and the empty-author source round-trips faithfully.
            pprChange.Author = pTcAuthor ?? "";
            if (!string.IsNullOrEmpty(pTcDate) && DateTime.TryParse(pTcDate, out var pTcDt))
                pprChange.Date = pTcDt;
            pprChange.Id = !string.IsNullOrEmpty(pTcId)
                ? pTcId
                : GenerateRevisionId();
            pprChange.AppendChild(new PreviousParagraphProperties());
            pProps.AppendChild(pprChange);
            // BUG-DUMP-R43-8: restore the prior-pPr snapshot the dump captured
            // (revision.beforeXml) so Reject-Change recovers the original
            // paragraph formatting instead of an empty <w:pPr/> marker.
            if (properties.TryGetValue("revision.beforeXml", out var pTcBeforeXml)
                && !string.IsNullOrWhiteSpace(pTcBeforeXml))
                ApplyBeforeXmlSnapshot(pprChange, pTcBeforeXml);
        }

        // High-level paragraph-insertion revision: ANY revision.* sub-key
        // (author/date/id) WITHOUT a `revision.type=<kind>` literal means "this
        // paragraph was just inserted as a tracked change". Mirrors the
        // equivalent in AddRun (Phase 1) and the inverse in Mutations.Remove
        // (Phase 4: remove paragraph + revision.author produces ¶ del +
        // content del wrappers).
        //
        // Word UI semantic: pressing Enter in revision mode inserts a new
        // paragraph and marks BOTH the new ¶ AND any typed content as ins.
        // We emit:
        //   1. <w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>  — ¶ mark
        //   2. <w:ins><w:r>...</w:r></w:ins>               — content wrapper
        //      (only when --prop text=... auto-created an inner run)
        // Each gets a distinct auto-allocated revision id sharing the same
        // author + date — accept-all sees them as related but independent.
        // An explicit `revision.type=ins` takes the same path as the bare-author
        // auto-insert (empty pTcKind); only `revision.type=format` diverges to
        // the pPrChange branch above.
        if (string.IsNullOrEmpty(pTcKind) || pTcKind.Trim().ToLowerInvariant() == "ins")
        {
            string? hTcAuthor = null, hTcDate = null, hTcId = null;
            properties.TryGetValue("revision.author", out hTcAuthor);
            properties.TryGetValue("revision.date", out hTcDate);
            properties.TryGetValue("revision.id", out hTcId);

            if (!string.IsNullOrEmpty(hTcAuthor) || !string.IsNullOrEmpty(hTcDate) || !string.IsNullOrEmpty(hTcId))
            {
                var author = string.IsNullOrEmpty(hTcAuthor) ? "OfficeCLI" : hTcAuthor!;
                // BUG-R4F-03: parse with RoundtripKind so a UTC (…Z) revision
                // date stays Utc and re-serializes as …Z, matching the source
                // byte-for-byte (default TryParse degrades Z to host Local,
                // making dump→batch→dump non-idempotent and timezone-dependent).
                // Mirrors the comment-date path in Add.Misc.cs.
                DateTime date = !string.IsNullOrEmpty(hTcDate)
                    && DateTime.TryParse(hTcDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var hd)
                    ? hd : DateTime.UtcNow;
                // ¶ mark: <w:pPr>…<w:rPr><w:ins/>…</w:rPr></w:pPr>
                // Append (not prepend) the mark rPr: in CT_PPr the paragraph-mark
                // rPr sits near the END of the sequence (after pStyle / numPr /
                // spacing / …). Prepending forced it to position 0, which both
                // mis-placed the rPr AND made the following numPr/spacing parse
                // as unexpected children. PREPEND the <w:ins> inside the rPr: in
                // CT_ParaRPr the ins/del/move group leads the sequence, so when
                // markRPr.* props (rFonts/sz/…) were already added the ins must
                // precede them rather than be appended last.
                var pMarkRPr = pProps.ParagraphMarkRunProperties
                              ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                pMarkRPr.PrependChild(new Inserted
                {
                    Author = author,
                    Date = date,
                    Id = !string.IsNullOrEmpty(hTcId) ? hTcId : GenerateRevisionId(),
                });

                // Content: wrap each direct-child Run that was auto-created
                // from --prop text=... (paragraph-level text/html/sym) in
                // <w:ins>. Skip Runs that are already inside an ins/del/move
                // wrapper. Each wrapper gets its own auto-allocated id when
                // hTcId was not explicit (otherwise reuse it for the ¶ mark
                // only, per the Phase 4 remove-paragraph convention).
                foreach (var r in para.Elements<Run>().ToList())
                {
                    var ins = new InsertedRun
                    {
                        Author = author,
                        Date = date,
                        Id = GenerateRevisionId(),
                    };
                    para.ReplaceChild(ins, r);
                    ins.AppendChild(r);
                }
            }
        }
        // BUG-DUMP-R44-6: paraMarkIns: <w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>
        // — paragraph-MARK insertion revision (only the pilcrow is a tracked
        // insertion; the run text is plain). Distinct from a content insertion
        // (<w:ins> wrapping runs) handled by the bare-attribution branch above.
        // The dump emits a paraMarkIns.* namespace (mirrors paraMarkDel.*); stamp
        // the mark rPr ONLY and never wrap the runs, so plain run text is not
        // promoted to a tracked insertion. Symmetric with the paraMarkDel block.
        string? pmiAuthor = null, pmiDate = null, pmiId = null;
        bool hasPmiNs = properties.TryGetValue("paraMarkIns.author", out pmiAuthor);
        hasPmiNs |= properties.TryGetValue("paraMarkIns.date", out pmiDate);
        hasPmiNs |= properties.TryGetValue("paraMarkIns.id", out pmiId);
        if (hasPmiNs)
        {
            var author = string.IsNullOrEmpty(pmiAuthor) ? "OfficeCLI" : pmiAuthor!;
            // BUG-R4F-03: RoundtripKind keeps a …Z date in Utc for byte-identical
            // round-trip (see the bare-attribution branch above).
            DateTime date = !string.IsNullOrEmpty(pmiDate)
                && DateTime.TryParse(pmiDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var hd3)
                ? hd3 : DateTime.UtcNow;
            // CONSISTENCY(pmrp-append): append the mark rPr (lands after pStyle);
            // prepend the <w:ins> inside the rPr (CT_ParaRPr ins/del/move group
            // leads). Mirrors the paraMarkDel block below.
            var pMarkRPr3 = pProps.ParagraphMarkRunProperties
                          ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            // BUG-DUMP-R71-PARAMARK-INSDEL-ORDER: seat via schema-order helper,
            // not PrependChild. A paragraph mark CAN carry both ins and del
            // (inserted by one reviewer, later deleted by another); two blind
            // prepends order them by execution (del first), but CT_ParaRPr
            // requires ins before del. The helper places each correctly.
            if (pMarkRPr3.GetFirstChild<Inserted>() == null)
            {
                InsertRunPropInSchemaOrder(pMarkRPr3, new Inserted
                {
                    Author = author,
                    Date = date,
                    Id = !string.IsNullOrEmpty(pmiId) ? pmiId : GenerateRevisionId(),
                });
            }
        }
        // paraMarkDel: <w:pPr><w:rPr><w:del .../></w:rPr></w:pPr> — paragraph-
        // mark deletion revision (paragraph join). Accept either explicit
        // `revision.type=paraMarkDel` (or schema-short `paraMarkDel`) plus
        // revision.author/.date/.id, OR a `paraMarkDel.*` namespace that
        // mirrors the paraMarkIns.* readback. Without this branch the dump
        // emitted by the new Get-side paraMarkDel.* readback round-tripped
        // back through AddParagraph as plain props and the ¶-del marker was
        // silently dropped.
        string? pmdAuthor = null, pmdDate = null, pmdId = null;
        bool hasPmdNs = properties.TryGetValue("paraMarkDel.author", out pmdAuthor);
        hasPmdNs |= properties.TryGetValue("paraMarkDel.date", out pmdDate);
        hasPmdNs |= properties.TryGetValue("paraMarkDel.id", out pmdId);
        bool isParaMarkDelType = pTcKind != null &&
            (pTcKind.Trim().Equals("paraMarkDel", StringComparison.OrdinalIgnoreCase)
             || pTcKind.Trim().Equals("paragraphMarkDeletion", StringComparison.OrdinalIgnoreCase)
             || pTcKind.Trim().Equals("del", StringComparison.OrdinalIgnoreCase));
        if (hasPmdNs || isParaMarkDelType)
        {
            // Fall through to revision.author/date/id when the namespaced
            // form wasn't provided (caller used revision.type=paraMarkDel).
            if (!hasPmdNs)
            {
                properties.TryGetValue("revision.author", out pmdAuthor);
                properties.TryGetValue("revision.date", out pmdDate);
                properties.TryGetValue("revision.id", out pmdId);
            }
            var author = string.IsNullOrEmpty(pmdAuthor) ? "OfficeCLI" : pmdAuthor!;
            // BUG-R4F-03: RoundtripKind keeps a …Z date in Utc (see above).
            DateTime date = !string.IsNullOrEmpty(pmdDate)
                && DateTime.TryParse(pmdDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var hd2)
                ? hd2 : DateTime.UtcNow;
            // Append (not prepend) the paragraph-mark rPr. In CT_PPr the
            // paragraph-mark <w:rPr> sits near the END of the sequence (after
            // pStyle / numPr / jc / …); PrependChild forced it to position 0,
            // ahead of a pStyle already set above, producing schema-invalid
            // `<w:pPr><w:rPr/><w:pStyle/></w:pPr>` that Word rejects on open.
            // AppendChild matches every other ParagraphMarkRunProperties site
            // in this file and lands the rPr after pStyle. CONSISTENCY(pmrp-append).
            var pMarkRPr2 = pProps.ParagraphMarkRunProperties
                          ?? pProps.AppendChild(new ParagraphMarkRunProperties());
            // Don't double-emit if a Deleted element already lives here.
            // BUG-DUMP-R71-PARAMARK-INSDEL-ORDER: seat the <w:del> via the
            // schema-order helper. CT_ParaRPr leads with the ins/del/move group,
            // so del must precede any markRPr.* props (rFonts/sz/…) already in
            // the rPr. The earlier assumption that a paragraph mark is "never
            // both inserted and deleted" is false — real review chains delete a
            // previously-inserted mark, leaving both ins (id A) and del (id B);
            // blind prepend then puts del before ins, which CT_ParaRPr rejects.
            // The helper orders ins-then-del regardless of application order.
            if (pMarkRPr2.GetFirstChild<Deleted>() == null)
            {
                InsertRunPropInSchemaOrder(pMarkRPr2, new Deleted
                {
                    Author = author,
                    Date = date,
                    Id = !string.IsNullOrEmpty(pmdId) ? pmdId : GenerateRevisionId(),
                });
            }
        }
        // BUG-DUMP-R49-1: numPrIns.*: <w:numPr><w:ins .../> records that the
        // paragraph's list-numbering assignment was inserted as a tracked change
        // (Reviewing pane: "Formatted: List Paragraph"). Stamp the <w:ins> child
        // inside the existing w:numPr so the revision is faithfully round-tripped.
        // Mirrors the paraMarkIns.* approach: no wrapping of text, only a marker
        // inside the structural element. Added after numId/numLevel (which must
        // be in numPr before this fires), so the w:ins sibling is present.
        {
            string? npiAuthor = null, npiDate = null, npiId = null;
            bool hasNpiNs = properties.TryGetValue("numPrIns.author", out npiAuthor);
            hasNpiNs |= properties.TryGetValue("numPrIns.date", out npiDate);
            hasNpiNs |= properties.TryGetValue("numPrIns.id", out npiId);
            if (hasNpiNs)
            {
                // BUG-DUMP-H77: a tracked numbering-insertion (<w:numPr><w:ins/></w:numPr>)
                // frequently carries NO numId and NO ilvl, so the numId/numLevel blocks
                // above never created the numPr. Materialize an empty numPr here so the
                // <w:ins> marker round-trips (CT_NumPr permits ilvl?, numId?, ins?).
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                if (numPr.GetFirstChild<Inserted>() == null)
                {
                    var author = string.IsNullOrEmpty(npiAuthor) ? "OfficeCLI" : npiAuthor!;
                    DateTime date = !string.IsNullOrEmpty(npiDate)
                        && DateTime.TryParse(npiDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var npiDt)
                        ? npiDt : DateTime.UtcNow;
                    // CT_NumPr schema order: ilvl?, numId?, ins? — append after
                    // ilvl/numId so the element lands at the correct position.
                    numPr.AppendChild(new Inserted
                    {
                        Author = author,
                        Date = date,
                        Id = !string.IsNullOrEmpty(npiId) ? npiId : GenerateRevisionId(),
                    });
                }
            }
        }
        return resultPath;
    }

    // BUG-R6A(BUG1): block-level containers that host w:p children and therefore
    // cannot carry m:oMath / m:oMathPara as a direct child. The equation-add path
    // must wrap math in a w:p for all of these. Body is included because the
    // inline path already wraps there; for display, Body uniquely tolerates a
    // bare m:oMathPara child (schema-legal) and is handled by its own branch.
    private static bool IsMathBlockContainer(OpenXmlElement parent) =>
        parent is Body or SdtBlock or Footnote or Endnote or Header or Footer
              or TextBoxContent or Comment or SdtContentBlock;

    // A plain-text SDT (<w:text/> in sdtPr) may legally hold only a single run of
    // plain text — adding an equation (math / a block paragraph) makes the file
    // one Word refuses to open, even though the SDK validator does not flag it.
    // Adding an equation implies rich content, so drop the <w:text/> restriction,
    // turning it into a rich-text SDT that accepts block-level content.
    private static void UpgradeTextSdtForBlockContent(OpenXmlElement parent)
    {
        var sdt = (parent as SdtContentBlock)?.Parent as SdtBlock
                  ?? parent.Ancestors<SdtBlock>().FirstOrDefault();
        sdt?.SdtProperties?.GetFirstChild<SdtContentText>()?.Remove();
    }

    // BUG-DUMP-COMMENT-IN-MATH: remove comment-range markers (commentRangeStart /
    // commentRangeEnd / commentReference, any prefix) from a verbatim math fragment
    // before it is reconstructed. These carry stale source comment ids that no
    // longer exist after comments.xml is renumbered; the owning comment is
    // re-anchored separately via AddComment. Leaving an empty run wrapper behind is
    // schema-legal (it renders nothing).
    // Strip self-closing comment-range markers (commentRangeStart/End/Reference)
    // from a verbatim XML slice. Used by both the equation (oMath) carrier and the
    // SDT carrier: a comment marker that rides along verbatim keeps its SOURCE id,
    // but comments are renumbered dense + re-anchored separately via EmitComments/
    // AddComment, so the verbatim marker becomes a dangling stale-id reference
    // (schema-invalid; the rebuilt doc fails to open in Word). The regex is
    // namespace-agnostic (\w+:) so it covers w:/any prefix.
    internal static string StripVerbatimCommentMarkers(string omml)
    {
        if (omml.IndexOf("comment", StringComparison.OrdinalIgnoreCase) < 0) return omml;
        return System.Text.RegularExpressions.Regex.Replace(
            omml,
            @"<\w+:comment(?:RangeStart|RangeEnd|Reference)\b[^>]*/>",
            string.Empty);
    }

    private string AddEquation(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        string resultPath;
        OpenXmlElement? newElement;
        // Accept `latex=` and `math=` as property aliases for `formula=` — both
        // are pervasive in docs/usage and read naturally for an equation, and the
        // TrackingPropertyDictionary marks them consumed so no false
        // unsupported_property warning fires (handler-as-truth).
        if (!properties.TryGetValue("formula", out var formula)
            && !properties.TryGetValue("text", out formula)
            && !properties.TryGetValue("latex", out formula)
            && !properties.TryGetValue("math", out formula))
            throw new ArgumentException(
                "'formula' (or 'text' / 'latex' / 'math') property is required for equation type");

        // A run (w:r) cannot host m:oMath/m:oMathPara — appending one produced a
        // schema-invalid file Word refuses to open. Reject with a clear message
        // rather than silently corrupting (exit 0).
        if (parent is Run)
            throw new ArgumentException(
                "Cannot add an equation to a run (w:r). Valid parents: /body, a paragraph, "
                + "a hyperlink, footnote, endnote, header/footer, textbox, comment, or sdtContent.");

        // If the equation lands in a plain-text SDT, relax it to rich-text so the
        // file stays openable in Word (see UpgradeTextSdtForBlockContent).
        UpgradeTextSdtForBlockContent(parent);

        // R2-fuzz-3: validate `mode` like Set does. Accept inline/display
        // case-insensitively (mode=INLINE works); any other value is reported
        // as unsupported (warning + exit 2) instead of silently defaulting to
        // display. CONSISTENCY: mirrors SetElementMPara/SetElementOMath's
        // "mode (valid: inline, display)" rejection.
        var mode = properties.GetValueOrDefault("mode", "display").ToLowerInvariant();
        if (properties.ContainsKey("mode") && mode is not ("inline" or "display"))
        {
            LastAddUnsupportedProps.Add("mode (valid: inline, display)");
            mode = "display"; // fall back so the equation is still written
        }

        // BUG-DUMP-EQVERBATIM: prefer the verbatim <m:oMath> the dump captured
        // (xml prop) over the LaTeX `formula` string. The formula string is lossy
        // — it drops the per-run <w:rPr> on every <m:r> (most consequentially
        // rFonts="Cambria Math", so a rebuilt equation renders in the body font at
        // the wrong size) and simplifies some structures. Fall back to
        // FormulaParser for the interactive `add equation formula=` path (no xml).
        M.OfficeMath BuildSourceOMath()
        {
            if ((properties.TryGetValue("xml", out var omml) || properties.TryGetValue("omml", out omml))
                && !string.IsNullOrEmpty(omml) && omml.Contains("oMath", StringComparison.Ordinal))
            {
                // BUG-DUMP-COMMENT-IN-MATH: a comment range whose End/Reference run
                // sits INSIDE this equation rides along in the verbatim <m:oMath>,
                // carrying the SOURCE comment id. But comments.xml is renumbered
                // dense on replay and the comment is re-anchored separately via
                // AddComment, so the math-borne marker keeps a now-nonexistent id —
                // a dangling reference (silent comment loss) plus a duplicate-id
                // desync (schema-invalid). Strip the comment-range markers from the
                // verbatim math; the comment survives through its AddComment anchor
                // (its end lands at the run boundary adjacent to the equation).
                omml = StripVerbatimCommentMarkers(omml);
                // BUG-DUMP-OLE-IN-OMATH: a MathType / Equation.DSMT4 OLE object
                // embedded inside the verbatim math references its binary payload
                // (<o:OLEObject r:id>) and preview image (<v:imagedata r:id>) by
                // relationship id. The equation emit base64-inlines those parts as
                // part{N}.* (mirroring the activex / vmlshape carriers); without
                // recreating them here the r:ids dangle on replay — a silent
                // embedding loss plus a validator NullReferenceException. Recreate
                // the parts on the host part and rewrite the math's r:ids to the
                // freshly assigned ones. Only the verbatim-xml path carries these.
                if (properties.ContainsKey("part1.relId"))
                {
                    var oleHostPart = ResolveImageHostPart(parent);
                    var rewriteOleIds = MaterializeInlinedParts(oleHostPart, properties, "equation");
                    omml = rewriteOleIds(omml);
                }
                try
                {
                    // Root is <m:oMath> → construct directly; root is <m:oMathPara>
                    // (display capture) → lift its inner <m:oMath>.
                    var frag = new M.OfficeMath(omml);
                    return frag;
                }
                catch
                {
                    try
                    {
                        var wrapped = new M.Paragraph(omml).GetFirstChild<M.OfficeMath>()
                            ?? new DocumentFormat.OpenXml.OpenXmlUnknownElement(omml)
                                .Descendants<M.OfficeMath>().FirstOrDefault();
                        if (wrapped != null) return (M.OfficeMath)wrapped.CloneNode(true);
                    }
                    catch { /* malformed — fall through to the formula string */ }
                }
            }
            // R3-fuzz-1: lenient parse — a too-deep/unparseable formula records
            // a warning (exit 2) and writes a placeholder instead of throwing
            // (exit 1 / whole-batch failure).
            var parsed = FormulaParser.ParseLenient(formula, LastUnrecognizedLatex);
            return parsed as M.OfficeMath ?? new M.OfficeMath(parsed.CloneNode(true));
        }

        if (mode == "inline" && parent is Paragraph inlinePara)
        {
            // Insert inline math into existing paragraph
            inlinePara.AppendChild(BuildSourceOMath());
            var mathCount = inlinePara.Elements<M.OfficeMath>().Count();
            resultPath = $"{parentPath}/oMath[{mathCount}]";
            newElement = inlinePara;
        }
        else if (mode == "inline" && parent is Hyperlink inlineHl)
        {
            // BUG-DUMP15-04: m:oMath nested inside w:hyperlink dump→batch
            // round-trip. AddEquation accepts a hyperlink parent so the
            // emitter can replay the equation INSIDE the hyperlink rather
            // than alongside it.
            inlineHl.AppendChild(BuildSourceOMath());
            var mathCount = inlineHl.Elements<M.OfficeMath>().Count();
            // BUG-R4-1: emit the RESOLVABLE oMath[N] segment (the resolver
            // matches the m:oMath element by LocalName). The non-hyperlink
            // inline branch above was fixed in R3-bt-1; this hyperlink-parent
            // branch still returned the unresolvable equation[N].
            resultPath = $"{parentPath}/oMath[{mathCount}]";
            newElement = inlineHl;
        }
        else if (mode == "inline" && IsMathBlockContainer(parent))
        {
            // BUG-R6A(BUG1): inline math under a block container (Body, SdtBlock,
            // footnote/endnote/header/footer) must be wrapped in a w:p — these
            // containers cannot host m:oMath/m:oMathPara as a direct child (only
            // Body tolerates a bare m:oMathPara, which masked the bug for the
            // others). Emit a bare m:oMath instead of m:oMathPara so the math
            // renders as inline-with-text rather than as a centered display
            // equation.
            M.OfficeMath inlineOMath = BuildSourceOMath();
            var hostPara = new Paragraph(inlineOMath);
            AssignParaId(hostPara);
            if (index.HasValue)
            {
                var children = parent.ChildElements.ToList();
                if (index.Value < children.Count)
                    parent.InsertBefore(hostPara, children[index.Value]);
                else
                    AppendToParent(parent, hostPara);
            }
            else
            {
                AppendToParent(parent, hostPara);
            }
            var pIdx = parent.Elements<Paragraph>().Count();
            resultPath = $"{parentPath}/{BuildParaPathSegment(hostPara, pIdx)}/oMath[1]";
            newElement = hostPara;
        }
        else
        {
            // Display mode: create m:oMathPara
            M.OfficeMath oMath = BuildSourceOMath();

            var mathPara = new M.Paragraph(oMath);

            // BUG-DUMP-EQDISPLAY-PPR: re-apply the source wrapper paragraph's line
            // spacing / before-after onto the rebuilt wrapper <w:p>. Call BEFORE any
            // bidi PrependChild / mark-rPr append so CT_PPr schema order holds
            // (bidi < spacing < jc < rPr). Without this a 1.5x display-equation
            // line collapsed to single spacing, compressing the page on round-trip.
            void ApplyEqWrapperSpacing(Paragraph wp)
            {
                if (properties == null) return;
                // CONSISTENCY(verbatim-ppr-supersede): when the dump carried the
                // whole wrapper <w:pPr> verbatim, restore it intact (spacing, jc,
                // pStyle AND the paragraph-mark <w:rPr> that sets the equation
                // line height). Re-applying only the granular spacing/jc keys
                // dropped the mark rPr and shifted the line box. The verbatim
                // pPr already contains bidi when the source had it, so the RTL
                // cascade below only ever adds a missing one.
                if (properties.TryGetValue("wrapperPpr", out var wpprXml)
                    && !string.IsNullOrWhiteSpace(wpprXml)
                    && wpprXml.Contains("pPr", StringComparison.Ordinal))
                {
                    try
                    {
                        var restored = new ParagraphProperties(wpprXml);
                        if (wp.ParagraphProperties != null) wp.ParagraphProperties.Remove();
                        wp.PrependChild(restored);
                        return;
                    }
                    catch { /* fall through to granular re-apply on malformed XML */ }
                }
                SpacingBetweenLines? EnsureSp()
                {
                    var sp = wp.ParagraphProperties ??= new ParagraphProperties();
                    return sp.SpacingBetweenLines ??= new SpacingBetweenLines();
                }
                if (properties.TryGetValue("lineSpacing", out var lsE) || properties.TryGetValue("linespacing", out lsE))
                {
                    var sbl = EnsureSp()!;
                    var (tw, mult) = SpacingConverter.ParseWordLineSpacing(lsE);
                    sbl.Line = tw.ToString();
                    sbl.LineRule = mult ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
                }
                if (properties.TryGetValue("lineRule", out var lrE) || properties.TryGetValue("linerule", out lrE))
                    EnsureSp()!.LineRule = ParseLineRule(lrE);
                if (properties.TryGetValue("spaceBefore", out var sbE) || properties.TryGetValue("spacebefore", out sbE))
                    EnsureSp()!.Before = SpacingConverter.ParseWordSpacing(sbE).ToString();
                if (properties.TryGetValue("spaceAfter", out var saE) || properties.TryGetValue("spaceafter", out saE))
                    EnsureSp()!.After = SpacingConverter.ParseWordSpacing(saE).ToString();
                // BUG-DUMP-EQWRAP-JC: re-apply the wrapper paragraph's own
                // justification (distinct from the math align). Schema order in
                // CT_PPr is spacing < jc, and this runs before bidi/mark-rPr, so
                // appending jc here is order-safe.
                if (properties.TryGetValue("wrapperAlign", out var waE) && !string.IsNullOrWhiteSpace(waE))
                {
                    var jc = waE.Trim().ToLowerInvariant() switch
                    {
                        "justify" or "both" => "both",
                        "center" => "center",
                        "right" or "end" => "right",
                        "left" or "start" => "left",
                        _ => waE.Trim()
                    };
                    var pp = wp.ParagraphProperties ??= new ParagraphProperties();
                    pp.Justification = new Justification { Val = new EnumValue<JustificationValues>(
                        new JustificationValues(jc)) };
                }
            }

            // BUG-DUMP19-02: apply m:oMathParaPr/m:jc when caller passes `align`
            // so block-equation alignment round-trips. Schema requires
            // m:oMathParaPr to precede m:oMath inside m:oMathPara.
            if (properties != null && properties.TryGetValue("align", out var alignVal)
                && !string.IsNullOrWhiteSpace(alignVal))
            {
                var jcVal = alignVal.Trim().ToLowerInvariant() switch
                {
                    "left" => M.JustificationValues.Left,
                    "right" => M.JustificationValues.Right,
                    "center" or "centre" => M.JustificationValues.Center,
                    "centergroup" => M.JustificationValues.CenterGroup,
                    _ => throw new ArgumentException(
                        $"Invalid equation align value: '{alignVal}'. Valid: left, center, right, centerGroup.")
                };
                mathPara.PrependChild(new M.ParagraphProperties(
                    new M.Justification { Val = jcVal }));
            }

            // Display equation must be a direct child of Body (wrapped in w:p).
            // If parent is a Paragraph, insert after that paragraph as a sibling.
            var insertTarget = parent;
            OpenXmlElement? insertAfter = null;
            if (parent is Paragraph parentPara)
            {
                insertTarget = parentPara.Parent ?? parent;
                insertAfter = parentPara;
            }

            if (insertTarget is Body || insertTarget is SdtBlock)
            {
                // Wrap m:oMathPara in w:p for schema validity
                var wrapPara = new Paragraph(mathPara);
                AssignParaId(wrapPara);
                ApplyEqWrapperSpacing(wrapPara);

                // CONSISTENCY(rtl-cascade): inherit pPr/bidi and paragraph-mark
                // rPr/rtl from the host paragraph so the wrapper preserves the
                // surrounding RTL flow. Without this, an equation inserted
                // into an Arabic paragraph silently breaks document direction
                // (mark anchors LTR, page side flips).
                if (parent is Paragraph parentParaForBidi
                    && parentParaForBidi.ParagraphProperties is { } parentPPr)
                {
                    var parentBidi = parentPPr.GetFirstChild<BiDi>();
                    var parentMarkRtl = parentPPr.ParagraphMarkRunProperties?
                        .GetFirstChild<RightToLeftText>();
                    if (parentBidi != null || parentMarkRtl != null)
                    {
                        var wrapPPr = wrapPara.ParagraphProperties ??= new ParagraphProperties();
                        if (parentBidi != null && wrapPPr.GetFirstChild<BiDi>() == null)
                            wrapPPr.PrependChild(new BiDi());
                        if (parentMarkRtl != null)
                        {
                            var markRPr = wrapPPr.ParagraphMarkRunProperties
                                ?? wrapPPr.AppendChild(new ParagraphMarkRunProperties());
                            if (markRPr.GetFirstChild<RightToLeftText>() == null)
                                markRPr.AppendChild(new RightToLeftText());
                        }
                    }
                }
                if (insertAfter != null)
                {
                    insertTarget.InsertAfter(wrapPara, insertAfter);
                }
                else if (index.HasValue)
                {
                    var children = insertTarget.ChildElements.ToList();
                    if (index.Value < children.Count)
                        insertTarget.InsertBefore(wrapPara, children[index.Value]);
                    else
                        AppendToParent(insertTarget, wrapPara);
                }
                else
                {
                    AppendToParent(insertTarget, wrapPara);
                }
                // Compute doc-order index matching NavigateToElement's /body/oMathPara[N]
                // resolution: enumerate bare M.Paragraph and pure oMathPara wrapper w:p's.
                var oMathParaOrdinal = 0;
                var found = 0;
                foreach (var el in insertTarget.ChildElements)
                {
                    if (el is M.Paragraph)
                    {
                        oMathParaOrdinal++;
                        if (ReferenceEquals(el, mathPara)) { found = oMathParaOrdinal; break; }
                    }
                    else if (el is Paragraph wp && IsOMathParaWrapperParagraph(wp))
                    {
                        oMathParaOrdinal++;
                        if (ReferenceEquals(el, wrapPara)) { found = oMathParaOrdinal; break; }
                    }
                }
                if (found == 0) found = oMathParaOrdinal; // fallback
                var bodyPath = insertAfter != null ? parentPath.Substring(0, parentPath.LastIndexOf('/')) : parentPath;
                resultPath = $"{bodyPath}/oMathPara[{found}]";
            }
            else if (IsMathBlockContainer(insertTarget))
            {
                // BUG-R6A(BUG1): block containers other than Body/SdtBlock
                // (footnote/endnote/header/footer) cannot host m:oMathPara as a
                // direct child — OOXML requires math to live inside a w:p. Wrap
                // the m:oMathPara in a w:p exactly like the Body path, but use a
                // simple paragraph-relative result path (these containers don't
                // flatten to /oMathPara[N] the way Body does in NavigateToElement).
                var wrapPara = new Paragraph(mathPara);
                AssignParaId(wrapPara);
                ApplyEqWrapperSpacing(wrapPara);
                if (index.HasValue)
                {
                    var children = insertTarget.ChildElements.ToList();
                    if (index.Value < children.Count)
                        insertTarget.InsertBefore(wrapPara, children[index.Value]);
                    else
                        AppendToParent(insertTarget, wrapPara);
                }
                else
                {
                    AppendToParent(insertTarget, wrapPara);
                }
                // When the caller targeted a PARAGRAPH inside the container
                // (insertAfter set), parentPath points at that child paragraph,
                // not the container — strip the trailing segment so the result is
                // /header/p[@paraId=X]/oMathPara[1], not the doubly-nested
                // /header/p[target]/p[@paraId=X]/oMathPara[1] (unresolvable).
                // Mirrors the Body branch's bodyPath derivation.
                var containerPath = insertAfter != null
                    ? parentPath.Substring(0, parentPath.LastIndexOf('/'))
                    : parentPath;
                var pIdx = insertTarget.Elements<Paragraph>().Count();
                resultPath = $"{containerPath}/{BuildParaPathSegment(wrapPara, pIdx)}/oMathPara[1]";
            }
            else
            {
                // Cell display equation: the m:oMathPara is appended INTO the
                // existing host paragraph (the cell paragraph), so that paragraph
                // IS the wrapper. Re-apply its spacing/justification here — the
                // new-wrapPara branches above never run for this case, so without
                // this a cell equation lost its line height + jc, collapsing the
                // line and drifting later content across page boundaries.
                if (parent is Paragraph cellWrapPara)
                    ApplyEqWrapperSpacing(cellWrapPara);
                AppendToParent(parent, mathPara);
                resultPath = $"{parentPath}/oMathPara[1]";
            }
            newElement = mathPara;
        }

        return resultPath;
    }

    private string AddRun(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // See RejectBareRevisionKey: the bare `revision=` literal was retired
        // when creation/action split into revision.type / revision.action.
        RejectBareRevisionKey(properties);
        string resultPath;
        // BUG-DUMP33-01: support <w:hyperlink> as a run parent so dump→batch
        // can round-trip tab-only / formatted runs that live inside a
        // hyperlink wrapper (Navigation surfaces them with hyperlink-scoped
        // _hyperlinkParent and WordBatchEmitter rebases the parent path).
        Hyperlink? targetHyperlink = null;
        Paragraph? targetPara = parent as Paragraph;
        if (targetPara == null && parent is Hyperlink hlParent && hlParent.Parent is Paragraph hlEnclosingPara)
        {
            targetHyperlink = hlParent;
            targetPara = hlEnclosingPara;
        }
        if (targetPara == null)
            throw new ArgumentException("Runs can only be added to paragraphs");

        // BUG-DUMP5-10: revision attribution from dump round-trip.
        // WordBatchEmitter emits revision / revision.author / revision.date
        // on the run when the source run sat inside a <w:ins>/<w:del>
        // wrapper. Without consuming these here, the dotted fallback below
        // dispatches them through TypedAttributeFallback.TrySet — which has
        // no rPr attribute to bind them to — and they're marked UNSUPPORTED,
        // dropping the wrapper entirely on replay.
        string? trackChangeKind = null;
        string? trackChangeAuthor = null;
        string? trackChangeDate = null;
        string? trackChangeId = null;
        if (properties.TryGetValue("revision.type", out var tcKindRaw))
            trackChangeKind = tcKindRaw?.Trim().ToLowerInvariant();
        properties.TryGetValue("revision.author", out trackChangeAuthor);
        properties.TryGetValue("revision.date", out trackChangeDate);
        properties.TryGetValue("revision.id", out trackChangeId);
        // BUG-DUMP-DELININS: a run that is BOTH inserted and deleted
        // (<w:ins><w:del><w:r><w:delText>) — one reviewer inserts text, another
        // deletes that insertion. The dump captures the outer ins as
        // revision.* and the inner del as revision.nested.*. Without rebuilding
        // both wrappers the deletion is lost and the text un-deletes.
        string? nestedTcKind = null, nestedTcAuthor = null, nestedTcDate = null, nestedTcId = null;
        if (properties.TryGetValue("revision.nested.type", out var nTcKindRaw))
            nestedTcKind = nTcKindRaw?.Trim().ToLowerInvariant();
        properties.TryGetValue("revision.nested.author", out nestedTcAuthor);
        properties.TryGetValue("revision.nested.date", out nestedTcDate);
        properties.TryGetValue("revision.nested.id", out nestedTcId);

        // High-level inference: if a revision.* sub-key is present
        // (author/date/id) without an explicit `revision.type=<kind>` literal,
        // default to "ins" — `add run + revision.author=X` means "create a
        // new run as a tracked insertion". Mirrors Word UI: any edit while
        // track-changes is on becomes a revision; for an `add run` op the
        // only natural revision kind is insertion. Format / moveFrom /
        // moveTo still require the explicit literal because they're not
        // implied by `add`.
        if (string.IsNullOrEmpty(trackChangeKind)
            && (!string.IsNullOrEmpty(trackChangeAuthor)
                || !string.IsNullOrEmpty(trackChangeDate)
                || !string.IsNullOrEmpty(trackChangeId)))
        {
            trackChangeKind = "ins";
        }

        var newRun = new Run();
        var newRProps = new RunProperties();
        // Per-script font slots (font.latin/ea/cs) compose with bare 'font'.
        // Mirrors AddParagraph's run-creation block.
        string? nrAscii = null, nrHAnsi = null, nrEa = null, nrCs = null;
        if (properties.TryGetValue("font", out var rFont) || properties.TryGetValue("font.name", out rFont))
        { nrAscii = rFont; nrHAnsi = rFont; nrEa = rFont; }
        if (properties.TryGetValue("font.latin", out var rfLatin))
        { nrAscii = rfLatin; nrHAnsi = rfLatin; }
        if (properties.TryGetValue("font.ea", out var rfEa)
            || properties.TryGetValue("font.eastasia", out rfEa)
            || properties.TryGetValue("font.eastasian", out rfEa))
        { nrEa = rfEa; }
        if (properties.TryGetValue("font.cs", out var rfCs)
            || properties.TryGetValue("font.complexscript", out rfCs)
            || properties.TryGetValue("font.complex", out rfCs))
        { nrCs = rfCs; }
        // BUG-DUMP24-01: theme-font slot support — bind a run to a theme
        // major/minor font (rFonts/@*Theme) instead of a literal face.
        // Mirrors AddParagraph text-bearing block.
        string? nrAsciiTheme = null, nrHAnsiTheme = null, nrEaTheme = null, nrCsTheme = null;
        if (properties.TryGetValue("font.asciiTheme", out var rfAT) || properties.TryGetValue("font.asciitheme", out rfAT))
            nrAsciiTheme = rfAT;
        if (properties.TryGetValue("font.hAnsiTheme", out var rfHAT) || properties.TryGetValue("font.hansitheme", out rfHAT))
            nrHAnsiTheme = rfHAT;
        if (properties.TryGetValue("font.eaTheme", out var rfEAT) || properties.TryGetValue("font.eatheme", out rfEAT) || properties.TryGetValue("font.eastasiatheme", out rfEAT))
            nrEaTheme = rfEAT;
        if (properties.TryGetValue("font.csTheme", out var rfCST) || properties.TryGetValue("font.cstheme", out rfCST))
            nrCsTheme = rfCST;
        // BUG-DUMP-R31-2: <w:rFonts w:hint> font-slot selector on an explicit
        // `add r` run. Bind it on the SAME curated RunFonts as the other slots
        // (not via the generic TypedAttributeFallback at the tail) so a run
        // carrying ONLY a hint (no ea/ascii/…) still materializes a <w:rFonts>
        // — TypedAttributeFallback's generic binding for a hint-only run is
        // fragile (the project conventions: schema reflection is a last-resort fallback,
        // not the canonical path). The hint composes with ascii/hAnsi/ea/cs.
        string? nrHint = properties.TryGetValue("font.hint", out var rfHintVal) ? rfHintVal : null;
        if (nrAscii != null || nrHAnsi != null || nrEa != null || nrCs != null
            || nrAsciiTheme != null || nrHAnsiTheme != null || nrEaTheme != null || nrCsTheme != null
            || nrHint != null)
        {
            var nrFonts = new RunFonts();
            if (nrAscii != null) nrFonts.Ascii = nrAscii;
            if (nrHAnsi != null) nrFonts.HighAnsi = nrHAnsi;
            if (nrEa != null) nrFonts.EastAsia = nrEa;
            if (nrCs != null) nrFonts.ComplexScript = nrCs;
            if (nrAsciiTheme != null)
                nrFonts.AsciiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(nrAsciiTheme));
            if (nrHAnsiTheme != null)
                nrFonts.HighAnsiTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(nrHAnsiTheme));
            if (nrEaTheme != null)
                nrFonts.EastAsiaTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(nrEaTheme));
            if (nrCsTheme != null)
                nrFonts.ComplexScriptTheme = new EnumValue<ThemeFontValues>(new ThemeFontValues(nrCsTheme));
            if (nrHint != null)
            {
                var nrHintLower = nrHint.Trim().ToLowerInvariant();
                var nrHintEnum = nrHintLower switch
                {
                    "eastasia" => (FontTypeHintValues?)FontTypeHintValues.EastAsia,
                    "cs" or "complexscript" or "complex" => FontTypeHintValues.ComplexScript,
                    "default" => FontTypeHintValues.Default,
                    _ => null,
                };
                if (nrHintEnum.HasValue) nrFonts.Hint = nrHintEnum.Value;
            }
            newRProps.AppendChild(nrFonts);
        }
        if (properties.TryGetValue("size", out var rSize) || properties.TryGetValue("font.size", out rSize) || properties.TryGetValue("fontsize", out rSize))
            newRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(rSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
        // CONSISTENCY(toggle-explicit-false): mirror AddParagraph text-bearing
        // (BUG-018) — explicit `false` must emit <w:b w:val="false"/> so the
        // run can override a style-asserted toggle. AddRun reaches this block
        // via dump→batch replay of any docx with run-level toggle overrides
        // (Heading1 + non-bold span, Quote + non-italic citation, …).
        if (properties.TryGetValue("bold", out var rBold) || properties.TryGetValue("font.bold", out rBold))
        {
            if (IsTruthy(rBold)) newRProps.Bold = new Bold();
            else if (IsExplicitFalseAddOverride(rBold))
                newRProps.Bold = new Bold { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("bold.cs", out var rBoldCs) || properties.TryGetValue("font.bold.cs", out rBoldCs))
        {
            // BUG-DUMP-BCS-FALSE: honor an explicit complex-script-bold OFF so a
            // run that overrides a bold style (<w:bCs w:val="0"/>) round-trips —
            // mirrors bare bold/italic. On-only dropped the override and the run
            // re-inherited the style's bold (Arabic headings rendered bold).
            if (IsTruthy(rBoldCs)) newRProps.BoldComplexScript = new BoldComplexScript();
            else if (IsExplicitFalseAddOverride(rBoldCs))
                newRProps.BoldComplexScript = new BoldComplexScript { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("italic", out var rItalic) || properties.TryGetValue("font.italic", out rItalic))
        {
            if (IsTruthy(rItalic)) newRProps.Italic = new Italic();
            else if (IsExplicitFalseAddOverride(rItalic))
                newRProps.Italic = new Italic { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("italic.cs", out var rItalicCs) || properties.TryGetValue("font.italic.cs", out rItalicCs))
        {
            // BUG-DUMP-BCS-FALSE: explicit complex-script-italic OFF override
            // (mirrors bold.cs above + bare italic).
            if (IsTruthy(rItalicCs)) newRProps.ItalicComplexScript = new ItalicComplexScript();
            else if (IsExplicitFalseAddOverride(rItalicCs))
                newRProps.ItalicComplexScript = new ItalicComplexScript { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("size.cs", out var rSizeCs) || properties.TryGetValue("font.size.cs", out rSizeCs))
        {
            newRProps.FontSizeComplexScript = new FontSizeComplexScript
            {
                Val = ((int)Math.Round(ParseFontSize(rSizeCs) * 2, MidpointRounding.AwayFromZero)).ToString()
            };
        }
        if (properties.TryGetValue("color", out var rColor) || properties.TryGetValue("font.color", out rColor))
        {
            // CONSISTENCY(theme-color): Add run color accepts scheme color
            // names (accent1, dark2, hyperlink, …); same logic as
            // ApplyRunFormatting in WordHandler.Helpers.cs.
            // CONSISTENCY(color-auto): see WordHandler.Helpers.cs ApplyRunFormatting.
            // BUG-DUMP-R44-1: split the ';themeColor=…' tail first so a direct
            // run color carrying both an explicit hex AND a theme linkage (e.g.
            // "#FFFFFF;themeColor=background1") keeps the hex as w:val and stamps
            // the theme attrs — instead of SanitizeHex mangling the whole string
            // (and the old code collapsing theme-only to garbage). Mirrors the
            // ApplyRunFormatting case "color" fix.
            var (rColorPos, rColorTheme) = ExtractThemeTail(rColor);
            if (string.Equals(rColorPos, "auto", StringComparison.OrdinalIgnoreCase))
            {
                newRProps.Color = new Color { Val = "auto" };
            }
            else if (rColorPos.Length == 0 && rColorTheme.Count > 0)
            {
                newRProps.Color = new Color { Val = "auto" };
            }
            else
            {
                var rSchemeName = OfficeCli.Core.ParseHelpers.NormalizeSchemeColorName(rColorPos);
                if (rSchemeName != null)
                    newRProps.Color = new Color { Val = "auto", ThemeColor = new EnumValue<ThemeColorValues>(new ThemeColorValues(rSchemeName)) };
                else
                    newRProps.Color = new Color { Val = SanitizeHex(rColorPos) };
            }
            ApplyColorTheme(newRProps.Color, rColorTheme);
        }
        if (properties.TryGetValue("underline", out var rUnderline) || properties.TryGetValue("font.underline", out rUnderline))
        {
            var ulVal = NormalizeUnderlineValue(rUnderline);
            newRProps.Underline = new Underline { Val = new UnderlineValues(ulVal) };
        }
        // CONSISTENCY(toggle-explicit-false): see bold/italic above.
        if (properties.TryGetValue("strike", out var rStrike)
                || properties.TryGetValue("strikethrough", out rStrike)
                || properties.TryGetValue("font.strike", out rStrike)
                || properties.TryGetValue("font.strikethrough", out rStrike))
        {
            if (IsTruthy(rStrike)) newRProps.Strike = new Strike();
            else if (IsExplicitFalseAddOverride(rStrike))
                newRProps.Strike = new Strike { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("highlight", out var rHighlight))
            newRProps.Highlight = new Highlight { Val = ParseHighlightColor(rHighlight) };
        if (properties.TryGetValue("caps", out var rCaps)
                || properties.TryGetValue("allcaps", out rCaps)
                || properties.TryGetValue("allCaps", out rCaps))
        {
            if (IsTruthy(rCaps)) newRProps.Caps = new Caps();
            else if (IsExplicitFalseAddOverride(rCaps))
                newRProps.Caps = new Caps { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("smallcaps", out var rSmallCaps) || properties.TryGetValue("smallCaps", out rSmallCaps))
        {
            if (IsTruthy(rSmallCaps)) newRProps.SmallCaps = new SmallCaps();
            else if (IsExplicitFalseAddOverride(rSmallCaps))
                newRProps.SmallCaps = new SmallCaps { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("dstrike", out var rDstrike)
            || properties.TryGetValue("doublestrike", out rDstrike)
            || properties.TryGetValue("doubleStrike", out rDstrike))
        {
            if (IsTruthy(rDstrike)) newRProps.DoubleStrike = new DoubleStrike();
            else if (IsExplicitFalseAddOverride(rDstrike))
                newRProps.DoubleStrike = new DoubleStrike { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("vanish", out var rVanish))
        {
            if (IsTruthy(rVanish)) newRProps.Vanish = new Vanish();
            else if (IsExplicitFalseAddOverride(rVanish))
                newRProps.Vanish = new Vanish { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("outline", out var rOutline))
        {
            if (IsTruthy(rOutline)) newRProps.Outline = new Outline();
            else if (IsExplicitFalseAddOverride(rOutline))
                newRProps.Outline = new Outline { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("shadow", out var rShadow))
        {
            if (IsTruthy(rShadow)) newRProps.Shadow = new Shadow();
            else if (IsExplicitFalseAddOverride(rShadow))
                newRProps.Shadow = new Shadow { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("emboss", out var rEmboss))
        {
            if (IsTruthy(rEmboss)) newRProps.Emboss = new Emboss();
            else if (IsExplicitFalseAddOverride(rEmboss))
                newRProps.Emboss = new Emboss { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("imprint", out var rImprint))
        {
            if (IsTruthy(rImprint)) newRProps.Imprint = new Imprint();
            else if (IsExplicitFalseAddOverride(rImprint))
                newRProps.Imprint = new Imprint { Val = OnOffValue.FromBoolean(false) };
        }
        if (properties.TryGetValue("noproof", out var rNoProof))
        {
            if (IsTruthy(rNoProof)) newRProps.NoProof = new NoProof();
            else if (IsExplicitFalseAddOverride(rNoProof))
                newRProps.NoProof = new NoProof { Val = OnOffValue.FromBoolean(false) };
        }
        // CONSISTENCY(add-set-symmetry): Set surfaces rStyle via the typed-attr
        // fallback; Add must accept it explicitly because the bare-key fallback
        // below skips dotless keys without warning. Without this, dump → batch
        // round-trips silently strip every <w:rStyle/> (BUG-R2-05 / BT-5).
        if (properties.TryGetValue("rStyle", out var rRStyle) || properties.TryGetValue("rstyle", out rRStyle))
        {
            if (!string.IsNullOrEmpty(rRStyle))
                newRProps.RunStyle = new RunStyle { Val = rRStyle };
        }
        if (properties.TryGetValue("rtl", out var rRtl) && IsTruthy(rRtl))
            ApplyRunFormatting(newRProps, "rtl", "true");
        // CONSISTENCY(canonical-key): accept "direction"=rtl|ltr as the
        // canonical alias for run-level rtl, matching paragraph/section
        // input vocabulary and the symmetric Get readback (R16-bt-1).
        else if (properties.TryGetValue("direction", out var rDir)
            || properties.TryGetValue("dir", out rDir))
        {
            var v = rDir?.Trim().ToLowerInvariant();
            if (v == "rtl") ApplyRunFormatting(newRProps, "rtl", "true");
            else if (v == "ltr") ApplyRunFormatting(newRProps, "rtl", "false");
        }
        if (properties.TryGetValue("vertAlign", out var rVertAlign) || properties.TryGetValue("vertalign", out rVertAlign))
        {
            newRProps.VerticalTextAlignment = new VerticalTextAlignment
            {
                Val = rVertAlign.ToLowerInvariant() switch
                {
                    "superscript" or "super" => VerticalPositionValues.Superscript,
                    "subscript" or "sub" => VerticalPositionValues.Subscript,
                    "baseline" => VerticalPositionValues.Baseline,
                    _ => throw new ArgumentException($"Invalid 'vertAlign' value: '{rVertAlign}'. Valid values: superscript, subscript, baseline."),
                }
            };
        }
        if (properties.TryGetValue("superscript", out var rSup) && IsTruthy(rSup))
            newRProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Superscript };
        if (properties.TryGetValue("subscript", out var rSub) && IsTruthy(rSub))
            newRProps.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
        if (properties.TryGetValue("charspacing", out var rCharSp) || properties.TryGetValue("charSpacing", out rCharSp)
            || properties.TryGetValue("letterspacing", out rCharSp) || properties.TryGetValue("letterSpacing", out rCharSp))
        {
            int rCsTwips = rCharSp.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                ? (int)Math.Round(ParseHelpers.SafeParseDouble(rCharSp[..^2], "charspacing") * 20, MidpointRounding.AwayFromZero)
                : (int)Math.Round(ParseHelpers.SafeParseDouble(rCharSp, "charspacing"), MidpointRounding.AwayFromZero);
            newRProps.Spacing = new Spacing { Val = rCsTwips };
        }
        if (properties.TryGetValue("shd", out var rShd) || properties.TryGetValue("shading", out rShd)
            || properties.TryGetValue("fill", out rShd))
        {
            // BUG-DUMP-R41-4: route through the shared ParseShadingValue so the
            // run-level <w:shd> theme-linkage (themeFill=…/themeColor=…) tail
            // round-trips; preserves the prior VAL;FILL;COLOR semantics.
            // CONSISTENCY(shd-canonical-fill): `fill` is the canonical Get key
            // for a solid run shading — accept it as an Add alias so dump→batch
            // (which now carries `fill`) replays via `add run --prop fill=…`.
            newRProps.Shading = ParseShadingValue(rShd);
        }

        // w14 text effects
        var tempRun = new Run();
        tempRun.PrependChild(newRProps);
        if (properties.TryGetValue("textOutline", out var toVal) || properties.TryGetValue("textoutline", out toVal))
            ApplyW14TextEffect(tempRun, "textOutline", toVal, BuildW14TextOutline);
        if (properties.TryGetValue("textFill", out var tfVal) || properties.TryGetValue("textfill", out tfVal))
            ApplyW14TextEffect(tempRun, "textFill", tfVal, BuildW14TextFill);
        if (properties.TryGetValue("w14shadow", out var w14sVal))
            ApplyW14TextEffect(tempRun, "shadow", w14sVal, BuildW14Shadow);
        if (properties.TryGetValue("w14glow", out var w14gVal))
            ApplyW14TextEffect(tempRun, "glow", w14gVal, BuildW14Glow);
        if (properties.TryGetValue("w14reflection", out var w14rVal))
            ApplyW14TextEffect(tempRun, "reflection", w14rVal, BuildW14Reflection);
        // Detach rPr from temp run for re-attachment to actual run
        newRProps.Remove();

        // Inherit default formatting from paragraph mark run properties.
        // CONSISTENCY(markRPr-inherit-opt-out): dump→batch sets the exact
        // run props it observed (no font.ea, no rFonts at all → no
        // inheritance wanted). Caller passes noMarkRPrInherit=true to
        // suppress the markRPr→rPr type-fill so the round-trip preserves
        // the source's "run has no rFonts even though para mark does" shape.
        bool noMarkInherit = properties.TryGetValue("nomarkrprinherit", out var nMri)
                          || properties.TryGetValue("noMarkRPrInherit", out nMri);
        var markRProps = targetPara.ParagraphProperties?.ParagraphMarkRunProperties;
        if (markRProps != null && !(noMarkInherit && IsTruthy(nMri)))
        {
            foreach (var child in markRProps.ChildElements)
            {
                var childType = child.GetType();
                if (newRProps.Elements().All(e => e.GetType() != childType))
                    newRProps.AppendChild(child.CloneNode(true));
            }
        }

        newRun.AppendChild(newRProps);
        // Run-level w14 effects + OpenType typographic toggles (textOutline/
        // textFill/w14shadow/w14glow/w14reflection/ligatures/numForm/numSpacing).
        // AddParagraph routes these through ApplyW14Effects for its implicit run;
        // the explicit `add r` path must do the same or a multi-run paragraph
        // (whose runs each emit as `add r`) drops them. These keys are listed in
        // addRunCuratedBare below so the bare-key fallback doesn't also flag them
        // UNSUPPORTED after ApplyW14Effects consumes them.
        ApplyW14Effects(newRun, properties);
        // BUG-DUMP7-01: a run carrying `sym=font:hex` carries a <w:sym/> glyph.
        // The dump surfaces the resolved Unicode codepoint of that glyph as the
        // LEADING character of `text` (GetRunText walks children in order: the
        // SymbolChar's PUA codepoint, then any literal <w:t>). So `text` is the
        // PUA glyph optionally FOLLOWED by real literal text. Emit the <w:sym/>,
        // then strip exactly the leading PUA glyph and append whatever literal
        // text remains. Appending the PUA glyph as a literal <w:t> would double
        // the visual output (cached glyph in body font + the <w:sym/>); dropping
        // the remaining text entirely (the old behaviour) silently lost a run
        // that mixed <w:sym/> + <w:t>WORLD</w:t> into a sym-only run.
        // BUG-DUMP-R40-2: a run carrying annotationRef=true is the comment
        // reference mark (<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
        // <w:annotationRef/></w:r>) that opens every Word-authored comment body.
        // The rPr (rStyle) is already built above; append the <w:annotationRef/>
        // mark and emit NO <w:t> (the source run carries no literal text — only
        // the mark). Dropping it lost the clickable comment-reference glyph.
        if (properties.TryGetValue("annotationRef", out var annRefRaw) && IsTruthy(annRefRaw))
        {
            newRun.AppendChild(new AnnotationReferenceMark());
        }
        // BUG-DUMP-HYPHEN-CELL: round-trip a STRUCTURAL hyphen element
        // (<w:noBreakHyphen/> / <w:softHyphen/>) so it survives in ANY host —
        // table cells, headers, footers — not just /body. The dump emits
        // `hyphen=noBreak|soft`; the cached glyph (U+2011 / U+00AD) sits at its
        // source position inside `text` (GetRunText surfaces it), so split `text`
        // at that glyph and emit text-before, <element>, text-after in source
        // order — mirroring the <w:sym> interleave handling above. A hyphen-only
        // source run (the common case: <w:r><w:noBreakHyphen/></w:r>) carries no
        // glyph in `text` and emits just the element. Replaces the lossy
        // degrade-to-literal-glyph path for non-/body hyphen runs.
        else if (properties.TryGetValue("hyphen", out var hyphenRaw) && !string.IsNullOrEmpty(hyphenRaw))
        {
            var hyphenKind = hyphenRaw.Trim().ToLowerInvariant();
            OpenXmlElement MakeHyphen() => hyphenKind switch
            {
                "soft" or "softhyphen" or "00ad" => new SoftHyphen(),
                _ => new NoBreakHyphen(), // "nobreak"/"nonbreaking"/"2011"/default
            };
            var glyph = hyphenKind is "soft" or "softhyphen" or "00ad" ? "­" : "‑";
            var runText = properties.GetValueOrDefault("text", "");
            int g = runText.IndexOf(glyph, StringComparison.Ordinal);
            if (g >= 0)
            {
                var before = runText[..g];
                var after = runText[(g + 1)..];
                if (!string.IsNullOrEmpty(before)) AppendTextWithBreaks(newRun, before);
                newRun.AppendChild(MakeHyphen());
                if (!string.IsNullOrEmpty(after)) AppendTextWithBreaks(newRun, after);
            }
            else
            {
                // No cached glyph in `text` — hyphen-only run (or text carries no
                // glyph): emit the element, then any literal text after it.
                newRun.AppendChild(MakeHyphen());
                if (!string.IsNullOrEmpty(runText))
                    AppendTextWithBreaks(newRun, runText);
            }
        }
        else if (properties.TryGetValue("sym", out var symRaw) && !string.IsNullOrEmpty(symRaw))
        {
            var colon = symRaw.LastIndexOf(':');
            string symFont = colon > 0 ? symRaw[..colon] : "";
            string symHex = colon >= 0 ? symRaw[(colon + 1)..] : symRaw;
            var sym = new SymbolChar();
            if (!string.IsNullOrEmpty(symFont)) sym.Font = symFont;
            if (!string.IsNullOrEmpty(symHex)) sym.Char = symHex.ToUpperInvariant();

            // BUG-DUMP-R44-2: the <w:sym> and any literal <w:t> must round-trip
            // in their SOURCE child order. GetRunText walks children in document
            // order, so the dump-cached glyph (the SymbolChar's PUA codepoint)
            // sits at the glyph's ACTUAL position within `text` — leading when
            // <w:sym> precedes <w:t> ("Symbol "), trailing when <w:t>
            // precedes <w:sym> ("Symbol "). The old code unconditionally
            // appended <w:sym> first and only stripped a LEADING glyph, which (a)
            // hoisted the symbol before the text and (b) left a trailing glyph
            // doubled as literal <w:t>. Split `text` at the glyph's index and
            // emit text-before, <w:sym>, text-after in order — preserving the
            // source interleave and stripping exactly the cached glyph wherever
            // it sits.
            var runText = properties.GetValueOrDefault("text", "");
            string glyph = "";
            if (!string.IsNullOrEmpty(symHex)
                && int.TryParse(symHex, System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture, out var symCode))
                glyph = char.ConvertFromUtf32(symCode);

            int g = glyph.Length > 0 ? runText.IndexOf(glyph, StringComparison.Ordinal) : -1;
            if (g >= 0)
            {
                // text-before-glyph → <w:sym> → text-after-glyph (source order).
                var before = runText[..g];
                var after = runText[(g + glyph.Length)..];
                if (!string.IsNullOrEmpty(before)) AppendTextWithBreaks(newRun, before);
                newRun.AppendChild(sym);
                if (!string.IsNullOrEmpty(after)) AppendTextWithBreaks(newRun, after);
            }
            else
            {
                // No cached glyph found in `text` (e.g. unresolvable codepoint)
                // — preserve the legacy shape: <w:sym> first, then any literal
                // text. Only emit a <w:t> when real text survives so a sym-only
                // run never gains a spurious empty <w:t>.
                newRun.AppendChild(sym);
                if (!string.IsNullOrEmpty(runText))
                    AppendTextWithBreaks(newRun, runText);
            }
        }
        else
        {
            var runText = properties.GetValueOrDefault("text", "");
            AppendTextWithBreaks(newRun, runText);
        }

        // BUG-DUMP-PTABTEXT: a run that mixed a <w:ptab/> with text/delText carries
        // the positional tab as inline props (Navigation kept it a `run` so the
        // co-resident text survived). Rebuild the <w:ptab/> ahead of the text
        // (ptab-then-text, the source convention) so BOTH round-trip. The later
        // ins/del wrapper (if any) leaves the ptab in place and only converts <w:t>.
        if (properties.TryGetValue("ptabInline", out var ptInline) && IsTruthy(ptInline))
        {
            var inlinePtab = new PositionalTab();
            if (properties.TryGetValue("ptabInline.align", out var piAlign) && !string.IsNullOrWhiteSpace(piAlign))
                inlinePtab.Alignment = ParsePtabAlignment(piAlign);
            if (properties.TryGetValue("ptabInline.relativeTo", out var piRel) && !string.IsNullOrWhiteSpace(piRel))
                inlinePtab.RelativeTo = ParsePtabRelativeTo(piRel);
            if (properties.TryGetValue("ptabInline.leader", out var piLead) && !string.IsNullOrWhiteSpace(piLead))
                inlinePtab.Leader = ParsePtabLeader(piLead);
            var firstContent = newRun.Elements().FirstOrDefault(e => e is not RunProperties);
            if (firstContent != null) newRun.InsertBefore(inlinePtab, firstContent);
            else newRun.AppendChild(inlinePtab);
        }

        // Dotted-key fallback: same generic helper as Set's run path.
        // Anything still unconsumed after the hand-rolled blocks above
        // gets routed through TypedAttributeFallback; failures land in
        // LastAddUnsupportedProps so the CLI surfaces a WARNING instead
        // of silently dropping. CONSISTENCY(add-set-symmetry).
        // BUG-R7-06: bare run-level keys (bdr / kern / lang shortcuts) that
        // the curated AddRun block above did not consume — route through
        // ApplyRunFormatting so batch replay actually applies them instead
        // of silently dropping. Mirrors the bare-key fallback in
        // AddParagraph (line 670). CONSISTENCY(add-set-symmetry).
        var addRunCuratedBare = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "type", "text", "html", "anchor", "anchorid",
            "font", "size", "fontsize", "fontSize", "bold", "italic", "color", "highlight",
            "underline", "strike", "strikethrough", "doublestrike", "dstrike",
            "vanish", "outline", "shadow", "emboss", "imprint", "noproof",
            "rtl", "vertalign", "superscript", "subscript",
            "charspacing", "letterspacing",
            "caps", "smallcaps", "allcaps",
            "boldcs", "italiccs", "sizecs",
            "shd", "shading", "fill",
            "rstyle", "rStyle",
            "annotationRef", "annotationref",
            "hyphen",
            "textoutline", "textfill", "w14shadow", "w14glow", "w14reflection",
            // OpenType typographic toggles applied via ApplyW14Effects above.
            "ligatures", "numform", "numspacing",
            // R53-A: link / href / url consumed by the post-insertion
            // hyperlink-wrap block below (mirrors pptx Add vocabulary).
            "link", "href", "url",
            // BUG-DUMP5-10: consumed up-front for the w:ins/w:del wrapper
            // emit at the bottom of this method. Bare `revision` is no
            // longer a valid key — creation = `revision.type`, action =
            // `revision.action`.
            "revision.type",
            // BUG-DUMP7-01: consumed up-front to emit <w:sym/> in place of <w:t>.
            "sym",
            // BUG-DUMP-PTABTEXT: consumed above to rebuild an inline <w:ptab/>
            // that shared a run with text.
            "ptabInline", "ptabInline.align", "ptabInline.relativeTo", "ptabInline.leader",
            // CONSISTENCY(markRPr-inherit-opt-out): consumed up-front (line ~1587)
            // to suppress markRPr→rPr type-fill on dump→batch replay. Not a real
            // OOXML attribute — pure inheritance toggle. Without this entry the
            // bare-key fallback flags it UNSUPPORTED on every dump-emitted `add r`.
            "nomarkrprinherit",
        };
        foreach (var (key, value) in properties)
        {
            if (key.Contains('.')) continue;
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            properties.ContainsKey(key);
            if (addRunCuratedBare.Contains(key)) continue;
            if (ApplyRunFormatting(newRProps, key, value)) continue;
            // BUG-DUMP8-07: rescue dump-emitted run props (specVanish,
            // webHidden, effect, em, fitText, position, …) that
            // ApplyRunFormatting has no curated case for but which are
            // typed scalar-val SDK elements. Mirrors the AddParagraph
            // bare-key fallback so dump→batch round-trips through. Only
            // genuinely unknown keys land in LastAddUnsupportedProps.
            if (Core.GenericXmlQuery.TryCreateTypedChild(newRProps, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            properties.ContainsKey(key);
            // CONSISTENCY(font-dotted-alias): font.name/font.bold/font.size/
            // font.italic/font.color/font.underline/font.strike are consumed
            // above by the curated alias blocks; skip the typed-attr fallback
            // so they don't get re-flagged as UNSUPPORTED.
            switch (key.ToLowerInvariant())
            {
                case "font.name":
                case "font.size":
                case "font.bold":
                case "font.italic":
                case "font.color":
                case "font.underline":
                case "font.strike":
                case "font.strikethrough":
                // Per-script slots and CS toggles already consumed above.
                case "font.latin":
                case "font.ea":
                case "font.eastasia":
                case "font.eastasian":
                case "font.cs":
                case "font.complexscript":
                case "font.complex":
                // BUG-DUMP24-01: theme-font slots consumed up-front by the
                // RunFonts theme block above (font.asciiTheme/hAnsiTheme/
                // eaTheme/csTheme); skip the typed-attr fallback so they
                // don't get re-flagged as UNSUPPORTED.
                case "font.asciitheme":
                case "font.hansitheme":
                case "font.eatheme":
                case "font.eastasiatheme":
                case "font.cstheme":
                // BUG-DUMP-R31-2: font.hint consumed by the curated RunFonts
                // block above; skip the typed-attr fallback so it isn't
                // re-applied or flagged UNSUPPORTED.
                case "font.hint":
                // CS run flags (<w:bCs/> / <w:iCs/> / <w:szCs/>) — the
                // run-add block above writes them through ApplyRunFormatting;
                // dotted-fallback can't resolve the dotted name into the
                // OpenXml element type.
                case "bold.cs":
                case "italic.cs":
                case "size.cs":
                case "font.bold.cs":
                case "font.italic.cs":
                case "font.size.cs":
                case "boldcs":
                case "italiccs":
                case "sizecs":
                // BUG-DUMP5-10: consumed up-front for the w:ins/w:del
                // wrapper emit at the bottom of this method.
                // revision.type is dotted (so falls into this loop, not the
                // bare-key loop) — the addRunCuratedBare allowlist above
                // includes "revision.type" but never fires because of the
                // `if (key.Contains('.')) continue;` filter; mirror it here.
                case "revision.type":
                case "revision.author":
                case "revision.date":
                case "revision.id":
                    continue;
            }
            // CONSISTENCY(add-set-symmetry / bcp47-validation): route lang.*
            // through ApplyRunFormatting so the BCP-47 validator that Set
            // applies also runs on Add (without this, malformed lang values
            // like "-" silently became <w:lang w:val="-"/>).
            switch (key.ToLowerInvariant())
            {
                case "lang.latin":
                case "lang.val":
                case "lang.ea":
                case "lang.eastasia":
                case "lang.eastasian":
                case "lang.cs":
                case "lang.complexscript":
                case "lang.bidi":
                // BUG-DUMP-R47-1: underline.color (and aliases) must route
                // through ApplyRunFormatting so the <w:u> lands in CT_RPr
                // schema order (InsertRunPropInSchemaOrder hoists it before any
                // w14 extension block). TypedAttributeFallback below appends at
                // the END of rPr — past an already-emitted <w14:textFill> — which
                // is schema-invalid ("unexpected child w:u"). Mirrors lang.*.
                case "underline.color":
                case "font.underline.color":
                    if (ApplyRunFormatting(newRProps, key, value)) continue;
                    break;
            }
            if (Core.TypedAttributeFallback.TrySet(newRProps, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        // BUG-DUMP-R71-RPR-ORDER: the run rPr was built across mixed paths
        // (SDK setters, ApplyRunFormatting, raw AppendChild for rFonts/sz, and
        // TypedAttributeFallback tail-appends), any of which can leave a child
        // out of CT_RPr order. Normalize once now so the emitted run validates.
        NormalizeRunPropsSchemaOrder(newRProps);

        // Use ChildElements for index lookup so ResolveAnchorPosition's
        // childElement-indexed result lines up. If index points at
        // ParagraphProperties, clamp forward so pPr stays first.
        // BUG-DUMP33-01: when targetHyperlink is set, append/insert inside
        // the hyperlink wrapper instead of directly into the paragraph.
        OpenXmlElement insertHost = (OpenXmlElement?)targetHyperlink ?? targetPara;
        var allChildren = insertHost.ChildElements.ToList();
        if (index.HasValue && index.Value < allChildren.Count)
        {
            var refElement = allChildren[index.Value];
            if (refElement is ParagraphProperties)
            {
                // insert after pPr — i.e. before whatever sits at index+1, else append
                if (index.Value + 1 < allChildren.Count)
                    insertHost.InsertBefore(newRun, allChildren[index.Value + 1]);
                else
                    insertHost.AppendChild(newRun);
            }
            else
            {
                insertHost.InsertBefore(newRun, refElement);
            }
            // CONSISTENCY(run-path-index): match navigation's r[N] enumeration
            // (Descendants<Run>() minus comment-reference runs) via GetAllRuns.
            var runPosIdx = PathIndex.FromArrayIndex(GetAllRuns(targetPara).IndexOf(newRun));
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            // For hyperlink-parented runs, parentPath already includes the
            // hyperlink segment; emit a hyperlink-scoped result path.
            if (targetHyperlink != null)
            {
                var hlIdx = targetPara.Elements<Hyperlink>()
                    .TakeWhile(h => !ReferenceEquals(h, targetHyperlink)).Count() + 1;
                var hlSubIdx = targetHyperlink.Elements<Run>()
                    .TakeWhile(r => !ReferenceEquals(r, newRun)).Count() + 1;
                var hlSegIdx = parentPath.LastIndexOf("/hyperlink[", StringComparison.Ordinal);
                var paraPathOnly = hlSegIdx > 0 ? parentPath.Substring(0, hlSegIdx) : parentPath;
                var paraOnly = ReplaceTrailingParaSegment(paraPathOnly, targetPara);
                resultPath = $"{paraOnly}/hyperlink[{hlIdx}]/r[{hlSubIdx}]";
            }
            else
            {
                resultPath = $"{ReplaceTrailingParaSegment(parentPath, targetPara)}/r[{runPosIdx}]";
            }
        }
        else
        {
            insertHost.AppendChild(newRun);
            if (targetHyperlink != null)
            {
                var hlIdx = targetPara.Elements<Hyperlink>()
                    .TakeWhile(h => !ReferenceEquals(h, targetHyperlink)).Count() + 1;
                var hlSubIdx = targetHyperlink.Elements<Run>()
                    .TakeWhile(r => !ReferenceEquals(r, newRun)).Count() + 1;
                var hlSegIdx = parentPath.LastIndexOf("/hyperlink[", StringComparison.Ordinal);
                var paraPathOnly = hlSegIdx > 0 ? parentPath.Substring(0, hlSegIdx) : parentPath;
                var paraOnly = ReplaceTrailingParaSegment(paraPathOnly, targetPara);
                resultPath = $"{paraOnly}/hyperlink[{hlIdx}]/r[{hlSubIdx}]";
            }
            else
            {
                var runCount = PathIndex.FromArrayIndex(GetAllRuns(targetPara).IndexOf(newRun));
                resultPath = $"{ReplaceTrailingParaSegment(parentPath, targetPara)}/r[{runCount}]";
            }
        }

        // R53-A: AddRun supports a `link=`/`href=`/`url=` shortcut that wraps
        // the newly inserted run in a <w:hyperlink> with the corresponding
        // relationship — same vocabulary the pptx Add accepts and the docx
        // hyperlink Add supports. Without this, link= surfaced as an
        // UNSUPPORTED warning and no rel was created.
        if (targetHyperlink == null
            && (properties.TryGetValue("link", out var runLink)
                || properties.TryGetValue("href", out runLink)
                || properties.TryGetValue("url", out runLink))
            && !string.IsNullOrWhiteSpace(runLink))
        {
            var hlRunHost = ResolveHostPart(targetPara);
            bool runLinkIsFragment = runLink.StartsWith('#');
            Uri? runLinkUri;
            if (runLinkIsFragment)
            {
                runLinkUri = new Uri(runLink, UriKind.Relative);
            }
            else if (Uri.TryCreate(runLink, UriKind.Absolute, out runLinkUri))
            {
                Core.HyperlinkUriValidator.RequireSafeScheme(runLink, "link");
                runLinkUri = new Uri(PercentEncodeUri(runLink), UriKind.Absolute);
            }
            else if (!Uri.TryCreate(runLink, UriKind.Relative, out runLinkUri))
            {
                throw new ArgumentException($"Invalid run link URL '{runLink}'. Expected an absolute URI, relative target, or fragment-only anchor (e.g. '#bookmark').");
            }
            string runHlRelId = hlRunHost.AddHyperlinkRelationship(runLinkUri!, isExternal: !runLinkIsFragment).Id;
            var runHlWrap = new Hyperlink { Id = runHlRelId };
            var newRunParent = newRun.Parent;
            if (newRunParent != null)
            {
                newRunParent.ReplaceChild(runHlWrap, newRun);
                runHlWrap.AppendChild(newRun);
                // Recompute resultPath to point at the run inside the hyperlink.
                var rebuiltHlIdx = targetPara.Elements<Hyperlink>()
                    .TakeWhile(h => !ReferenceEquals(h, runHlWrap)).Count() + 1;
                resultPath = $"{ReplaceTrailingParaSegment(parentPath, targetPara)}/hyperlink[{rebuiltHlIdx}]/r[1]";
            }
        }

        // BUG-DUMP5-10: wrap in w:ins / w:del when the dump asked for
        // track-change attribution. Replace newRun in its parent with the
        // wrapper containing newRun so author/date attribution survives the
        // dump→batch round-trip. The path computed above remains valid:
        // GetAllRuns walks Descendants<Run>() which descends into the
        // wrapper, so the run keeps its r[N] index.
        // v5.9: trackChange=format → <w:rPrChange> inside the run's rPr.
        // Carries author/date/id; the OLD rPr child is left empty (the
        // .doc-side sprmCPropRMark fires without the prior property
        // snapshot, so we just stamp the format-revision marker without
        // a recoverable before-state).
        if (trackChangeKind == "format")
        {
            var rPr = newRun.GetFirstChild<RunProperties>()
                   ?? newRun.PrependChild(new RunProperties());
            var rprChange = new RunPropertiesChange();
            // BUG-DUMP-PPRCHANGE-AUTHOR (run side): w:author is REQUIRED on
            // CT_TrackChange (rPrChange) — same schema rule as pPrChange above.
            // An empty-author source marker (w:author="") must round-trip as an
            // empty attribute, not a dropped one (which fails validation and
            // triggers Word repair-on-open).
            rprChange.Author = trackChangeAuthor ?? "";
            // BUG-R4F-03: RoundtripKind keeps a …Z date in Utc (see above).
            if (!string.IsNullOrEmpty(trackChangeDate)
                && DateTime.TryParse(trackChangeDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var tcfDate))
                rprChange.Date = tcfDate;
            rprChange.Id = !string.IsNullOrEmpty(trackChangeId)
                ? trackChangeId
                : GenerateRevisionId();
            // Schema: w:rPrChange child of w:rPr; ECMA-376 §17.13.5.31.
            // Empty inner rPr is schema-valid (means "no recorded prior
            // property set" — minimal marker form).
            rprChange.AppendChild(new RunProperties());
            rPr.AppendChild(rprChange);
            // BUG-DUMP-R43-8: restore the prior-property snapshot the dump
            // captured (revision.beforeXml) so Word's Reject-Change recovers
            // the original run formatting instead of an empty marker.
            if (properties.TryGetValue("revision.beforeXml", out var rTcBeforeXml)
                && !string.IsNullOrWhiteSpace(rTcBeforeXml))
                ApplyBeforeXmlSnapshot(rprChange, rTcBeforeXml);
        }
        if (trackChangeKind == "ins" || trackChangeKind == "del")
        {
            var parentEl = newRun.Parent;
            if (parentEl != null)
            {
                OpenXmlElement wrapper = trackChangeKind == "ins"
                    ? new InsertedRun()
                    : new DeletedRun();
                if (!string.IsNullOrEmpty(trackChangeAuthor))
                {
                    if (wrapper is InsertedRun insW) insW.Author = trackChangeAuthor;
                    else if (wrapper is DeletedRun delW) delW.Author = trackChangeAuthor;
                }
                // BUG-R4F-03: RoundtripKind keeps a …Z date in Utc (see above).
                if (!string.IsNullOrEmpty(trackChangeDate)
                    && DateTime.TryParse(trackChangeDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var tcDate))
                {
                    if (wrapper is InsertedRun insW2) insW2.Date = tcDate;
                    else if (wrapper is DeletedRun delW2) delW2.Date = tcDate;
                }
                if (!string.IsNullOrEmpty(trackChangeId))
                {
                    if (wrapper is InsertedRun insW3) insW3.Id = trackChangeId;
                    else if (wrapper is DeletedRun delW3) delW3.Id = trackChangeId;
                }
                else
                {
                    // Each ins/del needs a unique w:id. Allocated from the
                    // shared paraId pool (decimal form), guaranteed unique
                    // against all paraId/textId/revision ids in the document.
                    var fallbackId = GenerateRevisionId();
                    if (wrapper is InsertedRun insW4) insW4.Id = fallbackId;
                    else if (wrapper is DeletedRun delW4) delW4.Id = fallbackId;
                }
                // For w:del, the inner Run's <w:t> must become <w:delText>
                // so Word displays the strikethrough content. Convert
                // any Text children to DeletedText.
                if (trackChangeKind == "del")
                {
                    foreach (var t in newRun.Elements<Text>().ToList())
                    {
                        var dt = new DeletedText(t.Text ?? "") { Space = t.Space };
                        t.Parent?.ReplaceChild(dt, t);
                    }
                }
                // BUG-DUMP-DELININS: rebuild the <w:ins><w:del> stack for a run
                // that is both inserted and deleted. The wrapper above is the
                // OUTER ins; insert an INNER del between it and the run so the
                // shape is <w:ins><w:del><w:r><w:delText>. ECMA-376 permits only
                // ins⊃del nesting, so this fires only for revision.type=ins +
                // revision.nested.type=del.
                if (trackChangeKind == "ins" && nestedTcKind == "del")
                {
                    var innerDel = new DeletedRun();
                    if (!string.IsNullOrEmpty(nestedTcAuthor)) innerDel.Author = nestedTcAuthor;
                    if (!string.IsNullOrEmpty(nestedTcDate)
                        && DateTime.TryParse(nestedTcDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var ndDate))
                        innerDel.Date = ndDate;
                    innerDel.Id = !string.IsNullOrEmpty(nestedTcId) ? nestedTcId : GenerateRevisionId();
                    // The deleted run's text must ride in <w:delText>.
                    foreach (var t in newRun.Elements<Text>().ToList())
                    {
                        var dt = new DeletedText(t.Text ?? "") { Space = t.Space };
                        t.Parent?.ReplaceChild(dt, t);
                    }
                    // newRun is still a child of parentEl here — swap in the
                    // outer ins, then nest del then the run: <w:ins><w:del><w:r>.
                    parentEl.ReplaceChild(wrapper, newRun);
                    wrapper.AppendChild(innerDel);
                    innerDel.AppendChild(newRun);
                }
                else
                {
                    parentEl.ReplaceChild(wrapper, newRun);
                    wrapper.AppendChild(newRun);
                }
            }
        }
        // moveFrom / moveTo: low-level OOXML synthesis primitives for
        // dump/replay round-trip. The two sides MUST share the same w:id +
        // w:author + w:date to be recognised as a single move operation by
        // Word — across two independent `add run` calls the CLI cannot infer
        // which pair the caller means, so trackChange.id is REQUIRED here.
        // (For interactive authoring the high-level shape would be a single
        // compound `word move` command that emits both sides atomically.)
        if (trackChangeKind == "movefrom" || trackChangeKind == "moveto")
        {
            if (string.IsNullOrEmpty(trackChangeId))
                throw new InvalidOperationException(
                    $"revision.type={trackChangeKind} requires an explicit revision.id; "
                    + "moveFrom and moveTo must share the same id to be recognised as a "
                    + "pair by Word. Pass --prop revision.id=<n> on both sides.");

            // CONSISTENCY(move-range-markers): wrap via the same
            // WrapRunAsMoveFrom / WrapRunAsMoveTo helpers the `set
            // --prop revision.type=moveFrom` path uses, so the
            // moveFrom/moveTo run is BRACKETED by
            // moveFromRangeStart/End + moveToRangeStart/End carrying
            // Name="Move_{id}". Previously this Add path emitted a bare
            // <w:moveFrom>/<w:moveTo> wrapper with no range markers — on
            // a dump→batch round-trip the four range markers and the
            // shared w:name pairing were dropped, degrading the move to
            // an unpaired moveFrom + moveTo that Word's reviewing pane
            // can't pair (and that Word for Mac may refuse to open with
            // "Word found unreadable content"). The moveFrom and its
            // paired moveTo share one revision.id by design (see the
            // contract above + WordBatchEmitter pairing), so Move_{id}
            // matches across the two halves. Both wrappers keep <w:t>:
            // per ECMA-376 §17.3.3.34 w:delText is only valid inside
            // <w:del>, never inside <w:moveFrom>.
            if (newRun.Parent != null)
            {
                // BUG-R4F-03: RoundtripKind keeps a …Z date in Utc.
                DateTime moveDate = DateTime.UtcNow;
                if (!string.IsNullOrEmpty(trackChangeDate)
                    && DateTime.TryParse(trackChangeDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var mvDate))
                    moveDate = mvDate;
                var moveAuthor = string.IsNullOrEmpty(trackChangeAuthor) ? "OfficeCLI" : trackChangeAuthor!;
                if (trackChangeKind == "movefrom")
                    WrapRunAsMoveFrom(newRun, moveAuthor, moveDate, trackChangeId!);
                else
                    WrapRunAsMoveTo(newRun, moveAuthor, moveDate, trackChangeId!);
                // BUG-DUMP-MOVE-DEL: a run that is BOTH moved AND deleted
                // (<w:moveFrom|moveTo><w:del><w:r>) must keep its inner deletion —
                // otherwise the moved-and-deleted text resurfaces as live, accepted
                // content (meaning change). Insert a <w:del> between the move wrapper
                // and the run and convert <w:t> to <w:delText>, mirroring ins⊃del.
                if (nestedTcKind == "del" && newRun.Parent != null)
                {
                    var moveParent = newRun.Parent;
                    var innerDel = new DeletedRun
                    {
                        Id = !string.IsNullOrEmpty(nestedTcId) ? nestedTcId : GenerateRevisionId()
                    };
                    if (!string.IsNullOrEmpty(nestedTcAuthor)) innerDel.Author = nestedTcAuthor;
                    if (!string.IsNullOrEmpty(nestedTcDate)
                        && DateTime.TryParse(nestedTcDate, null, System.Globalization.DateTimeStyles.RoundtripKind, out var mvNdDate))
                        innerDel.Date = mvNdDate;
                    foreach (var t in newRun.Elements<Text>().ToList())
                        t.Parent?.ReplaceChild(new DeletedText(t.Text ?? "") { Space = t.Space }, t);
                    moveParent.ReplaceChild(innerDel, newRun);
                    innerDel.AppendChild(newRun);
                }
            }
        }

        // Refresh textId since paragraph content changed
        targetPara.TextId = GenerateParaId();

        return resultPath;
    }

    /// <summary>
    /// Append <paramref name="text"/> to <paramref name="run"/>, tokenizing on
    /// '\n' (w:br) and '\t' (w:tab) so the user-visible line breaks and tabs
    /// round-trip through Word instead of being collapsed to a single space.
    /// CRLF/CR are normalized to LF first.
    /// </summary>
    // Expand `{page}` / `{pages}` tokens in user-supplied paragraph text into
    // proper PAGE / NUMPAGES complex-field runs (begin / instrText / separate /
    // result / end). The pre-built `run` is reused for the first literal
    // segment so its rPr stays intact; subsequent literal segments and the
    // field-run sequences clone `rPropsTemplate` so formatting (font/size/
    // color/...) survives the split. Without this Word renders the tokens
    // verbatim instead of substituting page numbers.
    private static readonly System.Text.RegularExpressions.Regex PageFieldTokenRegex =
        new(@"\{(page|pages)\}", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

    internal static void AppendTextWithPageFields(Paragraph para, Run firstRun, RunProperties rPropsTemplate, string text)
    {
        if (string.IsNullOrEmpty(text) || !PageFieldTokenRegex.IsMatch(text))
        {
            AppendTextWithBreaks(firstRun, text);
            para.AppendChild(firstRun);
            return;
        }

        int cursor = 0;
        bool firstRunUsed = false;
        foreach (System.Text.RegularExpressions.Match m in PageFieldTokenRegex.Matches(text))
        {
            if (m.Index > cursor)
            {
                var segment = text.Substring(cursor, m.Index - cursor);
                var segRun = firstRunUsed ? new Run((RunProperties)rPropsTemplate.CloneNode(true)) : firstRun;
                AppendTextWithBreaks(segRun, segment);
                para.AppendChild(segRun);
                firstRunUsed = true;
            }
            var instr = m.Groups[1].Value.Equals("pages", StringComparison.OrdinalIgnoreCase) ? " NUMPAGES " : " PAGE ";
            para.AppendChild(new Run((RunProperties)rPropsTemplate.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.Begin }));
            para.AppendChild(new Run((RunProperties)rPropsTemplate.CloneNode(true), new FieldCode(instr) { Space = SpaceProcessingModeValues.Preserve }));
            para.AppendChild(new Run((RunProperties)rPropsTemplate.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.Separate }));
            para.AppendChild(new Run((RunProperties)rPropsTemplate.CloneNode(true), new Text("1") { Space = SpaceProcessingModeValues.Preserve }));
            para.AppendChild(new Run((RunProperties)rPropsTemplate.CloneNode(true), new FieldChar { FieldCharType = FieldCharValues.End }));
            firstRunUsed = true;
            cursor = m.Index + m.Length;
        }
        if (cursor < text.Length)
        {
            var tailRun = firstRunUsed ? new Run((RunProperties)rPropsTemplate.CloneNode(true)) : firstRun;
            AppendTextWithBreaks(tailRun, text.Substring(cursor));
            para.AppendChild(tailRun);
        }
    }

    internal static void AppendTextWithBreaks(Run run, string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            run.AppendChild(new Text("") { Space = SpaceProcessingModeValues.Preserve });
            return;
        }
        // CONSISTENCY(xml-text-validation): mirror Set's text= path — reject XML 1.0
        // illegal control chars before constructing Text nodes. Without this, the
        // resident process saves a corrupt DOM and surfaces "save failed — data may
        // be lost" only on close, costing the user their edits.
        Core.ParseHelpers.ValidateXmlText(text, "text", allowSoftBreakChar: true);
        // CONSISTENCY(escape-sequences): cross-handler convention — `\n` / `\t`
        // two-char escapes in --prop text= are interpreted as real newline /
        // tab. Mirrors PPTX shape-text and Excel cell-value handling. CRLF/CR
        // collapsed afterwards so all break forms route through <w:br/>.
        // CONSISTENCY(text-escape-boundary): \n / \t resolution at CLI --prop;
        // text arrives with real newlines already, just normalize CR / CRLF.
        // NEWLINE-SEMANTICS-V2: '\v' is the canonical soft-line-break char
        // (<w:br/>), matching Word's own object model (Chr(11)) and the
        // Google Docs API. '\n' still degrades to a soft break HERE because a
        // run physically cannot span paragraphs — block-level surfaces
        // (AddParagraph) consume '\n' as a paragraph boundary before text
        // ever reaches this run-scoped helper.
        var s = text.Replace("\r\n", "\n").Replace("\r", "\n");
        int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c == '\n' || c == '\v' || c == '\t')
            {
                if (i > start)
                    run.AppendChild(new Text(s.Substring(start, i - start)) { Space = SpaceProcessingModeValues.Preserve });
                if (c == '\t') run.AppendChild(new TabChar());
                else run.AppendChild(new Break());
                start = i + 1;
            }
        }
        if (start < s.Length)
            run.AppendChild(new Text(s.Substring(start)) { Space = SpaceProcessingModeValues.Preserve });
        else if (start == 0)
            run.AppendChild(new Text("") { Space = SpaceProcessingModeValues.Preserve });
    }

    // Add a tab stop. Parent must be a Paragraph or a paragraph/table-typed
    // Style; the helper finds or creates the pPr/Tabs container and appends
    // a TabStop. `pos` is required (twips, or any unit accepted by
    // SpacingConverter.ParseWordSpacing). `val` defaults to "left";
    // `leader` is optional. Returns the new tab's path under the
    // conventional /<parent>/tab[N] form — Navigation descends through
    // pPr/tabs (paragraph) or StyleParagraphProperties/tabs (style)
    // transparently for this segment shape.
    private string AddTab(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("pos", out var posStr) || string.IsNullOrWhiteSpace(posStr))
            throw new ArgumentException("tab requires 'pos' property (e.g. --prop pos=9360 or --prop pos=6cm)");

        // Tab positions may be negative (OOXML allows w:pos < 0 to place a tab
        // stop in the negative-indent / hanging region). Cannot reuse
        // SpacingConverter.ParseWordSpacing here because that helper enforces
        // a non-negative guard suitable for paragraph spacing but semantically
        // wrong for tab positions. Parse as signed twips with the same unit
        // suffix vocabulary as ParseWordSpacing (pt / cm / in / bare twips).
        var posTwips = ParseSignedTwips(posStr);

        var tabStop = new TabStop { Position = posTwips };
        if (properties.TryGetValue("val", out var valStr) && !string.IsNullOrEmpty(valStr))
        {
            var tabValNorm = valStr.ToLowerInvariant();
            // Validate before constructing the enum — an invalid string throws
            // ArgumentOutOfRangeException which the outer dispatcher catches and
            // surfaces as a misleading "Invalid index or anchor" error.
            var knownTabVals = new[] { "left", "center", "right", "decimal", "bar", "clear", "num", "start", "end" };
            if (!knownTabVals.Contains(tabValNorm))
                throw new ArgumentException($"Invalid tab val '{valStr}'. Valid: {string.Join(", ", knownTabVals)}.");
            tabStop.Val = new EnumValue<TabStopValues>(new TabStopValues(tabValNorm));
        }
        else
            tabStop.Val = TabStopValues.Left;
        if (properties.TryGetValue("leader", out var leaderStr) && !string.IsNullOrEmpty(leaderStr))
        {
            var leaderNorm = leaderStr.ToLowerInvariant();
            // BUG-DUMP10-06: TabStopLeaderCharValues enum strings are camelCase
            // ("middleDot"), not lowercase. Constructing
            // `new TabStopLeaderCharValues("middledot")` throws
            // ArgumentOutOfRangeException, which the outer dispatcher caught
            // and surfaced as the misleading "Invalid index or anchor" error.
            // Map explicitly to the SDK enum members instead — same pattern as
            // ptab leader resolution in WordHandler.Helpers.cs:858.
            tabStop.Leader = leaderNorm switch
            {
                "none"       => TabStopLeaderCharValues.None,
                "dot"        => TabStopLeaderCharValues.Dot,
                "heavy"      => TabStopLeaderCharValues.Heavy,
                "hyphen"     => TabStopLeaderCharValues.Hyphen,
                "middledot"  => TabStopLeaderCharValues.MiddleDot,
                "underscore" => TabStopLeaderCharValues.Underscore,
                _ => throw new ArgumentException(
                    $"Invalid tab leader '{leaderStr}'. Valid: none, dot, heavy, hyphen, middleDot, underscore."),
            };
        }

        // pPr children have a strict CT_PPr order; <w:tabs> sits early but
        // NOT first — pStyle (and keepNext/numPr/pBdr/…) precede it. Prepending
        // landed tabs before pStyle and produced schema-invalid pPr. Append,
        // then let SchemaOrder hoist it to the SDK-authoritative slot.
        Tabs tabs;
        if (parent is Paragraph para)
        {
            // pPr must come first inside <w:p> per CT_P schema
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            var tabsEl = pProps.GetFirstChild<Tabs>();
            if (tabsEl == null)
            {
                tabsEl = pProps.AppendChild(new Tabs());
                Core.SchemaOrder.Place(pProps, tabsEl);
            }
            tabs = tabsEl;
        }
        else if (parent is Style style)
        {
            // Type guard already enforced in Add.cs (paragraph/table only).
            // EnsureStyleParagraphProperties handles schema-correct insertion
            // before StyleRunProperties.
            var spProps = style.StyleParagraphProperties ?? EnsureStyleParagraphProperties(style);
            var tabsEl = spProps.GetFirstChild<Tabs>();
            if (tabsEl == null)
            {
                tabsEl = spProps.AppendChild(new Tabs());
                Core.SchemaOrder.Place(spProps, tabsEl);
            }
            tabs = tabsEl;
        }
        else
        {
            throw new ArgumentException(
                $"Cannot add 'tab' under {parentPath}: tab stops belong inside a paragraph or a paragraph-typed style.");
        }

        var existing = tabs.Elements<TabStop>().ToList();
        if (index.HasValue && index.Value >= 0 && index.Value < existing.Count)
            tabs.InsertBefore(tabStop, existing[index.Value]);
        else
            tabs.AppendChild(tabStop);

        var newIdx = PathIndex.FromArrayIndex(tabs.Elements<TabStop>().ToList().IndexOf(tabStop));
        return $"{parentPath}/tab[{newIdx}]";
    }

    // Signed twips parser for tab w:pos. Accepts the same unit suffixes as
    // SpacingConverter (pt / cm / in / bare twips) but permits negative values.
    private static int ParseSignedTwips(string value)
    {
        var trimmed = value.Trim();
        const double pointsPerCm = 72.0 / 2.54;
        const double pointsPerInch = 72.0;
        const int twipsPerPoint = 20;

        double points;
        if (trimmed.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            points = ParseSignedNumber(trimmed[..^2]);
        else if (trimmed.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            points = ParseSignedNumber(trimmed[..^2]) * pointsPerCm;
        else if (trimmed.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            points = ParseSignedNumber(trimmed[..^2]) * pointsPerInch;
        else
            // Bare number → twips (Word convention, matches ParseWordSpacing)
            return (int)Math.Round(ParseSignedNumber(trimmed));

        return (int)Math.Round(points * twipsPerPoint);
    }

    private static double ParseSignedNumber(string s)
    {
        var t = s.Trim();
        if (!double.TryParse(t, System.Globalization.CultureInfo.InvariantCulture, out var result)
            || double.IsNaN(result) || double.IsInfinity(result))
            throw new ArgumentException(
                $"Invalid tab 'pos' value '{s}'. Expected a finite number with optional unit (e.g. '-360', '6cm', '0.5in').");
        return result;
    }

    // CONSISTENCY(run-special-content): inline `<w:ptab>` (positional tab,
    // Word 2007+) wrapped in `<w:r>`. Used in headers/footers to anchor
    // left/center/right alignment regions. Mirrors AddBreak's "wrap an
    // inline structure in a Run, insert into paragraph" pattern.
    private string AddPtab(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Validate parent first (more fundamental than property contents) so
        // a misrouted call surfaces the real failure ("must be a paragraph")
        // instead of pushing the user through alignment/leader/relativeTo
        // diagnostics that wouldn't matter at the right path.
        if (parent is not Paragraph para)
            throw new ArgumentException("ptab parent must be a paragraph (got " + parent.GetType().Name + ").");

        if (!(properties.TryGetValue("align", out var alignment) || properties.TryGetValue("alignment", out alignment)) || string.IsNullOrWhiteSpace(alignment))
            throw new ArgumentException("ptab requires 'alignment' property (left, center, or right).");

        var ptab = new PositionalTab { Alignment = ParsePtabAlignment(alignment) };
        // CONSISTENCY(empty-prop-as-default): three optional ptab props use
        // matching IsNullOrWhiteSpace guards so empty-string is uniformly
        // treated as "unset / use default" — previously relativeTo passed
        // "" straight to ParsePtabRelativeTo, raising "Invalid relativeTo
        // ''" while leader silently defaulted, an asymmetry that bit
        // scripted callers building param dicts.
        if ((properties.TryGetValue("relativeTo", out var relTo)
             || properties.TryGetValue("relativeto", out relTo))
            && !string.IsNullOrWhiteSpace(relTo))
            ptab.RelativeTo = ParsePtabRelativeTo(relTo);
        else
            ptab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin;
        if (properties.TryGetValue("leader", out var leader) && !string.IsNullOrWhiteSpace(leader))
            ptab.Leader = ParsePtabLeader(leader);
        else
            ptab.Leader = AbsolutePositionTabLeaderCharValues.None;

        var ptabRun = new Run(ptab);
        // BUG-DUMP-TABRPR: a positional tab paints a leader in the run's font
        // and contributes to line height, so its typography is meaningful
        // (mirrors a plain tab). Apply any run-level props (font / size /
        // szCs / bold / …) onto the ptab run's rPr so dump→batch round-trips
        // them; EnsureRunProperties prepends <w:rPr> ahead of <w:ptab> per
        // schema order. ptab-structural keys are consumed above.
        foreach (var (k, v) in properties)
        {
            var kl = k.ToLowerInvariant();
            if (kl is "align" or "alignment" or "relativeto" or "leader") continue;
            ApplyRunFormatting(EnsureRunProperties(ptabRun), k, v);
        }
        InsertIntoParagraph(para, ptabRun, index);
        // CONSISTENCY(paraid-textid-refresh): paragraph contents changed,
        // so textId must regenerate to mark the paragraph as modified for
        // revision-tracking and diff tooling. Mirrors AddRun's behavior.
        para.TextId = GenerateParaId();
        var runIdx = PathIndex.FromArrayIndex(GetAllRuns(para).IndexOf(ptabRun));
        // CONSISTENCY(para-path-canonical): when parent is itself a
        // paragraph, parentPath already points at it — appending another
        // /p[N] would yield an illegal /p[1]/p[1]/r[N] path. Replace the
        // trailing /p[...] segment with paraId-form so the returned
        // path round-trips through Get unchanged.
        var canonicalParaPath = ReplaceTrailingParaSegment(parentPath, para);
        return $"{canonicalParaPath}/r[{runIdx}]";
    }
}
