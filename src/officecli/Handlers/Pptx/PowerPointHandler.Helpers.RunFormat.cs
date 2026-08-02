// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{

    private static double ParseFontSize(string value) =>
        ParseHelpers.ParseFontSize(value);

    private static bool ParsePptDirectionRtl(string value) => value.ToLowerInvariant() switch
    {
        "rtl" or "righttoleft" or "right-to-left" or "true" or "1" => true,
        "ltr" or "lefttoright" or "left-to-right" or "false" or "0" or "" => false,
        _ => throw new ArgumentException($"Invalid direction value: '{value}'. Valid values: rtl, ltr (also accepts true/false, 1/0, righttoleft/lefttoright, right-to-left/left-to-right; case-insensitive).")
    };

    /// <summary>
    /// Format an EMU value as points for round-trip with bare-number Add/Set input
    /// on PPTX paragraph indent. 12700 EMU = 1pt; output formatted with up to 2
    /// decimals (e.g. "1pt", "0.5pt", "-12pt"). CONSISTENCY(pptx-bare-as-points).
    /// </summary>
    private static string FormatPptIndentPoints(long emu)
    {
        var pt = emu / EmuConverter.EmuPerPointF;
        // 4 decimals: 2-decimal output rounded 216000 EMU → "17.01pt" → 216027
        // EMU on replay, drifting indents by ~2 EMU per round trip. 0.0001pt
        // = 1.27 EMU, so four decimals re-parse to the exact source EMU.
        return pt.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture) + "pt";
    }

    /// <summary>
    /// R65 bt-2: read <c>&lt;a:pPr&gt;/&lt;a:tabLst&gt;/&lt;a:tab pos algn/&gt;…</c> as the
    /// compact compound <c>"36pt:l,72pt:ctr,108pt:r,144pt:dec"</c> so custom tab
    /// stops survive dump → batch replay. Returns null when the paragraph has no
    /// <c>tabLst</c> or the list is empty (PowerPoint inherits master/layout tab
    /// stops in that case; emitting an empty string would replace inheritance
    /// with an explicit zero-tab list on replay). Position is rendered in
    /// points to round-trip with bare-number Add/Set input
    /// (CONSISTENCY(pptx-bare-as-points)). Alignment uses the OOXML literal
    /// values (<c>l</c>/<c>ctr</c>/<c>r</c>/<c>dec</c>) — the same vocabulary
    /// AddParagraph/SetParagraph accepts on input.
    /// </summary>
    private static string? ReadTabsFromPProps(Drawing.ParagraphProperties pProps)
    {
        var tabLst = pProps.GetFirstChild<Drawing.TabStopList>();
        if (tabLst == null) return null;
        var tabs = tabLst.Elements<Drawing.TabStop>().ToList();
        if (tabs.Count == 0) return null;
        var parts = new List<string>(tabs.Count);
        foreach (var tab in tabs)
        {
            if (tab.Position?.HasValue != true) continue;
            var posPt = FormatPptIndentPoints(tab.Position.Value);
            // Default alignment in OOXML schema is `l` when @algn is absent.
            var algn = tab.Alignment?.HasValue == true
                ? (tab.Alignment.InnerText ?? "l")
                : "l";
            parts.Add($"{posPt}:{algn}");
        }
        return parts.Count == 0 ? null : string.Join(",", parts);
    }

    /// <summary>
    /// R65 bt-2: parse the compact compound <c>"pos1:algn1,pos2:algn2,…"</c>
    /// surfaced by <see cref="ReadTabsFromPProps"/> into a fresh
    /// <c>&lt;a:tabLst&gt;</c>. Empty input yields null (caller skips the
    /// child entirely, preserving master/layout inheritance). Whitespace
    /// around delimiters is tolerated; alignment is optional and defaults
    /// to <c>l</c> when omitted (matches the OOXML schema default).
    /// </summary>
    internal static Drawing.TabStopList? ParseTabStopList(string spec)
    {
        if (string.IsNullOrWhiteSpace(spec)) return null;
        var list = new Drawing.TabStopList();
        var entries = spec.Split(',', StringSplitOptions.RemoveEmptyEntries);
        foreach (var raw in entries)
        {
            var entry = raw.Trim();
            if (entry.Length == 0) continue;
            string posPart;
            string algnPart;
            var colon = entry.IndexOf(':');
            if (colon < 0)
            {
                posPart = entry;
                algnPart = "l";
            }
            else
            {
                posPart = entry.Substring(0, colon).Trim();
                algnPart = entry.Substring(colon + 1).Trim();
                if (algnPart.Length == 0) algnPart = "l";
            }
            var posEmu = (int)Math.Round(SpacingConverter.ParsePointsSigned(posPart) * EmuConverter.EmuPerPointF);
            var algnEnum = algnPart.ToLowerInvariant() switch
            {
                "l" or "left" => Drawing.TextTabAlignmentValues.Left,
                "ctr" or "center" or "centre" => Drawing.TextTabAlignmentValues.Center,
                "r" or "right" => Drawing.TextTabAlignmentValues.Right,
                "dec" or "decimal" => Drawing.TextTabAlignmentValues.Decimal,
                _ => throw new ArgumentException($"Invalid tab alignment '{algnPart}' in tabs='{spec}'. Expected l|ctr|r|dec (OOXML a:tab/@algn).")
            };
            list.Append(new Drawing.TabStop { Position = posEmu, Alignment = algnEnum });
        }
        return list.HasChildren ? list : null;
    }

    /// <summary>
    /// Read the canonical `list` value from a paragraph's properties — mirrors
    /// the input alias table consumed by ApplyListStyle (Fill.cs). Returns null
    /// when the paragraph carries no <a:buChar>/<a:buAutoNum>/<a:buNone>.
    /// Used by both the shape-level `list` summary (first paragraph) AND the
    /// per-paragraph emit so dump→replay preserves per-paragraph bullets
    /// instead of collapsing every paragraph after the first to flush-left
    /// plain text. The canonical value can be re-fed to AddParagraph /
    /// Set paragraph via the existing `list` setter.
    /// </summary>
    // Bullet child local-names (the CT_TextParagraphProperties bullet group).
    // The lossy `list` keyword only captures buChar/buNone/buAutoNum and maps to
    // a single token; it drops the bullet font/color/size and the exact char. To
    // round-trip them faithfully, dump emits the whole bullet group verbatim as a
    // `bulletRaw` prop and the apply path splices it back in schema order.
    private static readonly string[] BulletChildLocalNames =
    {
        "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts",
        "buFontTx", "buFont", "buNone", "buAutoNum", "buChar", "buBlip",
    };

    /// <summary>
    /// Capture the full bullet element group (buClr/buFont/buSzPct/buChar/…) from
    /// a paragraph's properties as a concatenated raw-XML string, or null when the
    /// paragraph declares no bullet. Used by the dump readback so Wingdings/colored/
    /// sized bullets survive round-trip instead of degrading to the lossy keyword.
    /// </summary>
    private static string? ReadBulletRawFromPProps(Drawing.ParagraphProperties pProps)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var child in pProps.Elements())
        {
            // Canonicalize attribute order: the SDK preserves lexical attribute
            // order, so a buAutoNum built by ApplyListStyle (type first) and one
            // re-parsed from bulletRaw by ApplyBulletRaw (xmlns first) emitted
            // byte-different bulletRaw for the same bullet, breaking dump
            // idempotency. Same fix as the xlsx OLE anchor slices.
            if (Array.IndexOf(BulletChildLocalNames, child.LocalName) >= 0)
                sb.Append(ExcelHandler.CanonicalizeXmlAttributeOrder(child.OuterXml));
        }
        return sb.Length > 0 ? sb.ToString() : null;
    }

    /// <summary>
    /// R7-10: returns the re-feedable `list` companion value for a paragraph's bullet,
    /// or null when the bullet only round-trips via bulletRaw. A char bullet bound to a
    /// symbol font (buFont, e.g. Wingdings "l") is font-dependent and NOT re-feedable, so
    /// it is suppressed (preserving B5's bulletRaw-only contract for custom symbol bullets);
    /// a plain Unicode char bullet (e.g. "→") or a recognized keyword/auto-number IS emitted.
    /// </summary>
    private static string? ReadCanonicalListKeyword(Drawing.ParagraphProperties pProps)
    {
        // Char bullet with an explicit symbol font is not portable as list=.
        if (pProps.GetFirstChild<Drawing.CharacterBullet>() != null
            && pProps.GetFirstChild<Drawing.BulletFont>() != null)
            return null;
        return ReadListStyleFromPProps(pProps);
    }

    private static string? ReadListStyleFromPProps(Drawing.ParagraphProperties pProps)
    {
        var noBullet = pProps.GetFirstChild<Drawing.NoBullet>();
        if (noBullet != null) return "none";
        var charBullet = pProps.GetFirstChild<Drawing.CharacterBullet>();
        if (charBullet != null)
        {
            var charVal = charBullet.Char?.Value ?? "•";
            return charVal switch
            {
                "•" or "●" or "○" => "bullet",
                "–" or "—" or "-" => "dash",
                "►" or "▶" or "▸" or "➤" => "arrow",
                "✓" or "✔" => "check",
                "★" or "☆" or "⭐" => "star",
                _ => charVal
            };
        }
        var autoBullet = pProps.GetFirstChild<Drawing.AutoNumberedBullet>();
        if (autoBullet?.Type?.HasValue == true)
        {
            var autoVal = autoBullet.Type.InnerText;
            return autoVal switch
            {
                "arabicPeriod" or "arabicParenR" or "arabicPlain" or "arabicParenBoth" => "numbered",
                "romanLcPeriod" or "romanLcParenR" or "romanLcParenBoth" => "romanLc",
                "romanUcPeriod" or "romanUcParenR" or "romanUcParenBoth" => "romanUc",
                "alphaLcPeriod" or "alphaLcParenR" or "alphaLcParenBoth" => "alphaLc",
                "alphaUcPeriod" or "alphaUcParenR" or "alphaUcParenBoth" => "alphaUc",
                _ => autoVal
            };
        }
        return null;
    }

    /// <summary>
    /// Normalize DrawingML alignment abbreviations to human-readable values.
    /// OOXML stores "l", "r", "ctr", "just" etc. — we return "left", "right", "center", "justify".
    /// </summary>
    private static string NormalizeAlignment(string innerText) => innerText switch
    {
        "l" => "left",
        "r" => "right",
        "ctr" => "center",
        "just" => "justify",
        "dist" => "distributed",
        _ => innerText
    };

    /// <summary>
    /// Reorder children of a DrawingML RunProperties / EndParagraphRunProperties /
    /// DefaultRunProperties element into schema-valid order.
    /// Stable within the same order bucket to preserve relative order of existing fills.
    /// Unknown child types are pushed to the end (preserved but last).
    /// </summary>
    internal static void ReorderDrawingRunProperties(OpenXmlCompositeElement rPr)
    {
        if (rPr == null || !rPr.HasChildren) return;

        int OrderOf(OpenXmlElement el)
        {
            var t = el.GetType();
            foreach (var (type, order) in DrawingRunPropChildOrder)
                if (type == t) return order;
            return int.MaxValue;
        }

        var children = rPr.ChildElements.ToList();
        // Check if already sorted — avoid unnecessary reflows
        bool needsReorder = false;
        for (int i = 1; i < children.Count; i++)
        {
            if (OrderOf(children[i]) < OrderOf(children[i - 1]))
            {
                needsReorder = true;
                break;
            }
        }
        if (!needsReorder) return;

        // Stable sort by schema order
        var sorted = children
            .Select((el, idx) => (el, ord: OrderOf(el), idx))
            .OrderBy(t => t.ord)
            .ThenBy(t => t.idx)
            .Select(t => t.el)
            .ToList();

        foreach (var c in children) c.Remove();
        foreach (var c in sorted) rPr.AppendChild(c);
    }

    /// <summary>
    /// Read a GradientFill element and return a string representation (C1-C2[-angle] or radial:C1-C2[-focus]).
    /// </summary>
    /// <summary>
    /// Read a gradient stop color, handling both RgbColorModelHex and SchemeColor.
    /// Without this, scheme-color stops (accent1/dark1/...) read back as "#?" because
    /// FormatHexColor receives the literal "?" placeholder.
    /// </summary>
    private static string ReadGradientStopColor(Drawing.GradientStop gs)
    {
        var rgb = gs.GetFirstChild<Drawing.RgbColorModelHex>();
        if (rgb?.Val?.Value != null)
        {
            // CONSISTENCY(color-input-form): srgbClr with an a:alpha child
            // encodes a per-stop opacity (gradients use it for fade-in/out
            // stops). Wrap the trailing alpha byte into a CSS #RRGGBBAA when
            // the conversion is lossless; otherwise fall back to a raw
            // permille suffix (+alpha=N) so dump→replay preserves the OOXML
            // ST_FixedPercentage value byte-for-byte. Without the lossless
            // check, an alpha of 30000 permille → byte 0x4C → 29803 permille
            // on parse, drifting the stop color every round-trip.
            var rgbCopy = (Drawing.RgbColorModelHex)rgb.CloneNode(true);
            var hex = ParseHelpers.FormatHexColor(rgb.Val.Value);
            var alphaVal = rgb.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
            string baseHex = hex;
            string? alphaSuffix = null;
            if (alphaVal != null && alphaVal < 100000)
            {
                var alphaByte = (int)Math.Round(alphaVal.Value / 100000.0 * 255);
                alphaByte = Math.Clamp(alphaByte, 0, 255);
                // CONSISTENCY(gradient-alpha-permille): round-trip the byte
                // back through the same formula SanitizeColorForOoxml uses
                // on parse ((int)(byte/255.0*100000)) — if it lands on the
                // original permille we can safely use the compact 8-hex
                // form; otherwise the suffix carries the exact integer.
                var roundtrip = (int)(alphaByte / 255.0 * 100000);
                if (roundtrip == alphaVal.Value)
                    baseHex = $"{hex}{alphaByte:X2}";
                else
                    alphaSuffix = $"+alpha={alphaVal.Value}";
            }
            var withTransforms = AppendColorTransforms(baseHex, rgbCopy);
            return alphaSuffix == null ? withTransforms : withTransforms + alphaSuffix;
        }
        var scheme = gs.GetFirstChild<Drawing.SchemeColor>();
        // .Val.Value is an EnumValue<SchemeColorValues> — its ToString() returns the
        // enum object's CLR name ("SchemeColorValues { }"), not the semantic OOXML
        // name. Use InnerText to get "accent1"/"dark1"/... so the emitted gradient
        // string round-trips through BuildGradientFill's color parser.
        // CONSISTENCY(scheme-color-roundtrip): emit canonical long name
        // (dark1/light1/hyperlink/…) so OOXML internal short forms
        // (dk1/lt1/hlink/…) round-trip through Get the same way
        // ReadColorFromFill normalises them.
        if (scheme?.Val?.InnerText != null)
        {
            var name = ParseHelpers.NormalizeSchemeColorName(scheme.Val.InnerText) ?? scheme.Val.InnerText;
            return AppendColorTransforms(name, scheme);
        }
        var sys = gs.GetFirstChild<Drawing.SystemColor>();
        if (sys?.Val?.InnerText != null) return sys.Val.InnerText;
        var preset = gs.GetFirstChild<Drawing.PresetColor>();
        if (preset?.Val?.InnerText != null) return preset.Val.InnerText;
        return "?";
    }

    /// <summary>
    /// bt-B2: detect a captured <a:gradFill> that carries attributes or child
    /// elements ReadGradientString / BuildGradientFill don't model — namely
    /// the `flip` attribute (x / y / xy / none) on the gradFill element and
    /// the <a:tileRect> child (l/t/r/b offsets). Source-authored decks ship
    /// these for fine-tuned fills; emitting only the semantic stops/angle
    /// drops them on round-trip.
    /// </summary>
    internal static bool HasGradientNonSemanticTuning(Drawing.GradientFill gradFill)
    {
        if (gradFill.Flip?.HasValue == true) return true;
        if (gradFill.GetFirstChild<Drawing.TileRectangle>() != null) return true;
        // A path (radial/shape) gradient's <a:fillToRect> focus is only coarsely
        // represented by the semantic focus keyword (tl/tr/bl/br/center). An
        // off-center rect (e.g. t=150000 b=-50000, focus pushed below centre)
        // collapses to "center" and the focus is lost on replay (sample01).
        // Preserve the gradient verbatim whenever the fillToRect is not the
        // centred default (50000 on all four sides).
        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        var ftr = pathGrad?.GetFirstChild<Drawing.FillToRectangle>();
        if (ftr != null)
        {
            long l = ftr.Left?.Value ?? 50000, t = ftr.Top?.Value ?? 50000,
                 r = ftr.Right?.Value ?? 50000, b = ftr.Bottom?.Value ?? 50000;
            if (!(l == 50000 && t == 50000 && r == 50000 && b == 50000)) return true;
        }
        return false;
    }

    internal static string ReadGradientString(Drawing.GradientFill gradFill)
    {
        var stopEls = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stopEls == null || stopEls.Count == 0) return "gradient";

        var stopData = stopEls.Select(gs => (
            color: ReadGradientStopColor(gs),
            pos: gs.Position?.Value
        )).ToList();

        // CONSISTENCY(gradient-pos-permille): preserve the OOXML permille
        // pos verbatim. Previously the emit divided pos by 1000 (truncating
        // to whole-percent) and only emitted when the rounded percent
        // diverged from the even-distribution baseline by more than 1%.
        // That hid sub-percent drift: a 4-stop gradient with stops at
        // [0, 33000, 66000, 100000] (permille) compared each pos against
        // the even baseline [0, 33333, 66667, 100000], saw a 333-unit
        // (0.33%) gap, and decided NOT to emit pos — so BuildGradientFill
        // fell back to the even-distribution computation and replay shifted
        // stop 2 from 33000 to 33333. Reverse the comparison: emit pos
        // when ANY stop deviates from even-distribution (zero tolerance),
        // and use a `p` prefix on the raw permille value so the parser can
        // distinguish it from the legacy percent form (`@33` stays a
        // percent; `@p33000` is raw permille).
        bool hasCustomPos = false;
        int n = stopData.Count;
        for (int i = 0; i < n; i++)
        {
            var expectedPos = n == 1 ? 0 : (int)((long)i * 100000 / (n - 1));
            var actualPos = (int)(stopData[i].pos ?? 0);
            if (actualPos != expectedPos) { hasCustomPos = true; break; }
        }

        var stopStrs = stopData.Select((s, i) =>
            hasCustomPos && s.pos.HasValue
                ? $"{s.color}@p{s.pos.Value}"
                : s.color
        ).ToList();

        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        if (pathGrad != null)
        {
            var fillRect = pathGrad.GetFirstChild<Drawing.FillToRectangle>();
            var focus = "center";
            if (fillRect != null)
            {
                var fl = fillRect.Left?.Value ?? 50000;
                var ft = fillRect.Top?.Value ?? 50000;
                focus = (fl, ft) switch
                {
                    (0, 0) => "tl",
                    ( >= 100000, 0) => "tr",
                    (0, >= 100000) => "bl",
                    ( >= 100000, >= 100000) => "br",
                    _ => "center"
                };
            }
            // R24 — OOXML distinguishes "path" (shape-following) from "radial"
            // via the @path attribute. Background.cs reader already
            // distinguishes; this helper used to flatten everything to
            // "radial:" so dump→replay of a path gradient became a radial.
            var prefix = pathGrad.Path?.Value == Drawing.PathShadeValues.Shape ? "path" : "radial";
            return $"{prefix}:{string.Join("-", stopStrs)}-{focus}";
        }

        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        var deg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 0.0;
        var degStr = deg % 1 == 0 ? $"{(int)deg}" : $"{deg:0.##}";
        return $"linear;{string.Join(";", stopStrs)};{degStr}";
    }

    /// <summary>
    /// Apply run-level formatting to a PPT run's RunProperties.
    /// </summary>
    private static void ApplyPptRunFormatting(Drawing.Run run, string key, string value, Shape? shape = null)
    {
        var rPr = run.RunProperties ?? run.PrependChild(new Drawing.RunProperties());
        switch (key.ToLowerInvariant())
        {
            case "bold":
                rPr.Bold = IsTruthy(value);
                break;
            case "italic":
                rPr.Italic = IsTruthy(value);
                break;
            case "size":
                rPr.FontSize = (int)Math.Round(ParseFontSize(value) * 100, MidpointRounding.AwayFromZero);
                break;
            case "color":
                rPr.RemoveAllChildren<Drawing.SolidFill>();
                rPr.PrependChild(BuildSolidFill(value));
                break;
            case "font":
                // Bare 'font' targets all common scripts (Latin + EastAsian).
                // Use 'font.latin' / 'font.ea' / 'font.cs' for per-script control
                // (e.g. Japanese / Korean / Arabic documents).
                rPr.RemoveAllChildren<Drawing.LatinFont>();
                rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                ReorderDrawingRunProperties(rPr);
                break;
            case "font.latin":
                rPr.RemoveAllChildren<Drawing.LatinFont>();
                rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                ReorderDrawingRunProperties(rPr);
                break;
            case "font.ea" or "font.eastasia" or "font.eastasian":
                rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                ReorderDrawingRunProperties(rPr);
                break;
            case "font.cs" or "font.complexscript" or "font.complex":
                rPr.RemoveAllChildren<Drawing.ComplexScriptFont>();
                rPr.AppendChild(new Drawing.ComplexScriptFont { Typeface = value });
                ReorderDrawingRunProperties(rPr);
                break;
            case "underline":
                var ulVal = value.ToLowerInvariant() switch
                {
                    "true" or "single" => Drawing.TextUnderlineValues.Single,
                    "double" => Drawing.TextUnderlineValues.Double,
                    "heavy" => Drawing.TextUnderlineValues.Heavy,
                    "false" or "none" => Drawing.TextUnderlineValues.None,
                    _ => new Drawing.TextUnderlineValues(value)
                };
                rPr.Underline = ulVal;
                break;
            case "strikethrough" or "strike":
                var stVal = value.ToLowerInvariant() switch
                {
                    "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                    "double" => Drawing.TextStrikeValues.DoubleStrike,
                    "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                    _ => new Drawing.TextStrikeValues(value)
                };
                rPr.Strike = stVal;
                break;
            case "superscript":
                rPr.Baseline = IsTruthy(value) ? 30000 : 0;
                break;
            case "subscript":
                rPr.Baseline = IsTruthy(value) ? -25000 : 0;
                break;
            case "charspacing" or "spacing" or "letterspacing":
                var csPt = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? ParseHelpers.SafeParseDouble(value[..^2], "charspacing")
                    : ParseHelpers.SafeParseDouble(value, "charspacing");
                rPr.Spacing = (int)Math.Round(csPt * 100, MidpointRounding.AwayFromZero);
                break;
            case "highlight":
                rPr.RemoveAllChildren<Drawing.Highlight>();
                if (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
                {
                    var hl = new Drawing.Highlight();
                    hl.AppendChild(BuildSolidFillColor(value));
                    rPr.AppendChild(hl);
                    // CONSISTENCY(highlight): pin the CT_TextCharacterProperties
                    // slot (after effectLst, before uLn/uFill/latin) — same as
                    // the Set path in ShapeProperties.cs; PowerPoint silently
                    // drops out-of-order rPr children.
                    ReorderDrawingRunProperties(rPr);
                }
                break;
        }
    }

    // CT_TextParagraphProperties child schema rank (OOXML DrawingML):
    //   lnSpc, spcBef, spcAft, buClr*, buSzPct/Pts/Tx, buFontTx/buFont,
    //   buNone/buAutoNum/buChar/buBlip, tabLst, defRPr, extLst
    // PowerPoint silently drops out-of-order children. Any code that injects
    // a child into <a:pPr> after the element may already contain higher-rank
    // siblings (typical when the user calls Set repeatedly in reverse order)
    // must route through InsertPPrChild so the schema position is honoured.
    // CONSISTENCY(schema-order-pptx): mirrors the spPr fix pattern proven by
    // PptxSpPrSchemaOrderTests / PptxSchemaOrderR51Tests.
    private static readonly string[] PPrChildSchemaOrder =
    {
        "lnSpc", "spcBef", "spcAft",
        "buClr", "buClrTx",
        "buSzPct", "buSzPts", "buSzTx",
        "buFont", "buFontTx",
        "buNone", "buAutoNum", "buChar", "buBlip",
        "tabLst", "defRPr", "extLst",
    };

    private static int PPrChildRank(OpenXmlElement el)
    {
        var idx = Array.IndexOf(PPrChildSchemaOrder, el.LocalName);
        return idx < 0 ? int.MaxValue : idx;
    }

    /// <summary>
    /// Insert <paramref name="child"/> into a <c>&lt;a:pPr&gt;</c> at the
    /// schema-required position so the resulting XML validates regardless of
    /// the order in which properties were set. Caller is responsible for
    /// removing any pre-existing same-typed child first.
    /// </summary>
    // Apply the simple-valued CT_TextParagraphProperties attributes
    // (line-break / punctuation / font-alignment / default-tab-size). These are
    // plain pPr attributes, not child elements, so they're set directly on the
    // ParagraphProperties object (no schema-order insertion needed). Returns the
    // set of keys it consumed so callers can skip them in their own dispatch.
    // Shared by AddParagraph and SetParagraphOnShape for symmetric round-trip of
    // CJK line-break controls (eaLnBrk etc.).
    private static bool ApplyParagraphBreakProp(Drawing.ParagraphProperties pProps, string key, string value)
    {
        switch (key.ToLowerInvariant())
        {
            case "ealnbrk" or "ealinebreak":
                pProps.EastAsianLineBreak = IsTruthy(value);
                return true;
            case "latinlnbrk" or "latinlinebreak":
                pProps.LatinLineBreak = IsTruthy(value);
                return true;
            case "fontalgn" or "fontalignment":
                pProps.FontAlignment = value.Trim().ToLowerInvariant() switch
                {
                    "auto" => Drawing.TextFontAlignmentValues.Automatic,
                    "t" or "top" => Drawing.TextFontAlignmentValues.Top,
                    "ctr" or "center" => Drawing.TextFontAlignmentValues.Center,
                    "base" or "baseline" => Drawing.TextFontAlignmentValues.Baseline,
                    "b" or "bottom" => Drawing.TextFontAlignmentValues.Bottom,
                    _ => throw new ArgumentException($"Invalid fontAlgn value: '{value}'. Valid: auto, t, ctr, base, b.")
                };
                return true;
            case "deftabsz" or "defaulttabsize":
                pProps.DefaultTabSize = (int)Core.EmuConverter.ParseEmu(value);
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Apply a verbatim bullet-group XML string (buClr/buFont/buSzPct/buChar/…)
    /// captured by ReadBulletRawFromPProps. Removes any existing bullet children
    /// first, then parses the concatenated fragment and inserts each child in
    /// CT_TextParagraphProperties schema order (InsertPPrChild). This preserves
    /// colored/sized/Wingdings bullets that the lossy `list` keyword drops.
    /// </summary>
    private static void ApplyBulletRaw(Drawing.ParagraphProperties pProps, string rawXml)
    {
        // Validate and parse FIRST — every throw must happen before any
        // mutation, or the invalid_value error leaves a pPr stripped of its
        // bullet group behind it (the single-command path autosaves the dirty
        // DOM on Dispose). Same order contract as ApplyListStyle.
        List<OpenXmlElement>? newChildren = null;
        if (!string.IsNullOrWhiteSpace(rawXml))
        {
            // Reject non-XML input up front: a free-form string ("buFont=Wingdings")
            // survived the wrap-parse as TEXT CONTENT of <a:pPr> — an element with
            // no text model — producing a file schema validation passes but real
            // PowerPoint refuses to open (0x80070570).
            if (!rawXml.TrimStart().StartsWith("<"))
                throw new ArgumentException(
                    $"Invalid 'bulletRaw' value: '{rawXml}'. Expected verbatim bullet-group XML " +
                    "(e.g. <a:buChar char=\"•\"/> or <a:buAutoNum type=\"arabicPeriod\"/>) " +
                    "as emitted by Get; use list= for keyword bullets.");
            // Wrap in a throwaway pPr that declares the a: namespace so each child
            // fragment parses with its prefix bound; then lift the parsed children.
            const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var wrapped = $"<a:pPr xmlns:a=\"{aNs}\">{rawXml}</a:pPr>";
            // Materialize the children INSIDE the try: the SDK parses the
            // outer-XML constructor lazily, so a malformed fragment throws
            // XmlException only when ChildElements is first enumerated —
            // outside a constructor-only try it escaped the ArgumentException
            // wrap and surfaced as internal_error with a misleading
            // "corrupted OOXML part" message.
            try
            {
                newChildren = new Drawing.ParagraphProperties(wrapped).ChildElements.ToList();
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid 'bulletRaw' XML: '{rawXml}' ({ex.Message}).", ex);
            }
            foreach (var child in newChildren)
            {
                if (child is OpenXmlUnknownElement || child is OpenXmlMiscNode)
                    throw new ArgumentException(
                        $"Invalid 'bulletRaw' XML: unrecognized fragment '{child.OuterXml}'. " +
                        "Expected a:bu* bullet-group elements.");
            }
        }

        // Strip the prior bullet group so a re-apply is idempotent.
        pProps.RemoveAllChildren<Drawing.BulletColorText>();
        pProps.RemoveAllChildren<Drawing.BulletColor>();
        pProps.RemoveAllChildren<Drawing.BulletSizeText>();
        pProps.RemoveAllChildren<Drawing.BulletSizePercentage>();
        pProps.RemoveAllChildren<Drawing.BulletSizePoints>();
        pProps.RemoveAllChildren<Drawing.BulletFontText>();
        pProps.RemoveAllChildren<Drawing.BulletFont>();
        pProps.RemoveAllChildren<Drawing.NoBullet>();
        pProps.RemoveAllChildren<Drawing.AutoNumberedBullet>();
        pProps.RemoveAllChildren<Drawing.CharacterBullet>();
        pProps.RemoveAllChildren<Drawing.PictureBullet>();

        if (newChildren == null) return;
        foreach (var child in newChildren)
        {
            child.Remove();
            InsertPPrChild(pProps, child);
        }
    }

    // Round-trip a paragraph's <a:pPr><a:defRPr> verbatim. The defRPr is the
    // paragraph-level default run property: every run WITHOUT its own rPr (or
    // whose rPr omits a slot) inherits size/bold/font/color from here, BEFORE
    // falling back to the layout/master bodyStyle cascade. The granular
    // paragraph keys (align/lineSpacing/…) never captured it, so a paragraph
    // whose runs are bare <a:r> rendered at the master body size/weight instead
    // of the authored defRPr (e.g. a 40pt-bold-Helvetica body collapsing to the
    // master's 52pt-regular). Verbatim mirrors bulletRaw / lstStyleRaw.
    private static void ApplyDefRPrRaw(Drawing.ParagraphProperties pProps, string rawXml)
    {
        pProps.RemoveAllChildren<Drawing.DefaultRunProperties>();
        if (string.IsNullOrWhiteSpace(rawXml)) return;
        // Mirror ApplyBulletRaw: reject non-XML input instead of silently
        // no-opping (silent-accept hides the caller's mistake).
        if (!rawXml.TrimStart().StartsWith("<"))
            throw new ArgumentException(
                $"Invalid 'defRPrRaw' value: '{rawXml}'. Expected verbatim <a:defRPr .../> XML as emitted by Get.");
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var wrapped = $"<a:pPr xmlns:a=\"{aNs}\">{rawXml}</a:pPr>";
        Drawing.ParagraphProperties parsed;
        try { parsed = new Drawing.ParagraphProperties(wrapped); }
        catch (Exception ex)
        {
            throw new ArgumentException($"Invalid 'defRPrRaw' XML: '{rawXml}' ({ex.Message}).", ex);
        }
        var defRPr = parsed.GetFirstChild<Drawing.DefaultRunProperties>();
        if (defRPr == null)
            throw new ArgumentException(
                $"Invalid 'defRPrRaw' value: '{rawXml}'. No <a:defRPr> element found in the fragment.");
        defRPr.Remove();
        InsertPPrChild(pProps, defRPr);
    }

    internal static void InsertPPrChild(Drawing.ParagraphProperties pProps, OpenXmlElement child)
    {
        var newRank = PPrChildRank(child);
        // Find the first existing child whose rank is strictly greater — the
        // new element must precede it. Same idiom as spPr/PresetGeometry fix.
        OpenXmlElement? insertBefore = null;
        foreach (var existing in pProps.ChildElements)
        {
            if (PPrChildRank(existing) > newRank)
            {
                insertBefore = existing;
                break;
            }
        }
        if (insertBefore != null)
            pProps.InsertBefore(child, insertBefore);
        else
            pProps.AppendChild(child);
    }
}
