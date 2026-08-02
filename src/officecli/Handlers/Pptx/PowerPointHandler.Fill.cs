// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private static void InsertFillElement(ShapeProperties spPr, OpenXmlElement fillElement)
    {
        // Schema order: xfrm → (prstGeom | custGeom) → fill → ln → effectLst.
        // CT_ShapeProperties' geometry is a choice group: a shape carries EITHER
        // a:prstGeom OR a:custGeom, never both. Anchoring the fill only after
        // prstGeom placed the fill BEFORE a custGeom (xfrm → fill → custGeom),
        // which is out of schema order — PowerPoint refuses the file. Anchor
        // after whichever geometry element is present.
        var geom = (OpenXmlElement?)spPr.GetFirstChild<Drawing.PresetGeometry>()
                   ?? spPr.GetFirstChild<Drawing.CustomGeometry>();
        if (geom != null)
            spPr.InsertAfter(fillElement, geom);
        else
        {
            var xfrm = spPr.Transform2D;
            if (xfrm != null)
                spPr.InsertAfter(fillElement, xfrm);
            else
                spPr.PrependChild(fillElement);
        }
    }

    // ==================== Color Helpers ====================

    // Color/fill builders moved to Core/DrawingColorBuilder so ExcelHandler's
    // drawing-layer shapes can reuse the same scheme-color resolution.
    private static OpenXmlElement BuildColorElement(string value)
        => DrawingColorBuilder.BuildColorElement(value);

    private static Drawing.SolidFill BuildSolidFill(string colorValue)
        => DrawingColorBuilder.BuildSolidFill(colorValue);

    /// <summary>
    /// Build a <a:duotone> blip recolor element from a "c1,c2" spec.
    /// Each color may be hex (#RRGGBB / RRGGBB) or a scheme color name
    /// (accent1, dark1, …); BuildColorElement handles both. Throws on
    /// any other shape — duotone requires exactly two stops per OOXML.
    /// </summary>
    internal static Drawing.Duotone BuildDuotoneFromSpec(string spec)
    {
        var parts = spec.Split(',', StringSplitOptions.TrimEntries);
        if (parts.Length != 2 || string.IsNullOrEmpty(parts[0]) || string.IsNullOrEmpty(parts[1]))
            throw new ArgumentException($"Invalid 'duotone' value: '{spec}'. Expected 'color1,color2' (hex or scheme color).");
        var duo = new Drawing.Duotone();
        duo.AppendChild(BuildColorElement(parts[0]));
        duo.AppendChild(BuildColorElement(parts[1]));
        return duo;
    }

    private static Drawing.SchemeColorValues? TryParseSchemeColor(string value)
        => DrawingColorBuilder.TryParseSchemeColor(value);

    /// <summary>
    /// Read a color value from a SolidFill element, returning either hex RGB or scheme color name.
    /// </summary>
    internal static string? ReadColorFromFill(Drawing.SolidFill? solidFill)
    {
        if (solidFill == null) return null;
        var rgbEl = solidFill.GetFirstChild<Drawing.RgbColorModelHex>();
        if (rgbEl?.Val?.Value != null) return AppendColorTransforms(FormatHexWithAlpha(rgbEl), rgbEl);
        var schemeEl = solidFill.GetFirstChild<Drawing.SchemeColor>();
        if (schemeEl != null)
        {
            // CONSISTENCY(scheme-color-unknown): when the SDK's EnumValue can't
            // parse the schemeClr@val (custom themes with dk3/lt3/accent7+,
            // future OOXML additions) .Val.HasValue is false and InnerText is
            // empty. Fall back to the raw XML attribute so the color survives
            // round-trip instead of silently disappearing.
            var schemeVal = schemeEl.Val;
            string? raw = (schemeVal?.HasValue == true && !string.IsNullOrEmpty(schemeVal.InnerText))
                ? schemeVal.InnerText
                : schemeEl.GetAttribute("val", "").Value;
            if (!string.IsNullOrEmpty(raw))
            {
                var name = ParseHelpers.NormalizeSchemeColorName(raw) ?? raw;
                return AppendColorTransforms(name, schemeEl);
            }
        }
        return ReadSysOrPresetColor(solidFill);
    }

    /// <summary>
    /// Read an a:sysClr (system color) or a:prstClr (preset color) child from a
    /// color parent, returning a canonical hex (with any +lumMod/+shade/… transform
    /// suffix) or null when neither is present. sysClr resolves to its lastClr —
    /// the concrete RGB the host app last rendered, which is exactly what we want
    /// the round-trip to reproduce; prstClr resolves to its named-color hex when
    /// known, else the raw preset name (Add/Set resolves both forms).
    ///
    /// Before this, both element types fell through to a bare `return null`, so a
    /// shape filled with `sysClr "window"` (white) or `prstClr "black"` lost its
    /// fill entirely on Get → dump → rebuild and rendered as an INHERIT/no-fill
    /// transparent box.
    /// </summary>
    internal static string? ReadSysOrPresetColor(OpenXmlElement? parent)
    {
        if (parent == null) return null;
        var sysEl = parent.GetFirstChild<Drawing.SystemColor>();
        if (sysEl != null)
        {
            var last = sysEl.LastColor?.Value;
            string? hex = !string.IsNullOrEmpty(last)
                ? last
                : MapSystemColorFallback(sysEl.Val?.InnerText ?? sysEl.GetAttribute("val", "").Value);
            if (!string.IsNullOrEmpty(hex))
                return AppendColorTransforms(hex!.ToUpperInvariant(), sysEl);
        }
        var prstEl = parent.GetFirstChild<Drawing.PresetColor>();
        if (prstEl != null)
        {
            var name = prstEl.Val?.InnerText;
            if (string.IsNullOrEmpty(name)) name = prstEl.GetAttribute("val", "").Value;
            if (!string.IsNullOrEmpty(name))
                return AppendColorTransforms(ParseHelpers.TryGetNamedColorHex(name) ?? name, prstEl);
        }
        return null;
    }

    // sysClr without a lastClr attribute is rare (PowerPoint always writes one),
    // but fall back to the two system colors that actually appear in documents so
    // the fill never silently vanishes. Other ST_SystemColorVal values are
    // chrome-only and don't occur as shape fills.
    private static string? MapSystemColorFallback(string? val) => val switch
    {
        "window" => "FFFFFF",
        "windowText" => "000000",
        _ => null,
    };

    /// <summary>
    /// Read a color value from an a:highlight element, returning either hex RGB
    /// or scheme color name. CONSISTENCY(highlight): the highlight's color child
    /// has the same shape as a solidFill's (srgbClr / schemeClr), so wrap it in
    /// a throwaway SolidFill to reuse ReadColorFromFill — same trick as the
    /// HtmlPreview.Text.cs renderer.
    /// </summary>
    internal static string? ReadColorFromHighlight(Drawing.Highlight? highlight)
    {
        var colorChild = highlight?.GetFirstChild<Drawing.RgbColorModelHex>()
            ?? (OpenXmlElement?)highlight?.GetFirstChild<Drawing.SchemeColor>();
        if (colorChild == null) return null;
        return ReadColorFromFill(new Drawing.SolidFill(colorChild.CloneNode(true)));
    }

    // R8-4: encode a:lumMod / a:lumOff / a:shade / a:tint / a:satMod / a:satOff /
    // a:hueMod / a:hueOff color transforms as a chained "+name<intPercent>"
    // suffix on the canonical color string so they survive Get → Add/Set
    // round-trip. Pre-R8 these children were silently stripped: a slide's
    // accent1 with lumMod=50000 came back as bare "accent1", and a re-applied
    // round-trip lost the tint.
    private static readonly string[] ColorTransformLocalNames =
        { "lumMod", "lumOff", "shade", "tint", "satMod", "satOff", "hueMod", "hueOff", "alpha" };

    internal static string AppendColorTransforms(string baseColor, OpenXmlElement colorEl)
    {
        // a:srgbClr encodes alpha into the trailing AA byte of FormatHexWithAlpha
        // (RRGGBBAA), so the alpha child is already represented in the base hex.
        // a:schemeClr has no hex form, so its alpha child has nowhere else to
        // live and must be emitted as a "+alphaN" transform suffix to survive
        // round-trip — without this, accent1@alpha50000 came back as bare
        // "accent1" and re-applying the value lost the transparency.
        bool isRgb = colorEl is Drawing.RgbColorModelHex;
        var sb = new System.Text.StringBuilder(baseColor);
        foreach (var child in colorEl.Elements())
        {
            var ln = child.LocalName;
            if (Array.IndexOf(ColorTransformLocalNames, ln) < 0) continue;
            if (ln == "alpha" && isRgb) continue; // alpha already encoded into RRGGBBAA hex form
            var v = child.GetAttribute("val", "").Value;
            if (string.IsNullOrEmpty(v)) continue;
            // Convert OOXML ST_PositivePercentage (0..100000) → human percent.
            if (int.TryParse(v, out var n))
                sb.Append('+').Append(ln).Append(n / 1000);
            else
                sb.Append('+').Append(ln).Append(v);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Read a color value from any element that may contain RgbColorModelHex or SchemeColor.
    /// </summary>
    internal static string? ReadColorFromElement(OpenXmlElement? parent)
    {
        if (parent == null) return null;
        var rgbEl = parent.GetFirstChild<Drawing.RgbColorModelHex>();
        if (rgbEl?.Val?.Value != null) return AppendColorTransforms(FormatHexWithAlpha(rgbEl), rgbEl);
        var schemeEl = parent.GetFirstChild<Drawing.SchemeColor>();
        // CONSISTENCY(scheme-color-roundtrip): emit canonical long names
        // (dark1/light1/hyperlink/…) so OOXML internal short forms
        // (dk1/lt1/hlink/…) round-trip through Get the same way
        // ReadColorFromFill normalises them. Without this, shadow/glow/
        // gradient-stop schemeClr readback surfaced raw InnerText
        // ("dk1"/"hlink"/…), which Add/Set accepts but Get clients
        // following the documented vocabulary wouldn't recognise.
        if (schemeEl != null)
        {
            // CONSISTENCY(scheme-color-unknown): mirror ReadColorFromFill —
            // fall back to the raw @val attribute when EnumValue can't parse it
            // (custom themes, future OOXML additions).
            var schemeVal = schemeEl.Val;
            string? raw = (schemeVal?.HasValue == true && !string.IsNullOrEmpty(schemeVal.InnerText))
                ? schemeVal.InnerText
                : schemeEl.GetAttribute("val", "").Value;
            if (!string.IsNullOrEmpty(raw))
            {
                var name = ParseHelpers.NormalizeSchemeColorName(raw) ?? raw;
                return AppendColorTransforms(name, schemeEl);
            }
        }
        return ReadSysOrPresetColor(parent);
    }

    /// <summary>
    /// Format srgbClr hex, prefixing an AA byte when an a:alpha child is present and non-opaque.
    /// Alpha units are 0..100000 (100000 = opaque, matches OOXML ST_PositiveFixedPercentage).
    /// </summary>
    private static string FormatHexWithAlpha(Drawing.RgbColorModelHex rgbEl)
    {
        var hex = ParseHelpers.FormatHexColor(rgbEl.Val!.Value!);
        var alphaVal = rgbEl.GetFirstChild<Drawing.Alpha>()?.Val?.Value;
        if (alphaVal == null || alphaVal >= 100000) return hex;
        var alphaByte = (int)Math.Round(alphaVal.Value / 100000.0 * 255);
        alphaByte = Math.Clamp(alphaByte, 0, 255);
        // CONSISTENCY(color-input-form): emit CSS #RRGGBBAA so re-feeding the
        // value into Add/Set round-trips correctly (NormalizeArgbColor /
        // SanitizeColorForOoxml treat #-prefixed 8-hex as RRGGBBAA).
        return hex.StartsWith('#')
            ? $"{hex}{alphaByte:X2}"
            : $"{hex}{alphaByte:X2}";
    }

    private static void ApplyShapeFill(ShapeProperties spPr, string value)
    {
        // CONSISTENCY(fill-gradient-shorthand): accept gradient shorthand
        // ("C1-C2[-angle]", "radial:C1-C2", "path:C1-C2", and "LINEAR;C1;C2;angle")
        // directly on fill= — table cells and slide backgrounds already auto-detect
        // the same shorthand, so shape fill matches that input contract instead of
        // forcing callers to switch to the parallel gradient= key.
        var normalized = NormalizeGradientValue(value);
        OpenXmlElement newFill;
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
            newFill = new Drawing.NoFill();
        else if (normalized.StartsWith("radial:", StringComparison.OrdinalIgnoreCase)
              || normalized.StartsWith("path:", StringComparison.OrdinalIgnoreCase)
              || IsGradientColorString(normalized))
            newFill = BuildGradientFill(normalized);
        else
        {
            // CONSISTENCY(scheme-fill-transform-preserve): re-setting the same
            // scheme color (with no user-supplied +lumMod/+shade suffix) on a
            // shape that already carries lumMod/lumOff/shade/tint/satMod/hueMod
            // transforms must NOT strip them — the user expects a no-op when
            // they re-apply the same theme slot. R49 (756a9a13) only covered
            // mutations of an UNRELATED shape; this handles the same-shape
            // case. Skip when the new value carries its own transform chain
            // ("accent1+lumMod=75000" or "accent1:lumMod=75000") — that's an
            // explicit overwrite.
            if (!value.Contains('+') && !value.Contains(':'))
            {
                var existingSolid = spPr.GetFirstChild<Drawing.SolidFill>();
                var existingScheme = existingSolid?.GetFirstChild<Drawing.SchemeColor>();
                var newScheme = TryParseSchemeColor(value);
                if (existingScheme != null && newScheme.HasValue
                    && existingScheme.Val?.Value == newScheme.Value
                    && existingScheme.Elements().Any(c =>
                        Array.IndexOf(ColorTransformLocalNames, c.LocalName) >= 0))
                {
                    return; // preserve transforms — no-op
                }
            }
            newFill = BuildSolidFill(value);
        }

        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();
        spPr.RemoveAllChildren<Drawing.BlipFill>();

        InsertFillElement(spPr, newFill);
    }

    /// <summary>
    /// Apply gradient fill to ShapeProperties.
    /// Linear:  "color1-color2[-angle]"       e.g. "FF0000-0000FF", "FF0000-0000FF-90"
    /// Radial:  "radial:color1-color2"         e.g. "radial:4B0082-1E90FF"
    /// Radial with focus: "radial:color1-color2-tl" (tl/tr/bl/br/center)
    /// </summary>
    private static void ApplyGradientFill(ShapeProperties spPr, string value)
    {
        // Normalize alternative format: "LINEAR;C1;C2;angle" → "C1-C2-angle"
        value = NormalizeGradientValue(value);
        // Build new fill BEFORE removing old one (atomic: no data loss on invalid color)
        var newFill = BuildGradientFill(value);
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();
        spPr.RemoveAllChildren<Drawing.BlipFill>();
        InsertFillElement(spPr, newFill);
    }

    /// <summary>
    /// bt-7 dump→replay raw passthrough for gradient. Value is the captured
    /// <a:gradFill ...> verbatim including flip= / rotWithShape= attrs and
    /// any <a:tileRect/> child — attributes BuildGradientFill never re-emits.
    /// Mirrors the Set.gradientRaw branch so AddShape can consume the same
    /// key inline at create time rather than relying on a follow-up Set.
    /// Returns true on success, false on parse failure (caller surfaces as
    /// unsupported).
    /// </summary>
    internal static bool ApplyGradientRaw(ShapeProperties spPr, string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        try
        {
            var raw = value.Contains("xmlns:a=")
                ? value
                : value.Replace("<a:gradFill",
                    "<a:gradFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
            using var sr = new System.IO.StringReader(raw);
            using var xr = System.Xml.XmlReader.Create(sr);
            xr.MoveToContent();
            var grad = new Drawing.GradientFill();
            if (xr.HasAttributes)
            {
                while (xr.MoveToNextAttribute())
                {
                    if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                    grad.SetAttribute(new OpenXmlAttribute(
                        xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
                }
                xr.MoveToElement();
            }
            var gt = raw.IndexOf('>');
            var lt = raw.LastIndexOf('<');
            if (gt > 0 && lt > gt)
                grad.InnerXml = raw[(gt + 1)..lt];
            spPr.RemoveAllChildren<Drawing.SolidFill>();
            spPr.RemoveAllChildren<Drawing.NoFill>();
            spPr.RemoveAllChildren<Drawing.GradientFill>();
            spPr.RemoveAllChildren<Drawing.PatternFill>();
            spPr.RemoveAllChildren<Drawing.BlipFill>();
            InsertFillElement(spPr, grad);
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Apply pattern fill to ShapeProperties.
    /// Format: "<preset>" or "<preset>:<fgColor>" or "<preset>:<fgColor>:<bgColor>"
    ///   preset: e.g. pct25, ltHorz, cross, weave, zigZag (Drawing.PresetPatternValues)
    ///   fgColor / bgColor: lenient hex/named/scheme color (defaults: fg=000000, bg=FFFFFF)
    /// Examples: "pct25", "ltHorz:FF0000", "cross:red:white"
    /// </summary>
    private static void ApplyPatternFill(ShapeProperties spPr, string value)
    {
        // Build new fill BEFORE removing old one (atomic: no data loss on invalid input)
        var newFill = BuildPatternFill(value);
        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();
        spPr.RemoveAllChildren<Drawing.BlipFill>();
        InsertFillElement(spPr, newFill);
    }

    private static Drawing.PatternFill BuildPatternFill(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("pattern value cannot be empty.");

        var parts = value.Split(':');
        var presetName = parts[0].Trim();

        // Inherited pattern fill: a bare `<a:pattFill/>` (no preset, no colors)
        // means "pattern fill, inherit preset + colors from the style/theme
        // fillRef". NodeBuilder serializes this as "pattern=:" on dump; replay
        // round-trips it here instead of erroring on an empty preset. Any
        // explicitly-supplied fg/bg is still honored ("::FFFFFF" etc.).
        if (string.IsNullOrEmpty(presetName))
        {
            var bare = new Drawing.PatternFill();
            if (parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[1]))
            {
                var fgClrInherit = new Drawing.ForegroundColor();
                fgClrInherit.Append(BuildColorElement(parts[1].Trim()));
                bare.Append(fgClrInherit);
            }
            if (parts.Length > 2 && !string.IsNullOrWhiteSpace(parts[2]))
            {
                var bgClrInherit = new Drawing.BackgroundColor();
                bgClrInherit.Append(BuildColorElement(parts[2].Trim()));
                bare.Append(bgClrInherit);
            }
            return bare;
        }

        var fg = parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[1]) ? parts[1].Trim() : "000000";
        var bg = parts.Length > 2 && !string.IsNullOrWhiteSpace(parts[2]) ? parts[2].Trim() : "FFFFFF";

        var patternFill = new Drawing.PatternFill { Preset = ParsePresetPattern(presetName) };
        // Schema order: fgClr → bgClr
        var fgClr = new Drawing.ForegroundColor();
        fgClr.Append(BuildColorElement(fg));
        patternFill.Append(fgClr);
        var bgClr = new Drawing.BackgroundColor();
        bgClr.Append(BuildColorElement(bg));
        patternFill.Append(bgClr);
        return patternFill;
    }

    private static Drawing.PresetPatternValues ParsePresetPattern(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "pct5" => Drawing.PresetPatternValues.Percent5,
            "pct10" => Drawing.PresetPatternValues.Percent10,
            "pct20" => Drawing.PresetPatternValues.Percent20,
            "pct25" => Drawing.PresetPatternValues.Percent25,
            "pct30" => Drawing.PresetPatternValues.Percent30,
            "pct40" => Drawing.PresetPatternValues.Percent40,
            "pct50" => Drawing.PresetPatternValues.Percent50,
            "pct60" => Drawing.PresetPatternValues.Percent60,
            "pct70" => Drawing.PresetPatternValues.Percent70,
            "pct75" => Drawing.PresetPatternValues.Percent75,
            "pct80" => Drawing.PresetPatternValues.Percent80,
            "pct90" => Drawing.PresetPatternValues.Percent90,
            "dkhorz" => Drawing.PresetPatternValues.DarkHorizontal,
            "dkvert" => Drawing.PresetPatternValues.DarkVertical,
            "dkdndiag" => Drawing.PresetPatternValues.DarkDownwardDiagonal,
            "dkupdiag" => Drawing.PresetPatternValues.DarkUpwardDiagonal,
            "lthorz" => Drawing.PresetPatternValues.LightHorizontal,
            "ltvert" => Drawing.PresetPatternValues.LightVertical,
            "ltdndiag" => Drawing.PresetPatternValues.LightDownwardDiagonal,
            "ltupdiag" => Drawing.PresetPatternValues.LightUpwardDiagonal,
            "narhorz" => Drawing.PresetPatternValues.NarrowHorizontal,
            "narvert" => Drawing.PresetPatternValues.NarrowVertical,
            "horz" or "horizontal" => Drawing.PresetPatternValues.Horizontal,
            "vert" or "vertical" => Drawing.PresetPatternValues.Vertical,
            "dndiag" or "downdiag" => Drawing.PresetPatternValues.DownwardDiagonal,
            "updiag" => Drawing.PresetPatternValues.UpwardDiagonal,
            "wdupdiag" => Drawing.PresetPatternValues.WideUpwardDiagonal,
            "wddndiag" => Drawing.PresetPatternValues.WideDownwardDiagonal,
            "dashhorz" => Drawing.PresetPatternValues.DashedHorizontal,
            "dashvert" => Drawing.PresetPatternValues.DashedVertical,
            "dashdndiag" => Drawing.PresetPatternValues.DashedDownwardDiagonal,
            "dashupdiag" => Drawing.PresetPatternValues.DashedUpwardDiagonal,
            "smconfetti" => Drawing.PresetPatternValues.SmallConfetti,
            "lgconfetti" => Drawing.PresetPatternValues.LargeConfetti,
            "zigzag" => Drawing.PresetPatternValues.ZigZag,
            "wave" => Drawing.PresetPatternValues.Wave,
            "diagbrick" => Drawing.PresetPatternValues.DiagonalBrick,
            // R9b: `diagStripe` is a common user-facing alias for a diagonal
            // stripe; OOXML has no literal "diagStripe" token, so map it to the
            // closest preset (light upward diagonal stripes).
            "diagstripe" => Drawing.PresetPatternValues.LightUpwardDiagonal,
            "horzbrick" => Drawing.PresetPatternValues.HorizontalBrick,
            "weave" => Drawing.PresetPatternValues.Weave,
            "plaid" => Drawing.PresetPatternValues.Plaid,
            "divot" => Drawing.PresetPatternValues.Divot,
            "dotgrid" => Drawing.PresetPatternValues.DotGrid,
            "dotdiamond" => Drawing.PresetPatternValues.DottedDiamond,
            "shingle" => Drawing.PresetPatternValues.Shingle,
            "trellis" => Drawing.PresetPatternValues.Trellis,
            "sphere" => Drawing.PresetPatternValues.Sphere,
            "smgrid" => Drawing.PresetPatternValues.SmallGrid,
            "lggrid" => Drawing.PresetPatternValues.LargeGrid,
            "smcheck" => Drawing.PresetPatternValues.SmallCheck,
            "lgcheck" => Drawing.PresetPatternValues.LargeCheck,
            "openDmnd" or "opendmnd" => Drawing.PresetPatternValues.OpenDiamond,
            "solidDmnd" or "soliddmnd" => Drawing.PresetPatternValues.SolidDiamond,
            "cross" => Drawing.PresetPatternValues.Cross,
            "diagcross" => Drawing.PresetPatternValues.DiagonalCross,
            _ => throw new ArgumentException(
                $"Unknown pattern preset: '{name}'. Examples: pct25, ltHorz, cross, diagCross, weave, zigZag, wave, diagBrick, plaid.")
        };
    }

    /// <summary>
    /// Apply image (blip) fill to a shape.
    /// Format: file path to image, e.g. "/tmp/bg.png"
    /// </summary>
    private static void ApplyShapeImageFill(ShapeProperties spPr, string imagePath, SlidePart part,
        string? fillRectSpec = null, string? srcRectSpec = null)
    {
        var (stream, partType) = OfficeCli.Core.ImageSource.Resolve(imagePath);
        using var streamDispose = stream;

        var imagePart = part.AddImagePart(partType);
        imagePart.FeedData(stream);
        var relId = part.GetIdOfPart(imagePart);

        spPr.RemoveAllChildren<Drawing.SolidFill>();
        spPr.RemoveAllChildren<Drawing.NoFill>();
        spPr.RemoveAllChildren<Drawing.GradientFill>();
        spPr.RemoveAllChildren<Drawing.BlipFill>();
        spPr.RemoveAllChildren<Drawing.PatternFill>();

        var blipFill = new Drawing.BlipFill();
        blipFill.Append(new Drawing.Blip { Embed = relId });
        // CT_BlipFillProperties order: blip → srcRect → (tile|stretch). srcRect
        // (crop) and the stretch fillRect carry the image's framing; honor both
        // so dump→replay reproduces a banner image stretched past its shape
        // bounds (negative fillRect insets) instead of snapping to exact-fit.
        var sr = ParsePerMilleRect(srcRectSpec);
        if (sr.HasValue)
            blipFill.Append(new Drawing.SourceRectangle { Left = sr.Value.L, Top = sr.Value.T, Right = sr.Value.R, Bottom = sr.Value.B });
        var fr = new Drawing.FillRectangle();
        var frp = ParsePerMilleRect(fillRectSpec);
        if (frp.HasValue)
        { fr.Left = frp.Value.L; fr.Top = frp.Value.T; fr.Right = frp.Value.R; fr.Bottom = frp.Value.B; }
        blipFill.Append(new Drawing.Stretch(fr));
        InsertFillElement(spPr, blipFill);
    }

    /// <summary>
    /// Parse a "l,t,r,b" perMille rect spec (the form PictureToNode / shape
    /// blipFill readback emit) into four nullable int insets, or null when the
    /// spec is absent/malformed. Each component falls back to null (left unset)
    /// when it doesn't parse, so a partial spec degrades cleanly.
    /// </summary>
    private static (int? L, int? T, int? R, int? B)? ParsePerMilleRect(string? spec)
    {
        if (string.IsNullOrWhiteSpace(spec)) return null;
        var parts = spec.Split(',', StringSplitOptions.TrimEntries);
        if (parts.Length != 4) return null;
        static int? P(string s) => int.TryParse(s, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : (int?)null;
        return (P(parts[0]), P(parts[1]), P(parts[2]), P(parts[3]));
    }

    /// <summary>
    /// Apply text margin (padding) to a BodyProperties element.
    /// Supports: single value "0.5cm" (all sides), or "left,top,right,bottom" e.g. "0.5cm,0.3cm,0.5cm,0.3cm"
    /// </summary>
    private static void ApplyTextMargin(Drawing.BodyProperties bodyPr, string value)
    {
        // Maximum reasonable inset magnitude: ~142cm (max slide dimension in
        // OOXML = 51206400 EMU). Insets are ST_Coordinate32 (xsd:int) and MAY be
        // negative — PowerPoint uses a small negative lIns/tIns/rIns/bIns to let
        // text bleed slightly past the shape box (seen in the wild as lIns="-1").
        // The prior ParseEmuAsInt rejected any negative inset, which aborted the
        // whole `add shape` op on round-trip (and, for a grouped shape, threw off
        // the group's child count → a cascade of "Shape N not found in group").
        const int MaxInsetEmu = 51206400;
        static int ParseInset(string raw)
        {
            var v = Core.EmuConverter.ParseEmu(raw); // long; allows negative
            if (Math.Abs(v) > MaxInsetEmu)
                throw new ArgumentException($"Inset value {v} EMU exceeds maximum allowed magnitude ({MaxInsetEmu} EMU / ~142cm).");
            return (int)v;
        }

        var parts = value.Split(',');
        if (parts.Length == 1)
        {
            var emu = ParseInset(parts[0]);
            bodyPr.LeftInset = emu;
            bodyPr.TopInset = emu;
            bodyPr.RightInset = emu;
            bodyPr.BottomInset = emu;
        }
        else if (parts.Length == 4)
        {
            // CONSISTENCY(margin-sparse-roundtrip): "-" placeholder means
            // "leave this side unset" so dump->replay can preserve sparse
            // source bodyPr (e.g. `lIns=180000 tIns=180000 rIns=150000`
            // with NO bIns must NOT round-trip into a fabricated
            // bIns=45720 default). NodeBuilder emits the dash form when
            // any side is null on the source bodyPr.
            for (int i = 0; i < 4; i++)
            {
                var raw = parts[i].Trim();
                if (raw == "-" || raw.Length == 0) continue;
                var v = ParseInset(raw);
                switch (i)
                {
                    case 0: bodyPr.LeftInset = v; break;
                    case 1: bodyPr.TopInset = v; break;
                    case 2: bodyPr.RightInset = v; break;
                    case 3: bodyPr.BottomInset = v; break;
                }
            }
        }
        else
        {
            throw new ArgumentException("margin must be a single value or 4 comma-separated values (left,top,right,bottom)");
        }
    }

    private static Drawing.TextAlignmentTypeValues ParseTextAlignment(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" or "l" => Drawing.TextAlignmentTypeValues.Left,
            "center" or "c" or "ctr" => Drawing.TextAlignmentTypeValues.Center,
            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
            "justify" or "j" or "just" => Drawing.TextAlignmentTypeValues.Justified,
            // OOXML ST_TextAlignType also defines justLow (low-justify),
            // dist (distributed) and thaiDist (Thai distributed). Get emits the
            // raw token for these (the shape-level readback passes them through),
            // so accept both the token and a friendly alias to keep round-trip.
            "justlow" => Drawing.TextAlignmentTypeValues.JustifiedLow,
            "dist" or "distributed" => Drawing.TextAlignmentTypeValues.Distributed,
            "thdist" or "thaidist" or "thaidistributed" => Drawing.TextAlignmentTypeValues.ThaiDistributed,
            _ => throw new ArgumentException($"Invalid align: {value}. Use: left, center, right, justify, justLow, dist, thDist")
        };

    /// <summary>
    /// Apply list style (bullet/numbered) to ParagraphProperties.
    /// Values: "bullet" or "•", "numbered" or "1", "alpha" or "a", "roman" or "i", "none"
    /// </summary>
    private static void ApplyListStyle(Drawing.ParagraphProperties pProps, string value,
                                       bool preserveIndent = false)
    {
        // Validate and construct FIRST — an invalid value must throw before
        // any mutation. Stripping the existing bullet group ahead of the
        // value check left a half-mutated DOM behind the invalid_value error,
        // and the single-command path autosaves the dirty DOM on Dispose, so
        // the paragraph's explicit "no bullet" override was silently lost.
        OpenXmlElement bullet = value.ToLowerInvariant() switch
        {
            "bullet" or "•" or "disc" => new Drawing.CharacterBullet { Char = "•" },
            "dash" or "-" or "–" => new Drawing.CharacterBullet { Char = "–" },
            "arrow" or ">" or "→" => new Drawing.CharacterBullet { Char = "→" },
            "check" or "✓" => new Drawing.CharacterBullet { Char = "✓" },
            "star" or "★" => new Drawing.CharacterBullet { Char = "★" },
            "numbered" or "number" or "1" => new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.ArabicPeriod },
            "alpha" or "a" => new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod },
            "alphaupper" => new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod },
            "roman" or "i" => new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod },
            "romanupper" => new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod },
            "none" or "false" => new Drawing.NoBullet(),
            _ => value.Length <= 2
                ? new Drawing.CharacterBullet { Char = value }
                : throw new ArgumentException($"Invalid list style: {value}. Use: bullet, numbered, alpha, roman, none, or a single character"),
        };

        pProps.RemoveAllChildren<Drawing.CharacterBullet>();
        pProps.RemoveAllChildren<Drawing.AutoNumberedBullet>();
        pProps.RemoveAllChildren<Drawing.NoBullet>();
        pProps.RemoveAllChildren<Drawing.BulletFont>();

        // CONSISTENCY(schema-order-pptx): bu* children rank BEFORE tabLst and
        // defRPr in CT_TextParagraphProperties — a pPr that already carries
        // <a:defRPr> (kern/spacing set first) must not get the bullet appended
        // after it, or PowerPoint silently ignores it and the strict validator
        // flags the file. Route through InsertPPrChild like every other pPr
        // child injection.
        InsertPPrChild(pProps, bullet);

        if (bullet is Drawing.NoBullet)
        {
            // Interactive convenience: removing the bullet also clears the
            // hanging indent. Skipped when the same property bag carries an
            // explicit indent/marginLeft — key-iteration order is
            // undefined, so list=none must not erase a sibling indent=0pt
            // that was (or will be) applied in the same Set call.
            if (!preserveIndent)
            {
                pProps.LeftMargin = null;
                pProps.Indent = null;
            }
            return;
        }

        // Apply default hanging indent for bullet/numbered lists (matches PowerPoint defaults)
        if (pProps.LeftMargin == null)
            pProps.LeftMargin = 457200; // 0.5 inch
        if (pProps.Indent == null)
            pProps.Indent = -457200; // hanging indent
    }

    private static Drawing.ShapeTypeValues ParsePresetShape(string name) =>
        name.ToLowerInvariant() switch
        {
            "rect" or "rectangle" => Drawing.ShapeTypeValues.Rectangle,
            "roundrect" or "roundedrectangle" => Drawing.ShapeTypeValues.RoundRectangle,
            "ellipse" or "oval" => Drawing.ShapeTypeValues.Ellipse,
            "triangle" => Drawing.ShapeTypeValues.Triangle,
            "rtriangle" or "righttriangle" or "rttriangle" => Drawing.ShapeTypeValues.RightTriangle,
            "diamond" => Drawing.ShapeTypeValues.Diamond,
            "parallelogram" => Drawing.ShapeTypeValues.Parallelogram,
            "trapezoid" => Drawing.ShapeTypeValues.Trapezoid,
            "pentagon" => Drawing.ShapeTypeValues.Pentagon,
            "hexagon" => Drawing.ShapeTypeValues.Hexagon,
            "heptagon" => Drawing.ShapeTypeValues.Heptagon,
            "octagon" => Drawing.ShapeTypeValues.Octagon,
            "star4" => Drawing.ShapeTypeValues.Star4,
            "star5" => Drawing.ShapeTypeValues.Star5,
            "star6" => Drawing.ShapeTypeValues.Star6,
            "star8" => Drawing.ShapeTypeValues.Star8,
            "star10" => Drawing.ShapeTypeValues.Star10,
            "star12" => Drawing.ShapeTypeValues.Star12,
            "star16" => Drawing.ShapeTypeValues.Star16,
            "star24" => Drawing.ShapeTypeValues.Star24,
            "star32" => Drawing.ShapeTypeValues.Star32,
            // "arrow" alias mirrors PowerPoint's "Arrow: Right" UI label —
            // the unqualified short form users naturally reach for.
            "rightarrow" or "rarrow" or "arrow" => Drawing.ShapeTypeValues.RightArrow,
            "leftarrow" or "larrow" => Drawing.ShapeTypeValues.LeftArrow,
            "uparrow" => Drawing.ShapeTypeValues.UpArrow,
            "downarrow" => Drawing.ShapeTypeValues.DownArrow,
            "leftrightarrow" or "lrarrow" => Drawing.ShapeTypeValues.LeftRightArrow,
            "updownarrow" or "udarrow" => Drawing.ShapeTypeValues.UpDownArrow,
            "chevron" => Drawing.ShapeTypeValues.Chevron,
            "homeplat" or "homeplate" => Drawing.ShapeTypeValues.HomePlate,
            "plus" or "cross" => Drawing.ShapeTypeValues.Plus,
            "heart" => Drawing.ShapeTypeValues.Heart,
            "cloud" => Drawing.ShapeTypeValues.Cloud,
            "lightning" or "lightningbolt" => Drawing.ShapeTypeValues.LightningBolt,
            "sun" => Drawing.ShapeTypeValues.Sun,
            "moon" => Drawing.ShapeTypeValues.Moon,
            "arc" => Drawing.ShapeTypeValues.Arc,
            "donut" => Drawing.ShapeTypeValues.Donut,
            // blockArc is NOT noSmoking: blockArc (ST_ShapeType.blockArc) carries
            // three adjust handles (adj1/adj2/adj3), noSmoking carries one (adj).
            // Collapsing blockArc → NoSmoking emitted <a:prstGeom prst="noSmoking">
            // with three <a:gd> children, which is schema-invalid for noSmoking and
            // makes PowerPoint refuse the whole file. Let blockArc fall through to
            // the reflection lookup, which resolves Drawing.ShapeTypeValues.BlockArc.
            "nosmoking" => Drawing.ShapeTypeValues.NoSmoking,
            "cube" => Drawing.ShapeTypeValues.Cube,
            "can" or "cylinder" => Drawing.ShapeTypeValues.Can,
            "line" => Drawing.ShapeTypeValues.Line,
            "decagon" => Drawing.ShapeTypeValues.Decagon,
            "dodecagon" => Drawing.ShapeTypeValues.Dodecagon,
            "ribbon" => Drawing.ShapeTypeValues.Ribbon,
            "ribbon2" => Drawing.ShapeTypeValues.Ribbon2,
            "callout1" => Drawing.ShapeTypeValues.Callout1,
            "callout2" => Drawing.ShapeTypeValues.Callout2,
            "callout3" => Drawing.ShapeTypeValues.Callout3,
            "wedgeroundrectcallout" or "callout" => Drawing.ShapeTypeValues.WedgeRoundRectangleCallout,
            "wedgeellipsecallout" => Drawing.ShapeTypeValues.WedgeEllipseCallout,
            "cloudcallout" => Drawing.ShapeTypeValues.CloudCallout,
            "flowchartprocess" or "process" => Drawing.ShapeTypeValues.FlowChartProcess,
            "flowchartdecision" or "decision" => Drawing.ShapeTypeValues.FlowChartDecision,
            "flowchartterminator" or "terminator" => Drawing.ShapeTypeValues.FlowChartTerminator,
            "flowchartdocument" or "document" => Drawing.ShapeTypeValues.FlowChartDocument,
            "flowchartinputoutput" or "flowchartio" or "io" => Drawing.ShapeTypeValues.FlowChartInputOutput,
            "flowchartdata" or "data" => Drawing.ShapeTypeValues.FlowChartInputOutput,
            "flowchartpredefinedprocess" or "predefinedprocess" => Drawing.ShapeTypeValues.FlowChartPredefinedProcess,
            "flowchartpreparation" or "preparation" => Drawing.ShapeTypeValues.FlowChartPreparation,
            "flowchartmanualinput" or "manualinput" => Drawing.ShapeTypeValues.FlowChartManualInput,
            "flowchartmanualoperation" or "manualoperation" => Drawing.ShapeTypeValues.FlowChartManualOperation,
            "flowchartconnector" or "flowconnector" => Drawing.ShapeTypeValues.FlowChartConnector,
            "flowchartoffpageconnector" or "offpageconnector" => Drawing.ShapeTypeValues.FlowChartOffpageConnector,
            "flowchartmultidocument" or "multidocument" => Drawing.ShapeTypeValues.FlowChartMultidocument,
            "flowchartsort" or "sort" => Drawing.ShapeTypeValues.FlowChartSort,
            "flowchartmerge" or "merge" => Drawing.ShapeTypeValues.FlowChartMerge,
            "flowchartextract" or "extract" => Drawing.ShapeTypeValues.FlowChartExtract,
            "flowchartdelay" or "delay" => Drawing.ShapeTypeValues.FlowChartDelay,
            "flowchartdisplay" or "display" => Drawing.ShapeTypeValues.FlowChartDisplay,
            "flowchartalternateprocess" or "alternateprocess" => Drawing.ShapeTypeValues.FlowChartAlternateProcess,
            "brace" or "leftbrace" => Drawing.ShapeTypeValues.LeftBrace,
            "rightbrace" => Drawing.ShapeTypeValues.RightBrace,
            "leftbracket" => Drawing.ShapeTypeValues.LeftBracket,
            "rightbracket" => Drawing.ShapeTypeValues.RightBracket,
            "smileyface" or "smiley" => Drawing.ShapeTypeValues.SmileyFace,
            "foldedcorner" => Drawing.ShapeTypeValues.FoldedCorner,
            "frame" => Drawing.ShapeTypeValues.Frame,
            "gear6" => Drawing.ShapeTypeValues.Gear6,
            "gear9" => Drawing.ShapeTypeValues.Gear9,
            "notchedrightarrow" => Drawing.ShapeTypeValues.NotchedRightArrow,
            "bentuparrow" => Drawing.ShapeTypeValues.BentUpArrow,
            "curvedrightarrow" => Drawing.ShapeTypeValues.CurvedRightArrow,
            "stripedrightarrow" => Drawing.ShapeTypeValues.StripedRightArrow,
            "uturnarrow" => Drawing.ShapeTypeValues.UTurnArrow,
            "circulararrow" => Drawing.ShapeTypeValues.CircularArrow,
            // ParsePresetShape can't enumerate all ~180 OOXML preset values by hand.
            // Fall back to reflection lookup on Drawing.ShapeTypeValues static
            // properties so dump-replay survives preset names absent from the
            // hand-rolled list (pie, chord, blockArc, mathDivide, callouts, …).
            // Last-resort degrade is Rectangle so a single missing preset never
            // takes the whole shape add down (which would cascade positional refs).
            _ => ResolveShapeTypeByReflection(name) ?? Drawing.ShapeTypeValues.Rectangle,
        };

    private static Drawing.ShapeTypeValues? ResolveShapeTypeByReflection(string name)
    {
        var lower = name.ToLowerInvariant();
        var props = typeof(Drawing.ShapeTypeValues).GetProperties(
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        foreach (var p in props)
        {
            if (p.PropertyType != typeof(Drawing.ShapeTypeValues)) continue;
            // Match on the C# property name (e.g. "RoundRectangle") first.
            if (string.Equals(p.Name, lower, StringComparison.OrdinalIgnoreCase))
                return (Drawing.ShapeTypeValues?)p.GetValue(null);
            // …then on the OOXML serialized token (e.g. "round2SameRect"),
            // which is what dump emits and is what the prstGeom@prst attribute
            // actually uses. The C# property name diverges from the token for
            // many presets (Round2SameRectangle → round2SameRect, RoundRectangle
            // → roundRect, …), so matching only the property name dropped every
            // such preset to the Rectangle degrade. Read the token the same way
            // NodeBuilder does — via the rendered Preset.InnerText.
            var value = (Drawing.ShapeTypeValues?)p.GetValue(null);
            if (value != null)
            {
                var token = new Drawing.PresetGeometry { Preset = value }.Preset?.InnerText;
                if (string.Equals(token, lower, StringComparison.OrdinalIgnoreCase))
                    return value;
            }
        }
        return null;
    }

    // Strict variant used by Set to surface unknown preset names as an
    // unsupported_property error instead of silently degrading to Rectangle.
    // ParsePresetShape's last-resort degrade exists so a single bad preset
    // never takes a whole shape add (and its positional refs) down on import,
    // but Set is a single-property mutation — silently rewriting the user's
    // geometry to a rectangle is a worse outcome than telling them the name
    // wasn't recognised.
    internal static bool TryParsePresetShape(string name, out Drawing.ShapeTypeValues value)
    {
        try
        {
            var explicitHit = ParsePresetShape(name);
            // ParsePresetShape returns Rectangle for unknown names — to
            // distinguish a real "rect" input from the degraded fallback,
            // also check the explicit alias list / reflection path.
            var lower = name.ToLowerInvariant();
            if (lower is "rect" or "rectangle") { value = explicitHit; return true; }
            if (ResolveShapeTypeByReflection(name) != null) { value = explicitHit; return true; }
            // Heuristic: if ParsePresetShape gave us anything other than the
            // Rectangle fallback, it matched a hand-rolled alias. (Rectangle
            // is the only fallback target, so any non-Rectangle return is a
            // real hit; a Rectangle return without "rect"/"rectangle" input
            // is the silent degrade we want to reject.)
            if (explicitHit != Drawing.ShapeTypeValues.Rectangle) { value = explicitHit; return true; }
            value = explicitHit;
            return false;
        }
        catch
        {
            value = Drawing.ShapeTypeValues.Rectangle;
            return false;
        }
    }

    // BUG-FIX(B8): canonical names mirror OOXML LineEndValues so that the
    // value passed to Add/Set round-trips through Get. The previous mapping
    // had 'arrow' → Triangle (input) but Get emitted the OOXML name 'arrow'
    // for LineEndValues.Arrow, producing input/output asymmetry. Aliases
    // (open/closed/circle) are accepted but Get always returns the canonical
    // OOXML token (triangle, arrow, stealth, diamond, oval, none).
    private static Drawing.LineEndValues ParseLineEndType(string name) =>
        name.ToLowerInvariant() switch
        {
            "triangle" or "closed" => Drawing.LineEndValues.Triangle,
            "stealth" => Drawing.LineEndValues.Stealth,
            "diamond" => Drawing.LineEndValues.Diamond,
            "oval" or "circle" => Drawing.LineEndValues.Oval,
            "arrow" or "open" => Drawing.LineEndValues.Arrow,
            "none" => Drawing.LineEndValues.None,
            _ => throw new ArgumentException(
                $"Invalid line end type: '{name}'. Valid values: triangle, arrow, stealth, diamond, oval, none.")
        };

    // R4-5: map a size token to the @w (width) and @len (length) line-end enums.
    // CT_LineEndProperties models width and length as SEPARATE enums, so resolve
    // both. Returns false for an unrecognized token.
    private static bool TryParseLineEndSize(string value,
        out Drawing.LineEndWidthValues width, out Drawing.LineEndLengthValues length)
    {
        switch (value.Trim().ToLowerInvariant())
        {
            case "small" or "sm" or "s":
                width = Drawing.LineEndWidthValues.Small; length = Drawing.LineEndLengthValues.Small; return true;
            case "medium" or "med" or "m":
                width = Drawing.LineEndWidthValues.Medium; length = Drawing.LineEndLengthValues.Medium; return true;
            case "large" or "lg" or "l" or "big":
                width = Drawing.LineEndWidthValues.Large; length = Drawing.LineEndLengthValues.Large; return true;
            default:
                width = Drawing.LineEndWidthValues.Medium; length = Drawing.LineEndLengthValues.Medium; return false;
        }
    }

    // full prstDash enum (was clipped to 6 of 11 values; sysDot/sysDash/
    // sysDashDot/sysDashDotDot/lgDashDotDot threw "Invalid lineDash"). Mirrors
    // ST_PresetLineDashVal (DrawingML §20.1.10.49). Accepts canonical OOXML
    // tokens plus longstanding 'longdash[dot]' aliases for backward compat.
    internal static Drawing.PresetLineDashValues ParseLineDashValue(string value) =>
        value.ToLowerInvariant() switch
        {
            "solid" => Drawing.PresetLineDashValues.Solid,
            "dot" => Drawing.PresetLineDashValues.Dot,
            "dash" => Drawing.PresetLineDashValues.Dash,
            "dashdot" or "dash_dot" => Drawing.PresetLineDashValues.DashDot,
            "lgdash" or "lg_dash" or "longdash" => Drawing.PresetLineDashValues.LargeDash,
            "lgdashdot" or "lg_dash_dot" or "longdashdot" => Drawing.PresetLineDashValues.LargeDashDot,
            "lgdashdotdot" or "lg_dash_dot_dot" or "longdashdotdot" => Drawing.PresetLineDashValues.LargeDashDotDot,
            "sysdot" or "sys_dot" => Drawing.PresetLineDashValues.SystemDot,
            "sysdash" or "sys_dash" => Drawing.PresetLineDashValues.SystemDash,
            "sysdashdot" or "sys_dash_dot" => Drawing.PresetLineDashValues.SystemDashDot,
            "sysdashdotdot" or "sys_dash_dot_dot" => Drawing.PresetLineDashValues.SystemDashDotDot,
            _ => throw new ArgumentException(
                $"Invalid 'lineDash' value: '{value}'. Valid values: solid, dot, dash, dashdot, lgDash, lgDashDot, lgDashDotDot, sysDot, sysDash, sysDashDot, sysDashDotDot.")
        };

    // R64 bt-3: Parse a verbatim <a:custDash>…</a:custDash> string and
    // reconstruct a typed Drawing.CustomDash. Mirrors the lift-attrs +
    // InnerXml pattern used by shadowRaw / fillOverlayRaw / effectDagRaw /
    // effectsRaw passthrough installs in ShapeProperties — keeps a single
    // round-trip strategy for "no compressible string form" OOXML children.
    // <a:custDash> itself has no attributes; the payload is <a:ds d="N" sp="N"/>
    // stops.
    internal static Drawing.CustomDash BuildCustomDashFromRaw(string raw)
    {
        var xml = raw.Contains("xmlns:a=")
            ? raw
            : raw.Replace("<a:custDash",
                "<a:custDash xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"");
        var custDash = new Drawing.CustomDash();
        using var sr = new System.IO.StringReader(xml);
        using var xr = System.Xml.XmlReader.Create(sr);
        xr.MoveToContent();
        if (xr.HasAttributes)
        {
            while (xr.MoveToNextAttribute())
            {
                if (xr.Prefix == "xmlns" || xr.Name == "xmlns") continue;
                custDash.SetAttribute(new OpenXmlAttribute(
                    xr.Prefix, xr.LocalName, xr.NamespaceURI, xr.Value));
            }
            xr.MoveToElement();
        }
        if (!xr.IsEmptyElement)
        {
            var inner = xr.ReadInnerXml();
            if (!string.IsNullOrWhiteSpace(inner))
                custDash.InnerXml = inner;
        }
        return custDash;
    }
}
