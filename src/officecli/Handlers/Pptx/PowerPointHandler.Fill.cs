// Copyright 2025 OfficeCLI (officecli.ai)
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
        // Schema order: xfrm → prstGeom → fill → ln → effectLst
        var prstGeom = spPr.GetFirstChild<Drawing.PresetGeometry>();
        if (prstGeom != null)
            spPr.InsertAfter(fillElement, prstGeom);
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
        return null;
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
        return null;
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
    /// Apply pattern fill to ShapeProperties.
    /// Format: "<preset>" or "<preset>:<fgColor>" or "<preset>:<fgColor>:<bgColor>"
    ///   preset: e.g. pct25, ltHorz, dkCross, weave, zigZag (Drawing.PresetPatternValues)
    ///   fgColor / bgColor: lenient hex/named/scheme color (defaults: fg=000000, bg=FFFFFF)
    /// Examples: "pct25", "ltHorz:FF0000", "dkCross:red:white"
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
                $"Unknown pattern preset: '{name}'. Examples: pct25, ltHorz, dkCross, weave, zigZag, wave, diagBrick, plaid.")
        };
    }

    /// <summary>
    /// Apply image (blip) fill to a shape.
    /// Format: file path to image, e.g. "/tmp/bg.png"
    /// </summary>
    private static void ApplyShapeImageFill(ShapeProperties spPr, string imagePath, SlidePart part)
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
        blipFill.Append(new Drawing.Stretch(new Drawing.FillRectangle()));
        InsertFillElement(spPr, blipFill);
    }

    /// <summary>
    /// Apply text margin (padding) to a BodyProperties element.
    /// Supports: single value "0.5cm" (all sides), or "left,top,right,bottom" e.g. "0.5cm,0.3cm,0.5cm,0.3cm"
    /// </summary>
    private static void ApplyTextMargin(Drawing.BodyProperties bodyPr, string value)
    {
        // Maximum reasonable inset: ~142cm (max slide dimension in OOXML = 51206400 EMU)
        const int MaxInsetEmu = 51206400;

        var parts = value.Split(',');
        if (parts.Length == 1)
        {
            var emu = Core.EmuConverter.ParseEmuAsInt(parts[0]);
            if (emu > MaxInsetEmu)
                throw new ArgumentException($"Inset value {emu} EMU exceeds maximum allowed ({MaxInsetEmu} EMU / ~142cm).");
            bodyPr.LeftInset = emu;
            bodyPr.TopInset = emu;
            bodyPr.RightInset = emu;
            bodyPr.BottomInset = emu;
        }
        else if (parts.Length == 4)
        {
            var insets = new int[4];
            for (int i = 0; i < 4; i++)
            {
                insets[i] = Core.EmuConverter.ParseEmuAsInt(parts[i].Trim());
                if (insets[i] > MaxInsetEmu)
                    throw new ArgumentException($"Inset value {insets[i]} EMU exceeds maximum allowed ({MaxInsetEmu} EMU / ~142cm).");
            }
            bodyPr.LeftInset = insets[0];
            bodyPr.TopInset = insets[1];
            bodyPr.RightInset = insets[2];
            bodyPr.BottomInset = insets[3];
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
            "center" or "c" => Drawing.TextAlignmentTypeValues.Center,
            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
            "justify" or "j" => Drawing.TextAlignmentTypeValues.Justified,
            _ => throw new ArgumentException($"Invalid align: {value}. Use: left, center, right, justify")
        };

    /// <summary>
    /// Apply list style (bullet/numbered) to ParagraphProperties.
    /// Values: "bullet" or "•", "numbered" or "1", "alpha" or "a", "roman" or "i", "none"
    /// </summary>
    private static void ApplyListStyle(Drawing.ParagraphProperties pProps, string value)
    {
        pProps.RemoveAllChildren<Drawing.CharacterBullet>();
        pProps.RemoveAllChildren<Drawing.AutoNumberedBullet>();
        pProps.RemoveAllChildren<Drawing.NoBullet>();
        pProps.RemoveAllChildren<Drawing.BulletFont>();

        switch (value.ToLowerInvariant())
        {
            case "bullet" or "•" or "disc":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "•" });
                break;
            case "dash" or "-" or "–":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "–" });
                break;
            case "arrow" or ">" or "→":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "→" });
                break;
            case "check" or "✓":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "✓" });
                break;
            case "star" or "★":
                pProps.AppendChild(new Drawing.CharacterBullet { Char = "★" });
                break;
            case "numbered" or "number" or "1":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.ArabicPeriod });
                break;
            case "alpha" or "a":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod });
                break;
            case "alphaupper" or "A":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.AlphaUpperCharacterPeriod });
                break;
            case "roman" or "i":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.RomanLowerCharacterPeriod });
                break;
            case "romanupper" or "I":
                pProps.AppendChild(new Drawing.AutoNumberedBullet { Type = Drawing.TextAutoNumberSchemeValues.RomanUpperCharacterPeriod });
                break;
            case "none" or "false":
                pProps.AppendChild(new Drawing.NoBullet());
                pProps.LeftMargin = null;
                pProps.Indent = null;
                return;
            default:
                if (value.Length <= 2)
                    pProps.AppendChild(new Drawing.CharacterBullet { Char = value });
                else
                    throw new ArgumentException($"Invalid list style: {value}. Use: bullet, numbered, alpha, roman, none, or a single character");
                break;
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
            "nosmoking" or "blockarc" => Drawing.ShapeTypeValues.NoSmoking,
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
            if (string.Equals(p.Name, lower, StringComparison.OrdinalIgnoreCase))
                return (Drawing.ShapeTypeValues?)p.GetValue(null);
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
}
