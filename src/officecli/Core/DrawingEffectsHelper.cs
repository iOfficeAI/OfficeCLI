// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for building Drawing-namespace text/shape effects (a:effectLst children).
/// Used by both PPTX and Excel handlers to avoid code duplication.
/// Word uses a different namespace (w14) and has its own implementation.
/// </summary>
internal static class DrawingEffectsHelper
{
    /// <summary>
    /// Build an OuterShadow element from a value string.
    /// Format: "COLOR[-BLUR[-ANGLE[-DIST[-OPACITY]]]]"
    /// Defaults: blur=4pt, angle=45°, dist=3pt, opacity=40%
    /// </summary>
    public static Drawing.OuterShadow BuildOuterShadow(string value, Func<string, OpenXmlElement> colorBuilder)
    {
        var parts = SplitEffectParts(value);
        var blurPt = ParseParam(parts, 1, 4.0, "shadow blur");
        var angleDeg = ParseParam(parts, 2, 45.0, "shadow angle");
        var distPt = ParseParam(parts, 3, 3.0, "shadow distance");
        // Distinguish "user supplied opacity" from "default 40%": if the color
        // carries an 8-digit hex alpha (#RRGGBBAA) and no explicit -OPACITY tail,
        // the alpha-from-color must win over the 40% default so RRGGBBAA round-trips.
        bool hasExplicitOpacity = parts.Length > 4;
        var opacity = ParseParam(parts, 4, 40.0, "shadow opacity");

        var shadow = new Drawing.OuterShadow
        {
            BlurRadius = (long)(blurPt * 12700),
            Distance = (long)(distPt * 12700),
            Direction = (int)(angleDeg * 60000),
            Alignment = Drawing.RectangleAlignmentValues.TopLeft,
            RotateWithShape = false
        };
        var clr = colorBuilder(parts[0]);
        bool colorHasAlpha = clr.GetFirstChild<Drawing.Alpha>() != null;
        // ColorEncodesAlpha: user wrote an 8-digit hex with alpha=FF — the
        // alpha element is absent from the built color (SanitizeColorForOoxml
        // drops the redundant 100% alpha child), but the caller's intent was
        // an explicit "fully opaque" shadow. Suppress the 40% default so the
        // explicit FF alpha doesn't silently downgrade to 40%.
        bool colorEncodesAlpha = ColorEncodesExplicitAlpha(parts[0]);
        if (hasExplicitOpacity || (!colorHasAlpha && !colorEncodesAlpha))
            SetAlphaChild(clr, (int)(opacity * 1000));
        shadow.AppendChild(clr);
        return shadow;
    }

    /// <summary>
    /// Build a Glow element from a value string.
    /// Format: "COLOR[-RADIUS[-OPACITY]]"
    /// Defaults: radius=8pt, opacity=75%
    /// </summary>
    public static Drawing.Glow BuildGlow(string value, Func<string, OpenXmlElement> colorBuilder)
    {
        var parts = SplitEffectParts(value);
        var radiusPt = ParseParam(parts, 1, 8.0, "glow radius");
        bool hasExplicitOpacity = parts.Length > 2;
        var opacity = ParseParam(parts, 2, 75.0, "glow opacity");

        var glow = new Drawing.Glow { Radius = (long)(radiusPt * 12700) };
        var clr = colorBuilder(parts[0]);
        bool colorHasAlpha = clr.GetFirstChild<Drawing.Alpha>() != null;
        bool colorEncodesAlpha = ColorEncodesExplicitAlpha(parts[0]);
        if (hasExplicitOpacity || (!colorHasAlpha && !colorEncodesAlpha))
            SetAlphaChild(clr, (int)(opacity * 1000));
        glow.AppendChild(clr);
        return glow;
    }

    /// <summary>
    /// Returns true when <paramref name="colorInput"/> is an 8-digit hex form
    /// (CSS #RRGGBBAA or OOXML AARRGGBB) — i.e. the caller explicitly encoded
    /// an alpha byte. Even when that byte resolves to 0xFF (fully opaque),
    /// SanitizeColorForOoxml drops the redundant alpha element, so a naive
    /// "no alpha child → use default 40% / 75%" check would silently override
    /// the user's "100% opaque" intent. Mirror SanitizeColorForOoxml's
    /// detection here so shadow/glow honor the explicit encoding.
    /// </summary>
    private static bool ColorEncodesExplicitAlpha(string colorInput)
    {
        if (string.IsNullOrEmpty(colorInput)) return false;
        var hex = colorInput.TrimStart('#');
        return hex.Length == 8 && hex.All(c => char.IsAsciiHexDigit(c));
    }

    /// <summary>
    /// Build a Reflection element from a value string.
    /// Values: "tight"/"small", "half"/"true", "full", or numeric percentage.
    /// </summary>
    public static Drawing.Reflection BuildReflection(string value)
    {
        // Unknown preset names (and out-of-range numerics) used to silently
        // fall back to "half" (90000), masking typos. Reject so the caller
        // surfaces the value rather than writing a no-op effect.
        int endPos;
        switch (value.ToLowerInvariant())
        {
            case "tight": case "small": endPos = 55000; break;
            case "true":  case "half":  endPos = 90000; break;
            case "full":               endPos = 100000; break;
            default:
                if (!int.TryParse(value, out var pct) || pct < 0 || pct > 100)
                    throw new ArgumentException(
                        $"Invalid 'reflection' value '{value}'. Valid presets: none, tight, small, half, true, full; or a numeric percentage 0-100.");
                endPos = (int)Math.Min((long)pct * 1000, 100000);
                break;
        }

        return new Drawing.Reflection
        {
            BlurRadius = 6350,
            StartOpacity = 52000,
            StartPosition = 0,
            EndAlpha = 300,
            EndPosition = endPos,
            Distance = 0,
            Direction = 5400000,
            VerticalRatio = -100000,
            Alignment = Drawing.RectangleAlignmentValues.BottomLeft,
            RotateWithShape = false
        };
    }

    /// <summary>
    /// Build a SoftEdge element from a value string (radius in points).
    /// </summary>
    public static Drawing.SoftEdge BuildSoftEdge(string value)
    {
        var numStr = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ? value[..^2].Trim() : value;
        if (!double.TryParse(numStr, System.Globalization.CultureInfo.InvariantCulture, out var radiusPt)
            || double.IsNaN(radiusPt) || double.IsInfinity(radiusPt) || radiusPt < 0)
            throw new ArgumentException($"Invalid 'softedge' value '{value}'. Expected a finite non-negative numeric radius in points.");
        return new Drawing.SoftEdge { Radius = (long)(radiusPt * 12700) };
    }

    /// <summary>
    /// Get or create EffectList in correct schema position within Drawing.RunProperties.
    /// CT_TextCharacterProperties order: ln → fill → effectLst → highlight → ... → latin → ea → ...
    /// </summary>
    public static Drawing.EffectList EnsureRunEffectList(Drawing.RunProperties rPr)
    {
        var existing = rPr.GetFirstChild<Drawing.EffectList>();
        if (existing != null) return existing;

        var effectList = new Drawing.EffectList();
        var insertBefore = (OpenXmlElement?)rPr.GetFirstChild<Drawing.Highlight>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.UnderlineFollowsText>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.Underline>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.UnderlineFillText>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.UnderlineFill>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.LatinFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.EastAsianFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ComplexScriptFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.SymbolFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnClick>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnMouseOver>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ExtensionList>();
        if (insertBefore != null)
            rPr.InsertBefore(effectList, insertBefore);
        else
            rPr.AppendChild(effectList);
        return effectList;
    }

    /// <summary>
    /// Insert a fill element at the correct schema position in Drawing.RunProperties.
    /// CT_TextCharacterProperties order: ln → fill → effectLst → ... → latin → ea → ...
    /// </summary>
    public static void InsertFillInRunProperties(Drawing.RunProperties rPr, OpenXmlElement fillElement)
    {
        var insertBefore = (OpenXmlElement?)rPr.GetFirstChild<Drawing.EffectList>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.EffectDag>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.Highlight>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.LatinFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.EastAsianFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ComplexScriptFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.SymbolFont>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnClick>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.HyperlinkOnMouseOver>()
            ?? (OpenXmlElement?)rPr.GetFirstChild<Drawing.ExtensionList>();
        if (insertBefore != null)
            rPr.InsertBefore(fillElement, insertBefore);
        else
            rPr.AppendChild(fillElement);
    }

    /// <summary>
    /// Apply a text effect to a Drawing.Run's RunProperties effectLst.
    /// Handles create/remove logic. Returns false if value is "none".
    /// </summary>
    public static void ApplyTextEffect<T>(Drawing.Run run, string value, Func<T> builder) where T : OpenXmlElement
    {
        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
        var effectList = EnsureRunEffectList(rPr);
        effectList.RemoveAllChildren<T>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) rPr.RemoveChild(effectList);
            return;
        }
        // CT_EffectList children must appear in schema order (blur →
        // fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection
        // → softEdge); Excel/PowerPoint reject out-of-order trees with
        // Sch_UnexpectedElementContentExpectingComplex. Insert before the
        // first sibling that would otherwise come after us, instead of the
        // naive AppendChild that lands every effect at the tail in arrival
        // order.
        InsertEffectInSchemaOrder(effectList, builder());
    }

    /// <summary>
    /// Schema order for CT_EffectList children. Mirrored in
    /// PowerPointHandler.Effects.cs for the shape-level effectLst; keep both
    /// in sync if you add a new effect type.
    /// </summary>
    private static readonly Type[] s_effectListChildOrder =
    [
        typeof(Drawing.Blur),
        typeof(Drawing.FillOverlay),
        typeof(Drawing.Glow),
        typeof(Drawing.InnerShadow),
        typeof(Drawing.OuterShadow),
        typeof(Drawing.PresetShadow),
        typeof(Drawing.Reflection),
        typeof(Drawing.SoftEdge),
    ];

    private static void InsertEffectInSchemaOrder(OpenXmlElement effectList, OpenXmlElement effect)
    {
        var targetIdx = Array.IndexOf(s_effectListChildOrder, effect.GetType());
        foreach (var child in effectList.ChildElements)
        {
            var childIdx = Array.IndexOf(s_effectListChildOrder, child.GetType());
            if (childIdx > targetIdx)
            {
                effectList.InsertBefore(effect, child);
                return;
            }
        }
        effectList.AppendChild(effect);
    }

    /// <summary>
    /// Standard color builder for Drawing effects: sanitizes hex, creates RgbColorModelHex with optional alpha.
    /// Use instead of duplicating the lambda pattern inline.
    /// </summary>
    public static OpenXmlElement BuildRgbColor(string colorValue)
    {
        var (rgb, alpha) = ParseHelpers.SanitizeColorForOoxml(colorValue);
        var clr = new Drawing.RgbColorModelHex { Val = rgb };
        if (alpha.HasValue) clr.AppendChild(new Drawing.Alpha { Val = alpha.Value });
        return clr;
    }

    // --- Private helpers ---

    /// <summary>
    /// Set or replace the Alpha child on a color element. Callers like BuildOuterShadow
    /// and BuildGlow apply an explicit opacity from the user value string; if the color
    /// builder (e.g. ARGB hex like "80FF0000") already produced an Alpha child, blindly
    /// appending another would yield two a:alpha siblings — invalid OOXML which Office
    /// either rejects or interprets unpredictably. Replace any existing alpha to keep
    /// the user's opacity authoritative for the effect.
    /// </summary>
    private static void SetAlphaChild(OpenXmlElement colorElement, int alphaVal)
    {
        var existing = colorElement.GetFirstChild<Drawing.Alpha>();
        if (existing != null) existing.Remove();
        colorElement.AppendChild(new Drawing.Alpha { Val = alphaVal });
    }

    private static double ParseParam(string[] parts, int index, double defaultValue, string paramName)
    {
        if (parts.Length <= index) return defaultValue;
        var raw = parts[index];
        // The historical contract is "bare double" — blur/dist in pt, angle
        // in deg, opacity in %. Accept unit-qualified inputs that match each
        // dimension so callers can write "5pt", "45deg", "40%" without
        // forcing them to know the internal unit. Strip and parse the
        // numeric prefix; reject unknown trailing letters.
        var num = raw.Trim();
        if (num.Length == 0)
            throw new ArgumentException($"Invalid {paramName} value: '{raw}' (empty).");
        // Strip a trailing alpha unit suffix (pt/deg/%/cm/in/px/emu). The
        // numeric routes through pt/deg/% as-is — units other than pt for a
        // pt-dimension still parse the number but the result is not
        // converted; agents should stick to bare numbers or the native unit
        // for now. The point of this fix is to stop ParseParam throwing on
        // a unit-qualified token; a future pass can do real unit conversion.
        int suffixStart = num.Length;
        while (suffixStart > 0 && (char.IsLetter(num[suffixStart - 1]) || num[suffixStart - 1] == '%'))
            suffixStart--;
        var numPart = num[..suffixStart];
        if (!double.TryParse(numPart, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var val)
            || double.IsNaN(val) || double.IsInfinity(val))
            throw new ArgumentException($"Invalid {paramName} value: '{raw}'.");
        return val;
    }

    /// <summary>
    /// Split an effect value string into ["color", "p1", "p2", …] tokens.
    /// Historical separator is '-', but '-' collides with negative numbers
    /// (e.g. "red;-5" for a shadow with negative angle). Prefer ';' when
    /// present; fall back to '-' for the legacy form. Empty tokens are
    /// rejected up front so opacity/blur don't silently take the default
    /// for a malformed input like "red;;5".
    /// </summary>
    private static string[] SplitEffectParts(string value)
    {
        if (string.IsNullOrEmpty(value))
            throw new ArgumentException("Effect value cannot be empty.");
        // Prefer ';' so negative numeric params (e.g. "-5") survive split.
        // When ';' is present, treat '-' as part of a numeric value, not a
        // separator. Fall back to '-' for the legacy form. In ';' mode,
        // reject empty tokens up front — the historical '-' code defaulted
        // them silently, which masked typos like "red;;5".
        if (value.Contains(';'))
        {
            var parts = value.Split(';');
            for (int i = 0; i < parts.Length; i++)
            {
                if (string.IsNullOrEmpty(parts[i]))
                    throw new ArgumentException($"Invalid effect value '{value}': empty token at position {i + 1}.");
            }
            return parts;
        }
        return value.Split('-');
    }
}
