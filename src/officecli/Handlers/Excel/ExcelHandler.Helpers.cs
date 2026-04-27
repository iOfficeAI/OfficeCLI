// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Validate a sheet name against Excel's rules. Throws ArgumentException
    /// with a clear message on the first rule violation. Rules:
    ///   - non-empty, non-whitespace
    ///   - max 31 chars
    ///   - cannot contain  \  /  ?  *  :  [  ]
    ///   - cannot start or end with apostrophe '
    ///   - cannot equal reserved "History" (case-insensitive)
    /// </summary>
    internal static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Invalid sheet name: name cannot be empty or whitespace.");
        if (name.Length > 31)
            throw new ArgumentException(
                $"Invalid sheet name '{name}': length {name.Length} exceeds Excel's 31-char limit.");
        var forbidden = new[] { '\\', '/', '?', '*', ':', '[', ']' };
        var hit = name.IndexOfAny(forbidden);
        if (hit >= 0)
            throw new ArgumentException(
                $"Invalid sheet name '{name}': contains forbidden character '{name[hit]}'. Excel rejects any of: \\ / ? * : [ ]");
        if (name.StartsWith('\'') || name.EndsWith('\''))
            throw new ArgumentException(
                $"Invalid sheet name '{name}': cannot start or end with an apostrophe (').");
        if (name.Equals("History", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(
                "Invalid sheet name 'History': reserved by Excel for the change-history sheet.");
    }

    /// <summary>
    /// R35-3: cross-workbook cell formulas like "=[Other.xlsx]Sheet1!A1" or
    /// "=[1]Sheet1!A1" need an externalLinks part to resolve. Without one,
    /// Excel opens the file but the formula shows #REF!. Reject up-front
    /// rather than silently persist a broken formula.
    /// CONSISTENCY(cross-workbook-ref): mirrors the namedrange refersTo
    /// guard in ExcelHandler.Add.Tables.cs (R27-1).
    /// </summary>
    internal static void RejectCrossWorkbookFormula(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return;
        var trimmed = formula.TrimStart('=', ' ', '\t');
        if (System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^\["))
            throw new ArgumentException(
                $"Cross-workbook references like '{formula}' require an externalLinks part which officecli doesn't expose; use raw-set for this case");
    }

    /// <summary>
    /// Build an XDR BlipFill with an optional asvg:svgBlip extension when
    /// the caller wires in an SVG image part. Keeps Add/Set picture paths
    /// free of inline extension boilerplate.
    /// </summary>
    private static XDR.BlipFill BuildPictureBlipFill(string pngRelId, string? svgRelId)
        => BuildPictureBlipFill(pngRelId, svgRelId, null);

    private static XDR.BlipFill BuildPictureBlipFill(
        string pngRelId, string? svgRelId, Dictionary<string, string>? properties)
    {
        var blip = new Drawing.Blip { Embed = pngRelId };
        // P6: opacity → <a:alphaModFix amt="N"/> (0..100000 scale).
        // Accept percent (50, "50%") or fraction (0.5). 100/100%/1.0 → opaque (no node).
        if (properties != null
            && properties.TryGetValue("opacity", out var opRaw)
            && !string.IsNullOrWhiteSpace(opRaw))
        {
            var amt = ParseOpacityAmt(opRaw);
            if (amt.HasValue && amt.Value < 100000)
                blip.AppendChild(new Drawing.AlphaModulationFixed { Amount = amt.Value });
        }
        if (!string.IsNullOrEmpty(svgRelId))
            OfficeCli.Core.SvgImageHelper.AppendSvgExtension(blip, svgRelId);
        var blipFill = new XDR.BlipFill(blip);
        // P7: crop.l/r/t/b or srcRect=l=..,r=..,t=..,b=.. → <a:srcRect .../>
        // Values are percent (10 → 10000 in 1/1000 pct units). Emitted before <a:stretch>.
        var srcRect = ParseSrcRect(properties);
        if (srcRect != null)
            blipFill.AppendChild(srcRect);
        blipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));
        return blipFill;
    }

    // Parse crop.l/r/t/b (percent, 10 → 10000) and compound srcRect="l=10,r=10,..."
    // alias. Returns null when no crop props are present.
    internal static Drawing.SourceRectangle? ParseSrcRect(Dictionary<string, string>? properties)
    {
        if (properties == null) return null;
        int? l = null, r = null, t = null, b = null;
        if (properties.TryGetValue("srcRect", out var compound) && !string.IsNullOrWhiteSpace(compound))
        {
            foreach (var piece in compound.Split(',', StringSplitOptions.RemoveEmptyEntries))
            {
                var kv = piece.Split('=', 2);
                if (kv.Length != 2) continue;
                var key = kv[0].Trim().ToLowerInvariant();
                var val = ParseCropPercent(kv[1]);
                if (!val.HasValue) continue;
                switch (key) { case "l": l = val; break; case "r": r = val; break; case "t": t = val; break; case "b": b = val; break; }
            }
        }
        foreach (var (key, fld) in new[] { ("crop.l", "l"), ("crop.r", "r"), ("crop.t", "t"), ("crop.b", "b") })
        {
            if (properties.TryGetValue(key, out var vs) && !string.IsNullOrWhiteSpace(vs))
            {
                var v = ParseCropPercent(vs);
                if (!v.HasValue) continue;
                switch (fld) { case "l": l = v; break; case "r": r = v; break; case "t": t = v; break; case "b": b = v; break; }
            }
        }
        // CONSISTENCY(picture-crop): Office-API-style `cropLeft`/`cropRight`
        // /`cropTop`/`cropBottom` aliases. Accept fraction (<=1 → *100%) or
        // percent (>1 → as-is); e.g. `cropLeft=0.1` and `cropLeft=10` both
        // mean 10% crop from left.
        foreach (var (key, fld) in new[] { ("cropLeft", "l"), ("cropRight", "r"), ("cropTop", "t"), ("cropBottom", "b") })
        {
            if (properties.TryGetValue(key, out var vs) && !string.IsNullOrWhiteSpace(vs))
            {
                var v = ParseCropFractionOrPercent(vs);
                if (!v.HasValue) continue;
                switch (fld) { case "l": l = v; break; case "r": r = v; break; case "t": t = v; break; case "b": b = v; break; }
            }
        }
        if (l == null && r == null && t == null && b == null) return null;
        var sr = new Drawing.SourceRectangle();
        if (l.HasValue) sr.Left = l.Value;
        if (r.HasValue) sr.Right = r.Value;
        if (t.HasValue) sr.Top = t.Value;
        if (b.HasValue) sr.Bottom = b.Value;
        return sr;
    }

    private static int? ParseCropPercent(string raw)
    {
        var t = raw.Trim();
        if (t.EndsWith("%")) t = t[..^1].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (double.IsNaN(v) || double.IsInfinity(v)) return null;
        return (int)Math.Round(v * 1000.0);
    }

    // CONSISTENCY(picture-crop): For `cropLeft`/`cropRight`/`cropTop`/
    // `cropBottom` keys we treat input ambiguously: <=1 is a fraction
    // (0.1 → 10%), >1 is percent (10 → 10%). Trailing `%` is still
    // honored explicitly. Returns 1/1000 pct units, same as OOXML.
    private static int? ParseCropFractionOrPercent(string raw)
    {
        var t = raw.Trim();
        bool explicitPct = t.EndsWith("%");
        if (explicitPct) t = t[..^1].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (double.IsNaN(v) || double.IsInfinity(v)) return null;
        double pct = (!explicitPct && v > 0 && v <= 1.0) ? v * 100.0 : v;
        return (int)Math.Round(pct * 1000.0);
    }

    // Parse opacity percent/fraction to OOXML alphaModFix amt scale (0..100000).
    // Returns null if the input is not parseable; 100000 (fully opaque) is returned
    // as-is so the caller can decide to omit the node.
    internal static int? ParseOpacityAmt(string raw)
    {
        var t = raw.Trim();
        if (t.EndsWith("%")) t = t[..^1].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (double.IsNaN(v) || double.IsInfinity(v)) return null;
        // Fraction form (0..1) → treat as 0..100%; else percent.
        double pct = v <= 1.0 && v > 0 ? v * 100.0 : v;
        if (pct < 0) pct = 0; if (pct > 100) pct = 100;
        return (int)Math.Round(pct * 1000.0);
    }

    // Build an <xdr:pic> element with an initial Transform2D, applying any
    // user-supplied rotation/flip props. Keeps the Add.cs path readable.
    // CONSISTENCY(scheme-color): Map a scheme-color name
    // ("accent1"-"accent6", "lt1"/"dk1", "lt2"/"dk2", "bg1"/"tx1", "bg2"/"tx2",
    // "hlink", "folHlink") to the OOXML theme index used by TabColor.Theme,
    // color.Theme on fonts, etc. Returns null for non-scheme inputs — callers
    // then fall back to srgbClr (hex) handling.
    internal static uint? ExcelSchemeColorNameToThemeIndex(string s) =>
        s?.Trim().ToLowerInvariant() switch
        {
            "lt1" or "bg1" or "light1" or "background1" => 0u,
            "dk1" or "tx1" or "dark1" or "text1" => 1u,
            "lt2" or "bg2" or "light2" or "background2" => 2u,
            "dk2" or "tx2" or "dark2" or "text2" => 3u,
            "accent1" => 4u,
            "accent2" => 5u,
            "accent3" => 6u,
            "accent4" => 7u,
            "accent5" => 8u,
            "accent6" => 9u,
            "hlink" or "hyperlink" => 10u,
            "folhlink" or "followedhyperlink" => 11u,
            _ => null
        };

    // CONSISTENCY(rc-units): Row height is in points in OOXML; this helper
    // accepts bare numbers (treated as points, backward compat) as well as
    // unit-qualified "40pt", "40px", "1cm", "0.5in" and returns points.
    internal static double ParseRowHeightPoints(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Row height cannot be empty.");
        var trimmed = value.Trim();
        double pts;
        // Bare number → points (legacy behavior)
        if (double.TryParse(trimmed, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var bare)
            && !char.IsLetter(trimmed[^1]))
        {
            if (double.IsNaN(bare) || double.IsInfinity(bare))
                throw new ArgumentException($"Invalid 'height' value: '{value}'. Expected a finite number (row height in points, e.g. 15.75).");
            pts = bare;
        }
        else
        {
            // Unit-qualified: convert via EMU then back to points.
            try
            {
                var emu = OfficeCli.Core.EmuConverter.ParseEmu(trimmed);
                pts = emu / 12700.0;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid 'height' value: '{value}'. Expected a finite number or unit-qualified value (e.g. 15.75, 40pt, 40px, 1cm, 0.5in).", ex);
            }
        }
        // DEFERRED(xlsx/row-height-validation) RC2: Excel row height is bounded
        // [0, 409.5] points. Values outside this range are rejected by Excel at
        // open time (file silently repaired), so validate at Set time.
        if (pts < 0 || pts > 409.5)
            throw new ArgumentException($"Invalid 'height' value: '{value}'. Row height must be between 0 and 409.5 points.");
        return pts;
    }

    // CONSISTENCY(rc-units): Column width is in "maximum digit width" char
    // units (Calibri 11pt ≈ 7px per char). Accepts bare number (char units,
    // legacy) or unit-qualified px/cm/in/pt — physical sizes converted via
    // the 7-px-per-char approximation Excel uses internally.
    internal static double ParseColWidthChars(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Column width cannot be empty.");
        var trimmed = value.Trim();
        double chars;
        if (double.TryParse(trimmed, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var bare)
            && !char.IsLetter(trimmed[^1]))
        {
            if (double.IsNaN(bare) || double.IsInfinity(bare))
                throw new ArgumentException($"Invalid 'width' value: '{value}'. Expected a finite number (column width in char units, e.g. 8.43).");
            chars = bare;
        }
        else
        {
            try
            {
                var emu = OfficeCli.Core.EmuConverter.ParseEmu(trimmed);
                // 9525 EMU = 1 px; 7 px ≈ 1 char unit (Calibri 11pt MDW baseline)
                var px = emu / 9525.0;
                chars = px / 7.0;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid 'width' value: '{value}'. Expected a finite number or unit-qualified value (e.g. 8.43, 20px, 2cm, 1in, 60pt).", ex);
            }
        }
        // DEFERRED(xlsx/row-height-validation) RC2: Excel column width is bounded
        // [0, 255] character units. Validate at Set time.
        if (chars < 0 || chars > 255)
            throw new ArgumentException($"Invalid 'width' value: '{value}'. Column width must be between 0 and 255 character units.");
        return chars;
    }

    internal static XDR.Picture BuildPictureElementWithTransform(
        uint picId, string alt, string imgRelId, string? svgRelId,
        Dictionary<string, string> properties)
    {
        var xfrm = new Drawing.Transform2D(
            new Drawing.Offset { X = 0, Y = 0 },
            new Drawing.Extents { Cx = 0, Cy = 0 });
        ApplyTransform2DRotationFlip(xfrm, properties);
        // P13: accept user-supplied `name=` to override the auto-generated
        // "Picture {id}" label stamped into xdr:cNvPr @name.
        // P9: `altText=` alias for `alt=` (Description attribute).
        // P11: `title=` populates the OOXML @title attribute (distinct from alt).
        var picName = properties.GetValueOrDefault("name");
        if (string.IsNullOrWhiteSpace(picName))
            picName = $"Picture {picId}";
        var picTitle = properties.GetValueOrDefault("title");
        var cNvPr = new XDR.NonVisualDrawingProperties { Id = picId, Name = picName, Description = alt };
        if (!string.IsNullOrWhiteSpace(picTitle))
            cNvPr.Title = picTitle;
        return new XDR.Picture(
            new XDR.NonVisualPictureProperties(
                cNvPr,
                new XDR.NonVisualPictureDrawingProperties(new Drawing.PictureLocks { NoChangeAspect = true })
            ),
            BuildPictureBlipFill(imgRelId, svgRelId, properties),
            new XDR.ShapeProperties(
                xfrm,
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
            )
        );
    }

    // Map a table-column totals-row function token to its OOXML enum and the
    // SUBTOTAL function code Excel uses. Unknown tokens fall back to SUM (109)
    // — previously all non-"sum" tokens silently became SUM; this keeps the
    // same fallback for unknown tokens but routes known ones to the right
    // enum + SUBTOTAL code.
    internal static (TotalsRowFunctionValues, int) MapTotalsRowFunction(string tok) => tok switch
    {
        "sum" => (TotalsRowFunctionValues.Sum, 109),
        "average" or "avg" => (TotalsRowFunctionValues.Average, 101),
        "count" => (TotalsRowFunctionValues.Count, 103),
        "countnums" or "countnumbers" => (TotalsRowFunctionValues.CountNumbers, 102),
        "max" or "maximum" => (TotalsRowFunctionValues.Maximum, 104),
        "min" or "minimum" => (TotalsRowFunctionValues.Minimum, 105),
        "stddev" or "stdev" => (TotalsRowFunctionValues.StandardDeviation, 107),
        "var" or "variance" => (TotalsRowFunctionValues.Variance, 110),
        "none" or "label" or "" => (TotalsRowFunctionValues.None, 0),
        "custom" => (TotalsRowFunctionValues.Custom, 109),
        _ => (TotalsRowFunctionValues.Sum, 109)
    };

    // Apply `rotation=<deg>` / `flip=h|v|both|hv|vh` from the user properties
    // dict to a Drawing.Transform2D node. Silently no-op on missing props.
    // Mirrors PowerPointHandler's shape rotation semantics: angles are in
    // degrees (positive = clockwise), OOXML stores them as 60000ths of a
    // degree in the `rot` attribute. Values are normalized modulo 360.
    internal static void ApplyTransform2DRotationFlip(
        Drawing.Transform2D xfrm, Dictionary<string, string> properties)
    {
        if (xfrm == null) return;
        if (properties.TryGetValue("rotation", out var rotStr) && !string.IsNullOrWhiteSpace(rotStr))
        {
            if (double.TryParse(rotStr, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var deg))
            {
                var normalized = ((deg % 360) + 360) % 360;
                xfrm.Rotation = (int)Math.Round(normalized * 60000);
            }
        }
        if (properties.TryGetValue("flip", out var flipStr) && !string.IsNullOrWhiteSpace(flipStr))
        {
            var f = flipStr.Trim().ToLowerInvariant();
            bool flipH = f == "h" || f == "horizontal" || f == "both" || f == "hv" || f == "vh";
            bool flipV = f == "v" || f == "vertical" || f == "both" || f == "hv" || f == "vh";
            if (flipH) xfrm.HorizontalFlip = true;
            if (flipV) xfrm.VerticalFlip = true;
        }
        // CONSISTENCY(shape-flip): accept Office-API-style `flipH=true`,
        // `flipV=true`, `flipBoth=true` aliases in addition to the compact
        // `flip=h|v|both`. Boolean semantics follow IsTruthy (true/1/yes).
        if (properties.TryGetValue("flipH", out var flipHStr) && IsTruthy(flipHStr))
            xfrm.HorizontalFlip = true;
        if (properties.TryGetValue("flipV", out var flipVStr) && IsTruthy(flipVStr))
            xfrm.VerticalFlip = true;
        if (properties.TryGetValue("flipBoth", out var flipBothStr) && IsTruthy(flipBothStr))
        {
            xfrm.HorizontalFlip = true;
            xfrm.VerticalFlip = true;
        }
    }

    // SH6 — build a two/three-stop linear gradient fill for shape/textbox from
    // a "C1-C2[-C3][:angle]" spec. Mirrors the chart gradient parser used by
    // Core/Chart/ChartHelper.Builder.cs:BuildFillElement so chart and shape
    // gradient syntax stay consistent.
    internal static Drawing.GradientFill BuildShapeGradientFill(string spec)
    {
        var colonIdx = spec.LastIndexOf(':');
        var anglePart = 0;
        string colorsPart;
        if (colonIdx > 6 && int.TryParse(spec[(colonIdx + 1)..],
            System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var ang))
        {
            anglePart = ang;
            colorsPart = spec[..colonIdx];
        }
        else
        {
            colorsPart = spec;
        }
        var colors = colorsPart.Split('-').Select(c => c.Trim()).Where(c => c.Length > 0).ToArray();
        if (colors.Length < 2)
            throw new ArgumentException(
                $"gradientFill requires at least two '-' separated colors; got '{spec}'.");
        var gradFill = new Drawing.GradientFill { RotateWithShape = true };
        var gsLst = new Drawing.GradientStopList();
        for (int i = 0; i < colors.Length; i++)
        {
            var pos = (int)(i * 100000.0 / (colors.Length - 1));
            var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(colors[i]);
            var gs = new Drawing.GradientStop { Position = pos };
            gs.AppendChild(new Drawing.RgbColorModelHex { Val = rgb });
            gsLst.AppendChild(gs);
        }
        gradFill.AppendChild(gsLst);
        gradFill.AppendChild(new Drawing.LinearGradientFill
        {
            Angle = anglePart * 60000,
            Scaled = true
        });
        return gradFill;
    }

    // Normalize user-supplied data-validation formula values so Excel accepts
    // them. `type=list` auto-quotes bare lists. `type=time` accepts HH:MM /
    // HH:MM:SS and converts to the Excel time serial fraction. `type=date`
    // accepts YYYY-MM-DD and converts to the Excel date serial. `type=custom`
    // strips a leading '=' since OOXML `<x:formula1>` expects the formula body
    // without one.
    internal static string NormalizeValidationFormula(string value, DataValidationValues? type)
    {
        if (string.IsNullOrEmpty(value)) return value;
        if (type == DataValidationValues.List)
        {
            // list: wrap bare "a,b,c" in quotes; leave cell/range refs and
            // already-quoted literals alone. V1: a leading `=` signals a
            // formula-ref (e.g. `=VOpts`, `=$Z$1:$Z$5`) — strip the `=`
            // (OOXML `<x:formula1>` expects the body without one) and
            // pass through without quoting.
            if (value.StartsWith("="))
                return value.Substring(1);
            if (value.StartsWith("\"") || value.Contains("!") || value.Contains(":"))
                return value;
            if (value.Contains(','))
                return $"\"{value}\"";
            return value;
        }
        if (type == DataValidationValues.Time)
        {
            var m = System.Text.RegularExpressions.Regex.Match(value.Trim(), @"^(\d{1,2}):(\d{2})(?::(\d{2}))?$");
            if (m.Success)
            {
                var h = int.Parse(m.Groups[1].Value);
                var mn = int.Parse(m.Groups[2].Value);
                var s = m.Groups[3].Success ? int.Parse(m.Groups[3].Value) : 0;
                var frac = (h * 3600 + mn * 60 + s) / 86400.0;
                return frac.ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        if (type == DataValidationValues.Date)
        {
            if (System.DateTime.TryParseExact(value.Trim(), "yyyy-MM-dd",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var dt))
            {
                // Excel date serial: days since 1899-12-30 (accounts for the
                // 1900 leap bug baseline).
                var epoch = new System.DateTime(1899, 12, 30);
                return ((int)(dt - epoch).TotalDays).ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        if (type == DataValidationValues.Custom)
        {
            if (value.StartsWith("="))
                return value.Substring(1);
        }
        return value;
    }

    // Returns true if `s` would parse as a valid cell reference (e.g. A1,
    // TBL1, XFD1048576). Excel refuses to open files whose table names match
    // this pattern — the name is ambiguous with a cell address.
    internal static bool LooksLikeCellReference(string? s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        var m = System.Text.RegularExpressions.Regex.Match(s, @"^\$?([A-Za-z]{1,3})\$?([0-9]+)$");
        if (!m.Success) return false;
        var col = m.Groups[1].Value.ToUpperInvariant();
        var colIdx = 0;
        foreach (var ch in col) colIdx = colIdx * 26 + (ch - 'A' + 1);
        if (colIdx < 1 || colIdx > 16384) return false;
        if (!long.TryParse(m.Groups[2].Value, out var row) || row < 1 || row > 1048576) return false;
        return true;
    }

    // R7-3: heuristic — is `s` a formula body (SUM(...), A1+B1, IF(...)),
    // as opposed to a pure range-ref body (Sheet1!$A$1:$A$5, A1:A5, A1)?
    // Used to decide whether to flip <calcPr fullCalcOnLoad="1"/> so Excel
    // evaluates the defined name on first open. Range-only bodies don't
    // need forced recalc; function calls and operator expressions do.
    internal static bool LooksLikeFormulaBody(string? s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        var t = s.Trim();
        if (t.Length == 0) return false;
        // A function call or arithmetic expression contains '(' or an
        // operator outside a sheet-qualified range.
        if (t.Contains('(')) return true;
        if (t.IndexOfAny(new[] { '+', '-', '*', '/', '^', '&', '<', '>', '=', '%' }) >= 0)
            return true;
        return false;
    }

    // Make a string safe to use as an Excel table name, displayName, or
    // tableColumn name. Excel refuses to open files where these identifiers
    // look like a cell reference ("tbl1" → column TBL row 1) or are purely
    // numeric ("30").
    //
    // When `userProvided` is true (user explicitly passed --prop name=T1),
    // honor the name verbatim — callers who type `name=T1` expect a table
    // named `T1`, not `T1_`. Excel itself accepts these table identifiers
    // (the cell-reference ambiguity rule applies to defined names, not to
    // tables), so silently rewriting loses fidelity with no gain.
    //
    // When `userProvided` is false (auto-derived default such as
    // `Table{id}`, or tableColumn name read from a header cell) we suffix
    // "_" on cell-reference-shaped names to keep defaults safe.
    internal static string SanitizeTableIdentifier(string? name, bool userProvided = false)
    {
        if (string.IsNullOrEmpty(name)) return "_";
        if (userProvided) return name;
        var looksLikeRef = LooksLikeCellReference(name)
            || System.Text.RegularExpressions.Regex.IsMatch(name, @"^[0-9]+$");
        return looksLikeRef ? name + "_" : name;
    }

    // ==================== Path Normalization ====================

    /// <summary>
    /// Normalize Excel-native path notation to DOM style.
    /// Sheet1!A1 → /Sheet1/A1
    /// Sheet1!A1:D10 → /Sheet1/A1:D10
    /// Sheet1!row[2] → /Sheet1/row[2]
    /// Sheet1!1:1 → /Sheet1/row[1]   (whole row)
    /// Sheet1!A:A → /Sheet1/col[A]   (whole column)
    /// Paths already starting with '/' are returned unchanged.
    /// </summary>
    internal static string NormalizeExcelPath(string path)
    {
        // Handle "/Sheet1!A1" — strip leading '/' when '!' is present so native notation is parsed correctly
        if (path.StartsWith('/') && path.Contains('!'))
            path = path[1..];
        if (path.Equals("/workbook", StringComparison.OrdinalIgnoreCase)) return "/";
        if (path.StartsWith('/')) return path;
        var bang = path.IndexOf('!');
        if (bang > 0)
        {
            var sheet = path[..bang];
            var selector = path[(bang + 1)..];

            // Whole-row notation: "1:1" or "3:3"
            var wholeRow = System.Text.RegularExpressions.Regex.Match(selector, @"^(\d+):\1$");
            if (wholeRow.Success)
                return $"/{sheet}/row[{wholeRow.Groups[1].Value}]";

            // Whole-column notation: "A:A" or "AB:AB"
            var wholeCol = System.Text.RegularExpressions.Regex.Match(selector, @"^([A-Za-z]+):\1$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (wholeCol.Success)
                return $"/{sheet}/col[{wholeCol.Groups[1].Value.ToUpperInvariant()}]";

            return $"/{sheet}/{selector}";
        }
        return path;
    }

    /// <summary>
    /// Resolve sheet[N] index references in the first segment of a normalized path.
    /// E.g. /sheet[1]/A1 → /Sheet1/A1 (if the first sheet is named "Sheet1").
    /// Must be called after NormalizeExcelPath.
    /// </summary>
    private string ResolveSheetIndexInPath(string path)
    {
        if (!path.StartsWith('/')) return path;
        var trimmed = path[1..]; // remove leading '/'
        var slashIdx = trimmed.IndexOf('/');
        var firstSegment = slashIdx >= 0 ? trimmed[..slashIdx] : trimmed;
        var resolved = ResolveSheetName(firstSegment);
        if (resolved == firstSegment) return path;
        return slashIdx >= 0 ? $"/{resolved}/{trimmed[(slashIdx + 1)..]}" : $"/{resolved}";
    }

    // ==================== Private Helpers ====================

    private static Worksheet GetSheet(WorksheetPart part) =>
        part.Worksheet ?? throw new InvalidOperationException("Corrupt file: worksheet data missing");

    /// <summary>
    /// Insert a ConditionalFormatting element after all existing CF elements (preserving add order).
    /// Falls back to after sheetData if no CF exists yet.
    /// </summary>
    private static void InsertConditionalFormatting(Worksheet ws, ConditionalFormatting cfElement)
    {
        var lastCf = ws.Elements<ConditionalFormatting>().LastOrDefault();
        if (lastCf != null)
            lastCf.InsertAfterSelf(cfElement);
        else
        {
            var sheetData = ws.GetFirstChild<SheetData>();
            if (sheetData != null)
                sheetData.InsertAfterSelf(cfElement);
            else
                ws.AppendChild(cfElement);
        }
    }

    /// <summary>
    /// Compute the next available CF priority for a worksheet (max existing + 1).
    /// </summary>
    private static int NextCfPriority(Worksheet ws)
    {
        int max = 0;
        foreach (var cf in ws.Elements<ConditionalFormatting>())
            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
                if (rule.Priority?.HasValue == true && rule.Priority.Value > max)
                    max = rule.Priority.Value;
        return max + 1;
    }

    // T6 — built-in Excel table style names. Unknown names are rejected at
    // Add time rather than silently passed through to Excel.
    private static readonly HashSet<string> _builtInTableStyles = BuildBuiltInTableStyles();
    private static HashSet<string> BuildBuiltInTableStyles()
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        foreach (var tier in new[] { "Light", "Medium", "Dark" })
            for (int i = 1; i <= 28; i++)
                set.Add($"TableStyle{tier}{i}");
        // Pivot styles — users may apply a pivot style to a plain table.
        foreach (var tier in new[] { "Light", "Medium", "Dark" })
            for (int i = 1; i <= 28; i++)
                set.Add($"PivotStyle{tier}{i}");
        set.Add("TableStyleNone");
        return set;
    }

    internal void ValidateTableStyleName(string? styleName)
    {
        if (string.IsNullOrEmpty(styleName)) return;
        if (_builtInTableStyles.Contains(styleName)) return;
        // Workbook-level customStyles live under <x:tableStyles> on the stylesheet.
        var styles = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet;
        var tableStyles = styles?.GetFirstChild<TableStyles>();
        if (tableStyles != null)
        {
            foreach (var ts in tableStyles.Elements<TableStyle>())
                if (ts.Name?.Value == styleName) return;
        }
        throw new ArgumentException(
            $"Unknown table style: '{styleName}'. Use a built-in name like " +
            $"'TableStyleMedium2', or register a custom style on the workbook first.");
    }

    /// <summary>
    /// CF2: stamp the stopIfTrue attribute onto a CF rule when the user
    /// passed `stopIfTrue=true`. Centralized so every `add cf` branch
    /// (databar / colorscale / iconset / formulacf / cellIs / topN / ...)
    /// honors the same flag.
    /// </summary>
    internal static void ApplyStopIfTrue(ConditionalFormattingRule rule, Dictionary<string, string> properties)
    {
        if (properties.TryGetValue("stopIfTrue", out var v) && ParseHelpers.IsTruthy(v))
            rule.StopIfTrue = true;
    }

    /// <summary>
    /// Ensure the worksheet root declares `xmlns:x14` + `mc:Ignorable="x14"`.
    /// Without both, Excel silently drops the x14 extension block where
    /// sparklines, dataBar 2010+ extensions, and other Office2010 features
    /// live. CONSISTENCY(x14-ignorable): same pattern the sparkline branch
    /// uses inline.
    /// </summary>
    internal static void EnsureWorksheetX14Ignorable(Worksheet ws)
    {
        const string mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        const string x14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
        if (ws.LookupNamespace("mc") == null)
            ws.AddNamespaceDeclaration("mc", mcNs);
        if (ws.LookupNamespace("x14") == null)
            ws.AddNamespaceDeclaration("x14", x14Ns);
        var ignorable = ws.MCAttributes?.Ignorable?.Value ?? "";
        if (!ignorable.Split(' ').Contains("x14"))
        {
            ws.MCAttributes ??= new MarkupCompatibilityAttributes();
            ws.MCAttributes.Ignorable = string.IsNullOrEmpty(ignorable) ? "x14" : $"{ignorable} x14";
        }
    }

    /// <summary>
    /// Append an x14:conditionalFormatting block to the worksheet's extLst under
    /// ext URI `{78C0D931-6437-407d-A8EE-F0AAD7539E65}`. Creates the extension
    /// on first call, appends to the existing x14:conditionalFormattings
    /// container on subsequent calls. Also ensures mc:Ignorable="x14" is set.
    /// </summary>
    internal static void EnsureWorksheetX14ConditionalFormatting(Worksheet ws, X14.ConditionalFormatting x14Cf)
    {
        const string cfExtUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
        const string x14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

        EnsureWorksheetX14Ignorable(ws);

        var extList = ws.GetFirstChild<WorksheetExtensionList>() ?? ws.AppendChild(new WorksheetExtensionList());
        var ext = extList.Elements<WorksheetExtension>().FirstOrDefault(e => e.Uri == cfExtUri);
        X14.ConditionalFormattings cfContainer;
        if (ext != null)
        {
            cfContainer = ext.GetFirstChild<X14.ConditionalFormattings>()
                ?? ext.AppendChild(new X14.ConditionalFormattings());
        }
        else
        {
            ext = new WorksheetExtension { Uri = cfExtUri };
            ext.AddNamespaceDeclaration("x14", x14Ns);
            cfContainer = new X14.ConditionalFormattings();
            ext.Append(cfContainer);
            extList.Append(ext);
        }
        cfContainer.Append(x14Cf);
    }

    /// <summary>
    /// Mark a worksheet as dirty. The actual save (with schema-order reorder) is
    /// deferred to <see cref="FlushDirtyParts"/> which runs in Dispose().
    /// This replaces per-mutation Save() calls — batch operations over many cells
    /// previously triggered one disk write per cell (O(n) saves); now they all
    /// flush in a single pass at the end.
    /// </summary>
    private void SaveWorksheet(WorksheetPart part)
    {
        _dirtyWorksheets.Add(part);
    }

    /// <summary>
    /// Flush all pending worksheet and stylesheet saves. Called from Dispose().
    /// Each dirty WorksheetPart is reordered and saved exactly once regardless
    /// of how many mutations targeted it.
    /// </summary>
    private void FlushDirtyParts()
    {
        foreach (var part in _dirtyWorksheets)
        {
            ReorderWorksheetChildren(GetSheet(part));
            GetSheet(part).Save();
        }
        _dirtyWorksheets.Clear();
        if (_dirtyStylesheet)
        {
            _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet?.Save();
            _dirtyStylesheet = false;
        }
    }

    /// <summary>
    /// Get a sparkline group by 1-based index from a worksheet's extension list.
    /// Returns null if not found.
    /// </summary>
    internal X14.SparklineGroup? GetSparklineGroup(WorksheetPart worksheet, int index)
    {
        var ws = GetSheet(worksheet);
        var extList = ws.GetFirstChild<WorksheetExtensionList>();
        if (extList == null) return null;

        var spkExt = extList.Elements<WorksheetExtension>()
            .FirstOrDefault(e => e.Uri == "{05C60535-1F16-4fd2-B633-E4A46CF9E463}");
        if (spkExt == null) return null;

        var spkGroups = spkExt.GetFirstChild<X14.SparklineGroups>();
        if (spkGroups == null) return null;

        var groups = spkGroups.Elements<X14.SparklineGroup>().ToList();
        if (index < 1 || index > groups.Count) return null;
        return groups[index - 1];
    }

    /// <summary>
    /// Build a DocumentNode for a sparkline group.
    /// </summary>
    internal static DocumentNode SparklineGroupToNode(string sheetName, X14.SparklineGroup spkGroup, int index)
    {
        var node = new DocumentNode
        {
            Path = $"/{sheetName}/sparkline[{index}]",
            Type = "sparkline"
        };

        // Type: default is line when attribute is absent
        string spkType;
        if (spkGroup.Type?.HasValue == true)
        {
            var tv = spkGroup.Type.Value;
            spkType = tv == X14.SparklineTypeValues.Column ? "column"
                : tv == X14.SparklineTypeValues.Stacked ? "stacked"
                : "line";
        }
        else
        {
            spkType = "line";
        }
        node.Format["type"] = spkType;

        // Color
        var colorRgb = spkGroup.SeriesColor?.Rgb?.Value;
        node.Format["color"] = colorRgb != null
            ? ParseHelpers.FormatHexColor(colorRgb)
            : "#4472C4";

        // Negative color
        var negColorRgb = spkGroup.NegativeColor?.Rgb?.Value;
        if (negColorRgb != null)
            node.Format["negativeColor"] = ParseHelpers.FormatHexColor(negColorRgb);

        // Boolean flags
        if (spkGroup.Markers?.Value == true) node.Format["markers"] = true;
        if (spkGroup.High?.Value == true) node.Format["highPoint"] = true;
        if (spkGroup.Low?.Value == true) node.Format["lowPoint"] = true;
        if (spkGroup.First?.Value == true) node.Format["firstPoint"] = true;
        if (spkGroup.Last?.Value == true) node.Format["lastPoint"] = true;
        if (spkGroup.Negative?.Value == true) node.Format["negative"] = true;

        // Line weight
        if (spkGroup.LineWeight?.HasValue == true)
            node.Format["lineWeight"] = spkGroup.LineWeight.Value;

        // Cell / range from first sparkline element
        var firstSparkline = spkGroup.GetFirstChild<X14.Sparklines>()?.GetFirstChild<X14.Sparkline>();
        if (firstSparkline != null)
        {
            // CONSISTENCY(canonical-key): schema canonical keys are 'location'
            // (target cell) and 'dataRange' (source range). 'cell'/'range' are
            // legacy aliases retained on input.
            var cell = firstSparkline.ReferenceSequence?.Text ?? "";
            node.Format["location"] = cell;

            // Strip sheet prefix from range (Sheet1!A1:E1 → A1:E1)
            var formulaText = firstSparkline.Formula?.Text ?? "";
            var excl = formulaText.IndexOf('!');
            node.Format["dataRange"] = excl >= 0 ? formulaText[(excl + 1)..] : formulaText;
        }

        return node;
    }

    /// <summary>
    /// Delete the calculation chain part if present.
    /// Excel will recalculate and recreate it on next open.
    /// This avoids stale calc chain references after cell/formula mutations.
    /// </summary>
    private void DeleteCalcChainIfPresent()
    {
        var calcChainPart = _doc.WorkbookPart?.CalculationChainPart;
        if (calcChainPart != null)
            _doc.WorkbookPart!.DeletePart(calcChainPart);
    }

    /// <summary>
    /// Reorder worksheet children to match OpenXML schema sequence.
    /// Schema: sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData,
    ///   autoFilter, sortState, mergeCells, conditionalFormatting,
    ///   dataValidations, hyperlinks, printOptions, pageMargins, pageSetup,
    ///   headerFooter, drawing, legacyDrawing, tableParts, extLst
    /// </summary>
    private static void ReorderWorksheetChildren(Worksheet ws)
    {
        var order = new Dictionary<string, int>
        {
            ["sheetPr"] = 0, ["dimension"] = 1, ["sheetViews"] = 2, ["sheetFormatPr"] = 3,
            ["cols"] = 4, ["sheetData"] = 5, ["sheetCalcPr"] = 6, ["sheetProtection"] = 7,
            ["protectedRanges"] = 8, ["scenarios"] = 9, ["autoFilter"] = 10, ["sortState"] = 11,
            ["dataConsolidate"] = 12, ["customSheetViews"] = 13, ["mergeCells"] = 14,
            ["phoneticPr"] = 15, ["conditionalFormatting"] = 16, ["dataValidations"] = 17,
            ["hyperlinks"] = 18, ["printOptions"] = 19, ["pageMargins"] = 20,
            ["pageSetup"] = 21, ["headerFooter"] = 22, ["rowBreaks"] = 23, ["colBreaks"] = 24,
            ["drawing"] = 25, ["legacyDrawing"] = 26, ["tableParts"] = 27, ["extLst"] = 99
        };

        var children = ws.ChildElements.ToList();
        var sorted = children
            .OrderBy(c => order.TryGetValue(c.LocalName, out var idx) ? idx : 50)
            .ToList();

        bool needsReorder = false;
        for (int i = 0; i < children.Count; i++)
        {
            if (!ReferenceEquals(children[i], sorted[i]))
            {
                needsReorder = true;
                break;
            }
        }

        if (needsReorder)
        {
            foreach (var child in children) child.Remove();
            foreach (var child in sorted) ws.AppendChild(child);
        }
    }

    private Workbook GetWorkbook() =>
        _doc.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Corrupt file: workbook missing");

    private List<(string Name, WorksheetPart Part)> GetWorksheets() => GetWorksheets(_doc);

    private static List<(string Name, WorksheetPart Part)> GetWorksheets(SpreadsheetDocument doc)
    {
        var result = new List<(string, WorksheetPart)>();
        var workbook = doc.WorkbookPart?.Workbook;
        if (workbook == null) return result;

        var sheets = workbook.GetFirstChild<Sheets>();
        if (sheets == null) return result;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var name = sheet.Name?.Value ?? "?";
            var id = sheet.Id?.Value;
            if (id == null) continue;
            var part = (WorksheetPart)doc.WorkbookPart!.GetPartById(id);
            result.Add((name, part));
        }

        return result;
    }

    private static readonly System.Text.RegularExpressions.Regex SheetIndexPattern =
        new(@"^sheet\[(\d+)\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

    /// <summary>
    /// Resolve a sheet name that may be a 1-based index reference like "sheet[1]"
    /// to the actual sheet name. Returns the original name if not an index pattern.
    /// </summary>
    private string ResolveSheetName(string sheetName)
    {
        var m = SheetIndexPattern.Match(sheetName);
        if (m.Success && int.TryParse(m.Groups[1].Value, out var idx) && idx >= 1)
        {
            var sheets = GetWorksheets();
            if (idx <= sheets.Count)
                return sheets[idx - 1].Name;
        }
        return sheetName;
    }

    private WorksheetPart? FindWorksheet(string sheetName)
    {
        sheetName = ResolveSheetName(sheetName);
        foreach (var (name, part) in GetWorksheets())
        {
            if (name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                return part;
        }
        return null;
    }

    private ArgumentException SheetNotFoundException(string sheetName)
    {
        var available = GetWorksheets().Select(w => w.Name).ToList();
        var availableStr = available.Count > 0
            ? string.Join(", ", available)
            : "(none)";
        return new ArgumentException(
            $"Sheet not found: \"{sheetName}\". Available sheets: [{availableStr}]. " +
            $"Use DOM path \"/{available.FirstOrDefault() ?? "SheetName"}/A1\" or Excel notation \"{available.FirstOrDefault() ?? "SheetName"}!A1\".");
    }

    private string GetCellDisplayValue(Cell cell, Core.FormulaEvaluator? evaluator = null)
    {
        if (cell.DataType?.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? "";
        }

        var value = cell.CellValue?.Text ?? "";

        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sst = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sst?.SharedStringTable != null && int.TryParse(value, out int idx))
            {
                var item = sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }

        // Formula cells: if there's a cached value, return it.
        // If not, try to evaluate; last resort: show the formula expression.
        if (string.IsNullOrEmpty(value) && cell.CellFormula?.Text != null)
        {
            if (evaluator != null)
            {
                var evalResult = evaluator.TryEvaluateFull(cell.CellFormula.Text);
                if (evalResult != null && !evalResult.IsError)
                    return evalResult.ToCellValueText();
            }
            return "=" + Core.ModernFunctionQualifier.Unqualify(cell.CellFormula.Text);
        }

        // Apply number format to numeric cells (dates, percentages, etc.)
        // Mirrors POI DataFormatter: raw double + format code → display string
        if (cell.DataType == null && double.TryParse(value,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numVal))
        {
            var (numFmtId, formatCode) = ExcelDataFormatter.GetCellFormat(cell, _doc.WorkbookPart);
            if (numFmtId > 0)
            {
                var formatted = ExcelDataFormatter.TryFormat(numVal, numFmtId, formatCode);
                if (formatted != null) return formatted;
            }
        }

        return value;
    }

    private List<DocumentNode> GetSheetChildNodes(string sheetName, SheetData sheetData, int depth, WorksheetPart? worksheetPart = null)
    {
        var children = new List<DocumentNode>();
        var eval = depth > 0 && worksheetPart != null ? new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart) : null;
        // R6-5: dedupe by RowIndex. When a sheet contains both source data
        // rows and pivot-rendered rows (possible when a pivot is placed on
        // its own source sheet), the renderer appends additional <row> nodes
        // that can collide with existing RowIndex values. Children should
        // expose each logical row once.
        var seenRowIndices = new HashSet<uint>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var ridx = row.RowIndex?.Value ?? 0;
            if (ridx != 0 && !seenRowIndices.Add(ridx))
                continue;
            var rowIdx = row.RowIndex?.Value ?? 0;
            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{rowIdx}]",
                Type = "row",
                ChildCount = row.Elements<Cell>().Count()
            };
            if (row.Height?.Value != null)
                rowNode.Format["height"] = row.Height.Value;
            if (row.Hidden?.Value == true)
                rowNode.Format["hidden"] = true;

            if (depth > 0)
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    rowNode.Children.Add(CellToNode(sheetName, cell, worksheetPart, eval));
                }
            }

            children.Add(rowNode);
        }

        // Add chart children from DrawingsPart (following Apache POI pattern)
        if (worksheetPart?.DrawingsPart != null)
        {
            var chartParts = worksheetPart.DrawingsPart.ChartParts.ToList();
            for (int i = 0; i < chartParts.Count; i++)
            {
                var chart = chartParts[i].ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                var chartNode = new DocumentNode
                {
                    Path = $"/{sheetName}/chart[{i + 1}]",
                    Type = "chart"
                };
                if (chart != null)
                    ChartHelper.ReadChartProperties(chart, chartNode, 0);
                children.Add(chartNode);
            }
        }

        // R16-1: expose pivottable children so Get /Sheet1 lists them.
        // CONSISTENCY(sheet-children): same pattern as chart children above.
        if (worksheetPart != null)
        {
            var pivotParts = worksheetPart.PivotTableParts.ToList();
            for (int i = 0; i < pivotParts.Count; i++)
            {
                var ptNode = new DocumentNode
                {
                    Path = $"/{sheetName}/pivottable[{i + 1}]",
                    Type = "pivottable"
                };
                var pivotDef = pivotParts[i].PivotTableDefinition;
                if (pivotDef != null)
                    Core.PivotTableHelper.ReadPivotTableProperties(pivotDef, ptNode, pivotParts[i]);
                children.Add(ptNode);
            }
        }

        return children;
    }

    private DocumentNode CellToNode(string sheetName, Cell cell, WorksheetPart? part = null, Core.FormulaEvaluator? evaluator = null)
    {
        var cellRef = cell.CellReference?.Value ?? "?";
        var formula = cell.CellFormula?.Text is { } fText
            ? Core.ModernFunctionQualifier.Unqualify(fText)
            : null;
        string type;
        if (cell.DataType?.HasValue != true)
        {
            // R12-F2: a formula whose cached value is a non-numeric string
            // should report type=String, not the Number default. Excel itself
            // writes t="str" on such cells; external tools or our own writer
            // occasionally leave the attribute off, so infer from the cached
            // value content.
            var raw = cell.CellValue?.Text;
            if (formula != null
                && !string.IsNullOrEmpty(raw)
                && !double.TryParse(raw, System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out _))
                type = "String";
            else
                type = "Number";
        }
        else if (cell.DataType.Value == CellValues.String)
            type = "String";
        else if (cell.DataType.Value == CellValues.SharedString)
            type = "SharedString";
        else if (cell.DataType.Value == CellValues.Boolean)
            type = "Boolean";
        else if (cell.DataType.Value == CellValues.Error)
            type = "Error";
        else if (cell.DataType.Value == CellValues.InlineString)
            type = "InlineString";
        else if (cell.DataType.Value == CellValues.Date)
            type = "Date";
        else
            type = "Number";

        // Lazy-create evaluator if not provided and needed
        if (evaluator == null && formula != null && string.IsNullOrEmpty(cell.CellValue?.Text) && part != null)
        {
            var sheetData = GetSheet(part).GetFirstChild<SheetData>();
            if (sheetData != null)
                evaluator = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        }

        var displayText = GetCellDisplayValue(cell, evaluator);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{cellRef}",
            Type = "cell",
            Text = displayText,
            Preview = cellRef
        };

        node.Format["type"] = type;
        if (formula != null)
        {
            node.Format["formula"] = formula;
            // cachedValue: prefer XML cached value, then evaluated value
            var rawCached = cell.CellValue?.Text;
            if (!string.IsNullOrEmpty(rawCached))
                node.Format["cachedValue"] = rawCached;
            else if (displayText != null && !displayText.StartsWith("=") &&
                     !FormulaReferencesMissingSheet(formula))
            {
                // R9-1: do NOT fall back to an evaluated cachedValue when the
                // formula references a sheet that no longer exists in the
                // workbook. Otherwise cross-sheet refs whose target sheet
                // was removed silently evaluate to "0" (see
                // FormulaEvaluator.ResolveSheetCellResult), reporting a
                // stale/fake cached value where Excel would show #REF!.
                node.Format["cachedValue"] = displayText;
            }
        }
        // Array formula readback — keys match Set input
        if (cell.CellFormula?.FormulaType?.Value == CellFormulaValues.Array)
        {
            node.Format["arrayformula"] = true;
            if (cell.CellFormula.Reference?.Value != null)
                node.Format["arrayref"] = cell.CellFormula.Reference.Value;
        }
        if (string.IsNullOrEmpty(displayText) && formula == null) node.Format["empty"] = true;

        // Hyperlink readback
        if (part != null)
        {
            var hyperlink = GetSheet(part).GetFirstChild<Hyperlinks>()?.Elements<Hyperlink>()
                .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
            if (hyperlink?.Id?.Value != null)
            {
                try
                {
                    var rel = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == hyperlink.Id.Value);
                    if (rel != null)
                    {
                        var linkStr = rel.Uri.OriginalString;
                        // Strip trailing slash added by Uri normalization for bare authority URLs
                        if (linkStr.EndsWith("/") && rel.Uri.IsAbsoluteUri && rel.Uri.AbsolutePath == "/")
                            linkStr = linkStr.TrimEnd('/');
                        node.Format["link"] = linkStr;
                    }
                }
                catch { }
            }
            // Internal-location hyperlinks (Sheet1!B5, defined names) have no
            // external relationship — they live entirely in the @location
            // attribute. Without this branch, internal links round-trip
            // through Set but vanish from Get.
            else if (hyperlink?.Location?.Value is { Length: > 0 } loc)
            {
                node.Format["link"] = loc;
            }

            // Border readback from stylesheet
            var styleIndex = cell.StyleIndex?.Value ?? 0;
            var wbStylesPart = _doc.WorkbookPart?.WorkbookStylesPart;
            if (wbStylesPart?.Stylesheet != null && styleIndex > 0)
            {
                var cellFormats = wbStylesPart.Stylesheet.CellFormats;
                if (cellFormats != null && styleIndex < (uint)cellFormats.Elements<CellFormat>().Count())
                {
                    var xf = cellFormats.Elements<CellFormat>().ElementAt((int)styleIndex);
                    // Font readback
                    var fontId = xf.FontId?.Value ?? 0;
                    if (fontId > 0)
                    {
                        var fonts = wbStylesPart.Stylesheet.Fonts;
                        if (fonts != null && fontId < (uint)fonts.Elements<Font>().Count())
                        {
                            var font = fonts.Elements<Font>().ElementAt((int)fontId);
                            if (font.Bold != null) { node.Format["font.bold"] = true; }
                            if (font.Italic != null)
                            {
                                node.Format["font.italic"] = true;
                            }
                            if (font.Strike != null) node.Format["font.strike"] = true;
                            if (font.Underline != null)
                                node.Format["font.underline"] = font.Underline.Val?.InnerText == "double" ? "double" : "single";
                            if (font.Color?.Rgb?.Value != null)
                                node.Format["font.color"] = ParseHelpers.FormatHexColor(font.Color.Rgb.Value);
                            else if (font.Color?.Theme?.Value != null)
                            {
                                var themeName = ParseHelpers.ExcelThemeIndexToName(font.Color.Theme.Value);
                                if (themeName != null) node.Format["font.color"] = themeName;
                            }
                            // vertAlign (superscript/subscript) readback — R28-A3:
                            // use font.subscript/font.superscript to match font.bold/font.italic.
                            var vertAlign = font.GetFirstChild<VerticalTextAlignment>();
                            if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Superscript)
                            {
                                node.Format["font.superscript"] = true;
                            }
                            else if (vertAlign?.Val?.Value == VerticalAlignmentRunValues.Subscript)
                            {
                                node.Format["font.subscript"] = true;
                            }
                            if (font.FontSize?.Val?.Value != null)
                                node.Format["font.size"] = $"{font.FontSize.Val.Value:0.##}pt";
                            if (font.FontName?.Val?.Value != null) node.Format["font.name"] = font.FontName.Val.Value;
                            // Long-tail Font children (charset, family, outline,
                            // shadow, condense, extend, scheme, ...). Emit as
                            // `font.<localName>` symmetric with the Set-side
                            // GetOrCreateFont longTailFontProps path.
                            FillUnknownDottedProps(font, node, "font.", CuratedFontChildren);
                        }
                    }

                    // Fill readback
                    var fillId = xf.FillId?.Value ?? 0;
                    if (fillId > 0)
                    {
                        var fills = wbStylesPart.Stylesheet.Fills;
                        if (fills != null && fillId < (uint)fills.Elements<Fill>().Count())
                        {
                            var fill = fills.Elements<Fill>().ElementAt((int)fillId);
                            // Check gradient fill first
                            var gf = fill.GetFirstChild<GradientFill>();
                            if (gf != null)
                            {
                                var stops = gf.Elements<GradientStop>().ToList();
                                if (stops.Count >= 2)
                                {
                                    var validColors = stops
                                        .Select(s => s.Color?.Rgb?.Value)
                                        .Where(v => !string.IsNullOrEmpty(v))
                                        .Select(v => ParseHelpers.FormatHexColor(v!))
                                        .ToList();
                                    if (validColors.Count >= 2)
                                    {
                                        var colorParts = string.Join(";", validColors);
                                        int deg = (int)(gf.Degree?.Value ?? 0);
                                        node.Format["fill"] = $"gradient;{colorParts};{deg}";
                                    }
                                }
                            }
                            else
                            {
                                var pf = fill.PatternFill;
                                if (pf?.ForegroundColor?.Rgb?.Value != null)
                                    node.Format["fill"] = ParseHelpers.FormatHexColor(pf.ForegroundColor.Rgb.Value);
                                else if (pf?.ForegroundColor?.Theme?.Value != null)
                                {
                                    var themeName = ParseHelpers.ExcelThemeIndexToName(pf.ForegroundColor.Theme.Value);
                                    if (themeName != null) node.Format["fill"] = themeName;
                                }
                            }
                        }
                    }

                    var borderId = xf.BorderId?.Value ?? 0;
                    if (borderId > 0)
                    {
                        var borders = wbStylesPart.Stylesheet.Borders;
                        if (borders != null && borderId < (uint)borders.Elements<Border>().Count())
                        {
                            var border = borders.Elements<Border>().ElementAt((int)borderId);
                            var sides = new (string name, BorderPropertiesType? bp)[] {
                                ("left", border.LeftBorder), ("right", border.RightBorder),
                                ("top", border.TopBorder), ("bottom", border.BottomBorder)
                            };
                            foreach (var (side, b) in sides)
                            {
                                if (b?.Style?.Value != null && b.Style.Value != BorderStyleValues.None)
                                {
                                    node.Format[$"border.{side}"] = b.Style.InnerText;
                                    if (b.Color?.Rgb?.Value != null)
                                        node.Format[$"border.{side}.color"] = ParseHelpers.FormatHexColor(b.Color.Rgb.Value!);
                                }
                            }
                            // Diagonal border readback
                            var diag = border.DiagonalBorder;
                            if (diag?.Style?.Value != null && diag.Style.Value != BorderStyleValues.None)
                            {
                                node.Format["border.diagonal"] = diag.Style.InnerText;
                                if (diag.Color?.Rgb?.Value != null)
                                    node.Format["border.diagonal.color"] = ParseHelpers.FormatHexColor(diag.Color.Rgb.Value!);
                            }
                            if (border.DiagonalUp?.Value == true)
                                node.Format["border.diagonalUp"] = true;
                            if (border.DiagonalDown?.Value == true)
                                node.Format["border.diagonalDown"] = true;
                        }
                    }

                    // Alignment + wrap readback (like POI XSSFCellStyle.getWrapText)
                    var alignment = xf.Alignment;
                    if (alignment != null)
                    {
                        if (alignment.WrapText?.Value == true)
                            node.Format["alignment.wrapText"] = true;
                        if (alignment.Horizontal?.HasValue == true)
                            node.Format["alignment.horizontal"] = alignment.Horizontal.InnerText;
                        if (alignment.Vertical?.HasValue == true)
                        {
                            node.Format["alignment.vertical"] = alignment.Vertical.InnerText;
                        }
                        if (alignment.TextRotation?.HasValue == true && alignment.TextRotation.Value != 0)
                            node.Format["alignment.textRotation"] = alignment.TextRotation.Value.ToString();
                        if (alignment.Indent?.HasValue == true && alignment.Indent.Value > 0)
                            node.Format["alignment.indent"] = alignment.Indent.Value.ToString();
                        if (alignment.ShrinkToFit?.Value == true)
                            node.Format["alignment.shrinkToFit"] = true;
                        // DEFERRED(xlsx/cell-reading-order) CE10 — canonical
                        // readback as string form (context/ltr/rtl).
                        if (alignment.ReadingOrder?.HasValue == true && alignment.ReadingOrder.Value != 0)
                        {
                            node.Format["alignment.readingOrder"] = alignment.ReadingOrder.Value switch
                            {
                                1u => "ltr",
                                2u => "rtl",
                                _ => "context"
                            };
                        }
                        // Long-tail Alignment attributes (justifyLastLine,
                        // relativeIndent, ...). Symmetric with Set's default
                        // branch in ExcelStyleManager.ApplyStyle alignment loop.
                        FillUnknownAttrProps(alignment, node, "alignment.", CuratedAlignmentAttrs);
                    }

                    // Protection readback — both curated locked/hidden and any
                    // long-tail Protection attribute symmetric with Set.
                    var xfProt = xf.Protection;
                    if (xfProt != null)
                    {
                        if (xfProt.Locked?.Value != null) node.Format["protection.locked"] = xfProt.Locked.Value;
                        if (xfProt.Hidden?.Value != null) node.Format["protection.hidden"] = xfProt.Hidden.Value;
                        FillUnknownAttrProps(xfProt, node, "protection.", CuratedProtectionAttrs);
                    }

                    // R29: quotePrefix readback (set by leading apostrophe text mode)
                    if (xf.QuotePrefix?.Value == true)
                        node.Format["quotePrefix"] = true;

                    // Number format readback
                    var numFmtId = xf.NumberFormatId?.Value ?? 0;
                    if (numFmtId > 0)
                    {
                        node.Format["numFmtId"] = (int)numFmtId;
                        var numFmts = wbStylesPart.Stylesheet.NumberingFormats;
                        var customFmt = numFmts?.Elements<NumberingFormat>()
                            .FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
                        object fmtVal;
                        if (customFmt?.FormatCode?.Value != null)
                            fmtVal = customFmt.FormatCode.Value;
                        else
                        {
                            // Resolve built-in number format IDs to their format strings
                            // See ECMA-376 Part 1, 18.8.30 (numFmt) for built-in IDs
                            fmtVal = numFmtId switch
                            {
                                1 => "0",
                                2 => "0.00",
                                3 => "#,##0",
                                4 => "#,##0.00",
                                9 => "0%",
                                10 => "0.00%",
                                11 => "0.00E+00",
                                12 => "# ?/?",
                                13 => "# ??/??",
                                14 => "m/d/yy",
                                15 => "d-mmm-yy",
                                16 => "d-mmm",
                                17 => "mmm-yy",
                                18 => "h:mm AM/PM",
                                19 => "h:mm:ss AM/PM",
                                20 => "h:mm",
                                21 => "h:mm:ss",
                                22 => "m/d/yy h:mm",
                                37 => "#,##0 ;(#,##0)",
                                38 => "#,##0 ;[Red](#,##0)",
                                39 => "#,##0.00;(#,##0.00)",
                                40 => "#,##0.00;[Red](#,##0.00)",
                                45 => "mm:ss",
                                46 => "[h]:mm:ss",
                                47 => "mmss.0",
                                48 => "##0.0E+0",
                                49 => "@",
                                _ => (object)(int)numFmtId // fallback to ID for truly unknown formats
                            };
                        }
                        node.Format["numberformat"] = fmtVal;
                    }

                    // Protection readback handled above via the dotted
                    // canonical form (`protection.locked` / `protection.hidden`)
                    // — see CONSISTENCY(canonical-keys) in CLAUDE.md. Flat
                    // `locked` / `formulahidden` Get emission was removed to
                    // avoid double-emission alongside the dotted form. The
                    // Set side still accepts both flat shorthand and dotted
                    // input via IsStyleKey routing.
                }
            }

            // Merge cell readback
            var mergeCells = GetSheet(part).GetFirstChild<MergeCells>();
            if (mergeCells != null)
            {
                var mergeCell = mergeCells.Elements<MergeCell>()
                    .FirstOrDefault(m => IsCellInMergeRange(cellRef, m.Reference?.Value));
                if (mergeCell != null)
                {
                    var mergeRef = mergeCell.Reference?.Value ?? "";
                    node.Format["merge"] = mergeRef;
                    // Indicate if this cell is the top-left anchor of the merged range
                    if (mergeRef.Split(':')[0].Equals(cellRef, StringComparison.OrdinalIgnoreCase))
                        node.Format["mergeAnchor"] = true;
                }
            }
        }

        // Rich text (SST runs) readback
        if (cell.DataType?.Value == CellValues.SharedString &&
            int.TryParse(cell.CellValue?.Text, out var sstIdx2))
        {
            var sst2 = _doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            var ssi2 = sst2?.SharedStringTable?.Elements<SharedStringItem>().ElementAtOrDefault(sstIdx2);
            if (ssi2 != null)
            {
                var runs = ssi2.Elements<Run>().ToList();
                if (runs.Count > 0)
                {
                    node.Format["richtext"] = true;
                    node.ChildCount = runs.Count;
                    int runI = 1;
                    foreach (var run in runs)
                    {
                        node.Children.Add(RunToNode(run, $"/{sheetName}/{cellRef}/run[{runI}]"));
                        runI++;
                    }
                }
            }
        }

        return node;
    }

    private static DocumentNode RunToNode(Run run, string path)
    {
        var runNode = new DocumentNode { Path = path, Type = "run", Text = run.Text?.Text ?? "" };
        var rp = run.RunProperties;
        if (rp != null)
        {
            if (rp.GetFirstChild<Bold>() != null) runNode.Format["bold"] = true;
            if (rp.GetFirstChild<Italic>() != null) runNode.Format["italic"] = true;
            if (rp.GetFirstChild<Strike>() != null) runNode.Format["strike"] = true;
            var ul = rp.GetFirstChild<Underline>();
            if (ul != null) runNode.Format["underline"] = ul.Val?.InnerText == "double" ? "double" : "single";
            var va = rp.GetFirstChild<VerticalTextAlignment>();
            if (va?.Val?.Value == VerticalAlignmentRunValues.Superscript) runNode.Format["superscript"] = true;
            if (va?.Val?.Value == VerticalAlignmentRunValues.Subscript) runNode.Format["subscript"] = true;
            if (rp.GetFirstChild<FontSize>()?.Val?.Value != null)
                runNode.Format["size"] = $"{rp.GetFirstChild<FontSize>()!.Val!.Value:0.##}pt";
            if (rp.GetFirstChild<Color>()?.Rgb?.Value != null)
                runNode.Format["color"] = ParseHelpers.FormatHexColor(rp.GetFirstChild<Color>()!.Rgb!.Value!);
            if (rp.GetFirstChild<RunFont>()?.Val?.Value != null)
                runNode.Format["font"] = rp.GetFirstChild<RunFont>()!.Val!.Value!;
        }
        return runNode;
    }

    private static bool IsCellInMergeRange(string cellRef, string? rangeRef)
    {
        if (string.IsNullOrEmpty(rangeRef) || !rangeRef.Contains(':')) return false;
        var parts = rangeRef.Split(':');
        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var (cellCol, cellRow) = ParseCellReference(cellRef);

        var cellColIdx = ColumnNameToIndex(cellCol);
        return cellRow >= startRow && cellRow <= endRow
            && cellColIdx >= ColumnNameToIndex(startCol) && cellColIdx <= ColumnNameToIndex(endCol);
    }

    // T4 — rectangle intersection over A1:B2 style ranges (case-insensitive).
    // Returns true if two inclusive cell ranges share at least one cell.
    private static bool RangesOverlap(string rangeA, string rangeB)
    {
        if (string.IsNullOrEmpty(rangeA) || string.IsNullOrEmpty(rangeB)) return false;
        var (a1, a2) = SplitRange(rangeA);
        var (b1, b2) = SplitRange(rangeB);
        var (aSc, aSr) = ParseCellReference(a1);
        var (aEc, aEr) = ParseCellReference(a2);
        var (bSc, bSr) = ParseCellReference(b1);
        var (bEc, bEr) = ParseCellReference(b2);
        int aSci = ColumnNameToIndex(aSc), aEci = ColumnNameToIndex(aEc);
        int bSci = ColumnNameToIndex(bSc), bEci = ColumnNameToIndex(bEc);
        // Normalize (callers may pass B2:A1 theoretically)
        if (aSci > aEci) (aSci, aEci) = (aEci, aSci);
        if (bSci > bEci) (bSci, bEci) = (bEci, bSci);
        if (aSr > aEr) (aSr, aEr) = (aEr, aSr);
        if (bSr > bEr) (bSr, bEr) = (bEr, bSr);
        return aSci <= bEci && bSci <= aEci && aSr <= bEr && bSr <= aEr;
    }

    private static (string, string) SplitRange(string range)
    {
        if (!range.Contains(':')) return (range, range);
        var p = range.Split(':');
        return (p[0], p[1]);
    }

    private DocumentNode GetCellRange(string sheetName, SheetData sheetData, string range, int depth, WorksheetPart? part = null)
    {
        var parts = range.Split(':');
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range: {range}");

        var (startCol, startRow) = ParseCellReference(parts[0]);
        var (endCol, endRow) = ParseCellReference(parts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/{range}",
            Type = "range",
            Preview = range
        };

        // Build lookup of existing cells so we can fill empty stubs for missing positions
        var existingCells = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null)
                    existingCells[cell.CellReference.Value] = cell;
            }
        }

        // Enumerate every position in the range in row-major order,
        // materializing empty stubs for positions that have no cell element.
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        for (int r = startRow; r <= endRow; r++)
        {
            for (int c = startColIdx; c <= endColIdx; c++)
            {
                var cellRef = $"{IndexToColumnName(c)}{r}";
                if (existingCells.TryGetValue(cellRef, out var existingCell))
                    node.Children.Add(CellToNode(sheetName, existingCell, part, eval));
                else
                    node.Children.Add(new DocumentNode
                    {
                        Path = $"/{sheetName}/{cellRef}",
                        Type = "cell",
                        Text = "",
                        Preview = cellRef,
                        Format = { ["type"] = "Number", ["empty"] = true }
                    });
            }
        }

        node.ChildCount = node.Children.Count;
        return node;
    }

    /// <summary>
    /// Parse a cell value for sorting: returns a tuple (rank, numVal, strVal) so that
    /// nulls/empties sort last, numbers sort before strings, and cross-type comparison never occurs.
    /// rank=0 for numbers, rank=1 for strings, rank=2 for empty/null.
    /// </summary>
    private static (int Rank, double NumVal, string StrVal) ParseSortValue(string value)
    {
        if (string.IsNullOrEmpty(value)) return (2, 0.0, "");
        // Excel treats NaN / Infinity / -Infinity as text, not numbers. double.TryParse
        // happily accepts them though, which would make sort order dependent on whether
        // the exact casing matched double.TryParse's spec vs not — classify explicitly.
        if (value.Equals("NaN", StringComparison.Ordinal)
            || value.Equals("Infinity", StringComparison.Ordinal)
            || value.Equals("-Infinity", StringComparison.Ordinal)
            || value.Equals("+Infinity", StringComparison.Ordinal))
            return (1, 0.0, value);
        if (double.TryParse(value, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var num))
        {
            // Defensive: even non-literal inputs can produce non-finite doubles
            // (e.g. "1e999" overflows to +Infinity). Keep those in the string bucket.
            if (!double.IsFinite(num)) return (1, 0.0, value);
            return (0, num, "");
        }
        return (1, 0.0, value);
    }

    private static Cell? FindCell(SheetData sheetData, string cellRef)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
                    return cell;
            }
        }
        return null;
    }

    /// <summary>
    /// Find or create the Row for the given 1-based row index, using the per-SheetData
    /// row index cache to avoid O(n) linear scans. New rows are inserted in sorted order
    /// via binary search on the cache (O(log n)).
    /// </summary>
    private Row FindOrCreateRow(SheetData sheetData, uint rowIdx)
    {
        _rowIndex ??= new();
        if (!_rowIndex.TryGetValue(sheetData, out var rowMap))
        {
            rowMap = new SortedList<uint, Row>();
            foreach (var existingRow in sheetData.Elements<Row>())
                if (existingRow.RowIndex?.HasValue == true)
                    rowMap[existingRow.RowIndex.Value] = existingRow;
            _rowIndex[sheetData] = rowMap;
        }

        if (rowMap.TryGetValue(rowIdx, out var row))
            return row;

        row = new Row { RowIndex = rowIdx };
        // Binary search for predecessor in O(log n)
        var keys = rowMap.Keys;
        int lo = 0, hi = keys.Count - 1, predPos = -1;
        while (lo <= hi)
        {
            int mid = (lo + hi) / 2;
            if (keys[mid] < rowIdx) { predPos = mid; lo = mid + 1; }
            else hi = mid - 1;
        }
        if (predPos >= 0)
            rowMap.Values[predPos].InsertAfterSelf(row);
        else
            sheetData.InsertAt(row, 0);
        rowMap[rowIdx] = row;
        return row;
    }

    /// <summary>
    /// Invalidate the row index cache for a specific SheetData (or all sheets if null).
    /// Must be called whenever rows are structurally modified (removed, shifted).
    /// </summary>
    private void InvalidateRowIndex(SheetData? sheetData = null)
    {
        if (sheetData != null)
            _rowIndex?.Remove(sheetData);
        else
            _rowIndex = null;
    }

    private Cell FindOrCreateCell(SheetData sheetData, string cellRef)
    {
        var (colName, rowIdx) = ParseCellReference(cellRef);

        var row = FindOrCreateRow(sheetData, (uint)rowIdx);

        // Cell lookup within row — O(m) where m = cols per row (typically small)
        var cell = row.Elements<Cell>().FirstOrDefault(c =>
            c.CellReference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
        if (cell == null)
        {
            cell = new Cell { CellReference = cellRef.ToUpperInvariant() };
            // Insert in column order
            var afterCell = row.Elements<Cell>().LastOrDefault(c =>
            {
                var (cn, _) = ParseCellReference(c.CellReference?.Value ?? "A1");
                return ColumnNameToIndex(cn) < ColumnNameToIndex(colName);
            });
            if (afterCell != null)
                afterCell.InsertAfterSelf(cell);
            else
                row.InsertAt(cell, 0);
        }

        return cell;
    }

    // ==================== Conditional Formatting Helpers ====================

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    // CONSISTENCY(xlsx/comment-font): C8 — build the <x:rPr> for comment runs.
    // When no font.* properties are supplied, keep the legacy Tahoma 9 /
    // indexed-81 default for back-compat. When any font.* is present, honor
    // them and fall back to the defaults only for unspecified facets.
    // Input vocabulary mirrors the cell-level font handling: font.bold,
    // font.italic, font.underline (single|double), font.size (pt-qualified
    // or bare), font.color (#FF0000 / FF0000 / rgb() / named), font.name.
    internal static RunProperties BuildCommentRunProperties(Dictionary<string, string> properties)
    {
        bool hasAnyFont = properties.Keys.Any(k =>
            k.StartsWith("font.", StringComparison.OrdinalIgnoreCase));
        if (!hasAnyFont)
        {
            return new RunProperties(
                new FontSize { Val = 9 },
                new Color { Indexed = 81 },
                new RunFont { Val = "Tahoma" });
        }

        var rPr = new RunProperties();
        if (properties.TryGetValue("font.bold", out var fb) && IsTruthy(fb))
            rPr.AppendChild(new Bold());
        if (properties.TryGetValue("font.italic", out var fi) && IsTruthy(fi))
            rPr.AppendChild(new Italic());
        if (properties.TryGetValue("font.underline", out var fu) && !string.IsNullOrEmpty(fu)
            && !string.Equals(fu, "none", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(fu, "false", StringComparison.OrdinalIgnoreCase))
        {
            var uVal = string.Equals(fu, "double", StringComparison.OrdinalIgnoreCase)
                ? UnderlineValues.Double : UnderlineValues.Single;
            rPr.AppendChild(new Underline { Val = uVal });
        }
        // Size default 9pt
        var sizePt = properties.TryGetValue("font.size", out var fs)
            ? ParseHelpers.ParseFontSize(fs) : 9.0;
        rPr.AppendChild(new FontSize { Val = sizePt });
        // Color: explicit overrides default indexed=81
        if (properties.TryGetValue("font.color", out var fc) && !string.IsNullOrWhiteSpace(fc))
            rPr.AppendChild(new Color { Rgb = ParseHelpers.NormalizeArgbColor(fc) });
        else
            rPr.AppendChild(new Color { Indexed = 81 });
        // Name default Tahoma
        var fontName = properties.TryGetValue("font.name", out var fn) && !string.IsNullOrWhiteSpace(fn)
            ? fn : "Tahoma";
        rPr.AppendChild(new RunFont { Val = fontName });
        return rPr;
    }

    private static bool IsValidBooleanString(string? value) =>
        ParseHelpers.IsValidBooleanString(value);

    private static IconSetValues ParseIconSetValues(string name)
    {
        return name.ToLowerInvariant() switch
        {
            "3arrows" => IconSetValues.ThreeArrows,
            "3arrowsgray" => IconSetValues.ThreeArrowsGray,
            "3flags" => IconSetValues.ThreeFlags,
            "3trafficlights1" => IconSetValues.ThreeTrafficLights1,
            "3trafficlights2" => IconSetValues.ThreeTrafficLights2,
            "3signs" => IconSetValues.ThreeSigns,
            "3symbols" => IconSetValues.ThreeSymbols,
            "3symbols2" => IconSetValues.ThreeSymbols2,
            "4arrows" => IconSetValues.FourArrows,
            "4arrowsgray" => IconSetValues.FourArrowsGray,
            "4rating" => IconSetValues.FourRating,
            "4redtoblack" => IconSetValues.FourRedToBlack,
            "4trafficlights" => IconSetValues.FourTrafficLights,
            "5arrows" => IconSetValues.FiveArrows,
            "5arrowsgray" => IconSetValues.FiveArrowsGray,
            "5rating" => IconSetValues.FiveRating,
            "5quarters" => IconSetValues.FiveQuarters,
            _ => throw new ArgumentException($"Unknown icon set name: '{name}'. Valid names: 3Arrows, 3ArrowsGray, 3Flags, 3TrafficLights1, 3TrafficLights2, 3Signs, 3Symbols, 3Symbols2, 4Arrows, 4ArrowsGray, 4Rating, 4RedToBlack, 4TrafficLights, 5Arrows, 5ArrowsGray, 5Rating, 5Quarters")
        };
    }

    private static int GetIconCount(string name)
    {
        var lower = name.ToLowerInvariant();
        if (lower.StartsWith("5")) return 5;
        if (lower.StartsWith("4")) return 4;
        return 3;
    }

    // ==================== Data Validation Helpers ====================

    private DocumentNode TableToNode(string sheetName, WorksheetPart worksheetPart, int tableIndex, int depth)
    {
        var tableParts = worksheetPart.TableDefinitionParts.ToList();
        if (tableIndex < 1 || tableIndex > tableParts.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range (1..{tableParts.Count})");

        var tbl = tableParts[tableIndex - 1].Table
            ?? throw new ArgumentException($"Table {tableIndex} has no definition");

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/table[{tableIndex}]",
            Type = "table",
            Text = tbl.DisplayName?.Value ?? tbl.Name?.Value ?? $"Table{tableIndex}",
            Preview = $"{tbl.Name?.Value} ({tbl.Reference?.Value})"
        };

        node.Format["name"] = tbl.Name?.Value ?? "";
        node.Format["displayName"] = tbl.DisplayName?.Value ?? "";
        node.Format["ref"] = tbl.Reference?.Value ?? "";

        var styleInfo = tbl.GetFirstChild<TableStyleInfo>();
        if (styleInfo?.Name?.Value != null)
            node.Format["style"] = styleInfo.Name.Value;
        if (styleInfo != null)
        {
            if (styleInfo.ShowRowStripes is not null) node.Format["showRowStripes"] = styleInfo.ShowRowStripes.Value;
            if (styleInfo.ShowColumnStripes is not null) node.Format["showColumnStripes"] = styleInfo.ShowColumnStripes.Value;
            if (styleInfo.ShowFirstColumn is not null) node.Format["showFirstColumn"] = styleInfo.ShowFirstColumn.Value;
            if (styleInfo.ShowLastColumn is not null) node.Format["showLastColumn"] = styleInfo.ShowLastColumn.Value;
        }

        node.Format["headerRow"] = (tbl.HeaderRowCount?.Value ?? 1) != 0;
        node.Format["totalRow"] = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);

        var tableColumns = tbl.GetFirstChild<TableColumns>();
        if (tableColumns != null)
        {
            var colNames = tableColumns.Elements<TableColumn>()
                .Select(c => c.Name?.Value ?? "").ToArray();
            node.Format["columns"] = string.Join(",", colNames);
            node.ChildCount = colNames.Length;
        }

        return node;
    }

    private DocumentNode CommentToNode(string sheetName, Comment comment, Comments comments, int index)
    {
        var reference = comment.Reference?.Value ?? "?";
        var text = comment.CommentText?.InnerText ?? "";
        var authorId = comment.AuthorId?.Value ?? 0;

        var authors = comments.GetFirstChild<Authors>();
        var authorName = authors?.Elements<Author>().ElementAtOrDefault((int)authorId)?.Text ?? "Unknown";

        var node = new DocumentNode
        {
            Path = $"/{sheetName}/comment[{index}]",
            Type = "comment",
            Text = text,
            Preview = $"{reference}: {text}"
        };

        node.Format["ref"] = reference;
        node.Format["author"] = authorName;
        node.Format["anchoredTo"] = $"/{sheetName}/{reference}";

        // CONSISTENCY(xlsx/comment-font): C8 — surface font.* from first run's
        // rPr so Query/Get round-trips the Add-time formatting. Only report
        // non-default facets so Tahoma-9-indexed-81 comments stay unadorned.
        var firstRun = comment.CommentText?.Elements<Run>().FirstOrDefault();
        var rProps = firstRun?.RunProperties;
        if (rProps != null)
        {
            if (rProps.Elements<Bold>().Any()) node.Format["font.bold"] = true;
            if (rProps.Elements<Italic>().Any()) node.Format["font.italic"] = true;
            var u = rProps.Elements<Underline>().FirstOrDefault();
            if (u != null)
                node.Format["font.underline"] = u.Val?.InnerText == "double" ? "double" : "single";
            var clr = rProps.Elements<Color>().FirstOrDefault();
            if (clr?.Rgb?.HasValue == true)
                node.Format["font.color"] = ParseHelpers.FormatHexColor(clr.Rgb.Value!);
            var sz = rProps.Elements<FontSize>().FirstOrDefault();
            if (sz?.Val?.HasValue == true && sz.Val.Value != 9.0)
                node.Format["font.size"] = $"{sz.Val.Value:0.##}pt";
            var rf = rProps.Elements<RunFont>().FirstOrDefault();
            if (rf?.Val?.HasValue == true && rf.Val.Value != "Tahoma")
                node.Format["font.name"] = rf.Val.Value;
        }

        return node;
    }

    private static DocumentNode DataValidationToNode(string sheetName, DataValidation dv, int index)
    {
        var sqref = dv.SequenceOfReferences?.InnerText ?? "";
        var node = new DocumentNode
        {
            Path = $"/{sheetName}/validation[{index}]",
            Type = "validation",
            Text = sqref,
            Preview = $"validation[{index}] ({sqref})"
        };

        // CONSISTENCY(canonical-key): schema canonical key is 'ref', not 'sqref'.
        node.Format["ref"] = sqref;

        if (dv.Type?.HasValue == true)
            node.Format["type"] = dv.Type.InnerText;
        if (dv.Operator?.HasValue == true)
            node.Format["operator"] = dv.Operator.InnerText;

        if (dv.Formula1 != null)
        {
            // Preserve formula1 exactly as stored in XML so query→set round-trips:
            // list-type validations wrap literal options in "..." at Add time, and
            // stripping those quotes here made set(formula1=<stripped>) treat the
            // whole list as a single item. See DEFERRED(xlsx/validation-list-formula-roundtrip).
            node.Format["formula1"] = dv.Formula1.Text ?? "";
        }

        if (dv.Formula2 != null)
            node.Format["formula2"] = dv.Formula2.Text ?? "";

        if (dv.AllowBlank?.HasValue == true)
            node.Format["allowBlank"] = dv.AllowBlank.Value;
        if (dv.ShowErrorMessage?.HasValue == true)
            node.Format["showError"] = dv.ShowErrorMessage.Value;
        if (dv.ShowInputMessage?.HasValue == true)
            node.Format["showInput"] = dv.ShowInputMessage.Value;

        if (!string.IsNullOrEmpty(dv.ErrorTitle?.Value))
            node.Format["errorTitle"] = dv.ErrorTitle!.Value!;
        if (!string.IsNullOrEmpty(dv.Error?.Value))
            node.Format["error"] = dv.Error!.Value!;
        if (!string.IsNullOrEmpty(dv.PromptTitle?.Value))
            node.Format["promptTitle"] = dv.PromptTitle!.Value!;
        if (!string.IsNullOrEmpty(dv.Prompt?.Value))
            node.Format["prompt"] = dv.Prompt!.Value!;

        return node;
    }

    // ==================== Picture Helpers ====================

    private DocumentNode? GetPictureNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null) return null;

        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return null;

        var picAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Picture>().Any())
            .ToList();

        if (index < 1 || index > picAnchors.Count)
            return null;

        var anchor = picAnchors[index - 1];
        var picture = anchor.Descendants<XDR.Picture>().First();

        var node = new DocumentNode { Path = path, Type = "picture" };

        var nvProps = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        if (nvProps != null)
        {
            if (!string.IsNullOrEmpty(nvProps.Description?.Value))
            {
                node.Format["alt"] = nvProps.Description.Value;
                node.Text = nvProps.Description.Value;
            }
            if (!string.IsNullOrEmpty(nvProps.Name?.Value))
                node.Format["name"] = nvProps.Name.Value;
        }

        ReadAnchorPosition(anchor, node);

        // Rotation / flip readback from <xdr:spPr><a:xfrm rot=".." flipH=".." flipV="..">
        // CONSISTENCY(shape-flip): same canonical form as GetShapeNode.
        var picSpPr = picture.ShapeProperties;
        var xfrm = picSpPr?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Rotation?.HasValue == true && xfrm.Rotation.Value != 0)
            {
                var deg = xfrm.Rotation.Value / 60000.0;
                node.Format["rotation"] = Math.Round(deg, 2);
            }
            if (xfrm.HorizontalFlip?.HasValue == true && xfrm.VerticalFlip?.HasValue == true
                && xfrm.HorizontalFlip.Value && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "both";
            else if (xfrm.HorizontalFlip?.HasValue == true && xfrm.HorizontalFlip.Value)
                node.Format["flip"] = "h";
            else if (xfrm.VerticalFlip?.HasValue == true && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "v";
        }

        // CONSISTENCY(picture-crop): mirror PowerPointHandler.NodeBuilder.cs
        // crop readback. <a:srcRect l/t/r/b> stores values in 1000ths of a
        // percent (10000 = 10%); emit as comma-separated percent string.
        var picSrcRect = picture.BlipFill?.GetFirstChild<Drawing.SourceRectangle>();
        if (picSrcRect != null)
        {
            var cl = picSrcRect.Left?.Value ?? 0;
            var ct = picSrcRect.Top?.Value ?? 0;
            var cr = picSrcRect.Right?.Value ?? 0;
            var cb = picSrcRect.Bottom?.Value ?? 0;
            if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
                node.Format["crop"] = $"{cl / 1000.0:0.##},{ct / 1000.0:0.##},{cr / 1000.0:0.##},{cb / 1000.0:0.##}";
        }

        return node;
    }

    private DocumentNode? GetShapeNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null) return null;
        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return null;

        var shpAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();

        if (index < 1 || index > shpAnchors.Count)
            return null;

        var anchor = shpAnchors[index - 1];
        var shape = anchor.Descendants<XDR.Shape>().First();

        var node = new DocumentNode { Path = path, Type = "shape" };

        // Name
        var nvProps = shape.NonVisualShapeProperties?.GetFirstChild<XDR.NonVisualDrawingProperties>();
        if (nvProps?.Name?.Value != null)
            node.Format["name"] = nvProps.Name.Value;

        // Text — shape TextBody has one <a:p> per paragraph, each with
        // zero-or-more <a:r>/<a:t> runs. Concatenate runs within a
        // paragraph, then join paragraphs with '\n' so multi-line shape
        // text round-trips through Get.
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList();
        if (paragraphs != null && paragraphs.Count > 0)
        {
            node.Text = string.Join("\n", paragraphs.Select(p =>
                string.Join("", p.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""))));
        }
        var textRuns = shape.TextBody?.Descendants<Drawing.Run>().ToList();

        // Position/size
        ReadAnchorPosition(anchor, node);

        // Font properties from first run
        var firstRun = textRuns?.FirstOrDefault();
        var rPr = firstRun?.RunProperties;
        if (rPr != null)
        {
            if (rPr.FontSize?.HasValue == true)
                node.Format["size"] = $"{rPr.FontSize.Value / 100.0}pt";
            if (rPr.Bold?.HasValue == true && rPr.Bold.Value)
                node.Format["bold"] = true;
            if (rPr.Italic?.HasValue == true && rPr.Italic.Value)
                node.Format["italic"] = true;

            var solidFill = rPr.GetFirstChild<Drawing.SolidFill>();
            var colorHex = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (colorHex?.Val?.Value != null)
                node.Format["color"] = ParseHelpers.FormatHexColor(colorHex.Val.Value);
            else
            {
                var schemeClr = solidFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                if (schemeClr?.HasValue == true) node.Format["color"] = schemeClr.InnerText;
            }

            var latin = rPr.GetFirstChild<Drawing.LatinFont>();
            if (latin?.Typeface?.Value != null)
                node.Format["font"] = latin.Typeface.Value;
        }

        // Rotation / flip readback from <a:xfrm rot="..." flipH="..." flipV="...">
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Rotation?.HasValue == true && xfrm.Rotation.Value != 0)
            {
                // OOXML stores rotation in 60000ths of a degree; Add normalizes
                // into [0,360). Round-trip the same canonical form.
                var deg = xfrm.Rotation.Value / 60000.0;
                node.Format["rotation"] = Math.Round(deg, 2);
            }
            if (xfrm.HorizontalFlip?.HasValue == true && xfrm.VerticalFlip?.HasValue == true
                && xfrm.HorizontalFlip.Value && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "both";
            else if (xfrm.HorizontalFlip?.HasValue == true && xfrm.HorizontalFlip.Value)
                node.Format["flip"] = "h";
            else if (xfrm.VerticalFlip?.HasValue == true && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "v";
        }

        // Geometry preset (rect, ellipse, etc.) — `preset` is the canonical
        // key per shape help schema; `shape`/`geometry` are accepted as
        // Add/Set aliases. Match PPTX shape readback which also uses `preset`.
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
            node.Format["preset"] = presetGeom.Preset.InnerText;

        // Fill
        var spPr = shape.ShapeProperties;
        if (spPr?.GetFirstChild<Drawing.NoFill>() != null)
            node.Format["fill"] = "none";
        else
        {
            var shapeFill = spPr?.GetFirstChild<Drawing.SolidFill>();
            var fillColor = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (fillColor?.Val?.Value != null)
                node.Format["fill"] = ParseHelpers.FormatHexColor(fillColor.Val.Value);
            else
            {
                var schemeClr = shapeFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                if (schemeClr?.HasValue == true) node.Format["fill"] = schemeClr.InnerText;
            }
        }

        // Paragraph alignment — read first paragraph's a:pPr/@algn (mirrors
        // Set which writes to every paragraph). PPTX shape Get uses `align`
        // canonical key.
        var firstPara = shape.TextBody?.GetFirstChild<Drawing.Paragraph>();
        var firstPPr = firstPara?.ParagraphProperties;
        if (firstPPr?.Alignment?.HasValue == true)
        {
            // SDK v3 enum values are not compile-time constants; switch on InnerText.
            node.Format["align"] = firstPPr.Alignment.InnerText switch
            {
                "ctr" => "center",
                "r" => "right",
                "just" => "justify",
                "l" => "left",
                var s => s,
            };
        }

        // Vertical alignment — bodyPr/@anchor.
        var bodyPrForAnchor = shape.TextBody?.GetFirstChild<Drawing.BodyProperties>();
        if (bodyPrForAnchor?.Anchor?.HasValue == true)
        {
            node.Format["valign"] = bodyPrForAnchor.Anchor.InnerText switch
            {
                "t" => "top",
                "ctr" => "center",
                "b" => "bottom",
                var s => s,
            };
        }

        // Outline (line/border). Set writes "none" or "color[:width[:style]]".
        // Round-trip emits the same canonical form.
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            if (outline.GetFirstChild<Drawing.NoFill>() != null)
                node.Format["line"] = "none";
            else
            {
                var lineFill = outline.GetFirstChild<Drawing.SolidFill>();
                var lineRgb = lineFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                string? colorPart = null;
                if (lineRgb != null)
                    colorPart = ParseHelpers.FormatHexColor(lineRgb);
                else
                {
                    var schemeClr = lineFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                    if (schemeClr?.HasValue == true) colorPart = schemeClr.InnerText;
                }
                if (colorPart != null)
                {
                    var widthPt = outline.Width?.HasValue == true
                        ? $":{outline.Width.Value / 12700.0:0.##}"
                        : "";
                    node.Format["line"] = colorPart + widthPt;
                }
            }
        }

        // Margin (text body insets) — Add/Set accept points and write all four
        // sides uniformly; mirror that as a single points readback when all
        // four match. Stored as EMU on BodyProperties, 12700 EMU per point.
        var bodyPr = shape.TextBody?.GetFirstChild<Drawing.BodyProperties>();
        if (bodyPr != null)
        {
            var lIns = bodyPr.LeftInset?.Value;
            var rIns = bodyPr.RightInset?.Value;
            var tIns = bodyPr.TopInset?.Value;
            var bIns = bodyPr.BottomInset?.Value;
            if (lIns.HasValue || rIns.HasValue || tIns.HasValue || bIns.HasValue)
            {
                if (lIns == rIns && rIns == tIns && tIns == bIns && lIns.HasValue)
                    node.Format["margin"] = $"{lIns.Value / 12700.0:0.##}pt";
                else
                    node.Format["margin"] = $"{(lIns ?? 0) / 12700.0:0.##}pt,{(tIns ?? 0) / 12700.0:0.##}pt,{(rIns ?? 0) / 12700.0:0.##}pt,{(bIns ?? 0) / 12700.0:0.##}pt";
            }
        }

        // Effects — check shape-level then text-level
        var effectList = spPr?.GetFirstChild<Drawing.EffectList>();
        var textEffectList = (effectList == null || !effectList.HasChildren)
            ? rPr?.GetFirstChild<Drawing.EffectList>()
            : null;
        var activeEffects = effectList?.HasChildren == true ? effectList : textEffectList;
        if (activeEffects != null)
        {
            var shadow = activeEffects.GetFirstChild<Drawing.OuterShadow>();
            if (shadow != null)
            {
                var sColor = ParseHelpers.FormatHexColor(shadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000");
                node.Format["shadow"] = sColor;
            }
            var glow = activeEffects.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var gColor = ParseHelpers.FormatHexColor(glow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000");
                var gRadius = glow.Radius?.HasValue == true ? $"{glow.Radius.Value / 12700.0:0.##}" : "8";
                node.Format["glow"] = $"{gColor}-{gRadius}";
            }
        }

        return node;
    }

    // ==================== Shared Anchor Helpers ====================

    /// <summary>
    /// Set position/size properties (x, y, width, height) on a TwoCellAnchor.
    /// Returns true if the key was handled, false otherwise.
    /// </summary>
    private static bool TrySetAnchorPosition(XDR.TwoCellAnchor anchor, string key, string value)
    {
        switch (key)
        {
            case "x":
                if (anchor.FromMarker != null)
                {
                    // CONSISTENCY(ole-width-units): mirror Add — accept bare
                    // cell index OR unit-qualified offset ("2cm", "1in", "72pt").
                    var xVal = ParseAnchorOrigin(value, "x");
                    anchor.FromMarker.ColumnId!.Text = xVal.ToString();
                }
                return true;
            case "y":
                if (anchor.FromMarker != null)
                {
                    // CONSISTENCY(ole-width-units): see x case above.
                    var yVal = ParseAnchorOrigin(value, "y");
                    anchor.FromMarker.RowId!.Text = yVal.ToString();
                }
                return true;
            case "width":
                if (anchor.FromMarker != null && anchor.ToMarker != null)
                {
                    // CONSISTENCY(ole-width-units): mirror Add path's
                    // ParseAnchorDimension — accept bare integer cell spans
                    // OR unit-qualified strings ("6cm", "2in", "72pt").
                    var fromCol = int.TryParse(anchor.FromMarker.ColumnId?.Text, out var fc) ? fc : 0;
                    anchor.ToMarker.ColumnId!.Text = (fromCol + ParseAnchorDimension(value, "width")).ToString();
                }
                return true;
            case "height":
                if (anchor.FromMarker != null && anchor.ToMarker != null)
                {
                    // CONSISTENCY(ole-width-units): see width case above.
                    var fromRow = int.TryParse(anchor.FromMarker.RowId?.Text, out var fr) ? fr : 0;
                    anchor.ToMarker.RowId!.Text = (fromRow + ParseAnchorDimension(value, "height")).ToString();
                }
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Read position/size from a TwoCellAnchor into a DocumentNode's Format dictionary.
    /// </summary>
    private static void ReadAnchorPosition(XDR.TwoCellAnchor anchor, DocumentNode node)
    {
        var from = anchor.FromMarker;
        var to = anchor.ToMarker;
        if (from != null)
        {
            node.Format["x"] = from.ColumnId?.Text ?? "0";
            node.Format["y"] = from.RowId?.Text ?? "0";
        }
        if (to != null && from != null)
        {
            var fromCol = int.TryParse(from.ColumnId?.Text, out var fc) ? fc : 0;
            var toCol = int.TryParse(to.ColumnId?.Text, out var tc) ? tc : 0;
            var fromRow = int.TryParse(from.RowId?.Text, out var fr) ? fr : 0;
            var toRow = int.TryParse(to.RowId?.Text, out var tr2) ? tr2 : 0;
            node.Format["width"] = (toCol - fromCol).ToString();
            node.Format["height"] = (toRow - fromRow).ToString();
        }
    }

    /// <summary>
    /// Set rotation on a ShapeProperties element.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetRotation(XDR.ShapeProperties? spPr, string key, string value)
    {
        if (key is not ("rotation" or "rot")) return false;
        if (spPr == null) return true;

        var xfrm = spPr.GetFirstChild<Drawing.Transform2D>();
        if (xfrm == null)
        {
            xfrm = new Drawing.Transform2D(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 }
            );
            spPr.InsertAt(xfrm, 0);
        }
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var degrees))
            throw new ArgumentException($"Invalid 'rotation' value: '{value}'. Expected a number in degrees (e.g. 45, -90, 180.5).");
        xfrm.Rotation = (int)(degrees * 60000);
        return true;
    }

    /// <summary>
    /// Set horizontal / vertical flip on a shape's Transform2D. Accepts "h", "v", "both",
    /// or "none" to clear both. Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeFlip(XDR.ShapeProperties? spPr, string key, string value)
    {
        // Accept the compact `flip=h|v|both|hv|vh|none|false` form plus the
        // Office-API aliases `flipH=true`, `flipV=true`, `flipHorizontal=true`,
        // `flipVertical=true`, `flipBoth=true`. CONSISTENCY(shape-flip) — mirrors
        // ApplyTransform2DRotationFlip used on the Add path.
        if (key is not ("flip" or "fliph" or "flipv" or "fliphorizontal" or "flipvertical" or "flipboth"))
            return false;
        if (spPr == null) return true;
        var xfrm = spPr.GetFirstChild<Drawing.Transform2D>();
        if (xfrm == null)
        {
            xfrm = new Drawing.Transform2D(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 });
            spPr.InsertAt(xfrm, 0);
        }

        if (key == "flip")
        {
            var f = value.Trim().ToLowerInvariant();
            bool none = f is "none" or "false" or "";
            bool flipH = !none && (f is "h" or "horizontal" or "both" or "hv" or "vh");
            bool flipV = !none && (f is "v" or "vertical" or "both" or "hv" or "vh");
            xfrm.HorizontalFlip = flipH ? true : (bool?)null;
            xfrm.VerticalFlip = flipV ? true : (bool?)null;
            return true;
        }

        bool truthy = IsTruthy(value);
        if (key is "fliph" or "fliphorizontal")
            xfrm.HorizontalFlip = truthy ? true : (bool?)null;
        else if (key is "flipv" or "flipvertical")
            xfrm.VerticalFlip = truthy ? true : (bool?)null;
        else if (key == "flipboth")
        {
            xfrm.HorizontalFlip = truthy ? true : (bool?)null;
            xfrm.VerticalFlip = truthy ? true : (bool?)null;
        }
        return true;
    }

    /// <summary>
    /// Apply a dotted-form font property (`font.bold`, `font.italic`, `font.color`,
    /// `font.size`, `font.name`, `font.underline`) to every run in the shape's text body.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeFontProp(XDR.Shape shape, string key, string value)
    {
        if (!key.StartsWith("font.", StringComparison.Ordinal)) return false;
        var sub = key.Substring(5);
        foreach (var run in shape.Descendants<Drawing.Run>())
        {
            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
            switch (sub)
            {
                case "bold":
                    rPr.Bold = IsTruthy(value);
                    break;
                case "italic":
                    rPr.Italic = IsTruthy(value);
                    break;
                case "size":
                    rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(value) * 100);
                    break;
                case "name":
                    rPr.RemoveAllChildren<Drawing.LatinFont>();
                    rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                    rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                    rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                    break;
                case "color":
                {
                    rPr.RemoveAllChildren<Drawing.SolidFill>();
                    var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                    OfficeCli.Core.DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = cRgb }));
                    break;
                }
                case "underline":
                {
                    var uv = value.ToLowerInvariant();
                    rPr.Underline = uv switch
                    {
                        "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                        "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                        "none" or "false" => Drawing.TextUnderlineValues.None,
                        _ => Drawing.TextUnderlineValues.Single
                    };
                    break;
                }
                default:
                    return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Apply shape-level effects (shadow, glow, reflection, softedge) on a ShapeProperties element.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeEffect(XDR.ShapeProperties? spPr, string key, string value)
    {
        if (key is not ("shadow" or "glow" or "reflection" or "softedge")) return false;
        if (spPr == null) return true;

        var effectList = spPr.GetFirstChild<Drawing.EffectList>();
        var normalizedVal = value.Replace(':', '-');
        if (normalizedVal == "true") normalizedVal = key == "shadow" ? "000000" : key == "glow" ? "4472C4" : "half";

        if (normalizedVal.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            normalizedVal.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (effectList != null)
            {
                switch (key)
                {
                    case "shadow": effectList.RemoveAllChildren<Drawing.OuterShadow>(); break;
                    case "glow": effectList.RemoveAllChildren<Drawing.Glow>(); break;
                    case "reflection": effectList.RemoveAllChildren<Drawing.Reflection>(); break;
                    case "softedge": effectList.RemoveAllChildren<Drawing.SoftEdge>(); break;
                }
                if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            }
        }
        else
        {
            if (effectList == null) { effectList = new Drawing.EffectList(); spPr.AppendChild(effectList); }
            // CONSISTENCY(effect-list-schema-order): CT_EffectList order is
            // blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge.
            // Excel (and PPT) silently drops out-of-order children, so we must
            // InsertBefore the next-in-order sibling rather than AppendChild.
            OpenXmlElement newEffect;
            switch (key)
            {
                case "shadow":
                    effectList.RemoveAllChildren<Drawing.OuterShadow>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
                    break;
                case "glow":
                    effectList.RemoveAllChildren<Drawing.Glow>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
                    break;
                case "reflection":
                    effectList.RemoveAllChildren<Drawing.Reflection>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildReflection(normalizedVal);
                    break;
                case "softedge":
                    effectList.RemoveAllChildren<Drawing.SoftEdge>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildSoftEdge(normalizedVal);
                    break;
                default: return true;
            }
            InsertEffectInSchemaOrder(effectList, newEffect);
        }
        return true;
    }

    /// <summary>
    /// Insert an effectLst child at the correct DrawingML CT_EffectList schema position:
    /// blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge.
    /// </summary>
    private static void InsertEffectInSchemaOrder(Drawing.EffectList effectList, OpenXmlElement newEffect)
    {
        // Determine all types that must come AFTER newEffect per schema order.
        OpenXmlElement? insertBefore = newEffect switch
        {
            Drawing.Blur => (OpenXmlElement?)effectList.GetFirstChild<Drawing.FillOverlay>()
                ?? effectList.GetFirstChild<Drawing.Glow>()
                ?? effectList.GetFirstChild<Drawing.InnerShadow>()
                ?? effectList.GetFirstChild<Drawing.OuterShadow>()
                ?? effectList.GetFirstChild<Drawing.PresetShadow>()
                ?? (OpenXmlElement?)effectList.GetFirstChild<Drawing.Reflection>()
                ?? effectList.GetFirstChild<Drawing.SoftEdge>(),
            Drawing.FillOverlay => (OpenXmlElement?)effectList.GetFirstChild<Drawing.Glow>()
                ?? effectList.GetFirstChild<Drawing.InnerShadow>()
                ?? effectList.GetFirstChild<Drawing.OuterShadow>()
                ?? effectList.GetFirstChild<Drawing.PresetShadow>()
                ?? (OpenXmlElement?)effectList.GetFirstChild<Drawing.Reflection>()
                ?? effectList.GetFirstChild<Drawing.SoftEdge>(),
            Drawing.Glow => (OpenXmlElement?)effectList.GetFirstChild<Drawing.InnerShadow>()
                ?? effectList.GetFirstChild<Drawing.OuterShadow>()
                ?? effectList.GetFirstChild<Drawing.PresetShadow>()
                ?? (OpenXmlElement?)effectList.GetFirstChild<Drawing.Reflection>()
                ?? effectList.GetFirstChild<Drawing.SoftEdge>(),
            Drawing.InnerShadow => (OpenXmlElement?)effectList.GetFirstChild<Drawing.OuterShadow>()
                ?? effectList.GetFirstChild<Drawing.PresetShadow>()
                ?? (OpenXmlElement?)effectList.GetFirstChild<Drawing.Reflection>()
                ?? effectList.GetFirstChild<Drawing.SoftEdge>(),
            Drawing.OuterShadow => (OpenXmlElement?)effectList.GetFirstChild<Drawing.PresetShadow>()
                ?? (OpenXmlElement?)effectList.GetFirstChild<Drawing.Reflection>()
                ?? effectList.GetFirstChild<Drawing.SoftEdge>(),
            Drawing.PresetShadow => (OpenXmlElement?)effectList.GetFirstChild<Drawing.Reflection>()
                ?? effectList.GetFirstChild<Drawing.SoftEdge>(),
            Drawing.Reflection => (OpenXmlElement?)effectList.GetFirstChild<Drawing.SoftEdge>(),
            _ => null,
        };
        if (insertBefore != null) effectList.InsertBefore(newEffect, insertBefore);
        else effectList.AppendChild(newEffect);
    }

    /// <summary>
    /// Parse x, y, width, height from properties with given defaults. Used by both picture Add and shape Add.
    /// </summary>
    // CONSISTENCY(shape-preset): mirror PowerPointHandler.ParsePresetShape token
    // set so Excel `add shape preset=X` accepts the same vocabulary as PPT.
    //
    // Exhaustive map covering every OOXML preset token. Built once via
    // reflection over `Drawing.ShapeTypeValues` static properties — each
    // property's default `ToString()` (== OpenXml IEnumValue.Value) is the
    // OOXML token such as "smileyFace", "flowChartProcess", "lightningBolt".
    // We then overlay friendly aliases (oval, cylinder, rarrow, …).
    private static readonly Dictionary<string, Drawing.ShapeTypeValues> _shapePresetMap =
        BuildShapePresetMap();

    private static Dictionary<string, Drawing.ShapeTypeValues> BuildShapePresetMap()
    {
        var map = new Dictionary<string, Drawing.ShapeTypeValues>(StringComparer.Ordinal);
        foreach (var p in typeof(Drawing.ShapeTypeValues)
            .GetProperties(BindingFlags.Public | BindingFlags.Static)
            .Where(p => p.PropertyType == typeof(Drawing.ShapeTypeValues)))
        {
            if (p.GetValue(null) is not Drawing.ShapeTypeValues val) continue;
            // IEnumValue.Value is the OOXML token, e.g. "smileyFace". Do not
            // use ToString() — on OpenXml SDK 3.x record-struct wrappers it
            // returns "ShapeTypeValues { }" instead of the token.
            var token = (val as IEnumValue)?.Value;
            if (string.IsNullOrEmpty(token)) continue;
            map[token.ToLowerInvariant()] = val;
        }

        // Friendly aliases layered on top (key must be lowercase).
        void Alias(string alias, Drawing.ShapeTypeValues v) => map[alias] = v;
        Alias("rectangle", Drawing.ShapeTypeValues.Rectangle);
        Alias("roundedrectangle", Drawing.ShapeTypeValues.RoundRectangle);
        Alias("oval", Drawing.ShapeTypeValues.Ellipse);
        Alias("righttriangle", Drawing.ShapeTypeValues.RightTriangle);
        Alias("rtriangle", Drawing.ShapeTypeValues.RightTriangle);
        Alias("rarrow", Drawing.ShapeTypeValues.RightArrow);
        Alias("larrow", Drawing.ShapeTypeValues.LeftArrow);
        Alias("cross", Drawing.ShapeTypeValues.Plus);
        Alias("cylinder", Drawing.ShapeTypeValues.Can);
        return map;
    }

    /// <summary>
    /// Parse shape margin into 4 EMU insets (left, top, right, bottom).
    /// Accepts unit-qualified "14pt"/"0.5cm"/"0.2in"/bare-points for uniform
    /// inset, OR a 4-CSV "Lpt,Tpt,Rpt,Bpt" matching Get's readback format.
    /// CONSISTENCY(spacing-units): mirrors SpacingConverter usage so that
    /// margin's input vocabulary matches Get's "Npt"/"L,T,R,B" output.
    /// </summary>
    private static (int L, int T, int R, int B) ParseShapeMarginToEmu(string value)
    {
        var parts = (value ?? string.Empty).Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 4)
        {
            int Emu(string s) => (int)Math.Round(SpacingConverter.ParsePoints(s) * 12700);
            return (Emu(parts[0]), Emu(parts[1]), Emu(parts[2]), Emu(parts[3]));
        }
        if (parts.Length == 1)
        {
            var emu = (int)Math.Round(SpacingConverter.ParsePoints(parts[0]) * 12700);
            return (emu, emu, emu, emu);
        }
        throw new ArgumentException(
            $"Invalid 'margin' value '{value}'. Expected single length (e.g. '4pt', '0.5cm') or 4-CSV 'L,T,R,B'.");
    }

    private static Drawing.ShapeTypeValues ParseExcelShapePreset(string name)
    {
        var key = (name ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(key))
            return Drawing.ShapeTypeValues.Rectangle;
        if (_shapePresetMap.TryGetValue(key, out var val))
            return val;
        // R20-01: Unknown preset falls back to rectangle, but emit a stderr
        // warning so users notice (silent rect was found by audit). 'custom'
        // is the common case — it would require a custGeom path which
        // officecli doesn't expose, so suggest raw-set explicitly.
        if (key == "custom")
        {
            Console.Error.WriteLine(
                "Warning: preset='custom' requires a custGeom path which officecli does not expose; " +
                "falling back to preset='rect'. Use a 'rawset' / direct OOXML edit if you need a custom path.");
        }
        else
        {
            Console.Error.WriteLine(
                $"Warning: unknown shape preset '{name}'; falling back to preset='rect'. " +
                "Valid presets include rect, ellipse, roundRect, triangle, rightArrow, etc.");
        }
        return Drawing.ShapeTypeValues.Rectangle;
    }

    private static (int x, int y, int width, int height) ParseAnchorBounds(
        Dictionary<string, string> properties, string defX, string defY, string defW, string defH)
    {
        // CONSISTENCY(shape-h-w-alias): mirror PPTX shape Add — accept `w` as
        // alias for `width`, `h` as alias for `height`. Without this mapping,
        // ParseAnchorDimension never sees the user value and the negative-
        // number guard is silently bypassed (h=-100 left as default 3 cells).
        var widthRaw = properties.GetValueOrDefault("width")
            ?? properties.GetValueOrDefault("w")
            ?? defW;
        var heightRaw = properties.GetValueOrDefault("height")
            ?? properties.GetValueOrDefault("h")
            ?? defH;
        return (
            ParseAnchorOrigin(properties.GetValueOrDefault("x", defX) ?? defX, "x"),
            ParseAnchorOrigin(properties.GetValueOrDefault("y", defY) ?? defY, "y"),
            ParseAnchorDimension(widthRaw, "width"),
            ParseAnchorDimension(heightRaw, "height")
        );
    }

    /// <summary>
    /// Parse an anchor origin value (x/y) that is either a plain non-negative
    /// integer cell index ("0", "5") or a unit-qualified offset ("2cm", "1in",
    /// "72pt"). Unit-qualified values are converted to a cell index using the
    /// same approximate EMU/column and EMU/row factors as ParseAnchorDimension.
    /// CONSISTENCY(ole-width-units): symmetric with width/height units.
    /// </summary>
    private static int ParseAnchorOrigin(string value, string name)
    {
        if (int.TryParse(value, out var plainInt))
        {
            if (plainInt < 0)
                throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");
            return plainInt;
        }

        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected a non-negative cell index or a unit-qualified offset (e.g. '2cm', '1in') for {name}, got '{value}'.");
        }
        if (emu < 0)
            throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");

        const long emuPerColApprox = 609600;
        const long emuPerRowApprox = 190500;
        if (name == "y")
            return (int)(emu / emuPerRowApprox);
        return (int)(emu / emuPerColApprox);
    }

    /// <summary>
    /// Parse a width/height anchor value that is either a plain integer
    /// cell-count ("3", "5") or a unit-qualified size ("6cm", "2in", "72pt").
    /// Unit-qualified values are converted to an approximate cell count using
    /// Excel's default ~64px (~0.66cm) column width and ~15pt row height.
    /// CONSISTENCY(ole-width-units): Picture/Drawing elsewhere accept ParseEmu;
    /// anchor.x/y stay as cell coordinates, but width/height tolerate EMU units.
    /// </summary>
    private static int ParseAnchorDimension(string value, string name)
    {
        if (int.TryParse(value, out var plainInt))
        {
            // R30-1: negative cell-count is meaningless and silently
            // produced an invalid file. Reject up front. CONSISTENCY with
            // ParseAnchorDimensionEmu's negative-int guard.
            if (plainInt <= 0)
                throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");
            return plainInt;
        }

        // Not a plain integer — treat as EMU-convertible size string.
        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected an integer cell count or a unit-qualified size (e.g. '6cm', '2in') for {name}, got '{value}'.");
        }
        // R30-1: unit-qualified negative ("-2in") parses to a negative
        // EMU; reject so the shape branch matches picture behavior.
        if (emu <= 0)
            throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");

        // Rough conversion: 1 default Excel column ≈ 64px ≈ 0.677cm ≈ 609600 EMU.
        // 1 default Excel row    ≈ 15pt ≈ 0.529cm ≈ 190500 EMU.
        // For width/height passed as a unit, choose the larger of the two
        // converters so "6cm" yields a sensible ~9 columns result either axis.
        const long emuPerColApprox = 609600;
        const long emuPerRowApprox = 190500;
        if (name == "height")
            return Math.Max(1, (int)(emu / emuPerRowApprox));
        return Math.Max(1, (int)(emu / emuPerColApprox));
    }

    // CONSISTENCY(ole-width-units): OLE round-trip preserves sub-cell precision
    // by storing the full EMU extent in ObjectAnchor's From/To ColumnOffset and
    // RowOffset, instead of rounding to whole cells like ParseAnchorDimension.
    // Picture/shape branches keep the integer behavior for now.
    private const long EmuPerColApprox = 609600;
    private const long EmuPerRowApprox = 190500;

    /// <summary>
    /// Parse a width/height anchor value into EMU. Plain integers are treated
    /// as cell counts and multiplied by the default column/row EMU width.
    /// Unit-qualified values (e.g. "6cm", "2in") are parsed via EmuConverter.
    /// </summary>
    private static long ParseAnchorDimensionEmu(string value, string name)
    {
        if (long.TryParse(value, out var plainInt))
        {
            // R30-1: reject negative bare integers up front. Without this,
            // `width=-5` silently rounded to 0 (still invalid) and produced
            // an Excel-rejected file with cx=0/cy=0 anchors.
            if (plainInt <= 0)
                throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");
            // Bare integers are interpreted as cell counts (original grammar),
            // but values that exceed Excel's column max (16384) are clearly
            // EMU — for either axis. Using a single threshold (instead of
            // axis-specific MaxRows=1048576) keeps the heuristic symmetric
            // with ParseAnchorOriginCell so x/y/width/height all flip to
            // EMU at the same boundary.
            const int MaxCellIndex = 16384;
            // R39-2: cell-count form is rejected above the grid limit so
            // mistakes like `width=20000` raise a clear error instead of
            // being silently treated as raw EMU. Users passing EMU should
            // use a unit-qualified form (`914400emu`, `1in`) which is parsed
            // through EmuConverter further down. CONSISTENCY with
            // ParseAnchorOriginCell.
            if (plainInt > MaxCellIndex - 1)
                throw new ArgumentException(
                    $"Picture/shape {name} column/row index must be in [0, {MaxCellIndex - 1}] (got '{value}'). For EMU-scale sizes use a unit-qualified value like '1in' / '6cm' / '72pt'.");
            long perCell = (name == "height") ? EmuPerRowApprox : EmuPerColApprox;
            return plainInt * perCell;
        }

        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected an integer cell count or a unit-qualified size (e.g. '6cm', '2in') for {name}, got '{value}'.");
        }
        // R30-1: unit-qualified negatives (e.g. "-5cm") parse to a negative
        // EMU; reject so we don't write `<xdr:to><xdr:col>-2</xdr:col>...`
        // anchors that crash Excel on open.
        if (emu <= 0)
            throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");
        return emu;
    }

    /// <summary>
    /// Parse an <c>anchor=</c> prop value as a cell-reference or cell-range
    /// (e.g. <c>"B2"</c> or <c>"B2:F7"</c>) into 0-based XDR column/row
    /// coordinates. Returns <c>false</c> for anchor-mode strings like
    /// <c>oneCell</c>/<c>twoCell</c>/<c>absolute</c>, which the caller should
    /// route to the anchorMode path instead. Throws <see cref="ArgumentException"/>
    /// for syntactically invalid range strings.
    ///
    /// When only a single cell is supplied, <c>toCol</c>/<c>toRow</c> are set
    /// to <c>-1</c> so callers can fall back to a size-derived extent (e.g.
    /// width/height × EMU-per-cell). The regex mirrors the OLE branch grammar.
    ///
    /// CONSISTENCY(xdr-coords): XDR ColumnId/RowId are 0-based; ColumnNameToIndex
    /// returns 1-based, so this helper subtracts 1 on the way out.
    /// </summary>
    internal static bool TryParseCellRangeAnchor(
        string? value, out int fromCol, out int fromRow, out int toCol, out int toRow)
    {
        fromCol = fromRow = 0;
        toCol = toRow = -1;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var m = System.Text.RegularExpressions.Regex.Match(
            value, @"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return false;
        fromCol = ColumnNameToIndex(m.Groups[1].Value) - 1;
        fromRow = int.Parse(m.Groups[2].Value) - 1;
        if (m.Groups[3].Success)
        {
            toCol = ColumnNameToIndex(m.Groups[3].Value) - 1;
            toRow = int.Parse(m.Groups[4].Value) - 1;
        }
        return true;
    }

    /// <summary>
    /// Return true if the given anchor= value is one of the recognized
    /// anchorMode tokens (oneCell/twoCell/absolute). Used by the picture
    /// branch to disambiguate mode-strings from cell-range strings.
    /// </summary>
    internal static bool IsAnchorModeToken(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        var v = value.Trim().ToLowerInvariant();
        return v is "onecell" or "twocell" or "absolute";
    }

    /// <summary>
    /// Apply `x` / `y` / `width` / `height` to the N-th chart's
    /// <see cref="XDR.TwoCellAnchor"/> in a drawings part. Accepts the same
    /// value grammar as OLE objects and chart Add: integer cell counts, or
    /// unit-qualified EMU strings ("6cm", "2in", "720pt", raw EMU).
    ///
    /// Returns any keys from the input dict that couldn't be applied (parse
    /// failures, missing anchor, ...). Keys present but successfully applied
    /// are NOT returned — the caller is expected to strip them before
    /// forwarding to the chart content setter.
    ///
    /// CONSISTENCY(chart-position-set): mirrors the PPTX
    /// PowerPointHandler.Set.cs chart path — same vocabulary, same units —
    /// so one prop grammar covers chart position across all three document
    /// types. The mutation mechanic differs because Excel charts are pinned
    /// to cells via TwoCellAnchor.
    /// </summary>
    // BUG-R11-04: read the N-th chart's TwoCellAnchor as a "B2:F7" cell range
    // for chart Get. Mirrors ApplyChartPositionSet's GraphicFrame lookup so the
    // index semantics match. Returns null if the chart has no TwoCellAnchor
    // (e.g. absolute-anchored), in which case the caller omits the field.
    private static string? GetChartAnchorRange(DrawingsPart drawingsPart, int chartIdx)
    {
        if (drawingsPart.WorksheetDrawing == null) return null;
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count) return null;
        var gf = chartFrames[chartIdx - 1];
        if (gf.Parent is not XDR.TwoCellAnchor anchor) return null;
        var fromM = anchor.FromMarker;
        var toM = anchor.ToMarker;
        if (fromM == null || toM == null) return null;
        if (!int.TryParse(fromM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var fc)) return null;
        if (!int.TryParse(fromM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var fr)) return null;
        if (!int.TryParse(toM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var tc)) return null;
        if (!int.TryParse(toM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var tr)) return null;
        // XDR col/row are 0-based; IndexToColumnName expects 1-based.
        return $"{IndexToColumnName(fc + 1)}{fr + 1}:{IndexToColumnName(tc + 1)}{tr + 1}";
    }

    private static List<string> ApplyChartPositionSet(
        DrawingsPart drawingsPart, int chartIdx, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        if (drawingsPart.WorksheetDrawing == null) return unsupported;

        // Find the N-th chart frame (same order as GetExcelCharts).
        var chartFrames = drawingsPart.WorksheetDrawing
            .Descendants<XDR.GraphicFrame>()
            .Where(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf))
            .ToList();
        if (chartIdx < 1 || chartIdx > chartFrames.Count) return unsupported;
        var gf = chartFrames[chartIdx - 1];
        var anchor = gf.Parent as XDR.TwoCellAnchor;
        if (anchor?.FromMarker == null || anchor.ToMarker == null)
        {
            foreach (var k in new[] { "x", "y", "width", "height" })
                if (properties.ContainsKey(k)) unsupported.Add(k);
            return unsupported;
        }

        var fromM = anchor.FromMarker;
        var toM = anchor.ToMarker;

        // ---- Position (x, y) → FromMarker cell indices ----
        // `x` = column index (0-based), `y` = row index (0-based). Integer
        // only — sub-cell offset is not supported here (matches chart Add).
        if (properties.TryGetValue("x", out var xStr))
        {
            if (int.TryParse(xStr, out var newFromCol) && newFromCol >= 0)
            {
                var fromColChild = fromM.GetFirstChild<XDR.ColumnId>();
                var oldFromCol = int.TryParse(fromColChild?.Text ?? "0", out var ofc) ? ofc : 0;
                if (fromColChild != null) fromColChild.Text = newFromCol.ToString();
                // Shift ToMarker column by the same delta to preserve width.
                var toColChild = toM.GetFirstChild<XDR.ColumnId>();
                if (toColChild != null && int.TryParse(toColChild.Text ?? "0", out var oldToCol))
                    toColChild.Text = (oldToCol + (newFromCol - oldFromCol)).ToString();
                // Reset fromCol offset to 0 (align to cell boundary).
                var fromColOffChild = fromM.GetFirstChild<XDR.ColumnOffset>();
                if (fromColOffChild != null) fromColOffChild.Text = "0";
            }
            else unsupported.Add("x");
        }

        if (properties.TryGetValue("y", out var yStr))
        {
            if (int.TryParse(yStr, out var newFromRow) && newFromRow >= 0)
            {
                var fromRowChild = fromM.GetFirstChild<XDR.RowId>();
                var oldFromRow = int.TryParse(fromRowChild?.Text ?? "0", out var ofr) ? ofr : 0;
                if (fromRowChild != null) fromRowChild.Text = newFromRow.ToString();
                var toRowChild = toM.GetFirstChild<XDR.RowId>();
                if (toRowChild != null && int.TryParse(toRowChild.Text ?? "0", out var oldToRow))
                    toRowChild.Text = (oldToRow + (newFromRow - oldFromRow)).ToString();
                var fromRowOffChild = fromM.GetFirstChild<XDR.RowOffset>();
                if (fromRowOffChild != null) fromRowOffChild.Text = "0";
            }
            else unsupported.Add("y");
        }

        // ---- Dimensions (width, height) → rebuild ToMarker from FromMarker ----
        // Reuses the OLE-object path's EMU math (EmuPerColApprox / EmuPerRowApprox
        // approximation, sub-cell offset preserves precision).
        if (properties.TryGetValue("width", out var wStr))
        {
            long emuTotal;
            try { emuTotal = ParseAnchorDimensionEmu(wStr, "width"); }
            catch { unsupported.Add("width"); emuTotal = -1; }
            if (emuTotal >= 0)
            {
                int.TryParse(fromM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var fromCol);
                long.TryParse(fromM.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out var fromColOff);
                long wholeCols = emuTotal / EmuPerColApprox;
                long remCols = emuTotal % EmuPerColApprox;
                var toColChild = toM.GetFirstChild<XDR.ColumnId>();
                if (toColChild != null) toColChild.Text = (fromCol + (int)wholeCols).ToString();
                var toColOffChild = toM.GetFirstChild<XDR.ColumnOffset>();
                if (toColOffChild != null) toColOffChild.Text = (fromColOff + remCols).ToString();
            }
        }

        if (properties.TryGetValue("height", out var hStr))
        {
            long emuTotal;
            try { emuTotal = ParseAnchorDimensionEmu(hStr, "height"); }
            catch { unsupported.Add("height"); emuTotal = -1; }
            if (emuTotal >= 0)
            {
                int.TryParse(fromM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var fromRow);
                long.TryParse(fromM.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out var fromRowOff);
                long wholeRows = emuTotal / EmuPerRowApprox;
                long remRows = emuTotal % EmuPerRowApprox;
                var toRowChild = toM.GetFirstChild<XDR.RowId>();
                if (toRowChild != null) toRowChild.Text = (fromRow + (int)wholeRows).ToString();
                var toRowOffChild = toM.GetFirstChild<XDR.RowOffset>();
                if (toRowOffChild != null) toRowOffChild.Text = (fromRowOff + remRows).ToString();
            }
        }

        drawingsPart.WorksheetDrawing.Save();
        return unsupported;
    }

    /// <summary>
    /// Parse x, y (cell indices) + width, height (EMU) for OLE anchors that
    /// need sub-cell precision. See ParseAnchorDimensionEmu for width/height
    /// semantics.
    /// </summary>
    private static (int x, int y, long widthEmu, long heightEmu) ParseAnchorBoundsEmu(
        Dictionary<string, string> properties, string defX, string defY, string defW, string defH)
    {
        return (
            ParseAnchorOriginCell(properties.GetValueOrDefault("x", defX) ?? defX, "x"),
            ParseAnchorOriginCell(properties.GetValueOrDefault("y", defY) ?? defY, "y"),
            ParseAnchorDimensionEmu(properties.GetValueOrDefault("width", defW) ?? defW, "width"),
            ParseAnchorDimensionEmu(properties.GetValueOrDefault("height", defH) ?? defH, "height")
        );
    }

    /// <summary>
    /// Parse anchor x/y origin into a cell index. Plain integers are normally
    /// cell counts, but values that exceed the sheet's column/row max can only
    /// be EMU offsets — fall back to dividing by the per-cell EMU constant so
    /// users passing inch-EMU values (e.g. x=914400) land on a sensible cell
    /// instead of overflowing the FromMarker. CONSISTENCY(ole-width-units):
    /// mirrors ParseAnchorDimensionEmu's "large bare int = EMU" heuristic for
    /// width/height.
    /// </summary>
    private static int ParseAnchorOriginCell(string value, string name)
    {
        if (long.TryParse(value, out var plainInt))
        {
            // R30-1: x/y origins are 0-based cell indices; negative values
            // would write an invalid <xdr:col>/-row anchor. Reject up front.
            if (plainInt < 0)
                throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");
            // Excel's column max (16384) is the tightest sheet-coordinate
            // bound — anything beyond that is unambiguously an EMU offset
            // (rows go to 1048576 but a row index that high is also clearly
            // EMU in practice). Use the same threshold for x and y so users
            // passing inch-EMU (914400) consistently land on a sensible cell
            // on either axis.
            const int MaxCellIndex = 16384;
            // R39-2: bare cell-count form must reject above-grid values
            // outright. Previously, x=20000 hit the "large bare int = EMU"
            // branch and divided by 609600, silently coercing the origin
            // back to col=0 (or row=0 for y). Cell-count input is small
            // by definition; if a user passes a number above the column
            // max, it's either a typo or an EMU value mistakenly fed
            // without a unit suffix. Either way, refuse rather than silently
            // remap. CONSISTENCY with R30-1 negative guard.
            if (plainInt > MaxCellIndex - 1)
                throw new ArgumentException(
                    $"Picture/shape {name} column/row index must be in [0, {MaxCellIndex - 1}] (got '{value}'). For EMU-scale offsets use a unit-qualified value like '1in' / '6cm' / '72pt'.");
            return (int)plainInt;
        }

        // Unit-qualified ("1in", "2cm") → EMU → cell count via the same per-cell constants.
        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected an integer cell index or a unit-qualified offset (e.g. '1in', '2cm') for {name}, got '{value}'.");
        }
        if (emu < 0)
            throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");
        long perCellOut = (name == "y") ? EmuPerRowApprox : EmuPerColApprox;
        return (int)(emu / perCellOut);
    }

    /// <summary>
    /// Reorder RunProperties children to match CT_RPrElt schema order:
    /// b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, rFont, family, charset, scheme
    /// </summary>
    private static void ReorderRunProperties(RunProperties rpr)
    {
        if (rpr == null || !rpr.HasChildren) return;
        var children = rpr.ChildElements.ToList();
        var ordered = children.OrderBy(c => GetRunPropertyOrder(c)).ToList();
        rpr.RemoveAllChildren();
        foreach (var child in ordered) rpr.AppendChild(child);
    }

    private static int GetRunPropertyOrder(DocumentFormat.OpenXml.OpenXmlElement element) => element switch
    {
        Bold => 0,
        Italic => 1,
        Strike => 2,
        Condense => 3,
        Extend => 4,
        Outline => 5,
        Shadow => 6,
        Underline => 7,
        VerticalTextAlignment => 8,
        FontSize => 9,
        Color => 10,
        RunFont => 11,
        FontFamily => 12,
        RunPropertyCharSet => 13,
        FontScheme => 14,
        _ => 99
    };

    // ==================== Extended Chart Helpers ====================

    private const string ExcelChartExUri = "http://schemas.microsoft.com/office/drawing/2014/chartex";

    /// <summary>
    /// Load a chartEx sidecar resource (style / colors XML) bundled as an
    /// embedded resource. Files are copied verbatim from an Excel reference
    /// treemap and reused for every chartEx type — they carry default
    /// style/palette content that has no dependency on chart layout or data.
    /// See the chartex-sidecars CONSISTENCY note in ExcelHandler.Add.cs for
    /// why these sidecars are load-bearing (Excel deletes the whole drawing
    /// if they are missing from the relationships).
    /// </summary>
    private static Stream LoadChartExResource(string fileName)
    {
        var assembly = typeof(ExcelHandler).Assembly;
        var resourceName = $"OfficeCli.Resources.{fileName}";
        var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException(
                $"Embedded resource not found: {resourceName}. Ensure it is declared in officecli.csproj.");
        return stream;
    }

    /// <summary>
    /// Check if an XDR.GraphicFrame contains an extended chart (cx:chart).
    /// </summary>
    private static bool IsExtendedChartFrame(XDR.GraphicFrame gf)
    {
        return gf.Descendants<Drawing.GraphicData>()
            .Any(gd => gd.Uri == ExcelChartExUri);
    }

    /// <summary>
    /// Get the relationship ID from an extended chart GraphicFrame.
    /// </summary>
    private static string? GetExtendedChartRelId(XDR.GraphicFrame gf)
    {
        var gd = gf.Descendants<Drawing.GraphicData>().FirstOrDefault(g => g.Uri == ExcelChartExUri);
        if (gd == null) return null;
        var typed = gd.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId>().FirstOrDefault();
        if (typed?.Id?.Value != null) return typed.Id.Value;
        foreach (var child in gd.ChildElements)
        {
            var rId = child.GetAttributes().FirstOrDefault(a =>
                a.LocalName == "id" && a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (rId.Value != null) return rId.Value;
        }
        return null;
    }

    /// <summary>
    /// Count all charts (both standard ChartPart and ExtendedChartPart) in a DrawingsPart.
    /// </summary>
    private static int CountExcelCharts(DrawingsPart drawingsPart)
    {
        if (drawingsPart.WorksheetDrawing == null) return 0;
        return drawingsPart.WorksheetDrawing.Descendants<XDR.GraphicFrame>()
            .Count(gf => gf.Descendants<C.ChartReference>().Any() || IsExtendedChartFrame(gf));
    }

    /// <summary>
    /// Represents a chart in Excel that could be either a standard ChartPart or an ExtendedChartPart.
    /// </summary>
    private class ExcelChartInfo
    {
        public ChartPart? StandardPart { get; set; }
        public ExtendedChartPart? ExtendedPart { get; set; }
        public bool IsExtended => ExtendedPart != null;
    }

    /// <summary>
    /// Get all chart parts (standard + extended) in document order by walking GraphicFrame elements.
    /// </summary>
    private static List<ExcelChartInfo> GetExcelCharts(DrawingsPart drawingsPart)
    {
        var result = new List<ExcelChartInfo>();
        if (drawingsPart.WorksheetDrawing == null) return result;

        foreach (var gf in drawingsPart.WorksheetDrawing.Descendants<XDR.GraphicFrame>())
        {
            var chartRef = gf.Descendants<C.ChartReference>().FirstOrDefault();
            if (chartRef?.Id?.Value != null)
            {
                try
                {
                    var chartPart = (ChartPart)drawingsPart.GetPartById(chartRef.Id.Value);
                    result.Add(new ExcelChartInfo { StandardPart = chartPart });
                }
                catch { /* skip invalid references */ }
            }
            else if (IsExtendedChartFrame(gf))
            {
                var relId = GetExtendedChartRelId(gf);
                if (relId == null) continue;
                try
                {
                    var extPart = (ExtendedChartPart)drawingsPart.GetPartById(relId);
                    result.Add(new ExcelChartInfo { ExtendedPart = extPart });
                }
                catch { /* skip invalid references */ }
            }
        }

        return result;
    }

    /// <summary>
    /// Find and replace text across all sheets (or a specific sheet). Returns the number of replacements made.
    /// Handles SharedStringTable entries as well as inline strings and direct cell values.
    /// </summary>
    private int FindAndReplace(string find, string replace, WorksheetPart? targetSheet)
    {
        if (string.IsNullOrEmpty(find)) return 0;
        int totalCount = 0;
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return 0;

        // Replace in SharedStringTable (affects all sheets sharing these strings)
        if (targetSheet == null)
        {
            var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
            if (sst != null)
            {
                foreach (var si in sst.Elements<SharedStringItem>())
                {
                    // Handle simple text items
                    var textEl = si.GetFirstChild<Text>();
                    if (textEl?.Text != null && textEl.Text.Contains(find, StringComparison.Ordinal))
                    {
                        int count = CountOccurrences(textEl.Text, find);
                        textEl.Text = textEl.Text.Replace(find, replace, StringComparison.Ordinal);
                        totalCount += count;
                    }

                    // Handle rich text runs
                    foreach (var run in si.Elements<Run>())
                    {
                        var runText = run.GetFirstChild<Text>();
                        if (runText?.Text != null && runText.Text.Contains(find, StringComparison.Ordinal))
                        {
                            int count = CountOccurrences(runText.Text, find);
                            runText.Text = runText.Text.Replace(find, replace, StringComparison.Ordinal);
                            totalCount += count;
                        }
                    }
                }
                sst.Save();
            }
        }

        // Replace in inline strings and direct cell values
        var sheets = targetSheet != null
            ? [targetSheet]
            : workbookPart.WorksheetParts.ToList();

        foreach (var wsPart in sheets)
        {
            var sheetData = wsPart.Worksheet?.GetFirstChild<SheetData>();
            if (sheetData == null) continue;

            foreach (var row in sheetData.Elements<Row>())
            {
                foreach (var cell in row.Elements<Cell>())
                {
                    // Inline string
                    var inlineStr = cell.GetFirstChild<InlineString>();
                    if (inlineStr != null)
                    {
                        var t = inlineStr.GetFirstChild<Text>();
                        if (t?.Text != null && t.Text.Contains(find, StringComparison.Ordinal))
                        {
                            int count = CountOccurrences(t.Text, find);
                            t.Text = t.Text.Replace(find, replace, StringComparison.Ordinal);
                            totalCount += count;
                        }
                        // Rich text runs inside inline string
                        foreach (var run in inlineStr.Elements<Run>())
                        {
                            var runText = run.GetFirstChild<Text>();
                            if (runText?.Text != null && runText.Text.Contains(find, StringComparison.Ordinal))
                            {
                                int count = CountOccurrences(runText.Text, find);
                                runText.Text = runText.Text.Replace(find, replace, StringComparison.Ordinal);
                                totalCount += count;
                            }
                        }
                        continue;
                    }

                    // Direct string value (DataType is null or String)
                    if (cell.DataType?.Value == CellValues.String)
                    {
                        var cv = cell.CellValue;
                        if (cv?.Text != null && cv.Text.Contains(find, StringComparison.Ordinal))
                        {
                            int count = CountOccurrences(cv.Text, find);
                            cv.Text = cv.Text.Replace(find, replace, StringComparison.Ordinal);
                            totalCount += count;
                        }
                    }

                    // SharedStringTable reference — if targeting a specific sheet, replace inline
                    if (targetSheet != null && cell.DataType?.Value == CellValues.SharedString)
                    {
                        var sst = workbookPart.SharedStringTablePart?.SharedStringTable;
                        if (sst != null && cell.CellValue?.Text != null
                            && int.TryParse(cell.CellValue.Text, out var sstIdx))
                        {
                            var items = sst.Elements<SharedStringItem>().ToList();
                            if (sstIdx >= 0 && sstIdx < items.Count)
                            {
                                var si = items[sstIdx];
                                var siText = si.GetFirstChild<Text>();
                                if (siText?.Text != null && siText.Text.Contains(find, StringComparison.Ordinal))
                                {
                                    int count = CountOccurrences(siText.Text, find);
                                    siText.Text = siText.Text.Replace(find, replace, StringComparison.Ordinal);
                                    totalCount += count;
                                }
                                foreach (var run in si.Elements<Run>())
                                {
                                    var runText = run.GetFirstChild<Text>();
                                    if (runText?.Text != null && runText.Text.Contains(find, StringComparison.Ordinal))
                                    {
                                        int count = CountOccurrences(runText.Text, find);
                                        runText.Text = runText.Text.Replace(find, replace, StringComparison.Ordinal);
                                        totalCount += count;
                                    }
                                }
                                sst.Save();
                            }
                        }
                    }
                }
            }

            SaveWorksheet(wsPart);
        }

        return totalCount;
    }

    private static int CountOccurrences(string text, string find)
    {
        int count = 0;
        int idx = 0;
        while ((idx = text.IndexOf(find, idx, StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += find.Length;
        }
        return count;
    }

    /// <summary>
    /// Parse a dataRange (e.g. "Sheet1!A1:D5" or "A1:B3") and read cell data from the worksheet.
    /// Returns series data and populates properties with cell references for chart building.
    /// First row = category labels + series names, remaining rows = data.
    /// </summary>
    private (List<(string name, double[] values)> seriesData, string[]? categories) ParseDataRangeForChart(
        string dataRange, string defaultSheetName, Dictionary<string, string> properties)
    {
        // CONSISTENCY(defined-name-range): if dataRange has no '!' and no ':' and
        // looks like a workbook-defined name, resolve it to its referent range
        // (e.g. "MyData" -> "Sheet1!$A$1:$B$3"). Excel charts accept defined-name
        // references as a data source, so do the same here.
        var trimmedInput = dataRange.Trim();
        if (!trimmedInput.Contains('!') && !trimmedInput.Contains(':') &&
            System.Text.RegularExpressions.Regex.IsMatch(trimmedInput, @"^[A-Za-z_][A-Za-z0-9_\.]*$"))
        {
            var workbook = _doc.WorkbookPart?.Workbook;
            var defNames = workbook?.GetFirstChild<DefinedNames>();
            var match = defNames?.Elements<DefinedName>()
                .FirstOrDefault(dn => string.Equals(dn.Name?.Value, trimmedInput, StringComparison.OrdinalIgnoreCase));
            if (match == null || string.IsNullOrEmpty(match.Text))
                throw new ArgumentException($"DefinedName '{trimmedInput}' not found");
            dataRange = match.Text!;
        }

        // Parse sheet name and range
        string rangeSheetName = defaultSheetName;
        string rangePart = dataRange.Trim();
        var bangIdx = rangePart.IndexOf('!');
        if (bangIdx >= 0)
        {
            rangeSheetName = rangePart[..bangIdx].Trim('\'');
            rangePart = rangePart[(bangIdx + 1)..];
        }

        // Strip any $ signs for parsing
        var cleanRange = rangePart.Replace("$", "");
        var rangeParts = cleanRange.Split(':');
        if (rangeParts.Length != 2)
            throw new ArgumentException($"Invalid dataRange: '{dataRange}'. Expected format: 'Sheet1!A1:D5', 'A1:B3', or a defined-name");

        var (startCol, startRow) = ParseCellReference(rangeParts[0]);
        var (endCol, endRow) = ParseCellReference(rangeParts[1]);
        var startColIdx = ColumnNameToIndex(startCol);
        var endColIdx = ColumnNameToIndex(endCol);

        // Find the worksheet and read cells
        var ws = FindWorksheet(rangeSheetName)
            ?? throw new ArgumentException($"Sheet not found: {rangeSheetName}");
        var sheetData = GetSheet(ws).GetFirstChild<SheetData>();
        if (sheetData == null)
            throw new ArgumentException($"Sheet '{rangeSheetName}' has no data");

        // Build cell lookup. Track value, the originating Cell (for DataType),
        // and a "is blank" flag for cells that exist but carry no value.
        // R20-03: blank-vs-zero distinction is needed for dispBlanksAs=gap.
        // R20-04: DataType drives header detection — only string-typed
        // first-row cells are treated as series names.
        var cellLookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var cellTypeLookup = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        var cellPresent = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value != null)
                {
                    cellLookup[cell.CellReference.Value] = GetCellDisplayValue(cell);
                    cellTypeLookup[cell.CellReference.Value] = cell;
                    cellPresent.Add(cell.CellReference.Value);
                }
            }
        }

        // R20-04: a first-row cell counts as a header only when its DataType
        // is string-like (SharedString / InlineString / String). Numeric or
        // missing first-row cells mean "no header" — series starts at row 1.
        bool IsStringTypedHeader(string cellRef)
        {
            if (!cellTypeLookup.TryGetValue(cellRef, out var c)) return false;
            var dt = c.DataType?.Value;
            return dt == CellValues.SharedString
                || dt == CellValues.InlineString
                || dt == CellValues.String;
        }

        // Decide globally: if ANY non-corner cell in the first row is string-typed,
        // treat row 1 as headers; otherwise treat all rows as data and synthesize
        // series names. Picking globally keeps a single header convention
        // across columns (mixed string/number headers would be ambiguous).
        bool hasHeaderRow = false;
        for (int c = startColIdx + 1; c <= endColIdx; c++)
        {
            var headerRef = $"{IndexToColumnName(c)}{startRow}";
            if (IsStringTypedHeader(headerRef)) { hasHeaderRow = true; break; }
        }

        int dataStartRow = hasHeaderRow ? startRow + 1 : startRow;

        // First column (excluding header row if present) = category labels
        var categories = new List<string>();
        for (int r = dataStartRow; r <= endRow; r++)
        {
            var cellRef = $"{startCol}{r}";
            cellLookup.TryGetValue(cellRef, out var catVal);
            categories.Add(catVal ?? "");
        }

        var seriesData = new List<(string name, double[] values)>();
        int seriesIdx = 1;
        for (int c = startColIdx + 1; c <= endColIdx; c++)
        {
            var colName = IndexToColumnName(c);
            string seriesName;
            if (hasHeaderRow)
            {
                var headerRef = $"{colName}{startRow}";
                cellLookup.TryGetValue(headerRef, out var sn);
                seriesName = sn ?? $"Series {seriesIdx}";
            }
            else
            {
                seriesName = $"Series {seriesIdx}";
            }

            // Series values + per-index blank tracking. R20-03: under
            // dispBlanksAs=gap, blank source cells must be omitted from the
            // numCache; we forward the blank-index list via properties so
            // ApplySeriesReferences/numCache builder can honor it.
            var values = new List<double>();
            var blankIndexes = new List<int>();
            int idx = 0;
            for (int r = dataStartRow; r <= endRow; r++)
            {
                var cellRef = $"{colName}{r}";
                bool isBlank = !cellPresent.Contains(cellRef)
                    || string.IsNullOrEmpty(cellLookup.GetValueOrDefault(cellRef));
                cellLookup.TryGetValue(cellRef, out var valStr);
                if (double.TryParse(valStr, System.Globalization.CultureInfo.InvariantCulture, out var num))
                    values.Add(num);
                else
                    values.Add(0);
                if (isBlank) blankIndexes.Add(idx);
                idx++;
            }

            // Set up cell references in properties for ApplySeriesReferences
            var valuesRef = $"{rangeSheetName}!${colName}${dataStartRow}:${colName}${endRow}";
            var categoriesRef = $"{rangeSheetName}!${startCol}${dataStartRow}:${startCol}${endRow}";
            properties[$"series{seriesIdx}.name"] = seriesName;
            properties[$"series{seriesIdx}.values"] = valuesRef;
            properties[$"series{seriesIdx}.categories"] = categoriesRef;
            if (blankIndexes.Count > 0)
                properties[$"series{seriesIdx}._blankIndexes"] = string.Join(",", blankIndexes);

            seriesData.Add((seriesName, values.ToArray()));
            seriesIdx++;
        }

        return (seriesData, categories.ToArray());
    }

    // ==================== Binary Extraction ====================
    //
    // Support for `officecli get --save <dest>`. Parses the path to find
    // the owning worksheet and queries the node's relId. Both DrawingsPart
    // (pictures) and WorksheetPart (embedded ole/package) are consulted
    // because pictures live on DrawingsPart while OLE payloads live on
    // WorksheetPart directly.
    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        var node = Get(path, 0);
        if (node == null) return false;
        if (!node.Format.TryGetValue("relId", out var relObj) || relObj is not string relId
            || string.IsNullOrEmpty(relId))
            return false;

        // Path looks like /SheetName/... — find the worksheet.
        var normalized = NormalizeExcelPath(path);
        normalized = ResolveSheetIndexInPath(normalized);
        var segments = normalized.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheetPart = FindWorksheet(sheetName);
        if (worksheetPart == null) return false;

        DocumentFormat.OpenXml.Packaging.OpenXmlPart? part = null;
        try { part = worksheetPart.GetPartById(relId); } catch { /* try drawing */ }
        if (part == null && worksheetPart.DrawingsPart != null)
        {
            try { part = worksheetPart.DrawingsPart.GetPartById(relId); } catch { /* fall through */ }
        }
        if (part == null) return false;

        // BUG-R10-04: create the destination directory if missing so
        // `get --save ./outdir/file.bin` works when outdir doesn't exist.
        var destDir = Path.GetDirectoryName(destPath);
        if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            Directory.CreateDirectory(destDir);

        // CONSISTENCY(ole-cfb-wrap): non-Office OLE payloads are stored as
        // CFB containers with \x01Ole10Native; unwrap on read so the caller
        // gets back the bytes they fed in via `add ole src=...`.
        byte[] rawBytes;
        using (var src = part.GetStream())
        using (var ms = new MemoryStream())
        {
            src.CopyTo(ms);
            rawBytes = ms.ToArray();
        }
        var payload = OfficeCli.Core.OleHelper.UnwrapOle10NativeIfCfb(rawBytes);
        File.WriteAllBytes(destPath, payload);
        byteCount = payload.Length;
        contentType = part.ContentType;
        return true;
    }

    // ==================== OLE Object Writing Helpers ====================

    /// <summary>
    /// Ensure the given VmlDrawingPart contains a minimal v:shape with the
    /// specified shapeId so the schema-required <c>oleObject/@shapeId</c>
    /// attribute has a valid target. Modern Excel (2010+) renders OLE from
    /// the companion <c>objectPr/anchor</c>, but the shape itself still
    /// has to exist for a round-trip — otherwise opening the workbook in
    /// older Excel versions tends to drop the object silently.
    /// </summary>
    internal static void EnsureExcelVmlShapeForOle(VmlDrawingPart vmlPart, uint shapeId,
        int fromCol, int fromRow, int toCol, int toRow)
    {
        // Load the existing VML (may be absent on a freshly-created part).
        string existing;
        try
        {
            using var readStream = vmlPart.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
            using var reader = new StreamReader(readStream);
            existing = reader.ReadToEnd();
        }
        catch
        {
            existing = string.Empty;
        }

        // VML clientData carries the anchor (16 coordinates: from/to col/row + offsets).
        // Coordinates are in the legacy "left, top, right, bottom" pixel order.
        var anchorValue = $"{fromCol}, 0, {fromRow}, 0, {toCol}, 0, {toRow}, 0";
        var newShape = $"""
<v:shape id="_x0000_s{shapeId}" type="#_x0000_t75" style='position:absolute;margin-left:0;margin-top:0;width:100pt;height:40pt;visibility:hidden' o:oleicon="t" o:ole="" filled="f" stroked="f">
 <v:imagedata chromakey="white"/>
 <o:lock v:ext="edit" aspectratio="t"/>
 <x:ClientData ObjectType="Pict">
  <x:Anchor>{anchorValue}</x:Anchor>
  <x:CF>Pict</x:CF>
  <x:AutoPict/>
 </x:ClientData>
</v:shape>
""";

        string merged;
        if (string.IsNullOrWhiteSpace(existing))
        {
            // Build a minimal xml with shapetype + our shape.
            merged = $"""
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>
 <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
  <v:stroke joinstyle="miter"/>
  <v:formulas>
   <v:f eqn="if lineDrawn pixelLineWidth 0"/>
   <v:f eqn="sum @0 1 0"/>
   <v:f eqn="sum 0 0 @1"/>
   <v:f eqn="prod @2 1 2"/>
   <v:f eqn="prod @3 21600 pixelWidth"/>
   <v:f eqn="prod @3 21600 pixelHeight"/>
   <v:f eqn="sum @0 0 1"/>
   <v:f eqn="prod @6 1 2"/>
   <v:f eqn="prod @7 21600 pixelWidth"/>
   <v:f eqn="sum @8 21600 0"/>
   <v:f eqn="prod @7 21600 pixelHeight"/>
   <v:f eqn="sum @10 21600 0"/>
  </v:formulas>
  <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
  <o:lock v:ext="edit" aspectratio="t"/>
 </v:shapetype>
{newShape}
</xml>
""";
        }
        else
        {
            // Append our shape before the closing </xml> tag.
            var closeIdx = existing.LastIndexOf("</xml>", StringComparison.OrdinalIgnoreCase);
            if (closeIdx < 0) closeIdx = existing.Length;
            merged = existing.Substring(0, closeIdx) + newShape + "\n</xml>";
        }

        using var writeStream = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(writeStream);
        writer.Write(merged);
    }

    // ==================== OLE Object Reading ====================
    //
    // Enumerate all OLE objects attached to a worksheet. Excel stores these
    // as <x:oleObjects> inside the worksheet (each <x:oleObject> has
    // progId + shapeId + r:id), plus matching EmbeddedObjectPart /
    // EmbeddedPackagePart parts joined by rel id.
    //
    // CONSISTENCY(ole-orphan-indexing): orphan embedded parts (backing parts
    // with no matching x:oleObject XML element) are intentionally NOT
    // surfaced under the ole[N] index. Set/Remove dispatch on
    // ws.Descendants<OleObject>() which only yields schema-typed elements;
    // indexing orphans here would cause Get to return nodes that Set/Remove
    // cannot address. Orphans can still be audited via Validate() or raw
    // package inspection.
    internal List<DocumentNode> CollectOleNodesForSheet(string sheetName, WorksheetPart worksheetPart)
    {
        var nodes = new List<DocumentNode>();

        // Walk schema-typed <x:oleObject> elements (may live inside
        // <oleObjects>, directly under <worksheet>, or wrapped in an
        // <mc:AlternateContent><mc:Choice>...</mc:Choice></mc:AlternateContent>).
        // Descendants<OleObject> picks all of those up.
        var oleElements = GetSheet(worksheetPart).Descendants<OleObject>().ToList();
        for (int i = 0; i < oleElements.Count; i++)
        {
            var ole = oleElements[i];
            var node = new DocumentNode
            {
                Path = $"/{sheetName}/ole[{i + 1}]",
                Type = "ole",
                Text = ole.ProgId?.Value ?? "",
            };
            node.Format["objectType"] = "ole";
            // CONSISTENCY(ole-display): PPT and Word OLE Get both expose
            // Format["display"]. Excel worksheet OLE objects have no
            // DrawAspect concept — they always render as icons — so emit
            // a fixed "icon" value for schema symmetry.
            node.Format["display"] = "icon";
            if (ole.ProgId?.Value != null) node.Format["progId"] = ole.ProgId.Value;
            if (ole.ShapeId?.Value != null) node.Format["shapeId"] = (long)ole.ShapeId.Value;
            if (ole.Link?.Value != null) node.Format["link"] = ole.Link.Value;

            var relId = ole.Id?.Value;
            if (!string.IsNullOrEmpty(relId))
            {
                node.Format["relId"] = relId;
                try
                {
                    var part = worksheetPart.GetPartById(relId);
                    if (part != null)
                        OfficeCli.Core.OleHelper.PopulateFromPart(node, part, ole.ProgId?.Value);
                }
                catch
                {
                    // Relationship may be missing; leave part-sourced fields absent.
                }
            }

            // Expose anchor rectangle as unit-qualified width/height (cm).
            // CONSISTENCY(ole-width-units): mirrors PPTX/Word OLE which emit
            // EmuConverter.FormatEmu strings. Internally the anchor stores
            // only cell markers (col/row), so convert via the same rough
            // default-column/row → EMU constants used by ParseAnchorDimension
            // (Add-side). Known limitation: Excel's actual column widths are
            // ignored; this is a symmetric round-trip of the Add inputs.
            var objectPr = ole.GetFirstChild<EmbeddedObjectProperties>();
            var objAnchor = objectPr?.GetFirstChild<ObjectAnchor>();
            if (objAnchor != null)
            {
                var fromM = objAnchor.GetFirstChild<FromMarker>();
                var toM = objAnchor.GetFirstChild<ToMarker>();
                if (fromM != null && toM != null)
                {
                    int fromCol = 0, fromRow = 0, toCol = 0, toRow = 0;
                    long fromColOff = 0, fromRowOff = 0, toColOff = 0, toRowOff = 0;
                    int.TryParse(fromM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out fromCol);
                    int.TryParse(fromM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out fromRow);
                    int.TryParse(toM.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out toCol);
                    int.TryParse(toM.GetFirstChild<XDR.RowId>()?.Text ?? "0", out toRow);
                    long.TryParse(fromM.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out fromColOff);
                    long.TryParse(fromM.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out fromRowOff);
                    long.TryParse(toM.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out toColOff);
                    long.TryParse(toM.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out toRowOff);
                    // CONSISTENCY(ole-width-units): rebuild EMU extent from
                    // (cell-count * approx-per-cell) + (to-offset - from-offset)
                    // so sub-cell precision set on Add survives Get.
                    long widthEmu = Math.Max(0, (long)(toCol - fromCol)) * EmuPerColApprox
                        + (toColOff - fromColOff);
                    long heightEmu = Math.Max(0, (long)(toRow - fromRow)) * EmuPerRowApprox
                        + (toRowOff - fromRowOff);
                    if (widthEmu < 0) widthEmu = 0;
                    if (heightEmu < 0) heightEmu = 0;
                    node.Format["width"] = OfficeCli.Core.EmuConverter.FormatEmu(widthEmu);
                    node.Format["height"] = OfficeCli.Core.EmuConverter.FormatEmu(heightEmu);
                    // CONSISTENCY(ole-anchor-roundtrip): expose the cell-range
                    // form so `add ... anchor=B2:D4` survives Get/Query. XDR
                    // markers are 0-based; A1-style needs +1 on both axes.
                    node.Format["anchor"] =
                        $"{IndexToColumnName(fromCol + 1)}{fromRow + 1}:{IndexToColumnName(toCol + 1)}{toRow + 1}";
                }
            }

            nodes.Add(node);
        }

        return nodes;
    }

    // CONSISTENCY(xlsx/table-autoexpand): custom namespace marker stored on
    // the <x:table> root so `autoExpand=true` survives open/close cycles.
    // Real Excel ignores unknown-namespace attributes, so the file is still
    // opened cleanly on Windows — the flag only affects officecli's own
    // cell-write auto-grow behavior.
    private const string AutoExpandNamespaceUri = "https://officecli.ai/2025/autoexpand";
    private const string AutoExpandNamespacePrefix = "ae";
    private const string AutoExpandAttrName = "autoExpand";

    private static void SetTableAutoExpandMarker(Table table, bool enabled)
    {
        if (enabled)
        {
            table.AddNamespaceDeclaration(AutoExpandNamespacePrefix, AutoExpandNamespaceUri);
            table.SetAttribute(new OpenXmlAttribute(
                AutoExpandNamespacePrefix, AutoExpandAttrName, AutoExpandNamespaceUri, "1"));
        }
    }

    private static bool TableHasAutoExpand(Table? table)
    {
        if (table == null) return false;
        foreach (var attr in table.GetAttributes())
        {
            if (attr.NamespaceUri == AutoExpandNamespaceUri
                && attr.LocalName == AutoExpandAttrName
                && (attr.Value == "1" || string.Equals(attr.Value, "true", StringComparison.OrdinalIgnoreCase)))
                return true;
        }
        return false;
    }

    // Eager auto-grow on cell Add/Set. Called after writing `cellRef` on
    // `worksheet`. For each table on the sheet flagged with autoExpand:
    //   - if cell is in the row immediately below the table AND its column
    //     is within the table's column span → grow endRow by 1.
    //   - else if cell is in the column immediately right of the table AND
    //     its row is within the table's row span → grow endCol by 1 and
    //     append a blank tableColumn.
    // Both extensions are never applied at once (conservative).
    private void MaybeExpandTablesForCell(WorksheetPart worksheet, string cellRef)
    {
        var (cellCol, cellRow) = ParseCellReference(cellRef.ToUpperInvariant());
        var cellColIdx = ColumnNameToIndex(cellCol);

        foreach (var tdp in worksheet.TableDefinitionParts.ToList())
        {
            var table = tdp.Table;
            if (table == null) continue;
            if (!TableHasAutoExpand(table)) continue;
            if (table.Reference?.Value is not string rangeRef) continue;
            if (!rangeRef.Contains(':')) continue;

            var parts = rangeRef.Split(':');
            var (startColName, startRow) = ParseCellReference(parts[0]);
            var (endColName, endRow) = ParseCellReference(parts[1]);
            var startColIdx = ColumnNameToIndex(startColName);
            var endColIdx = ColumnNameToIndex(endColName);

            // Row below? (cell row == endRow + 1, within column span).
            if (cellRow == endRow + 1 && cellColIdx >= startColIdx && cellColIdx <= endColIdx)
            {
                endRow += 1;
                var newRef = $"{startColName}{startRow}:{endColName}{endRow}";
                table.Reference = newRef;
                var af = table.GetFirstChild<AutoFilter>();
                if (af != null) af.Reference = newRef;
                table.Save();
                continue;
            }

            // Column right? (cell col == endCol + 1, within row span).
            if (cellColIdx == endColIdx + 1 && cellRow >= startRow && cellRow <= endRow)
            {
                endColIdx += 1;
                var newEndColName = IndexToColumnName(endColIdx);
                var newRef = $"{startColName}{startRow}:{newEndColName}{endRow}";
                table.Reference = newRef;
                var af = table.GetFirstChild<AutoFilter>();
                if (af != null) af.Reference = newRef;

                var tableColumns = table.GetFirstChild<TableColumns>();
                if (tableColumns != null)
                {
                    var existing = tableColumns.Elements<TableColumn>().ToList();
                    var nextId = existing.Count == 0
                        ? 1u
                        : existing.Max(tc => tc.Id?.Value ?? 0u) + 1u;
                    var used = new HashSet<string>(
                        existing.Select(tc => tc.Name?.Value ?? "")
                                .Where(n => !string.IsNullOrEmpty(n)),
                        StringComparer.OrdinalIgnoreCase);
                    var baseName = $"Column{existing.Count + 1}";
                    var colName = baseName;
                    int dedupeIdx = 2;
                    while (!used.Add(colName))
                        colName = $"{baseName}{dedupeIdx++}";
                    tableColumns.AppendChild(new TableColumn
                    {
                        Id = nextId,
                        Name = colName
                    });
                    tableColumns.Count = (uint)tableColumns.Elements<TableColumn>().Count();
                }

                table.Save();
            }
        }
    }

    /// <summary>
    /// R9-1: scan a formula body for Sheet-qualified refs (bare `Sheet1!A1`
    /// or quoted `'My Data'!A1`) and return true if any referenced sheet
    /// name does not exist in the current workbook. Used to suppress the
    /// evaluator-based cachedValue fallback when cross-sheet refs point at
    /// a removed sheet — Real Excel shows `#REF!` there; we should not
    /// invent a "0".
    /// </summary>
    private bool FormulaReferencesMissingSheet(string formula)
    {
        if (string.IsNullOrEmpty(formula)) return false;
        var wb = _doc.WorkbookPart?.Workbook;
        if (wb == null) return false;
        var names = new HashSet<string>(
            wb.Descendants<Sheet>().Select(s => s.Name?.Value ?? "").Where(n => n.Length > 0),
            StringComparer.OrdinalIgnoreCase);

        // Quoted form: '...'! — inner single quotes escaped as ''
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(formula, @"'((?:[^']|'')+)'!"))
        {
            var name = m.Groups[1].Value.Replace("''", "'");
            if (!names.Contains(name)) return true;
        }
        // Bare form: Name! — letters/digits/underscore/period (Excel allows these unquoted)
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(formula, @"(?<![A-Za-z0-9_'.])([A-Za-z_][A-Za-z0-9_.]*)!"))
        {
            if (!names.Contains(m.Groups[1].Value)) return true;
        }
        return false;
    }

    // R13-1: Excel rejects cell values longer than 32767 chars (2^15 - 1) with
    // 0x800A03EC on save/open. Reject at write time with a clear error rather
    // than silently writing a file Excel will refuse to open.
    internal const int MaxCellTextLength = 32767;

    internal static void EnsureCellValueLength(string? value, string? cellRef = null)
    {
        if (value == null) return;
        if (value.Length > MaxCellTextLength)
        {
            var where = string.IsNullOrEmpty(cellRef) ? "" : $" at {cellRef}";
            throw new ArgumentException(
                $"Cell value{where} exceeds Excel's {MaxCellTextLength}-character limit (got {value.Length})");
        }
    }

    // R13-2: central ISO date parser accepting date-only, date+time, and the
    // common `T`-separator variants. Used by Add/Set cell value paths so
    // `2024-03-15T10:30:00` is converted to an OADate serial instead of being
    // written as a literal string (which Excel renders as text, not a date).
    internal static readonly string[] IsoDateFormats =
    {
        "yyyy-MM-dd",
        "yyyy/MM/dd",
        "yyyy-MM-dd HH:mm",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-MM-ddTHH:mm",
        "yyyy-MM-ddTHH:mm:ss",
        "yyyy-MM-ddTHH:mm:ssZ",
        "yyyy-MM-ddTHH:mm:ss.fff",
        "yyyy-MM-ddTHH:mm:ss.fffZ",
    };

    internal static bool TryParseIsoDateFlexible(string value, out System.DateTime result)
        => System.DateTime.TryParseExact(
            value,
            IsoDateFormats,
            System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None,
            out result);

    /// <summary>
    /// Build a <x:font> child for a dxf (differentialFormat) from font.* sub-props.
    /// Supports bold, italic, underline (single/double), strike, size, name, color.
    /// Returns null if no font sub-props were supplied.
    /// </summary>
    internal static Font? BuildFormulaCfFont(Dictionary<string, string> properties)
    {
        bool any = false;
        var font = new Font();
        if (properties.TryGetValue("font.bold", out var fBold) && ParseHelpers.IsTruthy(fBold))
        { font.Append(new Bold()); any = true; }
        if (properties.TryGetValue("font.italic", out var fItalic) && ParseHelpers.IsTruthy(fItalic))
        { font.Append(new Italic()); any = true; }
        if (properties.TryGetValue("font.strike", out var fStrike) && ParseHelpers.IsTruthy(fStrike))
        { font.Append(new Strike()); any = true; }
        if (properties.TryGetValue("font.underline", out var fUnder))
        {
            var ul = new Underline();
            var lv = fUnder.Trim().ToLowerInvariant();
            ul.Val = lv switch
            {
                "double" or "dbl" => UnderlineValues.Double,
                "singleaccounting" or "singleacct" => UnderlineValues.SingleAccounting,
                "doubleaccounting" or "doubleacct" => UnderlineValues.DoubleAccounting,
                "none" or "false" => UnderlineValues.None,
                _ => UnderlineValues.Single
            };
            font.Append(ul);
            any = true;
        }
        if (properties.TryGetValue("font.size", out var fSize))
        {
            // Accept "12", "12pt", "10.5pt" — strip trailing "pt" if present.
            var cleaned = fSize.Trim().TrimEnd('p', 't', 'P', 'T', ' ');
            if (double.TryParse(cleaned, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var sz))
            {
                font.Append(new FontSize { Val = sz });
                any = true;
            }
        }
        if (properties.TryGetValue("font.name", out var fName) && !string.IsNullOrWhiteSpace(fName))
        {
            font.Append(new FontName { Val = fName });
            any = true;
        }
        if (properties.TryGetValue("font.color", out var fColor))
        {
            var norm = ParseHelpers.NormalizeArgbColor(fColor);
            font.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = norm });
            any = true;
        }
        return any ? font : null;
    }

    // R37-B: detect whether a hyperlink target is an internal sheet/cell reference
    // (location-based) rather than an external URI. Recognises both the canonical
    // "#Sheet1!A1" form and the bare "Sheet1!A1" form (no leading '#'), as well
    // as the quoted variants used when the sheet name contains spaces or special
    // characters: "#'Multi Word'!A1" and "'Multi Word'!A1".
    //
    // Returns the location string (without leading '#') when matched, or null.
    // The location string is what gets written to the OOXML @location attribute.
    private static readonly System.Text.RegularExpressions.Regex s_internalLinkPattern =
        new System.Text.RegularExpressions.Regex(
            @"^#?(?:'(?:[^']|'')+'|[A-Za-z_][\w\.]*)![A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?$",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    internal static string? TryParseInternalHyperlinkLocation(string value)
    {
        if (string.IsNullOrEmpty(value)) return null;
        if (!s_internalLinkPattern.IsMatch(value)) return null;
        return value.StartsWith("#") ? value.Substring(1) : value;
    }

    // R24-1: detect whether a styleProps bag asks for the text number format
    // ("@"). All three accepted aliases are checked: numberformat, numfmt,
    // format. Whitespace is trimmed; quoting is not expected here because
    // ExcelStyleManager already strips surrounding quotes upstream.
    private static bool IsTextNumberFormat(Dictionary<string, string> styleProps)
    {
        foreach (var key in new[] { "numberformat", "numfmt", "format" })
        {
            if (styleProps.TryGetValue(key, out var v) && v != null
                && v.Trim() == "@")
                return true;
        }
        return false;
    }

    // OOXML local-names already mapped to canonical Format keys by the curated
    // Font reader. Skip in the long-tail fallback so we don't double-emit
    // (e.g. avoid `font.b: "1"` alongside `font.bold: true`).
    private static readonly System.Collections.Generic.HashSet<string> CuratedFontChildren =
        new(System.StringComparer.Ordinal)
    {
        "b", "i", "strike", "u", "vertAlign", "sz", "name", "color",
    };

    // CT_CellAlignment curated attribute set (handled by the alignment Get
    // reader above). Long-tail = anything else (justifyLastLine, relativeIndent).
    private static readonly System.Collections.Generic.HashSet<string> CuratedAlignmentAttrs =
        new(System.StringComparer.Ordinal)
    {
        "horizontal", "vertical", "wrapText", "textRotation",
        "indent", "shrinkToFit", "readingOrder",
    };

    // CT_CellProtection curated attribute set.
    private static readonly System.Collections.Generic.HashSet<string> CuratedProtectionAttrs =
        new(System.StringComparer.Ordinal)
    {
        "locked", "hidden",
    };

    // CT_Col curated attribute set (handled by the column Get reader).
    private static readonly System.Collections.Generic.HashSet<string> CuratedColAttrs =
        new(System.StringComparer.Ordinal)
    {
        "min", "max", "width", "hidden", "customWidth", "outlineLevel", "collapsed",
    };

    // CT_Row curated attribute set (handled by the row Get reader).
    private static readonly System.Collections.Generic.HashSet<string> CuratedRowAttrs =
        new(System.StringComparer.Ordinal)
    {
        "r", "ht", "height", "hidden", "outlineLevel", "collapsed",
        "customHeight", "spans",
    };

    // Long-tail OOXML fallback for sub-elements with rich child structure
    // (Font: `<charset val="1"/>`, `<family val="2"/>`, ...). Mirrors Word's
    // FillUnknownChildProps but emits keys with a dotted prefix
    // (`font.charset`) so they slot into Excel's existing canonical scheme.
    private static void FillUnknownDottedProps(DocumentFormat.OpenXml.OpenXmlElement? container,
        DocumentNode node, string prefix, System.Collections.Generic.HashSet<string> curatedNames)
    {
        if (container == null) return;
        foreach (var child in container.ChildElements)
        {
            var name = child.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (curatedNames.Contains(name)) continue;
            var key = prefix + name;
            if (node.Format.ContainsKey(key)) continue;
            if (child.ChildElements.Count > 0) continue;

            string? valAttr = null;
            int attrCount = 0;
            foreach (var a in child.GetAttributes())
            {
                attrCount++;
                if (a.LocalName.Equals("val", System.StringComparison.OrdinalIgnoreCase))
                    valAttr = a.Value;
            }
            if (valAttr != null) node.Format[key] = valAttr;
            else if (attrCount == 0) node.Format[key] = true;
        }
    }

    // Long-tail OOXML fallback for attribute-only elements (Alignment,
    // Protection — CT_CellAlignment / CT_CellProtection). Walks attributes
    // on the element itself, prefix-qualifying each.
    private static void FillUnknownAttrProps(DocumentFormat.OpenXml.OpenXmlElement? element,
        DocumentNode node, string prefix, System.Collections.Generic.HashSet<string> curatedNames)
    {
        if (element == null) return;
        foreach (var attr in element.GetAttributes())
        {
            var name = attr.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (curatedNames.Contains(name)) continue;
            var key = prefix + name;
            if (node.Format.ContainsKey(key)) continue;
            node.Format[key] = attr.Value;
        }
    }
}
