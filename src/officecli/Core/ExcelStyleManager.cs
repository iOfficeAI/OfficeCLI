// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Manages Excel cell styles via generic key=value properties.
/// Handles auto-creation of WorkbookStylesPart and deduplication of style entries.
///
/// Supported style keys:
///   numFmt          - number format string (e.g. "0%", "0.00", '#,##0.00"元"')
///   font.bold       - true/false
///   font.italic     - true/false
///   font.strike     - true/false
///   font.underline  - true/false or single/double
///   font.color      - hex RGB (e.g. "FF0000")
///   font.size       - point size (e.g. "11")
///   font.name       - font family name (e.g. "Calibri")
///   fill            - hex RGB background color (e.g. "4472C4")
///   border.all           - shorthand for all four sides (thin/medium/thick/double/dashed/dotted/none)
///   border.left/right/top/bottom - individual side style
///   border.color         - hex RGB color for all borders
///   border.left.color, border.right.color, etc. - per-side color
///   border.diagonal      - diagonal border style
///   border.diagonal.color - diagonal border color
///   border.diagonalUp    - true/false
///   border.diagonalDown  - true/false
///   alignment.horizontal - left/center/right
///   alignment.vertical   - top/center/bottom
///   alignment.wrapText   - true/false
/// </summary>
public class ExcelStyleManager
{
    private readonly WorkbookPart _workbookPart;

    public ExcelStyleManager(WorkbookPart workbookPart)
    {
        _workbookPart = workbookPart;
    }

    /// <summary>
    /// Ensure WorkbookStylesPart exists and return it.
    /// Creates a minimal default stylesheet if none exists.
    /// </summary>
    public WorkbookStylesPart EnsureStylesPart()
    {
        var stylesPart = _workbookPart.WorkbookStylesPart;
        if (stylesPart == null)
        {
            stylesPart = _workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = CreateDefaultStylesheet();
        }
        return stylesPart;
    }

    /// <summary>
    /// Ensure a Stylesheet exists on the WorkbookStylesPart and return it (non-null).
    /// </summary>
    private Stylesheet EnsureStylesheet()
    {
        var part = EnsureStylesPart();
        part.Stylesheet ??= CreateDefaultStylesheet();
        return part.Stylesheet;
    }

    /// <summary>
    /// Apply style properties to a cell. Merges with any existing cell style.
    /// Returns the style index to assign to the cell.
    /// </summary>
    public uint ApplyStyle(Cell cell, Dictionary<string, string> styleProps)
    {
        // Normalize keys to lowercase for case-insensitive matching
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in styleProps) props[k] = v;
        styleProps = props;

        var stylesheet = EnsureStylesheet();
        uint currentStyleIndex = cell.StyleIndex?.Value ?? 0;

        var cellFormats = EnsureCellFormats(stylesheet);
        var baseXf = currentStyleIndex < (uint)cellFormats.Elements<CellFormat>().Count()
            ? (CellFormat)cellFormats.Elements<CellFormat>().ElementAt((int)currentStyleIndex)
            : new CellFormat();

        // --- numFmt ---
        uint numFmtId = baseXf.NumberFormatId?.Value ?? 0;
        bool applyNumFmt = baseXf.ApplyNumberFormat?.Value ?? false;
        if (styleProps.TryGetValue("numfmt", out var numFmtStr) || styleProps.TryGetValue("numberformat", out numFmtStr)
            || styleProps.TryGetValue("format", out numFmtStr))
        {
            numFmtId = GetOrCreateNumFmt(stylesheet, numFmtStr);
            applyNumFmt = true;
        }

        // --- font ---
        uint fontId = baseXf.FontId?.Value ?? 0;
        bool applyFont = baseXf.ApplyFont?.Value ?? false;
        var fontProps = styleProps
            .Where(kv => kv.Key.StartsWith("font.", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key[5..].ToLowerInvariant(), kv => kv.Value);
        // Map "font" shorthand to font.name
        if (styleProps.TryGetValue("font", out var fontShorthand))
            fontProps["name"] = fontShorthand;
        // Map shorthand keys (bold, italic, strike, underline) to font.* equivalents
        foreach (var shortKey in new[] { "bold", "italic", "strike", "underline" })
        {
            if (styleProps.TryGetValue(shortKey, out var shortVal))
                fontProps[shortKey] = shortVal;
        }
        if (fontProps.Count > 0)
        {
            fontId = GetOrCreateFont(stylesheet, fontId, fontProps);
            applyFont = true;
        }

        // --- fill ---
        uint fillId = baseXf.FillId?.Value ?? 0;
        bool applyFill = baseXf.ApplyFill?.Value ?? false;
        if (styleProps.TryGetValue("fill", out var fillColor) || styleProps.TryGetValue("bgcolor", out fillColor))
        {
            if (fillColor.Contains('-') || fillColor.Contains(';'))
            {
                // Gradient fill: "FF0000-0000FF[-90]" or "radial:FF0000-0000FF"
                // Also handles semicolon format from Get: "gradient;FF0000;0000FF;90"
                var dashFormat = fillColor.Contains(';')
                    ? fillColor.TrimStart("gradient;".ToCharArray()).Replace(';', '-')
                    : fillColor;
                fillId = GetOrCreateGradientFill(stylesheet, dashFormat);
            }
            else
            {
                fillId = GetOrCreateFill(stylesheet, fillColor);
            }
            applyFill = true;
        }

        // --- border ---
        uint borderId = baseXf.BorderId?.Value ?? 0;
        bool applyBorder = baseXf.ApplyBorder?.Value ?? false;
        var borderProps = styleProps
            .Where(kv => kv.Key.StartsWith("border.", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key[7..].ToLowerInvariant(), kv => kv.Value);
        // Support "border" (without dot) as shorthand for "border.all"
        if (styleProps.TryGetValue("border", out var borderShorthand))
        {
            borderProps["all"] = borderShorthand;
        }
        if (borderProps.Count > 0)
        {
            borderId = GetOrCreateBorder(stylesheet, borderId, borderProps);
            applyBorder = true;
        }

        // --- alignment ---
        Alignment? alignment = baseXf.Alignment?.CloneNode(true) as Alignment;
        bool applyAlignment = baseXf.ApplyAlignment?.Value ?? false;
        var alignProps = styleProps
            .Where(kv => kv.Key.StartsWith("alignment.", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key[10..].ToLowerInvariant(), kv => kv.Value);
        // Handle shorthands: "wrap" → "wraptext", "halign" → "horizontal", "valign" → "vertical"
        if (styleProps.TryGetValue("wrap", out var wrapVal))
            alignProps["wraptext"] = wrapVal;
        if (styleProps.TryGetValue("wraptext", out var wrapVal2))
            alignProps["wraptext"] = wrapVal2;
        if (styleProps.TryGetValue("halign", out var halignVal))
            alignProps["horizontal"] = halignVal;
        if (styleProps.TryGetValue("valign", out var valignVal))
            alignProps["vertical"] = valignVal;
        if (styleProps.TryGetValue("rotation", out var rotVal))
            alignProps["rotation"] = rotVal;
        if (styleProps.TryGetValue("indent", out var indVal))
            alignProps["indent"] = indVal;
        if (styleProps.TryGetValue("shrinktofit", out var shrinkVal))
            alignProps["shrinktofit"] = shrinkVal;
        if (alignProps.Count > 0)
        {
            alignment ??= new Alignment();
            foreach (var (key, value) in alignProps)
            {
                switch (key)
                {
                    case "horizontal":
                        alignment.Horizontal = ParseHAlign(value);
                        break;
                    case "vertical":
                        alignment.Vertical = ParseVAlign(value);
                        break;
                    case "wraptext":
                        alignment.WrapText = IsTruthy(value);
                        break;
                    case "rotation" or "textrotation":
                        alignment.TextRotation = ParseHelpers.SafeParseUint(value, "rotation");
                        break;
                    case "indent":
                        alignment.Indent = ParseHelpers.SafeParseUint(value, "indent");
                        break;
                    case "shrinktofit" or "shrink":
                        alignment.ShrinkToFit = IsTruthy(value);
                        break;
                }
            }
            applyAlignment = true;
        }

        // --- find or create CellFormat ---
        uint xfIndex = FindOrCreateCellFormat(cellFormats,
            numFmtId, fontId, fillId, borderId, alignment,
            applyNumFmt, applyFont, applyFill, applyBorder, applyAlignment);

        stylesheet.Save();
        return xfIndex;
    }

    /// <summary>
    /// Identify which keys in a dictionary are style properties.
    /// </summary>
    public static bool IsStyleKey(string key)
    {
        var lower = key.ToLowerInvariant();
        return lower is "numfmt" or "fill" or "bgcolor" or "font" or "border"
            or "bold" or "italic" or "strike" or "underline"
            or "wrap" or "wraptext" or "numberformat" or "format" or "halign" or "valign"
            or "rotation" or "indent" or "shrinktofit"
            || lower.StartsWith("font.")
            || lower.StartsWith("alignment.")
            || lower.StartsWith("border.");
    }

    // ==================== NumberFormat ====================

    private static uint GetOrCreateNumFmt(Stylesheet stylesheet, string formatCode)
    {
        // Check built-in formats
        var builtinMap = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase)
        {
            ["general"] = 0, ["0"] = 1, ["0.00"] = 2, ["#,##0"] = 3, ["#,##0.00"] = 4,
            ["0%"] = 9, ["0.00%"] = 10,
        };
        if (builtinMap.TryGetValue(formatCode, out var builtinId))
            return builtinId;

        // Check existing custom formats
        var numFmts = stylesheet.NumberingFormats;
        if (numFmts != null)
        {
            foreach (var nf in numFmts.Elements<NumberingFormat>())
            {
                if (nf.FormatCode?.Value == formatCode)
                    return nf.NumberFormatId?.Value ?? 164;
            }
        }

        // Create new (custom IDs start at 164)
        if (numFmts == null)
        {
            numFmts = new NumberingFormats { Count = 0 };
            stylesheet.InsertAt(numFmts, 0);
        }

        uint newId = 164;
        foreach (var nf in numFmts.Elements<NumberingFormat>())
        {
            if (nf.NumberFormatId?.Value >= newId)
                newId = nf.NumberFormatId.Value + 1;
        }

        numFmts.Append(new NumberingFormat { NumberFormatId = newId, FormatCode = formatCode });
        numFmts.Count = (uint)numFmts.Elements<NumberingFormat>().Count();

        return newId;
    }

    // ==================== Font ====================

    private static uint GetOrCreateFont(Stylesheet stylesheet, uint baseFontId, Dictionary<string, string> fontProps)
    {
        var fonts = stylesheet.Fonts;
        if (fonts == null)
        {
            fonts = new Fonts(
                new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })
            ) { Count = 1 };
            // Insert after NumberingFormats if present, otherwise at start
            var numFmts = stylesheet.NumberingFormats;
            if (numFmts != null)
                numFmts.InsertAfterSelf(fonts);
            else
                stylesheet.InsertAt(fonts, 0);
        }

        // Get base font to merge with
        var baseFont = baseFontId < (uint)fonts.Elements<Font>().Count()
            ? fonts.Elements<Font>().ElementAt((int)baseFontId)
            : fonts.Elements<Font>().First();

        // Build target properties (merge: new props override base)
        bool bold = fontProps.TryGetValue("bold", out var bVal)
            ? IsTruthy(bVal) : baseFont.Bold != null;
        bool italic = fontProps.TryGetValue("italic", out var iVal)
            ? IsTruthy(iVal) : baseFont.Italic != null;
        bool strike = fontProps.TryGetValue("strike", out var sVal)
            ? IsTruthy(sVal) : baseFont.Strike != null;
        string? underline = fontProps.TryGetValue("underline", out var uVal)
            ? (uVal.ToLowerInvariant() is "double" ? "double" : (IsTruthy(uVal) || uVal.ToLowerInvariant() == "single" ? "single" : null))
            : (baseFont.Underline != null ? (baseFont.Underline.Val?.InnerText == "double" ? "double" : "single") : null);
        double size;
        if (fontProps.TryGetValue("size", out var szVal))
        {
            if (!double.TryParse(szVal, out var sz) || double.IsNaN(sz) || double.IsInfinity(sz))
                throw new ArgumentException($"Invalid font.size value: '{szVal}'. Expected a finite number.");
            size = sz;
        }
        else
        {
            size = baseFont.FontSize?.Val?.Value ?? 11;
        }
        string name = fontProps.GetValueOrDefault("name",
            baseFont.FontName?.Val?.Value ?? "Calibri");
        string? color = fontProps.TryGetValue("color", out var cVal)
            ? NormalizeColor(cVal) : baseFont.Color?.Rgb?.Value;

        // Search for existing match
        int idx = 0;
        foreach (var f in fonts.Elements<Font>())
        {
            if (FontMatches(f, bold, italic, strike, underline, size, name, color))
                return (uint)idx;
            idx++;
        }

        // Create new font (element order matters: b, i, strike, u, sz, color, name)
        var newFont = new Font();
        if (bold) newFont.Append(new Bold());
        if (italic) newFont.Append(new Italic());
        if (strike) newFont.Append(new Strike());
        if (underline != null)
        {
            var ul = new Underline();
            if (underline == "double")
                ul.Val = UnderlineValues.Double;
            newFont.Append(ul);
        }
        newFont.Append(new FontSize { Val = size });
        if (color != null)
            newFont.Append(new Color { Rgb = color });
        newFont.Append(new FontName { Val = name });

        fonts.Append(newFont);
        fonts.Count = (uint)fonts.Elements<Font>().Count();

        return (uint)(fonts.Elements<Font>().Count() - 1);
    }

    private static bool FontMatches(Font font, bool bold, bool italic, bool strike,
        string? underline, double size, string name, string? color)
    {
        if ((font.Bold != null) != bold) return false;
        if ((font.Italic != null) != italic) return false;
        if ((font.Strike != null) != strike) return false;
        if ((font.Underline != null) != (underline != null)) return false;
        if (font.Underline != null && underline != null)
        {
            var fontUlType = font.Underline.Val?.InnerText == "double" ? "double" : "single";
            if (fontUlType != underline) return false;
        }
        if (Math.Abs((font.FontSize?.Val?.Value ?? 11) - size) > 0.01) return false;
        if (!string.Equals(font.FontName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase)) return false;

        var fontColor = font.Color?.Rgb?.Value;
        if (color != null)
        {
            if (!string.Equals(fontColor, color, StringComparison.OrdinalIgnoreCase)) return false;
        }
        else if (fontColor != null) return false;

        return true;
    }

    // ==================== Fill ====================

    private static uint GetOrCreateFill(Stylesheet stylesheet, string hexColor)
    {
        var fills = stylesheet.Fills;
        if (fills == null)
        {
            fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            ) { Count = 2 };
            // Insert after Fonts
            var fonts = stylesheet.Fonts;
            if (fonts != null)
                fonts.InsertAfterSelf(fills);
            else
                stylesheet.Append(fills);
        }

        var normalizedColor = NormalizeColor(hexColor);

        // Search for existing match
        int idx = 0;
        foreach (var fill in fills.Elements<Fill>())
        {
            var pf = fill.PatternFill;
            if (pf?.PatternType?.Value == PatternValues.Solid &&
                string.Equals(pf.ForegroundColor?.Rgb?.Value, normalizedColor, StringComparison.OrdinalIgnoreCase))
                return (uint)idx;
            idx++;
        }

        // Create new fill
        fills.Append(new Fill(new PatternFill(
            new ForegroundColor { Rgb = normalizedColor }
        ) { PatternType = PatternValues.Solid }));
        fills.Count = (uint)fills.Elements<Fill>().Count();

        return (uint)(fills.Elements<Fill>().Count() - 1);
    }

    /// <summary>
    /// Create or find a gradient fill entry in the stylesheet.
    /// Format: "C1-C2[-angle]" (linear) or "radial:C1-C2" (radial).
    /// Reuses same parsing logic as PPTX gradient but outputs Spreadsheet.GradientFill.
    /// </summary>
    private static uint GetOrCreateGradientFill(Stylesheet stylesheet, string value)
    {
        var fills = stylesheet.Fills;
        if (fills == null)
        {
            fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            ) { Count = 2 };
            var fonts = stylesheet.Fonts;
            if (fonts != null) fonts.InsertAfterSelf(fills);
            else stylesheet.Append(fills);
        }

        // Parse gradient spec
        string gradType = "linear";
        string colorSpec = value;
        if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
        {
            gradType = "path";
            colorSpec = value[7..];
        }

        var parts = colorSpec.Split('-');
        var colors = parts.ToList();
        double degree = 90; // default top-to-bottom

        if (gradType == "linear" && colors.Count >= 2 &&
            double.TryParse(colors.Last(), System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var angleDeg) &&
            colors.Last().Length <= 3)
        {
            degree = angleDeg;
            colors.RemoveAt(colors.Count - 1);
        }

        if (colors.Count < 2) colors.Add(colors[0]);

        // Normalize colors
        for (int i = 0; i < colors.Count; i++)
            colors[i] = NormalizeColor(colors[i]);

        // Search for existing match
        int idx = 0;
        foreach (var existingFill in fills.Elements<Fill>())
        {
            var gf = existingFill.GetFirstChild<GradientFill>();
            if (gf != null)
            {
                var stops = gf.Elements<GradientStop>().ToList();
                if (stops.Count == colors.Count)
                {
                    bool match = true;
                    for (int i = 0; i < stops.Count; i++)
                    {
                        var stopColor = stops[i].Color?.Rgb?.Value;
                        if (!string.Equals(stopColor, colors[i], StringComparison.OrdinalIgnoreCase))
                        { match = false; break; }
                    }
                    if (match && Math.Abs((gf.Degree?.Value ?? 0) - degree) < 0.1)
                        return (uint)idx;
                }
            }
            idx++;
        }

        // Create new gradient fill
        var gradFill = new GradientFill();
        if (gradType == "path")
            gradFill.Type = GradientValues.Path;
        else
            gradFill.Degree = degree;

        for (int i = 0; i < colors.Count; i++)
        {
            double pos = colors.Count == 1 ? 0 : (double)i / (colors.Count - 1);
            gradFill.Append(new GradientStop(
                new Color { Rgb = new HexBinaryValue(colors[i]) }
            ) { Position = pos });
        }

        fills.Append(new Fill(gradFill));
        fills.Count = (uint)fills.Elements<Fill>().Count();
        return (uint)(fills.Elements<Fill>().Count() - 1);
    }

    // ==================== Border ====================

    private static uint GetOrCreateBorder(Stylesheet stylesheet, uint baseBorderId, Dictionary<string, string> borderProps)
    {
        var borders = stylesheet.Borders;
        if (borders == null)
        {
            borders = new Borders(
                new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())
            ) { Count = 1 };
            var fills = stylesheet.Fills;
            if (fills != null)
                fills.InsertAfterSelf(borders);
            else
                stylesheet.Append(borders);
        }

        // Get base border to merge with
        var baseBorder = baseBorderId < (uint)borders.Elements<Border>().Count()
            ? borders.Elements<Border>().ElementAt((int)baseBorderId)
            : borders.Elements<Border>().First();

        // Resolve styles: start from base, override with new props
        var leftStyle = baseBorder.LeftBorder?.Style?.Value ?? BorderStyleValues.None;
        var rightStyle = baseBorder.RightBorder?.Style?.Value ?? BorderStyleValues.None;
        var topStyle = baseBorder.TopBorder?.Style?.Value ?? BorderStyleValues.None;
        var bottomStyle = baseBorder.BottomBorder?.Style?.Value ?? BorderStyleValues.None;
        var diagonalStyle = baseBorder.DiagonalBorder?.Style?.Value ?? BorderStyleValues.None;

        string? leftColor = baseBorder.LeftBorder?.Color?.Rgb?.Value;
        string? rightColor = baseBorder.RightBorder?.Color?.Rgb?.Value;
        string? topColor = baseBorder.TopBorder?.Color?.Rgb?.Value;
        string? bottomColor = baseBorder.BottomBorder?.Color?.Rgb?.Value;
        string? diagonalColor = baseBorder.DiagonalBorder?.Color?.Rgb?.Value;

        bool diagonalUp = baseBorder.DiagonalUp?.Value ?? false;
        bool diagonalDown = baseBorder.DiagonalDown?.Value ?? false;

        // Apply "all" shorthand first (individual sides override later)
        if (borderProps.TryGetValue("all", out var allStyle))
        {
            var parsed = ParseBorderStyle(allStyle);
            leftStyle = rightStyle = topStyle = bottomStyle = parsed;
        }

        // Apply "color" shorthand
        if (borderProps.TryGetValue("color", out var allColor))
        {
            var normalized = NormalizeColor(allColor);
            leftColor = rightColor = topColor = bottomColor = normalized;
        }

        // Apply individual side styles
        if (borderProps.TryGetValue("left", out var lVal)) leftStyle = ParseBorderStyle(lVal);
        if (borderProps.TryGetValue("right", out var rVal)) rightStyle = ParseBorderStyle(rVal);
        if (borderProps.TryGetValue("top", out var tVal)) topStyle = ParseBorderStyle(tVal);
        if (borderProps.TryGetValue("bottom", out var bVal)) bottomStyle = ParseBorderStyle(bVal);
        if (borderProps.TryGetValue("diagonal", out var dVal)) diagonalStyle = ParseBorderStyle(dVal);

        // Apply individual side colors
        if (borderProps.TryGetValue("left.color", out var lcVal)) leftColor = NormalizeColor(lcVal);
        if (borderProps.TryGetValue("right.color", out var rcVal)) rightColor = NormalizeColor(rcVal);
        if (borderProps.TryGetValue("top.color", out var tcVal)) topColor = NormalizeColor(tcVal);
        if (borderProps.TryGetValue("bottom.color", out var bcVal)) bottomColor = NormalizeColor(bcVal);
        if (borderProps.TryGetValue("diagonal.color", out var dcVal)) diagonalColor = NormalizeColor(dcVal);

        // Diagonal direction flags
        if (borderProps.TryGetValue("diagonalup", out var duVal)) diagonalUp = IsTruthy(duVal);
        if (borderProps.TryGetValue("diagonaldown", out var ddVal)) diagonalDown = IsTruthy(ddVal);

        // Search for existing match
        int idx = 0;
        foreach (var b in borders.Elements<Border>())
        {
            if (BorderMatches(b, leftStyle, rightStyle, topStyle, bottomStyle, diagonalStyle,
                leftColor, rightColor, topColor, bottomColor, diagonalColor,
                diagonalUp, diagonalDown))
                return (uint)idx;
            idx++;
        }

        // Create new border
        var newBorder = new Border();

        newBorder.Append(CreateBorderElement<LeftBorder>(leftStyle, leftColor));
        newBorder.Append(CreateBorderElement<RightBorder>(rightStyle, rightColor));
        newBorder.Append(CreateBorderElement<TopBorder>(topStyle, topColor));
        newBorder.Append(CreateBorderElement<BottomBorder>(bottomStyle, bottomColor));
        newBorder.Append(CreateBorderElement<DiagonalBorder>(diagonalStyle, diagonalColor));

        if (diagonalUp) newBorder.DiagonalUp = true;
        if (diagonalDown) newBorder.DiagonalDown = true;

        borders.Append(newBorder);
        borders.Count = (uint)borders.Elements<Border>().Count();

        return (uint)(borders.Elements<Border>().Count() - 1);
    }

    private static T CreateBorderElement<T>(BorderStyleValues style, string? color) where T : BorderPropertiesType, new()
    {
        var element = new T();
        if (style != BorderStyleValues.None)
        {
            element.Style = style;
            if (color != null)
                element.Color = new Color { Rgb = color };
        }
        return element;
    }

    private static bool BorderMatches(Border border,
        BorderStyleValues leftStyle, BorderStyleValues rightStyle,
        BorderStyleValues topStyle, BorderStyleValues bottomStyle,
        BorderStyleValues diagonalStyle,
        string? leftColor, string? rightColor,
        string? topColor, string? bottomColor, string? diagonalColor,
        bool diagonalUp, bool diagonalDown)
    {
        if (!BorderSideMatches(border.LeftBorder, leftStyle, leftColor)) return false;
        if (!BorderSideMatches(border.RightBorder, rightStyle, rightColor)) return false;
        if (!BorderSideMatches(border.TopBorder, topStyle, topColor)) return false;
        if (!BorderSideMatches(border.BottomBorder, bottomStyle, bottomColor)) return false;
        if (!BorderSideMatches(border.DiagonalBorder, diagonalStyle, diagonalColor)) return false;
        if ((border.DiagonalUp?.Value ?? false) != diagonalUp) return false;
        if ((border.DiagonalDown?.Value ?? false) != diagonalDown) return false;
        return true;
    }

    private static bool BorderSideMatches(BorderPropertiesType? side, BorderStyleValues style, string? color)
    {
        var sideStyle = side?.Style?.Value ?? BorderStyleValues.None;
        if (sideStyle != style) return false;
        var sideColor = side?.Color?.Rgb?.Value;
        if (color != null)
        {
            if (!string.Equals(sideColor, color, StringComparison.OrdinalIgnoreCase)) return false;
        }
        else if (sideColor != null) return false;
        return true;
    }

    private static BorderStyleValues ParseBorderStyle(string value) =>
        value.ToLowerInvariant() switch
        {
            "thin" => BorderStyleValues.Thin,
            "medium" => BorderStyleValues.Medium,
            "thick" => BorderStyleValues.Thick,
            "double" => BorderStyleValues.Double,
            "dashed" => BorderStyleValues.Dashed,
            "dotted" => BorderStyleValues.Dotted,
            "dashdot" => BorderStyleValues.DashDot,
            "dashdotdot" => BorderStyleValues.DashDotDot,
            "hair" => BorderStyleValues.Hair,
            "mediumdashed" => BorderStyleValues.MediumDashed,
            "mediumdashdot" => BorderStyleValues.MediumDashDot,
            "mediumdashdotdot" => BorderStyleValues.MediumDashDotDot,
            "slantdashdot" => BorderStyleValues.SlantDashDot,
            "none" => BorderStyleValues.None,
            _ => throw new ArgumentException($"Invalid border style: '{value}'. Valid values: thin, medium, thick, double, dashed, dotted, dashdot, dashdotdot, hair, mediumdashed, mediumdashdot, mediumdashdotdot, slantdashdot, none."),
        };

    // ==================== CellFormat ====================

    private static uint FindOrCreateCellFormat(CellFormats cellFormats,
        uint numFmtId, uint fontId, uint fillId, uint borderId, Alignment? alignment,
        bool applyNumFmt, bool applyFont, bool applyFill, bool applyBorder, bool applyAlignment)
    {
        // Search for existing match
        int idx = 0;
        foreach (var xf in cellFormats.Elements<CellFormat>())
        {
            if ((xf.NumberFormatId?.Value ?? 0) == numFmtId &&
                (xf.FontId?.Value ?? 0) == fontId &&
                (xf.FillId?.Value ?? 0) == fillId &&
                (xf.BorderId?.Value ?? 0) == borderId &&
                AlignmentMatches(xf.Alignment, alignment))
                return (uint)idx;
            idx++;
        }

        // Create new CellFormat
        var newXf = new CellFormat
        {
            NumberFormatId = numFmtId,
            FontId = fontId,
            FillId = fillId,
            BorderId = borderId
        };
        if (applyNumFmt) newXf.ApplyNumberFormat = true;
        if (applyFont) newXf.ApplyFont = true;
        if (applyFill) newXf.ApplyFill = true;
        if (applyBorder) newXf.ApplyBorder = true;
        if (applyAlignment && alignment != null)
        {
            newXf.ApplyAlignment = true;
            newXf.Append(alignment);
        }

        cellFormats.Append(newXf);
        cellFormats.Count = (uint)cellFormats.Elements<CellFormat>().Count();

        return (uint)(cellFormats.Elements<CellFormat>().Count() - 1);
    }

    private static bool AlignmentMatches(Alignment? a, Alignment? b)
    {
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;
        return a.Horizontal?.Value == b.Horizontal?.Value &&
               a.Vertical?.Value == b.Vertical?.Value &&
               (a.WrapText?.Value ?? false) == (b.WrapText?.Value ?? false) &&
               (a.TextRotation?.Value ?? 0) == (b.TextRotation?.Value ?? 0) &&
               (a.Indent?.Value ?? 0) == (b.Indent?.Value ?? 0) &&
               (a.ShrinkToFit?.Value ?? false) == (b.ShrinkToFit?.Value ?? false);
    }

    // ==================== Helpers ====================

    private static Stylesheet CreateDefaultStylesheet()
    {
        return new Stylesheet(
            new NumberingFormats() { Count = 0 },
            new Fonts(
                new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })
            ) { Count = 1 },
            new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            ) { Count = 2 },
            new Borders(
                new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())
            ) { Count = 1 },
            new CellStyleFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
            ) { Count = 1 },
            new CellFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
            ) { Count = 1 },
            new CellStyles(
                new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }
            ) { Count = 1 }
        );
    }

    private static CellFormats EnsureCellFormats(Stylesheet stylesheet)
    {
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats == null)
        {
            cellFormats = new CellFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
            ) { Count = 1 };
            stylesheet.Append(cellFormats);
        }
        return cellFormats;
    }

    private static string NormalizeColor(string hex)
        => ParseHelpers.NormalizeArgbColor(hex);

    private static bool IsTruthy(string value) =>
        ParseHelpers.IsTruthy(value);

    private static HorizontalAlignmentValues ParseHAlign(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" => HorizontalAlignmentValues.Left,
            "center" => HorizontalAlignmentValues.Center,
            "right" => HorizontalAlignmentValues.Right,
            "justify" => HorizontalAlignmentValues.Justify,
            _ => throw new ArgumentException($"Invalid horizontal alignment: '{value}'. Valid values: left, center, right, justify.")
        };

    private static VerticalAlignmentValues ParseVAlign(string value) =>
        value.ToLowerInvariant() switch
        {
            "top" => VerticalAlignmentValues.Top,
            "center" => VerticalAlignmentValues.Center,
            "bottom" => VerticalAlignmentValues.Bottom,
            _ => throw new ArgumentException($"Invalid vertical alignment: '{value}'. Valid values: top, center, bottom.")
        };
}
