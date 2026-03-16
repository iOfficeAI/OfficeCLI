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
        if (styleProps.TryGetValue("numfmt", out var numFmtStr))
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
        if (fontProps.Count > 0)
        {
            fontId = GetOrCreateFont(stylesheet, fontId, fontProps);
            applyFont = true;
        }

        // --- fill ---
        uint fillId = baseXf.FillId?.Value ?? 0;
        bool applyFill = baseXf.ApplyFill?.Value ?? false;
        if (styleProps.TryGetValue("fill", out var fillColor))
        {
            fillId = GetOrCreateFill(stylesheet, fillColor);
            applyFill = true;
        }

        // --- border (keep existing) ---
        uint borderId = baseXf.BorderId?.Value ?? 0;

        // --- alignment ---
        Alignment? alignment = baseXf.Alignment?.CloneNode(true) as Alignment;
        bool applyAlignment = baseXf.ApplyAlignment?.Value ?? false;
        var alignProps = styleProps
            .Where(kv => kv.Key.StartsWith("alignment.", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key[10..].ToLowerInvariant(), kv => kv.Value);
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
                }
            }
            applyAlignment = true;
        }

        // --- find or create CellFormat ---
        uint xfIndex = FindOrCreateCellFormat(cellFormats,
            numFmtId, fontId, fillId, borderId, alignment,
            applyNumFmt, applyFont, applyFill, applyAlignment);

        stylesheet.Save();
        return xfIndex;
    }

    /// <summary>
    /// Identify which keys in a dictionary are style properties.
    /// </summary>
    public static bool IsStyleKey(string key)
    {
        var lower = key.ToLowerInvariant();
        return lower is "numfmt" or "fill"
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
            ? (uVal.ToLowerInvariant() is "double" ? "double" : (IsTruthy(uVal) ? "single" : null))
            : (baseFont.Underline != null ? "single" : null);
        double size = fontProps.TryGetValue("size", out var szVal) && double.TryParse(szVal, out var sz)
            ? sz : baseFont.FontSize?.Val?.Value ?? 11;
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

    // ==================== CellFormat ====================

    private static uint FindOrCreateCellFormat(CellFormats cellFormats,
        uint numFmtId, uint fontId, uint fillId, uint borderId, Alignment? alignment,
        bool applyNumFmt, bool applyFont, bool applyFill, bool applyAlignment)
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
               (a.WrapText?.Value ?? false) == (b.WrapText?.Value ?? false);
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
    {
        // Ensure AARRGGBB format
        hex = hex.TrimStart('#');
        if (hex.Length == 6) return "FF" + hex.ToUpperInvariant();
        return hex.ToUpperInvariant();
    }

    private static bool IsTruthy(string value) =>
        value.ToLowerInvariant() is "true" or "1" or "yes";

    private static HorizontalAlignmentValues? ParseHAlign(string value) =>
        value.ToLowerInvariant() switch
        {
            "left" => HorizontalAlignmentValues.Left,
            "center" => HorizontalAlignmentValues.Center,
            "right" => HorizontalAlignmentValues.Right,
            "justify" => HorizontalAlignmentValues.Justify,
            _ => null
        };

    private static VerticalAlignmentValues? ParseVAlign(string value) =>
        value.ToLowerInvariant() switch
        {
            "top" => VerticalAlignmentValues.Top,
            "center" => VerticalAlignmentValues.Center,
            "bottom" => VerticalAlignmentValues.Bottom,
            _ => null
        };
}
