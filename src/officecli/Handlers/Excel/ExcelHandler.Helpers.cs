// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    /// CONSISTENCY(axis-ref-compat): Excel-style whole-column (`B:B`, `A:C`)
    /// and whole-row (`1:1`, `2:5`) references are accepted as lenient input
    /// aliases for the canonical col[X]/row[N] path segments — same policy as
    /// A1-style ranges (`A1:D10`), which the path grammar already accepts.
    /// Returns the expanded canonical segments in axis order (one per
    /// column/row in the span), or null when the segment is not an axis
    /// reference. Canonical readback paths remain col[X]/row[N].
    /// Spans wider than 1024 are rejected — a set over B:XFD (16k columns)
    /// is almost certainly a mistake, not intent.
    /// </summary>
    private static List<string>? TryExpandAxisRef(string cellRef)
    {
        var colRef = System.Text.RegularExpressions.Regex.Match(
            cellRef, @"^([A-Za-z]{1,3}):([A-Za-z]{1,3})$");
        if (colRef.Success)
        {
            var from = ColumnNameToIndex(colRef.Groups[1].Value.ToUpperInvariant());
            var to = ColumnNameToIndex(colRef.Groups[2].Value.ToUpperInvariant());
            if (from > to) (from, to) = (to, from);
            if (to - from + 1 > 1024)
                throw new ArgumentException(
                    $"Column span {cellRef} covers {to - from + 1} columns (limit 1024). Narrow the span or address columns individually as col[X].");
            var cols = new List<string>();
            for (var i = from; i <= to; i++) cols.Add($"col[{IndexToColumnName(i)}]");
            return cols;
        }
        var rowRef = System.Text.RegularExpressions.Regex.Match(cellRef, @"^(\d+):(\d+)$");
        if (rowRef.Success
            && uint.TryParse(rowRef.Groups[1].Value, out var r1) && r1 >= 1
            && uint.TryParse(rowRef.Groups[2].Value, out var r2) && r2 >= 1)
        {
            if (r1 > r2) (r1, r2) = (r2, r1);
            if (r2 - r1 + 1 > 1024)
                throw new ArgumentException(
                    $"Row span {cellRef} covers {r2 - r1 + 1} rows (limit 1024). Narrow the span or address rows individually as row[N].");
            var rows = new List<string>();
            for (var i = r1; i <= r2; i++) rows.Add($"row[{i}]");
            return rows;
        }
        return null;
    }

    /// <summary>
    /// Parse a print-margin value into inches (PageMargins schema unit).
    /// Accepts "1in", "2.5cm", "1.27cm", "72pt", "10mm", or a bare number (inches).
    /// </summary>
    internal static double ParseMarginInches(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Invalid margin: empty value.");
        var v = value.Trim().ToLowerInvariant();
        double num;
        if (v.EndsWith("in"))
        {
            num = double.Parse(v[..^2].Trim(), System.Globalization.CultureInfo.InvariantCulture);
            return num;
        }
        if (v.EndsWith("cm"))
        {
            num = double.Parse(v[..^2].Trim(), System.Globalization.CultureInfo.InvariantCulture);
            return num / 2.54;
        }
        if (v.EndsWith("mm"))
        {
            num = double.Parse(v[..^2].Trim(), System.Globalization.CultureInfo.InvariantCulture);
            return num / 25.4;
        }
        if (v.EndsWith("pt"))
        {
            num = double.Parse(v[..^2].Trim(), System.Globalization.CultureInfo.InvariantCulture);
            return num / 72.0;
        }
        // Bare number = inches
        if (!double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out num))
            throw new ArgumentException($"Invalid margin value: '{value}' (use 1in, 2cm, 10mm, 72pt, or bare inches)");
        return num;
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
                pts = emu / EmuConverter.EmuPerPointF;
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
                var px = emu / EmuConverter.EmuPerPxF;
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

    // ==================== Conditional Formatting Helpers ====================

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    private static bool IsValidBooleanString(string? value) =>
        ParseHelpers.IsValidBooleanString(value);

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
            @"^#?(?:'(?:[^']|'')+'|[A-Za-z_][\w\.]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    internal static string? TryParseInternalHyperlinkLocation(string value)
    {
        if (string.IsNullOrEmpty(value)) return null;
        if (!s_internalLinkPattern.IsMatch(value)) return null;
        return value.StartsWith("#") ? value.Substring(1) : value;
    }

    /// <summary>
    /// Instance-aware internal-hyperlink resolution: everything
    /// <see cref="TryParseInternalHyperlinkLocation"/> accepts, plus a bare
    /// token that names an existing workbook defined name. Without the
    /// defined-name check, link=MyRange fell through to the external-URL
    /// path and produced a relationship targeting the literal string
    /// "MyRange" — a hyperlink that tries to open a file of that name.
    /// A '#'-prefixed defined name (#MyRange) also resolves here.
    /// </summary>
    private string? ResolveInternalHyperlinkLocation(string value)
    {
        var loc = TryParseInternalHyperlinkLocation(value);
        if (loc != null) return loc;
        if (string.IsNullOrEmpty(value)) return null;
        var bare = value.StartsWith('#') ? value[1..] : value;
        if (!System.Text.RegularExpressions.Regex.IsMatch(bare, @"^[\p{L}_\\][\p{L}\p{N}_.\\]*$"))
            return null;
        var isDefinedName = GetWorkbook().GetFirstChild<DefinedNames>()?
            .Elements<DefinedName>()
            .Any(d => d.Name?.Value?.Equals(bare, StringComparison.OrdinalIgnoreCase) == true) == true;
        return isDefinedName ? bare : null;
    }

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
            // OpenXmlMiscNode (XML comments, processing instructions, CDATA)
            // surfaces with synthetic LocalNames like "#comment" / "#text". They
            // are not OOXML elements and must not appear as Format keys.
            if (name.StartsWith("#") || child is DocumentFormat.OpenXml.OpenXmlMiscNode) continue;
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
