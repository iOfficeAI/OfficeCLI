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
internal class ExcelStyleManager
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
    public uint ApplyStyle(Cell cell, Dictionary<string, string> styleProps, List<string>? unsupportedOut = null)
    {
        // Normalize keys to lowercase for case-insensitive matching; skip null values
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in styleProps) if (v != null) props[k] = v;
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
        // Map shorthand keys (bold, italic, strike, underline, superscript, subscript, strikethrough, size) to font.* equivalents
        foreach (var shortKey in new[] { "bold", "italic", "strike", "underline", "superscript", "subscript", "strikethrough", "size" })
        {
            if (styleProps.TryGetValue(shortKey, out var shortVal))
                fontProps[shortKey == "strikethrough" ? "strike" : shortKey] = shortVal;
        }
        // Normalize "strikethrough" alias within font.* props
        if (fontProps.Remove("strikethrough", out var stVal))
            fontProps["strike"] = stVal;
        if (fontProps.Count > 0)
        {
            // Split into curated (handled by GetOrCreateFont's typed builder)
            // and long-tail (raw OOXML children appended via SDK schema-aware
            // AddChild, force-new in the dedup table).
            var longTailFontProps = fontProps
                .Where(kv => !CuratedFontKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            var curatedFontProps = fontProps
                .Where(kv => CuratedFontKeys.Contains(kv.Key))
                .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            // Preserve baseFont's existing long-tail children (charset, family,
            // outline, shadow, ...) — without this they'd be silently dropped
            // every time the cell's style is touched, since GetOrCreateFont
            // rebuilds Font from curated fields only.
            var baseFonts = stylesheet.Fonts;
            if (baseFonts != null && fontId < (uint)baseFonts.Elements<Font>().Count())
            {
                var baseFont = baseFonts.Elements<Font>().ElementAt((int)fontId);
                foreach (var child in baseFont.ChildElements)
                {
                    var name = child.LocalName;
                    if (CuratedFontChildLocalNames.Contains(name)) continue;
                    if (longTailFontProps.ContainsKey(name)) continue; // caller wins
                    string? valStr = null;
                    foreach (var a in child.GetAttributes())
                    {
                        if (a.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        { valStr = a.Value; break; }
                    }
                    if (valStr != null)
                        longTailFontProps[name] = valStr;
                }
            }
            fontId = GetOrCreateFont(stylesheet, fontId, curatedFontProps, longTailFontProps, unsupportedOut);
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
        // DEFERRED(xlsx/cell-reading-order) CE10: accept top-level `readingOrder`
        // as shorthand for `alignment.readingOrder`.
        if (styleProps.TryGetValue("readingorder", out var roVal))
            alignProps["readingorder"] = roVal;
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
                    {
                        // R39-3: OOXML §18.18.20 ST_TextRotation — valid values
                        // are 0..180 (degrees) plus the special sentinel 255
                        // (vertical text stack). Excel rejects 181..254 and
                        // anything above 255. R15 added the lower-bound guard
                        // (negative parsing throws via SafeParseUint), but the
                        // upper bound was missing, allowing files Excel later
                        // refuses to open.
                        var rot = ParseHelpers.SafeParseUint(value, "rotation");
                        if (!(rot <= 180 || rot == 255))
                            throw new ArgumentException(
                                $"Invalid 'rotation' value '{value}'. Must be 0..180 (degrees) or 255 (vertical stack).");
                        alignment.TextRotation = rot;
                        break;
                    }
                    case "indent":
                        alignment.Indent = ParseHelpers.SafeParseUint(value, "indent");
                        break;
                    case "shrinktofit" or "shrink":
                        alignment.ShrinkToFit = IsTruthy(value);
                        break;
                    case "readingorder":
                        // DEFERRED(xlsx/cell-reading-order) CE10: OOXML values
                        // 0=context, 1=ltr, 2=rtl. Accept numeric or string forms.
                        alignment.ReadingOrder = ParseReadingOrder(value);
                        break;
                    // Long-tail keys handled below (case-preserving) — see the
                    // styleProps walk after this loop. Skip in the lowered
                    // switch to avoid double-write.
                }
            }
            // Long-tail Alignment attributes (e.g. justifyLastLine,
            // relativeIndent). Walk styleProps directly to preserve original
            // case — OOXML attribute names are case-sensitive (Excel rejects
            // `justifylastline`, only accepts `justifyLastLine`).
            foreach (var (origKey, value) in styleProps)
            {
                if (!origKey.StartsWith("alignment.", StringComparison.OrdinalIgnoreCase)) continue;
                var subKey = origKey.Substring(10); // preserve case after "alignment."
                if (CuratedAlignmentSubKeysLower.Contains(subKey.ToLowerInvariant())) continue;
                alignment.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", subKey, "", value));
            }
            applyAlignment = true;
        }

        // --- quotePrefix ---
        // R28-B4 — quotePrefix=true marks the cell xf so Excel renders the
        // value literally (force-text). Used when the cell value starts with
        // a leading apostrophe; the apostrophe is stripped from the value
        // and quotePrefix carries the "force text" intent in the style.
        bool? quotePrefix = baseXf.QuotePrefix?.Value;
        if (styleProps.TryGetValue("quoteprefix", out var qpVal))
            quotePrefix = IsTruthy(qpVal);

        // --- protection ---
        Protection? protection = baseXf.Protection?.CloneNode(true) as Protection;
        bool applyProtection = baseXf.ApplyProtection?.Value ?? false;
        var protectionLongTail = styleProps
            .Where(kv => kv.Key.StartsWith("protection.", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(kv => kv.Key[11..].ToLowerInvariant(), kv => kv.Value);
        if (styleProps.TryGetValue("locked", out var lockedVal) ||
            styleProps.TryGetValue("formulahidden", out var fhVal) ||
            protectionLongTail.Count > 0)
        {
            protection ??= new Protection();
            if (styleProps.TryGetValue("locked", out var lv))
                protection.Locked = IsTruthy(lv);
            if (styleProps.TryGetValue("formulahidden", out var fv))
                protection.Hidden = IsTruthy(fv);
            // protection.locked and protection.hidden as canonical dotted keys
            // (mirror Get's `protection.locked` / `protection.hidden` output).
            if (protectionLongTail.TryGetValue("locked", out var pLocked))
                protection.Locked = IsTruthy(pLocked);
            if (protectionLongTail.TryGetValue("hidden", out var pHidden))
                protection.Hidden = IsTruthy(pHidden);
            // Anything else under protection.* is a raw long-tail attribute on
            // the Protection element. CT_CellProtection only has locked/hidden
            // today, but stay symmetric with Get's fallback if the schema grows.
            // Walk styleProps directly to preserve original case — OOXML
            // attributes are case-sensitive.
            foreach (var (origKey, value) in styleProps)
            {
                if (!origKey.StartsWith("protection.", StringComparison.OrdinalIgnoreCase)) continue;
                var subKey = origKey.Substring(11);
                if (subKey.Equals("locked", StringComparison.OrdinalIgnoreCase)) continue;
                if (subKey.Equals("hidden", StringComparison.OrdinalIgnoreCase)) continue;
                protection.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", subKey, "", value));
            }
            applyProtection = true;
        }

        // --- find or create CellFormat ---
        uint xfIndex = FindOrCreateCellFormat(cellFormats,
            numFmtId, fontId, fillId, borderId, alignment, protection,
            applyNumFmt, applyFont, applyFill, applyBorder, applyAlignment, applyProtection,
            quotePrefix);

        // Caller (ExcelHandler) is responsible for saving via _dirtyStylesheet flag.
        return xfIndex;
    }

    /// <summary>
    /// Ensure the workbook has the built-in "Hyperlink" cellStyle (builtinId=8)
    /// wired up with a blue underlined font, and return the cellXfs index that
    /// hyperlink cells should reference via `c/@s`.
    ///
    /// Creates (idempotently):
    ///   - a Font with color 0563C1 + underline
    ///   - a CellStyleFormats xf referencing that font (applyFont=true)
    ///   - a CellFormats xf inheriting from the cellStyleXf (xfId, applyFont=true)
    ///   - a CellStyles entry Name="Hyperlink" BuiltinId=8 pointing at the cellStyleXf
    ///
    /// Returns the cellXfs index to assign to the cell's StyleIndex.
    /// </summary>
    public uint EnsureHyperlinkCellStyle()
    {
        var stylesheet = EnsureStylesheet();

        // 1. Reuse existing "Hyperlink" cellStyle if already present.
        var cellStyles = stylesheet.CellStyles;
        if (cellStyles != null)
        {
            var existing = cellStyles.Elements<CellStyle>()
                .FirstOrDefault(cs => cs.BuiltinId?.Value == 8u);
            if (existing?.FormatId?.Value != null)
            {
                // FormatId is the cellStyleXfs index. Find a cellXfs that
                // references that cellStyleXf via xfId; if none, create one.
                uint styleXfId = existing.FormatId.Value;
                var cellFormats = EnsureCellFormats(stylesheet);
                int cIdx = 0;
                foreach (var xf in cellFormats.Elements<CellFormat>())
                {
                    if (xf.FormatId?.Value == styleXfId
                        && (xf.ApplyFont?.Value ?? false))
                        return (uint)cIdx;
                    cIdx++;
                }
                // Create a mirror cellXf pointing at the style xf.
                var styleXfs = stylesheet.CellStyleFormats!;
                var styleXf = (CellFormat)styleXfs.Elements<CellFormat>().ElementAt((int)styleXfId);
                var newXf = new CellFormat
                {
                    NumberFormatId = styleXf.NumberFormatId?.Value ?? 0,
                    FontId = styleXf.FontId?.Value ?? 0,
                    FillId = styleXf.FillId?.Value ?? 0,
                    BorderId = styleXf.BorderId?.Value ?? 0,
                    FormatId = styleXfId,
                    ApplyFont = true,
                };
                cellFormats.Append(newXf);
                cellFormats.Count = (uint)cellFormats.Elements<CellFormat>().Count();
                return (uint)(cellFormats.Elements<CellFormat>().Count() - 1);
            }
        }

        // 2. Create the hyperlink font (blue + underline), dedup by match.
        // Microsoft's canonical hyperlink color is 0563C1 (theme hyperlink).
        var hlFontProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["color"] = "0563C1",
            ["underline"] = "single",
        };
        uint hlFontId = GetOrCreateFont(stylesheet, 0, hlFontProps);

        // 3. Ensure CellStyleFormats exists and append a xf for the Hyperlink style.
        var cellStyleFormats = stylesheet.CellStyleFormats;
        if (cellStyleFormats == null)
        {
            cellStyleFormats = new CellStyleFormats(
                new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
            ) { Count = 1 };
            // Insert before CellFormats if possible.
            var cf = stylesheet.CellFormats;
            if (cf != null)
                cf.InsertBeforeSelf(cellStyleFormats);
            else
                stylesheet.Append(cellStyleFormats);
        }
        var hlStyleXf = new CellFormat
        {
            NumberFormatId = 0,
            FontId = hlFontId,
            FillId = 0,
            BorderId = 0,
            ApplyFont = true,
        };
        cellStyleFormats.Append(hlStyleXf);
        cellStyleFormats.Count = (uint)cellStyleFormats.Elements<CellFormat>().Count();
        uint hlStyleXfId = (uint)(cellStyleFormats.Elements<CellFormat>().Count() - 1);

        // 4. Add a CellFormats (cellXfs) entry that inherits from the style xf.
        var cellFormats2 = EnsureCellFormats(stylesheet);
        var hlCellXf = new CellFormat
        {
            NumberFormatId = 0,
            FontId = hlFontId,
            FillId = 0,
            BorderId = 0,
            FormatId = hlStyleXfId,
            ApplyFont = true,
        };
        cellFormats2.Append(hlCellXf);
        cellFormats2.Count = (uint)cellFormats2.Elements<CellFormat>().Count();
        uint hlCellXfIndex = (uint)(cellFormats2.Elements<CellFormat>().Count() - 1);

        // 5. Register the CellStyle name="Hyperlink" builtinId=8.
        if (cellStyles == null)
        {
            cellStyles = new CellStyles(
                new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }
            ) { Count = 1 };
            stylesheet.Append(cellStyles);
        }
        cellStyles.Append(new CellStyle
        {
            Name = "Hyperlink",
            FormatId = hlStyleXfId,
            BuiltinId = 8,
        });
        cellStyles.Count = (uint)cellStyles.Elements<CellStyle>().Count();

        return hlCellXfIndex;
    }

    /// <summary>
    /// Identify which keys in a dictionary are style properties.
    /// </summary>
    public static bool IsStyleKey(string key)
    {
        var lower = key.ToLowerInvariant();
        return lower is "numfmt" or "fill" or "bgcolor" or "font" or "border"
            or "bold" or "italic" or "strike" or "strikethrough" or "underline"
            or "superscript" or "subscript" or "size"
            or "wrap" or "wraptext" or "numberformat" or "format" or "halign" or "valign"
            or "rotation" or "indent" or "shrinktofit"
            or "locked" or "formulahidden"
            || lower == "readingorder"
            || lower == "quoteprefix"
            || lower.StartsWith("font.")
            || lower.StartsWith("alignment.")
            || lower.StartsWith("border.")
            || lower.StartsWith("protection.");
    }

    // DEFERRED(xlsx/cell-reading-order) CE10: Parse readingOrder values.
    // Accepts numeric (0/1/2) or string (context/contextDependent, ltr/leftToRight,
    // rtl/rightToLeft). Returns OOXML val to stamp as readingOrder="N".
    private static uint ParseReadingOrder(string value)
    {
        var v = value.Trim().ToLowerInvariant();
        return v switch
        {
            "0" or "context" or "contextdependent" => 0u,
            "1" or "ltr" or "lefttoright" => 1u,
            "2" or "rtl" or "righttoleft" => 2u,
            _ => throw new ArgumentException($"Invalid 'readingOrder' value: '{value}'. Expected 0/context, 1/ltr, or 2/rtl.")
        };
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

    // Font property keys handled by the curated builder in GetOrCreateFont
    // (matches the FontMatches dedup keyset). Anything else falls into the
    // long-tail bucket: raw OOXML children appended via SDK schema-aware
    // AddChild on a force-new Font record (skips dedup since the dedup
    // table doesn't track long-tail children).
    private static readonly HashSet<string> CuratedFontKeys =
        new(StringComparer.OrdinalIgnoreCase)
    {
        "bold", "italic", "strike", "underline",
        "superscript", "subscript", "vertalign",
        "size", "name", "color",
    };

    // Lowercased curated sub-key set for alignment.* dispatch — used by the
    // case-preserving long-tail walk to skip keys already handled by the
    // curated switch above.
    private static readonly HashSet<string> CuratedAlignmentSubKeysLower =
        new(StringComparer.Ordinal)
    {
        "horizontal", "vertical", "wraptext", "rotation", "textrotation",
        "indent", "shrinktofit", "shrink", "readingorder",
    };

    // OOXML local-names of Font children produced by the curated GetOrCreateFont
    // builder. baseFont long-tail preservation skips these (they'll be
    // rebuilt from current curated values).
    private static readonly HashSet<string> CuratedFontChildLocalNames =
        new(StringComparer.Ordinal)
    {
        "b", "i", "strike", "u", "vertAlign", "sz", "color", "name",
    };

    private static uint GetOrCreateFont(Stylesheet stylesheet, uint baseFontId,
        Dictionary<string, string> fontProps,
        Dictionary<string, string>? longTailFontProps = null,
        List<string>? unsupportedLongTail = null)
    {
        var fonts = stylesheet.Fonts;
        if (fonts == null)
        {
            fonts = new Fonts(
                new Font(new FontSize { Val = 11 }, new FontName { Val = OfficeDefaultFonts.MinorLatin })
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
            ? (uVal.ToLowerInvariant() is "double" ? "double" : (uVal.ToLowerInvariant() == "single" || (IsValidBooleanString(uVal) && IsTruthy(uVal)) ? "single" : null))
            : (baseFont.Underline != null ? (baseFont.Underline.Val?.InnerText == "double" ? "double" : "single") : null);
        // vertAlign: superscript / subscript / null (baseline)
        var baseVertAlign = baseFont.GetFirstChild<VerticalTextAlignment>();
        string? vertAlign;
        if (fontProps.TryGetValue("superscript", out var supVal))
            vertAlign = IsTruthy(supVal) ? "superscript" : null;
        else if (fontProps.TryGetValue("subscript", out var subVal))
            vertAlign = IsTruthy(subVal) ? "subscript" : null;
        else if (fontProps.TryGetValue("vertalign", out var vaVal))
            vertAlign = vaVal.ToLowerInvariant() is "superscript" or "subscript" ? vaVal.ToLowerInvariant() : null;
        else if (baseVertAlign?.Val?.Value == VerticalAlignmentRunValues.Superscript)
            vertAlign = "superscript";
        else if (baseVertAlign?.Val?.Value == VerticalAlignmentRunValues.Subscript)
            vertAlign = "subscript";
        else
            vertAlign = null;
        double size;
        if (fontProps.TryGetValue("size", out var szVal))
        {
            size = ParseHelpers.ParseFontSize(szVal);
            // R39-4: Excel UI caps font size at 409pt (ECMA-376 §17.4.18).
            // Values above silently render as default 11pt or open broken.
            // The lower bound (>0) is enforced in ParseFontSize; upper
            // bound is Excel-specific so it lives here, not in the shared
            // helper (Word/PPT have far higher limits).
            if (size > 409)
                throw new ArgumentException(
                    $"Invalid font size: '{szVal}'. Excel font size must be <= 409pt.");
        }
        else
        {
            size = baseFont.FontSize?.Val?.Value ?? 11;
        }
        string name = fontProps.GetValueOrDefault("name",
            baseFont.FontName?.Val?.Value ?? OfficeDefaultFonts.MinorLatin);
        // CONSISTENCY(scheme-color): font.color accepts scheme names
        // ("accent1"-"accent6", "lt1"/"dk1", "hlink", etc.) per CLAUDE.md.
        // When matched, store as <color theme="N"/> instead of rgb.
        string? color;
        uint? colorTheme = null;
        if (fontProps.TryGetValue("color", out var cVal))
        {
            var schemeIdx = OfficeCli.Handlers.ExcelHandler.ExcelSchemeColorNameToThemeIndex(cVal);
            if (schemeIdx.HasValue)
            {
                color = null;
                colorTheme = schemeIdx.Value;
            }
            else
            {
                color = NormalizeColor(cVal);
            }
        }
        else
        {
            color = baseFont.Color?.Rgb?.Value;
            colorTheme = baseFont.Color?.Theme?.Value;
        }

        // Long-tail children are added below (post-build) and dedup runs after
        // — that way SDK-rejected keys (e.g. font.bogus=xyz) don't influence
        // the dedup target, and a Font that ends up identical to an existing
        // record (because all long-tail attempts failed) reuses that record
        // instead of bloating the table.
        bool hasLongTail = longTailFontProps != null && longTailFontProps.Count > 0;

        // Create new font (element order: b, i, strike, u, vertAlign, sz, color, name)
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
        if (vertAlign != null)
        {
            newFont.Append(new VerticalTextAlignment
            {
                Val = vertAlign == "superscript"
                    ? VerticalAlignmentRunValues.Superscript
                    : VerticalAlignmentRunValues.Subscript
            });
        }
        newFont.Append(new FontSize { Val = size });
        if (colorTheme.HasValue)
            newFont.Append(new Color { Theme = (UInt32Value)colorTheme.Value });
        else if (color != null)
            newFont.Append(new Color { Rgb = color });
        newFont.Append(new FontName { Val = name });

        // Append long-tail children (charset, family, outline, shadow, condense,
        // extend, scheme, ...) via SDK schema-aware AddChild — orders correctly
        // per CT_Font even though the curated chain above used Append. Track
        // which keys actually landed (vs. SDK-rejected) so dedup runs against
        // the truly-resulting Font, not the input wishlist.
        var addedLongTail = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (hasLongTail)
        {
            foreach (var (key, value) in longTailFontProps!)
            {
                if (OfficeCli.Core.GenericXmlQuery.TryCreateTypedChild(newFont, key, value))
                    addedLongTail[key] = value;
                else
                    unsupportedLongTail?.Add($"font.{key}");
            }
        }

        // Dedup against existing fonts using the actually-built children.
        // Catches three cases the pre-append dedup would miss:
        // (a) repeated SAME long-tail Set on same cell — actualLongTail equals
        //     existing record -> reuse id, no bloat
        // (b) all long-tail rejected (e.g. font.bogus) — actualLongTail is
        //     empty so this matches a curated-only font
        // (c) different cells reaching the same curated+long-tail combo
        int existingIdx = 0;
        foreach (var f in fonts.Elements<Font>())
        {
            if (FontMatches(f, bold, italic, strike, underline, vertAlign, size, name, color, colorTheme)
                && LongTailChildrenMatch(f, addedLongTail))
                return (uint)existingIdx;
            existingIdx++;
        }

        fonts.Append(newFont);
        fonts.Count = (uint)fonts.Elements<Font>().Count();

        return (uint)(fonts.Elements<Font>().Count() - 1);
    }

    // Compare a Font's long-tail children (anything outside CuratedFontChildLocalNames)
    // against a target name->val map. Equal iff the sets match exactly (same keys,
    // same val attribute values). Used to extend FontMatches dedup with long-tail
    // awareness so repeated SAME-value Sets don't bloat the font table.
    private static bool LongTailChildrenMatch(Font font, Dictionary<string, string>? target)
    {
        var targetMap = target ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var fontLongTail = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var child in font.ChildElements)
        {
            var name = child.LocalName;
            if (CuratedFontChildLocalNames.Contains(name)) continue;
            string? valStr = null;
            foreach (var a in child.GetAttributes())
            {
                if (a.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                { valStr = a.Value; break; }
            }
            if (valStr != null) fontLongTail[name] = valStr;
        }
        if (fontLongTail.Count != targetMap.Count) return false;
        foreach (var (k, v) in targetMap)
        {
            if (!fontLongTail.TryGetValue(k, out var fv)) return false;
            if (!string.Equals(fv, v, StringComparison.Ordinal)) return false;
        }
        return true;
    }

    private static bool FontMatches(Font font, bool bold, bool italic, bool strike,
        string? underline, string? vertAlign, double size, string name, string? color, uint? colorTheme = null)
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
        // vertAlign comparison
        var fontVA = font.GetFirstChild<VerticalTextAlignment>();
        string? fontVertAlign = fontVA?.Val?.Value == VerticalAlignmentRunValues.Superscript ? "superscript"
            : fontVA?.Val?.Value == VerticalAlignmentRunValues.Subscript ? "subscript"
            : null;
        if (fontVertAlign != vertAlign) return false;

        if (Math.Abs((font.FontSize?.Val?.Value ?? 11) - size) > 0.01) return false;
        if (!string.Equals(font.FontName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase)) return false;

        var fontColor = font.Color?.Rgb?.Value;
        var fontColorTheme = font.Color?.Theme?.Value;
        if (colorTheme.HasValue)
        {
            if (fontColorTheme != colorTheme.Value) return false;
            if (fontColor != null) return false;
        }
        else if (color != null)
        {
            if (!string.Equals(fontColor, color, StringComparison.OrdinalIgnoreCase)) return false;
            if (fontColorTheme != null) return false;
        }
        else
        {
            if (fontColor != null) return false;
            if (fontColorTheme != null) return false;
        }

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

        // CONSISTENCY(border-dotted-style): R33-1 — accept the dotted form
        // `border.<side>.style=<value>` as alias for `border.<side>=<value>`.
        // Without this, `border.top.style=none` was silently swallowed (the
        // key reached here as `top.style` and matched no branch), reporting
        // success while leaving the border untouched. Same for `all.style`.
        // Per-side `*.color` already has explicit branches below.
        foreach (var sideKey in new[] { "all", "left", "right", "top", "bottom", "diagonal" })
        {
            var dottedStyleKey = sideKey + ".style";
            if (borderProps.TryGetValue(dottedStyleKey, out var dottedStyleVal)
                && !borderProps.ContainsKey(sideKey))
            {
                borderProps[sideKey] = dottedStyleVal;
            }
        }

        // Apply "all" shorthand first (individual sides override later)
        if (borderProps.TryGetValue("all", out var allStyle))
        {
            var parsed = ParseBorderStyle(allStyle);
            leftStyle = rightStyle = topStyle = bottomStyle = parsed;
        }

        // Apply "color" shorthand (border.color) and "all.color" (border.all.color)
        // Both fan out to all four sides. Per-side colors below can still override.
        if (borderProps.TryGetValue("color", out var allColor))
        {
            var normalized = NormalizeColor(allColor);
            leftColor = rightColor = topColor = bottomColor = normalized;
        }
        if (borderProps.TryGetValue("all.color", out var allColor2))
        {
            var normalized = NormalizeColor(allColor2);
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
        uint numFmtId, uint fontId, uint fillId, uint borderId, Alignment? alignment, Protection? protection,
        bool applyNumFmt, bool applyFont, bool applyFill, bool applyBorder, bool applyAlignment, bool applyProtection,
        bool? quotePrefix = null)
    {
        // Search for existing match
        int idx = 0;
        foreach (var xf in cellFormats.Elements<CellFormat>())
        {
            if ((xf.NumberFormatId?.Value ?? 0) == numFmtId &&
                (xf.FontId?.Value ?? 0) == fontId &&
                (xf.FillId?.Value ?? 0) == fillId &&
                (xf.BorderId?.Value ?? 0) == borderId &&
                AlignmentMatches(xf.Alignment, alignment) &&
                ProtectionMatches(xf.Protection, protection) &&
                (xf.QuotePrefix?.Value ?? false) == (quotePrefix ?? false))
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
        if (applyProtection && protection != null)
        {
            newXf.ApplyProtection = true;
            newXf.Append(protection);
        }
        if (quotePrefix == true) newXf.QuotePrefix = true;

        cellFormats.Append(newXf);
        cellFormats.Count = (uint)cellFormats.Elements<CellFormat>().Count();

        return (uint)(cellFormats.Elements<CellFormat>().Count() - 1);
    }

    private static bool ProtectionMatches(Protection? a, Protection? b)
    {
        if (a == null && b == null) return true;
        if (a == null || b == null) return false;
        return (a.Locked?.Value ?? true) == (b.Locked?.Value ?? true) &&
               (a.Hidden?.Value ?? false) == (b.Hidden?.Value ?? false) &&
               UnknownAttrsMatch(a, b, ProtectionCuratedAttrs);
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
               (a.ShrinkToFit?.Value ?? false) == (b.ShrinkToFit?.Value ?? false) &&
               (a.ReadingOrder?.Value ?? 0) == (b.ReadingOrder?.Value ?? 0) &&
               UnknownAttrsMatch(a, b, AlignmentCuratedAttrs);
    }

    // Curated attribute local-names already covered by the typed comparison
    // in AlignmentMatches / ProtectionMatches. The long-tail-aware comparison
    // (UnknownAttrsMatch) walks GetAttributes() and skips these so curated
    // values aren't double-compared via attribute reflection.
    private static readonly HashSet<string> AlignmentCuratedAttrs =
        new(StringComparer.Ordinal)
    {
        "horizontal", "vertical", "wrapText", "textRotation",
        "indent", "shrinkToFit", "readingOrder",
    };
    private static readonly HashSet<string> ProtectionCuratedAttrs =
        new(StringComparer.Ordinal) { "locked", "hidden" };

    // Compare unknown attributes (anything not in the curated set) on two
    // OpenXmlElements. Used by AlignmentMatches/ProtectionMatches so a second
    // Set with a different long-tail attribute value (e.g.
    // alignment.justifyLastLine flipped from "false" to "true") doesn't dedup
    // back to the prior xf and silently drop the new value (BUG-LT4).
    private static bool UnknownAttrsMatch(DocumentFormat.OpenXml.OpenXmlElement a,
        DocumentFormat.OpenXml.OpenXmlElement b, HashSet<string> curated)
    {
        var aAttrs = new Dictionary<string, string>(StringComparer.Ordinal);
        var bAttrs = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var attr in a.GetAttributes())
            if (!curated.Contains(attr.LocalName)) aAttrs[attr.LocalName] = attr.Value ?? "";
        foreach (var attr in b.GetAttributes())
            if (!curated.Contains(attr.LocalName)) bAttrs[attr.LocalName] = attr.Value ?? "";
        if (aAttrs.Count != bAttrs.Count) return false;
        foreach (var (k, v) in aAttrs)
        {
            if (!bAttrs.TryGetValue(k, out var bv)) return false;
            if (!string.Equals(v, bv, StringComparison.Ordinal)) return false;
        }
        return true;
    }

    // ==================== Helpers ====================

    private static Stylesheet CreateDefaultStylesheet()
    {
        return new Stylesheet(
            new NumberingFormats() { Count = 0 },
            new Fonts(
                new Font(new FontSize { Val = 11 }, new FontName { Val = OfficeDefaultFonts.MinorLatin })
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

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    private static bool IsValidBooleanString(string? value) =>
        ParseHelpers.IsValidBooleanString(value);

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
