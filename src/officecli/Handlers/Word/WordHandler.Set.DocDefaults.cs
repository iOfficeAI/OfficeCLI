// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Try to handle docDefaults.* keys. Returns true if handled.
    /// </summary>
    private bool TrySetDocDefaults(string key, string value)
    {
        switch (key)
        {
            // ==================== Default Run Properties ====================
            case "docdefaults.font" or "docdefaults.fontname":
            {
                var rPr = EnsureRunPropsDefault();
                var fonts = rPr.GetFirstChild<RunFonts>() ?? rPr.AppendChild(new RunFonts());
                fonts.Ascii = value;
                fonts.HighAnsi = value;
                fonts.EastAsia = value;
                fonts.ComplexScript = value;
                SaveStyles();
                return true;
            }
            case "docdefaults.font.latin" or "docdefaults.font.ascii":
            {
                var rPr = EnsureRunPropsDefault();
                var fonts = rPr.GetFirstChild<RunFonts>() ?? rPr.AppendChild(new RunFonts());
                fonts.Ascii = value;
                fonts.HighAnsi = value;
                SaveStyles();
                return true;
            }
            case "docdefaults.font.hansi" or "docdefaults.font.highansi":
            {
                var rPr = EnsureRunPropsDefault();
                var fonts = rPr.GetFirstChild<RunFonts>() ?? rPr.AppendChild(new RunFonts());
                fonts.HighAnsi = value;
                SaveStyles();
                return true;
            }
            case "docdefaults.font.eastasia":
            {
                var rPr = EnsureRunPropsDefault();
                var fonts = rPr.GetFirstChild<RunFonts>() ?? rPr.AppendChild(new RunFonts());
                fonts.EastAsia = value;
                SaveStyles();
                return true;
            }
            case "docdefaults.font.complexscript" or "docdefaults.font.cs":
            {
                var rPr = EnsureRunPropsDefault();
                var fonts = rPr.GetFirstChild<RunFonts>() ?? rPr.AppendChild(new RunFonts());
                fonts.ComplexScript = value;
                SaveStyles();
                return true;
            }
            case "docdefaults.fontsize" or "docdefaults.size":
            {
                var rPr = EnsureRunPropsDefault();
                var halfPts = ParseFontSizeToHalfPoints(value);
                var sz = rPr.GetFirstChild<FontSize>() ?? rPr.AppendChild(new FontSize());
                sz.Val = halfPts;
                var szCs = rPr.GetFirstChild<FontSizeComplexScript>() ?? rPr.AppendChild(new FontSizeComplexScript());
                szCs.Val = halfPts;
                SaveStyles();
                return true;
            }
            case "docdefaults.color":
            {
                var rPr = EnsureRunPropsDefault();
                var color = rPr.GetFirstChild<Color>();
                if (color == null)
                {
                    color = new Color();
                    // Schema order: color must come before sz, szCs
                    InsertRunPropBeforeSizeElements(rPr, color);
                }
                color.Val = SanitizeHex(value);
                SaveStyles();
                return true;
            }
            case "docdefaults.bold":
            {
                var rPr = EnsureRunPropsDefault();
                SetRunPropBoolInOrder<Bold>(rPr, IsTruthy(value));
                SaveStyles();
                return true;
            }
            case "docdefaults.italic":
            {
                var rPr = EnsureRunPropsDefault();
                SetRunPropBoolInOrder<Italic>(rPr, IsTruthy(value));
                SaveStyles();
                return true;
            }
            case "docdefaults.rtl" or "docdefaults.direction" or "docdefaults.dir":
            {
                // <w:rtl/> on rPrDefault makes RTL the document-wide default;
                // explicit run rtl=false overrides per-run. Mirrors bold/italic.
                var rPr = EnsureRunPropsDefault();
                bool rtlOn = key.ToLowerInvariant() == "docdefaults.rtl"
                    ? IsTruthy(value)
                    : ParseDirectionRtl(value);
                // <w:rtl/> sits late in CT_RPr (after vertAlign), so AppendChild
                // is schema-correct here — unlike Bold/Italic which must precede
                // Color/FontSize.
                rPr.RemoveAllChildren<RightToLeftText>();
                if (rtlOn) rPr.AppendChild(new RightToLeftText());
                SaveStyles();
                return true;
            }

            // ==================== Default Paragraph Properties ====================
            case "docdefaults.alignment" or "docdefaults.align":
            {
                var pPr = EnsureParaPropsDefault();
                // Use typed property setter to preserve OOXML schema element order
                // (Justification must precede AutoSpaceDE; AppendChild would place it last)
                if (pPr.Justification == null)
                    pPr.Justification = new Justification();
                pPr.Justification.Val = ParseJustification(value);
                SaveStyles();
                return true;
            }
            case "docdefaults.spacebefore":
            {
                var pPr = EnsureParaPropsDefault();
                // Use typed property setter to preserve OOXML schema element order
                if (pPr.SpacingBetweenLines == null)
                    pPr.SpacingBetweenLines = new SpacingBetweenLines();
                pPr.SpacingBetweenLines.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                SaveStyles();
                return true;
            }
            case "docdefaults.spaceafter":
            {
                var pPr = EnsureParaPropsDefault();
                if (pPr.SpacingBetweenLines == null)
                    pPr.SpacingBetweenLines = new SpacingBetweenLines();
                pPr.SpacingBetweenLines.After = SpacingConverter.ParseWordSpacing(value).ToString();
                SaveStyles();
                return true;
            }
            case "docdefaults.linespacing":
            {
                var pPr = EnsureParaPropsDefault();
                if (pPr.SpacingBetweenLines == null)
                    pPr.SpacingBetweenLines = new SpacingBetweenLines();
                var (twips, isMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                pPr.SpacingBetweenLines.Line = twips.ToString();
                pPr.SpacingBetweenLines.LineRule = isMultiplier
                    ? new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Auto)
                    : new DocumentFormat.OpenXml.EnumValue<LineSpacingRuleValues>(LineSpacingRuleValues.Exact);
                SaveStyles();
                return true;
            }

            default:
                return false;
        }
    }

    // ==================== Helpers ====================

    private RunPropertiesBaseStyle EnsureRunPropsDefault()
    {
        var mainPart = _doc.MainDocumentPart!;
        var stylesPart = mainPart.StyleDefinitionsPart
            ?? mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        var docDefaults = stylesPart.Styles.DocDefaults;
        if (docDefaults == null)
        {
            docDefaults = new DocDefaults();
            stylesPart.Styles.AppendChild(docDefaults);
        }

        var rPrDefault = docDefaults.RunPropertiesDefault;
        if (rPrDefault == null)
        {
            rPrDefault = new RunPropertiesDefault();
            // Schema order: rPrDefault must precede pPrDefault
            var pPrDefault = docDefaults.ParagraphPropertiesDefault;
            if (pPrDefault != null)
                pPrDefault.InsertBeforeSelf(rPrDefault);
            else
                docDefaults.AppendChild(rPrDefault);
        }

        var rPrBase = rPrDefault.RunPropertiesBaseStyle;
        if (rPrBase == null)
        {
            rPrBase = new RunPropertiesBaseStyle();
            rPrDefault.RunPropertiesBaseStyle = rPrBase;
        }

        return rPrBase;
    }

    /// <summary>
    /// Parse font size input (e.g. "14", "14pt", "10.5pt") to half-points string for OOXML.
    /// </summary>
    private static string ParseFontSizeToHalfPoints(string value)
    {
        // Route through ParseFontSize so the shared min/max guards
        // (>= 0.5pt, <= 4000pt) apply uniformly across handlers — previously
        // size=2147483647 overflowed `pts * 2` to a negative w:sz value.
        var pts = ParseHelpers.ParseFontSize(value);
        return ((int)Math.Round(pts * 2)).ToString();
    }

    private static void SetRunPropBool<T>(RunPropertiesBaseStyle rPr, bool value) where T : OnOffType, new()
    {
        var existing = rPr.GetFirstChild<T>();
        existing?.Remove();
        if (value)
            rPr.AppendChild(new T());
    }

    /// <summary>
    /// Set a Bold or Italic element in schema-correct order: before Color, FontSize, FontSizeComplexScript.
    /// </summary>
    private static void SetRunPropBoolInOrder<T>(RunPropertiesBaseStyle rPr, bool value) where T : OnOffType, new()
    {
        var existing = rPr.GetFirstChild<T>();
        existing?.Remove();
        if (value)
        {
            // b/i must appear before color, sz, szCs in w:rPr schema order
            InsertRunPropBeforeSizeElements(rPr, new T());
        }
    }

    /// <summary>
    /// Insert an element before the first of Color, FontSize, FontSizeComplexScript if any exist,
    /// otherwise append. This preserves schema order for w:rPrBase children.
    /// </summary>
    private static void InsertRunPropBeforeSizeElements(RunPropertiesBaseStyle rPr, DocumentFormat.OpenXml.OpenXmlElement elem)
    {
        // Schema order in w:rPr: rFonts → b → i → ... → color → sz → szCs → ...
        // Bold/Italic must come before Color; Color must come before FontSize/FontSizeComplexScript.
        // Find the earliest "later" element to insert before.
        DocumentFormat.OpenXml.OpenXmlElement? anchor = null;
        foreach (var child in rPr.ChildElements)
        {
            if (child is Color || child is FontSize || child is FontSizeComplexScript)
            {
                anchor = child;
                break;
            }
            // Bold/Italic also come before Color but after RunFonts — only apply anchor for
            // elements that must come after the one we're inserting.
            // For Color: only anchor on FontSize/FontSizeComplexScript (not Bold/Italic since those come before Color)
            // For Bold/Italic: anchor on Color, FontSize, FontSizeComplexScript
        }
        if (anchor != null)
            anchor.InsertBeforeSelf(elem);
        else
            rPr.AppendChild(elem);
    }

    private void SaveStyles()
    {
        _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Save();
    }

    /// <summary>
    /// Read DocDefaults into Format dictionary.
    /// </summary>
    private void PopulateDocDefaults(DocumentNode node)
    {
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        var docDefaults = stylesPart?.Styles?.DocDefaults;
        if (docDefaults == null) return;

        // Run properties defaults
        var rPr = docDefaults.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (rPr != null)
        {
            var fonts = rPr.GetFirstChild<RunFonts>();
            if (fonts?.Ascii?.Value != null)
            {
                node.Format["docDefaults.font"] = fonts.Ascii.Value;
                node.Format["defaultFont"] = fonts.Ascii.Value; // legacy alias for backward compat
            }
            if (fonts?.EastAsia?.Value != null)
                node.Format["docDefaults.font.eastAsia"] = fonts.EastAsia.Value;
            if (fonts?.HighAnsi?.Value != null && fonts.HighAnsi.Value != fonts.Ascii?.Value)
                node.Format["docDefaults.font.hAnsi"] = fonts.HighAnsi.Value;
            if (fonts?.ComplexScript?.Value != null)
                node.Format["docDefaults.font.complexScript"] = fonts.ComplexScript.Value;

            var sz = rPr.GetFirstChild<FontSize>();
            if (sz?.Val?.Value != null)
            {
                var halfPts = ParseHelpers.SafeParseDouble(sz.Val.Value, "fontSize");
                node.Format["docDefaults.fontSize"] = $"{halfPts / 2}pt";
            }

            var color = rPr.GetFirstChild<Color>();
            if (color?.Val?.Value != null)
                node.Format["docDefaults.color"] = ParseHelpers.FormatHexColor(color.Val.Value);

            if (rPr.GetFirstChild<Bold>() != null)
                node.Format["docDefaults.bold"] = true;
            if (rPr.GetFirstChild<Italic>() != null)
                node.Format["docDefaults.italic"] = true;
            if (rPr.GetFirstChild<RightToLeftText>() != null)
                node.Format["docDefaults.rtl"] = true;
        }

        // Paragraph properties defaults
        var pPr = docDefaults.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (pPr != null)
        {
            var jc = pPr.GetFirstChild<Justification>();
            if (jc?.Val?.Value != null)
                node.Format["docDefaults.alignment"] = jc.Val.InnerText;

            var spacing = pPr.GetFirstChild<SpacingBetweenLines>();
            if (spacing != null)
            {
                if (spacing.Before?.Value != null)
                    node.Format["docDefaults.spaceBefore"] = FormatTwipsToPt(uint.Parse(spacing.Before.Value));
                if (spacing.After?.Value != null)
                    node.Format["docDefaults.spaceAfter"] = FormatTwipsToPt(uint.Parse(spacing.After.Value));
                if (spacing.Line?.Value != null)
                {
                    var lineRule = spacing.LineRule?.InnerText ?? "auto";
                    var lineVal = int.Parse(spacing.Line.Value);
                    node.Format["docDefaults.lineSpacing"] = lineRule == "auto"
                        ? $"{lineVal / 240.0:0.##}x"
                        : $"{lineVal / 20.0:0.##}pt";
                }
            }
        }
    }

    private static string FormatTwipsToPt(uint twips)
    {
        var pt = twips / 20.0;
        return $"{pt:0.##}pt";
    }
}
