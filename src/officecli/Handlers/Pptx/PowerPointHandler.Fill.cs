// Copyright 2025 OfficeCli (officecli.ai)
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

    /// <summary>
    /// Parse a color string and return the appropriate OpenXML color element.
    /// Supports: hex RGB ("FF0000"), theme colors ("accent1", "dk1", "lt1", etc.)
    /// </summary>
    private static OpenXmlElement BuildColorElement(string value)
    {
        var schemeColor = TryParseSchemeColor(value);
        if (schemeColor.HasValue)
            return new Drawing.SchemeColor { Val = schemeColor.Value };

        var (rgb, alpha) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(value);
        var colorEl = new Drawing.RgbColorModelHex { Val = rgb };
        if (alpha.HasValue)
            colorEl.AppendChild(new Drawing.Alpha { Val = alpha.Value });
        return colorEl;
    }

    /// <summary>
    /// Build a SolidFill element with the appropriate color type.
    /// </summary>
    private static Drawing.SolidFill BuildSolidFill(string colorValue)
    {
        var solidFill = new Drawing.SolidFill();
        solidFill.Append(BuildColorElement(colorValue));
        return solidFill;
    }

    /// <summary>
    /// Try to parse a theme/scheme color name. Returns null if it's a hex RGB value.
    /// </summary>
    private static Drawing.SchemeColorValues? TryParseSchemeColor(string value)
    {
        return value.ToLowerInvariant().TrimStart('#') switch
        {
            "accent1" => Drawing.SchemeColorValues.Accent1,
            "accent2" => Drawing.SchemeColorValues.Accent2,
            "accent3" => Drawing.SchemeColorValues.Accent3,
            "accent4" => Drawing.SchemeColorValues.Accent4,
            "accent5" => Drawing.SchemeColorValues.Accent5,
            "accent6" => Drawing.SchemeColorValues.Accent6,
            "dk1" or "dark1" => Drawing.SchemeColorValues.Dark1,
            "dk2" or "dark2" => Drawing.SchemeColorValues.Dark2,
            "lt1" or "light1" => Drawing.SchemeColorValues.Light1,
            "lt2" or "light2" => Drawing.SchemeColorValues.Light2,
            "tx1" or "text1" => Drawing.SchemeColorValues.Text1,
            "tx2" or "text2" => Drawing.SchemeColorValues.Text2,
            "bg1" or "background1" => Drawing.SchemeColorValues.Background1,
            "bg2" or "background2" => Drawing.SchemeColorValues.Background2,
            "hlink" or "hyperlink" => Drawing.SchemeColorValues.Hyperlink,
            "folhlink" or "followedhyperlink" => Drawing.SchemeColorValues.FollowedHyperlink,
            _ => null
        };
    }

    /// <summary>
    /// Read a color value from a SolidFill element, returning either hex RGB or scheme color name.
    /// </summary>
    internal static string? ReadColorFromFill(Drawing.SolidFill? solidFill)
    {
        if (solidFill == null) return null;
        var rgb = solidFill.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null) return ParseHelpers.FormatHexColor(rgb);
        var scheme = solidFill.GetFirstChild<Drawing.SchemeColor>()?.Val;
        if (scheme?.HasValue == true) return scheme.InnerText;
        return null;
    }

    /// <summary>
    /// Read a color value from any element that may contain RgbColorModelHex or SchemeColor.
    /// </summary>
    internal static string? ReadColorFromElement(OpenXmlElement? parent)
    {
        if (parent == null) return null;
        var rgb = parent.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
        if (rgb != null) return ParseHelpers.FormatHexColor(rgb);
        var scheme = parent.GetFirstChild<Drawing.SchemeColor>()?.Val;
        if (scheme?.HasValue == true) return scheme.InnerText;
        return null;
    }

    private static void ApplyShapeFill(ShapeProperties spPr, string value)
    {
        // Build new fill element BEFORE removing old one (atomic: no data loss on validation failure)
        OpenXmlElement newFill = value.Equals("none", StringComparison.OrdinalIgnoreCase)
            ? new Drawing.NoFill()
            : BuildSolidFill(value);

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
                break;
            default:
                if (value.Length <= 2)
                    pProps.AppendChild(new Drawing.CharacterBullet { Char = value });
                else
                    throw new ArgumentException($"Invalid list style: {value}. Use: bullet, numbered, alpha, roman, none, or a single character");
                break;
        }
    }

    private static Drawing.ShapeTypeValues ParsePresetShape(string name) =>
        name.ToLowerInvariant() switch
        {
            "rect" or "rectangle" => Drawing.ShapeTypeValues.Rectangle,
            "roundrect" or "roundedrectangle" => Drawing.ShapeTypeValues.RoundRectangle,
            "ellipse" or "oval" => Drawing.ShapeTypeValues.Ellipse,
            "triangle" => Drawing.ShapeTypeValues.Triangle,
            "rtriangle" or "righttriangle" => Drawing.ShapeTypeValues.RightTriangle,
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
            "rightarrow" or "rarrow" => Drawing.ShapeTypeValues.RightArrow,
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
            "uturnArrow" => Drawing.ShapeTypeValues.UTurnArrow,
            "circularArrow" => Drawing.ShapeTypeValues.CircularArrow,
            _ => throw new ArgumentException(
                $"Unknown preset shape: '{name}'. Common presets: rect, roundRect, ellipse, triangle, diamond, " +
                "pentagon, hexagon, star5, rightArrow, leftArrow, chevron, plus, heart, cloud, cube, can, line, " +
                "callout, process, decision, smiley, frame, gear6")
        };

    private static Drawing.LineEndValues ParseLineEndType(string name) =>
        name.ToLowerInvariant() switch
        {
            "triangle" or "arrow" => Drawing.LineEndValues.Triangle,
            "stealth" => Drawing.LineEndValues.Stealth,
            "diamond" => Drawing.LineEndValues.Diamond,
            "oval" or "circle" => Drawing.LineEndValues.Oval,
            "open" => Drawing.LineEndValues.Arrow,
            "none" => Drawing.LineEndValues.None,
            _ => Drawing.LineEndValues.Triangle
        };
}
