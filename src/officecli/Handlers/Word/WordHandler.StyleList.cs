// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Style Inheritance ====================

    private RunProperties ResolveEffectiveRunProperties(Run run, Paragraph para)
    {
        var effective = new RunProperties();

        // 1. Start with docDefaults rPr
        var docDefaults = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults;
        var defaultRPr = docDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        if (defaultRPr != null)
            MergeRunProperties(effective, defaultRPr);

        // 2. Walk paragraph style basedOn chain (collect in order, apply from base to derived)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null)
        {
            var chain = new List<Style>();
            var visited = new HashSet<string>();
            var currentStyleId = styleId;
            while (currentStyleId != null && visited.Add(currentStyleId))
            {
                var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
                if (style == null) break;
                chain.Add(style);
                currentStyleId = style.BasedOn?.Val?.Value;
            }
            // Apply from base to derived (reverse order)
            for (int i = chain.Count - 1; i >= 0; i--)
            {
                var styleRPr = chain[i].StyleRunProperties;
                if (styleRPr != null)
                    MergeRunProperties(effective, styleRPr);
            }
        }

        // 3. Resolve character style (rStyle) from the run's rPr
        var rStyleId = run.RunProperties?.GetFirstChild<RunStyle>()?.Val?.Value;
        if (rStyleId != null)
        {
            var rStyleChain = new List<Style>();
            var rVisited = new HashSet<string>();
            var curRStyleId = rStyleId;
            while (curRStyleId != null && rVisited.Add(curRStyleId))
            {
                var rStyle = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                    ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == curRStyleId);
                if (rStyle == null) break;
                rStyleChain.Add(rStyle);
                curRStyleId = rStyle.BasedOn?.Val?.Value;
            }
            for (int i = rStyleChain.Count - 1; i >= 0; i--)
            {
                var sRPr = rStyleChain[i].StyleRunProperties;
                if (sRPr != null)
                    MergeRunProperties(effective, sRPr);
            }
        }

        // 4. Apply run's own direct rPr (highest priority, excluding rStyle which was resolved above)
        if (run.RunProperties != null)
            MergeRunProperties(effective, run.RunProperties);

        return effective;
    }

    private static void MergeRunProperties(RunProperties target, OpenXmlElement source)
    {
        // Merge each known property: source overwrites target
        var srcFonts = source.GetFirstChild<RunFonts>();
        if (srcFonts != null)
            target.RunFonts = srcFonts.CloneNode(true) as RunFonts;

        var srcSize = source.GetFirstChild<FontSize>();
        if (srcSize != null)
            target.FontSize = srcSize.CloneNode(true) as FontSize;

        var srcBold = source.GetFirstChild<Bold>();
        if (srcBold != null)
            target.Bold = srcBold.CloneNode(true) as Bold;

        var srcItalic = source.GetFirstChild<Italic>();
        if (srcItalic != null)
            target.Italic = srcItalic.CloneNode(true) as Italic;

        var srcUnderline = source.GetFirstChild<Underline>();
        if (srcUnderline != null)
            target.Underline = srcUnderline.CloneNode(true) as Underline;

        var srcStrike = source.GetFirstChild<Strike>();
        if (srcStrike != null)
            target.Strike = srcStrike.CloneNode(true) as Strike;

        var srcColor = source.GetFirstChild<Color>();
        if (srcColor != null)
            target.Color = srcColor.CloneNode(true) as Color;

        var srcHighlight = source.GetFirstChild<Highlight>();
        if (srcHighlight != null)
            target.Highlight = srcHighlight.CloneNode(true) as Highlight;
    }

    private static string? GetFontFromProperties(RunProperties? rProps)
    {
        if (rProps == null) return null;
        var fonts = rProps.RunFonts;
        return fonts?.EastAsia?.Value ?? fonts?.Ascii?.Value ?? fonts?.HighAnsi?.Value;
    }

    private static string? GetSizeFromProperties(RunProperties? rProps)
    {
        if (rProps == null) return null;
        var size = rProps.FontSize?.Val?.Value;
        if (size == null) return null;
        return $"{int.Parse(size) / 2}pt";
    }

    // ==================== Effective Properties Resolution ====================

    /// <summary>
    /// Populates effective.* format keys on a paragraph node for properties not explicitly set.
    /// Resolves from: paragraph style chain → document defaults.
    /// </summary>
    private void PopulateEffectiveParagraphProperties(DocumentNode node, Paragraph para)
    {
        // Resolve effective run properties from the first run (or an empty run for style-only resolution)
        var firstRun = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null)
            ?? new Run();
        var effective = ResolveEffectiveRunProperties(firstRun, para);

        // font.size
        if (!node.Format.ContainsKey("size") && effective.FontSize?.Val?.Value != null)
        {
            var sz = int.Parse(effective.FontSize.Val.Value) / 2.0;
            node.Format["effective.size"] = $"{sz:0.##}pt";
        }

        // font.name
        if (!node.Format.ContainsKey("font"))
        {
            var font = effective.RunFonts?.Ascii?.Value ?? effective.RunFonts?.HighAnsi?.Value
                ?? effective.RunFonts?.EastAsia?.Value;
            if (font != null)
                node.Format["effective.font"] = font;
        }

        // bold
        if (!node.Format.ContainsKey("bold") && effective.Bold != null)
            node.Format["effective.bold"] = true;

        // italic
        if (!node.Format.ContainsKey("italic") && effective.Italic != null)
            node.Format["effective.italic"] = true;

        // color
        if (!node.Format.ContainsKey("color"))
        {
            if (effective.Color?.Val?.Value != null)
                node.Format["effective.color"] = ParseHelpers.FormatHexColor(effective.Color.Val.Value);
            else if (effective.Color?.ThemeColor?.HasValue == true)
                node.Format["effective.color"] = effective.Color.ThemeColor.InnerText;
        }

        // underline
        if (!node.Format.ContainsKey("underline") && effective.Underline?.Val != null)
            node.Format["effective.underline"] = effective.Underline.Val.InnerText;

        // Resolve effective paragraph properties from style chain
        ResolveEffectiveParagraphStyleProperties(node, para);
    }

    /// <summary>
    /// Populates effective.* format keys on a run node for properties not explicitly set.
    /// </summary>
    private void PopulateEffectiveRunProperties(DocumentNode node, Run run, Paragraph para)
    {
        var effective = ResolveEffectiveRunProperties(run, para);

        if (!node.Format.ContainsKey("size") && effective.FontSize?.Val?.Value != null)
        {
            var sz = int.Parse(effective.FontSize.Val.Value) / 2.0;
            node.Format["effective.size"] = $"{sz:0.##}pt";
        }

        if (!node.Format.ContainsKey("font"))
        {
            var font = effective.RunFonts?.Ascii?.Value ?? effective.RunFonts?.HighAnsi?.Value
                ?? effective.RunFonts?.EastAsia?.Value;
            if (font != null)
                node.Format["effective.font"] = font;
        }

        if (!node.Format.ContainsKey("bold") && effective.Bold != null)
            node.Format["effective.bold"] = true;

        if (!node.Format.ContainsKey("italic") && effective.Italic != null)
            node.Format["effective.italic"] = true;

        if (!node.Format.ContainsKey("color"))
        {
            if (effective.Color?.Val?.Value != null)
                node.Format["effective.color"] = ParseHelpers.FormatHexColor(effective.Color.Val.Value);
            else if (effective.Color?.ThemeColor?.HasValue == true)
                node.Format["effective.color"] = effective.Color.ThemeColor.InnerText;
        }

        if (!node.Format.ContainsKey("underline") && effective.Underline?.Val != null)
            node.Format["effective.underline"] = effective.Underline.Val.InnerText;

        if (!node.Format.ContainsKey("strike") && effective.Strike != null)
            node.Format["effective.strike"] = true;

        if (!node.Format.ContainsKey("highlight") && effective.Highlight?.Val != null)
            node.Format["effective.highlight"] = effective.Highlight.Val.InnerText;
    }

    /// <summary>
    /// Resolves paragraph-level properties (alignment, spacing) from the paragraph style chain.
    /// </summary>
    private void ResolveEffectiveParagraphStyleProperties(DocumentNode node, Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return;

        var chain = new List<Style>();
        var visited = new HashSet<string>();
        var currentStyleId = styleId;
        while (currentStyleId != null && visited.Add(currentStyleId))
        {
            var style = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == currentStyleId);
            if (style == null) break;
            chain.Add(style);
            currentStyleId = style.BasedOn?.Val?.Value;
        }

        // Apply from base to derived (reverse order), collecting effective paragraph properties
        string? alignment = null;
        string? spaceBefore = null;
        string? spaceAfter = null;
        string? lineSpacing = null;

        for (int i = chain.Count - 1; i >= 0; i--)
        {
            var ppr = chain[i].StyleParagraphProperties;
            if (ppr == null) continue;

            if (ppr.Justification?.Val != null)
            {
                var txt = ppr.Justification.Val.InnerText;
                alignment = txt == "both" ? "justify" : txt;
            }
            if (ppr.SpacingBetweenLines?.Before?.Value != null)
                spaceBefore = SpacingConverter.FormatWordSpacing(ppr.SpacingBetweenLines.Before.Value);
            if (ppr.SpacingBetweenLines?.After?.Value != null)
                spaceAfter = SpacingConverter.FormatWordSpacing(ppr.SpacingBetweenLines.After.Value);
            if (ppr.SpacingBetweenLines?.Line?.Value != null)
                lineSpacing = SpacingConverter.FormatWordLineSpacing(
                    ppr.SpacingBetweenLines.Line.Value,
                    ppr.SpacingBetweenLines.LineRule?.InnerText);
        }

        if (!node.Format.ContainsKey("alignment") && alignment != null)
            node.Format["effective.alignment"] = alignment;
        if (!node.Format.ContainsKey("spaceBefore") && spaceBefore != null)
            node.Format["effective.spaceBefore"] = spaceBefore;
        if (!node.Format.ContainsKey("spaceAfter") && spaceAfter != null)
            node.Format["effective.spaceAfter"] = spaceAfter;
        if (!node.Format.ContainsKey("lineSpacing") && lineSpacing != null)
            node.Format["effective.lineSpacing"] = lineSpacing;
    }

    // ==================== List / Numbering ====================

    private string? GetParagraphListStyle(Paragraph para)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps == null) return null;
        var numId = numProps.NumberingId?.Val?.Value;
        if (numId == null || numId == 0) return null;
        var ilvl = numProps.NumberingLevelReference?.Val?.Value ?? 0;
        var numFmt = GetNumberingFormat(numId.Value, ilvl);
        return numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
    }

    private string GetListPrefix(Paragraph para)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        if (numProps == null) return "";

        var numId = numProps.NumberingId?.Val?.Value;
        var ilvl = numProps.NumberingLevelReference?.Val?.Value ?? 0;
        if (numId == null || numId == 0) return "";

        var indent = new string(' ', ilvl * 2);
        var numFmt = GetNumberingFormat(numId.Value, ilvl);

        return numFmt.ToLowerInvariant() switch
        {
            "bullet" => $"{indent}• ",
            "decimal" => $"{indent}1. ",
            "lowerletter" => $"{indent}a. ",
            "upperletter" => $"{indent}A. ",
            "lowerroman" => $"{indent}i. ",
            "upperroman" => $"{indent}I. ",
            _ => $"{indent}• "
        };
    }

    private string GetNumberingFormat(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return "bullet";

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return "bullet";

        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return "bullet";

        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        if (abstractNum == null) return "bullet";

        var level = abstractNum.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        var numFmt = level?.NumberingFormat?.Val;
        if (numFmt == null || !numFmt.HasValue) return "bullet";
        return numFmt.InnerText ?? "bullet";
    }

    private int? GetStartValue(int numId, int ilvl)
    {
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return null;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return null;

        // Check level override first
        var lvlOverride = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (lvlOverride?.StartOverrideNumberingValue?.Val?.Value is int overrideStart)
            return overrideStart;

        var abstractNumId = numInstance.AbstractNumId?.Val?.Value;
        if (abstractNumId == null) return null;

        var abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(a => a.AbstractNumberId?.Value == abstractNumId);
        var level = abstractNum?.Elements<Level>()
            .FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        return level?.StartNumberingValue?.Val?.Value;
    }

    /// <summary>
    /// Removes numbering from a paragraph.
    /// </summary>
    private static void RemoveListStyle(Paragraph para)
    {
        var pProps = para.ParagraphProperties;
        if (pProps?.NumberingProperties != null)
        {
            pProps.NumberingProperties.Remove();
        }
    }

    /// <summary>
    /// Finds an existing NumberingInstance that uses the same list type (bullet vs ordered),
    /// scanning the last paragraph in the body to support list continuation.
    /// </summary>
    private int? FindContinuationNumId(bool isBullet)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;

        // Check the last paragraph in the body
        var lastPara = body.Elements<Paragraph>().LastOrDefault();
        if (lastPara == null) return null;

        var numProps = lastPara.ParagraphProperties?.NumberingProperties;
        var prevNumId = numProps?.NumberingId?.Val?.Value;
        if (prevNumId == null || prevNumId == 0) return null;

        var fmt = GetNumberingFormat(prevNumId.Value, 0);
        var prevIsBullet = fmt.ToLowerInvariant() == "bullet";
        if (prevIsBullet == isBullet)
            return prevNumId.Value;

        return null;
    }

    private void ApplyListStyle(Paragraph para, string listStyleValue, int? startValue = null, int? listLevel = null)
    {
        // Handle "none" — remove numbering
        if (listStyleValue.ToLowerInvariant() is "none" or "remove" or "clear")
        {
            RemoveListStyle(para);
            return;
        }

        var isBullet = listStyleValue.ToLowerInvariant() is "bullet" or "unordered" or "ul";

        // Try to continue from a preceding list of the same type
        var continuationNumId = FindContinuationNumId(isBullet);
        if (continuationNumId != null && startValue == null)
        {
            var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
            var ilvl = listLevel ?? para.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
            pProps.NumberingProperties = new NumberingProperties
            {
                NumberingId = new NumberingId { Val = continuationNumId.Value },
                NumberingLevelReference = new NumberingLevelReference { Val = ilvl }
            };
            return;
        }

        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart;
        if (numberingPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }
        var numbering = numberingPart.Numbering
            ?? throw new InvalidOperationException("Corrupt file: numbering data missing");

        // Determine the next available IDs
        var maxAbstractId = numbering.Elements<AbstractNum>()
            .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
        var maxNumId = numbering.Elements<NumberingInstance>()
            .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;

        // Create abstract numbering definition with 9 levels
        var abstractNum = new AbstractNum { AbstractNumberId = maxAbstractId };
        abstractNum.AppendChild(new MultiLevelType { Val = MultiLevelValues.HybridMultilevel });

        var bulletChars = new[] { "\u2022", "\u25E6", "\u25AA" }; // •, ◦, ▪

        for (int lvl = 0; lvl < 9; lvl++)
        {
            var level = new Level { LevelIndex = lvl };
            level.AppendChild(new StartNumberingValue { Val = (lvl == 0 && startValue.HasValue) ? startValue.Value : 1 });

            if (isBullet)
            {
                level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Bullet });
                level.AppendChild(new LevelText { Val = bulletChars[lvl % bulletChars.Length] });
            }
            else
            {
                var fmt = (lvl % 3) switch
                {
                    0 => NumberFormatValues.Decimal,
                    1 => NumberFormatValues.LowerLetter,
                    _ => NumberFormatValues.LowerRoman
                };
                level.AppendChild(new NumberingFormat { Val = fmt });
                level.AppendChild(new LevelText { Val = $"%{lvl + 1}." });
            }

            level.AppendChild(new LevelJustification { Val = LevelJustificationValues.Left });
            level.AppendChild(new PreviousParagraphProperties(
                new Indentation { Left = ((lvl + 1) * 720).ToString(), Hanging = "360" }
            ));
            abstractNum.AppendChild(level);
        }

        // Insert AbstractNum before any NumberingInstance elements
        var firstNumInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstNumInstance != null)
            numbering.InsertBefore(abstractNum, firstNumInstance);
        else
            numbering.AppendChild(abstractNum);

        // Create numbering instance
        var numInstance = new NumberingInstance { NumberID = maxNumId };
        numInstance.AppendChild(new AbstractNumId { Val = maxAbstractId });
        numbering.AppendChild(numInstance);

        numbering.Save();

        // Apply to paragraph
        var pProps2 = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
        pProps2.NumberingProperties = new NumberingProperties
        {
            NumberingId = new NumberingId { Val = maxNumId },
            NumberingLevelReference = new NumberingLevelReference { Val = listLevel ?? 0 }
        };
    }

    /// <summary>
    /// Sets the start value override for a paragraph's numbering instance.
    /// </summary>
    private void SetListStartValue(Paragraph para, int startValue)
    {
        var numProps = para.ParagraphProperties?.NumberingProperties;
        var numId = numProps?.NumberingId?.Val?.Value;
        if (numId == null || numId == 0) return;

        var ilvl = numProps?.NumberingLevelReference?.Val?.Value ?? 0;
        var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
        if (numbering == null) return;

        var numInstance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(n => n.NumberID?.Value == numId);
        if (numInstance == null) return;

        // Find or create LevelOverride for this ilvl
        var lvlOverride = numInstance.Elements<LevelOverride>()
            .FirstOrDefault(o => o.LevelIndex?.Value == ilvl);
        if (lvlOverride == null)
        {
            lvlOverride = new LevelOverride { LevelIndex = ilvl };
            numInstance.AppendChild(lvlOverride);
        }
        lvlOverride.StartOverrideNumberingValue = new StartOverrideNumberingValue { Val = startValue };

        numbering.Save();
    }
}
