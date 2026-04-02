// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Populate Format dictionary on the root DocumentNode with document-level settings.
    /// Called from GetRootNode().
    /// </summary>
    private void PopulateDocSettings(DocumentNode node)
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var sectPr = _doc.MainDocumentPart?.Document?.Body?.GetFirstChild<SectionProperties>()
            ?? _doc.MainDocumentPart?.Document?.Body?.Descendants<SectionProperties>().LastOrDefault();

        // ==================== DocGrid ====================
        if (sectPr != null)
        {
            var grid = sectPr.GetFirstChild<DocGrid>();
            if (grid != null)
            {
                if (grid.Type?.Value != null)
                    node.Format["docGrid.type"] = grid.Type.InnerText;
                if (grid.LinePitch?.Value != null)
                    node.Format["docGrid.linePitch"] = grid.LinePitch.Value;
                if (grid.CharacterSpace?.Value != null)
                    node.Format["docGrid.charSpace"] = grid.CharacterSpace.Value;
            }

            // ==================== Columns ====================
            var cols = sectPr.GetFirstChild<Columns>();
            if (cols != null)
            {
                if (cols.ColumnCount?.Value != null)
                    node.Format["columns.count"] = (int)cols.ColumnCount.Value;
                if (cols.Space?.Value != null && double.TryParse(cols.Space.Value,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var spaceVal))
                    node.Format["columns.space"] = FormatTwipsToCm((uint)Math.Round(spaceVal));
                if (cols.EqualWidth?.Value != null)
                    node.Format["columns.equalWidth"] = cols.EqualWidth.Value;
                if (cols.Separator?.Value != null)
                    node.Format["columns.separator"] = cols.Separator.Value;
            }

            // ==================== SectionType ====================
            var sectType = sectPr.GetFirstChild<SectionType>();
            if (sectType?.Val?.Value != null)
                node.Format["section.type"] = sectType.Val.InnerText;
        }

        // ==================== CJK Layout (from DocDefaults ParagraphProperties) ====================
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        var pPrBase = stylesPart?.Styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        if (pPrBase != null)
        {
            if (pPrBase.AutoSpaceDE != null)
                node.Format["autoSpaceDE"] = pPrBase.AutoSpaceDE.Val?.Value ?? true;
            if (pPrBase.AutoSpaceDN != null)
                node.Format["autoSpaceDN"] = pPrBase.AutoSpaceDN.Val?.Value ?? true;
            if (pPrBase.Kinsoku != null)
                node.Format["kinsoku"] = pPrBase.Kinsoku.Val?.Value ?? true;
            if (pPrBase.OverflowPunctuation != null)
                node.Format["overflowPunct"] = pPrBase.OverflowPunctuation.Val?.Value ?? true;
        }

        if (settings == null) return;

        // ==================== CharacterSpacingControl ====================
        var charSpacing = settings.GetFirstChild<CharacterSpacingControl>();
        if (charSpacing?.Val?.Value != null)
            node.Format["charSpacingControl"] = charSpacing.Val.InnerText;

        // ==================== Print / Display ====================
        if (settings.GetFirstChild<DisplayBackgroundShape>() != null)
            node.Format["displayBackgroundShape"] = true;
        if (settings.GetFirstChild<DoNotDisplayPageBoundaries>() != null)
            node.Format["doNotDisplayPageBoundaries"] = true;
        if (settings.GetFirstChild<PrintFormsData>() != null)
            node.Format["printFormsData"] = true;
        if (settings.GetFirstChild<PrintPostScriptOverText>() != null)
            node.Format["printPostScriptOverText"] = true;
        if (settings.GetFirstChild<PrintFractionalCharacterWidth>() != null)
            node.Format["printFractionalCharacterWidth"] = true;
        if (settings.GetFirstChild<RemovePersonalInformation>() != null)
            node.Format["removePersonalInformation"] = true;
        if (settings.GetFirstChild<RemoveDateAndTime>() != null)
            node.Format["removeDateAndTime"] = true;

        // ==================== Font Embedding ====================
        if (settings.GetFirstChild<EmbedTrueTypeFonts>() != null)
            node.Format["embedFonts"] = true;
        if (settings.GetFirstChild<EmbedSystemFonts>() != null)
            node.Format["embedSystemFonts"] = true;
        if (settings.GetFirstChild<SaveSubsetFonts>() != null)
            node.Format["saveSubsetFonts"] = true;

        // ==================== Layout Flags ====================
        if (settings.GetFirstChild<MirrorMargins>() != null)
            node.Format["mirrorMargins"] = true;
        if (settings.GetFirstChild<GutterAtTop>() != null)
            node.Format["gutterAtTop"] = true;
        if (settings.GetFirstChild<BookFoldPrinting>() != null)
            node.Format["bookFoldPrinting"] = true;
        if (settings.GetFirstChild<BookFoldReversePrinting>() != null)
            node.Format["bookFoldReversePrinting"] = true;
        var bookFoldSheets = settings.GetFirstChild<BookFoldPrintingSheets>();
        if (bookFoldSheets?.Val?.Value != null)
            node.Format["bookFoldPrintingSheets"] = (int)bookFoldSheets.Val.Value;
        if (settings.GetFirstChild<EvenAndOddHeaders>() != null)
            node.Format["evenAndOddHeaders"] = true;
        var defTabStop = settings.GetFirstChild<DefaultTabStop>();
        if (defTabStop?.Val?.Value != null)
            node.Format["defaultTabStop"] = FormatTwipsToCm((uint)defTabStop.Val.Value);
    }
}
