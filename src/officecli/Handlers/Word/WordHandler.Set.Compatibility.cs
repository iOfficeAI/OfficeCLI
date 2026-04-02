// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// All Compatibility child element types (OnOffType), keyed by lowercase name.
    /// Used for generic Set/Get of compat flags.
    /// </summary>
    private static readonly Dictionary<string, Func<OnOffType>> CompatElementFactory = new(StringComparer.OrdinalIgnoreCase)
    {
        ["useSingleBorderForContiguousCells"] = () => new UseSingleBorderForContiguousCells(),
        ["wpJustification"] = () => new WordPerfectJustification(),
        ["noTabHangIndent"] = () => new NoTabHangIndent(),
        ["noLeading"] = () => new NoLeading(),
        ["spaceForUnderline"] = () => new SpaceForUnderline(),
        ["noColumnBalance"] = () => new NoColumnBalance(),
        ["balanceSingleByteDoubleByteWidth"] = () => new BalanceSingleByteDoubleByteWidth(),
        ["noExtraLineSpacing"] = () => new NoExtraLineSpacing(),
        ["doNotLeaveBackslashAlone"] = () => new DoNotLeaveBackslashAlone(),
        ["underlineTrailingSpaces"] = () => new UnderlineTrailingSpaces(),
        ["doNotExpandShiftReturn"] = () => new DoNotExpandShiftReturn(),
        ["spacingInWholePoints"] = () => new SpacingInWholePoints(),
        ["lineWrapLikeWord6"] = () => new LineWrapLikeWord6(),
        ["printBodyTextBeforeHeader"] = () => new PrintBodyTextBeforeHeader(),
        ["printColorBlackWhite"] = () => new PrintColorBlackWhite(),
        ["wordPerfectSpaceWidth"] = () => new WordPerfectSpaceWidth(),
        ["showBreaksInFrames"] = () => new ShowBreaksInFrames(),
        ["subFontBySize"] = () => new SubFontBySize(),
        ["suppressBottomSpacing"] = () => new SuppressBottomSpacing(),
        ["suppressTopSpacing"] = () => new SuppressTopSpacing(),
        ["suppressSpacingAtTopOfPage"] = () => new SuppressSpacingAtTopOfPage(),
        ["suppressTopSpacingWordPerfect"] = () => new SuppressTopSpacingWordPerfect(),
        ["suppressSpacingBeforeAfterPageBreak"] = () => new SuppressSpacingBeforeAfterPageBreak(),
        ["swapBordersFacingPages"] = () => new SwapBordersFacingPages(),
        ["convertMailMergeEscape"] = () => new ConvertMailMergeEscape(),
        ["truncateFontHeightsLikeWordPerfect"] = () => new TruncateFontHeightsLikeWordPerfect(),
        ["macWordSmallCaps"] = () => new MacWordSmallCaps(),
        ["usePrinterMetrics"] = () => new UsePrinterMetrics(),
        ["doNotSuppressParagraphBorders"] = () => new DoNotSuppressParagraphBorders(),
        ["wrapTrailSpaces"] = () => new WrapTrailSpaces(),
        ["footnoteLayoutLikeWord8"] = () => new FootnoteLayoutLikeWord8(),
        ["shapeLayoutLikeWord8"] = () => new ShapeLayoutLikeWord8(),
        ["alignTablesRowByRow"] = () => new AlignTablesRowByRow(),
        ["forgetLastTabAlignment"] = () => new ForgetLastTabAlignment(),
        ["adjustLineHeightInTable"] = () => new AdjustLineHeightInTable(),
        ["autoSpaceLikeWord95"] = () => new AutoSpaceLikeWord95(),
        ["noSpaceRaiseLower"] = () => new NoSpaceRaiseLower(),
        ["doNotUseHTMLParagraphAutoSpacing"] = () => new DoNotUseHTMLParagraphAutoSpacing(),
        ["layoutRawTableWidth"] = () => new LayoutRawTableWidth(),
        ["layoutTableRowsApart"] = () => new LayoutTableRowsApart(),
        ["useWord97LineBreakRules"] = () => new UseWord97LineBreakRules(),
        ["doNotBreakWrappedTables"] = () => new DoNotBreakWrappedTables(),
        ["doNotSnapToGridInCell"] = () => new DoNotSnapToGridInCell(),
        ["selectFieldWithFirstOrLastChar"] = () => new SelectFieldWithFirstOrLastChar(),
        ["applyBreakingRules"] = () => new ApplyBreakingRules(),
        ["doNotWrapTextWithPunctuation"] = () => new DoNotWrapTextWithPunctuation(),
        ["doNotUseEastAsianBreakRules"] = () => new DoNotUseEastAsianBreakRules(),
        ["useWord2002TableStyleRules"] = () => new UseWord2002TableStyleRules(),
        ["growAutofit"] = () => new GrowAutofit(),
        ["useFarEastLayout"] = () => new UseFarEastLayout(),
        ["useNormalStyleForList"] = () => new UseNormalStyleForList(),
        ["doNotUseIndentAsNumberingTabStop"] = () => new DoNotUseIndentAsNumberingTabStop(),
        ["useAltKinsokuLineBreakRules"] = () => new UseAltKinsokuLineBreakRules(),
        ["allowSpaceOfSameStyleInTable"] = () => new AllowSpaceOfSameStyleInTable(),
        ["doNotSuppressIndentation"] = () => new DoNotSuppressIndentation(),
        ["doNotAutofitConstrainedTables"] = () => new DoNotAutofitConstrainedTables(),
        ["autofitToFirstFixedWidthCell"] = () => new AutofitToFirstFixedWidthCell(),
        ["underlineTabInNumberingList"] = () => new UnderlineTabInNumberingList(),
        ["displayHangulFixedWidth"] = () => new DisplayHangulFixedWidth(),
        ["splitPageBreakAndParagraphMark"] = () => new SplitPageBreakAndParagraphMark(),
        ["doNotVerticallyAlignCellWithShape"] = () => new DoNotVerticallyAlignCellWithShape(),
        ["doNotBreakConstrainedForcedTable"] = () => new DoNotBreakConstrainedForcedTable(),
        ["doNotVerticallyAlignInTextBox"] = () => new DoNotVerticallyAlignInTextBox(),
        ["useAnsiKerningPairs"] = () => new UseAnsiKerningPairs(),
        ["cachedColumnBalance"] = () => new CachedColumnBalance(),
    };

    /// <summary>
    /// Preset definitions for compatibility.preset.
    /// Each preset is a set of compat flags + a compatibilityMode value.
    /// </summary>
    private static readonly Dictionary<string, (int CompatMode, string[] EnableFlags, string[] DisableFlags)> CompatPresets = new(StringComparer.OrdinalIgnoreCase)
    {
        ["word2019"] = (15, ["useFarEastLayout"], []),
        ["word2010"] = (14, ["doNotUseHTMLParagraphAutoSpacing", "useWord2002TableStyleRules", "useFarEastLayout"], []),
        ["css-layout"] = (15, ["adjustLineHeightInTable", "useFarEastLayout"], []),
    };

    /// <summary>
    /// Try to handle compatibility.* keys. Returns true if handled.
    /// </summary>
    private bool TrySetCompatibility(string key, string value)
    {
        // compatibility.preset — apply a batch of settings
        if (key == "compatibility.preset")
        {
            if (!CompatPresets.TryGetValue(value, out var preset))
                throw new ArgumentException($"Unknown compatibility preset: '{value}'. Valid: {string.Join(", ", CompatPresets.Keys)}");

            var compat = EnsureCompatibility();

            // Set compatibilityMode via CompatibilitySetting
            SetCompatibilityMode(compat, preset.CompatMode);

            // Enable flags
            foreach (var flag in preset.EnableFlags)
            {
                if (CompatElementFactory.TryGetValue(flag, out var factory))
                    SetCompatFlag(compat, factory, true);
            }
            // Disable flags
            foreach (var flag in preset.DisableFlags)
            {
                if (CompatElementFactory.TryGetValue(flag, out var factory))
                    SetCompatFlag(compat, factory, false);
            }

            SaveSettings();
            return true;
        }

        // compatibility.mode — set the w:compatSetting for compatibilityMode
        if (key == "compatibility.mode")
        {
            var compat = EnsureCompatibility();
            SetCompatibilityMode(compat, ParseHelpers.SafeParseInt(value, "compatibility.mode"));
            SaveSettings();
            return true;
        }

        // compatibility.<flagName> — individual flag
        if (key.StartsWith("compatibility."))
        {
            var flagName = key["compatibility.".Length..];
            if (!CompatElementFactory.TryGetValue(flagName, out var factory))
                return false;

            var compat = EnsureCompatibility();
            SetCompatFlag(compat, factory, IsTruthy(value));
            SaveSettings();
            return true;
        }

        return false;
    }

    private Compatibility EnsureCompatibility()
    {
        var settings = EnsureSettings();
        var compat = settings.GetFirstChild<Compatibility>();
        if (compat == null)
        {
            compat = new Compatibility();
            settings.AppendChild(compat);
        }
        return compat;
    }

    /// <summary>
    /// Set or remove a compat flag. Uses SetElement to maintain schema order.
    /// </summary>
    private static void SetCompatFlag(Compatibility compat, Func<OnOffType> factory, bool enable)
    {
        var sample = factory();
        var elementType = sample.GetType();

        // Remove existing
        var existing = compat.ChildElements.FirstOrDefault(e => e.GetType() == elementType);
        existing?.Remove();

        if (enable)
        {
            // Use SetElement to insert in schema order
            var newElem = factory();
            compat.AddChild(newElem);
        }
    }

    private static void SetCompatibilityMode(Compatibility compat, int mode)
    {
        // Remove existing compatibilityMode setting
        var existing = compat.Elements<CompatibilitySetting>()
            .FirstOrDefault(cs => cs.Name?.Value == CompatSettingNameValues.CompatibilityMode);
        existing?.Remove();

        compat.AppendChild(new CompatibilitySetting
        {
            Name = new EnumValue<CompatSettingNameValues>(CompatSettingNameValues.CompatibilityMode),
            Val = new StringValue(mode.ToString()),
            Uri = new StringValue("http://schemas.microsoft.com/office/word")
        });
    }

    private void SaveSettings()
    {
        _doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.Save();
    }

    /// <summary>
    /// Read compatibility settings into Format dictionary.
    /// </summary>
    private void PopulateCompatibility(DocumentNode node)
    {
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var compat = settings?.GetFirstChild<Compatibility>();
        if (compat == null) return;

        // Read compatibility mode
        var modeSetting = compat.Elements<CompatibilitySetting>()
            .FirstOrDefault(cs => cs.Name?.Value == CompatSettingNameValues.CompatibilityMode);
        if (modeSetting?.Val?.Value != null)
            node.Format["compatibility.mode"] = int.TryParse(modeSetting.Val.Value, out var m) ? (object)m : modeSetting.Val.Value;

        // Read all OnOffType compat flags that are present
        foreach (var (flagName, factory) in CompatElementFactory)
        {
            var sample = factory();
            var elementType = sample.GetType();
            var element = compat.ChildElements.FirstOrDefault(e => e.GetType() == elementType);
            if (element != null)
            {
                // OnOffType: presence means true, unless val="0" or val="false"
                var onOff = element as OnOffType;
                node.Format[$"compatibility.{flagName}"] = onOff?.Val?.Value ?? true;
            }
        }
    }
}
