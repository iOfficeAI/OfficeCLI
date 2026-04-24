// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// Set document-level settings: DocGrid, CJK layout, print/display, font embedding, layout flags, defaultTabStop.
    /// Called from SetDocumentProperties for keys with recognized names.
    /// Returns true if the key was handled.
    /// </summary>
    private bool TrySetDocSetting(string key, string value)
    {
        switch (key)
        {
            // ==================== DocGrid (lives in SectionProperties) ====================
            case "docgrid.type":
            {
                var grid = EnsureDocGridInSection();
                grid.Type = value.ToLowerInvariant() switch
                {
                    "default" or "none" => DocGridValues.Default,
                    "lines" => DocGridValues.Lines,
                    "linesandchars" or "linesandcharacters" => DocGridValues.LinesAndChars,
                    "snaptocharacters" or "snapchars" => DocGridValues.SnapToChars,
                    _ => throw new ArgumentException($"Invalid docGrid.type: '{value}'. Valid: default, lines, linesAndChars, snapToCharacters")
                };
                return true;
            }
            case "docgrid.linepitch":
            {
                var grid = EnsureDocGridInSection();
                grid.LinePitch = ParseHelpers.SafeParseInt(value, "docGrid.linePitch");
                return true;
            }
            case "docgrid.charspace" or "docgrid.characterspace":
            {
                var grid = EnsureDocGridInSection();
                grid.CharacterSpace = ParseHelpers.SafeParseInt(value, "docGrid.charSpace");
                return true;
            }

            // ==================== CJK Layout (lives in DocDefaults ParagraphProperties) ====================
            case "autospacede":
                SetParaDefault_AutoSpaceDE(IsTruthy(value));
                return true;
            case "autospacedn":
                SetParaDefault_AutoSpaceDN(IsTruthy(value));
                return true;
            case "kinsoku":
                SetParaDefault_Kinsoku(IsTruthy(value));
                return true;
            case "overflowpunct" or "overflowpunctuation":
                SetParaDefault_OverflowPunctuation(IsTruthy(value));
                return true;

            // ==================== CharacterSpacingControl (lives in Settings) ====================
            case "charspacingcontrol" or "characterspacingcontrol":
            {
                var settings = EnsureSettings();
                settings.GetFirstChild<CharacterSpacingControl>()?.Remove();
                var csc = new CharacterSpacingControl
                {
                    Val = value.ToLowerInvariant() switch
                    {
                        "donotcompress" or "none" => CharacterSpacingValues.DoNotCompress,
                        "compresspunctuation" or "punctuation" => CharacterSpacingValues.CompressPunctuation,
                        "compresspunctuationandjapanesekana" or "all" => CharacterSpacingValues.CompressPunctuationAndJapaneseKana,
                        _ => throw new ArgumentException($"Invalid charSpacingControl: '{value}'. Valid: doNotCompress, compressPunctuation, compressPunctuationAndJapaneseKana")
                    }
                };
                settings.AddChild(csc);
                EnsureSettings().Save();
                return true;
            }

            // ==================== Print / Display (lives in Settings) ====================
            case "displaybackgroundshape":
                SetOnOffSetting<DisplayBackgroundShape>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "donotdisplaypageboundaries":
                SetOnOffSetting<DoNotDisplayPageBoundaries>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "printformsdata":
                SetOnOffSetting<PrintFormsData>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "printpostscriptovertext":
                SetOnOffSetting<PrintPostScriptOverText>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "printfractionalcharacterwidth":
                SetOnOffSetting<PrintFractionalCharacterWidth>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "removepersonalinformation" or "removepersonalinfo":
                SetOnOffSetting<RemovePersonalInformation>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "removedateandtime":
                SetOnOffSetting<RemoveDateAndTime>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;

            // ==================== Font Embedding (lives in Settings) ====================
            case "embedfonts" or "embedtruetypefonts":
                SetOnOffSetting<EmbedTrueTypeFonts>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "embedsystemfonts":
                SetOnOffSetting<EmbedSystemFonts>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "savesubsetfonts":
                SetOnOffSetting<SaveSubsetFonts>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;

            // ==================== Review / Tracked Changes (lives in Settings) ====================
            // Toggle the top-level "Track Changes" flag independently of any
            // existing w:ins / w:del markup. Callers typically combine this
            // with acceptAllChanges / rejectAllChanges to finalize a review
            // session and take the document out of tracked-changes mode.
            case "trackrevisions" or "trackchanges":
                SetOnOffSetting<TrackRevisions>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;

            // ==================== Layout Flags (lives in Settings) ====================
            case "mirrormargins":
                SetOnOffSetting<MirrorMargins>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "gutterattop":
                SetOnOffSetting<GutterAtTop>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "bookfoldprinting":
                SetOnOffSetting<BookFoldPrinting>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "bookfoldreverseprinting":
                SetOnOffSetting<BookFoldReversePrinting>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "bookfoldprintingsheets":
            {
                var settings = EnsureSettings();
                settings.GetFirstChild<BookFoldPrintingSheets>()?.Remove();
                // Treat "false", "0", empty as remove; otherwise parse as int
                if (!string.IsNullOrEmpty(value) && value != "0" && !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
                    settings.AddChild(new BookFoldPrintingSheets { Val = (short)ParseHelpers.SafeParseInt(value, "bookFoldPrintingSheets") });
                settings.Save();
                return true;
            }
            case "evenandoddheaders":
                SetOnOffSetting<EvenAndOddHeaders>(EnsureSettings(), IsTruthy(value));
                EnsureSettings().Save();
                return true;
            case "defaulttabstop":
            {
                var settings = EnsureSettings();
                var twips = ParseTwips(value);
                if (twips > short.MaxValue)
                    throw new ArgumentException($"defaultTabStop value too large: {value} ({twips} twips, max {short.MaxValue})");
                settings.GetFirstChild<DefaultTabStop>()?.Remove();
                // AddChild respects OOXML schema particle order on composite elements
                settings.AddChild(new DefaultTabStop { Val = (short)twips });
                settings.Save();
                return true;
            }

            default:
                return false;
        }
    }

    // ==================== DocGrid Helper ====================

    private DocGrid EnsureDocGridInSection()
    {
        var sectPr = EnsureSectionProperties();
        var grid = sectPr.GetFirstChild<DocGrid>();
        if (grid == null)
        {
            grid = new DocGrid();
            sectPr.AppendChild(grid);
        }
        return grid;
    }

    // ==================== ParagraphPropertiesDefault Helpers ====================

    private ParagraphPropertiesBaseStyle EnsureParaPropsDefault()
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

        var pPrDefault = docDefaults.ParagraphPropertiesDefault;
        if (pPrDefault == null)
        {
            pPrDefault = new ParagraphPropertiesDefault();
            docDefaults.AppendChild(pPrDefault);
        }

        var pPrBase = pPrDefault.ParagraphPropertiesBaseStyle;
        if (pPrBase == null)
        {
            pPrBase = new ParagraphPropertiesBaseStyle();
            pPrDefault.ParagraphPropertiesBaseStyle = pPrBase;
        }

        return pPrBase;
    }

    private void SetParaDefault_AutoSpaceDE(bool value)
    {
        var pPr = EnsureParaPropsDefault();
        pPr.AutoSpaceDE = new AutoSpaceDE { Val = value };
        _doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Save();
    }

    private void SetParaDefault_AutoSpaceDN(bool value)
    {
        var pPr = EnsureParaPropsDefault();
        pPr.AutoSpaceDN = new AutoSpaceDN { Val = value };
        _doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Save();
    }

    private void SetParaDefault_Kinsoku(bool value)
    {
        var pPr = EnsureParaPropsDefault();
        pPr.Kinsoku = new Kinsoku { Val = value };
        _doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Save();
    }

    private void SetParaDefault_OverflowPunctuation(bool value)
    {
        var pPr = EnsureParaPropsDefault();
        pPr.OverflowPunctuation = new OverflowPunctuation { Val = value };
        _doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Save();
    }

    // ==================== Generic OnOff Setting Helper ====================

    /// <summary>
    /// Set or remove an OnOffType child element in Settings.
    /// When value is true, ensures the element exists in schema-correct position
    /// (before Compatibility, which must be near the end of w:settings).
    /// When false, removes it.
    /// </summary>
    private static void SetOnOffSetting<T>(Settings settings, bool value) where T : OnOffType, new()
    {
        var existing = settings.GetFirstChild<T>();
        existing?.Remove();
        if (value)
            settings.AddChild(new T()); // AddChild respects OOXML schema particle order
    }

    /// <summary>
    /// Insert an element at the schema-correct position in w:settings.
    /// Most settings elements must precede w:charSpacingControl and w:compat in the OOXML schema.
    /// Inserts before the first of CharacterSpacingControl or Compatibility if present,
    /// otherwise appends.
    /// </summary>
    private static void InsertBeforeCompatibility(Settings settings, DocumentFormat.OpenXml.OpenXmlElement elem)
    {
        // Find the earliest anchor (charSpacingControl comes before compat in schema,
        // and most other settings come before charSpacingControl)
        var anchor = (DocumentFormat.OpenXml.OpenXmlElement?)settings.GetFirstChild<CharacterSpacingControl>()
            ?? settings.GetFirstChild<Compatibility>();
        if (anchor != null)
            anchor.InsertBeforeSelf(elem);
        else
            settings.AppendChild(elem);
    }

    private Settings EnsureSettings()
    {
        var mainPart = _doc.MainDocumentPart!;
        var settingsPart = mainPart.DocumentSettingsPart
            ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        return settingsPart.Settings;
    }
}
