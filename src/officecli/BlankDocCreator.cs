// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli;

public static class BlankDocCreator
{
    public static void Create(string path, string? locale = null, bool minimal = false)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".xlsx":
                CreateExcel(path);
                break;
            case ".docx":
                CreateWord(path, locale, minimal);
                break;
            case ".pptx":
                CreatePowerPoint(path);
                break;
            default:
                if (TryCreateViaPlugin(path, ext)) break;
                throw new NotSupportedException($"Unsupported file type: {ext}. Supported: .docx, .xlsx, .pptx, or any extension served by an installed format-handler plugin that implements `create`.");
        }
    }

    /// <summary>
    /// Delegate creation of an unknown extension to a registered format-handler
    /// plugin, if one exists for that extension and exposes a `create &lt;path&gt;`
    /// CLI subcommand. Returns <c>true</c> if a plugin was found and produced
    /// the file successfully. Generic per docs/plugin-protocol.md — keeps
    /// BlankDocCreator format-agnostic; any plugin that implements the
    /// `create` subcommand on its executable participates.
    /// </summary>
    private static bool TryCreateViaPlugin(string path, string ext)
    {
        var plugin = OfficeCli.Core.Plugins.PluginRegistry.FindFor(
            OfficeCli.Core.Plugins.PluginKind.FormatHandler, ext);
        if (plugin is null) return false;
        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName = plugin.ExecutablePath,
            ArgumentList = { "create", System.IO.Path.GetFullPath(path) },
            UseShellExecute = false,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        using var proc = System.Diagnostics.Process.Start(psi);
        if (proc is null) return false;
        var stderr = proc.StandardError.ReadToEnd();
        proc.WaitForExit();
        if (proc.ExitCode != 0)
        {
            // Treat unknown-subcommand exit-64 as "plugin doesn't implement
            // create" — fall back to NotSupportedException so the user sees
            // the same error they'd see without any plugin installed.
            if (proc.ExitCode == 64) return false;
            throw new OfficeCli.Core.CliException(
                $"Format-handler plugin '{plugin.Manifest.Name}' failed to create {path}: {stderr.Trim()}")
            { Code = "plugin_create_failed" };
        }
        return true;
    }

    private static void CreateExcel(string path)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        worksheetPart.Worksheet.Save();

        workbookPart.Workbook = new Workbook(
            new Sheets(
                new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
            )
        );
        workbookPart.Workbook.Save();

        OfficeCliMetadata.StampOnCreate(doc);
    }

    private static void CreateWord(string path, string? locale = null, bool minimal = false)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();

        // Section with A4 page size, standard margins, and no docGrid snap
        var sectPr = new SectionProperties(
            new PageSize { Width = WordPageDefaults.A4WidthTwips, Height = WordPageDefaults.A4HeightTwips },
            new PageMargin { Top = 1440, Right = 1800U, Bottom = 1440, Left = 1800U },
            new DocGrid { Type = DocGridValues.Default }
        );

        // Compatibility: do not compress punctuation spacing
        // Schema order: characterSpacingControl must come before compat in w:settings
        var settings = new DocumentFormat.OpenXml.Wordprocessing.Settings(
            new CharacterSpacingControl { Val = CharacterSpacingValues.DoNotCompress },
            new Compatibility(
                new SpaceForUnderline(),
                new BalanceSingleByteDoubleByteWidth(),
                new DoNotLeaveBackslashAlone(),
                new UnderlineTrailingSpaces(),
                new DoNotExpandShiftReturn(),
                new AdjustLineHeightInTable(),
                new CompatibilitySetting
                {
                    Name = new EnumValue<CompatSettingNameValues>(CompatSettingNameValues.DoNotFlipMirrorIndents),
                    Val = new StringValue("1"),
                    Uri = new StringValue("http://schemas.microsoft.com/office/word")
                }
            )
        );
        // i18n: stamp themeFontLang from --locale so HTML preview, screen
        // readers, and Word's per-script font fallback know
        // the document's primary language. Routes the locale to EastAsia
        // (CJK), Bidi (Arabic / Hebrew / Persian / Urdu / Thai / Hindi),
        // or the bare Val attribute otherwise.
        if (!string.IsNullOrEmpty(locale))
        {
            var tfl = new DocumentFormat.OpenXml.Wordprocessing.ThemeFontLanguages();
            var langKey = locale.Replace('_', '-').ToLowerInvariant().Split('-')[0];
            switch (langKey)
            {
                case "zh":
                case "ja":
                case "ko":
                    tfl.EastAsia = locale;
                    break;
                case "ar":
                case "he":
                case "fa":
                case "ur":
                case "th":
                case "hi":
                    tfl.Bidi = locale;
                    break;
                default:
                    tfl.Val = locale;
                    break;
            }
            // ThemeFontLanguages must precede characterSpacingControl per
            // CT_Settings sequence — InsertBefore the existing first child.
            var firstChild = settings.FirstChild;
            if (firstChild != null) settings.InsertBefore(tfl, firstChild);
            else settings.AppendChild(tfl);
        }
        var settingsPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.DocumentSettingsPart>();
        settingsPart.Settings = settings;
        settingsPart.Settings.Save();

        var document = new Document(new Body(sectPr));
        // Declare common namespaces on <w:document> so later raw-set
        // injections (DrawingML textboxes <wps:wsp>, VML fallbacks <v:shape>,
        // pictures <pic:pic>, math <m:oMath>, ...) validate without each
        // call site re-declaring them. Mirrors what Word itself stamps on
        // save. Without this, mc:AlternateContent / mc:Choice Requires="wps"
        // fails MarkupCompatibility validation because the wps prefix is
        // not in scope at the AlternateContent element.
        document.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        document.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
        document.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
        document.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
        document.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
        document.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
        document.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        document.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
        document.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        document.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
        document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        document.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
        document.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
        document.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
        document.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
        // Mark 2010+/2012 namespaces as Ignorable so older readers degrade gracefully.
        document.MCAttributes ??= new DocumentFormat.OpenXml.MarkupCompatibilityAttributes();
        var existingIgnorable = document.MCAttributes.Ignorable?.Value;
        var ignorableTokens = new System.Collections.Generic.HashSet<string>(
            (existingIgnorable ?? "").Split(' ', System.StringSplitOptions.RemoveEmptyEntries));
        // Only mark prefixes that appear unwrapped (outside mc:AlternateContent)
        // as Ignorable — w14/wp14/w15 carry attributes like paraId/anchorId
        // directly. wps/wpg/wpi/wpc only appear inside mc:Choice and are
        // already gated by mc:Fallback, so they don't need (and shouldn't get)
        // Ignorable. Mirrors the docxexport MainXmlNamespaces.
        foreach (var p in new[] { "w14", "w15", "wp14" })
            ignorableTokens.Add(p);
        document.MCAttributes.Ignorable = string.Join(" ", ignorableTokens);
        mainPart.Document = document;

        // Two paths: full (default) emits Word-aligned baseline (Calibri 11pt
        // + Normal style + theme1.xml — matches the de-facto baseline, which
        // is what Word actually writes); minimal emits raw OOXML (TNR, no sz,
        // no Normal, no theme — matches POI's `new XWPFDocument()`). The
        // minimal path is the prior officecli behavior; the full path was
        // added so docs created by officecli render identically in Word /
        // / cli preview without relying on each renderer's
        // Normal.dotm fallback heuristics.
        //
        // Resolve locale-specific defaults from LocaleFontRegistry (POI/LO
        // pattern). Without a locale, only Latin slots are populated so the
        // host application's UI-locale defaults fill EastAsia / CS as needed.
        var (locLatin, locEa, locCs) = OfficeCli.Core.LocaleFontRegistry.Resolve(locale);

        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        if (minimal)
        {
            // Minimal path: docDefaults with rFonts only (Times New Roman),
            // no sz, no spacing, no Normal style, no theme. Use this for
            // testing the cli reader's fallback path or producing maximally
            // compact output. Matches `officecli create` output before the
            // Word-aligned baseline was added.
            var minDocDefaultFonts = new RunFonts
            {
                Ascii = locLatin ?? "Times New Roman",
                HighAnsi = locLatin ?? "Times New Roman",
            };
            if (!string.IsNullOrEmpty(locEa)) minDocDefaultFonts.EastAsia = locEa;
            if (!string.IsNullOrEmpty(locCs)) minDocDefaultFonts.ComplexScript = locCs;
            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(minDocDefaultFonts)),
                    new ParagraphPropertiesDefault()
                )
            );
            stylesPart.Styles.Save();
        }
        else
        {
            var docDefaultFonts = new RunFonts
            {
                Ascii = locLatin ?? OfficeDefaultFonts.MinorLatin,    // Calibri
                HighAnsi = locLatin ?? OfficeDefaultFonts.MinorLatin,
            };
            if (!string.IsNullOrEmpty(locEa)) docDefaultFonts.EastAsia = locEa;
            if (!string.IsNullOrEmpty(locCs)) docDefaultFonts.ComplexScript = locCs;

            // Normal style — default="1". Carry the Office 2013+ Normal
            // baseline (line=259/1.08 ×, no after) on the Normal pPr itself,
            // not on pPrDefault — cli's reader only walks the style chain via
            // ResolveSpacingFromStyle and doesn't yet inherit from pPrDefault.
            // Putting it on Normal keeps pPrDefault free for paragraph-shape
            // defaults (autoSpaceDE/DN, kinsoku, …) without spacing leakage.
            //
            // Why 1.08 × not 1.15 ×: empirical (stress-C measurement) — when
            // a list line has a 14 pt marker over 11 pt body, Word renders
            // the line at 14 × 1.08 × calibri-ratio = 18.45pt; cli with
            // 1.15 × renders at 14 × 1.15 × ratio = 19.65pt (1.3pt/paragraph drift
            // accumulating across the doc). Office 2013+ Normal IS 1.08 ×;
            // matching that here matches what Word actually does.
            var normalStyle = new Style(
                new StyleName { Val = "Normal" },
                new PrimaryStyle(),
                new StyleParagraphProperties(
                    new SpacingBetweenLines
                    {
                        After = "0",
                        Line = "259",
                        LineRule = LineSpacingRuleValues.Auto,
                    }
                )
            )
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true,
            };

            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(
                        new RunPropertiesBaseStyle(
                            docDefaultFonts,
                            new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "22" },              // 11pt
                            new FontSizeComplexScript { Val = "22" }
                        )
                    ),
                    new ParagraphPropertiesDefault()
                ),
                normalStyle
            );
            stylesPart.Styles.Save();
        }

        // theme1.xml — Office's minor=Calibri / major=Calibri Light. Without
        // a theme part, anything that looks up `themeFonts` (heading/body
        // theme references in styles.xml) gets nothing — emit a minimal
        // theme so future styles can reference it. Skipped on the minimal
        // path so its output stays free of theme dependencies.
        if (!minimal)
        {
        var themePart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemePart>();
        themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
            new DocumentFormat.OpenXml.Drawing.ThemeElements(
                new DocumentFormat.OpenXml.Drawing.ColorScheme(
                    new DocumentFormat.OpenXml.Drawing.Dark1Color(new DocumentFormat.OpenXml.Drawing.SystemColor { Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
                    new DocumentFormat.OpenXml.Drawing.Light1Color(new DocumentFormat.OpenXml.Drawing.SystemColor { Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new DocumentFormat.OpenXml.Drawing.Dark2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Dark2 }),
                    new DocumentFormat.OpenXml.Drawing.Light2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Light2 }),
                    new DocumentFormat.OpenXml.Drawing.Accent1Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent1 }),
                    new DocumentFormat.OpenXml.Drawing.Accent2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent2 }),
                    new DocumentFormat.OpenXml.Drawing.Accent3Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent3 }),
                    new DocumentFormat.OpenXml.Drawing.Accent4Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent4 }),
                    new DocumentFormat.OpenXml.Drawing.Accent5Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent5 }),
                    new DocumentFormat.OpenXml.Drawing.Accent6Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent6 }),
                    new DocumentFormat.OpenXml.Drawing.Hyperlink(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Hyperlink }),
                    new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.FollowedHyperlink })
                ) { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FontScheme(
                    new DocumentFormat.OpenXml.Drawing.MajorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = OfficeDefaultFonts.MajorLatin },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = locEa ?? "" },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = locCs ?? "" }
                    ),
                    new DocumentFormat.OpenXml.Drawing.MinorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = OfficeDefaultFonts.MinorLatin },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = locEa ?? "" },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = locCs ?? "" }
                    )
                ) { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FormatScheme(
                    new DocumentFormat.OpenXml.Drawing.FillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })
                    ),
                    new DocumentFormat.OpenXml.Drawing.LineStyleList(
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 6350, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat },
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 12700, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat },
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 19050, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat }
                    ),
                    new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList())
                    ),
                    new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })
                    )
                ) { Name = "Office" }
            )
        ) { Name = "Office Theme" };
        themePart.Theme.Save();
        }

        var numberingPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.NumberingDefinitionsPart>();
        numberingPart.Numbering = new DocumentFormat.OpenXml.Wordprocessing.Numbering();
        numberingPart.Numbering.Save();
        mainPart.Document.Save();

        OfficeCliMetadata.StampOnCreate(doc);
    }

    private static void CreatePowerPoint(string path)
    {
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation);
        var presentationPart = doc.AddPresentationPart();

        // Create SlideMaster + SlideLayout (required by spec)
        var slideMasterPart = presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideMasterPart>("rId1");
        var slideLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>("rId1");

        // Theme must be under presentationPart, then shared to slideMaster
        var themePart = presentationPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemePart>("rId2");
        slideMasterPart.AddPart(themePart);
        themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
            new DocumentFormat.OpenXml.Drawing.ThemeElements(
                new DocumentFormat.OpenXml.Drawing.ColorScheme(
                    new DocumentFormat.OpenXml.Drawing.Dark1Color(new DocumentFormat.OpenXml.Drawing.SystemColor { Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.WindowText, LastColor = "000000" }),
                    new DocumentFormat.OpenXml.Drawing.Light1Color(new DocumentFormat.OpenXml.Drawing.SystemColor { Val = DocumentFormat.OpenXml.Drawing.SystemColorValues.Window, LastColor = "FFFFFF" }),
                    new DocumentFormat.OpenXml.Drawing.Dark2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Dark2 }),
                    new DocumentFormat.OpenXml.Drawing.Light2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Light2 }),
                    new DocumentFormat.OpenXml.Drawing.Accent1Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent1 }),
                    new DocumentFormat.OpenXml.Drawing.Accent2Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent2 }),
                    new DocumentFormat.OpenXml.Drawing.Accent3Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent3 }),
                    new DocumentFormat.OpenXml.Drawing.Accent4Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent4 }),
                    new DocumentFormat.OpenXml.Drawing.Accent5Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent5 }),
                    new DocumentFormat.OpenXml.Drawing.Accent6Color(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Accent6 }),
                    new DocumentFormat.OpenXml.Drawing.Hyperlink(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.Hyperlink }),
                    new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = OfficeDefaultThemeColors.FollowedHyperlink })
                ) { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FontScheme(
                    new DocumentFormat.OpenXml.Drawing.MajorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = OfficeDefaultFonts.MajorLatin },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = "" },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = "" }
                    ),
                    new DocumentFormat.OpenXml.Drawing.MinorFont(
                        new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = OfficeDefaultFonts.MinorLatin },
                        new DocumentFormat.OpenXml.Drawing.EastAsianFont { Typeface = "" },
                        new DocumentFormat.OpenXml.Drawing.ComplexScriptFont { Typeface = "" }
                    )
                ) { Name = "Office" },
                new DocumentFormat.OpenXml.Drawing.FormatScheme(
                    new DocumentFormat.OpenXml.Drawing.FillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })
                    ),
                    new DocumentFormat.OpenXml.Drawing.LineStyleList(
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 6350, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat },
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 12700, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat },
                        new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })) { Width = 19050, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat }
                    ),
                    new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList()),
                        new DocumentFormat.OpenXml.Drawing.EffectStyle(new DocumentFormat.OpenXml.Drawing.EffectList())
                    ),
                    new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor }),
                        new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.PhColor })
                    )
                ) { Name = "Office" }
            )
        ) { Name = "Office Theme" };
        themePart.Theme.Save();

        // Layout 1: Blank
        slideLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties()
                )
            ) { Name = "Blank" }
        ) { Type = DocumentFormat.OpenXml.Presentation.SlideLayoutValues.Blank };
        slideLayoutPart.SlideLayout.Save();
        slideLayoutPart.AddPart(slideMasterPart);

        // Layout 2: Title Slide (title + subtitle)
        var titleLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>("rId2");
        titleLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(),
                    CreateLayoutPlaceholder(2, "Title", PlaceholderValues.CenteredTitle, 685800, 2130425, 7772400, 1470025),
                    CreateLayoutPlaceholder(3, "Subtitle", PlaceholderValues.SubTitle, 1371600, 3886200, 6400800, 1752600, idx: 1)
                )
            ) { Name = "Title Slide" }
        ) { Type = DocumentFormat.OpenXml.Presentation.SlideLayoutValues.Title };
        titleLayoutPart.SlideLayout.Save();
        titleLayoutPart.AddPart(slideMasterPart);

        // Layout 3: Title and Content
        var contentLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>("rId3");
        contentLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(),
                    CreateLayoutPlaceholder(2, "Title", PlaceholderValues.Title, 838200, 365125, 10515600, 1325563),
                    CreateLayoutPlaceholder(3, "Content", PlaceholderValues.Body, 838200, 1825625, 10515600, 4351338, idx: 1)
                )
            ) { Name = "Title and Content" }
        ) { Type = DocumentFormat.OpenXml.Presentation.SlideLayoutValues.ObjectText };
        contentLayoutPart.SlideLayout.Save();
        contentLayoutPart.AddPart(slideMasterPart);

        // Layout 4: Two Content
        var twoContentLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>("rId4");
        twoContentLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(),
                    CreateLayoutPlaceholder(2, "Title", PlaceholderValues.Title, 838200, 365125, 10515600, 1325563),
                    CreateLayoutPlaceholder(3, "Content Left", PlaceholderValues.Body, 838200, 1825625, 5181600, 4351338, idx: 1),
                    CreateLayoutPlaceholder(4, "Content Right", PlaceholderValues.Body, 6172200, 1825625, 5181600, 4351338, idx: 2)
                )
            ) { Name = "Two Content" }
        ) { Type = DocumentFormat.OpenXml.Presentation.SlideLayoutValues.TwoColumnText };
        twoContentLayoutPart.SlideLayout.Save();
        twoContentLayoutPart.AddPart(slideMasterPart);

        // Layout 5: Title Only (title placeholder, no body)
        var titleOnlyLayoutPart = slideMasterPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SlideLayoutPart>("rId5");
        titleOnlyLayoutPart.SlideLayout = new DocumentFormat.OpenXml.Presentation.SlideLayout(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties(),
                    CreateLayoutPlaceholder(2, "Title", PlaceholderValues.Title, 838200, 365125, 10515600, 1325563)
                )
            ) { Name = "Title Only" }
        ) { Type = DocumentFormat.OpenXml.Presentation.SlideLayoutValues.TitleOnly };
        titleOnlyLayoutPart.SlideLayout.Save();
        titleOnlyLayoutPart.AddPart(slideMasterPart);

        slideMasterPart.SlideMaster = new DocumentFormat.OpenXml.Presentation.SlideMaster(
            new DocumentFormat.OpenXml.Presentation.CommonSlideData(
                new DocumentFormat.OpenXml.Presentation.ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = 1, Name = "" },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
                    ),
                    new DocumentFormat.OpenXml.Presentation.GroupShapeProperties()
                )
            ),
            new DocumentFormat.OpenXml.Presentation.ColorMap
            {
                Background1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light1,
                Text1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark1,
                Background2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light2,
                Text2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark2,
                Accent1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent1,
                Accent2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent2,
                Accent3 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent3,
                Accent4 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent4,
                Accent5 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent5,
                Accent6 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent6,
                Hyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.FollowedHyperlink,
            },
            new DocumentFormat.OpenXml.Presentation.SlideLayoutIdList(
                new DocumentFormat.OpenXml.Presentation.SlideLayoutId { Id = 2147483649, RelationshipId = "rId1" },
                new DocumentFormat.OpenXml.Presentation.SlideLayoutId { Id = 2147483650, RelationshipId = "rId2" },
                new DocumentFormat.OpenXml.Presentation.SlideLayoutId { Id = 2147483651, RelationshipId = "rId3" },
                new DocumentFormat.OpenXml.Presentation.SlideLayoutId { Id = 2147483652, RelationshipId = "rId4" },
                new DocumentFormat.OpenXml.Presentation.SlideLayoutId { Id = 2147483653, RelationshipId = "rId5" }
            )
        );
        slideMasterPart.SlideMaster.Save();

        presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation(
            new DocumentFormat.OpenXml.Presentation.SlideMasterIdList(
                new DocumentFormat.OpenXml.Presentation.SlideMasterId { Id = 2147483648, RelationshipId = "rId1" }
            ),
            new SlideIdList(),
            new SlideSize { Cx = (int)SlideSizeDefaults.Widescreen16x9Cx, Cy = (int)SlideSizeDefaults.Widescreen16x9Cy },
            new NotesSize { Cx = SlideSizeDefaults.NotesPortraitCx, Cy = SlideSizeDefaults.NotesPortraitCy }
        );
        presentationPart.Presentation.Save();

        OfficeCliMetadata.StampOnCreate(doc);
    }

    private static Shape CreateLayoutPlaceholder(uint id, string name, PlaceholderValues phType,
        long x, long y, long cx, long cy, uint? idx = null)
    {
        var shape = new Shape();
        // OOXML convention (PowerPoint templates): Title/CenteredTitle placeholders
        // omit @idx (defaults to 0); SubTitle / Body / Footer / Date / SlideNumber
        // slots carry an explicit @idx so slide-side <p:ph idx="N"/> can bind back to
        // the matching layout placeholder during inheritance resolution.
        var placeholder = new PlaceholderShape { Type = phType };
        if (idx.HasValue) placeholder.Index = idx.Value;
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            new NonVisualShapeDrawingProperties(new DocumentFormat.OpenXml.Drawing.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(placeholder)
        );
        shape.ShapeProperties = new ShapeProperties(
            new DocumentFormat.OpenXml.Drawing.Transform2D(
                new DocumentFormat.OpenXml.Drawing.Offset { X = x, Y = y },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = cx, Cy = cy }
            )
        );
        shape.TextBody = new TextBody(
            new DocumentFormat.OpenXml.Drawing.BodyProperties(),
            new DocumentFormat.OpenXml.Drawing.ListStyle(),
            new DocumentFormat.OpenXml.Drawing.Paragraph(
                new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties { Language = "en-US" })
        );
        return shape;
    }
}
