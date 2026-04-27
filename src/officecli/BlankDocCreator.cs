// Copyright 2025 OfficeCli (officecli.ai)
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
    public static void Create(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".xlsx":
                CreateExcel(path);
                break;
            case ".docx":
                CreateWord(path);
                break;
            case ".pptx":
                CreatePowerPoint(path);
                break;
            default:
                throw new NotSupportedException($"Unsupported file type: {ext}. Supported: .docx, .xlsx, .pptx");
        }
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
    }

    private static void CreateWord(string path)
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
        var settingsPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.DocumentSettingsPart>();
        settingsPart.Settings = settings;
        settingsPart.Settings.Save();

        mainPart.Document = new Document(new Body(sectPr));

        // Default paragraph style: disable CJK-specific layout
        var stylesPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(
            new DocDefaults(
                new ParagraphPropertiesDefault(
                    new ParagraphPropertiesBaseStyle
                    {
                        AutoSpaceDE = new AutoSpaceDE { Val = false },
                        AutoSpaceDN = new AutoSpaceDN { Val = false },
                        Kinsoku = new DocumentFormat.OpenXml.Wordprocessing.Kinsoku { Val = false },
                        OverflowPunctuation = new OverflowPunctuation { Val = false },
                    }
                )
            )
        );
        stylesPart.Styles.Save();

        var numberingPart = mainPart.AddNewPart<DocumentFormat.OpenXml.Packaging.NumberingDefinitionsPart>();
        numberingPart.Numbering = new DocumentFormat.OpenXml.Wordprocessing.Numbering();
        numberingPart.Numbering.Save();
        mainPart.Document.Save();
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
                    CreateLayoutPlaceholder(3, "Subtitle", PlaceholderValues.SubTitle, 1371600, 3886200, 6400800, 1752600)
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
                    CreateLayoutPlaceholder(3, "Content", PlaceholderValues.Body, 838200, 1825625, 10515600, 4351338)
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
                    CreateLayoutPlaceholder(3, "Content Left", PlaceholderValues.Body, 838200, 1825625, 5181600, 4351338),
                    CreateLayoutPlaceholder(4, "Content Right", PlaceholderValues.Body, 6172200, 1825625, 5181600, 4351338)
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
    }

    private static Shape CreateLayoutPlaceholder(uint id, string name, PlaceholderValues phType,
        long x, long y, long cx, long cy)
    {
        var shape = new Shape();
        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
            new NonVisualDrawingProperties { Id = id, Name = name },
            new NonVisualShapeDrawingProperties(new DocumentFormat.OpenXml.Drawing.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = phType })
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
