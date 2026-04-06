// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddSection(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // Section break: adds SectionProperties to the last paragraph before the break point
        var breakType = properties.GetValueOrDefault("type", "nextPage").ToLowerInvariant();
        var sectType = breakType switch
        {
            "nextpage" or "next" => SectionMarkValues.NextPage,
            "continuous" => SectionMarkValues.Continuous,
            "evenpage" or "even" => SectionMarkValues.EvenPage,
            "oddpage" or "odd" => SectionMarkValues.OddPage,
            _ => throw new ArgumentException($"Invalid section break type: '{breakType}'. Valid values: nextPage, continuous, evenPage, oddPage.")
        };

        // Create a paragraph with section properties to mark the break
        var sectPara = new Paragraph();
        var sectPProps = new ParagraphProperties();
        var sectPr = new SectionProperties();
        sectPr.AppendChild(new SectionType { Val = sectType });

        // Ensure body-level sectPr has pgSz/pgMar (fix for docs created by older versions)
        var bodySectPr = body.GetFirstChild<SectionProperties>();
        if (bodySectPr != null && bodySectPr.GetFirstChild<PageSize>() == null)
        {
            bodySectPr.InsertBefore(new PageSize { Width = 11906, Height = 16838 },
                bodySectPr.GetFirstChild<DocGrid>());
        }
        if (bodySectPr != null && bodySectPr.GetFirstChild<PageMargin>() == null)
        {
            bodySectPr.InsertBefore(new PageMargin { Top = 1440, Right = 1800U, Bottom = 1440, Left = 1800U },
                bodySectPr.GetFirstChild<DocGrid>());
        }

        // Copy page size/margins from document section, or use A4 defaults
        var srcPageSize = bodySectPr?.GetFirstChild<PageSize>();
        sectPr.AppendChild(new PageSize
        {
            Width = srcPageSize?.Width ?? 11906,   // A4 width
            Height = srcPageSize?.Height ?? 16838,  // A4 height
            Orient = srcPageSize?.Orient
        });
        var srcMargin = bodySectPr?.GetFirstChild<PageMargin>();
        sectPr.AppendChild(new PageMargin
        {
            Top = srcMargin?.Top ?? 1440,
            Bottom = srcMargin?.Bottom ?? 1440,
            Left = srcMargin?.Left ?? 1800,
            Right = srcMargin?.Right ?? 1800
        });

        // Allow per-section overrides
        if (properties.TryGetValue("pagewidth", out var sw) || properties.TryGetValue("pageWidth", out sw) || properties.TryGetValue("width", out sw))
        {
            (sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize())).Width = ParseTwips(sw);
        }
        if (properties.TryGetValue("pageheight", out var sh) || properties.TryGetValue("pageHeight", out sh) || properties.TryGetValue("height", out sh))
        {
            (sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize())).Height = ParseTwips(sh);
        }
        if (properties.TryGetValue("orientation", out var orient))
        {
            var ps = sectPr.GetFirstChild<PageSize>() ?? sectPr.AppendChild(new PageSize());
            ps.Orient = orient.ToLowerInvariant() == "landscape"
                ? PageOrientationValues.Landscape
                : PageOrientationValues.Portrait;
            // Swap width/height if dimensions don't match orientation
            if (ps.Orient == PageOrientationValues.Landscape && ps.Width < ps.Height)
                (ps.Width!.Value, ps.Height!.Value) = (ps.Height.Value, ps.Width.Value);
            if (ps.Orient == PageOrientationValues.Portrait && ps.Width > ps.Height)
                (ps.Width!.Value, ps.Height!.Value) = (ps.Height.Value, ps.Width.Value);
        }

        // Columns support: "columns=2" or "columns=2,1cm"
        if (properties.TryGetValue("columns", out var colsVal) || properties.TryGetValue("columns.count", out colsVal))
        {
            var parts = colsVal.Split(',');
            var count = (short)int.Parse(parts[0].Trim());
            var cols = new Columns { ColumnCount = count, EqualWidth = true };
            if (parts.Length > 1)
                cols.Space = ParseTwips(parts[1].Trim()).ToString();
            sectPr.AppendChild(cols);
        }
        if (properties.TryGetValue("columns.space", out var colSpace))
        {
            var cols = sectPr.GetFirstChild<Columns>() ?? sectPr.AppendChild(new Columns());
            cols.Space = ParseTwips(colSpace).ToString();
        }

        sectPProps.AppendChild(sectPr);
        sectPara.AppendChild(sectPProps);
        AppendToParent(parent, sectPara);

        // Count section properties in document
        var secCount = body.Elements<Paragraph>()
            .Count(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null);
        var resultPath = $"/section[{secCount}]";
        return resultPath;
    }

    private string AddFootnote(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("text", out var fnText))
            throw new ArgumentException("'text' property is required for footnote type");

        if (parent is not Paragraph fnPara)
            throw new ArgumentException("Footnotes must be added to a paragraph: /body/p[N]");

        var mainPart2 = _doc.MainDocumentPart!;
        var fnPart = mainPart2.FootnotesPart ?? mainPart2.AddNewPart<FootnotesPart>();
        fnPart.Footnotes ??= new Footnotes(
            new Footnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Footnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
        );

        var fnId = (fnPart.Footnotes.Elements<Footnote>()
            .Where(f => f.Id?.Value > 0)
            .Select(f => f.Id!.Value)
            .DefaultIfEmpty(0).Max() + 1);

        var footnote = new Footnote { Id = fnId };
        var fnContentPara = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
            new Run(
                new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                new FootnoteReferenceMark()),
            new Run(new Text(" " + fnText) { Space = SpaceProcessingModeValues.Preserve })
        );
        footnote.AppendChild(fnContentPara);
        fnPart.Footnotes.AppendChild(footnote);
        fnPart.Footnotes.Save();

        // Insert reference in document body
        var fnRefRun = new Run(
            new RunProperties(new RunStyle { Val = "FootnoteReference" }),
            new FootnoteReference { Id = fnId }
        );
        fnPara.AppendChild(fnRefRun);

        var resultPath = $"/footnote[{fnId}]";
        return resultPath;
    }

    private string AddEndnote(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("text", out var enText))
            throw new ArgumentException("'text' property is required for endnote type");

        if (parent is not Paragraph enPara)
            throw new ArgumentException("Endnotes must be added to a paragraph: /body/p[N]");

        var mainPart3 = _doc.MainDocumentPart!;
        var enPart = mainPart3.EndnotesPart ?? mainPart3.AddNewPart<EndnotesPart>();
        enPart.Endnotes ??= new Endnotes(
            new Endnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
            new Endnote(new Paragraph(new Run(new Text("")))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 }
        );

        var enId = (enPart.Endnotes.Elements<Endnote>()
            .Where(e => e.Id?.Value > 0)
            .Select(e => e.Id!.Value)
            .DefaultIfEmpty(0).Max() + 1);

        var endnote = new Endnote { Id = enId };
        var enContentPara = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "EndnoteText" }),
            new Run(
                new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }),
                new EndnoteReferenceMark()),
            new Run(new Text(" " + enText) { Space = SpaceProcessingModeValues.Preserve })
        );
        endnote.AppendChild(enContentPara);
        enPart.Endnotes.AppendChild(endnote);
        enPart.Endnotes.Save();

        // Insert reference in document body
        var enRefRun = new Run(
            new RunProperties(new RunStyle { Val = "EndnoteReference" }),
            new EndnoteReference { Id = enId }
        );
        enPara.AppendChild(enRefRun);

        var resultPath = $"/endnote[{enId}]";
        return resultPath;
    }

    private string AddToc(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // Table of Contents field code
        var levels = properties.GetValueOrDefault("levels", "1-3");
        var tocTitle = properties.GetValueOrDefault("title", "");
        var hyperlinks = !properties.TryGetValue("hyperlinks", out var hlVal) || IsTruthy(hlVal);
        var pageNumbers = !properties.TryGetValue("pagenumbers", out var pnVal) || IsTruthy(pnVal);

        // Build field code instruction
        var instrBuilder = new StringBuilder($" TOC \\o \"{levels}\"");
        if (hyperlinks) instrBuilder.Append(" \\h");
        if (!pageNumbers) instrBuilder.Append(" \\z");
        instrBuilder.Append(" \\u ");

        var tocPara = new Paragraph();

        // Optional title
        if (!string.IsNullOrEmpty(tocTitle))
        {
            var titlePara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(new Text(tocTitle))
            );
            AppendToParent(parent, titlePara);
        }

        // Field begin
        tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        // Field code
        tocPara.AppendChild(new Run(new FieldCode(instrBuilder.ToString()) { Space = SpaceProcessingModeValues.Preserve }));
        // Field separate
        tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        // Placeholder text
        tocPara.AppendChild(new Run(new Text("Update field to see table of contents") { Space = SpaceProcessingModeValues.Preserve }));
        // Field end
        tocPara.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        AppendToParent(parent, tocPara);

        // Add UpdateFieldsOnOpen setting
        var settingsPart2 = _doc.MainDocumentPart!.DocumentSettingsPart
            ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
        settingsPart2.Settings ??= new Settings();
        if (settingsPart2.Settings.GetFirstChild<UpdateFieldsOnOpen>() == null)
        {
            settingsPart2.Settings.AddChild(new UpdateFieldsOnOpen { Val = true });
        }
        settingsPart2.Settings.Save();

        // Count TOC fields in document to determine index
        var tocCount = body.Elements<Paragraph>()
            .Count(p => p.Descendants<FieldCode>().Any(fc =>
                fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)));
        var resultPath = $"/toc[{tocCount}]";
        return resultPath;
    }

    private string AddStyle(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Create a new style in the styles part
        var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
            ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        var styleId = properties.GetValueOrDefault("id", properties.GetValueOrDefault("name", "CustomStyle"));
        var styleName = properties.GetValueOrDefault("name", styleId);
        var styleType = properties.GetValueOrDefault("type", "paragraph").ToLowerInvariant() switch
        {
            "character" or "char" => StyleValues.Character,
            "table" => StyleValues.Table,
            "numbering" => StyleValues.Numbering,
            "paragraph" or "para" => StyleValues.Paragraph,
            _ => throw new ArgumentException($"Invalid style type: '{properties.GetValueOrDefault("type", "paragraph")}'. Valid values: paragraph, character, table, numbering.")
        };

        var newStyle = new Style
        {
            Type = styleType,
            StyleId = styleId,
            CustomStyle = true
        };
        newStyle.AppendChild(new StyleName { Val = styleName });

        if ((properties.TryGetValue("basedon", out var basedOn) || properties.TryGetValue("basedOn", out basedOn)) && !string.IsNullOrEmpty(basedOn))
            newStyle.AppendChild(new BasedOn { Val = basedOn });
        if (properties.TryGetValue("next", out var nextStyle))
            newStyle.AppendChild(new NextParagraphStyle { Val = nextStyle });

        // Style paragraph properties
        var stylePPr = new StyleParagraphProperties();
        bool hasPPr = false;
        if (properties.TryGetValue("alignment", out var sAlign) || properties.TryGetValue("align", out sAlign))
        {
            stylePPr.Justification = new Justification { Val = ParseJustification(sAlign) };
            hasPPr = true;
        }
        if (properties.TryGetValue("spacebefore", out var sSBefore) || properties.TryGetValue("spaceBefore", out sSBefore))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.Before = SpacingConverter.ParseWordSpacing(sSBefore).ToString();
            hasPPr = true;
        }
        if (properties.TryGetValue("spaceafter", out var sSAfter) || properties.TryGetValue("spaceAfter", out sSAfter))
        {
            var sp = stylePPr.SpacingBetweenLines ?? (stylePPr.SpacingBetweenLines = new SpacingBetweenLines());
            sp.After = SpacingConverter.ParseWordSpacing(sSAfter).ToString();
            hasPPr = true;
        }
        if (hasPPr) newStyle.AppendChild(stylePPr);

        // Style run properties
        var styleRPr = new StyleRunProperties();
        bool hasRPr = false;
        if (properties.TryGetValue("font", out var sFont))
        {
            styleRPr.RunFonts = new RunFonts { Ascii = sFont, HighAnsi = sFont, EastAsia = sFont };
            hasRPr = true;
        }
        if (properties.TryGetValue("size", out var sSize))
        {
            styleRPr.FontSize = new FontSize { Val = ((int)Math.Round(ParseFontSize(sSize) * 2, MidpointRounding.AwayFromZero)).ToString() };
            hasRPr = true;
        }
        if (properties.TryGetValue("bold", out var sBold) && IsTruthy(sBold))
        {
            styleRPr.Bold = new Bold();
            hasRPr = true;
        }
        if (properties.TryGetValue("italic", out var sItalic) && IsTruthy(sItalic))
        {
            styleRPr.Italic = new Italic();
            hasRPr = true;
        }
        if (properties.TryGetValue("color", out var sColor))
        {
            styleRPr.Color = new Color { Val = SanitizeHex(sColor) };
            hasRPr = true;
        }
        if (hasRPr) newStyle.AppendChild(styleRPr);

        stylesPart.Styles.AppendChild(newStyle);
        stylesPart.Styles.Save();

        var resultPath = $"/styles/{styleId}";
        return resultPath;
    }

    private string AddHeader(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPartH = _doc.MainDocumentPart!;
        var headerPart = mainPartH.AddNewPart<HeaderPart>();

        var hPara = new Paragraph();
        AssignParaId(hPara);
        var hPProps = new ParagraphProperties();

        if (properties.TryGetValue("alignment", out var hAlign) || properties.TryGetValue("align", out hAlign))
            hPProps.Justification = new Justification { Val = ParseJustification(hAlign) };
        hPara.AppendChild(hPProps);

        // Build shared run properties for text and field runs
        RunProperties? hSharedRProps = null;
        if (properties.ContainsKey("font") || properties.ContainsKey("size") ||
            properties.ContainsKey("bold") || properties.ContainsKey("italic") || properties.ContainsKey("color"))
        {
            hSharedRProps = new RunProperties();
            if (properties.TryGetValue("font", out var hFont))
                hSharedRProps.AppendChild(new RunFonts { Ascii = hFont, HighAnsi = hFont, EastAsia = hFont });
            if (properties.TryGetValue("size", out var hSize))
                hSharedRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(hSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
            if (properties.TryGetValue("bold", out var hBold) && IsTruthy(hBold))
                hSharedRProps.Bold = new Bold();
            if (properties.TryGetValue("italic", out var hItalic) && IsTruthy(hItalic))
                hSharedRProps.Italic = new Italic();
            if (properties.TryGetValue("color", out var hColor))
                hSharedRProps.Color = new Color { Val = SanitizeHex(hColor) };
        }

        if (properties.TryGetValue("text", out var hText))
        {
            var hRun = new Run();
            if (hSharedRProps != null) hRun.AppendChild((RunProperties)hSharedRProps.CloneNode(true));
            hRun.AppendChild(new Text(hText) { Space = SpaceProcessingModeValues.Preserve });
            hPara.AppendChild(hRun);
        }

        // Support field=page|numpages|date etc. — generates fldChar complex field
        if (properties.TryGetValue("field", out var hFieldType))
        {
            var hFieldInstr = hFieldType.ToLowerInvariant() switch
            {
                "page" or "pagenum" or "pagenumber" => " PAGE ",
                "numpages" => " NUMPAGES ",
                "date" => " DATE \\@ \"yyyy-MM-dd\" ",
                "author" => " AUTHOR ",
                "title" => " TITLE ",
                "time" => " TIME ",
                "filename" => " FILENAME ",
                _ => $" {hFieldType.ToUpperInvariant()} "
            };
            var hBeginRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
            var hInstrRun = new Run(new FieldCode(hFieldInstr) { Space = SpaceProcessingModeValues.Preserve });
            var hSepRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
            var hResultRun = new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve });
            var hEndRun = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
            if (hSharedRProps != null)
            {
                hBeginRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hInstrRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hSepRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hResultRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
                hEndRun.PrependChild((RunProperties)hSharedRProps.CloneNode(true));
            }
            hPara.AppendChild(hBeginRun);
            hPara.AppendChild(hInstrRun);
            hPara.AppendChild(hSepRun);
            hPara.AppendChild(hResultRun);
            hPara.AppendChild(hEndRun);
        }

        headerPart.Header = new Header(hPara);
        headerPart.Header.Save();

        var hBody = mainPartH.Document!.Body!;
        var hSectPr = hBody.Elements<SectionProperties>().LastOrDefault()
            ?? hBody.AppendChild(new SectionProperties());

        var headerType = HeaderFooterValues.Default;
        if (properties.TryGetValue("type", out var hTypeStr))
        {
            headerType = hTypeStr.ToLowerInvariant() switch
            {
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                "default" => HeaderFooterValues.Default,
                _ => throw new ArgumentException($"Invalid header type: '{hTypeStr}'. Valid values: default, first, even.")
            };
        }

        var headerRef = new HeaderReference
        {
            Id = mainPartH.GetIdOfPart(headerPart),
            Type = headerType
        };
        hSectPr.PrependChild(headerRef);

        if (headerType == HeaderFooterValues.First)
        {
            if (hSectPr.GetFirstChild<TitlePage>() == null)
                hSectPr.AddChild(new TitlePage(), throwOnError: false);
        }

        var hIdx = mainPartH.HeaderParts.ToList().IndexOf(headerPart);
        return $"/header[{hIdx + 1}]";
    }

    private string AddFooter(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPartF = _doc.MainDocumentPart!;
        var footerPart = mainPartF.AddNewPart<FooterPart>();

        var fPara = new Paragraph();
        AssignParaId(fPara);
        var fPProps = new ParagraphProperties();

        if (properties.TryGetValue("alignment", out var fAlign) || properties.TryGetValue("align", out fAlign))
            fPProps.Justification = new Justification { Val = ParseJustification(fAlign) };
        fPara.AppendChild(fPProps);

        // Build shared run properties for text and field runs
        RunProperties? sharedRProps = null;
        if (properties.ContainsKey("font") || properties.ContainsKey("size") ||
            properties.ContainsKey("bold") || properties.ContainsKey("italic") || properties.ContainsKey("color"))
        {
            sharedRProps = new RunProperties();
            if (properties.TryGetValue("font", out var fFont))
                sharedRProps.AppendChild(new RunFonts { Ascii = fFont, HighAnsi = fFont, EastAsia = fFont });
            if (properties.TryGetValue("size", out var fSize))
                sharedRProps.AppendChild(new FontSize { Val = ((int)Math.Round(ParseFontSize(fSize) * 2, MidpointRounding.AwayFromZero)).ToString() });
            if (properties.TryGetValue("bold", out var fBold) && IsTruthy(fBold))
                sharedRProps.Bold = new Bold();
            if (properties.TryGetValue("italic", out var fItalic) && IsTruthy(fItalic))
                sharedRProps.Italic = new Italic();
            if (properties.TryGetValue("color", out var fColor))
                sharedRProps.Color = new Color { Val = SanitizeHex(fColor) };
        }

        if (properties.TryGetValue("text", out var fText))
        {
            var fRun = new Run();
            if (sharedRProps != null) fRun.AppendChild((RunProperties)sharedRProps.CloneNode(true));
            fRun.AppendChild(new Text(fText) { Space = SpaceProcessingModeValues.Preserve });
            fPara.AppendChild(fRun);
        }

        // Support field=page|numpages|date etc. — generates fldChar complex field
        if (properties.TryGetValue("field", out var fieldType))
        {
            var fieldInstr = fieldType.ToLowerInvariant() switch
            {
                "page" or "pagenum" or "pagenumber" => " PAGE ",
                "numpages" => " NUMPAGES ",
                "date" => " DATE \\@ \"yyyy-MM-dd\" ",
                "author" => " AUTHOR ",
                "title" => " TITLE ",
                "time" => " TIME ",
                "filename" => " FILENAME ",
                _ => $" {fieldType.ToUpperInvariant()} "
            };
            var beginRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Begin });
            var instrRun = new Run(new FieldCode(fieldInstr) { Space = SpaceProcessingModeValues.Preserve });
            var sepRun = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
            var resultRun = new Run(new Text("1") { Space = SpaceProcessingModeValues.Preserve });
            var endRun = new Run(new FieldChar { FieldCharType = FieldCharValues.End });
            if (sharedRProps != null)
            {
                beginRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                instrRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                sepRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                resultRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
                endRun.PrependChild((RunProperties)sharedRProps.CloneNode(true));
            }
            fPara.AppendChild(beginRun);
            fPara.AppendChild(instrRun);
            fPara.AppendChild(sepRun);
            fPara.AppendChild(resultRun);
            fPara.AppendChild(endRun);
        }

        footerPart.Footer = new Footer(fPara);
        footerPart.Footer.Save();

        var fBody = mainPartF.Document!.Body!;
        var fSectPr = fBody.Elements<SectionProperties>().LastOrDefault()
            ?? fBody.AppendChild(new SectionProperties());

        var footerType = HeaderFooterValues.Default;
        if (properties.TryGetValue("type", out var fTypeStr))
        {
            footerType = fTypeStr.ToLowerInvariant() switch
            {
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                "default" => HeaderFooterValues.Default,
                _ => throw new ArgumentException($"Invalid footer type: '{fTypeStr}'. Valid values: default, first, even.")
            };
        }

        var footerRef = new FooterReference
        {
            Id = mainPartF.GetIdOfPart(footerPart),
            Type = footerType
        };
        // Insert footerReference after the last headerReference to maintain schema order
        var lastHeaderRef = fSectPr.Elements<HeaderReference>().LastOrDefault();
        if (lastHeaderRef != null)
            fSectPr.InsertAfter(footerRef, lastHeaderRef);
        else
            fSectPr.PrependChild(footerRef);

        if (footerType == HeaderFooterValues.First)
        {
            if (fSectPr.GetFirstChild<TitlePage>() == null)
                fSectPr.AddChild(new TitlePage(), throwOnError: false);
        }

        var fIdx = mainPartF.FooterParts.ToList().IndexOf(footerPart);
        return $"/footer[{fIdx + 1}]";
    }
}
