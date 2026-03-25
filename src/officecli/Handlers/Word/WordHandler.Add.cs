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
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        OpenXmlElement parent;
        if (parentPath is "/" or "" or "/body")
        {
            parent = body;
            parentPath = "/body"; // Normalize so result paths are usable (e.g. /body/p[1] not //p[1])
        }
        else if (parentPath == "/styles")
        {
            // Ensure styles part exists for style operations
            var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
                ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles ??= new Styles();
            parent = stylesPart.Styles;
        }
        else
        {
            var parts = ParsePath(parentPath);
            parent = NavigateToElement(parts, out var ctx)
                ?? throw new ArgumentException($"Path not found: {parentPath}" + (ctx != null ? $". {ctx}" : ""));
        }

        var resultPath = type.ToLowerInvariant() switch
        {
            "paragraph" or "p" => AddParagraph(parent, parentPath, index, properties),
            "equation" or "formula" or "math" => AddEquation(parent, parentPath, index, properties),
            "run" or "r" => AddRun(parent, parentPath, index, properties),
            "table" or "tbl" => AddTable(parent, parentPath, index, properties),
            "row" or "tr" => AddRow(parent, parentPath, index, properties),
            "cell" or "tc" => AddCell(parent, parentPath, index, properties),
            "chart" => AddChart(parent, parentPath, index, properties),
            "picture" or "image" or "img" => AddPicture(parent, parentPath, index, properties),
            "comment" => AddComment(parent, parentPath, index, properties),
            "bookmark" => AddBookmark(parent, parentPath, index, properties),
            "hyperlink" or "link" => AddHyperlink(parent, parentPath, index, properties),
            "section" or "sectionbreak" => AddSection(parent, parentPath, index, properties),
            "footnote" => AddFootnote(parent, parentPath, index, properties),
            "endnote" => AddEndnote(parent, parentPath, index, properties),
            "toc" or "tableofcontents" => AddToc(parent, parentPath, index, properties),
            "style" => AddStyle(parent, parentPath, index, properties),
            "header" => AddHeader(parent, parentPath, index, properties),
            "footer" => AddFooter(parent, parentPath, index, properties),
            "field" or "pagenum" or "pagenumber" or "numpages" or "date" => AddField(parent, parentPath, index, properties, type),
            "pagebreak" or "columnbreak" or "break" => AddBreak(parent, parentPath, index, properties, type),
            "sdt" or "contentcontrol" => AddSdt(parent, parentPath, index, properties),
            "watermark" => AddWatermark(parent, parentPath, index, properties),
            _ => AddDefault(parent, parentPath, index, properties, type),
        };

        _doc.MainDocumentPart?.Document?.Save();
        return resultPath;
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var mainPart = _doc.MainDocumentPart!;

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                var chartPart = mainPart.AddNewPart<ChartPart>();
                var relId = mainPart.GetIdOfPart(chartPart);
                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new C.ChartSpace(
                    new C.Chart(new C.PlotArea(new C.Layout()))
                );
                chartPart.ChartSpace.Save();
                var chartIdx = mainPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/chart[{chartIdx + 1}]");

            case "header":
                var headerPart = mainPart.AddNewPart<HeaderPart>();
                var hRelId = mainPart.GetIdOfPart(headerPart);
                headerPart.Header = new Header(new Paragraph());
                headerPart.Header.Save();
                var hIdx = mainPart.HeaderParts.ToList().IndexOf(headerPart);
                return (hRelId, $"/header[{hIdx + 1}]");

            case "footer":
                var footerPart = mainPart.AddNewPart<FooterPart>();
                var fRelId = mainPart.GetIdOfPart(footerPart);
                footerPart.Footer = new Footer(new Paragraph());
                footerPart.Footer.Save();
                var fIdx = mainPart.FooterParts.ToList().IndexOf(footerPart);
                return (fRelId, $"/footer[{fIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, header, footer");
        }
    }


    private void SetDocumentProperties(Dictionary<string, string> properties)
    {
        var doc = _doc.MainDocumentPart?.Document
            ?? throw new InvalidOperationException("Document not found");

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "pagebackground" or "background":
                    doc.DocumentBackground = new DocumentBackground { Color = value };
                    // Enable background display in settings
                    var settingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings ??= new Settings();
                    if (settingsPart.Settings.GetFirstChild<DisplayBackgroundShape>() == null)
                        settingsPart.Settings.AppendChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "defaultfont":
                    var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart;
                    if (stylesPart?.Styles != null)
                    {
                        var defaultRunProps = stylesPart.Styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                        if (defaultRunProps != null)
                        {
                            var fonts = defaultRunProps.GetFirstChild<RunFonts>()
                                ?? defaultRunProps.AppendChild(new RunFonts());
                            fonts.Ascii = value;
                            fonts.HighAnsi = value;
                            fonts.EastAsia = value;
                            stylesPart.Styles.Save();
                        }
                    }
                    break;

                case "pagewidth":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = ParseTwips(value);
                    break;
                case "pageheight":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = ParseTwips(value);
                    break;
                case "margintop":
                    EnsurePageMargin().Top = ParseHelpers.SafeParseInt(value, "margintop");
                    break;
                case "marginbottom":
                    EnsurePageMargin().Bottom = ParseHelpers.SafeParseInt(value, "marginbottom");
                    break;
                case "marginleft":
                    EnsurePageMargin().Left = ParseHelpers.SafeParseUint(value, "marginleft");
                    break;
                case "marginright":
                    EnsurePageMargin().Right = ParseHelpers.SafeParseUint(value, "marginright");
                    break;

                // Core document properties
                case "title":
                    _doc.PackageProperties.Title = value;
                    break;
                case "author" or "creator":
                    _doc.PackageProperties.Creator = value;
                    break;
                case "subject":
                    _doc.PackageProperties.Subject = value;
                    break;
                case "keywords":
                    _doc.PackageProperties.Keywords = value;
                    break;
                case "description":
                    _doc.PackageProperties.Description = value;
                    break;
                case "category":
                    _doc.PackageProperties.Category = value;
                    break;
                case "lastmodifiedby":
                    _doc.PackageProperties.LastModifiedBy = value;
                    break;
                case "revision":
                    _doc.PackageProperties.Revision = value;
                    break;

                case "acceptallchanges":
                    if (IsTruthy(value))
                        AcceptAllChanges();
                    break;
                case "rejectallchanges":
                    if (IsTruthy(value))
                        RejectAllChanges();
                    break;
            }
        }
    }

    private SectionProperties EnsureSectionProperties()
    {
        var body = _doc.MainDocumentPart!.Document!.Body!;
        var sectPr = body.GetFirstChild<SectionProperties>();
        if (sectPr == null)
        {
            sectPr = new SectionProperties();
            body.AppendChild(sectPr);
        }
        if (sectPr.GetFirstChild<PageSize>() == null)
            sectPr.AppendChild(new PageSize { Width = 11906, Height = 16838 }); // A4 default
        return sectPr;
    }

    private PageMargin EnsurePageMargin()
    {
        var sectPr = EnsureSectionProperties();
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin == null)
        {
            margin = new PageMargin { Top = 1440, Bottom = 1440, Left = 1800, Right = 1800 };
            sectPr.AppendChild(margin);
        }
        return margin;
    }
}
