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
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        // CONSISTENCY(prop-key-case): property keys are case-insensitive
        // ("SRC"/"src"/"Src" all resolve the same). Normalize once at the
        // dispatch entry so every AddXxx helper can rely on TryGetValue("src").
        properties = properties == null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : properties.Comparer == StringComparer.OrdinalIgnoreCase
                ? properties
                : new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);

        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        OpenXmlElement parent;
        if (parentPath is "/" or "" or "/body")
        {
            parent = body;
            parentPath = "/body";
        }
        else if (parentPath == "/styles")
        {
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

        // Resolve --after/--before to index (handles find: prefix for text-based anchoring)
        var index = ResolveAnchorPosition(parent, parentPath, position);

        // Handle find: prefix — text-based anchoring
        if (index == FindAnchorIndex && position != null)
        {
            var anchorValue = (position.After ?? position.Before)!;
            var findValue = anchorValue["find:".Length..]; // strip "find:" prefix
            var isAfter = position.After != null;
            return AddAtFindPosition(parent, parentPath, type, findValue, isAfter, null, properties);
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
            "ole" or "oleobject" or "object" or "embed" => AddOle(parent, parentPath, index, properties),
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
            "field" or "pagenum" or "pagenumber" or "page" or "numpages" or "date" or "author" => AddField(parent, parentPath, index, properties, type),
            "pagebreak" or "columnbreak" or "break" => AddBreak(parent, parentPath, index, properties, type),
            "sdt" or "contentcontrol" => AddSdt(parent, parentPath, index, properties),
            "watermark" => AddWatermark(parent, parentPath, index, properties),
            "formfield" => AddFormField(parent, parentPath, index, properties),
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


    private void SetDocumentProperties(Dictionary<string, string> properties, List<string>? unsupported = null)
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
                        settingsPart.Settings.AddChild(new DisplayBackgroundShape());
                    settingsPart.Settings.Save();
                    break;

                case "defaultfont":
                    // Delegate to TrySetDocDefaults which uses EnsureRunPropsDefault()
                    // to create the DocDefaults chain when absent (e.g. blank documents).
                    TrySetDocDefaults("docdefaults.font", value);
                    break;
                case "defaultfontsize":
                    TrySetDocDefaults("docdefaults.fontsize", value);
                    break;

                case "pagewidth":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Width = ParseTwips(value);
                    break;
                case "pageheight":
                    EnsureSectionProperties().GetFirstChild<PageSize>()!.Height = ParseTwips(value);
                    break;
                case "margintop":
                    EnsurePageMargin().Top = (int)ParseTwips(value);
                    break;
                case "marginbottom":
                    EnsurePageMargin().Bottom = (int)ParseTwips(value);
                    break;
                case "marginleft":
                    EnsurePageMargin().Left = ParseTwips(value);
                    break;
                case "marginright":
                    EnsurePageMargin().Right = ParseTwips(value);
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

                case "protection":
                {
                    var protSettingsPart = _doc.MainDocumentPart!.DocumentSettingsPart
                        ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                    protSettingsPart.Settings ??= new Settings();

                    // Remove existing protection
                    var existing = protSettingsPart.Settings.GetFirstChild<DocumentProtection>();
                    existing?.Remove();

                    if (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        var editValue = value.ToLowerInvariant() switch
                        {
                            "forms" => DocumentProtectionValues.Forms,
                            "readonly" => DocumentProtectionValues.ReadOnly,
                            "comments" => DocumentProtectionValues.Comments,
                            "trackedchanges" => DocumentProtectionValues.TrackedChanges,
                            _ => DocumentProtectionValues.Forms
                        };
                        var prot = new DocumentProtection
                        {
                            Edit = new EnumValue<DocumentProtectionValues>(editValue),
                            Enforcement = new OnOffValue(true)
                        };
                        protSettingsPart.Settings.AppendChild(prot);
                    }

                    protSettingsPart.Settings.Save();
                    break;
                }

                case "acceptallchanges" or "accept-changes" or "acceptchanges":
                    if (value.Equals("all", StringComparison.OrdinalIgnoreCase) || IsTruthy(value))
                        AcceptAllChanges();
                    break;
                case "rejectallchanges" or "reject-changes" or "rejectchanges":
                    if (value.Equals("all", StringComparison.OrdinalIgnoreCase) || IsTruthy(value))
                        RejectAllChanges();
                    break;

                default:
                    // Try document settings, section layout, compatibility, and docDefaults
                    var lowerKey = key.ToLowerInvariant();
                    if (!TrySetDocSetting(lowerKey, value)
                        && !TrySetSectionLayout(lowerKey, value)
                        && !TrySetCompatibility(lowerKey, value)
                        && !TrySetDocDefaults(lowerKey, value)
                        && !Core.ThemeHandler.TrySetTheme(_doc.MainDocumentPart?.ThemePart, lowerKey, value)
                        && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                            Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                        unsupported?.Add(key);
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
        {
            var pgSz = new PageSize { Width = 11906, Height = 16838 }; // A4 default
            // Schema order: pgSz must come before pgMar, cols, and docGrid
            var firstNonRef = sectPr.ChildElements.FirstOrDefault(c =>
                c is not HeaderReference && c is not FooterReference && c is not SectionType);
            if (firstNonRef != null)
                firstNonRef.InsertBeforeSelf(pgSz);
            else
                sectPr.AppendChild(pgSz);
        }
        return sectPr;
    }

    private PageMargin EnsurePageMargin()
    {
        var sectPr = EnsureSectionProperties();
        var margin = sectPr.GetFirstChild<PageMargin>();
        if (margin == null)
        {
            margin = new PageMargin { Top = 1440, Bottom = 1440, Left = 1800, Right = 1800 };
            // Insert after PageSize to maintain CT_SectPr schema order: pgSz → pgMar → ...
            var pgSz = sectPr.GetFirstChild<PageSize>();
            if (pgSz != null)
                pgSz.InsertAfterSelf(margin);
            else
                sectPr.AddChild(margin, throwOnError: false);
        }
        return margin;
    }
}
