// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private static string HeaderFooterTypeName(HeaderFooterValues v)
    {
        if (v == HeaderFooterValues.First) return "first";
        if (v == HeaderFooterValues.Even) return "even";
        return "default";
    }

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
            bodySectPr.InsertBefore(new PageSize { Width = WordPageDefaults.A4WidthTwips, Height = WordPageDefaults.A4HeightTwips },
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
            Width = srcPageSize?.Width ?? WordPageDefaults.A4WidthTwips,
            Height = srcPageSize?.Height ?? WordPageDefaults.A4HeightTwips,
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
        if (properties.TryGetValue("columns.space", out var colSpace)
            || properties.TryGetValue("columnSpace", out colSpace))
        {
            var cols = sectPr.GetFirstChild<Columns>() ?? sectPr.AppendChild(new Columns());
            cols.Space = ParseTwips(colSpace).ToString();
        }

        // Per-section margin overrides — mutate the PageMargin child of the
        // new sectPr (not the body sectPr). Margins use Int32Value for Top/
        // Bottom and UInt32Value for Left/Right to match the schema.
        var pm = sectPr.GetFirstChild<PageMargin>() ?? sectPr.AppendChild(new PageMargin());
        if (properties.TryGetValue("marginTop", out var mTop) || properties.TryGetValue("margintop", out mTop))
            pm.Top = (int)ParseTwips(mTop);
        if (properties.TryGetValue("marginBottom", out var mBot) || properties.TryGetValue("marginbottom", out mBot))
            pm.Bottom = (int)ParseTwips(mBot);
        if (properties.TryGetValue("marginLeft", out var mLeft) || properties.TryGetValue("marginleft", out mLeft))
            pm.Left = ParseTwips(mLeft);
        if (properties.TryGetValue("marginRight", out var mRight) || properties.TryGetValue("marginright", out mRight))
            pm.Right = ParseTwips(mRight);

        // Line numbering — mirrors Set parser (WordHandler.Set.cs ~L608).
        if (properties.TryGetValue("lineNumbers", out var lnVal) || properties.TryGetValue("linenumbers", out lnVal))
        {
            var restart = lnVal.ToLowerInvariant() switch
            {
                "continuous" => LineNumberRestartValues.Continuous,
                "restartpage" or "page" => LineNumberRestartValues.NewPage,
                "restartsection" or "section" => LineNumberRestartValues.NewSection,
                _ => throw new ArgumentException(
                    $"Invalid lineNumbers value: '{lnVal}'. Valid values: continuous, restartPage, restartSection.")
            };
            var lnType = new LineNumberType { Restart = restart };
            if (properties.TryGetValue("lineNumberCountBy", out var lnBy)
                || properties.TryGetValue("linenumbercountby", out lnBy))
            {
                var by = int.Parse(lnBy);
                if (by > 1) lnType.CountBy = (short)by;
            }
            sectPr.AppendChild(lnType);
        }

        // Dotted-key fallback for sectPr-level attrs not modeled by the
        // hand-rolled blocks above (single-attr forms like docGrid.* or
        // future schema additions). CONSISTENCY(add-set-symmetry).
        // Skip the dotted curated keys that AddSection already consumes
        // explicitly to avoid double application.
        var sectionAlreadyConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "columns.count", "columns.space",
        };
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            if (sectionAlreadyConsumed.Contains(key)) continue;
            if (Core.TypedAttributeFallback.TrySet(sectPr, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        sectPProps.AppendChild(sectPr);
        sectPara.AppendChild(sectPProps);
        InsertAtIndexOrAppend(parent, sectPara, index);

        // Return the new section's document-order position (1-based) so the
        // path matches the NavigateToElement /section[N] resolver, which
        // walks body paragraphs with SectionProperties in document order.
        // Using the total count would break --before/--after (which insert
        // mid-document): the new section may not be the last one.
        var sectParas = body.Elements<Paragraph>()
            .Where(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
            .ToList();
        var secDocOrderIdx = sectParas.FindIndex(p => ReferenceEquals(p, sectPara));
        var resultPath = $"/section[{(secDocOrderIdx >= 0 ? secDocOrderIdx + 1 : sectParas.Count)}]";
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

        // Insert reference in document body at the requested index, keeping
        // pPr as first child (InsertIntoParagraph clamps forward past pPr).
        var fnRefRun = new Run(
            new RunProperties(new RunStyle { Val = "FootnoteReference" }),
            new FootnoteReference { Id = fnId }
        );
        InsertIntoParagraph(fnPara, fnRefRun, index);

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

        // Insert reference in document body at the requested index, keeping
        // pPr as first child (InsertIntoParagraph clamps forward past pPr).
        var enRefRun = new Run(
            new RunProperties(new RunStyle { Val = "EndnoteReference" }),
            new EndnoteReference { Id = enId }
        );
        InsertIntoParagraph(enPara, enRefRun, index);

        var resultPath = $"/endnote[{enId}]";
        return resultPath;
    }

    private string AddToc(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var body = _doc.MainDocumentPart?.Document?.Body
            ?? throw new InvalidOperationException("Document body not found");

        // TOC fields reference body-level heading styles; adding them in a
        // header/footer part is not meaningful and would yield an unnavigable
        // /toc[0] return path (body TOC count is 0). Reject early with a
        // clean error.
        if (parent is Header || parent is Footer
            || parent.Ancestors<Header>().Any() || parent.Ancestors<Footer>().Any())
        {
            throw new ArgumentException(
                "add --type toc is not supported inside a header or footer part. " +
                "TOC field codes reference body-level headings; insert into /body instead.");
        }

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

        // Insert TOC paragraph at the requested position first, then — if a
        // title was requested — insert the title paragraph immediately before
        // it so the title precedes the TOC field in reading order. Previously
        // the title was appended to the parent regardless of --index, ending
        // up after the TOC.
        InsertAtIndexOrAppend(parent, tocPara, index);
        if (!string.IsNullOrEmpty(tocTitle))
        {
            var titlePara = new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "TOCHeading" }),
                new Run(new Text(tocTitle))
            );
            tocPara.InsertBeforeSelf(titlePara);
        }

        // Add UpdateFieldsOnOpen setting
        var settingsPart2 = _doc.MainDocumentPart!.DocumentSettingsPart
            ?? _doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
        settingsPart2.Settings ??= new Settings();
        if (settingsPart2.Settings.GetFirstChild<UpdateFieldsOnOpen>() == null)
        {
            settingsPart2.Settings.AddChild(new UpdateFieldsOnOpen { Val = true });
        }
        settingsPart2.Settings.Save();

        // Determine TOC index in document order (not total count)
        var tocParas = body.Elements<Paragraph>()
            .Where(p => p.Descendants<FieldCode>().Any(fc =>
                fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)))
            .ToList();
        var tocIdx = tocParas.FindIndex(p => ReferenceEquals(p, tocPara));
        var resultPath = $"/toc[{(tocIdx >= 0 ? tocIdx + 1 : tocParas.Count)}]";
        return resultPath;
    }

    private string AddStyle(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // Create a new style in the styles part
        var stylesPart = _doc.MainDocumentPart!.StyleDefinitionsPart
            ?? _doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles ??= new Styles();

        // CONSISTENCY(style-dual-key): Get exposes the canonical readback
        // keys `styleId` and `styleName` on every paragraph (Round 2). Add
        // must accept the same alias trio (id / styleId / name / styleName)
        // or the readback writes back as `CustomStyle` — exactly the
        // silent-ignore alias trap that 19b3dd5b banned.
        var explicitId = properties.ContainsKey("id") || properties.ContainsKey("styleId") || properties.ContainsKey("styleid");
        var styleId = properties.GetValueOrDefault("id")
                   ?? properties.GetValueOrDefault("styleId")
                   ?? properties.GetValueOrDefault("styleid")
                   ?? properties.GetValueOrDefault("name")
                   ?? properties.GetValueOrDefault("styleName")
                   ?? properties.GetValueOrDefault("stylename")
                   ?? "CustomStyle";
        var styleName = properties.GetValueOrDefault("name")
                     ?? properties.GetValueOrDefault("styleName")
                     ?? properties.GetValueOrDefault("stylename")
                     ?? styleId;
        var styleType = properties.GetValueOrDefault("type", "paragraph").ToLowerInvariant() switch
        {
            "character" or "char" => StyleValues.Character,
            "table" => StyleValues.Table,
            "numbering" => StyleValues.Numbering,
            "paragraph" or "para" => StyleValues.Paragraph,
            _ => throw new ArgumentException($"Invalid style type: '{properties.GetValueOrDefault("type", "paragraph")}'. Valid values: paragraph, character, table, numbering.")
        };

        // Enforce unique styleId — schema requires unique w:styleId per styles.xml.
        // If the caller specified --prop id explicitly, reject; otherwise auto-suffix
        // to keep the call idempotent-ish for scripts that only pass --prop name.
        bool IdTaken(string candidate) => stylesPart.Styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleId?.Value, candidate, StringComparison.Ordinal));
        if (IdTaken(styleId))
        {
            if (explicitId)
                throw new ArgumentException(
                    $"Style '{styleId}' already exists. Pick a unique --prop id or --prop name.");
            var baseId = styleId;
            int suffix = 2;
            while (IdTaken(styleId)) styleId = $"{baseId}{suffix++}";
        }

        // OOXML requires w:name to be unique across styles.xml, same as w:styleId.
        // Reject duplicate display names — silently auto-suffixing the id while
        // leaving name unchanged produced two styles with identical UI labels
        // that users could not tell apart (BUG-R17-02).
        bool NameTaken(string candidate) => stylesPart.Styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleName?.Val?.Value, candidate, StringComparison.Ordinal));
        if (NameTaken(styleName))
            throw new ArgumentException(
                $"Style with name '{styleName}' already exists. Pick a unique --prop name.");

        // Built-in styles must not have customStyle=true, or Word won't recognize them
        // (e.g. TOC won't find Heading1 if it's marked as custom)
        var builtInIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Normal", "Heading1", "Heading2", "Heading3", "Heading4", "Heading5",
            "Heading6", "Heading7", "Heading8", "Heading9", "Title", "Subtitle",
            "Quote", "IntenseQuote", "ListParagraph", "NoSpacing", "TOCHeading"
        };
        var isBuiltIn = builtInIds.Contains(styleId);

        var newStyle = new Style
        {
            Type = styleType,
            StyleId = styleId,
        };
        if (!isBuiltIn)
            newStyle.CustomStyle = true;
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
        // Per-script font split. Each w:rFonts attr is independent — Word falls
        // back through the style chain / docDefaults for any unset attr, so we
        // only write what the caller passed and leave the rest alone. Dotted
        // keys layer on top of the bare `font=` shortcut: `font=Times,
        // font.eastAsia=SimSun` produces ascii/hAnsi=Times, eastAsia=SimSun.
        bool TrySetRFontsAttr(string key, Action<RunFonts, string> assign)
        {
            if (!properties.TryGetValue(key, out var v) || string.IsNullOrEmpty(v)) return false;
            styleRPr.RunFonts ??= new RunFonts();
            assign(styleRPr.RunFonts, v);
            hasRPr = true;
            return true;
        }
        TrySetRFontsAttr("font.ascii",    (rf, v) => rf.Ascii = v);
        TrySetRFontsAttr("font.hAnsi",    (rf, v) => rf.HighAnsi = v);
        TrySetRFontsAttr("font.eastAsia", (rf, v) => rf.EastAsia = v);
        TrySetRFontsAttr("font.cs",       (rf, v) => rf.ComplexScript = v);
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

        // Numbering linkage on the style itself (numPr inside StyleParagraphProperties).
        // Lets paragraphs inherit list editing without setting numPr on each paragraph,
        // which is the canonical pattern used by Heading1..9 in real templates.
        // Mirrors WordHandler.Set.cs paragraph-level numId/ilvl handling.
        bool hasStyleNumPr = (properties.TryGetValue("numId", out var sNumIdStr) || properties.TryGetValue("numid", out sNumIdStr))
                          || (properties.TryGetValue("ilvl", out _) || properties.TryGetValue("numLevel", out _) || properties.TryGetValue("numlevel", out _));
        if (hasStyleNumPr)
        {
            var pPrForNum = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
            var numPr = pPrForNum.NumberingProperties ?? (pPrForNum.NumberingProperties = new NumberingProperties());
            if (!string.IsNullOrEmpty(sNumIdStr))
            {
                var nid = ParseHelpers.SafeParseInt(sNumIdStr, "numId");
                if (nid < 0) throw new ArgumentException($"numId must be >= 0 (got {nid}).");
                // CONSISTENCY(numId-ref-check): mirror paragraph-level validation
                // in WordHandler.Add.Text.cs. Positive numIds must reference an
                // existing w:num so styles don't silently introduce dangling refs.
                if (nid > 0)
                {
                    var numberingPart = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                    var numExists = numberingPart?.Elements<NumberingInstance>()
                        .Any(n => n.NumberID?.Value == nid) ?? false;
                    if (!numExists)
                        throw new ArgumentException(
                            $"numId={nid} not found in /numbering. " +
                            "Create the num first (add /numbering --type num), or use numId=0 to remove numbering.");
                }
                numPr.NumberingId = new NumberingId { Val = nid };
            }
            string? ilvlRaw = null;
            if (properties.TryGetValue("ilvl", out var iRaw)
                || properties.TryGetValue("numLevel", out iRaw)
                || properties.TryGetValue("numlevel", out iRaw))
                ilvlRaw = iRaw;
            if (!string.IsNullOrEmpty(ilvlRaw))
            {
                var ilvl = ParseHelpers.SafeParseInt(ilvlRaw, "ilvl");
                if (ilvl < 0 || ilvl > 8)
                    throw new ArgumentException($"ilvl must be in range 0..8 (got {ilvl}).");
                numPr.NumberingLevelReference = new NumberingLevelReference { Val = ilvl };
            }
        }

        // CONSISTENCY(add-set-symmetry): mirror SetStylePath's ApplyRunFormatting
        // + generic OOXML fallback so `add` accepts the same prop surface as
        // `set` for any single-Val style property. Without this sweep, props
        // like underline/strike/highlight/contextualSpacing/kinsoku/snapToGrid
        // would be silently dropped on add (schema preflight lets them
        // through; AddStyle's TryGetValue list only covers ~13 keys).
        var addStyleConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // CONSISTENCY(style-dual-key): styleId / styleName are the
            // canonical readback keys Get surfaces (Round 2). The id/name
            // alias chain consumed them above; record both spellings here
            // so the per-key 'silent drop' sweep doesn't flag them as
            // unsupported even though they were honored.
            "id", "styleId", "styleid",
            "name", "styleName", "stylename",
            "type", "basedon", "basedOn", "next",
            "alignment", "align", "spacebefore", "spaceBefore",
            "spaceafter", "spaceAfter", "font", "size", "bold", "italic", "color",
            "font.ascii", "font.hAnsi", "font.eastAsia", "font.cs",
            "numId", "numid", "ilvl", "numLevel", "numlevel",
        };
        foreach (var (key, value) in properties)
        {
            if (addStyleConsumed.Contains(key)) continue;

            // 1) Run-formatting helper (covers underline/strike/highlight/caps/
            //    smallCaps/dstrike/vanish/shadow/emboss/imprint/noProof/rtl/
            //    superscript/subscript/charSpacing/shading/...).
            var rPrProbeAdd = new StyleRunProperties();
            if (ApplyRunFormatting(rPrProbeAdd, key, value))
            {
                ApplyRunFormatting(
                    newStyle.StyleRunProperties ?? newStyle.AppendChild(new StyleRunProperties()),
                    key, value);
                continue;
            }

            // 1b) Generic dotted "element.attr=value" fallback (e.g.
            //     ind.firstLine=240, shd.fill=FF0000, font.eastAsia=…).
            //     SDK-validated round-trip rejects unknown element/attr
            //     combinations. Runs ahead of the single-val fallback so
            //     dotted keys never accidentally get coerced into a
            //     <w:foo w:val="bar.baz"/> element.
            if (key.Contains('.'))
            {
                var pPrAttrProbe = new StyleParagraphProperties();
                if (Core.TypedAttributeFallback.TrySet(pPrAttrProbe, key, value))
                {
                    var pPrReal = newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle);
                    Core.TypedAttributeFallback.TrySet(pPrReal, key, value);
                    continue;
                }
                var rPrAttrProbe = new StyleRunProperties();
                if (Core.TypedAttributeFallback.TrySet(rPrAttrProbe, key, value))
                {
                    var rPrReal = newStyle.StyleRunProperties ?? newStyle.AppendChild(new StyleRunProperties());
                    Core.TypedAttributeFallback.TrySet(rPrReal, key, value);
                    continue;
                }
            }

            // 2) Generic OOXML single-Val fallback — pPr first, rPr second,
            //    matching SetStylePath's default branch. Detached probes
            //    avoid leaking empty containers on misses.
            var pPrProbeAdd = new StyleParagraphProperties();
            if (Core.GenericXmlQuery.TryCreateTypedChild(pPrProbeAdd, key, value))
            {
                Core.GenericXmlQuery.TryCreateTypedChild(
                    newStyle.StyleParagraphProperties ?? EnsureStyleParagraphProperties(newStyle),
                    key, value);
                continue;
            }
            var rPrProbeAdd2 = new StyleRunProperties();
            if (Core.GenericXmlQuery.TryCreateTypedChild(rPrProbeAdd2, key, value))
            {
                Core.GenericXmlQuery.TryCreateTypedChild(
                    newStyle.StyleRunProperties ?? newStyle.AppendChild(new StyleRunProperties()),
                    key, value);
                continue;
            }
            // Anything still unconsumed is a genuine silent drop — composites
            // (font.eastAsia, ind.firstLine, tabs, numId, ...) that the
            // curated AddStyle does not yet model. Record so the CLI layer
            // can surface a WARNING with targeted hints instead of a silent
            // "Added" lie. See StyleUnsupportedHints for the hint catalog.
            LastAddUnsupportedProps.Add(key);
        }

        stylesPart.Styles.AppendChild(newStyle);
        stylesPart.Styles.Save();

        var resultPath = $"/styles/{styleId}";
        return resultPath;
    }

    /// <summary>
    /// Add a numbering instance (&lt;w:num&gt;) under /numbering. A num is a thin
    /// pointer that references an existing &lt;w:abstractNum&gt; via abstractNumId.
    ///
    /// Mode B (current): requires --prop abstractNumId=N pointing at an existing
    /// abstractNum. Other modes (auto-create abstractNum, lvlOverride) follow.
    /// </summary>
    private string AddNum(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart
            ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        var numbering = numberingPart.Numbering;

        // Three modes:
        //   B/C: --prop abstractNumId=N (reuse existing template; optionally with start overrides)
        //   A:   --prop format=... (no abstractNumId; auto-create a matching abstractNum)
        //   neither: throw with guidance
        bool hasAbsId = properties.TryGetValue("abstractNumId", out var absIdStr) && !string.IsNullOrEmpty(absIdStr);
        bool hasFormat = properties.ContainsKey("format")
                       || properties.ContainsKey("text")
                       || properties.ContainsKey("indent")
                       || properties.ContainsKey("type")
                       || properties.ContainsKey("name")
                       || properties.ContainsKey("styleLink")
                       || properties.ContainsKey("numStyleLink")
                       || properties.Keys.Any(k =>
                            k.StartsWith("level", StringComparison.OrdinalIgnoreCase)
                            && k.Length > 5 && char.IsDigit(k[5]));
        if (hasAbsId && hasFormat)
            throw new ArgumentException(
                "--prop abstractNumId conflicts with --prop format/text/indent/type. " +
                "Either reuse an existing template (abstractNumId) or define a new one (format/text/indent/type), not both.");
        if (!hasAbsId && !hasFormat)
            throw new ArgumentException(
                "--type num requires either --prop abstractNumId=N (reuse existing template) " +
                "or --prop format=decimal|bullet|... (auto-create a matching abstractNum).");

        int abstractNumId;
        if (hasAbsId)
        {
            abstractNumId = ParseHelpers.SafeParseInt(absIdStr!, "abstractNumId");
            // Reject pointers that would dangle — Word silently drops numbering
            // when numId resolves to a missing abstractNum, which is a confusing
            // failure mode to debug. Catch it at write time.
            var abstractExists = numbering.Elements<AbstractNum>()
                .Any(a => a.AbstractNumberId?.Value == abstractNumId);
            if (!abstractExists)
                throw new ArgumentException(
                    $"abstractNumId={abstractNumId} not found in /numbering. " +
                    "Create the abstractNum first, or pick an existing one via 'officecli query <file> abstractNum'.");
        }
        else
        {
            abstractNumId = numbering.Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
            BuildAbstractNumElement(numbering, abstractNumId, properties);
        }

        // numId assignment: explicit collides → throw; otherwise max+1.
        // Mirrors AddStyle's IdTaken pattern, but numId is int (not string)
        // so there's no "auto-suffix" — just take next available.
        int numId;
        var explicitId = properties.ContainsKey("id");
        if (explicitId)
        {
            numId = ParseHelpers.SafeParseInt(properties["id"], "id");
            if (numId < 1)
                throw new ArgumentException($"numId must be >= 1 (got {numId}). numId=0 is reserved as 'no numbering'.");
            if (numbering.Elements<NumberingInstance>().Any(n => n.NumberID?.Value == numId))
                throw new ArgumentException(
                    $"numId {numId} already exists. Pick a unique --prop id, or omit --prop id for auto-assignment.");
        }
        else
        {
            numId = numbering.Elements<NumberingInstance>()
                .Select(n => n.NumberID?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
        }

        // Schema requires AbstractNum elements before NumberingInstance elements.
        // Append the new num at the end of the existing NumberingInstance run.
        var newNum = new NumberingInstance { NumberID = numId };
        newNum.AppendChild(new AbstractNumId { Val = abstractNumId });

        // Mode C: per-level start overrides. `start` is shorthand for
        // `startOverride.0`. `startOverride.N` (0..8) emits a <w:lvlOverride>
        // for that level. Each override is a fresh sibling element — no
        // collision logic needed since we're constructing a brand-new num.
        var startOverrides = new SortedDictionary<int, int>();
        if (properties.TryGetValue("start", out var startStr) && !string.IsNullOrEmpty(startStr))
            startOverrides[0] = ParseHelpers.SafeParseInt(startStr, "start");
        foreach (var kvp in properties)
        {
            const string prefix = "startOverride.";
            if (!kvp.Key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
            var lvlStr = kvp.Key.Substring(prefix.Length);
            var lvl = ParseHelpers.SafeParseInt(lvlStr, kvp.Key);
            if (lvl < 0 || lvl > 8)
                throw new ArgumentException($"{kvp.Key} level must be 0..8 (got {lvl}).");
            startOverrides[lvl] = ParseHelpers.SafeParseInt(kvp.Value, kvp.Key);
        }

        // Default-restart: Word's "two num instances on the same abstractNum"
        // behavior is "continue counting" unless the new num carries an
        // explicit <w:lvlOverride><w:startOverride/></w:lvlOverride>. That
        // contradicts what API users expect ("a new num instance = independent
        // counter"), so by default we inject a startOverride on level 0 with
        // the abstractNum's level0 start value (typically 1). Users who want
        // Word's literal continuation behavior pass --prop continue=true.
        bool wantsContinue = properties.TryGetValue("continue", out var contRaw) && IsTruthy(contRaw);
        if (!wantsContinue && !startOverrides.ContainsKey(0))
        {
            var srcAbs = numbering.Elements<AbstractNum>()
                .First(a => a.AbstractNumberId?.Value == abstractNumId);
            var lvl0 = srcAbs.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == 0);
            int defaultStart = lvl0?.StartNumberingValue?.Val?.Value ?? 1;
            startOverrides[0] = defaultStart;
        }

        foreach (var (lvl, startVal) in startOverrides)
        {
            var lvlOverride = new LevelOverride { LevelIndex = lvl };
            lvlOverride.AppendChild(new StartOverrideNumberingValue { Val = startVal });
            newNum.AppendChild(lvlOverride);
        }

        numbering.AppendChild(newNum);
        numbering.Save();
        return $"/numbering/num[@id={numId}]";
    }

    /// <summary>
    /// Add an AbstractNum (numbering template) under /numbering. This is the
    /// definition layer — what a list "looks like": 9 levels with their
    /// own format, marker text, indent, start, justification, marker font, etc.
    ///
    /// Per-level customization via dotted keys: --prop level0.format=decimal
    /// --prop level0.text=%1. --prop level0.indent=720 ... up through level8.
    /// Bare keys (format/text/indent/start) are aliases for level0.* for
    /// backward compatibility with --type num mode A.
    ///
    /// Levels not explicitly set fall back to a sensible cycle: bullet glyphs
    /// (•/◦/▪) for bullet types, decimal/lowerLetter/lowerRoman cycle for ordered.
    /// </summary>
    private string AddAbstractNum(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPart = _doc.MainDocumentPart!;
        var numberingPart = mainPart.NumberingDefinitionsPart
            ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering ??= new Numbering();
        var numbering = numberingPart.Numbering;

        int abstractNumId;
        if (properties.ContainsKey("id"))
        {
            abstractNumId = ParseHelpers.SafeParseInt(properties["id"], "id");
            if (abstractNumId < 0)
                throw new ArgumentException($"abstractNumId must be >= 0 (got {abstractNumId}).");
            if (numbering.Elements<AbstractNum>().Any(a => a.AbstractNumberId?.Value == abstractNumId))
                throw new ArgumentException(
                    $"abstractNumId {abstractNumId} already exists. Pick a unique --prop id, or omit --prop id for auto-assignment.");
        }
        else
        {
            abstractNumId = numbering.Elements<AbstractNum>()
                .Select(a => a.AbstractNumberId?.Value ?? 0).DefaultIfEmpty(-1).Max() + 1;
        }

        BuildAbstractNumElement(numbering, abstractNumId, properties);
        numbering.Save();
        return $"/numbering/abstractNum[@id={abstractNumId}]";
    }

    /// <summary>
    /// Build a fully-populated AbstractNum and insert it into Numbering in
    /// schema-correct order. Used by both the dedicated AddAbstractNum and
    /// AddNum mode A (auto-create template). Returns nothing — caller already
    /// chose abstractNumId and just needs the side effect.
    /// </summary>
    private static void BuildAbstractNumElement(Numbering numbering, int abstractNumId, Dictionary<string, string> properties)
    {
        var abstractNum = new AbstractNum { AbstractNumberId = abstractNumId };

        // Schema order inside abstractNum:
        // nsid → multiLevelType → tmpl → name → styleLink → numStyleLink → lvl[0..8]
        var multiLevelType = properties.GetValueOrDefault("type", "hybridMultilevel").ToLowerInvariant() switch
        {
            "hybridmultilevel" or "hybrid" => MultiLevelValues.HybridMultilevel,
            "multilevel" or "multi" => MultiLevelValues.Multilevel,
            "singlelevel" or "single" => MultiLevelValues.SingleLevel,
            _ => throw new ArgumentException($"Unknown multiLevelType '{properties["type"]}'. Valid: hybridMultilevel, multilevel, singleLevel.")
        };
        abstractNum.AppendChild(new MultiLevelType { Val = multiLevelType });

        if (properties.TryGetValue("name", out var anName) && !string.IsNullOrEmpty(anName))
            abstractNum.AppendChild(new AbstractNumDefinitionName { Val = anName });
        if (properties.TryGetValue("styleLink", out var anSL) && !string.IsNullOrEmpty(anSL))
            abstractNum.AppendChild(new StyleLink { Val = anSL });
        if (properties.TryGetValue("numStyleLink", out var anNSL) && !string.IsNullOrEmpty(anNSL))
            abstractNum.AppendChild(new NumberingStyleLink { Val = anNSL });

        // Top-level format determines level fallback cycle. Bare keys map to level0
        // (backward compat: format=bullet, text=•, indent=720, start=N).
        var topFormatRaw = properties.GetValueOrDefault("format", "decimal").ToLowerInvariant();
        var topIsBullet = topFormatRaw is "bullet" or "unordered" or "ul";
        var bulletChars = new[] { "•", "◦", "▪" };

        for (int lvl = 0; lvl < 9; lvl++)
        {
            var level = new Level { LevelIndex = lvl };
            var prefix = $"level{lvl}.";

            // Per-level format with fallback cycle
            string levelFormatRaw;
            if (lvl == 0 && properties.TryGetValue("format", out var bareFmt))
                levelFormatRaw = bareFmt;
            else if (properties.TryGetValue(prefix + "format", out var perLvlFmt))
                levelFormatRaw = perLvlFmt;
            else if (topIsBullet)
                levelFormatRaw = "bullet";
            else
                levelFormatRaw = (lvl % 3) switch { 0 => "decimal", 1 => "lowerLetter", _ => "lowerRoman" };
            var numFmt = ParseNumberingFormat(levelFormatRaw);
            var isBulletAtThisLvl = numFmt.Equals(NumberFormatValues.Bullet);

            // start (default 1)
            int start = 1;
            if (lvl == 0 && properties.TryGetValue("start", out var bareStart))
                start = ParseHelpers.SafeParseInt(bareStart, "start");
            else if (properties.TryGetValue(prefix + "start", out var perLvlStart))
                start = ParseHelpers.SafeParseInt(perLvlStart, prefix + "start");
            level.AppendChild(new StartNumberingValue { Val = start });
            level.AppendChild(new NumberingFormat { Val = numFmt });

            // suff (tab|space|nothing) — default tab in OOXML, omit unless overridden
            if (properties.TryGetValue(prefix + "suff", out var suffRaw) && !string.IsNullOrEmpty(suffRaw))
            {
                var suffVal = suffRaw.ToLowerInvariant() switch
                {
                    "tab" => LevelSuffixValues.Tab,
                    "space" => LevelSuffixValues.Space,
                    "nothing" or "none" => LevelSuffixValues.Nothing,
                    _ => throw new ArgumentException($"Invalid {prefix}suff '{suffRaw}'. Valid: tab, space, nothing.")
                };
                level.AppendChild(new LevelSuffix { Val = suffVal });
            }

            // lvlText
            string lvlText;
            if (lvl == 0 && properties.TryGetValue("text", out var bareText))
                lvlText = bareText;
            else if (properties.TryGetValue(prefix + "text", out var perLvlText))
                lvlText = perLvlText;
            else if (isBulletAtThisLvl)
                lvlText = bulletChars[lvl % bulletChars.Length];
            else
                lvlText = $"%{lvl + 1}.";
            level.AppendChild(new LevelText { Val = lvlText });

            // lvlJc (justification): left|center|right (default left)
            var jcRaw = properties.GetValueOrDefault(prefix + "justification",
                properties.GetValueOrDefault(prefix + "jc", "left")).ToLowerInvariant();
            var jcVal = jcRaw switch
            {
                "left" or "start" => LevelJustificationValues.Left,
                "center" => LevelJustificationValues.Center,
                "right" or "end" => LevelJustificationValues.Right,
                _ => throw new ArgumentException($"Invalid {prefix}justification '{jcRaw}'. Valid: left, center, right.")
            };
            level.AppendChild(new LevelJustification { Val = jcVal });

            // pPr/ind (indent + hanging)
            int leftIndent;
            if (lvl == 0 && properties.TryGetValue("indent", out var bareIndent))
                leftIndent = ParseHelpers.SafeParseInt(bareIndent, "indent");
            else if (properties.TryGetValue(prefix + "indent", out var perLvlIndent))
                leftIndent = ParseHelpers.SafeParseInt(perLvlIndent, prefix + "indent");
            else
                leftIndent = (lvl + 1) * 720;
            int hanging = properties.TryGetValue(prefix + "hanging", out var hangingRaw)
                ? ParseHelpers.SafeParseInt(hangingRaw, prefix + "hanging")
                : 360;
            level.AppendChild(new PreviousParagraphProperties(
                new Indentation { Left = leftIndent.ToString(), Hanging = hanging.ToString() }
            ));

            // rPr — marker font/size/color/bold/italic. Only emit when caller
            // supplied at least one rPr-relevant prop, otherwise let Word use
            // defaults (don't write a stray empty <w:rPr/>).
            bool hasRpr = properties.ContainsKey(prefix + "font")
                       || properties.ContainsKey(prefix + "size")
                       || properties.ContainsKey(prefix + "color")
                       || properties.ContainsKey(prefix + "bold")
                       || properties.ContainsKey(prefix + "italic");
            if (hasRpr)
            {
                var nspr = new NumberingSymbolRunProperties();
                if (properties.TryGetValue(prefix + "font", out var fontRaw) && !string.IsNullOrEmpty(fontRaw))
                {
                    nspr.AppendChild(new RunFonts { Ascii = fontRaw, HighAnsi = fontRaw, EastAsia = fontRaw });
                }
                if (properties.TryGetValue(prefix + "size", out var sizeRaw) && !string.IsNullOrEmpty(sizeRaw))
                {
                    var halfPt = (int)Math.Round(ParseFontSize(sizeRaw) * 2, MidpointRounding.AwayFromZero);
                    nspr.AppendChild(new FontSize { Val = halfPt.ToString() });
                }
                if (properties.TryGetValue(prefix + "color", out var colorRaw) && !string.IsNullOrEmpty(colorRaw))
                {
                    nspr.AppendChild(new Color { Val = SanitizeHex(colorRaw) });
                }
                if (properties.TryGetValue(prefix + "bold", out var boldRaw) && IsTruthy(boldRaw))
                    nspr.AppendChild(new Bold());
                if (properties.TryGetValue(prefix + "italic", out var italRaw) && IsTruthy(italRaw))
                    nspr.AppendChild(new Italic());
                level.AppendChild(nspr);
            }

            abstractNum.AppendChild(level);
        }

        // Schema requires AbstractNum before NumberingInstance.
        var firstNumInstance = numbering.GetFirstChild<NumberingInstance>();
        if (firstNumInstance != null)
            numbering.InsertBefore(abstractNum, firstNumInstance);
        else
            numbering.AppendChild(abstractNum);
    }

    /// <summary>
    /// Add a single &lt;w:lvl&gt; under an existing &lt;w:abstractNum&gt;. Distinct from
    /// AddDefault → TryCreateTypedElement, which uses schema-aware AddChild and
    /// silently REPLACES any existing lvl in the same parent (data loss when a
    /// caller adds ilvl=0 then ilvl=1 — only ilvl=1 survives). This helper uses
    /// AppendChild so multiple levels coexist, validates ilvl ∈ 0..8 and
    /// start as Int32, and accepts the same per-lvl props (lvlText/format/start/
    /// indent/...) the abstractNum builder accepts via levelN.* prefix.
    /// </summary>
    private string AddLvl(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not AbstractNum abstractNum)
            throw new ArgumentException(
                $"--type lvl requires parent /numbering/abstractNum[@id=N], got {parentPath}.");

        if (!properties.TryGetValue("ilvl", out var ilvlRaw) || string.IsNullOrEmpty(ilvlRaw))
            throw new ArgumentException("--type lvl requires --prop ilvl=N (0..8).");

        // ilvl: must be integer in 0..8 (OOXML ST_DecimalNumber for lvl is 0..8).
        if (!int.TryParse(ilvlRaw, System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var ilvl))
            throw new ArgumentException($"ilvl must be an integer 0..8 (got '{ilvlRaw}').");
        if (ilvl < 0 || ilvl > 8)
            throw new ArgumentException($"ilvl must be in range 0..8 (got {ilvl}).");

        // If a lvl with this ilvl already exists (typically from
        // AddAbstractNum's default lvl[0..8] pre-population), replace it in
        // place. New ilvl values are appended. The schema-aware AddChild path
        // in AddDefault collapsed every lvl onto a single slot; this dedicated
        // helper keeps siblings distinct and only swaps when ilvl matches.
        var existing = abstractNum.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == ilvl);

        // start: integer (no float, no overflow). Default 1.
        int start = 1;
        if (properties.TryGetValue("start", out var startRaw) && !string.IsNullOrEmpty(startRaw))
        {
            if (!int.TryParse(startRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out start))
                throw new ArgumentException(
                    $"start must be a 32-bit integer (got '{startRaw}'). Floats and values outside Int32 range are not accepted.");
        }

        var level = new Level { LevelIndex = ilvl };

        // numFmt: default decimal. Also accept 'numFmt' alias.
        var fmtRaw = properties.GetValueOrDefault("format",
            properties.GetValueOrDefault("numFmt", "decimal"));
        var numFmt = ParseNumberingFormat(fmtRaw);

        level.AppendChild(new StartNumberingValue { Val = start });
        level.AppendChild(new NumberingFormat { Val = numFmt });

        // lvlRestart (optional). CT_Lvl schema order places lvlRestart after
        // numFmt, before pStyle/isLgl/suff/lvlText.
        if (properties.TryGetValue("lvlRestart", out var lvlRestartRaw) && !string.IsNullOrEmpty(lvlRestartRaw))
        {
            if (!int.TryParse(lvlRestartRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var lrV))
                throw new ArgumentException($"lvlRestart must be a 32-bit integer (got '{lvlRestartRaw}').");
            level.AppendChild(new LevelRestart { Val = lrV });
        }

        // isLgl (optional). Schema order: after pStyle, before suff/lvlText.
        if (properties.TryGetValue("isLgl", out var isLglRaw) && IsTruthy(isLglRaw))
        {
            level.AppendChild(new IsLegalNumberingStyle());
        }

        // suff (optional)
        if (properties.TryGetValue("suff", out var suffRaw) && !string.IsNullOrEmpty(suffRaw))
        {
            var suffVal = suffRaw.ToLowerInvariant() switch
            {
                "tab" => LevelSuffixValues.Tab,
                "space" => LevelSuffixValues.Space,
                "nothing" or "none" => LevelSuffixValues.Nothing,
                _ => throw new ArgumentException($"Invalid suff '{suffRaw}'. Valid: tab, space, nothing.")
            };
            level.AppendChild(new LevelSuffix { Val = suffVal });
        }

        // lvlText: accept both 'text' and 'lvlText' aliases. Default: %{ilvl+1}. for
        // ordered, • for bullet.
        string lvlText;
        if (properties.TryGetValue("lvlText", out var ltRaw) && !string.IsNullOrEmpty(ltRaw))
            lvlText = ltRaw;
        else if (properties.TryGetValue("text", out var tRaw) && !string.IsNullOrEmpty(tRaw))
            lvlText = tRaw;
        else
            lvlText = numFmt.Equals(NumberFormatValues.Bullet) ? "•" : $"%{ilvl + 1}.";
        level.AppendChild(new LevelText { Val = lvlText });

        // jc (optional)
        if (properties.TryGetValue("justification", out var jcRaw) ||
            properties.TryGetValue("jc", out jcRaw))
        {
            var jcVal = jcRaw.ToLowerInvariant() switch
            {
                "left" or "start" => LevelJustificationValues.Left,
                "center" => LevelJustificationValues.Center,
                "right" or "end" => LevelJustificationValues.Right,
                _ => throw new ArgumentException($"Invalid justification '{jcRaw}'. Valid: left, center, right.")
            };
            level.AppendChild(new LevelJustification { Val = jcVal });
        }

        // pPr/ind (optional)
        int? leftIndent = null;
        if (properties.TryGetValue("indent", out var indRaw))
        {
            if (!int.TryParse(indRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var lv))
                throw new ArgumentException($"indent must be an integer in twips (got '{indRaw}').");
            leftIndent = lv;
        }
        int? hanging = null;
        if (properties.TryGetValue("hanging", out var hangRaw))
        {
            if (!int.TryParse(hangRaw, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var hv))
                throw new ArgumentException($"hanging must be an integer in twips (got '{hangRaw}').");
            hanging = hv;
        }
        if (leftIndent.HasValue || hanging.HasValue)
        {
            var ind = new Indentation();
            if (leftIndent.HasValue) ind.Left = leftIndent.Value.ToString();
            if (hanging.HasValue) ind.Hanging = hanging.Value.ToString();
            level.AppendChild(new PreviousParagraphProperties(ind));
        }

        // CRITICAL: AppendChild — NOT AddChild. Schema-aware AddChild treats
        // <w:lvl> as a single-instance child slot (the SDK's metadata says
        // "lvl[0..8]" but its schema model still flags them all as the same
        // child kind), so it would silently replace whatever lvl already
        // exists. AppendChild keeps every level distinct.
        if (existing != null)
        {
            existing.InsertBeforeSelf(level);
            existing.Remove();
        }
        else
        {
            abstractNum.AppendChild(level);
        }

        var numberingPart = _doc.MainDocumentPart?.NumberingDefinitionsPart;
        numberingPart?.Numbering?.Save();

        var absId = abstractNum.AbstractNumberId?.Value ?? 0;
        return $"/numbering/abstractNum[@id={absId}]/lvl[@ilvl={ilvl}]";
    }

    private static NumberFormatValues ParseNumberingFormat(string raw)
    {
        return raw.ToLowerInvariant() switch
        {
            "decimal" => NumberFormatValues.Decimal,
            "bullet" or "unordered" or "ul" => NumberFormatValues.Bullet,
            "lowerletter" or "loweralpha" => NumberFormatValues.LowerLetter,
            "upperletter" or "upperalpha" => NumberFormatValues.UpperLetter,
            "lowerroman" => NumberFormatValues.LowerRoman,
            "upperroman" => NumberFormatValues.UpperRoman,
            "ordinal" => NumberFormatValues.Ordinal,
            "cardinaltext" => NumberFormatValues.CardinalText,
            "ordinaltext" => NumberFormatValues.OrdinalText,
            "chinesecounting" => NumberFormatValues.ChineseCounting,
            "chineselegalsimplified" => NumberFormatValues.ChineseLegalSimplified,
            "chinesecountingthousand" => NumberFormatValues.ChineseCountingThousand,
            "ideographdigital" => NumberFormatValues.IdeographDigital,
            "japanesecounting" => NumberFormatValues.JapaneseCounting,
            "decimalzero" => NumberFormatValues.DecimalZero,
            "decimalenclosedcircle" => NumberFormatValues.DecimalEnclosedCircle,
            "decimalfullwidth" => NumberFormatValues.DecimalFullWidth,
            "none" => NumberFormatValues.None,
            _ => throw new ArgumentException($"Unknown numbering format '{raw}'. Common values: decimal, bullet, lowerLetter, upperLetter, lowerRoman, upperRoman, chineseCounting.")
        };
    }

    // Resolve the SectionProperties that a header/footer reference should
    // attach to, based on the parent path. `/section[N]` targets the carrier
    // paragraph's sectPr (mirrors NavigateToElement); `/`, `/body`, or any
    // other path falls back to the body-level (final) sectPr.
    private SectionProperties? ResolveTargetSectPrForHeaderFooter(string parentPath)
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return null;
        if (!string.IsNullOrEmpty(parentPath))
        {
            var m = System.Text.RegularExpressions.Regex.Match(
                parentPath, @"^/section\[(\d+)\]/?$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (m.Success && int.TryParse(m.Groups[1].Value, out var n))
            {
                var sectParas = body.Elements<Paragraph>()
                    .Where(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
                    .ToList();
                if (n >= 1 && n <= sectParas.Count)
                    return sectParas[n - 1].ParagraphProperties!.GetFirstChild<SectionProperties>();
            }
        }
        return body.Elements<SectionProperties>().LastOrDefault();
    }

    private string AddHeader(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPartH = _doc.MainDocumentPart!;

        // Resolve requested header type first, so we can reject duplicates before
        // creating an orphaned HeaderPart.
        var preHeaderType = HeaderFooterValues.Default;
        if (properties.TryGetValue("type", out var preHTypeStr) ||
            properties.TryGetValue("kind", out preHTypeStr) ||
            properties.TryGetValue("ref", out preHTypeStr))
        {
            preHeaderType = preHTypeStr.ToLowerInvariant() switch
            {
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                "default" => HeaderFooterValues.Default,
                _ => throw new ArgumentException($"Invalid header type: '{preHTypeStr}'. Valid values: default, first, even.")
            };
        }
        var preSectPr = ResolveTargetSectPrForHeaderFooter(parentPath);
        if (preSectPr != null && preSectPr.Elements<HeaderReference>()
                .Any(r => r.Type != null && r.Type.Value == preHeaderType))
        {
            throw new ArgumentException(
                $"Header of type '{HeaderFooterTypeName(preHeaderType)}' already exists in this section. " +
                "Remove the existing one first or use --prop type=<first|even>.");
        }

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

        // AssignParaId stamps w14:paraId / w14:textId on each w:p. Those
        // attributes are MS-2010 extensions and OpenXmlValidator rejects
        // them with Sch_UndeclaredAttribute unless the part declares the
        // w14 namespace and lists it in mc:Ignorable. The body part
        // (document.xml) does this at the document root; header/footer
        // parts need the same so paragraphs validated independently
        // accept the extension attrs.
        var hRoot = new Header(hPara);
        hRoot.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        hRoot.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        hRoot.SetAttribute(new OpenXmlAttribute("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006", "w14"));
        headerPart.Header = hRoot;
        headerPart.Header.Save();

        var hBody = mainPartH.Document!.Body!;
        var hSectPr = ResolveTargetSectPrForHeaderFooter(parentPath)
            ?? hBody.AppendChild(new SectionProperties());

        var headerType = preHeaderType;

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
        // CONSISTENCY(headerfooter-effective-toggle): mirror the type=first
        // → titlePg auto-write pattern. Without /settings/evenAndOddHeaders,
        // Word silently ignores the even header reference at render time.
        if (headerType == HeaderFooterValues.Even)
        {
            var hSettingsPart = mainPartH.DocumentSettingsPart
                ?? mainPartH.AddNewPart<DocumentSettingsPart>();
            hSettingsPart.Settings ??= new Settings();
            if (hSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>() == null)
                hSettingsPart.Settings.AddChild(new EvenAndOddHeaders(), throwOnError: false);
            hSettingsPart.Settings.Save();
        }

        var hIdx = mainPartH.HeaderParts.ToList().IndexOf(headerPart);
        return $"/header[{hIdx + 1}]";
    }

    private string AddFooter(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mainPartF = _doc.MainDocumentPart!;

        // Resolve requested footer type first, so we can reject duplicates before
        // creating an orphaned FooterPart.
        var preFooterType = HeaderFooterValues.Default;
        if (properties.TryGetValue("type", out var preFTypeStr) ||
            properties.TryGetValue("kind", out preFTypeStr) ||
            properties.TryGetValue("ref", out preFTypeStr))
        {
            preFooterType = preFTypeStr.ToLowerInvariant() switch
            {
                "first" => HeaderFooterValues.First,
                "even" => HeaderFooterValues.Even,
                "default" => HeaderFooterValues.Default,
                _ => throw new ArgumentException($"Invalid footer type: '{preFTypeStr}'. Valid values: default, first, even.")
            };
        }
        var preFSectPr = ResolveTargetSectPrForHeaderFooter(parentPath);
        if (preFSectPr != null && preFSectPr.Elements<FooterReference>()
                .Any(r => r.Type != null && r.Type.Value == preFooterType))
        {
            throw new ArgumentException(
                $"Footer of type '{HeaderFooterTypeName(preFooterType)}' already exists in this section. " +
                "Remove the existing one first or use --prop type=<first|even>.");
        }

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

        // Same w14 / mc:Ignorable declaration as AddHeader: paragraphs
        // here also carry w14:paraId / w14:textId from AssignParaId, and
        // OpenXmlValidator rejects them as undeclared without this.
        var fRoot = new Footer(fPara);
        fRoot.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        fRoot.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        fRoot.SetAttribute(new OpenXmlAttribute("Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006", "w14"));
        footerPart.Footer = fRoot;
        footerPart.Footer.Save();

        var fBody = mainPartF.Document!.Body!;
        var fSectPr = ResolveTargetSectPrForHeaderFooter(parentPath)
            ?? fBody.AppendChild(new SectionProperties());

        var footerType = preFooterType;

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
        // CONSISTENCY(headerfooter-effective-toggle): even-footer also needs
        // settings.xml/w:evenAndOddHeaders to render.
        if (footerType == HeaderFooterValues.Even)
        {
            var fSettingsPart = mainPartF.DocumentSettingsPart
                ?? mainPartF.AddNewPart<DocumentSettingsPart>();
            fSettingsPart.Settings ??= new Settings();
            if (fSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>() == null)
                fSettingsPart.Settings.AddChild(new EvenAndOddHeaders(), throwOnError: false);
            fSettingsPart.Settings.Save();
        }

        var fIdx = mainPartF.FooterParts.ToList().IndexOf(footerPart);
        return $"/footer[{fIdx + 1}]";
    }
}
