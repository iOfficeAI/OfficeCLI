// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Batch Set: if path looks like a selector (not starting with /), Query → Set each
        if (!string.IsNullOrEmpty(path) && !path.StartsWith("/"))
        {
            var targets = Query(path);
            if (targets.Count == 0)
                throw new ArgumentException($"No elements matched selector: {path}");
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        // Unified find: if 'find' key is present (at any path level), route to ProcessFind
        if (properties.TryGetValue("find", out var findText))
        {
            var replace = properties.TryGetValue("replace", out var r) ? r : null;
            // Separate run-level format properties from paragraph-level properties
            var formatProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var paraProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (key, value) in properties)
            {
                var k = key.ToLowerInvariant();
                if (k is "find" or "replace" or "scope" or "regex") continue;
                // Paragraph-level properties go to paraProps
                if (k is "style" or "alignment" or "align" or "firstlineindent" or "leftindent" or "indentleft"
                    or "indent" or "rightindent" or "indentright" or "hangingindent" or "spacebefore"
                    or "spaceafter" or "linespacing" or "keepnext" or "keeplines" or "pagebreakbefore"
                    or "widowcontrol" or "liststyle" or "start" or "text" or "formula"
                    or "contextualspacing")
                    paraProps[key] = value;
                else
                    formatProps[key] = value;
            }

            if (replace == null && formatProps.Count == 0 && paraProps.Count == 0)
                throw new ArgumentException("'find' requires either 'replace' and/or format properties (e.g. bold, highlight, color).");

            // CONSISTENCY(find-regex): canonical site for the `regex=true` → `r"..."`
            // raw-string normalization. `mark` and the other handlers' Set paths all
            // copy this pattern verbatim. To change the find/regex protocol,
            // grep "CONSISTENCY(find-regex)" and update every site project-wide;
            // do not diverge in a single handler.
            if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findText.StartsWith("r\"") && !findText.StartsWith("r'"))
                findText = $"r\"{findText}\"";

            var effectivePath = (path is "" or "/") ? "/body" : path;
            var matchCount = ProcessFind(effectivePath, findText, replace, formatProps.Count > 0 ? formatProps : new Dictionary<string, string>());
            LastFindMatchCount = matchCount;

            // Apply paragraph-level properties to the matched paragraphs
            if (paraProps.Count > 0)
            {
                var paragraphs = ResolveParagraphsForFind(effectivePath);
                foreach (var para in paragraphs)
                {
                    var pProps = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                    foreach (var (key, value) in paraProps)
                        ApplyParagraphLevelProperty(pProps, key, value);
                }
            }

            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Document-level properties
        if (path == "/" || path == "" || path.Equals("/body", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            _doc.MainDocumentPart?.Document?.Save();
            return unsupported;
        }

        // Handle /settings path — route to SetDocumentProperties which calls TrySetDocSetting
        if (path.Equals("/settings", StringComparison.OrdinalIgnoreCase))
        {
            SetDocumentProperties(properties, unsupported);
            EnsureSettings().Save();
            return unsupported;
        }

        // Handle /watermark path
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase))
            return SetWatermarkPath(properties);

        // FormField paths: /formfield[N] or /formfield[name]
        // Routed BEFORE ParsePath because the generic predicate validator
        // only accepts positive-integer / last() / [@attr=v] predicates and
        // would reject the documented /formfield[name] form.
        var ffSetMatchEarly = System.Text.RegularExpressions.Regex.Match(path, @"^/formfield\[(\w+)\]$");
        if (ffSetMatchEarly.Success)
        {
            var allFormFields = FindFormFields();
            var indexOrName = ffSetMatchEarly.Groups[1].Value;
            (FieldInfo Field, FormFieldData FfData) target;
            if (int.TryParse(indexOrName, out var ffIdx))
            {
                if (ffIdx < 1 || ffIdx > allFormFields.Count)
                    throw new ArgumentException($"FormField {ffIdx} not found (total: {allFormFields.Count})");
                target = allFormFields[ffIdx - 1];
            }
            else
            {
                target = allFormFields.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == indexOrName);
                if (target.Field == null)
                    throw new ArgumentException($"FormField '{indexOrName}' not found");
            }
            return SetFormField(target, properties);
        }

        // Positional aliases /numbering/abstractNum[N] and /numbering/num[N]
        // translate to canonical [@id=K] form (mirrors Get's normalization in
        // commit 0257e8ca). Without this, Set on positional paths fell
        // through to generic Navigation, which has no NumberingInstance
        // branch — and CLI printed "Updated …" while nothing changed on
        // disk. Tagged CONSISTENCY(numbering-positional-normalize).
        var numPosSetMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/(abstractNum|num)\[(\d+)\](.*)$");
        if (numPosSetMatch.Success)
        {
            var kind = numPosSetMatch.Groups[1].Value;
            var posIdx = int.Parse(numPosSetMatch.Groups[2].Value); // 1-based
            var rest = numPosSetMatch.Groups[3].Value;
            var nb = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
            int? canonId = null;
            if (kind == "abstractNum")
            {
                var abs = nb?.Elements<AbstractNum>().ElementAtOrDefault(posIdx - 1);
                canonId = abs?.AbstractNumberId?.Value;
            }
            else
            {
                var inst = nb?.Elements<NumberingInstance>().ElementAtOrDefault(posIdx - 1);
                canonId = inst?.NumberID?.Value;
            }
            if (canonId != null)
                return Set($"/numbering/{kind}[@id={canonId}]{rest}", properties);
        }

        // Numbering paths: /numbering/abstractNum[@id=N] and
        // /numbering/abstractNum[@id=N]/level[L]. Intercept BEFORE the generic
        // ParsePath call below — those paths use [@id=...] / [N starting at 0]
        // predicates that ParsePath's 1-based positional rule rejects.
        // Accept both /level[N] (positional, 0-based ilvl) and /lvl[@ilvl=N]
        // (canonical form returned by Get/Query — see R2 commit 48ee8c8c, R3
        // commit 2a634aeb). Without the @ilvl branch, Set silently no-ops on
        // the canonical path: the CLI prints "Updated" but numbering.Save()
        // never runs because the path falls through to generic Navigation
        // which has no Level branch in SetElement.
        var absNumSetMatchEarly = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/abstractNum\[@id=(\d+)\](?:/(?:level|lvl)\[(?:@ilvl=)?(\d+)\])?$");
        if (absNumSetMatchEarly.Success) return SetAbstractNumPath(absNumSetMatchEarly, properties);

        // /numbering/num[@id=N] — set abstractNumId on a NumberingInstance.
        // Without this intercept, generic Navigation finds the <w:num> element
        // but SetElement has no NumberingInstance branch, so the call returns
        // an empty unsupported list and the CLI prints "Updated …" while
        // nothing changes on disk.
        var numSetMatchEarly = System.Text.RegularExpressions.Regex.Match(
            path, @"^/numbering/num\[@id=(\d+)\]$");
        if (numSetMatchEarly.Success) return SetNumPath(numSetMatchEarly, properties);

        // Handle header/footer paths
        var hfParts = ParsePath(path);
        if (hfParts.Count >= 1)
        {
            var firstName = hfParts[0].Name.ToLowerInvariant();
            if ((firstName == "header" || firstName == "footer") && hfParts.Count == 1)
            {
                SetHeaderFooter(firstName, (hfParts[0].Index ?? 1) - 1, properties, unsupported);
                return unsupported;
            }
        }

        // Chart axis-by-role sub-path: /chart[N]/axis[@role=ROLE].
        var chartAxisSetMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisPath(chartAxisSetMatch, properties);

        // Chart paths: /chart[N] or /chart[N]/series[K]
        var chartMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success) return SetChartPath(chartMatch, properties);

        // Field paths: /field[N]
        var fieldSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/field\[(\d+)\]$");
        if (fieldSetMatch.Success) return SetFieldPath(fieldSetMatch, properties);

        // TOC paths: /toc[N], /toc (= first), /tableofcontents alias.
        var tocMatch = System.Text.RegularExpressions.Regex.Match(path,
            @"^/(?:toc|tableofcontents)(?:\[(\d+)\])?$");
        if (tocMatch.Success) return SetTocPath(tocMatch, properties);

        // Footnote paths: /footnote[N], /footnote[@footnoteId=N] (incl. -1/0
        // structural ids — separator/continuation/continuationNotice).
        var fnSetMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"/footnote\[(?:@footnoteId=)?(-?\d+)\]$");
        if (fnSetMatch.Success) return SetFootnotePath(fnSetMatch, properties);

        // Endnote paths: same shape as footnote.
        var enSetMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"/endnote\[(?:@endnoteId=)?(-?\d+)\]$");
        if (enSetMatch.Success) return SetEndnotePath(enSetMatch, properties);

        // Section paths: /section[N] or /body/sectPr[N] (canonical form returned by Get/Query)
        var secSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^(?:/section\[(\d+)\]|/body/sectPr(?:\[(\d+)\])?)$");
        if (secSetMatch.Success) return SetSectionPath(secSetMatch, properties);

        // Style paths: /styles/StyleId (set props on the style itself).
        // Restrict to a single segment so deeper paths like /styles/<id>/tab[N]
        // fall through to generic Navigation + SetElement (TabStop branch).
        var styleSetMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/styles/([^/]+)$");
        if (styleSetMatch.Success) return SetStylePath(styleSetMatch, properties);

        // CONSISTENCY(ole-shorthand-set): mirror the /body/ole[N] shorthand
        // already supported in Get (WordHandler.Query.cs) and Remove
        // (WordHandler.Mutations.cs). Without this intercept, Set falls through
        // to NavigateToElement which hits "No ole found at /body" because OLE
        // lives inside a run, not as a direct child of the body.
        var wordOleSetMatch = System.Text.RegularExpressions.Regex.Match(
            path,
            @"^(?<parent>/body|/header\[\d+\]|/footer\[\d+\])?/(?:ole|object|embed)\[(?<idx>\d+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (wordOleSetMatch.Success) return SetWordOlePath(wordOleSetMatch, properties);

        var parts = ParsePath(path);
        var element = NavigateToElement(parts, out var ctx);
        if (element == null)
            throw new ArgumentException($"Path not found: {path}" + (ctx != null ? $". {ctx}" : ""));

        // Clone element for rollback on failure (atomic: no partial modifications)
        var elementBackup = element.CloneNode(true);
        try
        {
        return SetElement(element, properties);
        }
        catch
        {
            // Rollback: restore element to pre-modification state
            element.Parent?.ReplaceChild(elementBackup, element);
            throw;
        }
    }

    private List<string> SetElement(OpenXmlElement element, Dictionary<string, string> properties)
    {
        if (element is Comment cmt) return SetElementComment(cmt, properties);
        if (element is BookmarkStart bk) return SetElementBookmark(bk, properties);
        if (element is SdtBlock || element is SdtRun) return SetElementSdt(element, properties);
        if (element is Run run) return SetElementRun(run, properties);
        if (element is Hyperlink hl) return SetElementHyperlink(hl, properties);
        if (element is M.Paragraph mPara) return SetElementMPara(mPara, properties);
        if (element is Paragraph para) return SetElementParagraph(para, properties);
        if (element is TableCell cell) return SetElementTableCell(cell, properties);
        if (element is TableRow row) return SetElementTableRow(row, properties);
        if (element is Table tbl) return SetElementTable(tbl, properties);
        if (element is TabStop tabStop) return SetElementTabStop(tabStop, properties);
        return new List<string>();
    }

    private void SetHeaderFooter(string kind, int index, Dictionary<string, string> properties, List<string> unsupported)
    {
        var mainPart = _doc.MainDocumentPart!;
        OpenXmlCompositeElement? container;
        OpenXmlPart partRef;

        if (kind == "header")
        {
            var part = mainPart.HeaderParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Header not found: /header[{index + 1}]");
            container = part.Header;
            partRef = part;
        }
        else
        {
            var part = mainPart.FooterParts.ElementAtOrDefault(index)
                ?? throw new ArgumentException($"Footer not found: /footer[{index + 1}]");
            container = part.Footer;
            partRef = part;
        }

        if (container == null)
            throw new ArgumentException($"{kind} content not found at index {index + 1}");

        var firstPara = container.Elements<Paragraph>().FirstOrDefault();
        if (firstPara == null)
        {
            firstPara = new Paragraph();
            container.AppendChild(firstPara);
        }
        var pProps = firstPara.ParagraphProperties ?? firstPara.PrependChild(new ParagraphProperties());

        foreach (var (key, value) in properties)
        {
            var k = key.ToLowerInvariant();
            if (ApplyParagraphLevelProperty(pProps, key, value))
            {
                // handled by paragraph-level helper
            }
            else switch (k)
            {
                case "text":
                {
                    RunProperties? existingRProps = null;
                    var existingRun = firstPara.Elements<Run>().FirstOrDefault();
                    if (existingRun?.RunProperties != null)
                        existingRProps = (RunProperties)existingRun.RunProperties.CloneNode(true);
                    foreach (var r in firstPara.Elements<Run>().ToList()) r.Remove();
                    var newRun = new Run();
                    if (existingRProps != null)
                        newRun.AppendChild(existingRProps);
                    newRun.AppendChild(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
                    firstPara.AppendChild(newRun);
                    break;
                }
                case "size" or "font" or "bold" or "italic" or "color" or "highlight" or "underline" or "strike":
                    // Apply run-level formatting to all runs in the container
                    foreach (var run in container.Descendants<Run>())
                        ApplyRunFormatting(EnsureRunProperties(run), key, value);
                    // Also update paragraph mark run properties so new runs inherit formatting
                    var markRPr = pProps.ParagraphMarkRunProperties ?? pProps.AppendChild(new ParagraphMarkRunProperties());
                    ApplyRunFormatting(markRPr, key, value);
                    break;
                case "type":
                {
                    // Mutate the HeaderReference/FooterReference Type attribute
                    // pointing at this part. Read side (WordHandler.Query.cs:660-666,
                    // 717-723) only inspects body-level SectionProperties, so the
                    // write side stays scoped to the same set for round-trip parity.
                    var newType = value.ToLowerInvariant() switch
                    {
                        "first" => HeaderFooterValues.First,
                        "even" => HeaderFooterValues.Even,
                        "default" => HeaderFooterValues.Default,
                        _ => throw new ArgumentException(
                            $"Invalid {kind} type: '{value}'. Valid values: default, first, even.")
                    };
                    var partRid = mainPart.GetIdOfPart(partRef);
                    var body = mainPart.Document?.Body
                        ?? throw new InvalidOperationException("Document body not found");
                    bool found = false;
                    foreach (var sp in body.Elements<SectionProperties>())
                    {
                        if (kind == "header")
                        {
                            var ownRef = sp.Elements<HeaderReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<HeaderReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Header of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        else
                        {
                            var ownRef = sp.Elements<FooterReference>().FirstOrDefault(r => r.Id?.Value == partRid);
                            if (ownRef == null) continue;
                            if (ownRef.Type?.Value == newType) { found = true; continue; }
                            if (sp.Elements<FooterReference>().Any(r => r != ownRef && r.Type?.Value == newType))
                                throw new ArgumentException(
                                    $"Footer of type '{value}' already exists in this section.");
                            ownRef.Type = newType;
                            found = true;
                        }
                        // Mirrors AddHeader: Title-page header requires <w:titlePg/> on the section.
                        if (newType == HeaderFooterValues.First && sp.GetFirstChild<TitlePage>() == null)
                            sp.AddChild(new TitlePage(), throwOnError: false);
                    }
                    if (!found) unsupported.Add(key);
                    break;
                }
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        if (kind == "header")
            mainPart.HeaderParts.ElementAt(index).Header?.Save();
        else
            mainPart.FooterParts.ElementAt(index).Footer?.Save();
    }

    // Border style format: "style" or "style;size" or "style;size;color" or "style;size;color;space"
    // Styles: none, single, thick, double, dotted, dashed, dotDash, dotDotDash, triple,
    //         thinThickSmallGap, thickThinSmallGap, thinThickThinSmallGap,
    //         thinThickMediumGap, thickThinMediumGap, thinThickThinMediumGap,
    //         thinThickLargeGap, thickThinLargeGap, thinThickThinLargeGap, wave, doubleWave, threeDEmboss, threeDEngrave
    /// <summary>Insert StyleParagraphProperties before StyleRunProperties to maintain OOXML schema order.</summary>
    private static StyleParagraphProperties EnsureStyleParagraphProperties(Style style)
    {
        var pPr = new StyleParagraphProperties();
        var rPr = style.StyleRunProperties;
        if (rPr != null)
            style.InsertBefore(pPr, rPr);
        else
            style.AppendChild(pPr);
        return pPr;
    }

    private static BorderValues ParseBorderStyle(string style) => style.ToLowerInvariant() switch
    {
        "none" => BorderValues.None,
        "nil" => BorderValues.Nil,
        "single" or "thin" => BorderValues.Single,
        "thick" or "medium" => BorderValues.Thick,
        "double" => BorderValues.Double,
        "dotted" => BorderValues.Dotted,
        "dashed" => BorderValues.Dashed,
        "dotdash" => BorderValues.DotDash,
        "dotdotdash" => BorderValues.DotDotDash,
        "triple" => BorderValues.Triple,
        "thinthicksmallgap" => BorderValues.ThinThickSmallGap,
        "thickthinsmallgap" => BorderValues.ThickThinSmallGap,
        "thinthickthinsmallgap" => BorderValues.ThinThickThinSmallGap,
        "thinthickmediumgap" => BorderValues.ThinThickMediumGap,
        "thickthinmediumgap" => BorderValues.ThickThinMediumGap,
        "thinthickthinmediumgap" => BorderValues.ThinThickThinMediumGap,
        "thinthicklargegap" => BorderValues.ThinThickLargeGap,
        "thickthinlargegap" => BorderValues.ThickThinLargeGap,
        "thinthickthinlargegap" => BorderValues.ThinThickThinLargeGap,
        "wave" => BorderValues.Wave,
        "doublewave" => BorderValues.DoubleWave,
        "threedembed" or "3demboss" => BorderValues.ThreeDEmboss,
        "threedengrave" or "3dengrave" => BorderValues.ThreeDEngrave,
        _ => throw new ArgumentException($"Invalid border style: '{style}'. Valid values: single, thick, double, dotted, dashed, none, triple, wave, etc.")
    };

    private static (BorderValues style, uint size, string? color, uint space) ParseBorderValue(string value)
    {
        var parts = value.Split(';');
        var style = ParseBorderStyle(parts[0]);
        uint size;
        if (parts.Length > 1)
        {
            if (!uint.TryParse(parts[1], out size))
                throw new ArgumentException($"Invalid border size '{parts[1]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
        }
        else
            size = style == BorderValues.Nil ? 0u : style == BorderValues.Thick ? 12u : 4u;
        string? color = parts.Length > 2 ? SanitizeHex(parts[2]) : null;
        uint space = 0u;
        if (parts.Length > 3 && !uint.TryParse(parts[3], out space))
            throw new ArgumentException($"Invalid border space '{parts[3]}', expected integer. Format: STYLE[;SIZE[;COLOR[;SPACE]]]");
        return (style, size, color, space);
    }

    private static T MakeBorder<T>(BorderValues style, uint size, string? color, uint space) where T : BorderType, new()
    {
        var b = new T { Val = style, Size = size, Space = space };
        if (color != null) b.Color = color;
        return b;
    }

    /// <summary>
    /// Apply a paragraph-level property. Returns true if handled, false if not recognized.
    /// Handles: style, alignment, indent, spacing, keepNext, keepLines, pageBreakBefore, widowControl, shading, pbdr.
    /// </summary>
    private bool ApplyParagraphLevelProperty(ParagraphProperties pProps, string key, string? value)
    {
        if (value is null) return false;
        switch (key.ToLowerInvariant())
        {
            case "style":
                pProps.ParagraphStyleId = new ParagraphStyleId { Val = value };
                return true;
            case "alignment" or "align":
                pProps.Justification = new Justification { Val = ParseJustification(value) };
                return true;
            case "firstlineindent":
                var indent = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                // Lenient input: accept "2cm", "0.5in", "18pt", or bare twips.
                indent.FirstLine = SpacingConverter.ParseWordSpacing(value).ToString();
                indent.Hanging = null;
                return true;
            case "leftindent" or "indentleft" or "indent":
                var indentL = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                // CONSISTENCY(lenient-spacing): mirror Add — accept cm/in/pt/twips via SpacingConverter.
                indentL.Left = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "rightindent" or "indentright":
                var indentR = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentR.Right = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "hangingindent" or "hanging":
                var indentH = pProps.Indentation ?? (pProps.Indentation = new Indentation());
                indentH.Hanging = SpacingConverter.ParseWordSpacing(value).ToString();
                indentH.FirstLine = null;
                return true;
            // Toggle props: always replace the element (don't `??=`) so an
            // existing `<w:foo w:val="false"/>` written by a previous Set or
            // by external tooling is correctly overridden when the new value
            // is true. With `??=` the val=false sticks and the toggle never
            // flips back to true (BUG-LT3).
            case "keepnext" or "keepwithnext":
                if (IsTruthy(value)) pProps.KeepNext = new KeepNext();
                else pProps.KeepNext = null;
                return true;
            case "keeplines" or "keeptogether":
                if (IsTruthy(value)) pProps.KeepLines = new KeepLines();
                else pProps.KeepLines = null;
                return true;
            case "pagebreakbefore":
                if (IsTruthy(value)) pProps.PageBreakBefore = new PageBreakBefore();
                else pProps.PageBreakBefore = null;
                return true;
            // fuzz-2: 'break=newPage' is the natural paragraph-context spelling
            // (mirrors section-context CONSISTENCY(section-type-alias) in
            // WordHandler.Set.Dispatch.cs:387). For a paragraph this maps to
            // pageBreakBefore=true; bare break=true also accepted.
            case "break":
                bool pbb = value.ToLowerInvariant() switch
                {
                    "newpage" or "page" or "nextpage" or "pagebreak" => true,
                    "none" or "" => false,
                    _ => IsTruthy(value)
                };
                if (pbb) pProps.PageBreakBefore = new PageBreakBefore();
                else pProps.PageBreakBefore = null;
                return true;
            case "widowcontrol" or "widoworphan":
                if (IsTruthy(value)) pProps.WidowControl = new WidowControl();
                else pProps.WidowControl = new WidowControl { Val = false };
                return true;
            case "contextualspacing" or "contextualSpacing":
                if (IsTruthy(value)) pProps.ContextualSpacing = new ContextualSpacing();
                else pProps.ContextualSpacing = null;
                return true;
            case "shading" or "shd":
                pProps.Shading = ParseShadingValue(value);
                return true;
            case "spacebefore":
                var spacingBefore = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingBefore.Before = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "spaceafter":
                var spacingAfter = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                spacingAfter.After = SpacingConverter.ParseWordSpacing(value).ToString();
                return true;
            case "linespacing":
                var spacingLine = pProps.SpacingBetweenLines ?? (pProps.SpacingBetweenLines = new SpacingBetweenLines());
                var (lsTwips, lsIsMultiplier) = SpacingConverter.ParseWordLineSpacing(value);
                spacingLine.Line = lsTwips.ToString();
                spacingLine.LineRule = lsIsMultiplier ? LineSpacingRuleValues.Auto : LineSpacingRuleValues.Exact;
                return true;
            case "numId" or "numid":
                var numPr = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                var numIdVal = ParseHelpers.SafeParseInt(value, "numId");
                if (numIdVal < 0)
                    throw new ArgumentException($"numId must be >= 0 (got {numIdVal}). Use numId=0 to remove numbering.");
                if (numIdVal > 0)
                {
                    var numbering = _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering;
                    var numExists = numbering?.Elements<NumberingInstance>()
                        .Any(n => n.NumberID?.Value == numIdVal) ?? false;
                    if (!numExists)
                        throw new ArgumentException(
                            $"numId={numIdVal} not found in /numbering. " +
                            "Create the num first (add /numbering --type num), or use numId=0 to remove numbering.");
                }
                numPr.NumberingId = new NumberingId { Val = numIdVal };
                return true;
            case "numLevel" or "numlevel" or "ilvl" or "listLevel" or "listlevel":
                var numPr2 = pProps.NumberingProperties ?? (pProps.NumberingProperties = new NumberingProperties());
                var ilvlSetVal = ParseHelpers.SafeParseInt(value, "numLevel");
                if (ilvlSetVal < 0 || ilvlSetVal > 8)
                    throw new ArgumentException($"ilvl must be in range 0..8 (got {ilvlSetVal}).");
                numPr2.NumberingLevelReference = new NumberingLevelReference { Val = ilvlSetVal };
                return true;
            case "pbdr.top" or "pbdr.bottom" or "pbdr.left" or "pbdr.right" or "pbdr.between" or "pbdr.bar" or "pbdr.all" or "pbdr":
            case "border.all" or "border" or "border.top" or "border.bottom" or "border.left" or "border.right":
                ApplyParagraphBorders(pProps, key, value);
                return true;
            default:
                return false;
        }
    }


    private static void ApplyParagraphBorders(ParagraphProperties pProps, string key, string value)
    {
        var borders = pProps.ParagraphBorders;
        if (borders == null)
        {
            borders = new ParagraphBorders();
            pProps.ParagraphBorders = borders; // typed setter maintains CT_PPr schema order
        }
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "pbdr.right" or "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "pbdr.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyStyleParagraphBorders(StyleParagraphProperties spPr, string key, string value)
    {
        var borders = spPr.GetFirstChild<ParagraphBorders>();
        if (borders == null)
        {
            borders = new ParagraphBorders();
            // StyleParagraphProperties is also OneSequence — use SetElement pattern
            // ParagraphBorders element order index is after Indentation and before Shading
            var afterRef = (OpenXmlElement?)spPr.GetFirstChild<Indentation>()
                ?? (OpenXmlElement?)spPr.GetFirstChild<SpacingBetweenLines>()
                ?? (OpenXmlElement?)spPr.GetFirstChild<Justification>();
            if (afterRef != null)
                spPr.InsertAfter(borders, afterRef);
            else
                spPr.PrependChild(borders);
        }
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "pbdr.all" or "pbdr" or "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.top" or "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "pbdr.bottom" or "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "pbdr.left" or "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "pbdr.right" or "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "pbdr.between":
                borders.BetweenBorder = MakeBorder<BetweenBorder>(style, size, color, space);
                break;
            case "pbdr.bar":
                borders.BarBorder = MakeBorder<BarBorder>(style, size, color, space);
                break;
        }
    }

    private static void ApplyTableBorders(TableProperties tblPr, string key, string value)
    {
        var borders = tblPr.TableBorders ?? tblPr.AppendChild(new TableBorders());
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space);
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.insideh" or "border.horizontal":
                borders.InsideHorizontalBorder = MakeBorder<InsideHorizontalBorder>(style, size, color, space);
                break;
            case "border.insidev" or "border.vertical":
                borders.InsideVerticalBorder = MakeBorder<InsideVerticalBorder>(style, size, color, space);
                break;
        }
    }

    /// <summary>
    /// CT_TcPr child schema order. Used by InsertTcPrChildInOrder to insert
    /// new tcPr children at their schema position rather than the tail.
    /// Children whose type isn't on this list (mc:AlternateContent and
    /// extensions, for instance) are tolerated — they sort to the end via
    /// the IndexOf == -1 sentinel.
    /// </summary>
    private static readonly Type[] s_tcPrChildOrder =
    [
        typeof(ConditionalFormatStyle),
        typeof(TableCellWidth),
        typeof(GridSpan),
        typeof(HorizontalMerge),
        typeof(VerticalMerge),
        typeof(TableCellBorders),
        typeof(Shading),
        typeof(NoWrap),
        typeof(TableCellMargin),
        typeof(TextDirection),
        typeof(TableCellFitText),
        typeof(TableCellVerticalAlignment),
        typeof(HideMark),
        // headers/cellIns/cellDel/cellMerge/tcPrChange follow but are rare
        // enough that we let the SDK's own setters handle them; they get
        // sentinel positions (-1) and end up at the tail, which is correct
        // when nothing else past tcPr has been written.
    ];

    private static void InsertTcPrChildInOrder(TableCellProperties tcPr, OpenXmlElement child)
    {
        var targetIdx = Array.IndexOf(s_tcPrChildOrder, child.GetType());
        if (targetIdx < 0)
        {
            tcPr.AppendChild(child);
            return;
        }
        foreach (var sibling in tcPr.ChildElements)
        {
            var sibIdx = Array.IndexOf(s_tcPrChildOrder, sibling.GetType());
            if (sibIdx > targetIdx)
            {
                tcPr.InsertBefore(child, sibling);
                return;
            }
        }
        tcPr.AppendChild(child);
    }

    private static void ApplyCellBorders(TableCellProperties tcPr, string key, string value)
    {
        // CT_TcPr child sequence is strict: cnfStyle → tcW → gridSpan →
        // hMerge → vMerge → tcBorders → shd → noWrap → tcMar →
        // textDirection → tcFitText → vAlign → hideMark → ... → tcPrChange.
        // Plain AppendChild lands tcBorders at the tail, after shd/vAlign/
        // tcMar that earlier setter calls already wrote, producing
        // Sch_UnexpectedElementContentExpectingComplex on tcBorders. Insert
        // before the first existing sibling that should come after tcBorders.
        var borders = tcPr.TableCellBorders;
        if (borders == null)
        {
            borders = new TableCellBorders();
            InsertTcPrChildInOrder(tcPr, borders);
        }
        var (style, size, color, space) = ParseBorderValue(value);

        switch (key.ToLowerInvariant())
        {
            case "border.all" or "border":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.top":
                borders.TopBorder = MakeBorder<TopBorder>(style, size, color, space);
                break;
            case "border.bottom":
                borders.BottomBorder = MakeBorder<BottomBorder>(style, size, color, space);
                break;
            case "border.left":
                borders.LeftBorder = MakeBorder<LeftBorder>(style, size, color, space);
                break;
            case "border.right":
                borders.RightBorder = MakeBorder<RightBorder>(style, size, color, space);
                break;
            case "border.tl2br":
                borders.TopLeftToBottomRightCellBorder = MakeBorder<TopLeftToBottomRightCellBorder>(style, size, color, space);
                break;
            case "border.tr2bl":
                borders.TopRightToBottomLeftCellBorder = MakeBorder<TopRightToBottomLeftCellBorder>(style, size, color, space);
                break;
        }
    }

    /// <summary>
    /// Apply gradient fill to a Word table cell using mc:AlternativeContent with w14:gradFill.
    /// Fallback is a solid shading with the start color.
    /// </summary>
    private static void ApplyCellGradient(TableCellProperties tcPr, string startColor, string endColor, int angleDeg)
    {
        // Sanitize colors: strip 8-char RRGGBBAA to 6-char RGB (w14:srgbClr requires 6 chars)
        var (startRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(startColor);
        var (endRgb, _) = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(endColor);

        // Remove existing shading/gradient
        RemoveCellGradient(tcPr);
        tcPr.Shading?.Remove();

        // Set fallback solid fill
        tcPr.Shading = new Shading { Val = ShadingPatternValues.Clear, Fill = startRgb };

        // Build w14:gradFill XML via raw OpenXml
        var w14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        // Convert angle to OOXML 60000ths of a degree
        var angleOoxml = angleDeg * 60000;

        var acElement = new OpenXmlUnknownElement("mc", "AlternateContent", mcNs);
        acElement.InnerXml = $@"<mc:Choice xmlns:mc=""{mcNs}"" xmlns:w14=""{w14Ns}"" Requires=""w14"">
    <w14:gradFill>
      <w14:gsLst>
        <w14:gs w14:pos=""0"">
          <w14:srgbClr w14:val=""{startRgb}""/>
        </w14:gs>
        <w14:gs w14:pos=""100000"">
          <w14:srgbClr w14:val=""{endRgb}""/>
        </w14:gs>
      </w14:gsLst>
      <w14:lin w14:ang=""{angleOoxml}"" w14:scaled=""1""/>
    </w14:gradFill>
  </mc:Choice>";

        tcPr.AppendChild(acElement);
    }

    /// <summary>
    /// Remove any existing gradient mc:AlternateContent from table cell properties.
    /// </summary>
    private static void RemoveCellGradient(TableCellProperties tcPr)
    {
        var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var existing = tcPr.ChildElements
            .Where(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs)
            .ToList();
        foreach (var e in existing) e.Remove();
    }

    /// <summary>
    /// Parse twips from a string with optional unit suffix: "1.5cm", "0.5in", "36pt", or raw twips.
    /// 1 inch = 1440 twips, 1 cm = 567 twips, 1 pt = 20 twips.
    /// </summary>
    private static TablePositionProperties EnsureTablePositionProperties(TableProperties tblPr)
    {
        var tpp = tblPr.GetFirstChild<TablePositionProperties>();
        if (tpp == null)
        {
            tpp = new TablePositionProperties
            {
                VerticalAnchor = VerticalAnchorValues.Page,
                HorizontalAnchor = HorizontalAnchorValues.Page
            };
            // CT_TblPr schema order: tblStyle → tblpPr → tblOverlap → ...
            var tblStyle = tblPr.GetFirstChild<TableStyle>();
            if (tblStyle != null)
                tblStyle.InsertAfterSelf(tpp);
            else
                tblPr.PrependChild(tpp);
        }
        return tpp;
    }

    internal static uint ParseTwips(string value)
    {
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (cm)");
            return (uint)Math.Round(num * 1440.0 / 2.54);
        }
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (in)");
            return (uint)Math.Round(num * 1440);
        }
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            var num = ParseHelpers.SafeParseDouble(value[..^2], "twips (pt)");
            return (uint)Math.Round(num * 20);
        }
        return ParseHelpers.SafeParseUint(value, "twips");
    }
}
