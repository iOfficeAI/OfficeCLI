using System.IO.Compression;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Set Layer ====================

    /// <summary>
    /// Apply a set of properties to the element at the given path.
    /// Returns names of properties that could not be applied (unsupported).
    /// </summary>
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();

        // Batch Set: @selector path → Query + Set each result
        if (path.StartsWith("@"))
        {
            var matches = Query(path);
            foreach (var match in matches)
            {
                if (!string.IsNullOrEmpty(match.Path))
                    Set(match.Path, new Dictionary<string, string>(properties));
            }
            return unsupported;
        }

        // Find/replace: supports any scope path, regex, format filter
        if (properties.ContainsKey("find"))
        {
            var findText = properties["find"];
            var replaceText = properties.GetValueOrDefault("replace") ?? "";

            XElement? scope = null;
            if (path is not ("/" or "" or "/body"))
            {
                try { scope = ResolvePath(path); } catch { /* fall through to full doc */ }
            }

            var formatFilter = new Dictionary<string, string>();
            foreach (var fk in new[] { "bold", "italic", "color", "fontsize" })
            {
                if (properties.TryGetValue(fk, out var fv))
                    formatFilter[fk] = fv;
            }

            FindAndReplace(findText, replaceText, scope,
                formatFilter.Count > 0 ? formatFilter : null);

            var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
            foreach (var k in new[] { "find", "replace", "bold", "italic", "color", "fontsize" })
                remaining.Remove(k);
            if (remaining.Count > 0)
                unsupported.AddRange(remaining.Keys);
            return unsupported;
        }

        // Label-based table fill: fill:라벨=값 (Plan 70)
        var fillKeys = properties.Keys
            .Where(k => k.StartsWith("fill:", StringComparison.OrdinalIgnoreCase)).ToList();
        if (fillKeys.Count > 0)
        {
            var fillProps = fillKeys.ToDictionary(
                k => k["fill:".Length..],
                k => properties[k],
                StringComparer.OrdinalIgnoreCase);
            FillByLabel(fillProps);
            // If only fill: props were provided, return (don't mark as unsupported)
            if (fillKeys.Count == properties.Count) return unsupported;
        }

        // /table/fill pseudo-path: all props are label=value
        if (path.Equals("/table/fill", StringComparison.OrdinalIgnoreCase))
        {
            FillByLabel(properties);
            return unsupported;
        }

        // Document-level properties
        if (path is "/" or "" or "/body")
        {

            // Document-level properties (default font, default font size)
            var docHandled = false;
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "defaultfont" or "basefont":
                        var charPrFont = FindCharPr("0");
                        if (charPrFont != null) { ApplyCharPrProperty(charPrFont, "fonthangul", value); docHandled = true; }
                        break;
                    case "defaultfontsize" or "basefontsize":
                        var charPrSize = FindCharPr("0");
                        if (charPrSize != null) { ApplyCharPrProperty(charPrSize, "fontsize", value); docHandled = true; }
                        break;
                    case "title" or "doctitle":
                        SetMetadata("title", value); docHandled = true; break;
                    case "creator" or "author":
                        SetMetadata("creator", value); docHandled = true; break;
                    case "subject":
                        SetMetadata("subject", value); docHandled = true; break;
                    case "description":
                        SetMetadata("description", value); docHandled = true; break;
                    case "language":
                        SetMetadata("language", value); docHandled = true; break;
                    case "keyword" or "keywords":
                        SetMetadata("keyword", value); docHandled = true; break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            if (docHandled) { _dirty = true; SaveHeader(); }
            return unsupported;
        }

        // Form field editing: /formfield[id] or /clickhere[id]
        if (path.StartsWith("/formfield[", StringComparison.OrdinalIgnoreCase)
            || path.StartsWith("/clickhere[", StringComparison.OrdinalIgnoreCase))
        {
            var fieldId = path[(path.IndexOf('[') + 1)..].TrimEnd(']');
            return SetFormFieldValue(fieldId, properties);
        }

        // Style editing: /header/style[N] path — handle before generic resolution
        if (path.StartsWith("/header/style", StringComparison.OrdinalIgnoreCase))
        {
            var style = ResolvePath(path);
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "name":
                        style.SetAttributeValue("name", value);
                        break;
                    case "engname":
                        style.SetAttributeValue("engName", value);
                        break;
                    case "font" or "fontfamily" or "fonthangul":
                        var sCharPrIdRef = style.Attribute("charPrIDRef")?.Value;
                        if (sCharPrIdRef != null)
                        {
                            var sCharPr = FindCharPr(sCharPrIdRef);
                            if (sCharPr != null)
                                ApplyCharPrProperty(sCharPr, "fonthangul", value);
                        }
                        break;
                    case "fontlatin":
                        var sCharPrIdRef2 = style.Attribute("charPrIDRef")?.Value;
                        if (sCharPrIdRef2 != null)
                        {
                            var sCharPr2 = FindCharPr(sCharPrIdRef2);
                            if (sCharPr2 != null)
                                ApplyCharPrProperty(sCharPr2, "fontlatin", value);
                        }
                        break;
                    case "size" or "fontsize":
                        var sCharPrIdRef3 = style.Attribute("charPrIDRef")?.Value;
                        if (sCharPrIdRef3 != null)
                        {
                            var sCharPr3 = FindCharPr(sCharPrIdRef3);
                            if (sCharPr3 != null)
                                ApplyCharPrProperty(sCharPr3, "fontsize", value);
                        }
                        break;
                    case "bold" or "italic":
                        var sCharPrId4 = style.Attribute("charPrIDRef")?.Value;
                        if (sCharPrId4 != null)
                        {
                            var sCharPr4 = FindCharPr(sCharPrId4);
                            if (sCharPr4 != null)
                                ApplyCharPrProperty(sCharPr4, key.ToLowerInvariant(), value);
                        }
                        break;
                    case "color":
                        var sCharPrId5 = style.Attribute("charPrIDRef")?.Value;
                        if (sCharPrId5 != null)
                        {
                            var sCharPr5 = FindCharPr(sCharPrId5);
                            if (sCharPr5 != null)
                                ApplyCharPrProperty(sCharPr5, "textcolor", value);
                        }
                        break;
                    case "alignment" or "align":
                        var sParaPrId = style.Attribute("paraPrIDRef")?.Value;
                        if (sParaPrId != null)
                        {
                            var sParaPr = _doc.Header!.Root!.Descendants(HwpxNs.Hh + "paraPr")
                                .FirstOrDefault(p => p.Attribute("id")?.Value == sParaPrId);
                            var alignEl = sParaPr?.Element(HwpxNs.Hh + "align");
                            if (alignEl == null && sParaPr != null)
                            {
                                alignEl = new XElement(HwpxNs.Hh + "align");
                                sParaPr.Add(alignEl);
                            }
                            alignEl?.SetAttributeValue("horizontal", value.ToUpperInvariant());
                        }
                        break;
                    default:
                        unsupported.Add(key);
                        break;
                }
            }
            _dirty = true;
            SaveHeader();
            return unsupported;
        }

        var element = ResolvePath(path);

        foreach (var (key, value) in properties)
        {
            switch (element.Name.LocalName)
            {
                case "p":
                    if (!SetParagraphProp(element, key, value))
                        unsupported.Add(key);
                    break;
                case "run":
                    if (!SetRunProp(element, key, value))
                        unsupported.Add(key);
                    break;
                case "t":
                    if (key.Equals("text", StringComparison.OrdinalIgnoreCase))
                        SetTextProp(element, value);
                    else
                        unsupported.Add(key);  // Don't silently coerce unsupported keys to text
                    break;
                case "tc":
                    if (!SetCellProp(element, key, value))
                        unsupported.Add(key);
                    break;
                case "tr":
                    if (!SetRowProp(element, key, value))
                        unsupported.Add(key);
                    break;
                case "tbl":
                    if (!SetTableProp(element, key, value))
                        unsupported.Add(key);
                    break;
                case "sec":
                    if (!SetSectionProp(element, key, value))
                        unsupported.Add(key);
                    break;
                case "line" or "rect" or "ellipse" or "polygon" or "pic" or "connectLine":
                    if (!SetShapeProp(element, key, value))
                        unsupported.Add(key);
                    break;
                default:
                    SetGenericAttr(element, key, value);
                    break;
            }
        }

        _dirty = true;
        // Save to correct part: header elements live in header.xml, not a section
        if (element.Document?.Root == _doc.Header?.Root)
            SaveHeader();
        else
            SaveSection(element);
        return unsupported;
    }

    // ==================== Text ====================

    /// <summary>
    /// Remove stale linesegarray from the nearest ancestor <c>&lt;hp:p&gt;</c>.
    /// HWPX linesegarray is a layout cache generated by Hancom's renderer.
    /// When text content changes, the cache becomes stale and causes Hancom to
    /// render text on a single compressed line (overlapping characters).
    /// Removing it forces Hancom to recalculate layout on open.
    /// </summary>
    private static void InvalidateLinesegarray(XElement element)
    {
        var para = element.Name.LocalName == "p"
            ? element
            : element.AncestorsAndSelf().FirstOrDefault(e => e.Name.LocalName == "p");
        para?.Elements(HwpxNs.Hp + "linesegarray").Remove();
    }

    /// <summary>
    /// Replace the text content of an &lt;hp:t&gt; element.
    /// </summary>
    private void SetTextProp(XElement tElement, string value)
    {
        InvalidateLinesegarray(tElement);
        tElement.Value = value;
    }

    // ==================== Table ====================

    /// <summary>
    /// Dispatch table property by name.
    /// </summary>
    private bool SetTableProp(XElement tbl, string property, string value)
    {
        var lower = property.ToLowerInvariant();
        if (lower.StartsWith("colwidth"))
            return SetIndividualColWidth(tbl, lower, value);
        return lower switch
        {
            "borderfillid" or "borderfillidref" => SetAttribute(tbl, "borderFillIDRef", value),
            "cellspacing" => SetAttribute(tbl, "cellSpacing", value),
            "align" or "tablealign" => SetTableAlignment(tbl, value),
            _ => false
        };
    }

    // ==================== Table Row ====================

    private bool SetRowProp(XElement tr, string property, string value)
    {
        return property.ToLowerInvariant() switch
        {
            "height" or "rowheight" => SetRowHeight(tr, value),
            _ => false
        };
    }

    private static bool SetRowHeight(XElement tr, string value)
    {
        if (!int.TryParse(value, out var h)) return false;
        foreach (var tc in tr.Elements(HwpxNs.Hp + "tc"))
        {
            var cellSz = tc.Element(HwpxNs.Hp + "cellSz");
            cellSz?.SetAttributeValue("height", h.ToString());
        }
        return true;
    }

    private bool SetTableAlignment(XElement tbl, string value)
    {
        var parentP = tbl.Ancestors(HwpxNs.Hp + "p").FirstOrDefault();
        if (parentP != null)
            return SetParagraphProp(parentP, "align", value) == true;
        return false;
    }

    private static bool SetIndividualColWidth(XElement tbl, string propName, string value)
    {
        var indexStr = propName.Replace("colwidth", "");
        if (!int.TryParse(indexStr, out var colIdx)) return false;
        colIdx--;
        var colSzElements = tbl.Elements(HwpxNs.Hp + "colSz").ToList();
        if (colIdx < 0 || colIdx >= colSzElements.Count) return false;
        colSzElements[colIdx].SetAttributeValue("width", value);
        return true;
    }

    // ==================== Table Cell ====================

    /// <summary>
    /// Dispatch table cell property by name.
    /// Supports: text, colspan, rowspan, borderfillid.
    /// </summary>
    private bool SetCellProp(XElement tc, string property, string value)
    {
        return property.ToLowerInvariant() switch
        {
            "text" => SetCellText(tc, value),
            "colspan" => SetCellSpan(tc, "colSpan", value),
            "rowspan" => SetCellSpan(tc, "rowSpan", value),
            "borderfillid" or "borderfillidref" => SetAttribute(tc, "borderFillIDRef", value),
            "valign" or "verticalalign" or "vertical-align" => SetCellVertAlign(tc, value),
            "align" or "halign" or "textalign" => SetCellHorzAlign(tc, value),
            "shading" or "bgcolor" or "fillcolor" => SetCellShading(tc, value),
            "bordercolor" => SetCellBorder(tc, color: value),
            "borderwidth" => SetCellBorder(tc, width: value),
            "bordertype" or "borderstyle" => SetCellBorder(tc, type: value),
            _ => false
        };
    }

    /// <summary>
    /// Set text content of a table cell by navigating tc → subList → p → run → t.
    /// </summary>
    private static bool SetCellVertAlign(XElement tc, string value)
    {
        var subList = tc.Element(HwpxNs.Hp + "subList");
        if (subList == null) return false;
        subList.SetAttributeValue("vertAlign", value.ToUpperInvariant());
        return true;
    }

    /// <summary>Set horizontal text alignment inside a cell by modifying the cell's paragraph paraPr.</summary>
    private bool SetCellHorzAlign(XElement tc, string value)
    {
        var subList = tc.Element(HwpxNs.Hp + "subList");
        var para = subList?.Element(HwpxNs.Hp + "p") ?? tc.Element(HwpxNs.Hp + "p");
        if (para == null) return false;
        return SetParagraphProp(para, "align", value) == true;
    }

    private bool SetCellText(XElement tc, string text)
    {
        var subList = tc.Element(HwpxNs.Hp + "subList");
        if (subList == null) return false;

        var paragraphs = subList.Elements(HwpxNs.Hp + "p").ToList();
        if (paragraphs.Count == 0) return false;

        // Set text on the first paragraph
        var result = SetParagraphText(paragraphs[0], text);

        // Remove ALL remaining paragraphs (guide text, placeholders, etc.)
        // This ensures template guide text like "※ 내용을 입력하세요" is cleared.
        foreach (var extra in paragraphs.Skip(1))
            extra.Remove();

        return result;
    }

    /// <summary>
    /// Set rowSpan or colSpan on a cell. Prefers the separate &lt;hp:cellSpan&gt; element
    /// (Hancom native format); falls back to cellAddr attributes for legacy documents.
    /// </summary>
    private static bool SetCellSpan(XElement tc, string spanAttr, string value)
    {
        if (!int.TryParse(value, out var spanVal) || spanVal < 1)
            return false;

        // Prefer separate <hp:cellSpan> element (Hancom native format)
        var cellSpan = tc.Element(HwpxNs.Hp + "cellSpan");
        if (cellSpan != null)
        {
            cellSpan.SetAttributeValue(spanAttr, spanVal.ToString());
            return true;
        }

        // Fallback: create cellSpan element if cellAddr exists
        var cellAddr = tc.Element(HwpxNs.Hp + "cellAddr");
        if (cellAddr == null) return false;

        // Check if span was on cellAddr (legacy)
        if (cellAddr.Attribute(spanAttr) != null)
        {
            cellAddr.SetAttributeValue(spanAttr, spanVal.ToString());
            return true;
        }

        // Create new cellSpan element after cellAddr
        var newCellSpan = new XElement(HwpxNs.Hp + "cellSpan",
            new XAttribute("colSpan", spanAttr == "colSpan" ? spanVal.ToString() : "1"),
            new XAttribute("rowSpan", spanAttr == "rowSpan" ? spanVal.ToString() : "1"));
        cellAddr.AddAfterSelf(newCellSpan);
        return true;
    }

    // ==================== Section ====================

    /// <summary>
    /// Dispatch section-level property by name.
    /// Section properties live in secPr (child of section root).
    /// </summary>
    private bool SetSectionProp(XElement sectionRoot, string property, string value)
    {
        return property.ToLowerInvariant() switch
        {
            "pagebackground" or "pagebg" or "backgroundcolor" => SetPageBackground(sectionRoot, value),
            "orientation" => SetOrientation(sectionRoot, value),
            "pagewidth" => SetPageDimension(sectionRoot, "width", value),
            "pageheight" => SetPageDimension(sectionRoot, "height", value),
            "margintop" or "margin-top" => SetPageMargin(sectionRoot, "top", value),
            "marginbottom" or "margin-bottom" => SetPageMargin(sectionRoot, "bottom", value),
            "marginleft" or "margin-left" => SetPageMargin(sectionRoot, "left", value),
            "marginright" or "margin-right" => SetPageMargin(sectionRoot, "right", value),
            _ => false
        };
    }

    private bool SetOrientation(XElement sectionRoot, string value)
    {
        var pagePr = sectionRoot.Descendants(HwpxNs.Hp + "pagePr").FirstOrDefault();
        if (pagePr == null) return false;
        // Hancom: NARROWLY = landscape, WIDELY = portrait. Dimensions DON'T change.
        var isLandscape = value.Equals("LANDSCAPE", StringComparison.OrdinalIgnoreCase)
            || value.Equals("NARROWLY", StringComparison.OrdinalIgnoreCase);
        pagePr.SetAttributeValue("landscape", isLandscape ? "NARROWLY" : "WIDELY");
        return true;
    }

    private static bool SetPageDimension(XElement sectionRoot, string attr, string value)
    {
        var pagePr = sectionRoot.Descendants(HwpxNs.Hp + "pagePr").FirstOrDefault();
        pagePr?.SetAttributeValue(attr, value);
        return pagePr != null;
    }

    private static bool SetPageMargin(XElement sectionRoot, string side, string value)
    {
        var margin = sectionRoot.Descendants(HwpxNs.Hp + "margin")
            .FirstOrDefault(m => m.Parent?.Name == HwpxNs.Hp + "pagePr");
        margin?.SetAttributeValue(side, value);
        return margin != null;
    }

    /// <summary>
    /// Set page background color via secPr > pageBorderFill.
    /// Creates a borderFill with no borders and the specified fill color.
    /// </summary>
    private bool SetPageBackground(XElement sectionRoot, string color)
    {
        var bfId = CreateCustomBorderFill(
            borderColor: "#000000", borderWidth: "0.00mm", borderType: "NONE",
            fillColor: color);

        var secPr = sectionRoot.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
        if (secPr == null) return false;

        var pageBf = secPr.Element(HwpxNs.Hp + "pageBorderFill");
        if (pageBf == null)
        {
            pageBf = new XElement(HwpxNs.Hp + "pageBorderFill",
                new XAttribute("type", "BOTH"),
                new XAttribute("borderFillIDRef", bfId));
            secPr.Add(pageBf);
        }
        else
        {
            pageBf.SetAttributeValue("borderFillIDRef", bfId);
        }
        return true;
    }

    // ==================== Paragraph ====================

    /// <summary>
    /// Dispatch paragraph property by name.
    /// Returns true if the property was recognized and applied.
    /// </summary>
    private bool SetParagraphProp(XElement p, string property, string value)
    {
        var lower = property.ToLowerInvariant();
        var result = lower switch
        {
            "text" => SetParagraphText(p, value),
            "style" or "styleidref" => SetAttribute(p, "styleIDRef", value),
            "align" or "alignment" => SetParagraphAlignment(p, value),
            "indent" or "leftindent" => SetParagraphIndent(p, value, "left"),
            "rightindent" => SetParagraphIndent(p, value, "right"),
            "parapridref" => SetAttribute(p, "paraPrIDRef", value),
            "spacebefore" or "spacingbefore" => SetParaPrSpacing(p, "before", value),
            "spaceafter" or "spacingafter" => SetParaPrSpacing(p, "after", value),
            "linespacing" or "lineheight" => SetParaPrSpacing(p, "lineSpacing", value),
            "linespacingtype" => SetParaPrSpacing(p, "lineSpacingType", value),
            "outlinelevel" or "heading" => SetParaPrHeadingLevel(p, value),
            "liststyle" or "list" or "bullet" => SetListStyle(p, value),
            "keepnext" or "keepwithnext" => SetBreakSetting(p, "keepWithNext", value),
            "keeplines" => SetBreakSetting(p, "keepLines", value),
            "pagebreakbefore" => SetBreakSetting(p, "pageBreakBefore", value),
            "widowcontrol" or "widoworphan" => SetBreakSetting(p, "widowOrphan", value),
            "hangingindent" or "hanging" => SetParagraphHangingIndent(p, value),
            "style" or "stylename" => SetStyleByName(p, value),
            _ => (bool?)null // not a paragraph-level prop
        };
        if (result.HasValue) return result.Value;

        // Delegate run-level properties (bold, italic, superscript, highlight, etc.)
        // to ALL runs inside the paragraph
        var runs = p.Elements(HwpxNs.Hp + "run").ToList();
        if (runs.Count == 0) return false;
        bool any = false;
        foreach (var run in runs)
        {
            if (SetRunProp(run, property, value))
                any = true;
        }
        return any;
    }

    /// <summary>
    /// Clear existing runs and set new text in a single run.
    /// </summary>
    private bool SetParagraphText(XElement para, string text)
    {
        // Preserve first run's charPrIDRef if available
        var existingRun = para.Elements(HwpxNs.Hp + "run").FirstOrDefault();
        var charPrIdRef = existingRun?.Attribute("charPrIDRef")?.Value ?? "0";

        // Remove stale linesegarray (layout cache) — see InvalidateLinesegarray().
        InvalidateLinesegarray(para);

        // CRITICAL: preserve runs that contain secPr, ctrl, or other structural elements
        // Only remove runs that are purely text-bearing
        var runs = para.Elements(HwpxNs.Hp + "run").ToList();
        var structuralRuns = runs.Where(r =>
            r.Elements(HwpxNs.Hp + "secPr").Any() ||
            r.Elements(HwpxNs.Hp + "ctrl").Any()).ToList();
        var textRuns = runs.Except(structuralRuns).ToList();

        // Remove only text runs, strip text from structural runs
        foreach (var tr in textRuns)
            tr.Remove();
        foreach (var sr in structuralRuns)
            sr.Elements(HwpxNs.Hp + "t").Remove();

        // Add new text run
        var run = new XElement(HwpxNs.Hp + "run",
            new XAttribute("charPrIDRef", charPrIdRef),
            new XElement(HwpxNs.Hp + "t", text));
        // Insert text run before structural runs so text appears first
        var firstStructural = para.Elements(HwpxNs.Hp + "run").FirstOrDefault();
        if (firstStructural != null)
            firstStructural.AddBeforeSelf(run);
        else
            para.Add(run);
        return true;
    }

    /// <summary>
    /// Set paragraph alignment via header.xml paraPr.
    /// Alignment values: "left", "center", "right", "justify", "distribute".
    /// Real HWPX stores alignment as a CHILD ELEMENT: &lt;hh:align horizontal="LEFT" vertical="BASELINE"/&gt;
    /// Values are UPPERCASE: LEFT, CENTER, RIGHT, JUSTIFY, DISTRIBUTE.
    /// </summary>
    private bool SetParagraphAlignment(XElement para, string alignment)
    {
        if (_doc.Header?.Root == null)
            return false;

        // HWPX uses uppercase alignment values
        var normalizedAlign = alignment.ToLowerInvariant() switch
        {
            "left" or "l" => "LEFT",
            "center" or "c" => "CENTER",
            "right" or "r" => "RIGHT",
            "justify" or "j" => "JUSTIFY",
            "distribute" or "d" => "DISTRIBUTE",
            _ => alignment.ToUpperInvariant()
        };

        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null)
            return false;

        // Alignment is a child element <hh:align horizontal="..." vertical="..."/>
        var alignEl = paraPr.Element(HwpxNs.Hh + "align");
        if (alignEl == null)
        {
            alignEl = new XElement(HwpxNs.Hh + "align",
                new XAttribute("horizontal", normalizedAlign),
                new XAttribute("vertical", "BASELINE"));
            paraPr.AddFirst(alignEl);
        }
        else
        {
            alignEl.SetAttributeValue("horizontal", normalizedAlign);
        }

        SaveHeader();
        return true;
    }

    /// <summary>
    /// Set paragraph indentation via header.xml paraPr.
    /// Units are in HWPUNIT (1 HWPUNIT ≈ 1/7200 inch; 1000 ≈ 10pt at 7200 DPI).
    /// </summary>
    private bool SetParagraphIndent(XElement para, string value, string side)
    {
        if (_doc.Header?.Root == null)
            return false;

        if (!int.TryParse(value, out var indentValue))
            return false;

        // Map user-facing side names to HWPX element local names
        var elementName = side.ToLowerInvariant() switch
        {
            "left" => "left",
            "right" => "right",
            "indent" or "intent" => "intent",   // first-line indent
            "before" or "prev" => "prev",        // space before paragraph
            "after" or "next" => "next",          // space after paragraph
            _ => side
        };

        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null)
            return false;

        // Find <hh:margin>. If inside <hp:switch>/<hp:default>, target the default.
        var margin = paraPr.Element(HwpxNs.Hh + "margin")
            ?? paraPr.Descendants(HwpxNs.Hh + "margin")
                .FirstOrDefault(m => m.Parent?.Name.LocalName == "default");
        if (margin == null)
        {
            margin = new XElement(HwpxNs.Hh + "margin");
            paraPr.Add(margin);
        }

        // Margin values are child elements: <hc:left value="3000" unit="HWPUNIT"/>
        var child = margin.Element(HwpxNs.Hc + elementName);
        if (child == null)
        {
            child = new XElement(HwpxNs.Hc + elementName,
                new XAttribute("value", indentValue.ToString()),
                new XAttribute("unit", "HWPUNIT"));
            margin.Add(child);
        }
        else
        {
            child.SetAttributeValue("value", indentValue.ToString());
        }

        SaveHeader();
        return true;
    }

    // ==================== Run ====================

    /// <summary>
    /// Dispatch run property by name.
    /// Run properties are stored on the charPr in header.xml.
    /// </summary>
    private bool SetRunProp(XElement run, string property, string value)
    {
        return property.ToLowerInvariant() switch
        {
            "text" => SetRunText(run, value),
            "charpridref" => SetAttribute(run, "charPrIDRef", value),
            "bold" or "italic" or "underline" or "strikeout"
                or "fontsize" or "textcolor" or "color"
                or "fonthangul" or "fontlatin"
                or "superscript" or "subscript"
                or "charspacing" or "letterspacing" or "spacing"
                or "shadecolor" or "shading"
                => EnsureCharPrProp(run, property.ToLowerInvariant(), value),
            "highlight" or "markpen" => SetHighlight(run, value),
            _ => false
        };
    }

    /// <summary>
    /// Replace text content of all &lt;hp:t&gt; children in a run.
    /// </summary>
    private bool SetRunText(XElement run, string text)
    {
        InvalidateLinesegarray(run);
        var tElements = run.Elements(HwpxNs.Hp + "t").ToList();
        if (tElements.Count == 0)
        {
            run.Add(new XElement(HwpxNs.Hp + "t", text));
        }
        else
        {
            // Set text on first <t>, remove the rest
            tElements[0].Value = text;
            foreach (var extra in tElements.Skip(1))
                extra.Remove();
        }
        return true;
    }

    // ==================== Paragraph Spacing ====================

    /// <summary>
    /// Set a spacing attribute on the paragraph's paraPr in header.xml.
    /// Spacing is stored as attributes on &lt;hh:spacing&gt; element (not child elements).
    /// attrName: "before", "after", "lineSpacing", "lineSpacingType".
    /// lineSpacingType: PERCENT, FIXED, BETWEEN_LINES.
    /// </summary>
    private bool SetParaPrSpacing(XElement para, string attrName, string value)
    {
        if (_doc.Header?.Root == null)
            return false;

        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null)
            return false;

        // Remove old-style <hh:spacing> element (incorrect structure from prior implementation)
        paraPr.Element(HwpxNs.Hh + "spacing")?.Remove();

        // HWPX spacing uses <hp:switch> > <hp:case> / <hp:default> blocks
        // containing <hh:margin> (with <hc:prev>/<hc:next>) and <hh:lineSpacing>
        var hpSwitch = paraPr.Element(HwpxNs.Hp + "switch");
        if (hpSwitch == null)
        {
            hpSwitch = new XElement(HwpxNs.Hp + "switch");
            var border = paraPr.Element(HwpxNs.Hh + "border");
            if (border != null)
                border.AddBeforeSelf(hpSwitch);
            else
                paraPr.Add(hpSwitch);
        }

        var hpCase = hpSwitch.Element(HwpxNs.Hp + "case");
        if (hpCase == null)
        {
            hpCase = new XElement(HwpxNs.Hp + "case",
                new XAttribute(HwpxNs.Hp + "required-namespace",
                    "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar"));
            hpSwitch.AddFirst(hpCase);
        }

        var hpDefault = hpSwitch.Element(HwpxNs.Hp + "default");
        if (hpDefault == null)
        {
            hpDefault = new XElement(HwpxNs.Hp + "default");
            hpSwitch.Add(hpDefault);
        }

        if (attrName == "lineSpacing")
        {
            // lineSpacing value is same in both case and default
            SetLineSpacingInBlock(hpCase, "value", value);
            SetLineSpacingInBlock(hpDefault, "value", value);
        }
        else if (attrName == "lineSpacingType")
        {
            SetLineSpacingInBlock(hpCase, "type", value);
            SetLineSpacingInBlock(hpDefault, "type", value);
        }
        else
        {
            // before → prev, after → next
            var marginChild = attrName == "before" ? "prev" : "next";
            if (!int.TryParse(value, out var caseVal))
                return false;

            var defaultVal = caseVal * 2; // default block = 2× case value
            SetMarginChild(hpCase, marginChild, caseVal.ToString());
            SetMarginChild(hpDefault, marginChild, defaultVal.ToString());
        }

        SaveHeader();
        return true;
    }

    /// <summary>Set a child element value inside &lt;hh:margin&gt; within a switch block.</summary>
    private static void SetMarginChild(XElement switchBlock, string childName, string value)
    {
        var margin = switchBlock.Element(HwpxNs.Hh + "margin");
        if (margin == null)
        {
            margin = new XElement(HwpxNs.Hh + "margin");
            switchBlock.AddFirst(margin);
        }

        var child = margin.Element(HwpxNs.Hc + childName);
        if (child == null)
        {
            child = new XElement(HwpxNs.Hc + childName,
                new XAttribute("value", value),
                new XAttribute("unit", "HWPUNIT"));
            margin.Add(child);
        }
        else
        {
            child.SetAttributeValue("value", value);
        }
    }

    /// <summary>Set lineSpacing attribute inside a switch block.</summary>
    private static void SetLineSpacingInBlock(XElement switchBlock, string attrName, string value)
    {
        var ls = switchBlock.Element(HwpxNs.Hh + "lineSpacing");
        if (ls == null)
        {
            ls = new XElement(HwpxNs.Hh + "lineSpacing",
                new XAttribute("type", "PERCENT"),
                new XAttribute("value", "160"),
                new XAttribute("unit", "HWPUNIT"));
            switchBlock.Add(ls);
        }
        ls.SetAttributeValue(attrName, value);
    }

    // ==================== Paragraph Heading / Outline Level ====================

    /// <summary>
    /// Set the outline/heading level on a paragraph's paraPr.
    /// Value "0" or "none" removes the heading. Values 1-9 set the heading level.
    /// </summary>
    private bool SetParaPrHeadingLevel(XElement para, string value)
    {
        if (_doc.Header?.Root == null)
            return false;

        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null)
            return false;

        var heading = paraPr.Element(HwpxNs.Hh + "heading");

        if (value == "0" || value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            heading?.Remove();
        }
        else
        {
            if (heading == null)
            {
                heading = new XElement(HwpxNs.Hh + "heading",
                    new XAttribute("type", "OUTLINE"),
                    new XAttribute("idRef", "0"),
                    new XAttribute("level", value));
                paraPr.Add(heading);
            }
            else
            {
                heading.SetAttributeValue("level", value);
                heading.SetAttributeValue("type", "OUTLINE");
            }
        }

        SaveHeader();
        return true;
    }

    // ==================== Form Field ====================

    private List<string> SetFormFieldValue(string idOrIndex, Dictionary<string, string> props)
    {
        var unsupported = new List<string>();
        var value = props.GetValueOrDefault("value") ?? props.GetValueOrDefault("text") ?? "";

        foreach (var sec in _doc.Sections)
        {
            foreach (var run in sec.Root.Descendants(HwpxNs.Hp + "run").ToList())
            {
                var ctrl = run.Element(HwpxNs.Hp + "ctrl");
                var fieldBegin = ctrl?.Element(HwpxNs.Hp + "fieldBegin");
                if (fieldBegin?.Attribute("type")?.Value != "CLICK_HERE") continue;
                var instId = fieldBegin.Attribute("id")?.Value;
                if (instId != idOrIndex) continue;

                // Update display text in the next run's <hp:t>
                var nextRun = run.ElementsAfterSelf(HwpxNs.Hp + "run").FirstOrDefault();
                var t = nextRun?.Element(HwpxNs.Hp + "t");
                if (t != null)
                {
                    InvalidateLinesegarray(t);
                    t.Value = value;
                    fieldBegin.SetAttributeValue("dirty", "1");
                    _dirty = true;
                    SaveSection(sec.Root);
                }
                return unsupported;
            }
        }
        unsupported.Add($"formfield [{idOrIndex}] not found");
        return unsupported;
    }

    // ==================== Shape Properties ====================

    private static bool SetShapeProp(XElement shape, string property, string value)
    {
        return property.ToLowerInvariant() switch
        {
            "wrap" or "textwrap" => SetShapeWrap(shape, value),
            "width" => SetShapeDimension(shape, "width", value),
            "height" => SetShapeDimension(shape, "height", value),
            _ => false
        };
    }

    /// <summary>Set text wrap mode. "char"/"inline" = 글자처럼 취급 (treatAsChar=1).</summary>
    private static bool SetShapeWrap(XElement shape, string value)
    {
        var isInline = value.Equals("char", StringComparison.OrdinalIgnoreCase)
            || value.Equals("inline", StringComparison.OrdinalIgnoreCase);
        var wrapValue = value.ToUpperInvariant() switch
        {
            "CHAR" or "INLINE" => "TOP_AND_BOTTOM",
            "SQUARE" => "SQUARE",
            "BEHIND" => "BEHIND_TEXT",
            "FRONT" => "IN_FRONT_OF_TEXT",
            "TIGHT" or "WRAP" => "TOP_AND_BOTTOM",
            _ => value.ToUpperInvariant()
        };
        shape.SetAttributeValue("textWrap", wrapValue);
        var pos = shape.Element(HwpxNs.Hp + "pos");
        pos?.SetAttributeValue("treatAsChar", isInline ? "1" : "0");
        return true;
    }

    private static bool SetShapeDimension(XElement shape, string attr, string value)
    {
        var sz = shape.Element(HwpxNs.Hp + "sz");
        sz?.SetAttributeValue(attr, value);
        return sz != null;
    }

    // ==================== Style by Name ====================

    private bool SetStyleByName(XElement p, string styleName)
    {
        var style = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "style")
            .FirstOrDefault(s => s.Attribute("name")?.Value == styleName
                || s.Attribute("engName")?.Value?.Equals(styleName, StringComparison.OrdinalIgnoreCase) == true);
        if (style == null) return false;
        p.SetAttributeValue("styleIDRef", style.Attribute("id")?.Value);
        return true;
    }

    // ==================== Break Settings / Indent ====================

    /// <summary>
    /// Set a breakSetting attribute on a paragraph's paraPr.
    /// XML: hh:paraPr > hh:breakSetting keepWithNext="0|1" ...
    /// </summary>
    private bool SetBreakSetting(XElement para, string attr, string value)
    {
        if (_doc.Header?.Root == null) return false;
        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null) return false;

        var bs = paraPr.Element(HwpxNs.Hh + "breakSetting");
        if (bs == null)
        {
            bs = new XElement(HwpxNs.Hh + "breakSetting",
                new XAttribute("breakLatinWord", "KEEP_WORD"),
                new XAttribute("breakNonLatinWord", "BREAK_WORD"),
                new XAttribute("widowOrphan", "0"),
                new XAttribute("keepWithNext", "0"),
                new XAttribute("keepLines", "0"),
                new XAttribute("pageBreakBefore", "0"),
                new XAttribute("lineWrap", "BREAK"));
            paraPr.Add(bs);
        }
        var boolVal = value.Equals("true", StringComparison.OrdinalIgnoreCase) || value == "1" ? "1" : "0";
        bs.SetAttributeValue(attr, boolVal);
        SaveHeader();
        return true;
    }

    /// <summary>
    /// Set hanging indent on a paragraph's paraPr margin.
    /// Hanging indent = negative indent + positive left. Value in HWPML units (283 ≈ 1mm).
    /// </summary>
    private bool SetParagraphHangingIndent(XElement para, string value)
    {
        if (!int.TryParse(value, out var hangVal) || hangVal <= 0) return false;
        if (_doc.Header?.Root == null) return false;
        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null) return false;

        // Update direct margin element
        var margin = paraPr.Element(HwpxNs.Hh + "margin");
        if (margin == null)
        {
            margin = new XElement(HwpxNs.Hh + "margin",
                new XAttribute("indent", "0"),
                new XAttribute("left", "0"), new XAttribute("right", "0"),
                new XAttribute("prev", "0"), new XAttribute("next", "0"));
            paraPr.Add(margin);
        }
        margin.SetAttributeValue("indent", (-hangVal).ToString());
        var currentLeft = (int?)margin.Attribute("left") ?? 0;
        if (currentLeft < hangVal)
            margin.SetAttributeValue("left", hangVal.ToString());

        // Also update margins inside <hp:switch> blocks (Hancom reads these)
        foreach (var switchMargin in paraPr.Descendants(HwpxNs.Hh + "margin"))
        {
            if (switchMargin == margin) continue;
            // Update child elements: <hc:intent value="N"> and <hc:left value="N">
            var intentEl = switchMargin.Element(HwpxNs.Hc + "intent");
            intentEl?.SetAttributeValue("value", (-hangVal).ToString());
            var leftEl = switchMargin.Element(HwpxNs.Hc + "left");
            leftEl?.SetAttributeValue("value", hangVal.ToString());
        }

        SaveHeader();
        return true;
    }

    // ==================== Numbering / List ====================

    /// <summary>
    /// Set list style on a paragraph. Creates a numbering definition in header.xml if needed,
    /// then links the paragraph's paraPr to it via heading element.
    /// Values: "bullet" (●), "number" or "decimal" (1. 2. 3.), "circle" (○),
    ///         "dash" (–), "none" (remove list).
    /// </summary>
    private bool SetListStyle(XElement para, string style)
    {
        if (_doc.Header?.Root == null) return false;

        var lower = style.ToLowerInvariant();
        if (lower == "none" || lower == "false" || lower == "0")
        {
            // Remove list: clear heading from paraPr
            return SetParaPrHeadingLevel(para, "0");
        }

        // Determine numbering format and text pattern
        var (format, textPattern) = lower switch
        {
            "bullet" or "disc" => ("BULLET", "●"),
            "circle" => ("BULLET", "○"),
            "dash" => ("BULLET", "–"),
            "number" or "decimal" or "numbered" => ("DIGIT", "%d."),
            "roman" => ("ROMAN_CAPITAL", "%d."),
            "romanlower" or "roman_small" => ("ROMAN_SMALL", "%d."),
            "hangul" => ("HANGUL", "%d."),
            "hanja" => ("HANJA", "%d."),
            _ => ("BULLET", "●")
        };

        // Find or create numbering definition in header.xml
        var numId = EnsureNumberingDef(format, textPattern);

        // Set paraPr heading to reference the numbering
        var paraPr = CloneParaPrIfShared(para);
        if (paraPr == null) return false;

        var heading = paraPr.Element(HwpxNs.Hh + "heading");
        if (heading == null)
        {
            heading = new XElement(HwpxNs.Hh + "heading",
                new XAttribute("type", "NUMBER"),
                new XAttribute("idRef", numId),
                new XAttribute("level", "1"));
            paraPr.Add(heading);
        }
        else
        {
            heading.SetAttributeValue("type", "NUMBER");
            heading.SetAttributeValue("idRef", numId);
            heading.SetAttributeValue("level", "1");
        }

        // Set left indent for list items (standard 800 HWPUNIT indent)
        var margin = paraPr.Element(HwpxNs.Hh + "margin");
        if (margin == null)
        {
            margin = new XElement(HwpxNs.Hh + "margin");
            paraPr.Add(margin);
        }
        var leftChild = margin.Element(HwpxNs.Hc + "left");
        if (leftChild == null)
        {
            margin.Add(new XElement(HwpxNs.Hc + "left",
                new XAttribute("value", "800"),
                new XAttribute("unit", "HWPUNIT")));
        }

        SaveHeader();
        return true;
    }

    /// <summary>
    /// Find or create a numbering definition in header.xml.
    /// Returns the numbering id string.
    /// </summary>
    private string EnsureNumberingDef(string format, string textPattern)
    {
        var header = _doc.Header!.Root!;
        var refList = header.Element(HwpxNs.Hh + "refList");

        // Find numberings container
        var numberings = refList?.Element(HwpxNs.Hh + "numberings");
        if (numberings == null)
        {
            // Create numberings container
            numberings = new XElement(HwpxNs.Hh + "numberings", new XAttribute("itemCnt", "0"));
            if (refList == null)
            {
                refList = new XElement(HwpxNs.Hh + "refList");
                header.Add(refList);
            }
            refList.Add(numberings);
        }

        // Check for existing matching numbering
        foreach (var num in numberings.Elements(HwpxNs.Hh + "numbering"))
        {
            var paraHead = num.Element(HwpxNs.Hh + "paraHead");
            if (paraHead != null)
            {
                var existingFormat = paraHead.Attribute("format")?.Value;
                var existingText = paraHead.Element(HwpxNs.Hh + "text")?.Value;
                if (existingFormat == format && existingText == textPattern)
                    return num.Attribute("id")?.Value ?? "1";
            }
        }

        // Create new numbering definition
        var maxId = numberings.Elements(HwpxNs.Hh + "numbering")
            .Select(n => int.TryParse(n.Attribute("id")?.Value, out var id) ? id : 0)
            .DefaultIfEmpty(0).Max();
        var newId = (maxId + 1).ToString();

        var newNumbering = new XElement(HwpxNs.Hh + "numbering",
            new XAttribute("id", newId),
            new XAttribute("start", "1"),
            new XElement(HwpxNs.Hh + "paraHead",
                new XAttribute("start", "1"),
                new XAttribute("level", "1"),
                new XAttribute("format", format),
                new XAttribute("alignment", "LEFT"),
                new XAttribute("useInstWidth", "1"),
                new XAttribute("autoIndent", "1"),
                new XAttribute("textOffset", "0"),
                new XAttribute("numFormat", "1"),
                new XElement(HwpxNs.Hh + "text", textPattern)));

        numberings.Add(newNumbering);
        var count = numberings.Elements(HwpxNs.Hh + "numbering").Count();
        numberings.SetAttributeValue("itemCnt", count.ToString());

        SaveHeader();
        return newId;
    }

    // ==================== Highlight (Markpen) ====================

    /// <summary>
    /// Set highlight (markpen) on a run by inserting markpenBegin/markpenEnd markers
    /// around the text content. This is NOT a charPr property — it's inline markers.
    /// Value: color hex (e.g. "#FFFF00" for yellow), "none"/"false" to remove.
    /// </summary>
    private bool SetHighlight(XElement run, string color)
    {
        var textElem = run.Element(HwpxNs.Hp + "t");
        if (textElem == null) return false;

        // Remove existing markpen markers from INSIDE <hp:t> (correct location)
        textElem.Elements(HwpxNs.Hp + "markpenBegin").ToList().ForEach(e => e.Remove());
        textElem.Elements(HwpxNs.Hp + "markpenEnd").ToList().ForEach(e => e.Remove());
        // Also clean up old-style sibling markers (wrong location from prior bug)
        run.Elements(HwpxNs.Hp + "markpenBegin").ToList().ForEach(e => e.Remove());
        run.Elements(HwpxNs.Hp + "markpenEnd").ToList().ForEach(e => e.Remove());

        var lower = color.ToLowerInvariant();
        if (lower != "none" && lower != "false" && lower != "0")
        {
            // Map common color names to hex
            var hexColor = lower switch
            {
                "yellow" => "#FFFF00",
                "green" => "#00FF00",
                "cyan" => "#00FFFF",
                "magenta" or "pink" => "#FF00FF",
                "red" => "#FF0000",
                "blue" => "#0000FF",
                _ => color // assume hex
            };

            // Golden structure: markers INSIDE <hp:t>, wrapping text content
            // <hp:t><hp:markpenBegin color="#FFFF00"/>text<hp:markpenEnd/></hp:t>
            textElem.AddFirst(
                new XElement(HwpxNs.Hp + "markpenBegin",
                    new XAttribute("color", hexColor)));
            textElem.Add(
                new XElement(HwpxNs.Hp + "markpenEnd"));
        }

        _dirty = true;
        SaveSection(run);
        return true;
    }

    // ==================== Cell Shading & Border ====================

    /// <summary>
    /// Set cell background color by creating a new borderFill with the fill color
    /// and assigning it to the cell's borderFillIDRef.
    /// </summary>
    private bool SetCellShading(XElement tc, string fillColor)
    {
        if (_doc.Header?.Root == null) return false;

        // Get current borderFill to preserve border settings
        var currentBfId = tc.Attribute("borderFillIDRef")?.Value ?? "1";
        var currentBf = _doc.Header.Root.Descendants(HwpxNs.Hh + "borderFill")
            .FirstOrDefault(e => e.Attribute("id")?.Value == currentBfId);

        // Clone existing border settings
        var borderType = currentBf?.Element(HwpxNs.Hh + "leftBorder")?.Attribute("type")?.Value ?? "SOLID";
        var borderWidth = currentBf?.Element(HwpxNs.Hh + "leftBorder")?.Attribute("width")?.Value ?? "0.12mm";
        var borderColor = currentBf?.Element(HwpxNs.Hh + "leftBorder")?.Attribute("color")?.Value ?? "#000000";

        var newBfId = CreateCustomBorderFill(borderColor, borderWidth, borderType, fillColor);
        tc.SetAttributeValue("borderFillIDRef", newBfId);
        return true;
    }

    /// <summary>
    /// Set cell border properties by creating a new borderFill and assigning it.
    /// Only the specified parameters are changed; others are preserved from the current borderFill.
    /// </summary>
    private bool SetCellBorder(XElement tc, string? color = null, string? width = null, string? type = null)
    {
        if (_doc.Header?.Root == null) return false;

        var currentBfId = tc.Attribute("borderFillIDRef")?.Value ?? "1";
        var currentBf = _doc.Header.Root.Descendants(HwpxNs.Hh + "borderFill")
            .FirstOrDefault(e => e.Attribute("id")?.Value == currentBfId);

        var borderType = type ?? currentBf?.Element(HwpxNs.Hh + "leftBorder")?.Attribute("type")?.Value ?? "SOLID";
        var borderWidth = width ?? currentBf?.Element(HwpxNs.Hh + "leftBorder")?.Attribute("width")?.Value ?? "0.12mm";
        var borderColor = color ?? currentBf?.Element(HwpxNs.Hh + "leftBorder")?.Attribute("color")?.Value ?? "#000000";

        // Check for existing fill color to preserve
        string? fillColor = null;
        var existingFill = currentBf?.Element(HwpxNs.Hc + "fillBrush")?.Element(HwpxNs.Hc + "winBrush");
        if (existingFill != null)
            fillColor = existingFill.Attribute("faceColor")?.Value;

        var newBfId = CreateCustomBorderFill(borderColor, borderWidth, borderType, fillColor);
        tc.SetAttributeValue("borderFillIDRef", newBfId);
        return true;
    }

    /// <summary>
    /// Create a custom borderFill in header.xml with specified border and optional fill settings.
    /// Returns the new borderFill ID.
    /// </summary>
    private string CreateCustomBorderFill(
        string borderColor = "#000000",
        string borderWidth = "0.12mm",
        string borderType = "SOLID",
        string? fillColor = null)
    {
        var borderFills = _doc.Header!.Root!.Descendants(HwpxNs.Hh + "borderFill");
        var newId = NextBorderFillId();

        var bf = new XElement(HwpxNs.Hh + "borderFill",
            new XAttribute("id", newId),
            new XAttribute("threeD", "0"),
            new XAttribute("shadow", "0"),
            new XAttribute("centerLine", "NONE"),
            new XAttribute("breakCellSeparateLine", "0"),
            new XElement(HwpxNs.Hh + "slash",
                new XAttribute("type", "NONE"), new XAttribute("crooked", "0"), new XAttribute("isCounter", "0")),
            new XElement(HwpxNs.Hh + "backSlash",
                new XAttribute("type", "NONE"), new XAttribute("crooked", "0"), new XAttribute("isCounter", "0")),
            MakeBorder("leftBorder", borderType, borderWidth, borderColor),
            MakeBorder("rightBorder", borderType, borderWidth, borderColor),
            MakeBorder("topBorder", borderType, borderWidth, borderColor),
            MakeBorder("bottomBorder", borderType, borderWidth, borderColor),
            MakeBorder("diagonal", "NONE", "0.00mm", "#000000"));

        if (fillColor != null)
        {
            bf.Add(new XElement(HwpxNs.Hc + "fillBrush",
                new XElement(HwpxNs.Hc + "winBrush",
                    new XAttribute("faceColor", fillColor),
                    new XAttribute("hatchColor", "#FFFFFF"),
                    new XAttribute("alpha", "0"))));
        }

        // Add to borderFills container
        var container = _doc.Header!.Root!.Descendants(HwpxNs.Hh + "borderFills").FirstOrDefault();
        if (container != null)
        {
            container.Add(bf);
            var count = container.Elements(HwpxNs.Hh + "borderFill").Count();
            container.SetAttributeValue("itemCnt", count.ToString());
        }
        else if (borderFills.Any())
        {
            borderFills.Last().AddAfterSelf(bf);
        }

        SaveHeader();
        return newId;
    }

    // ==================== CharPr Clone-or-Modify ====================

    /// <summary>
    /// CRITICAL: Set a character property on a run's charPr in header.xml.
    ///
    /// Algorithm:
    /// 1. Get current charPrIDRef from the run.
    /// 2. Find &lt;hh:charPr id="N"&gt; in header.xml.
    /// 3. Scan ALL sections to check if this charPr is referenced by ANY other run.
    ///    → If yes: CLONE the charPr (assign NextCharPrId()), update run's charPrIDRef.
    ///    → If no: modify the charPr in place.
    /// 4. Set the requested property on the (possibly cloned) charPr.
    /// </summary>
    private bool EnsureCharPrProp(XElement run, string prop, string value)
    {
        if (_doc.Header?.Root == null)
            return false;

        var charPrIdRef = run.Attribute("charPrIDRef")?.Value;
        if (charPrIdRef == null)
            return false;

        // Find the charPr in header.xml
        var charPr = _doc.Header.Root.Descendants(HwpxNs.Hh + "charPr")
            .FirstOrDefault(cp => cp.Attribute("id")?.Value == charPrIdRef);
        if (charPr == null)
            return false;

        // Count how many runs across ALL sections reference this charPr
        int refCount = 0;
        foreach (var section in _doc.Sections)
        {
            foreach (var r in section.Root.Descendants(HwpxNs.Hp + "run"))
            {
                if (r.Attribute("charPrIDRef")?.Value == charPrIdRef)
                    refCount++;
            }
        }

        // Clone if shared: either multiple runs reference this charPr,
        // or it's charPr 0 (the global default used by all new elements).
        // Without this, modifying charPr 0 via fontsize=22 on one paragraph
        // contaminates ALL paragraphs and table cells that use the default.
        if (refCount > 1 || charPrIdRef == "0")
        {
            var newId = NextCharPrId();
            var cloned = new XElement(charPr);
            cloned.SetAttributeValue("id", newId.ToString());
            // CRITICAL: Hancom uses POSITIONAL indexing (array index), not id-based lookup.
            // Append at END of container so position matches the new ID.
            var container = charPr.Parent!;
            container.Add(cloned);

            // Update this run to point to the clone
            run.SetAttributeValue("charPrIDRef", newId.ToString());
            charPr = cloned;

            // Update itemCnt on the parent <hh:charProperties> container
            var count = container.Elements(HwpxNs.Hh + "charPr").Count();
            container.SetAttributeValue("itemCnt", count.ToString());
        }

        // Apply the property to the charPr
        ApplyCharPrProperty(charPr, prop, value);
        SaveHeader();
        return true;
    }

    /// <summary>
    /// Apply a named property to a charPr element.
    /// </summary>
    private static void ApplyCharPrProperty(XElement charPr, string prop, string value)
    {
        switch (prop)
        {
            case "bold":
                ToggleCharPrFlag(charPr, HwpxNs.Hh + "bold", value);
                break;

            case "italic":
                ToggleCharPrFlag(charPr, HwpxNs.Hh + "italic", value);
                break;

            case "underline":
                ToggleCharPrFlag(charPr, HwpxNs.Hh + "underline", value);
                break;

            case "strikeout":
                ToggleCharPrFlag(charPr, HwpxNs.Hh + "strikeout", value);
                break;

            case "superscript":
                ToggleCharPrFlag(charPr, HwpxNs.Hh + "supscript", value);
                // Remove subscript if enabling superscript
                if (value.Equals("true", StringComparison.OrdinalIgnoreCase) || value == "1")
                    charPr.Element(HwpxNs.Hh + "subscript")?.Remove();
                break;

            case "subscript":
                ToggleCharPrFlag(charPr, HwpxNs.Hh + "subscript", value);
                // Remove superscript if enabling subscript
                if (value.Equals("true", StringComparison.OrdinalIgnoreCase) || value == "1")
                    charPr.Element(HwpxNs.Hh + "supscript")?.Remove();
                break;

            case "fontsize":
                // HWPX font size in centi-points (1/100 pt): 1000 = 10pt, 2000 = 20pt
                // User input is in pt — convert to centi-points
                if (double.TryParse(value, out var ptSize))
                    charPr.SetAttributeValue("height", ((int)(ptSize * 100)).ToString());
                break;

            case "textcolor" or "color":
                charPr.SetAttributeValue("textColor", value);
                break;

            case "charspacing" or "letterspacing" or "spacing":
                // Character spacing in percent per script. 0 = normal, -5 = 5% tighter.
                if (int.TryParse(value, out var spacingVal))
                {
                    var spacingEl = charPr.Element(HwpxNs.Hh + "spacing");
                    if (spacingEl == null)
                    {
                        spacingEl = new XElement(HwpxNs.Hh + "spacing",
                            new XAttribute("hangul", "0"), new XAttribute("latin", "0"),
                            new XAttribute("hanja", "0"), new XAttribute("japanese", "0"),
                            new XAttribute("other", "0"), new XAttribute("symbol", "0"),
                            new XAttribute("user", "0"));
                        charPr.Add(spacingEl);
                    }
                    spacingEl.SetAttributeValue("hangul", spacingVal.ToString());
                    spacingEl.SetAttributeValue("latin", spacingVal.ToString());
                }
                break;

            case "shadecolor" or "shading":
                charPr.SetAttributeValue("shadeColor", value);
                break;

            case "fonthangul":
                var fontRef = charPr.Element(HwpxNs.Hh + "fontRef");
                if (fontRef == null)
                {
                    fontRef = new XElement(HwpxNs.Hh + "fontRef");
                    charPr.Add(fontRef);
                }
                fontRef.SetAttributeValue("hangul", value);
                break;

            case "fontlatin":
                var fontRefLatin = charPr.Element(HwpxNs.Hh + "fontRef");
                if (fontRefLatin == null)
                {
                    fontRefLatin = new XElement(HwpxNs.Hh + "fontRef");
                    charPr.Add(fontRefLatin);
                }
                fontRefLatin.SetAttributeValue("latin", value);
                break;
        }
    }

    /// <summary>
    /// Normalize fontRef attributes to "0" (first declared font).
    /// Golden XML shows sup/subscript charPrs always use fontRef="0".
    /// </summary>
    private static void NormalizeFontRef(XElement charPr)
    {
        var fontRef = charPr.Element(HwpxNs.Hh + "fontRef");
        if (fontRef == null) return;
        foreach (var attr in fontRef.Attributes())
            attr.Value = "0";
    }

    /// <summary>
    /// Toggle a boolean charPr flag element.
    /// "true"/"1" → add element if missing; "false"/"0" → remove if present.
    /// </summary>
    private static void ToggleCharPrFlag(XElement charPr, XName flagName, string value)
    {
        var isTruthy = value.Equals("true", StringComparison.OrdinalIgnoreCase)
                    || value == "1";
        var existing = charPr.Element(flagName);

        if (isTruthy && existing == null)
        {
            charPr.Add(new XElement(flagName));
        }
        else if (!isTruthy && existing != null)
        {
            existing.Remove();
        }
    }

    // ==================== Style Clone Helpers ====================

    /// <summary>
    /// Clone charPr 0 to a new independent ID for use in new style definitions.
    /// Prevents global default contamination when editing style properties.
    /// </summary>
    private int CloneCharPrForNewStyle()
    {
        var charPr0 = _doc.Header?.Root?
            .Descendants(HwpxNs.Hh + "charPr")
            .FirstOrDefault(e => e.Attribute("id")?.Value == "0");
        if (charPr0 == null) return 0;

        var newId = NextCharPrId();
        var cloned = new XElement(charPr0);
        cloned.SetAttributeValue("id", newId.ToString());
        var container = charPr0.Parent!;
        container.Add(cloned);
        var count = container.Elements(HwpxNs.Hh + "charPr").Count();
        container.SetAttributeValue("itemCnt", count.ToString());
        return newId;
    }

    /// <summary>
    /// Clone paraPr 0 to a new independent ID for use in new style definitions.
    /// Prevents global default contamination when editing style properties.
    /// </summary>
    private int CloneParaPrForNewStyle()
    {
        var paraPr0 = _doc.Header?.Root?
            .Descendants(HwpxNs.Hh + "paraPr")
            .FirstOrDefault(e => e.Attribute("id")?.Value == "0");
        if (paraPr0 == null) return 0;

        var newId = NextParaPrId();
        var cloned = new XElement(paraPr0);
        cloned.SetAttributeValue("id", newId.ToString());
        var container = paraPr0.Parent!;
        container.Add(cloned);
        var count = container.Elements(HwpxNs.Hh + "paraPr").Count();
        container.SetAttributeValue("itemCnt", count.ToString());
        return newId;
    }

    // ==================== ID Generators ====================

    /// <summary>
    /// Return max charPrIDRef across ALL sections + header, then add 1.
    /// </summary>
    private int NextCharPrId()
    {
        int maxId = 0;

        // Scan all run elements across all sections
        foreach (var section in _doc.Sections)
        {
            foreach (var run in section.Root.Descendants(HwpxNs.Hp + "run"))
            {
                if (int.TryParse(run.Attribute("charPrIDRef")?.Value, out var id))
                    maxId = Math.Max(maxId, id);
            }
        }

        // Scan header.xml charPr definitions
        if (_doc.Header?.Root != null)
        {
            foreach (var charPr in _doc.Header.Root.Descendants(HwpxNs.Hh + "charPr"))
            {
                if (int.TryParse(charPr.Attribute("id")?.Value, out var id))
                    maxId = Math.Max(maxId, id);
            }
        }

        return maxId + 1;
    }

    /// <summary>
    /// Return max paraPrIDRef across ALL sections + header, then add 1.
    /// </summary>
    private int NextParaPrId()
    {
        int maxId = 0;

        foreach (var section in _doc.Sections)
        {
            foreach (var p in section.Root.Descendants(HwpxNs.Hp + "p"))
            {
                if (int.TryParse(p.Attribute("paraPrIDRef")?.Value, out var id))
                    maxId = Math.Max(maxId, id);
            }
        }

        if (_doc.Header?.Root != null)
        {
            foreach (var paraPr in _doc.Header.Root.Descendants(HwpxNs.Hh + "paraPr"))
            {
                if (int.TryParse(paraPr.Attribute("id")?.Value, out var id))
                    maxId = Math.Max(maxId, id);
            }
        }

        return maxId + 1;
    }

    /// <summary>
    /// Check if a paraPr is referenced by any paragraph OTHER than the given one.
    /// </summary>
    private bool IsParaPrShared(string paraPrIdRef, XElement excludeParagraph)
    {
        foreach (var section in _doc.Sections)
        {
            foreach (var p in section.Root.Descendants(HwpxNs.Hp + "p"))
            {
                if (p == excludeParagraph) continue;
                if (p.Attribute("paraPrIDRef")?.Value == paraPrIdRef)
                    return true;
            }
        }
        // Also check if any style references this paraPr
        if (_doc.Header?.Root != null)
        {
            foreach (var style in _doc.Header.Root.Descendants(HwpxNs.Hh + "style"))
            {
                if (style.Attribute("paraPrIDRef")?.Value == paraPrIdRef)
                    return true;
            }
        }
        return false;
    }

    // ==================== Generic ====================

    /// <summary>
    /// Set an XML attribute directly on the element.
    /// Fallback for element types without specialized property handling.
    /// </summary>
    private static bool SetGenericAttr(XElement element, string property, string value)
    {
        element.SetAttributeValue(property, value);
        return true;
    }

    /// <summary>Set a named attribute to a value. Always returns true.</summary>
    private static bool SetAttribute(XElement element, string name, string value)
    {
        element.SetAttributeValue(name, value);
        return true;
    }

    // ==================== Find & Replace ====================

    /// <summary>
    // ==================== Label-Based Table Fill (Plan 70) ====================

    /// <summary>
    /// Fill table cells by label matching. For each mapping, find a cell whose text
    /// matches the label, then set the adjacent cell's text to the value.
    /// </summary>
    private void FillByLabel(Dictionary<string, string> mappings)
    {
        var anyFilled = false;
        var filledSections = new HashSet<XElement>();

        foreach (var (key, value) in mappings)
        {
            var (label, direction) = ParseLabelSpec(key);
            var tc = FindCellByLabel(label, direction);
            if (tc == null) continue;

            SetCellText(tc, value);
            anyFilled = true;

            // Track which sections need saving
            var sectionRoot = tc.AncestorsAndSelf()
                .FirstOrDefault(e => e.Name.LocalName == "sec");
            if (sectionRoot != null) filledSections.Add(sectionRoot);
        }

        if (anyFilled)
        {
            foreach (var sec in filledSections)
                SaveSection(sec);
            _dirty = true;
        }
    }

    // ==================== Find & Replace ====================

    /// <summary>
    /// Replace all occurrences of <paramref name="find"/> with <paramref name="replace"/>
    /// across all sections' &lt;hp:t&gt; elements. Returns the number of replacements made.
    /// Known limitation: text split across multiple runs will not be matched.
    /// </summary>
    private int FindAndReplace(string find, string replace,
        XElement? scope = null, Dictionary<string, string>? formatFilter = null)
    {
        if (string.IsNullOrEmpty(find)) return 0;

        IEnumerable<XElement> searchRoots = scope != null
            ? new[] { scope }
            : _doc.Sections.Select(s => s.Root);

        int count = 0;
        var isRegex = find.StartsWith("regex:", StringComparison.OrdinalIgnoreCase);
        var regex = isRegex
            ? new System.Text.RegularExpressions.Regex(find[6..])
            : null;

        foreach (var root in searchRoots)
        {
            foreach (var run in root.Descendants(HwpxNs.Hp + "run").ToList())
            {
                if (formatFilter != null && !MatchesCharPrFormat(run, formatFilter))
                    continue;

                foreach (var t in run.Elements(HwpxNs.Hp + "t").ToList())
                {
                    var text = t.Value;
                    if (isRegex)
                    {
                        if (regex!.IsMatch(text))
                        {
                            InvalidateLinesegarray(t);
                            t.Value = regex.Replace(text, replace);
                            count++;
                        }
                    }
                    else if (text.Contains(find, StringComparison.Ordinal))
                    {
                        InvalidateLinesegarray(t);
                        t.Value = text.Replace(find, replace, StringComparison.Ordinal);
                        count++;
                    }
                }
            }
        }

        if (count > 0)
        {
            foreach (var sec in _doc.Sections)
                SaveSection(sec.Root);
            _dirty = true;
        }
        return count;
    }

    /// <summary>Check if a run's charPr matches the format filter.</summary>
    private bool MatchesCharPrFormat(XElement run, Dictionary<string, string> filter)
    {
        var charPrId = run.Attribute("charPrIDRef")?.Value ?? "0";
        var charPr = FindCharPr(charPrId);
        if (charPr == null) return false;

        foreach (var (key, expected) in filter)
        {
            switch (key.ToLowerInvariant())
            {
                case "bold":
                    var hasBold = charPr.Element(HwpxNs.Hh + "bold") != null;
                    if (hasBold != (expected == "true" || expected == "1")) return false;
                    break;
                case "italic":
                    var hasItalic = charPr.Element(HwpxNs.Hh + "italic") != null;
                    if (hasItalic != (expected == "true" || expected == "1")) return false;
                    break;
                case "color":
                    if (charPr.Attribute("textColor")?.Value != expected) return false;
                    break;
                case "fontsize":
                    if (charPr.Attribute("height")?.Value != expected) return false;
                    break;
            }
        }
        return true;
    }

    // ==================== Save Helpers ====================

    /// <summary>
    /// Save header.xml back to the ZIP archive.
    /// Uses delete-and-recreate pattern (avoids trailing bytes from SetLength(0)).
    /// </summary>
    private void SaveHeader()
    {
        if (_doc.Header == null || _doc.HeaderEntryPath == null) return;

        var entry = _doc.Archive.GetEntry(_doc.HeaderEntryPath);
        if (entry == null) return;

        var entryName = entry.FullName;
        entry.Delete();
        var newEntry = _doc.Archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        // CRITICAL: Hancom requires single-line (minified) XML without BOM.
        var xmlStr = HwpxPacker.MinifyXml(_doc.Header.ToString(SaveOptions.DisableFormatting));
        xmlStr = HwpxPacker.RestoreOriginalNamespaces(xmlStr);
        xmlStr = "<?xml version='1.0' encoding='UTF-8'?>" + xmlStr;
        var bytes = System.Text.Encoding.UTF8.GetBytes(xmlStr);
        stream.Write(bytes, 0, bytes.Length);
    }
}
