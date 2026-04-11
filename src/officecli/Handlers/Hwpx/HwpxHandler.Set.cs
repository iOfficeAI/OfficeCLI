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

        // Document-level properties (including find/replace)
        if (path is "/" or "" or "/body")
        {
            if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
            {
                var count = FindAndReplace(findText, replaceText);
                var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
                remaining.Remove("find");
                remaining.Remove("replace");
                if (remaining.Count > 0)
                    unsupported.AddRange(remaining.Keys);
                return unsupported;
            }
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
                case "tbl":
                    if (!SetTableProp(element, key, value))
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
    /// Replace the text content of an &lt;hp:t&gt; element.
    /// </summary>
    private void SetTextProp(XElement tElement, string value)
    {
        tElement.Value = value;
    }

    // ==================== Table ====================

    /// <summary>
    /// Dispatch table property by name.
    /// </summary>
    private bool SetTableProp(XElement tbl, string property, string value)
    {
        return property.ToLowerInvariant() switch
        {
            "borderfillid" or "borderfillidref" => SetAttribute(tbl, "borderFillIDRef", value),
            "cellspacing" => SetAttribute(tbl, "cellSpacing", value),
            _ => false
        };
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
    private bool SetCellText(XElement tc, string text)
    {
        var subList = tc.Element(HwpxNs.Hp + "subList");
        if (subList == null) return false;

        var para = subList.Element(HwpxNs.Hp + "p");
        if (para == null) return false;

        return SetParagraphText(para, text);
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

    // ==================== Paragraph ====================

    /// <summary>
    /// Dispatch paragraph property by name.
    /// Returns true if the property was recognized and applied.
    /// </summary>
    private bool SetParagraphProp(XElement p, string property, string value)
    {
        return property.ToLowerInvariant() switch
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
            _ => false
        };
    }

    /// <summary>
    /// Clear existing runs and set new text in a single run.
    /// </summary>
    private bool SetParagraphText(XElement para, string text)
    {
        // Preserve first run's charPrIDRef if available
        var existingRun = para.Elements(HwpxNs.Hp + "run").FirstOrDefault();
        var charPrIdRef = existingRun?.Attribute("charPrIDRef")?.Value ?? "0";

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

        var spacing = paraPr.Element(HwpxNs.Hh + "spacing");
        if (spacing == null)
        {
            spacing = new XElement(HwpxNs.Hh + "spacing");
            paraPr.Add(spacing);
        }

        spacing.SetAttributeValue(attrName, value);
        SaveHeader();
        return true;
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

        // Remove existing markpen markers
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

            textElem.AddBeforeSelf(
                new XElement(HwpxNs.Hp + "markpenBegin",
                    new XAttribute("beginColor", hexColor)));
            textElem.AddAfterSelf(
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

        // If more than one run uses this charPr, we must clone
        if (refCount > 1)
        {
            var newId = NextCharPrId();
            var cloned = new XElement(charPr);
            cloned.SetAttributeValue("id", newId.ToString());
            charPr.AddAfterSelf(cloned);

            // Update this run to point to the clone
            run.SetAttributeValue("charPrIDRef", newId.ToString());
            charPr = cloned;
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
                // HWPX font size is in hundredths of a point: 1000 = 10pt
                if (double.TryParse(value, out var ptSize))
                    charPr.SetAttributeValue("height", ((int)(ptSize * 100)).ToString());
                break;

            case "textcolor" or "color":
                charPr.SetAttributeValue("textColor", value);
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
    /// Replace all occurrences of <paramref name="find"/> with <paramref name="replace"/>
    /// across all sections' &lt;hp:t&gt; elements. Returns the number of replacements made.
    /// Known limitation: text split across multiple runs will not be matched.
    /// </summary>
    private int FindAndReplace(string find, string replace)
    {
        if (string.IsNullOrEmpty(find)) return 0;
        int totalCount = 0;

        foreach (var section in _doc.Sections)
        {
            foreach (var t in section.Document.Descendants(HwpxNs.Hp + "t"))
            {
                var text = t.Value;
                if (text.Contains(find, StringComparison.Ordinal))
                {
                    t.Value = text.Replace(find, replace, StringComparison.Ordinal);
                    totalCount++;
                }
            }
            if (totalCount > 0)
                SaveSection(section.Document.Root!);
        }

        _dirty = true;
        return totalCount;
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
