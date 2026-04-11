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
                => EnsureCharPrProp(run, property.ToLowerInvariant(), value),
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
