// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml.Linq;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    private XElement? FindCharPr(string idRef)
    {
        return _doc.Header?.Root?
            .Descendants(HwpxNs.Hh + "charPr")
            .FirstOrDefault(e => e.Attribute("id")?.Value == idRef);
    }

    private static double GetFontSizePt(XElement charPr)
        => ((double?)charPr.Attribute("height") ?? 1000) / 100.0;

    /// <summary>
    /// Extract cell address and span from multiple possible formats:
    /// 1. Modern: &lt;hp:cellAddr colAddr rowAddr/&gt; + &lt;hp:cellSpan colSpan rowSpan/&gt; (separate elements)
    /// 2. Combined: &lt;hp:cellAddr colAddr rowAddr colSpan rowSpan/&gt; (span attrs on cellAddr)
    /// 3. Legacy: attributes directly on &lt;hp:tc&gt;
    /// </summary>
    // ==================== Common Helpers (Plan 39) ====================

    /// <summary>
    /// Wrap content in an &lt;hp:run&gt; element with the given charPrIDRef.
    /// Used by CreateHyperlink, CreateFootnote, AddHeaderFooter, etc.
    /// </summary>
    private static XElement WrapInRun(XElement content, string charPrIDRef = "0")
        => new XElement(HwpxNs.Hp + "run",
            new XAttribute("charPrIDRef", charPrIDRef),
            content);

    /// <summary>
    /// Create a standard &lt;hp:subList&gt; element containing a single paragraph with text.
    /// Used by BuildCell, CreateFootnote, AddHeaderFooter.
    /// </summary>
    private XElement CreateSubList(string text, string vertAlign = "CENTER")
        => new XElement(HwpxNs.Hp + "subList",
            new XAttribute("id", NewId()),
            new XAttribute("textDirection", "HORIZONTAL"),
            new XAttribute("lineWrap", "BREAK"),
            new XAttribute("vertAlign", vertAlign),
            new XAttribute("linkListIDRef", "0"),
            new XAttribute("linkListNextIDRef", "0"),
            new XAttribute("textWidth", "0"),
            new XAttribute("textHeight", "0"),
            new XAttribute("hasTextRef", "0"),
            new XAttribute("hasNumRef", "0"),
            CreateParagraph(new() { ["text"] = text }));

    /// <summary>
    /// If the paraPr referenced by the paragraph is shared with other paragraphs,
    /// clone it with a new ID and update the paragraph's paraPrIDRef.
    /// Returns the (possibly cloned) paraPr, or null if not found.
    /// </summary>
    private XElement? CloneParaPrIfShared(XElement para)
    {
        var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;
        if (paraPrIdRef == null) return null;

        var paraPr = _doc.Header?.Root?
            .Descendants(HwpxNs.Hh + "paraPr")
            .FirstOrDefault(e => e.Attribute("id")?.Value == paraPrIdRef);
        if (paraPr == null) return null;

        if (IsParaPrShared(paraPrIdRef, para))
        {
            var newId = NextParaPrId();
            var cloned = new XElement(paraPr);
            cloned.SetAttributeValue("id", newId.ToString());
            paraPr.AddAfterSelf(cloned);
            para.SetAttributeValue("paraPrIDRef", newId.ToString());
            paraPr = cloned;
        }

        return paraPr;
    }

    /// <summary>
    /// Return the next available borderFill ID based on max existing ID (not count).
    /// Fixes the count-based ID generation bug that could cause ID collisions.
    /// </summary>
    private string NextBorderFillId()
    {
        var borderFills = _doc.Header!.Root!.Descendants(HwpxNs.Hh + "borderFill");
        var maxId = borderFills.Any()
            ? borderFills.Max(bf => int.TryParse(bf.Attribute("id")?.Value, out var n) ? n : 0)
            : 0;
        return (maxId + 1).ToString();
    }

    /// <summary>
    /// Create a border element (leftBorder, rightBorder, topBorder, bottomBorder, diagonal).
    /// </summary>
    private static XElement MakeBorder(string name, string type, string width, string color)
        => new XElement(HwpxNs.Hh + name,
            new XAttribute("type", type),
            new XAttribute("width", width),
            new XAttribute("color", color));

    // ==================== Cell Address Helpers ====================

    internal static (int Row, int Col, int RowSpan, int ColSpan) GetCellAddr(XElement tc)
    {
        var cellAddr = tc.Element(HwpxNs.Hp + "cellAddr");
        if (cellAddr != null)
        {
            int row = (int?)cellAddr.Attribute("rowAddr") ?? 0;
            int col = (int?)cellAddr.Attribute("colAddr") ?? 0;

            // Try separate <hp:cellSpan> element first (Hancom native format)
            var cellSpan = tc.Element(HwpxNs.Hp + "cellSpan");
            if (cellSpan != null)
            {
                return (row, col,
                    (int?)cellSpan.Attribute("rowSpan") ?? 1,
                    (int?)cellSpan.Attribute("colSpan") ?? 1);
            }

            // Fallback: span attrs on cellAddr itself
            return (row, col,
                (int?)cellAddr.Attribute("rowSpan") ?? 1,
                (int?)cellAddr.Attribute("colSpan") ?? 1);
        }

        // Fallback: attributes directly on <hp:tc>
        return (
            (int?)tc.Attribute("rowAddr") ?? 0,
            (int?)tc.Attribute("colAddr") ?? 0,
            (int?)tc.Attribute("rowSpan") ?? 1,
            (int?)tc.Attribute("colSpan") ?? 1
        );
    }
}
