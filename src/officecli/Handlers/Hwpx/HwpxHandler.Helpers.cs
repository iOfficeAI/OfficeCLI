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
    /// Extract cell address from either:
    /// 1. Child element: &lt;hp:cellAddr colAddr="0" rowAddr="0" colSpan="1" rowSpan="1"/&gt;
    /// 2. Legacy attributes on tc: &lt;hp:tc colAddr="0" rowAddr="0" .../&gt;
    /// </summary>
    internal static (int Row, int Col, int RowSpan, int ColSpan) GetCellAddr(XElement tc)
    {
        var cellAddr = tc.Element(HwpxNs.Hp + "cellAddr");
        if (cellAddr != null)
        {
            return (
                (int?)cellAddr.Attribute("rowAddr") ?? 0,
                (int?)cellAddr.Attribute("colAddr") ?? 0,
                (int?)cellAddr.Attribute("rowSpan") ?? 1,
                (int?)cellAddr.Attribute("colSpan") ?? 1
            );
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
