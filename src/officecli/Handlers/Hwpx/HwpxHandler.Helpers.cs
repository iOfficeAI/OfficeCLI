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
            // CRITICAL: Hancom uses POSITIONAL indexing (array index), not id-based lookup.
            // Append at END so position matches the new ID.
            var container = paraPr.Parent!;
            container.Add(cloned);
            para.SetAttributeValue("paraPrIDRef", newId.ToString());
            paraPr = cloned;

            // Update itemCnt on the parent <hh:paraProperties> container
            var count = container.Elements(HwpxNs.Hh + "paraPr").Count();
            container.SetAttributeValue("itemCnt", count.ToString());
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

    // ==================== Label-Based Table Fill (Plan 70) ====================

    /// <summary>
    /// Extract all text from a table cell: tc → subList → p* → run* → t*.
    /// Reuses <see cref="ExtractParagraphText"/> for consistency.
    /// </summary>
    internal static string ExtractCellText(XElement tc)
    {
        var subList = tc.Element(HwpxNs.Hp + "subList");
        var paragraphs = subList?.Elements(HwpxNs.Hp + "p")
                      ?? tc.Elements(HwpxNs.Hp + "p");

        var sb = new System.Text.StringBuilder();
        foreach (var p in paragraphs)
        {
            var text = ExtractParagraphText(p);
            if (sb.Length > 0 && !string.IsNullOrEmpty(text))
                sb.Append('\n');
            sb.Append(text);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Normalize a label for matching: trim, collapse whitespace,
    /// strip trailing colon/fullwidth colon/spaces.
    /// </summary>
    internal static string NormalizeLabel(string label)
    {
        if (string.IsNullOrEmpty(label)) return "";
        var normalized = System.Text.RegularExpressions.Regex.Replace(label.Trim(), @"\s+", " ");
        normalized = normalized.TrimEnd(':', ' ', '\t', '\u00A0', '\uFF1A'); // ASCII colon + fullwidth colon
        return normalized.Trim();
    }

    /// <summary>
    /// Parse "라벨>direction" syntax. Default direction is "right".
    /// Examples: "대표자>down" → ("대표자", "down"), "대표자" → ("대표자", "right").
    /// </summary>
    internal static (string Label, string Direction) ParseLabelSpec(string key)
    {
        var idx = key.IndexOf('>');
        if (idx > 0 && idx < key.Length - 1)
            return (key[..idx].Trim(), key[(idx + 1)..].Trim().ToLowerInvariant());
        return (key.Trim(), "right");
    }

    /// <summary>
    /// Find a table cell whose text matches <paramref name="label"/>,
    /// then return the adjacent cell in the specified <paramref name="direction"/>.
    /// Searches all tables in all sections.
    /// Handles merged cells via <see cref="GetCellAddr"/>.
    /// </summary>
    internal XElement? FindCellByLabel(string label, string direction = "right")
    {
        var normalizedLabel = NormalizeLabel(label);
        if (string.IsNullOrEmpty(normalizedLabel)) return null;

        foreach (var sec in _doc.Sections)
        {
            foreach (var tbl in sec.Tables)
            {
                var result = FindCellInTable(tbl, normalizedLabel, direction);
                if (result != null) return result;
            }
        }
        return null;
    }

    /// <summary>
    /// Build a 2D grid from table cells using cellAddr. Handles merged cells via rowSpan/colSpan.
    /// Reused by: FindCellInTable (Plan 70), RecognizeFormFields (Plan 70.2), Table Map (Plan 71).
    /// </summary>
    internal static (XElement?[,] Grid, List<(XElement Tc, int Row, int Col, int RowSpan, int ColSpan)> Cells)
        BuildTableGrid(XElement tbl)
    {
        var rows = tbl.Elements(HwpxNs.Hp + "tr").ToList();
        if (rows.Count == 0) return (new XElement?[0, 0], new());

        int maxRow = 0, maxCol = 0;
        var cellList = new List<(XElement tc, int row, int col, int rowSpan, int colSpan)>();

        foreach (var tr in rows)
        {
            foreach (var tc in tr.Elements(HwpxNs.Hp + "tc"))
            {
                var (row, col, rowSpan, colSpan) = GetCellAddr(tc);
                cellList.Add((tc, row, col, rowSpan, colSpan));
                if (row + rowSpan > maxRow) maxRow = row + rowSpan;
                if (col + colSpan > maxCol) maxCol = col + colSpan;
            }
        }

        var grid = new XElement?[maxRow, maxCol];
        foreach (var (tc, row, col, rowSpan, colSpan) in cellList)
        {
            for (int r = row; r < row + rowSpan && r < maxRow; r++)
                for (int c = col; c < col + colSpan && c < maxCol; c++)
                    grid[r, c] = tc;
        }

        return (grid, cellList);
    }

    /// <summary>
    /// Search a single table for a label match and return the adjacent cell.
    /// Uses <see cref="BuildTableGrid"/> for merged cell handling.
    /// </summary>
    private static XElement? FindCellInTable(XElement tbl, string normalizedLabel, string direction)
    {
        var (grid, cellList) = BuildTableGrid(tbl);
        if (cellList.Count == 0) return null;
        int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);

        // Search for label match
        foreach (var (tc, row, col, rowSpan, colSpan) in cellList)
        {
            var cellText = ExtractCellText(tc);
            var normalizedCell = NormalizeLabel(cellText);

            if (!normalizedCell.Equals(normalizedLabel, StringComparison.OrdinalIgnoreCase))
                continue;

            // Calculate target position based on direction
            int targetRow = row, targetCol = col;
            switch (direction)
            {
                case "right": targetCol = col + colSpan; break;
                case "left":  targetCol = col - 1; break;
                case "down":  targetRow = row + rowSpan; break;
                case "up":    targetRow = row - 1; break;
            }

            // Bounds check and return
            if (targetRow >= 0 && targetRow < maxRow &&
                targetCol >= 0 && targetCol < maxCol)
            {
                var target = grid[targetRow, targetCol];
                if (target != null && target != tc) return target;
            }
        }

        return null;
    }

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

    // ==================== Form Recognition (Plan 70.2) ====================

    /// <summary>A single recognized form field from auto-detection.</summary>
    internal record RecognizedField(
        string Label, string Value, string Path, int Row, int Col, string Strategy);

    /// <summary>Korean government form label keywords (~40 items).</summary>
    private static readonly string[] LabelKeywords = [
        "성명", "이름", "주소", "전화", "전화번호", "휴대폰", "연락처",
        "생년월일", "주민등록번호", "소속", "직위", "직급", "부서",
        "이메일", "팩스", "학교", "학년", "반", "학번",
        "신청인", "대표자", "담당자", "작성자",
        "일시", "날짜", "기간", "장소", "목적", "사유", "비고",
        "금액", "수량", "단가", "합계", "계", "소계",
        "동아리명", "사업분야", "참가구분", "인원수",
        // Regulation/expenditure keywords (Plan 70.2 — regulation doc support)
        "비목", "항목해설", "증빙", "집행", "비용항목", "지출",
        "결제일", "결제금액", "카드번호", "승인번호", "사용처",
        "구분", "내용", "지도교수", "검수자", "검수일"
    ];

    /// <summary>
    /// Determine if a cell's text looks like a form label.
    /// Keyword substring match + short Korean heuristic (2-8 chars, no digits).
    /// </summary>
    internal static bool IsLabelCell(string text)
    {
        var trimmed = NormalizeLabel(text);
        if (string.IsNullOrEmpty(trimmed) || trimmed.Length > 30) return false;

        if (LabelKeywords.Any(kw => trimmed.Contains(kw))) return true;

        if (System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^[\uAC00-\uD7A3\s()·]{2,8}$")
            && !System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"\d"))
            return true;

        return false;
    }

    /// <summary>
    /// Recognize form fields from all tables in the document.
    /// Strategy 1: Adjacent cell label-value (left→right).
    /// Strategy 2: Header row + data rows (first row all short text → headers).
    /// </summary>
    internal List<RecognizedField> RecognizeFormFields()
    {
        var fields = new List<RecognizedField>();

        foreach (var (sec, tbl, localTblIdx) in _doc.AllTables())
        {
            var (grid, cellList) = BuildTableGrid(tbl);
            if (cellList.Count == 0) continue;

            int maxRow = grid.GetLength(0), maxCol = grid.GetLength(1);
            var tableFields = new List<RecognizedField>();

            // Strategy 1: Adjacent cell label-value (label left, value right)
            if (maxCol >= 2)
            {
                var seen = new HashSet<XElement>();
                foreach (var (tc, row, col, rowSpan, colSpan) in cellList)
                {
                    if (seen.Contains(tc)) continue;
                    seen.Add(tc);

                    var cellText = ExtractCellText(tc);
                    if (!IsLabelCell(cellText)) continue;

                    int targetCol = col + colSpan;
                    if (targetCol < maxCol)
                    {
                        var valueCell = grid[row, targetCol];
                        if (valueCell != null && valueCell != tc)
                        {
                            var value = ExtractCellText(valueCell).Trim();
                            if (!string.IsNullOrEmpty(value))
                            {
                                var path = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]/tr[{row + 1}]/tc[{col + 1}]";
                                tableFields.Add(new RecognizedField(
                                    NormalizeLabel(cellText), value, path, row, col, "adjacent"));
                            }
                        }
                    }
                }
            }

            // Strategy 2: Header+data (first row all short text → treat as headers)
            if (tableFields.Count == 0 && maxRow >= 2 && maxCol >= 2)
            {
                bool allLabels = true;
                for (int c = 0; c < maxCol; c++)
                {
                    var headerCell = grid[0, c];
                    if (headerCell == null) { allLabels = false; break; }
                    var ht = ExtractCellText(headerCell).Trim();
                    if (string.IsNullOrEmpty(ht) || ht.Length > 20) { allLabels = false; break; }
                }

                if (allLabels)
                {
                    for (int r = 1; r < maxRow; r++)
                        for (int c = 0; c < maxCol; c++)
                        {
                            var headerCell = grid[0, c];
                            var dataCell = grid[r, c];
                            if (headerCell == null || dataCell == null) continue;
                            var label = ExtractCellText(headerCell).Trim();
                            var value = ExtractCellText(dataCell).Trim();
                            if (!string.IsNullOrEmpty(label) && !string.IsNullOrEmpty(value))
                            {
                                var path = $"/section[{sec.Index + 1}]/tbl[{localTblIdx + 1}]/tr[{r + 1}]/tc[{c + 1}]";
                                tableFields.Add(new RecognizedField(
                                    NormalizeLabel(label), value, path, r, c, "header-data"));
                            }
                        }
                }
            }

            fields.AddRange(tableFields);
        }

        return fields;
    }
}
