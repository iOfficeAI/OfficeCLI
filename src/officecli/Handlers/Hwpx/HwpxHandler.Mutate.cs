// File: src/officecli/Handlers/Hwpx/HwpxHandler.Mutate.cs
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Add ====================

    /// <summary>
    /// Add a new element under the parent at the given path.
    /// Returns the path of the newly created element.
    /// </summary>
    /// <param name="parentPath">Path to the parent element.</param>
    /// <param name="type">Element type: "paragraph", "table", "run" (lowercase).</param>
    /// <param name="index">Optional 1-based insertion index. null = append.</param>
    /// <param name="properties">Optional properties for the new element.</param>
    public string Add(string parentPath, string type, int? index,
                      Dictionary<string, string> properties)
    {
        var parent = ResolvePath(parentPath);

        // Header/footer: special handling — adds to secPr, not to parent directly
        if (type.Equals("header", StringComparison.OrdinalIgnoreCase) || type.Equals("footer", StringComparison.OrdinalIgnoreCase))
        {
            var isHeader = type.Equals("header", StringComparison.OrdinalIgnoreCase);
            var hfElement = AddHeaderFooter(parent, properties, isHeader);
            _dirty = true;
            // SaveSection needs an element from the target section document
            SaveSection(hfElement);
            return $"/{(isHeader ? "header" : "footer")}[1]";
        }

        // Memo: special handling — adds to <hp:memogroup> at section level, not inline
        if (type.Equals("comment", StringComparison.OrdinalIgnoreCase) || type.Equals("memo", StringComparison.OrdinalIgnoreCase))
        {
            var memoElement = AddMemoToGroup(parent, properties);
            _dirty = true;
            SaveSection(memoElement);
            var memoCount = memoElement.Parent?.Elements(HwpxNs.Hp + "memo").Count() ?? 1;
            return $"/memo[{memoCount}]";
        }

        var newElement = type.ToLowerInvariant() switch
        {
            "paragraph" or "p" => CreateParagraph(properties),
            "table" or "tbl"   => CreateTable(properties),
            "run"              => CreateRun(properties),
            "row" or "tr"      => CreateRow(parent, properties),
            "cell" or "tc"     => CreateCell(parent, properties),
            "picture" or "image" or "pic" => CreatePicture(properties),
            "hyperlink" or "link"         => CreateHyperlink(properties),
            "pagebreak" or "page-break"   => CreatePageBreak(),
            "footnote"                    => CreateFootnote(properties),
            "endnote"                     => CreateFootnote(properties, isEndnote: true),
            "pagenum" or "pagenumber"     => CreatePageNum(properties),
            "bookmark"                    => CreateBookmark(properties),
            _ => throw new CliException($"Unsupported element type: {type}")
        };

        // Hancom requires tables and pictures to be wrapped: <hp:p><hp:run><hp:tbl/pic>...</hp:tbl/pic></hp:run></hp:p>
        // If adding to a section (or section-like parent), wrap in p>run.
        var needsWrap = (newElement.Name == HwpxNs.Hp + "tbl" || newElement.Name == HwpxNs.Hp + "pic")
                        && IsSectionLike(parent);
        if (needsWrap)
        {
            newElement = WrapTableInParagraph(newElement);
        }

        // Insert at index or append
        if (index.HasValue)
        {
            var targetName = newElement.Name;
            var siblings = parent.Elements(targetName).ToList();
            var insertIdx = index.Value - 1; // convert 1-based to 0-based

            if (insertIdx <= 0 || siblings.Count == 0)
            {
                // Insert as first child of this type
                var firstOfType = parent.Elements(targetName).FirstOrDefault();
                if (firstOfType != null)
                    firstOfType.AddBeforeSelf(newElement);
                else
                    parent.Add(newElement);
            }
            else if (insertIdx >= siblings.Count)
            {
                // Append after last sibling of this type
                siblings.Last().AddAfterSelf(newElement);
            }
            else
            {
                // Insert before the element currently at this index
                siblings[insertIdx].AddBeforeSelf(newElement);
            }
        }
        else
        {
            parent.Add(newElement);
        }

        _dirty = true;
        SaveSection(parent);
        return BuildPath(newElement);
    }

    // ==================== Element Factories ====================

    /// <summary>
    /// Create a new paragraph element with optional text content.
    /// Props: "text" → paragraph text, "styleidref" → style ID, "charpridref" → char property ID.
    /// </summary>
    private XElement CreateParagraph(Dictionary<string, string>? props)
    {
        var id = NewId();
        var text = props?.GetValueOrDefault("text") ?? "";
        var styleIdRef = props?.GetValueOrDefault("styleidref") ?? props?.GetValueOrDefault("styleIDRef") ?? "0";
        var charPrIdRef = props?.GetValueOrDefault("charpridref") ?? props?.GetValueOrDefault("charPrIDRef") ?? "0";
        var paraPrIdRef = props?.GetValueOrDefault("parapridref") ?? props?.GetValueOrDefault("paraPrIDRef") ?? "0";

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", id),
            new XAttribute("styleIDRef", styleIdRef),
            new XAttribute("paraPrIDRef", paraPrIdRef),
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", charPrIdRef),
                new XElement(HwpxNs.Hp + "t", text)
            )
        );
    }

    /// <summary>
    /// Create a new table element with full DOCX-parity features.
    ///
    /// CRITICAL: Every &lt;hp:tc&gt; MUST have ALL of the following children
    /// in this exact order, or Hancom will crash on open:
    ///   - &lt;hp:subList vertAlign="CENTER" ...&gt;&lt;hp:p .../&gt;&lt;/hp:subList&gt;
    ///   - &lt;hp:cellAddr colAddr="C" rowAddr="R"/&gt;
    ///   - &lt;hp:cellSpan colSpan="1" rowSpan="1"/&gt;
    ///   - &lt;hp:cellSz width="W" height="H"/&gt;
    ///   - &lt;hp:cellMargin left="510" right="510" top="141" bottom="141"/&gt;
    ///
    /// Props:
    ///   "rows"           → row count (default 2)
    ///   "cols"           → col count (default 2)
    ///   "width"          → total table width in HWPML units (default 42520 ≈ A4 body)
    ///   "data"           → cell data: "H1,H2;R1C1,R1C2" or CSV file path
    ///   "colWidths"      → per-column widths: "10000,15000,17520"
    ///   "merge"          → merge spec: "startRow,startCol,endRow,endCol;..."
    ///   "borderFillIDRef"→ table-level border fill ID (default "1")
    /// </summary>
    private XElement CreateTable(Dictionary<string, string>? props)
    {
        var id = NewId();

        // Parse data: "H1,H2;R1C1,R1C2;R2C1,R2C2" or CSV file path
        string[][]? tableData = null;
        if (props?.TryGetValue("data", out var dataStr) == true && !string.IsNullOrEmpty(dataStr))
        {
            if (File.Exists(dataStr))
                tableData = File.ReadAllLines(dataStr)
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(l => l.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
            else
                tableData = dataStr.Split(';')
                    .Select(r => r.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
        }

        // Determine dimensions
        int rows, cols;
        if (tableData != null)
        {
            rows = tableData.Length;
            cols = tableData.Max(r => r.Length);
            // Allow explicit overrides to be larger
            if (int.TryParse(props?.GetValueOrDefault("rows"), out var r2) && r2 > rows) rows = r2;
            if (int.TryParse(props?.GetValueOrDefault("cols"), out var c2) && c2 > cols) cols = c2;
        }
        else
        {
            rows = int.TryParse(props?.GetValueOrDefault("rows"), out var r) && r > 0 ? r : 2;
            cols = int.TryParse(props?.GetValueOrDefault("cols"), out var c) && c > 0 ? c : 2;
        }

        var totalWidth = int.TryParse(props?.GetValueOrDefault("width"), out var w) && w > 0 ? w : 42520;
        var defaultCellWidth = totalWidth / Math.Max(cols, 1);

        // Parse per-column widths: "10000,15000,17520"
        int[]? colWidthArr = null;
        if ((props?.TryGetValue("colwidths", out var cwStr) == true
            || props?.TryGetValue("colWidths", out cwStr) == true)
            && !string.IsNullOrEmpty(cwStr))
        {
            colWidthArr = cwStr.Split(',')
                .Select(s => int.TryParse(s.Trim(), out var v) ? v : defaultCellWidth)
                .ToArray();
        }

        var borderFillRef = props?.GetValueOrDefault("borderfillid")
            ?? props?.GetValueOrDefault("borderFillIDRef")
            ?? EnsureTableBorderFill();

        var cellHeight = 1000;
        var totalHeight = rows * cellHeight;

        var tbl = new XElement(HwpxNs.Hp + "tbl",
            new XAttribute("id", id),
            new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "TABLE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"),
            new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"),
            new XAttribute("dropcapstyle", "None"),
            new XAttribute("pageBreak", "CELL"),
            new XAttribute("repeatHeader", "1"),
            new XAttribute("rowCnt", rows.ToString()),
            new XAttribute("colCnt", cols.ToString()),
            new XAttribute("cellSpacing", "0"),
            new XAttribute("borderFillIDRef", borderFillRef),
            new XAttribute("noAdjust", "0"),
            // Table size — required by Hancom for rendering
            new XElement(HwpxNs.Hp + "sz",
                new XAttribute("width", totalWidth.ToString()),
                new XAttribute("widthRelTo", "ABSOLUTE"),
                new XAttribute("height", totalHeight.ToString()),
                new XAttribute("heightRelTo", "ABSOLUTE"),
                new XAttribute("protect", "0")),
            // Position — treatAsChar=1 makes table inline with text
            new XElement(HwpxNs.Hp + "pos",
                new XAttribute("treatAsChar", "1"),
                new XAttribute("affectLSpacing", "0"),
                new XAttribute("flowWithText", "1"),
                new XAttribute("allowOverlap", "0"),
                new XAttribute("holdAnchorAndSO", "0"),
                new XAttribute("vertRelTo", "PARA"),
                new XAttribute("horzRelTo", "COLUMN"),
                new XAttribute("vertAlign", "TOP"),
                new XAttribute("horzAlign", "LEFT"),
                new XAttribute("vertOffset", "0"),
                new XAttribute("horzOffset", "0")),
            // Outer margin
            new XElement(HwpxNs.Hp + "outMargin",
                new XAttribute("left", "283"),
                new XAttribute("right", "283"),
                new XAttribute("top", "283"),
                new XAttribute("bottom", "283")),
            // Inner margin
            new XElement(HwpxNs.Hp + "inMargin",
                new XAttribute("left", "510"),
                new XAttribute("right", "510"),
                new XAttribute("top", "141"),
                new XAttribute("bottom", "141"))
        );

        // Column widths — Hancom uses these to distribute column sizes
        for (int col = 0; col < cols; col++)
        {
            var cw = colWidthArr != null && col < colWidthArr.Length ? colWidthArr[col] : defaultCellWidth;
            tbl.Add(new XElement(HwpxNs.Hp + "colSz",
                new XAttribute("width", cw.ToString())));
        }

        // Parse merge instructions: "0,0,0,3;1,0,2,0" (startRow,startCol,endRow,endCol)
        var mergedCells = new HashSet<(int row, int col)>();
        var spanMap = new Dictionary<(int row, int col), (int rowSpan, int colSpan)>();
        if (props?.TryGetValue("merge", out var mergeStr) == true && !string.IsNullOrEmpty(mergeStr))
        {
            foreach (var m in mergeStr.Split(';'))
            {
                var parts = m.Trim().Split(',');
                if (parts.Length == 4
                    && int.TryParse(parts[0].Trim(), out var sr)
                    && int.TryParse(parts[1].Trim(), out var sc)
                    && int.TryParse(parts[2].Trim(), out var er)
                    && int.TryParse(parts[3].Trim(), out var ec))
                {
                    for (int mr = sr; mr <= er; mr++)
                        for (int mc = sc; mc <= ec; mc++)
                        {
                            if (mr == sr && mc == sc)
                                spanMap[(mr, mc)] = (er - sr + 1, ec - sc + 1);
                            else
                                mergedCells.Add((mr, mc));
                        }
                }
            }
        }

        // Build rows and cells
        for (int row = 0; row < rows; row++)
        {
            var tr = new XElement(HwpxNs.Hp + "tr");

            for (int col = 0; col < cols; col++)
            {
                // Skip cells covered by a merge
                if (mergedCells.Contains((row, col)))
                    continue;

                var cellId = NewId();
                var rowSpan = 1;
                var colSpan = 1;
                if (spanMap.TryGetValue((row, col), out var span))
                {
                    rowSpan = span.rowSpan;
                    colSpan = span.colSpan;
                }

                // Cell width = sum of spanned columns
                var cellW = 0;
                for (int ci = col; ci < col + colSpan && ci < cols; ci++)
                    cellW += colWidthArr != null && ci < colWidthArr.Length ? colWidthArr[ci] : defaultCellWidth;

                // Cell text from data prop or positional prop
                var cellText = "";
                if (tableData != null && row < tableData.Length && col < tableData[row].Length)
                    cellText = tableData[row][col];
                else if (props?.TryGetValue($"r{row + 1}c{col + 1}", out var rc) == true)
                    cellText = rc;

                // Per-cell borderFillIDRef: check "r{row+1}c{col+1}borderfillid" prop, fallback to table-level
                var cellBorderFill = props?.GetValueOrDefault($"r{row + 1}c{col + 1}borderfillid")
                    ?? props?.GetValueOrDefault($"r{row + 1}borderfillid")
                    ?? borderFillRef;

                // Per-cell vertAlign: check "r{row+1}c{col+1}valign" prop
                var cellVertAlign = props?.GetValueOrDefault($"r{row + 1}c{col + 1}valign")
                    ?? props?.GetValueOrDefault("valign") ?? "CENTER";

                var tc = BuildCell(row, col, rowSpan, colSpan, cellW, cellHeight,
                    cellText, cellBorderFill, isHeader: row == 0, vertAlign: cellVertAlign);

                tr.Add(tc);
            }

            tbl.Add(tr);
        }

        return tbl;
    }

    /// <summary>
    /// Create a new run element with optional text content.
    /// Props: "text" → run text, "charpridref" → char property ID.
    /// </summary>
    private XElement CreateRun(Dictionary<string, string>? props)
    {
        var text = props?.GetValueOrDefault("text") ?? "";
        var charPrIdRef = props?.GetValueOrDefault("charpridref") ?? props?.GetValueOrDefault("charPrIDRef") ?? "0";

        return new XElement(HwpxNs.Hp + "run",
            new XAttribute("charPrIDRef", charPrIdRef),
            new XElement(HwpxNs.Hp + "t", text)
        );
    }

    /// <summary>
    /// Create a new table row with cells. Parent MUST be a &lt;hp:tbl&gt;.
    /// Props: "cols" → cell count (default from table colCnt),
    ///        "c1", "c2", ... → cell text for each column.
    /// </summary>
    private XElement CreateRow(XElement parent, Dictionary<string, string>? props)
    {
        if (parent.Name.LocalName != "tbl")
            throw new CliException("Rows can only be added to a table element");

        var colCnt = int.TryParse(parent.Attribute("colCnt")?.Value, out var cc) ? cc : 1;
        var cols = int.TryParse(props?.GetValueOrDefault("cols"), out var c) && c > 0 ? c : colCnt;
        var existingRows = parent.Elements(HwpxNs.Hp + "tr").Count();

        // Get column widths from existing colSz elements
        var colSizes = parent.Elements(HwpxNs.Hp + "colSz")
            .Select(e => int.TryParse(e.Attribute("width")?.Value, out var w) ? w : 42520 / cols)
            .ToList();

        var tr = new XElement(HwpxNs.Hp + "tr");

        for (int col = 0; col < cols; col++)
        {
            var cellText = props?.GetValueOrDefault($"c{col + 1}") ?? "";
            var cellWidth = col < colSizes.Count ? colSizes[col] : 42520 / cols;

            tr.Add(BuildCell(existingRows, col, 1, 1, cellWidth, 1000, cellText, "1"));
        }

        // Update table rowCnt
        parent.SetAttributeValue("rowCnt", (existingRows + 1).ToString());

        return tr;
    }

    /// <summary>
    /// Create a new table cell. Parent MUST be a &lt;hp:tr&gt;.
    /// Props: "text" → cell text, "width" → cell width.
    /// </summary>
    private XElement CreateCell(XElement parent, Dictionary<string, string>? props)
    {
        if (parent.Name.LocalName != "tr")
            throw new CliException("Cells can only be added to a table row element");

        var existingCells = parent.Elements(HwpxNs.Hp + "tc").Count();
        var text = props?.GetValueOrDefault("text") ?? "";
        var cellWidth = int.TryParse(props?.GetValueOrDefault("width"), out var w) && w > 0 ? w : 10000;

        // Determine row address from parent's position in the table
        var tbl = parent.Parent;
        var rowAddr = tbl?.Elements(HwpxNs.Hp + "tr").ToList().IndexOf(parent) ?? 0;

        return BuildCell(rowAddr, existingCells, 1, 1, cellWidth, 1000, text, "1");
    }

    // ==================== Remove ====================

    /// <summary>
    /// Remove the element at the given path.
    /// Returns null on success. Throws CliException if not found.
    /// </summary>
    public string? Remove(string path)
    {
        var element = ResolvePath(path);

        // Capture parent for SaveSection before detaching
        var parent = element.Parent
            ?? throw new CliException($"Cannot remove root element at: {path}");

        element.Remove();
        _dirty = true;
        SaveSection(parent);
        return null;
    }

    // ==================== Move ====================

    /// <summary>
    /// Move an element from sourcePath to a new position under targetParentPath.
    ///
    /// CORRECT detach-then-insert pattern:
    /// 1. Resolve targetParentPath FIRST (validate before modifying the tree).
    /// 2. Resolve sourcePath.
    /// 3. Detach: call source.Remove() BEFORE re-inserting.
    ///    Bad pattern: target.Add(source) when source is still parented —
    ///    XLinq silently moves it, but only within the same XDocument.
    ///    Cross-section moves fail silently without detach.
    /// 4. Insert at the specified index under the target parent.
    /// </summary>
    /// <returns>New path of the moved element.</returns>
    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        if (string.IsNullOrEmpty(targetParentPath))
            throw new CliException("Target parent path is required for move");

        // 1. Resolve target FIRST (before tree modification)
        var target = ResolvePath(targetParentPath);

        // 2. Resolve source
        var source = ResolvePath(sourcePath);
        var sourceParent = source.Parent;

        // 3. Detach source — NEVER re-parent directly
        source.Remove();

        // 4. Insert at position
        if (index.HasValue)
        {
            var siblings = target.Elements(source.Name).ToList();
            var insertIdx = index.Value - 1;

            if (insertIdx <= 0 || siblings.Count == 0)
            {
                var firstOfType = target.Elements(source.Name).FirstOrDefault();
                if (firstOfType != null)
                    firstOfType.AddBeforeSelf(source);
                else
                    target.Add(source);
            }
            else if (insertIdx >= siblings.Count)
            {
                siblings.Last().AddAfterSelf(source);
            }
            else
            {
                siblings[insertIdx].AddBeforeSelf(source);
            }
        }
        else
        {
            target.Add(source);
        }

        _dirty = true;

        // Save both affected sections
        if (sourceParent != null)
            SaveSection(sourceParent);
        SaveSection(target);

        return BuildPath(source);
    }

    // ==================== CopyFrom ====================

    /// <summary>
    /// Deep-clone the element at sourcePath and insert the copy under targetParentPath.
    /// Assigns a new id attribute to the clone to avoid duplicate IDs.
    /// </summary>
    /// <returns>Path of the newly created copy.</returns>
    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var source = ResolvePath(sourcePath);
        var target = ResolvePath(targetParentPath);

        // Deep clone
        var clone = new XElement(source);

        // Assign new IDs to clone and all descendants with id attributes
        AssignNewIds(clone);

        // Insert at position
        if (index.HasValue)
        {
            var siblings = target.Elements(clone.Name).ToList();
            var insertIdx = index.Value - 1;

            if (insertIdx <= 0 || siblings.Count == 0)
            {
                var firstOfType = target.Elements(clone.Name).FirstOrDefault();
                if (firstOfType != null)
                    firstOfType.AddBeforeSelf(clone);
                else
                    target.Add(clone);
            }
            else if (insertIdx >= siblings.Count)
            {
                siblings.Last().AddAfterSelf(clone);
            }
            else
            {
                siblings[insertIdx].AddBeforeSelf(clone);
            }
        }
        else
        {
            target.Add(clone);
        }

        _dirty = true;
        SaveSection(target);
        return BuildPath(clone);
    }

    /// <summary>
    /// Recursively assign new IDs to an element and all descendants
    /// that have an "id" attribute. Prevents duplicate IDs in the document.
    /// </summary>
    private void AssignNewIds(XElement element)
    {
        if (element.Attribute("id") != null)
        {
            element.SetAttributeValue("id", NewId());
        }

        foreach (var child in element.Elements())
        {
            AssignNewIds(child);
        }
    }

    // ==================== Helpers ====================

    /// <summary>
    /// Ensure a borderFill with SOLID black borders exists in header.xml.
    /// Returns the borderFill ID. If one already exists, returns its ID.
    /// If not, creates a new one and returns the new ID.
    /// </summary>
    private string EnsureTableBorderFill()
    {
        var header = _doc.Header?.Root;
        if (header == null) return "1"; // fallback

        var refList = header.Element(HwpxNs.Hh + "refList");
        if (refList == null) return "1";

        var borderFills = refList.Element(HwpxNs.Hh + "borderFills");
        if (borderFills == null) return "1";

        // Check if any existing borderFill has SOLID borders on all 4 sides
        foreach (var bf in borderFills.Elements(HwpxNs.Hh + "borderFill"))
        {
            var left = bf.Element(HwpxNs.Hh + "leftBorder");
            var right = bf.Element(HwpxNs.Hh + "rightBorder");
            var top = bf.Element(HwpxNs.Hh + "topBorder");
            var bottom = bf.Element(HwpxNs.Hh + "bottomBorder");
            if (left?.Attribute("type")?.Value == "SOLID"
                && right?.Attribute("type")?.Value == "SOLID"
                && top?.Attribute("type")?.Value == "SOLID"
                && bottom?.Attribute("type")?.Value == "SOLID")
            {
                return bf.Attribute("id")?.Value ?? "1";
            }
        }

        // None found — create a new one with SOLID black borders
        var newId = NextBorderFillId();

        var newBorderFill = new XElement(HwpxNs.Hh + "borderFill",
            new XAttribute("id", newId),
            new XAttribute("threeD", "0"),
            new XAttribute("shadow", "0"),
            new XAttribute("centerLine", "NONE"),
            new XAttribute("breakCellSeparateLine", "0"),
            new XElement(HwpxNs.Hh + "slash", new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
            new XElement(HwpxNs.Hh + "backSlash", new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
            MakeBorder("leftBorder", "SOLID", "0.12 mm", "#000000"),
            MakeBorder("rightBorder", "SOLID", "0.12 mm", "#000000"),
            MakeBorder("topBorder", "SOLID", "0.12 mm", "#000000"),
            MakeBorder("bottomBorder", "SOLID", "0.12 mm", "#000000"),
            MakeBorder("diagonal", "SOLID", "0.12 mm", "#000000")
        );

        borderFills.Add(newBorderFill);
        var existingCount = borderFills.Elements(HwpxNs.Hh + "borderFill").Count();
        borderFills.SetAttributeValue("itemCnt", existingCount.ToString());

        // Save the modified header
        SaveHeader();

        return newId;
    }

    /// <summary>
    /// Check if parent element is a section or section-like container.
    /// Tables added to these containers must be wrapped in p>run.
    /// </summary>
    private static bool IsSectionLike(XElement parent)
    {
        var localName = parent.Name.LocalName;
        return localName is "sec" or "section" or "body";
    }

    /// <summary>
    /// Wrap a &lt;hp:tbl&gt; in &lt;hp:p&gt;&lt;hp:run&gt;...&lt;/hp:run&gt;&lt;/hp:p&gt;
    /// so Hancom renders it. Without this wrapper, tables are invisible.
    /// </summary>
    private XElement WrapTableInParagraph(XElement tbl)
    {
        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                tbl)
        );
    }

    /// <summary>
    /// Build a Hancom-compatible <hp:tc> element with correct child ordering:
    /// subList → cellAddr → cellSpan → cellSz → cellMargin.
    /// This matches the structure produced by Hancom Office (2011 namespace).
    /// </summary>
    private XElement BuildCell(int rowAddr, int colAddr, int rowSpan, int colSpan,
                               int width, int height, string text,
                               string borderFillIDRef, bool isHeader = false,
                               string vertAlign = "CENTER")
    {
        return new XElement(HwpxNs.Hp + "tc",
            new XAttribute("name", ""),
            new XAttribute("header", isHeader ? "1" : "0"),
            new XAttribute("hasMargin", "0"),
            new XAttribute("protect", "0"),
            new XAttribute("editable", "0"),
            new XAttribute("dirty", "0"),
            new XAttribute("borderFillIDRef", borderFillIDRef),
            CreateSubList(text, vertAlign),
            new XElement(HwpxNs.Hp + "cellAddr",
                new XAttribute("colAddr", colAddr.ToString()),
                new XAttribute("rowAddr", rowAddr.ToString())),
            new XElement(HwpxNs.Hp + "cellSpan",
                new XAttribute("colSpan", colSpan.ToString()),
                new XAttribute("rowSpan", rowSpan.ToString())),
            new XElement(HwpxNs.Hp + "cellSz",
                new XAttribute("width", width.ToString()),
                new XAttribute("height", height.ToString())),
            new XElement(HwpxNs.Hp + "cellMargin",
                new XAttribute("left", "510"),
                new XAttribute("right", "510"),
                new XAttribute("top", "141"),
                new XAttribute("bottom", "141"))
        );
    }

    // ==================== Picture ====================

    /// <summary>
    /// Create a picture element. The image file is registered in the ZIP (BinData/) and content.hpf manifest.
    /// Golden template based on real Hancom documents: uses hc:img (NOT hp:img), hc:pt0-pt3 for imgRect.
    /// Props: path (required), width (e.g. "2in"), height (e.g. "1in"), alt.
    /// </summary>
    private XElement CreatePicture(Dictionary<string, string>? props)
    {
        var path = props?.GetValueOrDefault("path")
            ?? throw new CliException("picture requires 'path' property");
        if (!File.Exists(path))
            throw new CliException($"Image file not found: {path}");

        var widthHwp = ParseDimensionToHwpUnit(props?.GetValueOrDefault("width") ?? "2in");
        var heightHwp = ParseDimensionToHwpUnit(props?.GetValueOrDefault("height") ?? "1in");

        // 1. Read image bytes and determine format
        var imageBytes = File.ReadAllBytes(path);
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        if (ext == "jpg") ext = "jpeg";
        var mediaType = ext switch
        {
            "png" => "image/png",
            "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" or "tif" => "image/tiff",
            _ => $"image/{ext}"
        };

        // 2. Find next available image ID in content.hpf
        var imageId = GetNextImageId();
        var binFileName = $"image{imageId}.{ext}";

        // 3. Add image to ZIP at BinData/ (root level, NOT Contents/BinData/)
        var binEntry = _doc.Archive.CreateEntry($"BinData/{binFileName}", System.IO.Compression.CompressionLevel.Optimal);
        using (var binStream = binEntry.Open())
            binStream.Write(imageBytes, 0, imageBytes.Length);

        // 4. Register in content.hpf manifest
        RegisterImageInManifest($"image{imageId}", $"BinData/{binFileName}", mediaType);

        // 5. Create <hp:pic> element (golden template structure from real Hancom docs)
        var id = NewId();
        var instId = NewId();
        return new XElement(HwpxNs.Hp + "pic",
            new XAttribute("id", id),
            new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"),
            new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"),
            new XAttribute("dropcapstyle", "None"),
            new XAttribute("href", ""),
            new XAttribute("groupLevel", "0"),
            new XAttribute("instid", instId),
            new XAttribute("reverse", "0"),
            new XElement(HwpxNs.Hp + "offset", new XAttribute("x", "0"), new XAttribute("y", "0")),
            new XElement(HwpxNs.Hp + "orgSz",
                new XAttribute("width", widthHwp), new XAttribute("height", heightHwp)),
            new XElement(HwpxNs.Hp + "curSz",
                new XAttribute("width", widthHwp), new XAttribute("height", heightHwp)),
            new XElement(HwpxNs.Hp + "flip", new XAttribute("horizontal", "0"), new XAttribute("vertical", "0")),
            new XElement(HwpxNs.Hp + "rotationInfo",
                new XAttribute("angle", "0"),
                new XAttribute("centerX", (widthHwp / 2).ToString()),
                new XAttribute("centerY", (heightHwp / 2).ToString()),
                new XAttribute("rotateimage", "1")),
            new XElement(HwpxNs.Hp + "renderingInfo",
                new XElement(HwpxNs.Hc + "transMatrix",
                    new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                    new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0")),
                new XElement(HwpxNs.Hc + "scaMatrix",
                    new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                    new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0")),
                new XElement(HwpxNs.Hc + "rotMatrix",
                    new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
                    new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0"))),
            // CRITICAL: hc:img, NOT hp:img (core namespace)
            new XElement(HwpxNs.Hc + "img",
                new XAttribute("binaryItemIDRef", $"image{imageId}"),
                new XAttribute("bright", "0"),
                new XAttribute("contrast", "0"),
                new XAttribute("effect", "REAL_PIC"),
                new XAttribute("alpha", "0")),
            new XElement(HwpxNs.Hp + "imgRect",
                new XElement(HwpxNs.Hc + "pt0", new XAttribute("x", "0"), new XAttribute("y", "0")),
                new XElement(HwpxNs.Hc + "pt1", new XAttribute("x", widthHwp), new XAttribute("y", "0")),
                new XElement(HwpxNs.Hc + "pt2", new XAttribute("x", widthHwp), new XAttribute("y", heightHwp)),
                new XElement(HwpxNs.Hc + "pt3", new XAttribute("x", "0"), new XAttribute("y", heightHwp))),
            new XElement(HwpxNs.Hp + "imgClip",
                new XAttribute("left", "0"), new XAttribute("right", widthHwp),
                new XAttribute("top", "0"), new XAttribute("bottom", heightHwp)),
            new XElement(HwpxNs.Hp + "inMargin",
                new XAttribute("left", "0"), new XAttribute("right", "0"),
                new XAttribute("top", "0"), new XAttribute("bottom", "0")),
            new XElement(HwpxNs.Hp + "imgDim",
                new XAttribute("dimwidth", widthHwp), new XAttribute("dimheight", heightHwp)),
            new XElement(HwpxNs.Hp + "effects"),
            new XElement(HwpxNs.Hp + "sz",
                new XAttribute("width", widthHwp), new XAttribute("widthRelTo", "ABSOLUTE"),
                new XAttribute("height", heightHwp), new XAttribute("heightRelTo", "ABSOLUTE"),
                new XAttribute("protect", "0")),
            new XElement(HwpxNs.Hp + "pos",
                new XAttribute("treatAsChar", "1"), new XAttribute("affectLSpacing", "0"),
                new XAttribute("flowWithText", "1"), new XAttribute("allowOverlap", "0"),
                new XAttribute("holdAnchorAndSO", "0"),
                new XAttribute("vertRelTo", "PARA"), new XAttribute("horzRelTo", "COLUMN"),
                new XAttribute("vertAlign", "TOP"), new XAttribute("horzAlign", "LEFT"),
                new XAttribute("vertOffset", "0"), new XAttribute("horzOffset", "0")),
            new XElement(HwpxNs.Hp + "outMargin",
                new XAttribute("left", "0"), new XAttribute("right", "0"),
                new XAttribute("top", "0"), new XAttribute("bottom", "0"))
        );
    }

    // ==================== Hyperlink ====================

    /// <summary>
    /// Create a hyperlink using the OWPML fieldBegin/fieldEnd pattern (3-run structure).
    /// Golden template based on OWPML schema + python-hwpx implementation.
    /// Props: url/href (required), text (default=url).
    /// </summary>
    private XElement CreateHyperlink(Dictionary<string, string>? props)
    {
        var url = props?.GetValueOrDefault("url") ?? props?.GetValueOrDefault("href")
            ?? throw new CliException("hyperlink requires 'url' property");
        var text = props?.GetValueOrDefault("text") ?? url;
        var fieldId = NewId();
        var fieldIdNum = NewId();

        // Determine link category and command encoding (golden template 2026-04-11)
        string category, command;
        if (url.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
        {
            category = "HWPHYPERLINK_TYPE_EMAIL";
            command = EscapeHyperlinkCommand(url) + ";2;0;0;";
        }
        else if (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
              || url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            category = "HWPHYPERLINK_TYPE_URL";
            command = EscapeHyperlinkCommand(url) + ";1;0;0;";
        }
        else
        {
            category = "HWPHYPERLINK_TYPE_EX";
            command = EscapeHyperlinkCommand(url) + ";3;0;0;";
        }

        // Ensure hyperlink charPr exists in header.xml (blue underline)
        var linkCharPrId = EnsureHyperlinkCharPr();

        // Build parameters element
        var parameters = new XElement(HwpxNs.Hp + "parameters",
            new XAttribute("cnt", "6"),
            new XAttribute("name", ""),
            new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "0"),
            new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Command"), command),
            new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Path"), url),
            new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Category"), category),
            new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "TargetType"), "HWPHYPERLINK_TARGET_BOOKMARK"),
            new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "DocOpenType"), "HWPHYPERLINK_JUMP_CURRENTTAB"));

        // Hyperlinks in HWPX use fieldBegin/fieldEnd (golden template confirmed).
        // URL type uses double nesting; email/file use single nesting.
        var para = new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            // Run 1: fieldBegin with parameters
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldBegin",
                        new XAttribute("id", fieldId),
                        new XAttribute("type", "HYPERLINK"),
                        new XAttribute("name", ""),
                        new XAttribute("editable", "0"),
                        new XAttribute("dirty", "1"),
                        new XAttribute("zorder", "-1"),
                        new XAttribute("fieldid", fieldIdNum),
                        new XAttribute("metaTag", ""),
                        parameters))),
            // Run 2: visible text with hyperlink charPr (blue underline)
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", linkCharPrId),
                new XElement(HwpxNs.Hp + "t", text)),
            // Run 3: fieldEnd
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldEnd",
                        new XAttribute("beginIDRef", fieldId),
                        new XAttribute("fieldid", fieldIdNum))),
                new XElement(HwpxNs.Hp + "t"))
        );
        return para;
    }

    /// <summary>
    /// Escape special characters in hyperlink Command parameter.
    /// Colons and semicolons are escaped with backslash (golden template verified).
    /// </summary>
    private static string EscapeHyperlinkCommand(string url)
        => url.Replace(":", "\\:").Replace(";", "\\;");

    /// <summary>
    /// Ensure a hyperlink charPr (blue text, bottom underline) exists in header.xml.
    /// Returns the charPr id string. Creates one if not found.
    /// </summary>
    private string EnsureHyperlinkCharPr()
    {
        // Look for existing hyperlink charPr (textColor=#0000FF with underline BOTTOM)
        var charPrs = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "charPr");
        if (charPrs != null)
        {
            foreach (var cp in charPrs)
            {
                if (cp.Attribute("textColor")?.Value == "#0000FF")
                {
                    var underline = cp.Element(HwpxNs.Hh + "underline");
                    if (underline?.Attribute("type")?.Value == "BOTTOM")
                        return cp.Attribute("id")?.Value ?? "0";
                }
            }
        }

        // Create new hyperlink charPr
        var newId = NextCharPrId();

        // Clone from charPr id=0 as base
        var baseCharPr = FindCharPr("0");
        XElement newCharPr;
        if (baseCharPr != null)
        {
            newCharPr = new XElement(baseCharPr);
            newCharPr.SetAttributeValue("id", newId.ToString());
        }
        else
        {
            newCharPr = new XElement(HwpxNs.Hh + "charPr",
                new XAttribute("id", newId.ToString()),
                new XAttribute("height", "1000"),
                new XAttribute("shadeColor", "none"),
                new XAttribute("useFontSpace", "0"),
                new XAttribute("useKerning", "0"),
                new XAttribute("symMark", "NONE"),
                new XAttribute("borderFillIDRef", "2"));
        }

        // Set blue text color
        newCharPr.SetAttributeValue("textColor", "#0000FF");

        // Set underline to BOTTOM SOLID blue
        var underlineEl = newCharPr.Element(HwpxNs.Hh + "underline");
        if (underlineEl != null)
        {
            underlineEl.SetAttributeValue("type", "BOTTOM");
            underlineEl.SetAttributeValue("shape", "SOLID");
            underlineEl.SetAttributeValue("color", "#0000FF");
        }
        else
        {
            newCharPr.Add(new XElement(HwpxNs.Hh + "underline",
                new XAttribute("type", "BOTTOM"),
                new XAttribute("shape", "SOLID"),
                new XAttribute("color", "#0000FF")));
        }

        // Add to header.xml
        var lastCharPr = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "charPr").LastOrDefault();
        if (lastCharPr != null)
            lastCharPr.AddAfterSelf(newCharPr);

        SaveHeader();
        return newId.ToString();
    }

    // ==================== Page Break ====================

    /// <summary>
    /// Create a page break paragraph. In HWPX, page break is simply a paragraph
    /// with pageBreak="1" attribute.
    /// </summary>
    private XElement CreatePageBreak()
    {
        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "1"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"));
    }

    // ==================== Footnote ====================

    /// <summary>
    /// Create a footnote or endnote. Uses hp:ctrl > hp:footNote/endNote > hp:subList structure.
    /// The marker appears at the insertion point; footnote text at page bottom, endnote at document end.
    /// Props: text (required), number (auto if omitted).
    /// </summary>
    private XElement CreateFootnote(Dictionary<string, string>? props, bool isEndnote = false)
    {
        var text = props?.GetValueOrDefault("text")
            ?? throw new CliException($"{(isEndnote ? "endnote" : "footnote")} requires 'text' property");
        var number = props?.GetValueOrDefault("number") ?? "0"; // 0 = auto-number
        var tagName = isEndnote ? "endNote" : "footNote";

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            WrapInRun(
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + tagName,
                        new XAttribute("number", number),
                        CreateSubList(text, "TOP")))));
    }

    // ==================== Comment / Memo ====================

    /// <summary>
    /// Add a memo to the section-level memogroup container.
    /// HWPX memos live in: section > hp:memogroup > hp:memo > hp:paraList > hp:p
    /// NOT inside hp:ctrl inline (that causes Hancom to crash).
    /// Props: text (required).
    /// </summary>
    private XElement AddMemoToGroup(XElement sectionParent, Dictionary<string, string>? props)
    {
        var text = props?.GetValueOrDefault("text")
            ?? throw new CliException("comment/memo requires 'text' property");

        // Ensure memoPr exists in header
        var memoShapeId = EnsureMemoPr();

        // Find or create the section root (hs:sec)
        var section = sectionParent;
        if (section.Name != HwpxNs.Hs + "sec")
            section = sectionParent.AncestorsAndSelf(HwpxNs.Hs + "sec").FirstOrDefault() ?? sectionParent;

        // Find or create <hp:memogroup>
        var memoGroup = section.Element(HwpxNs.Hp + "memogroup");
        if (memoGroup == null)
        {
            memoGroup = new XElement(HwpxNs.Hp + "memogroup");
            section.Add(memoGroup);
        }

        // Create memo with paraList structure (NOT subList)
        var memoId = $"memo{memoGroup.Elements(HwpxNs.Hp + "memo").Count()}";
        var memo = new XElement(HwpxNs.Hp + "memo",
            new XAttribute("id", memoId),
            new XAttribute("memoShapeIDRef", memoShapeId),
            new XElement(HwpxNs.Hp + "paraList",
                new XElement(HwpxNs.Hp + "p",
                    new XAttribute("id", NewId()),
                    new XAttribute("paraPrIDRef", "0"),
                    new XAttribute("styleIDRef", "0"),
                    new XElement(HwpxNs.Hp + "run",
                        new XAttribute("charPrIDRef", "0"),
                        new XElement(HwpxNs.Hp + "t", text)))));

        memoGroup.Add(memo);
        return memo;
    }

    /// <summary>
    /// Ensure a memoProperties/memoPr definition exists in header.xml.
    /// Returns the memoPr ID to use as memoShapeIDRef.
    /// </summary>
    private string EnsureMemoPr()
    {
        var refList = _doc.Header!.Root!.Element(HwpxNs.Hh + "refList");
        if (refList == null)
        {
            refList = new XElement(HwpxNs.Hh + "refList");
            _doc.Header.Root.Add(refList);
        }

        var memoProps = refList.Element(HwpxNs.Hh + "memoProperties");
        if (memoProps != null)
        {
            var existing = memoProps.Elements(HwpxNs.Hh + "memoPr").FirstOrDefault();
            if (existing != null)
                return existing.Attribute("id")?.Value ?? "0";
        }

        // Create memoProperties with default memoPr
        memoProps = new XElement(HwpxNs.Hh + "memoProperties",
            new XAttribute("itemCnt", "1"),
            new XElement(HwpxNs.Hh + "memoPr",
                new XAttribute("id", "0"),
                new XAttribute("width", "15591"),
                new XAttribute("lineWidth", "0.6mm"),
                new XAttribute("lineType", "SOLID"),
                new XAttribute("lineColor", "#B6D7AE"),
                new XAttribute("fillColor", "#F0FFE9"),
                new XAttribute("activeColor", "#CFF1C7"),
                new XAttribute("memoType", "NORMAL")));
        refList.Add(memoProps);
        SaveHeader();

        return "0";
    }

    // ==================== Page Numbering ====================

    /// <summary>
    /// Create a page number element. HWPX uses hp:ctrl > hp:pageNum structure.
    /// Props: pos (default BOTTOM_CENTER), format (default DIGIT).
    /// formatType: DIGIT, CIRCLED_DIGIT, ROMAN_CAPITAL, ROMAN_SMALL, HANGUL, HANJA.
    /// pos: TOP_LEFT, TOP_CENTER, TOP_RIGHT, BOTTOM_LEFT, BOTTOM_CENTER, BOTTOM_RIGHT,
    ///      OUTSIDE_TOP, OUTSIDE_BOTTOM, INSIDE_TOP, INSIDE_BOTTOM.
    /// </summary>
    private XElement CreatePageNum(Dictionary<string, string>? props)
    {
        var pos = props?.GetValueOrDefault("pos") ?? "BOTTOM_CENTER";
        var format = props?.GetValueOrDefault("format") ?? "DIGIT";

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            WrapInRun(
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "pageNum",
                        new XAttribute("pos", pos),
                        new XAttribute("formatType", format),
                        new XAttribute("sideChar", "")))));
    }

    // ==================== Bookmark ====================

    /// <summary>
    /// Create a point bookmark element. HWPX uses hp:ctrl > hp:bookmark structure.
    /// Props: name (required).
    /// Note: Range bookmarks (fieldBegin/fieldEnd) require start/end at different positions
    /// and are not supported in this version.
    /// </summary>
    private XElement CreateBookmark(Dictionary<string, string>? props)
    {
        var name = props?.GetValueOrDefault("name")
            ?? throw new CliException("bookmark requires 'name' property");

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            WrapInRun(
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "bookmark",
                        new XAttribute("name", name)))));
    }

    // ==================== Header / Footer ====================

    /// <summary>
    /// Add header or footer to the section using the ctrl pattern (golden template verified 2026-04-11).
    /// Structure: hp:run > hp:ctrl > hp:header/footer > hp:subList > hp:p
    /// The ctrl is inserted into the first paragraph's secPr run (second position).
    /// Props: text (required), type (BOTH/ODD/EVEN, default=BOTH).
    /// </summary>
    private XElement AddHeaderFooter(XElement sectionRoot, Dictionary<string, string>? props, bool isHeader)
    {
        var text = props?.GetValueOrDefault("text") ?? "";
        var applyPageType = props?.GetValueOrDefault("type") ?? "BOTH";
        var tagName = isHeader ? "header" : "footer";
        var vertAlign = isHeader ? "TOP" : "BOTTOM";

        // Find secPr in the section document
        var doc = sectionRoot.Document ?? sectionRoot.AncestorsAndSelf().Last().Document;
        var searchRoot = doc?.Root ?? sectionRoot;

        var secPr = searchRoot.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault()
            ?? searchRoot.Descendants().FirstOrDefault(e => e.Name.LocalName == "secPr");

        if (secPr == null)
            throw new CliException("Cannot find <secPr> in section to add header/footer");

        // Calculate textWidth/textHeight from pagePr margins
        var pagePr = secPr.Element(HwpxNs.Hp + "pagePr");
        var marginEl = pagePr?.Element(HwpxNs.Hp + "margin");
        var pageWidth = (int?)pagePr?.Attribute("width") ?? 59528;
        var marginLeft = (int?)marginEl?.Attribute("left") ?? 8504;
        var marginRight = (int?)marginEl?.Attribute("right") ?? 8504;
        var marginHf = isHeader
            ? ((int?)marginEl?.Attribute("header") ?? 4252)
            : ((int?)marginEl?.Attribute("footer") ?? 4252);
        var textWidth = pageWidth - marginLeft - marginRight;

        // Determine header/footer id — use incremental: headers start at 1, footers at 2
        var existingHfCount = searchRoot.Descendants(HwpxNs.Hp + "header").Count()
                            + searchRoot.Descendants(HwpxNs.Hp + "footer").Count();
        var hfId = (existingHfCount + 1).ToString();

        // Create subList with correct dimensions (golden template: id="" empty, textWidth/Height from pagePr)
        var subList = new XElement(HwpxNs.Hp + "subList",
            new XAttribute("id", ""),
            new XAttribute("textDirection", "HORIZONTAL"),
            new XAttribute("lineWrap", "BREAK"),
            new XAttribute("vertAlign", vertAlign),
            new XAttribute("linkListIDRef", "0"),
            new XAttribute("linkListNextIDRef", "0"),
            new XAttribute("textWidth", textWidth.ToString()),
            new XAttribute("textHeight", marginHf.ToString()),
            new XAttribute("hasTextRef", "0"),
            new XAttribute("hasNumRef", "0"),
            new XElement(HwpxNs.Hp + "p",
                new XAttribute("id", "0"),
                new XAttribute("paraPrIDRef", "0"),
                new XAttribute("styleIDRef", "0"),
                new XAttribute("pageBreak", "0"),
                new XAttribute("columnBreak", "0"),
                new XAttribute("merged", "0"),
                new XElement(HwpxNs.Hp + "run",
                    new XAttribute("charPrIDRef", "0"),
                    new XElement(HwpxNs.Hp + "t", text)),
                new XElement(HwpxNs.Hp + "linesegarray",
                    new XElement(HwpxNs.Hp + "lineseg",
                        new XAttribute("textpos", "0"),
                        new XAttribute("vertpos", "0"),
                        new XAttribute("vertsize", "1000"),
                        new XAttribute("textheight", "1000"),
                        new XAttribute("baseline", "850"),
                        new XAttribute("spacing", "600"),
                        new XAttribute("horzpos", "0"),
                        new XAttribute("horzsize", textWidth.ToString()),
                        new XAttribute("flags", "393216")))));

        // Create the ctrl element
        var hfElement = new XElement(HwpxNs.Hp + tagName,
            new XAttribute("id", hfId),
            new XAttribute("applyPageType", applyPageType),
            subList);

        var ctrlElement = new XElement(HwpxNs.Hp + "ctrl", hfElement);

        // Find the run that contains secPr and add the ctrl there
        var secPrRun = secPr.Parent;
        if (secPrRun?.Name == HwpxNs.Hp + "run")
        {
            // Insert ctrl after secPr run, or find existing body run to prepend
            var bodyRun = secPrRun.ElementsAfterSelf(HwpxNs.Hp + "run").FirstOrDefault();
            if (bodyRun != null)
            {
                // Add ctrl at the beginning of the body run
                bodyRun.AddFirst(ctrlElement);
            }
            else
            {
                // No body run yet — create one with the ctrl and body text placeholder
                var newRun = new XElement(HwpxNs.Hp + "run",
                    new XAttribute("charPrIDRef", "0"),
                    ctrlElement);
                secPrRun.AddAfterSelf(newRun);
            }
        }
        else
        {
            // Fallback: add as new run in first paragraph
            var firstP = searchRoot.Descendants(HwpxNs.Hp + "p").FirstOrDefault();
            if (firstP != null)
            {
                firstP.Add(new XElement(HwpxNs.Hp + "run",
                    new XAttribute("charPrIDRef", "0"),
                    ctrlElement));
            }
        }

        return hfElement;
    }

    // ==================== Image Helpers ====================

    /// <summary>Find the next available image index by scanning content.hpf manifest.</summary>
    private int GetNextImageId()
    {
        var hpfEntry = _doc.Archive.GetEntry("Contents/content.hpf");
        if (hpfEntry == null) return 1;

        using var stream = hpfEntry.Open();
        var hpf = LoadAndNormalize(stream);
        var maxId = 0;
        foreach (var item in hpf.Descendants().Where(e => e.Name.LocalName == "item"))
        {
            var id = item.Attribute("id")?.Value;
            if (id != null && id.StartsWith("image", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(id.AsSpan("image".Length), out var num) && num > maxId)
                    maxId = num;
            }
        }
        return maxId + 1;
    }

    /// <summary>Register an image item in content.hpf manifest.</summary>
    private void RegisterImageInManifest(string itemId, string href, string mediaType)
    {
        var hpfEntry = _doc.Archive.GetEntry("Contents/content.hpf");
        if (hpfEntry == null)
            throw new CliException("Cannot find Contents/content.hpf in HWPX archive");

        XDocument hpf;
        using (var stream = hpfEntry.Open())
            hpf = LoadAndNormalize(stream);

        // Add item to manifest (inside <opf:manifest>)
        var manifest = hpf.Descendants().FirstOrDefault(e => e.Name.LocalName == "manifest");
        if (manifest == null)
            throw new CliException("Cannot find <manifest> in content.hpf");

        manifest.Add(new XElement(HwpxNs.Opf + "item",
            new XAttribute("id", itemId),
            new XAttribute("href", href),
            new XAttribute("media-type", mediaType),
            new XAttribute("isEmbeded", "1")));

        // Save back to ZIP
        var entryName = hpfEntry.FullName;
        hpfEntry.Delete();
        var newEntry = _doc.Archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal);
        using var outStream = newEntry.Open();
        var xmlStr = HwpxPacker.MinifyXml(hpf.ToString(SaveOptions.DisableFormatting));
        xmlStr = HwpxPacker.RestoreOriginalNamespaces(xmlStr);
        xmlStr = "<?xml version='1.0' encoding='UTF-8'?>" + xmlStr;
        var bytes = System.Text.Encoding.UTF8.GetBytes(xmlStr);
        outStream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>
    /// Parse a dimension string (e.g. "2in", "50mm", "100pt", "5cm") to HWPX units (HWPUNIT).
    /// 1 inch = 7200 HWPUNIT, 1mm ≈ 283.46 HWPUNIT, 1pt = 100 HWPUNIT, 1cm = 2834.6 HWPUNIT.
    /// A4 width = 59528 HWPUNIT ≈ 210mm.
    /// </summary>
    private static int ParseDimensionToHwpUnit(string dim)
    {
        dim = dim.Trim();
        if (int.TryParse(dim, out var rawVal)) return rawVal; // already in HWPUNIT

        // Extract numeric part and unit
        var i = 0;
        while (i < dim.Length && (char.IsDigit(dim[i]) || dim[i] == '.'))
            i++;
        if (i == 0) return 14400; // default 2in

        var number = double.Parse(dim[..i], System.Globalization.CultureInfo.InvariantCulture);
        var unit = dim[i..].Trim().ToLowerInvariant();

        return unit switch
        {
            "in" or "inch" => (int)(number * 7200),
            "mm" => (int)(number * 283.46),
            "cm" => (int)(number * 2834.6),
            "pt" => (int)(number * 100),
            "hwp" => (int)number,
            _ => (int)(number * 7200) // default to inches
        };
    }

    /// <summary>
    /// Generate a unique numeric ID string.
    /// Hancom requires numeric IDs (not hex) for elements to render properly.
    /// Uses a high base + random offset to avoid collisions with existing IDs.
    /// </summary>
    private static long _idCounter = 2000000000L + Random.Shared.Next(0, 100000000);
    private string NewId()
    {
        return Interlocked.Increment(ref _idCounter).ToString();
    }
}
