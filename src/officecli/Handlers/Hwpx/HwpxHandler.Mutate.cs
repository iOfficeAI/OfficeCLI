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
    /// <param name="position">Optional insertion position. null = append.</param>
    /// <param name="properties">Optional properties for the new element.</param>
    public string Add(string parentPath, string type, InsertPosition? position,
                      Dictionary<string, string> properties)
    {
        var index = position?.Index;
        // Section: special handling — creates new section file + manifest entry (no parent needed)
        if (type.Equals("section", StringComparison.OrdinalIgnoreCase))
        {
            var newSection = AddNewSection(properties);
            _dirty = true;
            return $"/section[{newSection.Index + 1}]";
        }

        // Style: header-level, not section-level
        if (type.Equals("style", StringComparison.OrdinalIgnoreCase))
        {
            var newStyle = CreateStyleElement(properties);
            _dirty = true;
            return $"/header/style[{newStyle.Attribute("id")?.Value}]";
        }

        var parent = ResolvePath(parentPath);

        // Header/footer: special handling — adds to secPr, not to parent directly
        if (type.Equals("header", StringComparison.OrdinalIgnoreCase) || type.Equals("footer", StringComparison.OrdinalIgnoreCase))
        {
            var isHeader = type.Equals("header", StringComparison.OrdinalIgnoreCase);
            var hfElement = AddHeaderFooter(parent, properties, isHeader);
            _dirty = true;
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

        // TOC: special handling — generates multiple paragraphs from headings
        if (type.Equals("toc", StringComparison.OrdinalIgnoreCase) || type.Equals("tableofcontents", StringComparison.OrdinalIgnoreCase))
        {
            var mode = properties?.GetValueOrDefault("mode") ?? "static";
            var tocParas = mode.Equals("field", StringComparison.OrdinalIgnoreCase)
                ? CreateFieldToc(properties)
                : CreateStaticToc(properties);
            foreach (var p in tocParas)
                parent.Add(p);
            _dirty = true;
            SaveSection(parent);
            return $"/toc[{tocParas.Count} paragraphs]";
        }

        var newElement = type.ToLowerInvariant() switch
        {
            "paragraph" or "p" => CreateParagraph(properties),
            "table" or "tbl"   => CreateTable(properties),
            "run"              => CreateRun(properties),
            "row" or "tr"      => CreateRow(parent, properties),
            "cell" or "tc"     => CreateCell(parent, properties),
            "picture" or "image" or "pic" => CreatePicture(parent, properties),
            "hyperlink" or "link"         => CreateHyperlink(properties),
            "pagebreak" or "page-break"   => CreatePageBreak(),
            "columnbreak" or "column-break" => CreateColumnBreak(properties),
            "footnote"                    => CreateFootnote(properties),
            "endnote"                     => CreateFootnote(properties, isEndnote: true),
            "pagenum" or "pagenumber"     => CreatePageNum(properties),
            "bookmark"                    => CreateBookmark(properties),
            "equation" or "eq"            => CreateEquation(properties),
            "line"                        => CreateLine(properties),
            "rect" or "rectangle"         => CreateRect(properties),
            "ellipse" or "circle"         => CreateEllipse(properties),
            "textbox"                     => CreateTextBox(properties),
            "polygon"                     => CreatePolygon(properties),
            "triangle"                    => CreatePolygon(MergeProps(properties, "sides", "3")),
            "pentagon"                    => CreatePolygon(MergeProps(properties, "sides", "5")),
            "arrow"                       => CreateArrow(properties),
            "formfield" or "form"        => CreateFormField(properties),
            "field"                       => CreateField(properties),
            "date"                        => CreateField(MergeProps(properties, "type", "DATE")),
            "filepath" or "path"          => CreateField(MergeProps(properties, "type", "PATH")),
            "clickhere"                   => CreateField(MergeProps(properties, "type", "CLICK_HERE")),
            "checkbox" or "check"        => CreateField(MergeProps(properties, "type", "CHECKBOX")),
            "dropdown" or "drop"         => CreateField(MergeProps(properties, "type", "DROPDOWN")),
            "summary" or "summery"        => CreateField(MergeProps(properties, "type", "SUMMERY")),
            "author"                      => CreateField(MergeProps(properties, "type", "SUMMERY", "command", "$author")),
            "title"                       => CreateField(MergeProps(properties, "type", "SUMMERY", "command", "$title")),
            "lastsaveby"                  => CreateField(MergeProps(properties, "type", "SUMMERY", "command", "$lastsaveby")),
            "filename"                    => CreateField(MergeProps(properties, "type", "PATH", "command", "$F")),
            "watermark"                   => CreateWatermark(properties),
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

        // Auto-replace first empty paragraph: if parent is section-like and first paragraph
        // has no text (only secPr), replace it instead of appending after it.
        // This prevents the "always empty first line" issue with base.hwpx template.
        if (!index.HasValue && newElement.Name == HwpxNs.Hp + "p" && IsSectionLike(parent))
        {
            var firstP = parent.Elements(HwpxNs.Hp + "p").FirstOrDefault();
            if (firstP != null)
            {
                var hasText = firstP.Descendants(HwpxNs.Hp + "t").Any(t => !string.IsNullOrEmpty(t.Value));
                var isOnlyPara = parent.Elements(HwpxNs.Hp + "p").Count() == 1;
                if (!hasText && isOnlyPara)
                {
                    // Preserve secPr from the first paragraph (it's structurally required)
                    var secPrRun = firstP.Elements(HwpxNs.Hp + "run")
                        .FirstOrDefault(r => r.Element(HwpxNs.Hp + "secPr") != null);
                    if (secPrRun != null)
                        newElement.AddFirst(secPrRun);
                    firstP.ReplaceWith(newElement);
                    goto afterInsert;
                }
            }
        }

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
        afterInsert:

        // Apply formatting properties (fontsize, bold, italic, color, align, etc.)
        // to the newly created paragraph — CreateParagraph only handles text/style/IDRef
        var createdPara = newElement.Name == HwpxNs.Hp + "p" ? newElement : null;
        if (createdPara != null && properties != null)
        {
            var structuralKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "text", "styleidref", "styleIDRef", "charpridref", "charPrIDRef",
                  "parapridref", "paraPrIDRef" };
            foreach (var (key, value) in properties)
            {
                if (structuralKeys.Contains(key)) continue;
                SetParagraphProp(createdPara, key, value);
            }
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
    /// <summary>Swap two elements in the document (Plan 96).</summary>
    public (string NewPath1, string NewPath2) Swap(string pathA, string pathB)
    {
        var elemA = ResolvePath(pathA);
        var elemB = ResolvePath(pathB);
        var placeholder = new XElement(HwpxNs.Hp + "_swap");
        elemA.ReplaceWith(placeholder);
        elemB.ReplaceWith(elemA);
        placeholder.ReplaceWith(elemB);
        var secA = elemA.AncestorsAndSelf().FirstOrDefault(e => e.Name.LocalName == "sec");
        if (secA != null) SaveSection(secA);
        var secB = elemB.AncestorsAndSelf().FirstOrDefault(e => e.Name.LocalName == "sec");
        if (secB != null && secB != secA) SaveSection(secB);
        _dirty = true;
        return (BuildPath(elemA), BuildPath(elemB));
    }

    /// Remove the element at the given path with type-aware cascading cleanup.
    /// Special paths: /watermark, /pagebackground, /toc.
    /// Returns null on success. Throws CliException if not found.
    /// </summary>
    public string? Remove(string path)
    {
        // Special path handlers (not element-based)
        if (path.Equals("/watermark", StringComparison.OrdinalIgnoreCase)
            || path.Equals("/pagebackground", StringComparison.OrdinalIgnoreCase))
            return RemovePageBackground();
        if (path.Equals("/toc", StringComparison.OrdinalIgnoreCase))
            return RemoveToc();
        // Header/Footer removal: /header[N] or /footer[N]
        var hfMatch = System.Text.RegularExpressions.Regex.Match(
            path, @"^/(header|footer)\[(\d+)\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (hfMatch.Success)
            return RemoveHeaderFooter(hfMatch.Groups[1].Value.Equals("header", StringComparison.OrdinalIgnoreCase), int.Parse(hfMatch.Groups[2].Value));
        // Section removal: /section[N] with no further path segments
        if (System.Text.RegularExpressions.Regex.IsMatch(path, @"^/section\[\d+\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return RemoveSection(path);

        var element = ResolvePath(path);
        var parent = element.Parent
            ?? throw new CliException($"Cannot remove root element at: {path}");

        // Type-aware cascade
        var localName = element.Name.LocalName;
        switch (localName)
        {
            case "tbl":
                // Remove wrapper paragraph if tbl is the only content
                if (IsWrapperParagraph(parent))
                {
                    var wrapperParent = parent.Parent!;
                    parent.Remove();
                    _dirty = true;
                    SaveSection(wrapperParent);
                    return null;
                }
                break;

            case "pic" or "img":
                CleanupBinData(element);
                if (IsWrapperParagraph(parent))
                {
                    var wrapperParent = parent.Parent!;
                    parent.Remove();
                    _dirty = true;
                    SaveSection(wrapperParent);
                    return null;
                }
                break;

            case "equation":
                // equation is wrapped in p > run > equation
                var eqRun = parent; // run
                var eqP = eqRun?.Parent; // p
                if (eqP != null && IsWrapperParagraph(eqP))
                {
                    var eqParent = eqP.Parent!;
                    eqP.Remove();
                    _dirty = true;
                    SaveSection(eqParent);
                    return null;
                }
                break;

            case "memo":
                CleanupMemoMarkers(element);
                break;
        }

        element.Remove();
        _dirty = true;
        SaveSection(parent);
        return null;
    }

    // ==================== Remove Helpers ====================

    /// <summary>
    /// Check if a paragraph is just a wrapper for a single table/picture/equation.
    /// These wrappers should be removed together with their content.
    /// </summary>
    private static bool IsWrapperParagraph(XElement p)
    {
        if (p.Name != HwpxNs.Hp + "p") return false;
        var runs = p.Elements(HwpxNs.Hp + "run").ToList();
        if (runs.Count != 1) return false;
        var run = runs[0];
        var nonTextChildren = run.Elements()
            .Where(e => e.Name != HwpxNs.Hp + "t" || !string.IsNullOrWhiteSpace(e.Value))
            .ToList();
        return nonTextChildren.Count == 1 && (
            nonTextChildren[0].Name == HwpxNs.Hp + "tbl" ||
            nonTextChildren[0].Name == HwpxNs.Hp + "pic" ||
            nonTextChildren[0].Name == HwpxNs.Hp + "equation");
    }

    /// <summary>
    /// Remove header or footer ctrl element from section XML by 1-based index.
    /// </summary>
    private string? RemoveHeaderFooter(bool isHeader, int index)
    {
        var tagName = isHeader ? "header" : "footer";
        int found = 0;
        foreach (var section in _doc.Sections)
        {
            foreach (var hf in section.Root.Descendants(HwpxNs.Hp + tagName).ToList())
            {
                found++;
                if (found == index)
                {
                    var ctrl = hf.Parent;
                    if (ctrl?.Name == HwpxNs.Hp + "ctrl")
                        ctrl.Remove();
                    else
                        hf.Remove();
                    _dirty = true;
                    SaveSection(section.Root);
                    return null;
                }
            }
        }
        throw new CliException($"Cannot find {tagName}[{index}] in HWPX document");
    }

    /// <summary>
    /// Clean up BinData ZIP entry when an image is removed.
    /// </summary>
    private void CleanupBinData(XElement picElement)
    {
        var imgEl = picElement.Descendants()
            .FirstOrDefault(e => e.Name.LocalName == "img" || e.Name.LocalName == "binItem");
        var binRef = imgEl?.Attribute("binaryItemIDRef")?.Value
            ?? imgEl?.Attribute("src")?.Value;
        if (binRef == null) return;

        // Delete BinData entry — try Contents/ prefixed paths first, then root BinData/ (legacy)
        var entryPath = binRef.StartsWith("Contents/")
            ? binRef
            : binRef.StartsWith("BinData/")
                ? $"Contents/{binRef}"
                : $"Contents/BinData/{binRef}";
        var entry = _doc.Archive.GetEntry(entryPath)
            ?? _doc.Archive.GetEntry($"BinData/{binRef}");
        entry?.Delete();
    }

    /// <summary>
    /// Remove inline memo markers (memoBegin/memoEnd) from body text.
    /// </summary>
    private void CleanupMemoMarkers(XElement memoElement)
    {
        var memoId = memoElement.Attribute("id")?.Value;
        if (memoId == null) return;

        foreach (var sec in _doc.Sections)
        {
            foreach (var ctrl in sec.Root.Descendants(HwpxNs.Hp + "ctrl").ToList())
            {
                var fieldBegin = ctrl.Element(HwpxNs.Hp + "fieldBegin");
                if (fieldBegin?.Attribute("type")?.Value == "MEMO"
                    && fieldBegin.Attribute("instId")?.Value == memoId)
                {
                    ctrl.Remove();
                }
                var fieldEnd = ctrl.Element(HwpxNs.Hp + "fieldEnd");
                if (fieldEnd?.Attribute("type")?.Value == "MEMO")
                {
                    ctrl.Remove();
                }
            }
        }
    }

    /// <summary>Remove page background (pageBorderFill) from all sections.</summary>
    private string? RemovePageBackground()
    {
        foreach (var sec in _doc.Sections)
        {
            var pageBf = sec.Root.Descendants(HwpxNs.Hp + "pageBorderFill")
                .Where(e => e.Attribute("type")?.Value == "BOTH")
                .ToList();
            foreach (var bf in pageBf) bf.Remove();
            _dirty = true;
            SaveSection(sec.Root);
        }
        return null;
    }

    /// <summary>Add image watermark via pageBorderFill + imgBrush (Plan 98).</summary>
    private XElement CreateWatermark(Dictionary<string, string>? props)
    {
        var path = props?.GetValueOrDefault("path") ?? props?.GetValueOrDefault("src")
            ?? throw new CliException("watermark requires 'path' or 'src' property");
        if (!File.Exists(path))
            throw new CliException($"Watermark image not found: {path}");

        var bright = props?.GetValueOrDefault("bright") ?? "70";
        var contrast = props?.GetValueOrDefault("contrast") ?? "-50";

        // 1. Add image to BinData
        var imageBytes = File.ReadAllBytes(path);
        var ext = Path.GetExtension(path).TrimStart('.').ToLowerInvariant();
        if (ext == "jpg") ext = "jpeg";
        var mediaType = ext switch { "png" => "image/png", "jpeg" => "image/jpeg", _ => $"image/{ext}" };
        var imageId = GetNextImageId();
        var binFileName = $"image{imageId}.{ext}";

        // BinData lives at ZIP root (BinData/), NOT under Contents/.
        // Hancom resolves manifest href relative to ZIP root, not content.hpf location.
        var binEntry = _doc.Archive.CreateEntry($"BinData/{binFileName}",
            System.IO.Compression.CompressionLevel.Optimal);
        using (var binStream = binEntry.Open())
            binStream.Write(imageBytes, 0, imageBytes.Length);
        RegisterImageInManifest($"image{imageId}", $"BinData/{binFileName}", mediaType);

        // 2. Create borderFill with imgBrush (golden template pattern)
        var bfId = NextBorderFillId();
        var bf = new XElement(HwpxNs.Hh + "borderFill",
            new XAttribute("id", bfId),
            new XAttribute("threeD", "0"),
            new XAttribute("shadow", "0"),
            new XAttribute("centerLine", "NONE"),
            new XAttribute("breakCellSeparateLine", "0"),
            new XElement(HwpxNs.Hh + "slash",
                new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
            new XElement(HwpxNs.Hh + "backSlash",
                new XAttribute("type", "NONE"), new XAttribute("Crooked", "0"), new XAttribute("isCounter", "0")),
            MakeBorder("leftBorder", "NONE", "0.1 mm", "#000000"),
            MakeBorder("rightBorder", "NONE", "0.1 mm", "#000000"),
            MakeBorder("topBorder", "NONE", "0.1 mm", "#000000"),
            MakeBorder("bottomBorder", "NONE", "0.1 mm", "#000000"),
            MakeBorder("diagonal", "SOLID", "0.1 mm", "#000000"),
            new XElement(HwpxNs.Hc + "fillBrush",
                new XElement(HwpxNs.Hc + "imgBrush",
                    new XAttribute("mode", "TOTAL"),
                    new XElement(HwpxNs.Hc + "img",
                        new XAttribute("binaryItemIDRef", $"image{imageId}"),
                        new XAttribute("bright", bright),
                        new XAttribute("contrast", contrast),
                        new XAttribute("effect", "REAL_PIC"),
                        new XAttribute("alpha", "0")))));

        var container = _doc.Header!.Root!.Descendants(HwpxNs.Hh + "borderFills").FirstOrDefault();
        container?.Add(bf);
        if (container != null)
            container.SetAttributeValue("itemCnt", container.Elements(HwpxNs.Hh + "borderFill").Count().ToString());
        SaveHeader();

        // 3. Set pageBorderFill on all sections
        foreach (var sec in _doc.Sections)
        {
            var secPr = sec.Root.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
            if (secPr == null) continue;

            // Update BOTH type pageBorderFill
            var pageBf = secPr.Elements(HwpxNs.Hp + "pageBorderFill")
                .FirstOrDefault(e => e.Attribute("type")?.Value == "BOTH");
            if (pageBf != null)
                pageBf.SetAttributeValue("borderFillIDRef", bfId);
            else
                secPr.Add(new XElement(HwpxNs.Hp + "pageBorderFill",
                    new XAttribute("type", "BOTH"),
                    new XAttribute("borderFillIDRef", bfId),
                    new XAttribute("textBorder", "PAPER"),
                    new XAttribute("headerInside", "0"),
                    new XAttribute("footerInside", "0"),
                    new XAttribute("fillArea", "PAPER"),
                    new XElement(HwpxNs.Hp + "offset",
                        new XAttribute("left", "1417"),
                        new XAttribute("right", "1417"),
                        new XAttribute("top", "1417"),
                        new XAttribute("bottom", "1417"))));

            SaveSection(sec.Root);
        }

        _dirty = true;
        // Return empty paragraph as placeholder (watermark is page-level, not inline)
        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("styleIDRef", "0"));
    }

    /// <summary>Remove TOC paragraphs (static V1 or field-based V2).</summary>
    private string? RemoveToc()
    {
        foreach (var sec in _doc.Sections)
        {
            // V2: field-based TOC (fieldBegin type="TABLEOFCONTENTS" ... fieldEnd)
            var fieldBegins = sec.Root.Descendants(HwpxNs.Hp + "fieldBegin")
                .Where(fb => fb.Attribute("type")?.Value == "TABLEOFCONTENTS")
                .ToList();
            if (fieldBegins.Count > 0)
            {
                foreach (var fb in fieldBegins)
                {
                    var instId = fb.Attribute("instId")?.Value;
                    // Remove all paragraphs between fieldBegin and matching fieldEnd
                    var beginCtrl = fb.Parent; // ctrl
                    var beginPara = beginCtrl?.Parent; // p
                    if (beginPara == null) continue;

                    var parasToRemove = new List<XElement> { beginPara };
                    var sibling = beginPara.ElementsAfterSelf(HwpxNs.Hp + "p").GetEnumerator();
                    while (sibling.MoveNext())
                    {
                        parasToRemove.Add(sibling.Current);
                        var endField = sibling.Current.Descendants(HwpxNs.Hp + "fieldEnd")
                            .FirstOrDefault(fe => fe.Attribute("type")?.Value == "TABLEOFCONTENTS");
                        if (endField != null) break;
                    }
                    foreach (var p in parasToRemove) p.Remove();
                }
                _dirty = true;
                SaveSection(sec.Root);
                continue;
            }

            // V1: static TOC (title "목차"/"차례" + entries until blank line)
            var tocStart = -1;
            var tocEnd = -1;
            var paras = sec.Root.Elements(HwpxNs.Hp + "p").ToList();
            for (int i = 0; i < paras.Count; i++)
            {
                var text = ExtractParagraphText(paras[i]).Trim();
                if (text == "목차" || text == "목 차" || text == "차례"
                    || text.StartsWith("[목차]") || text.StartsWith("[차례]"))
                {
                    tocStart = i;
                    // Consume entries until blank line (CreateStaticToc adds trailing blank)
                    for (int j = i + 1; j < paras.Count; j++)
                    {
                        var t = ExtractParagraphText(paras[j]).Trim();
                        tocEnd = j;
                        if (string.IsNullOrWhiteSpace(t)) break;
                    }
                    break;
                }
            }
            if (tocStart >= 0 && tocEnd >= tocStart)
            {
                for (int i = tocEnd; i >= tocStart; i--)
                    paras[i].Remove();
                _dirty = true;
                SaveSection(sec.Root);
            }
        }
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
    public string Move(string sourcePath, string? targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
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
    public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position)
    {
        var index = position?.Index;
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
    private readonly record struct PictureReferenceBox(int Width, int Height);

    private XElement CreatePicture(XElement parent, Dictionary<string, string>? props)
    {
        var path = props?.GetValueOrDefault("path")
            ?? props?.GetValueOrDefault("src")
            ?? throw new CliException("picture requires 'path' property");
        if (!File.Exists(path))
            throw new CliException($"Image file not found: {path}");

        var widthHwp = ParseDimensionToHwpUnit(props?.GetValueOrDefault("width") ?? "2in");
        var heightHwp = ParseDimensionToHwpUnit(props?.GetValueOrDefault("height") ?? "1in");
        var wrapMode = ResolvePictureWrap(props);
        var anchorMode = ResolvePictureAnchor(props, wrapMode);
        var (horizontalAlign, verticalAlign) = ParsePictureAlignment(props);
        var sectionRoot = ResolvePictureSectionRoot(parent);
        var anchorParagraph = ResolvePictureAnchorParagraph(parent);
        var referenceBox = ResolvePictureReferenceBox(anchorMode, sectionRoot, anchorParagraph);
        var (offsetX, offsetY) = ResolvePictureOffsets(props, anchorMode, widthHwp, heightHwp,
            referenceBox, horizontalAlign, verticalAlign);
        var treatAsChar = wrapMode == "char";
        var textWrap = MapPictureWrap(wrapMode);
        var lockValue = ParseBoolProp(props, "lock") ? "1" : "0";
        var zOrder = ResolvePictureZOrder(props, wrapMode).ToString();
        var relativeTarget = anchorMode == "page" ? "PAPER" : "PARA";

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

        // 3. Add image to ZIP — BinData/ at archive root (not Contents/BinData/)
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
            new XAttribute("zOrder", zOrder),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", textWrap),
            new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", lockValue),
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
                new XAttribute("treatAsChar", treatAsChar ? "1" : "0"), new XAttribute("affectLSpacing", "0"),
                new XAttribute("flowWithText", "1"), new XAttribute("allowOverlap", "0"),
                new XAttribute("holdAnchorAndSO", "0"),
                new XAttribute("vertRelTo", relativeTarget), new XAttribute("horzRelTo", relativeTarget),
                new XAttribute("vertAlign", "TOP"), new XAttribute("horzAlign", "LEFT"),
                new XAttribute("vertOffset", offsetY), new XAttribute("horzOffset", offsetX)),
            new XElement(HwpxNs.Hp + "outMargin",
                new XAttribute("left", "0"), new XAttribute("right", "0"),
                new XAttribute("top", "0"), new XAttribute("bottom", "0"))
        );
    }

    private static string ResolvePictureWrap(Dictionary<string, string>? props)
    {
        var rawWrap = props?.GetValueOrDefault("wrap")
            ?? props?.GetValueOrDefault("textwrap")
            ?? "";
        var normalizedWrap = rawWrap.Trim().ToLowerInvariant() switch
        {
            "char" or "inline" => "char",
            "square" => "square",
            "front" or "infront" or "in_front" => "front",
            "behind" or "back" => "behind",
            "topbottom" or "top_bottom" or "top-and-bottom" or "tight" or "wrap" => "topbottom",
            _ => "char"
        };
        if (normalizedWrap != "char")
            return normalizedWrap;

        var anchor = props?.GetValueOrDefault("anchor")?.Trim().ToLowerInvariant();
        var hasPositioningProps = props != null && (props.ContainsKey("x") || props.ContainsKey("y")
            || props.ContainsKey("halign") || props.ContainsKey("valign"));
        return (anchor is "page" or "para") || hasPositioningProps ? "topbottom" : "char";
    }

    private static string ResolvePictureAnchor(Dictionary<string, string>? props, string wrapMode)
    {
        var anchor = props?.GetValueOrDefault("anchor")?.Trim().ToLowerInvariant();
        if (anchor is "page" or "paper")
            return "page";
        if (anchor == "para")
            return "para";
        return wrapMode == "char" ? "para" : "para";
    }

    private static (string Horizontal, string Vertical) ParsePictureAlignment(Dictionary<string, string>? props)
    {
        var horizontal = props?.GetValueOrDefault("halign")?.Trim().ToLowerInvariant() switch
        {
            "center" or "middle" => "center",
            "right" => "right",
            _ => "left"
        };
        var vertical = props?.GetValueOrDefault("valign")?.Trim().ToLowerInvariant() switch
        {
            "middle" or "center" => "middle",
            "bottom" => "bottom",
            _ => "top"
        };
        return (horizontal, vertical);
    }

    private XElement ResolvePictureSectionRoot(XElement parent)
    {
        if (parent.Name == HwpxNs.Hs + "sec")
            return parent;
        var sectionRoot = parent.AncestorsAndSelf()
            .FirstOrDefault(e => e.Name == HwpxNs.Hs + "sec");
        return sectionRoot ?? _doc.PrimarySection.Root;
    }

    private static XElement? ResolvePictureAnchorParagraph(XElement parent)
    {
        if (parent.Name == HwpxNs.Hp + "p")
            return parent;
        if (parent.Name == HwpxNs.Hp + "run" && parent.Parent?.Name == HwpxNs.Hp + "p")
            return parent.Parent;
        return null;
    }

    private static PictureReferenceBox ResolvePictureReferenceBox(string anchorMode, XElement sectionRoot,
        XElement? anchorParagraph)
    {
        var secPr = sectionRoot.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
        var pagePr = secPr?.Element(HwpxNs.Hp + "pagePr");
        var margin = pagePr?.Element(HwpxNs.Hp + "margin");
        var pageWidth = (int?)pagePr?.Attribute("width") ?? 59528;
        var pageHeight = (int?)pagePr?.Attribute("height") ?? 84186;
        var marginLeft = (int?)margin?.Attribute("left") ?? 8504;
        var marginRight = (int?)margin?.Attribute("right") ?? 8504;
        var marginTop = (int?)margin?.Attribute("top") ?? 5668;
        var marginBottom = (int?)margin?.Attribute("bottom") ?? 4252;

        if (anchorMode == "page")
            return new PictureReferenceBox(pageWidth, pageHeight);

        var bodyWidth = Math.Max(pageWidth - marginLeft - marginRight, 0);
        var bodyHeight = Math.Max(pageHeight - marginTop - marginBottom, 0);
        if (anchorParagraph == null)
            return new PictureReferenceBox(bodyWidth, bodyHeight);
        return new PictureReferenceBox(bodyWidth, bodyHeight);
    }

    private static (int X, int Y) ResolvePictureOffsets(Dictionary<string, string>? props, string anchorMode,
        int widthHwp, int heightHwp, PictureReferenceBox referenceBox, string horizontalAlign, string verticalAlign)
    {
        var explicitX = ParsePictureOffset(props?.GetValueOrDefault("x"));
        var explicitY = ParsePictureOffset(props?.GetValueOrDefault("y"));

        var baseX = horizontalAlign switch
        {
            "center" => (referenceBox.Width - widthHwp) / 2,
            "right" => referenceBox.Width - widthHwp,
            _ => 0
        };

        var baseY = 0;
        if (anchorMode == "page")
        {
            baseY = verticalAlign switch
            {
                "middle" => (referenceBox.Height - heightHwp) / 2,
                "bottom" => referenceBox.Height - heightHwp,
                _ => 0
            };
        }

        return (baseX + explicitX, baseY + explicitY);
    }

    private static int ParsePictureOffset(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return 0;
        return ParseDimensionToHwpUnit(value);
    }

    private static string MapPictureWrap(string wrapMode) => wrapMode switch
    {
        "square" => "SQUARE",
        "front" => "IN_FRONT_OF_TEXT",
        "behind" => "BEHIND_TEXT",
        _ => "TOP_AND_BOTTOM"
    };

    private static bool ParseBoolProp(Dictionary<string, string>? props, string key)
    {
        var value = props?.GetValueOrDefault(key);
        return value != null && (value.Equals("true", StringComparison.OrdinalIgnoreCase) || value == "1");
    }

    private static int ResolvePictureZOrder(Dictionary<string, string>? props, string wrapMode)
    {
        if (int.TryParse(props?.GetValueOrDefault("z"), out var zOrder))
            return zOrder;
        return wrapMode switch
        {
            "front" => 1,
            "behind" => 0,
            _ => 0
        };
    }

    // ==================== Hyperlink ====================

    private static readonly HashSet<string> SafeUrlSchemes = new(StringComparer.OrdinalIgnoreCase)
    {
        "http", "https", "mailto", "ftp", "ftps", "tel"
    };

    private static void ValidateUrlScheme(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            throw new ArgumentException("Hyperlink URL must not be empty.");

        // Reject obvious local/UNC paths
        if (url.StartsWith('\\') || url.StartsWith("//") ||
            (url.Length >= 2 && url[1] == ':'))
            throw new ArgumentException(
                $"Local or UNC path not allowed as hyperlink target: '{url}'.");

        if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
        {
            if (!SafeUrlSchemes.Contains(uri.Scheme))
                throw new ArgumentException(
                    $"Unsafe URL scheme '{uri.Scheme}' in '{url}'. " +
                    $"Allowed schemes: {string.Join(", ", SafeUrlSchemes)}.");
        }
        else
        {
            // Non-parseable as absolute URI — reject unless it looks like
            // a fragment (#anchor) or same-document reference
            if (!url.StartsWith('#'))
                throw new ArgumentException(
                    $"Hyperlink target must be an absolute URL with allowed scheme " +
                    $"({string.Join(", ", SafeUrlSchemes)}) or a fragment reference: '{url}'.");
        }
    }

    /// <summary>
    /// Create a hyperlink using the OWPML fieldBegin/fieldEnd pattern (3-run structure).
    /// Golden template based on OWPML schema + python-hwpx implementation.
    /// Props: url/href (required), text (default=url).
    /// </summary>
    private XElement CreateHyperlink(Dictionary<string, string>? props)
    {
        var url = props?.GetValueOrDefault("url") ?? props?.GetValueOrDefault("href")
            ?? throw new CliException("hyperlink requires 'url' property");

        // Plan 99.9.E3: Safe URL scheme whitelist
        ValidateUrlScheme(url);

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
        // CRITICAL: Hancom uses POSITIONAL indexing (array index), not id-based lookup.
        // Append at END of container so position matches the new ID.
        var lastCharPr = _doc.Header?.Root?.Descendants(HwpxNs.Hh + "charPr").LastOrDefault();
        if (lastCharPr != null)
        {
            var container = lastCharPr.Parent!;
            container.Add(newCharPr);
            var count = container.Elements(HwpxNs.Hh + "charPr").Count();
            container.SetAttributeValue("itemCnt", count.ToString());
        }

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

    /// <summary>Create column break — changes colCount via colPr (Plan 96).</summary>
    private XElement CreateColumnBreak(Dictionary<string, string>? props)
    {
        var cols = int.Parse(props?.GetValueOrDefault("cols") ?? "2");
        var gap = props?.GetValueOrDefault("gap") ?? (cols > 1 ? "2268" : "0");
        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "colPr",
                        new XAttribute("id", ""),
                        new XAttribute("type", "NEWSPAPER"),
                        new XAttribute("layout", "LEFT"),
                        new XAttribute("colCount", cols.ToString()),
                        new XAttribute("sameSz", cols > 1 ? "1" : "1"),
                        new XAttribute("sameGap", gap)))));
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
    /// Add a memo to the section-level memogroup container and attach a
    /// fieldBegin/fieldEnd anchor to the last paragraph so Hancom displays it.
    /// HWPX memos live in: section > hp:memogroup > hp:memo > hp:paraList > hp:p
    /// The anchor uses fieldBegin type="MEMO" with parameters linking to the memo.
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

        // Attach fieldBegin/fieldEnd anchor to last paragraph
        var lastPara = section.Elements(HwpxNs.Hp + "p").LastOrDefault();
        if (lastPara != null)
        {
            var fieldId = Guid.NewGuid().ToString("N");
            var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // fieldBegin run
            var runBegin = new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldBegin",
                        new XAttribute("id", fieldId),
                        new XAttribute("type", "MEMO"),
                        new XAttribute("editable", "true"),
                        new XAttribute("dirty", "false"),
                        new XAttribute("fieldid", fieldId),
                        new XElement(HwpxNs.Hp + "parameters",
                            new XAttribute("count", "5"),
                            new XAttribute("name", ""),
                            new XElement(HwpxNs.Hp + "stringParam",
                                new XAttribute("name", "ID"), memoId),
                            new XElement(HwpxNs.Hp + "integerParam",
                                new XAttribute("name", "Number"), "1"),
                            new XElement(HwpxNs.Hp + "stringParam",
                                new XAttribute("name", "CreateDateTime"), now),
                            new XElement(HwpxNs.Hp + "stringParam",
                                new XAttribute("name", "Author"), ""),
                            new XElement(HwpxNs.Hp + "stringParam",
                                new XAttribute("name", "MemoShapeID"), memoShapeId)),
                        new XElement(HwpxNs.Hp + "subList",
                            new XAttribute("id", $"memo-field-{memoId}"),
                            new XAttribute("textDirection", "HORIZONTAL"),
                            new XAttribute("lineWrap", "BREAK"),
                            new XAttribute("vertAlign", "TOP"),
                            new XElement(HwpxNs.Hp + "p",
                                new XAttribute("id", NewId()),
                                new XAttribute("paraPrIDRef", "0"),
                                new XAttribute("styleIDRef", "0"),
                                new XAttribute("pageBreak", "0"),
                                new XAttribute("columnBreak", "0"),
                                new XAttribute("merged", "0"),
                                new XElement(HwpxNs.Hp + "run",
                                    new XAttribute("charPrIDRef", "0"),
                                    new XElement(HwpxNs.Hp + "t", memoId)))))));

            // fieldEnd run
            var runEnd = new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldEnd",
                        new XAttribute("beginIDRef", fieldId),
                        new XAttribute("fieldid", fieldId))));

            // Insert at beginning and end of paragraph
            lastPara.AddFirst(runBegin);
            lastPara.Add(runEnd);
        }

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

    // ==================== Equation ====================

    /// <summary>
    /// Create an equation paragraph. Uses hp:equation (NOT hp:eqEdit — eqEdit is HWP5 class name).
    /// Structure: hp:p > hp:run > hp:equation (ShapeObject) > hp:script.
    /// Props: script (required), font, mode (LINE|CHAR), color.
    /// Based on hwp_recog/18 (hwpxlib confirmed structure).
    /// </summary>
    private XElement CreateEquation(Dictionary<string, string>? props)
    {
        var script = props?.GetValueOrDefault("script")
            ?? props?.GetValueOrDefault("formula")
            ?? throw new CliException("equation requires 'script' or 'formula' property")
                { Code = "invalid_prop" };

        // Golden template analysis (test_eq_golden.hwpx):
        // version="Equation Version 60", font="HancomEQN", lineMode="CHAR"
        // numberingType="EQUATION", textWrap="TOP_AND_BOTTOM", textFlow="BOTH_SIDES"
        // horzRelTo="PARA", vertAlign="TOP", outMargin top/bottom="0"
        // NO <hp:inMargin>, HAS <hp:shapeComment>, <hp:script xml:space="preserve">
        var font = props?.GetValueOrDefault("font") ?? "HancomEQN";
        var lineMode = props?.GetValueOrDefault("mode")?.ToUpperInvariant() ?? "CHAR";
        var textColor = props?.GetValueOrDefault("color") ?? "#000000";
        var baseUnit = props?.GetValueOrDefault("baseunit") ?? "1000";

        // Default sz: Hancom auto-calculates based on equation complexity.
        // Golden examples: simple eq ~3700x975, fraction ~5500x2250, sigma ~2000x2700
        // Use moderate defaults; Hancom recalculates on open.
        var width = props?.GetValueOrDefault("width") ?? "5000";
        var height = props?.GetValueOrDefault("height") ?? "2000";

        if (int.TryParse(width, out _) == false)
            width = ParseDimensionToHwpUnit(width).ToString();
        if (int.TryParse(height, out _) == false)
            height = ParseDimensionToHwpUnit(height).ToString();

        var equation = new XElement(HwpxNs.Hp + "equation",
            new XAttribute("id", NewId()),
            new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "EQUATION"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"),
            new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"),
            new XAttribute("dropcapstyle", "None"),
            new XAttribute("version", "Equation Version 60"),
            new XAttribute("baseLine", "85"),
            new XAttribute("textColor", textColor),
            new XAttribute("baseUnit", baseUnit),
            new XAttribute("lineMode", lineMode),
            new XAttribute("font", font),
            // ShapeObject children (order matches golden template)
            new XElement(HwpxNs.Hp + "sz",
                new XAttribute("width", width),
                new XAttribute("widthRelTo", "ABSOLUTE"),
                new XAttribute("height", height),
                new XAttribute("heightRelTo", "ABSOLUTE"),
                new XAttribute("protect", "0")),
            new XElement(HwpxNs.Hp + "pos",
                new XAttribute("treatAsChar", "1"),
                new XAttribute("affectLSpacing", "0"),
                new XAttribute("flowWithText", "1"),
                new XAttribute("allowOverlap", "0"),
                new XAttribute("holdAnchorAndSO", "0"),
                new XAttribute("vertRelTo", "PARA"),
                new XAttribute("horzRelTo", "PARA"),
                new XAttribute("vertAlign", "TOP"),
                new XAttribute("horzAlign", "LEFT"),
                new XAttribute("vertOffset", "0"),
                new XAttribute("horzOffset", "0")),
            new XElement(HwpxNs.Hp + "outMargin",
                new XAttribute("left", "56"),
                new XAttribute("right", "56"),
                new XAttribute("top", "0"),
                new XAttribute("bottom", "0")),
            // NO <hp:inMargin> — golden template doesn't have it
            new XElement(HwpxNs.Hp + "shapeComment", "수식입니다."),
            // xml:space="preserve" + trailing newline matches golden
            new XElement(HwpxNs.Hp + "script",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                script + "\n"));

        // Wrap: hp:p > hp:run > hp:equation + hp:t (empty, matches golden)
        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XAttribute("pageBreak", "0"),
            new XAttribute("columnBreak", "0"),
            new XAttribute("merged", "0"),
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", "0"),
                equation,
                new XElement(HwpxNs.Hp + "t")));
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

    // ==================== Fields (golden 기반: CLICK_HERE, SUMMERY, PATH) ====================

    private XElement CreateFormField(Dictionary<string, string>? props)
    {
        var formFieldType = props?.GetValueOrDefault("formfieldtype")
            ?? props?.GetValueOrDefault("type")
            ?? "text";

        return formFieldType.ToLowerInvariant() switch
        {
            "text" or "clickhere" or "click_here" => CreateField(MergeProps(props, "type", "CLICK_HERE")),
            "checkbox" or "check" => CreateField(MergeProps(props, "type", "CHECKBOX")),
            "dropdown" or "drop" => CreateField(MergeProps(props, "type", "DROPDOWN")),
            _ => throw new CliException($"Unsupported HWPX form field type: {formFieldType}")
            {
                Code = "invalid_prop",
                Suggestion = "Use type=text, type=checkbox, or type=dropdown."
            }
        };
    }

    /// <summary>
    /// Create a field (fieldBegin + display text + fieldEnd).
    /// Props: "type" (required), "text" (display), "command", "direction"
    /// Golden structure: hp:ctrl > hp:fieldBegin with parameters > hp:run > hp:t > hp:ctrl > hp:fieldEnd
    /// </summary>
    private XElement CreateField(Dictionary<string, string>? props)
    {
        var fieldType = props?.GetValueOrDefault("type")?.ToUpperInvariant()
            ?? throw new CliException("field requires 'type' property") { Code = "invalid_prop" };
        var displayText = props?.GetValueOrDefault("text") ?? GetDefaultFieldText(fieldType, props);
        var fieldId = NewId();
        var instId = NewId();
        var fieldName = props?.GetValueOrDefault("name") ?? "";
        var metaTag = props?.GetValueOrDefault("metatag") ?? props?.GetValueOrDefault("metaTag") ?? "";

        // Build fieldBegin with golden-accurate attributes
        var editable = fieldType is "CLICK_HERE" or "CHECKBOX" or "DROPDOWN"
            ? "1"
            : (props?.GetValueOrDefault("editable") ?? "0");
        var dirty = fieldType is "CLICK_HERE" or "CHECKBOX" or "DROPDOWN" ? "1" : "0";

        var fieldBeginEl = new XElement(HwpxNs.Hp + "fieldBegin",
            new XAttribute("id", instId),
            new XAttribute("type", fieldType),
            new XAttribute("name", fieldName),
            new XAttribute("editable", editable),
            new XAttribute("dirty", dirty),
            new XAttribute("zorder", "-1"),
            new XAttribute("fieldid", fieldId),
            new XAttribute("metaTag", metaTag));

        // Add parameters based on field type (golden structure)
        var parameters = BuildFieldParameters(fieldType, props);
        if (parameters != null)
            fieldBeginEl.Add(parameters);

        // Golden structure: separate runs for fieldBegin, text, fieldEnd
        // CLICK_HERE display text uses red italic charPr (golden: textColor=#FF0000 + italic)
        var textCharPr = "0";
        if (fieldType == "CLICK_HERE")
            textCharPr = EnsureClickHereCharPr().ToString();

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl", fieldBeginEl)),
            new XElement(HwpxNs.Hp + "run",
                new XAttribute("charPrIDRef", textCharPr),
                new XElement(HwpxNs.Hp + "t", displayText)),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldEnd",
                        new XAttribute("beginIDRef", instId),
                        new XAttribute("fieldid", fieldId))),
                new XElement(HwpxNs.Hp + "t")));
    }

    private XElement? BuildFieldParameters(string fieldType, Dictionary<string, string>? props)
    {
        return fieldType switch
        {
            "CLICK_HERE" => new XElement(HwpxNs.Hp + "parameters",
                new XAttribute("cnt", props != null && props.ContainsKey("maxlength") ? "4" : "3"), new XAttribute("name", ""),
                new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "9"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Command"),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    $"Clickhere:set:66:Direction:wstring:{GetClickHereDirection(props).Length}:{GetClickHereDirection(props)} HelpState:wstring:0:  "),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Direction"),
                    GetClickHereDirection(props)),
                CreateOptionalIntegerParam("MaxLength", props?.GetValueOrDefault("maxlength"))),

            "CHECKBOX" => new XElement(HwpxNs.Hp + "parameters",
                new XAttribute("cnt", "4"), new XAttribute("name", ""),
                new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "9"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Command"),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    $"CheckBox:set:7:Checked:int:{(IsChecked(props) ? "1" : "0")} Label:wstring:{GetCheckboxLabel(props).Length}:{GetCheckboxLabel(props)}"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Checked"),
                    IsChecked(props) ? "1" : "0"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Label"),
                    GetCheckboxLabel(props))),

            "DROPDOWN" => new XElement(HwpxNs.Hp + "parameters",
                new XAttribute("cnt", "4"), new XAttribute("name", ""),
                new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "9"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Command"),
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    $"Dropdown:set:12:Items:wstring:{GetDropdownItemsValue(props).Length}:{GetDropdownItemsValue(props)} SelectedIndex:int:{ResolveDropdownSelectedIndex(props)}"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Items"),
                    GetDropdownItemsValue(props)),
                new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "SelectedIndex"),
                    ResolveDropdownSelectedIndex(props).ToString())),

            "SUMMERY" => new XElement(HwpxNs.Hp + "parameters",
                new XAttribute("cnt", "3"), new XAttribute("name", ""),
                new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "8"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Command"),
                    props?.GetValueOrDefault("command") ?? "$createtime"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Property"),
                    props?.GetValueOrDefault("command") ?? "$createtime")),

            "PATH" => new XElement(HwpxNs.Hp + "parameters",
                new XAttribute("cnt", "3"), new XAttribute("name", ""),
                new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "8"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Command"),
                    props?.GetValueOrDefault("format") ?? "$P$F"),
                new XElement(HwpxNs.Hp + "stringParam", new XAttribute("name", "Format"),
                    props?.GetValueOrDefault("format") ?? "$P$F")),

            _ => null
        };
    }

    private static string GetDefaultFieldText(string type, Dictionary<string, string>? props) => type switch
    {
        "DATE" => DateTime.Now.ToString("yyyy-MM-dd"),
        "PATH" => "(filepath)",
        "CHECKBOX" => IsChecked(props)
            ? (props?.GetValueOrDefault("checkedtext") ?? "☑")
            : (props?.GetValueOrDefault("uncheckedtext") ?? "☐"),
        "DROPDOWN" => ResolveDropdownValue(props),
        "CLICK_HERE" => props?.GetValueOrDefault("text")
            ?? props?.GetValueOrDefault("defaultvalue")
            ?? props?.GetValueOrDefault("direction")
            ?? "이곳을 마우스로 누르고 내용을 입력하세요.",
        "SUMMERY" => props?.GetValueOrDefault("command") switch
        {
            "$lastsaveby" => "(author)",
            "$createtime" => DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
            _ => "(summary)"
        },
        _ => $"({type})"
    };

    /// <summary>Ensure a red italic charPr exists for CLICK_HERE display text. Returns its ID.</summary>
    private int EnsureClickHereCharPr()
    {
        if (_doc.Header?.Root == null) return 0;
        // Check if a red italic charPr already exists
        var charPrs = _doc.Header.Root.Descendants(HwpxNs.Hh + "charPr").ToList();
        foreach (var cp in charPrs)
        {
            if (cp.Attribute("textColor")?.Value == "#FF0000"
                && cp.Element(HwpxNs.Hh + "italic") != null)
                return int.TryParse(cp.Attribute("id")?.Value, out var existingId) ? existingId : 0;
        }
        // Create new charPr: clone charPr 0 + set red + italic
        var baseCharPr = charPrs.FirstOrDefault(c => c.Attribute("id")?.Value == "0");
        if (baseCharPr == null) return 0;
        var newCharPr = new XElement(baseCharPr);
        var newId = charPrs.Select(c => int.TryParse(c.Attribute("id")?.Value, out var i) ? i : 0).Max() + 1;
        newCharPr.SetAttributeValue("id", newId.ToString());
        newCharPr.SetAttributeValue("textColor", "#FF0000");
        // Add italic if missing
        if (newCharPr.Element(HwpxNs.Hh + "italic") == null)
            newCharPr.Add(new XElement(HwpxNs.Hh + "italic"));
        var container = baseCharPr.Parent!;
        container.Add(newCharPr);
        container.SetAttributeValue("itemCnt", container.Elements(HwpxNs.Hh + "charPr").Count().ToString());
        SaveHeader();
        return newId;
    }

    private static string GetClickHereDirection(Dictionary<string, string>? props)
    {
        return props?.GetValueOrDefault("direction")
            ?? props?.GetValueOrDefault("defaultvalue")
            ?? props?.GetValueOrDefault("text")
            ?? "이곳을 마우스로 누르고 내용을 입력하세요.";
    }

    private static bool IsChecked(Dictionary<string, string>? props)
    {
        var raw = props?.GetValueOrDefault("checked")
            ?? props?.GetValueOrDefault("value")
            ?? props?.GetValueOrDefault("text");
        return raw != null && ParseHelpers.IsTruthy(raw);
    }

    private static string GetCheckboxLabel(Dictionary<string, string>? props)
    {
        return props?.GetValueOrDefault("label")
            ?? props?.GetValueOrDefault("name")
            ?? "";
    }

    private static string[] ParseDropdownOptions(Dictionary<string, string>? props)
    {
        var raw = props?.GetValueOrDefault("options")
            ?? props?.GetValueOrDefault("items")
            ?? "";
        return raw.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    }

    private static string GetDropdownItemsValue(Dictionary<string, string>? props)
    {
        return string.Join("|", ParseDropdownOptions(props));
    }

    private static int ResolveDropdownSelectedIndex(Dictionary<string, string>? props)
    {
        var options = ParseDropdownOptions(props);
        if (options.Length == 0) return 0;

        if (int.TryParse(props?.GetValueOrDefault("selectedindex"), out var parsedIndex))
            return Math.Clamp(parsedIndex, 0, options.Length - 1);

        var selectedValue = props?.GetValueOrDefault("value") ?? props?.GetValueOrDefault("text");
        if (!string.IsNullOrEmpty(selectedValue))
        {
            var matchedIndex = Array.FindIndex(options, option =>
                string.Equals(option, selectedValue, StringComparison.Ordinal));
            if (matchedIndex >= 0) return matchedIndex;
        }

        return 0;
    }

    private static string ResolveDropdownValue(Dictionary<string, string>? props)
    {
        var explicitValue = props?.GetValueOrDefault("text") ?? props?.GetValueOrDefault("value");
        if (!string.IsNullOrEmpty(explicitValue))
            return explicitValue;

        var options = ParseDropdownOptions(props);
        if (options.Length == 0) return "(dropdown)";

        return options[ResolveDropdownSelectedIndex(props)];
    }

    private static XElement? CreateOptionalIntegerParam(string name, string? value)
    {
        return int.TryParse(value, out var parsed) && parsed > 0
            ? new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", name), parsed.ToString())
            : null;
    }

    private static Dictionary<string, string> MergeProps(Dictionary<string, string>? props, string key, string value)
    {
        var result = props != null ? new Dictionary<string, string>(props, StringComparer.OrdinalIgnoreCase) : new(StringComparer.OrdinalIgnoreCase);
        result[key] = value;
        return result;
    }

    private static Dictionary<string, string> MergeProps(Dictionary<string, string>? props, string key1, string value1, string key2, string value2)
    {
        var result = MergeProps(props, key1, value1);
        result[key2] = value2;
        return result;
    }

    // ==================== Style ====================

    private XElement CreateStyleElement(Dictionary<string, string> props)
    {
        var name = props.GetValueOrDefault("name")
            ?? throw new CliException("style requires 'name'") { Code = "invalid_prop" };
        var engName = props.GetValueOrDefault("engname") ?? name;
        var type = props.GetValueOrDefault("type")?.ToUpperInvariant() ?? "PARA";

        var styles = _doc.Header!.Root!.Descendants(HwpxNs.Hh + "style").ToList();
        var maxId = styles.Select(s => int.TryParse(s.Attribute("id")?.Value, out var i) ? i : 0)
            .DefaultIfEmpty(0).Max();
        var newId = (maxId + 1).ToString();

        var style = new XElement(HwpxNs.Hh + "style",
            new XAttribute("id", newId),
            new XAttribute("type", type),
            new XAttribute("name", name),
            new XAttribute("engName", engName),
            new XAttribute("paraPrIDRef", CloneParaPrForNewStyle().ToString()),
            new XAttribute("charPrIDRef", CloneCharPrForNewStyle().ToString()),
            new XAttribute("nextStyleIDRef", newId));

        var container = _doc.Header.Root.Descendants(HwpxNs.Hh + "styles").FirstOrDefault();
        if (container != null)
        {
            container.Add(style);
            container.SetAttributeValue("itemCnt",
                container.Elements(HwpxNs.Hh + "style").Count().ToString());
        }
        SaveHeader();
        return style;
    }

    // ==================== Shapes (golden54 기반) ====================

    /// <summary>Create common shape child elements matching golden template structure.</summary>
    private XElement[] CreateShapeCommonChildren(int w, int h, Dictionary<string, string>? props)
    {
        var x = int.TryParse(props?.GetValueOrDefault("x"), out var xv) ? xv : 0;
        var y = int.TryParse(props?.GetValueOrDefault("y"), out var yv) ? yv : 0;
        var treatAsChar = props?.GetValueOrDefault("wrap")?.Equals("char", StringComparison.OrdinalIgnoreCase) == true;
        var cx = w / 2; var cy = h / 2;

        XElement identity(string name) => new XElement(HwpxNs.Hc + name,
            new XAttribute("e1", "1"), new XAttribute("e2", "0"), new XAttribute("e3", "0"),
            new XAttribute("e4", "0"), new XAttribute("e5", "1"), new XAttribute("e6", "0"));

        return new XElement[]
        {
            new XElement(HwpxNs.Hp + "offset", new XAttribute("x", "0"), new XAttribute("y", "0")),
            new XElement(HwpxNs.Hp + "orgSz", new XAttribute("width", w), new XAttribute("height", h)),
            new XElement(HwpxNs.Hp + "curSz", new XAttribute("width", "0"), new XAttribute("height", "0")),
            new XElement(HwpxNs.Hp + "flip", new XAttribute("horizontal", "0"), new XAttribute("vertical", "0")),
            new XElement(HwpxNs.Hp + "rotationInfo",
                new XAttribute("angle", "0"), new XAttribute("centerX", cx), new XAttribute("centerY", cy),
                new XAttribute("rotateimage", "1")),
            new XElement(HwpxNs.Hp + "renderingInfo", identity("transMatrix"), identity("scaMatrix"), identity("rotMatrix")),
        };
    }

    private XElement CreateShapeSzPosMargin(int w, int h, Dictionary<string, string>? props)
    {
        // Returns a container — caller extracts children
        var x = int.TryParse(props?.GetValueOrDefault("x"), out var xv) ? xv : 0;
        var y = int.TryParse(props?.GetValueOrDefault("y"), out var yv) ? yv : 0;
        var treatAsChar = props?.GetValueOrDefault("wrap")?.Equals("char", StringComparison.OrdinalIgnoreCase) == true;

        return new XElement("_tmp",
            new XElement(HwpxNs.Hp + "sz",
                new XAttribute("width", w), new XAttribute("widthRelTo", "ABSOLUTE"),
                new XAttribute("height", h), new XAttribute("heightRelTo", "ABSOLUTE"),
                new XAttribute("protect", "0")),
            new XElement(HwpxNs.Hp + "pos",
                new XAttribute("treatAsChar", treatAsChar ? "1" : "0"),
                new XAttribute("affectLSpacing", "0"), new XAttribute("flowWithText", "1"),
                new XAttribute("allowOverlap", treatAsChar ? "1" : "0"),
                new XAttribute("holdAnchorAndSO", "0"),
                new XAttribute("vertRelTo", "PARA"), new XAttribute("horzRelTo", "COLUMN"),
                new XAttribute("vertAlign", "TOP"), new XAttribute("horzAlign", "LEFT"),
                new XAttribute("vertOffset", y), new XAttribute("horzOffset", x)),
            new XElement(HwpxNs.Hp + "outMargin",
                new XAttribute("left", "0"), new XAttribute("right", "0"),
                new XAttribute("top", "0"), new XAttribute("bottom", "0")));
    }

    private static XElement CreateLineShapeElement(Dictionary<string, string>? props)
    {
        var color = props?.GetValueOrDefault("color") ?? props?.GetValueOrDefault("linecolor") ?? "#000000";
        var width = props?.GetValueOrDefault("linewidth") ?? "33";
        var style = props?.GetValueOrDefault("linestyle") ?? "SOLID";
        return new XElement(HwpxNs.Hp + "lineShape",
            new XAttribute("color", color), new XAttribute("width", width),
            new XAttribute("style", style), new XAttribute("endCap", "FLAT"),
            new XAttribute("headStyle", "NORMAL"), new XAttribute("tailStyle", "NORMAL"),
            new XAttribute("headfill", "1"), new XAttribute("tailfill", "1"),
            new XAttribute("headSz", "MEDIUM_MEDIUM"), new XAttribute("tailSz", "MEDIUM_MEDIUM"),
            new XAttribute("outlineStyle", "NORMAL"), new XAttribute("alpha", "0"));
    }

    private static XElement? CreateFillBrush(Dictionary<string, string>? props)
    {
        var fill = props?.GetValueOrDefault("fillcolor") ?? props?.GetValueOrDefault("fill");
        if (fill == null) return null;
        return new XElement(HwpxNs.Hc + "fillBrush",
            new XElement(HwpxNs.Hc + "winBrush",
                new XAttribute("faceColor", fill), new XAttribute("hatchColor", "#000000"),
                new XAttribute("alpha", "0")));
    }

    private XElement CreateLine(Dictionary<string, string>? props)
    {
        var w = int.TryParse(props?.GetValueOrDefault("width"), out var wv) ? wv : 20000;
        var h = 43; // golden: line height is minimal
        var common = CreateShapeCommonChildren(w, h, props);
        var szPosMargin = CreateShapeSzPosMargin(w, h, props);

        var line = new XElement(HwpxNs.Hp + "line",
            new XAttribute("id", NewId()), new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"), new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"), new XAttribute("dropcapstyle", "None"),
            new XAttribute("href", ""), new XAttribute("groupLevel", "0"),
            new XAttribute("instid", NewId()), new XAttribute("isReverseHV", "0"));
        line.Add(common);
        line.Add(CreateLineShapeElement(props));
        line.Add(new XElement(HwpxNs.Hp + "shadow",
            new XAttribute("type", "NONE"), new XAttribute("color", "#B2B2B2"),
            new XAttribute("offsetX", "0"), new XAttribute("offsetY", "0"), new XAttribute("alpha", "0")));
        line.Add(new XElement(HwpxNs.Hc + "startPt", new XAttribute("x", "0"), new XAttribute("y", h)));
        line.Add(new XElement(HwpxNs.Hc + "endPt", new XAttribute("x", w), new XAttribute("y", "0")));
        line.Add(szPosMargin.Elements());
        line.Add(new XElement(HwpxNs.Hp + "shapeComment", "선입니다."));

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"), line,
                new XElement(HwpxNs.Hp + "t", "")));
    }

    private XElement CreateRect(Dictionary<string, string>? props)
    {
        var w = int.TryParse(props?.GetValueOrDefault("width"), out var wv) ? wv : 20000;
        var h = int.TryParse(props?.GetValueOrDefault("height"), out var hv) ? hv : 10000;
        var text = props?.GetValueOrDefault("text");
        var common = CreateShapeCommonChildren(w, h, props);
        var szPosMargin = CreateShapeSzPosMargin(w, h, props);
        var fillBrush = CreateFillBrush(props);

        var rect = new XElement(HwpxNs.Hp + "rect",
            new XAttribute("id", NewId()), new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"), new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"), new XAttribute("dropcapstyle", "None"),
            new XAttribute("href", ""), new XAttribute("groupLevel", "0"),
            new XAttribute("instid", NewId()), new XAttribute("ratio", "0"));
        rect.Add(common);
        rect.Add(CreateLineShapeElement(props));
        if (fillBrush != null) rect.Add(fillBrush);
        rect.Add(new XElement(HwpxNs.Hp + "shadow",
            new XAttribute("type", "NONE"), new XAttribute("color", "#B2B2B2"),
            new XAttribute("offsetX", "0"), new XAttribute("offsetY", "0"), new XAttribute("alpha", "178")));
        // drawText (text inside shape)
        if (text != null)
        {
            rect.Add(new XElement(HwpxNs.Hp + "drawText",
                new XAttribute("lastWidth", w), new XAttribute("name", ""), new XAttribute("editable", "0"),
                new XElement(HwpxNs.Hp + "subList",
                    new XAttribute("id", ""), new XAttribute("textDirection", "HORIZONTAL"),
                    new XAttribute("lineWrap", "BREAK"), new XAttribute("vertAlign", "CENTER"),
                    new XAttribute("linkListIDRef", "0"), new XAttribute("linkListNextIDRef", "0"),
                    new XAttribute("textWidth", "0"), new XAttribute("textHeight", "0"),
                    new XAttribute("hasTextRef", "0"), new XAttribute("hasNumRef", "0"),
                    new XElement(HwpxNs.Hp + "p",
                        new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
                        new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"),
                            new XElement(HwpxNs.Hp + "t", text)))),
                new XElement(HwpxNs.Hp + "textMargin",
                    new XAttribute("left", "283"), new XAttribute("right", "283"),
                    new XAttribute("top", "283"), new XAttribute("bottom", "283"))));
        }
        // Corner points
        rect.Add(new XElement(HwpxNs.Hc + "pt0", new XAttribute("x", "0"), new XAttribute("y", "0")));
        rect.Add(new XElement(HwpxNs.Hc + "pt1", new XAttribute("x", w), new XAttribute("y", "0")));
        rect.Add(new XElement(HwpxNs.Hc + "pt2", new XAttribute("x", w), new XAttribute("y", h)));
        rect.Add(new XElement(HwpxNs.Hc + "pt3", new XAttribute("x", "0"), new XAttribute("y", h)));
        rect.Add(szPosMargin.Elements());

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"), rect,
                new XElement(HwpxNs.Hp + "t", "")));
    }

    private XElement CreateEllipse(Dictionary<string, string>? props)
    {
        var w = int.TryParse(props?.GetValueOrDefault("width"), out var wv) ? wv : 15000;
        var h = int.TryParse(props?.GetValueOrDefault("height"), out var hv) ? hv : 10000;
        var common = CreateShapeCommonChildren(w, h, props);
        var szPosMargin = CreateShapeSzPosMargin(w, h, props);
        var fillBrush = CreateFillBrush(props);
        var cx = w / 2; var cy = h / 2;

        var ellipse = new XElement(HwpxNs.Hp + "ellipse",
            new XAttribute("id", NewId()), new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"), new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"), new XAttribute("dropcapstyle", "None"),
            new XAttribute("href", ""), new XAttribute("groupLevel", "0"),
            new XAttribute("instid", NewId()),
            new XAttribute("intervalDirty", "0"), new XAttribute("hasArcPr", "0"),
            new XAttribute("arcType", "NORMAL"));
        ellipse.Add(common);
        ellipse.Add(CreateLineShapeElement(props));
        if (fillBrush != null) ellipse.Add(fillBrush);
        ellipse.Add(new XElement(HwpxNs.Hp + "shadow",
            new XAttribute("type", "NONE"), new XAttribute("color", "#B2B2B2"),
            new XAttribute("offsetX", "0"), new XAttribute("offsetY", "0"), new XAttribute("alpha", "178")));
        // Ellipse geometry
        ellipse.Add(new XElement(HwpxNs.Hc + "center", new XAttribute("x", cx), new XAttribute("y", cy)));
        ellipse.Add(new XElement(HwpxNs.Hc + "ax1", new XAttribute("x", w), new XAttribute("y", cy)));
        ellipse.Add(new XElement(HwpxNs.Hc + "ax2", new XAttribute("x", cx), new XAttribute("y", "0")));
        ellipse.Add(new XElement(HwpxNs.Hc + "start1", new XAttribute("x", "0"), new XAttribute("y", "0")));
        ellipse.Add(new XElement(HwpxNs.Hc + "end1", new XAttribute("x", "0"), new XAttribute("y", "0")));
        ellipse.Add(new XElement(HwpxNs.Hc + "start2", new XAttribute("x", "0"), new XAttribute("y", "0")));
        ellipse.Add(new XElement(HwpxNs.Hc + "end2", new XAttribute("x", "0"), new XAttribute("y", "0")));
        ellipse.Add(szPosMargin.Elements());

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"), ellipse,
                new XElement(HwpxNs.Hp + "t", "")));
    }

    private XElement CreateTextBox(Dictionary<string, string>? props)
    {
        // TextBox = rect with drawText and default fill
        var merged = new Dictionary<string, string>(props ?? new(), StringComparer.OrdinalIgnoreCase);
        if (!merged.ContainsKey("fill") && !merged.ContainsKey("fillcolor"))
            merged["fillcolor"] = "#FFFFFF";
        if (!merged.ContainsKey("text"))
            merged["text"] = "";
        return CreateRect(merged);
    }

    private XElement CreatePolygon(Dictionary<string, string>? props)
    {
        var w = int.TryParse(props?.GetValueOrDefault("width"), out var wv) ? wv : 15000;
        var h = int.TryParse(props?.GetValueOrDefault("height"), out var hv) ? hv : 15000;
        var sides = int.TryParse(props?.GetValueOrDefault("sides"), out var sv) ? sv : 5;
        var text = props?.GetValueOrDefault("text");
        var common = CreateShapeCommonChildren(w, h, props);
        var szPosMargin = CreateShapeSzPosMargin(w, h, props);
        var fillBrush = CreateFillBrush(props);

        var polygon = new XElement(HwpxNs.Hp + "polygon",
            new XAttribute("id", NewId()), new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"), new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"), new XAttribute("dropcapstyle", "None"),
            new XAttribute("href", ""), new XAttribute("groupLevel", "0"),
            new XAttribute("instid", NewId()));
        polygon.Add(common);
        polygon.Add(CreateLineShapeElement(props));
        if (fillBrush != null) polygon.Add(fillBrush);
        polygon.Add(new XElement(HwpxNs.Hp + "shadow",
            new XAttribute("type", "NONE"), new XAttribute("color", "#B2B2B2"),
            new XAttribute("offsetX", "0"), new XAttribute("offsetY", "0"), new XAttribute("alpha", "0")));
        // drawText
        if (text != null)
        {
            polygon.Add(new XElement(HwpxNs.Hp + "drawText",
                new XAttribute("lastWidth", w), new XAttribute("name", ""), new XAttribute("editable", "0"),
                new XElement(HwpxNs.Hp + "subList",
                    new XAttribute("id", ""), new XAttribute("textDirection", "HORIZONTAL"),
                    new XAttribute("lineWrap", "BREAK"), new XAttribute("vertAlign", "CENTER"),
                    new XAttribute("linkListIDRef", "0"), new XAttribute("linkListNextIDRef", "0"),
                    new XAttribute("textWidth", "0"), new XAttribute("textHeight", "0"),
                    new XAttribute("hasTextRef", "0"), new XAttribute("hasNumRef", "0"),
                    new XElement(HwpxNs.Hp + "p",
                        new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
                        new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"),
                            new XElement(HwpxNs.Hp + "t", text)))),
                new XElement(HwpxNs.Hp + "textMargin",
                    new XAttribute("left", "283"), new XAttribute("right", "283"),
                    new XAttribute("top", "283"), new XAttribute("bottom", "283"))));
        }
        // Generate regular polygon vertices in orgSz coordinate space
        var cx = w / 2.0; var cy = h / 2.0;
        var r = Math.Min(cx, cy);
        for (int i = 0; i <= sides; i++)
        {
            var angle = -Math.PI / 2 + 2 * Math.PI * i / sides;
            var px = (int)(cx + r * Math.Cos(angle));
            var py = (int)(cy + r * Math.Sin(angle));
            polygon.Add(new XElement(HwpxNs.Hc + "pt", new XAttribute("x", px), new XAttribute("y", py)));
        }
        polygon.Add(szPosMargin.Elements());

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"), polygon,
                new XElement(HwpxNs.Hp + "t", "")));
    }

    private XElement CreateArrow(Dictionary<string, string>? props)
    {
        // Arrow = line with tailStyle arrow
        var merged = new Dictionary<string, string>(props ?? new(), StringComparer.OrdinalIgnoreCase);
        merged["linestyle"] = merged.GetValueOrDefault("linestyle") ?? "SOLID";
        var w = int.TryParse(merged.GetValueOrDefault("width"), out var wv) ? wv : 20000;
        var h = 43;
        var common = CreateShapeCommonChildren(w, h, merged);
        var szPosMargin = CreateShapeSzPosMargin(w, h, merged);
        var color = merged.GetValueOrDefault("color") ?? "#000000";
        var lineWidth = merged.GetValueOrDefault("linewidth") ?? "33";

        var line = new XElement(HwpxNs.Hp + "line",
            new XAttribute("id", NewId()), new XAttribute("zOrder", "0"),
            new XAttribute("numberingType", "PICTURE"),
            new XAttribute("textWrap", "TOP_AND_BOTTOM"), new XAttribute("textFlow", "BOTH_SIDES"),
            new XAttribute("lock", "0"), new XAttribute("dropcapstyle", "None"),
            new XAttribute("href", ""), new XAttribute("groupLevel", "0"),
            new XAttribute("instid", NewId()), new XAttribute("isReverseHV", "0"));
        line.Add(common);
        // Arrow lineShape: tailStyle = arrow
        line.Add(new XElement(HwpxNs.Hp + "lineShape",
            new XAttribute("color", color), new XAttribute("width", lineWidth),
            new XAttribute("style", "SOLID"), new XAttribute("endCap", "FLAT"),
            new XAttribute("headStyle", "NORMAL"), new XAttribute("tailStyle", "ARROW"),
            new XAttribute("headfill", "1"), new XAttribute("tailfill", "1"),
            new XAttribute("headSz", "MEDIUM_MEDIUM"), new XAttribute("tailSz", "MEDIUM_MEDIUM"),
            new XAttribute("outlineStyle", "NORMAL"), new XAttribute("alpha", "0")));
        line.Add(new XElement(HwpxNs.Hp + "shadow",
            new XAttribute("type", "NONE"), new XAttribute("color", "#B2B2B2"),
            new XAttribute("offsetX", "0"), new XAttribute("offsetY", "0"), new XAttribute("alpha", "0")));
        line.Add(new XElement(HwpxNs.Hc + "startPt", new XAttribute("x", "0"), new XAttribute("y", "0")));
        line.Add(new XElement(HwpxNs.Hc + "endPt", new XAttribute("x", w), new XAttribute("y", "0")));
        line.Add(szPosMargin.Elements());
        line.Add(new XElement(HwpxNs.Hp + "shapeComment", "화살표입니다."));

        return new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()), new XAttribute("styleIDRef", "0"), new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"), line,
                new XElement(HwpxNs.Hp + "t", "")));
    }

    // ==================== TOC ====================

    /// <summary>
    /// Create a static Table of Contents from document headings.
    /// Returns multiple paragraph elements — one title + one per heading.
    /// Props: "maxlevel" → max heading depth (default 3), "title" → TOC title (default "목차").
    /// </summary>
    /// <summary>
    /// Create a field-based TOC. Hancom can regenerate via "도구 > 필드 업데이트".
    /// Golden structure: fieldBegin(TABLEOFCONTENTS) with Command parameter → content → fieldEnd
    /// </summary>
    private List<XElement> CreateFieldToc(Dictionary<string, string>? props)
    {
        var maxLevel = int.TryParse(props?.GetValueOrDefault("maxlevel"), out var ml) ? ml : 3;
        var title = props?.GetValueOrDefault("title") ?? "차례";
        var headings = CollectHeadings(maxLevel);
        if (headings.Count == 0)
            throw new CliException("No headings found in document.") { Code = "no_headings" };

        var instId = NewId();
        var fieldId = NewId();
        var result = new List<XElement>();

        // Golden-accurate fieldBegin with Command parameter
        var commandStr = $"TableOfContents:set:140:ContentsMake:uint:31 ContentsStyles:wstring:0: ContentsLevel:int:{maxLevel} ContentsAutoTabRight:int:0 ContentsLeader:int:3 ContentsHyperlink:bool:1  ";

        // Field begin paragraph
        result.Add(new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldBegin",
                        new XAttribute("id", instId),
                        new XAttribute("type", "TABLEOFCONTENTS"),
                        new XAttribute("name", ""),
                        new XAttribute("editable", "1"),
                        new XAttribute("dirty", "0"),
                        new XAttribute("zorder", "-1"),
                        new XAttribute("fieldid", fieldId),
                        new XAttribute("metaTag", ""),
                        new XElement(HwpxNs.Hp + "parameters",
                            new XAttribute("cnt", "2"), new XAttribute("name", ""),
                            new XElement(HwpxNs.Hp + "integerParam", new XAttribute("name", "Prop"), "8"),
                            new XElement(HwpxNs.Hp + "stringParam",
                                new XAttribute("name", "Command"),
                                new XAttribute(XNamespace.Xml + "space", "preserve"),
                                commandStr)))))));

        // TOC title
        var titleProps = new Dictionary<string, string>
            { ["text"] = title, ["fontsize"] = "14", ["bold"] = "true" };
        var titlePara = CreateParagraph(titleProps);
        var structuralKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "text", "styleidref", "styleIDRef", "charpridref", "charPrIDRef", "parapridref", "paraPrIDRef" };
        foreach (var (k, v) in titleProps)
            if (!structuralKeys.Contains(k)) SetParagraphProp(titlePara, k, v);
        result.Add(titlePara);

        // Heading entries
        foreach (var (level, text) in headings)
        {
            var indent = new string('\u3000', level - 1);
            var entryProps = new Dictionary<string, string> { ["text"] = $"{indent}{text}", ["fontsize"] = "9" };
            var entryPara = CreateParagraph(entryProps);
            foreach (var (k, v) in entryProps)
                if (!structuralKeys.Contains(k)) SetParagraphProp(entryPara, k, v);
            result.Add(entryPara);
        }

        // Field end paragraph
        result.Add(new XElement(HwpxNs.Hp + "p",
            new XAttribute("id", NewId()),
            new XAttribute("styleIDRef", "0"),
            new XAttribute("paraPrIDRef", "0"),
            new XElement(HwpxNs.Hp + "run", new XAttribute("charPrIDRef", "0"),
                new XElement(HwpxNs.Hp + "ctrl",
                    new XElement(HwpxNs.Hp + "fieldEnd",
                        new XAttribute("beginIDRef", instId),
                        new XAttribute("fieldid", fieldId))),
                new XElement(HwpxNs.Hp + "t"))));

        return result;
    }

    private List<XElement> CreateStaticToc(Dictionary<string, string>? props)
    {
        var maxLevel = int.TryParse(props?.GetValueOrDefault("maxlevel"), out var ml) ? ml : 3;
        var title = props?.GetValueOrDefault("title") ?? "목차";

        var headings = CollectHeadings(maxLevel);
        if (headings.Count == 0)
            throw new CliException("No headings found in document. Set outlineLevel on paragraphs first.")
                { Code = "no_headings" };

        var result = new List<XElement>();
        var structuralKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "text", "styleidref", "styleIDRef", "charpridref", "charPrIDRef",
              "parapridref", "paraPrIDRef" };

        // TOC title paragraph — bold, 14pt
        var titleProps = new Dictionary<string, string>
        {
            ["text"] = title,
            ["fontsize"] = "14",
            ["bold"] = "true"
        };
        var titlePara = CreateParagraph(titleProps);
        foreach (var (k, v) in titleProps)
            if (!structuralKeys.Contains(k)) SetParagraphProp(titlePara, k, v);
        result.Add(titlePara);

        // One paragraph per heading with indentation
        foreach (var (level, text) in headings)
        {
            var indent = new string('\u3000', level - 1); // fullwidth space for visual indent
            var entryProps = new Dictionary<string, string>
            {
                ["text"] = $"{indent}{text}"
            };
            // Sub-headings slightly smaller
            if (level >= 2)
                entryProps["fontsize"] = "9";
            var entryPara = CreateParagraph(entryProps);
            foreach (var (k, v) in entryProps)
                if (!structuralKeys.Contains(k)) SetParagraphProp(entryPara, k, v);
            result.Add(entryPara);
        }

        // Empty paragraph after TOC
        result.Add(CreateParagraph(new Dictionary<string, string> { ["text"] = " " }));

        return result;
    }

    /// <summary>
    /// Collect headings from all sections.
    /// Detection: style name match ("개요 N" / "Heading N") OR paraPr > heading element.
    /// </summary>
    private List<(int Level, string Text)> CollectHeadings(int maxLevel)
    {
        var headings = new List<(int Level, string Text)>();

        foreach (var (section, para, localIdx) in _doc.AllParagraphs())
        {
            int? level = null;

            // Method 1: Style-based detection (same as View.cs GetParagraphStyleInfo)
            var styleIdRef = para.Attribute("styleIDRef")?.Value;
            if (_doc.Header != null && styleIdRef != null)
            {
                var style = _doc.Header.Root!.Descendants(HwpxNs.Hh + "style")
                    .FirstOrDefault(s => s.Attribute("id")?.Value == styleIdRef);
                if (style != null)
                {
                    var name = style.Attribute("name")?.Value ?? "";
                    var m = System.Text.RegularExpressions.Regex.Match(name, @"개요\s*(\d+)");
                    if (!m.Success)
                        m = System.Text.RegularExpressions.Regex.Match(name, @"(?i)heading\s*(\d+)");
                    if (m.Success)
                        level = int.Parse(m.Groups[1].Value);
                }
            }

            // Method 2: paraPr > heading element with type="OUTLINE"
            if (level == null)
            {
                var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;
                if (_doc.Header != null && paraPrIdRef != null)
                {
                    var paraPr = _doc.Header.Root!.Descendants(HwpxNs.Hh + "paraPr")
                        .FirstOrDefault(pp => pp.Attribute("id")?.Value == paraPrIdRef);
                    var heading = paraPr?.Element(HwpxNs.Hh + "heading");
                    if (heading?.Attribute("type")?.Value == "OUTLINE"
                        && int.TryParse(heading.Attribute("level")?.Value, out var hl))
                        level = hl;
                }
            }

            if (level.HasValue && level.Value >= 1 && level.Value <= maxLevel)
            {
                var text = ExtractParagraphText(para);
                if (!string.IsNullOrWhiteSpace(text))
                    headings.Add((level.Value, HwpxKorean.Normalize(text)));
            }
        }

        return headings;
    }

    // ==================== Multi-Section ====================

    /// <summary>
    /// Add a new section to the document. Creates section XML, updates manifest, increments secCnt.
    /// Props: "orientation" (PORTRAIT/LANDSCAPE), "pageWidth", "pageHeight", "marginTop/Bottom/Left/Right"
    /// </summary>
    private HwpxSection AddNewSection(Dictionary<string, string>? props)
    {
        var newIndex = _doc.Sections.Count;
        var entryPath = $"Contents/section{newIndex}.xml";

        // Clone the first section's structure (secPr, namespaces, etc.) for compatibility
        var sourceSection = _doc.PrimarySection;
        var sourceRoot = sourceSection.Root;

        // Deep-copy the section root element (preserves all namespaces + child structure)
        var newRoot = new XElement(sourceRoot);

        // Strip all content paragraphs except the first one (which holds secPr)
        var paras = newRoot.Elements(HwpxNs.Hp + "p").ToList();
        for (int i = 1; i < paras.Count; i++)
            paras[i].Remove();

        // Clear text in the first paragraph (keep secPr structure)
        var firstPara = newRoot.Elements(HwpxNs.Hp + "p").First();
        firstPara.SetAttributeValue("id", NewId());
        // Remove linesegarray (stale layout cache)
        firstPara.Elements(HwpxNs.Hp + "linesegarray").Remove();
        // Clear text content in runs but keep secPr
        foreach (var run in firstPara.Elements(HwpxNs.Hp + "run"))
        {
            foreach (var t in run.Elements(HwpxNs.Hp + "t"))
                t.Value = "";
        }

        // Apply orientation if requested
        // Hancom: landscape="NARROWLY" (가로), "WIDELY" (세로) — dimensions stay the same
        var landscape = props?.GetValueOrDefault("orientation")?.Equals("LANDSCAPE", StringComparison.OrdinalIgnoreCase) == true;
        if (landscape)
        {
            var pagePr = newRoot.Descendants(HwpxNs.Hp + "pagePr").FirstOrDefault();
            pagePr?.SetAttributeValue("landscape", "NARROWLY");
        }

        // Apply custom page dimensions/margins if specified
        if (props != null)
        {
            var pagePr = newRoot.Descendants(HwpxNs.Hp + "pagePr").FirstOrDefault();
            var margin = pagePr?.Element(HwpxNs.Hp + "margin");
            if (props.ContainsKey("pagewidth")) pagePr?.SetAttributeValue("width", props["pagewidth"]);
            if (props.ContainsKey("pageheight")) pagePr?.SetAttributeValue("height", props["pageheight"]);
            if (props.ContainsKey("margintop")) margin?.SetAttributeValue("top", props["margintop"]);
            if (props.ContainsKey("marginbottom")) margin?.SetAttributeValue("bottom", props["marginbottom"]);
            if (props.ContainsKey("marginleft")) margin?.SetAttributeValue("left", props["marginleft"]);
            if (props.ContainsKey("marginright")) margin?.SetAttributeValue("right", props["marginright"]);
        }

        var secDoc = new XDocument(new XDeclaration("1.0", "UTF-8", null), newRoot);

        var section = new HwpxSection
        {
            Index = newIndex,
            EntryPath = entryPath,
            Document = secDoc
        };
        _doc.Sections.Add(section);

        // Update header.xml secCnt
        var headEl = _doc.Header?.Root;
        if (headEl != null)
        {
            var currentCnt = (int?)headEl.Attribute("secCnt") ?? 1;
            headEl.SetAttributeValue("secCnt", (currentCnt + 1).ToString());
            SaveHeader();
        }

        // Update manifest (content.hpf)
        AddManifestEntry(entryPath);

        // Save new section to ZIP
        SaveSection(section.Root);
        return section;
    }

    /// <summary>Add an item+spine entry to content.hpf manifest.</summary>
    private void AddManifestEntry(string entryPath)
    {
        var opfDoc = _doc.ManifestDoc;
        if (opfDoc?.Root == null) return;

        var opfNs = opfDoc.Root.Name.Namespace;
        var itemId = Path.GetFileNameWithoutExtension(entryPath);

        var manifest = opfDoc.Root.Element(opfNs + "manifest");
        manifest?.Add(new XElement(opfNs + "item",
            new XAttribute("id", itemId),
            new XAttribute("href", entryPath),
            new XAttribute("media-type", "application/xml")));

        var spine = opfDoc.Root.Element(opfNs + "spine");
        spine?.Add(new XElement(opfNs + "itemref",
            new XAttribute("idref", itemId)));

        SaveManifest();
    }

    /// <summary>Remove an item+spine entry from content.hpf manifest.</summary>
    private void RemoveManifestEntry(string entryPath)
    {
        var opfDoc = _doc.ManifestDoc;
        if (opfDoc?.Root == null) return;

        var opfNs = opfDoc.Root.Name.Namespace;
        var itemId = Path.GetFileNameWithoutExtension(entryPath);

        var manifest = opfDoc.Root.Element(opfNs + "manifest");
        manifest?.Elements(opfNs + "item")
            .FirstOrDefault(e => e.Attribute("id")?.Value == itemId
                || e.Attribute("href")?.Value == entryPath)
            ?.Remove();

        var spine = opfDoc.Root.Element(opfNs + "spine");
        spine?.Elements(opfNs + "itemref")
            .FirstOrDefault(e => e.Attribute("idref")?.Value == itemId)
            ?.Remove();

        SaveManifest();
    }

    // G6: Korean government standard margins (행정문서 양식)
    private static class GovStandardMargins
    {
        public const int PageWidth = 59528;   // A4: 210mm
        public const int PageHeight = 84186;  // A4: 297mm
        public const int Top = 5668;          // 20mm
        public const int Bottom = 4252;       // 15mm
        public const int Left = 5668;         // 20mm
        public const int Right = 5668;        // 20mm
        public const int Header = 4252;       // 15mm
        public const int Footer = 4252;       // 15mm
        public const int Gutter = 0;
    }

    /// <summary>Apply Korean government standard margins to section 1.</summary>
    private bool ApplyGovStandardMargins()
    {
        var sec = _doc.Sections.FirstOrDefault();
        if (sec == null) return false;
        var secPr = sec.Root.Descendants(HwpxNs.Hp + "secPr").FirstOrDefault();
        var pagePr = secPr?.Element(HwpxNs.Hp + "pagePr");
        var margin = pagePr?.Element(HwpxNs.Hp + "margin");
        if (margin == null) return false;

        margin.SetAttributeValue("top", GovStandardMargins.Top.ToString());
        margin.SetAttributeValue("bottom", GovStandardMargins.Bottom.ToString());
        margin.SetAttributeValue("left", GovStandardMargins.Left.ToString());
        margin.SetAttributeValue("right", GovStandardMargins.Right.ToString());
        margin.SetAttributeValue("header", GovStandardMargins.Header.ToString());
        margin.SetAttributeValue("footer", GovStandardMargins.Footer.ToString());
        margin.SetAttributeValue("gutter", GovStandardMargins.Gutter.ToString());

        _dirty = true;
        return true;
    }

    /// <summary>Save content.hpf manifest to ZIP archive.</summary>
    private void SaveManifest()
    {
        if (_doc.ManifestDoc == null || _doc.ManifestEntryPath == null) return;

        var entryName = _doc.ManifestEntryPath;
        var entry = _doc.Archive.GetEntry(entryName);
        if (entry == null) return;

        entry.Delete();
        var newEntry = _doc.Archive.CreateEntry(entryName, System.IO.Compression.CompressionLevel.Optimal);
        using var stream = newEntry.Open();
        var xmlStr = HwpxPacker.MinifyXml(_doc.ManifestDoc.ToString(SaveOptions.DisableFormatting));
        xmlStr = HwpxPacker.RestoreOriginalNamespaces(xmlStr);
        xmlStr = "<?xml version='1.0' encoding='UTF-8'?>" + xmlStr;
        var bytes = System.Text.Encoding.UTF8.GetBytes(xmlStr);
        stream.Write(bytes, 0, bytes.Length);
    }

    // ==================== Metadata ====================

    /// <summary>Set a metadata value in content.hpf. Matches Hancom's OPF meta structure.</summary>
    private bool SetMetadata(string name, string value)
    {
        var opfDoc = _doc.ManifestDoc;
        if (opfDoc?.Root == null) return false;
        var ns = opfDoc.Root.Name.Namespace;
        var metadata = opfDoc.Root.Element(ns + "metadata");
        if (metadata == null)
        {
            metadata = new XElement(ns + "metadata");
            opfDoc.Root.AddFirst(metadata);
        }

        // title and language are direct elements; others use <opf:meta name="...">
        if (name is "title" or "language")
        {
            var el = metadata.Element(ns + name);
            if (el == null) { el = new XElement(ns + name); metadata.Add(el); }
            el.Value = value;
        }
        else
        {
            var el = metadata.Elements(ns + "meta")
                .FirstOrDefault(e => e.Attribute("name")?.Value == name);
            if (el == null)
            {
                el = new XElement(ns + "meta",
                    new XAttribute("name", name), new XAttribute("content", "text"));
                metadata.Add(el);
            }
            el.Value = value;
        }

        SaveManifest();
        _dirty = true;
        return true;
    }

    /// <summary>Read all metadata from content.hpf.</summary>
    public Dictionary<string, string> GetMetadata()
    {
        var opfDoc = _doc.ManifestDoc;
        if (opfDoc?.Root == null) return new();
        var ns = opfDoc.Root.Name.Namespace;
        var metadata = opfDoc.Root.Element(ns + "metadata");
        if (metadata == null) return new();

        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // Direct elements
        var title = metadata.Element(ns + "title")?.Value;
        if (!string.IsNullOrEmpty(title)) result["title"] = title;
        var lang = metadata.Element(ns + "language")?.Value;
        if (!string.IsNullOrEmpty(lang)) result["language"] = lang;

        // G4: Dublin Core elements (dc:title, dc:creator, dc:subject, etc.)
        var dcFields = new[] { "title", "creator", "subject", "description",
                               "publisher", "contributor", "date", "type",
                               "format", "identifier", "source", "language",
                               "relation", "coverage", "rights" };
        foreach (var field in dcFields)
        {
            if (result.ContainsKey(field)) continue;
            var dcEl = metadata.Element(HwpxNs.Dc + field);
            if (dcEl != null && !string.IsNullOrEmpty(dcEl.Value))
                result[field] = dcEl.Value;
        }

        // Meta elements
        foreach (var meta in metadata.Elements(ns + "meta"))
        {
            var n = meta.Attribute("name")?.Value;
            var v = meta.Value;
            if (n != null && !string.IsNullOrEmpty(v)) result[n] = v;
        }
        return result;
    }

    /// <summary>Remove a section by path (e.g. /section[2]).</summary>
    private string? RemoveSection(string path)
    {
        var segments = ParsePath(path);
        var secIdx = (segments[0].Index ?? 1) - 1;
        if (_doc.Sections.Count <= 1)
            throw new CliException("Cannot remove last section") { Code = "invalid_op" };
        if (secIdx < 0 || secIdx >= _doc.Sections.Count)
            throw new CliException($"Section {secIdx + 1} not found");

        var section = _doc.Sections[secIdx];
        _doc.Sections.RemoveAt(secIdx);

        // Reindex remaining sections
        for (int i = secIdx; i < _doc.Sections.Count; i++)
            _doc.Sections[i].Index = i;

        // Update secCnt
        var headEl = _doc.Header?.Root;
        if (headEl != null)
        {
            headEl.SetAttributeValue("secCnt", _doc.Sections.Count.ToString());
            SaveHeader();
        }

        // Remove manifest entry + ZIP entry
        RemoveManifestEntry(section.EntryPath);
        var zipEntry = _doc.Archive.GetEntry(section.EntryPath);
        zipEntry?.Delete();
        _dirty = true;
        return null;
    }
}
