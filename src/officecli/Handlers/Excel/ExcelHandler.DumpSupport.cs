// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using ThreadedCmt = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Dump support ====================
    //
    // CONSISTENCY(emit-X-mirror): read-only enumeration surface consumed by
    // ExcelBatchEmitter, mirroring the public helper methods PowerPointHandler
    // grew for PptxBatchEmitter (GetSlideBulletImageParts, GetTimingAudioRels,
    // ...). The emitter lives outside the handler's partial-class family, so
    // everything it needs beyond Get/Query is exposed here.

    /// <summary>Sheet names in workbook (sldIdLst-equivalent) order.</summary>
    public List<string> GetDumpSheetNames() => GetWorksheets().Select(t => t.Name).ToList();

    /// <summary>
    /// Workbook-level settings node (date1904, calc.*, activeTab, protection).
    /// Same Format keys as PopulateWorkbookSettings emits on Get.
    /// </summary>
    public DocumentNode GetDumpWorkbookNode()
    {
        var node = new DocumentNode { Path = "/workbook", Type = "workbook" };
        PopulateWorkbookSettings(node);
        // Core document properties (settable via `set / --prop title=…`):
        // the dump silently dropped them. Timestamps and lastModifiedBy are
        // deliberately excluded — save-time stamping would flip them every
        // replay cycle and break dump idempotency.
        var pkgProps = _doc.PackageProperties;
        if (!string.IsNullOrEmpty(pkgProps.Title)) node.Format["title"] = pkgProps.Title!;
        if (!string.IsNullOrEmpty(pkgProps.Creator)) node.Format["author"] = pkgProps.Creator!;
        if (!string.IsNullOrEmpty(pkgProps.Subject)) node.Format["subject"] = pkgProps.Subject!;
        if (!string.IsNullOrEmpty(pkgProps.Description)) node.Format["description"] = pkgProps.Description!;
        if (!string.IsNullOrEmpty(pkgProps.Keywords)) node.Format["keywords"] = pkgProps.Keywords!;
        if (!string.IsNullOrEmpty(pkgProps.Category)) node.Format["category"] = pkgProps.Category!;
        return node;
    }

    /// <summary>
    /// Enumerate every row of a sheet with ALL cells that carry content OR
    /// style. The bulk Get path (GetSheetChildNodes) intentionally omits
    /// styled-empty cells (&lt;c s="1"/&gt;, issue #149 bloat guard); a dump
    /// must include them because their xf holds user-visible formatting
    /// (filled header bands, bordered empty grids). Each cell node is built
    /// by the same CellToNode Get uses, so Format keys match Get exactly.
    /// A dump-only <c>__raw</c> Format key carries the raw stored
    /// &lt;x:v&gt; text so the emitter can reproduce numbers/dates without
    /// going through display formatting.
    /// </summary>
    public List<DocumentNode> GetDumpRowNodes(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var rows = new List<DocumentNode>();
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null) return rows;

        // One evaluator per sheet: CellToNode lazily creates a fresh
        // FormulaEvaluator per formula cell when none is passed, which is
        // O(cells × sheet-size) on formula-heavy sheets.
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);
        // For unresolved-shared-string detection: a cell with t="s" whose
        // index has no entry (missing/truncated sharedStrings part) would
        // otherwise surface its INDEX as the cell text — confidently-wrong
        // data. Mark such cells so the emitter warns and skips them.
        var sstCount = _doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable?
            .Elements<SharedStringItem>().Count() ?? 0;
        var seenRowIndices = new HashSet<uint>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var ridx = row.RowIndex?.Value ?? 0;
            if (ridx != 0 && !seenRowIndices.Add(ridx)) continue;

            var rowNode = new DocumentNode
            {
                Path = $"/{sheetName}/row[{ridx}]",
                Type = "row"
            };
            // CONSISTENCY(unit-qualified-readback): pt-suffix row height,
            // mirroring GetSheetChildNodes.
            if (row.Height?.Value != null)
                rowNode.Format["height"] = $"{row.Height.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}pt";
            if (row.Hidden?.Value == true)
                rowNode.Format["hidden"] = true;
            if (row.OutlineLevel?.Value is { } rol && rol > 0)
                rowNode.Format["outlineLevel"] = (int)rol;
            if (row.Collapsed?.Value == true)
                rowNode.Format["collapsed"] = true;
            // Row-level default style (CT_Row @customFormat + @s). @customFormat
            // is the gate that makes @s the default cellXf for the row's
            // phantom (never-materialized) cells; without it the dump dropped
            // the row's default style entirely. Mirrors the Get readback keys.
            if (row.CustomFormat?.Value == true)
            {
                rowNode.Format["customFormat"] = true;
                if (row.StyleIndex?.Value is { } rsi)
                    rowNode.Format["s"] = rsi.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            // Column cursor for cells whose r= attribute is absent — a legal
            // OOXML variant (position is implied: one right of the previous
            // cell). Without synthesis such cells got a "?" path and the
            // emitter silently dropped their values from the dump.
            int impliedCol = 0;
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value is { } cr && cr.Length > 0)
                {
                    try { impliedCol = ColumnNameToIndex(ParseCellReference(cr).Item1); }
                    catch { impliedCol++; }
                }
                else
                {
                    impliedCol++;
                    cell.CellReference = $"{IndexToColumnName(impliedCol)}{ridx}";
                }
                var hasContent = CellHasContent(cell);
                var hasStyle = cell.StyleIndex != null && cell.StyleIndex.Value != 0;
                if (!hasContent && !hasStyle) continue;
                var cellNode = CellToNode(sheetName, cell, worksheet, eval);
                if (cell.DataType?.Value == CellValues.SharedString
                    && (!int.TryParse(cell.CellValue?.Text, out var sstIdx)
                        || sstIdx < 0 || sstIdx >= sstCount))
                    cellNode.Format["__unresolvedSst"] = true;
                var raw = cell.CellValue?.Text;
                if (!string.IsNullOrEmpty(raw))
                    cellNode.Format["__raw"] = raw;
                // CONSISTENCY(picture-inline-base64): in-cell image bytes ride
                // the same inline `data:<contentType>;base64,<bytes>` carrier as
                // the pptx/word picture round-trip (PptxBatchEmitter.Media.cs) —
                // no sidecar file, so a dump script stays self-contained.
                // Image bytes are dump-only: normal Get keeps its compact
                // metadata surface instead of embedding a large data URI.
                if (cellNode.Format.TryGetValue("type", out var cellType)
                    && cellType is string typeName && typeName == "Image"
                    && TryGetInCellImage(cell, out var imageInfo))
                {
                    using var imageStream = imageInfo.Part.GetStream(FileMode.Open, FileAccess.Read);
                    using var imageBytes = new MemoryStream();
                    imageStream.CopyTo(imageBytes);
                    cellNode.Format["__imageDataUri"] =
                        $"data:{imageInfo.ContentType};base64,{Convert.ToBase64String(imageBytes.ToArray())}";
                }
                // Rich-text carrier: inline-string runs or a rich shared-string
                // entry can't ride the CSV baseline. Serialize the runs into
                // the `runs=<json>` shape ApplyRichTextToCell consumes so the
                // emitter can replay via `set type=richtext runs=...`.
                if (HasRichTextContent(cell))
                    cellNode.Format["__richruns"] = SerializeRichTextRuns(cell);
                rowNode.Children.Add(cellNode);
            }

            if (rowNode.Children.Count == 0 && rowNode.Format.Count == 0) continue;
            rowNode.ChildCount = rowNode.Children.Count;
            rows.Add(rowNode);
        }
        return rows;
    }

    private bool HasRichTextContent(Cell cell)
    {
        var inline = cell.GetFirstChild<InlineString>();
        if (inline != null && inline.Elements<Run>().Any()) return true;
        return GetRichTextHost(cell)?.Elements<Run>().Any() == true;
    }

    private OpenXmlElement? GetRichTextHost(Cell cell)
    {
        var inline = cell.GetFirstChild<InlineString>();
        if (inline != null && inline.Elements<Run>().Any()) return inline;
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(cell.CellValue?.Text, out var ssIdx))
        {
            var ssItems = _doc.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
            return ssItems?.Elements<SharedStringItem>().ElementAtOrDefault(ssIdx);
        }
        return null;
    }

    /// <summary>
    /// Serialize a rich-text cell's runs into the `runs=<json>` array
    /// ApplyRichTextToCell consumes. Key vocabulary mirrors RunToNode's Get
    /// output (bold/italic/strike/underline/superscript/subscript/size/color/
    /// font) — the Add side accepts each of these verbatim.
    /// </summary>
    private string SerializeRichTextRuns(Cell cell)
    {
        var host = GetRichTextHost(cell);
        var arr = new System.Text.Json.Nodes.JsonArray();
        if (host == null) return "[]";
        // A rich shared-string item may open with a bare <t> (unformatted
        // leading segment) before the first <r>; carry it as a plain run.
        var leading = host.GetFirstChild<Text>();
        if (leading != null && !string.IsNullOrEmpty(leading.Text))
            arr.Add((System.Text.Json.Nodes.JsonNode)new System.Text.Json.Nodes.JsonObject { ["text"] = leading.Text });
        foreach (var run in host.Elements<Run>())
        {
            var o = new System.Text.Json.Nodes.JsonObject { ["text"] = run.Text?.Text ?? "" };
            var rp = run.RunProperties;
            if (rp != null)
            {
                if (rp.GetFirstChild<Bold>() != null) o["bold"] = true;
                if (rp.GetFirstChild<Italic>() != null) o["italic"] = true;
                if (rp.GetFirstChild<Strike>() != null) o["strike"] = true;
                var ul = rp.GetFirstChild<Underline>();
                if (ul != null) o["underline"] = ul.Val?.InnerText == "double" ? "double" : "single";
                var va = rp.GetFirstChild<VerticalTextAlignment>();
                if (va?.Val?.Value == VerticalAlignmentRunValues.Superscript) o["superscript"] = true;
                if (va?.Val?.Value == VerticalAlignmentRunValues.Subscript) o["subscript"] = true;
                var sz = rp.GetFirstChild<FontSize>();
                if (sz?.Val?.Value != null) o["size"] = $"{sz.Val.Value:0.##}pt";
                var clr = rp.GetFirstChild<Color>();
                if (clr?.Rgb?.Value != null) o["color"] = ParseHelpers.FormatHexColor(clr.Rgb.Value!);
                var rf = rp.GetFirstChild<RunFont>();
                if (rf?.Val?.Value != null) o["font"] = rf.Val.Value!;
            }
            arr.Add((System.Text.Json.Nodes.JsonNode)o);
        }
        return arr.ToJsonString();
    }

    /// <summary>Per-sheet element counts for the indexed Get paths the batch
    /// emitter transcribes (cf[N] / validation[N] / comment[N] / table[N] /
    /// chart[N] / sparkline[N]).</summary>
    public (int Tables, int Cfs, int Validations, int Comments, int Charts, int Sparklines, bool HasExtendedChart)
        GetDumpElementCounts(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var ws = GetSheet(worksheet);
        var tables = worksheet.TableDefinitionParts.Count();
        var cfs = ws.Elements<ConditionalFormatting>().Count();
        var validations = ws.GetFirstChild<DataValidations>()?.Elements<DataValidation>().Count() ?? 0;
        var comments = worksheet.WorksheetCommentsPart?.Comments?
            .GetFirstChild<CommentList>()?.Elements<Comment>().Count() ?? 0;
        // chart[N] Get indexes over GetExcelCharts (standard + extended in
        // drawing order); extended (chartEx) charts have no Add vocabulary.
        int charts = 0; bool hasExtended = false;
        if (worksheet.DrawingsPart is { } dp)
        {
            var chartInfos = GetExcelCharts(dp);
            charts = chartInfos.Count;
            hasExtended = chartInfos.Any(c => c.IsExtended);
        }
        int sparklines = 0;
        var extLst = ws.GetFirstChild<WorksheetExtensionList>();
        if (extLst != null)
            sparklines = extLst.Descendants<DocumentFormat.OpenXml.Office2010.Excel.SparklineGroup>().Count();
        return (tables, cfs, validations, comments, charts, sparklines, hasExtended);
    }

    /// <summary>
    /// Column definitions of a sheet, one node per column LETTER (Column
    /// elements with min/max ranges are expanded). Format keys mirror the
    /// /Sheet/col[X] Get surface: width, hidden, customWidth, outlineLevel,
    /// collapsed. Range expansion is capped per definition so a stray
    /// A:XFD-wide entry cannot emit 16 000 rows — the overflow is reported
    /// through <paramref name="truncated"/>.
    /// </summary>
    public List<DocumentNode> GetDumpColumnNodes(string sheetName, out bool truncated)
    {
        truncated = false;
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var nodes = new List<DocumentNode>();
        var cols = GetSheet(worksheet).GetFirstChild<Columns>();
        if (cols == null) return nodes;

        const int MaxExpandPerDef = 256;
        foreach (var col in cols.Elements<Column>())
        {
            var min = (int)(col.Min?.Value ?? 0);
            var max = (int)(col.Max?.Value ?? 0);
            if (min < 1 || max < min) continue;
            var span = max - min + 1;
            if (span > MaxExpandPerDef)
            {
                truncated = true;
                max = min + MaxExpandPerDef - 1;
            }
            for (int i = min; i <= max; i++)
            {
                var letter = IndexToColumnName(i);
                var node = new DocumentNode
                {
                    Path = $"/{sheetName}/col[{letter}]",
                    Type = "column",
                    Preview = letter
                };
                if (col.Width?.Value != null && col.CustomWidth?.Value == true)
                    node.Format["width"] = col.Width.Value;
                if (col.Hidden?.Value == true) node.Format["hidden"] = true;
                if (col.OutlineLevel?.Value is { } ol && ol > 0)
                    node.Format["outlineLevel"] = (int)ol;
                if (col.Collapsed?.Value == true) node.Format["collapsed"] = true;
                // Column-level number format lives on a style reference
                // (<col s="N">), NOT a CT_Col attribute. Resolve it to the
                // canonical `numberformat` code so the emitter can replay it
                // via `set col[X] --prop numberformat=...` (mirrors the column
                // Get reader in ExcelHandler.Query.cs).
                if (col.Style?.Value is uint colStyleIdx && colStyleIdx != 0)
                {
                    // A whole-column style ref (<col style="N">) can carry
                    // font/fill/border/alignment facets, not just a number
                    // format. The numberformat-only readback below cannot
                    // capture those, so when the referenced cellXf applies any
                    // non-numberformat facet, emit the raw style index (mirrors
                    // the row @s fix). Replay realigns the index because styled
                    // cells are emitted before column settings, so the same
                    // cellXf is recreated at the same slot. numberformat-only
                    // columns keep the portable code emission.
                    var colXf = _doc.WorkbookPart?.WorkbookStylesPart?.Stylesheet?
                        .CellFormats?.ElementAtOrDefault((int)colStyleIdx) as CellFormat;
                    bool colHasFacets = colXf != null && (
                        colXf.ApplyFont?.Value == true || colXf.ApplyFill?.Value == true
                        || colXf.ApplyBorder?.Value == true || colXf.ApplyAlignment?.Value == true
                        || (colXf.FontId?.Value ?? 0) != 0 || (colXf.FillId?.Value ?? 0) > 1
                        || (colXf.BorderId?.Value ?? 0) != 0);
                    if (colHasFacets)
                    {
                        node.Format["style"] = colStyleIdx.ToString();
                    }
                    else
                    {
                        var (colNumFmtId, colFmtCode) = ExcelDataFormatter.GetCellFormat(
                            new Cell { StyleIndex = colStyleIdx }, _doc.WorkbookPart);
                        if (colNumFmtId > 0)
                        {
                            var colCode = !string.IsNullOrEmpty(colFmtCode)
                                ? colFmtCode
                                : ExcelDataFormatter.ResolveBuiltInFormatCode(colNumFmtId);
                            if (!string.IsNullOrEmpty(colCode))
                                node.Format["numberformat"] = colCode;
                        }
                    }
                }
                if (node.Format.Count > 0) nodes.Add(node);
            }
        }
        return nodes;
    }

    /// <summary>
    /// All merged ranges of a sheet, straight from the worksheet's
    /// MergeCells element. Per-cell Format["merge"] readback cannot drive
    /// this — a merge whose cells are all empty and unstyled has no cell
    /// node at all in the dump row enumeration.
    /// </summary>
    public List<string> GetDumpMergeRanges(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var merges = GetSheet(worksheet).GetFirstChild<MergeCells>();
        if (merges == null) return new List<string>();
        return merges.Elements<MergeCell>()
            .Select(m => m.Reference?.Value)
            .Where(r => !string.IsNullOrEmpty(r))
            .Select(r => r!)
            .ToList();
    }

    /// <summary>
    /// Resolve a dump subtree token ("SheetName" or "sheet[N]") to the
    /// canonical sheet name; null when it does not resolve.
    /// </summary>
    public string? ResolveDumpSheetName(string token)
    {
        try
        {
            var resolved = ResolveSheetName(token);
            return FindWorksheet(resolved) != null ? resolved : null;
        }
        catch { return null; }
    }

    /// <summary>Pivot-table locations on a sheet ("E1:H10" ranges). The dump
    /// value baseline EXCLUDES cells inside these ranges — they are derived
    /// render output that `add pivottable` regenerates on replay; importing
    /// them as static values would collide with the rebuilt pivot.</summary>
    public List<string> GetDumpPivotLocations(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var result = new List<string>();
        foreach (var pp in worksheet.PivotTableParts)
        {
            var loc = pp.PivotTableDefinition?.Location?.Reference?.Value;
            if (!string.IsNullOrEmpty(loc)) result.Add(loc!);
        }
        return result;
    }

    /// <summary>Number of slicers on a sheet (slicer[N] index space).</summary>
    public int GetDumpSlicerCount(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        // Count slicer DEFINITIONS, not the worksheet extLst <x14:slicer r:id>
        // references. Excel groups every slicer bound to a sheet into one
        // shared SlicersPart with N <x14:slicer> children but only ONE extLst
        // reference to that part — counting references returned 1 for N
        // slicers and the dump silently dropped slicer[2..N]. Mirror the
        // slicer[N] Get index space (Helpers.Node.cs / TryFindSlicerByIndex),
        // which enumerates the first SlicersPart's Slicers collection.
        var slicersPart = worksheet.GetPartsOfType<SlicersPart>().FirstOrDefault();
        return slicersPart?.Slicers?
            .Elements<DocumentFormat.OpenXml.Office2010.Excel.Slicer>().Count() ?? 0;
    }

    /// <summary>
    /// Per-shape flags (aligned with the shape[N] index space) marking
    /// shapes that live inside an mc:AlternateContent Fallback — e.g. the
    /// "Slicer (requires Excel 2010 or later)" placeholder rectangle a
    /// slicer carries for down-level clients. The dump must skip them:
    /// AddSlicer regenerates the fallback on replay, and re-adding it as a
    /// REAL shape both duplicates it and breaks dump idempotency.
    /// </summary>
    public List<bool> GetDumpShapeFallbackFlags(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var wsDrawing = worksheet.DrawingsPart?.WorksheetDrawing;
        var flags = new List<bool>();
        if (wsDrawing == null) return flags;
        foreach (var (shape, _) in EnumerateLeafShapes(wsDrawing))
        {
            var inFallback = false;
            for (var p = shape.Parent; p != null; p = p.Parent)
                if (p.LocalName == "Fallback") { inFallback = true; break; }
            flags.Add(inFallback);
        }
        return flags;
    }

    /// <summary>Number of pivot tables on a sheet (pivottable[N] index space).</summary>
    public int GetDumpPivotCount(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        return worksheet.PivotTableParts.Count();
    }

    /// <summary>Per-chart flags (chart[N] index space) marking extended (cx:)
    /// charts — the semantic chart emit must skip them (they ride the
    /// verbatim chartex carrier instead).</summary>
    public List<bool> GetDumpChartExtendedFlags(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        if (worksheet.DrawingsPart is not { } dp) return new List<bool>();
        return GetExcelCharts(dp).Select(c => c.IsExtended).ToList();
    }

    /// <summary>One verbatim OLE carrier per embedded object: pinned rIds,
    /// payload/icon bytes, the VML anchor shape, and the oleObjects child
    /// element XML. Consumed by ExcelBatchEmitter.EmitOles → add-part ole.</summary>
    public sealed record DumpOleSlice(
        string RelId, string DataBase64, string Kind, string ContentType, string Extension,
        string? IconRelId, string? IconBase64, string? IconContentType,
        string? VmlShapeXml, string ObjectXml);

    public List<DumpOleSlice> GetDumpOleSlices(string sheetName)
    {
        var result = new List<DumpOleSlice>();
        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null) return result;
        var oleObjects = GetSheet(worksheet).GetFirstChild<OleObjects>();
        if (oleObjects == null) return result;

        string? vmlXml = null;
        if (worksheet.VmlDrawingParts.FirstOrDefault() is { } vmlPart)
        {
            using var r = new StreamReader(vmlPart.GetStream());
            vmlXml = r.ReadToEnd();
        }

        foreach (var child in oleObjects.ChildElements)
        {
            var ole = child as OleObject ?? child.Descendants<OleObject>().FirstOrDefault();
            if (ole == null) continue;
            var relId = ole.Id?.Value;
            if (string.IsNullOrEmpty(relId)) continue;
            OpenXmlPart payloadPart;
            try { payloadPart = worksheet.GetPartById(relId!); }
            catch { continue; }

            using var ps = payloadPart.GetStream();
            using var pms = new MemoryStream();
            ps.CopyTo(pms);
            var ext = Path.GetExtension(payloadPart.Uri.ToString());
            if (string.IsNullOrEmpty(ext)) ext = ".bin";

            string? iconRid = null, iconB64 = null, iconCt = null;
            var objectPr = ole.GetFirstChild<EmbeddedObjectProperties>();
            var iconRelRaw = objectPr?.Id?.Value;
            if (!string.IsNullOrEmpty(iconRelRaw)
                && worksheet.GetPartById(iconRelRaw!) is ImagePart iconPart)
            {
                using var isr = iconPart.GetStream();
                using var ims = new MemoryStream();
                isr.CopyTo(ims);
                iconRid = iconRelRaw;
                iconB64 = Convert.ToBase64String(ims.ToArray());
                iconCt = iconPart.ContentType;
            }

            // VML anchor shape: matched by the conventional _x0000_s{shapeId}
            // id AddOle and Excel both write.
            string? shapeXml = null;
            if (vmlXml != null && ole.ShapeId?.Value is { } spid)
            {
                var m = System.Text.RegularExpressions.Regex.Match(vmlXml,
                    $@"<v:shape [^>]*id=""_x0000_s{spid}"".*?</v:shape>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (m.Success) shapeXml = m.Value;
            }

            // Canonicalize attribute order so the slice is byte-stable
            // regardless of how the source element was authored: the SDK
            // preserves authoring order (AddOle's typed construction vs the
            // outerXml-ctor parse path differ on where xmlns declarations
            // land), which would flip on every dump→replay→dump cycle and
            // break idempotency.
            var objectXml = CanonicalizeXmlAttributeOrder(child.OuterXml);
            // Kind from the source part type, mirroring the docx dump: content
            // type alone cannot classify (legacy .xls package parts carry
            // application/vnd.ms-excel, not an openxmlformats-officedocument CT).
            var kind = payloadPart is EmbeddedPackagePart ? "package" : "object";
            result.Add(new DumpOleSlice(
                relId!, Convert.ToBase64String(pms.ToArray()), kind, payloadPart.ContentType, ext,
                iconRid, iconB64, iconCt,
                shapeXml, objectXml));
        }
        return result;
    }

    /// <summary>Deterministic attribute order (namespace declarations first,
    /// then by namespace + local name) for XML slices whose byte-equality
    /// drives dump idempotency checks. Semantics are unchanged.</summary>
    internal static string CanonicalizeXmlAttributeOrder(string xml)
    {
        var root = System.Xml.Linq.XElement.Parse(xml);
        void Normalize(System.Xml.Linq.XElement e)
        {
            var attrs = e.Attributes()
                .OrderBy(a => a.IsNamespaceDeclaration ? 0 : 1)
                .ThenBy(a => a.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(a => a.Name.LocalName, StringComparer.Ordinal)
                .ToList();
            e.RemoveAttributes();
            foreach (var a in attrs) e.Add(a);
            foreach (var c in e.Elements()) Normalize(c);
        }
        Normalize(root);
        return root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    /// <summary>One verbatim chartEx (cx:) carrier per extended chart:
    /// pinned rIds + base64 part payloads + the hosting TwoCellAnchor XML
    /// slice. No semantic vocabulary exists for waterfall/funnel/sunburst —
    /// dump carries the parts byte-for-byte (mirrors pptx SmartArt).</summary>
    public sealed record DumpChartExSlice(
        string RelId, string XmlBase64,
        string? ColorsRelId, string? ColorsXmlBase64,
        string? StyleRelId, string? StyleXmlBase64,
        string AnchorXml);

    public List<DumpChartExSlice> GetDumpChartExSlices(string sheetName)
    {
        var result = new List<DumpChartExSlice>();
        var worksheet = FindWorksheet(sheetName);
        var drawingsPart = worksheet?.DrawingsPart;
        var wsDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheet == null || drawingsPart == null || wsDrawing == null) return result;

        static string ReadPartBase64(OpenXmlPart part)
        {
            using var s = part.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return Convert.ToBase64String(ms.ToArray());
        }

        foreach (var anchor in wsDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
        {
            // A chartEx frame references its part via <cx:chart r:id> inside
            // a graphicData with the chartex uri.
            var cxRef = anchor.Descendants()
                .FirstOrDefault(e => e.LocalName == "chart"
                    && e.NamespaceUri.Contains("chartex", StringComparison.OrdinalIgnoreCase));
            if (cxRef == null) continue;
            var relId = cxRef.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "id").Value;
            if (string.IsNullOrEmpty(relId)) continue;
            if (drawingsPart.GetPartById(relId!) is not ExtendedChartPart extPart) continue;

            string? colorsRid = null, colorsXml = null, styleRid = null, styleXml = null;
            foreach (var sub in extPart.Parts)
            {
                switch (sub.OpenXmlPart)
                {
                    case ChartColorStylePart cp:
                        colorsRid = sub.RelationshipId; colorsXml = ReadPartBase64(cp); break;
                    case ChartStylePart sp:
                        styleRid = sub.RelationshipId; styleXml = ReadPartBase64(sp); break;
                }
            }
            // Canonicalize attribute order for the same reason as the OLE
            // slice: the SDK preserves authoring order, so xmlns placement
            // flipped between the typed-construction and outerXml-parse
            // paths, breaking dump→replay→dump byte idempotency.
            result.Add(new DumpChartExSlice(
                relId!, ReadPartBase64(extPart),
                colorsRid, colorsXml, styleRid, styleXml,
                CanonicalizeXmlAttributeOrder(anchor.OuterXml)));
        }
        return result;
    }

    /// <summary>Total-row rectangles ("A4:D4") of tables with a totals row.
    /// The totals row is derived render output that `add table totalRow=true`
    /// regenerates on replay — the value baseline must exclude it (same rule
    /// as pivot location rectangles).</summary>
    public List<string> GetDumpTableTotalRowRects(string sheetName)
    {
        var result = new List<string>();
        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null) return result;
        foreach (var tdp in worksheet.TableDefinitionParts)
        {
            var tbl = tdp.Table;
            if (tbl == null || (tbl.TotalsRowCount?.Value ?? 0) == 0) continue;
            var refStr = tbl.Reference?.Value;
            if (string.IsNullOrEmpty(refStr) || !refStr!.Contains(':')) continue;
            var parts = refStr.Split(':');
            var (c1, _) = ParseCellReference(parts[0]);
            var (c2, r2) = ParseCellReference(parts[1]);
            result.Add($"{c1}{r2}:{c2}{r2}");
        }
        return result;
    }

    /// <summary>Per-column totals-row function tokens ("none,sum,average")
    /// for table[index] (1-based, TableDefinitionParts order — same space as
    /// Get's table[N]). Null when the table has no totals row. Feeds the
    /// emitter's `totalsRowFunction=` prop so replay rebuilds each column's
    /// aggregation exactly instead of AddTable's label+SUM default.</summary>
    public string? GetDumpTableTotalsTokens(string sheetName, int index)
    {
        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null) return null;
        var tdps = worksheet.TableDefinitionParts.ToList();
        if (index < 1 || index > tdps.Count) return null;
        var tbl = tdps[index - 1].Table;
        if (tbl == null || (tbl.TotalsRowCount?.Value ?? 0) == 0) return null;
        var cols = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>().ToList();
        if (cols == null) return null;
        return string.Join(",", cols.Select(c =>
            c.TotalsRowFunction?.InnerText?.ToLowerInvariant() ?? "none"));
    }

    /// <summary>Drawing-layer counts (pictures / leaf shapes) matching the
    /// picture[N] / shape[N] Get index spaces.</summary>
    public (int Pictures, int Shapes) GetDumpDrawingCounts(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var wsDrawing = worksheet.DrawingsPart?.WorksheetDrawing;
        if (wsDrawing == null) return (0, 0);
        var pictures = EnumeratePictureAnchors(wsDrawing).Count();
        var shapes = EnumerateLeafShapes(wsDrawing).Count();
        return (pictures, shapes);
    }

    /// <summary>Exact twoCellAnchor rectangle for a drawing, in the x/y +
    /// EMU width/height vocabulary AddPicture/AddShape consume. Width and
    /// height invert the Add path's whole-cells + remainder-offset split
    /// (EmuPerColApprox / EmuPerRowApprox), so replaying `width=NNNemu`
    /// reconstructs the source's To marker bit-for-bit — unlike the Get
    /// surface's whole-cell col/row deltas, which discard sub-cell offsets.
    /// HasFromOffset flags anchors whose From marker carries a non-zero
    /// offset (real-Excel-authored); the add vocabulary cannot express that,
    /// so the emitter surfaces it as a warning.</summary>
    /// <summary>Mode is the anchor kind ("twoCell" / "oneCell" / "absolute").
    /// For absolute anchors X/Y are unused; XEmu/YEmu carry the EMU position
    /// (the add vocabulary takes x/y as EMU for anchorMode=absolute).</summary>
    public sealed record DumpAnchorEmu(int X, int Y, long WidthEmu, long HeightEmu, bool HasFromOffset,
        string Mode = "twoCell", long XEmu = 0, long YEmu = 0);

    public DumpAnchorEmu? GetDumpPictureAnchorEmu(string sheetName, int index)
    {
        var wsDrawing = FindWorksheet(sheetName)?.DrawingsPart?.WorksheetDrawing;
        if (wsDrawing == null) return null;
        var picAnchors = EnumeratePictureAnchors(wsDrawing).ToList();
        if (index < 1 || index > picAnchors.Count) return null;
        return AnchorToEmuRect(picAnchors[index - 1]);
    }

    public DumpAnchorEmu? GetDumpShapeAnchorEmu(string sheetName, int index)
    {
        var wsDrawing = FindWorksheet(sheetName)?.DrawingsPart?.WorksheetDrawing;
        if (wsDrawing == null) return null;
        var shapes = EnumerateLeafShapes(wsDrawing).ToList();
        if (index < 1 || index > shapes.Count) return null;
        return AnchorToEmuRect(shapes[index - 1].anchor);
    }

    private static DumpAnchorEmu? AnchorToEmuRect(OpenXmlCompositeElement anchorEl)
    {
        // oneCell / absolute anchors carry size in <xdr:ext> (EMU) instead of
        // a To marker; absolute additionally carries position as EMU <xdr:pos>.
        if (anchorEl is DocumentFormat.OpenXml.Drawing.Spreadsheet.OneCellAnchor one)
        {
            var oFrom = one.FromMarker;
            var oExt = one.GetFirstChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent>();
            if (oFrom == null || oExt?.Cx?.HasValue != true || oExt.Cy?.HasValue != true) return null;
            static int OP(string? s) => int.TryParse(s, out var v) ? v : 0;
            static long OPL(string? s) => long.TryParse(s, out var v) ? v : 0;
            return new DumpAnchorEmu(OP(oFrom.ColumnId?.Text), OP(oFrom.RowId?.Text),
                oExt.Cx.Value, oExt.Cy.Value,
                OPL(oFrom.ColumnOffset?.Text) != 0 || OPL(oFrom.RowOffset?.Text) != 0,
                Mode: "oneCell");
        }
        if (anchorEl is DocumentFormat.OpenXml.Drawing.Spreadsheet.AbsoluteAnchor abs)
        {
            var aPos = abs.GetFirstChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.Position>();
            var aExt = abs.GetFirstChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent>();
            if (aExt?.Cx?.HasValue != true || aExt.Cy?.HasValue != true) return null;
            return new DumpAnchorEmu(0, 0, aExt.Cx.Value, aExt.Cy.Value, false,
                Mode: "absolute", XEmu: aPos?.X?.Value ?? 0, YEmu: aPos?.Y?.Value ?? 0);
        }
        if (anchorEl is not DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor anchor) return null;
        var from = anchor.FromMarker;
        var to = anchor.ToMarker;
        if (from == null || to == null) return null;
        static int P(string? s) => int.TryParse(s, out var v) ? v : 0;
        static long PL(string? s) => long.TryParse(s, out var v) ? v : 0;
        int fc = P(from.ColumnId?.Text), fr = P(from.RowId?.Text);
        long fcOff = PL(from.ColumnOffset?.Text), frOff = PL(from.RowOffset?.Text);
        long wEmu = (P(to.ColumnId?.Text) - fc) * EmuPerColApprox + PL(to.ColumnOffset?.Text) - fcOff;
        long hEmu = (P(to.RowId?.Text) - fr) * EmuPerRowApprox + PL(to.RowOffset?.Text) - frOff;
        if (wEmu <= 0 || hEmu <= 0) return null;
        return new DumpAnchorEmu(fc, fr, wEmu, hEmu, fcOff != 0 || frOff != 0);
    }

    /// <summary>
    /// Extract picture[index]'s image bytes as a data URI for the emit's
    /// `add picture src=` prop (ImageSource.Resolve round-trips data URIs).
    /// SVG dual-representation aware: when the blip carries an
    /// asvg:svgBlip extension, the TRUE source is the SVG part (the r:embed
    /// PNG is just the fallback AddPicture regenerates on replay).
    ///
    /// CONSISTENCY(picture-inline-base64): floating-picture bytes ride the same
    /// inline `data:<contentType>;base64,<bytes>` carrier as the in-cell image
    /// and the pptx/word picture round-trip (PptxBatchEmitter.Media.cs) — no
    /// sidecar file, so a dump script stays self-contained.
    /// </summary>
    public string? GetDumpPictureDataUri(string sheetName, int index)
    {
        var worksheet = FindWorksheet(sheetName);
        var drawingsPart = worksheet?.DrawingsPart;
        var wsDrawing = drawingsPart?.WorksheetDrawing;
        if (worksheet == null || drawingsPart == null || wsDrawing == null) return null;
        var picAnchors = EnumeratePictureAnchors(wsDrawing).ToList();
        if (index < 1 || index > picAnchors.Count) return null;
        var picture = picAnchors[index - 1]
            .Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().First();
        var blip = picture.BlipFill?.Blip;
        if (blip == null) return null;

        // SVG extension takes precedence over the PNG fallback embed.
        var svgBlip = blip.Descendants()
            .FirstOrDefault(e => e.LocalName == "svgBlip");
        var relId = svgBlip?.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "embed").Value
            ?? blip.Embed?.Value;
        if (string.IsNullOrEmpty(relId)) return null;

        if (drawingsPart.GetPartById(relId!) is not ImagePart imgPart) return null;
        using var s = imgPart.GetStream();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return $"data:{imgPart.ContentType};base64,{Convert.ToBase64String(ms.ToArray())}";
    }

    /// <summary>
    /// Cheap package-part scan for per-sheet content the batch emit does not
    /// round-trip yet. Mirrors PptxBatchEmitter's EmitAuxiliaryPartsScan
    /// contract: silent data loss is worse than a noisy warning.
    /// </summary>
    // Worksheet children the batch emitter round-trips (or that replay
    // regenerates as a side effect). Anything outside this set is content
    // dump will silently lose — surface it as an unsupported_feature warning
    // instead (mirrors PptxBatchEmitter.EmitAuxiliaryPartsScan).
    private static readonly HashSet<string> DumpHandledWorksheetChildren = new(StringComparer.Ordinal)
    {
        "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols",
        "sheetData", "sheetProtection", "autoFilter", "sortState",
        "mergeCells", "conditionalFormatting", "dataValidations",
        "hyperlinks", "printOptions", "pageMargins", "pageSetup",
        "headerFooter", "rowBreaks", "colBreaks", "drawing",
        "legacyDrawing", "legacyDrawingHF", "oleObjects", "tableParts",
        "extLst", "phoneticPr", "sheetCalcPr", "ignoredErrors",
    };

    // extLst ext URIs the emitter consumes (sparklines, slicer list, x14
    // conditional formatting, x14 data validations, timeline).
    private static readonly HashSet<string> DumpHandledExtUris = new(StringComparer.OrdinalIgnoreCase)
    {
        "{05C60535-1F16-4fd2-B633-E4A46CF9E463}",
        "{725AE2AE-9491-48be-B2B4-4EB974FC3084}",
        "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
        "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}",
        "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",
    };

    public List<(string Element, string Reason)> GetDumpUnsupportedFeatures(string sheetName)
    {
        var result = new List<(string, string)>();
        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null) return result;
        var ws = GetSheet(worksheet);

        foreach (var child in ws.ChildElements)
        {
            var name = child.LocalName;
            if (DumpHandledWorksheetChildren.Contains(name)) continue;
            result.Add((name,
                $"worksheet element <{name}> is not round-tripped by dump; it will be missing after replay"));
        }

        var wsExtLst = ws.GetFirstChild<WorksheetExtensionList>();
        if (wsExtLst != null)
        {
            foreach (var ext in wsExtLst.Elements<WorksheetExtension>())
            {
                var uri = ext.Uri?.Value ?? "";
                if (DumpHandledExtUris.Contains(uri)) continue;
                result.Add(($"extLst ext {uri}",
                    $"worksheet extension {uri} is not round-tripped by dump; its content will be missing after replay"));
            }
        }

        // Threaded comments (Excel 365) live in a relationship-linked
        // WorksheetThreadedCommentsPart, not in the worksheet XML, so the
        // child/extLst scans above never see them. The batch emitter only
        // transcribes the legacy comments part, so any threaded comment is
        // silently lost on replay — surface it here.
        var threadedCount = worksheet.WorksheetThreadedCommentsParts
            .Sum(p => p.ThreadedComments?.Elements<ThreadedCmt.ThreadedComment>().Count() ?? 0);
        if (threadedCount > 0)
        {
            result.Add(("threadedComments",
                $"{threadedCount} threaded comment(s) are not round-tripped by dump; they will be missing after replay"));
        }
        return result;
    }
}
