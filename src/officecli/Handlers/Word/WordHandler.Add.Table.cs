// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddTable(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var table = new Table();
        // BUG-R2-P1-5: always seed all 6 default borders (top/bottom/left/right/
        // insideH/insideV at Single/4), then apply user-supplied border.* props
        // on top. Previously a partial border spec (e.g. just border.top +
        // border.left) wiped the other four sides, surprising users who
        // expected partial-override semantics. To express a genuine three-line
        // table (top/bottom only), pass border=none first to wipe defaults,
        // then border.top + border.bottom. CONSISTENCY(border-default-overlay).
        TableProperties tblProps = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            )
        );
        table.AppendChild(tblProps);
        // Apply user-supplied border.* props in order; "border" / "border.all"
        // (with value "none") wipes defaults before per-side props overlay.
        var orderedBorderProps = properties
            .Where(kv => kv.Key.StartsWith("border", StringComparison.OrdinalIgnoreCase))
            .OrderBy(kv =>
            {
                var k = kv.Key.ToLowerInvariant();
                return (k == "border" || k == "border.all") ? 0 : 1;
            })
            .ToList();
        foreach (var (bk, bv) in orderedBorderProps)
        {
            ApplyTableBorders(tblProps, bk, bv);
        }

        // Parse data if provided: "H1,H2;R1C1,R1C2;R2C1,R2C2" or CSV file/URL/data-URI
        string[][]? tableData = null;
        if (properties.TryGetValue("data", out var dataStr))
        {
            if (OfficeCli.Core.FileSource.IsResolvable(dataStr))
                tableData = OfficeCli.Core.FileSource.ResolveLines(dataStr)
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .Select(l => l.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
            else
                tableData = dataStr.Split(';')
                    .Select(r => r.Split(',').Select(c => c.Trim()).ToArray())
                    .ToArray();
        }

        int rows, cols;
        if (tableData != null)
        {
            rows = tableData.Length;
            cols = tableData.Max(r => r.Length);
        }
        else
        {
            rows = 1;
            if (properties.TryGetValue("rows", out var rowsStr))
            {
                if (!int.TryParse(rowsStr, out rows))
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Expected a positive integer.");
                if (rows <= 0)
                    throw new ArgumentException($"Invalid 'rows' value: '{rowsStr}'. Must be a positive integer (> 0).");
            }
            cols = 1;
            if (properties.TryGetValue("cols", out var colsStr))
            {
                cols = ParseHelpers.SafeParseInt(colsStr, "cols");
                if (cols <= 0)
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Must be a positive integer (> 0).");
            }
        }

        // Parse per-column widths: colWidths="3000,2000,5000"
        int[]? colWidthArr = null;
        if (properties.TryGetValue("colwidths", out var cwStr) || properties.TryGetValue("colWidths", out cwStr))
        {
            var parts = cwStr.Split(',');
            colWidthArr = new int[parts.Length];
            for (int ci = 0; ci < parts.Length; ci++)
            {
                if (!int.TryParse(parts[ci].Trim(), out colWidthArr[ci]))
                    throw new ArgumentException($"Invalid 'colwidths' value: '{parts[ci].Trim()}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
                // BUG-R1-01: reject negative or zero up front (Set already
                // does this; Add did not). Invalid OOXML otherwise.
                if (colWidthArr[ci] <= 0)
                    throw new ArgumentException($"Invalid 'colwidths' value: '{parts[ci].Trim()}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
            }
        }

        // BUG-R9-B1: when caller passes colWidths=... without cols=, infer
        // the column count from colWidths.Length so the tblGrid + downstream
        // row-cell loops produce the right number of columns. Previously
        // cols defaulted to 1 and only one column was emitted, silently
        // dropping the rest of the widths.
        if (colWidthArr != null && cols < colWidthArr.Length)
            cols = colWidthArr.Length;

        // gridCols=N opt-out for sources whose <w:tblGrid/> is intentionally
        // empty (cells carry their own tcW, or auto-fit + no explicit widths).
        // Without this, AddTable always seeded `cols` GridColumn entries which
        // dump's back-fill (ReadCellProps width-from-gridCol mitigation) then
        // surfaced as `set tc width=…` rows that the source dump never had —
        // up to N×M extra commands per round-trip on test.docx-style tables.
        // gridCols defaults to cols; gridCols=0 yields an empty <w:tblGrid/>.
        int gridCols = cols;
        bool gridColsExplicit = false;
        if (properties.TryGetValue("gridcols", out var gridColsStr)
            || properties.TryGetValue("gridCols", out gridColsStr))
        {
            gridCols = ParseHelpers.SafeParseInt(gridColsStr, "gridCols");
            if (gridCols < 0)
                throw new ArgumentException($"Invalid 'gridCols' value: '{gridColsStr}'. Must be 0 or a positive integer.");
            gridColsExplicit = true;
        }

        // Add table grid
        // BUG-R1-P0-4: when colWidths is not specified, default per-column
        // width should be computed from the section's usable body width
        // (page width − left/right margins) divided by `cols`. The previous
        // hard-coded 2400-twips default overflowed the page once cols > 3
        // on default A4 / Letter section properties.
        long defaultColTwips = 2400;
        if (colWidthArr == null)
        {
            var sectPr = _doc.MainDocumentPart?.Document?.Body?
                .Descendants<SectionProperties>().LastOrDefault();
            var pgSz = sectPr?.GetFirstChild<PageSize>();
            var pgMar = sectPr?.GetFirstChild<PageMargin>();
            long pageW = pgSz?.Width?.Value ?? 12240u;
            long mL = pgMar?.Left?.Value ?? 1440u;
            long mR = pgMar?.Right?.Value ?? 1440u;
            long usable = Math.Max(1, pageW - mL - mR);
            defaultColTwips = Math.Max(1, usable / Math.Max(1, cols));
        }

        var tblGrid = new TableGrid();
        for (int gc = 0; gc < gridCols; gc++)
        {
            // BUG-R1-01: reject negative or zero gridCol widths up front
            // (Set already does this; Add did not). Invalid OOXML otherwise.
            if (colWidthArr != null && gc < colWidthArr.Length)
            {
                if (colWidthArr[gc] <= 0)
                    throw new ArgumentException($"Invalid 'colwidths' value: '{colWidthArr[gc]}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
            }
            var w = colWidthArr != null && gc < colWidthArr.Length
                ? colWidthArr[gc].ToString()
                : defaultColTwips.ToString();
            tblGrid.AppendChild(new GridColumn { Width = w });
        }
        table.AppendChild(tblGrid);

        // BUG-R8-H1: default <w:tblW> from sum of gridCol widths when the user
        // did not provide width=... explicitly. Without tblW, Word switches to
        // auto-fit and squashes columns to the visible text width, ignoring the
        // tblGrid we just wrote. The user-supplied width= path below overrides
        // this default when present (assignment to tblProps.TableWidth wins).
        // Skip the auto-tblW path when gridCols=0 was explicitly requested
        // (empty tblGrid → no widths to sum → emitting `<w:tblW w:w="0"/>` would
        // poison the dump round-trip with a width key the source never had).
        // Also skip when caller passed `skipTblW=true` (dump-replay path for
        // sources whose <w:tbl> has no <w:tblW> element — leaves Word to
        // auto-fit from gridCol widths, matching the source's behaviour).
        bool skipTblW = properties.TryGetValue("skiptblw", out var stbw)
                     || properties.TryGetValue("skipTblW", out stbw);
        bool skipTblWRequested = skipTblW && IsTruthy(stbw);
        if (!properties.ContainsKey("width") && !(gridColsExplicit && gridCols == 0) && !skipTblWRequested)
        {
            long totalTwips = 0;
            for (int gc = 0; gc < gridCols; gc++)
            {
                totalTwips += colWidthArr != null && gc < colWidthArr.Length
                    ? colWidthArr[gc]
                    : defaultColTwips;
            }
            tblProps.TableWidth = new TableWidth
            {
                Width = totalTwips.ToString(),
                Type = TableWidthUnitValues.Dxa
            };
        }

        // Apply table-level properties from Add parameters
        // Set of keys the switch below consumes. Used to mark a key as
        // accessed via ContainsKey only when a case actually matched, so
        // genuine typos still fall through to the tracker's UnusedKeys.
        var tblBareConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "align", "alignment", "width", "indent", "cellspacing", "layout",
            "padding", "padding.top", "padding.bottom", "padding.left", "padding.right",
            "style", "shd", "shading", "direction", "dir", "bidi",
            // CONSISTENCY(add-set-symmetry): mirror Set's tblPr-level cases.
            "overlap",
        };
        foreach (var (tk, tv) in properties)
        {
            var tkl = tk.ToLowerInvariant();
            // BUG-R9 (tbllook.* compound key): strip the "tbllook." namespace
            // prefix so callers can write tblLook.firstRow=true alongside the
            // bare `firstRow=true` form. Sub-keys must resolve to a known
            // tblLook leaf — unknown sub-keys raise instead of being silently
            // dropped (and falsely reporting "Updated" via Set).
            if (tkl.StartsWith("tbllook."))
            {
                var sub = tkl.Substring("tbllook.".Length);
                if (sub is "firstrow" or "lastrow"
                        or "firstcol" or "firstcolumn"
                        or "lastcol" or "lastcolumn"
                        or "bandrow" or "bandedrows" or "bandrows"
                        or "bandcol" or "bandedcols" or "bandcols"
                        or "nohband" or "nohorizontalband"
                        or "novband" or "noverticalband")
                    tkl = sub;
                else
                    throw new ArgumentException(
                        $"Unknown tblLook sub-key: '{tk}'. Valid sub-keys: " +
                        $"firstRow, lastRow, firstCol, lastCol, bandRow, bandCol, " +
                        $"noHBand, noVBand. Or use the bare hex form tblLook=04A0.");
            }
            if (tkl is "rows" or "cols" or "colwidths" or "gridcols" or "skiptblw" || tkl.StartsWith("border")) continue;
            // ACCOUNTING(handler-as-truth): see AddStyle. ContainsKey only
            // when the switch will consume this key — otherwise typos would
            // leak past UnusedKeys detection.
            if (tblBareConsumed.Contains(tkl)) properties.ContainsKey(tk);
            switch (tkl)
            {
                case "align" or "alignment":
                    tblProps.TableJustification = new TableJustification
                    {
                        Val = tv.ToLowerInvariant() switch
                        {
                            "center" => TableRowAlignmentValues.Center,
                            "right" => TableRowAlignmentValues.Right,
                            "left" => TableRowAlignmentValues.Left,
                            _ => throw new ArgumentException($"Invalid table alignment value: '{tv}'. Valid values: left, center, right.")
                        }
                    };
                    break;
                case "width":
                    // BUG-DUMP19-03: accept "auto" so dump round-trip preserves
                    // <w:tblW w:type="auto"/>. Without this, SafeParseUint("auto")
                    // throws and the prop is silently dropped/normalized.
                    if (string.Equals(tv, "auto", StringComparison.OrdinalIgnoreCase))
                    {
                        tblProps.TableWidth = new TableWidth { Width = "0", Type = TableWidthUnitValues.Auto };
                    }
                    else if (tv.EndsWith('%'))
                    {
                        var pct = ParseHelpers.SafeParseInt(tv.TrimEnd('%'), "width") * 50;
                        tblProps.TableWidth = new TableWidth { Width = pct.ToString(), Type = TableWidthUnitValues.Pct };
                    }
                    else
                    {
                        // BUG-R8-H1: accept unit-qualified widths (cm/in/pt/dxa)
                        // mirror Set cell-width path. Previously SafeParseUint
                        // rejected width=10cm even though help docs showed cm.
                        // CONSISTENCY(unit-twips): ParseTwips is the canonical
                        // input-side twips converter for Word.
                        tblProps.TableWidth = new TableWidth { Width = WordHandler.ParseTwips(tv).ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "indent":
                    tblProps.TableIndentation = new TableIndentation { Width = ParseHelpers.SafeParseInt(tv, "indent"), Type = TableWidthUnitValues.Dxa };
                    break;
                case "cellspacing":
                    tblProps.TableCellSpacing = new TableCellSpacing { Width = ParseHelpers.SafeParseUint(tv, "cellspacing").ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                case "layout":
                    tblProps.TableLayout = new TableLayout
                    {
                        Type = tv.ToLowerInvariant() == "fixed" ? TableLayoutValues.Fixed : TableLayoutValues.Autofit
                    };
                    break;
                case "padding":
                {
                    // CONSISTENCY(tblpr-schema-order): tblCellMar must be inserted
                    // at rank 13 in CT_TblPrBase; raw AppendChild produced
                    // schema-invalid OOXML when other props (style→tblLook at rank
                    // 14) appeared earlier in argv. Same fix for padding.{top,
                    // bottom,left,right} below and Set padding path.
                    var paddingVal = ParseHelpers.SafeParseInt(tv, "padding");
                    if (paddingVal < 0)
                        throw new ArgumentException($"Invalid 'padding' value: '{tv}'. Table cell margins must be non-negative (OOXML w:tblCellMar).");
                    var dxa = paddingVal.ToString();
                    var cm = EnsureTableCellMarginDefault(tblProps);
                    cm.TopMargin = new TopMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    cm.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                    cm.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(paddingVal, short.MaxValue), Type = TableWidthValues.Dxa };
                    break;
                }
                // BUG-DUMP13-04: per-side default cell margins. WordBatchEmitter
                // passes asymmetric padding.* keys through unfolded when sides
                // differ; without these cases AddTable warned UNSUPPORTED and
                // the values became zero on round-trip. Mirrors the per-cell
                // tcMar handling in Set.Element.cs.
                case "padding.top":
                    {
                        var tv2 = ParseHelpers.SafeParseInt(tv, "padding.top");
                        if (tv2 < 0)
                            throw new ArgumentException($"Invalid 'padding.top' value: '{tv}'. Table cell margins must be non-negative.");
                        var cmt = EnsureTableCellMarginDefault(tblProps);
                        cmt.TopMargin = new TopMargin { Width = tv2.ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding.bottom":
                    {
                        var bv2 = ParseHelpers.SafeParseInt(tv, "padding.bottom");
                        if (bv2 < 0)
                            throw new ArgumentException($"Invalid 'padding.bottom' value: '{tv}'. Table cell margins must be non-negative.");
                        var cmb = EnsureTableCellMarginDefault(tblProps);
                        cmb.BottomMargin = new BottomMargin { Width = bv2.ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding.left":
                    {
                        var lv = ParseHelpers.SafeParseInt(tv, "padding.left");
                        if (lv < 0)
                            throw new ArgumentException($"Invalid 'padding.left' value: '{tv}'. Table cell margins must be non-negative.");
                        var cml = EnsureTableCellMarginDefault(tblProps);
                        cml.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(lv, short.MaxValue), Type = TableWidthValues.Dxa };
                    }
                    break;
                case "padding.right":
                    {
                        var rv = ParseHelpers.SafeParseInt(tv, "padding.right");
                        if (rv < 0)
                            throw new ArgumentException($"Invalid 'padding.right' value: '{tv}'. Table cell margins must be non-negative.");
                        var cmr = EnsureTableCellMarginDefault(tblProps);
                        cmr.TableCellRightMargin = new TableCellRightMargin { Width = (short)Math.Min(rv, short.MaxValue), Type = TableWidthValues.Dxa };
                    }
                    break;
                case "style":
                case "tablestyle":
                case "tablestyleid":
                    // BUG-R3 P1-#6: schema declares tableStyle/tableStyleId as
                    // aliases for `style`; honor them here so Add doesn't flag
                    // them UNSUPPORTED.
                    tblProps.TableStyle = new TableStyle { Val = tv };
                    // Add TableLook so built-in styles apply banding correctly.
                    // CONSISTENCY(tblpr-schema-order): tblLook is rank 14 and
                    // must precede tblCaption/tblDescription/tblPrChange — raw
                    // AppendChild produced schema-invalid order when those
                    // higher-ranked elements existed first (and was the root
                    // cause of issue #105 when combined with `padding`).
                    tblProps.RemoveAllChildren<TableLook>();
                    InsertTblPrChildInOrder(tblProps, new TableLook { Val = "04A0" });
                    break;
                case "shd" or "shading":
                    {
                        // BUG-DUMP21-01: w:tblPr/w:shd table-level shading
                        // round-trip. Mirrors paragraph/cell `shading` parsing
                        // — accepts FILL, VAL;FILL, or VAL;FILL;COLOR.
                        var shdParts = tv.Split(';');
                        var tShd = new Shading();
                        if (shdParts.Length == 1)
                        {
                            tShd.Val = ShadingPatternValues.Clear;
                            tShd.Fill = SanitizeHex(shdParts[0]);
                        }
                        else if (shdParts.Length >= 2)
                        {
                            var pat = shdParts[0].TrimStart('#');
                            if (pat.Length >= 6 && pat.All(char.IsAsciiHexDigit))
                            {
                                tShd.Val = ShadingPatternValues.Clear;
                                tShd.Fill = SanitizeHex(shdParts[0]);
                            }
                            else
                            {
                                tShd.Val = new ShadingPatternValues(shdParts[0]);
                                tShd.Fill = SanitizeHex(shdParts[1]);
                                if (shdParts.Length >= 3)
                                    tShd.Color = SanitizeHex(shdParts[2]);
                            }
                        }
                        tblProps.Shading = tShd;
                    }
                    break;
                case "overlap":
                {
                    // CONSISTENCY(add-set-symmetry): mirror Set's overlap case
                    // (Set.Element.cs:1752). CT_TblPr schema:
                    // tblStyle → tblpPr → tblOverlap → ...
                    tblProps.RemoveAllChildren<TableOverlap>();
                    if (!tv.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        var overlapEl = new TableOverlap
                        {
                            Val = tv.ToLowerInvariant() switch
                            {
                                "overlap" or "true" or "always" => TableOverlapValues.Overlap,
                                "never" or "false" => TableOverlapValues.Never,
                                _ => throw new ArgumentException($"Invalid overlap: '{tv}'. Valid: overlap, never, none.")
                            }
                        };
                        var tppRef = tblProps.GetFirstChild<TablePositionProperties>();
                        if (tppRef != null) tppRef.InsertAfterSelf(overlapEl);
                        else
                        {
                            var styleRef = tblProps.GetFirstChild<TableStyle>();
                            if (styleRef != null) styleRef.InsertAfterSelf(overlapEl);
                            else tblProps.PrependChild(overlapEl);
                        }
                    }
                    break;
                }
                case "direction" or "dir" or "bidi":
                    // Table-level bidi: emit <w:bidiVisual/> on tblPr in schema
                    // order. Mirrors paragraph/cell direction=rtl vocabulary.
                    // CONSISTENCY(rtl-cascade).
                    tblProps.RemoveAllChildren<BiDiVisual>();
                    if (ParseDirectionRtl(tv))
                        InsertTblPrChildInOrder(tblProps, new BiDiVisual());
                    break;
                // BUG-R4-02/08: tblLook props at Add time. Mirrors the Set.Element.cs
                // tblLook switch — accepts lowercase + camelCase aliases as input.
                // Without this, dump→batch round-trip silently lost firstRow etc.
                // CONSISTENCY(add-set-symmetry).
                case "firstrow":
                case "lastrow":
                case "firstcol" or "firstcolumn":
                case "lastcol" or "lastcolumn":
                case "bandrow" or "bandedrows" or "bandrows":
                case "bandcol" or "bandedcols" or "bandcols":
                case "nohband" or "nohorizontalband":
                case "novband" or "noverticalband":
                case "tbllook":
                    {
                        var tblLook = tblProps.GetFirstChild<TableLook>();
                        if (tblLook == null)
                        {
                            tblLook = new TableLook { Val = "04A0" };
                            InsertTblPrChildInOrder(tblProps, tblLook);
                        }
                        if (tkl == "tbllook")
                        {
                            // raw hex passthrough (e.g. tblLook=04A0)
                            tblLook.Val = tv;
                            break;
                        }
                        var bv = IsTruthy(tv);
                        switch (tkl)
                        {
                            case "firstrow": tblLook.FirstRow = bv; break;
                            case "lastrow": tblLook.LastRow = bv; break;
                            case "firstcol" or "firstcolumn": tblLook.FirstColumn = bv; break;
                            case "lastcol" or "lastcolumn": tblLook.LastColumn = bv; break;
                            case "bandrow" or "bandedrows" or "bandrows": tblLook.NoHorizontalBand = !bv; break;
                            case "bandcol" or "bandedcols" or "bandcols": tblLook.NoVerticalBand = !bv; break;
                            case "nohband" or "nohorizontalband": tblLook.NoHorizontalBand = bv; break;
                            case "novband" or "noverticalband": tblLook.NoVerticalBand = bv; break;
                        }
                        break;
                    }
            }
        }

        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                var cellText = tableData != null && r < tableData.Length && c < tableData[r].Length
                    ? tableData[r][c] : (properties.TryGetValue($"r{r + 1}c{c + 1}", out var rc) ? rc : "");
                // CONSISTENCY(table-cell-defaults): do not stamp explicit
                // spaceAfter=0 / lineSpacing=240 Auto on freshly-created cell
                // paragraphs — let them inherit from style/docDefaults like
                // regular body paragraphs. Otherwise dump→batch round-trip
                // grows 67 extra `set spaceAfter=0pt lineSpacing=1x` commands
                // per cell (BUG-R3-3).
                var cellPara = new Paragraph();
                AssignParaId(cellPara);
                if (!string.IsNullOrEmpty(cellText))
                    cellPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
                var cell = new TableCell(cellPara);
                // BUG-R6-06 / BUG-R6-01: do NOT stamp an explicit
                // <w:tcW> on every cell when the user supplied colWidths
                // — w:tblGrid/w:gridCol already encodes the column
                // widths, and per-cell tcW makes dump→batch→dump
                // non-idempotent (each round-trip emits N×M extra
                // `set width=…` commands). Cells without a tcW inherit
                // the column width from tblGrid as the schema intends.
                row.AppendChild(cell);
            }
            table.AppendChild(row);
        }

        // Dotted-key fallback for tblPr-level attrs not modeled by the
        // hand-rolled blocks above (single-attr forms like tblpPr.* or
        // future schema additions). CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            properties.ContainsKey(key);
            // border.{top,bottom,left,right,insideH,insideV,all} were already
            // applied at the top of AddTable via ApplyTableBorders. Skip them
            // here so they don't get mis-flagged UNSUPPORTED by the generic
            // TypedAttributeFallback (which doesn't model border.*).
            // CONSISTENCY(add-set-symmetry).
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase)) continue;
            // BUG-DUMP14-04: padding.{top,bottom,left,right} are handled by
            // the main switch above (round-13 added tblCellMar emit). Skip
            // them here so they aren't double-tagged as UNSUPPORTED by the
            // generic TypedAttributeFallback. Mirrors border.* skip.
            if (key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase)) continue;
            if (Core.TypedAttributeFallback.TrySet(tblProps, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
            InsertAtPosition(parent, table, index);
        else
            AppendToParent(parent, table);
        var tbls = parent.Elements<Table>().ToList();
        var idx = tbls.FindIndex(t => ReferenceEquals(t, table));
        return $"{parentPath}/tbl[{(idx >= 0 ? idx + 1 : tbls.Count)}]";
    }

    private string AddRow(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not Table targetTable)
            throw new ArgumentException("Rows can only be added to a table: /body/tbl[N]");

        var grid = targetTable.GetFirstChild<TableGrid>()
            ?? targetTable.PrependChild(new TableGrid());
        var existingGridCols = grid.Elements<GridColumn>().ToList();
        var existingCols = existingGridCols.Count > 0 ? existingGridCols.Count : 1;
        int newCols = existingCols;
        if (properties.TryGetValue("cols", out var colsVal))
        {
            newCols = ParseHelpers.SafeParseInt(colsVal, "cols");
            // BUG-R1-P0-3a: cols=0 silently produces an empty <w:tr> with no
            // cells; per OOXML spec a row must contain at least one cell.
            if (newCols <= 0)
                throw new ArgumentException($"Invalid 'cols' value: '{colsVal}'. Must be a positive integer (> 0); a row with 0 cells is invalid OOXML.");
        }

        // BUG-R1-P0-3b: cols > existing tblGrid count must expand tblGrid
        // to keep tcW / gridCol in agreement. Otherwise the extra cells
        // have no column-width definition and Word misaligns them.
        // BUG-R2-P0-2: extending the grid alone leaves already-existing rows
        // with fewer cells than the grid claims. Word renders the missing
        // slots as a half-collapsed final column. Pad each existing row with
        // empty placeholder cells so per-row cell count tracks the new grid.
        if (existingGridCols.Count > 0 && newCols > existingGridCols.Count)
        {
            // Width: average of existing cols, falling back to 2400.
            long avg = (long)existingGridCols.Average(gc =>
                long.TryParse(gc.Width?.Value, out var w) ? w : 2400L);
            int oldCount = existingGridCols.Count;
            for (int extra = oldCount; extra < newCols; extra++)
                grid.AppendChild(new GridColumn { Width = avg.ToString() });

            int padPerRow = newCols - oldCount;
            foreach (var existingRow in targetTable.Elements<TableRow>())
            {
                for (int i = 0; i < padPerRow; i++)
                {
                    var pad = new TableCell(new Paragraph());
                    AssignParaId(pad.GetFirstChild<Paragraph>()!);
                    existingRow.AppendChild(pad);
                }
            }
        }

        var newRow = new TableRow();
        TableRowProperties? newRowProps = null;
        if (properties.TryGetValue("height", out var rowHeight))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeight), HeightType = HeightRuleValues.AtLeast });
        }
        if (properties.TryGetValue("height.exact", out var rowHeightExact))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.GetFirstChild<TableRowHeight>()?.Remove();
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeightExact), HeightType = HeightRuleValues.Exact });
        }
        if (properties.TryGetValue("header", out var headerVal) && IsTruthy(headerVal))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableHeader());
        }
        // CONSISTENCY(add-set-symmetry): mirror Set's cantsplit case
        // (Set.Element.cs:1504). Row stays together across pages when true.
        if (properties.TryGetValue("cantSplit", out var cantSplitVal)
            || properties.TryGetValue("cantsplit", out cantSplitVal))
        {
            if (IsTruthy(cantSplitVal))
            {
                newRowProps ??= newRow.AppendChild(new TableRowProperties());
                if (newRowProps.GetFirstChild<CantSplit>() == null)
                    newRowProps.AppendChild(new CantSplit());
            }
        }

        for (int c = 0; c < newCols; c++)
        {
            var cellText = properties.TryGetValue($"c{c + 1}", out var ct) ? ct : "";
            var cellPara = new Paragraph();
            AssignParaId(cellPara);
            if (!string.IsNullOrEmpty(cellText))
                cellPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
            newRow.AppendChild(new TableCell(cellPara));
        }

        // Dotted-key fallback for trPr-level attrs (trHeight.*, etc.) not
        // modeled by hand-rolled blocks. Lazy-create trPr if any dotted
        // attr binds. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            properties.ContainsKey(key);
            var trPrTarget = newRowProps ?? new TableRowProperties();
            if (Core.TypedAttributeFallback.TrySet(trPrTarget, key, value))
            {
                if (newRowProps == null)
                {
                    newRow.PrependChild(trPrTarget);
                    newRowProps = trPrTarget;
                }
                continue;
            }
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
        {
            var existingRows = targetTable.Elements<TableRow>().ToList();
            if (index.Value < existingRows.Count)
                targetTable.InsertBefore(newRow, existingRows[index.Value]);
            else
                targetTable.AppendChild(newRow);
        }
        else
        {
            targetTable.AppendChild(newRow);
        }

        var rowIdx = targetTable.Elements<TableRow>().ToList().IndexOf(newRow) + 1;
        return $"{parentPath}/tr[{rowIdx}]";
    }

    /// <summary>
    /// Insert a new virtual column into a Word table. OOXML has no <w:col>
    /// element, so this synthesizes one by inserting a <w:gridCol> in
    /// <w:tblGrid> and a fresh <w:tc> at the same positional index in every
    /// existing <w:tr>. Rejects when any affected row carries gridSpan or
    /// vMerge in that column slot — those merge directives reference column
    /// positions and would silently break.
    /// </summary>
    private string AddTableColumn(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not Table targetTable)
            throw new ArgumentException("Columns can only be added to a table: /body/tbl[N]");

        var grid = targetTable.GetFirstChild<TableGrid>()
            ?? targetTable.PrependChild(new TableGrid());
        var existingGridCols = grid.Elements<GridColumn>().ToList();
        var insertIdx = index.HasValue && index.Value >= 0 && index.Value < existingGridCols.Count
            ? index.Value
            : existingGridCols.Count; // append by default

        // Reject if any row at insertIdx straddles the boundary via merge.
        foreach (var row in targetTable.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            // Check the cell currently occupying slot `insertIdx` (the one
            // that will be pushed right). gridSpan or vMerge here means
            // re-indexing the column slot would split a merged region.
            if (insertIdx < cells.Count && CellHasMerge(cells[insertIdx]))
                throw new ArgumentException(
                    $"Cannot insert column at index {insertIdx} of {parentPath}: " +
                    $"a row contains a merged cell straddling that boundary (gridSpan/vMerge). " +
                    "Unmerge first or pick a different position.");
        }

        // Width: explicit, or average of existing cols, or default 2400 twips
        long defaultWidthTwips = 2400;
        long newWidth = properties.TryGetValue("width", out var wVal)
            ? ParseTwips(wVal)
            : (existingGridCols.Count > 0
                ? (long)existingGridCols.Average(gc => long.TryParse(gc.Width?.Value, out var w) ? w : defaultWidthTwips)
                : defaultWidthTwips);

        var newGridCol = new GridColumn { Width = newWidth.ToString() };
        if (insertIdx < existingGridCols.Count)
            grid.InsertBefore(newGridCol, existingGridCols[insertIdx]);
        else
            grid.AppendChild(newGridCol);

        var cellText = properties.GetValueOrDefault("text", "");
        foreach (var row in targetTable.Elements<TableRow>())
        {
            var newPara = new Paragraph();
            AssignParaId(newPara);
            if (!string.IsNullOrEmpty(cellText))
                newPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
            var newCell = new TableCell(newPara);

            var cells = row.Elements<TableCell>().ToList();
            if (insertIdx < cells.Count)
                row.InsertBefore(newCell, cells[insertIdx]);
            else
                row.AppendChild(newCell);
        }

        var newColIdx = grid.Elements<GridColumn>().ToList().IndexOf(newGridCol) + 1;
        return $"{parentPath}/col[{newColIdx}]";
    }

    /// <summary>
    /// True if the cell carries gridSpan > 1 (horizontal merge) or any
    /// vMerge directive (vertical merge — restart or continue).
    /// </summary>
    private static bool CellHasMerge(TableCell cell)
    {
        var tcPr = cell.GetFirstChild<TableCellProperties>();
        if (tcPr == null) return false;
        var span = tcPr.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
        if (span > 1) return true;
        if (tcPr.GetFirstChild<VerticalMerge>() != null) return true;
        return false;
    }

    private string AddCell(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (parent is not TableRow targetRow)
            throw new ArgumentException("Cells can only be added to a table row: /body/tbl[N]/tr[M]");

        // BUG-R1-P0-2: AddCell on an existing row must keep tblGrid in sync.
        // Without this, the new cell has no matching <w:gridCol> and the
        // last "virtual column" collapses in Word. We synchronize lazily:
        // if the row's total grid-column occupancy after appending exceeds
        // the existing tblGrid, append matching gridCol entries averaging
        // the existing widths. Mirrors AddTableColumn's width logic.
        Table? cellParentTable = targetRow.Parent as Table;
        TableGrid? cellGrid = cellParentTable?.GetFirstChild<TableGrid>();

        var cellParagraph = new Paragraph();
        AssignParaId(cellParagraph);
        if (properties.TryGetValue("text", out var cellTxt))
            cellParagraph.AppendChild(new Run(new Text(cellTxt) { Space = SpaceProcessingModeValues.Preserve }));

        // Reading direction (Arabic / Hebrew). Mirrors AddParagraph: 'rtl'
        // writes <w:bidi/> on the cell paragraph's pPr and stamps <w:rtl/>
        // on the paragraph mark + any text run that was just appended.
        // CONSISTENCY(rtl-cascade).
        if (properties.TryGetValue("direction", out var cellDirRaw)
            || properties.TryGetValue("dir", out cellDirRaw)
            || properties.TryGetValue("bidi", out cellDirRaw))
        {
            bool cellRtl = ParseDirectionRtl(cellDirRaw);
            var cellPProps = cellParagraph.ParagraphProperties ?? cellParagraph.PrependChild(new ParagraphProperties());
            if (cellRtl) cellPProps.BiDi = new BiDi();
            var cellMarkRPr = cellPProps.ParagraphMarkRunProperties ?? cellPProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(cellMarkRPr, "direction", cellRtl ? "rtl" : "ltr");
            foreach (var existingRun in cellParagraph.Descendants<Run>())
                ApplyRunFormatting(EnsureRunProperties(existingRun), "direction", cellRtl ? "rtl" : "ltr");
        }

        var newCell = new TableCell(cellParagraph);

        if (properties.TryGetValue("width", out var cellWidth))
        {
            // BUG-DUMP6-04: accept "N%" alongside bare twips so dump→batch
            // round-trips pct cell widths. OOXML stores pct as fifths-of-percent.
            TableCellWidth tcw;
            if (cellWidth.EndsWith('%') &&
                double.TryParse(cellWidth.AsSpan(0, cellWidth.Length - 1),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var pctCw))
            {
                tcw = new TableCellWidth
                {
                    Width = ((int)Math.Round(pctCw * 50)).ToString(),
                    Type = TableWidthUnitValues.Pct
                };
            }
            else
            {
                tcw = new TableCellWidth { Width = cellWidth, Type = TableWidthUnitValues.Dxa };
            }
            newCell.PrependChild(new TableCellProperties(tcw));
        }

        // BUG-R2-P3-6: bare `fill` / `shd` / `shading` on AddCell were
        // silently dropped because the dotted-key fallback below only
        // visits keys containing '.'. Schema declares add:true for `fill`
        // on docx table-cell, so honour the contract. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            var keyLower = key.ToLowerInvariant();
            if (keyLower is "fill" or "shd" or "shading")
            {
                var tcPrFill = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                var shd = new Shading();
                var shdParts = value.Split(';');
                if (shdParts.Length == 1)
                {
                    shd.Val = ShadingPatternValues.Clear;
                    shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[0]).Rgb;
                }
                else if (shdParts.Length >= 2)
                {
                    var pat = shdParts[0].TrimStart('#');
                    if (pat.Length >= 6 && pat.All(char.IsAsciiHexDigit))
                    {
                        shd.Val = ShadingPatternValues.Clear;
                        shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[0]).Rgb;
                    }
                    else
                    {
                        shd.Val = new ShadingPatternValues(shdParts[0]);
                        shd.Fill = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[1]).Rgb;
                        if (shdParts.Length >= 3)
                            shd.Color = OfficeCli.Core.ParseHelpers.SanitizeColorForOoxml(shdParts[2]).Rgb;
                    }
                }
                tcPrFill.Shading = shd;
            }
        }

        // CONSISTENCY(add-set-symmetry): mirror Set's noWrap / hideMark cases
        // (Set.Element.cs:1342 + TryCreateTypedChild path).
        if (properties.TryGetValue("noWrap", out var noWrapVal)
            || properties.TryGetValue("nowrap", out noWrapVal))
        {
            if (IsTruthy(noWrapVal))
            {
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                if (tcPr.NoWrap == null) tcPr.NoWrap = new NoWrap();
            }
        }
        if (properties.TryGetValue("hideMark", out var hideMarkVal)
            || properties.TryGetValue("hidemark", out hideMarkVal))
        {
            if (IsTruthy(hideMarkVal))
            {
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                if (tcPr.GetFirstChild<HideMark>() == null)
                    tcPr.AppendChild(new HideMark());
            }
        }
        // CONSISTENCY(add-set-symmetry): bare vMerge / gridSpan / colspan
        // were Set-only; Add forced callers through the dotted form
        // (vMerge.val=restart, gridSpan.val=N) which works but is
        // surprising. Mirror Set's vocabulary so doc2 (and humans) can
        // write the natural form on Add too. v5.3 paragraph property.
        if (properties.TryGetValue("vMerge", out var vMergeAddVal)
            || properties.TryGetValue("vmerge", out vMergeAddVal))
        {
            if (!string.IsNullOrEmpty(vMergeAddVal))
            {
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                tcPr.VerticalMerge = vMergeAddVal.ToLowerInvariant() == "restart"
                    ? new VerticalMerge { Val = MergedCellValues.Restart }
                    : new VerticalMerge();
            }
        }
        if (properties.TryGetValue("gridSpan", out var gridSpanAddVal)
            || properties.TryGetValue("gridspan", out gridSpanAddVal)
            || properties.TryGetValue("colspan", out gridSpanAddVal))
        {
            var span = OfficeCli.Core.ParseHelpers.SafeParseInt(gridSpanAddVal, "gridSpan");
            if (span > 0)
            {
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                tcPr.GridSpan = new GridSpan { Val = span };
            }
        }
        // v6.4: horizontal merge on Add (CONSISTENCY(add-set-symmetry)).
        // Mirrors WordHandler.Set.Element.cs "hmerge" case but writes the
        // raw <w:hMerge/> attribute rather than gridSpan: dump replay from
        // doc2 already emits per-cell `hMerge=continue` for every
        // continuation cell, so we must preserve them as their own
        // <w:tc> nodes (gridSpan-collapsed continuations are unrepresented
        // in the source .doc and would corrupt subsequent vMerge anchors).
        // Set's gridSpan redirect is a UX optimisation for human callers;
        // for round-trip fidelity we want the legacy hMerge form.
        if (properties.TryGetValue("hMerge", out var hMergeAddVal)
            || properties.TryGetValue("hmerge", out hMergeAddVal))
        {
            if (!string.IsNullOrEmpty(hMergeAddVal))
            {
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                tcPr.HorizontalMerge = hMergeAddVal.ToLowerInvariant() == "restart"
                    ? new HorizontalMerge { Val = MergedCellValues.Restart }
                    : new HorizontalMerge();
            }
        }
        // CONSISTENCY(add-set-symmetry): cell-level w:tcMar (padding) was
        // Set-only; Add silently rejected it via UnusedKeys. Mirror Set's
        // padding / padding.{top,bottom,left,right} cases so doc2 round-
        // trips per-cell margins (sprmTCellPadding 0xD632/0xD634 -> tcMar).
        foreach (var (cmKey, cmVal) in properties)
        {
            var cmKl = cmKey.ToLowerInvariant();
            if (cmKl is not ("padding" or "padding.top" or "padding.bottom" or "padding.left" or "padding.right"))
                continue;
            properties.ContainsKey(cmKey);
            var tcPrCm = newCell.GetFirstChild<TableCellProperties>()
                ?? newCell.PrependChild(new TableCellProperties());
            var mar = tcPrCm.TableCellMargin ?? (tcPrCm.TableCellMargin = new TableCellMargin());
            if (cmKl == "padding")
            {
                var dxa = OfficeCli.Core.ParseHelpers.SafeParseUint(cmVal, "padding").ToString();
                mar.TopMargin    = new TopMargin    { Width = dxa, Type = TableWidthUnitValues.Dxa };
                mar.BottomMargin = new BottomMargin { Width = dxa, Type = TableWidthUnitValues.Dxa };
                mar.LeftMargin   = new LeftMargin   { Width = dxa, Type = TableWidthUnitValues.Dxa };
                mar.RightMargin  = new RightMargin  { Width = dxa, Type = TableWidthUnitValues.Dxa };
            }
            else
            {
                var v = OfficeCli.Core.ParseHelpers.SafeParseInt(cmVal, cmKl);
                if (v < 0)
                    throw new ArgumentException($"Invalid '{cmKl}' value: '{cmVal}'. Cell margins must be non-negative (OOXML w:tcMar).");
                var w = v.ToString();
                switch (cmKl)
                {
                    case "padding.top":    mar.TopMargin    = new TopMargin    { Width = w, Type = TableWidthUnitValues.Dxa }; break;
                    case "padding.bottom": mar.BottomMargin = new BottomMargin { Width = w, Type = TableWidthUnitValues.Dxa }; break;
                    case "padding.left":   mar.LeftMargin   = new LeftMargin   { Width = w, Type = TableWidthUnitValues.Dxa }; break;
                    case "padding.right":  mar.RightMargin  = new RightMargin  { Width = w, Type = TableWidthUnitValues.Dxa }; break;
                }
            }
        }

        // Dotted-key fallback for tcPr-level attrs (shd.fill, etc.) not
        // modeled by hand-rolled blocks. Lazy-create tcPr if any dotted
        // attr binds. CONSISTENCY(add-set-symmetry).
        foreach (var (key, value) in properties)
        {
            if (!key.Contains('.')) continue;
            // v5.3: padding.* already consumed by the cell-margin block
            // above; skip so the unsupported tracker doesn't double-flag.
            var kLower = key.ToLowerInvariant();
            if (kLower is "padding.top" or "padding.bottom" or "padding.left" or "padding.right")
                continue;
            // ACCOUNTING(handler-as-truth): see AddStyle for rationale.
            properties.ContainsKey(key);
            var tcPr = newCell.GetFirstChild<TableCellProperties>();
            var lazyTcPr = tcPr ?? new TableCellProperties();
            // CONSISTENCY(add-set-symmetry): route border.{top,bottom,left,
            // right,all,tl2br,tr2bl} through the same ApplyCellBorders helper
            // Set uses, instead of falling through to TypedAttributeFallback
            // which doesn't model border.* and would mis-flag UNSUPPORTED.
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase)
                || key.Equals("border", StringComparison.OrdinalIgnoreCase))
            {
                ApplyCellBorders(lazyTcPr, key, value);
                if (tcPr == null) newCell.PrependChild(lazyTcPr);
                continue;
            }
            if (Core.TypedAttributeFallback.TrySet(lazyTcPr, key, value))
            {
                if (tcPr == null) newCell.PrependChild(lazyTcPr);
                continue;
            }
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
        {
            var cells = targetRow.Elements<TableCell>().ToList();
            if (index.Value < cells.Count)
                targetRow.InsertBefore(newCell, cells[index.Value]);
            else
                targetRow.AppendChild(newCell);
        }
        else
        {
            targetRow.AppendChild(newCell);
        }

        // BUG-R1-P0-2: expand tblGrid if this row's grid-column occupancy
        // (sum of gridSpan) now exceeds existing gridCol count.
        // BUG-R1-table-merge: when expanding tblGrid, pad sibling rows with
        // empty placeholder cells so they remain aligned to the new column
        // count. CONSISTENCY(table-grid-pad): mirrors AddRow at lines 471-489.
        if (cellGrid != null && cellParentTable != null)
        {
            var existingGridCount = cellGrid.Elements<GridColumn>().Count();
            var rowSpan = targetRow.Elements<TableCell>().Sum(tc =>
                tc.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
            if (rowSpan > existingGridCount)
            {
                long avgWidth;
                var existingWidths = cellGrid.Elements<GridColumn>().ToList();
                avgWidth = existingWidths.Count > 0
                    ? (long)existingWidths.Average(gc => long.TryParse(gc.Width?.Value, out var w) ? w : 2400L)
                    : 2400L;
                for (int i = existingGridCount; i < rowSpan; i++)
                    cellGrid.AppendChild(new GridColumn { Width = avgWidth.ToString() });

                int padPerRow = rowSpan - existingGridCount;
                foreach (var siblingRow in cellParentTable.Elements<TableRow>())
                {
                    if (siblingRow == targetRow) continue;
                    for (int i = 0; i < padPerRow; i++)
                    {
                        var pad = new TableCell(new Paragraph());
                        AssignParaId(pad.GetFirstChild<Paragraph>()!);
                        siblingRow.AppendChild(pad);
                    }
                }
            }
        }

        var cellIdx = targetRow.Elements<TableCell>().ToList().IndexOf(newCell) + 1;
        return $"{parentPath}/tc[{cellIdx}]";
    }
}
