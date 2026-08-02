// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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
    // BUG-DUMP-R40-5: true when the caller supplied an explicit <w:tblLook>
    // value (bare hex form or any decomposed boolean facet, including the
    // `tblLook.` compound-key prefix). Used to suppress the default 04A0 seed
    // that the `style` case would otherwise force, so a source whose table had
    // no tblLook (or a non-default one) round-trips faithfully.
    private static bool TableHasExplicitTblLook(Dictionary<string, string> properties)
    {
        foreach (var k in properties.Keys)
        {
            var kl = k.ToLowerInvariant();
            if (kl.StartsWith("tbllook")) return true; // tbllook + tbllook.*
            switch (kl)
            {
                case "firstrow":
                case "lastrow":
                case "firstcol" or "firstcolumn":
                case "lastcol" or "lastcolumn":
                case "bandrow" or "bandedrows" or "bandrows":
                case "bandcol" or "bandedcols" or "bandcols":
                case "nohband" or "nohorizontalband":
                case "novband" or "noverticalband":
                    return true;
            }
        }
        return false;
    }

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
        //
        // `skipDefaultBorders=true` (dump→batch D1 round-trip; see
        // WordBatchEmitter.Table.cs) suppresses the seed so a follow-up
        // `set tbl --prop border=...` can snapshot a bare tblPr as the
        // tblPrChange "before" state — without this, the seeded borders
        // would already match the source's borders and the set's snapshot
        // would equal the current state, hiding the revision from Word's
        // reviewing pane.
        bool skipDefaultBorders =
            (properties.TryGetValue("skipDefaultBorders", out var sdbA) && IsTruthy(sdbA))
            || (properties.TryGetValue("skipdefaultborders", out var sdbB) && IsTruthy(sdbB));
        TableProperties tblProps;
        if (skipDefaultBorders)
        {
            tblProps = new TableProperties();
        }
        else
        {
            tblProps = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                )
            );
        }
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
            // Both forms are quote-aware: a cell wrapped in double quotes may
            // contain the separator, so `"Doe, John",30` is two cells. A plain
            // Split(',') made it three and shifted the rest of the row.
            // CONSISTENCY(table-data-parse): mirrored in the pptx add-table path.
            if (OfficeCli.Core.FileSource.IsResolvable(dataStr))
                tableData = OfficeCli.Core.DelimitedText.ParseGrid(
                    OfficeCli.Core.FileSource.ResolveText(dataStr), ',', '\n');
            else
                tableData = OfficeCli.Core.DelimitedText.ParseGrid(dataStr, ',', ';');
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
            // Accept both `cols` (canonical) and `columns` (the natural
            // English spelling AI agents reach for first). Without the
            // alias, `--prop columns=4` was silently dropped and the table
            // defaulted to 1 column — Add iterates `properties` via foreach
            // which marks every key consumed in the tracker, so the typo
            // didn't surface as an UNSUPPORTED warning either.
            if (properties.TryGetValue("cols", out var colsStr)
                || properties.TryGetValue("columns", out colsStr))
            {
                cols = ParseHelpers.SafeParseInt(colsStr, "cols");
                if (cols <= 0)
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. Must be a positive integer (> 0).");
                if (cols > 63)
                    throw new ArgumentException($"Invalid 'cols' value: '{colsStr}'. OOXML limits table columns to 63 (w:tblGrid max). Got: {cols}.");
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
                var part = parts[ci].Trim();
                // BUG-R2b: accept the unit-qualified forms Get now emits
                // (dxa suffix) as well as cm/in/pt, so a colWidths value read
                // from Get round-trips back through Add. ParseTwips handles all
                // four units and bare twips; mirror the Set colWidths path.
                uint parsedTwips;
                try { parsedTwips = ParseTwips(part); }
                catch (ArgumentException)
                {
                    throw new ArgumentException($"Invalid 'colwidths' value: '{part}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
                }
                colWidthArr[ci] = (int)parsedTwips;
                // BUG-R1-01: reject negative or zero up front (Set already
                // does this; Add did not). Invalid OOXML otherwise.
                if (colWidthArr[ci] <= 0)
                    throw new ArgumentException($"Invalid 'colwidths' value: '{part}'. Each column width must be a positive integer (in twips). Example: colwidths=3000,2000,5000");
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

        // Apply table-level properties from Add parameters.
        // explicitDirection: did the user pass direction=ltr|rtl|bidi=...?
        // If not, fall back to the surrounding section's bidi state below
        // (CONSISTENCY(rtl-cascade) — same intent as the paragraph cascade).
        bool explicitDirection = false;
        // BUG-DUMP-R36-2: band sizes are collected during the prop loop and
        // materialized once afterwards (see the mc:AlternateContent guard
        // below) — CT_TblPr has no slot for them, so they cannot be inserted
        // directly like the other tblPr children.
        int? rowBandSize = null, colBandSize = null;
        // Set of keys the switch below consumes. Used to mark a key as
        // accessed via ContainsKey only when a case actually matched, so
        // genuine typos still fall through to the tracker's UnusedKeys.
        var tblBareConsumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "align", "alignment", "width", "indent", "cellspacing", "layout",
            "padding", "padding.top", "padding.bottom", "padding.left", "padding.right",
            "style", "shd", "shading", "cellShading", "direction", "dir", "bidi",
            // CONSISTENCY(add-set-symmetry): mirror Set's tblPr-level cases.
            // BUG-DUMP-H89: `tblOverlap.val` is the Get readback key (Navigation.cs);
            // accept it as an alias for `overlap` so dump→batch round-trips.
            "overlap", "tblOverlap.val", "tblOverlap", "caption", "description",
            // BUG-DUMP-R36-2: band stripe widths.
            "rowbandsize", "colbandsize", "columnbandsize",
            // BUG-DUMP-R40-5: tblLook bitmask + decomposed facets. Without these
            // in the consumed set, the access-accounting never marks them and the
            // round-trip `tableLook=…` key surfaced as a spurious UNSUPPORTED
            // warning even though the switch below applies it.
            "tbllook", "firstrow", "lastrow", "firstcol", "firstcolumn",
            "lastcol", "lastcolumn", "bandrow", "bandedrows", "bandrows",
            "bandcol", "bandedcols", "bandcols", "nohband", "nohorizontalband",
            "novband", "noverticalband",
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
            // `data` (inline cell content), like rows/cols/colwidths, is fully
            // consumed earlier (TryGetValue("data") above) — skip it here so the
            // tblPr switch default doesn't falsely flag it as unsupported_property.
            if (tkl is "rows" or "cols" or "columns" or "colwidths" or "gridcols" or "skiptblw" or "skiptbllook" or "skipdefaultborders" or "data" || tkl.StartsWith("border")) continue;
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
                        // Fractional percentages are exact in OOXML (pct unit =
                        // 0.02%); mirror the cell-width branch's double parse.
                        var pct = (int)Math.Round(ParseHelpers.SafeParseDouble(tv.TrimEnd('%'), "width") * 50);
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
                    // BUG-DUMP-R34-TBLIND: honour a pct-typed indent ("2%") so it
                    // round-trips as <w:tblInd w:type="pct"> instead of collapsing
                    // to dxa twips (which shifts the whole table).
                    // CONSISTENCY(length-units): non-pct indent accepts pt/cm/in
                    // (and bare = twips) via SpacingConverter, matching tc padding
                    // and every other length slot. tblInd is signed (ST_SignedTwipsMeasure).
                    tblProps.TableIndentation = tv.TrimEnd().EndsWith("%", StringComparison.Ordinal)
                        ? new TableIndentation { Width = (int)Math.Round(ParseHelpers.SafeParseDouble(tv.TrimEnd().TrimEnd('%'), "indent") * 50), Type = TableWidthUnitValues.Pct }
                        : new TableIndentation { Width = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(tv), Type = TableWidthUnitValues.Dxa };
                    break;
                case "cellspacing":
                    tblProps.TableCellSpacing = new TableCellSpacing { Width = OfficeCli.Core.SpacingConverter.ParseWordSpacing(tv).ToString(), Type = TableWidthUnitValues.Dxa };
                    break;
                case "layout":
                    tblProps.TableLayout = new TableLayout
                    {
                        Type = tv.ToLowerInvariant() switch
                        {
                            "fixed"   => TableLayoutValues.Fixed,
                            "autofit" or "auto" => TableLayoutValues.Autofit,
                            _ => throw new ArgumentException($"Invalid 'layout' value: '{tv}'. Valid values: fixed, autofit."),
                        }
                    };
                    break;
                case "padding":
                {
                    // CONSISTENCY(tblpr-schema-order): tblCellMar must be inserted
                    // at rank 13 in CT_TblPrBase; raw AppendChild produced
                    // schema-invalid OOXML when other props (style→tblLook at rank
                    // 14) appeared earlier in argv. Same fix for padding.{top,
                    // bottom,left,right} below and Set padding path.
                    var paddingVal = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(tv);
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
                        var tv2 = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(tv);
                        if (tv2 < 0)
                            throw new ArgumentException($"Invalid 'padding.top' value: '{tv}'. Table cell margins must be non-negative.");
                        var cmt = EnsureTableCellMarginDefault(tblProps);
                        cmt.TopMargin = new TopMargin { Width = tv2.ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding.bottom":
                    {
                        var bv2 = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(tv);
                        if (bv2 < 0)
                            throw new ArgumentException($"Invalid 'padding.bottom' value: '{tv}'. Table cell margins must be non-negative.");
                        var cmb = EnsureTableCellMarginDefault(tblProps);
                        cmb.BottomMargin = new BottomMargin { Width = bv2.ToString(), Type = TableWidthUnitValues.Dxa };
                    }
                    break;
                case "padding.left":
                    {
                        var lv = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(tv);
                        if (lv < 0)
                            throw new ArgumentException($"Invalid 'padding.left' value: '{tv}'. Table cell margins must be non-negative.");
                        var cml = EnsureTableCellMarginDefault(tblProps);
                        cml.TableCellLeftMargin = new TableCellLeftMargin { Width = (short)Math.Min(lv, short.MaxValue), Type = TableWidthValues.Dxa };
                    }
                    break;
                case "padding.right":
                    {
                        var rv = OfficeCli.Core.SpacingConverter.ParseWordSpacingSigned(tv);
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
                    // BUG-DUMP-R40-5: only seed the default 04A0 when the caller
                    // did NOT supply an explicit tblLook (bare hex `tableLook=…`
                    // or a decomposed `firstRow`/`tblLook.firstRow`/… key). The
                    // dump→batch round-trip emits `tableLook=<source hex>`; force-
                    // seeding 04A0 here (firstRow+firstColumn) leaked first-column
                    // conditional formatting onto tables whose source tblLook had
                    // it off. The tblLook case below applies the explicit value
                    // (it runs after this case regardless of prop order, so the
                    // seed it builds would be the authoritative hex anyway — but
                    // suppressing the default keeps a no-tblLook source clean).
                    // BUG-DUMP-TBLLOOK-INJECT: skipTblLook=true (dump-replay of a
                    // source whose <w:tblPr> had no <w:tblLook>) also suppresses
                    // the default seed — otherwise 04A0 leaks the style's
                    // first-row/first-column conditional formatting onto every
                    // styled table. Mirrors skipTblW.
                    if (!TableHasExplicitTblLook(properties)
                        && !((properties.TryGetValue("skiptbllook", out var stl)
                              || properties.TryGetValue("skipTblLook", out stl)) && IsTruthy(stl)))
                    {
                        tblProps.RemoveAllChildren<TableLook>();
                        InsertTblPrChildInOrder(tblProps, new TableLook { Val = "04A0" });
                    }
                    break;
                case "shd" or "shading" or "cellshading":
                    {
                        // BUG-DUMP21-01: w:tblPr/w:shd table-level shading
                        // round-trip. Route through the shared ParseShadingValue
                        // so the full form is honored — FILL, VAL;FILL,
                        // VAL;FILL;COLOR, plus the themeFill=/themeFillTint=/
                        // themeFillShade= theme-linkage tail the dump emits. The
                        // old hand-rolled split dropped the theme tail, losing the
                        // theme reference on round-trip.
                        tblProps.Shading = ParseShadingValue(tv);
                    }
                    break;
                // BUG-DUMP-H89: `tbloverlap.val` is the Get readback key; accept it
                // (and bare `tbloverlap`) as aliases for `overlap` so the dump's
                // emitted key round-trips through batch instead of being ignored.
                case "overlap":
                case "tbloverlap.val":
                case "tbloverlap":
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
                    explicitDirection = true;
                    break;
                // Accessibility caption / description (<w:tblCaption>/<w:tblDescription>).
                // Mirrors the SetElementTable cases so dump→batch round-trips them;
                // without these the emitter's `add table` carried caption/description
                // but AddTable silently dropped them. CONSISTENCY(add-set-symmetry).
                case "caption":
                    tblProps.RemoveAllChildren<TableCaption>();
                    if (!string.IsNullOrEmpty(tv))
                        InsertTblPrChildInOrder(tblProps, new TableCaption { Val = tv });
                    break;
                case "description":
                    tblProps.RemoveAllChildren<TableDescription>();
                    if (!string.IsNullOrEmpty(tv))
                        InsertTblPrChildInOrder(tblProps, new TableDescription { Val = tv });
                    break;
                // BUG-DUMP-R36-2: tblStyleRowBandSize / tblStyleColBandSize —
                // band stripe width (rows/cols per band) for banded table styles.
                // Mirrors the Navigation readback; without these the emitter's
                // `add table` carried rowBandSize/colBandSize but AddTable dropped
                // them, flattening the visible striping. CONSISTENCY(add-set-symmetry).
                // Collected here, written after the loop (mc guard below).
                case "rowbandsize":
                    if (int.TryParse(tv, out var rbs)) rowBandSize = rbs;
                    break;
                case "colbandsize" or "columnbandsize":
                    if (int.TryParse(tv, out var cbs)) colBandSize = cbs;
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
                default:
                    // Bare key that didn't match any known table prop. Common
                    // typos previously silently failed (e.g. `columns=4`
                    // before the `cols` alias was added; `column`, `rowcount`,
                    // `border-all`, …). Surface via LastAddUnsupportedProps
                    // so the CLI emits a WARNING — the table is still built
                    // (best-effort) but the user sees that their key was
                    // ignored. Skip dynamic patterns: r{N}c{M} cell-content
                    // keys (consumed in the row-building loop below) and any
                    // dotted key (the second foreach further down routes
                    // those through TypedAttributeFallback and tracks
                    // unsupporteds itself).
                    if (tk.Contains('.')) break;
                    if (System.Text.RegularExpressions.Regex.IsMatch(
                            tkl, @"^r\d+c\d+$")) break;
                    LastAddUnsupportedProps.Add(tk);
                    break;
            }
        }

        // BUG-DUMP-R36-2: materialize tblStyleRowBandSize / tblStyleColBandSize.
        // The document-level CT_TblPr particle (SDK ground truth — see the
        // generated TableProperties class) does NOT admit these elements in ANY
        // position; they are only schema-valid inside a table STYLE's tblPr
        // (CT_TblPrStyle). Inserting them bare therefore always validates as
        // "invalid child element 'tblStyleRowBandSize'", no matter the order.
        // wml-aware consumers (Word and other editors — whose DOCX export writes
        // them document-side at the CT_TblPrBase rank-4/5 slot) DO honor the
        // elements, so we keep them document-side but wrap them in an
        // mc:AlternateContent guard with Requires="w": every wordprocessingml
        // consumer selects the Choice (the band sizes resolve in place, in
        // their canonical pre-tblW slot), while strict schema validators skip
        // MC content and stay green. Navigation's table readback unwraps the
        // guard, so Get/dump still surface rowBandSize/colBandSize and the
        // dump→batch round-trip regenerates this exact shape.
        if (rowBandSize.HasValue || colBandSize.HasValue)
        {
            var bandChoice = new AlternateContentChoice { Requires = "w" };
            if (rowBandSize is int rbv)
                bandChoice.Append(new TableStyleRowBandSize { Val = rbv });
            if (colBandSize is int cbv)
                bandChoice.Append(new TableStyleColumnBandSize { Val = cbv });
            var bandGuard = new AlternateContent(bandChoice, new AlternateContentFallback());
            // Place the guard where the band sizes themselves belong
            // (CT_TblPrBase rank 4/5: after tblStyle/tblpPr/tblOverlap/
            // bidiVisual, before tblW) so the MC-resolved document keeps
            // canonical order for order-sensitive consumers.
            OpenXmlElement? bandSuccessor = null;
            foreach (var child in tblProps.ChildElements)
            {
                if (child is TableStyle or TablePositionProperties or TableOverlap or BiDiVisual) continue;
                bandSuccessor = child;
                break;
            }
            if (bandSuccessor != null)
                bandSuccessor.InsertBeforeSelf(bandGuard);
            else
                tblProps.AppendChild(bandGuard);
        }

        // Auto-RTL: when the user didn't pin direction explicitly and the
        // surrounding section / doc-defaults are RTL (Arabic/Hebrew/… via
        // --locale or the OS culture snapshot), stamp <w:bidiVisual/> so
        // the table's column order matches the rest of the document.
        // Without this, every `add table` in an RTL doc rendered with
        // left-to-right column order — the agent had to remember to pass
        // --prop direction=rtl on every table, inconsistent with how
        // paragraphs inherit section bidi automatically.
        if (!explicitDirection && IsTableContextRtl(parent))
        {
            tblProps.RemoveAllChildren<BiDiVisual>();
            InsertTblPrChildInOrder(tblProps, new BiDiVisual());
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
            // CONSISTENCY(add-set-symmetry): tblp.horzAnchor / tblp.vertAnchor
            // are enum-valued (ST_HAnchor / ST_VAnchor). The Set side curated
            // them at Set.Element.cs:2659 and throws on invalid input; the
            // generic TypedAttributeFallback below would write raw strings
            // like "column" verbatim. Validate here so Add matches Set.
            var kLowAdd = key.ToLowerInvariant();
            if (kLowAdd is "tblp.horzanchor" or "tblp.horizontalanchor"
                or "tblppr.horzanchor" or "tblppr.horizontalanchor")
            {
                // CONSISTENCY(tblpr-schema-order): tblpPr is rank 1 in CT_TblPr.
                // Appending it (as the old code did) lands it last, after
                // tblLook → schema-invalid floating table. Reuse the Set-side
                // helper so Add inserts it in order, same as set tblp.*.
                var tpp = EnsureTablePositionProperties(tblProps);
                tpp.HorizontalAnchor = value.ToLowerInvariant() switch
                {
                    "margin" => HorizontalAnchorValues.Margin,
                    "page" => HorizontalAnchorValues.Page,
                    "text" => HorizontalAnchorValues.Text,
                    _ => throw new ArgumentException($"Invalid 'tblp.horzAnchor' value: '{value}'. Valid: margin, page, text."),
                };
                continue;
            }
            if (kLowAdd is "tblp.vertanchor" or "tblp.verticalanchor"
                or "tblppr.vertanchor" or "tblppr.verticalanchor")
            {
                // CONSISTENCY(tblpr-schema-order): tblpPr is rank 1 in CT_TblPr.
                // Appending it (as the old code did) lands it last, after
                // tblLook → schema-invalid floating table. Reuse the Set-side
                // helper so Add inserts it in order, same as set tblp.*.
                var tpp = EnsureTablePositionProperties(tblProps);
                tpp.VerticalAnchor = value.ToLowerInvariant() switch
                {
                    "margin" => VerticalAnchorValues.Margin,
                    "page" => VerticalAnchorValues.Page,
                    "text" => VerticalAnchorValues.Text,
                    _ => throw new ArgumentException($"Invalid 'tblp.vertAnchor' value: '{value}'. Valid: margin, page, text."),
                };
                continue;
            }
            // tblpPr.*FromText / position.*FromText — ST_TwipsMeasure with the
            // short-range SDK property. TypedAttributeFallback would write the
            // overflowed value verbatim (32768 cast → -32768); validate up-front.
            if (kLowAdd is "tblppr.leftfromtext" or "tblppr.rightfromtext"
                or "tblppr.topfromtext" or "tblppr.bottomfromtext"
                or "position.leftfromtext" or "position.rightfromtext"
                or "position.topfromtext" or "position.bottomfromtext")
            {
                var ftTwips = ParseTwips(value);
                if (ftTwips > 32767)
                    throw new ArgumentException($"Invalid '{key}' value: '{value}'. Must be 0..32767 twips (OOXML ST_TwipsMeasure short range).");
                // CONSISTENCY(tblpr-schema-order): tblpPr is rank 1 in CT_TblPr.
                // Appending it (as the old code did) lands it last, after
                // tblLook → schema-invalid floating table. Reuse the Set-side
                // helper so Add inserts it in order, same as set tblp.*.
                var tpp = EnsureTablePositionProperties(tblProps);
                if (kLowAdd.EndsWith("leftfromtext")) tpp.LeftFromText = (short)ftTwips;
                else if (kLowAdd.EndsWith("rightfromtext")) tpp.RightFromText = (short)ftTwips;
                else if (kLowAdd.EndsWith("topfromtext")) tpp.TopFromText = (short)ftTwips;
                else tpp.BottomFromText = (short)ftTwips;
                continue;
            }
            if (Core.TypedAttributeFallback.TrySet(tblProps, key, value)) continue;
            LastAddUnsupportedProps.Add(key);
        }

        if (index.HasValue)
            InsertAtPosition(parent, table, index);
        else
            AppendToParent(parent, table);
        // OOXML §17.4.66: every <w:tc> must end with <w:p>. Word rejects
        // a cell ending with <w:tbl> as the last child. Nested table Add
        // appends the inner <w:tbl> as the cell's last child — restore the
        // required trailing empty paragraph here so Word can open the doc.
        if (parent is TableCell && parent.LastChild is Table)
            parent.AppendChild(new Paragraph());
        _lastAddedTable = table;
        var tbls = parent.Elements<Table>().ToList();
        var idx = tbls.FindIndex(t => ReferenceEquals(t, table));
        return $"{parentPath}/tbl[{(idx >= 0 ? idx + 1 : tbls.Count)}]";
    }

    /// <summary>True when properties carries any trackChange.* sub-key.</summary>
    private static bool HasRowTrackChangeProps(Dictionary<string, string> properties)
    {
        foreach (var k in properties.Keys)
        {
            if (k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
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
            // BUG-DUMP-R25-1: bare `height` = AUTO row-sizing (no @w:hRule).
            // Explicit rule via height.atleast / height.exact below.
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeight) });
        }
        if (properties.TryGetValue("height.atleast", out var rowHeightAtLeast))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            newRowProps.GetFirstChild<TableRowHeight>()?.Remove();
            newRowProps.AppendChild(new TableRowHeight { Val = ParseTwips(rowHeightAtLeast), HeightType = HeightRuleValues.AtLeast });
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
        // BUG-DUMP-R37-3: <w:hidden/> — row not displayed/printed.
        // CONSISTENCY(add-set-symmetry): mirrors SetElementTableRow `hidden`.
        if (properties.TryGetValue("hidden", out var rowHiddenVal) && IsTruthy(rowHiddenVal))
        {
            newRowProps ??= newRow.AppendChild(new TableRowProperties());
            if (newRowProps.GetFirstChild<Hidden>() == null)
                newRowProps.AppendChild(new Hidden());
        }
        // BUG-DUMP-R24-1: row-level <w:jc> (whole-row alignment). jc ranks
        // after cantSplit/trHeight/tblHeader in CT_TrPr, so AppendChild keeps
        // schema order. CONSISTENCY(add-set-symmetry): mirrors SetElementTableRow.
        if (properties.TryGetValue("rowAlign", out var rowAlignVal)
            || properties.TryGetValue("rowalign", out rowAlignVal))
        {
            if (!string.IsNullOrEmpty(rowAlignVal))
            {
                newRowProps ??= newRow.AppendChild(new TableRowProperties());
                var rav = rowAlignVal.ToLowerInvariant() switch
                {
                    "left" => TableRowAlignmentValues.Left,
                    "center" => TableRowAlignmentValues.Center,
                    "right" => TableRowAlignmentValues.Right,
                    _ => throw new ArgumentException(
                        $"Invalid rowAlign '{rowAlignVal}': must be left, center, or right")
                };
                newRowProps.AppendChild(new TableJustification { Val = rav });
            }
        }

        // BUG-DUMP-R42-2: leading/trailing grid-column skips + their preferred
        // widths (ragged/indented table edge). CONSISTENCY(add-set-symmetry):
        // mirrors SetElementTableRow's gridBefore/wBefore/gridAfter/wAfter cases.
        // Use InsertTrPrChildInOrder so each lands at its CT_TrPr rank regardless
        // of which other trPr children already exist.
        if (properties.TryGetValue("gridBefore", out var gridBeforeVal)
            || properties.TryGetValue("gridbefore", out gridBeforeVal))
        {
            if (!string.IsNullOrEmpty(gridBeforeVal))
            {
                newRowProps ??= newRow.AppendChild(new TableRowProperties());
                InsertTrPrChildInOrder(newRowProps, new GridBefore { Val = ParseHelpers.SafeParseInt(gridBeforeVal, "gridBefore") });
            }
        }
        if (properties.TryGetValue("wBefore", out var wBeforeVal)
            || properties.TryGetValue("wbefore", out wBeforeVal))
        {
            if (!string.IsNullOrEmpty(wBeforeVal))
            {
                newRowProps ??= newRow.AppendChild(new TableRowProperties());
                InsertTrPrChildInOrder(newRowProps, BuildRowWidth<WidthBeforeTableRow>(wBeforeVal, "wBefore"));
            }
        }
        if (properties.TryGetValue("gridAfter", out var gridAfterVal)
            || properties.TryGetValue("gridafter", out gridAfterVal))
        {
            if (!string.IsNullOrEmpty(gridAfterVal))
            {
                newRowProps ??= newRow.AppendChild(new TableRowProperties());
                InsertTrPrChildInOrder(newRowProps, new GridAfter { Val = ParseHelpers.SafeParseInt(gridAfterVal, "gridAfter") });
            }
        }
        if (properties.TryGetValue("wAfter", out var wAfterVal)
            || properties.TryGetValue("wafter", out wAfterVal))
        {
            if (!string.IsNullOrEmpty(wAfterVal))
            {
                newRowProps ??= newRow.AppendChild(new TableRowProperties());
                InsertTrPrChildInOrder(newRowProps, BuildRowWidth<WidthAfterTableRow>(wAfterVal, "wAfter"));
            }
        }

        // BUG-DUMP-R24-4: per-row <w:tblPrEx> (table property exceptions),
        // verbatim element captured by Navigation.ReadRowProps. Insert as the
        // row's FIRST child (CT_Row order: tblPrEx?, trPr?, cells).
        // CONSISTENCY(add-set-symmetry): mirrors SetElementTableRow.
        if (properties.TryGetValue("tblPrEx", out var rowTblPrEx)
            && !string.IsNullOrWhiteSpace(rowTblPrEx))
        {
            newRow.PrependChild(new TablePropertyExceptions(rowTblPrEx));
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

        // High-level row-insertion revision: any trackChange.* sub-key (author/
        // date/id) marks the newly-added row as inserted by placing a bare
        // <w:ins/> marker inside <w:trPr>. Mirrors the `add run + trackChange.*
        // → <w:ins> wrapper` pattern (Phase 1) but uses OOXML's row-level
        // marker-in-Pr form rather than a wrapper.
        //
        // KNOWN LIMITATION (A-route boundary): we do NOT cascade by wrapping
        // every cell's inner run in <w:ins> too. Word UI requires the cascade
        // to display the row as "newly inserted" visually; without it the row
        // appears unmarked in Word even though accept-all / reject-all still
        // identify it as a revision. Authoring a fully-rendered row insertion
        // is out of CLI scope — use Word directly. The trPr/ins marker remains
        // useful for: programmatic accept/reject, query revision, and
        // round-trip preservation.
        if (HasRowTrackChangeProps(properties))
        {
            string? rowTcAuthor = null, rowTcDate = null, rowTcId = null;
            properties.TryGetValue("revision.author", out rowTcAuthor);
            properties.TryGetValue("revision.date", out rowTcDate);
            properties.TryGetValue("revision.id", out rowTcId);

            newRowProps ??= newRow.PrependChild(new TableRowProperties());
            var marker = new Inserted
            {
                Author = string.IsNullOrEmpty(rowTcAuthor) ? "OfficeCLI" : rowTcAuthor,
                Date = !string.IsNullOrEmpty(rowTcDate) && DateTime.TryParse(rowTcDate, out var rowD)
                    ? rowD : DateTime.UtcNow,
                Id = !string.IsNullOrEmpty(rowTcId) ? rowTcId : GenerateRevisionId(),
            };
            newRowProps.AppendChild(marker);

            // Remove trackChange.* from unsupported list (consumed).
            LastAddUnsupportedProps.RemoveAll(k =>
                k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase));
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

        var rowIdx = PathIndex.FromArrayIndex(targetTable.Elements<TableRow>().ToList().IndexOf(newRow));
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

        // Reject if any row's merge makes the insertion slot un-addressable by
        // position. BUG-COLOP-GRIDIDX: the old check inspected cells[insertIdx]
        // by ORDINAL, but a preceding gridSpan shifts the ordinal off the grid
        // slot — so a straddling merge before insertIdx was missed AND the
        // insertion at cells[insertIdx] below would push the wrong cell right.
        // The slot-aware guard rejects a merged target slot OR a preceding
        // horizontal span (ordinal ≠ slot). Appending (insertIdx == count) has
        // no occupied slot to check and is always safe.
        if (insertIdx < existingGridCols.Count)
            GuardColumnSlotAddressable(targetTable, insertIdx, "insert column at index " + insertIdx + " of " + parentPath + ";");

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

        // Phase 6: virtual-column insertion revision. trackChange.* sub-keys
        // mark every newly-inserted cell with <w:tcPr><w:cellIns/></w:tcPr>.
        // The N cells produced (one per existing row) is the column op's
        // natural output — NOT a cascade into pre-existing content. All
        // cellIns markers share author/date but get distinct auto-allocated
        // ids (no explicit trackChange.id support across N cells — it would
        // be ambiguous).
        bool colHasTc = HasRowTrackChangeProps(properties);
        string colTcAuthor = "OfficeCLI";
        DateTime colTcDate = DateTime.UtcNow;
        if (colHasTc)
        {
            properties.TryGetValue("revision.author", out var aRaw);
            if (!string.IsNullOrEmpty(aRaw)) colTcAuthor = aRaw;
            properties.TryGetValue("revision.date", out var dRaw);
            if (!string.IsNullOrEmpty(dRaw) && DateTime.TryParse(dRaw, out var parsed))
                colTcDate = parsed;
        }

        foreach (var row in targetTable.Elements<TableRow>())
        {
            var newPara = new Paragraph();
            AssignParaId(newPara);
            if (!string.IsNullOrEmpty(cellText))
                newPara.AppendChild(new Run(new Text(cellText) { Space = SpaceProcessingModeValues.Preserve }));
            var newCell = new TableCell(newPara);

            if (colHasTc)
            {
                var tcPr = newCell.PrependChild(new TableCellProperties());
                tcPr.AppendChild(new CellInsertion
                {
                    Author = colTcAuthor,
                    Date = colTcDate,
                    Id = GenerateRevisionId(),
                });
            }

            var cells = row.Elements<TableCell>().ToList();
            if (insertIdx < cells.Count)
                row.InsertBefore(newCell, cells[insertIdx]);
            else
                row.AppendChild(newCell);
        }

        if (colHasTc)
            LastAddUnsupportedProps.RemoveAll(k =>
                k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase));

        var newColIdx = PathIndex.FromArrayIndex(grid.Elements<GridColumn>().ToList().IndexOf(newGridCol));
        return $"{parentPath}/col[{newColIdx}]";
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
            // BUG-DUMP-R42-6: round-trip nil/auto cell-width types (mirror the
            // Set.Element.cs width branch). nil = "no preferred width"; without
            // this an AddCell with width=nil/auto coerced to a bogus dxa value.
            else if (string.Equals(cellWidth, "nil", StringComparison.OrdinalIgnoreCase))
            {
                tcw = new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Nil };
            }
            else if (string.Equals(cellWidth, "auto", StringComparison.OrdinalIgnoreCase))
            {
                tcw = new TableCellWidth { Width = "0", Type = TableWidthUnitValues.Auto };
            }
            else
            {
                // Strip a trailing "dxa" suffix (the form Get emits) so the bare
                // twips path works for dump→batch round-trips.
                var dxaVal = cellWidth.EndsWith("dxa", StringComparison.OrdinalIgnoreCase)
                    ? cellWidth[..^3] : cellWidth;
                tcw = new TableCellWidth { Width = dxaVal, Type = TableWidthUnitValues.Dxa };
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
            if (keyLower is "fill" or "shd" or "shading" or "cellshading")
            {
                // foreach uses the base Dictionary enumerator (bypasses the
                // TrackingPropertyDictionary override) — register the key so the
                // consumed shading prop isn't a false unsupported_property.
                properties.ContainsKey(key);
                var tcPrFill = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                // Route through the shared ParseShadingValue so a cell's
                // themeFill=/themeFillTint=/themeFillShade= theme-linkage tail
                // round-trips; the old hand-rolled split dropped it.
                tcPrFill.Shading = ParseShadingValue(value);
            }
        }

        // BUG-WB-ADD-CELL-RUN: bare run-level props (bold/italic/color/size/
        // font/underline/strike/highlight) on AddCell were silently dropped and
        // falsely warned, while Set on a cell applies them (SetElementTableCell
        // run-prop branch). CONSISTENCY(add-set-symmetry): fan out across the
        // cell's runs, or — when the cell has no text run — hoist onto the
        // paragraph-mark rPr so a future run inherits, mirroring AddParagraph's
        // no-text hoist and Set's !hasRuns branch.
        foreach (var (rkey, rvalue) in properties)
        {
            var rkl = rkey.ToLowerInvariant();
            if (rkl is not ("font" or "size" or "fontsize" or "bold" or "italic"
                or "color" or "highlight" or "underline" or "underline.color"
                or "underlinecolor" or "strike"))
                continue;
            // Register the key with the tracking comparer (foreach bypasses it).
            properties.ContainsKey(rkey);
            bool cellHasRuns = false;
            foreach (var existingRun in cellParagraph.Elements<Run>())
            {
                cellHasRuns = true;
                ApplyRunFormatting(EnsureRunProperties(existingRun), rkey, rvalue);
            }
            var rPProps = cellParagraph.ParagraphProperties ?? cellParagraph.PrependChild(new ParagraphProperties());
            var rMarkRPr = rPProps.ParagraphMarkRunProperties ?? rPProps.AppendChild(new ParagraphMarkRunProperties());
            ApplyRunFormatting(rMarkRPr, rkey, rvalue);
            if (!cellHasRuns && rMarkRPr.ChildElements.Count == 0) rMarkRPr.Remove();
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
        // BUG-DUMP-CELLTAIL: <w:tcFitText/> toggle. CONSISTENCY(add-set-symmetry):
        // mirror Set's tcFitText case. Appended before hideMark so the resulting
        // child order follows CT_TcPr (tcFitText → vAlign → hideMark).
        if (properties.TryGetValue("tcFitText", out var tcFitTextVal)
            || properties.TryGetValue("tcfittext", out tcFitTextVal))
        {
            if (IsTruthy(tcFitTextVal))
            {
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                if (tcPr.GetFirstChild<TableCellFitText>() == null)
                    tcPr.AppendChild(new TableCellFitText());
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

        // CONSISTENCY(add-set-symmetry): valign, align, textdirection, cnfstyle
        // were Set-only in SetElementTableCell; mirror them here so AddCell
        // accepts the same vocabulary without false UNSUPPORTED warnings.
        if (properties.TryGetValue("valign", out var valignAddVal))
        {
            var tcPr = newCell.GetFirstChild<TableCellProperties>()
                ?? newCell.PrependChild(new TableCellProperties());
            tcPr.TableCellVerticalAlignment = new TableCellVerticalAlignment
            {
                Val = valignAddVal.ToLowerInvariant() switch
                {
                    "top"    => TableVerticalAlignmentValues.Top,
                    "center" => TableVerticalAlignmentValues.Center,
                    "bottom" => TableVerticalAlignmentValues.Bottom,
                    _ => throw new ArgumentException($"Invalid valign value: '{valignAddVal}'. Valid values: top, center, bottom.")
                }
            };
        }
        if (properties.TryGetValue("align", out var cellAlignAddVal)
            || properties.TryGetValue("alignment", out cellAlignAddVal)
            || properties.TryGetValue("halign", out cellAlignAddVal))
        {
            var alignVal = ParseJustification(cellAlignAddVal);
            foreach (var cellAlignPara in newCell.Elements<Paragraph>())
            {
                var cpProps = cellAlignPara.ParagraphProperties ?? cellAlignPara.PrependChild(new ParagraphProperties());
                cpProps.Justification = new Justification { Val = alignVal };
            }
        }
        if (properties.TryGetValue("textdirection", out var textDirAddVal)
            || properties.TryGetValue("textdir", out textDirAddVal))
        {
            var tcPr = newCell.GetFirstChild<TableCellProperties>()
                ?? newCell.PrependChild(new TableCellProperties());
            // Reuse the shared section parser so the cell path also accepts the
            // canonical OOXML InnerText forms the dump emits (tbLrV / tbRlV /
            // lrTbV — rotated vertical variants). CONSISTENCY(add-set-symmetry)
            // with the cell Set textDirection path.
            tcPr.TextDirection = new TextDirection { Val = ParseSectionTextDirection(textDirAddVal) };
        }
        if (properties.TryGetValue("cnfstyle", out var cnfAddVal))
        {
            if (!string.IsNullOrEmpty(cnfAddVal))
            {
                var cnfVal = ValidateCnfStyleBitmask(cnfAddVal);
                var tcPr = newCell.GetFirstChild<TableCellProperties>()
                    ?? newCell.PrependChild(new TableCellProperties());
                var cnf = tcPr.GetFirstChild<ConditionalFormatStyle>();
                if (cnf == null)
                    tcPr.PrependChild(new ConditionalFormatStyle { Val = cnfVal });
                else
                    cnf.Val = cnfVal;
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
                if (!ApplyCellBorders(lazyTcPr, key, value))
                {
                    LastAddUnsupportedProps.Add(key);
                    continue;
                }
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

        // Phase 6: cell-insertion revision. Any trackChange.* sub-key marks
        // the newly-added cell with <w:tcPr><w:cellIns/></w:tcPr>. Mirrors
        // `add row + trackChange.author` (which puts <w:ins/> in trPr).
        if (HasRowTrackChangeProps(properties))  // reuse the row helper
        {
            string? cTcAuthor = null, cTcDate = null, cTcId = null;
            properties.TryGetValue("revision.author", out cTcAuthor);
            properties.TryGetValue("revision.date", out cTcDate);
            properties.TryGetValue("revision.id", out cTcId);

            var tcPr = newCell.GetFirstChild<TableCellProperties>()
                      ?? newCell.PrependChild(new TableCellProperties());
            tcPr.AppendChild(new CellInsertion
            {
                Author = string.IsNullOrEmpty(cTcAuthor) ? "OfficeCLI" : cTcAuthor,
                Date = !string.IsNullOrEmpty(cTcDate) && DateTime.TryParse(cTcDate, out var cellD)
                    ? cellD : DateTime.UtcNow,
                Id = !string.IsNullOrEmpty(cTcId) ? cTcId : GenerateRevisionId(),
            });
            LastAddUnsupportedProps.RemoveAll(k =>
                k.StartsWith("revision.", StringComparison.OrdinalIgnoreCase));
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

        var cellIdx = PathIndex.FromArrayIndex(targetRow.Elements<TableCell>().ToList().IndexOf(newCell));
        return $"{parentPath}/tc[{cellIdx}]";
    }
}
