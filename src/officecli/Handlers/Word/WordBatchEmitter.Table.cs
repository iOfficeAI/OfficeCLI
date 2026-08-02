// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    // BUG-DUMP-R27-6: enumerate a table cell's DIRECT block children (top-level
    // <w:p>, <w:tbl>, <w:sdt>, <w:customXml>) in document order from its raw
    // XML. A depth-tracked scan keeps paragraphs/tables nested inside a child
    // (a nested table's cells, a customXml's inner blocks) from being counted
    // as cell-level blocks. The first element open is the <w:tc> wrapper itself
    // (depth 0); its direct children are at depth 1. <w:tcPr> is skipped (it is
    // cell properties, not block content). Mirrors EnumerateNoteDirectChildren.
    // Needed because cellNode.Children (Navigation) surfaces only p/tbl/sdt and
    // OMITS customXml, so a cell whose content is wrapped in a block customXml
    // had its inner text silently dropped on dump (the body path flattens +
    // warns; the cell path did neither).
    private static List<string> EnumerateCellDirectChildren(string? cellXml)
    {
        var result = new List<string>();
        if (string.IsNullOrEmpty(cellXml)) return result;
        int depth = -1; // becomes 0 when the <w:tc> wrapper opens
        bool seenWrapper = false;
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(
                     cellXml!, @"<(/?)w:([A-Za-z]+)\b[^>]*?(/?)>"))
        {
            var closing = m.Groups[1].Value == "/";
            var name = m.Groups[2].Value;
            var selfClose = m.Groups[3].Value == "/";
            if (!seenWrapper)
            {
                if (!closing) { seenWrapper = true; depth = 0; }
                continue;
            }
            if (closing) { depth--; continue; }
            if (depth == 0 && (name == "p" || name == "tbl" || name == "sdt" || name == "customXml"))
                result.Add(name);
            if (!selfClose) depth++;
        }
        return result;
    }

    // BUG-DUMP-R27-6: return the Nth (1-based) cellNode child matching either of
    // two accepted Type spellings ("p"/"paragraph", "table", "sdt"), so the
    // document-ordered customXml plan can recover the corresponding navigation
    // node by ordinal. Returns null when out of range (defensive).
    private static DocumentNode? NthChildOfType(List<DocumentNode> children,
                                                string t1, string t2, int ordinal)
    {
        int seen = 0;
        foreach (var ch in children)
        {
            if (ch.Type == t1 || ch.Type == t2)
            {
                seen++;
                if (seen == ordinal) return ch;
            }
        }
        return null;
    }

    // BUG-DUMP-H84: true when the table has a ROW-LEVEL <w:sdt> — a content
    // control that is a DIRECT child of <w:tbl> (sibling of <w:tr>), wrapping one
    // or more rows. The canonical case is a <w15:repeatingSection> form template
    // (often with a <w15:repeatingSectionItem> SDT per row). EmitTable enumerates
    // bare <w:tr> children and has no carrier for such a wrapper, so it would be
    // silently flattened to a static table. Depth-scan mirrors
    // EnumerateCellDirectChildren: the <w:tbl> wrapper is depth 0, its direct
    // children at depth 0; a <w:sdt> seen at depth 0 is row-level (a cell-level
    // SDT sits inside <w:tc>, depth > 0, and must NOT trigger this).
    private static bool TableHasRowLevelSdt(string? tblXml)
    {
        if (string.IsNullOrEmpty(tblXml)
            || !tblXml!.Contains("<w:sdt", StringComparison.Ordinal))
            return false;
        int depth = -1;
        bool seenWrapper = false;
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(
                     tblXml!, @"<(/?)w:([A-Za-z]+)\b[^>]*?(/?)>"))
        {
            var closing = m.Groups[1].Value == "/";
            var name = m.Groups[2].Value;
            var selfClose = m.Groups[3].Value == "/";
            if (!seenWrapper)
            {
                if (!closing) { seenWrapper = true; depth = 0; }
                continue;
            }
            if (closing) { depth--; continue; }
            if (depth == 0 && name == "sdt") return true;
            if (!selfClose) depth++;
        }
        return false;
    }

    private static void EmitTable(WordHandler word, string sourcePath, int targetIndex,
                                  List<BatchItem> items, BodyEmitContext? ctx = null,
                                  string? parentTablePath = null,
                                  string containerPath = "/body",
                                  int depth = 0)
    {
        // CONSISTENCY(dos-hardening): nested-table emission recurses with no
        // structural bound; a crafted deeply-nested table would otherwise hang
        // (or overflow the stack) during `dump`. See DocumentLimits.
        OfficeCli.Core.DocumentLimits.EnsureDepth(depth);

        // BUG-R11A(BUG1): bump the document-order table ordinal BEFORE the
        // empty-table early-return so the count never desyncs from the
        // `(//w:tbl)[N]` selectors used by cell-SDT raw-sets. EmitTable recurses
        // in DFS document order, so this matches the target's //w:tbl indexing.
        int tableOrdinal = 0;
        if (ctx != null) tableOrdinal = ++ctx.TableOrdinalBox[0];

        var tableNode = word.Get(sourcePath);

        // BUG-DUMP-H84: a table whose rows are wrapped by a row-level <w:sdt>
        // (a <w15:repeatingSection> form template / per-row repeatingSectionItem)
        // can't round-trip through the typed `add table` path below — EmitTable
        // enumerates bare <w:tr> children and has no carrier for the wrapping
        // control, so the repeating-section structure was silently flattened to a
        // static table (markers 1→0, no warning). Route the whole <w:tbl> verbatim
        // via raw-set, mirroring the rich block-SDT path (EmitSdt). Restricted to
        // body tables with no external relationship (the common form-template
        // shape — verbatim injection can't recreate dangling rels); other cases
        // warn and fall through (rows + content survive as a static table, no
        // longer a silent loss).
        var tblRawXml = word.RawElementXml(sourcePath);
        if (!string.IsNullOrEmpty(tblRawXml) && TableHasRowLevelSdt(tblRawXml))
        {
            if (containerPath == "/body" && !HasExternalRelRef(tblRawXml!))
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = "//w:body/w:sectPr",
                    Action = "insertbefore",
                    Xml = tblRawXml!
                });
                // CONSISTENCY(tbl-ordinal): the verbatim table (and any nested
                // <w:tbl> in its cells) ships WITHOUT routing through EmitTable, so
                // EmitTable's `++TableOrdinalBox` (line above) counted only the
                // outer table. Bump by the shipped XML's remaining table count so
                // later `(//w:tbl)[N]` selectors stay in lockstep with replay.
                // Mirrors the EmitSdt verbatim / textbox carrier adjustment.
                if (ctx != null)
                {
                    int shipped = System.Text.RegularExpressions.Regex
                        .Matches(tblRawXml!, "<w:tbl[ >]").Count;
                    if (shipped > 1) ctx.TableOrdinalBox[0] += shipped - 1;
                }
                return;
            }
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "table.rowSdt",
                Path: sourcePath,
                Reason: "table rows wrapped by a content control (e.g. repeatingSection) — the wrapping control is dropped on dump→batch; the rows and their content are preserved as a static table"));
            // fall through to the typed table emit (content survives)
        }

        var rows = (tableNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "row")
            .ToList();
        if (rows.Count == 0) return;

        // Column count must cover the widest row including colspan effects.
        // Format["cols"] reflects gridCol; per-row effective width is
        // sum(colspan or 1) over each cell. Take the max so a first row
        // with merged cells (visible cell count < grid width) doesn't
        // truncate the table shape and break later `set tc[N]` rows.
        var rowEffectiveWidths = new List<int>(rows.Count);
        var rowCellNodes = new List<List<DocumentNode>>(rows.Count);
        var rowNodes = new List<DocumentNode>(rows.Count);
        foreach (var rowChild in rows)
        {
            var rowNode = word.Get(rowChild.Path);
            rowNodes.Add(rowNode);
            var cells = (rowNode.Children ?? new List<DocumentNode>())
                .Where(c => c.Type == "cell")
                .ToList();
            rowCellNodes.Add(cells);
            int width = 0;
            foreach (var cell in cells)
            {
                int span = 1;
                if (cell.Format.TryGetValue("colspan", out var sp) &&
                    int.TryParse(sp?.ToString(), out var n) && n > 0)
                {
                    span = n;
                }
                width += span;
            }
            rowEffectiveWidths.Add(width);
        }
        int colsFromRows = rowEffectiveWidths.Count > 0 ? rowEffectiveWidths.Max() : 0;
        int colsFromGrid = 0;
        if (tableNode.Format.TryGetValue("cols", out var gridColObj) &&
            int.TryParse(gridColObj?.ToString(), out var gridCols))
        {
            colsFromGrid = gridCols;
        }
        // Format["cols"] back-fills from first-row cell count when source has
        // no <w:tblGrid> at all, so it can't tell us "source had zero gridCol".
        // _gridCols is the unbiased count (Navigation emits 0 when TableGrid
        // is missing or empty). EmitTable uses this to drive the gridCols=0
        // opt-out on the dumped `add table`.
        int actualGridCols = colsFromGrid;
        if (tableNode.Format.TryGetValue("_gridCols", out var actualGridObj) &&
            int.TryParse(actualGridObj?.ToString(), out var ag))
        {
            actualGridCols = ag;
        }
        int cols = Math.Max(colsFromGrid, colsFromRows);
        if (cols == 0) return;

        var tableProps = FilterEmittableProps(tableNode.Format);
        // Strip the revision-marker surface keys so they don't ride on
        // `add table` — they're consumed by a follow-up EmitTrackChangeMarker
        // call below. Without this, AddTable's schema fallback would create
        // a phantom <w:tblPrChange> on the new table, and the follow-up
        // `set trackChange.author=...` would then trip the
        // "element already has a pending tblPrChange" guard.
        tableProps.Remove("tblPrChange.author");
        tableProps.Remove("tblPrChange.date");
        // BUG-R24-TBLLOOK: the table node surfaces BOTH the authoritative hex
        // tblLook bitmask AND the decomposed boolean facets (firstRow / lastRow /
        // … — emitted only for the facets that are ON, see Navigation BUG-R3-01).
        // Forwarding both to AddTable produces a MIXED
        // <w:tblLook w:val="04A0" w:firstRow="true" w:firstColumn="true" …/> that
        // lists only the enabled facets and omits the disabled ones. Word DEFAULTS
        // an omitted tblLook facet to ON, so a table whose source explicitly
        // disabled lastRow (w:lastRow="0") had the last-row conditional style
        // (e.g. a GridTable lastRow bold) wrongly applied on rebuild — the last
        // row rendered bold and reflowed. The hex val encodes every facet
        // authoritatively and Word reads it correctly on its own, so drop the
        // redundant decomposed keys and let the bare val drive AddTable.
        if (tableProps.ContainsKey("tblLook"))
        {
            tableProps.Remove("firstRow");
            tableProps.Remove("lastRow");
            tableProps.Remove("firstCol");
            tableProps.Remove("lastCol");
            tableProps.Remove("bandedRows");
            tableProps.Remove("bandedCols");
        }
        else
        {
            // BUG-DUMP-TBLLOOK-INJECT: source <w:tblPr> had no <w:tblLook>.
            // AddTable's style-case seeds a default 04A0 (firstRow+firstColumn)
            // so interactive `add table style=…` applies built-in banding — but
            // on replay that injects first-row/first-column conditional
            // formatting onto a table whose source never had a tblLook, shifting
            // every styled table's first row/column. Signal AddTable to leave
            // tblLook absent. Mirrors the gridCols=0 / skipTblW opt-out flags.
            tableProps["skipTblLook"] = "true";
        }
        tableProps["rows"] = rows.Count.ToString();
        tableProps["cols"] = cols.ToString();
        // Source had no <w:tblGrid> or an empty one — cells (if any) carry
        // their own tcW, or the table is auto-fit. Without an explicit
        // `gridCols=0`, AddTable would seed `cols` default GridColumn entries
        // which ReadCellProps then back-fills as per-cell widths on the next
        // dump, producing N×M extra `set tc width=…` rows the source never
        // had (test.docx tbl[1]). Signal AddTable to leave tblGrid empty.
        if (actualGridCols == 0)
            tableProps["gridCols"] = "0";
        // Source had no <w:tblW> — surface a `skipTblW=true` user-facing
        // flag (mirrors `gridCols=0`). AddTable's default-tblW stamp
        // path defers to this when set, so replay won't grow a phantom
        // <w:tblW>. Skip when source had any explicit width (auto / dxa /
        // pct) — those round-trip through the existing `width=` key.
        bool sourceHadNoTblW = tableNode.Format.TryGetValue("_noTblW", out var noTblW)
            && noTblW is bool b && b;
        // BUG-DUMP-AUTOFITW: a table is autofit unless layout is explicitly
        // "fixed" (OOXML default, and what ReadTableProps emits when the
        // <w:tblLayout> element is absent). EmitTable passes this to
        // ExtractCellOnlyProps so tcW-less cells in an autofit table keep
        // their tcW-less state (no fabricated width) — only fixed-layout
        // tables and cells with real source tcW emit a width.
        bool tableIsAutofit = !(tableNode.Format.TryGetValue("layout", out var layoutObj)
            && layoutObj is string layoutStr
            && layoutStr.Equals("fixed", StringComparison.OrdinalIgnoreCase));
        if (sourceHadNoTblW && !tableProps.ContainsKey("width"))
            tableProps["skipTblW"] = "true";
        // Drop the internal-only markers from emitted props (BatchItem.Props
        // never carries them; only Navigation→EmitTable consumes them).
        tableProps.Remove("_gridCols");
        tableProps.Remove("_noTblW");
        // BUG-BORDER-PARTIAL: AddTable seeds all 6 default borders and overlays user
        // props on top, so a partial border spec (e.g. only border.top +
        // border.bottom for a banner-line table) replays as 6 single-borders.
        // If the source table emits only a subset of the 6 sides, prepend an
        // explicit `border=none` wipe so the visible result round-trips.
        // CONSISTENCY(border-default-overlay).
        //
        // The same fix applies to the zero-sides case: source tables with no
        // <w:tblBorders> at all (Word treats as no rules) used to replay as
        // 6 single-borders because EmitTable emitted no border prop and
        // AddTable's default-overlay won. The second dump then saw the
        // stamped borders and emitted six border.* props that the first
        // dump didn't — a 6× length asymmetry per affected table. Extend
        // the wipe to fire whenever no per-side / no-border-all key is
        // present in source's emit.
        {
            var sideKeys = new[] { "border.top", "border.bottom", "border.left",
                "border.right", "border.insideH", "border.insideV" };
            int presentSides = sideKeys.Count(s => tableProps.ContainsKey(s));
            bool hasBorderAll = tableProps.ContainsKey("border") || tableProps.ContainsKey("border.all");
            // BUG-STYLE-BORDER: the none-wipe (and AddTable's default-border
            // seed it counteracts) exist for tables that are GENUINELY
            // borderless (no inline <w:tblBorders>, no style). A table whose
            // borders come from a TABLE STYLE (a <w:tblStyle w:val="…"/> ref,
            // no inline borders) ALSO emits <6 per-side keys here — but it must
            // NOT get any inline override: writing all-none borders renders it
            // borderless, and AddTable's default single-border seed paints a
            // generic thin grid that masks the style's GridTable5Dark design.
            // For style-driven tables, instead tell AddTable to skip the
            // default-border seed entirely (skipDefaultBorders) so the style's
            // borders apply unmodified; any explicit inline border.* keys the
            // source DID carry still overlay on top. Genuinely borderless,
            // style-less tables keep the original none-wipe idempotency fix.
            // CONSISTENCY(border-default-overlay).
            bool hasTableStyleRef = tableProps.TryGetValue("style", out var styRef)
                && !string.IsNullOrEmpty(styRef);
            if (hasTableStyleRef)
            {
                // Suppress AddTable's generic 6-side single-border seed. The
                // table style supplies borders; explicit inline border.* keys
                // (if any) still apply via ApplyTableBorders after the seed is
                // skipped. No all-none wipe — that would defeat the style.
                tableProps["skipDefaultBorders"] = "true";
                // Still collapse 6 identical explicit per-side keys to the
                // compact form for idempotency (fall through to the else-if
                // below would be skipped otherwise).
                if (presentSides == 6 && !hasBorderAll)
                {
                    var firstS = tableProps[sideKeys[0]];
                    if (sideKeys.All(s => tableProps[s] == firstS))
                    {
                        foreach (var s in sideKeys) tableProps.Remove(s);
                        tableProps["border"] = firstS;
                    }
                }
            }
            else if (presentSides < 6 && !hasBorderAll)
            {
                // BUG-BORDER-NONE-INJECT: source <w:tblBorders> defined only a
                // SUBSET of the six sides (e.g. just insideH on a horizontal-
                // rule table), or had no <w:tblBorders> at all. The absent
                // sides emit no key — Word leaves them undefined, which on a
                // style-less table renders as no line. The old fix prepended a
                // `border=none;4` all-sides wipe so AddTable's default 6-single
                // seed wouldn't paint a generic grid; but that FABRICATES five
                // explicit <w:none> sides the source never had. Visually
                // identical (none == absent on a style-less table), yet a
                // re-dump reads the stamped none sides back as five extra
                // per-side `border.*=none;4` rows the first dump didn't emit —
                // a non-idempotent +5/+6 length asymmetry per table (frc tbl0:
                // insideH-only → 5 spurious none rows on the second dump).
                //
                // Suppress AddTable's default-border seed instead, so the
                // absent sides stay absent in the rebuilt XML, byte-matching
                // the source. The real per-side keys still overlay via
                // ApplyTableBorders. Same mechanism the styled-table branch
                // above already uses — extended to the style-less subset case.
                // CONSISTENCY(border-default-overlay).
                tableProps["skipDefaultBorders"] = "true";
            }
            // Symmetric collapse: when all 6 sides carry the IDENTICAL folded
            // value (same style + sz + color + space), prefer the compact
            // `border=<v>` form so dump round-trips that started from
            // "no <w:tblBorders>" (whose first emit becomes `border=none`)
            // re-emit the same single key after replay rather than fanning
            // out to six explicit per-side rows. ApplyTableBorders interprets
            // `border=<v>` as "set all 6 sides to <v>", so the visible result
            // is identical either way.
            else if (presentSides == 6 && !hasBorderAll)
            {
                var first = tableProps[sideKeys[0]];
                if (sideKeys.All(s => tableProps[s] == first))
                {
                    foreach (var s in sideKeys) tableProps.Remove(s);
                    tableProps["border"] = first;
                }
            }
        }
        // Nested tables sit inside a parent table cell; AddTable accepts
        // /body/tbl[N]/tr[M]/tc[K] as a parent. Outer-level tables target
        // /body. parentTablePath, when set, is a cell target path
        // (/body/tbl[X]/tr[Y]/tc[Z]) that we emit nested tables under.
        var tableParentPath = parentTablePath ?? containerPath;

        // tblPrChange round-trip — D1 emit shape.
        //
        // OLD path was: (1) `add table` with all props (rows/cols/width/
        // borders/...) AND (2) a follow-up no-op `set` carrying only
        // `revision.author=…` to re-stamp the marker. The second step's
        // BeginTrackChangeIfRequested snapshotted the just-finalized tblPr
        // as the "before" state — but that's also the "after" state since
        // nothing changed between (1) and (2). Word's reviewing pane then
        // silently dropped the revision because before==after.
        //
        // NEW path: when the source carried a tblPrChange, split into
        //   (1) `add table` with structural-only keys (rows / cols /
        //       gridCols / skipTblW) so the seeded tblPr is bare,
        //   (2) `set table` with EVERY other prop the source had + the
        //       attribution. BeginTrackChangeIfRequested now snapshots the
        //       bare-tblPr from (1) as "before" and captures the props
        //       applied in (2) as the diff — Word's reviewing pane sees a
        //       real change and surfaces it as Alice's tracked edit.
        //
        // Snapshot is still over-attributed (the source might only have
        // changed `width`, but all 10 current props get marked as
        // "changed" because we can't recover which subset was the real
        // edit). Over-attribute is the lesser evil; current behavior
        // under-attributes to silent loss.
        var tblPrAuthor = TryStringFormat(tableNode.Format, "tblPrChange.author");
        var tblPrDate = TryStringFormat(tableNode.Format, "tblPrChange.date");
        bool hasTblPrChange = !string.IsNullOrEmpty(tblPrAuthor);

        Dictionary<string, string> tableAddProps;
        Dictionary<string, string>? tableSetProps = null;
        if (hasTblPrChange)
        {
            tableAddProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            tableSetProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in tableProps)
            {
                if (k is "rows" or "cols" or "gridCols")
                    tableAddProps[k] = v;
                else
                    tableSetProps[k] = v;
            }
            // Force AddTable to seed an empty tblPr (no default tblW, no
            // default 6-border block). Without these, AddTable's defaults
            // tend to coincide with the source's set values (auto-fit
            // tblW = grid sum; single-4 borders are the most common case),
            // making the follow-up set a no-op in OOXML terms — the
            // snapshot then equals current state and Word's reviewing
            // pane silently hides the tblPrChange. With the seed
            // suppressed, the snapshot captures the bare tblPr and the
            // set's props produce a real diff Word will surface.
            tableAddProps["skipTblW"] = "true";
            tableAddProps["skipDefaultBorders"] = "true";
            tableSetProps["revision.type"] = "format";
            tableSetProps["revision.author"] = tblPrAuthor!;
            if (!string.IsNullOrEmpty(tblPrDate))
                tableSetProps["revision.date"] = tblPrDate!;
            // BUG-DUMP-R43-9: carry the verbatim prior-tblPr snapshot so the
            // tblPrChange records the real pre-change properties.
            var tblPrBeforeXml = TryStringFormat(tableNode.Format, "tblPrChange.beforeXml");
            if (tblPrBeforeXml != null)
                tableSetProps["revision.beforeXml"] = tblPrBeforeXml;
            // BUG-DUMP-R71-TBLPREX-CASCADE: suppress the apply-side per-row
            // tblPrEx cascade. That cascade is an interactive Mac-Word
            // visibility hack; on round-trip the source's real per-row tblPrEx
            // already replay verbatim via per-row `set tr --prop tblPrEx`, so
            // letting the cascade also run injects spurious tblPrEx into every
            // row (tables with a table-level tblPrChange but no per-row
            // exceptions went 0 → rows×2 tblPrEx and failed validation).
            tableSetProps["revision.skipRowCascade"] = "true";
        }
        else
        {
            tableAddProps = tableProps;
        }

        // Pin the column direction explicitly. AddTable's interactive
        // convenience auto-stamps <w:bidiVisual/> when the surrounding
        // context is RTL and no direction was passed — correct for a user
        // typing `add table` into an Arabic document, wrong for replay: a
        // source table WITHOUT bidiVisual (LTR columns inside an RTL doc)
        // came back visually mirrored. The reader emits direction=rtl only
        // when bidiVisual is present, so absence here means LTR — say so.
        if (!tableAddProps.ContainsKey("direction"))
            tableAddProps["direction"] = "ltr";

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tableParentPath,
            Type = "table",
            Props = tableAddProps
        });

        // BUG-R3 (nested-table multi-instance): a single cell may hold MORE
        // THAN ONE nested table stacked back-to-back (some editors' export
        // splits a logical table across several <w:tbl> siblings). The old
        // `tbl[1]` target hardcoded the FIRST nested table for every nested
        // emit, so the 2nd..Nth tables' per-cell `set tc[K]` ops resolved
        // against table #1 (wrong rows/cols) and produced thousands of
        // "Path not found …/tbl[1]/tr[N]/tc[K]". `add table` on a cell parent
        // appends (tbl[1], tbl[2], …, verified), so the just-added nested
        // table is always the cell's LAST table — mirror the outer-table
        // `tbl[last()]` convention so each nested table addresses itself.
        var tablePath = parentTablePath != null
            ? $"{parentTablePath}/tbl[last()]"
            : $"{containerPath}/tbl[last()]";

        if (tableSetProps != null && tableSetProps.Count > 0)
        {
            items.Add(new BatchItem
            {
                Command = "set",
                Path = tablePath,
                Props = tableSetProps
            });
        }

        // BUG-DUMP-R43-9: re-apply a <w:tblGridChange> (prior column-grid
        // tracked-change) verbatim. It has no author/date attribution to fold
        // and no Set key, so — like cellMerge.xml — it round-trips via a raw-set
        // append into the table's <w:tblGrid>. Guarded to body tables (the
        // (//w:tbl)[N] selector targets top-level body tables, same restriction
        // as the cell raw-set above); the SDK reorders tblGrid children to
        // schema order on save, so append is safe. Only emitted when the grid
        // doesn't already carry a tblGridChange from a width-Set side effect
        // (the snapshot here is the authoritative source state).
        // BUG-DUMP-R71-TBLGRIDCHANGE-DUP: when the table also carries a
        // tracked table-properties change (hasTblPrChange), the follow-up
        // `set` step replays the source colWidths under track-changes, and the
        // colWidths-Set-under-revision side effect ALREADY re-creates the
        // <w:tblGridChange> in the grid (see RestorePropsFromChange/gridChange
        // in Set.Revision.cs). Appending the verbatim snapshot here too then
        // duplicates it — two <w:tblGridChange> in one <w:tblGrid>, which
        // CT_TblGrid (gridCol* + tblGridChange?) rejects. Only emit the raw-set
        // append when the set side effect won't produce one (no tblPrChange, or
        // no colWidths to drive the grid change).
        bool gridChangeFromSetSideEffect = hasTblPrChange
            && tableSetProps != null
            && (tableSetProps.ContainsKey("colWidths") || tableSetProps.ContainsKey("colwidths"));
        if (containerPath == "/body"
            && !gridChangeFromSetSideEffect
            && tableNode.Format.TryGetValue("tblGridChange.xml", out var gridChangeRaw)
            && gridChangeRaw?.ToString() is { Length: > 0 } gridChangeXml)
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/document",
                Xpath = $"(//w:tbl)[{tableOrdinal}]/w:tblGrid",
                Action = "append",
                Xml = gridChangeXml,
            });
        }

        for (int r = 0; r < rows.Count; r++)
        {
            // Emit row-level properties (header / height / height.rule) as a
            // `set` on the row path — `add table` only seeds rows, it doesn't
            // surface per-row props (BUG-ROWPROPS). Without this, `dump→batch`
            // silently strips repeating-header rows and explicit row heights.
            var rowNode = rowNodes[r];
            // trPrChange D1 round-trip: fold the attribution into the
            // row's prop-set step so BeginTrackChangeIfRequested snapshots
            // the bare-trPr "before" vs the props-applied "after". If the
            // source had a trPrChange but no row-only props, still emit
            // the bare-attribution set so the marker exists (over-attributed
            // tail case, same lesser-evil as the tblPrChange edge case).
            var rowProps = ExtractRowOnlyProps(rowNode.Format);
            bool rowHadRevision = FoldRevisionIntoProps(rowNode.Format, "trPrChange", rowProps);
            // BUG-DUMP-R40-6: row-level tracked-change marker (<w:trPr><w:ins>/
            // <w:del>). The row reader stamps revision.type=ins|del +
            // revision.author/.date/.id directly on rowNode.Format (not under a
            // *PrChange prefix). Fold those creation keys into the row's `set tr`
            // step so SetElementTableRow → BeginTrackChangeIfRequested re-creates
            // the marker. Distinct from the trPrChange (format-change) fold above.
            bool rowHadInsDel = FoldRowRevisionInsDel(rowNode.Format, rowProps);
            if (rowProps.Count > 0 || rowHadRevision || rowHadInsDel)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{tablePath}/tr[{r + 1}]",
                    Props = rowProps
                });
            }
            var cells = rowCellNodes[r];
            for (int c = 0; c < cells.Count; c++)
            {
                var cellNode = word.Get(cells[c].Path);
                var cellTargetPath = $"{tablePath}/tr[{r + 1}]/tc[{c + 1}]";

                // BUG-DUMP-R26-7: the global-ordinal XPath of THIS cell, used by
                // EmitParagraph's inline raw-set fallbacks (rich field result,
                // nested SDT, VML textbox) to append verbatim content into the
                // correct cell instead of falling back to the lossy typed emit.
                // BUG-DUMP-R35-HFCELL: carried for body AND header/footer-hosted
                // tables. For a header/footer part the hfCtx's TableOrdinalBox is
                // fresh per part, so `(//w:tbl)[tableOrdinal]` is the part-local
                // DFS index the raw-set resolves against (same form the cell-SDT
                // block raw-set already uses with rawPart=containerPath). The
                // owning part travels alongside in cellRawPart so the cell raw-set
                // sites below — and ResolveRawSetHost for inline SDTs — target the
                // header/footer part instead of hardcoding "/document". Other
                // containers (footnote/endnote parts threaded as "/body") keep the
                // body form; their cell raw-set targeting is unchanged.
                bool cellRawAddressable = containerPath == "/body"
                    || IsHeaderFooterHost(containerPath);
                string? cellRawXPath = cellRawAddressable
                    ? $"(//w:tbl)[{tableOrdinal}]/w:tr[{r + 1}]/w:tc[{c + 1}]"
                    : null;
                string cellRawPart = containerPath == "/body" ? "/document" : containerPath;

                // Cell-level tcPr properties (fill, valign, width, borders,
                // padding, colspan, …) are surfaced on cellNode.Format but
                // were previously dropped — only the inner paragraph was
                // emitted. Push them via a `set` on the cell path before
                // the paragraph emits so cell shading / merges / widths
                // round-trip. Skip keys that EmitParagraph will re-apply
                // to the first paragraph (align/direction/run leak-throughs)
                // to avoid double-application.
                // tcPrChange D1 round-trip: fold attribution into the
                // cell's prop-set step (mirrors the row branch above).
                var cellProps = ExtractCellOnlyProps(cellNode.Format, tableIsAutofit);
                bool cellHadRevision = FoldRevisionIntoProps(cellNode.Format, "tcPrChange", cellProps);
                // BUG-DUMP-R51-2: carry a cell-level tracked insert/delete marker
                // (<w:tcPr><w:cellIns>/<w:cellDel>) onto the cell's `set` step so
                // SetElementTableCell re-stamps it. Distinct from tcPrChange (a
                // cell-property change) and from row-level trIns/trDel — the
                // reader surfaces it under cellRevision.* (see ReadCellProps).
                foreach (var crk in new[] { "cellRevision.type", "cellRevision.author", "cellRevision.date", "cellRevision.id" })
                    if (cellNode.Format.TryGetValue(crk, out var crv) && crv != null)
                        cellProps[crk] = crv.ToString()!;
                if (cellProps.Count > 0 || cellHadRevision)
                {
                    // CONSISTENCY(tblgrid-preserve): tcW values in the source
                    // are allowed to disagree with the gridCol widths (Word
                    // renders by tcW; tblGrid is a layout hint). Suppress
                    // Set.tc's tblGrid-sync side effect so AddTable's
                    // authoritative colWidths survives subsequent per-cell
                    // width sets.
                    if (cellProps.ContainsKey("width"))
                        cellProps["skipGridSync"] = "true";
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cellTargetPath,
                        Props = cellProps
                    });
                }

                // BUG-DUMP-R32-3: re-apply a <w:cellMerge> tracked-change marker
                // (cell split/merge under Track Changes) verbatim. It is not a
                // curated tcPr Set key, so it round-trips via raw-set into the
                // cell's <w:tcPr>. Appending INTO an existing tcPr is safe
                // (cellMerge ranks near the end of CT_TcPr), but a fresh tcPr
                // must be PREPENDED to the cell: CT_Tc requires tcPr as the
                // first child, and raw-set does not reorder tc children —
                // appending placed it after <w:p>, which the schema validator
                // rejects ("unexpected child element tcPr").
                // CONSISTENCY(tcpr-first): mirrors the
                // `cell.PrependChild(new TableCellProperties())` pattern used
                // by every tcPr-creation site in Add.Table/Set.Element.
                // When the cell already got a tcPr from the cellProps `set` above
                // (emitted earlier in item order), append into that existing
                // tcPr; otherwise wrap the marker in a fresh <w:tcPr> prepended
                // to the cell. Guarded to body-hosted tables (cellRawXPath != null),
                // matching the cell-SDT raw-set restriction; header/footer cells
                // emit a warning instead so the loss is never silent.
                if (cellNode.Format.TryGetValue("cellMerge.xml", out var cellMergeRaw)
                    && cellMergeRaw?.ToString() is { Length: > 0 } cellMergeXml)
                {
                    if (cellRawXPath != null)
                    {
                        bool tcPrExists = cellProps.Count > 0 || cellHadRevision;
                        items.Add(tcPrExists
                            ? new BatchItem
                            {
                                Command = "raw-set",
                                Part = cellRawPart,
                                Xpath = $"{cellRawXPath}/w:tcPr",
                                Action = "append",
                                Xml = cellMergeXml,
                            }
                            : new BatchItem
                            {
                                Command = "raw-set",
                                Part = cellRawPart,
                                Xpath = cellRawXPath,
                                Action = "prepend",
                                Xml = $"<w:tcPr>{cellMergeXml}</w:tcPr>",
                            });
                    }
                    else
                    {
                        ctx?.Warnings.Add(new DocxUnsupportedWarning(
                            Element: "cellMerge",
                            Path: cells[c].Path,
                            Reason: "A <w:cellMerge> tracked-change marker (cell split/merge "
                            + "under Track Changes) in a header/footer-hosted table "
                            + "cell was dropped on rebuild (cell raw-set targeting is "
                            + "limited to body tables)."));
                    }
                }

                // Each cell carries auto-generated paragraphs (Add table seeds
                // one empty paragraph per cell). Update the first one in place
                // and append further paragraphs as fresh adds. Nested tables
                // and paragraphs are emitted in document order so footnote/
                // chart cursors (carried in ctx) advance correctly through
                // the table cell content. Without ctx threading, body-level
                // footnote/chart references after a table would resolve
                // against the wrong note text.
                var cellChildren = cellNode.Children ?? new List<DocumentNode>();
                int cellParaIdx = 0;
                int nestedTblIdx = 0;
                bool firstParaSeen = false;
                bool cellSdtLeftSeed = false; // BUG-DUMP-R36-CELLSDT: SDT raw-set ahead of the cell's auto-seed paragraph
                // BUG-DUMP-CELLSDT-TRAILP: the typed `add sdt` path (plain /
                // text-shaped block SDT) CONSUMES the cell's auto-seed paragraph,
                // unlike the raw-set insert-before-seed path (cellSdtLeftSeed)
                // which preserves it. When the seed is consumed before any real
                // paragraph claims it, a following sibling paragraph must be a
                // fresh `add p` — otherwise it inherits autoPresent=true and its
                // `set p[last()]` targets a paragraph that no longer exists, the
                // step fails, and the trailing cell paragraph is silently dropped.
                bool cellSdtConsumedSeed = false;
                // BUG-DUMP-CELLSDT-2ND: EmitCellSdt's `cellHasContent` decides
                // insert-before-the-auto-seed (false) vs append (true). It was fed
                // `firstParaSeen`, which only tracks PARAGRAPHS — so a second SDT
                // after a first SDT (or after a nested table) still tried to insert
                // before the cell's seed <w:p>, but that seed was already consumed
                // (typed `add sdt`) or displaced, so the raw-set targeted a missing
                // paragraph and the SDT was dropped. Track ANY emitted cell content
                // (paragraph / table / SDT) so a non-leading SDT appends instead.
                bool cellHasAnyContent = false;

                // BUG-DUMP-R27-6: a block-level <w:customXml> wrapper that is a
                // DIRECT cell child is omitted from cellNode.Children (Navigation
                // surfaces only p/tbl/sdt for cells), so its inner paragraph text
                // was silently dropped — and, unlike the body path, no warning
                // fired. Detect direct customXml children from the cell's raw XML;
                // when present, drive the cell emit off a document-ordered plan
                // (EnumerateCellDirectChildren) that interleaves the customXml's
                // FLATTENED inner paragraphs with the normal p/tbl/sdt children,
                // matching the body's flatten+warn contract (inner text survives,
                // the custom-XML binding loss is reported loudly). The fast path
                // below stays unchanged when no customXml is present.
                var cellSourcePath = cells[c].Path;
                var cellDirectKinds = EnumerateCellDirectChildren(word.RawElementXml(cellSourcePath));
                bool cellHasCustomXml = cellDirectKinds.Contains("customXml");

                // BUG-DUMP-NESTED-TBL-TRAILING: OOXML requires every cell to
                // end with a paragraph (not a table). When a cell would
                // otherwise end with a table, the SDK auto-inserts a trailing
                // paragraph on save — so the cell's LAST paragraph following
                // a nested table is structurally auto-present on the target
                // side too, regardless of whether source's iteration already
                // used its autoPresent slot on a leading paragraph. Without
                // this, source [table, p] dumps `set p[last()]`
                // (autoPresent=true) but target [auto-p, table, p] re-dumps
                // `set p[1]` + `add p` and diverges by one row.
                int trailingAutoP = -1;
                for (int k = cellChildren.Count - 1; k >= 0; k--)
                {
                    var ct = cellChildren[k].Type;
                    if (ct != "paragraph" && ct != "p") continue;
                    if (k > 0 && cellChildren[k - 1].Type == "table")
                        trailingAutoP = k;
                    break;
                }

                // Cross-paragraph field spans INSIDE this cell (a TOC whose
                // fldChar begin lives in the first cell paragraph and its end
                // paragraphs later). The per-paragraph field collapse cannot
                // pair them, so the begin chain was warn-dropped and the end
                // silently filtered — the rebuilt TOC lost its field wrapper
                // (entries restyled as bare hyperlinks, lead text duplicated).
                // Mirror the body-level span machinery: raw-set each member
                // paragraph verbatim into the cell.
                var cellSpanParas = cellRawXPath != null
                    ? GetCellCrossParagraphFieldParaOrdinals(word, cellChildren)
                    : new HashSet<int>();

                if (!cellHasCustomXml)
                for (int k = 0; k < cellChildren.Count; k++)
                {
                    var cc = cellChildren[k];
                    if (cc.Type == "paragraph" || cc.Type == "p")
                    {
                        cellParaIdx++;
                        if (cellSpanParas.Contains(cellParaIdx))
                        {
                            var rawSpanP = word.GetElementXml(cc.Path);
                            if (!string.IsNullOrEmpty(rawSpanP) && !HasExternalRelRef(rawSpanP))
                            {
                                var rawSpanPart = containerPath == "/body" ? "/document" : containerPath;
                                items.Add(new BatchItem
                                {
                                    Command = "raw-set",
                                    Part = rawSpanPart,
                                    // The first cell paragraph replaces the
                                    // seeded empty paragraph AddTable created;
                                    // later members append after it.
                                    Xpath = firstParaSeen ? cellRawXPath! : $"{cellRawXPath}/w:p[1]",
                                    Action = firstParaSeen ? "append" : "replace",
                                    Xml = rawSpanP
                                });
                                firstParaSeen = true;
                                continue;
                            }
                            // Unresolvable (external rel inside the span):
                            // fall through to the typed emit and its warning.
                        }
                        // BUG-R4 (DBF-R4-02): a display equation (<m:oMathPara>)
                        // inside a cell surfaces here as a plain paragraph child
                        // whose Get returns an empty paragraph — EmitParagraph
                        // would emit `set p[N]` with no content and the formula
                        // would be lost. Mirror the body walker's typed routing:
                        // detect the oMathPara-wrapper and emit `add equation`
                        // targeting the cell paragraph instead. `add equation`
                        // (display) on an existing cell paragraph appends the
                        // m:oMathPara into it, reproducing the wrapper shape.
                        var cellEq = word.TryGetDisplayEquationAtParagraph(cc.Path);
                        if (cellEq != null)
                        {
                            // BUG-DUMP-CELLEQ-NESTEDTBL: match the plain-paragraph
                            // trailing-auto-p test — an equation paragraph directly
                            // after a nested table reuses the empty paragraph the
                            // SDK seeds AFTER that table, exactly like a plain
                            // paragraph does. The old test (k == trailingAutoP only)
                            // missed the equation when it wasn't the cell's LAST
                            // paragraph.
                            bool eqIsTrailingAutoP = k == trailingAutoP
                                || (k > 0 && cellChildren[k - 1].Type == "table");
                            // First cell paragraph (or the SDK auto-trailing one)
                            // reuses an auto-present seeded paragraph; otherwise
                            // create a fresh host paragraph for the equation.
                            if ((firstParaSeen || cellSdtConsumedSeed) && !eqIsTrailingAutoP)
                                items.Add(new BatchItem
                                {
                                    Command = "add",
                                    Parent = cellTargetPath,
                                    Type = "paragraph",
                                });
                            // When reusing the post-table trailing paragraph, target
                            // p[last()] — NOT p[cellParaIdx]. For a [nested-table,
                            // equation] cell, p[cellParaIdx] resolves to the cell's
                            // leading outer-seed paragraph, which the nested-lead
                            // `remove p[1]` then deletes, silently dropping the
                            // equation. p[last()] is the seeded trailing paragraph
                            // that survives (mirrors how a plain paragraph after a
                            // nested table reuses it via set p[last()]).
                            var eqTargetPath = eqIsTrailingAutoP
                                ? $"{cellTargetPath}/p[last()]"
                                : $"{cellTargetPath}/p[{cellParaIdx}]";
                            EmitCellDisplayEquation(word, cellEq, eqTargetPath, items);
                            firstParaSeen = true;
                            cellHasAnyContent = true;
                            continue;
                        }
                        // The FIRST paragraph after ANY nested table is also
                        // auto-present: at replay time `add table` momentarily
                        // leaves the cell ending in a table, so AddTable seeds
                        // an empty paragraph right after it. A plain `add p`
                        // then stacked a second paragraph and every following
                        // block shifted down (an extra blank line per nested
                        // table, eventually reflowing pages).
                        bool isTrailingAutoP = k == trailingAutoP
                            || (k > 0 && cellChildren[k - 1].Type == "table");
                        // BUG-DUMP-R26-7: publish THIS cell's raw-set XPath so the
                        // paragraph's inline raw-set fallbacks target the right
                        // cell. Re-set per paragraph because a preceding nested
                        // table recursion overwrote the box with its own cell.
                        if (ctx != null) { ctx.CurrentCellXPathBox[0] = cellRawXPath; ctx.CurrentCellPartBox[0] = cellRawPart; }
                        EmitParagraph(word, cc.Path, cellTargetPath, cellParaIdx, items,
                                      autoPresent: (!firstParaSeen && !cellSdtConsumedSeed) || isTrailingAutoP, ctx);
                        firstParaSeen = true;
                        cellHasAnyContent = true;
                    }
                    else if (cc.Type == "table")
                    {
                        nestedTblIdx++;
                        EmitTable(word, cc.Path, nestedTblIdx, items, ctx,
                                  parentTablePath: cellTargetPath, depth: depth + 1);
                        cellHasAnyContent = true;
                    }
                    else if (cc.Type == "sdt" && ctx != null)
                    {
                        // BUG-R11A(BUG1): a block-level <w:sdt> that is a direct
                        // child of this cell. Previously the cell walk recognised
                        // only paragraphs and nested tables, so the SDT (and its
                        // inner content) was dropped on dump. Emit it via the
                        // shared cell-SDT helper (typed `add sdt` for text-shaped
                        // controls, raw-set verbatim for rich block content).
                        // The raw-set xpath resolves to THIS cell by the table's
                        // document-order ordinal plus the current row/cell index.
                        var rawPart = containerPath == "/body" ? "/document" : containerPath;
                        var cellXPath = $"(//w:tbl)[{tableOrdinal}]/w:tr[{r + 1}]/w:tc[{c + 1}]";
                        // BUG-DUMP-H85: EmitCellSdt's `cellHasContent` chooses
                        // insert-before-the-auto-seed-<w:p> (false) vs append-to-cell
                        // (true). Feeding it `cellHasAnyContent` was wrong: a leading
                        // rich SDT inserts BEFORE the seed and PRESERVES it (returns
                        // sdtLeftSeed), yet it flipped cellHasAnyContent true, so a
                        // SECOND SDT that still precedes the cell's trailing paragraph
                        // appended (landing AFTER that paragraph) — silently reordering
                        // [sdt, sdt, p] to [sdt, p, sdt]. The real question is whether
                        // the seed <w:p> is still available as an insert-before anchor:
                        // it is, until a real paragraph claims it (firstParaSeen), a
                        // typed `add sdt` consumes it (cellSdtConsumedSeed), or a nested
                        // table is emitted (nestedTblIdx > 0). While the seed survives,
                        // successive `insertbefore w:p[1]` raw-sets stack in document
                        // order ([sdt1, sdt2, seed]); the trailing paragraph then claims
                        // the seed, yielding [sdt1, sdt2, p]. The BUG-DUMP-CELLSDT-2ND
                        // case (a typed SDT consumed the seed) still appends, via
                        // cellSdtConsumedSeed.
                        bool seedUnavailable = firstParaSeen || cellSdtConsumedSeed || nestedTblIdx > 0;
                        bool sdtLeftSeed = EmitCellSdt(word, cc.Path, cellTargetPath, cellXPath, rawPart,
                                    cellHasContent: seedUnavailable, items, ctx);
                        cellSdtLeftSeed |= sdtLeftSeed;
                        // Typed `add sdt` (returns false here with no prior cell
                        // content) consumed the auto-seed; the raw-set seed-left
                        // path returns true and keeps it. Flag the consumed case so
                        // a following sibling paragraph emits a fresh `add p`.
                        if (!sdtLeftSeed && !cellHasAnyContent) cellSdtConsumedSeed = true;
                        cellHasAnyContent = true;
                    }
                }

                // BUG-DUMP-R36-CELLSDT: a cell whose SOLE content is a block SDT
                // (e.g. a checkbox content control filling a rating-grid cell) has
                // no direct <w:p>, so no paragraph consumed the auto-seed paragraph
                // AddTable creates per cell. EmitCellSdt raw-set the SDT after the
                // <w:tcPr> (ahead of that seed), leaving a spurious trailing empty
                // paragraph the source never had — across a 39-cell grid that
                // inflated row heights enough to reflow a 5-page form to 6 pages.
                // When that insert-ahead-of-seed path fired (cellSdtLeftSeed) and
                // no real paragraph took the seed, remove the cell's direct
                // auto-seed <w:p>. The block SDT ends with a paragraph internally,
                // so the cell stays schema-valid (matches the source shape:
                // <w:tc><w:tcPr/><w:sdt/></w:tc>). Gated on cellSdtLeftSeed so the
                // typed `add sdt` / append paths (which leave no bare seed) aren't
                // hit with a remove that matches nothing.
                if (!cellHasCustomXml && cellSdtLeftSeed && !firstParaSeen && cellRawXPath != null)
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = cellRawPart,
                        Xpath = $"{cellRawXPath}/w:p",
                        Action = "remove",
                    });
                }

                // BUG-DUMP-R27-6: document-ordered cell walk used ONLY when the
                // cell carries a direct <w:customXml> child (the fast path above
                // — cellNode.Children — omits customXml entirely). Drive emission
                // off the raw-XML child order so a customXml interleaves correctly
                // with surrounding p/tbl/sdt children; p/tbl/sdt map by ordinal to
                // the cellNode.Children entries, customXml is flattened (its inner
                // paragraphs emit as cell paragraphs) and a deterministic warning
                // is recorded (matching the body customXmlPr arm in EmitBody).
                if (cellHasCustomXml)
                {
                    int planParaIdx = 0, planTblIdx = 0, planSdtIdx = 0, planCxIdx = 0;
                    foreach (var kind in cellDirectKinds)
                    {
                        if (kind == "p")
                        {
                            planParaIdx++;
                            var ccNode = NthChildOfType(cellChildren, "p", "paragraph", planParaIdx);
                            if (ccNode == null) continue;
                            cellParaIdx++;
                            if (ctx != null) { ctx.CurrentCellXPathBox[0] = cellRawXPath; ctx.CurrentCellPartBox[0] = cellRawPart; }
                            EmitParagraph(word, ccNode.Path, cellTargetPath, cellParaIdx, items,
                                          autoPresent: !firstParaSeen && !cellSdtConsumedSeed, ctx);
                            firstParaSeen = true;
                        }
                        else if (kind == "tbl")
                        {
                            planTblIdx++;
                            var ccNode = NthChildOfType(cellChildren, "table", "table", planTblIdx);
                            if (ccNode == null) continue;
                            nestedTblIdx++;
                            EmitTable(word, ccNode.Path, nestedTblIdx, items, ctx,
                                      parentTablePath: cellTargetPath, depth: depth + 1);
                        }
                        else if (kind == "sdt" && ctx != null)
                        {
                            planSdtIdx++;
                            var ccNode = NthChildOfType(cellChildren, "sdt", "sdt", planSdtIdx);
                            if (ccNode == null) continue;
                            var rawPart = containerPath == "/body" ? "/document" : containerPath;
                            var cellXPath = $"(//w:tbl)[{tableOrdinal}]/w:tr[{r + 1}]/w:tc[{c + 1}]";
                            // BUG-DUMP-H85: same seed-availability predicate as the fast
                            // path — a leading rich SDT preserves the seed, so a second
                            // SDT preceding the trailing paragraph must still insert
                            // before it, not append.
                            bool seedUnavailableCx = firstParaSeen || cellSdtConsumedSeed || planTblIdx > 0;
                            if (!EmitCellSdt(word, ccNode.Path, cellTargetPath, cellXPath, rawPart,
                                        cellHasContent: seedUnavailableCx, items, ctx) && !seedUnavailableCx)
                                cellSdtConsumedSeed = true;
                        }
                        else if (kind == "customXml")
                        {
                            planCxIdx++;
                            var cxPath = $"{cellSourcePath}/customXml[{planCxIdx}]";
                            // Warn first (loss of the element/uri/placeholder/attr
                            // binding) — mirrors the body customXmlPr arm.
                            if (ctx != null)
                            {
                                string? cxEl = null;
                                try
                                {
                                    var cxNode = word.Get(cxPath);
                                    if (cxNode.Format.TryGetValue("element", out var ev) && ev != null)
                                        cxEl = ev.ToString();
                                }
                                catch { /* best-effort descriptor */ }
                                var descr = string.IsNullOrEmpty(cxEl) ? "" : $" (element=\"{cxEl}\")";
                                ctx.Warnings.Add(new DocxUnsupportedWarning(
                                    Element: "customXml",
                                    Path: cxPath,
                                    Reason: $"block-level customXml wrapper{descr} in a table cell (custom-XML data binding: element/uri/placeholder/attr) dropped on dump→batch round-trip; the wrapped content's text survives but the binding does not"));
                            }
                            // Flatten the customXml's inner paragraphs into the
                            // cell, preserving their text (the body path does the
                            // same via Navigation's WalkBodyChild recursion). Inner
                            // tables/SDTs inside a cell customXml are out of scope
                            // this round — paragraphs cover the text-survival
                            // contract the warning advertises.
                            int innerP = 0;
                            while (true)
                            {
                                innerP++;
                                var innerPath = $"{cxPath}/p[{innerP}]";
                                DocumentNode? innerNode = null;
                                try { innerNode = word.Get(innerPath); }
                                catch { break; }
                                if (innerNode == null) break;
                                cellParaIdx++;
                                if (ctx != null) { ctx.CurrentCellXPathBox[0] = cellRawXPath; ctx.CurrentCellPartBox[0] = cellRawPart; }
                                EmitParagraph(word, innerPath, cellTargetPath, cellParaIdx, items,
                                              autoPresent: !firstParaSeen && !cellSdtConsumedSeed, ctx);
                                firstParaSeen = true;
                            }
                        }
                    }
                }

                // BUG-DUMP-R2-NESTED-LEAD: a cell whose FIRST source child is a
                // table has no leading source paragraph to reuse the empty
                // paragraph `add table` auto-seeds, so that seed survives as a
                // phantom blank line above the nested table (source [table, p] →
                // replay [p, table, p]). The trailing paragraph already landed
                // on the SDK's auto-trailing paragraph via trailingAutoP, so the
                // leading seed (cell's p[1]) is unconsumed — remove it. Validates
                // clean either way; this restores source structure.
                if (cellChildren.Count > 0 && cellChildren[0].Type == "table")
                {
                    items.Add(new BatchItem
                    {
                        Command = "remove",
                        Path = $"{cellTargetPath}/p[1]",
                    });
                }

                // BUG-DUMP-H97: cell-level (direct <w:tc> child) bookmark / perm
                // markers — between <w:tcPr> and the first paragraph (Google Docs cell
                // nav anchors, often column-span colFirst/colLast), or between/after
                // cell paragraphs — are skipped by the p/tbl/sdt cell walk above and
                // were silently dropped. Replay each verbatim (preserving id/name/
                // colFirst/colLast) at its paragraph-relative position via raw-set,
                // mirroring the header/footer-root structural-bookmark path. Emitted
                // AFTER the cell's paragraphs so the w:p[K] anchor resolves. Body
                // tables only (cellRawXPath != null); header/footer cells warn.
                var cellBms = word.GetCellStructuralBookmarks(cellSourcePath);
                if (cellBms.Count > 0)
                {
                    if (cellRawXPath != null)
                    {
                        foreach (var (bmXml, relXpath, action) in cellBms)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "raw-set",
                                Part = cellRawPart,
                                Xpath = relXpath == "." ? cellRawXPath : $"{cellRawXPath}/{relXpath}",
                                Action = action == "before" ? "insertbefore"
                                       : action == "after" ? "insertafter" : "append",
                                Xml = bmXml,
                            });
                        }
                    }
                    else
                    {
                        ctx?.Warnings.Add(new DocxUnsupportedWarning(
                            Element: "bookmark",
                            Path: cellSourcePath,
                            Reason: "cell-level bookmark/perm marker (direct <w:tc> child) in a "
                            + "header/footer-hosted table cell was dropped on rebuild "
                            + "(cell raw-set targeting is limited to body tables)."));
                    }
                }
            }
            // Trim trailing cells when source row is underfilled (sum of
            // source spans < gridCols). AddTable seeds `cols` cells per row;
            // `set tc[i] colspan=N` removes excess cells DOWN TO gridCols but
            // also PADS UP TO gridCols when the post-set total is short — so
            // a source row like [colspan=3] in a 4-col grid lands at 2 cells
            // post-replay (1 spanning + 1 pad). Source-shape preservation
            // demands removing (gridCols - sum_of_source_spans) trailing
            // cells AFTER all per-cell sets. The remove path is non-padding,
            // so the final cell count matches source. CONSISTENCY(table-row-
            // cell-count).
            int excessTrail = cols - rowEffectiveWidths[r];
            for (int e = 0; e < excessTrail; e++)
            {
                items.Add(new BatchItem
                {
                    Command = "remove",
                    Path = $"{tablePath}/tr[{r + 1}]/tc[last()]",
                });
            }
        }

        // Cell-level content controls: a <w:sdt> direct child of <w:tr> whose
        // sdtContent wraps the <w:tc> (Word's dropdown-bound cell). Navigation
        // flattens those to plain cells, and the inline-SDT emit demotes the
        // control to a run-level sdt INSIDE the cell — dropping the binding
        // from the other wrapped cells entirely. Patch each wrapped cell back
        // to its verbatim <w:sdt> block after all typed cell content has been
        // applied. Per row, replace in DESCENDING cell order: replacing
        // w:tc[3] with w:sdt removes it from the w:tc axis, which would shift
        // the index of every later w:tc in that row.
        if (containerPath == "/body")
        {
            for (int r = 0; r < rows.Count; r++)
            {
                var wrapped = word.GetSdtWrappedCellsOfRow(rowNodes[r].Path);
                for (int wi = wrapped.Count - 1; wi >= 0; wi--)
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = "/document",
                        Xpath = $"(//w:tbl)[{tableOrdinal}]/w:tr[{r + 1}]/w:tc[{wrapped[wi].CellOrdinal}]",
                        Action = "replace",
                        Xml = wrapped[wi].SdtXml,
                    });
                }
            }
        }
        // Row-level content controls: a <w:sdt> (SdtRow) direct child of
        // <w:tbl> whose sdtContent wraps a whole <w:tr> — Word's locked-row
        // shape, used by government forms to make an entire row read-only.
        // Navigation flattens these to plain rows (GetTableRowsFlattened) so
        // their cells/text round-trip via the typed emit above, but the SDT
        // wrapper and its <w:lock> would be lost. Patch each wrapped row back
        // to its verbatim <w:sdt> block after all typed row/cell content has
        // been applied. Replace in DESCENDING row order: replacing w:tr[N]
        // with <w:sdt> removes it from the w:tr axis, shifting the index of
        // every later w:tr. Mirrors the cell-wrapped pass above.
        // CONSISTENCY(sdt-wrapped-table).
        if (containerPath == "/body")
        {
            var wrappedRows = word.GetSdtWrappedRowsOfTable(sourcePath);
            for (int wi = wrappedRows.Count - 1; wi >= 0; wi--)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = $"(//w:tbl)[{tableOrdinal}]/w:tr[{wrappedRows[wi].RowOrdinal}]",
                    Action = "replace",
                    Xml = wrappedRows[wi].SdtXml,
                });
            }
        }
        // BUG-DUMP-TABLE-STRUCT-BOOKMARK: re-insert any <w:bookmarkStart>/<w:bookmarkEnd>
        // that sat at table-structure level (a direct child of <w:tbl> between rows,
        // or of <w:tr> between cells). The typed emit above only walks rows/cells, so
        // these cross-reference targets were dropped, leaving dangling PAGEREF/REF
        // ("Error! Bookmark not defined."). Replay each verbatim at its source
        // position via raw-set. Restricted to body tables, where the (//w:tbl)[N]
        // selector + /document part are reliable (same restriction as the tblGrid
        // raw-set above); header/footer/nested-table structural bookmarks are rare
        // and deferred.
        if (containerPath == "/body")
        {
            // BUG-DUMP-FF-BOOKMARK-DUP: a row-level bookmark that WRAPS a legacy
            // form field (FORMTEXT/FORMCHECKBOX) is emitted by TWO paths — the form
            // field's own `add formfield` recreates its wrapping bookmark (consuming
            // one unit of the per-name bookmark budget), and this structural
            // re-injection would emit it a SECOND time, duplicating the bookmark
            // name (Word de-dups/drops one, breaking the form field / REF). The cell
            // emit runs before this pass, so a form-field bookmark's budget is
            // already spent here: skip a lone named start whose budget is exhausted,
            // and its matching lone end (by id). A genuinely structural-only bookmark
            // (e.g. a _Toc heading anchor with no form field) still has budget and is
            // emitted (and accounted). Coalesced zero-length bookmarks (start+end in
            // one fragment) are structural-only and pass through.
            var skippedBmIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var (bmXml, relXpath, action) in word.GetTableStructuralBookmarks(sourcePath))
            {
                var starts = System.Text.RegularExpressions.Regex.Matches(bmXml, "<w:bookmarkStart\\b");
                var ends = System.Text.RegularExpressions.Regex.Matches(bmXml, "<w:bookmarkEnd\\b");
                // Lone named start: claim a budget unit; if none remain it was already
                // emitted by a form field — skip it and remember its id.
                if (starts.Count == 1 && ends.Count == 0)
                {
                    var nameM = System.Text.RegularExpressions.Regex.Match(bmXml, "w:name=\"([^\"]*)\"");
                    var idM = System.Text.RegularExpressions.Regex.Match(bmXml, "<w:bookmarkStart\\b[^>]*w:id=\"(\\d+)\"");
                    if (nameM.Success && ctx != null && !ctx.ConsumeBookmarkBudget(word, nameM.Groups[1].Value))
                    {
                        if (idM.Success) skippedBmIds.Add(idM.Groups[1].Value);
                        continue;
                    }
                }
                // Lone end whose matching start was skipped: drop it too.
                else if (ends.Count == 1 && starts.Count == 0)
                {
                    var idM = System.Text.RegularExpressions.Regex.Match(bmXml, "<w:bookmarkEnd\\b[^>]*w:id=\"(\\d+)\"");
                    if (idM.Success && skippedBmIds.Contains(idM.Groups[1].Value))
                        continue;
                }
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = $"(//w:tbl)[{tableOrdinal}]/{relXpath}",
                    Action = action,
                    Xml = bmXml,
                });
            }
        }

        // BUG-DUMP-R26-7: clear the cell-XPath context once this table is fully
        // emitted so body/header/footer content AFTER the table (or a parent
        // cell's content after a nested table) doesn't inherit a stale cell
        // address. A parent cell re-publishes its own XPath before its next
        // paragraph (see the per-paragraph set above), so null here is safe.
        if (ctx != null) { ctx.CurrentCellXPathBox[0] = null; ctx.CurrentCellPartBox[0] = null; }
    }

    // BUG-R4 (DBF-R4-02): emit a typed `add equation` (display) targeting a cell
    // paragraph path. Mirrors TryEmitDisplayEquation (WordBatchEmitter.Paragraph.cs)
    // but for an arbitrary cell-paragraph parent (TryEmitDisplayEquation is hard-
    // coded to parent "/body"). `add equation` on an existing cell paragraph
    // appends the m:oMathPara into it, reproducing the source wrapper shape.
    private static void EmitCellDisplayEquation(WordHandler word, DocumentNode eqNode, string parentPath, List<BatchItem> items)
    {
        var mode = eqNode.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode!
        };
        if (!string.IsNullOrEmpty(eqNode.Text))
            eqProps["formula"] = eqNode.Text!;
        // BUG-DUMP-CELLEQ-VERBATIM: forward the verbatim <m:oMath> so a cell
        // display equation keeps its math-run rPr (Cambria Math, sizes) instead
        // of being reparsed from the lossy LaTeX string. Without this the
        // equation rendered in the body font at the wrong metrics, shifting the
        // surrounding lines and drifting later content across page boundaries.
        // Mirrors TryEmitDisplayEquation (WordBatchEmitter.Paragraph.cs); the
        // body path was fixed in c0b0f015 but this cell path was missed.
        if (eqNode.Format.TryGetValue("xml", out var eqXml)
            && eqXml != null && eqXml.ToString() is { Length: > 0 } eqXmlS
            && eqXmlS.Contains("oMath", StringComparison.Ordinal))
            eqProps["xml"] = eqXmlS;
        // Carry any OLE/preview-image parts referenced inside the verbatim math
        // (MathType/Equation objects) so they don't dangle on replay.
        AddMathInlinedPartProps(word, eqNode.Path, eqProps);
        if (eqNode.Format.TryGetValue("align", out var eqAlign)
            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
            eqProps["align"] = eqAlign.ToString()!;
        // BUG-DUMP-CELLEQ-PPR: forward the wrapper paragraph's spacing/justification
        // so the rebuilt cell equation keeps its line height and alignment. Mirrors
        // TryEmitDisplayEquation.
        foreach (var sk in new[] { "lineSpacing", "lineRule", "spaceBefore", "spaceAfter", "wrapperAlign", "wrapperPpr" })
            if (eqNode.Format.TryGetValue(sk, out var sv)
                && sv != null && sv.ToString() is { Length: > 0 } svs)
                eqProps[sk] = svs;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentPath,
            Type = "equation",
            Props = eqProps
        });
    }

    // Cell Format includes both true tcPr keys and "leaked" keys read from
    // the first inner paragraph/run (align, direction, font, size, bold, …).
    // EmitParagraph re-emits those for the first paragraph, so emitting them
    // here too would double-apply. Whitelist genuine cell-level keys only.
    private static readonly HashSet<string> CellOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "width", "valign", "vmerge", "hmerge", "colspan", "nowrap", "textDirection",
        "cnfStyle",
        // BUG-DUMP-CELLTAIL: forward the long-tail tcPr toggles so dump→batch
        // round-trips them (Add.Table.cs / Set.cs apply both).
        "hideMark", "tcFitText",
    };

    private static Dictionary<string, string> ExtractCellOnlyProps(
        Dictionary<string, object?> raw, bool tableIsAutofit)
    {
        // BUG-DUMP-AUTOFITW: drop the fabricated width when the table is
        // autofit AND this cell's width was derived from tblGrid (no source
        // <w:tcW>). Re-emitting a synthetic tcW over-constrains Word's column
        // solver and shifts boundaries. Fixed-layout tables and cells with a
        // real source tcW (no _widthDerived marker) are unchanged — they keep
        // their width + the existing skipGridSync handling.
        bool widthDerived = raw.TryGetValue("_widthDerived", out var wd)
            && wd is bool wdb && wdb;
        bool suppressWidth = tableIsAutofit && widthDerived;
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (suppressWidth && key.Equals("width", StringComparison.OrdinalIgnoreCase))
                continue;
            if (CellOnlyKeys.Contains(key) ||
                key.StartsWith("border.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("shading.", StringComparison.OrdinalIgnoreCase))
            {
                filtered[key] = val;
            }
        }
        // BUG-DUMP21-02: when shading.* sub-keys are present, the
        // FilterEmittableProps shading-fold will emit a folded `shading`
        // key carrying val+fill+color. The legacy `fill` alias surfaced by
        // ReadCellProps duplicates the same color and would cause Set tc
        // to apply the bare-color form on top of the folded shading,
        // overwriting val/color. Drop it here so only the canonical folded
        // form replays.
        if (filtered.Keys.Any(k => k.StartsWith("shading.", StringComparison.OrdinalIgnoreCase)))
        {
            filtered.Remove("fill");
        }
        // Negative cell margins (<w:tcMar w:w="-13">) are schema-valid and real
        // Word produces them (tight tables whose text bleeds slightly into the
        // border zone), but the Set/Add padding path deliberately rejects a
        // negative w:tcMar (BUG-R1-07). A verbatim dump of a negative margin
        // therefore emits a `set tc padding=-13` step the rebuild rejects —
        // round-trip self-conflict (2 failed steps). Clamp to 0 on emit: a
        // -13-twip (~0.02cm) margin is visually indistinguishable from 0, so the
        // rebuild renders identically and stays clean. (Accepting negative tcMar
        // project-wide is the alternative but would reverse the BUG-R1-07 cell
        // padding contract — out of scope for a fidelity round-trip.)
        foreach (var pk in filtered.Keys
                     .Where(k => k.Equals("padding", StringComparison.OrdinalIgnoreCase)
                              || k.StartsWith("padding.", StringComparison.OrdinalIgnoreCase))
                     .ToList())
        {
            if (filtered[pk]?.ToString() is { } pv
                && int.TryParse(pv, out var pn) && pn < 0)
                filtered[pk] = "0";
        }
        return FilterEmittableProps(filtered);
    }

    // Row-level keys emitted by Navigation.ReadRowProps. Used by EmitTable
    // so dump→batch round-trips header rows / heights / cantSplit. Cell
    // children are emitted separately via ExtractCellOnlyProps.
    private static readonly HashSet<string> RowOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "header", "height", "cantSplit", "cnfStyle",
        // BUG-DUMP-R37-3: row-level <w:hidden/> (row not displayed/printed).
        "hidden",
        // BUG-DUMP-R24-1: row-level <w:jc> (whole-row alignment).
        "rowAlign",
        // BUG-DUMP-R24-4: per-row <w:tblPrEx> overrides (verbatim element).
        "tblPrEx",
        // BUG-DUMP-R42-2: leading/trailing grid-column skips + their preferred
        // widths (ragged/indented table edge). Carried through `set tr` so
        // SetElementTableRow re-emits <w:gridBefore>/<w:wBefore>/<w:gridAfter>/<w:wAfter>.
        "gridBefore", "wBefore", "gridAfter", "wAfter",
        // BUG-DUMP-R62-ROWCELLSPACING: row-level <w:tblCellSpacing> (inter-cell
        // gap for this row). SetElementTableRow re-emits it onto the row's trPr.
        "cellSpacing",
    };

    /// <summary>Read a string-valued key from a DocumentNode.Format dict
    /// (Format values are typed as <c>object?</c>). Returns null when
    /// the key is missing, the value is null, or the string is empty.
    /// Used by the D1 round-trip path to detect whether a host element
    /// carried a `*PrChange` marker in source.</summary>
    private static string? TryStringFormat(Dictionary<string, object?> format, string key)
    {
        if (!format.TryGetValue(key, out var obj) || obj == null) return null;
        var s = obj.ToString();
        return string.IsNullOrEmpty(s) ? null : s;
    }

    /// <summary>Fold the source's `<paramref name="prefix"/>.author` /
    /// `.date` keys into an existing prop bag as a `revision.type=format`
    /// + `revision.author` + `revision.date` triplet. Called by the
    /// row / cell emit paths so the structural-prop `set` and the
    /// revision attribution travel in one batch step — the
    /// BeginTrackChangeIfRequested snapshot then captures the bare-pr
    /// "before" state vs the just-applied "after" state, producing a
    /// real *PrChange diff Word's reviewing pane will surface (instead
    /// of the legacy two-step emit's `before==after` lie).
    ///
    /// Returns true when a revision was folded in. Mutates
    /// <paramref name="props"/> in place.</summary>
    private static bool FoldRevisionIntoProps(
        Dictionary<string, object?> format,
        string prefix,
        Dictionary<string, string> props)
    {
        var author = TryStringFormat(format, $"{prefix}.author");
        if (author == null) return false;
        props["revision.type"] = "format";
        props["revision.author"] = author;
        var date = TryStringFormat(format, $"{prefix}.date");
        if (date != null) props["revision.date"] = date;
        // BUG-DUMP-R43-9: carry the verbatim prior-properties snapshot so
        // BeginTrackChangeIfRequested restores the REAL pre-change tcPr/trPr
        // (what Reject-Change recovers) instead of the over-attributed
        // current-state snapshot. One mechanism, all *PrChange hosts.
        var beforeXml = TryStringFormat(format, $"{prefix}.beforeXml");
        if (beforeXml != null) props["revision.beforeXml"] = beforeXml;
        return true;
    }

    /// <summary>BUG-DUMP-R40-6: fold a row-level ins/del tracked-change marker
    /// (read from <c>revision.type=ins|del</c> + <c>revision.author/.date/.id</c>
    /// stamped directly on the row node) into the row's `set tr` prop bag so
    /// SetElementTableRow re-creates <c>&lt;w:trPr&gt;&lt;w:ins&gt;/&lt;w:del&gt;</c>.
    /// Returns true when an ins/del marker was folded in. No-op (returns false)
    /// for any other revision.type — the trPrChange format-change fold owns
    /// those.</summary>
    private static bool FoldRowRevisionInsDel(
        Dictionary<string, object?> format,
        Dictionary<string, string> props)
    {
        var type = TryStringFormat(format, "revision.type");
        if (type is not ("ins" or "del")) return false;
        props["revision.type"] = type!;
        var author = TryStringFormat(format, "revision.author");
        if (author != null) props["revision.author"] = author;
        var date = TryStringFormat(format, "revision.date");
        if (date != null) props["revision.date"] = date;
        var id = TryStringFormat(format, "revision.id");
        if (id != null) props["revision.id"] = id;
        return true;
    }

    private static Dictionary<string, string> ExtractRowOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        // BUG-DUMP-R25-1: translate the readback's height + height.rule into the
        // rule-specific apply key. Absent height.rule = AUTO row-sizing → bare
        // `height` (no @w:hRule injected). Exact/atLeast map to the explicit
        // keys so the source rule round-trips faithfully.
        string? heightRule = null;
        if (raw.TryGetValue("height.rule", out var ruleObj))
            heightRule = ruleObj?.ToString();
        foreach (var (key, val) in raw)
        {
            if (!RowOnlyKeys.Contains(key)) continue;
            if (string.Equals(key, "height", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(heightRule, "exact", StringComparison.OrdinalIgnoreCase))
                    filtered["height.exact"] = val;
                else if (string.Equals(heightRule, "atLeast", StringComparison.OrdinalIgnoreCase))
                    filtered["height.atleast"] = val;
                else
                    filtered["height"] = val;
            }
            else
            {
                filtered[key] = val;
            }
        }
        return FilterEmittableProps(filtered);
    }

    // Cross-paragraph field spans inside a single table cell: returns the
    // 1-based paragraph-child ordinals covered by any span whose fldChar
    // begin/end pair straddles paragraph boundaries. Mirrors
    // WordHandler.GetCrossParagraphFieldSpanRanges (body-level); a non-
    // paragraph child interrupts an open span, and an unterminated span is
    // abandoned so its paragraphs fall back to the typed emit.
    private static HashSet<int> GetCellCrossParagraphFieldParaOrdinals(
        WordHandler word, List<DocumentNode> cellChildren)
    {
        var members = new HashSet<int>();
        var pending = new List<int>();
        int paraOrdinal = 0, depth = 0;
        bool open = false;
        foreach (var cc in cellChildren)
        {
            if (cc.Type != "paragraph" && cc.Type != "p")
            {
                if (open) { open = false; depth = 0; pending.Clear(); }
                continue;
            }
            paraOrdinal++;
            var xml = word.GetElementXml(cc.Path) ?? "";
            int begins = System.Text.RegularExpressions.Regex.Matches(
                xml, "fldCharType=\"begin\"").Count;
            int ends = System.Text.RegularExpressions.Regex.Matches(
                xml, "fldCharType=\"end\"").Count;
            if (!open)
            {
                if (begins > ends)
                {
                    open = true;
                    depth = begins - ends;
                    pending.Clear();
                    pending.Add(paraOrdinal);
                }
            }
            else
            {
                pending.Add(paraOrdinal);
                depth += begins - ends;
                if (depth <= 0)
                {
                    foreach (var o in pending) members.Add(o);
                    open = false; depth = 0; pending.Clear();
                }
            }
        }
        return members;
    }
}
