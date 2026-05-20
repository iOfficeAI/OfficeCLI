// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    private static void EmitTable(WordHandler word, string sourcePath, int targetIndex,
                                  List<BatchItem> items, BodyEmitContext? ctx = null,
                                  string? parentTablePath = null,
                                  string containerPath = "/body")
    {
        var tableNode = word.Get(sourcePath);
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
            if (presentSides < 6 && !hasBorderAll)
            {
                // Use the canonical "style;size" form: ApplyTableBorders'
                // ParseBorderValue defaults size=4 for `none`, so writing
                // `none;4` matches what the round-trip produces (six explicit
                // <w:none w:sz="4"> elements collapsed by the all-same fold
                // below). Without the `;4`, the FIRST dump emits `border=none`
                // and the SECOND dump emits `border=none;4` — non-idempotent
                // value shape.
                tableProps["border"] = "none;4";
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
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tableParentPath,
            Type = "table",
            Props = tableProps
        });

        // For nested tables, the target path is parent_cell/tbl[1] (first
        // table in the cell). For outer tables, it's /body/tbl[N].
        var tablePath = parentTablePath != null
            ? $"{parentTablePath}/tbl[1]"
            : $"{containerPath}/tbl[last()]";
        for (int r = 0; r < rows.Count; r++)
        {
            // Emit row-level properties (header / height / height.rule) as a
            // `set` on the row path — `add table` only seeds rows, it doesn't
            // surface per-row props (BUG-ROWPROPS). Without this, `dump→batch`
            // silently strips repeating-header rows and explicit row heights.
            var rowNode = rowNodes[r];
            var rowProps = ExtractRowOnlyProps(rowNode.Format);
            if (rowProps.Count > 0)
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

                // Cell-level tcPr properties (fill, valign, width, borders,
                // padding, colspan, …) are surfaced on cellNode.Format but
                // were previously dropped — only the inner paragraph was
                // emitted. Push them via a `set` on the cell path before
                // the paragraph emits so cell shading / merges / widths
                // round-trip. Skip keys that EmitParagraph will re-apply
                // to the first paragraph (align/direction/run leak-throughs)
                // to avoid double-application.
                var cellProps = ExtractCellOnlyProps(cellNode.Format);
                if (cellProps.Count > 0)
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
                foreach (var cc in cellChildren)
                {
                    if (cc.Type == "paragraph" || cc.Type == "p")
                    {
                        cellParaIdx++;
                        EmitParagraph(word, cc.Path, cellTargetPath, cellParaIdx, items,
                                      autoPresent: !firstParaSeen, ctx);
                        firstParaSeen = true;
                    }
                    else if (cc.Type == "table")
                    {
                        nestedTblIdx++;
                        EmitTable(word, cc.Path, nestedTblIdx, items, ctx,
                                  parentTablePath: cellTargetPath);
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
    }

    // Cell Format includes both true tcPr keys and "leaked" keys read from
    // the first inner paragraph/run (align, direction, font, size, bold, …).
    // EmitParagraph re-emits those for the first paragraph, so emitting them
    // here too would double-apply. Whitelist genuine cell-level keys only.
    private static readonly HashSet<string> CellOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "width", "valign", "vmerge", "hmerge", "colspan", "nowrap", "textDirection",
    };

    private static Dictionary<string, string> ExtractCellOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
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
        return FilterEmittableProps(filtered);
    }

    // Row-level keys emitted by Navigation.ReadRowProps. Used by EmitTable
    // so dump→batch round-trips header rows / heights / cantSplit. Cell
    // children are emitted separately via ExtractCellOnlyProps.
    private static readonly HashSet<string> RowOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "header", "height", "cantSplit",
    };

    private static Dictionary<string, string> ExtractRowOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        bool heightExact = false;
        if (raw.TryGetValue("height.rule", out var ruleObj) &&
            string.Equals(ruleObj?.ToString(), "exact", StringComparison.OrdinalIgnoreCase))
        {
            heightExact = true;
        }
        foreach (var (key, val) in raw)
        {
            if (!RowOnlyKeys.Contains(key)) continue;
            // height + height.rule=exact → SetElementTableRow expects key
            // `height.exact`. Translate so dump output applies cleanly.
            if (heightExact && string.Equals(key, "height", StringComparison.OrdinalIgnoreCase))
            {
                filtered["height.exact"] = val;
            }
            else
            {
                filtered[key] = val;
            }
        }
        return FilterEmittableProps(filtered);
    }
}
