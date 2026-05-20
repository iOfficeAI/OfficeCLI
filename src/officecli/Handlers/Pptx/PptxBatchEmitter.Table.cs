// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(emit-table-mirror): mirrors WordBatchEmitter.Table.cs in
    // shape — emit `add table` with rows/cols, then per-row `set tr[N]`,
    // per-cell `set tc[K]`, and finally cell text via `set tc[K] text=...`.
    // PPT tables are simpler than Word tables (no nested tables, no
    // tblGrid/tblBorders aggregate elements), so this is a much smaller
    // method than the docx version.

    private static void EmitTable(PowerPointHandler ppt, DocumentNode tableNode,
                                  string parentSlidePath, string replayPath,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        // depth=2 so /slide/table/tr/tc cell nodes materialize with text.
        var fullTable = ppt.Get(tableNode.Path, depth: 2);
        var props = FilterEmittableProps(fullTable.Format);
        if (!props.ContainsKey("rows") || !props.ContainsKey("cols")) return;

        // AddTable seeds rows×cols empty cells; per-cell text + per-row
        // height + per-cell tcPr (fill/borders/padding/valign/spans) get
        // pushed via subsequent `set` rows. Avoid re-emitting the `data=`
        // shortcut form — it's mutually exclusive with rows/cols and would
        // hide per-cell formatting we want to preserve.
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "table",
            Props = props.Count > 0 ? props : null,
        });

        var tablePath = replayPath;
        if (fullTable.Children == null) return;
        var rows = fullTable.Children.Where(c => c.Type == "tr").ToList();

        int rIdx = 0;
        foreach (var row in rows)
        {
            rIdx++;
            var rowProps = FilterEmittableProps(row.Format);
            // Row height — Set tr accepts `height=`; other row-level keys
            // round-trip through the per-cell set path so emit only the
            // narrow whitelist here.
            var emittedRow = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (rowProps.TryGetValue("height", out var h))
                emittedRow["height"] = h;
            if (emittedRow.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{tablePath}/tr[{rIdx}]",
                    Props = emittedRow,
                });
            }

            int cIdx = 0;
            foreach (var cell in row.Children ?? new List<DocumentNode>())
            {
                if (cell.Type != "tc") continue;
                cIdx++;
                var cellProps = FilterEmittableProps(cell.Format);
                // Set tc accepts text= for replacing the cell's text body.
                if (!string.IsNullOrEmpty(cell.Text))
                    cellProps["text"] = cell.Text!;
                if (cellProps.Count == 0) continue;
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{tablePath}/tr[{rIdx}]/tc[{cIdx}]",
                    Props = cellProps,
                });
            }
        }
    }
}
