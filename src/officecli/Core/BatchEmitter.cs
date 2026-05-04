// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Handlers;

namespace OfficeCli.Core;

/// <summary>
/// Walks an opened handler's document tree and emits a sequence of BatchItem
/// rows that, when replayed against a blank document of the same format,
/// reconstruct the original document.
///
/// <para>
/// This is the core of the `officecli dump --format batch` pipeline. The
/// emit relies on the OOXML schema reflection fallback in
/// <see cref="TypedAttributeFallback"/> + <see cref="GenericXmlQuery"/>:
/// any leaf property that Get reads can be re-applied via Add/Set, so
/// emit just transcribes Format keys directly without per-property
/// allowlisting.
/// </para>
///
/// <para>
/// Scope (v0.5): docx body paragraphs (with run formatting) + tables (single
/// paragraph + single run per cell, common case). Resources (styles,
/// numbering, theme, headers, footers, sections, comments, footnotes,
/// endnotes) and richer cell contents are NOT yet emitted — follow-up
/// passes will add them.
/// </para>
/// </summary>
public static class BatchEmitter
{
    /// <summary>Emit a batch sequence for a Word document.</summary>
    public static List<BatchItem> EmitWord(WordHandler word)
    {
        var items = new List<BatchItem>();

        // Phase order matters: resources first so body refs (style=Heading1,
        // numId=3, etc.) resolve when the paragraph adds reach them on replay.
        EmitStyles(word, items);
        EmitSection(word, items);
        var paraIdToTargetIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        EmitBody(word, items, paraIdToTargetIdx);
        EmitComments(word, items, paraIdToTargetIdx);
        return items;
    }

    private static void EmitComments(WordHandler word, List<BatchItem> items,
                                     Dictionary<string, int> paraIdToTargetIdx)
    {
        var comments = word.Query("comment");
        foreach (var c in comments)
        {
            var props = FilterEmittableProps(c.Format);
            if (!string.IsNullOrEmpty(c.Text))
                props["text"] = c.Text!;
            // Map anchoredTo (source paraId path) -> target paragraph index.
            // anchoredTo looks like "/body/p[@paraId=00100000]"; parse and
            // resolve via the paraId map we built during EmitBody.
            string parentTarget = "/body/p[1]";  // safe fallback to first body para
            if (props.TryGetValue("anchoredTo", out var anchor))
            {
                var pid = ExtractParaId(anchor);
                if (pid != null && paraIdToTargetIdx.TryGetValue(pid, out var idx))
                    parentTarget = $"/body/p[{idx}]";
                props.Remove("anchoredTo");
            }
            // The comment id is allocated by AddComment on the target side;
            // do not propagate the source id (would conflict on replay).
            props.Remove("id");
            // Date is auto-stamped by the SDK on add — emitting it would
            // overwrite the user's local "now" with the source moment, which
            // is rarely the desired round-trip behaviour.
            props.Remove("date");

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentTarget,
                Type = "comment",
                Props = props
            });
        }
    }

    private static string? ExtractParaId(string anchorPath)
    {
        var m = System.Text.RegularExpressions.Regex.Match(anchorPath, @"@paraId=([0-9A-Fa-f]+)");
        return m.Success ? m.Groups[1].Value : null;
    }

    // Section-level keys that root.Format exposes. Theme / docDefaults /
    // settings / protection live on root too but each gets its own emit
    // phase, so this list intentionally omits them.
    private static readonly HashSet<string> SectionKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "pageWidth", "pageHeight", "orientation",
        "marginTop", "marginBottom", "marginLeft", "marginRight",
        "pageStart", "pageNumFmt",
        "titlePage", "direction", "rtlGutter",
        "lineNumbers", "lineNumberCountBy",
    };

    private static void EmitSection(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in root.Format)
        {
            if (!SectionKeys.Contains(k)) continue;
            if (v == null) continue;
            var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s.Length > 0) props[k] = s;
        }
        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props
        });
    }

    private static void EmitStyles(WordHandler word, List<BatchItem> items)
    {
        // Use query() rather than walking Get("/styles").Children — the
        // positional /styles/style[N] children Get returns are not
        // addressable on the Get side (style paths resolve by id, not by
        // index). Query produces id-based paths and excludes docDefaults.
        var styles = word.Query("style");
        foreach (var stub in styles)
        {
            DocumentNode full;
            try { full = word.Get(stub.Path); }
            catch { continue; }
            var props = FilterEmittableProps(full.Format);
            // Ensure id is present (Add requires it for /styles target).
            if (!props.ContainsKey("id") && !props.ContainsKey("styleId"))
            {
                if (props.TryGetValue("name", out var n)) props["id"] = n;
                else continue;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/styles",
                Type = "style",
                Props = props
            });
        }
    }

    private sealed class NoteCursor { public int Index; }

    private sealed record BodyEmitContext(
        List<string> FootnoteTexts,
        List<string> EndnoteTexts,
        NoteCursor FootnoteCursor,
        NoteCursor EndnoteCursor,
        Dictionary<string, int>? ParaIdToTargetIdx);

    private static void EmitBody(WordHandler word, List<BatchItem> items,
                                 Dictionary<string, int>? paraIdToTargetIdx = null)
    {
        var bodyNode = word.Get("/body");
        if (bodyNode.Children == null) return;

        // Footnotes/endnotes are referenced by runs (rStyle=FootnoteReference)
        // inside body paragraphs but the run carries no id back to the
        // notes part. We assume notes are listed in document order matching
        // reference order — the typical case since AddFootnote/AddEndnote
        // allocate ids sequentially.
        var ctx = new BodyEmitContext(
            FootnoteTexts: word.Query("footnote").Select(n => n.Text ?? "").ToList(),
            EndnoteTexts: word.Query("endnote").Select(n => n.Text ?? "").ToList(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ParaIdToTargetIdx: paraIdToTargetIdx);

        int pIndex = 0, tblIndex = 0;
        foreach (var child in bodyNode.Children)
        {
            switch (child.Type)
            {
                case "paragraph":
                case "p":
                    pIndex++;
                    EmitParagraph(word, child.Path, "/body", pIndex, items, autoPresent: false, ctx);
                    break;
                case "table":
                    tblIndex++;
                    EmitTable(word, child.Path, tblIndex, items);
                    break;
                case "section":
                case "sectPr":
                    // The body always carries one trailing sectPr that the
                    // blank document already provides; for v0.5 we rely on
                    // that default and skip emitting section properties.
                    // Section emit is a follow-up.
                    break;
                default:
                    // Unknown body-level child types (sdt, etc.) — skip for v0.5.
                    break;
            }
        }
    }

    /// <summary>
    /// Emit a paragraph at the target index under <paramref name="parentPath"/>.
    /// When <paramref name="autoPresent"/> is true, the parent already has a
    /// pre-existing paragraph at that index (e.g. an auto-created table cell
    /// paragraph); we issue a `set` instead of a fresh `add` so the existing
    /// paragraph gets reused rather than duplicated.
    /// </summary>
    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent,
                                      BodyEmitContext? ctx = null)
    {
        var pNode = word.Get(sourcePath);

        // Track source paraId -> target index so comments anchored on this
        // paragraph can be retargeted on replay (paraIds regenerate in the
        // target document, so positional indices are the stable handle).
        if (ctx?.ParaIdToTargetIdx != null && parentPath == "/body" &&
            pNode.Format.TryGetValue("paraId", out var paraIdVal) && paraIdVal != null)
        {
            ctx.ParaIdToTargetIdx[paraIdVal.ToString()!] = targetIndex;
        }

        var props = FilterEmittableProps(pNode.Format);
        var runs = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "run" || c.Type == "r")
            .ToList();

        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        if (runs.Count <= 1)
        {
            if (runs.Count == 1)
            {
                var runProps = FilterEmittableProps(runs[0].Format);
                foreach (var (k, v) in runProps)
                {
                    if (!props.ContainsKey(k)) props[k] = v;
                }
                if (!string.IsNullOrEmpty(runs[0].Text))
                    props["text"] = runs[0].Text!;
            }

            if (autoPresent)
            {
                // Replace the auto-created paragraph in place — only push the
                // set when there is something to apply, otherwise the empty
                // skeleton is already correct.
                if (props.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{parentPath}/p[{targetIndex}]",
                        Props = props
                    });
                }
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = parentPath,
                    Type = "p",
                    Props = props.Count > 0 ? props : null
                });
            }
            return;
        }

        // Multi-run paragraph: emit/set the paragraph empty first, then add
        // each run as an explicit child.
        if (autoPresent)
        {
            if (props.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{parentPath}/p[{targetIndex}]",
                    Props = props
                });
            }
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "p",
                Props = props.Count > 0 ? props : null
            });
        }

        var paraTargetPath = $"{parentPath}/p[{targetIndex}]";
        foreach (var run in runs)
        {
            // Detect footnote/endnote reference runs. The OOXML model marks
            // them with a w:rStyle = FootnoteReference / EndnoteReference;
            // the run itself carries no visible text. Emit them as a
            // typed footnote/endnote add anchored on the host paragraph and
            // pull the body text from the pre-resolved ordered list — see
            // BodyEmitContext for the document-order assumption.
            var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
            if (ctx != null && rStyle == "FootnoteReference")
            {
                var noteText = ctx.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                    ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index]
                    : "";
                ctx.FootnoteCursor.Index++;
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "footnote",
                    Props = new() { ["text"] = noteText }
                });
                continue;
            }
            if (ctx != null && rStyle == "EndnoteReference")
            {
                var noteText = ctx.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                    ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index]
                    : "";
                ctx.EndnoteCursor.Index++;
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "endnote",
                    Props = new() { ["text"] = noteText }
                });
                continue;
            }

            var rProps = FilterEmittableProps(run.Format);
            if (!string.IsNullOrEmpty(run.Text))
                rProps["text"] = run.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    private static void EmitTable(WordHandler word, string sourcePath, int targetIndex, List<BatchItem> items)
    {
        var tableNode = word.Get(sourcePath);
        var rows = (tableNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "row")
            .ToList();
        if (rows.Count == 0) return;

        // Pull cell count from the first row. Column count emitted by Get
        // (Format["cols"]) reflects the gridCol count, which can drift from
        // actual cells if the table has merges; the row's own count is the
        // safer bet for shape during replay.
        int cellsInFirstRow = 0;
        var row0 = word.Get(rows[0].Path);
        if (row0.Children != null)
            cellsInFirstRow = row0.Children.Count(c => c.Type == "cell");
        if (cellsInFirstRow == 0) return;

        var tableProps = FilterEmittableProps(tableNode.Format);
        tableProps["rows"] = rows.Count.ToString();
        tableProps["cols"] = cellsInFirstRow.ToString();
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "table",
            Props = tableProps
        });

        var tablePath = $"/body/tbl[{targetIndex}]";
        for (int r = 0; r < rows.Count; r++)
        {
            var rowNode = word.Get(rows[r].Path);
            var cells = (rowNode.Children ?? new List<DocumentNode>())
                .Where(c => c.Type == "cell")
                .ToList();
            for (int c = 0; c < cells.Count; c++)
            {
                var cellNode = word.Get(cells[c].Path);
                var cellTargetPath = $"{tablePath}/tr[{r + 1}]/tc[{c + 1}]";

                // Each cell carries auto-generated paragraphs (Add table seeds
                // one empty paragraph per cell). Update the first one in place
                // and append further paragraphs as fresh adds.
                var cellParas = (cellNode.Children ?? new List<DocumentNode>())
                    .Where(x => x.Type == "paragraph" || x.Type == "p")
                    .ToList();
                for (int p = 0; p < cellParas.Count; p++)
                {
                    EmitParagraph(word, cellParas[p].Path, cellTargetPath, p + 1, items,
                                  autoPresent: p == 0);
                }
            }
        }
    }

    // Format keys that must NOT be emitted: derived (computed by Get, not
    // user-set), unstable (regenerate on save), or coordinate-system
    // (paths that only make sense in the source document).
    private static readonly HashSet<string> SkipKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "basedOn.path",
        "paraId", "textId", "rsidR", "rsidRDefault", "rsidRPr", "rsidP", "rsidTr",
    };

    private static Dictionary<string, string> FilterEmittableProps(Dictionary<string, object?> raw)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (SkipKeys.Contains(key)) continue;
            if (key.StartsWith("effective.", StringComparison.OrdinalIgnoreCase)) continue;
            if (key.EndsWith(".cs.source", StringComparison.OrdinalIgnoreCase)) continue;

            // BORDER subattr asymmetry: Get exposes `border.top: single` AND
            // `border.top.sz: 4` / `border.top.color: 808080` as separate keys,
            // but Set's case table stops at the 2-segment level — the 3-segment
            // sub-attribute keys would be misrouted through ApplyTableBorders'
            // dotted fallback and crash on `Invalid border style: '4'`. Drop
            // them here as a known lossy projection until Set grows the
            // matching cases (border width / color readback survive only via
            // the main `border.*` style key for now).
            if (key.StartsWith("border.", StringComparison.OrdinalIgnoreCase) &&
                key.Count(ch => ch == '.') >= 2)
            {
                continue;
            }

            if (val == null) continue;
            var s = val switch
            {
                bool b => b ? "true" : "false",
                _ => val.ToString() ?? ""
            };
            if (s.Length > 0) result[key] = s;
        }
        return result;
    }
}
