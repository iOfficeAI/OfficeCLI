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
        EmitThemeRaw(word, items);
        EmitNumberingRaw(word, items);
        EmitSettingsRaw(word, items);
        EmitSection(word, items);
        EmitHeadersFooters(word, items);
        var paraIdToTargetIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        EmitBody(word, items, paraIdToTargetIdx);
        EmitComments(word, items, paraIdToTargetIdx);
        return items;
    }

    private static void EmitThemeRaw(WordHandler word, List<BatchItem> items)
    {
        // Theme carries clrScheme + fontScheme + fmtScheme — pure structured
        // XML that users rarely modify property-by-property; the natural
        // operation is "swap the entire theme block". Raw-set replace fits
        // that model exactly. Word.Raw returns the literal string
        // "(no theme)" when the part is missing — gate on a leading '<' so
        // we only emit when there's real XML to ship.
        string xml;
        try { xml = word.Raw("/theme"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/theme",
            Xpath = "/a:theme",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitSettingsRaw(WordHandler word, List<BatchItem> items)
    {
        // Settings carries dozens of feature flags + compat shims that
        // surface on root.Format only piecemeal — and not all of them are
        // wired through Set's case table. Wholesale raw-set is the simplest
        // way to keep Word feature toggles (evenAndOddHeaders, mirrorMargins,
        // schema-pegged compat options, …) round-tripped without
        // per-property allowlisting.
        string xml;
        try { xml = word.Raw("/settings"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/settings",
            Xpath = "/w:settings",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitNumberingRaw(WordHandler word, List<BatchItem> items)
    {
        // Numbering models list templates (abstractNum + num pairs, each
        // abstractNum holds 9 levels with their own pPr / numFmt / lvlText).
        // Reconstructing this through typed Add would mean another emitter
        // in itself; for v0.5 we ship the entire <w:numbering> XML wholesale
        // via raw-set. The blank document creates an empty numbering part,
        // so a single replace on the part root is sufficient.
        string xml;
        try { xml = word.Raw("/numbering"); }
        catch { return; }
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return;
        // Skip when numbering is empty (just `<w:numbering/>` with no children).
        if (!xml.Contains("<w:abstractNum") && !xml.Contains("<w:num "))
            return;

        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/numbering",
            Xpath = "/w:numbering",
            Action = "replace",
            Xml = xml
        });
    }

    private static void EmitHeadersFooters(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        if (root.Children == null) return;
        int hIdx = 0, fIdx = 0;
        foreach (var child in root.Children)
        {
            if (child.Type == "header")
            {
                hIdx++;
                EmitHeaderFooterPart(word, child.Path, "header", hIdx, items);
            }
            else if (child.Type == "footer")
            {
                fIdx++;
                EmitHeaderFooterPart(word, child.Path, "footer", fIdx, items);
            }
        }
    }

    private static void EmitHeaderFooterPart(WordHandler word, string sourcePath, string kind,
                                             int targetIndex, List<BatchItem> items)
    {
        var partNode = word.Get(sourcePath);
        var paras = (partNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "paragraph" || c.Type == "p")
            .ToList();
        var subType = partNode.Format.TryGetValue("type", out var t) ? t?.ToString() ?? "default" : "default";

        // Seed the part with the first paragraph's text (AddHeader/AddFooter
        // create a single auto paragraph and accept text/align/style on it).
        // Multi-run first paragraphs collapse into a flat text string here —
        // run-level formatting on the seed paragraph is a v0.5 lossy item.
        var seedProps = new Dictionary<string, string> { ["type"] = subType };
        if (paras.Count > 0)
        {
            // Get on /header[1] returns paragraph stubs without their run
            // children — re-Get the first paragraph to surface its runs.
            var firstPara = word.Get(paras[0].Path);
            var firstRuns = (firstPara.Children ?? new List<DocumentNode>())
                .Where(c => c.Type == "run" || c.Type == "r")
                .ToList();
            if (firstRuns.Count == 1 && !string.IsNullOrEmpty(firstRuns[0].Text))
            {
                seedProps["text"] = firstRuns[0].Text!;
                var runProps = FilterEmittableProps(firstRuns[0].Format);
                foreach (var (k, v) in runProps)
                    if (!seedProps.ContainsKey(k)) seedProps[k] = v;
            }
            else if (firstRuns.Count >= 1)
            {
                // Multi-run: collapse plain text only, drop per-run formatting.
                seedProps["text"] = string.Join("", firstRuns.Select(r => r.Text ?? ""));
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/",
            Type = kind,
            Props = seedProps
        });

        // Additional paragraphs (>= 2nd) appended to the part directly.
        var partTargetPath = $"/{kind}[{targetIndex}]";
        for (int p = 1; p < paras.Count; p++)
        {
            EmitParagraph(word, paras[p].Path, partTargetPath, p + 1, items, autoPresent: false);
        }
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

    // Root-level keys that round-trip via `set /`. Includes section page
    // layout, document protection, doc-level grid + defaults. Excludes
    // metadata that auto-updates on save (created/modified timestamps,
    // lastModifiedBy, package author/title — those re-stamp anyway).
    private static readonly HashSet<string> RootScalarKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // Section page layout (mirrors body's trailing sectPr)
        "pageWidth", "pageHeight", "orientation",
        "marginTop", "marginBottom", "marginLeft", "marginRight",
        "pageStart", "pageNumFmt",
        "titlePage", "direction", "rtlGutter",
        "lineNumbers", "lineNumberCountBy",
        // Document protection
        "protection", "protectionEnforced",
        // Document grid (CJK-aware line layout)
        "charSpacingControl",
    };

    // Dotted-prefix groups that round-trip wholesale via `set /`. Each
    // sub-key is forwarded as-is; the schema-reflection layer routes the
    // dotted path into the right OOXML target.
    private static readonly string[] RootPrefixGroups = new[]
    {
        "docDefaults.",
        "docGrid.",
    };

    private static void EmitSection(WordHandler word, List<BatchItem> items)
    {
        var root = word.Get("/");
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (k, v) in root.Format)
        {
            bool include = RootScalarKeys.Contains(k);
            if (!include)
            {
                foreach (var pref in RootPrefixGroups)
                {
                    if (k.StartsWith(pref, StringComparison.OrdinalIgnoreCase))
                    {
                        include = true;
                        break;
                    }
                }
            }
            if (!include) continue;
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

    private sealed record ChartSpec(Dictionary<string, object?> Format, IReadOnlyList<DocumentNode> Series);

    private sealed record BodyEmitContext(
        List<string> FootnoteTexts,
        List<string> EndnoteTexts,
        NoteCursor FootnoteCursor,
        NoteCursor EndnoteCursor,
        List<ChartSpec> ChartSpecs,
        NoteCursor ChartCursor,
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
        // Charts: query("chart") returns /chart[N] in document order, which
        // matches the order chart-bearing runs appear in body. Pre-resolve
        // each chart's properties + series children so EmitParagraph can
        // emit a typed `add chart` row when it walks across each ref.
        var charts = word.Query("chart");
        var chartSpecs = charts.Select(c =>
        {
            var full = word.Get(c.Path);
            return new ChartSpec(full.Format, full.Children ?? new List<DocumentNode>());
        }).ToList();

        var ctx = new BodyEmitContext(
            FootnoteTexts: word.Query("footnote").Select(n => n.Text ?? "").ToList(),
            EndnoteTexts: word.Query("endnote").Select(n => n.Text ?? "").ToList(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: chartSpecs,
            ChartCursor: new NoteCursor(),
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

        // Inline section break: a paragraph carrying <w:sectPr> is the
        // OOXML representation of a mid-document section boundary.
        // AddSection on /body produces this same shape, so we emit
        // `add /body --type section` (which creates a fresh break paragraph)
        // rather than emitting a regular `add p`. The companion
        // sectionBreak.* keys map back to AddSection's prop vocabulary.
        if (parentPath == "/body" &&
            pNode.Format.TryGetValue("sectionBreak", out var breakKind) && breakKind != null)
        {
            var sectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = breakKind.ToString() ?? "nextPage"
            };
            foreach (var (k, v) in pNode.Format)
            {
                if (!k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase)) continue;
                if (v == null) continue;
                var keyTail = k["sectionBreak.".Length..];
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) sectProps[keyTail] = s;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "section",
                Props = sectProps
            });
            return;
        }

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
            .Where(c => c.Type == "run" || c.Type == "r" || c.Type == "picture")
            .ToList();

        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        // Picture runs need their own typed `add picture` row, so the
        // collapse only applies when the sole run is a regular text run.
        bool collapseSingleRun = runs.Count <= 1 &&
            !(runs.Count == 1 && runs[0].Type == "picture");
        if (collapseSingleRun)
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
            // Drawing-bearing runs surface as type=="picture" regardless of
            // whether the Drawing wraps an image (Blip) or a chart
            // (c:chart). Try the image path first; if there's no embedded
            // image part the run is a chart anchor — pull the next
            // pre-resolved ChartSpec and emit a typed `add chart` row.
            if (run.Type == "picture")
            {
                var binary = word.GetImageBinary(run.Path);
                if (binary.HasValue)
                {
                    var (bytes, contentType) = binary.Value;
                    var dataUri = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
                    var picProps = FilterEmittableProps(run.Format);
                    picProps.Remove("id");
                    picProps.Remove("contentType");
                    picProps.Remove("fileSize");
                    picProps["src"] = dataUri;
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "picture",
                        Props = picProps
                    });
                    continue;
                }

                if (ctx != null && ctx.ChartCursor.Index < ctx.ChartSpecs.Count)
                {
                    var spec = ctx.ChartSpecs[ctx.ChartCursor.Index];
                    ctx.ChartCursor.Index++;
                    var chartProps = BuildChartProps(spec);
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "chart",
                        Props = chartProps
                    });
                    continue;
                }
                // Drawing without image part and no matching chart spec —
                // unsupported anchor (OLE/SmartArt). Skip silently.
                continue;
            }

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

        // Column count must cover the widest row including colspan effects.
        // Format["cols"] reflects gridCol; per-row effective width is
        // sum(colspan or 1) over each cell. Take the max so a first row
        // with merged cells (visible cell count < grid width) doesn't
        // truncate the table shape and break later `set tc[N]` rows.
        var rowEffectiveWidths = new List<int>(rows.Count);
        var rowCellNodes = new List<List<DocumentNode>>(rows.Count);
        foreach (var rowChild in rows)
        {
            var rowNode = word.Get(rowChild.Path);
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
        int cols = Math.Max(colsFromGrid, colsFromRows);
        if (cols == 0) return;

        var tableProps = FilterEmittableProps(tableNode.Format);
        tableProps["rows"] = rows.Count.ToString();
        tableProps["cols"] = cols.ToString();
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
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cellTargetPath,
                        Props = cellProps
                    });
                }

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

    // Cell Format includes both true tcPr keys and "leaked" keys read from
    // the first inner paragraph/run (align, direction, font, size, bold, …).
    // EmitParagraph re-emits those for the first paragraph, so emitting them
    // here too would double-apply. Whitelist genuine cell-level keys only.
    private static readonly HashSet<string> CellOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "width", "valign", "vmerge", "colspan", "nowrap", "textDirection",
    };

    private static Dictionary<string, string> ExtractCellOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (CellOnlyKeys.Contains(key) ||
                key.StartsWith("border.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase))
            {
                filtered[key] = val;
            }
        }
        return FilterEmittableProps(filtered);
    }

    private static Dictionary<string, string> BuildChartProps(ChartSpec spec)
    {
        // AddChart ingests data series via a single `data="Name1:v1,v2;Name2:v1,v2"`
        // string. Reconstruct that string from the series children Get
        // exposes; categories come from the chart's own Format key.
        var props = FilterEmittableProps(spec.Format);
        // Strip Get-only / SDK-managed keys that AddChart neither expects
        // nor accepts.
        props.Remove("id");
        props.Remove("seriesCount");

        // Build data="Name:v1,v2;..." from series children.
        var seriesParts = new List<string>();
        foreach (var s in spec.Series)
        {
            if (s.Type != "series") continue;
            if (!s.Format.TryGetValue("name", out var nObj) || nObj == null) continue;
            if (!s.Format.TryGetValue("values", out var vObj) || vObj == null) continue;
            var name = nObj.ToString() ?? "";
            var vals = vObj.ToString() ?? "";
            if (name.Length == 0 || vals.Length == 0) continue;
            seriesParts.Add($"{name}:{vals}");
        }
        if (seriesParts.Count > 0)
        {
            props["data"] = string.Join(";", seriesParts);
        }
        return props;
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
