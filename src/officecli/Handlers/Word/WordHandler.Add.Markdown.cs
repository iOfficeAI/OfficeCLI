// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Core.Markdown;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    /// <summary>
    /// `add --type markdown` — the docx analogue of `add --type diagram`.
    /// Parses a Markdown subset via the shared, format-neutral
    /// <see cref="MarkdownParser"/> (Core/Markdown, like MermaidParser) and
    /// expands it into ordinary paragraphs / tables at <paramref name="parentPath"/>.
    ///
    /// Like `diagram`, this is ADD-ONLY: there is no persistent "markdown"
    /// node — the expansion produces normal Word elements, so edit those
    /// afterwards (no matching Set/Get/Query/Remove). The block→element and
    /// inline→run mapping reuses the handler's OWN <c>Add</c>/<c>Set</c> entry
    /// points (each returns the real path), so no OOXML is hand-built here and
    /// no Core→handler dependency is introduced.
    ///
    /// Input mirrors `diagram`: canonical `markdown` (+ aliases `text`/`md`)
    /// inline, or `src`/`path` to a .md file.
    /// </summary>
    private string AddMarkdown(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        var mdText = properties.GetValueOrDefault("markdown")
                     ?? properties.GetValueOrDefault("text")
                     ?? properties.GetValueOrDefault("md");
        if (string.IsNullOrWhiteSpace(mdText)
            && (properties.TryGetValue("src", out var srcFile) || properties.TryGetValue("path", out srcFile))
            && !string.IsNullOrWhiteSpace(srcFile))
        {
            if (!System.IO.File.Exists(srcFile))
                throw new ArgumentException($"markdown source file not found: '{srcFile}'.");
            mdText = System.IO.File.ReadAllText(srcFile);
        }
        if (string.IsNullOrWhiteSpace(mdText))
            throw new ArgumentException("markdown requires inline 'markdown' text (aliases: text, md) or a 'src' .md file path.");

        var doc = MarkdownParser.Parse(mdText);

        // Position: the first block honors the caller's --index/--after/--before;
        // each subsequent block chains AfterElement(previous) so document order
        // is preserved when inserting mid-document. With no position, every
        // block appends (pos = null). Requires the exit-invalidate guard in
        // Add() — anchor navigation used to leave a stale body child-index
        // (see WordAddAnchorCacheTests) which broke the per-cell Set after an
        // anchored `add table`.
        string? firstPath = null;
        string? lastPath = null;

        // Two-phase: phase 1 adds every block (pure appends stay on the
        // append-fast path), phase 2 applies all inline range formatting.
        // Interleaving Add + Set-by-path was O(n²): every body-level Add
        // clears the body child-index/paraId caches, so each Set's path
        // navigation rebuilt them from scratch (8k formatted lines ≈ 147s).
        // Deferred, the first Set rebuilds once and the rest hit the cache.
        var pendingInline = new List<(string Path, List<MdSpan> Inlines)>();

        InsertPosition? PosFor()
        {
            // No explicit index: every block appends at the document end, in
            // document order — so a plain append (null) keeps each block on the
            // O(1) append hot path. Chaining InsertPosition.AfterElement(lastPath)
            // instead routed every block through the anchor path, which navigates
            // to (and rebuilds the body child-index cache from) the pre-mutation
            // tree on each call: O(n) per block, O(n²) overall (an 8000-line list
            // took ~37s). Only when an explicit start index was given must blocks
            // chain after one another to stay contiguous and ordered at that index.
            if (index.HasValue)
                return lastPath != null ? InsertPosition.AfterElement(lastPath) : InsertPosition.AtIndex(index.Value);
            return null;
        }

        void Record(string path)
        {
            firstPath ??= path;
            lastPath = path;
        }

        // Defer per-operation saves across BOTH phases. Every block below goes
        // through the public Add()/Set() entry points, each of which ends in
        // SaveDoc() — a full main-part re-serialization. Left eager, an N-block
        // document serializes the whole (growing) part N times: O(N²) — the
        // dominant cost (an 8000-line list spent ~27s almost entirely here).
        // The outer Add() that dispatched to us performs the single real save
        // after we return (its SaveDoc() runs with the flag restored). Restore
        // the caller's flag so batch-replay / resident semantics are untouched.
        var _savedDeferSave = DeferSave;
        DeferSave = true;

        // Atomicity: markdown expansion is many block Adds under a single
        // deferred save. If a later block throws (e.g. a GFM table exceeding the
        // OOXML column limit) the blocks already appended must NOT survive — the
        // operation reported failure, so a half-applied document is corruption.
        // Snapshot the target's direct children up front; on any failure, remove
        // every child appended since and rethrow. Directed undo (not
        // DiscardOnDispose) so a long-lived resident session's OTHER unsaved
        // edits are untouched. KNOWN LIMITATION: list blocks may have minted
        // numbering abstractNum/num definitions that this DOM-child removal does
        // not reclaim — those become orphan (unreferenced) numbering entries,
        // invisible and harmless in Word, not part of the rendered content.
        var _preExisting = new HashSet<OpenXmlElement>(parent.ChildElements);

        // Accumulate per-block diagnostics. Every block expands through the
        // public Add(), which RESETS LastAddWarnings / LastAddUnsupportedProps /
        // LastUnrecognizedLatex at its entry (WordHandler.Add.cs) — so left
        // as-is only the final block's diagnostics survive to the caller (a
        // heading's "style not found" warning vanished the moment a trailing
        // list block ran its own Add()). Union each block's diagnostics into
        // locals as we go, then write the merged set back before returning.
        var accWarnings = new List<string>();
        var accUnsupported = new List<string>();
        var accLatex = new List<string>();
        void HarvestBlockDiagnostics()
        {
            if (LastAddWarnings.Count > 0) accWarnings.AddRange(LastAddWarnings);
            if (LastAddUnsupportedProps.Count > 0) accUnsupported.AddRange(LastAddUnsupportedProps);
            if (LastUnrecognizedLatex.Count > 0) accLatex.AddRange(LastUnrecognizedLatex);
        }
        try
        {
        foreach (var block in doc.Blocks)
        {
            switch (block)
            {
                case MdHeading h:
                    Record(AddHeading(parentPath, h, PosFor(), pendingInline));
                    HarvestBlockDiagnostics();
                    break;
                case MdParagraph p:
                    Record(AddFlowParagraph(parentPath, p.Inlines, null, pendingInline, pos: PosFor()));
                    HarvestBlockDiagnostics();
                    break;
                case MdBlockQuote q:
                    // Blockquote: left-indented + italic direct formatting. The
                    // built-in "Quote" style is absent from a blank docx, so we
                    // format directly rather than reference a style that renders
                    // as plain body text.
                    Record(AddFlowParagraph(parentPath, q.Inlines, null, pendingInline,
                        extra: new(StringComparer.OrdinalIgnoreCase) { ["indentLeft"] = "720", ["italic"] = "true" },
                        pos: PosFor()));
                    HarvestBlockDiagnostics();
                    break;
                case MdList list:
                    // AddFlowList makes one Add() per item, each resetting the
                    // diagnostics; harvest after every item so no item's warning
                    // is clobbered by the next.
                    AddFlowList(parentPath, list, depth: 0, Record, PosFor, pendingInline, HarvestBlockDiagnostics);
                    break;
                case MdCodeBlock code:
                    // Monospace via direct font; built-in code styles are absent
                    // from a blank docx.
                    foreach (var line in code.Code.Split('\n'))
                    {
                        Record(Add(parentPath, "paragraph", PosFor(),
                            new(StringComparer.OrdinalIgnoreCase) { ["font"] = "Consolas", ["text"] = line }));
                        HarvestBlockDiagnostics();
                    }
                    break;
                case MdHorizontalRule:
                    Record(Add(parentPath, "paragraph", PosFor(),
                        new(StringComparer.OrdinalIgnoreCase) { ["borderBottom"] = "single" }));
                    HarvestBlockDiagnostics();
                    break;
                case MdTable t:
                    Record(AddFlowTable(parentPath, t, PosFor()));
                    HarvestBlockDiagnostics();
                    break;
            }
        }

        // Phase 2: all structural adds done — apply inline formatting in one
        // sweep so path navigation builds the body caches once and reuses them.
        // ApplyInlineSpans mutates the DOM directly (no SaveDoc), so this sweep
        // never re-serialized per line; it stays inside the DeferSave scope only
        // so the empty-markdown fallback Add() below also defers.
        foreach (var (path, inlines) in pendingInline)
            ApplyInlineSpans(path, inlines);

        // Empty markdown still yields at least an anchor so `add` returns a path.
        // Compute the result FIRST (the fallback Add resets diagnostics, but it
        // only fires when no block was added, so nothing accumulated is lost),
        // then surface the merged per-block diagnostics to the caller.
        var result = firstPath ?? Add(parentPath, "paragraph", PosFor(),
            new(StringComparer.OrdinalIgnoreCase) { ["text"] = "" });

        // Advisory: links/images are flattened to plain text in this expander
        // (no w:hyperlink relationship, no embedded picture — a documented
        // limitation, real import is a v2 item). Surface a warning instead of
        // truncating silently (the gap has been re-reported repeatedly).
        int flattened = CountFlattenedReferences(doc);
        if (flattened > 0)
            accWarnings.Add($"{flattened} link/image reference(s) flattened to plain text (hyperlink/image import lands in v2)");

        LastAddWarnings = accWarnings;
        LastAddUnsupportedProps = accUnsupported;
        LastUnrecognizedLatex = accLatex;
        return result;
        }
        catch
        {
            // Directed rollback: drop every block appended during this failed
            // expansion so nothing partially-applied is left to be persisted by
            // the outer save / Dispose autosave.
            foreach (var added in parent.ChildElements.Where(c => !_preExisting.Contains(c)).ToList())
                added.Remove();
            // Bulk removal invalidates the append-monotonic body caches.
            InvalidateBodyParaCache();
            throw;
        }
        finally
        {
            DeferSave = _savedDeferSave;
        }
    }

    // Heading font sizes (pt) per level — a compact hierarchy that reads as
    // headings even in a blank docx with no built-in Heading styles defined.
    private static readonly int[] HeadingSizes = { 18, 16, 14, 13, 12, 11 };

    private string AddHeading(string parentPath, MdHeading h, InsertPosition? pos,
        List<(string Path, List<MdSpan> Inlines)> pendingInline)
    {
        int size = HeadingSizes[Math.Clamp(h.Level, 1, 6) - 1];
        // Reference the built-in style id AND apply direct bold+size. The style
        // ref keeps TOC / navigation working if the doc later gains Heading
        // definitions; the direct formatting guarantees the visual hierarchy
        // now (a blank docx ships Heading1..9 as bare scaffolding that renders
        // like body text).
        var extra = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["style"] = $"Heading{Math.Clamp(h.Level, 1, 6)}",
            ["bold"] = "true",
            ["size"] = size.ToString(),
        };
        return AddFlowParagraph(parentPath, h.Inlines, null, pendingInline, extra, pos);
    }

    private string AddFlowParagraph(string parentPath, List<MdSpan> inlines, string? style,
        List<(string Path, List<MdSpan> Inlines)> pendingInline,
        Dictionary<string, string>? extra = null, InsertPosition? pos = null)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["text"] = string.Concat(inlines.Select(s => s.Text)),
        };
        if (style != null) props["style"] = style;
        if (extra != null)
            foreach (var kv in extra) props[kv.Key] = kv.Value;

        var path = Add(parentPath, "paragraph", pos, props);
        if (inlines.Any(s => s.Bold || s.Italic || s.Code || s.Strike))
            pendingInline.Add((path, inlines)); // formatted in phase 2 (perf)
        return path;
    }

    private void AddFlowList(string parentPath, MdList list, int depth, Action<string> record,
        Func<InsertPosition?> posFor,
        List<(string Path, List<MdSpan> Inlines)> pendingInline,
        Action harvestDiagnostics)
    {
        // Use liststyle (real numbering: bullets / 1. 2.) rather than a
        // ListBullet/ListNumber style id — those style ids are absent from a
        // blank docx and would render as plain paragraphs with no marker.
        // Nesting maps to the numbering ilvl via `level` (0-based, capped 0-8).
        var listStyle = list.Ordered ? "ordered" : "unordered";
        var level = Math.Clamp(depth, 0, 8).ToString();

        // Decide THIS list's numId once. The first item's marker paragraph is
        // created via `liststyle` (which mints — or continues from a preceding
        // block-level list — a numId); its assigned numId is captured and reused
        // EXPLICITLY (numId=) for every subsequent item. Otherwise an item whose
        // content includes code paragraphs (Normal style) would break
        // FindContinuationNumId's "scan back to the previous list paragraph"
        // guess — the code paragraph terminates the scan, so the next item minted
        // a fresh numId and real Word restarted numbering (e.g. "1." then "1.").
        // Reusing one numId for all siblings keeps a single list. (Nested child
        // lists recurse and decide their own numId independently.)
        int? listNumId = null;

        string AddMarkerParagraph(Dictionary<string, string> props)
        {
            if (listNumId is int reuse)
            {
                props["numId"] = reuse.ToString();
                props["numlevel"] = level;
            }
            else
            {
                props["liststyle"] = listStyle;
                props["level"] = level;
                // CommonMark: an ordered list starts at its first item's ordinal.
                // Pass start= only for a non-default start so the numId is minted
                // with the right <w:start>; a start of 1 is left implicit to
                // preserve cross-block list continuation (start= forces a mint).
                if (list.Ordered && list.Start != 1)
                    props["start"] = list.Start.ToString();
            }
            var p = Add(parentPath, "paragraph", posFor(), props);
            harvestDiagnostics(); // this Add() reset diagnostics — capture before the next
            record(p);
            if (listNumId == null
                && NavigateToElement(ParsePath(p)) is Paragraph para
                && para.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value is int nid)
                listNumId = nid;
            return p;
        }

        foreach (var item in list.Items)
        {
            // markerPlaced tracks whether the list number/bullet has been
            // attached to a paragraph for this item yet. A normal item places it
            // on its text paragraph; a pure-fence item (no inline text, only a
            // code block) places it on the FIRST code line instead — so the item
            // stays numbered without emitting an empty marker paragraph.
            bool markerPlaced = false;

            if (item.Inlines.Count > 0 || item.CodeBlocks.Count == 0)
            {
                var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["text"] = string.Concat(item.Inlines.Select(s => s.Text)),
                };
                var path = AddMarkerParagraph(props);
                if (item.Inlines.Any(s => s.Bold || s.Italic || s.Code || s.Strike))
                    pendingInline.Add((path, item.Inlines)); // formatted in phase 2 (perf)
                markerPlaced = true;
            }

            // Fenced code content of the item — each line its own Consolas
            // paragraph (monospace, indent already stripped by the parser). The
            // first line carries the list number when the item had no text; the
            // rest are plain (no numbering).
            foreach (var cb in item.CodeBlocks)
            {
                foreach (var codeLine in cb.Code.Split('\n'))
                {
                    var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["font"] = "Consolas",
                        ["text"] = codeLine,
                    };
                    if (!markerPlaced)
                    {
                        AddMarkerParagraph(props);
                        markerPlaced = true;
                    }
                    else
                    {
                        var path = Add(parentPath, "paragraph", posFor(), props);
                        harvestDiagnostics();
                        record(path);
                    }
                }
            }

            // All nested segments, in order — one item can own several
            // (marker switch / partial dedent), see MdListItem.Children.
            foreach (var child in item.Children)
                AddFlowList(parentPath, child, depth + 1, record, posFor, pendingInline, harvestDiagnostics);
        }
    }

    /// <summary>
    /// Count link/image references that the expander flattens to plain text
    /// (every span carries the source Href; the handler drops it — no
    /// w:hyperlink / picture). Counts maximal runs of consecutive same-Href
    /// spans, so one construct (even with split inline formatting) counts once.
    /// </summary>
    private static int CountFlattenedReferences(MarkdownDocument doc)
    {
        int n = 0;
        void Scan(List<MdSpan> spans)
        {
            string? prev = null;
            foreach (var s in spans)
            {
                if (s.Href != null && s.Href != prev) n++;
                prev = s.Href;
            }
        }
        void ScanList(MdList l)
        {
            foreach (var it in l.Items)
            {
                Scan(it.Inlines);
                foreach (var ch in it.Children) ScanList(ch);
            }
        }
        foreach (var b in doc.Blocks)
        {
            switch (b)
            {
                case MdParagraph p: Scan(p.Inlines); break;
                case MdHeading h: Scan(h.Inlines); break;
                case MdBlockQuote q: Scan(q.Inlines); break;
                case MdList l: ScanList(l); break;
                case MdTable t:
                    foreach (var cell in t.Header) Scan(cell);
                    foreach (var row in t.Rows)
                        foreach (var cell in row) Scan(cell);
                    break;
            }
        }
        return n;
    }

    private string AddFlowTable(string parentPath, MdTable t, InsertPosition? pos)
    {
        int cols = t.Header.Count;
        int rows = t.Rows.Count + 1;
        var tablePath = Add(parentPath, "table", pos,
            new(StringComparer.OrdinalIgnoreCase) { ["rows"] = rows.ToString(), ["cols"] = cols.ToString() });

        var grid = new List<List<List<MdSpan>>> { t.Header };
        grid.AddRange(t.Rows);

        // Populate cells by mutating the freshly-added table element DIRECTLY
        // rather than issuing a per-cell `Set("/body/tbl[N]/row/cell")`. Each of
        // those Set calls re-navigates from /body, and the first cell of every
        // table rebuilds the body child-index (O(body)) — cleared by the
        // heading/paragraph/list body-Adds between tables — so an M-table
        // interleave was O(M²). Navigating to the table once and writing cell
        // text in-place keeps table fill at O(cells-in-this-table), no body
        // navigation. Cell text here is plain (markdown table cells carry no
        // inline formatting in this expander), so a single run per cell matches
        // what the Set text path produced.
        // Use the element AddTable just created (captured in _lastAddedTable)
        // instead of re-navigating "/body/tbl[N]", which would rebuild the body
        // child-index (O(body)) once per table — the residual O(M²) in an
        // M-table interleave. This keeps table fill fully off the /body index.
        if (_lastAddedTable is Table tableEl)
        {
            var rowEls = tableEl.Elements<TableRow>().ToList();
            for (int r = 0; r < grid.Count && r < rowEls.Count; r++)
            {
                var row = grid[r];
                var cellEls = rowEls[r].Elements<TableCell>().ToList();
                for (int c = 0; c < cols && c < row.Count && c < cellEls.Count; c++)
                {
                    var text = string.Concat(row[c].Select(s => s.Text));
                    if (text.Length == 0) continue;
                    var cellPara = cellEls[c].GetFirstChild<Paragraph>()
                                   ?? cellEls[c].AppendChild(new Paragraph());
                    foreach (var run in cellPara.Elements<Run>().ToList()) run.Remove();
                    cellPara.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
                }
            }
        }
        return tablePath;
    }

    /// <summary>
    /// Rebuild a freshly-added paragraph's runs so each inline span becomes its
    /// own run with direct formatting (bold/italic/Consolas). Spans tile the
    /// paragraph text in order, so runs are CONSTRUCTED once instead of
    /// range-split after the fact — a paragraph with k formatted spans used to
    /// issue k `set range=` calls, each rescanning every already-split run:
    /// O(k²) per paragraph, ~2.5min for a 2000-span paragraph (blank-line-free
    /// LLM output gathers into exactly that). Direct construction is O(k).
    /// The base run's properties (e.g. a heading's bold+size) are cloned onto
    /// every span run, then span flags overlay them.
    /// </summary>
    private void ApplyInlineSpans(string paraPath, List<MdSpan> inlines)
    {
        if (NavigateToElement(ParsePath(paraPath)) is not Paragraph para) return;

        var baseProps = para.Elements<Run>().FirstOrDefault()?.RunProperties;
        foreach (var r in para.Elements<Run>().ToList()) r.Remove();

        foreach (var span in inlines)
        {
            if (span.Text.Length == 0) continue;
            var run = new Run();
            var rPr = baseProps != null ? (RunProperties)baseProps.CloneNode(true) : new RunProperties();
            if (span.Bold) rPr.Bold ??= new Bold();
            if (span.Italic) rPr.Italic ??= new Italic();
            if (span.Strike) rPr.Strike ??= new Strike();
            if (span.Code) rPr.RunFonts = new RunFonts { Ascii = "Consolas", HighAnsi = "Consolas" };
            if (rPr.HasChildren) run.AppendChild(rPr);

            // \t → <w:tab/>, mirroring AddParagraph's text conversion.
            var parts = span.Text.Split('\t');
            for (int p = 0; p < parts.Length; p++)
            {
                if (p > 0) run.AppendChild(new TabChar());
                if (parts[p].Length > 0)
                    run.AppendChild(new Text(parts[p]) { Space = SpaceProcessingModeValues.Preserve });
            }
            para.AppendChild(run);
        }
    }
}
