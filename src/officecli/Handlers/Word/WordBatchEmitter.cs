// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

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
public static partial class WordBatchEmitter
{
    /// <summary>
    /// Emit a batch sequence for a subtree of a Word document.
    /// <para>
    /// Path semantics: dump scopes purely to "what's under this path".
    /// `/` = whole document including all parts (styles, numbering, theme,
    /// settings, body, headers/footers, comments). A subtree path like
    /// `/body/p[5]` emits only that paragraph — styles/numbering/theme are
    /// NOT included because they live at sibling paths (`/styles`,
    /// `/numbering`, etc.), not under the requested subtree. References
    /// such as `style=Heading1` or `numId=3` are emitted as-is; replay
    /// onto a target document that already defines them works, otherwise
    /// the reference falls back to the target's defaults.
    /// </para>
    /// <para>
    /// Known limitations of subtree (non-`/`) dumps:
    /// — Footnote/endnote/chart references inside the emitted paragraph
    ///   resolve to the first N items in the source document's notes/charts,
    ///   not the original positions (cursors start at 0). Use `/` if the
    ///   subtree contains such references.
    /// — Image rels (rIds) reference the source package; the resource itself
    ///   is not bundled.
    /// </para>
    /// </summary>
    public static List<BatchItem> EmitWord(WordHandler word, string path)
    {
        if (string.IsNullOrEmpty(path))
            throw new CliException("dump path cannot be empty. Use '/' for the full document or a subtree path like /body/p[1].")
                { Code = "invalid_path" };
        if (path == "/") return EmitWord(word);

        var items = new List<BatchItem>();
        switch (path.ToLowerInvariant())
        {
            case "/theme": EmitThemeRaw(word, items); return items;
            case "/settings": EmitSettingsRaw(word, items); return items;
            case "/numbering": EmitNumberingRaw(word, items); return items;
            case "/styles": EmitStyles(word, items); return items;
            case "/body":
                EmitBody(word, items, new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase));
                return items;
        }

        // Reject bare /body/p and /body/tbl (no [N]). WordHandler.Get resolves
        // bare name segments to FirstOrDefault, which would silently dump the
        // first paragraph/table — almost never what the caller meant.
        var lastSeg = path.Substring(path.LastIndexOf('/') + 1);
        if (string.Equals(lastSeg, "p", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lastSeg, "tbl", StringComparison.OrdinalIgnoreCase))
        {
            throw new CliException(
                $"dump path not supported: {path} (missing index predicate). " +
                "Supported: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles")
            { Code = "unsupported_path" };
        }

        // Reject deep paths (e.g. /body/tbl[1]/tr[1]/tc[1]/p[1]). The dispatch
        // below assumes parent="/body" and would silently emit a wrongly
        // re-parented node. Supported subtree paths at this point are
        // /body/p[N] or /body/tbl[N] — exactly 2 segments below root.
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length > 2)
        {
            throw new CliException(
                $"dump path not supported: {path} (nested below /body). " +
                "Supported: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles")
            { Code = "unsupported_path" };
        }

        DocumentNode node;
        try { node = word.Get(path); }
        catch (Exception ex)
        {
            throw new CliException($"dump path not found: {path} ({ex.Message})") { Code = "path_not_found" };
        }

        if (node.Type != "paragraph" && node.Type != "p" && node.Type != "table")
        {
            throw new CliException(
                $"dump path not supported: {path} (type={node.Type}). " +
                "Supported: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles")
            { Code = "unsupported_path" };
        }

        var ctx = new BodyEmitContext(
            FootnoteTexts: word.Query("footnote").Select(n => n.Text ?? "").ToList(),
            EndnoteTexts: word.Query("endnote").Select(n => n.Text ?? "").ToList(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: word.Query("chart").Select(c =>
            {
                var full = word.Get(c.Path);
                return new ChartSpec(full.Format, full.Children ?? new List<DocumentNode>());
            }).ToList(),
            ChartCursor: new NoteCursor(),
            ParaIdToTargetIdx: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            DeferredBookmarks: new List<BatchItem>());

        if (node.Type == "table")
            EmitTable(word, path, 1, items, ctx);
        else
            EmitParagraph(word, path, "/body", 1, items, autoPresent: false, ctx);

        items.AddRange(ctx.DeferredBookmarks);
        return items;
    }

    /// <summary>Emit a batch sequence for a Word document (full document, equivalent to path "/").</summary>
    public static List<BatchItem> EmitWord(WordHandler word)
    {
        var items = new List<BatchItem>();

        // Phase order matters: resources first so body refs (style=Heading1,
        // numId=3, etc.) resolve when the paragraph adds reach them on replay.
        // Numbering must come BEFORE styles — list-style definitions
        // (Heading paragraphs with numPr) reference numId values, so style
        // adds that carry `numId=N` need /numbering to already hold N.
        EmitNumberingRaw(word, items);
        EmitStyles(word, items);
        EmitThemeRaw(word, items);
        EmitSettingsRaw(word, items);
        EmitSection(word, items);
        // Headers/footers run AFTER body: multi-section docs now emit
        // `add header parent="/section[N]"` (see EmitHeaderFooterPart), and
        // the /section[N] resolver only finds the carrier paragraph after
        // EmitBody has added it. Without body in place, every /section[N]
        // resolved to the body-level sectPr (the last section's), so
        // adding header type=default to two different sections collided
        // ("already exists in this section"). Body→header direction has
        // no replay-time dependency: header parts (PAGE/PAGEREF fields,
        // etc.) resolve their cross-refs at render time, not at batch-
        // apply time.
        var paraIdToTargetIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        EmitBody(word, items, paraIdToTargetIdx);
        EmitHeadersFooters(word, items);
        EmitComments(word, items, paraIdToTargetIdx);
        // CONSISTENCY(markRPr-inherit-opt-out): dump emits each run's props
        // verbatim from the source; we never want AddRun's UX-convenience
        // markRPr→rPr type-fill to add a w:rFonts (or any other) child the
        // source never had. Stamp the opt-out on every emitted `add r` once
        // here, instead of threading it through five EmitParagraph branches.
        foreach (var it in items)
        {
            if (it.Command == "add" && (it.Type == "r" || it.Type == "run"))
            {
                it.Props ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (!it.Props.ContainsKey("noMarkRPrInherit") && !it.Props.ContainsKey("nomarkrprinherit"))
                    it.Props["noMarkRPrInherit"] = "true";
            }
        }
        return items;
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
        // BUG-DUMP11-01: chapter-numbering attributes on w:pgNumType.
        "chapStyle", "chapSep",
        "titlePage", "direction", "rtlGutter",
        // BUG-DUMP11-03: <w:noEndnote/> section flag.
        "noEndnote",
        "lineNumbers", "lineNumberCountBy",
        // BUG-DUMP11-02: lnNumType/@w:start (first line number when counting).
        "lineNumberStart",
        // Multi-column section layout. Get exposes these as canonical keys
        // (columns, columnSpace, columns.equalWidth) and Set's case table
        // accepts all three (WordHandler.Set.SectionLayout.cs). Without them
        // here, multi-column documents silently revert to single column on
        // round-trip.
        "columns", "columnSpace",
        // Document-level final-section break type (oddPage / evenPage /
        // continuous). Set / accepts section.type but the canonical Get
        // surfaces it bare; emit so the trailing sectPr's type survives.
        "section.type",
        // Document protection
        "protection", "protectionEnforced",
        // BUG-DUMP10-03: document-level page background color
        // (<w:document><w:background w:color="…"/>). Set already accepts
        // this canonical key (WordHandler.Add.cs:565); without inclusion
        // here, dump silently dropped the page background on round-trip.
        "background",
        // Document grid (CJK-aware line layout)
        "charSpacingControl",
        // pPrDefault CJK toggles — without these, Word inserts an automatic
        // space between Latin runs and adjacent CJK glyphs ("2025年" →
        // "2025 年"). Templates that explicitly disable autoSpaceDE/DN
        // depend on these surviving the round-trip.
        "kinsoku", "overflowPunct", "autoSpaceDE", "autoSpaceDN",
    };

    // Dotted-prefix groups that round-trip wholesale via `set /`. Each
    // sub-key is forwarded as-is; the schema-reflection layer routes the
    // dotted path into the right OOXML target.
    private static readonly string[] RootPrefixGroups = new[]
    {
        "docDefaults.",
        "docGrid.",
        // columns.equalWidth / columns.separator etc. roundtrip via the
        // canonical dotted form Get already emits.
        "columns.",
    };

    // Captured once per process: blank doc's `Get("/")` root Format, normalized
    // to string values. Used by EmitSection to skip keys whose source value
    // matches what BlankDocCreator stamps — those keys would otherwise leak
    // from blank into the replay target and re-appear on the next dump,
    // breaking dump-then-replay-then-dump idempotency.
    private static readonly Lazy<IReadOnlyDictionary<string, string>> _blankRootBaseline =
        new(ComputeBlankRootBaseline);

    private static IReadOnlyDictionary<string, string> ComputeBlankRootBaseline()
    {
        var tempPath = Path.Combine(Path.GetTempPath(),
            $"officecli_blank_baseline_{Guid.NewGuid():N}.docx");
        try
        {
            OfficeCli.BlankDocCreator.Create(tempPath);
            using var handler = new OfficeCli.Handlers.WordHandler(tempPath, editable: false);
            var root = handler.Get("/");
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in root.Format)
            {
                if (v == null) continue;
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) result[k] = s;
            }
            return result;
        }
        catch
        {
            // If baseline computation fails (test harness with no temp path
            // access, etc.), fall back to an empty baseline. EmitSection then
            // behaves as it did before this change.
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            try { File.Delete(tempPath); } catch { }
        }
    }

    private sealed record ChartSpec(Dictionary<string, object?> Format, IReadOnlyList<DocumentNode> Series);

    private sealed record BodyEmitContext(
        List<string> FootnoteTexts,
        List<string> EndnoteTexts,
        NoteCursor FootnoteCursor,
        NoteCursor EndnoteCursor,
        List<ChartSpec> ChartSpecs,
        NoteCursor ChartCursor,
        Dictionary<string, int>? ParaIdToTargetIdx,
        // BUG-DUMP10-04: cross-paragraph bookmarks (endPara > 0) need to be
        // emitted *after* every host paragraph already exists on replay,
        // because AddBookmark relocates the BookmarkEnd to siblings[N+endPara]
        // and that sibling does not exist yet during the in-order walk.
        // EmitParagraph stashes the deferred `add bookmark` rows here;
        // EmitBody appends them once all paragraphs are emitted.
        List<BatchItem> DeferredBookmarks);

    private static void EmitBody(WordHandler word, List<BatchItem> items,
                                 Dictionary<string, int>? paraIdToTargetIdx = null)
    {
        // BUG-DUMP-X6-02: word.Get("/body") raises "Path not found: /body" on
        // a zip lacking word/document.xml. Surface a CliException pointing at
        // the file rather than leaking an internal path the user never asked
        // for (common when dumping "/" on a corrupt or non-Word zip).
        DocumentNode bodyNode;
        try
        {
            bodyNode = word.Get("/body");
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw new CliException(
                "dump failed: word/document.xml is missing — the file may not be a valid Word document")
                { Code = "invalid_document" };
        }
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
            ParaIdToTargetIdx: paraIdToTargetIdx,
            DeferredBookmarks: new List<BatchItem>());

        int pIndex = 0, tblIndex = 0;
        foreach (var child in bodyNode.Children)
        {
            switch (child.Type)
            {
                case "paragraph":
                case "p":
                    // BUG-X4-FUZZ-1: display-mode equations surface in
                    // bodyNode.Children as type="paragraph" but the path
                    // resolver addresses them as /body/oMathPara[N], NOT as
                    // /body/p[N]. Incrementing pIndex for them would offset
                    // every subsequent inline-child path (hyperlink/footnote/
                    // run) by +1 per preceding equation, breaking round-trip.
                    // Detect the wrapper via path and route to EmitParagraph
                    // without bumping pIndex — EmitParagraph's equation branch
                    // re-emits the equation as `add /body --type equation`.
                    if (child.Path.Contains("/oMathPara[", StringComparison.OrdinalIgnoreCase))
                    {
                        EmitParagraph(word, child.Path, "/body", pIndex + 1, items, autoPresent: false, ctx);
                    }
                    else
                    {
                        pIndex++;
                        EmitParagraph(word, child.Path, "/body", pIndex, items, autoPresent: false, ctx);
                    }
                    break;
                case "table":
                    tblIndex++;
                    EmitTable(word, child.Path, tblIndex, items, ctx);
                    break;
                case "section":
                case "sectPr":
                    // The body always carries one trailing sectPr that the
                    // blank document already provides; for v0.5 we rely on
                    // that default and skip emitting section properties.
                    // Section emit is a follow-up.
                    break;
                case "sdt":
                    EmitSdt(word, child.Path, items);
                    break;
                case "bookmark":
                    // Standalone body-level <w:bookmarkStart> (e.g. an anchor
                    // added with `add /body --type bookmark`). Inline bookmarks
                    // inside paragraphs are handled by EmitParagraph; without
                    // this case, body-level bookmark anchors were silently
                    // dropped on dump.
                    {
                        var bmFull = word.Get(child.Path);
                        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        if (bmFull.Format.TryGetValue("name", out var nm)
                            && nm != null && !string.IsNullOrEmpty(nm.ToString()))
                            bmProps["name"] = nm.ToString()!;
                        else
                            break; // BookmarkStart with no name is unusable
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = "/body",
                            Type = "bookmark",
                            Props = bmProps
                        });
                    }
                    break;
                case "bookmarkEnd":
                    // Paired with the body-level bookmarkStart emit above. The
                    // matching `add bookmark` re-creates start+end together so
                    // the standalone end node needs no emit.
                    break;
                case "equation":
                    // BUG-DUMP13-03: a bare <m:oMathPara> direct child of
                    // <w:body> (not wrapped in a w:p) surfaces in
                    // bodyNode.Children as type="equation". Without this case
                    // it fell to `default: break` and was silently dropped.
                    // Mirror the EmitParagraph equation branch shape.
                    {
                        var eqFull = word.Get(child.Path);
                        var mode = eqFull.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
                        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode
                        };
                        if (!string.IsNullOrEmpty(eqFull.Text))
                            eqProps["formula"] = eqFull.Text!;
                        // BUG-DUMP19-02: forward block-equation alignment.
                        if (eqFull.Format.TryGetValue("align", out var eqAlign)
                            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
                            eqProps["align"] = eqAlign.ToString()!;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = "/body",
                            Type = "equation",
                            Props = eqProps
                        });
                    }
                    break;
                default:
                    // Unknown body-level child types — skip for v0.5.
                    break;
            }
        }

        // BUG-DUMP10-04: flush deferred cross-paragraph bookmark rows. They
        // are emitted last so AddBookmark sees the full sibling list when
        // walking forward to the BookmarkEnd's target paragraph.
        items.AddRange(ctx.DeferredBookmarks);
    }
}
