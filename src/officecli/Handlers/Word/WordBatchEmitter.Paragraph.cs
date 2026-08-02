// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    /// <summary>
    /// Emit a paragraph at the target index under <paramref name="parentPath"/>.
    /// When <paramref name="autoPresent"/> is true, the parent already has a
    /// pre-existing paragraph at that index (e.g. an auto-created table cell
    /// paragraph); we issue a `set` instead of a fresh `add` so the existing
    /// paragraph gets reused rather than duplicated.
    /// </summary>
    // BUG-DUMP26-01 / BUG-DUMP-SECTNUM: a paragraph's numbering props must never
    // ride on an `add p` / `set p` as ad-hoc numbering. (1) numId/numLevel that came
    // from style inheritance (ResolveNumPrFromStyle, no direct w:numPr) must be
    // dropped — the style already supplies them and emitting them would promote
    // inherited→explicit on replay. (2) When a direct numId is present, the
    // abstractNum/num pair is already in /numbering (raw-set wholesale by
    // EmitNumberingRaw); forwarding numFmt/listStyle/start to AddParagraph triggers
    // ad-hoc numbering-definition creation — Word allocates a FRESH numId, orphaning
    // the original abstract numbering's level rPr (color/bold/custom marker). Drop
    // those so the paragraph just attaches by numId+numLevel to the existing def.
    // Applied by BOTH the normal paragraph emit AND the section-carrier paragraph
    // `set` (TryEmitInlineSectionBreak), which builds its pPr props independently.
    private static void ApplyNumberingInheritanceFilters(IDictionary<string, string> props, DocumentNode pNode)
    {
        bool numInherited = pNode.Format.TryGetValue("numInherited", out var niVal)
            && string.Equals(niVal?.ToString(), "true", StringComparison.OrdinalIgnoreCase);
        if (numInherited)
        {
            props.Remove("numId");
            props.Remove("numLevel");
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        if (props.ContainsKey("numId"))
        {
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
    }

    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent,
                                      BodyEmitContext? ctx = null)
    {
        var pNode = word.Get(sourcePath);

        if (TryEmitDisplayEquation(word, pNode, parentPath, autoPresent, items)) return;

        // Track source paraId -> target index BEFORE any early-return path
        // (section break, TOC, …). Comments anchored on a section-break or
        // TOC paragraph would otherwise miss the mapping and fall back to
        // /body/p[1], silently retargeting the comment.
        if (ctx?.ParaIdToTargetIdx != null && parentPath == "/body" &&
            pNode.Format.TryGetValue("paraId", out var earlyParaId) && earlyParaId != null)
        {
            ctx.ParaIdToTargetIdx[earlyParaId.ToString()!] = targetIndex;
        }

        if (TryEmitInlineSectionBreak(word, pNode, parentPath, items, ctx)) return;
        if (TryEmitTocParagraph(pNode, parentPath, items)) return;
        if (TryEmitTextboxOnlyParagraph(word, pNode, parentPath, autoPresent, items, ctx)) return;

        var props = FilterEmittableProps(pNode.Format);
        // BUG-DUMP-R44-6: paraMarkIns.* must round-trip as a MARK-ONLY tracked
        // insertion — <w:pPr><w:rPr><w:ins/></w:rPr></w:pPr> on the pilcrow
        // ALONE — never as a content insertion that wraps the (plain) run text.
        // The former path here rewrote paraMarkIns.* into a bare revision.author
        // (no revision.type); AddParagraph's bare-attribution branch treats that
        // as "this whole paragraph was inserted" and wraps every auto-created
        // <w:r> in <w:ins>, promoting plain run text to a tracked insertion it
        // never was (Reject Changes would then delete text the source keeps).
        // Pass the paraMarkIns.* keys through verbatim instead — AddParagraph's
        // dedicated paraMarkIns block stamps the mark rPr only, exactly mirroring
        // the paraMarkDel.* handling just below. A genuine run/paragraph content
        // insertion still arrives as revision.type=ins on the run/paragraph and
        // wraps correctly (unaffected by this branch).
        // (No remove here — let paraMarkIns.* pass through to AddParagraph.)
        // BUG-DUMP-R43-8: pPrChange's PreviousParagraphProperties snapshot now
        // round-trips verbatim via revision.beforeXml (set by the pPrChange
        // readback in Navigation.cs and consumed by AddParagraph). The former
        // revision.beforeLost warn-and-drop path is retired — the prior-pPr
        // payload is preserved, not lost. (A defensive strip remains in case a
        // stale dump still carries the legacy key.)
        props.Remove("revision.beforeLost", out var _);
        // paraMarkDel.* — surfaces the dump path through AddParagraph's
        // paraMarkDel block (added alongside this readback). Pass the keys
        // through unchanged; AddParagraph allowlists the prefix and consumes
        // them at end-of-function. Don't fold into revision.* — that
        // namespace already routes to ins/format paths and a paragraph can
        // legitimately carry BOTH paraMarkDel (¶ join) and revision.type=
        // format (pPrChange).
        // (No remove here — let the props pass through to AddParagraph.)
        // BUG-DUMP26-01: numId/numLevel that came from style inheritance
        // (ResolveNumPrFromStyle, no direct w:numPr on the paragraph) must
        // not ride on `add p` — the style already supplies them, and emitting
        // them would semantically promote inherited→explicit on replay.
        // Mirrors the first-run hoist precedent for run-character props
        // inherited from styles.
        ApplyNumberingInheritanceFilters(props, pNode);
        // BUG-R4F-02: a paragraph may carry a numId that does not resolve to any
        // <w:num> in /numbering (dangling reference). This is valid OOXML — Word
        // renders the paragraph, just without a list marker — but the Add-side
        // dangling-numId guard (WordHandler.Add.Text.cs) rejects it, and because
        // `add p` is atomic the whole paragraph (TEXT included) is lost on replay.
        // Drop the numbering props so the `add p` succeeds with its text, and
        // surface a warning so the dropped numbering is visible. The numbering is
        // kept whenever the numId IS defined (the common case). Mirrors the
        // out-of-window font-size / zero-width-column clamps in Filters.cs.
        if (props.TryGetValue("numId", out var numIdStr)
            && int.TryParse(numIdStr, out var numIdInt)
            && numIdInt > 0
            && !word.IsNumIdDefined(numIdInt))
        {
            props.Remove("numId");
            props.Remove("numLevel");
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "numId",
                Path: pNode.Path,
                Reason: $"paragraph references numId={numIdInt} which is not defined in /numbering (dangling reference); the numbering was dropped so the paragraph text survives dump→batch round-trip"));
        }
        // Collapse non-TOC field chains (fldChar(begin) + instrText(" PAGE ")
        // + fldChar(separate) + display run(s) + fldChar(end)) into a single
        // synthetic "field" entry. Without this collapse, the subsequent
        // `runs` filter sees only the cached display run and emits the field
        // value as static text — PAGE/REF/SEQ/HYPERLINK/NUMPAGES degrade to
        // their evaluated string and stop auto-updating (BUG-X2-05 / X2-1).
        var fieldEntries = CollapseFieldChains(pNode.Children ?? new List<DocumentNode>(), word);
        // R14-bug1+2: a legacy form field MAY embed a BookmarkStart/End of its
        // own name (Word wraps form fields in a bookmark so REF fields can target
        // them, but a plain FORMCHECKBOX/FORMTEXT authored without that wrap has
        // NONE). AddFormField recreates the wrapping bookmark internally — if the
        // emit pipeline also drops an `add bookmark name=X` row before the
        // `add formfield name=X`, AddFormField throws on the duplicate. Filter
        // bookmarks whose name matches a sibling formfield synth's ffName.
        var formFieldNames = fieldEntries
            .Where(e => e.Type == "formfield" && e.Format.TryGetValue("ffName", out _))
            .Select(e => e.Format["ffName"]?.ToString() ?? "")
            .Where(n => !string.IsNullOrEmpty(n))
            .ToHashSet(StringComparer.Ordinal);
        // Gate on ANY form field, not only named ones: a paragraph holding only
        // nameless fields still needs the noBookmark pin pass below, else each
        // nameless field gains a fabricated ff_<guid> bookmark on rebuild
        // (BUG-DUMP-R72-FF-BOOKMARK-COUNT).
        if (fieldEntries.Any(e => e.Type == "formfield"))
        {
            // BUG-DUMP-FFCHECKBOX-BOOKMARK: a form field whose SOURCE had no
            // wrapping bookmark must NOT gain a fabricated one on rebuild.
            // AddFormField wraps every field in a <w:bookmarkStart name=ffName>
            // unconditionally; a 54-checkbox grid with no source bookmarks then
            // gained 54 fabricated Check1/Check1_N bookmarks (and a uniquify
            // pass), which alters the checkbox cells' content and nudges row
            // heights → table reflow → page drift. Mark each formfield synth
            // with whether a matching bookmark actually sits among its siblings;
            // TryEmitFormFieldRun forwards a `noBookmark` pin to AddFormField so
            // a bookmark-less source stays bookmark-less. (A field whose source
            // HAS the bookmark keeps the existing behaviour — the bookmark sibling
            // is filtered below and AddFormField recreates it.)
            var bookmarkNamesPresent = fieldEntries
                .Where(e => e.Type == "bookmark"
                    && e.Format.TryGetValue("name", out var bnm) && bnm != null)
                .Select(e => e.Format["name"]!.ToString() ?? "")
                .ToHashSet(StringComparer.Ordinal);
            // BUG-DUMP-FF-ROWLEVEL-BOOKMARK / BUG-DUMP-R72-FF-BOOKMARK-COUNT: a
            // form field's wrapping bookmark may sit at ROW level (a <w:tr> child
            // between cells) — invisible to the same-paragraph set, and dropped by
            // the table emitter — so pinning noBookmark purely on the same-paragraph
            // check would erase every row-level bookmark. The earlier fix consulted
            // a document-wide NAME SET ("does any bookmark with this name exist?"),
            // but that over-fires when many fields share one name: a doc with ONE
            // <w:bookmarkStart name="Check1"> and 26 checkbox fields all named
            // "Check1" then recreated 26 Check1 bookmarks (+a uniquify cascade).
            // Use a count-aware BUDGET instead: each name may hand out only as many
            // wrapping bookmarks as the source actually had. A same-paragraph match
            // is a real bookmark, so it always recreates AND reserves one budget
            // unit; a field with no same-paragraph bookmark keeps one only while the
            // remaining budget (row-level / other-paragraph source bookmarks) lasts;
            // an unnamed field — which cannot carry a named bookmark — and a field
            // whose budget is exhausted are pinned noBookmark.
            foreach (var ffSynth in fieldEntries.Where(e => e.Type == "formfield"))
            {
                var ffn = ffSynth.Format.TryGetValue("ffName", out var ffnObj)
                    ? (ffnObj?.ToString() ?? "")
                    : "";
                if (string.IsNullOrEmpty(ffn))
                {
                    // A nameless source field had no wrapping bookmark (a bookmark
                    // needs a name), yet AddFormField would auto-generate an
                    // ff_<guid> name + bookmark for it (the interactive default).
                    // Pin noBookmark on round-trip so a bookmark-less field stays
                    // bookmark-less instead of gaining a fabricated ff_<guid> one.
                    ffSynth.Format["_noBookmark"] = true;
                    continue;
                }
                if (bookmarkNamesPresent.Contains(ffn))
                {
                    // Real same-paragraph wrapping bookmark: always recreate, but
                    // reserve its budget so a later same-named field can't reuse it.
                    ctx?.ConsumeBookmarkBudget(word, ffn);
                    continue;
                }
                if (ctx == null || !ctx.ConsumeBookmarkBudget(word, ffn))
                    ffSynth.Format["_noBookmark"] = true;
            }
            if (ctx != null)
                foreach (var ffn in formFieldNames) ctx.FormFieldBookmarkNames.Add(ffn);
            fieldEntries = fieldEntries
                .Where(e => !((e.Type == "bookmark" || e.Type == "bookmarkEnd")
                    && e.Format.TryGetValue("name", out var bn)
                    && bn != null
                    && formFieldNames.Contains(bn.ToString() ?? "")))
                .ToList();
        }
        // BUG-DUMP5-01/02: include break-typed children in the same ordered
        // list as runs so document-order is preserved on emit.
        var runs = fieldEntries
            .Where(c => c.Type == "run" || c.Type == "r" || c.Type == "picture" || c.Type == "field" || c.Type == "formfield" || c.Type == "ptab" || c.Type == "break"
                || c.Type == "equation"
                || c.Type == "tab"
                || c.Type == "bookmark"
                || c.Type == "bookmarkEnd"
                // BUG-DUMP-PERM: ranged editing-permission markers are
                // positioned paragraph children — keep them in the ordered run
                // list so TryEmitPermRun replays them at their source offset.
                || c.Type == "permStart"
                || c.Type == "permEnd"
                // BUG-DUMP-RUBY: ruby (phonetic guide) child surfaces the
                // verbatim <w:r><w:ruby> XML for a raw-set append.
                || c.Type == "ruby"
                // BUG-DUMP-R42-9: bdo (bidirectional override) child surfaces the
                // verbatim <w:bdo> wrapper XML for a raw-set append.
                || c.Type == "bdo"
                // BUG-DUMP-R43-7: dir (bidirectional embedding) child surfaces the
                // verbatim <w:dir> wrapper XML for a raw-set append.
                || c.Type == "dir"
                // R10-bug1: include ole children so TryEmitOleRun can fire
                // a warning instead of letting them be silently filtered
                // out of the run list (full round-trip is a backlog item).
                || c.Type == "ole")
            .ToList();
        var breaks = runs.Where(c => c.Type == "break").ToList();
        var bookmarks = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "bookmark"
                && !(c.Format.TryGetValue("name", out var bn) && bn != null
                    && formFieldNames.Contains(bn.ToString() ?? "")))
            .ToList();
        var inlineSdts = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "sdt")
            .ToList();

        bool collapseSingleRun = ShouldCollapseSingleRun(word, runs, breaks.Count, bookmarks.Count, inlineSdts.Count);
        pNode.Format.TryGetValue("tabs", out var pTabs);

        if (collapseSingleRun)
        {
            if (runs.Count == 1)
            {
                // BUG-DUMP-R35-2: a wrapper-flattened run that collapses into the
                // paragraph's own `text` prop bypasses EmitPlainOrHyperlinkRun, so
                // the deterministic smartTag/customXml flatten warning never fired
                // for a single-run wrapped paragraph (text survived, loss silent).
                // CONSISTENCY(wrapper-flatten-warning): same emit as
                // EmitPlainOrHyperlinkRun's _wrapperFlattened branch.
                WarnWrapperFlattened(runs[0], ctx);
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
                if (props.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{parentPath}/p[last()]",
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
            EmitTabStops($"{parentPath}/p[last()]", pTabs, items);
            return;
        }

        // Multi-run paragraph: emit the paragraph empty first, then add each
        // run as an explicit child. See BUG-DUMP-HOIST in
        // StripRunCharacterPropsFromParagraph — for multi-run paragraphs the
        // firstRun hoist would re-apply formatting to every sibling on
        // replay, so strip run-level keys before emit.
        //
        // BUG-DUMP-R42-4: but a RUN-LESS paragraph that nonetheless reaches this
        // path (e.g. it carries a bookmark marker, which ShouldCollapseSingleRun
        // keeps off the collapse path so TryEmitBookmarkRun replays the marker)
        // has NO source run to hoist from — Navigation's firstRun-fallback read
        // those bare size/size.cs/bold.cs/font.* keys straight off the paragraph
        // MARK rPr (the ¶ glyph's formatting), not off a run. Stripping them here
        // drops the markRPr entirely, leaving `add p {}` and collapsing a
        // cover-page paragraph's large-size ¶ on rebuild. Only strip when a real
        // text/format-bearing run exists (the genuine hoist source); a run-less
        // paragraph keeps its bare markRPr keys, same as the non-bookmark empty
        // paragraph that rides the collapse path with full markRPr.
        // BUG-DUMP-R26: a field chain swallows the paragraph's text runs in
        // CollapseFieldChains, so a field-result paragraph has NO run-typed
        // children left — yet the paragraph node's bare character keys were
        // harvested (firstRun-fallback) from the field's RESULT runs, and the
        // field emit (raw-set verbatim / add field) replays that formatting
        // itself. Leaving the harvested keys on `add p` duplicates them onto
        // the ¶ mark on rebuild (<w:b/> count 2). Field entries are therefore
        // format-bearing hoist sources too.
        // BUG-DUMP-MARKSZ-DEL: a paragraph whose only runs are tracked-revision
        // runs (<w:del>/<w:ins>/<w:moveFrom>/<w:moveTo>) has NO direct hoist
        // source — Navigation's firstRun (para.Elements<Run>()) skips revision-
        // wrapped runs, so the paragraph's bare size/size.cs/font.* keys were
        // read off the ¶-mark rPr, not off a run. Treating the del/ins run as
        // the hoist source and stripping here dropped the mark's font size, so
        // a deleted-content table cell collapsed to default line height on
        // rebuild — pushing every later row down (cumulative drift, +1 page).
        // Only a non-revision run is a genuine hoist source.
        static bool IsRevisionWrappedRun(DocumentNode c) =>
            c.Format.TryGetValue("revision.type", out var rvt)
            && rvt?.ToString() is "del" or "ins" or "moveFrom" or "moveTo";
        bool hasFormatBearingRun = runs.Any(c =>
            (c.Type == "run" || c.Type == "r" || c.Type == "field")
            && !IsRevisionWrappedRun(c));
        if (hasFormatBearingRun)
            StripRunCharacterPropsFromParagraph(props);
        if (autoPresent)
        {
            if (props.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{parentPath}/p[last()]",
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

        var paraTargetPath = $"{parentPath}/p[last()]";
        EmitTabStops(paraTargetPath, pTabs, items);

        // BUG-DUMP4-06 / BUG-R12B(BUG1): emit each inline SdtRun child AT its
        // real intra-paragraph position interleaved with the runs, not hoisted
        // ahead of them. Navigation appends every SdtRun at the tail of
        // pNode.Children (it sits outside the DOM-ordered run/bookmark/eq merge),
        // so emitting all SDTs before the runs loop scrambled document order: a
        // content control sitting between two runs ("Video [sdt] a powerful…")
        // came back as "[sdt] Video a powerful…". Recover each child's true
        // document rank from the source paragraph XML (top-level child order),
        // then flush each SDT just before the first run whose rank exceeds it.
        var childDocOrder = ComputeParagraphChildDocOrder(word, pNode.Path);
        // Local SDT emit (formerly the standalone foreach above). Returns after
        // appending the appropriate `add sdt` / rich raw-set op to `items`.
        void EmitInlineSdt(DocumentNode sdt)
        {
            // BUG-R12A(BUG1): an inline/run-level <w:sdt> whose content carries
            // more than one run OR any run-level rPr (bold/color/size/…) cannot
            // round-trip through the flat `add sdt text=` path — AddSdt seeds a
            // single unformatted run from `text`, so "FIRSTSECOND" comes back as
            // one plain run and the bold+red is lost. Mirror the R11 block-SDT
            // fix: raw-set the <w:sdt> verbatim into the just-emitted host
            // paragraph (a run-level SDT has no inner <w:p>, so the rich-BLOCK
            // detector doesn't apply; use a run-level richness check). Restricted
            // to /body hosts + no external rels (same constraints as the inline
            // textbox raw-set) so dangling r:id/r:embed can't be produced; other
            // hosts fall back to the flat text emit.
            // BUG-DUMP-R26-7: rich/nested inline SDTs now round-trip verbatim in
            // header/footer/cell hosts too (ResolveRawSetHost), not only /body.
            if (TryEmitRichInlineSdt(word, sdt, parentPath, items, ctx))
                return;
            var sdtProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // BUG-DUMP-SDTPROPS: forward the form-control sdtPr children the typed
            // emit previously dropped (lock / placeholder / date-picker value /
            // combo+dropdown selection). Same whitelist as the block-SDT path
            // (EmitSdtTyped) so inline and block controls round-trip identically.
            foreach (var key in SdtTypedEmitKeys)
            {
                if (sdt.Format.TryGetValue(key, out var v) && v != null)
                {
                    var s = NormalizeSdtEmitValue(v);
                    if (s.Length > 0) sdtProps[key] = s;
                }
            }
            if (!string.IsNullOrEmpty(sdt.Text))
                sdtProps["text"] = sdt.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "sdt",
                Props = sdtProps
            });
        }

        // BUG-DUMP6-05: collapse N runs that share a hyperlink wrapper into
        // one synthetic hyperlink-typed entry — see CoalesceHyperlinkRuns.
        runs = CoalesceHyperlinkRuns(runs);
        // BUG-D1-MULTIDRAWING-HOST: when this paragraph hosts ≥2 drawing-
        // bearing runs (side-by-side card layout), every textbox must attach
        // to the SAME host paragraph just emitted by the `add p` above so the
        // side-by-side relationship survives round-trip. A single drawing
        // either reached the wrapper-coalesce shortcut (children.Count == 1)
        // OR shares its source paragraph with sibling text/runs — in the
        // latter case attaching to the host paragraph is still correct
        // (preserves the inline relationship that the source had).
        int drawingBearingCount = runs.Count(r =>
        {
            if (r.Type == "picture") return true;
            if (r.Type != "run" && r.Type != "r") return false;
            var probe = word.GetElementXml(r.Path);
            return !string.IsNullOrEmpty(probe) && IsTextboxDrawing(probe);
        });
        // BUG-DUMP-R25-7: a header/footer paragraph that hosts a textbox drawing
        // (e.g. the centred page-number box in a footer) must keep the drawing
        // run INSIDE its styled host paragraph. The body-only
        // TryEmitTextboxOnlyParagraph shortcut does not fire for header/footer
        // hosts, so the textbox otherwise fell through to AddTextbox creating a
        // NEW unstyled paragraph — splitting the footer into (a) an empty
        // pStyle-carrying paragraph and (b) an unstyled drawing paragraph that
        // reverted to Normal's taller exact line height, growing the reserved
        // footer area and reflowing the body (+1 rendered page). Attaching to
        // the just-emitted paraTargetPath (which carries the pPr/pStyle/ind/jc)
        // keeps the drawing in its styled host. Restrict to non-body hosts —
        // /body single-drawing paragraphs already round-trip via the
        // wrapper-coalesce shortcut, and the ≥2 side-by-side case is unchanged.
        string? sharedAttachPara =
            (drawingBearingCount >= 2
             || (drawingBearingCount == 1 && parentPath != "/body"
                 && IsHeaderFooterHost(parentPath)))
            ? paraTargetPath : null;
        // BUG-R14B: hyperlink rows already emitted at this parent belong to
        // earlier paragraphs (paraTargetPath is the same "/body/p[last()]"
        // literal for every <w:p>). Capture that count so multi-run hyperlinks
        // in THIS paragraph re-index from 1 — see EmitStructuredHyperlink.
        int hlBaseline = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        // BUG-R12B(BUG1): inline SDTs sorted by their source document rank,
        // each flushed just before the first run that sits after it. A rank of
        // int.MaxValue (no XML position recovered) falls through to the post-loop
        // tail flush — same behavior as the old hoist-to-end ordering, so a
        // paragraph whose XML can't be probed degrades to the prior shape rather
        // than dropping the SDT.
        var pendingSdts = inlineSdts
            .Select(s => (sdt: s, rank: ChildDocRank(childDocOrder, s.Path)))
            .OrderBy(t => t.rank)
            .ToList();
        int sdtCursor = 0;
        foreach (var run in runs)
        {
            int runRank = ChildDocRank(childDocOrder, run.Path);
            while (sdtCursor < pendingSdts.Count && pendingSdts[sdtCursor].rank < runRank)
            {
                EmitInlineSdt(pendingSdts[sdtCursor].sdt);
                sdtCursor++;
            }
            if (TryEmitBookmarkRun(run, paraTargetPath, items, ctx)) continue;
            if (TryEmitPermRun(run, paraTargetPath, items)) continue;
            if (TryEmitPgNumRun(word, run, parentPath, items, ctx)) continue;
            if (TryEmitDateFieldRun(word, run, parentPath, items, ctx)) continue;
            if (TryEmitHyphenRun(word, run, parentPath, paraTargetPath, items, ctx, hlBaseline)) continue;
            if (TryEmitRubyRun(run, parentPath, paraTargetPath, items, ctx)) continue;
            if (TryEmitBdoRun(run, parentPath, items, ctx)) continue;
            if (TryEmitDirRun(run, parentPath, items, ctx)) continue;
            if (TryEmitBreakRun(word, run, parentPath, paraTargetPath, items, ctx)) continue;
            if (TryEmitTabRun(run, paraTargetPath, items)) continue;
            if (TryEmitPtabRun(run, paraTargetPath, items)) continue;
            if (TryEmitEquationRun(word, run, paraTargetPath, items)) continue;
            if (TryEmitFormFieldRun(run, paraTargetPath, items)) continue;
            if (TryEmitFieldRun(word, run, paraTargetPath, parentPath, items, ctx)) continue;
            // OLE/embedded-object runs surface as type="ole" (see CreateOleNode
            // in WordHandler.ImageHelpers.cs). TryEmitOleRun base64-inlines the
            // embedded payload + icon and the VML frame metadata into a
            // self-contained `add ole` (picture-run style), so the object
            // round-trips with no external file; it warns only when the payload
            // can't be resolved.
            if (TryEmitOleRun(run, paraTargetPath, items, ctx, word)) continue;
            if (TryEmitPictureRun(word, run, paraTargetPath, parentPath, targetIndex, items, ctx, sharedAttachPara)) continue;
            if (TryEmitNoteRefRun(word, run, paraTargetPath, items, ctx)) continue;
            if (TryEmitMixedBreakRun(word, run, parentPath, paraTargetPath, items, ctx)) continue;
            EmitPlainOrHyperlinkRun(word, run, paraTargetPath, items, ctx, hlBaseline);
        }
        // Flush any SDTs that sit after the last run (or whose rank could not be
        // recovered from the XML — int.MaxValue lands here).
        for (; sdtCursor < pendingSdts.Count; sdtCursor++)
            EmitInlineSdt(pendingSdts[sdtCursor].sdt);
    }

    // BUG-R12B(BUG1): recover the document-order rank of each top-level
    // paragraph child (run / sdt / bookmark / …) from the source paragraph XML.
    // Navigation surfaces inline SdtRun children at the tail of pNode.Children
    // (outside the DOM-ordered run merge), so the only place the true
    // interleaving survives is the OOXML element order itself. Returns a map
    // from positional child segment ("r[2]", "sdt[1]", "bookmark[1]") to a
    // 0-based document rank. Depth-tracked so a <w:r> nested INSIDE a <w:sdt>
    // (the SDT's own content run) is not counted as a sibling.
    private static Dictionary<string, int> ComputeParagraphChildDocOrder(WordHandler word, string? paragraphPath)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(paragraphPath)) return map;
        string xml;
        try { xml = word.GetElementXml(paragraphPath) ?? ""; }
        catch { return map; }
        if (string.IsNullOrEmpty(xml)) return map;
        // Strip the leading <w:p …> open and trailing </w:p> so we walk only the
        // paragraph's content; track nesting depth to keep to top-level children.
        var perNameIdx = new Dictionary<string, int>(StringComparer.Ordinal);
        int rank = 0;
        bool seenParaOpen = false;
        // BUG-DUMP-SDTORDER-HYPERLINK: a <w:r> nested inside a <w:hyperlink> (or
        // an ins/del/smartTag/customXml/dir/bdo run-wrapper) IS surfaced by
        // Navigation as a paragraph-level /r[N] — its run resolver flattens
        // Descendants<Run>() excluding only SdtRun-nested runs (see the "r" case
        // in WordHandler.Navigation.cs). So r[N] must count runs THROUGH those
        // transparent wrappers; counting only literal top-level children
        // desynced r[N] (a paragraph with hyperlinks numbered ". If the
        // assessment" as r[3] here but r[5] in Navigation) and scrambled the
        // inline-SDT flush order — a content control between two runs came back
        // attached to the wrong run. Only <w:pPr> and <w:sdt> are opaque (pPr's
        // children aren't content; an inline SDT's runs surface under the sdt
        // node, not as paragraph runs); every other run-container is transparent.
        var transparentWrappers = new HashSet<string>(StringComparer.Ordinal)
        {
            "hyperlink", "ins", "del", "moveFrom", "moveTo",
            "smartTag", "customXml", "dir", "bdo",
        };
        var openStack = new Stack<bool>(); // true = this open incremented suppress
        int suppress = 0; // >0 ⇒ inside an opaque container (pPr / sdt / a run)
        // Match element opens/closes/self-closes for the w: and m: namespaces
        // (m:oMathPara / m:oMath surface as paragraph children too).
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(xml, @"<(/?)(?:w|m):([A-Za-z]+)\b[^>]*?(/?)>"))
        {
            var closing = m.Groups[1].Value == "/";
            var name = m.Groups[2].Value;
            var selfClose = m.Groups[3].Value == "/";
            if (!seenParaOpen)
            {
                // The first open we encounter is the paragraph element itself.
                if (!closing) { seenParaOpen = true; }
                continue;
            }
            if (closing)
            {
                if (openStack.Count > 0 && openStack.Pop()) suppress--;
                continue;
            }
            if (suppress == 0)
            {
                // Paragraph-level child (only transparent wrappers above it) —
                // assign the next document rank under its OOXML local name
                // (matches the /r[N], /sdt[N], … path segments Navigation builds).
                var seg = name switch
                {
                    "r" => "r",
                    "sdt" => "sdt",
                    "bookmarkStart" => "bookmark",
                    "bookmarkEnd" => "bookmarkEnd",
                    "hyperlink" => "hyperlink",
                    "fldSimple" => "field",
                    _ => name,
                };
                int idx = perNameIdx.TryGetValue(seg, out var c) ? c : 0;
                perNameIdx[seg] = idx + 1;
                int thisRank = rank++;
                map[$"{seg}[{idx + 1}]"] = thisRank;
                // BUG-DUMP-INLINESDT-BMEND-RANK: Navigation builds a bookmarkEnd
                // child's path as bookmarkEnd[@id=N] (id-keyed, not positional —
                // BUG-DUMP-BMEND-IDPATH), so the positional map key alone misses in
                // ChildDocRank → the run loop reads rank=int.MaxValue for the
                // bookmarkEnd child and prematurely flushes a still-pending inline
                // SDT before the run that precedes it. That silently reorders the
                // run across the content control ("with [SDT]" -> "[SDT]with"),
                // garbling text with no validate flag and no visible render change.
                // Register an id-keyed alias at the SAME rank so the lookup hits and
                // run<->inline-SDT document order round-trips.
                if (seg == "bookmarkEnd")
                {
                    var idm = System.Text.RegularExpressions.Regex.Match(m.Value, "w:id=\"(\\d+)\"");
                    if (idm.Success)
                        map[$"bookmarkEnd[@id={idm.Groups[1].Value}]"] = thisRank;
                }
            }
            if (!selfClose)
            {
                // Transparent run-wrappers do NOT suppress their children (inner
                // runs still count as paragraph runs); everything else (pPr, sdt,
                // a run and its rPr/text) is opaque.
                bool opaque = !transparentWrappers.Contains(name);
                if (opaque) suppress++;
                openStack.Push(opaque);
            }
        }
        return map;
    }

    // Look up a child node's document rank from the trailing positional path
    // segment (e.g. "/body/p[…]/r[2]" → "r[2]"). Falls back to int.MaxValue when
    // the path uses a non-positional segment (e.g. r[@…]) or isn't in the map,
    // so unrecoverable children sort to the tail rather than to the front.
    private static int ChildDocRank(Dictionary<string, int> docOrder, string? path)
    {
        if (string.IsNullOrEmpty(path) || docOrder.Count == 0) return int.MaxValue;
        int slash = path.LastIndexOf('/');
        var seg = slash >= 0 ? path[(slash + 1)..] : path;
        return docOrder.TryGetValue(seg, out var r) ? r : int.MaxValue;
    }

    // ── Extracted helpers (behavior unchanged from inline original) ──

    private static bool TryEmitDisplayEquation(WordHandler word, DocumentNode pNode, string parentPath, bool autoPresent, List<BatchItem> items)
    {
        // Display-mode equations (<m:oMathPara>) surface in EmitBody's
        // bodyNode.Children as type=paragraph, but a direct Get on the
        // path returns type=equation with the LaTeX-ish formula in
        // DocumentNode.Text. EmitParagraph would otherwise emit an empty
        // `add p` and lose the entire formula. Route to typed
        // `add /body --type equation` instead.
        if (pNode.Type != "equation" || parentPath != "/body" || autoPresent) return false;
        var mode = pNode.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode
        };
        if (!string.IsNullOrEmpty(pNode.Text))
            eqProps["formula"] = pNode.Text!;
        // BUG-DUMP-EQVERBATIM (display): forward the verbatim <m:oMath> so the
        // rebuilt equation keeps its math-run rPr (Cambria Math, sizes) instead
        // of being reparsed from the lossy LaTeX string.
        if (pNode.Format.TryGetValue("xml", out var eqXml)
            && eqXml != null && eqXml.ToString() is { Length: > 0 } eqXmlS
            && eqXmlS.Contains("oMath", StringComparison.Ordinal))
            eqProps["xml"] = eqXmlS;
        // Carry any OLE/preview-image parts referenced inside the verbatim math
        // (MathType/Equation objects) so they don't dangle on replay.
        AddMathInlinedPartProps(word, pNode.Path, eqProps);
        // BUG-DUMP19-02: forward block-equation alignment.
        if (pNode.Format.TryGetValue("align", out var eqAlign)
            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
            eqProps["align"] = eqAlign.ToString()!;
        // BUG-DUMP-EQDISPLAY-PPR: forward the wrapper paragraph's spacing so the
        // rebuilt display-equation paragraph keeps its line height (e.g. 1.5x);
        // dropping it collapsed the equation line and compressed the page.
        foreach (var sk in new[] { "lineSpacing", "lineRule", "spaceBefore", "spaceAfter", "wrapperAlign", "wrapperPpr" })
            if (pNode.Format.TryGetValue(sk, out var sv)
                && sv != null && sv.ToString() is { Length: > 0 } svs)
                eqProps[sk] = svs;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    private static bool TryEmitInlineSectionBreak(WordHandler word, DocumentNode pNode, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Inline section break: a paragraph carrying <w:sectPr> is the
        // OOXML representation of a mid-document section boundary.
        // AddSection on /body produces this same shape, so we emit
        // `add /body --type section` (which creates a fresh break paragraph)
        // rather than emitting a regular `add p`. The companion
        // sectionBreak.* keys map back to AddSection's prop vocabulary.
        if (parentPath != "/body" ||
            !pNode.Format.TryGetValue("sectionBreak", out var breakKind) || breakKind == null)
            return false;
        {
            // BUG-DUMP-R31-1: a childless mid-document <w:sectPr/> must round-trip
            // bare — no fabricated <w:type>, no default pgSz/pgMar. Navigation
            // emits sectionBreak.type ONLY when the source carried a real
            // <w:type>, and sectionBreak.empty=true when the sectPr was truly
            // childless. Emit `type` only when the source had one; forward the
            // `empty` flag so AddSection produces a bare sectPr instead of
            // stamping the body section's geometry.
            bool sourceHadType = pNode.Format.ContainsKey("sectionBreak.type");
            var sectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (sourceHadType)
                sectProps["type"] = breakKind.ToString() ?? "nextPage";
            foreach (var (k, v) in pNode.Format)
            {
                if (!k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase)) continue;
                if (v == null) continue;
                var keyTail = k["sectionBreak.".Length..];
                // `sectionBreak.type` already drove the `type` key above; don't
                // also emit a stray `type` (it's the same value) — skip the alias.
                if (string.Equals(keyTail, "type", StringComparison.OrdinalIgnoreCase)) continue;
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) sectProps[keyTail] = s;
            }
            // Fold the carrier sectPr's pgBorders.<side>.sz/.color/.space sub-keys
            // (now prefix-stripped to bare pgBorders.* form) into the single
            // STYLE;SIZE;COLOR;SPACE value AddSection's pgBorders.<side> case
            // parses. Mirrors EmitSection's FoldPgBordersProps for the body sectPr.
            FoldPgBordersProps(sectProps);
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "section",
                Props = sectProps
            });
            // BUG-DUMP-R37-1: AddSection only captures the sectPr geometry
            // (page/margin/grid/column via sectionBreak.* keys). The HOST
            // paragraph that carries the <w:sectPr> in its pPr also holds its
            // own paragraph-level properties (spacing/ind/jc/widowControl/…) and
            // a paragraph-mark rPr (bold/size/color/font.ea) — the empty
            // section-terminating paragraph's mark-rPr size/spacing set its
            // rendered line height, so dropping them shifts pagination near
            // section boundaries. Re-apply them with a `set` on the rebuilt
            // section paragraph (now at /body/p[last()]), mirroring how a normal
            // paragraph round-trips its pPr: reuse FilterEmittableProps (same
            // pPr/mark-rPr vocabulary AddParagraph/Set understand) and strip the
            // sectionBreak.* keys already consumed by `add section`. On a
            // run-less section paragraph the bare bold/size/color/font.* keys
            // land on the ParagraphMarkRunProperties (SetElement Paragraph branch).
            var sectPProps = FilterEmittableProps(pNode.Format);
            sectPProps.Remove("sectionBreak");
            foreach (var k in sectPProps.Keys
                         .Where(k => k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase))
                         .ToList())
                sectPProps.Remove(k);
            // BUG-DUMP-SECTNUM: this `set` reuses AddParagraph/SetElement's numbering
            // vocabulary, so it must apply the SAME inheritance filters as the normal
            // emit (line ~70). Without it a section-carrier paragraph that inherits
            // numbering from its style emitted numId+numFmt+start on the `set`,
            // triggering ad-hoc numbering-definition creation: a spurious num +
            // abstractNum in numbering.xml and a direct numPr stamped on a paragraph
            // that had none in the source.
            ApplyNumberingInheritanceFilters(sectPProps, pNode);
            if (sectPProps.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = "/body/p[last()]",
                    Props = sectPProps
                });
            }
            // BUG-R12C: the section-carrier paragraph holds its own custom tab
            // stops in Format["tabs"], which the `set` above can't express (tab
            // stops round-trip as separate `add tab` ops, never as a pPr prop).
            // A title-page paragraph that centres its text by tabbing to a
            // centre-aligned tab stop (jc=left + <w:tab w:val="center"/>) lost
            // the stop here and the text fell back to the default left tab —
            // "YEAR OF SUBMISSION" rendered left-aligned instead of centred.
            // Mirror the normal paragraph path's EmitTabStops at line 287.
            pNode.Format.TryGetValue("tabs", out var carrierTabs);
            EmitTabStops("/body/p[last()]", carrierTabs, items);
            // BUG-DUMP4-04: a section-break paragraph can also carry visible
            // text runs (the carrier paragraph is just a regular paragraph
            // with sectPr in its pPr). AddSection appends a fresh paragraph
            // at /body/p[targetIndex]; emit each text-bearing run as
            // `add r` against that paragraph.
            var carrierRuns = (pNode.Children ?? new List<DocumentNode>())
                .Where(c =>
                {
                    // BUG-DUMP7-11: include inline w:sdt carrier children.
                    if (c.Type == "sdt") return true;
                    // BUG-R12C: a tab / positional-tab run surfaces as type
                    // "tab"/"ptab" with empty Text, so the !IsNullOrEmpty(Text)
                    // gate below dropped it. A title-page paragraph that centres
                    // its content by tabbing to a centre tab stop (jc=left +
                    // leading <w:tab/>) then rendered left-aligned. Keep the tab
                    // run; the loop emits it via TryEmitTabRun/TryEmitPtabRun.
                    if (c.Type == "tab" || c.Type == "ptab") return true;
                    // Spanning bookmarks anchored on the section paragraph:
                    // dropping the start leaves the matching end dangling
                    // (replay fails with "no matching open bookmarkStart").
                    if (c.Type == "bookmark" || c.Type == "bookmarkEnd") return true;
                    // BUG-DUMP-SECTBR: a pure page/column/line break run surfaces
                    // as type "break" with empty Text (the main run loop routes it
                    // through TryEmitBreakRun). The "page break, then a new section"
                    // idiom — <w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:br
                    // w:type="page"/></w:r></w:p> — puts that break on the
                    // section-carrier paragraph, where the run/r/picture gate below
                    // dropped it, collapsing the forced page break and reflowing the
                    // section boundary. Keep it for the TryEmitBreakRun call below.
                    if (c.Type == "break") return true;
                    // Anchored cover art surfaces as type="picture" when the
                    // drawing carries an image blip — include it for the
                    // drawing-carrier branch below.
                    if (c.Type != "run" && c.Type != "r" && c.Type != "picture") return false;
                    if (!string.IsNullOrEmpty(c.Text)) return true;
                    // BUG-DUMP5-08 / BUG-R7B(BUG1): include empty footnote /
                    // endnote reference runs (their visible text comes via the
                    // typed emit branch below, not from c.Text). Detect via the
                    // actual reference child, not the arbitrary rStyle name.
                    var rsv = c.Format.TryGetValue("rStyle", out var rsraw) ? rsraw?.ToString() : null;
                    if (ClassifyNoteRefRun(word, c, rsv) != NoteRefKind.None)
                        return true;
                    // A cover-page section paragraph can host anchored
                    // drawings (textboxes, picture-filled shapes) in
                    // otherwise text-less runs; dropping them deleted the
                    // whole cover art. Include drawing/pict-bearing runs.
                    var crx = word.GetElementXml(c.Path);
                    if (!string.IsNullOrEmpty(crx)
                        && (crx.Contains("<w:drawing", StringComparison.Ordinal)
                            || crx.Contains("<w:pict", StringComparison.Ordinal)))
                        return true;
                    return false;
                })
                .ToList();
            if (carrierRuns.Count > 0)
            {
                var carrierPath = $"/body/p[last()]";
                // BUG-DUMP-SECTHL: a hyperlink hosted in the section-carrier
                // paragraph (e.g. a List-of-Figures entry whose paragraph also
                // carries the <w:sectPr>) must round-trip through the same
                // structured path the main run loop uses. Coalesce consecutive
                // hyperlink runs (text + leader tabs + page number sharing one
                // <w:hyperlink> wrapper) into a single synthetic node so
                // EmitPlainOrHyperlinkRun emits the `add hyperlink` wrapper before
                // its trailing runs — without this the carrier loop emitted each
                // run as a bare `add r`, the wrapper was never created, and the
                // trailing tab/page-number runs (which target .../hyperlink[1])
                // were dropped. Capture the prior-paragraph hyperlink count so the
                // wrapper re-indexes from 1 (BUG-R14B), same as the main loop.
                carrierRuns = CoalesceHyperlinkRuns(carrierRuns);
                int carrierHlBaseline = items.Count(it => it.Type == "hyperlink"
                    && string.Equals(it.Parent, carrierPath, StringComparison.Ordinal));
                foreach (var run in carrierRuns)
                {
                    if (run.Type == "bookmark" || run.Type == "bookmarkEnd")
                    {
                        TryEmitBookmarkRun(run, carrierPath, items, ctx);
                        continue;
                    }
                    // Coalesced hyperlink run (or a lone hyperlink-wrapped run):
                    // emit the structured wrapper, not a bare `add r`.
                    if ((run.Type == "run" || run.Type == "r")
                        && (run.Format.ContainsKey("url") || run.Format.ContainsKey("anchor")))
                    {
                        EmitPlainOrHyperlinkRun(word, run, carrierPath, items, ctx, carrierHlBaseline);
                        continue;
                    }
                    // BUG-R12C: tab / positional-tab runs round-trip through the
                    // same helpers the main run loop uses, so a leading centre
                    // tab survives on the section-carrier paragraph.
                    if (TryEmitTabRun(run, carrierPath, items)) continue;
                    if (TryEmitPtabRun(run, carrierPath, items)) continue;
                    // BUG-DUMP-SECTBR: a pure page/column/line break run on the
                    // section-carrier paragraph round-trips through the same helper
                    // the main run loop uses, so a forced page break that precedes
                    // a section boundary survives.
                    if (TryEmitBreakRun(word, run, parentPath, carrierPath, items, ctx)) continue;
                    // BUG-DUMP7-11: inline SDT carrier — same prop whitelist
                    // as the body-paragraph inline-SDT branch.
                    if (run.Type == "sdt")
                    {
                        // BUG-R12C: a rich inline SDT (per-run rPr/color,
                        // multi-run, nested, or special sdtPr) must round-trip
                        // verbatim here too. The section-carrier branch only had
                        // the flat `add sdt text=` path, which dropped the sdtPr
                        // rStyle/placeholder and the content runs' formatting —
                        // a title-page "year OF SUBMISSION" control styled by a
                        // smallCaps character style came back as literal lowercase
                        // text. Mirror EmitInlineSdt: try the rich raw-set first,
                        // fall through to the flat emit when it isn't rich (or the
                        // host isn't raw-set-addressable).
                        if (TryEmitRichInlineSdt(word, run, parentPath, items, ctx))
                            continue;
                        var sdtCarrierProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
                        {
                            if (run.Format.TryGetValue(key, out var v) && v != null)
                            {
                                var s = v.ToString() ?? "";
                                if (s.Length > 0) sdtCarrierProps[key] = s;
                            }
                        }
                        if (!string.IsNullOrEmpty(run.Text))
                            sdtCarrierProps["text"] = run.Text!;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "sdt",
                            Props = sdtCarrierProps
                        });
                        continue;
                    }
                    var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
                    // BUG-R7B(BUG1): detect via the actual reference child, not
                    // the (arbitrary) rStyle name.
                    var carrierNoteKind = ctx != null
                        ? ClassifyNoteRefRun(word, run, rStyle) : NoteRefKind.None;
                    // BUG-R12A(BUG3): structural note-body emit (per-run rPr +
                    // multi-paragraph), same as TryEmitNoteRefRun. The cursor is
                    // pre-incremented to the 1-based source/target note index.
                    if (carrierNoteKind == NoteRefKind.Footnote)
                    {
                        int idx = ++ctx!.FootnoteCursor.Index;
                        EmitNoteReference(word, "footnote", idx, idx, carrierPath, items, run);
                        continue;
                    }
                    if (carrierNoteKind == NoteRefKind.Endnote)
                    {
                        int idx = ++ctx!.EndnoteCursor.Index;
                        EmitNoteReference(word, "endnote", idx, idx, carrierPath, items, run);
                        continue;
                    }
                    // Drawing/pict-bearing carrier run: route through the
                    // full picture pipeline first (charts, textboxes, wps
                    // shapes, inline pictures all have typed/carrier paths
                    // there); only a run the pipeline declines falls back to
                    // the inlined-parts carriers or — for a drawing with no
                    // relationship references — a verbatim raw-set.
                    var carrierRunXml = word.GetElementXml(run.Path);
                    if (!string.IsNullOrEmpty(carrierRunXml)
                        && (carrierRunXml.Contains("<w:drawing", StringComparison.Ordinal)
                            || carrierRunXml.Contains("<w:pict", StringComparison.Ordinal)))
                    {
                        if (TryEmitPictureRun(word, run, carrierPath, "/body", 0, items, ctx))
                            continue;
                        if (word.GetDrawingShapeEmitData(run.Path) is { } csData)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = carrierPath,
                                Type = "inlinedparts",
                                Props = PackInlinedPartsProps(csData),
                            });
                            continue;
                        }
                        if (word.GetVmlShapeEmitData(run.Path) is { } cvData)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = carrierPath,
                                Type = "inlinedparts",
                                Props = PackInlinedPartsProps(cvData),
                            });
                            continue;
                        }
                        if (!HasExternalRelRef(carrierRunXml))
                        {
                            items.Add(new BatchItem
                            {
                                Command = "raw-set",
                                Part = "/document",
                                Xpath = "/w:document/w:body/w:p[last()]",
                                Action = "append",
                                Xml = carrierRunXml
                            });
                        }
                        else
                        {
                            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                                Element: "drawing",
                                Path: run.Path,
                                Reason: "section-paragraph drawing references relationships that cannot be reconstructed; it is dropped on replay"));
                        }
                        continue;
                    }
                    var rProps = FilterEmittableProps(run.Format);
                    if (!string.IsNullOrEmpty(run.Text))
                        rProps["text"] = run.Text!;
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = carrierPath,
                        Type = "r",
                        Props = rProps
                    });
                }
            }
            // BUG-DUMP-R61-TOCSECT: a cross-paragraph field (canonically a TOC or
            // INDEX) can OPEN or CLOSE in a section-carrier paragraph — the
            // paragraph holds both the section's <w:sectPr> and a field marker run
            // (<w:fldChar w:fldCharType="begin"/> + <w:instrText> + separate for
            // the opener, or the terminating end fldChar for the closer). The
            // carrier branch emits the section + visible runs but never the bare
            // field marker runs (they carry no visible text, so the carrierRuns
            // filter drops them). Worse, the field's entry paragraphs round-trip
            // verbatim via EmitCrossParagraphFieldMember but this opener/closer
            // paragraph reaches THIS branch instead: the sectPr's
            // <w:headerReference>/<w:footerReference r:id="…"/> trips that member's
            // HasExternalRelRef guard (a false positive — the header/footer ref is
            // recreated by the section emit, not a dangling content rel), so it
            // falls back to the typed section emit here. Dropping the begin fldChar
            // + instrText leaves an empty INDEX/TOC instruction (Word regenerates
            // an EMPTY field — the whole index/toc disappears, reflowing the doc);
            // dropping the end fldChar leaves the field unterminated (Word renders
            // the raw field code). Re-emit every fldChar AND instrText run verbatim
            // via a rel-free raw-set append onto the rebuilt section paragraph
            // (these marker runs carry no relationship of their own), restoring the
            // field's opener instruction / terminator.
            foreach (var fcRun in (pNode.Children ?? new List<DocumentNode>())
                         .Where(c => c.Type == "fieldChar" || c.Type == "instrText"))
            {
                var fcXml = word.GetElementXml(fcRun.Path);
                if (string.IsNullOrEmpty(fcXml)) continue;
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = "/w:document/w:body/w:p[last()]",
                    Action = "append",
                    Xml = fcXml
                });
            }
            return true;
        }
    }

    /// <summary>
    /// BUG-DUMP-TXBX-WRAPPER: a body paragraph whose only meaningful child is
    /// a textbox-bearing Drawing run (textboxes ship inside
    /// <c>&lt;mc:AlternateContent&gt;</c>, so Get reports the run as
    /// type=&quot;run&quot; with no Format hints) used to emit BOTH an empty
    /// <c>add p</c> wrapper AND a typed <c>add textbox</c> row. On replay
    /// AddTextbox creates its own host paragraph, leaving the target with
    /// one extra empty paragraph per textbox. Detect the
    /// textbox-only-paragraph shape here and emit only the textbox row.
    /// </summary>
    private static bool TryEmitTextboxOnlyParagraph(
        WordHandler word, DocumentNode pNode, string parentPath, bool autoPresent,
        List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Wrapper coalescing only makes sense at /body — header/footer/cell
        // hosts of a textbox have their own pattern and we don't want to
        // skip wrapping paragraphs that carry visible run formatting.
        if (parentPath != "/body" || autoPresent) return false;
        if (!string.IsNullOrEmpty(pNode.Text)) return false;
        var children = pNode.Children ?? new List<DocumentNode>();
        // Need exactly one drawing-bearing child (run / picture) and nothing
        // else. Bookmarks / sdts / breaks need the paragraph wrapper to
        // anchor against and must not coalesce.
        if (children.Count != 1) return false;
        var run = children[0];
        // Source-side: AlternateContent wraps the drawing so Get reports the
        // run as plain "run"/"r" with no Format hints.
        // Target-side (after AddTextbox replay): Drawing sits directly under
        // Run with no AlternateContent, so Get reports it as "picture".
        // Both shapes must collapse here — otherwise source and target dumps
        // disagree on whether to emit the `add p` wrapper and the round-trip
        // drift grows on every textbox.
        if (run.Type != "run" && run.Type != "r" && run.Type != "picture") return false;
        // Don't gate on run.Text here: picture/textbox runs surface their
        // docPr name in DocumentNode.Text (e.g. "文本框 1") which is not
        // visible body text — it doesn't disqualify the wrapper-coalesce.
        var rawXml = word.GetElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml) || !IsTextboxDrawing(rawXml)) return false;
        // BUG-DUMP-R26-6: a legacy VML textbox round-trips via a verbatim
        // raw-set APPEND into the host paragraph (TryEmitTextbox), which needs
        // an `add p` to exist first. The host-less shortcut here (which relies
        // on AddTextbox creating its own modern host) would leave the raw-set
        // with no /body/p[last()] target. Bail so the normal EmitParagraph flow
        // emits `add p`, then routes the run through TryEmitPictureRun →
        // TryEmitTextbox, which raw-sets the VML into the just-added paragraph.
        if (IsVmlTextbox(rawXml)) return false;
        // Delegate to the same emit path TryEmitPictureRun uses so geometry
        // props + inner-paragraph recursion stay identical.
        return TryEmitTextbox(word, run, rawXml, parentPath, items, ctx);
    }

    private static bool TryEmitTocParagraph(DocumentNode pNode, string parentPath, List<BatchItem> items)
    {
        // TOC field-bearing paragraph: a fldChar(begin) + instrText("TOC ...")
        // + fldChar(separate) + placeholder run + fldChar(end) chain. Get
        // exposes only the placeholder text on the parent paragraph, so
        // emitting a regular `add p text=...` would drop the field structure
        // entirely and Word would no longer auto-update the TOC on open.
        if (parentPath != "/body" || pNode.Children == null) return false;
        var instrChild = pNode.Children
            .FirstOrDefault(c => c.Type == "instrText"
                && (c.Format.TryGetValue("instruction", out var iv)
                    && iv?.ToString()?.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true));
        if (instrChild == null) return false;
        // BUG-DUMP-TOC-COLOCATED-PICTURE: the typed `add toc` fast path emits ONLY
        // the TOC field and returns, so any OTHER content co-located in the same
        // paragraph is dropped. The canonical case is a background/letterhead
        // picture anchored on the TOC's first paragraph (a behindDoc logo) — it
        // vanished silently. Bail to the generic EmitParagraph path when the
        // paragraph also carries a drawing: that path emits the picture via
        // TryEmitPictureRun AND round-trips the TOC field verbatim through the
        // generic field-emit (instr=…), so Word still regenerates the TOC.
        if (pNode.Children.Any(c => c.Type == "picture")) return false;
        var instr = instrChild.Format["instruction"]!.ToString()!;
        // BUG-DUMP-TOC-LOSSY: the typed `add toc` path does NOT round-trip an
        // arbitrary TOC field. AddToc reconstructs a CANONICAL instruction —
        // it always emits ` TOC \o "{levels}"` (defaulting levels to "1-3"),
        // always appends ` \u `, and always writes the "Update field to see
        // table of contents" placeholder as the field result. ParseTocInstruction
        // only understands \o \h \z \t \b, so any other switch is silently
        // dropped. Concrete corruptions this caused:
        //  - ` TOC \h \z \c "Table" ` (a Table-of-Tables / table-of-captions)
        //    became ` TOC \o "1-3" \h \z \u ` — \c "Table" dropped, a bogus
        //    \o "1-3" fabricated, so the field switched from a caption index to
        //    a heading index.
        //  - ` TOC \h \z \t "Style,1" ` (a custom-style index with no \o) gained
        //    a fabricated ` \o "1-3"`.
        //  - A bare ` TOC \o "1-3" ` (no \u) gained a stray ` \u `.
        // Only fire the typed path for an instruction AddToc reproduces BYTE-FOR-
        // BYTE from the parsed props (same switch set, in canonical form) that is
        // self-contained in THIS paragraph. Anything AddToc would reshape — a
        // missing \o or \u that it fabricates, or any switch it can't represent —
        // bails to the generic field-emit path (CollapseFieldChains →
        // BuildFieldAddProps default arm), which preserves the full instruction
        // verbatim via `instr=`. A cross-paragraph TOC (the canonical Word shape,
        // whose begin/instr/separate open here, whose cached entries live in
        // following paragraphs or a top-level <w:sdt>, and whose end closes
        // elsewhere) is routed verbatim by the cross-paragraph field-span
        // machinery (GetCrossParagraphFieldSpanRanges → EmitCrossParagraphFieldMember)
        // and never reaches this method; the self-contained guard here is a safety
        // net for a span that escapes that detection. The authored `add toc`
        // shape (` TOC \o "1-3" \h \u ` + placeholder, all in one paragraph)
        // round-trips byte-for-byte through AddToc and stays on the typed path.
        if (!TocInstructionRoundTripsThroughAddToc(instr)) return false;
        var fldCharTypes = pNode.Children
            .Where(c => c.Type == "fieldChar")
            .Select(c => c.Format.TryGetValue("fieldCharType", out var fv) ? fv?.ToString() : null)
            .ToList();
        bool selfContained = fldCharTypes.Any(t => string.Equals(t, "begin", StringComparison.OrdinalIgnoreCase))
            && fldCharTypes.Any(t => string.Equals(t, "end", StringComparison.OrdinalIgnoreCase));
        if (!selfContained) return false;
        var tocProps = ParseTocInstruction(instr);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "toc",
            Props = tocProps
        });
        return true;
    }

    // True when AddToc would reproduce this TOC instruction BYTE-FOR-BYTE from
    // the props ParseTocInstruction extracts. AddToc emits a fixed canonical
    // shape — ` TOC \o "{levels}"` (+ ` \h` if hyperlinks, + ` \z` if page
    // numbers suppressed, + ` \t "{cs}"`, + ` \b "{bm}"`) and ALWAYS appends
    // ` \u `. So the typed path is faithful only when the source instruction
    // already has \o and \u, optionally h/z/t/b, and nothing else (any switch
    // ParseTocInstruction can't represent — \c \f \a \n \p \s \d \l \w \x \# —
    // or a missing \o / \u that AddToc would fabricate, makes it lossy).
    // Reconstruct AddToc's exact output from the parsed props and compare,
    // normalizing only inter-token whitespace so source switch ORDER doesn't
    // matter (AddToc emits o,h,z,t,b,u; the source may list them differently).
    private static bool TocInstructionRoundTripsThroughAddToc(string instruction)
    {
        var props = ParseTocInstruction(instruction);
        // ParseTocInstruction sets levels only when \o is present; AddToc would
        // default it to "1-3" and so fabricate an \o the source lacks.
        if (!props.ContainsKey("levels")) return false;
        // AddToc always appends \u; a source without \u would gain one.
        if (!System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\u\\b")) return false;
        // Reject any switch AddToc can't represent (would be silently dropped).
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(instruction, "\\\\([A-Za-z#])"))
        {
            if (!"ohztbu".Contains(m.Groups[1].Value, StringComparison.Ordinal)) return false;
        }
        // Rebuild AddToc's canonical instruction (mirrors AddToc in
        // WordHandler.Add.Structure.cs) and compare on a whitespace-normalized,
        // switch-sorted basis so ordering differences alone don't force a bail.
        var rebuilt = new System.Text.StringBuilder($" TOC \\o \"{props["levels"]}\"");
        if (props.TryGetValue("hyperlinks", out var h) && h == "true") rebuilt.Append(" \\h");
        if (props.TryGetValue("pageNumbers", out var z) && z == "false") rebuilt.Append(" \\z");
        if (props.TryGetValue("customStyles", out var cs) && !string.IsNullOrEmpty(cs))
            rebuilt.Append($" \\t \"{cs}\"");
        if (props.TryGetValue("bookmark", out var bm) && !string.IsNullOrEmpty(bm))
            rebuilt.Append($" \\b \"{bm}\"");
        rebuilt.Append(" \\u ");
        return TocCanonicalForm(rebuilt.ToString()) == TocCanonicalForm(instruction);
    }

    // Canonical comparison form for a TOC instruction: collapse runs of
    // whitespace to a single space, trim, and sort the switch tokens (each a
    // `\x` optionally followed by its quoted/bare argument) so two instructions
    // that differ only in switch order compare equal.
    private static string TocCanonicalForm(string instruction)
    {
        var collapsed = System.Text.RegularExpressions.Regex
            .Replace(instruction.Trim(), "\\s+", " ");
        // Split into the leading "TOC" token plus each `\x [arg]` switch.
        var m = System.Text.RegularExpressions.Regex.Match(collapsed, "^TOC\\b");
        if (!m.Success) return collapsed;
        var rest = collapsed[m.Length..];
        var switches = System.Text.RegularExpressions.Regex
            .Matches(rest, "\\\\[A-Za-z#](?:\\s+(?:\"[^\"]*\"|[^\\\\\\s]+))?")
            .Select(s => System.Text.RegularExpressions.Regex.Replace(s.Value.Trim(), "\\s+", " "))
            .OrderBy(s => s, StringComparer.Ordinal)
            .ToList();
        return "TOC " + string.Join(" ", switches);
    }

    private static bool ShouldCollapseSingleRun(WordHandler word, List<DocumentNode> runs, int breaksCount, int bookmarksCount, int inlineSdtsCount)
    {
        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        if (runs.Count > 1) return false;
        if (breaksCount > 0 || bookmarksCount > 0 || inlineSdtsCount > 0) return false;
        if (runs.Count == 0) return true;
        // BUG-DUMP-PERM: a paragraph holding a ranged editing-permission marker
        // must stay on the explicit-run path so TryEmitPermRun replays the
        // marker at its offset; collapsing into `add p` would drop it.
        if (runs.Any(rr => rr.Type == "permStart" || rr.Type == "permEnd")) return false;
        // BUG-DUMP-R40-7: a run-LESS paragraph whose only content is a
        // bookmark/bookmarkEnd marker (the close end of a bookmark that spans
        // into this paragraph) must stay on the explicit-run path so
        // TryEmitBookmarkRun replays the `add bookmark end=true` op. The
        // `bookmarks` collector (pNode.Children where Type=="bookmark") only
        // catches the START side, so a lone bookmarkEnd left bookmarksCount==0
        // and the single-run collapse folded the marker's name= into `add p`
        // props (leaking a phantom paragraph name and dropping the bookmarkEnd
        // entirely — the bookmark was left unterminated). Bookmark markers carry
        // no paragraph-level text/format, so there is nothing to collapse anyway.
        if (runs.Any(rr => rr.Type == "bookmark" || rr.Type == "bookmarkEnd")) return false;
        var r = runs[0];
        // Picture / ptab runs need their own typed `add` rows.
        if (r.Type == "picture" || r.Type == "ptab") return false;
        // BUG-DUMP-R25-2: a sole run whose only content is a tab CHARACTER
        // (<w:r><w:tab/></w:r>, surfaced as Type=="tab" with empty Text) must
        // stay on the explicit-run path so TryEmitTabRun replays `add r
        // text="\t"`. Collapsing into `add p` flattens it via GetRunText and
        // — with no <w:t> text to carry — the tab character vanishes. In
        // multi-run paragraphs the tab survives because sibling run ops keep
        // the structure; only this sole-run case was lost.
        if (r.Type == "tab") return false;
        // OLE / embedded-object runs must stay on the explicit-run path so
        // TryEmitOleRun fires. Collapsing folds the ole metadata (progId,
        // fileSize, drawAspect, …) into `add p` props that AddParagraph
        // silently ignores — the embedded object vanishes with no warning.
        // The explicit path surfaces the documented "ole run dropped" warning
        // (full binary round-trip is a separate, deferred feature) and keeps
        // the host paragraph intact.
        if (r.Type == "ole") return false;
        // A single run carrying a drawing/shape payload (textbox, connector
        // line, autoshape — surfaced as a plain "run" when AlternateContent-
        // wrapped) must NOT collapse into `add p`: collapsing flattens the
        // paragraph and silently drops the shape. Keep it on the explicit-run
        // path so TryEmitPictureRun preserves it (typed textbox or raw-set).
        if ((r.Type == "run" || r.Type == "r"))
        {
            var probe = word.GetElementXml(r.Path);
            if (!string.IsNullOrEmpty(probe) && IsDrawingBearingRun(probe)) return false;
            // BUG-DUMP-R24-3: a sole run carrying a page/column <w:br> MIXED with
            // <w:t> text must stay on the explicit-run path so TryEmitMixedBreakRun
            // raw-passes the verbatim <w:r>. Collapsing to `add p text="…"`
            // flattens the run via GetRunText (which drops page/column breaks —
            // they have no \n representation), silently losing the break.
            if (!string.IsNullOrEmpty(probe)
                && System.Text.RegularExpressions.Regex.IsMatch(probe, @"<w:br\b[^>]*\bw:type=""(?:page|column)""")
                && System.Text.RegularExpressions.Regex.IsMatch(probe, @"<w:t[\s>]"))
                return false;
        }
        // R14-bug1+2: legacy form field synth — needs its own typed
        // `add formfield` row; collapsing into `add p` flattens the
        // ffData wrapper into paragraph props that AddParagraph ignores.
        if (r.Type == "formfield") return false;
        // BUG-DUMP6-05 / BUG-DUMP10-05: hyperlink-wrapped run (url, anchor,
        // or tooltip-only via isHyperlink sentinel) must re-emit as
        // `add hyperlink` — `add p` does not consume url/anchor.
        if (r.Format.ContainsKey("url") || r.Format.ContainsKey("anchor")
            || r.Format.ContainsKey("isHyperlink")) return false;
        // BUG-FUZZ-2 / BUG-R7B(BUG1): footnote/endnote reference runs need the
        // typed `add footnote/endnote` branch; AddParagraph doesn't consume the
        // reference. Detect via the actual reference child, not the arbitrary
        // rStyle name (real docs use ids like "a5", not "FootnoteReference").
        {
            var srStyle = r.Format.TryGetValue("rStyle", out var srraw) ? srraw?.ToString() : null;
            if (ClassifyNoteRefRun(word, r, srStyle) != NoteRefKind.None)
                return false;
        }
        // BUG-W14-EFFECTS / BUG-DUMP5-09 / 7-01 / 5-10: run-level w14 effects /
        // OpenType properties / sym / trackChange — AddParagraph's
        // ApplyRunFormatting fallback has no cases for these; they'd
        // surface as UNSUPPORTED on replay.
        if (r.Format.ContainsKey("w14shadow")
            || r.Format.ContainsKey("textOutline")
            || r.Format.ContainsKey("textFill")
            || r.Format.ContainsKey("w14glow")
            || r.Format.ContainsKey("w14reflection")
            || r.Format.ContainsKey("ligatures")
            || r.Format.ContainsKey("numForm")
            || r.Format.ContainsKey("numSpacing")
            || r.Format.ContainsKey("revision.type")
            // BUG-DUMP-PGNUM: a sole run containing <w:pgNum/> (even one that
            // ALSO carries <w:t> text and/or a <w:cr/>) must stay on the
            // explicit-run path so TryEmitPgNumRun raw-passes the verbatim
            // <w:r>. Collapsing to `add p text="…"` flattens the run into
            // paragraph props and silently drops the pgNum placeholder.
            || r.Format.ContainsKey("_hasPgNum")
            // BUG-DUMP-DATEFIELD: a run containing a date-component placeholder
            // (<w:dayLong/> etc.) must stay on the explicit-run path so
            // TryEmitDateFieldRun raw-passes the verbatim <w:r>. Collapsing to
            // `add p text="…"` flattens the run into paragraph props and persists
            // GetRunText's "[dayLong]" sentinel as literal text, losing the
            // element Word substitutes the date against.
            || r.Format.ContainsKey("_hasDateField")
            // BUG-DUMP-R40-3: a run containing <w:noBreakHyphen/>/<w:softHyphen/>
            // must stay on the explicit-run path so TryEmitHyphenRun raw-passes
            // the verbatim <w:r>. Collapsing to `add p text="…"` persists
            // GetRunText's glyph as literal text and drops the structural hyphen.
            || r.Format.ContainsKey("_hasHyphen")
            || r.Format.ContainsKey("sym")) return false;
        // BUG-RSHD-PROMOTE: a sole run carrying run-level character shading
        // (<w:rPr><w:shd>) must NOT collapse into `add p`. AddParagraph routes
        // `shading`/`shd` to PARAGRAPH properties (<w:pPr><w:shd>) only — there
        // is no run-level shading routing on `add p` (BUG-DUMP22-03 deliberately
        // suppresses stamping pPr shading onto the inline run). Collapsing the
        // run's shading.* keys would therefore hoist a tight character highlight
        // into a full page-width paragraph band. Unlike `bdr`/`highlight` —
        // which `add p` forwards to ApplyRunFormatting (rPr) so they keep their
        // run-level identity in the collapse — `shading` semantics diverge
        // between pPr and rPr, so the only correct round-trip is the explicit
        // `add r shading=…` path (WordHandler.Add.Text.cs routes a run's
        // `shading`/`shd` to rPr). Keep this run on the explicit-run path.
        // Genuine paragraph-level shading (read from pPr onto the paragraph
        // node, not the run) still rides on `add p` unaffected.
        if (r.Format.Keys.Any(k => k.StartsWith("shading.", StringComparison.OrdinalIgnoreCase)))
            return false;
        // BUG-FIELD-COLLAPSE: a synthetic field run carries `instruction=…` —
        // collapse would lose the field chain on replay.
        if (r.Type == "field") return false;
        // BUG-DUMP-RUBY: a sole ruby (phonetic-guide) run must stay on the
        // explicit-run path so TryEmitRubyRun raw-sets the <w:ruby> verbatim.
        // Collapsing into `add p` would fold the base text into a plain run
        // and drop the <w:ruby>/<w:rt>/<w:rubyBase> wrapper.
        if (r.Type == "ruby") return false;
        // BUG-DUMP-R42-9: a sole <w:bdo> (bidirectional override) child must stay
        // on the explicit-run path so TryEmitBdoRun raw-sets the wrapper verbatim.
        // Collapsing into `add p` would fold the inner text into a plain run and
        // drop the <w:bdo> wrapper (and its load-bearing w:val direction).
        if (r.Type == "bdo") return false;
        // BUG-DUMP-R43-7: a sole <w:dir> (bidirectional embedding) child must stay
        // on the explicit-run path so TryEmitDirRun raw-sets the wrapper verbatim.
        if (r.Type == "dir") return false;
        // BUG-DUMP7-03: inline equation must emit `add equation` explicitly.
        if (r.Type == "equation") return false;
        // BUG-DUMP-R42-5: a sole run carrying its own reading direction
        // (<w:rPr><w:rtl/></w:rtl>, surfaced as Format["direction"]) must NOT
        // collapse into `add p`. The run-prop hoist would copy `direction` into
        // the paragraph prop bag, and AddParagraph routes `direction` through
        // ApplyDirectionCascade — stamping a paragraph-level <w:bidi/> AND a ¶
        // mark <w:rtl/>, flipping the whole paragraph's base direction. rtl is a
        // RUN-level property; keep it on the explicit-run path so it re-emits as
        // `add r --prop direction=…` (run rPr only). A genuine paragraph-level
        // <w:bidi/> rides on the paragraph node's own `direction` key (read from
        // pPr.BiDi) and is unaffected.
        if (r.Format.ContainsKey("direction")) return false;
        return true;
    }

    private static bool TryEmitBookmarkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP25-01: bookmark child emitted in DOM order so a
        // BookmarkStart between runs survives round-trip at its original
        // intra-paragraph offset.
        //
        // BUG-DUMP-BMSPAN: a content-WRAPPING bookmark (`_spanOpen=true`, set
        // by Navigation when the range holds runs/equations/fields) is split
        // into TWO positioned ops: an `open=true` start here (places only
        // <w:bookmarkStart>) and a matching `end=true` op emitted at the
        // BookmarkEnd's own DOM position (a separate `bookmarkEnd` child, even
        // a downstream paragraph for cross-paragraph spans). Document-order
        // replay then keeps the wrapped content INSIDE the range so
        // REF/PAGEREF/TOC anchors survive. A zero-length bookmark (no
        // `_spanOpen`) keeps the single combined op so it stays empty.
        if (run.Type == "bookmarkEnd")
        {
            var endProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (run.Format.TryGetValue("name", out var enm) && enm != null
                && enm.ToString() is { Length: > 0 } ens)
                endProps["name"] = ens;
            else
                return true; // unnamed end marker — start emit recreates pair
            // A legacy form field's embedded bookmark closes in a LATER
            // paragraph than the field run; AddFormField already recreated
            // the whole pair, so this stray end would fail with "no matching
            // open bookmarkStart" on replay.
            if (ctx != null && ctx.FormFieldBookmarkNames.Contains(endProps["name"]))
                return true;
            endProps["end"] = "true";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "bookmark",
                Props = endProps
            });
            return true;
        }
        if (run.Type != "bookmark") return false;
        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (run.Format.TryGetValue("name", out var bmName) && bmName != null)
        {
            var s = bmName.ToString();
            if (!string.IsNullOrEmpty(s)) bmProps["name"] = s;
        }
        if (bmProps.Count == 0) return true; // skip unnamed/anonymous bookmarks
        if (run.Format.TryGetValue("_spanOpen", out var sp) && sp is bool bsp && bsp)
        {
            bmProps["open"] = "true";
            // BUG-DUMP-R47-5: forward the SOURCE bookmark id for a span-open
            // start so a matching <w:bookmarkEnd> that round-trips verbatim via a
            // raw-set (e.g. inside a TOC <w:sdt> block) still pairs with it.
            // Without the id, AddBookmark allocates a fresh one and the start is
            // left unpaired (the raw-set end keeps the source id). AddBookmark
            // honors this id only on the open=true path; EnsureBookmarkIds dedupes
            // any collision by renumbering the matched pair.
            if (run.Format.TryGetValue("id", out var bmSrcId)
                && bmSrcId?.ToString() is { Length: > 0 } bmSrcIdStr)
                bmProps["id"] = bmSrcIdStr;
        }
        // BUG-DUMP-R32-4: forward colFirst/colLast for an inline
        // table-column-range bookmark (mirrors the body-direct case).
        ForwardBookmarkColRange(run.Format, bmProps);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "bookmark",
            Props = bmProps
        });
        return true;
    }

    // BUG-DUMP-R32-4: copy a bookmark node's colFirst/colLast (a rectangular
    // table-column-range bookmark) into the `add bookmark` props so AddBookmark
    // re-stamps w:colFirst/w:colLast. Shared by the body-direct and inline
    // bookmark emit paths.
    private static void ForwardBookmarkColRange(
        Dictionary<string, object?> format, Dictionary<string, string> bmProps)
    {
        if (format.TryGetValue("colFirst", out var cf)
            && cf?.ToString() is { Length: > 0 } cfs)
            bmProps["colFirst"] = cfs;
        if (format.TryGetValue("colLast", out var cl)
            && cl?.ToString() is { Length: > 0 } cls)
            bmProps["colLast"] = cls;
        // BUG-DUMP-BMDISPLACED: forward w:displacedByCustomXml ("next"/"prev")
        // so a bookmark adjacent to an SDT/custom-XML boundary keeps it — losing
        // it shifts the marker across the boundary and PAGEREF/TOC entries to the
        // bookmark render "Error! Bookmark not defined."
        if (format.TryGetValue("displacedByCustomXml", out var dbcx)
            && dbcx?.ToString() is { Length: > 0 } dbcxs)
            bmProps["displacedByCustomXml"] = dbcxs;
    }

    // BUG-DUMP-PERM: emit a ranged editing-permission marker (<w:permStart>/
    // <w:permEnd>) at its DOM position so an editable-region delimiter survives
    // round-trip. Mirrors TryEmitBookmarkRun: each marker is a positioned
    // `add permStart`/`add permEnd` op carrying its source attributes.
    private static bool TryEmitPermRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        if (run.Type != "permStart" && run.Type != "permEnd") return false;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in run.Type == "permStart"
            ? new[] { "id", "edGrp", "ed", "colFirst", "colLast" }
            : new[] { "id" })
        {
            if (run.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString() ?? "";
                if (s.Length > 0) props[key] = s;
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = run.Type,
            Props = props
        });
        return true;
    }

    private static bool TryEmitBreakRun(WordHandler word, DocumentNode run, string parentPath, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP5-01/02: a soft <w:br/> with NO type attribute is a line
        // break, not a page break — fall back to type=line. Emitted inline
        // from the unified runs loop so each break stays at its source
        // position instead of being hoisted to the front of the paragraph.
        if (run.Type != "break") return false;
        var breakType = run.Format.TryGetValue("breakType", out var bt) ? bt?.ToString() : null;
        // BUG-R12B(BUG3): a break run can carry its own <w:rPr> (<w:noProof/>,
        // font, color, …). Navigation strips every typography/proofing key from
        // a `break` node (TypographyOnlyKeys), and `add pagebreak` builds a bare
        // <w:r><w:br/></w:r> with no rPr, so the run-properties were dropped on
        // round-trip. The break has no scalar add-API for arbitrary rPr, so —
        // mirroring the ruby / rich-inline-SDT raw-set fallback — re-insert the
        // verbatim <w:r><w:rPr>…</w:rPr><w:br/></w:r> when the source run carries
        // a non-empty rPr. Restricted to /body hosts with no external rels (same
        // constraints as the other raw-set fallbacks) so no r:id/r:embed can
        // dangle; everything else stays on the lossless typed `add pagebreak`.
        // BUG-DUMP-DELBREAK: a break run wrapped in <w:del>/<w:ins>/move (e.g. a
        // tracked-DELETED page break — invisible in Word's final view) must NOT
        // take the verbatim raw-set path below: RawElementXml(run.Path) returns
        // the bare <w:r>, NOT the surrounding <w:del>, so the deletion wrapper is
        // dropped and the break resurrects as a LIVE page break (each one adds a
        // page). Route revision-wrapped breaks through the typed `add pagebreak`
        // path, which forwards revision.* so AddBreak re-wraps the rebuilt run.
        bool isRevisionWrapped = run.Format.ContainsKey("revision.type");
        if (parentPath == "/body" && !isRevisionWrapped)
        {
            var rawXml = word.RawElementXml(run.Path);
            if (!string.IsNullOrEmpty(rawXml)
                && BreakRunHasMeaningfulRunProps(rawXml!)
                && !HasExternalRelRef(rawXml!))
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = "/w:document/w:body/w:p[last()]",
                    Action = "append",
                    Xml = rawXml!
                });
                return true;
            }
        }
        var brkProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["type"] = string.IsNullOrEmpty(breakType) ? "line" : breakType!
        };
        if (run.Format.TryGetValue("breakClear", out var brkClear)
            && brkClear?.ToString() is { Length: > 0 } brkClearS)
            brkProps["breakClear"] = brkClearS;
        // BUG-DUMP-BREAKRPR: the verbatim raw-set fallback above only runs for
        // /body hosts. For breaks in other containers (table cells, header/
        // footer) the typed `add pagebreak` builds a bare <w:r><w:br/></w:r> and
        // drops the run's rPr — whose font/size sets the height of the line the
        // break starts, so the line collapsed to the default size and inflated
        // cell/row height. Forward the source run's <w:rPr> so AddBreak re-applies
        // it. Extract from the raw run XML (Navigation strips typography keys off
        // break nodes, so there is no scalar Format to read).
        var brkRawXml = word.RawElementXml(run.Path);
        if (!string.IsNullOrEmpty(brkRawXml))
        {
            try
            {
                var brkRunEl = System.Xml.Linq.XElement.Parse(brkRawXml!);
                var wNsBrk = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var brkRPrEl = brkRunEl.Element(wNsBrk + "rPr");
                if (brkRPrEl != null)
                    brkProps["breakRunRpr"] = brkRPrEl.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            }
            catch { /* bare break: no rPr to forward */ }
        }
        // BUG-DUMP-DELBREAK: forward the tracked-change attribution so AddBreak
        // re-wraps the rebuilt break run in <w:del>/<w:ins>/move. Without this a
        // deleted (invisible) page break replays as a live break and inflates the
        // page count. Mirrors the deleted-field forwarding (WrapRunsInRevision).
        foreach (var rk in new[] { "revision.type", "revision.author", "revision.date", "revision.id" })
            if (run.Format.TryGetValue(rk, out var rv)
                && rv?.ToString() is { Length: > 0 } rvs)
                brkProps[rk] = rvs;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "pagebreak",
            Props = brkProps
        });
        return true;
    }

    // BUG-DUMP-R24-3: a page/column break that lives MIXED inside a
    // text-bearing run (<w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After
    // </w:t></w:r>) is dropped on dump→batch. GetRunText flattens the run text
    // to "BeforeAfter" (page/column <w:br> has no \n representation, unlike a
    // textWrapping break) and the `add r` replay loses the break entirely.
    // A break that is the SOLE content of its run already round-trips via the
    // Navigation `break`-node upgrade + TryEmitBreakRun. Here we cover only the
    // mixed case: when the run is a plain text run whose verbatim XML carries a
    // page/column <w:br>, re-insert the whole <w:r> verbatim via a raw-set
    // append (mirrors the rich-break / ruby / pgNum raw-set fallback), so text
    // AND the break — with full run formatting — survive intact.
    private static bool TryEmitMixedBreakRun(WordHandler word, DocumentNode run, string parentPath, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Only plain text runs reach here (break-only runs are Type=="break"
        // and consumed by TryEmitBreakRun upstream). Picture/field/etc. runs
        // were already handled by their own TryEmit* probes before this point.
        if (run.Type is not ("run" or "r")) return false;
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return false;
        // Mixed only: a page/column <w:br> alongside a <w:t> (text) child. A
        // break-only run never has a <w:t>, so it won't match and stays on the
        // typed path. textWrapping/line breaks (no w:type, or w:type="textWrapping")
        // already round-trip as \n and must NOT be raw-set here.
        bool hasPageOrColumnBreak = System.Text.RegularExpressions.Regex.IsMatch(
            rawXml, @"<w:br\b[^>]*\bw:type=""(?:page|column)""");
        if (!hasPageOrColumnBreak) return false;
        bool hasText = System.Text.RegularExpressions.Regex.IsMatch(rawXml, @"<w:t[\s>]");
        if (!hasText) return false;
        // BUG-DUMP-R24-3 (FIX-3 follow-up): the verbatim raw-set targets
        // /w:document/w:body/w:p[last()]; only a body host has that addressable
        // last() paragraph. A mixed-break run inside a header/footer/cell would
        // mis-anchor. Return FALSE (not true) so the run falls through to the
        // normal text path (EmitPlainOrHyperlinkRun), which preserves the run's
        // <w:t> content via GetRunText — only the page/column <w:br> is dropped.
        // Returning true here would consume the run and DROP ITS TEXT TOO,
        // which is worse than the pre-fix flatten-but-keep-text behaviour.
        // Surface the deterministic "break lost" warning, then defer.
        if (parentPath != "/body")
        {
            // BUG-DUMP-R27 (BUG-DUMP-R24-3 follow-up): a SINGLE page/column break
            // that PRECEDES all text in its run (<w:r><w:br w:type="page"/><w:t>…
            // </w:t></w:r>) — pervasive in table cells whose leading page break
            // splits the row across pages — still round-trips in a non-body host:
            // emit the break as a typed `add pagebreak` on the paragraph's OWN
            // resolvable path (paraTargetPath = "{parentPath}/p[last()]" works for
            // cell/header/footer paragraphs, unlike the body-only raw-set xpath),
            // then return FALSE so the normal text path emits the run's <w:t>
            // AFTER it — preserving break-then-text order. Mid-run / trailing /
            // multi-break runs can't be ordered this way and keep warn-and-defer.
            var brkMatches = System.Text.RegularExpressions.Regex.Matches(
                rawXml, @"<w:br\b[^>]*\bw:type=""(page|column)""");
            var firstTextIdx = rawXml.IndexOf("<w:t", StringComparison.Ordinal);
            if (brkMatches.Count == 1
                && (firstTextIdx < 0 || brkMatches[0].Index < firstTextIdx))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "pagebreak",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["type"] = brkMatches[0].Groups[1].Value
                    }
                });
                return false; // text path (EmitPlainOrHyperlinkRun) emits <w:t> after
            }
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "break",
                Path: run.Path,
                Reason: "page/column break mixed inside a text run within a header/footer/table cell could not be serialized for round-trip; the break is dropped from the replayed document (the run's text is preserved)"));
            return false;
        }
        if (HasExternalRelRef(rawXml!)) return false;
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    // A break run is "rich" when it carries a non-empty <w:rPr> (noProof, font,
    // color, …) that the typed `add pagebreak` path cannot reproduce. An empty
    // <w:rPr/> (or none at all) is lossless on the typed path, so only a
    // populated rPr forces the verbatim raw-set fallback.
    private static bool BreakRunHasMeaningfulRunProps(string runXml)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            runXml, @"<w:rPr\b[^>]*?(?:/>|>(.*?)</w:rPr>)", System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!m.Success) return false;
        // Self-closing <w:rPr/> has no inner content → nothing to preserve.
        var inner = m.Groups[1].Value;
        return !string.IsNullOrWhiteSpace(inner);
    }

    private static bool TryEmitRubyRun(DocumentNode run, string parentPath, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-RUBY: a <w:ruby> (CJK phonetic guide) has no scalar
        // representation on add/set — its <w:rt> (furigana) and <w:rubyBase>
        // (base text) carry independent runs with their own rPr, and the
        // <w:rubyPr> alignment/hps/hpsRaise/hpsBaseText/lid attributes have no
        // add-API. Re-insert the captured <w:r><w:ruby>…</w:r> verbatim via a
        // raw-set append, mirroring the textbox/connector raw-set fallback in
        // TryEmitPictureRun (same shape: append the run's outer XML to the
        // just-emitted host paragraph at /w:document/w:body/w:p[last()]).
        if (run.Type != "ruby") return false;
        var rawXml = run.Format.TryGetValue("_rawRubyXml", out var rx) ? rx?.ToString() : null;
        if (string.IsNullOrEmpty(rawXml))
            return true; // nothing to emit (shouldn't happen) — consumed anyway
        // The verbatim raw-set targets /w:document/w:body/w:p[last()]; only a
        // body host has that addressable last() paragraph. A ruby inside a
        // header/footer/cell would need a different anchor — flag the loss
        // rather than mis-anchoring (full non-body round-trip is a backlog
        // item; same conservatism as the non-body textbox raw-set).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "ruby",
                Path: run.Path,
                Reason: "ruby (phonetic guide) inside a header/footer/table cell could not be serialized for round-trip; the base text is lost from the replayed document"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitBdoRun(DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-R42-9: a <w:bdo> (bidirectional override — forces the visual
        // RTL/LTR character ordering of its wrapped runs) has no scalar add/set
        // representation; the typed `add r` path drops the wrapper, losing the
        // load-bearing w:val direction. Re-insert the captured <w:bdo>…</w:bdo>
        // verbatim via a raw-set append, mirroring the ruby/pgNum raw-set
        // fallback (same shape: append the wrapper's outer XML to the just-
        // emitted host paragraph at /w:document/w:body/w:p[last()]).
        if (run.Type != "bdo") return false;
        var rawXml = run.Format.TryGetValue("_rawBdoXml", out var rx) ? rx?.ToString() : null;
        if (string.IsNullOrEmpty(rawXml))
            return true; // nothing to emit (shouldn't happen) — consumed anyway
        // The verbatim raw-set targets /w:document/w:body/w:p[last()]; only a
        // body host has that addressable last() paragraph. A bdo inside a
        // header/footer/cell would need a different anchor — flag the loss
        // rather than mis-anchoring (same conservatism as the ruby fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "bdo",
                Path: run.Path,
                Reason: "bidirectional override (w:bdo) inside a header/footer/table cell could not be serialized for round-trip; the wrapped runs survive flattened but the forced character ordering is lost from the replayed document"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitDirRun(DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-R43-7: a <w:dir> (bidirectional embedding — sets the bidi
        // embedding direction of its wrapped runs, distinct from <w:bdo> override
        // and the run-level <w:rtl> toggle) has no scalar add/set representation;
        // the typed `add r` path drops the wrapper, losing the load-bearing w:val
        // direction. Re-insert the captured <w:dir>…</w:dir> verbatim via a
        // raw-set append, mirroring the bdo fallback exactly.
        if (run.Type != "dir") return false;
        var rawXml = run.Format.TryGetValue("_rawDirXml", out var rx) ? rx?.ToString() : null;
        if (string.IsNullOrEmpty(rawXml))
            return true; // nothing to emit (shouldn't happen) — consumed anyway
        // Only a body host has an addressable last() paragraph; a dir inside a
        // header/footer/cell would mis-anchor — flag the loss instead (same
        // conservatism as the bdo/ruby fallbacks).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "dir",
                Path: run.Path,
                Reason: "bidirectional embedding (w:dir) inside a header/footer/table cell could not be serialized for round-trip; the wrapped runs survive flattened but the embedding direction is lost from the replayed document"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitPgNumRun(WordHandler word, DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-PGNUM: a run containing <w:pgNum/> (page-number placeholder)
        // has no scalar add/set representation — the typed `add r` path drops the
        // <w:pgNum/> entirely (text/other content survives but the placeholder
        // vanishes with no warning). Mirroring the rich-break / ruby raw-set
        // fallback, re-insert the verbatim <w:r> via a raw-set append so the
        // <w:pgNum/> — and any co-located <w:cr/> the typed path would demote to
        // <w:br/> — survive the round-trip. RunToNode stamps Format["_hasPgNum"].
        if (!run.Format.ContainsKey("_hasPgNum")) return false;
        // Only the /body host has the addressable last() paragraph anchor the
        // raw-set targets; a header/footer-hosted pgNum needs a different anchor
        // (backlog, same conservatism as the non-body ruby/textbox fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "pgNum",
                Path: run.Path,
                Reason: "page-number placeholder (w:pgNum) inside a header/footer/table cell could not be serialized for round-trip; the placeholder is lost from the replayed document"));
            return true;
        }
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return true; // nothing to emit
        // Same external-rel guard as the other raw-set fallbacks — a dangling
        // r:id/r:embed would not resolve in the rebuilt document.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "pgNum",
                Path: run.Path,
                Reason: "page-number placeholder run carries an external relationship reference and could not be serialized verbatim for round-trip"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitDateFieldRun(WordHandler word, DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-DATEFIELD: a run containing a Word date-component placeholder
        // (<w:dayLong/> / <w:dayShort/> / <w:monthLong/> / <w:monthShort/> /
        // <w:yearLong/> / <w:yearShort/>) has no scalar add/set representation —
        // the typed `add r` path persists GetRunText's "[dayLong]" human sentinel
        // as literal <w:t> text and drops the element entirely. Mirroring the
        // pgNum raw-set fallback, re-insert the verbatim <w:r> via a raw-set
        // append so the date element — and any co-located <w:t> text in the same
        // run — survive the round-trip. RunToNode stamps Format["_hasDateField"].
        if (!run.Format.ContainsKey("_hasDateField")) return false;
        // Only the /body host has the addressable last() paragraph anchor the
        // raw-set targets; a header/footer-hosted date field needs a different
        // anchor (backlog, same conservatism as the pgNum/ruby/textbox fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "dateField",
                Path: run.Path,
                Reason: "date-component placeholder (w:dayLong/w:monthLong/…) inside a header/footer/table cell could not be serialized for round-trip; the placeholder is lost from the replayed document"));
            return true;
        }
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return true; // nothing to emit
        // Same external-rel guard as the other raw-set fallbacks — a dangling
        // r:id/r:embed would not resolve in the rebuilt document.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "dateField",
                Path: run.Path,
                Reason: "date-component placeholder run carries an external relationship reference and could not be serialized verbatim for round-trip"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitHyphenRun(WordHandler word, DocumentNode run, string parentPath, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx, int hlBaseline = 0)
    {
        // BUG-DUMP-R40-3: a run containing <w:noBreakHyphen/> (non-breaking
        // hyphen) or <w:softHyphen/> (discretionary hyphen) — siblings of <w:t>
        // inside the run — has no scalar add/set representation. The typed
        // `add r`/`add p text=` path persists GetRunText's Unicode glyph
        // (U+2011 / U+00AD) as literal <w:t> text and drops the structural hyphen
        // element. Mirroring the pgNum/dateField raw-set fallback, re-insert the
        // verbatim <w:r> via a raw-set append so the hyphen element — and any
        // co-located <w:t> text in the same run — survive the round-trip.
        // RunToNode stamps Format["_hasHyphen"].
        if (!run.Format.ContainsKey("_hasHyphen")) return false;
        // BUG-DUMP-HYPHEN-CELL + BUG-DUMP-HYPHEN-RESERIALIZE: emit a structural
        // hyphen run (<w:noBreakHyphen/> / <w:softHyphen/>) as a typed
        // `add r --prop hyphen=noBreak|soft` for EVERY host — body, header,
        // footer, table cell. AddRun rebuilds the element via the SDK at apply
        // time (splitting `text` at the cached glyph in source order), so it
        // survives in any host.
        //
        // The OLD /body path used a raw-set append of the verbatim <w:r> against
        // /w:body/w:p[last()]. That inserted the element correctly, but when a
        // LATER `add r` in the SAME paragraph mutated p[last()], the SDK
        // re-serialized the paragraph and silently converted the raw-injected
        // <w:noBreakHyphen/> back to its U+2011 glyph (tester: 5/159 in a real
        // FRC report degraded this way). The typed `add r hyphen=` path has no
        // such re-serialization hazard — AddRun builds the element through the
        // live DOM in document order, like any other run. Dropping the raw-set
        // also removes the external-rel guard that path needed; a structural
        // hyphen run carries no relationship reference.
        string hyKind = run.Format.TryGetValue("_hasHyphen", out var hk)
            && string.Equals(hk?.ToString(), "soft", StringComparison.OrdinalIgnoreCase)
            ? "soft" : "noBreak";
        var hyProps = FilterEmittableProps(run.Format);
        hyProps.Remove("_hasHyphen");
        hyProps["hyphen"] = hyKind;
        if (!string.IsNullOrEmpty(run.Text)) hyProps["text"] = run.Text!;
        else hyProps.Remove("text");
        var hyParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = hyParent,
            Type = "r",
            Props = hyProps
        });
        return true;
    }

    private static bool TryEmitTabRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP14-02: tab-only run (<w:r><w:tab/></w:r>) surfaces as
        // type="tab" with empty Text. AddText splits "\t" into TabChar, so
        // emit `add r text="\t"` to round-trip the tab character.
        if (run.Type != "tab") return false;
        var tabParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        var tabProps = FilterEmittableProps(run.Format);
        tabProps["text"] = "\t";
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tabParent,
            Type = "r",
            Props = tabProps
        });
        return true;
    }

    private static bool TryEmitPtabRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-PTAB: ptab (positional tab) — Navigation surfaces its own run
        // type with align/relativeTo/leader on Format. Without an explicit
        // emit branch the runs filter would drop it and round-trip would
        // silently lose right-align/header-style tabs.
        if (run.Type != "ptab") return false;
        // BUG-DUMP-TABRPR: carry the ptab run's own rPr (font / size / szCs /
        // …) alongside its align/relativeTo/leader. Like a tab, a positional
        // tab paints a leader in the run's font and contributes to line
        // height, so its typography is meaningful — RunToNode keeps it on
        // run.Format and AddPtab applies it on replay.
        var ptabProps = FilterEmittableProps(run.Format);
        if (run.Format.TryGetValue("align", out var pAlign) && pAlign != null)
            ptabProps["alignment"] = pAlign.ToString() ?? "";
        if (run.Format.TryGetValue("relativeTo", out var pRel) && pRel != null)
            ptabProps["relativeTo"] = pRel.ToString() ?? "";
        if (run.Format.TryGetValue("leader", out var pLead) && pLead != null)
            ptabProps["leader"] = pLead.ToString() ?? "";
        var ptabParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = ptabParent,
            Type = "ptab",
            Props = ptabProps.Count > 0 ? ptabProps : null
        });
        return true;
    }

    // Build `add hyperlink` props from a source hyperlink node so a hyperlink
    // that carries no emittable runs (e.g. one wrapping only an <m:oMath>) can
    // still be materialized. Returns null when the hyperlink has no addressable
    // destination (neither url nor anchor) — the caller then routes the child
    // under the paragraph so its content still survives.
    private static Dictionary<string, string>? TryBuildHyperlinkAddProps(WordHandler word, string srcHlPath)
    {
        DocumentNode hl;
        try { hl = word.Get(srcHlPath); }
        catch { return null; }
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in new[] { "url", "anchor", "tooltip", "tgtFrame", "history" })
        {
            if (hl.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString();
                if (!string.IsNullOrEmpty(s)) props[key] = s!;
            }
        }
        return (props.ContainsKey("url") || props.ContainsKey("anchor")) ? props : null;
    }

    private static bool TryEmitEquationRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP7-03: inline <m:oMath> as paragraph child. Get surfaces it
        // as type="equation" with mode=inline and the LaTeX-ish formula in
        // Text. AddEquation accepts a paragraph parent for inline mode.
        // BUG-DUMP15-04: m:oMath inside w:hyperlink surfaces with a
        // hyperlink-scoped path (.../p[N]/hyperlink[K]/equation[M]). Strip
        // the trailing /equation[M] segment so the emitted Parent places the
        // equation INSIDE the hyperlink on replay.
        if (run.Type != "equation") return false;
        var eqMode = run.Format.TryGetValue("mode", out var emv) ? emv?.ToString() : "inline";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(eqMode) ? "inline" : eqMode!
        };
        // Always emit `formula` (even when empty); ToLatex may legitimately
        // return "" for minimal m:oMath.
        eqProps["formula"] = run.Text ?? "";
        // BUG-DUMP-EQVERBATIM: also carry the verbatim <m:oMath> so AddEquation
        // can restore the math EXACTLY — the LaTeX formula string drops every
        // math-run <w:rPr> (rFonts="Cambria Math" / sz) and simplifies some
        // structures, so a formatted equation rebuilt from the string alone
        // renders at the wrong font/size. AddEquation falls back to `formula`
        // when this is absent or unparseable, so the interactive path is unchanged.
        var eqXml = run.Format.TryGetValue("_omathXml", out var exv) ? exv?.ToString() : null;
        if (!string.IsNullOrEmpty(eqXml) && eqXml.Contains("oMath", StringComparison.Ordinal))
            eqProps["xml"] = eqXml;
        // Carry any OLE/preview-image parts referenced inside the verbatim math
        // (MathType/Equation objects) so they don't dangle on replay.
        AddMathInlinedPartProps(word, run.Path, eqProps);
        var eqParent = paraTargetPath;
        if (!string.IsNullOrEmpty(run.Path))
        {
            // R3-bt-1: the NodeBuilder now emits the resolvable `/oMath[N]`
            // segment for inline equations (was `/equation[N]`). Strip whichever
            // is present so the derived hyperlink-parent logic keeps working for
            // both the new paths and any legacy `/equation[` producer.
            var idxEq = run.Path.LastIndexOf("/oMath[", StringComparison.Ordinal);
            if (idxEq < 0)
                idxEq = run.Path.LastIndexOf("/equation[", StringComparison.Ordinal);
            if (idxEq > 0)
            {
                var derived = run.Path.Substring(0, idxEq);
                var hlIdx = derived.LastIndexOf("/hyperlink[", StringComparison.Ordinal);
                if (hlIdx > 0)
                {
                    // BUG-DBF-R3-02: the equation lives inside a <w:hyperlink>.
                    // A hyperlink is normally materialized via its child runs, but
                    // a hyperlink whose ONLY content is an <m:oMath> has no runs —
                    // so no `add hyperlink` row is emitted and replaying
                    // `add equation parent=…/hyperlink[K]` fails ("path not
                    // found"), dropping the math. Emit the missing `add hyperlink`
                    // (fetched from the source) before the equation so its parent
                    // exists; if the hyperlink can't be resolved, fall back to the
                    // paragraph so the math survives (the rare link wrapper is lost
                    // rather than the content).
                    var srcHlPath = run.Path.Substring(0, idxEq); // …/hyperlink[K]
                    var hlSeg = derived.Substring(hlIdx); // /hyperlink[K]
                    // BUG-DUMP-FIELDHL-XPARA: count hyperlink rows since THIS
                    // paragraph's own `add p` — paraTargetPath ("/…/p[last()]") is
                    // shared by every paragraph, so a global count let an earlier
                    // paragraph's hyperlink suppress the missing-hyperlink emit here
                    // and route the equation to a non-existent /hyperlink[K].
                    int eqLastParaAdd = items.FindLastIndex(it =>
                        it.Command == "add" && it.Type == "p");
                    int alreadyEmitted = 0;
                    for (int hi = eqLastParaAdd + 1; hi < items.Count; hi++)
                        if (items[hi].Type == "hyperlink"
                            && string.Equals(items[hi].Parent, paraTargetPath, StringComparison.Ordinal))
                            alreadyEmitted++;
                    var rebasedHl = paraTargetPath + hlSeg;
                    int wantK = 0;
                    var kStr = hlSeg.Length > 11 ? hlSeg[11..^1] : "";
                    int.TryParse(kStr, out wantK);
                    if (alreadyEmitted < wantK)
                    {
                        var hlProps = TryBuildHyperlinkAddProps(word, srcHlPath);
                        if (hlProps != null)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = paraTargetPath,
                                Type = "hyperlink",
                                Props = hlProps
                            });
                            eqParent = rebasedHl;
                        }
                        // else: leave eqParent = paraTargetPath (math survives).
                    }
                    else
                    {
                        eqParent = rebasedHl;
                    }
                }
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = eqParent,
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    // R14-bug1+2: legacy form field synth from CollapseFieldChains. The
    // begin run's <w:ffData> was unpacked onto ff*-prefixed Format keys
    // (ffName, ffType, ffDefault, ffMaxLength, ffChecked, ffItems,
    // ffHelpText, ffStatusText, ffEntryMacro, ffExitMacro, ffCalcOnExit,
    // ffTextType, ffTextFormat, ffCheckBoxSize, ffEnabled). Map back to
    // AddFormField's accepted keys (drop the `ff` prefix, lowercase the
    // first letter) so replay rebuilds the FieldChar + FormFieldData +
    // Bookmark chain with every original wrapper intact.
    private static bool TryEmitFormFieldRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        if (run.Type != "formfield") return false;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // CONSISTENCY(formfield-keys): AddFormField accepts type/formfieldtype,
        // name, default, maxLength, checked, items, helpText, statusText,
        // entryMacro, exitMacro, calcOnExit, textType, textFormat,
        // checkBoxSize, enabled. Mirror the AddFormField accepted vocabulary.
        foreach (var (k, v) in run.Format)
        {
            if (v == null) continue;
            if (!k.StartsWith("ff", StringComparison.OrdinalIgnoreCase)) continue;
            // Strip the "ff" prefix and lowercase the first remaining char so
            // ffName→name, ffDefault→default, etc.
            if (k.Length <= 2) continue;
            var bare = char.ToLowerInvariant(k[2]) + k.Substring(3);
            var sv = v.ToString();
            if (string.IsNullOrEmpty(sv)) continue;
            props[bare] = sv!;
        }
        // Field-run formatting forwarded by CollapseFieldChains (theme/literal
        // fonts, size, bold, …) — AddFormField stamps it on every field run.
        foreach (var (k, v) in run.Format)
        {
            if (v == null) continue;
            if (!FieldResultFormatKeys.Contains(k) || props.ContainsKey(k)) continue;
            var s2 = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s2.Length > 0) props[k] = s2;
        }
        // Preserve cached display text where AddFormField would otherwise
        // emit the placeholder symbol (text input) or "false" (checkbox).
        if (!props.ContainsKey("text"))
        {
            if (!string.IsNullOrEmpty(run.Text))
                props["text"] = run.Text!;
            else
                // Explicit empty pin: the source field has NO cached result
                // run; without the pin AddFormField fabricates an NBSP
                // placeholder and every empty form row gains a glyph.
                props["text"] = "";
        }
        // BUG-DUMP-FFCHECKBOX-BOOKMARK: the source field had no wrapping
        // bookmark (marked by EmitParagraph when no matching bookmark sibling
        // exists). Pin noBookmark so AddFormField does NOT fabricate one.
        if (run.Format.TryGetValue("_noBookmark", out var nbObj) && nbObj is bool nbB && nbB)
            props["noBookmark"] = "true";
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "formfield",
            Props = props,
        });
        return true;
    }

    private static bool TryEmitFieldRun(WordHandler word, DocumentNode run, string paraTargetPath, string parentPath, List<BatchItem> items, BodyEmitContext? ctx = null)
    {
        if (run.Type != "field") return false;
        // BUG-DUMP-R28-INCLUDEPICTURE: a marker-only synth emitted by
        // CollapseFieldChains when it decomposed a drawing-result field
        // (INCLUDEPICTURE). It carries the source paths of one contiguous run of
        // fldChar/instrText (and any text result) runs to round-trip verbatim via
        // raw-set, interleaved with the real `picture` nodes the same decompose
        // produced. The markers carry no relationships, so the append is safe.
        if (run.Format.TryGetValue("_fieldMarkerRaw", out var fmr) && fmr is bool fmrB && fmrB)
        {
            // BUG-DUMP-FLDSIMPLE-IMG: a fldSimple decomposed into a complex field
            // carries synthesized begin/instr/separate/end fldChar markers as inline
            // raw XML (no source slice paths exist for them). Append that verbatim.
            if (run.Format.TryGetValue("_markerInlineXml", out var mixObj)
                && mixObj is string mix && !string.IsNullOrEmpty(mix)
                && ResolveRawSetHost(parentPath, ctx) is { } inlineHost)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = inlineHost.Part,
                    Xpath = inlineHost.XPath,
                    Action = "append",
                    Xml = mix
                });
                return true;
            }
            var markerPaths = run.Format.TryGetValue("_markerSlicePaths", out var mspObj)
                ? mspObj as List<string> : null;
            if (markerPaths is { Count: > 0 } && ResolveRawSetHost(parentPath, ctx) is { } markerHost)
            {
                var sb = new System.Text.StringBuilder();
                foreach (var mp in markerPaths)
                {
                    // Entries starting with '<' are pre-synthesized verbatim XML
                    // (bookmarkStart/End interleaved with the markers — their
                    // node paths don't resolve here); the rest are source paths.
                    var mx = mp.StartsWith('<') ? mp : word.GetElementXml(mp);
                    if (!string.IsNullOrEmpty(mx)) sb.Append(mx);
                }
                if (sb.Length > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = markerHost.Part,
                        Xpath = markerHost.XPath,
                        Action = "append",
                        Xml = sb.ToString()
                    });
                }
            }
            return true;
        }
        // BUG-DUMP-R26-2: a field whose cached result has multiple distinctly-
        // formatted runs can't round-trip through `add field` (single-rPr model
        // collapses the runs and leaks the first run's bold onto the fldChar
        // markers). CollapseFieldChains flagged it and stashed the field-slice
        // run paths (begin..end). Raw-set the whole chain verbatim so per-run
        // formatting survives.
        // BUG-DUMP-R26-7: fire for /body, header/footer AND table-cell hosts
        // (ResolveRawSetHost), not only /body — previously a rich field result
        // in a header/footer/cell silently fell to the lossy typed emit.
        if (run.Format.TryGetValue("_richFieldResult", out var rfr) && rfr is bool rfrB && rfrB
            && run.Format.TryGetValue("_fieldSlicePaths", out var spObj) && spObj is string spStr
            && !string.IsNullOrEmpty(spStr)
            && ResolveRawSetHost(parentPath, ctx) is { } fieldHost)
        {
            var sb = new System.Text.StringBuilder();
            bool ok = true;
            var rfSlice = spStr.Split('\n').Where(p => !string.IsNullOrEmpty(p)).ToList();
            // BUG-DUMP-H78: a tracked deletion inside the result is a <w:del> sibling
            // BETWEEN field runs; per-path extraction resolves each path to its inner
            // <w:r> and strips the <w:del> wrapper. The emitter sets
            // _fieldSliceForceRange so we take the contiguous sibling-range extraction
            // (begin..end) directly, capturing the <w:del>/<w:delText> in place.
            bool forceRange = run.Format.TryGetValue("_fieldSliceForceRange", out var frv)
                && frv is bool frvB && frvB;
            if (forceRange && rfSlice.Count > 0)
            {
                ok = false; // skip per-path; fall to the range resolve below
            }
            else
            {
                foreach (var p in rfSlice)
                {
                    var xml = word.GetElementXml(p);
                    if (string.IsNullOrEmpty(xml)) { ok = false; break; }
                    sb.Append(xml);
                }
            }
            // BUG-DUMP-R56-NESTEDFORMFIELD: same bookmark-in-slice fragility as
            // the nested-field branch — fall back to a contiguous sibling-range
            // resolve when a per-child path (e.g. a bookmark indexed by w:id)
            // doesn't navigate.
            if (!ok && rfSlice.Count > 0)
            {
                var rangeXml = word.GetSiblingRangeXml(rfSlice[0], rfSlice[^1]);
                if (!string.IsNullOrEmpty(rangeXml))
                {
                    sb.Clear();
                    sb.Append(rangeXml);
                    ok = true;
                }
            }
            if (ok && sb.Length > 0)
            {
                var chainXml = sb.ToString();
                // BUG-DUMP-R26-7 (PART B): an external relationship inside the
                // cached result (e.g. a hyperlink r:id) would DANGLE in the
                // rebuilt part — the raw-set can't recreate the rel. Do NOT fall
                // silently to the lossy extractor; emit a deterministic warning
                // naming the loss. Full rel preservation is a separate effort.
                if (HasExternalRelRef(chainXml))
                {
                    ctx?.Warnings.Add(new DocxUnsupportedWarning(
                        Element: "field.richResult",
                        Path: run.Path,
                        Reason: "field cached result with per-run formatting AND an external relationship (hyperlink/image) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the formatted result is flattened to uniform text on replay"));
                    // Fall through to the typed path (preserves instruction +
                    // cached value with uniform formatting; loss is now visible).
                }
                else
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = fieldHost.Part,
                        Xpath = fieldHost.XPath,
                        Action = "append",
                        Xml = chainXml
                    });
                    return true;
                }
            }
            // Reconstruction failed — fall through to the typed path so the
            // field still round-trips (cached value + uniform formatting).
        }
        // BUG-DUMP-R26-7 (PART B): a single-run field result that wraps a
        // hyperlink (external rel) isn't "rich" (so the verbatim raw-set above
        // doesn't fire) but the typed path below still drops the hyperlink
        // wrapper + rel silently (text + bold survive). Emit a deterministic
        // warning so the loss is visible. (The multi-run rich+rel case is
        // already warned on the raw-set branch above, so gate on NOT-rich to
        // avoid a double warning.)
        if (run.Format.TryGetValue("_fieldResultHasExternalRel", out var frer)
            && frer is bool frerB && frerB
            && !(run.Format.TryGetValue("_richFieldResult", out var rfr2) && rfr2 is bool rfr2B && rfr2B))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "field.richResult",
                Path: run.Path,
                Reason: "field cached result wraps a hyperlink whose external relationship target is not carried through dump→batch; the hyperlink is flattened (its link is dropped) on replay"));
        }
        // Synthetic field entry from CollapseFieldChains. Format carries
        // `instruction` (raw fldSimple/instrText) and Text holds the cached
        // display value. AddField parses the instruction and rebuilds the
        // fldChar chain on replay.
        // BUG-DUMP18-02: w:fldSimple / fldChar-chain field inside w:hyperlink
        // should replay INSIDE the hyperlink — but only when a prior
        // `add hyperlink` row actually landed at the target paragraph
        // (BUG-DUMP9-03 fldSimple-only hyperlinks never surface a hyperlink
        // row, and routing the field there would fail path lookup on replay).
        if (run.Type != "field") return false;
        // R10-bug7: CollapseFieldChains flagged a nested field (IF/REF with
        // an inner DATE/PAGE/MERGEFIELD branch). AddField rebuilds a flat
        // begin/instr/separate/display/end chain and cannot model the
        // nested branches — emitting an `add field` row here would either
        // throw (parser sees garbage), drop the inner branches, OR merge
        // the inner instruction into the outer expression. Backlog item:
        // teach AddField to accept a tree representation. Until then, the
        // cheapest correct behavior is to flag the loss in envelope
        // warnings — same model as the OLE warning above — so callers
        // don't ship a doc with the IF false-branch silently stripped.
        // R10-bug8: malformed field (begin without matching end) surfaces as
        // a synth from CollapseFieldChains with _unmatchedFieldBegin=true.
        // Same warning model as _nestedField — preserve cached display,
        // flag the partial instruction in envelope.warnings.
        if (run.Format.TryGetValue("_unmatchedFieldBegin", out var ufbObj) && ufbObj is bool ufbB && ufbB)
        {
            if (ctx != null)
            {
                var partialInstr = run.Format.TryGetValue("instruction", out var pIv)
                    ? pIv?.ToString() ?? "" : "";
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.unmatched_begin",
                    Path: run.Path,
                    Reason: $"fldChar(begin) without matching end; partial instruction='{partialInstr}' dropped"));
            }
            if (!string.IsNullOrEmpty(run.Text))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["text"] = run.Text!
                    }
                });
            }
            return true;
        }
        if (run.Format.TryGetValue("_nestedField", out var nfObj) && nfObj is bool nfB && nfB)
        {
            // BUG-DUMP-R43-5: round-trip the nested field's begin..end run slice
            // verbatim via raw-set so the full fldChar/instrText sequence (and the
            // inner field) survives, instead of collapsing to cached display text.
            // Each slice run carries a resolvable source Path; concatenate their
            // OuterXml and append to the just-emitted host paragraph.
            // BUG-DUMP-R47-6: resolve the host via ResolveRawSetHost so a nested
            // field inside a HEADER/FOOTER/table cell round-trips too — not only
            // /body. A running-header with a nested IF/QUOTE/STYLEREF field (the
            // common "Chapter N — Title" header) was flattening to bare cached
            // text on the non-/body fallback, which changed the rendered header
            // height and cascaded body pagination across every page of the
            // section. Mirrors the _richFieldResult host resolution just above.
            var slicePaths = run.Format.TryGetValue("_nestedFieldSlicePaths", out var nfSpObj)
                ? nfSpObj as List<string> : null;
            if (slicePaths is { Count: > 0 } && ResolveRawSetHost(parentPath, ctx) is { } nestedHost)
            {
                var sb = new System.Text.StringBuilder();
                bool allResolved = true;
                foreach (var sp in slicePaths)
                {
                    var rx = word.GetElementXml(sp);
                    if (string.IsNullOrEmpty(rx)) { allResolved = false; break; }
                    sb.Append(rx);
                }
                // BUG-DUMP-R56-NESTEDFORMFIELD: the slice may interleave
                // bookmarkStart/bookmarkEnd children (the inner field carries a
                // form-field bookmark), whose query paths (indexed by w:id) don't
                // round-trip through NavigateToElement — per-child resolution then
                // bails to the lossy cached-text fallback, dropping the nested
                // FORMTEXT structure, its bookmark AND the field-run formatting
                // (bold/size/font). Resolve the contiguous begin..end sibling
                // range from the parent instead; it captures every interleaved
                // element verbatim regardless of individual path navigability.
                if (!allResolved)
                {
                    var rangeXml = word.GetSiblingRangeXml(slicePaths[0], slicePaths[^1]);
                    if (!string.IsNullOrEmpty(rangeXml))
                    {
                        sb.Clear();
                        sb.Append(rangeXml);
                        allResolved = true;
                    }
                }
                if (allResolved && sb.Length > 0)
                {
                    var nestedXml = sb.ToString();
                    // External rel inside the nested field (e.g. a hyperlink r:id)
                    // would dangle in the rebuilt part — the raw-set can't recreate
                    // the rel. Warn + fall through rather than emit a broken ref.
                    if (HasExternalRelRef(nestedXml))
                    {
                        ctx?.Warnings.Add(new DocxUnsupportedWarning(
                            Element: "field.nested",
                            Path: run.Path,
                            Reason: "nested field carrying an external relationship (hyperlink/image) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the inner field codes are dropped on replay"));
                    }
                    else
                    {
                        items.Add(new BatchItem
                        {
                            Command = "raw-set",
                            Part = nestedHost.Part,
                            Xpath = nestedHost.XPath,
                            Action = "append",
                            Xml = nestedXml
                        });
                        return true;
                    }
                }
            }
            // Fallback (non-body host or unresolvable slice): preserve the cached
            // display and flag the loss, as before.
            if (ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.nested",
                    Path: run.Path,
                    Reason: "nested field (begin inside a field's branch) inside a header/footer/table cell could not be serialized for round-trip; cached display preserved but inner field codes dropped"));
            }
            // Still emit the cached display so the paragraph isn't empty.
            if (!string.IsNullOrEmpty(run.Text))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["text"] = run.Text!
                    }
                });
            }
            return true;
        }
        var instr = run.Format.TryGetValue("instruction", out var iv)
            ? iv?.ToString() ?? "" : "";
        var fieldProps = BuildFieldAddProps(instr, run.Text ?? "");
        if (fieldProps != null
            && run.Format.TryGetValue("_noFieldSeparator", out var nfs)
            && nfs is bool nfsB && nfsB)
        {
            fieldProps["noSeparator"] = "true";
        }
        // BUG-DUMP-R37-4: forward the source field's lock bit so AddField
        // recreates it locked (begin fldChar @w:fldLock). Carried on the synth
        // Format by CollapseFieldChains (complex) / Navigation (fldSimple).
        if (fieldProps != null
            && run.Format.TryGetValue("fldLock", out var flkv) && flkv != null
            && string.Equals(flkv.ToString(), "true", StringComparison.OrdinalIgnoreCase))
        {
            fieldProps["fldLock"] = "true";
        }
        // BUG-DUMP-R24-2: source field had a separator but an empty cached
        // result. Pass an explicit empty `text` so AddField emits an empty
        // result run instead of fabricating a «name» placeholder.
        if (fieldProps != null
            && !fieldProps.ContainsKey("text")
            && run.Format.TryGetValue("_emptyFieldResult", out var efr)
            && efr is bool efrB && efrB)
        {
            fieldProps["text"] = "";
        }
        // BUG-R12A(BUG1): forward the captured cached-result-run formatting
        // (stashed under `_resultFmt.` by CollapseFieldChains) onto the `add
        // field` prop bag. AddField applies font/size/bold/color uniformly to
        // every rebuilt field run, so a bold/red/20pt PAGE field round-trips
        // styled instead of as a plain "1". Keys AddField can't express
        // (italic/underline/strike) are surfaced as a warning — the field still
        // round-trips (instruction + cached value preserved), only the extra
        // emphasis is lost; raw-set field passthrough is a backlog item.
        if (fieldProps != null)
        {
            var unsupportedFmt = new List<string>();
            foreach (var (k, v) in run.Format)
            {
                if (v == null) continue;
                if (!k.StartsWith("_resultFmt.", StringComparison.OrdinalIgnoreCase)) continue;
                var bare = k["_resultFmt.".Length..];
                if (!FieldAddSupportedFormatKeys.Contains(bare))
                {
                    unsupportedFmt.Add(bare);
                    continue;
                }
                // Map the literal face slots (font.latin/font.ascii/font.hAnsi)
                // onto AddField's uniform `font`. BUG-DUMP-FIELDHINT: font.hint is
                // NOT a face — it's the rFonts hint attribute; collapsing it wrote
                // a bogus font="eastAsia" face and dropped the hint, so the cached
                // result glyph (a GB3 ①②③) re-rendered in the Latin face and the
                // cell reflowed. Pass it through so AddField's per-slot loop applies
                // it via ApplyRunFormatting (RunFonts.Hint).
                var target = bare.Equals("font.hint", StringComparison.OrdinalIgnoreCase)
                    ? "font.hint"
                    : (bare.StartsWith("font.", StringComparison.OrdinalIgnoreCase) ? "font" : bare);
                if (!fieldProps.ContainsKey(target))
                {
                    var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                    if (s.Length > 0) fieldProps[target] = s;
                }
            }
            if (unsupportedFmt.Count > 0 && ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.resultFormat",
                    Path: run.Path,
                    Reason: $"cached field-result run formatting ({string.Join(", ", unsupportedFmt)}) cannot be expressed via add field; field instruction + bold/color/size/font preserved, extra emphasis dropped"));
            }
        }
        // BUG-DUMP-DELFIELD: forward the revision attribution that
        // CollapseFieldChains propagated from the <w:del>/<w:ins> wrapper onto
        // the synth. AddField wraps the rebuilt field chain in the matching
        // <w:del>/<w:ins> when it sees revision.type — so a deleted HYPERLINK
        // round-trips as a deletion (delInstrText + delText inside <w:del>)
        // rather than collapsing to live text. Forward to both the `add field`
        // and the plain-text fallback below.
        if (fieldProps != null)
        {
            foreach (var rk in new[] { "revision.type", "revision.author", "revision.date", "revision.id" })
            {
                if (fieldProps.ContainsKey(rk)) continue;
                if (run.Format.TryGetValue(rk, out var rv) && rv != null)
                {
                    var s = rv.ToString();
                    if (!string.IsNullOrEmpty(s)) fieldProps[rk] = s;
                }
            }
        }
        var fldParent = paraTargetPath;
        string? candidateHlParent = null;
        if (!string.IsNullOrEmpty(run.Path))
        {
            var idxFld = run.Path.LastIndexOf("/field[", StringComparison.Ordinal);
            if (idxFld > 0)
            {
                var derived = run.Path.Substring(0, idxFld);
                if (derived.Contains("/hyperlink["))
                    candidateHlParent = derived;
            }
        }
        // fldChar-chain fields surface with a flat /…/r[N] path; the
        // hyperlink hint is in Format._hyperlinkParent.
        if (candidateHlParent == null
            && run.Format.TryGetValue("_hyperlinkParent", out var fhlpObj)
            && fhlpObj != null)
        {
            var hint = fhlpObj.ToString();
            if (!string.IsNullOrEmpty(hint)) candidateHlParent = hint;
        }
        if (candidateHlParent != null)
        {
            // Re-base the candidate path onto paraTargetPath and verify a
            // prior `add hyperlink` row landed under that same paragraph.
            const string hlMarker = "/hyperlink[";
            var hlIdxStart = candidateHlParent.LastIndexOf(hlMarker, StringComparison.Ordinal);
            if (hlIdxStart > 0)
            {
                var hlEnd = candidateHlParent.IndexOf(']', hlIdxStart);
                if (hlEnd > hlIdxStart)
                {
                    var kStr = candidateHlParent.Substring(hlIdxStart + hlMarker.Length,
                        hlEnd - hlIdxStart - hlMarker.Length);
                    if (int.TryParse(kStr, out var kIdx))
                    {
                        var rebased = paraTargetPath
                            + candidateHlParent.Substring(hlIdxStart);
                        // BUG-DUMP-FIELDHL-XPARA: paraTargetPath is the literal
                        // "/…/p[last()]" — IDENTICAL for every paragraph (dump always
                        // targets the most-recently-added p). Counting hyperlink rows
                        // by Parent==paraTargetPath therefore tallied hyperlinks from
                        // ALL prior paragraphs, so a field-ONLY hyperlink (no separate
                        // display run → no `add hyperlink` row of its own) inherited a
                        // phantom "hyperlink exists" from an unrelated earlier
                        // paragraph and routed the field to /hyperlink[1] that doesn't
                        // exist in THIS paragraph — "Path not found" on replay dropped
                        // the whole REF/PAGEREF field and its visible text. Count only
                        // hyperlink rows emitted SINCE the current paragraph's own
                        // `add p` boundary so the tally is paragraph-local.
                        int lastParaAdd = items.FindLastIndex(it =>
                            it.Command == "add" && it.Type == "p");
                        int emittedHls = 0;
                        for (int hi = lastParaAdd + 1; hi < items.Count; hi++)
                            if (items[hi].Type == "hyperlink"
                                && string.Equals(items[hi].Parent, paraTargetPath, StringComparison.Ordinal))
                                emittedHls++;
                        if (emittedHls >= kIdx)
                            fldParent = rebased;
                    }
                }
            }
        }
        if (fieldProps != null)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = fldParent,
                Type = "field",
                Props = fieldProps
            });
        }
        else if (!string.IsNullOrEmpty(run.Text))
        {
            // Unparseable instruction — fall back to plain text so the
            // paragraph still renders the cached value rather than going empty.
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = fldParent,
                Type = "r",
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["text"] = run.Text! }
            });
        }
        return true;
    }

    private static bool TryEmitPictureRun(WordHandler word, DocumentNode run, string paraTargetPath, string parentPath, int targetIndex, List<BatchItem> items, BodyEmitContext? ctx, string? sharedAttachPara = null)
    {
        // Drawing-bearing runs surface as type="picture" regardless of
        // whether the Drawing wraps an image (Blip) or a chart (c:chart).
        // Try the image path first; if no embedded image part the run is a
        // chart anchor — pull the next pre-resolved ChartSpec and emit a
        // typed `add chart` row.
        // Drawings wrapped in <mc:AlternateContent>/<mc:Choice> surface as a
        // plain "run" node (Run.GetFirstChild<Drawing>() returns null because
        // the Drawing lives inside the AlternateContent wrapper), so we also
        // accept "run" / "r" when the raw XML carries an obvious textbox
        // marker. Non-drawing runs without those markers short-circuit out
        // of the textbox/picture path immediately.
        if (run.Type != "picture")
        {
            if (run.Type != "run" && run.Type != "r") return false;
            var probeXml = word.GetElementXml(run.Path);
            if (string.IsNullOrEmpty(probeXml)) return false;
            bool isTextbox = IsTextboxDrawing(probeXml);
            // A genuine text/hyperlink run (no drawing payload at all) belongs
            // to the plain-run path — short-circuit out so EmitPlainOrHyperlink
            // run handles it. Only drawing-bearing runs continue here.
            if (!isTextbox && !IsDrawingBearingRun(probeXml)) return false;
            // BUG-R3 (dump emits `add textbox` into a table cell): see the
            // sibling site below — a cell-hosted textbox must attach to its
            // containing paragraph (paraTargetPath inside the cell), never as a
            // direct <w:tc> child. Some editors export textboxes wrapped in
            // <mc:AlternateContent>, so they reach this "run"-typed branch.
            string? attachParaAc = sharedAttachPara
                ?? (parentPath.Contains("/tc[", StringComparison.Ordinal) ? paraTargetPath : null);
            if (isTextbox && TryEmitTextbox(word, run, probeXml, parentPath, items, ctx, attachParaAc, paraTargetPath))
                return true;
            // AlternateContent-wrapped non-textbox shapes — connector lines
            // (<wps:cNvCnPr>, e.g. a letterhead separator), autoshapes, groups
            // — and textboxes whose typed emit failed fall back to a raw-set
            // append so the shape survives round-trip instead of vanishing.
            // BUG-DUMP-R47-1: the host now resolves via ResolveRawSetHost
            // (body / header / footer / table cell), mirroring the wps:wsp
            // DrawingML-shape path below — a cell/header/footer-hosted legacy
            // VML/non-textbox shape was previously dropped because this gate
            // hardcoded parentPath == "/body" and the body anchor. The external-rel
            // guards are unchanged: drawings carrying an r:embed/r:id we can't
            // re-anchor still fall through to the warn+drop below.
            // BUG-R1-01 (preserved): ResolveRawSetHost("/body") returns exactly
            // ("/document", "/w:document/w:body/w:p[last()]") — attach to the
            // most-recently-emitted paragraph via last(), NOT the navigation
            // pIndex (targetIndex). pIndex deliberately does not advance for an
            // oMathPara-bearing paragraph (EmitBody: display equations resolve as
            // /body/oMathPara[N], not /body/p[N]), but in the LITERAL XML the
            // equation is a real <w:p> wrapping <m:oMathPara>. So a numeric
            // w:p[{pIndex}] under-counts by one per preceding equation and the
            // shape lands on the wrong paragraph. The host paragraph was just
            // added by this same EmitParagraph call, so it is always the last
            // w:p at replay time — the same last()-relative attach the typed
            // picture/textbox path uses. So body-hosted shapes round-trip
            // byte-identically to the previous hardcoded anchor.
            if (!probeXml.Contains("r:embed") && !probeXml.Contains("r:id")
                && !DrawingHasUnreconstructableRel(probeXml)
                && ctx != null
                && ResolveRawSetHost(parentPath, ctx) is { } drawHost)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = drawHost.Part,
                    Xpath = drawHost.XPath,
                    Action = "append",
                    Xml = probeXml
                });
                return true;
            }
            // BUG-DUMP-R47-3: a legacy VML IMAGE pict (<w:pict><v:shape
            // type="#_x0000_t75"><v:imagedata r:id>) — NOT a textbox, so the
            // IsVmlTextbox path above never fired — carries its bitmap through an
            // image-part rel. The no-rel raw-set just above skips it (it DOES carry
            // r:id), so it used to fall straight into the warn-drop below and the
            // image vanished from the rebuild (a full-width Gantt/diagram image
            // disappearing shifts body pagination by several pages). Ship it through
            // the same inlined-parts vmlshape carrier the rel-bearing VML-textbox
            // path uses: GetVmlShapeEmitData inlines the v:imagedata image part(s)
            // and AddVmlShape rewrites the rel id on replay. Fires for any
            // pict-bearing run whose references all resolve.
            if (probeXml.Contains("<w:pict", StringComparison.Ordinal)
                && word.GetVmlShapeEmitData(run.Path) is { } vmlImgData)
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = sharedAttachPara ?? paraTargetPath,
                    Type = "inlinedparts",
                    Props = PackInlinedPartsProps(vmlImgData),
                });
                return true;
            }
            // BUG-DUMP-WPG-GROUP: an mc:AlternateContent-wrapped DrawingML group
            // (<wpg:wgp> group of pictures) or shape surfaces as a plain "run"
            // node — the Drawing lives inside the AltContent so there is no typed
            // picture node, the no-rel raw-set above skipped it (its blips carry
            // r:embed), and it is a <w:drawing> not a <w:pict> so the VML carrier
            // skipped it too. It then fell into the warn-drop below and the WHOLE
            // group (every nested image) vanished. Ship it through the inlined-
            // parts carrier: GetDrawingShapeEmitData inlines every referenced
            // image part and rewrites the rel ids on replay, so the group drawing
            // round-trips verbatim. Mirrors the wps:wsp shape carrier in the
            // type=="picture" branch below. GuardCarrierContentTypes returns null
            // for a drawing referencing an unsupported part (a chart, …), so those
            // correctly fall through to the warn-drop instead of being mis-routed.
            if (probeXml.Contains("<w:drawing", StringComparison.Ordinal)
                && word.GetDrawingShapeEmitData(run.Path) is { } grpData)
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = sharedAttachPara ?? paraTargetPath,
                    Type = "inlinedparts",
                    Props = PackInlinedPartsProps(grpData),
                });
                return true;
            }
            // Drawing-bearing but not safely raw-set-able inline (lives in a
            // header/footer/cell, or carries an external relationship we can't
            // re-anchor). Flag the loss rather than silently dropping it.
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                "drawing", run.Path,
                "non-textbox drawing/shape could not be serialized for round-trip; it will be missing from the replayed document"));
            return true;
        }
        // BUG-DUMP-R31-4: a run whose drawing is a <wps:wsp> DrawingML SHAPE
        // (preset geometry + outline + spPr fill — e.g. an ellipse with a
        // blipFill image and a red a:ln) surfaces as type="picture" because
        // GetImageBinary finds the blipFill's embedded image. Routed through the
        // `add picture` path below it is FLATTENED to a plain <pic:pic> rect: the
        // prstGeom, the a:ln outline and the shape semantics are all lost. A
        // wps:wsp shape must round-trip AS A SHAPE — raw-set its drawing verbatim
        // so geometry + outline + fill structure survive — not be downgraded to a
        // picture. Distinguish a genuine inline <pic:pic> (keep the picture path)
        // from a wps:wsp shape (raw-set passthrough). Mirrors the VML-textbox /
        // non-textbox-shape raw-set convention.
        {
            var shapeXml = word.GetElementXml(run.Path);
            // BUG-DUMP-WPG-GROUP: a DrawingML GROUP (<wpg:wgp> — multiple pictures/
            // shapes grouped, often mc:AlternateContent-wrapped) surfaces as
            // type="picture" because GetImageBinary finds the FIRST nested blip.
            // The picture path below would flatten the whole group to that single
            // <pic:pic>, dropping every other grouped image AND the group structure.
            // Route the group through the inlined-parts carrier instead, so all
            // nested image parts + the verbatim group drawing round-trip. Mirrors
            // the wps:wsp shape carrier just below; GuardCarrierContentTypes nulls
            // out a group referencing an unsupported part so it falls through.
            if (!string.IsNullOrEmpty(shapeXml)
                && shapeXml.Contains("<wpg:wgp", StringComparison.Ordinal)
                && word.GetDrawingShapeEmitData(run.Path) is { } wpgData)
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = sharedAttachPara ?? paraTargetPath,
                    Type = "inlinedparts",
                    Props = PackInlinedPartsProps(wpgData),
                });
                return true;
            }
            if (!string.IsNullOrEmpty(shapeXml)
                && IsWpsShapeDrawing(shapeXml)
                && ctx != null
                && ResolveRawSetHost(parentPath, ctx) is { } shapeHost)
            {
                // The shape's blipFill references an embedded image part via
                // r:embed; a verbatim raw-set would dangle it. Ship the run
                // through the inlined-parts carrier (verbatim runXml +
                // part{N} image bytes, rel ids rewritten on replay) so the
                // bitmap fill survives — same shape as the vmlshape carrier.
                // A wps:wsp with NO external rel (plain fill) raw-sets verbatim.
                if (HasExternalRelRef(shapeXml)
                    && word.GetDrawingShapeEmitData(run.Path) is { } shpData)
                {
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "inlinedparts",
                        Props = PackInlinedPartsProps(shpData),
                    });
                    return true;
                }
                // Fallback (unresolvable reference): scrub the blipFill rel
                // (neutral solidFill placeholder keeps the shape valid) and
                // warn that the image bitmap is dropped — geometry and outline
                // still round-trip.
                var scrubbed = ScrubDrawingBlipFillRels(shapeXml, out var blipDropped);
                if (blipDropped)
                    ctx.Warnings.Add(new DocxUnsupportedWarning(
                        Element: "shape.blipFill",
                        Path: run.Path,
                        Reason: "DrawingML shape (wps:wsp) image fill (a:blipFill r:embed) dropped — shape-image round-trip is not supported; the shape's geometry and outline are preserved and the fill was replaced with a neutral placeholder to keep the rebuilt drawing valid"));
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = shapeHost.Part,
                    Xpath = shapeHost.XPath,
                    Action = "append",
                    Xml = scrubbed
                });
                return true;
            }
        }
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
            // BUG-DUMP-R28-1: the picture node's width/height come from
            // CreateImageNode formatted as 1-decimal CENTIMETRES (EMU /
            // EmuPerCmF, "F1"). Replaying that cm string through ParseEmu snaps
            // cx/cy back to a 360000-EMU (0.1cm) grid, shifting every inline
            // drawing by up to ~17800 EMU. The cumulative drift over many
            // drawings moves where page breaks land — the whole document's
            // text reflows vertically. Emit the EXACT cx/cy straight from
            // <wp:extent> as "<emu>emu" so ParseEmu reconstructs the original
            // value byte-for-byte (mirrors the textbox path, which already
            // emits raw EMU). AddPicture applies the same parsed EMU to both
            // <wp:extent> and the matching <a:ext> in <a:xfrm>. Covers inline
            // drawings in /body AND inside table cells (same emit path).
            var picXml = word.GetElementXml(run.Path);
            if (!string.IsNullOrEmpty(picXml))
            {
                var extMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"wp:extent\b[^>]*\bcx=""(\d+)""[^>]*\bcy=""(\d+)""");
                if (extMatch.Success)
                {
                    picProps["width"] = extMatch.Groups[1].Value + "emu";
                    picProps["height"] = extMatch.Groups[2].Value + "emu";
                }
                // BUG-DUMP-R29-1: Word's inline-picture layout HEIGHT depends on
                // <wp:effectExtent l/t/r/b> (the drawing's visual overflow/effect
                // margin) even when <wp:extent> is identical. The rebuild
                // hardcoded effectExtent to 0/0/0/0 (CreateImageRun /
                // CreateAnchorImageRun), so each affected drawing rendered
                // ~35px shorter, pulling downstream content up across page
                // boundaries — a visible document-wide vertical drift. Capture
                // the source l/t/r/b as a single "l,t,r,b" EMU prop so AddPicture
                // can restore it byte-for-byte. Absent (or all-zero) effectExtent
                // is the interactive default and needs no prop.
                var eeMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"wp:effectExtent\b[^>]*\bl=""(-?\d+)""[^>]*\bt=""(-?\d+)""[^>]*\br=""(-?\d+)""[^>]*\bb=""(-?\d+)""");
                if (eeMatch.Success
                    && !(eeMatch.Groups[1].Value == "0" && eeMatch.Groups[2].Value == "0"
                         && eeMatch.Groups[3].Value == "0" && eeMatch.Groups[4].Value == "0"))
                {
                    picProps["effectExtent"] =
                        $"{eeMatch.Groups[1].Value},{eeMatch.Groups[2].Value}," +
                        $"{eeMatch.Groups[3].Value},{eeMatch.Groups[4].Value}";
                }
                // BUG-DUMP-R39-1: an anchored (floating) picture positioned with
                // an absolute offset stores it as <wp:positionH><wp:posOffset>EMU.
                // CreateImageNode emits Format["hPosition"]/["vPosition"] as
                // 1-decimal CENTIMETRES (EmuPerCmF, "F1"). Replaying that cm string
                // through AddPicture's ParseEmu snaps the offset back to a
                // 360000-EMU (0.1cm) grid — e.g. posOffset 1234567 -> 1224000,
                // visibly shifting the floating image. Mirror the R28/R38 raw-EMU
                // pattern: pull the EXACT posOffset straight from the source XML,
                // scoped to its own positionH/positionV block so H->hPosition and
                // V->vPosition map correctly, and emit "<emu>emu" so ParseEmu
                // reconstructs the original offset byte-for-byte. Only override
                // when a posOffset is present — anchors using <wp:align>
                // (left/center/right) carry no posOffset, and inline pictures have
                // no positionH/V at all, so the regex simply won't match (safe).
                var hPosMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"<wp:positionH\b[^>]*>.*?<wp:posOffset>(-?\d+)</wp:posOffset>.*?</wp:positionH>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (hPosMatch.Success)
                    picProps["hPosition"] = hPosMatch.Groups[1].Value + "emu";
                var vPosMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"<wp:positionV\b[^>]*>.*?<wp:posOffset>(-?\d+)</wp:posOffset>.*?</wp:positionV>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (vPosMatch.Success)
                    picProps["vPosition"] = vPosMatch.Groups[1].Value + "emu";
                // BUG-DUMP-R45-4: the reconstructor rebuilds <a:blip>/<pic:spPr>
                // from a FIXED subset, silently dropping image-level visual
                // content the source carried — recolor/alpha inside <a:blip>
                // (duotone/biLevel/alphaModFix) and the spPr drop-shadow
                // (<a:effectLst><a:outerShdw>). Mirror the chart verbatim-capture
                // pattern: grab them as verbatim XML props so AddPicture can
                // re-inject at schema-correct positions. Only set each prop when
                // the source actually has that content (a plain picture stays
                // plain — no spurious empty effectLst / blip children).
                var blipInner = CapturePicBlipInnerXml(picXml);
                if (!string.IsNullOrEmpty(blipInner))
                    picProps["blipEffects"] = blipInner!;
                // BUG-R13C consistency: CapturePicBlipInnerXml/StripRelReferencingBlipExts
                // drops the SVG companion (<asvg:svgBlip r:embed=…>) because its dangling
                // relationship would abort the whole `add picture`; the PNG raster
                // fallback still renders, so content is conserved but the VECTOR layer is
                // lost. Surface that as a warning — mirroring the theme-image / OLE
                // fallback-drop warnings — instead of dropping it silently.
                if (picXml.Contains("svgBlip", StringComparison.Ordinal))
                    ctx?.Warnings.Add(new DocxUnsupportedWarning(
                        "picture", run.Path,
                        "SVG vector layer (svgBlip) dropped on round-trip; PNG raster fallback preserved"));
                // BUG-DUMP-H82: StripRelReferencingBlipExts drops ANY <a:ext> that
                // references a relationship, not only the SVG companion — most
                // notably the Office 2010 artistic-effect extension
                // (<a14:imgProps><a14:imgLayer r:embed=…><a14:imgEffect>…) whose
                // r:embed points at a `.wdp` (HD Photo) backing layer. That drop was
                // silent (the svgBlip warning above didn't match), losing the
                // editable effect + its .wdp source with no signal. Warn for any
                // rel-referencing ext drop that ISN'T the already-warned svgBlip,
                // mirroring the drop-but-warn model for lossy media extensions.
                else if (BlipHasRelReferencingExt(picXml))
                    ctx?.Warnings.Add(new DocxUnsupportedWarning(
                        "picture", run.Path,
                        "image effect layer (e.g. Office artistic effect / HD-photo .wdp backing layer) dropped on round-trip; raster image preserved"));
                var spEffectLst = CapturePicSpPrEffectLst(picXml);
                if (!string.IsNullOrEmpty(spEffectLst))
                    picProps["spEffects"] = spEffectLst!;
                // The fixed spPr rebuild also drops xfrm flip flags
                // (<a:xfrm flipH="1"> — mirrored logos), a content extent
                // (<a:ext>) that legitimately differs from the frame's
                // wp:extent, bwMode, and explicit <a:noFill/>/<a:ln> blocks.
                // Capture the whole <pic:spPr> verbatim; AddPicture swaps its
                // rebuilt spPr for this block (and then skips the narrower
                // spEffects injection — the effectLst already rides inside).
                var spPrWhole = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"<pic:spPr[^>]*?>.*?</pic:spPr>|<pic:spPr[^>]*?/>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (spPrWhole.Success)
                    picProps["spPrXml"] = spPrWhole.Value;
                // Anchor wrap distances (distT/distB/distL/distR — the gap
                // between a floating image and the text wrapping around it).
                // CreateAnchorImageRun hardcodes T/B=0, L/R=114300; a figure
                // with asymmetric distances shifted every adjacent line.
                // Capture as "T,B,L,R" so AddPicture restores them. Inline
                // pictures have no <wp:anchor>, so the match simply won't fire.
                var anchorMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml, @"<wp:anchor\b([^>]*)>");
                if (anchorMatch.Success)
                {
                    string DistAttr(string n) =>
                        System.Text.RegularExpressions.Regex.Match(
                            anchorMatch.Groups[1].Value, n + "=\"(\\d+)\"") is { Success: true } mm
                            ? mm.Groups[1].Value : "0";
                    var wt = DistAttr("distT"); var wb = DistAttr("distB");
                    var wl = DistAttr("distL"); var wr = DistAttr("distR");
                    if (!(wt == "0" && wb == "0" && wl == "114300" && wr == "114300"))
                        picProps["wrapDist"] = $"{wt},{wb},{wl},{wr}";
                }
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "picture",
                Props = picProps
            });
            return true;
        }

        // Only consume a ChartSpec if the run is genuinely a chart. Picture-
        // typed runs that aren't images can also be background images, OLE
        // objects, SmartArt, watermark anchors etc — falling through
        // unconditionally would misalign chart positions.
        if (ctx != null && word.IsChartRun(run.Path)
            && ctx.ChartCursor.Index < ctx.ChartSpecs.Count)
        {
            var spec = ctx.ChartSpecs[ctx.ChartCursor.Index];
            ctx.ChartCursor.Index++;
            // VERBATIM-FIRST: carry the chart part + its sidecars byte-for-byte
            // instead of rebuilding from semantic props. The typed BuildChartProps
            // path below de-references the chart data (numRef→numLit, drops strRef
            // category labels / ptCount data points / dLbls / externalData) and
            // renders a visibly compressed chart. The verbatim <w:drawing> also
            // preserves the host wrapper (wp:extent / effectExtent / anchor) for
            // free, so the R38-1 / anchor width fix-ups below are unnecessary on
            // this path. Falls through to the typed path when the carrier can't
            // resolve every referenced part (return null) — same conservative
            // fallback as the other inlined-parts carriers.
            if (word.GetChartVerbatimEmitData(run.Path) is { } chartVerbatim)
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "inlinedparts",
                    Props = PackInlinedPartsProps(chartVerbatim),
                });
                return true;
            }
            var chartProps = BuildChartProps(spec);
            // BUG-DUMP-R38-1: the chart node's width/height come from
            // WordHandler.Query formatted as 1-decimal CENTIMETRES (cx/cy /
            // EmuPerCmF, "F1"). Replaying that cm string through ParseEmu in
            // AddChart snaps the <wp:extent cx cy> back to a 360000-EMU (0.1cm)
            // grid — e.g. cx=5486400 (15.24cm) -> 5472000 (15.2cm), a 14400-EMU
            // width loss that visibly reflows the chart's title/bars/labels.
            // Emit the EXACT cx/cy straight from <wp:extent> as "<emu>emu" so
            // ParseEmu reconstructs the original value byte-for-byte. Mirrors
            // the inline-picture path (R28) above. effectExtent is out of scope:
            // AddChart accepts no effectExtent prop and these chart anchors
            // carry all-zero effectExtent.
            var chartXml = word.GetElementXml(run.Path);
            if (!string.IsNullOrEmpty(chartXml))
            {
                var chartExtMatch = System.Text.RegularExpressions.Regex.Match(
                    chartXml,
                    @"wp:extent\b[^>]*\bcx=""(\d+)""[^>]*\bcy=""(\d+)""");
                if (chartExtMatch.Success)
                {
                    chartProps["width"] = chartExtMatch.Groups[1].Value + "emu";
                    chartProps["height"] = chartExtMatch.Groups[2].Value + "emu";
                }
                // A chart wrapped in <wp:anchor> is a FLOATING chart — capture
                // its wrap + position so AddChart rebuilds the anchor instead of
                // flattening it to an inline frame (which drops it entirely on
                // replay, since only inline charts ever got a spec). Inline
                // charts have no <wp:anchor>, so none of these fire. Mirrors the
                // floating-picture capture above.
                CaptureChartAnchorProps(chartXml!, chartProps);
            }
            // A chart may carry a c:userShapes overlay drawing (a logo / photo /
            // annotation drawn on top of the chart in Word's chart editor) on a
            // ChartDrawingPart. AddChart rebuilds the chart from scratch, so
            // without shipping that part the overlay vanishes. Pack it (verbatim
            // chartshapes XML + embedded images) under a `userShapes.` prefix;
            // AddChart re-creates the part and the <c:userShapes> reference.
            var usData = word.GetChartUserShapesEmitData(run.Path);
            if (usData != null)
            {
                chartProps["userShapesXml"] = usData.RunXml;
                int upi = 0;
                foreach (var part in usData.Parts)
                {
                    upi++;
                    chartProps[$"userShapes.part{upi}.relId"] = part.RelId;
                    chartProps[$"userShapes.part{upi}.data"] =
                        $"data:{part.ContentType};base64,{System.Convert.ToBase64String(part.Bytes)}";
                }
                int uei = 0;
                foreach (var ext in usData.Externals)
                {
                    uei++;
                    chartProps[$"userShapes.ext{uei}.relId"] = ext.RelId;
                    chartProps[$"userShapes.ext{uei}.type"] = ext.Type;
                    chartProps[$"userShapes.ext{uei}.target"] = ext.Target;
                }
            }
            // BUG-DUMP-CHART-SIDECARS: carry the native chart's sidecar parts
            // (chartStyle / chartColorStyle / themeOverride / embedded data
            // workbook) so AddChart re-attaches them instead of rebuilding a
            // chart stripped of its theme, custom colours, and editable data.
            var sidecars = word.GetChartSidecarEmitData(run.Path);
            if (sidecars != null)
            {
                foreach (var (role, ct, bytes) in sidecars)
                {
                    // One part per role for native charts (style/colors/
                    // themeOverride/package each appear at most once).
                    chartProps[$"sidecar.{role}.data"] =
                        $"data:{ct};base64,{System.Convert.ToBase64String(bytes)}";
                }
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "chart",
                Props = chartProps
            });
            return true;
        }
        // Drawing without image part and not a chart — most likely a wps
        // shape. BUG-DUMP-TXBX: textbox-bearing drawings get a typed
        // `add textbox` row plus recursive inner-paragraph/run emits so
        // round-trip preserves structure (raw-set fallback was emitting
        // BOTH the full <w:drawing> XML AND flattening the textbox's
        // inner runs back onto the host paragraph). Non-textbox shapes
        // still fall through to the raw-set append.
        var rawXml = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(rawXml) && IsTextboxDrawing(rawXml))
        {
            // BUG-R3 (dump emits `add textbox` into a table cell): a textbox is
            // a drawing carried by a run inside a paragraph, never a direct
            // <w:tc> child — `add textbox` with a bare cell parent is rejected
            // ("table cells only accept paragraphs, tables, or SDTs"). When the
            // host is a cell, attach the textbox to the run's containing
            // paragraph (paraTargetPath, which lives inside the cell) so
            // AddTextbox's paragraph-attach path (ResolveDrawingHostFromParagraph)
            // anchors it correctly — mirrors the body single-drawing path. Use
            // any already-computed sharedAttachPara (side-by-side multi-drawing
            // host) first; otherwise fall back to paraTargetPath for cells.
            string? attachPara = sharedAttachPara
                ?? (parentPath.Contains("/tc[", StringComparison.Ordinal) ? paraTargetPath : null);
            if (TryEmitTextbox(word, run, rawXml, parentPath, items, ctx, attachPara, paraTargetPath))
                return true;
        }
        // BUG-R3 (linked external image): a <w:drawing> carrying an
        // unreconstructable relationship (linked-image r:link, SmartArt
        // r:dm/r:lo/r:qs/r:cs) must NOT be raw-set verbatim — its relationship
        // target isn't recreated, so the replayed drawing would dangle and fail
        // [Semantic] validation. Drop it cleanly and surface the loss (mirrors
        // the SmartArt/non-textbox warn-and-skip on the AlternateContent path
        // above) instead of falling through to a silent raw-set or a silent
        // return.
        // SmartArt: the diagram's data/layout/quickStyle/colors parts (plus the
        // data part's rendered-drawing child) ship base64-inlined in a
        // self-contained `add diagram`, mirroring the `add activex` carrier —
        // previously the whole diagram was warn-dropped as unreconstructable.
        if (!string.IsNullOrEmpty(rawXml) && rawXml.Contains("relIds")
            && word.GetDiagramEmitData(run.Path) is { } dgmData)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "inlinedparts",
                Props = PackInlinedPartsProps(dgmData),
            });
            return true;
        }
        if (!string.IsNullOrEmpty(rawXml) && DrawingHasUnreconstructableRel(rawXml))
        {
            string detail = rawXml.Contains("r:link")
                ? "linked external image (<a:blip r:link>) references an external relationship that is not carried through dump→batch; the image will be missing from the replayed document"
                : "drawing references a relationship (SmartArt/diagram) that cannot be reconstructed; it will be missing from the replayed document";
            ctx?.Warnings.Add(new DocxUnsupportedWarning("drawing", run.Path, detail));
            return true;
        }
        if (!string.IsNullOrEmpty(rawXml) &&
            parentPath == "/body" &&
            !rawXml.Contains("r:embed") && !rawXml.Contains("r:id"))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/document",
                // BUG-R1-01 (second site, plain <w:drawing> wps shape): same
                // last()-relative attach as above — pIndex under-counts literal
                // w:p when a display equation precedes this shape's paragraph.
                Xpath = "/w:document/w:body/w:p[last()]",
                Action = "append",
                Xml = rawXml
            });
        }
        return true;
    }

    // SmartArt (diagram) drawings reference their data / layout / colors /
    // quickStyle parts via r:dm / r:lo / r:qs / r:cs — relationships, but NOT
    // r:embed/r:id. The dump never reconstructs those parts, so raw-setting the
    // drawing verbatim leaves dangling relationships: the SDK validator NREs in
    // RelationshipTypeConstraint and real Word refuses to open the file ("may be
    // corrupt"). Treat such drawings as unreconstructable so the caller flags a
    // loss warning instead of emitting a corrupt file. (Plain shapes/connectors
    // carry no relationships and still round-trip via raw-set.)
    private static bool DrawingHasUnreconstructableRel(string xml) =>
        xml.Contains("r:dm") || xml.Contains("r:lo") ||
        xml.Contains("r:qs") || xml.Contains("r:cs") ||
        // BUG-R3 (linked external image): <a:blip r:link="rIdN"> references an
        // EXTERNAL image relationship (TargetMode="External"), not an embedded
        // image part. The dump never recreates that relationship, so raw-setting
        // the drawing verbatim leaves a dangling r:link → [Semantic] "the
        // relationship referenced by r:link does not exist". Full linked-image
        // round-trip (carrying the external rel) is a separate feature; treat
        // the drawing as unreconstructable so the raw-set sites skip it cleanly
        // and the caller surfaces a loss warning (mirrors the SmartArt path).
        xml.Contains("r:link");

    // Capture a floating chart's <wp:anchor> wrap + positioning into props that
    // AddChart/BuildChartFrame consume (anchor / wrap / hposition / vposition /
    // halign / valign / hrelative / vrelative / behindtext / relativeHeight /
    // effectExtent / wrapDist). No-op when the drawing is an inline chart (no
    // <wp:anchor>). Regex-scoped per positionH/positionV block so H→hposition
    // and V→vposition map correctly. Mirrors the floating-picture capture.
    private static void CaptureChartAnchorProps(string chartXml, Dictionary<string, string> props)
    {
        var anchorMatch = System.Text.RegularExpressions.Regex.Match(chartXml, @"<wp:anchor\b([^>]*)>");
        if (!anchorMatch.Success) return;
        props["anchor"] = "true";
        var attrs = anchorMatch.Groups[1].Value;

        string Attr(string scope, string n) =>
            System.Text.RegularExpressions.Regex.Match(scope, n + "=\"(-?\\d+)\"") is { Success: true } m
                ? m.Groups[1].Value : "0";

        if (System.Text.RegularExpressions.Regex.Match(attrs, "behindDoc=\"(1|true)\"").Success)
            props["behindtext"] = "true";
        var rh = System.Text.RegularExpressions.Regex.Match(attrs, "relativeHeight=\"(\\d+)\"");
        if (rh.Success) props["relativeHeight"] = rh.Groups[1].Value;

        var wt = Attr(attrs, "distT"); var wb = Attr(attrs, "distB");
        var wl = Attr(attrs, "distL"); var wr = Attr(attrs, "distR");
        if (!(wt == "0" && wb == "0" && wl == "114300" && wr == "114300"))
            props["wrapDist"] = $"{wt},{wb},{wl},{wr}";

        var ee = System.Text.RegularExpressions.Regex.Match(
            chartXml,
            @"wp:effectExtent\b[^>]*\bl=""(-?\d+)""[^>]*\bt=""(-?\d+)""[^>]*\br=""(-?\d+)""[^>]*\bb=""(-?\d+)""");
        if (ee.Success && !(ee.Groups[1].Value == "0" && ee.Groups[2].Value == "0"
            && ee.Groups[3].Value == "0" && ee.Groups[4].Value == "0"))
            props["effectExtent"] = $"{ee.Groups[1].Value},{ee.Groups[2].Value},{ee.Groups[3].Value},{ee.Groups[4].Value}";

        // Wrap mode — which wp:wrap* child is present.
        props["wrap"] =
            chartXml.Contains("<wp:wrapSquare") ? "square" :
            chartXml.Contains("<wp:wrapTight") ? "tight" :
            chartXml.Contains("<wp:wrapThrough") ? "through" :
            chartXml.Contains("<wp:wrapTopAndBottom") ? "topandbottom" :
            chartXml.Contains("<wp:wrapNone") ? "none" : "square";

        // Per-axis relativeFrom + posOffset/align, scoped to each block.
        var hBlock = System.Text.RegularExpressions.Regex.Match(
            chartXml, @"<wp:positionH\b[^>]*>.*?</wp:positionH>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        var vBlock = System.Text.RegularExpressions.Regex.Match(
            chartXml, @"<wp:positionV\b[^>]*>.*?</wp:positionV>",
            System.Text.RegularExpressions.RegexOptions.Singleline);

        void AxisProps(System.Text.RegularExpressions.Match block, string relKey, string posKey, string alignKey)
        {
            if (!block.Success) return;
            var rel = System.Text.RegularExpressions.Regex.Match(block.Value, "relativeFrom=\"([^\"]+)\"");
            if (rel.Success) props[relKey] = rel.Groups[1].Value;
            var align = System.Text.RegularExpressions.Regex.Match(block.Value, @"<wp:align>([^<]+)</wp:align>");
            if (align.Success) { props[alignKey] = align.Groups[1].Value; return; }
            var off = System.Text.RegularExpressions.Regex.Match(block.Value, @"<wp:posOffset>(-?\d+)</wp:posOffset>");
            if (off.Success) props[posKey] = off.Groups[1].Value + "emu";
        }
        AxisProps(hBlock, "hrelative", "hposition", "halign");
        AxisProps(vBlock, "vrelative", "vposition", "valign");
    }

    // BUG-DUMP-R45-4: capture the inner XML of the FIRST <a:blip> inside a
    // picture's drawing (the recolor/alpha children — duotone / biLevel /
    // alphaModFix / lum* / clrChange — that AddPicture's fixed blip rebuild
    // drops). The r:embed is an ATTRIBUTE of <a:blip>, not a child, so it is
    // preserved automatically by AddPicture and excluded here. Returns null
    // when the blip is self-closing or empty (a plain picture stays plain).
    private static string? CapturePicBlipInnerXml(string picXml)
    {
        // Match the first non-self-closing <a:blip …>…</a:blip> and grab the
        // inner content. A self-closing <a:blip … /> has no children → skip.
        var m = System.Text.RegularExpressions.Regex.Match(
            picXml,
            @"<a:blip\b[^>]*?>(.*?)</a:blip>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!m.Success) return null;
        var inner = m.Groups[1].Value.Trim();
        inner = StripRelReferencingBlipExts(inner);
        return inner.Length > 0 ? inner : null;
    }

    // BUG-R13C: an <a:blip>'s <a:extLst> can carry an <a:ext> whose child
    // references an external package relationship — most commonly the SVG
    // companion <asvg:svgBlip r:embed="rIdN"/> (a PNG fallback + SVG original).
    // The dump inlines only the raster <a:blip> source (the base64 `src`); the
    // SVG part and its relationship are not carried through dump→batch, so the
    // r:embed="rIdN" both dangles AND, injected as a bare fragment, fails to
    // parse ("'r' is an undeclared prefix") — the whole `add picture` step
    // aborts and the image is lost. Drop any <a:ext> block that references a
    // relationship (r:embed / r:link / r:id); the raster fallback still renders.
    // Exts with no relationship reference (e.g. a14:useLocalDpi) are kept.
    // BUG-DUMP-H82: true when the picture's first <a:blip> carries an <a:ext>
    // that references a relationship (r:embed/r:link/r:id) — i.e. an extension
    // that StripRelReferencingBlipExts will drop. Used to warn on the non-svgBlip
    // drops (Office 2010 artistic effects + their .wdp backing layer) that were
    // previously stripped silently. Scoped to the <a:blip> inner so the main
    // r:embed attribute on <a:blip> (carried through by AddPicture) is not counted.
    private static bool BlipHasRelReferencingExt(string picXml)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            picXml, @"<a:blip\b[^>]*?>(.*?)</a:blip>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!m.Success) return false;
        foreach (System.Text.RegularExpressions.Match ext in
                 System.Text.RegularExpressions.Regex.Matches(
                     m.Groups[1].Value, @"<a:ext\b[^>]*>.*?</a:ext>",
                     System.Text.RegularExpressions.RegexOptions.Singleline))
            if (System.Text.RegularExpressions.Regex.IsMatch(ext.Value, @"\br:(embed|link|id)\s*="))
                return true;
        return false;
    }

    private static string StripRelReferencingBlipExts(string blipInner)
    {
        if (string.IsNullOrEmpty(blipInner)
            || !blipInner.Contains("<a:ext", StringComparison.Ordinal))
            return blipInner;
        var cleaned = System.Text.RegularExpressions.Regex.Replace(
            blipInner,
            @"<a:ext\b[^>]*>.*?</a:ext>",
            mm => System.Text.RegularExpressions.Regex.IsMatch(mm.Value, @"\br:(embed|link|id)\s*=")
                ? string.Empty
                : mm.Value,
            System.Text.RegularExpressions.RegexOptions.Singleline);
        // Drop a now-empty <a:extLst></a:extLst> (or self-closed) so a plain
        // raster blip emits no spurious empty wrapper.
        cleaned = System.Text.RegularExpressions.Regex.Replace(
            cleaned, @"<a:extLst>\s*</a:extLst>", string.Empty,
            System.Text.RegularExpressions.RegexOptions.Singleline);
        return cleaned.Trim();
    }

    // BUG-DUMP-R45-4: capture the verbatim <a:effectLst>…</a:effectLst> sitting
    // inside the picture's <pic:spPr> (the drop-shadow / glow / reflection that
    // AddPicture's fixed spPr rebuild drops). Returns null when no effectLst is
    // present so a plain picture stays plain.
    private static string? CapturePicSpPrEffectLst(string picXml)
    {
        // Scope to the <pic:spPr> block first so we don't accidentally grab an
        // effectLst from an unrelated sibling (none today, but keeps the regex
        // honest if spPr structure grows).
        var spMatch = System.Text.RegularExpressions.Regex.Match(
            picXml,
            @"<pic:spPr\b[^>]*?>(.*?)</pic:spPr>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        var scope = spMatch.Success ? spMatch.Groups[1].Value : picXml;
        var m = System.Text.RegularExpressions.Regex.Match(
            scope,
            @"<a:effectLst\b[^>]*?>.*?</a:effectLst>|<a:effectLst\b[^>]*?/>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        return m.Success ? m.Value : null;
    }

    // BUG-DUMP-R26-6: a legacy VML textbox is a <w:pict> carrying a
    // <v:textbox> (with <w:txbxContent>) or a <v:shape type="#_x0000_t202">
    // (the VML textbox preset). Distinct from the modern DrawingML textbox
    // (wps:wsp/wps:txbx), which the typed `add textbox` path handles. We
    // detect VML so it can round-trip verbatim via raw-set instead of being
    // force-converted (and emptied) through the DrawingML emit.
    private static bool IsVmlTextbox(string rawXml)
    {
        if (!rawXml.Contains("<w:pict", StringComparison.Ordinal)
            && !rawXml.Contains("<v:", StringComparison.Ordinal))
            return false;
        return rawXml.Contains("<v:textbox", StringComparison.Ordinal)
            || rawXml.Contains("_x0000_t202", StringComparison.Ordinal);
    }

    private static bool IsTextboxDrawing(string rawXml)
    {
        // A wpg:wgp GROUP (e.g. a `diagram`, whose node shapes are themselves
        // textboxes) is NOT a plain textbox. Classifying it as one makes the
        // textbox-only-paragraph shortcut flatten the whole group down to its
        // first child textbox, dropping every other node + connector. Let a
        // group fall through to the general drawing path, which round-trips the
        // whole <w:drawing> verbatim (raw-set) since a native diagram carries no
        // external relationship references.
        if (rawXml.Contains("<wpg:wgp", StringComparison.Ordinal)) return false;

        // Mirrors WordHandler.CountTextboxesInHost / Navigation's textbox
        // selector — a textbox is a wps:wsp with txBox=1 cNvSpPr or a
        // wps:txbx child carrying w:txbxContent.
        // BUG-DUMP-TXBXCONTENT-LITERAL (H31 family): the third probe must match
        // the <w:txbxContent> ELEMENT open tag, not the bare token — a run whose
        // visible text literally contains "txbxContent" (docs describing OOXML
        // textbox internals) would otherwise be misrouted through the textbox
        // drawing path, which extracts no drawing payload and silently drops the
        // plain text. The sibling clauses already use element-anchored forms.
        return rawXml.Contains("txBox=\"1\"")
            || rawXml.Contains("<wps:txbx")
            || System.Text.RegularExpressions.Regex.IsMatch(rawXml, @"<\w*:?txbxContent[\s/>]");
    }

    /// <summary>
    /// True when a run carries any drawing/shape payload (image, chart,
    /// textbox, connector line, autoshape, group) — i.e. a
    /// <c>&lt;w:drawing&gt;</c>, legacy VML <c>&lt;w:pict&gt;</c>, or an
    /// <c>&lt;mc:AlternateContent&gt;</c>/<c>&lt;wps:&gt;</c> wrapper around one.
    /// Used to decide whether a non-textbox run still needs preserving via a
    /// raw-set append rather than being treated as a plain text run.
    /// </summary>
    private static bool IsDrawingBearingRun(string rawXml)
    {
        return rawXml.Contains("<w:drawing")
            || rawXml.Contains("<w:pict")
            || rawXml.Contains("<mc:AlternateContent")
            || rawXml.Contains("<wps:");
    }

    /// <summary>
    /// BUG-DUMP-R31-4: True when a run's drawing is a modern DrawingML SHAPE
    /// (<c>&lt;wps:wsp&gt;</c> with preset geometry / spPr) rather than a plain
    /// inline <c>&lt;pic:pic&gt;</c> picture. Such a shape may carry a blipFill
    /// image (so it surfaces as type="picture"), but its geometry + outline are
    /// shape semantics that the picture-emit path would flatten away. Textboxes
    /// (which carry their own typed/raw-set path) are excluded — they are handled
    /// upstream by IsTextboxDrawing.
    /// </summary>
    private static bool IsWpsShapeDrawing(string rawXml)
    {
        if (string.IsNullOrEmpty(rawXml)) return false;
        if (IsTextboxDrawing(rawXml)) return false;          // textbox has its own path
        // A drawing whose graphicData is a CHART must take the typed chart
        // emit further down — claiming it here raw-set the chart's r:id
        // verbatim with no chart part behind it, producing a file real Word
        // refuses to open.
        if (rawXml.Contains("drawingml/2006/chart", StringComparison.Ordinal)) return false;
        if (!rawXml.Contains("<wps:wsp", StringComparison.Ordinal)) return false;
        // A genuine shape carries a preset/custom geometry or its own shape
        // properties — that's what the picture path cannot represent.
        return rawXml.Contains("prstGeom", StringComparison.Ordinal)
            || rawXml.Contains("custGeom", StringComparison.Ordinal)
            || rawXml.Contains("<wps:spPr", StringComparison.Ordinal);
    }

    /// <summary>
    /// BUG-DUMP-R31-4: replace every <c>&lt;a:blipFill&gt;</c> whose blip
    /// carries a relationship reference (r:embed / r:link) with a neutral
    /// <c>&lt;a:solidFill&gt;&lt;a:srgbClr val="FFFFFF"/&gt;&lt;/a:solidFill&gt;</c>
    /// placeholder. The dump cannot carry the referenced image binary + part
    /// rels through a raw-set, so a verbatim passthrough would dangle the rel and
    /// corrupt the rebuilt drawing. Scrubbing keeps the shape a schema-valid
    /// filled shape (geometry + outline intact) at the cost of the image bitmap.
    /// Mirrors StripDanglingThemeBlipRefs in WordBatchEmitter.Resources.cs.
    /// </summary>
    private static string ScrubDrawingBlipFillRels(string xml, out bool dropped)
    {
        dropped = false;
        if (string.IsNullOrEmpty(xml) || !xml.Contains("blipFill", StringComparison.Ordinal))
            return xml;
        if (!xml.Contains(":embed=", StringComparison.Ordinal)
            && !xml.Contains(":link=", StringComparison.Ordinal))
            return xml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (doc.Root == null) return xml;
            var aNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
            var rNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            foreach (var blipFill in doc.Descendants(aNs + "blipFill").ToList())
            {
                var blip = blipFill.Element(aNs + "blip");
                var hasRel = blip != null && blip.Attributes().Any(a => a.Name.Namespace == rNs);
                if (!hasRel) continue;
                var placeholder = new System.Xml.Linq.XElement(aNs + "solidFill",
                    new System.Xml.Linq.XElement(aNs + "srgbClr",
                        new System.Xml.Linq.XAttribute("val", "FFFFFF")));
                blipFill.ReplaceWith(placeholder);
                dropped = true;
            }
            if (!dropped) return xml;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return xml;
        }
    }

    /// <summary>
    /// BUG-DUMP-TXBX: emit a typed <c>add textbox</c> row for the host of
    /// the current drawing run, followed by recursive inner-paragraph/run
    /// emits under <c>/&lt;host&gt;/textbox[N]</c>. Geometry props
    /// (width/height/wrap/anchor.x/anchor.y/fill) are extracted from the
    /// raw drawing XML so the rebuilt textbox keeps its layout.
    /// </summary>
    private static bool TryEmitTextbox(WordHandler word, DocumentNode run, string rawXml,
                                       string parentPath, List<BatchItem> items, BodyEmitContext? ctx,
                                       string? attachParaPath = null, string? paraTargetPath = null)
    {
        if (ctx == null) return false;

        // BUG-DUMP-TEXTBOX-INDEX-DESYNC: the typed `add textbox` path reads source
        // inner content via word.Get on the SOURCE textbox index, which must track
        // Navigation's source /<host>/textbox[N]. Navigation counts ONLY a
        // <w:drawing> whose inner XML carries `<wps:txbx` or `txBox="1"` (a
        // DrawingML textbox) — NOT a bare legacy <w:pict><v:textbox> VML box. So a
        // verbatim path bumps the SOURCE ordinal only when its rawXml matches that
        // same rule; bumping for a bare-VML box (which Navigation does not index)
        // would itself desync the following typed textbox's read. (TextboxCounters —
        // the REBUILD index for the emit target — counts only typed `add textbox`
        // rows and is deliberately NOT bumped here.) Keyed by parentPath, matching
        // the typed path's host key.
        void BumpSourceTextboxOrdinalForVerbatim()
        {
            // BUG-DUMP-TBLORDINAL-TEXTBOX: a textbox shipped VERBATIM (raw-set /
            // inlined-parts carrier) carries its <w:txbxContent> tables WITHOUT
            // going through EmitTable, so EmitTable's `++TableOrdinalBox` never
            // fires for them — yet the later `(//w:tbl)[N]` cell raw-set selectors
            // count ALL tables in document order (including textbox-nested ones).
            // Leaving the ordinal short made every following table's selector land
            // N tables early, so a cell-content raw-set targeted the wrong table —
            // here a tr[57] cell-merge XPath hit a 7-row table and the cell text
            // was dropped. Bump the ordinal by the shipped XML's table count so the
            // `(//w:tbl)` numbering stays in lockstep with replay. Mirrors the
            // EmitSdt carrier's identical adjustment. (Unconditional — any shipped
            // table must count, regardless of whether the box is Navigation-indexed.)
            int tblCount = System.Text.RegularExpressions.Regex
                .Matches(rawXml, "<w:tbl[ >]").Count;
            if (tblCount > 0) ctx.TableOrdinalBox[0] += tblCount;

            bool navigationCountsIt = rawXml.Contains("<w:drawing", StringComparison.Ordinal)
                && (rawXml.Contains("<wps:txbx", StringComparison.Ordinal)
                    || rawXml.Contains("txBox=\"1\"", StringComparison.Ordinal));
            if (!navigationCountsIt) return;
            ctx.SourceTextboxCounters[parentPath] =
                (ctx.SourceTextboxCounters.TryGetValue(parentPath, out var _pv) ? _pv : 0) + 1;
        }

        // BUG-DUMP-R26-6: a LEGACY VML textbox (<w:pict> with <v:shape
        // type="#_x0000_t202"> / <v:textbox><w:txbxContent>) is a different
        // shape family than the modern DrawingML box `add textbox` produces.
        // The typed emit below parses DrawingML namespaces (wp:/wps:/a:) that
        // don't exist in VML, so it extracted ZERO props (no text, fill, or
        // stroke) and Navigation can't surface the VML txbxContent under
        // /<host>/textbox[N] for the recursive inner-content emit — the box came
        // back as an empty modern textbox with its content + fillcolor/
        // strokecolor gone. The faithful (and lossless) round-trip is a verbatim
        // raw-set of the whole <w:pict> run into the just-emitted host paragraph
        // — mirrors the non-textbox-shape / rich-inline-SDT raw-set append.
        // BUG-DUMP-R26-7: fire for /body, header/footer AND table-cell hosts
        // (ResolveRawSetHost), not only /body — VML page-number / watermark
        // boxes in headers are the most common real case.
        if (IsVmlTextbox(rawXml) && ResolveRawSetHost(parentPath, ctx) is { } vmlHost)
        {
            // BUG-DUMP-R26-7 (PART B): a VML shape with an external relationship
            // (e.g. <v:imagedata r:id> referencing an image part) would dangle in
            // the rebuilt part. Don't silently flatten — emit a deterministic
            // warning naming the loss, then fall through. Plain VML textboxes
            // (fillcolor/strokecolor, no r:id) still round-trip verbatim.
            if (HasExternalRelRef(rawXml))
            {
                // The rel refs are resolvable (hyperlinks inside the textbox
                // content, v:imagedata image parts): ship them through the
                // inlined-parts carrier so the shape round-trips with its
                // relationships recreated — previously warn-dropped.
                var vmlParent = attachParaPath ?? paraTargetPath;
                var vmlData = vmlParent != null ? word.GetVmlShapeEmitData(run.Path) : null;
                if (vmlData != null)
                {
                    BumpSourceTextboxOrdinalForVerbatim();
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = vmlParent,
                        Type = "inlinedparts",
                        Props = PackInlinedPartsProps(vmlData),
                    });
                    return true;
                }
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "textbox.vmlContent",
                    Path: run.Path,
                    Reason: "legacy VML shape with an external relationship (image r:id / linked content) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the shape content is dropped on replay"));
            }
            else
            {
                BumpSourceTextboxOrdinalForVerbatim();
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = vmlHost.Part,
                    Xpath = vmlHost.XPath,
                    Action = "append",
                    Xml = rawXml
                });
                return true;
            }
        }

        // BUG-DUMP-TEXTBOX-IMG: a MODERN DrawingML textbox shape (wps:wsp +
        // txbxContent, often mc:AlternateContent/wpg-wrapped — e.g. a letterhead
        // shape pairing a caption box with a logo) that ALSO carries an embedded
        // picture (<a:blip r:embed>) loses that image on the typed `add textbox`
        // path below, which extracts only geometry + text. The image binary then
        // vanishes from the rebuild (this is the wpg-group image-loss class: a
        // page of grouped letterhead shapes silently dropping their logos). Route
        // such a shape through the inlined-parts carrier so the embedded image
        // part + verbatim shape XML (box + text + picture) round-trip, mirroring
        // the VML carrier above. Plain textboxes (no embedded image) keep the
        // typed `add textbox` path so they stay cleanly editable.
        if (rawXml.Contains("r:embed", StringComparison.Ordinal)
            && (attachParaPath ?? paraTargetPath) is { } tbImgParent
            && word.GetDrawingShapeEmitData(run.Path) is { } tbImgData)
        {
            BumpSourceTextboxOrdinalForVerbatim();
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = tbImgParent,
                Type = "inlinedparts",
                Props = PackInlinedPartsProps(tbImgData),
            });
            return true;
        }

        // Only emit a typed `add textbox` for hosts AddTextbox itself
        // supports: /body, /body/tbl[..]/tc[N], /header[N], /footer[N].
        // Other parents fall through to the raw-set append.
        string hostPath = parentPath;
        if (!IsTextboxHostPath(hostPath)) return false;

        // BUG-D1-MULTIDRAWING-HOST: when N textboxes share a source
        // paragraph (side-by-side card layout), attach each to the same
        // already-emitted host paragraph (attachParaPath = /body/p[last()])
        // instead of /body — otherwise AddTextbox creates a fresh host per
        // textbox and the side-by-side layout fans out into N stacked
        // paragraphs. The textbox INDEX still scopes to hostPath so
        // /body/textbox[K] addressing remains continuous across the doc.
        string emitParent = attachParaPath ?? hostPath;

        // Allocate the REBUILD textbox index (n): only typed `add textbox` rows
        // count, because that is what the replayed SET ops can address.
        int n = ctx.TextboxCounters.TryGetValue(hostPath, out var prev) ? prev + 1 : 1;
        ctx.TextboxCounters[hostPath] = n;
        string textboxPath = hostPath == "/" ? "/textbox[" + n + "]" : $"{hostPath}/textbox[{n}]";

        // BUG-DUMP-TEXTBOX-INDEX-DESYNC: allocate the SOURCE textbox index (sourceN)
        // separately — it counts EVERY textbox (verbatim + typed) so it matches
        // Navigation's source /<host>/textbox[N]. The inner-content recursion below
        // READS from sourceReadPath (sourceN) but EMITS into textboxPath (n). When
        // no verbatim sibling precedes this textbox the two indices coincide
        // (sourceN == n) and behaviour is unchanged.
        int sourceN = (ctx.SourceTextboxCounters.TryGetValue(hostPath, out var sprev) ? sprev : 0) + 1;
        ctx.SourceTextboxCounters[hostPath] = sourceN;
        string sourceReadPath = hostPath == "/" ? "/textbox[" + sourceN + "]" : $"{hostPath}/textbox[{sourceN}]";

        // Extract geometry / wrap / fill / anchor from the drawing XML so the
        // rebuilt textbox keeps its layout. Conservative best-effort — any
        // attribute we can't parse falls back to AddTextbox's defaults.
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(rawXml);
            System.Xml.Linq.XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            var anchor = doc.Descendants(wp + "anchor").FirstOrDefault()
                      ?? (System.Xml.Linq.XElement?)doc.Descendants(wp + "inline").FirstOrDefault();
            var extent = doc.Descendants(wp + "extent").FirstOrDefault();
            if (extent != null)
            {
                var cx = extent.Attribute("cx")?.Value;
                var cy = extent.Attribute("cy")?.Value;
                if (!string.IsNullOrEmpty(cx)) props["width"] = cx + "emu";
                if (!string.IsNullOrEmpty(cy)) props["height"] = cy + "emu";
            }
            if (anchor != null)
            {
                var posHEl = anchor.Element(wp + "positionH");
                var posVEl = anchor.Element(wp + "positionV");
                var posH = posHEl?.Element(wp + "posOffset")?.Value;
                var posV = posVEl?.Element(wp + "posOffset")?.Value;
                // A floating drawing positions each axis EITHER by an absolute
                // posOffset OR by a relative <wp:align> (left/center/right /
                // top/bottom). Capture the align form too — without it a
                // center-aligned textbox round-trips as posOffset=0 (i.e. jumps
                // hard left). posOffset wins when both somehow appear.
                var alignH = posHEl?.Element(wp + "align")?.Value;
                var alignV = posVEl?.Element(wp + "align")?.Value;
                if (!string.IsNullOrEmpty(posH)) props["anchor.x"] = posH + "emu";
                else if (!string.IsNullOrEmpty(alignH)) props["hAlign"] = alignH;
                if (!string.IsNullOrEmpty(posV)) props["anchor.y"] = posV + "emu";
                else if (!string.IsNullOrEmpty(alignV)) props["vAlign"] = alignV;
                // Anchor reference frames. AddTextbox hardcodes column/paragraph;
                // without forwarding these a textbox anchored relativeFrom="page"
                // round-trips as relativeFrom="paragraph" and floats off-position
                // (the source posOffset is measured from a different origin).
                var hRel = posHEl?.Attribute("relativeFrom")?.Value;
                var vRel = posVEl?.Attribute("relativeFrom")?.Value;
                if (!string.IsNullOrEmpty(hRel)) props["hRelative"] = hRel;
                if (!string.IsNullOrEmpty(vRel)) props["vRelative"] = vRel;
                // wrap token
                if (anchor.Element(wp + "wrapSquare") != null) props["wrap"] = "square";
                else if (anchor.Element(wp + "wrapTight") != null) props["wrap"] = "tight";
                else if (anchor.Element(wp + "wrapTopAndBottom") != null) props["wrap"] = "topAndBottom";
                else if (anchor.Element(wp + "wrapNone") != null) props["wrap"] = "none";
            }
            var spPr = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "spPr");
            // Geometry preset (rect default). roundRect etc. otherwise reverted
            // to a sharp rectangle on rebuild.
            var prst = spPr?.Element(a + "prstGeom")?.Attribute("prst")?.Value;
            if (!string.IsNullOrEmpty(prst) && prst != "rect") props["geometry"] = prst;
            // Rotation (<a:xfrm rot>, 60000ths of a degree). Round-trips raw.
            var rot = spPr?.Element(a + "xfrm")?.Attribute("rot")?.Value;
            if (!string.IsNullOrEmpty(rot) && rot != "0") props["rotation"] = rot;
            // Fill: solidFill (with optional alpha) or gradFill inside wps:spPr.
            var solidFill = spPr?.Element(a + "solidFill");
            var srgbClr = solidFill?.Element(a + "srgbClr");
            if (srgbClr?.Attribute("val")?.Value is string fillHex && !string.IsNullOrEmpty(fillHex))
            {
                props["fill"] = fillHex;
                var alpha = srgbClr.Element(a + "alpha")?.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(alpha)) props["fill.opacity"] = alpha;
            }
            else
            {
                var grad = spPr?.Element(a + "gradFill");
                if (grad != null)
                {
                    var stops = grad.Element(a + "gsLst")?.Elements(a + "gs")
                        .Select(gs => (gs.Element(a + "srgbClr")?.Attribute("val")?.Value, gs.Attribute("pos")?.Value))
                        .Where(t => !string.IsNullOrEmpty(t.Item1))
                        .Select(t => $"{t.Item1}@{t.Item2 ?? "0"}");
                    if (stops != null)
                    {
                        var joined = string.Join(";", stops);
                        if (!string.IsNullOrEmpty(joined)) props["fill.gradient"] = joined;
                    }
                }
            }
            // Shadow (<a:effectLst><a:outerShdw>): emit the compact tuple the
            // add path understands so the drop shadow round-trips faithfully.
            var shdw = spPr?.Element(a + "effectLst")?.Element(a + "outerShdw");
            if (shdw != null)
            {
                var sClr = shdw.Element(a + "srgbClr");
                props["shadow"] = string.Join(";",
                    shdw.Attribute("blurRad")?.Value ?? "50800",
                    shdw.Attribute("dist")?.Value ?? "38100",
                    shdw.Attribute("dir")?.Value ?? "5400000",
                    sClr?.Attribute("val")?.Value ?? "000000",
                    sClr?.Element(a + "alpha")?.Attribute("val")?.Value ?? "40000");
            }
            // Line / border (<a:ln>): width (EMU, round-trips through ParseEmu),
            // solidFill color, and dash style. Without this the textbox outline
            // was dropped entirely on dump→batch — borders vanished and content
            // reflowed. <a:noFill/> means the box explicitly has no border.
            var ln = spPr?.Element(a + "ln");
            if (ln != null)
            {
                if (ln.Element(a + "noFill") != null)
                {
                    props["line.style"] = "none";
                }
                else
                {
                    var lnW = ln.Attribute("w")?.Value;
                    if (!string.IsNullOrEmpty(lnW)) props["line.width"] = lnW;
                    var lnClr = ln.Element(a + "solidFill")?.Element(a + "srgbClr")?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(lnClr)) props["line.color"] = lnClr;
                    // a:prstDash@val is already an OOXML dash name; AddTextbox's
                    // MapDashStyle accepts it (lower-cased) and re-emits it.
                    var dash = ln.Element(a + "prstDash")?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(dash)) props["line.style"] = dash;
                }
            }
            // bodyPr text insets. AddTextbox hardcodes Word defaults
            // (91440/45720); a source with zero insets (common on tight
            // letterhead title boxes) otherwise loses ~0.2in of usable width
            // and its text rewraps. Forward each present inset verbatim (EMU).
            var bodyPr = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "bodyPr");
            if (bodyPr != null)
            {
                foreach (var (attr, key) in new[] { ("lIns", "inset.left"), ("tIns", "inset.top"), ("rIns", "inset.right"), ("bIns", "inset.bottom") })
                {
                    var v = bodyPr.Attribute(attr)?.Value;
                    if (!string.IsNullOrEmpty(v)) props[key] = v;
                }
                // Vertical text flow + vertical anchor. Without these a vertical
                // (eaVert) box renders as char-wrapped horizontal text and a
                // centered box anchors to the top.
                var vert = bodyPr.Attribute("vert")?.Value;
                if (!string.IsNullOrEmpty(vert) && vert != "horz") props["textDirection"] = vert;
                var bAnchor = bodyPr.Attribute("anchor")?.Value;
                if (!string.IsNullOrEmpty(bAnchor) && bAnchor != "t") props["textAnchor"] = bAnchor;
                // BUG-DUMP-R25-6: wps:bodyPr/@wrap is the IN-shape text-wrap
                // mode (distinct from wp:wrapNone/wp:wrapSquare, the around-shape
                // wrap forwarded via `wrap`). AddTextbox hardcoded wrap="square";
                // a source wrap="none"/"tight"/"through" was clobbered to square.
                // Forward the value verbatim. The source bodyPr ALWAYS exists
                // here (we're inside `if (bodyPr != null)`), so emit an explicit
                // empty-string sentinel when @wrap is ABSENT — AddTextbox then
                // omits the attribute entirely, preserving absence instead of
                // re-injecting the schema default. (A `null` key, by contrast,
                // means "interactive caller never touched wrap" → legacy square.)
                var bodyWrap = bodyPr.Attribute("wrap")?.Value;
                props["bodyWrap"] = bodyWrap ?? "";
                // BUG-DUMP-R25-6: <a:spAutoFit/> (resize shape to fit text) is a
                // child of bodyPr. AddTextbox never emitted it, so the box kept
                // its fixed cy instead of shrinking to the content — same
                // vertical-extent / reflow defect. Forward as an autoFit flag.
                if (bodyPr.Elements().Any(e => e.Name.LocalName == "spAutoFit"))
                    props["autoFit"] = "true";
            }
            // docPr name → alt
            var docPr = doc.Descendants(wp + "docPr").FirstOrDefault();
            var altName = docPr?.Attribute("name")?.Value;
            if (!string.IsNullOrEmpty(altName) && altName != "Text Box") props["alt"] = altName;
        }
        catch
        {
            // Parsing failures: still emit the `add textbox` row with whatever
            // we managed to extract; defaults cover the rest.
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = emitParent,
            Type = "textbox",
            Props = props.Count > 0 ? props : null
        });

        // Recurse over inner content. Get on /<host>/textbox[N] returns the
        // <w:txbxContent>; its children are the inner <w:p>. AddTextbox auto-
        // seeds one empty <w:p>, so the first source paragraph uses set-on-
        // existing (autoPresent: true) and the rest emit as fresh adds.
        try
        {
            var txbxNode = word.Get(sourceReadPath);
            var children = txbxNode.Children ?? new List<DocumentNode>();
            int innerPIdx = 0;
            int innerTblIdx = 0;
            bool firstParaSeen = false;
            foreach (var child in children)
            {
                if (child.Type == "paragraph" || child.Type == "p")
                {
                    innerPIdx++;
                    // The generic fallback fabricates child paths from the
                    // OOXML LocalName ("/body/txbxContent[N]/p[M]") which the
                    // Navigation layer can't re-resolve — the user-facing
                    // path segment is "textbox", not "txbxContent". Use the
                    // canonical /body/textbox[N]/p[M] form instead.
                    var sourceParaPath = $"{sourceReadPath}/p[{innerPIdx}]";
                    EmitParagraph(word, sourceParaPath, textboxPath, innerPIdx, items,
                                  autoPresent: !firstParaSeen, ctx);
                    firstParaSeen = true;
                }
                else if (child.Type == "table" || child.Type == "tbl")
                {
                    // BUG-D1-TXBX-TABLE: tables nested INSIDE a textbox were
                    // silently dropped on dump — the children loop only
                    // recognised paragraph types. Reuse EmitTable with the
                    // textbox path as containerPath so the resulting
                    // `add table` rows target /body/textbox[N]/tbl[K]
                    // (AddTable already accepts a TextBoxContent parent).
                    innerTblIdx++;
                    var sourceTblPath = $"{sourceReadPath}/tbl[{innerTblIdx}]";
                    EmitTable(word, sourceTblPath, innerTblIdx, items, ctx,
                              parentTablePath: null, containerPath: textboxPath);
                }
                else if (child.Type == "sdt")
                {
                    // BUG-R5B(BUG1): a block-level <w:sdt> nested inside a
                    // textbox (the canonical centered page-number footer:
                    // textbox → sdt → <w:p> with a PAGE field) was silently
                    // dropped — the inner walk only recognised paragraph and
                    // table types, so the SDT (and the field inside it) vanished
                    // on dump→batch and the footer lost its page number. There
                    // is no typed `add sdt` parent for a txbxContent host, and a
                    // PAGE-field SDT is "rich" content the typed text emit cannot
                    // reproduce anyway. Inject the SDT verbatim via raw-set into
                    // the just-created textbox's <w:txbxContent> (mirrors the
                    // rich-block-SDT raw-set in EmitSdt). The part is the host's
                    // own part (/document for body, /header[N] or /footer[N]
                    // otherwise); the (//w:txbxContent)[n] index selects the
                    // textbox we just emitted.
                    var sdtRaw = word.RawElementXml(child.Path);
                    if (!string.IsNullOrEmpty(sdtRaw))
                    {
                        var rawPart = hostPath == "/body" ? "/document" : hostPath;
                        // `add textbox` auto-seeds one empty <w:p> in the
                        // txbxContent. To keep document order, an SDT that
                        // precedes every source paragraph (the canonical
                        // page-number footer: sdt-then-empty-p) is prepended so
                        // it lands ahead of the seed paragraph; an SDT that
                        // follows already-emitted paragraphs is appended.
                        items.Add(new BatchItem
                        {
                            Command = "raw-set",
                            Part = rawPart,
                            Xpath = $"(//w:txbxContent)[{n}]",
                            Action = firstParaSeen ? "append" : "prepend",
                            Xml = sdtRaw
                        });
                    }
                    else
                    {
                        ctx.Warnings.Add(new DocxUnsupportedWarning(
                            Element: "textbox.sdt",
                            Path: child.Path,
                            Reason: "content control nested inside a textbox could not be serialized for round-trip; it will be missing from the replayed document"));
                    }
                }
            }
        }
        catch
        {
            // If the inner walk fails for any reason, the typed `add textbox`
            // still landed — round-trip recreates an empty textbox with the
            // right geometry, which beats the previous double-emit.
        }
        return true;
    }

    private static bool IsTextboxHostPath(string parentPath)
    {
        // Matches ResolveDrawingHost: /body, /body/tbl[..]/tr[..]/tc[N],
        // /header[N], /footer[N]. Reject anything else so non-supported
        // hosts fall through to the raw-set append.
        if (string.Equals(parentPath, "/body", StringComparison.Ordinal)) return true;
        if (parentPath.StartsWith("/header[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)
            && !parentPath.Substring(8).Contains('/')) return true;
        if (parentPath.StartsWith("/footer[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)
            && !parentPath.Substring(8).Contains('/')) return true;
        if (parentPath.Contains("/tc[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)) return true;
        return false;
    }

    // BUG-DUMP-R25-7: a bare /header[N] or /footer[N] host (NOT a cell, which
    // already attaches its textbox to the containing paragraph via the
    // "/tc[" branch in TryEmitPictureRun). Used to keep a single footer/header
    // textbox inside its styled host paragraph instead of letting AddTextbox
    // fork a new unstyled paragraph.
    private static bool IsHeaderFooterHost(string parentPath)
    {
        // Reject nested (cell) paths: a bare /footer[N] has its only '/' at 0.
        if (parentPath.IndexOf('/', 1) >= 0) return false;
        return (parentPath.StartsWith("/header[", StringComparison.Ordinal)
                || parentPath.StartsWith("/footer[", StringComparison.Ordinal))
            && parentPath.EndsWith("]", StringComparison.Ordinal);
    }

    // BUG-DUMP-R26-7: resolve the (Part, XPath) for a raw-set APPEND into the
    // last paragraph of the host identified by <paramref name="parentPath"/>.
    // The verbatim-content raw-set fallbacks (rich field result FIX-2, nested
    // SDT FIX-5, VML textbox FIX-6) previously fired only for /body; this
    // extends them to header / footer / table-cell hosts so the content
    // survives there too instead of falling silently to the lossy extractor.
    //   /body                         -> ("/document",  "/w:document/w:body/w:p[last()]")
    //   /header[N]                    -> ("/header[N]", "/w:hdr/w:p[last()]")
    //   /footer[N]                    -> ("/footer[N]", "/w:ftr/w:p[last()]")
    //   table cell (ctx box set)      -> ("/document",  "(//w:tbl)[N]/w:tr[M]/w:tc[K]/w:p[last()]")
    // Returns null when the host isn't one we can address (the caller then
    // keeps its existing behaviour — typically the lossy typed emit).
    private static (string Part, string XPath)? ResolveRawSetHost(string parentPath, BodyEmitContext? ctx)
    {
        if (parentPath == "/body")
            return ("/document", "/w:document/w:body/w:p[last()]");
        if (IsHeaderFooterHost(parentPath))
        {
            var root = parentPath.StartsWith("/header[", StringComparison.Ordinal) ? "/w:hdr" : "/w:ftr";
            return (parentPath, $"{root}/w:p[last()]");
        }
        // Table cell: EmitTable stashes the current cell's ordinal XPath in the
        // context box while walking the cell's paragraphs. Append into that
        // cell's last paragraph. CurrentCellPartBox carries the owning part:
        // "/document" for a body table, the header/footer part path for a
        // header/footer-hosted table (BUG-DUMP-R35-HFCELL — previously the part
        // was hardcoded "/document" and header/footer cells were not carried at
        // all, so a rich inline SDT there fell through to the lossy typed emit).
        if (parentPath.Contains("/tc[", StringComparison.Ordinal)
            && ctx?.CurrentCellXPathBox is { } box && box[0] is { } cellXPath)
        {
            var cellPart = ctx.CurrentCellPartBox is { } pbox && pbox[0] is { } p
                ? p : "/document";
            return (cellPart, $"{cellXPath}/w:p[last()]");
        }
        return null;
    }

    /// <summary>
    /// R10-bug1: OLE / embedded-object runs (Type=="ole"). Full round-trip:
    /// pull the embedded payload + VML icon binaries and the frame metadata
    /// (progId / DrawAspect / dimensions / alt-name) and emit a self-contained
    /// `add ole` carrying the bytes as data: URIs (mirrors picture-run base64
    /// inlining). AddOle rebuilds the embedded part, icon part and &lt;w:object&gt;
    /// wrapper from these — no external src file needed.
    ///
    /// If the payload can't be resolved (orphaned relationship / unreadable
    /// part), fall back to keeping the host paragraph and emitting a warning
    /// naming the path, rather than silently dropping the object.
    /// </summary>
    private static bool TryEmitOleRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx, WordHandler word)
    {
        if (run.Type != "ole") return false;

        var data = word.GetOleEmitData(run.Path);
        if (data != null)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["src"] = $"data:{data.EmbeddedContentType};base64,{Convert.ToBase64String(data.EmbeddedBytes)}",
                ["oleKind"] = data.OleKind,
                ["contentType"] = data.EmbeddedContentType,
            };
            if (!string.IsNullOrEmpty(data.EmbeddedExt)) props["embedExt"] = data.EmbeddedExt;
            if (!string.IsNullOrEmpty(data.ProgId)) props["progId"] = data.ProgId!;
            if (!string.IsNullOrEmpty(data.Display)) props["display"] = data.Display!;
            if (!string.IsNullOrEmpty(data.Width)) props["width"] = data.Width!;
            if (!string.IsNullOrEmpty(data.Height)) props["height"] = data.Height!;
            if (!string.IsNullOrEmpty(data.Name)) props["name"] = data.Name!;
            // Floating-OLE positioning: the verbatim v:shape style (position:
            // absolute + margin + z-index + wrap) keeps the object out of the text
            // flow on replay, plus the original native object box.
            if (!string.IsNullOrEmpty(data.ShapeStyle)) props["shapeStyle"] = data.ShapeStyle!;
            if (!string.IsNullOrEmpty(data.DxaOrig)) props["dxaOrig"] = data.DxaOrig!;
            if (!string.IsNullOrEmpty(data.DyaOrig)) props["dyaOrig"] = data.DyaOrig!;
            // BUG-DUMP-OLECROP: forward the VML imagedata crop so AddOle re-applies
            // it — an uncropped preview renders larger and pushes later pages down.
            if (!string.IsNullOrEmpty(data.Crop)) props["crop"] = data.Crop!;
            // BUG-DUMP-DELOLE: forward tracked-change attribution so AddOle re-wraps
            // the rebuilt OLE run in <w:del>/<w:ins>/move. A tracked-DELETED figure
            // (invisible in Word's final view) otherwise resurrects as a LIVE
            // full-size object, inflating the page count and cascading a render
            // drift. Mirrors the deleted-break (TryEmitBreakRun) / deleted-field
            // forwarding. revision.* live on the run node (set by the Get-side
            // DeletedRun/InsertedRun ancestor walk), not on GetOleEmitData.
            foreach (var rk in new[] { "revision.type", "revision.author", "revision.date", "revision.id" })
                if (run.Format.TryGetValue(rk, out var rv)
                    && rv?.ToString() is { Length: > 0 } rvs)
                    props[rk] = rvs;
            if (data.IconBytes is { Length: > 0 })
                props["icon"] = $"data:{data.IconContentType ?? "image/png"};base64,{Convert.ToBase64String(data.IconBytes)}";
            // BUG-DUMP-OLERPR: forward the OLE run's <w:rPr> so AddOle re-applies
            // it. The run wrapping <w:object> can carry run typography that affects
            // layout — most visibly a <w:bdr> border box around the object, but
            // also rFonts/sz that set the host line height. AddOle otherwise builds
            // a bare <w:r> and the lost border/line-height nudged every following
            // line, reflowing the page. Mirrors the break-run rPr forwarding above;
            // extract from raw XML since Navigation strips typography off ole nodes.
            var oleRawXml = word.RawElementXml(run.Path);
            if (!string.IsNullOrEmpty(oleRawXml))
            {
                try
                {
                    var oleRunEl = System.Xml.Linq.XElement.Parse(oleRawXml!);
                    var wNsOle = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    var oleRPrEl = oleRunEl.Element(wNsOle + "rPr");
                    if (oleRPrEl != null)
                        props["runRpr"] = oleRPrEl.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                }
                catch { /* no rPr to forward */ }
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "ole",
                Props = props,
            });
            return true;
        }

        // An ActiveX form control (<w:object> hosting <w:control r:id> with a
        // VML preview image, no o:OLEObject) has no embedded OLE payload, so
        // GetOleEmitData returns null — previously every such control (radio
        // buttons, checkboxes in form templates) was warn-dropped and its table
        // cell rebuilt empty. Emit a self-contained `add activex` instead:
        // verbatim run XML plus every referenced part's bytes base64-inlined,
        // mirroring the `add ole` data: URI design.
        var axData = word.GetActiveXEmitData(run.Path);
        if (axData != null)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "inlinedparts",
                Props = PackInlinedPartsProps(axData),
            });
            return true;
        }

        if (ctx != null)
        {
            var progId = run.Format.TryGetValue("progId", out var pid) ? pid?.ToString() : null;
            var reason = progId != null
                ? $"ole run dropped (progId={progId}); embedded payload could not be resolved for round-trip"
                : "ole run dropped; embedded payload could not be resolved for round-trip";
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "ole",
                Path: run.Path,
                Reason: reason));
        }
        return true;
    }

    // Shared prop packing for the inlined-parts carriers (`add activex`,
    // `add diagram`): verbatim run XML + one part{N}.relId/part{N}.data pair
    // per referenced package part, with part{N}.child{M}.* for nested parts.
    // BUG-DUMP-OLE-IN-OMATH: a MathType / Equation OLE object embedded inside the
    // verbatim <m:oMath> carrier references its binary (<o:OLEObject r:id>) and
    // preview image (<v:imagedata r:id>) by relationship id. The `xml` carrier
    // ships those refs but not the parts, so they dangle on replay (a silent
    // embedding loss plus a validator NullReferenceException). Base64-inline the
    // referenced parts as part{N}.* (the same carrier shape as activex/vmlshape)
    // so AddEquation rematerializes them and rewrites the r:ids. No-op when the
    // math carries no verbatim xml or references no parts (the common case), so
    // the interactive `add equation formula=` path is untouched.
    private static void AddMathInlinedPartProps(WordHandler word, string? mathPath, Dictionary<string, string> eqProps)
    {
        if (string.IsNullOrEmpty(mathPath)
            || !eqProps.TryGetValue("xml", out var xml)
            || string.IsNullOrEmpty(xml)
            || !xml.Contains(":id=\"", StringComparison.Ordinal))
            return;
        var inlined = word.GetMathInlinedPartsEmitData(mathPath);
        if (inlined == null) return;
        foreach (var kv in PackInlinedPartsProps(inlined))
            if (!string.Equals(kv.Key, "runXml", StringComparison.Ordinal))
                eqProps[kv.Key] = kv.Value;
    }

    private static Dictionary<string, string> PackInlinedPartsProps(WordHandler.ActiveXEmitData data)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["runXml"] = data.RunXml,
        };
        int pi = 0;
        foreach (var part in data.Parts)
        {
            pi++;
            props[$"part{pi}.relId"] = part.RelId;
            props[$"part{pi}.data"] =
                $"data:{part.ContentType};base64,{Convert.ToBase64String(part.Bytes)}";
            int ci = 0;
            foreach (var child in part.Children)
            {
                ci++;
                props[$"part{pi}.child{ci}.relId"] = child.RelId;
                props[$"part{pi}.child{ci}.data"] =
                    $"data:{child.ContentType};base64,{Convert.ToBase64String(child.Bytes)}";
                // BUG-DUMP-R71-USERSHAPES-IMG: a child can own further parts (a
                // chart userShapes drawing -> its image). Emit that grandchild
                // level so the drawing's r:embed isn't left dangling on replay.
                int gi = 0;
                foreach (var gc in child.Children)
                {
                    gi++;
                    props[$"part{pi}.child{ci}.gc{gi}.relId"] = gc.RelId;
                    props[$"part{pi}.child{ci}.gc{gi}.data"] =
                        $"data:{gc.ContentType};base64,{Convert.ToBase64String(gc.Bytes)}";
                }
            }
            // Per-part external rels (e.g. a chart's <c:externalData r:id> ->
            // external oleObject workbook). Recreated on the part itself, with
            // the original id, since the verbatim part bytes reference it.
            int pei = 0;
            foreach (var ext in part.Externals)
            {
                pei++;
                props[$"part{pi}.ext{pei}.relId"] = ext.RelId;
                props[$"part{pi}.ext{pei}.type"] = ext.Type;
                props[$"part{pi}.ext{pei}.target"] = ext.Target;
            }
        }
        int ei = 0;
        foreach (var ext in data.Externals)
        {
            ei++;
            props[$"ext{ei}.relId"] = ext.RelId;
            props[$"ext{ei}.type"] = ext.Type;
            props[$"ext{ei}.target"] = ext.Target;
        }
        return props;
    }

    private static bool TryEmitNoteRefRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Footnote/endnote reference runs carry a <w:footnoteReference> /
        // <w:endnoteReference> child. Emit them as a typed footnote/endnote
        // add anchored on the host paragraph and pull the body text from the
        // pre-resolved ordered list — see BodyEmitContext for the
        // document-order assumption.
        //
        // BUG-R7B(BUG1): detection cannot rely on rStyle == "FootnoteReference":
        // that is only the default style name. Real documents reference the
        // note with an arbitrary style id (e.g. rStyle="a5"), so the run's
        // rStyle never matched and the reference (plus its note body) was
        // silently dropped — Get surfaces no Format key for the reference
        // element, so probe the raw run XML for the actual reference child.
        if (ctx == null) return false;
        var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
        var noteKind = ClassifyNoteRefRun(word, run, rStyle);
        // BUG-R12A(BUG3): emit the note body STRUCTURALLY (per-run rPr +
        // multi-paragraph) instead of flattening it to one `text` prop. The
        // cursor is the 0-based document-order reference index; source AND target
        // note are the (cursor+1)-th note (Query and the body walk both run in
        // source order, one `add <kind>` per reference).
        if (noteKind == NoteRefKind.Footnote)
        {
            int idx = ++ctx.FootnoteCursor.Index; // 1-based source/target index
            EmitNoteReference(word, "footnote", idx, idx, paraTargetPath, items, run);
            return true;
        }
        if (noteKind == NoteRefKind.Endnote)
        {
            int idx = ++ctx.EndnoteCursor.Index; // 1-based source/target index
            EmitNoteReference(word, "endnote", idx, idx, paraTargetPath, items, run);
            return true;
        }
        return false;
    }

    private enum NoteRefKind { None, Footnote, Endnote }

    // BUG-R7B(BUG1): a run is a footnote/endnote reference when it contains a
    // <w:footnoteReference>/<w:endnoteReference> child — the rStyle is only a
    // weak hint (defaults to FootnoteReference/EndnoteReference but is an
    // arbitrary style id in real documents). Probe the raw XML; fall back to
    // the rStyle name when raw XML is unavailable.
    private static NoteRefKind ClassifyNoteRefRun(WordHandler word, DocumentNode run, string? rStyle)
    {
        if (run.Type != "run" && run.Type != "r") return NoteRefKind.None;
        var raw = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(raw))
        {
            // BUG-DUMP-NOTEREF-LITERAL: match the <w:footnoteReference> ELEMENT
            // open tag, not the bare token. A run whose visible text literally
            // contains the word "footnoteReference"/"endnoteReference" (common in
            // docs that describe OOXML/Word internals) is NOT a note anchor — the
            // old substring probe matched the word inside <w:t>…</w:t> and replaced
            // the whole run with a synthesized note reference, silently dropping all
            // its text. Require the element open-tag form `<[prefix:]footnoteReference`
            // followed by a tag-terminating char so text content can never match.
            if (System.Text.RegularExpressions.Regex.IsMatch(raw, @"<\w*:?footnoteReference[\s/>]"))
                return NoteRefKind.Footnote;
            if (System.Text.RegularExpressions.Regex.IsMatch(raw, @"<\w*:?endnoteReference[\s/>]"))
                return NoteRefKind.Endnote;
            return NoteRefKind.None;
        }
        if (string.Equals(rStyle, "FootnoteReference", StringComparison.OrdinalIgnoreCase)) return NoteRefKind.Footnote;
        if (string.Equals(rStyle, "EndnoteReference", StringComparison.OrdinalIgnoreCase)) return NoteRefKind.Endnote;
        return NoteRefKind.None;
    }

    // BUG-DUMP-R35-2: deterministic flatten warning for a run synthesized from
    // inside a <w:smartTag>/<w:customXml> wrapper (Navigation marks it with
    // Format["_wrapperFlattened"]). Shared by the per-run emit path
    // (EmitPlainOrHyperlinkRun) and the single-run paragraph collapse path so
    // the wrapper loss is never silent regardless of which shape the
    // paragraph takes. CONSISTENCY(wrapper-flatten-warning).
    private static void WarnWrapperFlattened(DocumentNode run, BodyEmitContext? ctx)
    {
        if (run.Format.TryGetValue("_wrapperFlattened", out var wfObj)
            && wfObj is bool wfB && wfB && ctx != null)
        {
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "smartTag/customXml",
                Path: run.Path,
                Reason: "inline smartTag/customXml wrapper flattened on dump→batch round-trip; the wrapped run text and formatting are preserved, only the wrapper element is dropped"));
        }
    }

    // BUG-DUMP-R28-REWRITTEN: mirror AddHyperlink's accept logic
    // (WordHandler.Add.Misc.cs): a url is emittable iff it is a fragment anchor
    // (#name), a safe-scheme absolute URI, or a relative target. The SDK leaves a
    // "rewritten://<guid>" placeholder when a Target is so malformed it cannot be
    // parsed into a System.Uri at all (the canonical case is a mailto: whose
    // address part is free text typed into the link field). Emitting `add
    // hyperlink url=…` for such a value aborts the batch step (AddHyperlink
    // throws "Invalid hyperlink URL"), dropping the run. Returns false so the
    // caller drops the url and degrades to a plain run, preserving the text.
    private static bool IsEmittableHyperlinkUrl(string? url)
    {
        if (string.IsNullOrEmpty(url)) return false;
        if (url.StartsWith("rewritten:", StringComparison.OrdinalIgnoreCase)) return false;
        if (url.StartsWith("#", StringComparison.Ordinal)) return true;
        if (Uri.TryCreate(url, UriKind.Absolute, out _))
            return Core.HyperlinkUriValidator.IsSafeScheme(url);
        return Uri.TryCreate(url, UriKind.Relative, out _);
    }

    private static void EmitPlainOrHyperlinkRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx = null, int hlBaseline = 0)
    {
        // BUG-R12A(BUG1): a hyperlink wrapper with >1 run or any per-run rPr was
        // stashed by CoalesceHyperlinkRuns with its original runs in Children.
        // Emit the wrapper carrying the FIRST run's text + rPr, then one
        // structured `add run` per remaining run targeting the hyperlink path so
        // each run's bold/color/size/font survives round-trip (the flat
        // `add hyperlink text=` path would flatten them into one unformatted run).
        // Mirrors the R9 comment-body structured-add-run fix; raw-set is not an
        // option here because the <w:hyperlink> r:id relationship can't be
        // recreated by verbatim XML injection.
        if (run.Format.TryGetValue("_hlStructured", out var hlsObj) && hlsObj is bool hlsB && hlsB
            && run.Children is { Count: > 0 } hlRuns)
        {
            EmitStructuredHyperlink(word, hlRuns, paraTargetPath, items, ctx, hlBaseline);
            return;
        }
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;
        // BUG-DUMP-R35-2: a run synthesized from inside a <w:smartTag>/
        // <w:customXml> wrapper. We FLATTEN the wrapper (drop the smartTag/
        // customXml element) but PRESERVE the inner run's text + formatting —
        // consistent with how Word often strips these and with the project's
        // flatten precedents. Surface a deterministic warning so the wrapper
        // loss isn't silent (matches the external-rel / picBullet convention).
        WarnWrapperFlattened(run, ctx);
        // CONSISTENCY(move-range-markers): a moveFrom/moveTo run's own w:id in
        // the source usually differs from its paired half (the pairing lives on
        // the bracketing range markers' shared w:name, not on the run id). Rewrite
        // both halves to one SHARED revision.id so AddRun's WrapRunAsMove* helpers
        // synthesize Move_{id} range markers that pair the moveFrom to its moveTo.
        if (ctx is { MovePairIds.Count: > 0 }
            && rProps.TryGetValue("revision.type", out var mvType)
            && (mvType.Equals("moveFrom", StringComparison.OrdinalIgnoreCase)
                || mvType.Equals("moveTo", StringComparison.OrdinalIgnoreCase))
            && rProps.TryGetValue("revision.id", out var mvId)
            && ctx.MovePairIds.TryGetValue(mvId, out var sharedId))
        {
            rProps["revision.id"] = sharedId;
        }
        // BUG-DUMP-R43-8: rPrChange's PreviousRunProperties snapshot now
        // round-trips verbatim via revision.beforeXml (see EmitParagraph
        // counterpart). The former warn-and-drop path is retired — the
        // prior-rPr payload is preserved. Defensive strip only, for stale dumps.
        rProps.Remove("revision.beforeLost");

        // Hyperlink-wrapped run: Get flattens a <w:hyperlink>'s child run
        // into a regular run-typed node but copies the resolved URL onto
        // Format["url"]. AddRun does not consume `url` — emitting type="r"
        // would silently drop the hyperlink wrapper. Re-emit as a typed
        // `add hyperlink` so the <w:hyperlink>+rel-relationship round-trip
        // rebuilds correctly.
        // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
        // on both Get readback and Add input.
        if (rProps.ContainsKey("url") || rProps.ContainsKey("anchor")
            || rProps.ContainsKey("isHyperlink"))
        {
            // AddHyperlink writes its own color/underline defaults from theme;
            // drop the inferred `color: hyperlink` / `underline: single` Get
            // echoes back so we don't override those defaults. Track the drops:
            // a dropped key means the SOURCE had an explicit element that the
            // defaults reproduce — the "inherit" sentinel below must NOT fire
            // for it, or the explicit single underline / theme color vanishes.
            bool hlColorDropped = false, hlUnderlineDropped = false;
            if (rProps.TryGetValue("color", out var hlColor)
                && string.Equals(hlColor, "hyperlink", StringComparison.OrdinalIgnoreCase))
            {
                rProps.Remove("color");
                hlColorDropped = true;
            }
            if (rProps.TryGetValue("underline", out var hlUl)
                && string.Equals(hlUl, "single", StringComparison.OrdinalIgnoreCase))
            {
                rProps.Remove("underline");
                hlUnderlineDropped = true;
            }
            rProps.Remove("isHyperlink");
            // BUG-DUMP-R28-REWRITTEN: a malformed Target (SDK rewritten://
            // placeholder, or a mailto whose address is free text) cannot be
            // expressed as `add hyperlink` — the url fails AddHyperlink's
            // validation and aborts the batch step, dropping the run. Drop the
            // unusable url (+ warn) so the wrapper degrades to a plain run and the
            // visible text survives. A tooltip/tgtFrame/history-only wrapper still
            // emits (those don't need a parseable url).
            if (rProps.TryGetValue("url", out var hlUrlCheck) && !IsEmittableHyperlinkUrl(hlUrlCheck))
            {
                rProps.Remove("url");
                ctx?.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "hyperlink.url",
                    Path: run.Path,
                    Reason: $"hyperlink target '{hlUrlCheck}' is malformed (not a valid absolute/relative URI or anchor) and cannot round-trip; the link is dropped and its text is preserved as plain text"));
            }
            // Bare <w:hyperlink> wrapper with no url/anchor/tooltip/tgtFrame
            // /history carries no round-trippable property — AddHyperlink
            // would reject it. Fall through and emit as a plain run.
            if (!rProps.ContainsKey("url") && !rProps.ContainsKey("anchor")
                && !rProps.ContainsKey("tooltip") && !rProps.ContainsKey("tgtFrame")
                && !rProps.ContainsKey("tgtframe") && !rProps.ContainsKey("history"))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = rProps.Count > 0 ? rProps : null
                });
                return;
            }
            // The SOURCE run has no <w:color>/<w:u> at all (a TOC leader row
            // that deliberately looks like plain text). Without the sentinel,
            // AddHyperlink stamps its theme-blue + single-underline defaults
            // and the dotted leader comes back blue. Injected only on the real
            // `add hyperlink` path — the bare-wrapper fallback above emits a
            // plain `add r`, whose color parser must not see "inherit".
            if (!hlColorDropped && !rProps.ContainsKey("color")) rProps["color"] = "inherit";
            if (!hlUnderlineDropped && !rProps.ContainsKey("underline")) rProps["underline"] = "inherit";
            // BUG-DUMP-HYPERLINK-EMPTYTEXT: the wrapper's display text comes from
            // the source run's <w:t>. When that run is EMPTY, line 3600 above
            // omits the `text` key (empty text isn't emitted) — and AddHyperlink's
            // `GetValueOrDefault("text", url ?? anchor ?? "link")` then injects the
            // URL/anchor as VISIBLE text (e.g. a 3-run hyperlink whose first run is
            // <w:t></w:t> rebuilt as "ex191….htmInsider Trading…"). Emit an
            // explicit empty `text` so the wrapper's first run stays empty and the
            // real display text comes from the trailing structured `add r` runs.
            if (!rProps.ContainsKey("text")) rProps["text"] = "";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "hyperlink",
                Props = rProps,
            });
            return;
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "r",
            Props = rProps.Count > 0 ? rProps : null
        });
    }

    // BUG-R12A(BUG1): emit a hyperlink wrapper + one structured `add run` per
    // wrapped run so per-run rPr survives. The first run materializes the
    // wrapper (AddHyperlink seeds the wrapper's first run from `text` + rPr
    // props); subsequent runs are appended via `add r` targeting the
    // hyperlink[K] path (verified working — AddRun accepts a hyperlink parent
    // and preserves bold/color/size/font). hlIndex is the 1-based ordinal of
    // this hyperlink among the rows already emitted under the host paragraph.
    // BUG-R14B: hlBaseline is the number of hyperlink rows already present
    // at paraTargetPath *before this paragraph's own run-emit pass began*.
    // paraTargetPath is the literal "/body/p[last()]" for every paragraph, so
    // a raw items.Count() of hyperlink rows at that parent also tallies the
    // hyperlinks from previously-emitted paragraphs — at replay time those
    // live under earlier <w:p> elements, not the current last() paragraph, so
    // the current paragraph's hyperlinks re-index from 1. Subtracting the
    // baseline yields the wrapper's LIVE 1-based index inside this paragraph,
    // which is what the trailing `add r` rows must target.
    private static void EmitStructuredHyperlink(WordHandler word, List<DocumentNode> hlRuns, string paraTargetPath,
                                                List<BatchItem> items, BodyEmitContext? ctx, int hlBaseline = 0)
    {
        // Build the wrapper add from the first run's props (url/anchor/tooltip/…
        // + the run's own rPr). Reuse the existing flat-emit logic by routing
        // the first run through EmitPlainOrHyperlinkRun with the structured flag
        // cleared, so the theme-default scrubbing + bare-wrapper fallback stay
        // identical.
        var first = hlRuns[0];
        var firstClone = new DocumentNode
        {
            Path = first.Path,
            Type = "run",
            // A tab-first hyperlink (TOC leader): the wrapper's text "\t" is
            // split into a real <w:tab/> by AddText, so the leader stays the
            // hyperlink's first child in source order.
            Text = first.Type == "tab" ? "\t" : first.Text,
            Format = new Dictionary<string, object?>(first.Format, StringComparer.OrdinalIgnoreCase),
        };
        firstClone.Format.Remove("_hlStructured");
        int hlBefore = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        EmitPlainOrHyperlinkRun(word, firstClone, paraTargetPath, items, ctx);
        int hlAfter = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        // If the first run did not materialize a hyperlink row (bare wrapper with
        // no url/anchor/tooltip → AddHyperlink can't represent it, so it fell back
        // to a plain run), the wrapper is gone — emit the remaining runs as plain
        // runs too so their text/rPr still survive.
        if (hlAfter == hlBefore)
        {
            for (int k = 1; k < hlRuns.Count; k++)
            {
                var clone = new DocumentNode
                {
                    Path = hlRuns[k].Path,
                    Type = hlRuns[k].Type,
                    Text = hlRuns[k].Text,
                    Format = new Dictionary<string, object?>(hlRuns[k].Format, StringComparer.OrdinalIgnoreCase),
                };
                clone.Format.Remove("_hlStructured");
                clone.Format.Remove("url");
                clone.Format.Remove("anchor");
                EmitPlainOrHyperlinkRun(word, clone, paraTargetPath, items, ctx);
            }
            return;
        }
        // BUG-R14B: live in-paragraph index = total hyperlink rows at this
        // parent minus the count that belonged to earlier paragraphs.
        var hlPath = $"{paraTargetPath}/hyperlink[{hlAfter - hlBaseline}]";
        for (int k = 1; k < hlRuns.Count; k++)
        {
            // BUG-DUMP-FOOTNOTE-IN-HYPERLINK: a footnote/endnote REFERENCE run
            // nested INSIDE a hyperlink (e.g. a linked phrase that also carries a
            // footnote) reaches here as a wrapped child. Emitting it as a bare
            // `add r` drops the <w:footnoteReference> element AND fails to advance
            // the per-reference note cursor — so a LATER note body (the highest id)
            // silently disappears (the visible symptom is "the last footnote went
            // missing"). Classify it and route through the note-reference emit
            // (targeting the hyperlink path so the mark stays inside the link),
            // advancing the cursor exactly like a top-level note ref.
            if (ctx != null)
            {
                var khStyle = hlRuns[k].Format.TryGetValue("rStyle", out var khrs) ? khrs?.ToString() : null;
                var khNote = ClassifyNoteRefRun(word, hlRuns[k], khStyle);
                // AddFootnote/AddEndnote require a PARAGRAPH parent (they reject a
                // hyperlink path), so anchor the note ref on paraTargetPath rather
                // than hlPath. The reference lands in the host paragraph adjacent to
                // the link instead of strictly inside it — the same "good enough"
                // boundary trade-off the comment-in-SDT/oMath strips accept — but the
                // note body + continuous numbering are preserved (cursor advances in
                // document order), which is what was silently lost before.
                if (khNote == NoteRefKind.Footnote)
                {
                    int fidx = ++ctx.FootnoteCursor.Index;
                    EmitNoteReference(word, "footnote", fidx, fidx, paraTargetPath, items, hlRuns[k]);
                    continue;
                }
                if (khNote == NoteRefKind.Endnote)
                {
                    int eidx = ++ctx.EndnoteCursor.Index;
                    EmitNoteReference(word, "endnote", eidx, eidx, paraTargetPath, items, hlRuns[k]);
                    continue;
                }
            }
            var rProps = FilterEmittableProps(hlRuns[k].Format);
            // The hyperlink-wrapper keys belong to the <w:hyperlink>, not its
            // child runs — strip them so `add r` doesn't choke / re-wrap.
            // BUG-DUMP-R43-4: rStyle/rstyle is a PER-RUN character style (e.g.
            // Internetlink), NOT a hyperlink-wrapper attribute. It lives in
            // HyperlinkWrapperOnlyKeys only so a sole rStyle-bearing run stays on
            // the lossless fast path; here, on the structured (multi-run) path,
            // each trailing run must re-emit its OWN rStyle (the first run already
            // got it via the wrapper add). Keep it — strip only the genuine
            // wrapper attrs.
            foreach (var wk in HyperlinkWrapperOnlyKeys)
            {
                if (string.Equals(wk, "rStyle", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(wk, "rstyle", StringComparison.OrdinalIgnoreCase))
                    continue;
                rProps.Remove(wk);
            }
            // Drop the theme-default echoes Get stamps on every hyperlink run;
            // the wrapper's own run already carries the real Hyperlink style.
            if (rProps.TryGetValue("color", out var c)
                && string.Equals(c, "hyperlink", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("color");
            if (rProps.TryGetValue("underline", out var u)
                && string.Equals(u, "single", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("underline");
            if (hlRuns[k].Type == "tab")
                rProps["text"] = "\t";
            else if (!string.IsNullOrEmpty(hlRuns[k].Text))
                rProps["text"] = hlRuns[k].Text!;
            // BUG-DUMP-HYPHEN-RESERIALIZE: a structural hyphen run
            // (<w:noBreakHyphen/>/<w:softHyphen/>) INSIDE a hyperlink reaches the
            // structured-hyperlink trailing-run emit, not TryEmitHyphenRun. Emit
            // it with the hyphen= prop so AddRun rebuilds the element instead of
            // persisting the U+2011/U+00AD glyph as literal <w:t> text (the glyph
            // wraps differently and degrades the round-trip). Mirrors the
            // host-agnostic hyphen emit; AddRun accepts a hyperlink parent.
            // Read _hasHyphen from the ORIGINAL run Format — FilterEmittableProps
            // (rProps) strips it via SkipKeys, so it's gone from rProps here.
            if (hlRuns[k].Format.TryGetValue("_hasHyphen", out var khKindObj))
            {
                rProps["hyphen"] = string.Equals(khKindObj?.ToString(), "soft", StringComparison.OrdinalIgnoreCase)
                    ? "soft" : "noBreak";
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = hlPath,
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    // BUG-R12A(BUG1): raw-set a rich inline (run-level) <w:sdt> into the host
    // paragraph so its per-run formatting survives. Returns true when it emitted
    // a raw-set (or a warning + fall-through is desired); false when the SDT is
    // simple enough for the flat `add sdt text=` fast path. The host paragraph
    // was just added by EmitParagraph, so it is the last <w:p> in the body at
    // replay time — the same last()-relative attach the inline-textbox raw-set
    // uses.
    private static bool TryEmitRichInlineSdt(WordHandler word, DocumentNode sdt,
                                             string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        var rawXml = word.RawElementXml(sdt.Path);
        if (string.IsNullOrEmpty(rawXml) || !IsRichInlineSdt(rawXml!)) return false;
        // BUG-DUMP-R26-7: target /body, header/footer OR table-cell hosts, not
        // only /body. When the host isn't raw-set-addressable, fall back to the
        // typed emit (no regression vs the old /body-only guard).
        if (ResolveRawSetHost(parentPath, ctx) is not { } sdtHost) return false;
        // External relationship references (hyperlink r:id, image r:embed/r:link)
        // would dangle in the rebuilt part — raw injection does not recreate the
        // matching rels. Emit a deterministic warning naming the loss instead of
        // silently flattening (BUG-DUMP-R26-7 PART B), then fall back to the flat
        // text emit. Full rel preservation is a separate, larger effort.
        if (HasExternalRelRef(rawXml!))
        {
            // Ship the SDT through the inlined-parts carrier (verbatim sdtXml +
            // part{N}/ext{N} data, rel ids rewritten on replay) — same shape as
            // the block-level EmitSdt and cell-level EmitCellSdt carriers, so
            // an inline picture control keeps its image. The replay `add sdt`
            // targets the host paragraph just added by EmitParagraph
            // (p[last()]), where AddSdt injects a run-level SdtRun. Body host
            // only: header/footer/cell hosts route through their own emitters.
            if (parentPath == "/body" && word.GetSdtEmitData(sdt.Path) is { } inlineSdtData)
            {
                var carrierProps = PackInlinedPartsProps(inlineSdtData);
                carrierProps["sdtXml"] = WordHandler.StripVerbatimCommentMarkers(carrierProps["runXml"]);
                carrierProps.Remove("runXml");
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = "/body/p[last()]",
                    Type = "sdt",
                    Props = carrierProps,
                });
                return true;
            }
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "sdt.richContent",
                Path: sdt.Path,
                Reason: "inline content control with formatted/nested content AND an external relationship (hyperlink/image) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the control is flattened to text on replay"));
            return false;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = sdtHost.Part,
            Xpath = sdtHost.XPath,
            Action = "append",
            Xml = rawXml
        });
        return true;
    }

    // A run-level <w:sdt> is "rich" when its sdtContent carries formatting the
    // text-only typed emit can't reproduce: more than one run, any run-level
    // <w:rPr>, a hyperlink/field, or an intra-run break/tab. Single plain run →
    // flat `add sdt text=` stays lossless. (Distinct from IsRichBlockSdt, which
    // keys on inner <w:p> count — a run-level SDT has no inner paragraph.)
    private static bool IsRichInlineSdt(string sdtXml)
    {
        // BUG-DUMP-R27-4: a run-level SDT carrying a repeatingSection /
        // repeatingSectionItem / docPartObj sdtPr marker must raw-set verbatim
        // (the typed path can't express the special type). See
        // HasSpecialSdtTypeMarker in WordBatchEmitter.Resources.cs.
        if (HasSpecialSdtTypeMarker(sdtXml))
            return true;
        // BUG-DUMP-R26-5: nested inline SDT. The outer <w:sdt> wraps one or more
        // child <w:sdt> in its sdtContent (L1>L2>L3 tag/id/alias nesting). The
        // flat `add sdt text=` path seeds a single run from the innermost text
        // and drops every nesting level's tag/id/alias plus the structure. A
        // second <w:sdt> anywhere in the XML means at least one nested control,
        // so raw-set the whole tree verbatim to preserve depth + per-level props.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:sdt[ >]").Count > 1)
            return true;
        // More than one content run.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:r[ >]").Count > 1)
            return true;
        // Any run-level run properties (bold/color/size/font/…) inside content.
        if (sdtXml.Contains("<w:rPr", StringComparison.Ordinal)
            // exclude sdtPr/endPr-only rPr false positives: those live in
            // <w:sdtPr>/<w:sdtEndPr><w:rPr>; a content rPr sits under <w:r>.
            && System.Text.RegularExpressions.Regex.IsMatch(sdtXml, "<w:r[ >].*?<w:rPr"))
            return true;
        // BUG-DUMP-H80-1: a run-level SDT whose content carries a tracked change
        // (<w:del>/<w:ins>/<w:moveFrom>/<w:moveTo>) cannot round-trip through the
        // flat `add sdt text=` path — a del-only content paragraph (no live run)
        // flattens to an empty run and the deletion is silently dropped. This is
        // the inline-path counterpart of the block-SDT fix (IsRichBlockSdt in
        // WordBatchEmitter.Resources.cs); the block fix did not cover this path.
        if (sdtXml.Contains("<w:del", StringComparison.Ordinal)
            || sdtXml.Contains("<w:ins", StringComparison.Ordinal)
            || sdtXml.Contains("<w:moveFrom", StringComparison.Ordinal)
            || sdtXml.Contains("<w:moveTo", StringComparison.Ordinal))
            return true;
        // BUG-DUMP-H94: a run-level SDT whose content carries a range/anchor marker
        // (<w:bookmarkStart/End>, <w:commentRangeStart/End> / <w:commentReference>,
        // <w:permStart/End>) loses those markers through the flat `add sdt text=`
        // path (seeds only text). Inline-path counterpart of the IsRichBlockSdt
        // fix; force the verbatim raw-set path.
        if (sdtXml.Contains("<w:bookmark", StringComparison.Ordinal)
            || sdtXml.Contains("<w:comment", StringComparison.Ordinal)
            || sdtXml.Contains("<w:perm", StringComparison.Ordinal))
            return true;
        return sdtXml.Contains("<w:hyperlink", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldChar", StringComparison.Ordinal)
            || sdtXml.Contains("w:instrText", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldSimple", StringComparison.Ordinal)
            || sdtXml.Contains("<w:drawing", StringComparison.Ordinal)
            || sdtXml.Contains("<w:br", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:cr", StringComparison.Ordinal)
            // BUG-DUMP-H95: text-less run-content the typed path drops (produce no
            // <w:t>) — symbol, positional tab (<w:ptab> ≠ <w:tab>), hyphen markers.
            || sdtXml.Contains("<w:sym", StringComparison.Ordinal)
            || sdtXml.Contains("<w:ptab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:noBreakHyphen", StringComparison.Ordinal)
            || sdtXml.Contains("<w:softHyphen", StringComparison.Ordinal)
            // BUG-DUMP-EQUATION-SDT: an EQUATION content control (<w:sdtPr><w:equation/>
            // … <w:sdtContent><m:oMathPara>/<m:oMath>) carries its math in m: runs
            // (<m:r>/<m:t>), not <w:r>/<w:t>, so none of the run checks above fire and
            // the typed `add sdt` path silently dropped the entire equation (wrapper +
            // math). Treat any SDT carrying math content — or the <w:equation/> sdtPr
            // type marker — as rich so it raw-sets verbatim, preserving the control
            // type and the equation.
            || sdtXml.Contains("<m:oMath", StringComparison.Ordinal)
            || sdtXml.Contains("<w:equation", StringComparison.Ordinal);
    }

    // Collapse OOXML complex field chains (fldChar(begin) + instrText + …
    // + fldChar(end)) into a single synthetic "field" DocumentNode with
    // Format["instruction"] (raw code) and Text (cached display value).
    // Non-field children pass through untouched in original order. The TOC
    // chain is handled by the dedicated EmitParagraph branch above and never
    // reaches this collapsing step (early-return in that branch).
    // BUG-DUMP6-05: collapse consecutive runs sharing the same url/anchor
    // into a single synthetic node so dump emits ONE `add hyperlink` per
    // <w:hyperlink>, regardless of how many runs the source wrapped. The
    // synthesized node carries the merged Text (for AddHyperlink's `text`
    // prop) and the shared url/anchor/Hyperlink-style format keys.
    // Mirrors the field-emit hyperlink-parent rebase logic for tab/ptab runs.
    // Navigation marks tab-only runs that live inside w:hyperlink with a
    // Format["_hyperlinkParent"] hint (e.g. /body/p[1]/hyperlink[2]); without
    // re-routing on emit they would replay under the bare paragraph and lose
    // the hyperlink wrapper. The candidate-verify step (a prior `add hyperlink`
    // row must have landed under paraTargetPath) avoids dangling paths when
    // the hyperlink has no emittable runs and so was never added.
    private static string ResolveHyperlinkParent(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        string? candidateHlParent = null;
        if (run.Format.TryGetValue("_hyperlinkParent", out var hlpObj) && hlpObj != null)
        {
            var hint = hlpObj.ToString();
            if (!string.IsNullOrEmpty(hint)) candidateHlParent = hint;
        }
        if (candidateHlParent == null) return paraTargetPath;

        const string hlMarker = "/hyperlink[";
        var hlIdxStart = candidateHlParent.LastIndexOf(hlMarker, StringComparison.Ordinal);
        if (hlIdxStart <= 0) return paraTargetPath;
        var hlEnd = candidateHlParent.IndexOf(']', hlIdxStart);
        if (hlEnd <= hlIdxStart) return paraTargetPath;
        var kStr = candidateHlParent.Substring(hlIdxStart + hlMarker.Length,
            hlEnd - hlIdxStart - hlMarker.Length);
        if (!int.TryParse(kStr, out var kIdx)) return paraTargetPath;
        var rebased = paraTargetPath + candidateHlParent.Substring(hlIdxStart);
        // BUG-DUMP-FIELDHL-XPARA: paraTargetPath ("/…/p[last()]") is identical for
        // every paragraph, so counting hyperlink rows by Parent==paraTargetPath
        // tallied hyperlinks from ALL prior paragraphs — a tab/ptab/equation run in
        // a hyperlink that emitted no `add hyperlink` row of its own would inherit a
        // phantom hyperlink from an earlier paragraph and route to a non-existent
        // /hyperlink[K]. Count only rows since this paragraph's own `add p`.
        int lastParaAdd = items.FindLastIndex(it => it.Command == "add" && it.Type == "p");
        int emittedHls = 0;
        for (int hi = lastParaAdd + 1; hi < items.Count; hi++)
            if (items[hi].Type == "hyperlink"
                && string.Equals(items[hi].Parent, paraTargetPath, StringComparison.Ordinal))
                emittedHls++;
        return emittedHls >= kIdx ? rebased : paraTargetPath;
    }

    // BUG-R12A(BUG1): hyperlink-wrapper keys + the theme defaults AddHyperlink
    // re-applies on its own. A run carrying ONLY these keys has no per-run
    // formatting worth preserving structurally, so the flat `add hyperlink
    // text=` fast path stays lossless. Any OTHER key (bold, color≠hyperlink,
    // size, font, …) means the run's rPr must survive — route to structured emit.
    private static readonly HashSet<string> HyperlinkWrapperOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "url", "anchor", "tooltip", "tgtFrame", "tgtframe", "history",
        "docLocation", "doclocation", "isHyperlink", "_hyperlinkParent",
        "rStyle", "rstyle",
    };

    // True when a hyperlink-wrapped run carries run-level formatting (rPr) that
    // AddHyperlink's flat `text=` path cannot reproduce. The theme-default
    // color=hyperlink / underline=single echoes that Get stamps back are NOT
    // real formatting (AddHyperlink re-applies them), so they don't count.
    private static bool HyperlinkRunHasRunFormatting(DocumentNode run)
    {
        var props = FilterEmittableProps(run.Format);
        foreach (var (k, v) in props)
        {
            if (HyperlinkWrapperOnlyKeys.Contains(k)) continue;
            if (string.Equals(k, "text", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(k, "color", StringComparison.OrdinalIgnoreCase)
                && string.Equals(v, "hyperlink", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(k, "underline", StringComparison.OrdinalIgnoreCase)
                && string.Equals(v, "single", StringComparison.OrdinalIgnoreCase)) continue;
            return true;
        }
        return false;
    }

    private static List<DocumentNode> CoalesceHyperlinkRuns(List<DocumentNode> runs)
    {
        var result = new List<DocumentNode>(runs.Count);
        int i = 0;
        while (i < runs.Count)
        {
            var run = runs[i];
            string? url = null, anchor = null, hlParent = null;
            // Tab runs INSIDE a hyperlink (a TOC row whose link wraps the
            // leader tab + page number) carry the same anchor/_hyperlinkParent
            // markers — include them so the wrapper add is emitted before (and
            // contains) the tab. Excluding them split the link: the tab run
            // raced ahead of the wrapper (its parent guard miscounted across
            // paragraphs sharing the literal /body/p[last()] parent) and the
            // replay dropped it.
            if (run.Type == "run" || run.Type == "r" || run.Type == "tab")
            {
                if (run.Format.TryGetValue("url", out var u))
                    url = u?.ToString();
                if (run.Format.TryGetValue("anchor", out var a))
                    anchor = a?.ToString();
                // `_hyperlinkParent` (e.g. `/body/p[1]/hyperlink[2]`) is the
                // unique-per-wrapper marker NodeBuilder stamps on every run
                // hosted inside a `<w:hyperlink>`. Two sibling hyperlinks
                // sharing the same target URL (e.g. two "Read more →" links
                // pointing at /pricing) carry identical `url`/`anchor` but
                // distinct `_hyperlinkParent` paths, so the URL-only equality
                // pre-R12 collapsed them into one item whose text was
                // "Read more →Read more →" — silent loss of the second link.
                if (run.Format.TryGetValue("_hyperlinkParent", out var hp))
                    hlParent = hp?.ToString();
            }
            if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(anchor))
            {
                result.Add(run);
                i++;
                continue;
            }
            // Walk forward over consecutive runs with the same url/anchor AND
            // the same hyperlink-wrapper path. Differing `_hyperlinkParent`
            // values mark a hyperlink boundary even when URL/anchor match.
            int j = i + 1;
            var group = new List<DocumentNode> { run };
            // BUG-DUMP-H99: zero-width position anchors (bookmark/perm markers) that
            // sit BETWEEN two runs of the SAME hyperlink are stashed here and
            // re-emitted right after the merged wrapper, so an interior anchor does
            // not fragment the one <w:hyperlink> into two.
            var pendingInterior = new List<DocumentNode>();
            var sb = new System.Text.StringBuilder(run.Text ?? "");
            while (j < runs.Count)
            {
                var next = runs[j];
                if (next.Type != "run" && next.Type != "r" && next.Type != "tab")
                {
                    // BUG-DUMP-H99: a <w:bookmarkStart/End> or <w:permStart/End>
                    // interior to a hyperlink (e.g. an _Hlk edit anchor Word inserts
                    // mid-link) previously broke the coalesce walk here, splitting one
                    // hyperlink into two on round-trip (silent injection 1→2). When a
                    // later run with the SAME url/anchor/_hyperlinkParent follows, the
                    // marker is interior: stash it (re-emitted just after the merged
                    // wrapper, where the zero-width anchor is position-equivalent) and
                    // keep coalescing. Any other non-run node — a second hyperlink,
                    // field, sdt, drawing — is a genuine boundary and still breaks.
                    if (IsInteriorHyperlinkMarker(next)
                        && HasSameHyperlinkRunAhead(runs, j + 1, url, anchor, hlParent))
                    {
                        pendingInterior.Add(next);
                        j++;
                        continue;
                    }
                    break;
                }
                next.Format.TryGetValue("url", out var nUrlObj);
                next.Format.TryGetValue("anchor", out var nAncObj);
                next.Format.TryGetValue("_hyperlinkParent", out var nHlpObj);
                var nUrl = nUrlObj?.ToString();
                var nAnchor = nAncObj?.ToString();
                var nHlParent = nHlpObj?.ToString();
                if (!string.Equals(nUrl, url, StringComparison.Ordinal)) break;
                if (!string.Equals(nAnchor, anchor, StringComparison.Ordinal)) break;
                if (!string.Equals(nHlParent, hlParent, StringComparison.Ordinal)) break;
                group.Add(next);
                sb.Append(next.Type == "tab" ? "\t" : (next.Text ?? ""));
                j++;
            }
            // BUG-R12A(BUG1): the flat `add hyperlink text=Part1Part2` fast path
            // drops every per-run rPr on the wrapped runs (a 2nd bold+red run
            // round-trips as plain text). Keep the fast path ONLY for the lossless
            // case: a single run with no run-level formatting. Otherwise stash the
            // original group runs in Children and flag the node so
            // EmitPlainOrHyperlinkRun emits the wrapper + structured `add run`
            // rows (mirrors the R9 comment-body fix), preserving each run's rPr.
            bool needsStructured = group.Count > 1 || group.Any(HyperlinkRunHasRunFormatting);
            var merged = new DocumentNode
            {
                Path = run.Path,
                // Normalize to "run": a group whose FIRST member is a tab must
                // not surface as a tab-typed node, or TryEmitTabRun would
                // intercept it ahead of the hyperlink emit.
                Type = "run",
                Text = sb.ToString(),
                Format = new Dictionary<string, object?>(run.Format, StringComparer.OrdinalIgnoreCase),
            };
            if (needsStructured)
            {
                merged.Format["_hlStructured"] = true;
                merged.Children = group;
            }
            result.Add(merged);
            // BUG-DUMP-H99: re-emit any interior bookmark/perm markers right after
            // the merged wrapper (their original document position fell between the
            // coalesced runs; a zero-width anchor is position-equivalent there).
            result.AddRange(pendingInterior);
            i = j;
        }
        return result;
    }

    // BUG-DUMP-H99: zero-width position anchors that may sit interior to a
    // hyperlink's run sequence. Splitting the coalesce walk at one of these
    // fragmented a single <w:hyperlink> into two on round-trip.
    private static bool IsInteriorHyperlinkMarker(DocumentNode n) =>
        n.Type is "bookmark" or "bookmarkEnd" or "permStart" or "permEnd";

    // True when, looking forward from `from`, the next run/r/tab carries the same
    // url + anchor + _hyperlinkParent as the group being coalesced — i.e. the
    // intervening marker(s) are interior to ONE hyperlink rather than separating
    // two distinct same-target hyperlinks (which carry different _hyperlinkParent).
    private static bool HasSameHyperlinkRunAhead(List<DocumentNode> runs, int from,
        string? url, string? anchor, string? hlParent)
    {
        for (int k = from; k < runs.Count; k++)
        {
            var n = runs[k];
            if (n.Type == "run" || n.Type == "r" || n.Type == "tab")
            {
                n.Format.TryGetValue("url", out var uo);
                n.Format.TryGetValue("anchor", out var ao);
                n.Format.TryGetValue("_hyperlinkParent", out var ho);
                return string.Equals(uo?.ToString(), url, StringComparison.Ordinal)
                    && string.Equals(ao?.ToString(), anchor, StringComparison.Ordinal)
                    && string.Equals(ho?.ToString(), hlParent, StringComparison.Ordinal);
            }
            // Skip further interior markers; any other node is a boundary.
            if (IsInteriorHyperlinkMarker(n)) continue;
            return false;
        }
        return false;
    }

    // BUG-DUMP-HOIST: run-level character properties that WordHandler.Navigation
    // surfaces on the paragraph node (via the firstRun fallback) but which must
    // NOT ride on `add p` for multi-run paragraphs — every individual run gets
    // its own `add r` carrying its real props.
    private static readonly HashSet<string> RunCharacterPropsHoistedFromFirstRun = new(StringComparer.OrdinalIgnoreCase)
    {
        "bold", "italic", "size", "color", "underline", "underline.color",
        "strike", "highlight",
        "font.latin", "font.ea", "font.ascii", "font.hAnsi",
        // complex-script siblings populated by ReadComplexScriptRunFormatting
        "bold.cs", "italic.cs", "size.cs", "font.cs",
    };

    private static void StripRunCharacterPropsFromParagraph(Dictionary<string, string> props)
    {
        foreach (var k in RunCharacterPropsHoistedFromFirstRun)
            props.Remove(k);
    }

    // Layer per-stop `add tab` rows under a parent path that already has the
    // host paragraph/style created. tabs is the flat List<Dict> Get exposes.
    private static void EmitTabStops(string parentPath, object? tabsVal, List<BatchItem> items)
    {
        if (tabsVal is not IEnumerable<Dictionary<string, object?>> list) return;
        foreach (var t in list)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (t.TryGetValue("pos", out var p) && p != null) props["pos"] = p.ToString() ?? "";
            if (t.TryGetValue("val", out var v) && v != null) props["val"] = v.ToString() ?? "";
            if (t.TryGetValue("leader", out var l) && l != null) props["leader"] = l.ToString() ?? "";
            if (props.Count == 0 || !props.ContainsKey("pos")) continue;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "tab",
                Props = props
            });
        }
    }
}
