// Copyright 2025 OfficeCLI (officecli.ai)
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
    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent,
                                      BodyEmitContext? ctx = null)
    {
        var pNode = word.Get(sourcePath);

        if (TryEmitDisplayEquation(pNode, parentPath, autoPresent, items)) return;

        // Track source paraId -> target index BEFORE any early-return path
        // (section break, TOC, …). Comments anchored on a section-break or
        // TOC paragraph would otherwise miss the mapping and fall back to
        // /body/p[1], silently retargeting the comment.
        if (ctx?.ParaIdToTargetIdx != null && parentPath == "/body" &&
            pNode.Format.TryGetValue("paraId", out var earlyParaId) && earlyParaId != null)
        {
            ctx.ParaIdToTargetIdx[earlyParaId.ToString()!] = targetIndex;
        }

        if (TryEmitInlineSectionBreak(pNode, parentPath, items, ctx)) return;
        if (TryEmitTocParagraph(pNode, parentPath, items)) return;

        var props = FilterEmittableProps(pNode.Format);
        // BUG-DUMP26-01: numId/numLevel that came from style inheritance
        // (ResolveNumPrFromStyle, no direct w:numPr on the paragraph) must
        // not ride on `add p` — the style already supplies them, and emitting
        // them would semantically promote inherited→explicit on replay.
        // Mirrors the first-run hoist precedent for run-character props
        // inherited from styles.
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
        // When a paragraph carries numId, the abstractNum/num pair is already
        // in /numbering (raw-set wholesale by EmitNumberingRaw). Forwarding
        // numFmt/listStyle/start to AddParagraph triggers ad-hoc
        // numbering-definition creation in WordHandler.Add — Word allocates
        // a fresh numId (1→9, 2→16, …) and the paragraph references the
        // new one, orphaning the original abstract numbering's level rPr
        // (color, bold, custom marker text). Drop those keys so the
        // paragraph just attaches by numId+numLevel to the existing def.
        if (props.ContainsKey("numId"))
        {
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        // Collapse non-TOC field chains (fldChar(begin) + instrText(" PAGE ")
        // + fldChar(separate) + display run(s) + fldChar(end)) into a single
        // synthetic "field" entry. Without this collapse, the subsequent
        // `runs` filter sees only the cached display run and emits the field
        // value as static text — PAGE/REF/SEQ/HYPERLINK/NUMPAGES degrade to
        // their evaluated string and stop auto-updating (BUG-X2-05 / X2-1).
        var fieldEntries = CollapseFieldChains(pNode.Children ?? new List<DocumentNode>());
        // BUG-DUMP5-01/02: include break-typed children in the same ordered
        // list as runs so document-order is preserved on emit.
        var runs = fieldEntries
            .Where(c => c.Type == "run" || c.Type == "r" || c.Type == "picture" || c.Type == "field" || c.Type == "ptab" || c.Type == "break"
                || c.Type == "equation"
                || c.Type == "tab"
                || c.Type == "bookmark")
            .ToList();
        var breaks = runs.Where(c => c.Type == "break").ToList();
        var bookmarks = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "bookmark")
            .ToList();
        var inlineSdts = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "sdt")
            .ToList();

        bool collapseSingleRun = ShouldCollapseSingleRun(runs, breaks.Count, bookmarks.Count, inlineSdts.Count);
        pNode.Format.TryGetValue("tabs", out var pTabs);

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

        // BUG-DUMP4-06: emit inline SdtRun children before the runs loop.
        foreach (var sdt in inlineSdts)
        {
            var sdtProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
            {
                if (sdt.Format.TryGetValue(key, out var v) && v != null)
                {
                    var s = v.ToString() ?? "";
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
        foreach (var run in runs)
        {
            if (TryEmitBookmarkRun(run, paraTargetPath, items, ctx)) continue;
            if (TryEmitBreakRun(run, paraTargetPath, items)) continue;
            if (TryEmitTabRun(run, paraTargetPath, items)) continue;
            if (TryEmitPtabRun(run, paraTargetPath, items)) continue;
            if (TryEmitEquationRun(run, paraTargetPath, items)) continue;
            if (TryEmitFieldRun(run, paraTargetPath, items)) continue;
            if (TryEmitPictureRun(word, run, paraTargetPath, parentPath, targetIndex, items, ctx)) continue;
            if (TryEmitNoteRefRun(run, paraTargetPath, items, ctx)) continue;
            EmitPlainOrHyperlinkRun(run, paraTargetPath, items);
        }
    }

    // ── Extracted helpers (behavior unchanged from inline original) ──

    private static bool TryEmitDisplayEquation(DocumentNode pNode, string parentPath, bool autoPresent, List<BatchItem> items)
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
        // BUG-DUMP19-02: forward block-equation alignment.
        if (pNode.Format.TryGetValue("align", out var eqAlign)
            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
            eqProps["align"] = eqAlign.ToString()!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    private static bool TryEmitInlineSectionBreak(DocumentNode pNode, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
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
                    if (c.Type != "run" && c.Type != "r") return false;
                    if (!string.IsNullOrEmpty(c.Text)) return true;
                    // BUG-DUMP5-08: include empty rStyle-bearing footnote /
                    // endnote refs (their visible text comes via the typed
                    // emit branch below, not from c.Text).
                    if (c.Format.TryGetValue("rStyle", out var rsv)
                        && rsv != null
                        && (string.Equals(rsv.ToString(), "FootnoteReference", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(rsv.ToString(), "EndnoteReference", StringComparison.OrdinalIgnoreCase)))
                        return true;
                    return false;
                })
                .ToList();
            if (carrierRuns.Count > 0)
            {
                var carrierPath = $"/body/p[last()]";
                foreach (var run in carrierRuns)
                {
                    // BUG-DUMP7-11: inline SDT carrier — same prop whitelist
                    // as the body-paragraph inline-SDT branch.
                    if (run.Type == "sdt")
                    {
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
                    if (ctx != null && rStyle == "FootnoteReference")
                    {
                        var noteText = ctx.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                            ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index] : "";
                        ctx.FootnoteCursor.Index++;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "footnote",
                            Props = new() { ["text"] = noteText }
                        });
                        continue;
                    }
                    if (ctx != null && rStyle == "EndnoteReference")
                    {
                        var noteText = ctx.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                            ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index] : "";
                        ctx.EndnoteCursor.Index++;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
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
                        Parent = carrierPath,
                        Type = "r",
                        Props = rProps
                    });
                }
            }
            return true;
        }
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
        var instr = instrChild.Format["instruction"]!.ToString()!;
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

    private static bool ShouldCollapseSingleRun(List<DocumentNode> runs, int breaksCount, int bookmarksCount, int inlineSdtsCount)
    {
        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        if (runs.Count > 1) return false;
        if (breaksCount > 0 || bookmarksCount > 0 || inlineSdtsCount > 0) return false;
        if (runs.Count == 0) return true;
        var r = runs[0];
        // Picture / ptab runs need their own typed `add` rows.
        if (r.Type == "picture" || r.Type == "ptab") return false;
        // BUG-DUMP6-05 / BUG-DUMP10-05: hyperlink-wrapped run (url, anchor,
        // or tooltip-only via isHyperlink sentinel) must re-emit as
        // `add hyperlink` — `add p` does not consume url/anchor.
        if (r.Format.ContainsKey("url") || r.Format.ContainsKey("anchor")
            || r.Format.ContainsKey("isHyperlink")) return false;
        // BUG-FUZZ-2: footnote/endnote reference runs need the typed
        // `add footnote/endnote` branch; AddParagraph doesn't consume rStyle.
        if (r.Format.TryGetValue("rStyle", out var srStyle)
            && (string.Equals(srStyle?.ToString(), "FootnoteReference", StringComparison.OrdinalIgnoreCase)
                || string.Equals(srStyle?.ToString(), "EndnoteReference", StringComparison.OrdinalIgnoreCase)))
            return false;
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
            || r.Format.ContainsKey("trackChange")
            || r.Format.ContainsKey("sym")) return false;
        // BUG-FIELD-COLLAPSE: a synthetic field run carries `instruction=…` —
        // collapse would lose the field chain on replay.
        if (r.Type == "field") return false;
        // BUG-DUMP7-03: inline equation must emit `add equation` explicitly.
        if (r.Type == "equation") return false;
        return true;
    }

    private static bool TryEmitBookmarkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP25-01: bookmark child emitted in DOM order so a
        // BookmarkStart between runs survives round-trip at its original
        // intra-paragraph offset. Deferred bookmarks (endPara=true) are
        // pushed onto ctx.DeferredBookmarks so the End sibling can land in
        // a downstream paragraph.
        if (run.Type != "bookmark") return false;
        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (run.Format.TryGetValue("name", out var bmName) && bmName != null)
        {
            var s = bmName.ToString();
            if (!string.IsNullOrEmpty(s)) bmProps["name"] = s;
        }
        if (bmProps.Count == 0) return true; // skip unnamed/anonymous bookmarks
        bool deferred = false;
        if (run.Format.TryGetValue("endPara", out var bmEnd) && bmEnd != null)
        {
            var s = bmEnd.ToString();
            if (!string.IsNullOrEmpty(s) && s != "0")
            {
                bmProps["endPara"] = s;
                deferred = true;
            }
        }
        var bmItem = new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "bookmark",
            Props = bmProps
        };
        if (deferred && ctx != null)
            ctx.DeferredBookmarks.Add(bmItem);
        else
            items.Add(bmItem);
        return true;
    }

    private static bool TryEmitBreakRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP5-01/02: a soft <w:br/> with NO type attribute is a line
        // break, not a page break — fall back to type=line. Emitted inline
        // from the unified runs loop so each break stays at its source
        // position instead of being hoisted to the front of the paragraph.
        if (run.Type != "break") return false;
        var breakType = run.Format.TryGetValue("breakType", out var bt) ? bt?.ToString() : null;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "pagebreak",
            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = string.IsNullOrEmpty(breakType) ? "line" : breakType!
            }
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
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tabParent,
            Type = "r",
            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["text"] = "\t"
            }
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
        var ptabProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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

    private static bool TryEmitEquationRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
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
        var eqParent = paraTargetPath;
        if (!string.IsNullOrEmpty(run.Path))
        {
            var idxEq = run.Path.LastIndexOf("/equation[", StringComparison.Ordinal);
            if (idxEq > 0)
            {
                var derived = run.Path.Substring(0, idxEq);
                if (derived.Contains("/hyperlink["))
                    eqParent = derived;
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

    private static bool TryEmitFieldRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
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
        var instr = run.Format.TryGetValue("instruction", out var iv)
            ? iv?.ToString() ?? "" : "";
        var fieldProps = BuildFieldAddProps(instr, run.Text ?? "");
        if (fieldProps != null
            && run.Format.TryGetValue("_noFieldSeparator", out var nfs)
            && nfs is bool nfsB && nfsB)
        {
            fieldProps["noSeparator"] = "true";
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
                        int emittedHls = items.Count(it => it.Type == "hyperlink"
                            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
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

    private static bool TryEmitPictureRun(WordHandler word, DocumentNode run, string paraTargetPath, string parentPath, int targetIndex, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Drawing-bearing runs surface as type="picture" regardless of
        // whether the Drawing wraps an image (Blip) or a chart (c:chart).
        // Try the image path first; if no embedded image part the run is a
        // chart anchor — pull the next pre-resolved ChartSpec and emit a
        // typed `add chart` row.
        if (run.Type != "picture") return false;
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
            var chartProps = BuildChartProps(spec);
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
        // shape. No typed Add path exists yet, but the XML is self-contained
        // (no rId/embed back-references) so round-trip via raw-set append is
        // safe. Targets the already-created paragraph by xpath positional
        // index. Caveats: drawings with embedded image references would also
        // land here and silently lose their image part — acceptable v0.5
        // lossy mode.
        var rawXml = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(rawXml) &&
            parentPath == "/body" &&
            !rawXml.Contains("r:embed") && !rawXml.Contains("r:id"))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/document",
                Xpath = $"/w:document/w:body/w:p[{targetIndex}]",
                Action = "append",
                Xml = rawXml
            });
        }
        return true;
    }

    private static bool TryEmitNoteRefRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Footnote/endnote reference runs are empty <w:r> elements with
        // rStyle = FootnoteReference / EndnoteReference. Emit them as a
        // typed footnote/endnote add anchored on the host paragraph and
        // pull the body text from the pre-resolved ordered list — see
        // BodyEmitContext for the document-order assumption.
        if (ctx == null) return false;
        var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
        if (rStyle == "FootnoteReference")
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
            return true;
        }
        if (rStyle == "EndnoteReference")
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
            return true;
        }
        return false;
    }

    private static void EmitPlainOrHyperlinkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;

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
            // echoes back so we don't override those defaults.
            if (rProps.TryGetValue("color", out var hlColor)
                && string.Equals(hlColor, "hyperlink", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("color");
            if (rProps.TryGetValue("underline", out var hlUl)
                && string.Equals(hlUl, "single", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("underline");
            rProps.Remove("isHyperlink");
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
        int emittedHls = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        return emittedHls >= kIdx ? rebased : paraTargetPath;
    }

    private static List<DocumentNode> CoalesceHyperlinkRuns(List<DocumentNode> runs)
    {
        var result = new List<DocumentNode>(runs.Count);
        int i = 0;
        while (i < runs.Count)
        {
            var run = runs[i];
            string? url = null, anchor = null;
            if (run.Type == "run" || run.Type == "r")
            {
                if (run.Format.TryGetValue("url", out var u))
                    url = u?.ToString();
                if (run.Format.TryGetValue("anchor", out var a))
                    anchor = a?.ToString();
            }
            if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(anchor))
            {
                result.Add(run);
                i++;
                continue;
            }
            // Walk forward over consecutive runs with the same url/anchor.
            int j = i + 1;
            var sb = new System.Text.StringBuilder(run.Text ?? "");
            while (j < runs.Count)
            {
                var next = runs[j];
                if (next.Type != "run" && next.Type != "r") break;
                next.Format.TryGetValue("url", out var nUrlObj);
                next.Format.TryGetValue("anchor", out var nAncObj);
                var nUrl = nUrlObj?.ToString();
                var nAnchor = nAncObj?.ToString();
                if (!string.Equals(nUrl, url, StringComparison.Ordinal)) break;
                if (!string.Equals(nAnchor, anchor, StringComparison.Ordinal)) break;
                sb.Append(next.Text ?? "");
                j++;
            }
            if (j == i + 1)
            {
                // No coalescing — emit the single run as-is.
                result.Add(run);
            }
            else
            {
                var merged = new DocumentNode
                {
                    Path = run.Path,
                    Type = run.Type,
                    Text = sb.ToString(),
                    Format = new Dictionary<string, object?>(run.Format, StringComparer.OrdinalIgnoreCase),
                };
                result.Add(merged);
            }
            i = j;
        }
        return result;
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
