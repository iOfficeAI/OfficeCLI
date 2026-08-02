// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Find / Format / Replace ====================

    /// <summary>
    /// Build a flat list of (Run, Text, charStart, charEnd) spans for a paragraph.
    /// Uses Descendants to include runs inside hyperlinks, w:ins, w:del, etc.
    /// Shared by ProcessFindInParagraph, SplitRunsAtRange, etc.
    /// </summary>
    private static List<(Run Run, Text TextElement, int Start, int End)> BuildRunTexts(Paragraph para)
    {
        var runTexts = new List<(Run Run, Text TextElement, int Start, int End)>();
        int pos = 0;
        foreach (var run in para.Descendants<Run>())
        {
            foreach (var text in run.Elements<Text>())
            {
                var len = text.Text?.Length ?? 0;
                if (len > 0)
                    runTexts.Add((run, text, pos, pos + len));
                pos += len;
            }
        }
        return runTexts;
    }

    // ==================== NEWLINE-SEMANTICS-V2 (issue #262) ====================
    //
    // find/replace speaks the same newline vocabulary as add/get: '\v' is a
    // soft line break (<w:br/>), '\n' is a paragraph boundary. Breaks carry
    // no text, so to make '\v' matchable the paragraph's breaks are
    // MATERIALIZED into literal '\v' chars inside Text nodes before the
    // find; after the replace, one normalization pass converts every '\v'
    // char (leftover or newly inserted by the replacement string) back into
    // a <w:br/> element and every '\n' char into a paragraph split via
    // SplitParagraphAtOffset — each new paragraph inheriting the original
    // pPr (style/numbering/indent), which is exactly the Shift+Enter →
    // Enter conversion issue #262 asked for:
    //   officecli set doc.docx /body --find $'\v' --replace $'\n'
    // The chars exist only in-memory between the two passes; nothing
    // XML-illegal is ever serialized.

    /// <summary>Convert textWrapping breaks / CarriageReturns into literal '\v' Text nodes.</summary>
    private static void MaterializeSoftBreakChars(Paragraph para)
    {
        foreach (var run in para.Descendants<Run>().ToList())
        {
            foreach (var child in run.ChildElements.ToList())
            {
                bool isSoftBreak = child is CarriageReturn
                    || (child is Break b && (b.Type == null || b.Type.Value == BreakValues.TextWrapping));
                if (!isSoftBreak) continue;
                var marker = new Text("\v") { Space = SpaceProcessingModeValues.Preserve };
                run.ReplaceChild(marker, child);
            }
        }
    }

    /// <summary>
    /// Convert literal '\v' chars back into <w:br/> elements and '\n' chars
    /// into paragraph splits. Returns the number of paragraph splits made.
    /// </summary>
    private static int NormalizeNewlineChars(Paragraph para)
    {
        // Pass 1: '\v' → <w:br/>, rebuilt in place inside each Text's run.
        foreach (var run in para.Descendants<Run>().ToList())
        {
            foreach (var textEl in run.Elements<Text>().ToList())
            {
                var s = textEl.Text ?? "";
                if (s.IndexOf('\v') < 0) continue;
                OpenXmlElement anchor = textEl;
                var segments = s.Split('\v');
                for (int i = 0; i < segments.Length; i++)
                {
                    if (segments[i].Length > 0)
                    {
                        var t = new Text(segments[i]) { Space = SpaceProcessingModeValues.Preserve };
                        run.InsertAfter(t, anchor); anchor = t;
                    }
                    if (i < segments.Length - 1)
                    {
                        var br = new Break();
                        run.InsertAfter(br, anchor); anchor = br;
                    }
                }
                textEl.Remove();
            }
        }

        // Pass 2: '\n' → paragraph split. Work on the CURRENT tail paragraph
        // each round; every split moves the remaining text (and any further
        // '\n' chars) into the tail.
        int splits = 0;
        var current = para;
        while (true)
        {
            var runTexts = BuildRunTexts(current);
            var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
            var nl = fullText.IndexOf('\n');
            if (nl < 0) break;

            // Delete the '\n' char itself from its Text node first, so the
            // split offset addresses the char-free stream.
            foreach (var rt in runTexts)
            {
                if (nl >= rt.Start && nl < rt.End)
                {
                    var local = nl - rt.Start;
                    rt.TextElement.Text = rt.TextElement.Text!.Remove(local, 1);
                    rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    break;
                }
            }

            var total = fullText.Length - 1;
            splits++;
            if (nl <= 0)
            {
                // Boundary: '\n' at paragraph start → empty head, content tail.
                var head = new Paragraph();
                if (current.ParagraphProperties != null)
                    head.PrependChild((ParagraphProperties)current.ParagraphProperties.CloneNode(true));
                current.Parent!.InsertBefore(head, current);
                continue; // remaining '\n's still in `current`
            }
            if (nl >= total)
            {
                // Boundary: '\n' at paragraph end → content head, empty tail.
                var tail = new Paragraph();
                if (current.ParagraphProperties != null)
                    tail.PrependChild((ParagraphProperties)current.ParagraphProperties.CloneNode(true));
                current.Parent!.InsertAfter(tail, current);
                current = tail;
                continue;
            }
            current = SplitParagraphAtOffset(current, nl);
        }
        return splits;
    }

    /// <summary>
    /// Split a paragraph at the given character offset, producing a head
    /// paragraph (the original <paramref name="para"/>, now holding
    /// runs/content up to <paramref name="charOffset"/>) followed by a tail
    /// paragraph inserted as its immediate next sibling (holding content
    /// from <paramref name="charOffset"/> onward). The tail inherits a
    /// clone of the head's paragraph properties so style/numbering/heading
    /// is preserved on both halves — matching Word's own Enter-key split.
    /// Preconditions: 0 &lt; charOffset &lt; fullText length (boundary cases
    /// should be handled by the caller without splitting).
    /// </summary>
    private static Paragraph SplitParagraphAtOffset(Paragraph para, int charOffset)
    {
        var runTexts = BuildRunTexts(para);

        // Split the run that straddles charOffset so a clean run boundary
        // exists at the split point. After this call, runTexts is stale.
        // Two straddle shapes: the offset falls INSIDE one Text node, or —
        // when a run holds several Text nodes (e.g. after
        // MaterializeSoftBreakChars) — exactly ON the boundary BETWEEN two
        // Text nodes of the same run. The old strict inside-a-Text check
        // missed the second shape, so the run-level partition below kept the
        // whole run on the head side and produced an empty tail.
        foreach (var runGroup in runTexts.GroupBy(rt => rt.Run))
        {
            var runStart = runGroup.Min(rt => rt.Start);
            var runEnd = runGroup.Max(rt => rt.End);
            if (charOffset <= runStart || charOffset >= runEnd) continue;

            var inside = runGroup.FirstOrDefault(rt => charOffset > rt.Start && charOffset < rt.End);
            if (inside.TextElement != null)
                SplitRunAtOffset(inside.Run, inside.TextElement, charOffset - inside.Start);
            else
            {
                // Boundary between two Text nodes of the same run: split the
                // run BEFORE the boundary Text (SplitRunAtOffset no-ops on a
                // local offset of 0 by contract). The boundary Text plus every
                // subsequent child moves into a fresh right-hand run.
                var atBoundary = runGroup.First(rt => rt.Start == charOffset);
                var hostRun = atBoundary.Run;
                var rightRun = new Run();
                if (hostRun.RunProperties != null)
                    rightRun.AppendChild((RunProperties)hostRun.RunProperties.CloneNode(true));
                bool after = false;
                var move = new List<OpenXmlElement>();
                foreach (var child in hostRun.ChildElements)
                {
                    if (ReferenceEquals(child, atBoundary.TextElement)) after = true;
                    if (after && child is not RunProperties) move.Add(child);
                }
                foreach (var el in move) { el.Remove(); rightRun.AppendChild(el); }
                hostRun.InsertAfterSelf(rightRun);
            }
            break;
        }

        // Recompute run positions and partition runs into head (< charOffset)
        // and tail (>= charOffset). Inline children other than Run
        // (hyperlink/bookmark/field/sdt/…) are routed by their document
        // order relative to the cumulative text length: anything whose
        // text footprint falls entirely on the tail side moves with the
        // tail paragraph. Runs with zero-length text at the boundary stay
        // with the head (matches Enter-key behavior in Word).
        var tail = new Paragraph();
        if (para.ParagraphProperties != null)
            tail.PrependChild((ParagraphProperties)para.ParagraphProperties.CloneNode(true));

        // Walk children in document order. For Run, compute its text range
        // and decide; for non-Run inline children, treat their text contribution
        // as zero-length at the current cumulative offset (consistent with how
        // BuildRunTexts ignores them).
        int cumulative = 0;
        var toMove = new List<OpenXmlElement>();
        foreach (var child in para.ChildElements.ToList())
        {
            if (child is ParagraphProperties) continue;
            if (child is Run run)
            {
                var runLen = run.Elements<Text>().Sum(t => t.Text?.Length ?? 0);
                if (cumulative >= charOffset)
                {
                    toMove.Add(child);
                }
                cumulative += runLen;
            }
            else
            {
                // Non-run inline content: keep on head side if we're still
                // before the split point, move to tail if we've crossed it.
                if (cumulative >= charOffset)
                    toMove.Add(child);
            }
        }

        foreach (var el in toMove)
        {
            el.Remove();
            tail.AppendChild(el);
        }

        para.InsertAfterSelf(tail);
        return tail;
    }

    private static Run SplitRunAtOffset(Run run, Text splitText, int charOffset)
    {
        // Split `run` at `charOffset` INSIDE `splitText` (text-local, the
        // offset space every caller already computes from BuildRunTexts).
        // The right half is a fresh run carrying only cloned run properties;
        // the split text's right part plus EVERY subsequent child (w:t,
        // w:tab, w:br, …) moves into it, in document order.
        //
        // The old implementation cloned the whole run and patched Text
        // elements by index: non-Text children (tabs, breaks) ended up
        // duplicated on BOTH sides (one tab became three across a two-cut
        // range split), post-split Texts stayed on the left as residue, and
        // it interpreted the offset as run-local while callers passed
        // text-local — so a range set on a tab-carrying paragraph corrupted
        // content and dropped the formatting.
        var full = splitText.Text ?? "";
        if (charOffset <= 0 || charOffset >= full.Length)
            return run; // boundary — nothing to split (callers guard, keep old behavior)

        var rightRun = new Run();
        if (run.RunProperties != null)
            rightRun.AppendChild((RunProperties)run.RunProperties.CloneNode(true));

        splitText.Text = full[..charOffset];
        splitText.Space = SpaceProcessingModeValues.Preserve;

        var toMove = new List<OpenXmlElement>();
        bool afterSplit = false;
        foreach (var child in run.ChildElements)
        {
            if (ReferenceEquals(child, splitText)) { afterSplit = true; continue; }
            if (afterSplit && child is not RunProperties) toMove.Add(child);
        }

        rightRun.AppendChild(new Text(full[charOffset..]) { Space = SpaceProcessingModeValues.Preserve });
        foreach (var el in toMove)
        {
            el.Remove();
            rightRun.AppendChild(el);
        }

        run.InsertAfterSelf(rightRun);
        return rightRun;
    }

    /// <summary>
    /// Apply run formatting to an explicit character range. This is <c>find</c>
    /// with the text-match step short-circuited: the caller supplies the
    /// [start,end) offsets directly (0-based, half-open) instead of a pattern,
    /// and the same per-paragraph split-run + ApplyRunFormatting path runs. Kept
    /// isomorphic with <see cref="ProcessFind(string,string,string?,Dictionary{string,string},Dictionary{string,string}?,out List{Paragraph})"/>:
    /// same scope resolution (ResolveParagraphsForFind, incl. header/footer/note
    /// sweep), same paraId regeneration on changed paragraphs, and — like find —
    /// it does NOT save (the handler's normal save path flushes). Offsets are
    /// relative to the concatenated run text of the resolved scope; a range that
    /// straddles a paragraph boundary formats each covered paragraph's slice.
    /// Run-level only — keys ApplyRunFormatting does not handle (paragraph-level
    /// props) are returned as unsupported for the caller to self-report
    /// (handler-as-truth). Sets LastFindMatchCount to the number of runs formatted.
    /// </summary>
    private List<string> ProcessWordRange(string path, IReadOnlyList<(int Start, int End)> ranges, Dictionary<string, string> formatProps)
    {
        var paragraphs = ResolveParagraphsForFind(path);
        if (paragraphs.Count == 0)
            throw new ArgumentException($"No paragraphs found at path: {path}");

        int totalLen = 0;
        foreach (var para in paragraphs)
        {
            var rts = BuildRunTexts(para);
            totalLen += rts.Count > 0 ? rts[^1].End : 0;
        }
        foreach (var (s, e) in ranges)
            if (s > totalLen || e > totalLen)
                throw new ArgumentException(
                    $"range end {e} out of bounds (scope text has {totalLen} chars).");

        var unsupported = new List<string>();
        var changedParas = new HashSet<Paragraph>();
        int applied = 0;
        foreach (var (start, end) in ranges)
        {
            int cursor = 0;
            foreach (var para in paragraphs)
            {
                var rts = BuildRunTexts(para);
                int paraLen = rts.Count > 0 ? rts[^1].End : 0;
                int paraStart = cursor;
                int paraEnd = cursor + paraLen;
                cursor = paraEnd;

                int localStart = Math.Max(start, paraStart) - paraStart;
                int localEnd = Math.Min(end, paraEnd) - paraStart;
                if (localStart >= localEnd) continue; // no (non-empty) overlap here

                var targetRuns = SplitRunsAtRange(para, localStart, localEnd);
                if (targetRuns.Count == 0) continue;
                foreach (var run in targetRuns)
                {
                    var rPr = EnsureRunProperties(run);
                    foreach (var (key, value) in formatProps)
                        if (!ApplyRunFormatting(rPr, key, value) && !unsupported.Contains(key))
                            unsupported.Add(key);
                    applied++;
                }
                changedParas.Add(para);
            }
        }

        // Paragraph content structurally changed (runs split) — regenerate paraId
        // exactly as ProcessFind does for a matched paragraph.
        foreach (var para in changedParas)
            para.TextId = GenerateParaId();

        LastFindMatchCount = applied;
        return unsupported;
    }

    /// <summary>
    /// Split runs in a paragraph so that the character range [charStart, charEnd)
    /// is covered by dedicated runs. Returns the list of runs covering that range.
    /// </summary>
    private static List<Run> SplitRunsAtRange(Paragraph para, int charStart, int charEnd)
    {
        // Split at charEnd first (so charStart offsets remain valid)
        var runTexts = BuildRunTexts(para);
        foreach (var rt in runTexts)
        {
            if (charEnd > rt.Start && charEnd < rt.End)
            {
                var localOffset = charEnd - rt.Start;
                SplitRunAtOffset(rt.Run, rt.TextElement, localOffset);
                break;
            }
        }

        // Rebuild after split, then split at charStart
        runTexts = BuildRunTexts(para);
        foreach (var rt in runTexts)
        {
            if (charStart > rt.Start && charStart < rt.End)
            {
                var localOffset = charStart - rt.Start;
                SplitRunAtOffset(rt.Run, rt.TextElement, localOffset);
                break;
            }
        }

        // Rebuild and collect runs covering [charStart, charEnd)
        runTexts = BuildRunTexts(para);
        var result = new List<Run>();
        foreach (var rt in runTexts)
        {
            if (rt.Start >= charStart && rt.End <= charEnd)
                result.Add(rt.Run);
        }
        return result;
    }

    /// <summary>
    /// Unified find operation on a paragraph: replace text and/or apply formatting.
    /// Returns the number of matches processed.
    ///
    /// When <paramref name="revisionProps"/> is non-null, every change becomes a
    /// tracked revision:
    ///   - text replace → matched runs wrapped in w:del, replacement run wrapped in w:ins
    ///   - format-only  → each matched run gets a w:rPrChange snapshot of its prior rPr
    /// Wrapping reuses <c>WrapRunAsDeleted</c> / <c>WrapRunAsInserted</c> + the same
    /// rPrChange decorator used by <c>set /body/p[N]/r[M] --prop font.color=… --prop
    /// revision.author=…</c> (see WordHandler.Set.Revision.cs), so the marker shape is
    /// byte-equivalent to the non-find path. Each match gets fresh revision ids
    /// (one for the w:del span, one for the w:ins) so accept/reject by id works.
    /// </summary>
    private int ProcessFindInParagraph(
        Paragraph para,
        string pattern,
        bool isRegex,
        string? replace,
        Dictionary<string, string>? formatProps,
        Dictionary<string, string>? revisionProps,
        (int Start, int End)? runScope = null)
    {
        // NEWLINE-SEMANTICS-V2: materialize soft breaks as '\v' chars when the
        // pattern wants to match them or the replacement inserts newlines —
        // the whole existing offset machinery then works on plain chars, and
        // one normalization pass at the end restores element form (and
        // performs '\n' paragraph splits). Run-scoped finds skip this: their
        // span was computed against the unmaterialized stream.
        bool newlineAware = !runScope.HasValue && (
            pattern.Contains('\v') || (isRegex && pattern.Contains(@"\v"))
            || (replace != null && (replace.Contains('\v') || replace.Contains('\n'))));
        if (newlineAware)
        {
            if (replace != null && replace.Contains('\n') && revisionProps is { Count: > 0 })
                throw new ArgumentException(
                    "find/replace with revision tracking cannot insert a paragraph boundary ('\\n' in replace). "
                    + "Use '\\v' for a soft line break, or run the replace without revision.* props.");
            MaterializeSoftBreakChars(para);
        }

        var runTexts = BuildRunTexts(para);
        if (runTexts.Count == 0) { if (newlineAware) NormalizeNewlineChars(para); return 0; }

        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
        // CONSISTENCY(regex-backref-expand): collect Match objects in regex mode so we can
        // call Match.Result(replace) — which expands backreferences against the original
        // match captures, and unlike re-running Regex.Replace on the substring, correctly
        // handles lookaround anchors (e.g. r"foo(?=bar)") whose context is lost in isolation.
        // BUG-TESTER+FUZZER R31: wrap with try/catch so RegexMatchTimeoutException is
        // converted to ArgumentException (consistent with FindMatchRanges), and avoid
        // a second Regex.Matches call by deriving ranges from the same Match list.
        List<System.Text.RegularExpressions.Match>? matchObjs = null;
        List<(int Start, int Length)> matches;
        if (isRegex)
        {
            try
            {
                matchObjs = System.Text.RegularExpressions.Regex.Matches(
                        fullText,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.None,
                        FindHelpers.RegexMatchTimeout)
                    .Cast<System.Text.RegularExpressions.Match>()
                    .Where(m => m.Length > 0)
                    .ToList();
            }
            catch (System.Text.RegularExpressions.RegexParseException ex)
            {
                throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}", ex);
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException ex)
            {
                throw new ArgumentException(
                    $"Regex pattern '{pattern}' exceeded {FindHelpers.RegexMatchTimeout.TotalSeconds}s match timeout (catastrophic backtracking?)",
                    ex);
            }
            matches = matchObjs.Select(m => (m.Index, m.Length)).ToList();
        }
        else
        {
            matches = FindHelpers.FindMatchRanges(fullText, pattern, isRegex);
        }
        if (matches.Count == 0) { if (newlineAware) NormalizeNewlineChars(para); return 0; }

        // CONSISTENCY(find-run-scope): when the find scope is a single run
        // (/body/p[N]/r[K]), keep only matches fully contained in that run's
        // character span — mirrors PPTX ProcessFindInPptParagraph's runIndexFilter
        // (R32). The span is passed in (computed from the resolved run element in
        // the same BuildRunTexts coordinate) rather than a run index, because a
        // Word run may hold several <w:t> children — index-based counting would
        // misalign. Filter both `matches` and `matchObjs` together to keep the
        // regex-backref list in step.
        if (runScope.HasValue)
        {
            var (scopeStart, scopeEnd) = runScope.Value;
            var keepIdx = new HashSet<int>();
            for (int k = 0; k < matches.Count; k++)
            {
                var (s, l) = matches[k];
                if (s >= scopeStart && s + l <= scopeEnd) keepIdx.Add(k);
            }
            matches = matches.Where((_, k) => keepIdx.Contains(k)).ToList();
            if (matchObjs != null)
                matchObjs = matchObjs.Where((_, k) => keepIdx.Contains(k)).ToList();
            if (matches.Count == 0) return 0;
        }

        // Process from end to start to preserve character offsets
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            var (matchStart, matchLen) = matches[i];
            var matchEnd = matchStart + matchLen;

            // ---- find + revision: branch off BEFORE the legacy non-tracked
            //      paths so a stray revisionProps can't silently degrade into
            //      a destructive in-place edit. Layered:
            //        replace != null  → w:del fragments + w:ins replacement
            //        else             → w:rPrChange per matched run
            //      Format props (if also present with replace) are applied to
            //      the inserted run so the new text gets the requested look.
            if (revisionProps != null && revisionProps.Count > 0)
            {
                // Resolve revision attribution defaults once per match.
                string author = revisionProps.TryGetValue("revision.author", out var a) && !string.IsNullOrEmpty(a)
                    ? a : "OfficeCLI";
                DateTime date = DateTime.UtcNow;
                if (revisionProps.TryGetValue("revision.date", out var dStr)
                    && !string.IsNullOrEmpty(dStr)
                    && DateTime.TryParse(dStr, out var parsedDate))
                    date = parsedDate;

                if (replace != null)
                {
                    string effectiveReplace = replace;
                    if (isRegex && matchObjs != null && i < matchObjs.Count)
                        effectiveReplace = matchObjs[i].Result(replace);

                    // Cross-hyperlink replacement is still rejected — the wrapped
                    // form would corrupt the URL/format-binding of the hyperlink
                    // structure just as the unwrapped form did.
                    {
                        var affected = BuildRunTexts(para)
                            .Where(rt => rt.End > matchStart && rt.Start < matchEnd)
                            .Select(rt => rt.Run.Ancestors<Hyperlink>().FirstOrDefault())
                            .Distinct()
                            .ToList();
                        if (affected.Count > 1)
                            throw new ArgumentException(
                                $"find/replace+revision cannot span a hyperlink boundary "
                                + $"(match at offset {matchStart}, length {matchLen})");
                    }

                    // Split the runs so the matched span is a contiguous list of
                    // sibling runs we can wrap individually.
                    var targetRuns = SplitRunsAtRange(para, matchStart, matchEnd);
                    if (targetRuns.Count == 0) continue;

                    // Guard: matched runs must not already be inside a revision
                    // wrapper — stacking ins/del muddies accept/reject semantics
                    // (mirrors the BeginTrackChangeIfRequested guard for `set`).
                    foreach (var run in targetRuns)
                    {
                        if (run.Ancestors<InsertedRun>().Any()
                            || run.Ancestors<DeletedRun>().Any()
                            || run.Ancestors<MoveFromRun>().Any()
                            || run.Ancestors<MoveToRun>().Any())
                            throw new InvalidOperationException(
                                $"find/replace+revision: matched run at offset {matchStart} "
                                + "is already inside a revision wrapper; accept/reject the "
                                + "existing marker first");
                    }

                    // Template rPr for the inserted run: clone the first matched
                    // run's rPr so the replacement inherits the original look
                    // (font, size, color), then layer formatProps on top.
                    RunProperties templateRPr;
                    var firstRPr = targetRuns[0].GetFirstChild<RunProperties>();
                    templateRPr = firstRPr != null
                        ? (RunProperties)firstRPr.CloneNode(true)
                        : new RunProperties();
                    // Strip any prior rPrChange off the clone — it belongs to
                    // the source run's history, not to the inserted run.
                    foreach (var rprc in templateRPr.Elements<RunPropertiesChange>().ToList())
                        rprc.Remove();
                    if (formatProps != null)
                    {
                        foreach (var (key, value) in formatProps)
                            ApplyRunFormatting(templateRPr, key, value);
                    }

                    // Wrap each matched run as w:del. WrapRunAsDeleted returns the
                    // wrapper so we can locate the insertion point for the w:ins
                    // (immediately after the last w:del wrapper).
                    DeletedRun? lastDelWrapper = null;
                    foreach (var run in targetRuns)
                    {
                        var w = WrapRunAsDeleted(run, author, date, null);
                        if (w != null) lastDelWrapper = w;
                    }

                    // Insert replacement (skip if user passed --prop replace="" — a
                    // deletion-only operation). The new w:ins sibling sits right
                    // after the last w:del wrapper.
                    if (!string.IsNullOrEmpty(effectiveReplace) && lastDelWrapper?.Parent != null)
                    {
                        var newRun = new Run(
                            templateRPr,
                            new Text(effectiveReplace) { Space = SpaceProcessingModeValues.Preserve });
                        lastDelWrapper.Parent.InsertAfter(newRun, lastDelWrapper);
                        WrapRunAsInserted(newRun, author, date, null);
                    }
                }
                else
                {
                    // format-only + revision: per matched run, snapshot rPr →
                    // apply format → append w:rPrChange. Reuses
                    // BeginTrackChangeIfRequested so the snapshot shape is
                    // byte-identical to the `set /body/p[N]/r[M] --prop … --prop
                    // revision.author=…` path.
                    if (formatProps == null || formatProps.Count == 0) continue;

                    var targetRuns = SplitRunsAtRange(para, matchStart, matchEnd);
                    foreach (var run in targetRuns)
                    {
                        // Each run uses its own freshly-generated id so accept/reject
                        // by /revision[@id=N] addresses them individually.
                        var combined = new Dictionary<string, string>(formatProps, StringComparer.OrdinalIgnoreCase);
                        foreach (var (rk, rv) in revisionProps)
                            combined[rk] = rv;
                        var (stripped, wrap) = BeginTrackChangeIfRequested(run, combined);
                        var rPr = EnsureRunProperties(run);
                        foreach (var (key, value) in stripped)
                            ApplyRunFormatting(rPr, key, value);
                        wrap();
                    }
                }
                continue;
            }

            if (replace != null)
            {
                // For regex replace, expand backreferences ($1, ${name}, etc.) via
                // Match.Result so lookaround context is preserved.
                string effectiveReplace = replace;
                if (isRegex && matchObjs != null && i < matchObjs.Count)
                {
                    effectiveReplace = matchObjs[i].Result(replace);
                }

                // BUG-BT-2: detect cross-hyperlink-boundary replacement. If the
                // match spans runs whose Hyperlink ancestors differ (e.g. one
                // run inside a Hyperlink, another in plain paragraph body),
                // a naive cross-run text edit destroys the hyperlink structure
                // (URL + blue/underline formatting are lost). Reject up-front
                // with a clear error rather than silently corrupting the doc.
                {
                    var affected = BuildRunTexts(para)
                        .Where(rt => rt.End > matchStart && rt.Start < matchEnd)
                        .Select(rt => rt.Run.Ancestors<Hyperlink>().FirstOrDefault())
                        .Distinct()
                        .ToList();
                    if (affected.Count > 1)
                    {
                        throw new ArgumentException(
                            $"find/replace cannot span a hyperlink boundary (match at offset {matchStart}, length {matchLen}): " +
                            $"the match crosses into or out of a <w:hyperlink>, which would destroy its URL and formatting. " +
                            $"Narrow the pattern to stay inside or outside the hyperlink, or edit the hyperlink text directly.");
                    }
                }

                // Step 1: Replace text in affected runs (same logic as old ReplaceInParagraph)
                var currentRunTexts = BuildRunTexts(para);
                bool first = true;
                foreach (var rt in currentRunTexts)
                {
                    if (rt.End <= matchStart || rt.Start >= matchEnd)
                        continue;

                    var textStr = rt.TextElement.Text ?? "";
                    var localStart = Math.Max(0, matchStart - rt.Start);
                    var localEnd = Math.Min(textStr.Length, matchEnd - rt.Start);

                    if (first)
                    {
                        rt.TextElement.Text = textStr[..localStart] + effectiveReplace + textStr[localEnd..];
                        rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                        first = false;
                    }
                    else
                    {
                        rt.TextElement.Text = textStr[..Math.Max(0, matchStart - rt.Start)] + textStr[localEnd..];
                        rt.TextElement.Space = SpaceProcessingModeValues.Preserve;
                    }
                }

                // BUG-TESTER fuzz-1: cross-run replace consumes intermediate runs leaving
                // them with empty <w:t/> — drop those orphan runs so persisted XML stays clean.
                // Only remove runs whose Text element is now empty AND have no other
                // semantic children (Break, TabChar, Drawing, FieldChar, Picture, etc.).
                // RunProperties (rPr) alone is not semantic content.
                var emptyRunsToRemove = new List<Run>();
                foreach (var run in para.Descendants<Run>())
                {
                    bool hasContent = false;
                    bool hasEmptyText = false;
                    foreach (var child in run.ChildElements)
                    {
                        if (child is RunProperties)
                            continue;
                        if (child is Text t)
                        {
                            if (string.IsNullOrEmpty(t.Text))
                                hasEmptyText = true;
                            else
                                hasContent = true;
                        }
                        else
                        {
                            hasContent = true;
                        }
                    }
                    if (hasEmptyText && !hasContent)
                        emptyRunsToRemove.Add(run);
                }
                foreach (var run in emptyRunsToRemove)
                    run.Remove();

                // Step 2: If format props, split at the replaced text position and apply
                if (formatProps != null && formatProps.Count > 0)
                {
                    // The replaced text now starts at matchStart with length = effectiveReplace.Length
                    var replacedEnd = matchStart + effectiveReplace.Length;
                    if (effectiveReplace.Length > 0)
                    {
                        var targetRuns = SplitRunsAtRange(para, matchStart, replacedEnd);
                        foreach (var run in targetRuns)
                        {
                            var rPr = EnsureRunProperties(run);
                            foreach (var (key, value) in formatProps)
                                ApplyRunFormatting(rPr, key, value);
                        }
                    }
                }
            }
            else if (formatProps != null && formatProps.Count > 0)
            {
                // No replace, just split and format
                var targetRuns = SplitRunsAtRange(para, matchStart, matchEnd);
                foreach (var run in targetRuns)
                {
                    var rPr = EnsureRunProperties(run);
                    foreach (var (key, value) in formatProps)
                        ApplyRunFormatting(rPr, key, value);
                }
            }
        }

        // NEWLINE-SEMANTICS-V2: restore '\v' chars to <w:br/> elements and
        // apply '\n' paragraph splits introduced by the replacement string.
        if (newlineAware)
            NormalizeNewlineChars(para);

        return matches.Count;
    }

    /// <summary>
    /// Unified find operation: process find/replace/format across paragraphs resolved from a path.
    /// Called from Set when 'find' key is present.
    /// Returns (matchCount, unsupportedKeys).
    /// </summary>
    private int ProcessFind(
        string path,
        string findValue,
        string? replace,
        Dictionary<string, string> formatProps)
        => ProcessFind(path, findValue, replace, formatProps, null, out _);

    /// <summary>
    /// Overload that surfaces the set of paragraphs whose text actually matched
    /// the find pattern. Callers that follow up with paragraph-scope mutations
    /// (e.g. <c>direction</c>) must filter by this set rather than re-resolving
    /// every paragraph under the path — otherwise <c>find=X --prop direction=rtl</c>
    /// silently rewrites every paragraph in the document. R8-fuzz-1 / R8-fuzz-2.
    ///
    /// <paramref name="revisionProps"/> threaded through to
    /// <see cref="ProcessFindInParagraph"/>; null = legacy non-tracked mode.
    /// </summary>
    private int ProcessFind(
        string path,
        string findValue,
        string? replace,
        Dictionary<string, string> formatProps,
        Dictionary<string, string>? revisionProps,
        out List<Paragraph> matchedParagraphs)
    {
        matchedParagraphs = new List<Paragraph>();
        var (pattern, isRegex) = FindHelpers.ParseFindPattern(findValue);
        if (string.IsNullOrEmpty(pattern) && !isRegex) return 0;

        // Resolve paragraphs from path. A /body/p[N]/r[K] path also surfaces the
        // target run so we can confine matches to its character span; the span is
        // computed from the run element (not an index — a Word run can hold several
        // <w:t>) in the same BuildRunTexts coordinate the match offsets use. A run
        // with no text yields no span → whole-paragraph scope is kept (unchanged).
        var paragraphs = ResolveParagraphsForFind(path, out var scopeRun);
        Paragraph? scopePara = scopeRun?.Ancestors<Paragraph>().FirstOrDefault();
        (int Start, int End)? scopeSpan =
            (scopeRun != null && scopePara != null) ? RunCharSpan(scopePara, scopeRun) : null;

        int totalCount = 0;
        foreach (var para in paragraphs)
        {
            var count = ProcessFindInParagraph(
                para,
                pattern,
                isRegex,
                replace,
                formatProps.Count > 0 ? formatProps : null,
                revisionProps,
                (scopeSpan.HasValue && ReferenceEquals(para, scopePara)) ? scopeSpan : null);
            if (count > 0)
            {
                para.TextId = GenerateParaId();
                matchedParagraphs.Add(para);
            }
            totalCount += count;
        }

        return totalCount;
    }

    /// <summary>
    /// Character span [start,end) of <paramref name="run"/> within
    /// <paramref name="para"/>, in the BuildRunTexts coordinate the find match
    /// offsets use. Null when the run carries no text (nothing to scope to — the
    /// caller then keeps whole-paragraph scope). Robust to a run holding several
    /// &lt;w:t&gt; children (their spans are contiguous, so first.Start..last.End).
    /// </summary>
    private static (int Start, int End)? RunCharSpan(Paragraph para, Run run)
    {
        int? start = null;
        int end = 0;
        foreach (var rt in BuildRunTexts(para))
            if (ReferenceEquals(rt.Run, run)) { start ??= rt.Start; end = rt.End; }
        return start.HasValue ? (start.Value, end) : null;
    }

    /// <summary>
    /// Resolve paragraphs for a find operation based on path.
    /// "/" or "/body" → body paragraphs; "/header[N]" → header N; "/footer[N]" → footer N;
    /// "/paragraph[N]" → specific paragraph; selector → query results.
    ///
    /// BUG-TESTER+FUZZER R33: out-of-bound indices and unrecognized Word
    /// roots (e.g. /slide[1]) must throw ArgumentException instead of
    /// silently returning an empty paragraph list. Mirrors the PPTX
    /// ResolvePptParagraphsForFind contract — see commit 898f9284.
    /// CONSISTENCY(find-strict-path): Word + PPTX share this strict-path
    /// behaviour; if the contract is relaxed, update both sites in one pass.
    /// </summary>
    private List<Paragraph> ResolveParagraphsForFind(string path)
        => ResolveParagraphsForFind(path, out _);

    /// <summary>
    /// Resolve paragraphs, and — when the path is a single run (/body/p[N]/r[K])
    /// — surface that run in <paramref name="scopeRun"/> so the caller can confine
    /// find matches to the run's character span (CONSISTENCY(find-run-scope)).
    /// scopeRun is null for every non-run scope (whole paragraph, table cell,
    /// header/footer, selector, …), which fall back to the full resolved scope.
    /// </summary>
    private List<Paragraph> ResolveParagraphsForFind(string path, out Run? scopeRun)
    {
        scopeRun = null;
        var paragraphs = new List<Paragraph>();
        var mainPart = _doc.MainDocumentPart;

        if (path is "/" or "" or "/body")
        {
            // R21-1: root find/replace must sweep EVERY part that holds
            // paragraphs, not just the body — header/footer/footnote/endnote/
            // comment text was silently left unreplaced. Mirror the part
            // fan-out in EnsureAllParaIds (the canonical full-part list).
            if (mainPart?.Document?.Body != null)
                paragraphs.AddRange(mainPart.Document.Body.Descendants<Paragraph>());
            if (mainPart != null)
            {
                foreach (var headerPart in mainPart.HeaderParts)
                    if (headerPart.Header != null)
                        paragraphs.AddRange(headerPart.Header.Descendants<Paragraph>());
                foreach (var footerPart in mainPart.FooterParts)
                    if (footerPart.Footer != null)
                        paragraphs.AddRange(footerPart.Footer.Descendants<Paragraph>());
                if (mainPart.FootnotesPart?.Footnotes != null)
                    paragraphs.AddRange(mainPart.FootnotesPart.Footnotes.Descendants<Paragraph>());
                if (mainPart.EndnotesPart?.Endnotes != null)
                    paragraphs.AddRange(mainPart.EndnotesPart.Endnotes.Descendants<Paragraph>());
                // R46 Blocker-3: comment bodies are author-only annotations and must
                // NOT be swept by body-scoped find/replace. Excluding the comments part
                // here keeps "/" and "/body" scope to document text + header/footer/
                // footnote/endnote (the previously-added R21-1 fan-out). Comment text
                // is still mutable via explicit /comment[N] paths.
            }
            return paragraphs;
        }

        if (path.StartsWith("/header[", StringComparison.OrdinalIgnoreCase))
        {
            var idx = ParseHelpers.SafeParseInt(path.Split('[', ']')[1], "header index") - 1;
            var headers = mainPart?.HeaderParts.ToList() ?? new List<HeaderPart>();
            if (idx < 0 || idx >= headers.Count)
                throw new ArgumentException($"Header index out of range: {idx + 1} (have {headers.Count} header(s)).");
            var headerPart = headers[idx];
            if (headerPart.Header != null)
                paragraphs.AddRange(headerPart.Header.Descendants<Paragraph>());
            return paragraphs;
        }

        if (path.StartsWith("/footer[", StringComparison.OrdinalIgnoreCase))
        {
            var idx = ParseHelpers.SafeParseInt(path.Split('[', ']')[1], "footer index") - 1;
            var footers = mainPart?.FooterParts.ToList() ?? new List<FooterPart>();
            if (idx < 0 || idx >= footers.Count)
                throw new ArgumentException($"Footer index out of range: {idx + 1} (have {footers.Count} footer(s)).");
            var footerPart = footers[idx];
            if (footerPart.Footer != null)
                paragraphs.AddRange(footerPart.Footer.Descendants<Paragraph>());
            return paragraphs;
        }

        if (path.StartsWith("/"))
        {
            // Specific element path — navigate to it. NavigateToElement returns
            // null for both unknown roots (e.g. /slide[1]) and out-of-bound
            // indices (e.g. /body/p[999]); both must throw, never silently
            // resolve to zero paragraphs.
            var element = NavigateToElement(ParsePath(path));
            if (element == null)
                throw new ArgumentException(
                    $"Cannot resolve find scope path: '{path}'. "
                    + "Expected /, /body, /body/p[N], /body/p[N]/r[K], /body/tbl[N], "
                    + "/body/tbl[N]/tr[R]/tc[C], /header[N], or /footer[N]; "
                    + "or a CSS-style selector (e.g. paragraph, run).");

            if (element is Paragraph p)
            {
                paragraphs.Add(p);
                return paragraphs;
            }

            // BUG-BT-1: when path resolves to an inline element (e.g. a Run
            // under /body/p[N]/r[K], or a Hyperlink), Descendants<Paragraph>()
            // is empty — the find would silently match nothing. Walk up to
            // the containing paragraph instead so /run paths still work,
            // and also harvest any paragraphs nested inside (e.g. tables).
            var nestedParas = element.Descendants<Paragraph>().ToList();
            if (nestedParas.Count > 0)
            {
                paragraphs.AddRange(nestedParas);
            }
            else
            {
                var ancestorPara = element.Ancestors<Paragraph>().FirstOrDefault();
                if (ancestorPara != null)
                    paragraphs.Add(ancestorPara);
                // A /r[K] path resolves to a Run: confine find to that run's span
                // (CONSISTENCY(find-run-scope)). A Hyperlink (or other inline
                // wrapper) is not a run and keeps whole-paragraph scope.
                if (element is Run rn)
                    scopeRun = rn;
            }
            return paragraphs;
        }

        // Selector — query and resolve each result's paragraphs
        var targets = Query(path);
        foreach (var target in targets)
        {
            var elem = NavigateToElement(ParsePath(target.Path));
            if (elem is Paragraph tp)
                paragraphs.Add(tp);
            else if (elem != null)
                paragraphs.AddRange(elem.Descendants<Paragraph>());
        }

        return paragraphs;
    }

    // ==================== Add at find position ====================

    private static readonly HashSet<string> InlineTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "run", "r", "picture", "image", "img", "hyperlink", "link",
        "field", "pagenum", "pagenumber", "page", "numpages", "sectionpages", "section",
        "date", "createdate", "savedate", "printdate", "edittime", "time",
        "author", "lastsavedby", "title", "subject", "filename",
        "numwords", "numchars", "revnum", "template", "comments", "doccomments", "keywords",
        "mergefield", "ref", "pageref", "noteref", "seq", "styleref", "docproperty", "if",
        "pagebreak", "columnbreak", "break", "footnote", "endnote",
        "equation", "formula", "math", "bookmark", "formfield",
        "comment", "sdt", "contentcontrol", "chart"
    };

    /// <summary>
    /// Add an element at a text-find position within a paragraph.
    /// For inline types: split the run at the find position and insert inline.
    /// For block types: split the paragraph at the find position and insert the block element between.
    /// </summary>
    private string AddAtFindPosition(
        OpenXmlElement parent,
        string parentPath,
        string type,
        string findValue,
        bool isAfter, // true = after-find, false = before-find
        InsertPosition? position,
        Dictionary<string, string> properties)
    {
        // Support regex=true prop as alternative to r"..." prefix
        // CONSISTENCY(find-regex): mirror of WordHandler.Set.cs:60-61. grep
        // "CONSISTENCY(find-regex)" for every project-wide call site.
        if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findValue.StartsWith("r\"") && !findValue.StartsWith("r'"))
            findValue = $"r\"{findValue}\"";

        var (pattern, isRegex) = FindHelpers.ParseFindPattern(findValue);

        // Guard: empty find pattern would produce unbounded matches and blow
        // up downstream regex/plain-text scans. Surface a clean error instead
        // of leaking the raw .NET exception.
        if (string.IsNullOrEmpty(pattern))
            throw new ArgumentException("find: pattern must not be empty. Example: --after \"find:hello\".");

        // Resolve to a paragraph — either the parent itself, or the first
        // descendant paragraph of a container (body/cell/sdt) whose text
        // matches the pattern.
        Paragraph para;
        string paraPath;
        if (parent is Paragraph p)
        {
            para = p;
            paraPath = parentPath;
        }
        else
        {
            var hit = FindParagraphContainingText(parent, parentPath, pattern, isRegex)
                ?? throw new ArgumentException(
                    $"Text '{findValue}' not found in any paragraph under {parentPath}.");
            para = hit.Para;
            paraPath = hit.Path;
        }

        var runTexts = BuildRunTexts(para);
        if (runTexts.Count == 0)
            throw new ArgumentException("Paragraph has no text content to search.");

        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
        var matches = FindHelpers.FindMatchRanges(fullText, pattern, isRegex);
        if (matches.Count == 0)
            throw new ArgumentException($"Text '{findValue}' not found in paragraph.");

        // Use first match
        var (matchStart, matchLen) = matches[0];
        var splitPoint = isAfter ? matchStart + matchLen : matchStart;

        bool isInline = InlineTypes.Contains(type);

        if (isInline)
        {
            return AddInlineAtSplitPoint(para, paraPath, splitPoint, type, position, properties);
        }
        else
        {
            // Block types (paragraph/table/section/toc/…) under a `find:`
            // anchor: honor the literal position. When the anchor lands at
            // a paragraph boundary (splitPoint == 0 or == full length),
            // insert as a sibling before/after the matched paragraph
            // (no split needed). When the anchor lands mid-paragraph,
            // split the paragraph at that offset and insert the new block
            // between the two halves as body-level siblings.
            //
            // This mirrors Word's native "cursor mid-sentence → Insert →
            // Table" behavior: the user asked for position X, they get
            // the block at position X, even if that requires splitting
            // the containing paragraph.
            var container = para.Parent
                ?? throw new InvalidOperationException("Matched paragraph has no parent container.");
            var containerPath = paraPath.Contains('/')
                ? paraPath[..paraPath.LastIndexOf('/')]
                : "/body";
            var siblings = container.Elements<OpenXmlElement>().ToList();
            var paraIdx = siblings.IndexOf(para);
            if (paraIdx < 0)
                throw new InvalidOperationException("Matched paragraph not found among its parent's children.");

            var totalLen = fullText.Length;
            bool atBoundary = splitPoint == 0 || splitPoint == totalLen;

            if (atBoundary)
            {
                var insertIdx = (splitPoint == totalLen) ? paraIdx + 1 : paraIdx;
                return Add(containerPath, type, InsertPosition.AtIndex(insertIdx), properties);
            }

            // Mid-paragraph: split the paragraph, inherit pPr on the tail,
            // then insert the new block between the head and tail paragraphs.
            SplitParagraphAtOffset(para, splitPoint);
            // Head paragraph is now `para`; tail paragraph is its immediate
            // following sibling. Insert the new block between them.
            var insertIdxMid = paraIdx + 1;
            return Add(containerPath, type, InsertPosition.AtIndex(insertIdxMid), properties);
        }
    }

    /// <summary>
    /// Walk the child paragraphs of a container and return the first paragraph
    /// (plus its constructed path) whose text matches the given pattern.
    /// Used to let body-level find: anchors resolve without requiring the
    /// caller to spell out a specific paragraph path.
    /// </summary>
    private (Paragraph Para, string Path)? FindParagraphContainingText(
        OpenXmlElement container, string containerPath, string pattern, bool isRegex)
    {
        var paragraphs = container.Elements<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var candidate = paragraphs[i];
            var runTexts = BuildRunTexts(candidate);
            if (runTexts.Count == 0) continue;

            var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
            if (FindHelpers.FindMatchRanges(fullText, pattern, isRegex).Count > 0)
            {
                var paraPath = $"{containerPath}/{BuildParaPathSegment(candidate, i + 1)}";
                return (candidate, paraPath);
            }
        }
        return null;
    }

    /// <summary>
    /// Insert an inline element at a character split point within a paragraph.
    /// Splits the run at the position and inserts the element.
    /// </summary>
    private string AddInlineAtSplitPoint(
        Paragraph para,
        string parentPath,
        int splitPoint,
        string type,
        InsertPosition? position,
        Dictionary<string, string> properties)
    {
        // Split runs at the point
        var runTexts = BuildRunTexts(para);
        Run? insertAfterRun = null;

        foreach (var rt in runTexts)
        {
            if (splitPoint >= rt.Start && splitPoint <= rt.End)
            {
                if (splitPoint == rt.Start)
                {
                    // Insert before this run — find previous run
                    insertAfterRun = rt.Run.PreviousSibling<Run>();
                }
                else if (splitPoint == rt.End)
                {
                    // Insert after this run
                    insertAfterRun = rt.Run;
                }
                else
                {
                    // Split the run at the offset
                    var localOffset = splitPoint - rt.Start;
                    SplitRunAtOffset(rt.Run, rt.TextElement, localOffset);
                    insertAfterRun = rt.Run; // insert after the left portion
                }
                break;
            }
        }

        // Calculate run-based index for insertion
        var runs = para.Elements<Run>().ToList();
        int runIndex;
        if (insertAfterRun != null)
        {
            var idx = runs.IndexOf(insertAfterRun);
            runIndex = idx >= 0 ? idx + 1 : runs.Count;
        }
        else
        {
            runIndex = 0; // insert before all runs
        }

        // Convert run-count index → ChildElements-index so downstream handlers
        // (which read parent.ChildElements[index]) land at the right slot. When
        // the paragraph has a ParagraphProperties child, the ChildElements
        // index is shifted by one; when inserting before all runs, point at
        // the first run's ChildElements index rather than 0 (which is pPr).
        var childElems = para.ChildElements.ToList();
        int childIndex;
        if (runIndex >= runs.Count)
        {
            childIndex = childElems.Count;
        }
        else
        {
            var targetRun = runs[runIndex];
            childIndex = childElems.IndexOf(targetRun);
            if (childIndex < 0) childIndex = childElems.Count;
        }

        return Add(parentPath, type, InsertPosition.AtIndex(childIndex), properties);
    }

    /// <summary>
    /// Insert a block element at a character split point within a paragraph.
    /// Splits the paragraph into two and inserts the block element between them.
    /// </summary>
    private string AddBlockAtSplitPoint(
        Paragraph para,
        string parentPath,
        int splitPoint,
        string type,
        InsertPosition? position,
        Dictionary<string, string> properties)
    {
        var runTexts = BuildRunTexts(para);
        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));

        // If split point is at the very end, just insert after the paragraph
        if (splitPoint >= fullText.Length)
        {
            var bodyPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
            return Add(bodyPath, type, InsertPosition.AfterElement(parentPath.Split('/').Last()), properties);
        }

        // If split point is at the very beginning, just insert before the paragraph
        if (splitPoint <= 0)
        {
            var bodyPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
            return Add(bodyPath, type, InsertPosition.BeforeElement(parentPath.Split('/').Last()), properties);
        }

        // Split runs at the point
        foreach (var rt in runTexts)
        {
            if (splitPoint > rt.Start && splitPoint < rt.End)
            {
                var localOffset = splitPoint - rt.Start;
                SplitRunAtOffset(rt.Run, rt.TextElement, localOffset);
                break;
            }
        }

        // Rebuild run list after split
        runTexts = BuildRunTexts(para);
        fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));

        // Find the first run that starts at or after splitPoint
        Run? firstRightRun = null;
        foreach (var rt in runTexts)
        {
            if (rt.Start >= splitPoint)
            {
                firstRightRun = rt.Run;
                break;
            }
        }

        if (firstRightRun == null)
        {
            // All text before split — insert after paragraph
            var bodyPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
            return Add(bodyPath, type, InsertPosition.AfterElement(parentPath.Split('/').Last()), properties);
        }

        // Create a new paragraph for the right portion, inheriting paragraph properties
        var rightPara = new Paragraph();
        if (para.ParagraphProperties != null)
            rightPara.ParagraphProperties = (ParagraphProperties)para.ParagraphProperties.CloneNode(true);
        AssignParaId(rightPara);

        // Move runs from firstRightRun onwards to the new paragraph
        var runsToMove = new List<OpenXmlElement>();
        OpenXmlElement? current = firstRightRun;
        while (current != null)
        {
            runsToMove.Add(current);
            current = current.NextSibling();
            // Stop if we hit another paragraph-level structure (shouldn't happen normally)
        }
        // Filter: only move runs and inline elements, not ParagraphProperties
        foreach (var elem in runsToMove)
        {
            if (elem is ParagraphProperties) continue;
            elem.Remove();
            rightPara.AppendChild(elem);
        }

        // Collect existing children before Add, so we can find the newly added element
        var parentOfPara = para.Parent!;
        var childrenBefore = new HashSet<OpenXmlElement>(parentOfPara.ChildElements);

        // Insert rightPara after the original paragraph
        para.InsertAfterSelf(rightPara);

        // Add the block element via normal Add (appends before sectPr)
        var bodyParentPath = parentPath.Contains('/') ? parentPath[..parentPath.LastIndexOf('/')] : "/body";
        var result = Add(bodyParentPath, type, null, properties);

        // Find the newly added element (the one not in childrenBefore and not rightPara)
        OpenXmlElement? addedElement = null;
        foreach (var child in parentOfPara.ChildElements)
        {
            if (!childrenBefore.Contains(child) && child != rightPara)
            {
                addedElement = child;
                break;
            }
        }

        // Move it between para and rightPara
        if (addedElement != null)
        {
            addedElement.Remove();
            parentOfPara.InsertAfter(addedElement, para);
        }

        SaveDoc();
        return result;
    }
}
