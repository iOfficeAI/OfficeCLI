// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // PR2: emit speaker notes as a single `add notes parent=/slide[N]` row
    // carrying the concatenated body text. AddNotes accepts only `text=` /
    // `direction=` / `lang=` for the notes body — there is no typed Add
    // path for arbitrary shapes on a notesSlide, so the docx-mirror
    // "walk spTree as shape tree" approach has no replay surface to land
    // on. Emit only the body-placeholder text + direction/lang carried by
    // the `/slide[N]/notes` Get node. Mirrors AddNotes's input vocabulary.
    private static void EmitNotes(PowerPointHandler ppt, string slidePath,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        var slideMatch = System.Text.RegularExpressions.Regex.Match(slidePath, @"^/slide\[(\d+)\]$");
        if (!slideMatch.Success) return;
        var slideIdx = int.Parse(slideMatch.Groups[1].Value);
        if (!ppt.SlideHasNotes(slideIdx)) return;

        DocumentNode notes;
        try { notes = ppt.Get($"{slidePath}/notes"); }
        catch { return; }
        if (notes.Type == "error") return;

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrEmpty(notes.Text))
            props["text"] = notes.Text!;
        // Direction/lang are surfaced on the notes node Format bag by
        // AddNotes round-trip; forward them through the same canonical
        // keys AddNotes accepts.
        foreach (var key in new[] { "direction", "lang" })
        {
            if (notes.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString() ?? "";
                if (s.Length > 0) props[key] = s;
            }
        }
        if (props.Count == 0) return;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = slidePath,
            Type = "notes",
            Props = props,
        });
    }

    // Slide-level legacy comments (`<p:cm>`) live in SlideCommentsPart, not
    // the shape tree, so the standard EmitSlide walk never reaches them —
    // dump silently lost every author/date/anchor on a deck that carried
    // review comments. Re-emit each as `add comment parent=/slide[N]` using
    // the same vocabulary AddSlideComment accepts (text/author/initials/x/y/
    // date). Index-1 is emitted with no `--index`, so AddSlideComment appends
    // monotonically and the source order is preserved on replay.
    private static void EmitComments(PowerPointHandler ppt, string slidePath,
                                     List<BatchItem> items, SlideEmitContext ctx)
    {
        var slideMatch = System.Text.RegularExpressions.Regex.Match(slidePath, @"^/slide\[(\d+)\]$");
        if (!slideMatch.Success) return;
        var slideIdx = int.Parse(slideMatch.Groups[1].Value);

        List<DocumentNode> commentNodes;
        try { commentNodes = ppt.EnumerateComments(slideIdx); }
        catch { return; }
        if (commentNodes.Count == 0) return;

        foreach (var cmt in commentNodes)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(cmt.Text))
                props["text"] = cmt.Text!;
            // Mirror the AddSlideComment vocabulary verbatim. `index` is a
            // node-level Get-only field (the per-author monotonic counter
            // PowerPoint assigns); replaying it would force-collide with the
            // counter AddSlideComment maintains on the target deck.
            foreach (var key in new[] { "author", "initials", "x", "y", "date" })
            {
                if (cmt.Format.TryGetValue(key, out var v) && v != null)
                {
                    var s = v.ToString() ?? "";
                    if (s.Length > 0) props[key] = s;
                }
            }

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = slidePath,
                Type = "comment",
                Props = props.Count > 0 ? props : null,
            });
        }
    }

    // Modern p188 threaded comments — distinct OOXML part from legacy p:cm.
    // Emit one `add modernComment parent=/slide[N]` row per top-level thread
    // followed by `add modernComment parent=/slide[N]` rows with
    // parent=/slide[N]/modernComment[K] for each reply, in document order so
    // replay rebuilds the thread tree in shape.
    private static void EmitModernComments(PowerPointHandler ppt, string slidePath,
                                           List<BatchItem> items, SlideEmitContext ctx)
    {
        var slideMatch = System.Text.RegularExpressions.Regex.Match(slidePath, @"^/slide\[(\d+)\]$");
        if (!slideMatch.Success) return;
        var slideIdx = int.Parse(slideMatch.Groups[1].Value);

        List<DocumentNode> threads;
        try { threads = ppt.EnumerateModernComments(slideIdx); }
        catch { return; }
        if (threads.Count == 0) return;

        int topIdx = 0;
        foreach (var top in threads)
        {
            topIdx++;
            // Top-level row.
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(top.Text)) props["text"] = top.Text!;
            foreach (var key in new[] { "author", "initials", "created" })
            {
                if (top.Format.TryGetValue(key, out var v) && v != null)
                {
                    var s = v.ToString() ?? "";
                    if (s.Length > 0) props[key] = s;
                }
            }
            // resolved is bool — only emit when true (false is the default).
            if (top.Format.TryGetValue("resolved", out var rv) && rv is bool rb && rb)
                props["resolved"] = "true";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = slidePath,
                Type = "modernComment",
                Props = props.Count > 0 ? props : null,
            });

            // Reply rows. The top-level rows we just emitted are indexed
            // 1..N on the replayed deck in the same order we emit them, so
            // parent= can reference /slide[N]/modernComment[topIdx].
            var parentPath = $"/slide[{slideIdx}]/modernComment[{topIdx}]";
            foreach (var r in top.Children)
            {
                var rp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrEmpty(r.Text)) rp["text"] = r.Text!;
                foreach (var key in new[] { "author", "initials", "created" })
                {
                    if (r.Format.TryGetValue(key, out var v) && v != null)
                    {
                        var s = v.ToString() ?? "";
                        if (s.Length > 0) rp[key] = s;
                    }
                }
                rp["parent"] = parentPath;
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = slidePath,
                    Type = "modernComment",
                    Props = rp,
                });
            }
        }
    }
}
