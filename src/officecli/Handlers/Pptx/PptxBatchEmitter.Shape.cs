// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(emit-shape-mirror): mirrors WordBatchEmitter.Paragraph.cs
    // logic shape — get the node, filter props, decide collapsed-single-run
    // vs multi-run, emit the parent then iterate children. PowerPoint
    // shapes can carry many paragraphs (a slide text body is a list of
    // <a:p> elements), so the collapse heuristic is per-paragraph, not
    // per-shape.

    // Forward slide-jump form emitted by NodeBuilder ("slide[3]"). Internal
    // PowerPoint actions (firstslide/lastslide/nextslide/previousslide/endshow)
    // don't depend on a relationship and replay fine at shape-add time.
    private static readonly System.Text.RegularExpressions.Regex SlideJumpLink =
        new(@"^slide\[\d+\]$", System.Text.RegularExpressions.RegexOptions.IgnoreCase
                              | System.Text.RegularExpressions.RegexOptions.Compiled);

    // Strip-only variant for nested bags (paragraph/run) — the shape-level
    // emit owns the deferred slide-jump set; nested bags must not re-emit it.
    private static void DummyCtxStripSlideJump(Dictionary<string, string> props)
    {
        if (props.TryGetValue("link", out var v) && SlideJumpLink.IsMatch(v ?? ""))
            props.Remove("link");
    }

    // Run-level analogue of DeferSlideJumpLink. Target path is the run's
    // positional path under its paragraph parent; tooltip rides along.
    private static void DeferRunSlideJumpLink(Dictionary<string, string> props, string paraPath,
                                              int runIndex, SlideEmitContext ctx)
    {
        if (!props.TryGetValue("link", out var linkVal) || string.IsNullOrEmpty(linkVal)) return;
        if (!SlideJumpLink.IsMatch(linkVal)) return;
        props.Remove("link");
        var deferredProps = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
        {
            ["link"] = linkVal,
        };
        if (props.TryGetValue("tooltip", out var tt) && !string.IsNullOrEmpty(tt))
        {
            deferredProps["tooltip"] = tt;
            props.Remove("tooltip");
        }
        ctx.DeferredLinks.Add(new BatchItem
        {
            Command = "set",
            Path = $"{paraPath}/run[{runIndex}]",
            Props = deferredProps,
        });
    }

    // R24 — a:pPr accepts none of these (ECMA-376 §21.1.2.2.7 lvlLPr /
    // §21.1.2.2.6 defaultLevelParagraphProperties — language is part of
    // a:rPr only). The single-run-collapse path used to spill these onto
    // the paragraph set bag, which Set then routed into `unsupported`.
    private static readonly HashSet<string> RunOnlyRprAttrs =
        new(StringComparer.OrdinalIgnoreCase)
    {
        "lang", "altLang", "kern", "kumimoji", "normalizeH",
        "smtClean", "smtId", "bmk", "dirty", "err", "baseline",
    };

    // Pull a `link=slide[N]` prop out of the bag and queue a deferred `set`
    // BatchItem so the link write runs after every slide has been added.
    // External URLs and named actions stay in the prop bag for the normal
    // shape-add path. `enqueue=false` is used for nested para/run prop bags
    // where the shape-level emit already handles the deferred set — we just
    // need to drop the prop so the nested `set` doesn't fail.
    private static void DeferSlideJumpLink(Dictionary<string, string> props, string replayPath,
                                           SlideEmitContext ctx, bool enqueue = true)
    {
        if (!props.TryGetValue("link", out var linkVal) || string.IsNullOrEmpty(linkVal)) return;
        if (!SlideJumpLink.IsMatch(linkVal)) return;
        props.Remove("link");
        if (!enqueue) return;
        var deferredProps = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
        {
            ["link"] = linkVal,
        };
        if (props.TryGetValue("tooltip", out var tt) && !string.IsNullOrEmpty(tt))
        {
            // Tooltip is meaningful only with a link; carry it along.
            deferredProps["tooltip"] = tt;
            props.Remove("tooltip");
        }
        ctx.DeferredLinks.Add(new BatchItem
        {
            Command = "set",
            Path = replayPath,
            Props = deferredProps,
        });
    }

    private static void EmitShape(PowerPointHandler ppt, DocumentNode shapeNode, string parentSlidePath,
                                  string replayPath, List<BatchItem> items, SlideEmitContext ctx)
    {
        // depth=3 so paragraph -> run -> any inline runs all materialize. The
        // single-run collapse heuristic needs the run nodes present to read
        // their text / char-prop bag.
        var fullShape = ppt.Get(shapeNode.Path, depth: 3);
        var shapeProps = FilterEmittableProps(fullShape.Format);
        DeferSlideJumpLink(shapeProps, replayPath, ctx);

        // NodeBuilder emits `geometry=rect` for every shape with the implicit
        // <a:prstGeom prst="rect"/> body — including plain text boxes and
        // bare `--type shape` calls (no styling). Strip the rect default for
        // textbox/title (they don't "own" a geometry concept, so echoing it
        // back would attach a shape signal to a textbox on replay) and for
        // bare default-flavor shapes that carry no other distinguishing
        // styling. When the source explicitly set fill/line/etc., keep
        // `geometry=rect` so the replay path sees the same prop bag.
        if (shapeProps.TryGetValue("geometry", out var geomVal)
            && geomVal.Equals("rect", StringComparison.OrdinalIgnoreCase))
        {
            bool stripRect = shapeNode.Type == "textbox" || shapeNode.Type == "title";
            if (!stripRect && shapeNode.Type == "shape")
            {
                bool hasExplicitStyling =
                    shapeProps.ContainsKey("fill")
                    || shapeProps.ContainsKey("gradient")
                    || shapeProps.ContainsKey("pattern")
                    || shapeProps.ContainsKey("line")
                    || shapeProps.ContainsKey("lineWidth")
                    || shapeProps.ContainsKey("lineDash")
                    || shapeProps.ContainsKey("opacity");
                stripRect = !hasExplicitStyling;
            }
            if (stripRect) shapeProps.Remove("geometry");
        }

        // Emit type matches Add dispatch: "title" / "equation" both reduce to
        // "shape" or "textbox" on Add, and the emitted shape carries its
        // distinguishing prop (isTitle=true / formula=...). For now use
        // "textbox" for plain text shapes (no geometry) and "shape" otherwise.
        // CONSISTENCY(equation-emit-degrade): AddEquation throws when neither
        // `formula` nor `text` is present. NodeBuilder.ShapeToNode emits
        // Format["formula"] (LaTeX from OMath) when available; if it isn't
        // (exotic OMath that ToLatex can't render), degrade to a plain textbox
        // emit rather than crash replay.
        bool isEquation = shapeNode.Type == "equation" && shapeProps.ContainsKey("formula");
        // Preserve the shape/textbox distinction on emit: NodeBuilder's Type
        // already reflects the on-disk txBox flag, so route by Type rather
        // than reverse-engineering it from geometry presence (which we may
        // have just stripped above).
        string emitType = shapeNode.Type switch
        {
            "title" => "shape",
            "equation" => isEquation ? "equation" : "shape",
            "shape" => "shape",
            _ => "textbox",
        };

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = emitType,
            Props = shapeProps.Count > 0 ? shapeProps : null,
        });

        // Equation shapes' text body is AlternateContent (a14:m + readable
        // fallback run); the math content is fully captured by `formula`.
        // Emitting paragraphs/runs here would inject the fallback string as
        // user text — skip the body walk for equations entirely.
        if (isEquation) return;

        EmitTextBody(ppt, fullShape, replayPath, items, ctx: ctx);
    }

    private static void EmitPlaceholder(PowerPointHandler ppt, DocumentNode phNode, string parentSlidePath,
                                        string replayPath, List<BatchItem> items, SlideEmitContext ctx)
    {
        var full = ppt.Get(phNode.Path, depth: 3);
        var props = FilterEmittableProps(full.Format);
        DeferSlideJumpLink(props, replayPath, ctx);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "placeholder",
            Props = props.Count > 0 ? props : null,
        });

        // AddPlaceholder seeds the first paragraph with <a:endParaRPr> only —
        // no <a:r>. Emitting the first run via `set run[1]` (the shape/textbox
        // path) targets a non-existent run and fails the batch. Tell
        // EmitTextBody the seeded paragraph has zero runs so it issues `add
        // run` for the first run instead.
        EmitTextBody(ppt, full, replayPath, items, seededFirstParaHasRun: false, ctx: ctx);
    }

    private static void EmitConnector(PowerPointHandler ppt, DocumentNode cxnNode, string parentSlidePath,
                                      List<BatchItem> items, SlideEmitContext ctx)
    {
        var full = ppt.Get(cxnNode.Path);
        var props = FilterEmittableProps(full.Format);

        // R24 — NodeBuilder emits startShape / endShape as raw OOXML shape IDs.
        // Replay reassigns IDs through AcquireShapeId, so the original numeric
        // ID will reference the wrong shape (or be out of range) by the time
        // Add runs on a fresh deck. Translate to the positional path form that
        // ResolveShapeId already accepts (`/slide[N]/shape[K]`) so the endpoint
        // re-resolves against whatever shape sits at that ordinal in the
        // rebuilt slide. The translation is done eagerly against the source
        // slide because the source still has the original IDs.
        TranslateConnectorEndpoint(ppt, cxnNode, props, "startShape", "from");
        TranslateConnectorEndpoint(ppt, cxnNode, props, "endShape", "to");

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "connector",
            Props = props.Count > 0 ? props : null,
        });
    }

    private static void TranslateConnectorEndpoint(PowerPointHandler ppt,
        DocumentNode cxnNode, Dictionary<string, string> props,
        string srcKey, string dstKey)
    {
        if (!props.TryGetValue(srcKey, out var idStr)) return;
        if (!uint.TryParse(idStr, out var id)) return;
        // cxnNode.Path is /slide[N]/connector[K]; derive the slide number.
        var slideMatch = System.Text.RegularExpressions.Regex.Match(cxnNode.Path ?? "", @"^/slide\[(\d+)\]");
        if (!slideMatch.Success) return;
        var slideIdx = int.Parse(slideMatch.Groups[1].Value);
        var shapePathIdx = ppt.ResolveShapeOrdinalById(slideIdx, id);
        if (shapePathIdx == null) return; // Endpoint refers to a shape we
                                          // can't find on this slide (cross-
                                          // slide cxn, group-nested, etc.);
                                          // leave the raw id and let Add
                                          // emit a warning instead.
        props.Remove(srcKey);
        props[dstKey] = $"/slide[{slideIdx}]/shape[{shapePathIdx}]";
        // Drop the auxiliary index — Add re-derives the connection point.
        props.Remove(srcKey == "startShape" ? "startIdx" : "endIdx");
    }

    private static void EmitGroup(PowerPointHandler ppt, DocumentNode grpNode, string parentSlidePath,
                                  string replayPath, List<BatchItem> items, SlideEmitContext ctx)
    {
        var full = ppt.Get(grpNode.Path);
        var props = FilterEmittableProps(full.Format);
        // CONSISTENCY(zorder): direct Get on /slide[N]/group[K] strips zorder
        // because the NodeBuilder branch that emits it only runs when the
        // group surfaces as a *child* of the slide enumeration (the source
        // grpNode passed in). Without preserving zorder, a slide with
        // [group, shape] at zorders [1, 2] replays as [shape, group] = [1, 2]
        // — the group lands AFTER the shape because AddGroup defaults to
        // append. Mirror group.json (now declares add/set=true on zorder).
        if (!props.ContainsKey("zorder")
            && grpNode.Format.TryGetValue("zorder", out var grpZ) && grpZ != null)
        {
            var s = grpZ.ToString();
            if (!string.IsNullOrEmpty(s)) props["zorder"] = s!;
        }
        DeferSlideJumpLink(props, replayPath, ctx);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "group",
            Props = props.Count > 0 ? props : null,
        });

        if (full.Children == null) return;

        // Group children resolve through the same dispatch as slide-level
        // children. Replay parent for the group's children is the group's
        // positional path; children get fresh ordinals within the group scope.
        var ord = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var child in full.Children)
        {
            switch (child.Type)
            {
                case "textbox":
                case "title":
                case "shape":
                case "equation":
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    EmitShape(ppt, child, replayPath, $"{replayPath}/shape[{ord["shape"]}]", items, ctx);
                    break;
                case "connector":
                    ord["connector"] = ord.GetValueOrDefault("connector", 0) + 1;
                    EmitConnector(ppt, child, replayPath, items, ctx);
                    break;
                case "group":
                    ord["group"] = ord.GetValueOrDefault("group", 0) + 1;
                    EmitGroup(ppt, child, replayPath, $"{replayPath}/group[{ord["group"]}]", items, ctx);
                    break;
                case "placeholder":
                    // CONSISTENCY(unified-shape-counter): placeholders and
                    // plain shapes share <p:sp> sibling positions.
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    EmitPlaceholder(ppt, child, replayPath, $"{replayPath}/shape[{ord["shape"]}]", items, ctx);
                    break;
                default:
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: replayPath,
                        Reason: "group child type deferred to PR2 / unrecognized"));
                    break;
            }
        }
    }

    // Walk an emitted shape's text body. Each paragraph becomes an `add
    // paragraph` entry under the shape; runs become `add run` children of the
    // paragraph (with text carried as the canonical "text" prop). Single-run
    // paragraphs collapse run props onto the paragraph itself, mirroring the
    // docx single-run optimization.
    private static void EmitTextBody(PowerPointHandler ppt, DocumentNode shapeNode, string shapeParent, List<BatchItem> items,
                                     bool seededFirstParaHasRun = true, SlideEmitContext? ctx = null)
    {
        if (shapeNode.Children == null) return;
        var paragraphs = shapeNode.Children.Where(c => c.Type == "paragraph" || c.Type == "p").ToList();
        if (paragraphs.Count == 0) return;
        // shapeParent is the positional replay path (e.g. /slide[1]/shape[2]),
        // computed by the caller from per-slide ordinal counters. Replaces
        // the previous shapeNode.Path which carried @id= and broke replay.

        int pIdx = 0;
        foreach (var para in paragraphs)
        {
            pIdx++;
            // PPTX-SPECIFIC(shape-auto-empty-paragraph): AddShape / AddTextbox /
            // AddPlaceholder seed the txBody with one empty <a:p>. If we emit
            // every paragraph as `add`, replay produces an off-by-one empty
            // paragraph[1] that accumulates across round-trips. So the first
            // paragraph under a shape rewrites the seeded one via `set`, and
            // subsequent paragraphs append via `add`. docx body has no
            // equivalent auto-empty seed (AddSection initializes an empty body
            // and AddParagraph appends), so WordBatchEmitter uses pure `add`.
            EmitParagraph(ppt, para, shapeParent, pIdx, items,
                firstParagraph: pIdx == 1,
                seededParaHasRun: pIdx == 1 && seededFirstParaHasRun,
                ctx: ctx);
        }
    }

    private static void EmitParagraph(PowerPointHandler ppt, DocumentNode paraNode, string shapeParent,
                                      int paraIdx, List<BatchItem> items, bool firstParagraph,
                                      bool seededParaHasRun = true,
                                      SlideEmitContext? ctx = null)
    {
        var props = FilterEmittableProps(paraNode.Format);
        // CONSISTENCY(slide-jump-defer): the shape-level emit already deferred
        // the canonical `set link=slide[N]`; strip slide-jump links from any
        // bubbled-through para/run bag so the inline set doesn't fire too
        // early and trip "Slide jump target out of range".
        DummyCtxStripSlideJump(props);
        var runs = (paraNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "run" || c.Type == "r").ToList();

        // CONSISTENCY(single-run-collapse): mirrors WordBatchEmitter.Paragraph
        // collapseSingleRun — fold a lone run's text + char props onto the
        // paragraph add so simple cases stay one BatchItem.
        bool collapseSingleRun = runs.Count == 1
            && (paraNode.Children?.Count ?? 0) == 1;

        if (collapseSingleRun)
        {
            var runProps = FilterEmittableProps(runs[0].Format);
            DummyCtxStripSlideJump(runProps);
            // R24 — run-only rPr attributes (lang, altLang, kern, kumimoji,
            // normalizeH, smtClean, smtId, bmk, dirty, err, baseline) are not
            // valid on a:pPr. The collapse used to dump them onto the
            // paragraph set, which then routed them into `unsupported`. Split
            // them out and apply them via a follow-up `set …/run[1]` so the
            // round-trip still captures the rPr attribute on the right node.
            var runOnly = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var k in RunOnlyRprAttrs)
            {
                if (runProps.TryGetValue(k, out var v))
                {
                    runOnly[k] = v;
                    runProps.Remove(k);
                }
            }
            foreach (var (k, v) in runProps)
            {
                if (!props.ContainsKey(k)) props[k] = v;
            }
            if (!string.IsNullOrEmpty(runs[0].Text))
                props["text"] = runs[0].Text!;
            string collapsedParaPath;
            if (firstParagraph)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{shapeParent}/paragraph[1]",
                    Props = props.Count > 0 ? props : null,
                });
                collapsedParaPath = $"{shapeParent}/paragraph[1]";
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = shapeParent,
                    Type = "paragraph",
                    Props = props.Count > 0 ? props : null,
                });
                collapsedParaPath = $"{shapeParent}/paragraph[{paraIdx}]";
            }
            if (runOnly.Count > 0)
            {
                // CONSISTENCY(empty-paragraph-no-run): when the paragraph carried
                // no text and the run-only attrs reduce to AddRun's hard-coded
                // seeds (lang/altLang), there is no useful round-trip — emitting
                // `add run` would only re-seed the same values and grow noise
                // each cycle. Skip the row entirely (matches EmitRun /
                // EmitFirstRunAsSet's RunDefaultOnlyKeys filter).
                bool collapseHasText = props.ContainsKey("text");
                if (!collapseHasText
                    && runOnly.Keys.All(k => RunDefaultOnlyKeys.Contains(k)))
                {
                    return;
                }
                if (firstParagraph && !seededParaHasRun && !collapseHasText)
                {
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = collapsedParaPath,
                        Type = "run",
                        Props = runOnly,
                    });
                }
                else
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{collapsedParaPath}/run[1]",
                        Props = runOnly,
                    });
                }
            }
            return;
        }

        // Multi-run path: emit the paragraph empty (or with paragraph-level
        // props only) then a run per child. First paragraph rewrites the
        // shape's auto-seeded empty <a:p> via `set`; later paragraphs append.
        string paraParent;
        if (firstParagraph)
        {
            items.Add(new BatchItem
            {
                Command = "set",
                Path = $"{shapeParent}/paragraph[1]",
                Props = props.Count > 0 ? props : null,
            });
            paraParent = $"{shapeParent}/paragraph[1]";
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = shapeParent,
                Type = "paragraph",
                Props = props.Count > 0 ? props : null,
            });
            // Target parent path for runs is the just-emitted paragraph at
            // its known positional index (paraIdx). Earlier code used
            // paragraph[last()] but the resolver doesn't walk into a
            // placeholder's txBody to find paragraphs, so explicit index
            // is the portable form.
            // /body/p[last()].
            paraParent = $"{shapeParent}/paragraph[{paraIdx}]";
        }

        // R8-7: AddShape / AddTextbox / AddPlaceholder seed the txBody with
        // one paragraph carrying one empty <a:r>. If we emit every run as
        // `add`, replay produces a phantom empty run[1] before our content
        // and drifts by +1 run per round-trip. Mirror the single-paragraph
        // rewrite: the FIRST run of the FIRST paragraph rewrites the seeded
        // empty run via `set .../run[1]` rather than `add run`.
        // AddPlaceholder's seeded paragraph has zero <a:r> elements (only
        // <a:endParaRPr>), so `set run[1]` would target a missing run. Only
        // rewrite-the-seed when an actual run was seeded (shape/textbox path).
        bool firstRunOnSeededParagraph = firstParagraph && runs.Count > 0 && seededParaHasRun;
        for (int ri = 0; ri < runs.Count; ri++)
        {
            if (ri == 0 && firstRunOnSeededParagraph)
                EmitFirstRunAsSet(runs[ri], paraParent, items, ctx);
            else
                EmitRun(runs[ri], paraParent, items, ctx, runIndex: ri + 1);
        }
    }

    // R8-7: rewrite the seeded <a:r> via `set` rather than appending another
    // one. Mirrors EmitRun but emits a single set item against
    // <paraParent>/run[1]. Empty/lang-only seeded runs in the source are
    // filtered the same way EmitRun filters; an empty rewrite is a no-op set
    // with no props.
    private static void EmitFirstRunAsSet(DocumentNode runNode, string paraParent, List<BatchItem> items,
                                          SlideEmitContext? ctx = null)
    {
        var props = FilterEmittableProps(runNode.Format);
        if (ctx != null) DeferRunSlideJumpLink(props, paraParent, 1, ctx);
        else DummyCtxStripSlideJump(props);
        bool hasText = !string.IsNullOrEmpty(runNode.Text);
        if (!hasText && props.Count > 0
            && props.Keys.All(k => RunDefaultOnlyKeys.Contains(k)))
            return;
        if (!hasText && props.Count == 0) return;
        if (hasText) props["text"] = runNode.Text!;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = $"{paraParent}/run[1]",
            Props = props.Count > 0 ? props : null,
        });
    }

    // Run-level Format keys that AddRun seeds on every new <a:r> regardless
    // of caller input — emitting them adds nothing but noise on round-trip
    // AND triggers drift when the source had MORE than one default-only run.
    // `lang` is the canonical culprit: AddRun hard-codes Language="en-US",
    // so a paragraph carrying N empty <a:r> elements with only lang=en-US
    // produces N+M runs on every dump→replay (M = newly-seeded defaults on
    // the freshly-added paragraph). Treat the empty/lang-only run as a
    // no-op marker and skip it entirely.
    private static readonly HashSet<string> RunDefaultOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "lang", "altLang",
    };

    private static void EmitRun(DocumentNode runNode, string paraParent, List<BatchItem> items,
                                SlideEmitContext? ctx = null, int runIndex = 0)
    {
        var props = FilterEmittableProps(runNode.Format);
        // Defer run-level slide-jump links the same way shape-level links are
        // deferred — emit a follow-up `set` BatchItem against the run path
        // once every target slide is materialized. Without this, run-internal
        // `link=slide[N]` was silently stripped and the rendered run lost its
        // hyperlink on replay. External URLs / named actions / mailto stay in
        // the run prop bag and AddRun.ApplyRunHyperlink handles them inline.
        if (ctx != null) DeferRunSlideJumpLink(props, paraParent, runIndex, ctx);
        else DummyCtxStripSlideJump(props);
        bool hasText = !string.IsNullOrEmpty(runNode.Text);

        // Drop runs that carry no text and only default attributes AddRun
        // would seed anyway. Without this, a deck with N lang-only empty
        // runs accumulates N more on each round-trip — the source's N stay
        // (faithfully re-emitted), and AddRun's hard-coded Language="en-US"
        // seeds a fresh lang on every newly-added <a:r>, so the next dump
        // sees N+M runs per paragraph and drifts by M each cycle.
        if (!hasText && props.Count > 0
            && props.Keys.All(k => RunDefaultOnlyKeys.Contains(k)))
        {
            return;
        }
        // Fully empty <a:r> (no text, no props after filtering) — same
        // logic: AddRun would just seed its defaults, no useful content
        // round-trips.
        if (!hasText && props.Count == 0)
            return;

        if (hasText)
            props["text"] = runNode.Text!;

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraParent,
            Type = "run",
            Props = props.Count > 0 ? props : null,
        });
    }
}
