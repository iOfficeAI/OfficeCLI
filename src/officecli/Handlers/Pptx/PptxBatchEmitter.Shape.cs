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
        // <a:prstGeom prst="rect"/> body — including plain text boxes. Replay
        // routes any shape carrying `geometry=` through AddShape, which (per
        // bbe1a0c8) seeds a default outline when the caller picks a geometry
        // but supplies no explicit fill/line. The result is a round-trip
        // drift: a clean textbox grows a 1pt black border on every dump+replay.
        // Strip the rect default for textbox/title sources; explicit `shape`
        // types keep the geometry so AddShape sees the user's intent.
        if ((shapeNode.Type == "textbox" || shapeNode.Type == "title")
            && shapeProps.TryGetValue("geometry", out var geomVal)
            && geomVal.Equals("rect", StringComparison.OrdinalIgnoreCase))
        {
            shapeProps.Remove("geometry");
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
        string emitType = shapeNode.Type switch
        {
            "title" => "shape",
            "equation" => isEquation ? "equation" : (shapeProps.ContainsKey("geometry") ? "shape" : "textbox"),
            _ => shapeProps.ContainsKey("geometry") ? "shape" : "textbox",
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

        EmitTextBody(ppt, fullShape, replayPath, items);
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

        EmitTextBody(ppt, full, replayPath, items);
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
    private static void EmitTextBody(PowerPointHandler ppt, DocumentNode shapeNode, string shapeParent, List<BatchItem> items)
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
            EmitParagraph(ppt, para, shapeParent, pIdx, items, firstParagraph: pIdx == 1);
        }
    }

    private static void EmitParagraph(PowerPointHandler ppt, DocumentNode paraNode, string shapeParent,
                                      int paraIdx, List<BatchItem> items, bool firstParagraph)
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
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{collapsedParaPath}/run[1]",
                    Props = runOnly,
                });
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
        bool firstRunOnSeededParagraph = firstParagraph && runs.Count > 0;
        for (int ri = 0; ri < runs.Count; ri++)
        {
            if (ri == 0 && firstRunOnSeededParagraph)
                EmitFirstRunAsSet(runs[ri], paraParent, items);
            else
                EmitRun(runs[ri], paraParent, items);
        }
    }

    // R8-7: rewrite the seeded <a:r> via `set` rather than appending another
    // one. Mirrors EmitRun but emits a single set item against
    // <paraParent>/run[1]. Empty/lang-only seeded runs in the source are
    // filtered the same way EmitRun filters; an empty rewrite is a no-op set
    // with no props.
    private static void EmitFirstRunAsSet(DocumentNode runNode, string paraParent, List<BatchItem> items)
    {
        var props = FilterEmittableProps(runNode.Format);
        DummyCtxStripSlideJump(props);
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

    private static void EmitRun(DocumentNode runNode, string paraParent, List<BatchItem> items)
    {
        var props = FilterEmittableProps(runNode.Format);
        DummyCtxStripSlideJump(props);
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
