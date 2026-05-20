// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// CONSISTENCY(emit-X-mirror): scaffold mirrors WordBatchEmitter.cs — same
// public entry shape (full-doc + subtree overloads), same Get-driven
// transcription, same partial-class split (entry / Filters / Shape / Notes).
//
// PR1 scope (text-only): slide / shape / textbox / title / connector /
// group / placeholder + paragraph + run. Tables, pictures, charts, notes
// bodies, layout/master/theme raw — PR2.
public static partial class PptxBatchEmitter
{
    /// <summary>
    /// Carry-state for one emit run. Mirrors WordBatchEmitter.BodyEmitContext
    /// but trimmed for PR1 (no footnote/endnote/chart cursors yet —
    /// PowerPoint has no notes-with-numbering concept; chart/table content
    /// lands in PR2).
    /// </summary>
    internal sealed record SlideEmitContext(
        List<UnsupportedWarning> Unsupported)
    {
        // Forward slide-jump links (e.g. shape[1] on slide[1] linking to
        // slide[3]) must replay AFTER every slide is added — otherwise the
        // `link=slide[N]` prop on shape Add resolves against a deck where
        // the target slide does not yet exist and ResolveHyperlinkTarget
        // throws "Slide jump target out of range". Defer those props into
        // a second set-pass appended at the end of EmitPptx.
        public List<BatchItem> DeferredLinks { get; } = new();
    }

    /// <summary>
    /// Captured at emit time when a slide carries content we cannot round-trip
    /// through the existing handler vocabulary (animations, SmartArt, OLE,
    /// video/audio, exotic transitions). The slide itself is emitted; the
    /// unsupported element is dropped silently from `items` but recorded
    /// here so the CLI can surface a warning bundle to the caller.
    /// </summary>
    public sealed record UnsupportedWarning(string Element, string SlidePath, string Reason);

    /// <summary>
    /// Emit a full PowerPoint document as a sequence of BatchItem rows.
    /// Returns the items plus any unsupported-element warnings.
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitPptx(PowerPointHandler ppt)
    {
        var items = new List<BatchItem>();
        var ctx = new SlideEmitContext(new List<UnsupportedWarning>());

        // Clear the target deck's slides FIRST so replay onto a non-empty
        // target lands on a clean slate. Without this, `add slide` items
        // append after existing slides while every `add shape parent=/slide[N]`
        // path still resolves to the original slide[N] — the target ends up
        // with 2× the slide count (existing + freshly added empties) on each
        // round-trip. `remove /slide[*]` is a no-op on a deck with 0 slides,
        // so this is safe for the clean-target case too.
        items.Add(new BatchItem { Command = "remove", Path = "/slide[*]" });

        // Resource parts FIRST — theme, notesMaster, masters, layouts.
        // Order matters: replay's raw-set must overwrite the blank deck's
        // seeded baseline before slide content is added so per-slide
        // layout refs (sld@layout="rId4") resolve against the source's
        // layout set, not blank's. Mirrors docx's
        // settings → theme → numbering → styles → body ordering.
        EmitThemeRaw(ppt, items);
        EmitNotesMasterRaw(ppt, items);
        EmitMasterRaw(ppt, items);
        EmitLayoutRaw(ppt, items);
        // R8-5: emit presentation-level slide dimensions so custom sldSz
        // round-trips through dump → batch. Previously EmitPptx skipped the
        // root node entirely; replay always landed on the blank-deck default
        // (33.87cm × 19.05cm widescreen), silently resizing decks built for
        // 4:3, A4, custom banners, etc.
        EmitPresentationProps(ppt, items);

        // CONSISTENCY(slide-order): always iterate via the handler's
        // GetSlideParts() (sldIdLst-driven). Walking SlideParts off the
        // package returns parts in zip URI order — `slide12.xml` sorts
        // before `slide3.xml`, scrambling user-visible order.
        // CONSISTENCY(emit-skip-on-validate): a non-standard attribute or
        // element on a single slide must not abort the whole dump. The
        // OpenXml SDK throws a flat InvalidOperationException ("The element
        // does not allow the specified attribute.") when its strict-mode
        // validator catches a foreign/extension attribute (common in vendor
        // templates: gov_bja_template, 1.pptx, ...). Iterate slides one by
        // one and surface OOXML validation failures as unsupported_element
        // warnings instead of crashing the whole dump.
        var slideCount = ppt.SlideCount;
        for (int slideNum = 1; slideNum <= slideCount; slideNum++)
        {
            var slidePath = $"/slide[{slideNum}]";
            // CONSISTENCY(slide-ordinal-stub): every iteration MUST contribute
            // exactly one `add slide` so subsequent set paths /slide[N+1]/…
            // resolve to the same N+1 slot on replay. Pre-R5 we just
            // `continue`d on validation failure, emitting zero items for the
            // skipped slide — every later set drifted by one slot and
            // dump → batch on a deck with one bad slide could orphan
            // hundreds of items.
            DocumentNode slideNode;
            int preCount = items.Count;
            try { slideNode = ppt.Get(slidePath); }
            catch (Exception ex) when (ex.Message.Contains("does not allow", StringComparison.Ordinal)
                                    || ex.Message.Contains("not allowed", StringComparison.Ordinal))
            {
                ctx.Unsupported.Add(new UnsupportedWarning(
                    Element: "slide.ooxml_validation",
                    SlidePath: slidePath,
                    Reason: ex.Message));
                items.Add(new BatchItem { Command = "add", Parent = "/", Type = "slide" });
                continue;
            }
            try
            {
                EmitSlide(ppt, slideNode, slideNum, items, ctx);
            }
            catch (Exception ex) when (ex.Message.Contains("does not allow", StringComparison.Ordinal)
                                    || ex.Message.Contains("not allowed", StringComparison.Ordinal))
            {
                ctx.Unsupported.Add(new UnsupportedWarning(
                    Element: "slide.ooxml_validation",
                    SlidePath: slidePath,
                    Reason: ex.Message));
                // Roll back partial emits from the failing slide and replace
                // with a single blank-slide stub to keep ordinals aligned.
                if (items.Count > preCount)
                    items.RemoveRange(preCount, items.Count - preCount);
                items.Add(new BatchItem { Command = "add", Parent = "/", Type = "slide" });
            }
        }

        // Flush deferred slide-jump link sets — every target slide now exists,
        // so `ResolveHyperlinkTarget` can map slide[N] to the relationship.
        if (ctx.DeferredLinks.Count > 0)
            items.AddRange(ctx.DeferredLinks);

        return (items, ctx.Unsupported);
    }

    // R8-5: emit a single `set /` carrying slideWidth/slideHeight when the
    // source deck deviates from the blank-baseline 33.87cm × 19.05cm
    // widescreen. The blank-doc default is hard-coded inside BlankDocCreator,
    // not surfaced by Get, so we string-compare the canonical FormatEmu
    // output. EmitPresentationProps is a no-op for the default case to keep
    // unchanged decks from gaining a spurious item on round-trip.
    private const string DefaultSlideWidth = "33.87cm";
    private const string DefaultSlideHeight = "19.05cm";

    // Presentation-level Format keys that TrySetPresentationSetting accepts
    // on `set /`. The Get side surfaces these via PopulatePresentationSettings
    // (Set.Presentation.cs); without this allowlist, only slideWidth/Height
    // round-tripped — firstSlideNum, show.loop, print.*, compatMode, etc.
    // were silently dropped on dump.
    //
    // Get emits `direction = rtl` for RTL presentations but the setter case
    // key is `rtl`. We rewrite the key on emit so replay's TrySetPresentationSetting
    // accepts it. Mirrors the `direction → rtl` alias that already lives in
    // Set.cs path-pattern dispatch.
    private static readonly HashSet<string> PresentationEmitKeys =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "firstSlideNum", "compatMode", "removePersonalInfo",
            "print.what", "print.colorMode", "print.hiddenSlides",
            "print.scaleToFitPaper", "print.frameSlides",
            "show.loop", "show.narration", "show.animation", "show.useTimings",
        };

    private static void EmitPresentationProps(PowerPointHandler ppt, List<BatchItem> items)
    {
        DocumentNode root;
        try { root = ppt.Get("/"); }
        catch { return; }
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.Format.TryGetValue("slideWidth", out var wObj) && wObj is string w
            && !string.Equals(w, DefaultSlideWidth, StringComparison.OrdinalIgnoreCase))
            props["slideWidth"] = w;
        if (root.Format.TryGetValue("slideHeight", out var hObj) && hObj is string h
            && !string.Equals(h, DefaultSlideHeight, StringComparison.OrdinalIgnoreCase))
            props["slideHeight"] = h;

        // Presentation attributes / print / show settings — only emit non-default
        // values (Get omits keys that match the OOXML defaults).
        foreach (var key in PresentationEmitKeys)
        {
            if (!root.Format.TryGetValue(key, out var v) || v == null) continue;
            var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s.Length == 0) continue;
            props[key] = s;
        }

        // direction → rtl: Get emits `direction = rtl`, setter accepts `rtl`.
        if (root.Format.TryGetValue("direction", out var dObj) && dObj is string ds
            && ds.Equals("rtl", StringComparison.OrdinalIgnoreCase))
            props["rtl"] = "true";

        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props,
        });
    }

    /// <summary>
    /// Emit a subtree of a PowerPoint document. Supported subtree paths:
    /// `/slide[N]`, `/theme`, `/notesMaster`, `/slideMaster[N]`, `/slideLayout[N]`,
    /// `/noteSlide[N]`, `/presentation`. Resource subtrees emit a single raw-set
    /// replace; replay onto a foreign deck does NOT carry cross-part dependency
    /// closure (e.g. a `/slideLayout[K]` dump only stamps the layout's XML — the
    /// referenced master, theme, and per-slide layout rId rewiring are NOT
    /// included). Mirrors WordBatchEmitter's raw-emit subtree surface
    /// (/theme, /settings, /numbering, /styles).
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitPptx(
        PowerPointHandler ppt, string path)
    {
        const string SupportedHint = "Supported: /, /presentation, /slide[N], /theme, /notesMaster, /slideMaster[N], /slideLayout[N], /noteSlide[N]";

        if (string.IsNullOrEmpty(path))
            throw new CliException($"dump path cannot be empty. Use '/' for the full document or a subtree path like /slide[N]. {SupportedHint}")
                { Code = "invalid_path" };
        if (path == "/") return EmitPptx(ppt);

        var items = new List<BatchItem>();
        var ctx = new SlideEmitContext(new List<UnsupportedWarning>());

        // CONSISTENCY(case-insensitive-subtree): paths with [N] go through
        // regex `IgnoreCase`, so case-folding the literal-prefix branches too
        // aligns the dispatcher to a single rule and matches the docx subtree
        // dispatcher (WordBatchEmitter uses `path.ToLowerInvariant()`).
        var lp = path.ToLowerInvariant();

        if (lp == "/presentation")
        {
            EmitPresentationProps(ppt, items);
            return (items, ctx.Unsupported);
        }
        if (lp == "/theme")
        {
            EmitThemeRaw(ppt, items);
            return (items, ctx.Unsupported);
        }
        if (lp == "/notesmaster")
        {
            EmitNotesMasterRaw(ppt, items);
            return (items, ctx.Unsupported);
        }

        // Index parsing: regex restricts to ASCII [0-9]+ (not Unicode \d, which
        // matches Arabic-Indic numerals etc. that int.Parse rejects under
        // InvariantCulture). int.TryParse guards against Int32 overflow.
        static int ParseIndexOrThrow(string raw, string fullPath)
        {
            if (!int.TryParse(raw, System.Globalization.NumberStyles.Integer,
                              System.Globalization.CultureInfo.InvariantCulture, out var n))
                throw new CliException($"dump path not found: {fullPath} (index '{raw}' out of range or not an integer)")
                    { Code = "path_not_found" };
            return n;
        }

        var slideMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slide\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (slideMatch.Success)
        {
            var idx = ParseIndexOrThrow(slideMatch.Groups[1].Value, path);
            DocumentNode slideNode;
            try { slideNode = ppt.Get(path); }
            catch (Exception ex)
            {
                throw new CliException($"dump path not found: {path} ({ex.Message})") { Code = "path_not_found" };
            }
            EmitSlide(ppt, slideNode, idx, items, ctx);
            // CONSISTENCY(deferred-link-flush): subtree slide dump must flush
            // ctx.DeferredLinks before returning, otherwise any `link=slide[N]`
            // prop on a shape inside the dumped slide is silently dropped from
            // the output (DeferSlideJumpLink moves it out of the shape's
            // prop bag into ctx.DeferredLinks expecting the full-doc EmitPptx
            // tail flush, which the subtree path never reaches).
            //
            // Cross-slide targets (e.g. dump /slide[1] when the shape links
            // to /slide[3]) still emit the set row — replay against a deck
            // missing the target slide fails with ResolveHyperlinkTarget's
            // "Slide jump target out of range", which is a clearer error
            // than silent prop loss.
            if (ctx.DeferredLinks.Count > 0)
                items.AddRange(ctx.DeferredLinks);
            return (items, ctx.Unsupported);
        }

        var masterMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slideMaster\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (masterMatch.Success)
        {
            var idx = ParseIndexOrThrow(masterMatch.Groups[1].Value, path);
            if (idx < 1 || idx > ppt.SlideMasterCount)
                throw new CliException($"dump path not found: {path} (total slideMasters: {ppt.SlideMasterCount})")
                    { Code = "path_not_found" };
            if (!EmitMasterRawOne(ppt, idx, items))
                throw new CliException($"dump path not found: {path} (slideMaster {idx} raw read failed)")
                    { Code = "path_not_found" };
            return (items, ctx.Unsupported);
        }

        var layoutMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slideLayout\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (layoutMatch.Success)
        {
            var idx = ParseIndexOrThrow(layoutMatch.Groups[1].Value, path);
            if (idx < 1 || idx > ppt.SlideLayoutCount)
                throw new CliException($"dump path not found: {path} (total slideLayouts: {ppt.SlideLayoutCount})")
                    { Code = "path_not_found" };
            if (!EmitLayoutRawOne(ppt, idx, items))
                throw new CliException($"dump path not found: {path} (slideLayout {idx} raw read failed)")
                    { Code = "path_not_found" };
            return (items, ctx.Unsupported);
        }

        var noteMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/noteSlide\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (noteMatch.Success)
        {
            var idx = ParseIndexOrThrow(noteMatch.Groups[1].Value, path);
            if (idx < 1 || idx > ppt.SlideCount)
                throw new CliException($"dump path not found: {path} (total slides: {ppt.SlideCount})")
                    { Code = "path_not_found" };
            if (!EmitNoteSlideRawOne(ppt, idx, items))
                throw new CliException($"dump path not found: {path} (slide {idx} has no notes)")
                    { Code = "path_not_found" };
            return (items, ctx.Unsupported);
        }

        throw new CliException(
            $"dump path not supported: {path}. {SupportedHint}")
            { Code = "unsupported_path" };
    }

    private static void EmitSlide(PowerPointHandler ppt, DocumentNode slideNode, int slideNum,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        var slidePath = slideNode.Path;
        ProbeUnsupportedOnSlide(ppt, slidePath, ctx);

        // Pull the full slide node so layout / hidden / background etc. surface
        // even when the entry passed us a depth-truncated tree from "/".
        var fullSlide = ppt.Get(slidePath);
        var slideProps = FilterEmittableProps(fullSlide.Format);

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/",
            Type = "slide",
            Props = slideProps.Count > 0 ? slideProps : null,
        });

        // ShapeToNode tags placeholder shapes as plain "textbox"/"title". To
        // emit them as `add placeholder` we cross-reference each shape's cNvPr
        // id with the slide's Query("placeholder") result.
        // Only index placeholders defined on the slide itself. Query also
        // returns layout-inherited placeholders (Format["inheritedFrom"]
        // = "layout") whose ph index/id can collide with auto-assigned
        // textbox cNvPr ids on the slide (python-pptx starts at 2, layout
        // ftr/dt/sldNum live at id 2..4) — without this filter, the second
        // textbox would be misclassified as `ftr` and crash placeholder
        // type parsing, or silently disappear in dump.
        var placeholderById = new Dictionary<string, DocumentNode>(StringComparer.Ordinal);
        foreach (var ph in ppt.Query("placeholder"))
        {
            if (!ph.Path.StartsWith(slidePath + "/", StringComparison.Ordinal)) continue;
            if (ph.Format.TryGetValue("inheritedFrom", out var inh) && inh as string == "layout") continue;
            if (ph.Format.TryGetValue("id", out var phId) && phId != null)
                placeholderById[phId.ToString()!] = ph;
        }

        // Children: walk shape-tree level. Get already routed group/connector/
        // textbox/title/equation into typed nodes, so just iterate and dispatch.
        if (fullSlide.Children == null) return;
        // CONSISTENCY(positional-emit): dump references its own added elements
        // by positional `/slide[N]/shape[K]` (mirrors docx /body/p[K]) rather
        // than cNvPr `@id=N`. Add accepts caller-supplied id but emit chooses
        // not to use it — id collisions with layout-inherited placeholders
        // would otherwise break replay (animations/video deck cascade).
        //
        // CONSISTENCY(unified-shape-counter): placeholders are <p:sp> siblings
        // of plain shapes in the OOXML shape tree, so ResolveShape counts them
        // together. AddPlaceholder also appends a <p:sp> and returns
        // `/slide[N]/shape[<count>]` (Add.Misc.cs). The emitter must therefore
        // share a SINGLE positional counter across textbox/title/shape/equation
        // /placeholder and emit replay paths as `/slide[N]/shape[K]` for ALL
        // of them — otherwise a placeholder dispatched first leaves the
        // shape counter at 1, and the next textbox emits `set
        // /slide[N]/shape[1]/...` which on replay clobbers the placeholder.
        // Previously the emitter kept separate `shape`/`placeholder` counters
        // and emitted `/slide[N]/placeholder[K]` for placeholders, but the
        // replay paths for paragraph/run inside that placeholder still used
        // the same `/slide[N]/shape[K]` form — see EmitTextBody — so every
        // shape after a placeholder collided.
        // Pre-build the per-slide animation index keyed by source shape @id
        // (or positional fallback). EmitAnimationsForShape pulls per-shape
        // entries from this map as we emit each <p:sp>.
        var animIndex = BuildSlideAnimationIndex(ppt, slideNum);

        var ord = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var child in fullSlide.Children)
        {
            // Placeholder dispatch first — overrides textbox/title type.
            if ((child.Type == "textbox" || child.Type == "title" || child.Type == "shape")
                && child.Format.TryGetValue("id", out var cid) && cid != null
                && placeholderById.TryGetValue(cid.ToString()!, out var phNode))
            {
                ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                var phReplay = $"{slidePath}/shape[{ord["shape"]}]";
                EmitPlaceholder(ppt, phNode, slidePath, phReplay, items, ctx);
                EmitAnimationsForShape(GetAnimationsForChild(animIndex, child, ord["shape"]), phReplay, items);
                continue;
            }
            switch (child.Type)
            {
                case "textbox":
                case "title":
                case "shape":
                case "equation":
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    {
                        var shReplay = $"{slidePath}/shape[{ord["shape"]}]";
                        EmitShape(ppt, child, slidePath, shReplay, items, ctx);
                        EmitAnimationsForShape(GetAnimationsForChild(animIndex, child, ord["shape"]), shReplay, items);
                    }
                    break;
                case "placeholder":
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    {
                        var phReplay2 = $"{slidePath}/shape[{ord["shape"]}]";
                        EmitPlaceholder(ppt, child, slidePath, phReplay2, items, ctx);
                        EmitAnimationsForShape(GetAnimationsForChild(animIndex, child, ord["shape"]), phReplay2, items);
                    }
                    break;
                case "connector":
                    ord["connector"] = ord.GetValueOrDefault("connector", 0) + 1;
                    EmitConnector(ppt, child, slidePath, items, ctx);
                    break;
                case "group":
                    ord["group"] = ord.GetValueOrDefault("group", 0) + 1;
                    EmitGroup(ppt, child, slidePath, $"{slidePath}/group[{ord["group"]}]", items, ctx);
                    break;
                case "table":
                    ord["table"] = ord.GetValueOrDefault("table", 0) + 1;
                    EmitTable(ppt, child, slidePath, $"{slidePath}/table[{ord["table"]}]", items, ctx);
                    break;
                case "picture":
                    ord["picture"] = ord.GetValueOrDefault("picture", 0) + 1;
                    EmitPicture(ppt, child, slidePath, $"{slidePath}/picture[{ord["picture"]}]", items, ctx);
                    break;
                case "chart":
                    ord["chart"] = ord.GetValueOrDefault("chart", 0) + 1;
                    EmitChart(ppt, child, slidePath, items, ctx, ord["chart"]);
                    break;
                case "ole":
                case "video":
                case "audio":
                case "3dmodel":
                case "model3d":
                case "zoom":
                    // PR3+ scope. ProbeUnsupportedOnSlide already records the
                    // OLE/video/audio/3D markers via raw-XML sniff; this branch
                    // catches the children that surfaced via the typed Get
                    // tree (when NodeBuilder learns to tag them).
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: slidePath,
                        Reason: "deferred to later PR"));
                    break;
                default:
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: slidePath,
                        Reason: "unrecognized child type"));
                    break;
            }
        }

        // Notes body content — stub for PR1. Notes part presence does not
        // surface in the slide subtree's children today (notes live under
        // /slide[N]/notes); PR2 will reach in and emit them.
        EmitNotes(ppt, slidePath, items, ctx);

        // Legacy slide comments — also off the shape tree (SlideCommentsPart).
        // Emit AFTER notes so the per-slide row order is stable: shapes →
        // notes → comments, mirroring how a reader would traverse the slide.
        EmitComments(ppt, slidePath, items, ctx);

        // Modern p188 threaded comments — distinct from legacy p:cm; live in
        // PowerPointCommentPart. Emit after legacy comments to keep a stable
        // per-slide row ordering.
        EmitModernComments(ppt, slidePath, items, ctx);
    }

    // Touch the raw slide XML to find content that has no handler vocabulary
    // yet. Each match adds an UnsupportedWarning entry; we never throw.
    private static void ProbeUnsupportedOnSlide(PowerPointHandler ppt, string slidePath,
                                                SlideEmitContext ctx)
    {
        string xml;
        try { xml = ppt.Raw(slidePath); }
        catch { return; }

        // <p:timing> = slide animation. EmitAnimationsForShape now emits the
        // entrance/exit/emphasis effects per shape via the `animation` Query
        // surface, so the timing tree no longer aborts to an unsupported
        // warning. Exotic timing constructs (motion paths, sequence groupings)
        // still go through the Query — animations the Query doesn't enumerate
        // are silently dropped.

        // SmartArt sits inside a graphicFrame as a dgm:relIds element.
        if (xml.Contains("dgm:relIds", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("smartArt", slidePath,
                "diagram (SmartArt) graphic frame present"));

        // OLE / video / audio / 3D — element names are distinctive enough.
        if (xml.Contains("<p:oleObj", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("oleObj", slidePath,
                "embedded OLE object present"));
        if (xml.Contains("<p:video", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("video", slidePath, "video element present"));
        if (xml.Contains("<p:audio", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("audio", slidePath, "audio element present"));
        if (xml.Contains("p:model3d", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("model3D", slidePath, "3D model present"));

        // Exotic transitions. Morph is most common; conveyor/ferris/honeycomb/
        // gallery live under p:transition's p15: extension list. Sniff the
        // transition element if present and tag by extension hint.
        // Vanilla transitions (fade/push/wipe/cut) already round-trip via
        // the `transition` prop, so they are NOT unsupported.
        var tIdx = xml.IndexOf("<p:transition", StringComparison.Ordinal);
        if (tIdx >= 0)
        {
            var tEnd = xml.IndexOf("</p:transition>", tIdx, StringComparison.Ordinal);
            var tSlice = tEnd > tIdx ? xml.Substring(tIdx, tEnd - tIdx) : xml.Substring(tIdx);
            if (tSlice.Contains("p159:morph", StringComparison.Ordinal)
                || tSlice.Contains("p15:morph", StringComparison.Ordinal)
                || tSlice.Contains("<p159:morph", StringComparison.Ordinal))
            {
                ctx.Unsupported.Add(new UnsupportedWarning("transition.morph", slidePath,
                    "morph transition uses p15: extension"));
            }
        }
    }

    // Emit one `add animation` BatchItem per effect attached to this shape.
    // Replay parent is the shape's positional path in the emitted document
    // (caller-supplied — must match the just-emitted `add shape/placeholder`).
    //
    // Previously animations were caught by ProbeUnsupportedOnSlide and surfaced
    // only as a warning, so dump→batch→replay lost every entrance/exit/emphasis
    // effect plus its trigger/delay/duration. The animation Query surface
    // already produces fine-grained nodes (effect/class/trigger/duration/delay/
    // direction/easein/easeout via PopulateAnimationNode); this helper just
    // forwards each animation's emittable props as an `add animation` row.
    //
    // Direction was added to PopulateAnimationNode in this same change — without
    // it, fly-down would round-trip as fly-up (AddAnimation default).
    //
    // Motion-path animations are excluded by the Query (presetClass="motion"
    // never surfaces under selector "animation"). Other exotic timing
    // constructs (sequence groupings, conditional triggers) are silently
    // dropped — the visible effects round-trip.
    // Per-shape animation emit. Accepts a pre-filtered list of animation
    // nodes whose shape segment matches this shape (resolved by the caller
    // via the @id → positional map built from fullSlide.Children).
    private static void EmitAnimationsForShape(List<DocumentNode> animsForShape,
                                               string replayShapePath, List<BatchItem> items)
    {
        foreach (var anim in animsForShape)
        {
            var animProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // Map Format keys → AddAnimation accepted keys. presetId is
            // derived from effect+class on Add, so emitting it would either
            // be ignored or trigger an unsupported_property warning.
            foreach (var (k, v) in anim.Format)
            {
                if (v == null) continue;
                if (k.Equals("presetId", StringComparison.OrdinalIgnoreCase)) continue;
                var s = v.ToString() ?? "";
                if (s.Length == 0) continue;
                animProps[k] = s;
            }
            if (animProps.Count == 0) continue;

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = replayShapePath,
                Type = "animation",
                Props = animProps,
            });
        }
    }

    // Build a map from source @id (or source positional) to the list of
    // animation nodes on that shape. Query("animation") paths use either
    // /slide[N]/shape[@id=X]/animation[A] or /slide[N]/shape[K]/animation[A]
    // depending on whether cNvPr.Id is present.
    private static Dictionary<string, List<DocumentNode>> BuildSlideAnimationIndex(
        PowerPointHandler ppt, int slideNum)
    {
        var map = new Dictionary<string, List<DocumentNode>>(StringComparer.Ordinal);
        List<DocumentNode> all;
        try { all = ppt.Query("animation"); }
        catch { return map; }

        var slidePrefix = $"/slide[{slideNum}]/";
        var rx = new System.Text.RegularExpressions.Regex(
            @"^/slide\[\d+\]/shape\[([^\]]+)\]/animation\[\d+\]$");
        foreach (var anim in all)
        {
            if (!anim.Path.StartsWith(slidePrefix, StringComparison.Ordinal)) continue;
            var m = rx.Match(anim.Path);
            if (!m.Success) continue;
            var key = m.Groups[1].Value; // either "5" (positional) or "@id=10" form
            if (!map.TryGetValue(key, out var list))
            {
                list = new List<DocumentNode>();
                map[key] = list;
            }
            list.Add(anim);
        }
        return map;
    }

    // Resolve the animation list for the shape currently being emitted.
    // child.Format["id"] (when present) maps to @id=X; otherwise positional.
    private static List<DocumentNode> GetAnimationsForChild(
        Dictionary<string, List<DocumentNode>> map, DocumentNode child, int sourcePositional)
    {
        // Try @id= form first when child carries id.
        if (child.Format.TryGetValue("id", out var cidObj) && cidObj != null)
        {
            var idKey = $"@id={cidObj}";
            if (map.TryGetValue(idKey, out var byId)) return byId;
        }
        // Fall back to positional.
        if (map.TryGetValue(sourcePositional.ToString(System.Globalization.CultureInfo.InvariantCulture),
                            out var byPos))
            return byPos;
        return new List<DocumentNode>();
    }
}
