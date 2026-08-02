// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(picture-inline-base64): mirrors
    // WordBatchEmitter.Paragraph.TryEmitPictureRun — no size threshold, no
    // sidecar file, always emit `src="data:<contentType>;base64,<bytes>"`.
    // A 50MB picture produces a 70MB batch JSON; accepted by design.
    private static void EmitPicture(PowerPointHandler ppt, DocumentNode picNode,
                                    string parentSlidePath, string replayPath,
                                    List<BatchItem> items,
                                    SlideEmitContext ctx)
    {
        var fullPic = ppt.Get(picNode.Path);
        var props = FilterEmittableProps(fullPic.Format);
        DeferSlideJumpLink(props, replayPath, ctx);

        // <a:clrChange> recolor adjustment — there is no typed Set
        // vocabulary today, so capture the source element's outer XML and
        // re-inject it on replay via raw-set after the picture has been
        // added. Pulled here so the raw-set lands on the correct
        // picture[K]/blipFill/blip xpath.
        string? clrChangeXml = ppt.GetPictureBlipClrChangeXml(picNode.Path);

        // R56 bt-6: capture every other untyped blip child (alphaMod / blur /
        // hsl / tint / <a:colorMod> channel modulation / ...) so dump→batch
        // doesn't silently drop them. The typed children (alphaModFix /
        // biLevel / duotone / lum / clrChange) ride the Format-key surface
        // already and are filtered out by GetPictureBlipPassthroughChildrenXml.
        var passthroughBlipChildren = ppt.GetPictureBlipPassthroughChildrenXml(picNode.Path);

        // Companion binary parts the blip references from its extLst (HD Photo
        // .wdp backup layer, SVG companion). The passthrough above re-appends the
        // extLst verbatim with <... r:embed="rIdN">, so each companion part must
        // be re-created with the SAME source rId or the rebuilt picture dangles
        // (lost effects layer; strict consumers reject the deck). Captured here
        // and emitted as add-part extpart rows below, after the picture add.
        var blipCompanions = ppt.GetPictureBlipCompanionParts(picNode.Path);

        // Picture-in-placeholder: a picture that fills a layout placeholder
        // carries <p:ph type="pic" idx="N"/> in its nvPr and an empty <p:spPr/>,
        // inheriting its geometry from the layout. The plain `add picture` below
        // drops the <p:ph> and AddPicture stamps a default xfrm, so the picture
        // lands at the wrong size/offset on replay. Capture the placeholder
        // marker (and the source spPr when it had no explicit xfrm) so we can
        // re-inject them via raw-set and let the layout drive geometry again.
        var (phXml, inheritSpPrXml) = ppt.GetPicturePlaceholderRoundtripXml(picNode.Path);

        var binary = ppt.GetImageBinary(picNode.Path);
        if (binary.HasValue)
        {
            var (bytes, contentType) = binary.Value;
            props["src"] = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
        }
        else
        {
            // No embedded part — picture is unresolvable on round-trip.
            // Drop to an unsupported warning rather than emit a half-row
            // that AddPicture would reject for missing src.
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "picture",
                SlidePath: parentSlidePath,
                Reason: "picture has no resolvable embedded image part"));
            return;
        }

        // Drop Get-only diagnostic keys that AddPicture neither expects nor
        // accepts (mirrors docx WordBatchEmitter picture emit).
        // CONSISTENCY(shape-id-high-range): KEEP the source cNvPr.Id —
        // AcquireShapeId in AddPicture honors caller-supplied id and the
        // auto-assign base is 100000+, so the source id (typically a single-
        // or low-double-digit number for PowerPoint-authored pictures)
        // never collides with the counter. Without preserving it, the
        // picture's cNvPr id rewrites to 100000+ on round-trip, drifting
        // against the source and breaking any animation/spTgt that targeted
        // the picture by id. Mirrors EmitPlaceholder / EmitShape behavior.
        props.Remove("contentType");
        props.Remove("fileSize");
        props.Remove("alt");
        // Re-add alt only if it was the explicit user-set value (not the
        // "(missing)" placeholder PictureToNode stamps in).
        var altRaw = fullPic.Format.TryGetValue("alt", out var av) ? av?.ToString() : null;
        if (!string.IsNullOrEmpty(altRaw) && altRaw != "(missing)")
            props["alt"] = altRaw;

        // Schema declares brightness/contrast/shadow/glow as add:false, set:true
        // on pptx/picture. AddPicture rejects them with UNSUPPORTED on replay
        // and the values are silently lost. Lift them out of the add bag and
        // defer to a follow-up `set` on the same replay path. Mirrors the
        // DeferSlideJumpLink pattern (deferred-set after every add).
        // Hard-coded picture-level drop list — same precedent as `image=true`,
        // `background=image`, `fill=gradient` drops elsewhere in this emitter.
        var deferredEffects = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in PictureSetOnlyEffectKeys)
        {
            if (props.TryGetValue(key, out var val))
            {
                deferredEffects[key] = val;
                props.Remove(key);
            }
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "picture",
            Props = props.Count > 0 ? props : null,
        });

        if (deferredEffects.Count > 0)
        {
            ctx.DeferredLinks.Add(new BatchItem
            {
                Command = "set",
                Path = replayPath,
                Props = deferredEffects,
            });
        }

        // Carry the blip's companion parts (HD Photo .wdp layer, SVG companion)
        // BEFORE the extLst passthrough raw-set below re-introduces their
        // r:embed references. The companion relationship is slide-level, so the
        // host is /slide[N] (extracted from replayPath, which may be nested in a
        // group). Pin the SOURCE rId + relationship type via add-part extpart.
        if (blipCompanions.Count > 0
            && System.Text.RegularExpressions.Regex.Match(replayPath, @"^/slide\[(\d+)\]")
                is { Success: true } slideM)
        {
            var slideHostPath = $"/slide[{slideM.Groups[1].Value}]";
            foreach (var comp in blipCompanions)
            {
                items.Add(new BatchItem
                {
                    Command = "add-part",
                    Parent = slideHostPath,
                    Type = "extpart",
                    Props = new Dictionary<string, string>
                    {
                        ["rid"] = comp.RelId,
                        ["rel-type"] = comp.RelType,
                        ["content-type"] = comp.ContentType,
                        ["ext"] = comp.TargetExt,
                        ["data"] = comp.Base64Data,
                    },
                });
            }
        }

        // CONSISTENCY(picture-clrchange-rawset): inject <a:clrChange> back
        // onto the just-added picture's <a:blip>. Schema position: the
        // clrChange child sits between the optional <a:alphaBiLevel> /
        // <a:alphaCeiling> / etc. siblings and before any fill-mode
        // elements; appending to the blip element keeps the same relative
        // ordering as the source for the common one-effect case. The
        // replayPath form is `/slide[N]/picture[K]` where K is the
        // 1-based picture ordinal within spTree's <p:pic> siblings — the
        // same scope `p:pic[K]` resolves through the xpath engine.
        if ((clrChangeXml != null || passthroughBlipChildren.Count > 0)
            && System.Text.RegularExpressions.Regex.Match(replayPath,
                @"^/slide\[(\d+)\]/picture\[(\d+)\]$") is { Success: true } picM)
        {
            var picOrd = int.Parse(picM.Groups[2].Value);
            var blipXpath = $"/p:sld/p:cSld/p:spTree/p:pic[{picOrd}]/p:blipFill/a:blip";
            if (clrChangeXml != null)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = parentSlidePath,
                    Xpath = blipXpath,
                    Action = "append",
                    Xml = clrChangeXml,
                });
            }
            // R56 bt-6: each untyped blip child is replayed as its own append.
            // Source order is preserved (foreach over ChildElements upstream).
            foreach (var childXml in passthroughBlipChildren)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = parentSlidePath,
                    Xpath = blipXpath,
                    Action = "append",
                    Xml = childXml,
                });
            }
        }

        // Re-inject the placeholder marker + inherited spPr captured above.
        // Appending <p:ph> into the rebuilt picture's empty <p:nvPr> restores
        // the placeholder binding; replacing the rebuilt <p:spPr> with the
        // source's xfrm-less spPr drops AddPicture's default xfrm so the layout
        // placeholder geometry is inherited again.
        if (phXml != null
            && System.Text.RegularExpressions.Regex.Match(replayPath,
                @"^/slide\[(\d+)\]/picture\[(\d+)\]$") is { Success: true } phPicM)
        {
            var phPicOrd = int.Parse(phPicM.Groups[2].Value);
            var picXpath = $"/p:sld/p:cSld/p:spTree/p:pic[{phPicOrd}]";
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = parentSlidePath,
                Xpath = $"{picXpath}/p:nvPicPr/p:nvPr",
                Action = "append",
                Xml = phXml,
            });
            if (inheritSpPrXml != null)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = parentSlidePath,
                    Xpath = $"{picXpath}/p:spPr",
                    Action = "replace",
                    Xml = inheritSpPrXml,
                });
            }
        }
    }

    // Picture effect props with schema `add: false, set: true`. Must NOT ride
    // along inside the add picture op props bag — AddPicture rejects them.
    private static readonly HashSet<string> PictureSetOnlyEffectKeys =
        new(StringComparer.OrdinalIgnoreCase)
    {
        "brightness", "contrast", "shadow", "glow",
    };

    // Phase 3c-media. Mirrors EmitSmartArtsForSlide (Phase 3b). Per slide,
    // scan for <p:pic> hosts that carry <a:videoFile> or <a:audioFile>;
    // emit an `add-part video|audio` row that creates the underlying
    // MediaDataPart + Video/AudioReferenceRelationship + MediaReferenceRel
    // + thumbnail ImagePart with SOURCE rIds pinned via --prop. Then emit
    // one raw-set append on /p:sld/p:cSld/p:spTree carrying the <p:pic>
    // XML verbatim — the pinned rIds make the videoFile/audioFile/p14:media/
    // blip references all resolve to the just-created parts.
    //
    // Skipped by the typed walk: the dispatch in EmitSlide's switch routes
    // child.Type == "video"|"audio" away from EmitPicture (which would
    // re-emit a plain picture without the media rels) into a no-op,
    // letting THIS pass own the entire <p:pic> emit.
    //
    // Audit caveat: the SDK's CreateMediaDataPart allocates a URI like
    // /ppt/media/media1.mp4, NOT the source's /media/mediadata.mp4
    // (legacy zip-root layout). The binary content survives byte-equal;
    // the audit's content-loss check is by content hash (see
    // tools/pptx-roundtrip-audit.py).
    internal static void EmitMediaForSlide(PowerPointHandler ppt, int slideNum,
                                           string slidePath, List<BatchItem> items,
                                           SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.MediaInfo> medias;
        try { medias = ppt.GetMediaOnSlide(slideNum); }
        catch { return; }
        if (medias.Count == 0) return;

        foreach (var m in medias)
        {
            var partType = m.IsVideo ? "video" : "audio";
            var ridKey   = m.IsVideo ? "video-rid" : "audio-rid";
            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["data"] = Convert.ToBase64String(m.MediaBytes),
                ["content-type"] = m.MediaContentType,
                ["extension"] = m.MediaExtension,
                ["thumbnail-data"] = Convert.ToBase64String(m.ThumbnailBytes),
                ["thumbnail-content-type"] = m.ThumbnailContentType,
                [ridKey] = m.LinkRelId,
                ["media-rid"] = m.MediaEmbedRelId,
                ["thumbnail-rid"] = m.ThumbnailRelId,
            };
            items.Add(new BatchItem
            {
                Command = "add-part",
                Parent = slidePath,
                Type = partType,
                Props = props,
            });

            // Append the <p:pic> verbatim into the spTree. Canonicalise
            // via the slide-slice canonicaliser so post-replay re-emit
            // hits byte-equal (same trick SmartArt uses).
            string picCanon;
            try { picCanon = NormalizeSlideRawSlice(m.PicXml); }
            catch { picCanon = m.PicXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = picCanon,
            });
        }
    }

    // Phase 3c-media (legacy/external). Companion to EmitMediaForSlide for
    // <p:pic> video/audio hosts that GetMediaOnSlide rejects because they carry
    // no embedded MediaDataPart — the classic case is a PowerPoint 2007 movie
    // linked to an external file (<a:videoFile r:link="rIdN"/> where rIdN is a
    // TargetMode="External" file:// relationship, plus a local poster image in
    // the blipFill). The typed walk skips video/audio children (EmitSlide's
    // switch routes them here), and GetMediaOnSlide skips no-embed pics, so
    // without this pass the whole picture — poster and all — is silently lost.
    //
    // For each such pic we emit: one `add-part extrel` per external link rel
    // (re-creating the TargetMode="External" relationship with its pinned rId
    // so <a:videoFile r:link> no longer dangles), one `add-part image` per
    // local poster/blipFill image (pinned rId), then a raw-set append of the
    // <p:pic> verbatim. Same append-at-spTree-end model as EmitMediaForSlide.
    internal static void EmitExternalMediaForSlide(PowerPointHandler ppt, int slideNum,
                                                   string slidePath, List<BatchItem> items,
                                                   SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.ExternalMediaPicInfo> pics;
        try { pics = ppt.GetExternalMediaPicsOnSlide(slideNum); }
        catch { return; }
        if (pics.Count == 0) return;

        foreach (var p in pics)
        {
            foreach (var (rid, relType, target) in p.ExternalRels)
            {
                items.Add(new BatchItem
                {
                    Command = "add-part",
                    Parent = slidePath,
                    Type = "extrel",
                    Props = new Dictionary<string, string>(StringComparer.Ordinal)
                    {
                        ["rid"] = rid,
                        ["rel-type"] = relType,
                        ["target"] = target,
                    },
                });
            }

            if (p.ImageRids.Count > 0)
            {
                foreach (var img in ppt.GetSlideImagePartsByRelId(slideNum, p.ImageRids.ToList()))
                    items.Add(new BatchItem
                    {
                        Command = "add-part",
                        Parent = slidePath,
                        Type = "image",
                        Props = new Dictionary<string, string>(StringComparer.Ordinal)
                        {
                            ["rid"] = img.RelId,
                            ["content-type"] = img.ContentType,
                            ["data"] = img.Base64Data,
                        },
                    });
            }

            string picCanon;
            try { picCanon = NormalizeSlideRawSlice(p.PicXml); }
            catch { picCanon = p.PicXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = picCanon,
            });
        }
    }

    // Phase 3c-3d. Mirrors EmitMediaForSlide (Phase 3c-media). Per slide,
    // scan for <mc:AlternateContent> blocks whose <mc:Choice Requires="am3d">
    // carries <am3d:model3d>; emit an `add-part model3d` row that creates
    // the underlying ExtendedPart (.glb via AddExtendedPart, since the SDK
    // has no typed Model3DPart) + thumbnail ImagePart with SOURCE rIds
    // pinned via --prop. Then emit a raw-set append on
    // /p:sld/p:cSld/p:spTree carrying the AlternateContent XML verbatim —
    // the pinned rIds make the am3d:model3d r:embed, am3d:raster's
    // am3d:blip r:embed, AND the Fallback p:pic's a:blip r:embed all
    // resolve to the just-created parts.
    //
    // Slice handling: the AlternateContent block is canonicalised via the
    // same NormalizeSlideRawSlice pass as smartart / media. The am3d
    // namespace is NOT in the ambient list (p / a / r / mc), so its
    // xmlns:am3d declaration travels with the slice on the <mc:Choice>
    // — exactly the source form, so first- and post-replay rounds match.
    internal static void EmitModel3dForSlide(PowerPointHandler ppt, int slideNum,
                                              string slidePath, List<BatchItem> items,
                                              SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.Model3dInfo> models;
        try { models = ppt.GetModel3dOnSlide(slideNum); }
        catch { return; }
        if (models.Count == 0) return;

        foreach (var m in models)
        {
            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["data"] = Convert.ToBase64String(m.Model3dBytes),
                ["content-type"] = m.Model3dContentType,
                ["extension"] = m.Model3dExtension,
                ["model3d-rid"] = m.Model3dRelId,
                ["thumbnail-data"] = Convert.ToBase64String(m.ThumbnailBytes),
                ["thumbnail-content-type"] = m.ThumbnailContentType,
                ["thumbnail-rid"] = m.ThumbnailRelId,
            };
            items.Add(new BatchItem
            {
                Command = "add-part",
                Parent = slidePath,
                Type = "model3d",
                Props = props,
            });

            string acCanon;
            try { acCanon = NormalizeSlideRawSlice(m.AlternateContentXml); }
            catch { acCanon = m.AlternateContentXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = acCanon,
            });
        }
    }

    // Phase 3c-ole. Mirrors EmitMediaForSlide (Phase 3c-media) and
    // EmitModel3dForSlide (Phase 3c-3d). Per slide, scan for
    // <p:graphicFrame> blocks whose <a:graphicData uri=".../ole"> carries
    // a <p:oleObj>; emit an `add-part ole` row that creates the underlying
    // EmbeddedPackagePart (OOXML containers) or EmbeddedObjectPart (generic
    // binaries) + thumbnail icon ImagePart with SOURCE rIds pinned via
    // --prop. Then emit a raw-set append on /p:sld/p:cSld/p:spTree
    // carrying the graphicFrame XML verbatim — the pinned rIds make the
    // p:oleObj r:id AND the inner p:pic's a:blip r:embed both resolve to
    // the just-created parts.
    //
    // Slice handling: the graphicFrame block is canonicalised via the
    // same NormalizeSlideRawSlice pass as smartart / media / 3d. The
    // ambient namespaces (p / a / r) are stripped at the slice root;
    // anything exotic the OLE shape brings in stays on the slice as a
    // local xmlns declaration.
    internal static void EmitOleForSlide(PowerPointHandler ppt, int slideNum,
                                          string slidePath, List<BatchItem> items,
                                          SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.OleInfo> oles;
        try { oles = ppt.GetOlesOnSlide(slideNum); }
        catch { return; }
        if (oles.Count == 0) return;

        foreach (var o in oles)
        {
            if (o.LinkedTarget != null)
            {
                // LINKED OLE (TargetMode=External, <p:link/>): no payload part;
                // recreate the external relationship + the thumbnail image so
                // the verbatim graphicFrame's r:id / r:embed resolve.
                items.Add(new BatchItem
                {
                    Command = "add-part",
                    Parent = slidePath,
                    Type = "extrel",
                    Props = new Dictionary<string, string>(StringComparer.Ordinal)
                    {
                        ["rid"] = o.OleRelId,
                        ["rel-type"] = o.LinkedRelType ?? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
                        ["target"] = o.LinkedTarget,
                    },
                });
                items.Add(new BatchItem
                {
                    Command = "add-part",
                    Parent = slidePath,
                    Type = "image",
                    Props = new Dictionary<string, string>(StringComparer.Ordinal)
                    {
                        ["rid"] = o.ThumbnailRelId,
                        ["content-type"] = o.ThumbnailContentType,
                        ["data"] = Convert.ToBase64String(o.ThumbnailBytes),
                    },
                });
                string lgfCanon;
                try { lgfCanon = NormalizeSlideRawSlice(o.GraphicFrameXml); }
                catch { lgfCanon = o.GraphicFrameXml; }
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = slidePath,
                    Xpath = "/p:sld/p:cSld/p:spTree",
                    Action = "append",
                    Xml = lgfCanon,
                });
                continue;
            }

            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["data"] = Convert.ToBase64String(o.OleBytes),
                ["content-type"] = o.OleContentType,
                ["extension"] = o.OleExtension,
                ["ole-rid"] = o.OleRelId,
                ["thumbnail-data"] = Convert.ToBase64String(o.ThumbnailBytes),
                ["thumbnail-content-type"] = o.ThumbnailContentType,
                ["thumbnail-rid"] = o.ThumbnailRelId,
            };
            items.Add(new BatchItem
            {
                Command = "add-part",
                Parent = slidePath,
                Type = "ole",
                Props = props,
            });

            string gfCanon;
            try { gfCanon = NormalizeSlideRawSlice(o.GraphicFrameXml); }
            catch { gfCanon = o.GraphicFrameXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = gfCanon,
            });
        }
    }

    // Generic <mc:AlternateContent> catch-all. Scans the slide's raw XML for
    // AlternateContent blocks directly under <p:spTree> that match none of
    // the specific emitters (am3d:model3d / SmartArt / media / OLE — each
    // already detects its own marker and emits add-part + raw-set
    // passthrough). The remaining AlternateContent blocks carry
    // emerging-feature content the semantic walk doesn't model
    // (NodeBuilder's EnumerateRenderableElements explicitly skips the
    // mc:AlternateContent wrapper to avoid double-counting Choice + Fallback
    // <p:sp> children) — without this catch-all, dump→replay silently drops
    // them. Replay re-appends the verbatim XML at the same spTree level via
    // raw-set; the source byte form survives.
    //
    // CONSISTENCY(mc-alt-skip-recovery): pairs with NodeBuilder's
    // mc:AlternateContent skip in EnumerateRenderableElements — the skip
    // suppresses semantic enumeration, this emitter re-injects the raw block.
    //
    // Skip markers carry the specific-handler signal so a deck with
    // model3d / 3D content goes through EmitModel3dForSlide, not this
    // catch-all. AmThirdD/SmartArt/Media/OLE markers come from the source
    // XML directly; matching is a substring scan to avoid a full XML parse.
    internal static void EmitGenericAlternateContentForSlide(
        PowerPointHandler ppt, int slideNum, string slidePath,
        List<BatchItem> items, SlideEmitContext ctx)
    {
        string xml;
        try { xml = ppt.Raw(slidePath); }
        catch { return; }

        // Bound the scan to spTree contents so transition/timing
        // AlternateContent (handled by the exotic-content scanner) is
        // off-limits.
        var spTreeOpen = xml.IndexOf("<p:spTree", StringComparison.Ordinal);
        var spTreeClose = xml.LastIndexOf("</p:spTree>", StringComparison.Ordinal);
        if (spTreeOpen < 0 || spTreeClose <= spTreeOpen) return;
        var spTreeRegion = xml.Substring(spTreeOpen, spTreeClose - spTreeOpen);

        // CONSISTENCY(mc-alt-zorder): preserve sibling order so 3D-model /
        // emerging-content AlternateContent lands at the same z-order
        // position the source intended. The semantic walk emits regular
        // <p:sp>/<p:pic>/<p:graphicFrame> first, then this catch-all ran
        // with action=append — pushing AlternateContent to the end of
        // spTree and inverting z-order (3D model floated above text boxes
        // it was meant to sit behind). Count how many shape-like siblings
        // precede each AlternateContent in source order; insert before the
        // (N+1)th renderable child on replay.
        int cursor = 0;
        while (true)
        {
            var altIdx = spTreeRegion.IndexOf("<mc:AlternateContent", cursor, StringComparison.Ordinal);
            if (altIdx < 0) break;
            var altEnd = spTreeRegion.IndexOf("</mc:AlternateContent>", altIdx, StringComparison.Ordinal);
            if (altEnd < 0) break;
            altEnd += "</mc:AlternateContent>".Length;
            var slice = spTreeRegion.Substring(altIdx, altEnd - altIdx);

            // Only AlternateContent that is a DIRECT child of <p:spTree>
            // is in scope here. Equation shapes (<p:sp> containing
            // <p:txBody><a:p><mc:AlternateContent> with <a14:m>/<m:oMath>)
            // already round-trip through AddEquation — re-emitting their
            // nested AlternateContent at slide root would duplicate the
            // math block as a loose <p:spTree> child, which PowerPoint
            // renders as plain runs without proper sup/sub binding (i 2,
            // x 2, α 1, β 2 instead of i², x², α₁, β²). Same applies to
            // any <p:sp>-nested AlternateContent (e.g. inline svg fallback
            // patterns). Detect nesting by counting unclosed <p:sp> /
            // <p:grpSp> opening tags before altIdx.
            bool nested = IsInsideShapeOrGroup(spTreeRegion, altIdx);

            // Skip blocks owned by a specific emitter.
            bool skip = nested
                || slice.Contains("am3d:", StringComparison.Ordinal)
                || slice.Contains("Requires=\"am3d\"", StringComparison.Ordinal)
                || slice.Contains("<dgm:relIds", StringComparison.Ordinal)
                || slice.Contains("<p:oleObj", StringComparison.Ordinal)
                || slice.Contains("<p:audio", StringComparison.Ordinal)
                || slice.Contains("<p:video", StringComparison.Ordinal);

            int precedingShapeCount = skip
                ? 0
                : CountRenderableSiblingsBefore(spTreeRegion, altIdx);
            // TYPED siblings after this AlternateContent in source order —
            // those exist on the replay slide before any raw-set runs, so if
            // any follow we can pin a positional `insertbefore` target. A
            // following AlternateContent does NOT count: it is emitted after
            // us in this same loop, so pinning on it forward-references an
            // element that doesn't exist yet ("XPath matched no elements",
            // zoom-test / prezi-demo: two consecutive trailing zoom frames).
            // When only AlternateContent (or nothing) follows, append — the
            // later blocks also append in source order, preserving layout.
            int followingShapeCount = skip
                ? 0
                : CountTypedRenderableSiblingsAfter(spTreeRegion, altEnd);

            cursor = altEnd;
            if (skip) continue;

            string canon;
            try { canon = NormalizeSlideRawSlice(slice); }
            catch { canon = slice; }

            // Carry image parts referenced (r:embed / r:link) inside the block.
            // The block is re-inserted verbatim, so any <a:blip r:embed> in a
            // Fallback <p:pic> (e.g. a math-equation AlternateContent: Choice
            // a14 <m:oMath> + Fallback picture) must resolve — otherwise the
            // rId dangles and PowerPoint refuses the deck (0x80070570,
            // sample07). Pin the SOURCE rId so the verbatim r:embed resolves.
            var altRids = new HashSet<string>(StringComparer.Ordinal);
            foreach (System.Text.RegularExpressions.Match rm in
                     System.Text.RegularExpressions.Regex.Matches(slice, @"r:(?:embed|link)=""([^""]+)"""))
                altRids.Add(rm.Groups[1].Value);

            // Internal slide-jump links inside the block (<a:hlinkClick
            // r:id="rIdN" action="…hlinksldjump"> — every zoom frame carries
            // one next to its preview blip). The verbatim slice keeps the
            // source rId, but the relationship targets another SLIDE, which
            // is re-added later in the batch — so DEFER the pinned slide
            // relationship to the end of the emit (ctx.DeferredLinks replays
            // after every slide exists). Without this the jump r:id dangled
            // and the zoom frames were functionally dead (semantic validate
            // reds on every zoom/prezi deck). Mirrors the table txBodyRaw
            // slide-jump carrier.
            try
            {
                var altIdRids = new HashSet<string>(StringComparer.Ordinal);
                foreach (System.Text.RegularExpressions.Match rm in
                         System.Text.RegularExpressions.Regex.Matches(slice, @"r:id=""([^""]+)"""))
                    altIdRids.Add(rm.Groups[1].Value);
                foreach (var (jumpRelId, targetOrd) in ppt.GetSlideInternalSlideJumpRels(slideNum, altIdRids))
                {
                    ctx.DeferredLinks.Add(new BatchItem
                    {
                        Command = "add-part",
                        Parent = slidePath,
                        Type = "sliderel",
                        Props = new Dictionary<string, string>
                        {
                            ["rid"] = jumpRelId,
                            ["target"] = targetOrd.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        },
                    });
                }
            }
            catch { /* best-effort */ }

            // chartEx blocks (cx: extension charts — funnel/sunburst/treemap):
            // the Choice's <cx:chart r:id> references an ExtendedChartPart that
            // no other pass re-creates; carry it (plus colors/style sidecars
            // and the embedded xlsx) or the verbatim slice's rId dangles and
            // PowerPoint refuses the deck (funnel-pp1).
            if (slice.Contains("chartex", StringComparison.OrdinalIgnoreCase))
            {
                var cxRids = new HashSet<string>(StringComparer.Ordinal);
                foreach (System.Text.RegularExpressions.Match rm in
                         System.Text.RegularExpressions.Regex.Matches(slice, @"r:id=""(rId\d+)"""))
                    cxRids.Add(rm.Groups[1].Value);
                foreach (var cx in ppt.GetChartExPartsByRelId(slideNum, cxRids))
                {
                    var cxProps = new Dictionary<string, string>(StringComparer.Ordinal)
                    {
                        ["rid"] = cx.RelId,
                        ["xml"] = cx.XmlBase64,
                    };
                    if (cx.ColorsRelId != null && cx.ColorsBase64 != null)
                    { cxProps["colors-rid"] = cx.ColorsRelId; cxProps["colors"] = cx.ColorsBase64; }
                    if (cx.StyleRelId != null && cx.StyleBase64 != null)
                    { cxProps["style-rid"] = cx.StyleRelId; cxProps["style"] = cx.StyleBase64; }
                    if (cx.PackageRelId != null && cx.PackageBase64 != null)
                    {
                        cxProps["package-rid"] = cx.PackageRelId;
                        cxProps["package"] = cx.PackageBase64;
                        if (cx.PackageContentType != null) cxProps["package-content-type"] = cx.PackageContentType;
                    }
                    items.Add(new BatchItem
                    {
                        Command = "add-part",
                        Parent = slidePath,
                        Type = "chartex",
                        Props = cxProps,
                    });
                }
            }

            if (altRids.Count > 0)
            {
                foreach (var img in ppt.GetSlideImagePartsByRelId(slideNum, altRids))
                    items.Add(new BatchItem
                    {
                        Command = "add-part",
                        Parent = slidePath,
                        Type = "image",
                        Props = new Dictionary<string, string>
                        {
                            ["rid"] = img.RelId,
                            ["content-type"] = img.ContentType,
                            ["data"] = img.Base64Data,
                        },
                    });
            }

            if (followingShapeCount > 0)
            {
                // Replay spTree (built by the semantic walk) lists
                // structural <p:nvGrpSpPr>/<p:grpSpPr> followed by the
                // appended shapes/pictures/graphicFrames/groupshapes/
                // connectionShapes in source order. precedingShapeCount
                // counts renderable elements before this AlternateContent
                // in the source spTree, so the (N+1)th renderable child on
                // replay is the right insertion target — earlier
                // AlternateContent siblings (emitted before us in this
                // same loop) bump the index when they sit before more
                // renderable shapes too.
                var renderableXpath =
                    "/p:sld/p:cSld/p:spTree/*[self::p:sp or self::p:pic "
                    + "or self::p:cxnSp or self::p:graphicFrame or self::p:grpSp "
                    + "or self::mc:AlternateContent]"
                    + $"[{precedingShapeCount + 1}]";
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = slidePath,
                    Xpath = renderableXpath,
                    Action = "insertbefore",
                    Xml = canon,
                });
            }
            else
            {
                // No renderable sibling follows in the source; append is
                // the right action even relative to other AlternateContent
                // blocks (which also append in source order).
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = slidePath,
                    Xpath = "/p:sld/p:cSld/p:spTree",
                    Action = "append",
                    Xml = canon,
                });
            }
        }
    }

    // True when <paramref name="offset"/> falls inside an unclosed
    // <p:sp> / <p:grpSp> / <p:pic> / <p:cxnSp> / <p:graphicFrame> region —
    // i.e. the AlternateContent at that position is a descendant of a shape
    // (e.g. an equation txBody, or a Mac-authored <p:pic> whose blipFill is
    // wrapped in mc:AlternateContent) rather than a direct child of
    // <p:spTree>. Such nested AlternateContent is owned by that element's own
    // emitter (EmitPicture / EmitShape / chart / smartart / table / ole) and
    // must NOT be re-emitted as a loose spTree child (double-injection +
    // schema-invalid <p:blipFill> under <p:spTree>). Counts opening minus
    // closing tags via a regex sweep; treats self-closing forms as
    // immediately balanced.
    private static bool IsInsideShapeOrGroup(string spTreeRegion, int offset)
    {
        var rx = new System.Text.RegularExpressions.Regex(
            @"<(/?)(p:sp|p:grpSp|p:pic|p:cxnSp|p:graphicFrame)(\s[^/>]*?/?|/?)>",
            System.Text.RegularExpressions.RegexOptions.Compiled);
        int depth = 0;
        foreach (System.Text.RegularExpressions.Match m in rx.Matches(spTreeRegion))
        {
            if (m.Index >= offset) break;
            bool isClose = m.Groups[1].Value == "/";
            bool isSelfClose = m.Groups[3].Value.EndsWith('/');
            if (isClose) depth--;
            else if (!isSelfClose) depth++;
        }
        return depth > 0;
    }

    // Count <p:sp>/<p:pic>/<p:cxnSp>/<p:graphicFrame>/<p:grpSp>/
    // <mc:AlternateContent> opening tags that are DIRECT children of <p:spTree>
    // and occur before <paramref name="beforeOffset"/>. Tags nested inside an
    // mc:AlternateContent / p:grpSp / p:sp / p:graphicFrame (e.g. the <p:sp>
    // pair living under mc:Choice + mc:Fallback that wraps an OOMath equation
    // shape) must NOT count — they are descendants, not siblings, of any
    // spTree-level AlternateContent we are positioning. R65 fuzzer-1: a
    // blank slide whose only direct-child renderable is one mc:AlternateContent
    // produced precedingShapeCount=0 / followingShapeCount=2 (the two nested
    // <p:sp>), routed through the insertbefore[1] branch — but the replay
    // slide had no first-shape anchor, xpath matched nothing, fragment lost.
    private static int CountRenderableSiblingsBefore(string spTreeRegion, int beforeOffset)
    {
        // Sweep every open/close/self-close of any container or renderable
        // tag in source order, maintain a depth counter, and only count
        // renderable opens that land at depth 0 (direct spTree children).
        var rx = new System.Text.RegularExpressions.Regex(
            @"<(/?)(p:sp|p:pic|p:cxnSp|p:graphicFrame|p:grpSp|mc:AlternateContent|mc:Choice|mc:Fallback)\b([^>]*?)(/?)>",
            System.Text.RegularExpressions.RegexOptions.Compiled);
        int depth = 0;
        int count = 0;
        var renderable = new HashSet<string>(StringComparer.Ordinal)
        {
            "p:sp", "p:pic", "p:cxnSp", "p:graphicFrame", "p:grpSp", "mc:AlternateContent",
        };
        foreach (System.Text.RegularExpressions.Match m in rx.Matches(spTreeRegion))
        {
            if (m.Index >= beforeOffset) break;
            bool isClose = m.Groups[1].Value == "/";
            string tag = m.Groups[2].Value;
            bool isSelfClose = m.Groups[4].Value == "/";
            if (isClose)
            {
                depth--;
            }
            else
            {
                if (depth == 0 && renderable.Contains(tag)) count++;
                if (!isSelfClose) depth++;
            }
        }
        return count;
    }

    // Depth-0 TYPED renderable children (sp/pic/cxnSp/graphicFrame/grpSp —
    // excluding mc:AlternateContent) at or after <paramref name="afterOffset"/>.
    // Used to decide insertbefore vs append for a raw AlternateContent slice:
    // typed elements are created by the semantic walk BEFORE any raw-set runs,
    // so they are valid positional anchors — a later AlternateContent sibling
    // is NOT (it is emitted after us in the same loop), and pinning on it
    // forward-references an element that doesn't exist yet at execution time.
    private static int CountTypedRenderableSiblingsAfter(string spTreeRegion, int afterOffset)
    {
        var rx = new System.Text.RegularExpressions.Regex(
            @"<(/?)(p:sp|p:pic|p:cxnSp|p:graphicFrame|p:grpSp|mc:AlternateContent|mc:Choice|mc:Fallback)\b([^>]*?)(/?)>",
            System.Text.RegularExpressions.RegexOptions.Compiled);
        int depth = 0;
        int count = 0;
        var typed = new HashSet<string>(StringComparer.Ordinal)
        {
            "p:sp", "p:pic", "p:cxnSp", "p:graphicFrame", "p:grpSp",
        };
        foreach (System.Text.RegularExpressions.Match m in rx.Matches(spTreeRegion))
        {
            bool isClose = m.Groups[1].Value == "/";
            string tag = m.Groups[2].Value;
            bool isSelfClose = m.Groups[4].Value == "/";
            if (isClose)
            {
                depth--;
            }
            else
            {
                if (depth == 0 && m.Index >= afterOffset && typed.Contains(tag)) count++;
                if (!isSelfClose) depth++;
            }
        }
        return count;
    }
}
