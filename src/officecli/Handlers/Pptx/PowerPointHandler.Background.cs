// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Slide Background ====================

    /// <summary>
    /// Apply a background to a slide, slide layout, or slide master.
    ///
    /// Supported values for the "background" property:
    ///   RRGGBB               solid color        e.g. "FF0000"
    ///   none / transparent   remove background
    ///   C1-C2                gradient           e.g. "FF0000-0000FF"
    ///   C1-C2-angle          gradient + angle   e.g. "FF0000-0000FF-45"
    ///   C1-C2-C3             3-stop gradient    e.g. "FF0000-FFFF00-0000FF"
    ///   image:path           image fill         e.g. "image:/tmp/bg.png"
    ///
    /// Accepts SlidePart, SlideLayoutPart, or SlideMasterPart — all three parts share
    /// the same p:bg / p:bgPr schema inside CommonSlideData.
    /// </summary>
    internal record BackgroundImageOptions(string? Mode = null, int? Alpha = null, int? Scale = null);

    /// <summary>
    /// If properties contain only background.mode/alpha/scale (no "background" key),
    /// mutate the existing image fill in place — preserves Blip.Embed so the image
    /// part is not duplicated.
    /// </summary>
    internal static void MaybeMutateExistingBackgroundImage(
        OpenXmlPart part, Dictionary<string, string> properties)
    {
        bool hasBackground = properties.Keys.Any(k => k.Equals("background", StringComparison.OrdinalIgnoreCase));
        if (hasBackground) return;
        var opts = ReadBackgroundImageOptions(properties);
        if (opts == null) return;
        MutateBackgroundImageFill(part, opts);
    }

    internal static BackgroundImageOptions? ReadBackgroundImageOptions(Dictionary<string, string> properties)
    {
        string? Lookup(string k) => properties
            .Where(p => p.Key.Equals(k, StringComparison.OrdinalIgnoreCase))
            .Select(p => p.Value).FirstOrDefault();

        var mode = Lookup("background.mode");
        var alphaStr = Lookup("background.alpha");
        var scaleStr = Lookup("background.scale");
        if (mode == null && alphaStr == null && scaleStr == null) return null;

        int? alpha = null, scale = null;
        if (alphaStr != null && !int.TryParse(alphaStr, out var a))
            throw new ArgumentException($"background.alpha must be an integer 0..100, got '{alphaStr}'");
        else if (alphaStr != null) alpha = int.Parse(alphaStr);
        if (scaleStr != null && !int.TryParse(scaleStr, out var s))
            throw new ArgumentException($"background.scale must be an integer 1..500, got '{scaleStr}'");
        else if (scaleStr != null) scale = int.Parse(scaleStr);
        return new BackgroundImageOptions(mode, alpha, scale);
    }

    private static void ApplyBackground(OpenXmlPart part, string value, BackgroundImageOptions? imgOpts = null)
    {
        // Normalize alternative gradient format: "LINEAR;C1;C2;angle" → "C1-C2-angle"
        value = NormalizeGradientValue(value);

        // background.mode/alpha/scale are image-only; reject early if paired with a
        // non-image value so the user isn't fooled by a success echo for a no-op.
        var isImage = value.StartsWith("image:", StringComparison.OrdinalIgnoreCase);
        var isClear = value.Equals("none", StringComparison.OrdinalIgnoreCase)
                   || value.Equals("transparent", StringComparison.OrdinalIgnoreCase)
                   || value.Equals("clear", StringComparison.OrdinalIgnoreCase);
        if (imgOpts != null && !isImage)
        {
            var opt = imgOpts.Mode != null ? "background.mode"
                    : imgOpts.Alpha != null ? "background.alpha"
                    : "background.scale";
            var kind = isClear ? "none/transparent" : "solid/gradient";
            throw new ArgumentException(
                $"{opt} is only valid with an image background (current background={kind}); " +
                "pair with background=image:<path>");
        }

        var cSld = GetCommonSlideData(part)
            ?? throw new InvalidOperationException($"{part.GetType().Name} has no CommonSlideData");

        // Build the new background element (or pre-buffered image bytes) BEFORE mutating
        // the existing bg. A validation failure (bad color, missing image, bad options)
        // must not destroy the prior bg — matches the atomicity contract of ApplyShapeFill
        // and the build-first-then-swap pattern used in MutateBackgroundImageFill.
        Background? newBg = null;
        (byte[] Bytes, PartTypeInfo PartType, Background Bg)? prepared = null;

        if (isClear)
        {
            // sentinel: leave newBg null; handled below as "remove-only".
        }
        else if (value.StartsWith("image:", StringComparison.OrdinalIgnoreCase))
        {
            var imagePath = value[6..].Trim();
            prepared = PrepareBackgroundImage(imagePath, imgOpts);
        }
        else
        {
            var bgPr = new BackgroundProperties();
            if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase) ||
                value.StartsWith("path:", StringComparison.OrdinalIgnoreCase))
            {
                if (!IsGradientColorString(value))
                    throw new ArgumentException(
                        $"Invalid gradient specification: '{value}'. " +
                        "Radial/path gradients require at least 2 hex colors, e.g. 'radial:FF0000-0000FF'");
                bgPr.Append(BuildGradientFill(value));
            }
            else if (IsGradientColorString(value))
            {
                bgPr.Append(BuildGradientFill(value));
            }
            else
            {
                bgPr.Append(BuildSolidFill(value));
            }
            newBg = new Background();
            newBg.Append(bgPr);
        }

        // All validation passed — now safe to tear down the old bg.
        DeleteBackgroundImageParts(cSld, part);
        cSld.Background?.Remove();

        if (isClear) return;

        if (prepared is (byte[] bytes, PartTypeInfo pt, Background imgBg))
        {
            var imagePart = AddBackgroundImagePart(part, pt);
            using (var ms = new MemoryStream(bytes))
                imagePart.FeedData(ms);
            var relId = GetBackgroundImageRelId(part, imagePart);
            // Set the rel id on the prepared Blip (placeholder at build time).
            var blip = imgBg.Descendants<Drawing.Blip>().First();
            blip.Embed = relId;
            newBg = imgBg;
        }

        // Insert before ShapeTree — schema order: p:bg → p:spTree. If spTree is missing
        // (externally corrupted), create a minimal one so the resulting p:cSld is still
        // schema-valid (spTree is mandatory; PrependChild without it writes invalid XML).
        var shapeTree = cSld.ShapeTree;
        if (shapeTree == null)
        {
            shapeTree = new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(new Drawing.TransformGroup()));
            cSld.AppendChild(shapeTree);
        }
        cSld.InsertBefore(newBg!, shapeTree);
    }

    // CONSISTENCY(slide-background-part): SlidePart/SlideLayoutPart/SlideMasterPart all
    // share the p:bg schema but have no common API. Each overload keeps the call-site simple.
    private static void ApplySlideBackground(SlidePart slidePart, string value)
        => ApplyBackground(slidePart, value);

    private static CommonSlideData? GetCommonSlideData(OpenXmlPart part) => part switch
    {
        SlidePart sp => sp.Slide?.CommonSlideData,
        SlideLayoutPart lp => lp.SlideLayout?.CommonSlideData,
        SlideMasterPart mp => mp.SlideMaster?.CommonSlideData,
        _ => null
    };

    internal static void SaveBackgroundRoot(OpenXmlPart part)
    {
        switch (part)
        {
            case SlidePart sp: sp.Slide?.Save(); break;
            case SlideLayoutPart lp: lp.SlideLayout?.Save(); break;
            case SlideMasterPart mp: mp.SlideMaster?.Save(); break;
        }
    }

    private static void DeleteBackgroundImageParts(CommonSlideData cSld, OpenXmlPart part)
    {
        var bgPr = cSld.Background?.BackgroundProperties;
        if (bgPr == null) return;
        foreach (var bf in bgPr.Elements<Drawing.BlipFill>().ToList())
        {
            var embed = bf.GetFirstChild<Drawing.Blip>()?.Embed?.Value;
            if (string.IsNullOrEmpty(embed)) continue;
            try
            {
                var refPart = part.GetPartById(embed);
                if (refPart is ImagePart ip)
                    part.DeletePart(ip);
            }
            catch { /* rel may be missing or already gone */ }
        }
    }

    private static ImagePart AddBackgroundImagePart(OpenXmlPart part, PartTypeInfo partType) => part switch
    {
        SlidePart sp => sp.AddImagePart(partType),
        SlideLayoutPart lp => lp.AddImagePart(partType),
        SlideMasterPart mp => mp.AddImagePart(partType),
        _ => throw new NotSupportedException($"{part.GetType().Name} does not support image parts")
    };

    private static string GetBackgroundImageRelId(OpenXmlPart part, ImagePart imagePart) => part switch
    {
        SlidePart sp => sp.GetIdOfPart(imagePart),
        SlideLayoutPart lp => lp.GetIdOfPart(imagePart),
        SlideMasterPart mp => mp.GetIdOfPart(imagePart),
        _ => throw new NotSupportedException($"{part.GetType().Name} does not support image parts")
    };

    /// <summary>
    /// Resolve an image source and build a Background element with a placeholder Blip
    /// (Embed to be filled in once an ImagePart actually exists). Does not mutate the
    /// document — if anything throws here, the caller's prior bg is still intact.
    /// </summary>
    private static (byte[] Bytes, PartTypeInfo PartType, Background Bg) PrepareBackgroundImage(
        string imagePath, BackgroundImageOptions? opts)
    {
        // Validate options up-front.
        if (opts?.Scale != null)
        {
            var m = (opts.Mode ?? "stretch").ToLowerInvariant();
            if (m != "tile")
                throw new ArgumentException(
                    $"background.scale is only valid with background.mode=tile (got mode={m}); " +
                    "set background.mode=tile together with background.scale");
        }
        if (opts?.Alpha is int preAlpha && (preAlpha < 0 || preAlpha > 100))
            throw new ArgumentException($"background.alpha must be 0..100, got {preAlpha}");
        // Mode + scale validation via BuildBlipFillMode (throws on bad mode / scale range).
        var modeChild = BuildBlipFillMode(opts);

        var (stream, partType) = OfficeCli.Core.ImageSource.Resolve(imagePath);
        byte[] bytes;
        using (stream)
        using (var buf = new MemoryStream())
        {
            stream.CopyTo(buf);
            bytes = buf.ToArray();
        }

        var blip = new Drawing.Blip(); // Embed set later, once an ImagePart exists
        if (opts?.Alpha is int alpha && alpha < 100)
            blip.Append(new Drawing.AlphaModulationFixed { Amount = alpha * 1000 });

        var blipFill = new Drawing.BlipFill();
        blipFill.Append(blip);
        blipFill.Append(modeChild);

        var bgPr = new BackgroundProperties();
        bgPr.Append(blipFill);
        var bg = new Background();
        bg.Append(bgPr);
        return (bytes, partType, bg);
    }

    private static void ApplyBackgroundImageFill(
        BackgroundProperties bgPr, OpenXmlPart part, string imagePath,
        BackgroundImageOptions? opts = null)
    {
        // Kept for legacy call sites that invoke ApplyBackgroundImageFill directly.
        // Validate up-front so the image part isn't created just to be orphaned by a later throw.
        if (opts?.Scale != null)
        {
            var m = (opts.Mode ?? "stretch").ToLowerInvariant();
            if (m != "tile")
                throw new ArgumentException(
                    $"background.scale is only valid with background.mode=tile (got mode={m}); " +
                    "set background.mode=tile together with background.scale");
        }
        if (opts?.Alpha is int preAlpha && (preAlpha < 0 || preAlpha > 100))
            throw new ArgumentException($"background.alpha must be 0..100, got {preAlpha}");

        var (stream, partType) = OfficeCli.Core.ImageSource.Resolve(imagePath);
        using var streamDispose = stream;

        var imagePart = AddBackgroundImagePart(part, partType);
        imagePart.FeedData(stream);
        var relId = GetBackgroundImageRelId(part, imagePart);

        var blip = new Drawing.Blip { Embed = relId };
        // Alpha: a:alphaModFix inside a:blip. amt is 0..100000 (100000 = opaque).
        // Skip emitting when alpha=100 so apply/mutate both converge to the same XML.
        if (opts?.Alpha is int alpha && alpha < 100)
        {
            blip.Append(new Drawing.AlphaModulationFixed { Amount = alpha * 1000 });
        }

        var blipFill = new Drawing.BlipFill();
        blipFill.Append(blip);
        // Schema order inside a:blipFill: a:blip → a:srcRect → {a:tile | a:stretch}.
        blipFill.Append(BuildBlipFillMode(opts));
        bgPr.Append(blipFill);
    }

    /// <summary>
    /// Modify mode/alpha/scale of an existing image background in place without
    /// touching the Blip.Embed rel — so the image part is not duplicated or orphaned.
    /// Throws if the current background is not an image fill.
    /// </summary>
    internal static void MutateBackgroundImageFill(OpenXmlPart part, BackgroundImageOptions opts)
    {
        var cSld = GetCommonSlideData(part)
            ?? throw new InvalidOperationException($"{part.GetType().Name} has no CommonSlideData");
        var bgPr = cSld.Background?.BackgroundProperties
            ?? throw new ArgumentException(
                "background.mode/alpha/scale requires an existing image background; " +
                "set background=image:<path> first");
        var blipFill = bgPr.GetFirstChild<Drawing.BlipFill>()
            ?? throw new ArgumentException(
                "background.mode/alpha/scale requires an image background, but the current " +
                "background is solid/gradient; set background=image:<path> first");
        var blip = blipFill.GetFirstChild<Drawing.Blip>()
            ?? throw new InvalidOperationException("BlipFill has no Blip child");
        if (string.IsNullOrEmpty(blip.Embed?.Value))
            throw new ArgumentException(
                "Cannot mutate background image: the existing blip has no r:embed rel. " +
                "Re-set background=image:<path> to rebind.");

        // Alpha: remove any existing alphaModFix, then re-add if specified.
        // Null alpha means "leave existing alpha alone" — matches the partial-update semantic.
        if (opts.Alpha is int alpha)
        {
            if (alpha < 0 || alpha > 100)
                throw new ArgumentException($"background.alpha must be 0..100, got {alpha}");
            blip.Elements<Drawing.AlphaModulationFixed>().ToList().ForEach(e => e.Remove());
            if (alpha < 100) // 100 = opaque, default, skip emitting
                blip.Append(new Drawing.AlphaModulationFixed { Amount = alpha * 1000 });
        }

        // Mode/scale: replace the existing tile/stretch child. If either is specified,
        // we need current values for the other to preserve them.
        if (opts.Mode != null || opts.Scale != null)
        {
            var (curMode, curScale) = ReadCurrentBlipFillMode(blipFill);
            // Normalize incoming mode so the scale-compat check doesn't reject "TILE"
            // simply because it wasn't lowercased. BuildBlipFillMode also lowercases.
            var effectiveMode = (opts.Mode ?? curMode).Trim().ToLowerInvariant();
            // Scale is meaningful only in tile mode — reject scale-on-stretch/center to
            // prevent a silent no-op. Callers must set mode=tile to use scale.
            if (opts.Scale != null && effectiveMode != "tile")
                throw new ArgumentException(
                    $"background.scale is only valid with background.mode=tile (current mode: {effectiveMode}); " +
                    "set background.mode=tile together with background.scale");
            var merged = new BackgroundImageOptions(
                Mode: effectiveMode,
                Scale: opts.Scale ?? curScale);
            // Build first, then swap — BuildBlipFillMode validates and may throw, so we
            // must not remove the existing child before the new one is ready.
            var newChild = BuildBlipFillMode(merged);
            blipFill.Elements<Drawing.Tile>().ToList().ForEach(e => e.Remove());
            blipFill.Elements<Drawing.Stretch>().ToList().ForEach(e => e.Remove());
            blipFill.Append(newChild);
        }
    }

    private static (string Mode, int Scale) ReadCurrentBlipFillMode(Drawing.BlipFill blipFill)
    {
        var tile = blipFill.GetFirstChild<Drawing.Tile>();
        if (tile == null) return ("stretch", 100);
        var sx = tile.HorizontalRatio?.Value ?? 100000;
        var algn = tile.Alignment?.Value;
        if (algn == Drawing.RectangleAlignmentValues.Center && sx == 100000)
            return ("center", 100);
        return ("tile", (int)Math.Round(sx / 1000.0));
    }

    private static OpenXmlElement BuildBlipFillMode(BackgroundImageOptions? opts)
    {
        var mode = (opts?.Mode ?? "stretch").Trim().ToLowerInvariant();
        var scale = opts?.Scale ?? 100;
        if (scale < 1 || scale > 500)
            throw new ArgumentException($"background.scale must be 1..500, got {scale}");
        var sxSy = scale * 1000; // 100% == 100000

        return mode switch
        {
            "stretch" => new Drawing.Stretch(new Drawing.FillRectangle()),
            "tile" => new Drawing.Tile
            {
                HorizontalRatio = sxSy,
                VerticalRatio = sxSy,
                Alignment = Drawing.RectangleAlignmentValues.TopLeft,
                Flip = Drawing.TileFlipValues.None,
            },
            // Center = tile anchored at center with no scaling. Matches LibreOffice's
            // FillBitmapMode_NO_REPEAT → oox export pattern (WriteXGraphicTile algn=ctr).
            "center" => new Drawing.Tile
            {
                HorizontalRatio = 100000,
                VerticalRatio = 100000,
                Alignment = Drawing.RectangleAlignmentValues.Center,
                Flip = Drawing.TileFlipValues.None,
            },
            _ => throw new ArgumentException($"background.mode must be stretch/tile/center, got '{mode}'"),
        };
    }

    // ==================== Read back ====================

    /// <summary>
    /// Populate Format["background"] on a slide DocumentNode.
    /// Values mirror the input format: hex for solid, "C1-C2[-angle]" for gradient, "image" for blip.
    /// </summary>
    private static void ReadSlideBackground(Slide slide, DocumentNode node)
        => ReadBackground(slide.CommonSlideData, node);

    internal static void ReadBackground(CommonSlideData? cSld, DocumentNode node)
    {
        if (cSld?.Background == null) return;

        var bgPr = cSld.Background.BackgroundProperties;
        if (bgPr == null)
        {
            // Theme-referenced background (p:bgRef). Not settable via our set commands,
            // but should surface on get so users see that a bg exists.
            var bgRef = cSld.Background.GetFirstChild<BackgroundStyleReference>();
            if (bgRef != null)
            {
                var color = ReadColorFromElement(bgRef);
                node.Format["background"] = color != null ? $"ref:{color}" : "ref";
                if (bgRef.Index?.HasValue == true)
                    node.Format["background.ref"] = (int)bgRef.Index.Value;
            }
            return;
        }

        var solidFill = bgPr.GetFirstChild<Drawing.SolidFill>();
        var gradFill  = bgPr.GetFirstChild<Drawing.GradientFill>();
        var blipFill  = bgPr.GetFirstChild<Drawing.BlipFill>();

        if (solidFill != null)
        {
            var bgColor = ReadColorFromFill(solidFill);
            if (bgColor != null) node.Format["background"] = bgColor;
        }
        else if (gradFill != null)
        {
            var stopEls = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
            // Emit @pct only when the stop deviates from the uniform default so the common
            // case round-trips to bare "C1-C2[-Cn]". Scheme colors are handled via
            // ReadColorFromElement; a hex-only read dropped them as "?".
            var stops = stopEls?.Select((gs, i) =>
            {
                var color = ReadColorFromElement(gs) ?? "?";
                if (gs.Position?.Value is int pos)
                {
                    var n = stopEls.Count;
                    var expected = n <= 1 ? 0 : (int)((long)i * 100000 / (n - 1));
                    if (pos != expected)
                        return $"{color}@{(int)Math.Round(pos / 1000.0)}";
                }
                return color;
            }).ToList();
            if (stops?.Count > 0)
            {
                var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
                if (pathGrad != null)
                {
                    var fillRect = pathGrad.GetFirstChild<Drawing.FillToRectangle>();
                    var focus = "center";
                    if (fillRect != null)
                    {
                        var fl = fillRect.Left?.Value ?? 50000;
                        var ft = fillRect.Top?.Value ?? 50000;
                        focus = (fl, ft) switch
                        {
                            (0, 0) => "tl",
                            ( >= 100000, 0) => "tr",
                            (0, >= 100000) => "bl",
                            ( >= 100000, >= 100000) => "br",
                            _ => "center"
                        };
                    }
                    var prefix = pathGrad.Path?.Value == Drawing.PathShadeValues.Shape ? "path" : "radial";
                    node.Format["background"] = $"{prefix}:{string.Join("-", stops)}-{focus}";
                }
                else
                {
                    var gradStr = string.Join("-", stops);
                    var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
                    if (linear?.Angle?.HasValue == true)
                    {
                        var bgDeg = linear.Angle.Value / 60000.0;
                        gradStr += bgDeg % 1 == 0 ? $"-{(int)bgDeg}" : $"-{bgDeg:0.##}";
                    }
                    node.Format["background"] = gradStr;
                }
            }
        }
        else if (blipFill != null)
        {
            node.Format["background"] = "image";

            var blip = blipFill.GetFirstChild<Drawing.Blip>();
            var alphaMod = blip?.GetFirstChild<Drawing.AlphaModulationFixed>();
            if (alphaMod?.Amount?.HasValue == true)
            {
                // amt is 0..100000 (100000 = opaque). Expose as 0..100.
                var amt = alphaMod.Amount.Value;
                node.Format["background.alpha"] = (int)Math.Round(amt / 1000.0);
            }

            var tile = blipFill.GetFirstChild<Drawing.Tile>();
            if (tile != null)
            {
                // LibreOffice convention: algn=ctr + sx=sy=100000 → "center",
                // anything else with tile → "tile".
                var algn = tile.Alignment?.Value;
                var sx = tile.HorizontalRatio?.Value ?? 100000;
                if (algn == Drawing.RectangleAlignmentValues.Center && sx == 100000)
                {
                    node.Format["background.mode"] = "center";
                }
                else
                {
                    node.Format["background.mode"] = "tile";
                    if (sx != 100000)
                        node.Format["background.scale"] = (int)Math.Round(sx / 1000.0);
                }
            }
            // Stretch is the default; only emit background.mode when non-default.

            // Surface srcRect crop bounds (1000ths of a percent) so third-party cropped
            // image backgrounds show up on get. Any side with a non-zero inset qualifies.
            var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
            if (srcRect != null)
            {
                var l = srcRect.Left?.Value ?? 0;
                var t = srcRect.Top?.Value ?? 0;
                var r = srcRect.Right?.Value ?? 0;
                var b = srcRect.Bottom?.Value ?? 0;
                if ((l | t | r | b) != 0)
                    node.Format["background.crop"] = $"{l},{t},{r},{b}";
            }
        }
    }

    // ==================== Helpers ====================

    /// <summary>
    /// Normalize alternative gradient formats to the canonical "-" separated form.
    /// Handles: "LINEAR;C1;C2;angle" → "C1-C2-angle", "RADIAL;C1;C2" → "radial:C1-C2"
    /// </summary>
    private static string NormalizeGradientValue(string value)
    {
        // Detect semicolon-separated format: TYPE;C1;C2[;angle/focus]
        if (!value.Contains(';')) return value;

        var parts = value.Split(';');
        if (parts.Length < 3) return value;

        var type = parts[0].Trim().ToUpperInvariant();
        var colorAndParams = parts.Skip(1).Select(p => p.Trim()).ToArray();

        // Dash is the separator in the canonical form, so a trailing signed angle
        // (e.g. "LINEAR;C1;C2;-90" or "LINEAR;C1;C2;+45") would splice into "C1-C2--90"
        // / "C1-C2-+45" and fail as an empty color token. Normalize a trailing signed
        // integer to its unsigned canonical form so the advertised semicolon syntax
        // stays usable.
        // Only linear form has a trailing angle; radial/path have a focus keyword, so
        // don't touch their trailing token — a trailing integer there is a color stop,
        // not an angle, and wrapping it would fabricate a fake color.
        if (colorAndParams.Length >= 2 && type == "LINEAR")
        {
            var tail = colorAndParams[^1];
            if (int.TryParse(tail, out var angleDeg) && angleDeg >= -360 && angleDeg <= 360
                && (tail.StartsWith('-') || tail.StartsWith('+')))
                colorAndParams[^1] = (((angleDeg % 360) + 360) % 360).ToString();
        }

        return type switch
        {
            "LINEAR" => string.Join("-", colorAndParams),
            "RADIAL" => "radial:" + string.Join("-", colorAndParams),
            "PATH" => "path:" + string.Join("-", colorAndParams),
            _ => value // unknown type, leave as-is
        };
    }

    /// <summary>
    /// Returns true if value looks like a gradient color string ("RRGGBB-RRGGBB[-angle]").
    /// </summary>
    private static bool IsGradientColorString(string value)
    {
        // The radial:/path: prefix is itself the gradient marker — don't second-guess
        // the color forms (hex/scheme/8-hex) here; BuildGradientFill validates them.
        var v = value;
        if (v.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
            return v.Length > 7;
        if (v.StartsWith("path:", StringComparison.OrdinalIgnoreCase))
            return v.Length > 5;

        var parts = v.Split('-');
        return parts.Length >= 2 && IsHexColorString(parts[0]);
    }

    private static bool IsHexColorString(string s)
    {
        s = s.TrimStart('#');
        // Strip @position suffix used for gradient stops (e.g. "FF0000@50").
        var at = s.IndexOf('@');
        if (at >= 0) s = s[..at];
        // Accept 3-digit shorthand (parity with SanitizeColorForOoxml) alongside
        // the canonical 6/8-digit forms so gradients can mix "F00-00F" consistently
        // with the solid-bg path.
        return (s.Length == 3 || s.Length == 6 || s.Length == 8) &&
               s.All(c => char.IsAsciiHexDigit(c));
    }

    /// <summary>
    /// Build a GradientFill element from a color string.
    /// Shared by both shape gradient and slide background gradient.
    ///
    /// Linear:  "C1-C2", "C1-C2-angle", "C1-C2-C3[-angle]"
    /// Radial:  "radial:C1-C2", "radial:C1-C2-tl" (focus: tl/tr/bl/br/center)
    /// Path:    "path:C1-C2", "path:C1-C2-tl"
    /// </summary>
    internal static Drawing.GradientFill BuildGradientFill(string value)
    {
        // Check for radial/path prefix
        string? gradientType = null;
        string colorSpec = value;

        if (value.StartsWith("radial:", StringComparison.OrdinalIgnoreCase))
        {
            gradientType = "radial";
            colorSpec = value[7..];
        }
        else if (value.StartsWith("path:", StringComparison.OrdinalIgnoreCase))
        {
            gradientType = "path";
            colorSpec = value[5..];
        }

        var parts = colorSpec.Split('-');
        // R10: Tolerate single-color gradient at the parser front-end too.
        // aaae88bf added duplicate-on-empty fallback after angle/focus stripping,
        // but this earlier guard rejected `gradient=FF0000` outright before that
        // code could run. Treating empty input as a hard error is still correct.
        if (parts.Length == 0 || (parts.Length == 1 && string.IsNullOrWhiteSpace(parts[0])))
            throw new ArgumentException(
                "Gradient requires at least one color, e.g. FF0000 or FF0000-0000FF");

        var colorParts = parts.ToList();
        string? focusPoint = null;
        int angle = 5400000; // default 90° = top→bottom

        if (gradientType != null)
        {
            // For radial/path: last segment may be a focus keyword (tl/tr/bl/br/center)
            var last = colorParts.Last().ToLowerInvariant();
            if (last is "tl" or "tr" or "bl" or "br" or "center" or "c")
            {
                focusPoint = last;
                colorParts.RemoveAt(colorParts.Count - 1);
            }
        }
        else
        {
            // For linear: last segment is angle if it's a short integer (with optional "deg" suffix)
            var lastPart = colorParts.Last();
            var angleCandidate = lastPart.EndsWith("deg", StringComparison.OrdinalIgnoreCase)
                ? lastPart[..^3] : lastPart;
            // "deg" suffix is an angle even if out of range — always strip it.
            var hasDegSuffix = lastPart.EndsWith("deg", StringComparison.OrdinalIgnoreCase);
            if (colorParts.Count >= 2 &&
                int.TryParse(angleCandidate, out var angleDeg) &&
                (hasDegSuffix || angleCandidate.Length <= 4))
            {
                // OOXML a:lin/@ang range is [0, 21600000) in 60000ths of a degree.
                // Accept only [-360, 360] — anything outside is almost certainly a
                // user typo; mod-wrapping would silently bake in a different fill.
                if (angleDeg < -360 || angleDeg > 360)
                    throw new ArgumentException(
                        $"gradient angle must be in [-360, 360], got {angleDeg}");
                angleDeg = ((angleDeg % 360) + 360) % 360;
                angle = angleDeg * 60000;
                colorParts.RemoveAt(colorParts.Count - 1);
            }
        }

        // R24-2: if only one color remains after removing angle/focus, tolerate
        // it by duplicating the color — the result is a visually solid fill
        // shaped as a 2-stop gradient. Throwing here was a user-facing crash
        // reachable from `Set` (e.g. gradient="FF0000:45" / "FF0000-90") and
        // surprised callers who expected lenient parsing.
        if (colorParts.Count == 1)
        {
            colorParts.Add(colorParts[0]);
        }

        var gradFill = new Drawing.GradientFill();
        var gsLst = new Drawing.GradientStopList();

        for (int i = 0; i < colorParts.Count; i++)
        {
            var cp = colorParts[i];
            int pos;
            var atIdx = cp.IndexOf('@');
            if (atIdx >= 0 && int.TryParse(cp[(atIdx + 1)..], out var pct))
            {
                pos = Math.Clamp(pct, 0, 100) * 1000;
                cp = cp[..atIdx];
            }
            else
            {
                pos = colorParts.Count == 1
                    ? 0
                    : (int)((long)i * 100000 / (colorParts.Count - 1));
            }
            var gs = new Drawing.GradientStop { Position = pos };
            gs.AppendChild(BuildColorElement(cp));
            gsLst.AppendChild(gs);
        }

        gradFill.AppendChild(gsLst);

        if (gradientType is "radial" or "path")
        {
            // Build path gradient fill with fillToRect controlling the focal point
            var (l, t, r, b) = (focusPoint ?? "center") switch
            {
                "tl" => (0, 0, 100000, 100000),       // top-left focal point
                "tr" => (100000, 0, 0, 100000),        // top-right
                "bl" => (0, 100000, 100000, 0),        // bottom-left
                "br" => (100000, 100000, 0, 0),        // bottom-right
                _ => (50000, 50000, 50000, 50000)       // center
            };

            // radial: → circular PathShade, path: → shape-following PathShade. Without
            // this split the two prefixes produce byte-identical XML, so path: used to
            // read back as radial:.
            var shadeKind = gradientType == "path"
                ? Drawing.PathShadeValues.Shape
                : Drawing.PathShadeValues.Circle;
            var pathFill = new Drawing.PathGradientFill { Path = shadeKind };
            pathFill.AppendChild(new Drawing.FillToRectangle
            {
                Left = l, Top = t, Right = r, Bottom = b
            });
            gradFill.AppendChild(pathFill);
        }
        else
        {
            gradFill.AppendChild(new Drawing.LinearGradientFill { Angle = angle, Scaled = true });
        }

        return gradFill;
    }
}
