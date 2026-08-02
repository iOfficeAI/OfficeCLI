// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{

    /// <summary>
    /// Build an XDR BlipFill with an optional asvg:svgBlip extension when
    /// the caller wires in an SVG image part. Keeps Add/Set picture paths
    /// free of inline extension boilerplate.
    /// </summary>
    /// <summary>
    /// CONSISTENCY(shape-line): build a shape <a:ln> outline from the compound
    /// 'color[:width[:style]]' form (mirrors the pptx shape line grammar and
    /// the xlsx Set path). 'none' → NoFill outline. width in points; style is
    /// a prstDash token (solid/dash/dot/dashdot/longdash). Single-place OOXML
    /// mutation so Add and Set stay identical.
    /// </summary>
    private static Drawing.Outline BuildShapeOutline(string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
            return new Drawing.Outline(new Drawing.NoFill());

        var parts = value.Split(':');
        var outline = new Drawing.Outline(DrawingColorBuilder.BuildSolidFill(parts[0]));
        if (parts.Length > 1
            && double.TryParse(parts[1], System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var wpt))
        {
            outline.Width = (int)Math.Round(wpt * EmuConverter.EmuPerPoint);
        }
        if (parts.Length > 2)
        {
            var dash = parts[2].ToLowerInvariant() switch
            {
                "dash" => Drawing.PresetLineDashValues.Dash,
                "dot" => Drawing.PresetLineDashValues.Dot,
                "dashdot" => Drawing.PresetLineDashValues.DashDot,
                "longdash" => Drawing.PresetLineDashValues.LargeDash,
                "solid" => Drawing.PresetLineDashValues.Solid,
                _ => (Drawing.PresetLineDashValues?)null
            };
            if (dash != null)
                outline.AppendChild(new Drawing.PresetDash { Val = dash });
        }
        return outline;
    }

    private static XDR.BlipFill BuildPictureBlipFill(string pngRelId, string? svgRelId)
        => BuildPictureBlipFill(pngRelId, svgRelId, null);

    private static XDR.BlipFill BuildPictureBlipFill(
        string pngRelId, string? svgRelId, Dictionary<string, string>? properties)
    {
        var blip = new Drawing.Blip { Embed = pngRelId };
        // P6: opacity → <a:alphaModFix amt="N"/> (0..100000 scale).
        // Accept percent (50, "50%") or fraction (0.5). 100/100%/1.0 → opaque (no node).
        if (properties != null
            && properties.TryGetValue("opacity", out var opRaw)
            && !string.IsNullOrWhiteSpace(opRaw))
        {
            var amt = ParseOpacityAmt(opRaw);
            if (amt.HasValue && amt.Value < 100000)
                blip.AppendChild(new Drawing.AlphaModulationFixed { Amount = amt.Value });
        }
        if (!string.IsNullOrEmpty(svgRelId))
            OfficeCli.Core.SvgImageHelper.AppendSvgExtension(blip, svgRelId);
        var blipFill = new XDR.BlipFill(blip);
        // P7: crop.l/r/t/b or srcRect=l=..,r=..,t=..,b=.. → <a:srcRect .../>
        // Values are percent (10 → 10000 in 1/1000 pct units). Emitted before <a:stretch>.
        var srcRect = ParseSrcRect(properties);
        if (srcRect != null)
            blipFill.AppendChild(srcRect);
        blipFill.AppendChild(new Drawing.Stretch(new Drawing.FillRectangle()));
        return blipFill;
    }

    // Parse crop.l/r/t/b (percent, 10 → 10000) and compound srcRect="l=10,r=10,..."
    // alias. Returns null when no crop props are present.
    internal static Drawing.SourceRectangle? ParseSrcRect(Dictionary<string, string>? properties)
    {
        if (properties == null) return null;
        int? l = null, r = null, t = null, b = null;
        if (properties.TryGetValue("srcRect", out var compound) && !string.IsNullOrWhiteSpace(compound))
        {
            // Track whether any piece parsed so we can throw a clear error
            // instead of silently no-oping (which would also wipe existing
            // srcRect because the caller replaces with ParseSrcRect's null).
            bool anyParsed = false;
            foreach (var piece in compound.Split(',', StringSplitOptions.RemoveEmptyEntries))
            {
                var kv = piece.Split('=', 2);
                if (kv.Length != 2) continue;
                var key = kv[0].Trim().ToLowerInvariant();
                var val = ParseCropPercent(kv[1]);
                if (!val.HasValue) continue;
                switch (key) { case "l": l = val; break; case "r": r = val; break; case "t": t = val; break; case "b": b = val; break; }
                anyParsed = true;
            }
            if (!anyParsed)
                throw new ArgumentException(
                    $"Invalid srcRect '{compound}'. Expected 'l=10,r=10,t=5,b=5' (any subset; values are percent 0-100). "
                    + "For raw l/t/r/b numbers use cropLeft/cropTop/cropRight/cropBottom keys.");
        }
        // CONSISTENCY(picture-crop): bare composite `crop=l,t,r,b` — the exact
        // form Get emits (and pptx Add already accepts). Without it, dump→batch
        // replay warned UNSUPPORTED and silently dropped the srcRect.
        if (properties.TryGetValue("crop", out var cropAll) && !string.IsNullOrWhiteSpace(cropAll)
            && !cropAll.Contains('='))
        {
            var cropParts = cropAll.Split(',');
            var cropVals = cropParts.Length == 4
                ? cropParts.Select(ParseCropPercent).ToArray()
                : null;
            if (cropVals == null || !cropVals.All(v => v.HasValue))
                throw new ArgumentException(
                    $"Invalid crop '{cropAll}'. Expected four comma-separated percentages in l,t,r,b order (e.g. '10,15,5,20').");
            l = cropVals[0]; t = cropVals[1]; r = cropVals[2]; b = cropVals[3];
        }
        foreach (var (key, fld) in new[] { ("crop.l", "l"), ("crop.r", "r"), ("crop.t", "t"), ("crop.b", "b") })
        {
            if (properties.TryGetValue(key, out var vs) && !string.IsNullOrWhiteSpace(vs))
            {
                var v = ParseCropPercent(vs);
                if (!v.HasValue) continue;
                switch (fld) { case "l": l = v; break; case "r": r = v; break; case "t": t = v; break; case "b": b = v; break; }
            }
        }
        // CONSISTENCY(picture-crop): Office-API-style `cropLeft`/`cropRight`
        // /`cropTop`/`cropBottom` aliases. Accept fraction (<=1 → *100%) or
        // percent (>1 → as-is); e.g. `cropLeft=0.1` and `cropLeft=10` both
        // mean 10% crop from left.
        foreach (var (key, fld) in new[] { ("cropLeft", "l"), ("cropRight", "r"), ("cropTop", "t"), ("cropBottom", "b") })
        {
            if (properties.TryGetValue(key, out var vs) && !string.IsNullOrWhiteSpace(vs))
            {
                var v = ParseCropFractionOrPercent(vs);
                if (!v.HasValue) continue;
                switch (fld) { case "l": l = v; break; case "r": r = v; break; case "t": t = v; break; case "b": b = v; break; }
            }
        }
        if (l == null && r == null && t == null && b == null) return null;
        var sr = new Drawing.SourceRectangle();
        if (l.HasValue) sr.Left = l.Value;
        if (r.HasValue) sr.Right = r.Value;
        if (t.HasValue) sr.Top = t.Value;
        if (b.HasValue) sr.Bottom = b.Value;
        return sr;
    }

    private static int? ParseCropPercent(string raw)
    {
        var t = raw.Trim();
        if (t.EndsWith("%")) t = t[..^1].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (double.IsNaN(v) || double.IsInfinity(v)) return null;
        return (int)Math.Round(v * 1000.0);
    }

    // CONSISTENCY(picture-crop): For `cropLeft`/`cropRight`/`cropTop`/
    // `cropBottom` keys we treat input ambiguously: <=1 is a fraction
    // (0.1 → 10%), >1 is percent (10 → 10%). Trailing `%` is still
    // honored explicitly. Returns 1/1000 pct units, same as OOXML.
    private static int? ParseCropFractionOrPercent(string raw)
    {
        var t = raw.Trim();
        bool explicitPct = t.EndsWith("%");
        if (explicitPct) t = t[..^1].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (double.IsNaN(v) || double.IsInfinity(v)) return null;
        double pct = (!explicitPct && v > 0 && v <= 1.0) ? v * 100.0 : v;
        return (int)Math.Round(pct * 1000.0);
    }

    // Parse opacity percent/fraction to OOXML alphaModFix amt scale (0..100000).
    // Returns null if the input is not parseable; 100000 (fully opaque) is returned
    // as-is so the caller can decide to omit the node.
    internal static int? ParseOpacityAmt(string raw)
    {
        var t = raw.Trim();
        if (t.EndsWith("%")) t = t[..^1].Trim();
        if (!double.TryParse(t, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var v))
            return null;
        if (double.IsNaN(v) || double.IsInfinity(v)) return null;
        // Fraction form (0..1) → treat as 0..100%; else percent.
        double pct = v <= 1.0 && v > 0 ? v * 100.0 : v;
        if (pct < 0) pct = 0; if (pct > 100) pct = 100;
        return (int)Math.Round(pct * 1000.0);
    }

    internal static XDR.Picture BuildPictureElementWithTransform(
        uint picId, string alt, string imgRelId, string? svgRelId,
        Dictionary<string, string> properties)
    {
        var xfrm = new Drawing.Transform2D(
            new Drawing.Offset { X = 0, Y = 0 },
            new Drawing.Extents { Cx = 0, Cy = 0 });
        ApplyTransform2DRotationFlip(xfrm, properties);
        // P13: accept user-supplied `name=` to override the auto-generated
        // "Picture {id}" label stamped into xdr:cNvPr @name.
        // P9: `altText=` alias for `alt=` (Description attribute).
        // P11: `title=` populates the OOXML @title attribute (distinct from alt).
        var picName = properties.GetValueOrDefault("name");
        if (string.IsNullOrWhiteSpace(picName))
            picName = $"Picture {picId}";
        var picTitle = properties.GetValueOrDefault("title");
        var cNvPr = new XDR.NonVisualDrawingProperties { Id = picId, Name = picName, Description = alt };
        if (!string.IsNullOrWhiteSpace(picTitle))
            cNvPr.Title = picTitle;
        return new XDR.Picture(
            new XDR.NonVisualPictureProperties(
                cNvPr,
                new XDR.NonVisualPictureDrawingProperties(new Drawing.PictureLocks { NoChangeAspect = true })
            ),
            BuildPictureBlipFill(imgRelId, svgRelId, properties),
            new XDR.ShapeProperties(
                xfrm,
                new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = Drawing.ShapeTypeValues.Rectangle }
            )
        );
    }

    // Apply `rotation=<deg>` / `flip=h|v|both|hv|vh` from the user properties
    // dict to a Drawing.Transform2D node. Silently no-op on missing props.
    // Mirrors PowerPointHandler's shape rotation semantics: angles are in
    // degrees (positive = clockwise), OOXML stores them as 60000ths of a
    // degree in the `rot` attribute. Values are normalized modulo 360.
    internal static void ApplyTransform2DRotationFlip(
        Drawing.Transform2D xfrm, Dictionary<string, string> properties)
    {
        if (xfrm == null) return;
        if (properties.TryGetValue("rotation", out var rotStr) && !string.IsNullOrWhiteSpace(rotStr))
        {
            if (double.TryParse(rotStr, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var deg))
            {
                var normalized = ((deg % 360) + 360) % 360;
                xfrm.Rotation = (int)Math.Round(normalized * 60000);
            }
        }
        if (properties.TryGetValue("flip", out var flipStr) && !string.IsNullOrWhiteSpace(flipStr))
        {
            var f = flipStr.Trim().ToLowerInvariant();
            bool flipH = f == "h" || f == "horizontal" || f == "both" || f == "hv" || f == "vh";
            bool flipV = f == "v" || f == "vertical" || f == "both" || f == "hv" || f == "vh";
            if (flipH) xfrm.HorizontalFlip = true;
            if (flipV) xfrm.VerticalFlip = true;
        }
        // CONSISTENCY(shape-flip): accept Office-API-style `flipH=true`,
        // `flipV=true`, `flipBoth=true` aliases in addition to the compact
        // `flip=h|v|both`. Boolean semantics follow IsTruthy (true/1/yes).
        if (properties.TryGetValue("flipH", out var flipHStr) && IsTruthy(flipHStr))
            xfrm.HorizontalFlip = true;
        if (properties.TryGetValue("flipV", out var flipVStr) && IsTruthy(flipVStr))
            xfrm.VerticalFlip = true;
        if (properties.TryGetValue("flipBoth", out var flipBothStr) && IsTruthy(flipBothStr))
        {
            xfrm.HorizontalFlip = true;
            xfrm.VerticalFlip = true;
        }
    }

    // SH6 — build a two/three-stop linear gradient fill for shape/textbox from
    // a "C1-C2[-C3][:angle]" spec. Mirrors the chart gradient parser used by
    // Core/Chart/ChartHelper.Builder.cs:BuildFillElement so chart and shape
    // gradient syntax stay consistent.
    internal static Drawing.GradientFill BuildShapeGradientFill(string spec)
    {
        var colonIdx = spec.LastIndexOf(':');
        var anglePart = 0;
        string colorsPart;
        if (colonIdx > 6 && int.TryParse(spec[(colonIdx + 1)..],
            System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var ang))
        {
            anglePart = ang;
            colorsPart = spec[..colonIdx];
        }
        else
        {
            colorsPart = spec;
        }
        var colors = colorsPart.Split('-').Select(c => c.Trim()).Where(c => c.Length > 0).ToArray();
        if (colors.Length < 2)
            throw new ArgumentException(
                $"gradientFill requires at least two '-' separated colors; got '{spec}'.");
        var gradFill = new Drawing.GradientFill { RotateWithShape = true };
        var gsLst = new Drawing.GradientStopList();
        for (int i = 0; i < colors.Length; i++)
        {
            var pos = (int)(i * 100000.0 / (colors.Length - 1));
            var (rgb, _) = ParseHelpers.SanitizeColorForOoxml(colors[i]);
            var gs = new Drawing.GradientStop { Position = pos };
            gs.AppendChild(new Drawing.RgbColorModelHex { Val = rgb });
            gsLst.AppendChild(gs);
        }
        gradFill.AppendChild(gsLst);
        gradFill.AppendChild(new Drawing.LinearGradientFill
        {
            Angle = ParseHelpers.GradientAngleToOoxmlUnits(anglePart),
            Scaled = true
        });
        return gradFill;
    }

    // ==================== Picture Helpers ====================

    // Pictures can hang off any of the three anchor kinds AddPicture emits
    // (twoCellAnchor / oneCellAnchor / absoluteAnchor). Enumerate in document
    // order so picture[N] indexing is stable across mixed anchor kinds — every
    // consumer (Get / Query / Set / Remove / dump) must use this enumerator or
    // their index spaces diverge and oneCell/absolute pictures silently vanish
    // from dump output.
    internal static IEnumerable<OpenXmlCompositeElement> EnumeratePictureAnchors(XDR.WorksheetDrawing wsDrawing)
    {
        foreach (var child in wsDrawing.ChildElements)
        {
            if (child is XDR.TwoCellAnchor or XDR.OneCellAnchor or XDR.AbsoluteAnchor
                && child.Descendants<XDR.Picture>().Any())
                yield return (OpenXmlCompositeElement)child;
        }
    }

    private DocumentNode? GetPictureNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null) return null;

        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return null;

        var picAnchors = EnumeratePictureAnchors(wsDrawing).ToList();

        if (index < 1 || index > picAnchors.Count)
            return null;

        var anchor = picAnchors[index - 1];
        var picture = anchor.Descendants<XDR.Picture>().First();

        var node = new DocumentNode { Path = path, Type = "picture" };

        var nvProps = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        if (nvProps != null)
        {
            if (!string.IsNullOrEmpty(nvProps.Description?.Value))
            {
                node.Format["alt"] = nvProps.Description.Value;
                node.Text = nvProps.Description.Value;
            }
            if (!string.IsNullOrEmpty(nvProps.Name?.Value))
                node.Format["name"] = nvProps.Name.Value;
            // P11 readback — Add writes `title=` into cNvPr @title; without
            // this the key was write-only and dump silently dropped it.
            if (!string.IsNullOrEmpty(nvProps.Title?.Value))
                node.Format["title"] = nvProps.Title.Value;
            // P8 readback — Add writes a picture hyperlink as <a:hlinkClick>
            // under cNvPr (external via a DrawingsPart relationship, or
            // internal via @location). Without this the hyperlink was
            // write-only and dropped on Get/dump.
            var picHlink = nvProps.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (picHlink != null)
            {
                if (!string.IsNullOrEmpty(picHlink.Id?.Value))
                {
                    var rel = worksheetPart.DrawingsPart?.HyperlinkRelationships
                        .FirstOrDefault(r => r.Id == picHlink.Id.Value);
                    if (rel != null) node.Format["hyperlink"] = rel.Uri.ToString();
                }
                else
                {
                    var loc = picHlink.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == "location").Value;
                    if (!string.IsNullOrEmpty(loc)) node.Format["hyperlink"] = "#" + loc;
                }
            }
        }

        ReadAnchorPosition(anchor, node);

        // Rotation / flip readback from <xdr:spPr><a:xfrm rot=".." flipH=".." flipV="..">
        // CONSISTENCY(shape-flip): same canonical form as GetShapeNode.
        var picSpPr = picture.ShapeProperties;
        var xfrm = picSpPr?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Rotation?.HasValue == true && xfrm.Rotation.Value != 0)
            {
                var deg = xfrm.Rotation.Value / 60000.0;
                node.Format["rotation"] = Math.Round(deg, 2);
            }
            if (xfrm.HorizontalFlip?.HasValue == true && xfrm.VerticalFlip?.HasValue == true
                && xfrm.HorizontalFlip.Value && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "both";
            else if (xfrm.HorizontalFlip?.HasValue == true && xfrm.HorizontalFlip.Value)
                node.Format["flip"] = "h";
            else if (xfrm.VerticalFlip?.HasValue == true && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "v";
        }

        // CONSISTENCY(picture-crop): mirror PowerPointHandler.NodeBuilder.cs
        // crop readback. <a:srcRect l/t/r/b> stores values in 1000ths of a
        // percent (10000 = 10%); emit as comma-separated percent string.
        var picSrcRect = picture.BlipFill?.GetFirstChild<Drawing.SourceRectangle>();
        if (picSrcRect != null)
        {
            var cl = picSrcRect.Left?.Value ?? 0;
            var ct = picSrcRect.Top?.Value ?? 0;
            var cr = picSrcRect.Right?.Value ?? 0;
            var cb = picSrcRect.Bottom?.Value ?? 0;
            if (cl != 0 || ct != 0 || cr != 0 || cb != 0)
                node.Format["crop"] = $"{cl / 1000.0:0.##},{ct / 1000.0:0.##},{cr / 1000.0:0.##},{cb / 1000.0:0.##}";
        }

        // P6 readback: opacity from <a:blip><a:alphaModFix amt="N"/> (0..100000
        // scale; Add takes percent). Mirrors BuildPictureBlipFill. Omitted when
        // fully opaque (no alphaModFix node).
        var picBlip = picture.BlipFill?.GetFirstChild<Drawing.Blip>();
        var picAlpha = picBlip?.GetFirstChild<Drawing.AlphaModulationFixed>();
        if (picAlpha?.Amount?.Value is { } alphaAmt && alphaAmt < 100000)
            node.Format["opacity"] = (int)Math.Round(alphaAmt / 1000.0);

        // P10 readback: decorative flag from <xdr:cNvPr><a:extLst><a:ext
        // uri="{FF2B5EF4-...}"><a16:decorative val="1"/>. Mirrors the Add
        // emit; the a16:decorative node is an unknown element.
        var picCNvPrRead = picture.NonVisualPictureProperties?.NonVisualDrawingProperties;
        if (picCNvPrRead != null)
        {
            // Search the whole cNvPr subtree — the a:extLst under cNvPr is a
            // distinct strongly-typed child, so a GetFirstChild<ExtensionList>
            // misses it; Descendants() reaches the unknown a16:decorative node.
            bool decorative = picCNvPrRead.Descendants()
                .Any(e => e.LocalName == "decorative"
                    && e.GetAttributes().Any(a => a.LocalName == "val" && a.Value == "1"));
            if (decorative) node.Format["decorative"] = true;
        }

        return node;
    }

    // CONSISTENCY(xlsx-group-flatten): enumerate every leaf XDR.Shape across
    // all TwoCellAnchors in worksheet-drawing order. Anchors that wrap a
    // GroupShape (`<xdr:grpSp>`) with multiple inner shapes contribute each
    // inner shape; non-grouped anchors contribute exactly one shape, so
    // existing files without groups see no index renumbering. Shapes anchored
    // alone (no group) and shapes inside a group are returned in document
    // order, which matches how Excel itself enumerates them on the canvas.
    internal static IEnumerable<(XDR.Shape shape, XDR.TwoCellAnchor anchor)> EnumerateLeafShapes(XDR.WorksheetDrawing wsDrawing)
    {
        foreach (var anchor in wsDrawing.Elements<XDR.TwoCellAnchor>())
        {
            foreach (var sp in anchor.Descendants<XDR.Shape>())
                yield return (sp, anchor);
        }
    }

    internal sealed class DumpDrawingHyperlinkSpec
    {
        public string Id { get; set; } = "";
        public string Target { get; set; } = "";
        public bool IsExternal { get; set; }
    }

    internal sealed class DumpShapeReplayItem
    {
        // Exactly one of ShapeIndex / GroupAnchorXml is populated.
        public int? ShapeIndex { get; set; }
        public string? GroupAnchorXml { get; set; }
        public List<DumpDrawingHyperlinkSpec> GroupHyperlinks { get; set; } = [];
        public string? Warning { get; set; }
    }

    // Compact, trimming-safe carrier used inside the dump JSON string value:
    // base64(id),base64(target),0|1 entries separated by '|'. Base64 never
    // contains ',' or '|', so no reflection-based JSON serializer is needed
    // inside the single-file published executable.
    internal static string EncodeDumpDrawingHyperlinks(
        IEnumerable<DumpDrawingHyperlinkSpec> hyperlinks)
    {
        static string B64(string value) => Convert.ToBase64String(
            System.Text.Encoding.UTF8.GetBytes(value));
        return string.Join("|", hyperlinks.Select(h =>
            $"{B64(h.Id)},{B64(h.Target)},{(h.IsExternal ? "1" : "0")}"));
    }

    internal static List<DumpDrawingHyperlinkSpec> DecodeDumpDrawingHyperlinks(
        string encoded)
    {
        static string FromB64(string value) => System.Text.Encoding.UTF8.GetString(
            Convert.FromBase64String(value));
        var result = new List<DumpDrawingHyperlinkSpec>();
        if (string.IsNullOrEmpty(encoded)) return result;
        foreach (var entry in encoded.Split('|', StringSplitOptions.RemoveEmptyEntries))
        {
            var fields = entry.Split(',');
            if (fields.Length != 3 || (fields[2] != "0" && fields[2] != "1"))
                throw new FormatException(
                    "Expected base64(id),base64(target),0|1 entries separated by '|'.");
            result.Add(new DumpDrawingHyperlinkSpec
            {
                Id = FromB64(fields[0]),
                Target = FromB64(fields[1]),
                IsExternal = fields[2] == "1",
            });
        }
        return result;
    }

    /// <summary>
    /// Build the dump replay sequence for worksheet shapes without destroying
    /// real DrawingML groups. A grouped TwoCellAnchor is carried verbatim when
    /// every relationship referenced inside it is a hyperlink relationship;
    /// those lightweight relationships can be recreated safely on replay.
    ///
    /// Groups that reference package parts (pictures/charts/etc.) fall back to
    /// the existing leaf-shape path with an explicit warning. Copying those
    /// parts requires a recursive part-graph carrier and raw XML alone would
    /// leave dangling r:embed/r:id values.
    /// </summary>
    internal List<DumpShapeReplayItem> GetDumpShapeReplayItems(string sheetName)
    {
        var result = new List<DumpShapeReplayItem>();
        var worksheet = FindWorksheet(sheetName);
        var drawingsPart = worksheet?.DrawingsPart;
        var wsDrawing = drawingsPart?.WorksheetDrawing;
        if (drawingsPart == null || wsDrawing == null) return result;

        const string relNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        int leafIndex = 0;

        foreach (var anchor in wsDrawing.Elements<XDR.TwoCellAnchor>())
        {
            var leaves = anchor.Descendants<XDR.Shape>().ToList();
            var firstLeafIndex = leafIndex + 1;
            leafIndex += leaves.Count;

            var group = anchor.GetFirstChild<XDR.GroupShape>();
            if (group == null)
            {
                for (int i = 0; i < leaves.Count; i++)
                    result.Add(new DumpShapeReplayItem { ShapeIndex = firstLeafIndex + i });
                continue;
            }

            var referencedIds = anchor
                .Descendants()
                .Prepend(anchor)
                .SelectMany(e => e.GetAttributes())
                .Where(a => a.NamespaceUri == relNs
                    && (a.LocalName == "id" || a.LocalName == "embed" || a.LocalName == "link")
                    && !string.IsNullOrEmpty(a.Value))
                .Select(a => a.Value!)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            var hyperlinks = new List<DumpDrawingHyperlinkSpec>();
            var unsupportedIds = new List<string>();
            foreach (var rid in referencedIds)
            {
                var hyperlink = drawingsPart.HyperlinkRelationships
                    .FirstOrDefault(r => r.Id == rid);
                if (hyperlink == null)
                {
                    unsupportedIds.Add(rid);
                    continue;
                }
                hyperlinks.Add(new DumpDrawingHyperlinkSpec
                {
                    Id = hyperlink.Id ?? rid,
                    Target = hyperlink.Uri.ToString(),
                    IsExternal = hyperlink.IsExternal,
                });
            }

            if (unsupportedIds.Count == 0)
            {
                result.Add(new DumpShapeReplayItem
                {
                    GroupAnchorXml = anchor.OuterXml,
                    GroupHyperlinks = hyperlinks,
                });
                continue;
            }

            var warning =
                $"grouped drawing references non-hyperlink relationship(s) {string.Join(", ", unsupportedIds)}; "
                + "the group was flattened because its dependent package parts cannot yet be carried safely";
            for (int i = 0; i < leaves.Count; i++)
            {
                result.Add(new DumpShapeReplayItem
                {
                    ShapeIndex = firstLeafIndex + i,
                    Warning = i == 0 ? warning : null,
                });
            }
        }

        return result;
    }

    private DocumentNode? GetShapeNode(string sheetName, WorksheetPart worksheetPart, int index, string path)
    {
        var drawingsPart = worksheetPart.DrawingsPart;
        if (drawingsPart == null) return null;
        var wsDrawing = drawingsPart.WorksheetDrawing;
        if (wsDrawing == null) return null;

        var shapes = EnumerateLeafShapes(wsDrawing).ToList();

        if (index < 1 || index > shapes.Count)
            return null;

        var (shape, anchor) = shapes[index - 1];

        var node = new DocumentNode { Path = path, Type = "shape" };

        // Name
        var nvProps = shape.NonVisualShapeProperties?.GetFirstChild<XDR.NonVisualDrawingProperties>();
        if (nvProps?.Name?.Value != null)
            node.Format["name"] = nvProps.Name.Value;

        // Shape hyperlinks live under <xdr:cNvPr><a:hlinkClick>. Group shapes
        // are flattened to leaf shapes for Get/dump, so preserve each leaf link.
        if (nvProps != null)
        {
            var shapeHlink = nvProps.GetFirstChild<Drawing.HyperlinkOnClick>();
            if (shapeHlink != null)
            {
                if (!string.IsNullOrEmpty(shapeHlink.Id?.Value))
                {
                    var rel = drawingsPart.HyperlinkRelationships
                        .FirstOrDefault(r => r.Id == shapeHlink.Id.Value);
                    if (rel != null) node.Format["hyperlink"] = rel.Uri.ToString();
                }
                else
                {
                    var loc = shapeHlink.GetAttributes()
                        .FirstOrDefault(a => a.LocalName == "location").Value;
                    if (!string.IsNullOrEmpty(loc)) node.Format["hyperlink"] = "#" + loc;
                }
            }
        }

        // Text — shape TextBody has one <a:p> per paragraph, each with
        // zero-or-more <a:r>/<a:t> runs. Concatenate runs within a
        // paragraph, then join paragraphs with '\n' so multi-line shape
        // text round-trips through Get.
        var paragraphs = shape.TextBody?.Elements<Drawing.Paragraph>().ToList();
        if (paragraphs != null && paragraphs.Count > 0)
        {
            node.Text = string.Join("\n", paragraphs.Select(p =>
                string.Join("", p.Elements<Drawing.Run>().Select(r => r.Text?.Text ?? ""))));
        }
        var textRuns = shape.TextBody?.Descendants<Drawing.Run>().ToList();

        // Position/size
        ReadAnchorPosition(anchor, node);

        // Font properties from first run
        var firstRun = textRuns?.FirstOrDefault();
        var rPr = firstRun?.RunProperties;
        if (rPr != null)
        {
            if (rPr.FontSize?.HasValue == true)
                node.Format["size"] = $"{rPr.FontSize.Value / 100.0}pt";
            if (rPr.Bold?.HasValue == true && rPr.Bold.Value)
                node.Format["bold"] = true;
            if (rPr.Italic?.HasValue == true && rPr.Italic.Value)
                node.Format["italic"] = true;
            if (rPr.Underline?.HasValue == true && rPr.Underline.Value != Drawing.TextUnderlineValues.None)
                node.Format["underline"] = rPr.Underline.Value == Drawing.TextUnderlineValues.Double ? "double" : "single";

            var solidFill = rPr.GetFirstChild<Drawing.SolidFill>();
            var colorHex = solidFill?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (colorHex?.Val?.Value != null)
                node.Format["color"] = ParseHelpers.FormatHexColor(colorHex.Val.Value);
            else
            {
                var schemeClr = solidFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                if (schemeClr?.HasValue == true) node.Format["color"] = schemeClr.InnerText;
            }

            var latin = rPr.GetFirstChild<Drawing.LatinFont>();
            if (latin?.Typeface?.Value != null)
                node.Format["font"] = latin.Typeface.Value;
        }

        // Rotation / flip readback from <a:xfrm rot="..." flipH="..." flipV="...">
        var xfrm = shape.ShapeProperties?.Transform2D;
        if (xfrm != null)
        {
            if (xfrm.Rotation?.HasValue == true && xfrm.Rotation.Value != 0)
            {
                // OOXML stores rotation in 60000ths of a degree; Add normalizes
                // into [0,360). Round-trip the same canonical form.
                var deg = xfrm.Rotation.Value / 60000.0;
                node.Format["rotation"] = Math.Round(deg, 2);
            }
            if (xfrm.HorizontalFlip?.HasValue == true && xfrm.VerticalFlip?.HasValue == true
                && xfrm.HorizontalFlip.Value && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "both";
            else if (xfrm.HorizontalFlip?.HasValue == true && xfrm.HorizontalFlip.Value)
                node.Format["flip"] = "h";
            else if (xfrm.VerticalFlip?.HasValue == true && xfrm.VerticalFlip.Value)
                node.Format["flip"] = "v";
        }

        // Geometry preset (rect, ellipse, etc.) — `preset` is the canonical
        // key per shape help schema; `preset`/`shape` are accepted as
        // Add/Set aliases. Aligns with PPTX shape readback (commit 9f72712a).
        var presetGeom = shape.ShapeProperties?.GetFirstChild<Drawing.PresetGeometry>();
        if (presetGeom?.Preset?.HasValue == true)
            node.Format["geometry"] = presetGeom.Preset.InnerText;

        // Fill
        var spPr = shape.ShapeProperties;
        if (spPr?.GetFirstChild<Drawing.NoFill>() != null)
            node.Format["fill"] = "none";
        else if (spPr?.GetFirstChild<Drawing.GradientFill>() is { } shapeGradFill)
        {
            // SH6 readback — reconstruct the "C1-C2[-C3][:angle]" spec that
            // BuildShapeGradientFill consumes, so dump→batch replays the
            // gradient instead of silently dropping it to the default fill.
            var stopColors = shapeGradFill.GetFirstChild<Drawing.GradientStopList>()?
                .Elements<Drawing.GradientStop>()
                .Select(gs => gs.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value)
                .Where(v => !string.IsNullOrEmpty(v))
                .ToList();
            if (stopColors is { Count: >= 2 })
            {
                var spec = string.Join("-", stopColors.Select(v => ParseHelpers.FormatHexColor(v!)));
                var linAngle = shapeGradFill.GetFirstChild<Drawing.LinearGradientFill>()?.Angle?.Value ?? 0;
                var angleDeg = (int)Math.Round(linAngle / 60000.0);
                if (angleDeg != 0) spec += $":{angleDeg}";
                node.Format["gradientFill"] = spec;
            }
        }
        else
        {
            var shapeFill = spPr?.GetFirstChild<Drawing.SolidFill>();
            var fillColor = shapeFill?.GetFirstChild<Drawing.RgbColorModelHex>();
            if (fillColor?.Val?.Value != null)
                node.Format["fill"] = ParseHelpers.FormatHexColor(fillColor.Val.Value);
            else
            {
                var schemeClr = shapeFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                if (schemeClr?.HasValue == true) node.Format["fill"] = schemeClr.InnerText;
            }
        }

        // Paragraph alignment — read first paragraph's a:pPr/@algn (mirrors
        // Set which writes to every paragraph). PPTX shape Get uses `align`
        // canonical key.
        var firstPara = shape.TextBody?.GetFirstChild<Drawing.Paragraph>();
        var firstPPr = firstPara?.ParagraphProperties;
        if (firstPPr?.Alignment?.HasValue == true)
        {
            // SDK v3 enum values are not compile-time constants; switch on InnerText.
            node.Format["align"] = firstPPr.Alignment.InnerText switch
            {
                "ctr" => "center",
                "r" => "right",
                "just" => "justify",
                "l" => "left",
                var s => s,
            };
        }

        // Vertical alignment — bodyPr/@anchor.
        var bodyPrForAnchor = shape.TextBody?.GetFirstChild<Drawing.BodyProperties>();
        if (bodyPrForAnchor?.Anchor?.HasValue == true)
        {
            node.Format["valign"] = bodyPrForAnchor.Anchor.InnerText switch
            {
                "t" => "top",
                "ctr" => "center",
                "b" => "bottom",
                var s => s,
            };
        }

        // Outline (line/border). Set writes "none" or "color[:width[:style]]".
        // Round-trip emits the same canonical form.
        var outline = spPr?.GetFirstChild<Drawing.Outline>();
        if (outline != null)
        {
            if (outline.GetFirstChild<Drawing.NoFill>() != null)
                node.Format["line"] = "none";
            else
            {
                var lineFill = outline.GetFirstChild<Drawing.SolidFill>();
                var lineRgb = lineFill?.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value;
                string? colorPart = null;
                if (lineRgb != null)
                    colorPart = ParseHelpers.FormatHexColor(lineRgb);
                else
                {
                    var schemeClr = lineFill?.GetFirstChild<Drawing.SchemeColor>()?.Val;
                    if (schemeClr?.HasValue == true) colorPart = schemeClr.InnerText;
                }
                if (colorPart != null)
                {
                    var widthPt = outline.Width?.HasValue == true
                        ? $":{outline.Width.Value / EmuConverter.EmuPerPointF:0.##}"
                        : "";
                    node.Format["line"] = colorPart + widthPt;
                }
            }
        }

        // Margin (text body insets) — Add/Set accept points and write all four
        // sides uniformly; mirror that as a single points readback when all
        // four match. Stored as EMU on BodyProperties, 12700 EMU per point.
        var bodyPr = shape.TextBody?.GetFirstChild<Drawing.BodyProperties>();
        if (bodyPr != null)
        {
            var lIns = bodyPr.LeftInset?.Value;
            var rIns = bodyPr.RightInset?.Value;
            var tIns = bodyPr.TopInset?.Value;
            var bIns = bodyPr.BottomInset?.Value;
            if (lIns.HasValue || rIns.HasValue || tIns.HasValue || bIns.HasValue)
            {
                if (lIns == rIns && rIns == tIns && tIns == bIns && lIns.HasValue)
                    node.Format["margin"] = $"{lIns.Value / EmuConverter.EmuPerPointF:0.##}pt";
                else
                    node.Format["margin"] = $"{(lIns ?? 0) / EmuConverter.EmuPerPointF:0.##}pt,{(tIns ?? 0) / EmuConverter.EmuPerPointF:0.##}pt,{(rIns ?? 0) / EmuConverter.EmuPerPointF:0.##}pt,{(bIns ?? 0) / EmuConverter.EmuPerPointF:0.##}pt";
            }
        }

        // Effects — check shape-level then text-level
        var effectList = spPr?.GetFirstChild<Drawing.EffectList>();
        var textEffectList = (effectList == null || !effectList.HasChildren)
            ? rPr?.GetFirstChild<Drawing.EffectList>()
            : null;
        var activeEffects = effectList?.HasChildren == true ? effectList : textEffectList;
        if (activeEffects != null)
        {
            var shadow = activeEffects.GetFirstChild<Drawing.OuterShadow>();
            if (shadow != null)
            {
                var sColor = ParseHelpers.FormatHexColor(shadow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000");
                node.Format["shadow"] = sColor;
            }
            var glow = activeEffects.GetFirstChild<Drawing.Glow>();
            if (glow != null)
            {
                var gColor = ParseHelpers.FormatHexColor(glow.GetFirstChild<Drawing.RgbColorModelHex>()?.Val?.Value ?? "000000");
                var gRadius = glow.Radius?.HasValue == true ? $"{glow.Radius.Value / EmuConverter.EmuPerPointF:0.##}" : "8";
                node.Format["glow"] = $"{gColor}-{gRadius}";
            }
            // softEdge readback — xlsx Add/Set build a:softEdge (Add.Drawings.cs)
            // but Get omitted it, so the shared shape schema's get:true over-
            // reported parity. Mirror the pptx emit form ("<radius>pt") so
            // set softEdge=<value> re-parses its own readback.
            var softEdge = activeEffects.GetFirstChild<Drawing.SoftEdge>();
            if (softEdge?.Radius?.HasValue == true)
                node.Format["softEdge"] = $"{softEdge.Radius.Value / EmuConverter.EmuPerPointF:0.##}pt";
            // reflection readback — Add/Set build a:reflection (via
            // DrawingEffectsHelper.BuildReflection) but Get omitted it, so the
            // Get-driven dump silently dropped it. BuildReflection encodes the
            // preset/strength in @endPos (1/1000 percent); emit the numeric
            // percent so `set reflection=<pct>` re-parses its own readback
            // (true/half/full/tight all normalize to the same endPos on replay).
            var reflection = activeEffects.GetFirstChild<Drawing.Reflection>();
            if (reflection?.EndPosition?.HasValue == true)
                node.Format["reflection"] = reflection.EndPosition.Value switch
                {
                    // Mirror PowerPointHandler.NodeBuilder reflection readback:
                    // preset names for the canonical strengths, numeric percent
                    // otherwise. Each re-parses via BuildReflection on replay.
                    55000 => "tight",
                    90000 => "half",
                    100000 => "full",
                    _ => (reflection.EndPosition.Value / 1000).ToString(System.Globalization.CultureInfo.InvariantCulture),
                };
        }

        return node;
    }

    // ==================== Shared Anchor Helpers ====================

    /// <summary>
    /// Set position/size properties (x, y, width, height) on a TwoCellAnchor.
    /// Returns true if the key was handled, false otherwise.
    /// </summary>
    // OOXML spreadsheet-drawing extents/positions (<xdr:ext>, <xdr:pos>) are
    // Int32-bounded EMU (ST_PositiveCoordinate32). Values beyond that pass
    // SDK save silently but fail schema validation (MaxInclusive) and make
    // real Excel refuse the workbook (0x800A03EC) — reject at input instead
    // of persisting an unopenable file. Only oneCell/absolute anchors hit
    // this: twoCell anchors split spans into whole-column/row markers.
    internal static long ValidateDrawingCoordEmu(long emu, string key)
    {
        if (emu > int.MaxValue)
            throw new ArgumentException(
                $"Picture/shape {key} is out of range for a oneCell/absolute drawing anchor: " +
                $"{emu} EMU exceeds the OOXML limit 2147483647 EMU (~23.5in per axis).");
        return emu;
    }

    // Anchor-kind dispatch mirroring ReadAnchorPosition: oneCell keeps x/y as
    // cell indices but sizes via <xdr:ext> EMU; absolute positions and sizes
    // are all EMU (ParseEmu accepts "2cm"/"1in"/"NNNemu"/bare EMU).
    private static bool TrySetAnchorPosition(OpenXmlCompositeElement anchorEl, string key, string value)
    {
        switch (anchorEl)
        {
            case XDR.TwoCellAnchor two:
                return TrySetAnchorPosition(two, key, value);
            case XDR.OneCellAnchor one:
                switch (key)
                {
                    case "x":
                        if (one.FromMarker?.ColumnId != null)
                            one.FromMarker.ColumnId.Text = ParseAnchorOrigin(value, "x").ToString();
                        return true;
                    case "y":
                        if (one.FromMarker?.RowId != null)
                            one.FromMarker.RowId.Text = ParseAnchorOrigin(value, "y").ToString();
                        return true;
                    case "width":
                    {
                        var ext = one.GetFirstChild<XDR.Extent>();
                        if (ext != null) ext.Cx = ValidateDrawingCoordEmu(EmuConverter.ParseEmu(value), key);
                        return true;
                    }
                    case "height":
                    {
                        var ext = one.GetFirstChild<XDR.Extent>();
                        if (ext != null) ext.Cy = ValidateDrawingCoordEmu(EmuConverter.ParseEmu(value), key);
                        return true;
                    }
                    default:
                        return false;
                }
            case XDR.AbsoluteAnchor abs:
                switch (key)
                {
                    case "x":
                    {
                        var pos = abs.GetFirstChild<XDR.Position>();
                        if (pos != null) pos.X = ValidateDrawingCoordEmu(EmuConverter.ParseEmu(value), key);
                        return true;
                    }
                    case "y":
                    {
                        var pos = abs.GetFirstChild<XDR.Position>();
                        if (pos != null) pos.Y = ValidateDrawingCoordEmu(EmuConverter.ParseEmu(value), key);
                        return true;
                    }
                    case "width":
                    {
                        var ext = abs.GetFirstChild<XDR.Extent>();
                        if (ext != null) ext.Cx = ValidateDrawingCoordEmu(EmuConverter.ParseEmu(value), key);
                        return true;
                    }
                    case "height":
                    {
                        var ext = abs.GetFirstChild<XDR.Extent>();
                        if (ext != null) ext.Cy = ValidateDrawingCoordEmu(EmuConverter.ParseEmu(value), key);
                        return true;
                    }
                    default:
                        return false;
                }
            default:
                return false;
        }
    }

    private static bool TrySetAnchorPosition(XDR.TwoCellAnchor anchor, string key, string value)
    {
        switch (key)
        {
            case "x":
                if (anchor.FromMarker != null)
                {
                    // CONSISTENCY(ole-width-units): mirror Add — accept bare
                    // cell index OR unit-qualified offset ("2cm", "1in", "72pt").
                    var xVal = ParseAnchorOrigin(value, "x");
                    anchor.FromMarker.ColumnId!.Text = xVal.ToString();
                }
                return true;
            case "y":
                if (anchor.FromMarker != null)
                {
                    // CONSISTENCY(ole-width-units): see x case above.
                    var yVal = ParseAnchorOrigin(value, "y");
                    anchor.FromMarker.RowId!.Text = yVal.ToString();
                }
                return true;
            case "width":
                if (anchor.FromMarker != null && anchor.ToMarker != null)
                {
                    // CONSISTENCY(ole-width-units): mirror Add path's
                    // ParseAnchorDimension — accept bare integer cell spans
                    // OR unit-qualified strings ("6cm", "2in", "72pt").
                    var fromCol = int.TryParse(anchor.FromMarker.ColumnId?.Text, out var fc) ? fc : 0;
                    anchor.ToMarker.ColumnId!.Text = (fromCol + ParseAnchorDimension(value, "width")).ToString();
                }
                return true;
            case "height":
                if (anchor.FromMarker != null && anchor.ToMarker != null)
                {
                    // CONSISTENCY(ole-width-units): see width case above.
                    var fromRow = int.TryParse(anchor.FromMarker.RowId?.Text, out var fr) ? fr : 0;
                    anchor.ToMarker.RowId!.Text = (fromRow + ParseAnchorDimension(value, "height")).ToString();
                }
                return true;
            default:
                return false;
        }
    }

    /// <summary>
    /// Read position/size from a TwoCellAnchor into a DocumentNode's Format dictionary.
    /// </summary>
    // Anchor-kind dispatch: oneCell/absolute anchors carry their size in an
    // <xdr:ext> (EMU) instead of a To marker, and absolute carries x/y as EMU
    // <xdr:pos>. Emit raw "NNNemu" for EMU-exact values (same exactness
    // convention as the dump emitter's width/height) plus an anchorMode key
    // (omitted for the twoCell default) so dump→batch replays the same anchor
    // kind.
    private static void ReadAnchorPosition(OpenXmlCompositeElement anchor, DocumentNode node)
    {
        switch (anchor)
        {
            case XDR.TwoCellAnchor two:
                ReadAnchorPosition(two, node);
                return;
            case XDR.OneCellAnchor one:
            {
                node.Format["anchorMode"] = "oneCell";
                var from = one.FromMarker;
                if (from != null)
                {
                    node.Format["x"] = from.ColumnId?.Text ?? "0";
                    node.Format["y"] = from.RowId?.Text ?? "0";
                }
                var ext = one.GetFirstChild<XDR.Extent>();
                if (ext?.Cx?.HasValue == true) node.Format["width"] = $"{ext.Cx.Value}emu";
                if (ext?.Cy?.HasValue == true) node.Format["height"] = $"{ext.Cy.Value}emu";
                return;
            }
            case XDR.AbsoluteAnchor abs:
            {
                node.Format["anchorMode"] = "absolute";
                var pos = abs.GetFirstChild<XDR.Position>();
                if (pos?.X?.HasValue == true) node.Format["x"] = $"{pos.X.Value}emu";
                if (pos?.Y?.HasValue == true) node.Format["y"] = $"{pos.Y.Value}emu";
                var ext = abs.GetFirstChild<XDR.Extent>();
                if (ext?.Cx?.HasValue == true) node.Format["width"] = $"{ext.Cx.Value}emu";
                if (ext?.Cy?.HasValue == true) node.Format["height"] = $"{ext.Cy.Value}emu";
                return;
            }
        }
    }

    private static void ReadAnchorPosition(XDR.TwoCellAnchor anchor, DocumentNode node)
    {
        var from = anchor.FromMarker;
        var to = anchor.ToMarker;
        if (from != null)
        {
            node.Format["x"] = from.ColumnId?.Text ?? "0";
            node.Format["y"] = from.RowId?.Text ?? "0";
        }
        if (to != null && from != null)
        {
            var fromCol = int.TryParse(from.ColumnId?.Text, out var fc) ? fc : 0;
            var toCol = int.TryParse(to.ColumnId?.Text, out var tc) ? tc : 0;
            var fromRow = int.TryParse(from.RowId?.Text, out var fr) ? fr : 0;
            var toRow = int.TryParse(to.RowId?.Text, out var tr2) ? tr2 : 0;
            node.Format["width"] = (toCol - fromCol).ToString();
            node.Format["height"] = (toRow - fromRow).ToString();
        }
    }

    /// <summary>
    /// Set rotation on a ShapeProperties element.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetRotation(XDR.ShapeProperties? spPr, string key, string value)
    {
        if (key is not ("rotation" or "rot")) return false;
        if (spPr == null) return true;

        var xfrm = spPr.GetFirstChild<Drawing.Transform2D>();
        if (xfrm == null)
        {
            xfrm = new Drawing.Transform2D(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 }
            );
            spPr.InsertAt(xfrm, 0);
        }
        if (!double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var degrees))
            throw new ArgumentException($"Invalid 'rotation' value: '{value}'. Expected a number in degrees (e.g. 45, -90, 180.5).");
        xfrm.Rotation = (int)(degrees * 60000);
        return true;
    }

    /// <summary>
    /// Set horizontal / vertical flip on a shape's Transform2D. Accepts "h", "v", "both",
    /// or "none" to clear both. Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeFlip(XDR.ShapeProperties? spPr, string key, string value)
    {
        // Accept the compact `flip=h|v|both|hv|vh|none|false` form plus the
        // Office-API aliases `flipH=true`, `flipV=true`, `flipHorizontal=true`,
        // `flipVertical=true`, `flipBoth=true`. CONSISTENCY(shape-flip) — mirrors
        // ApplyTransform2DRotationFlip used on the Add path.
        if (key is not ("flip" or "fliph" or "flipv" or "fliphorizontal" or "flipvertical" or "flipboth"))
            return false;
        if (spPr == null) return true;
        var xfrm = spPr.GetFirstChild<Drawing.Transform2D>();
        if (xfrm == null)
        {
            xfrm = new Drawing.Transform2D(
                new Drawing.Offset { X = 0, Y = 0 },
                new Drawing.Extents { Cx = 0, Cy = 0 });
            spPr.InsertAt(xfrm, 0);
        }

        if (key == "flip")
        {
            var f = value.Trim().ToLowerInvariant();
            bool none = f is "none" or "false" or "";
            bool flipH = !none && (f is "h" or "horizontal" or "both" or "hv" or "vh");
            bool flipV = !none && (f is "v" or "vertical" or "both" or "hv" or "vh");
            xfrm.HorizontalFlip = flipH ? true : (bool?)null;
            xfrm.VerticalFlip = flipV ? true : (bool?)null;
            return true;
        }

        bool truthy = IsTruthy(value);
        if (key is "fliph" or "fliphorizontal")
            xfrm.HorizontalFlip = truthy ? true : (bool?)null;
        else if (key is "flipv" or "flipvertical")
            xfrm.VerticalFlip = truthy ? true : (bool?)null;
        else if (key == "flipboth")
        {
            xfrm.HorizontalFlip = truthy ? true : (bool?)null;
            xfrm.VerticalFlip = truthy ? true : (bool?)null;
        }
        return true;
    }

    /// <summary>
    /// Apply a dotted-form font property (`font.bold`, `font.italic`, `font.color`,
    /// `font.size`, `font.name`, `font.underline`) to every run in the shape's text body.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeFontProp(XDR.Shape shape, string key, string value)
    {
        if (!key.StartsWith("font.", StringComparison.Ordinal)) return false;
        var sub = key.Substring(5);
        foreach (var run in shape.Descendants<Drawing.Run>())
        {
            var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
            switch (sub)
            {
                case "bold":
                    rPr.Bold = IsTruthy(value);
                    break;
                case "italic":
                    rPr.Italic = IsTruthy(value);
                    break;
                case "size":
                    rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(value) * 100);
                    break;
                case "name":
                    rPr.RemoveAllChildren<Drawing.LatinFont>();
                    rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                    rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                    rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                    break;
                case "color":
                {
                    rPr.RemoveAllChildren<Drawing.SolidFill>();
                    OfficeCli.Core.DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                        DrawingColorBuilder.BuildSolidFill(value));
                    break;
                }
                case "underline":
                {
                    var uv = value.ToLowerInvariant();
                    rPr.Underline = uv switch
                    {
                        "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                        "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                        "none" or "false" => Drawing.TextUnderlineValues.None,
                        _ => Drawing.TextUnderlineValues.Single
                    };
                    break;
                }
                default:
                    return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Apply shape-level effects (shadow, glow, reflection, softedge) on a ShapeProperties element.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TrySetShapeEffect(XDR.ShapeProperties? spPr, string key, string value)
    {
        if (key is not ("shadow" or "glow" or "reflection" or "softedge")) return false;
        if (spPr == null) return true;

        var effectList = spPr.GetFirstChild<Drawing.EffectList>();
        var normalizedVal = value.Replace(':', '-');
        if (normalizedVal == "true") normalizedVal = key == "shadow" ? "000000" : key == "glow" ? "4472C4" : "half";

        if (normalizedVal.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            normalizedVal.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (effectList != null)
            {
                switch (key)
                {
                    case "shadow": effectList.RemoveAllChildren<Drawing.OuterShadow>(); break;
                    case "glow": effectList.RemoveAllChildren<Drawing.Glow>(); break;
                    case "reflection": effectList.RemoveAllChildren<Drawing.Reflection>(); break;
                    case "softedge": effectList.RemoveAllChildren<Drawing.SoftEdge>(); break;
                }
                if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            }
        }
        else
        {
            if (effectList == null) { effectList = new Drawing.EffectList(); spPr.AppendChild(effectList); }
            // CONSISTENCY(effect-list-schema-order): CT_EffectList order is
            // blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge.
            // Excel (and PPT) silently drops out-of-order children, so we must
            // InsertBefore the next-in-order sibling rather than AppendChild.
            OpenXmlElement newEffect;
            switch (key)
            {
                case "shadow":
                    effectList.RemoveAllChildren<Drawing.OuterShadow>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
                    break;
                case "glow":
                    effectList.RemoveAllChildren<Drawing.Glow>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor);
                    break;
                case "reflection":
                    effectList.RemoveAllChildren<Drawing.Reflection>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildReflection(normalizedVal);
                    break;
                case "softedge":
                    effectList.RemoveAllChildren<Drawing.SoftEdge>();
                    newEffect = OfficeCli.Core.DrawingEffectsHelper.BuildSoftEdge(normalizedVal);
                    break;
                default: return true;
            }
            OfficeCli.Core.DrawingEffectsHelper.InsertEffectInSchemaOrder(effectList, newEffect);
        }
        return true;
    }

    /// <summary>
    /// Parse x, y, width, height from properties with given defaults. Used by both picture Add and shape Add.
    /// </summary>
    // CONSISTENCY(shape-preset): mirror PowerPointHandler.ParsePresetShape token
    // set so Excel `add shape preset=X` accepts the same vocabulary as PPT.
    //
    // Exhaustive map covering every OOXML preset token. Built once via
    // reflection over `Drawing.ShapeTypeValues` static properties — each
    // property's default `ToString()` (== OpenXml IEnumValue.Value) is the
    // OOXML token such as "smileyFace", "flowChartProcess", "lightningBolt".
    // We then overlay friendly aliases (oval, cylinder, rarrow, …).
    private static readonly Dictionary<string, Drawing.ShapeTypeValues> _shapePresetMap =
        BuildShapePresetMap();

    private static Dictionary<string, Drawing.ShapeTypeValues> BuildShapePresetMap()
    {
        var map = new Dictionary<string, Drawing.ShapeTypeValues>(StringComparer.Ordinal);
        foreach (var p in typeof(Drawing.ShapeTypeValues)
            .GetProperties(BindingFlags.Public | BindingFlags.Static)
            .Where(p => p.PropertyType == typeof(Drawing.ShapeTypeValues)))
        {
            if (p.GetValue(null) is not Drawing.ShapeTypeValues val) continue;
            // IEnumValue.Value is the OOXML token, e.g. "smileyFace". Do not
            // use ToString() — on OpenXml SDK 3.x record-struct wrappers it
            // returns "ShapeTypeValues { }" instead of the token.
            var token = (val as IEnumValue)?.Value;
            if (string.IsNullOrEmpty(token)) continue;
            map[token.ToLowerInvariant()] = val;
        }

        // Friendly aliases layered on top (key must be lowercase).
        void Alias(string alias, Drawing.ShapeTypeValues v) => map[alias] = v;
        Alias("rectangle", Drawing.ShapeTypeValues.Rectangle);
        Alias("roundedrectangle", Drawing.ShapeTypeValues.RoundRectangle);
        Alias("oval", Drawing.ShapeTypeValues.Ellipse);
        Alias("righttriangle", Drawing.ShapeTypeValues.RightTriangle);
        Alias("rtriangle", Drawing.ShapeTypeValues.RightTriangle);
        Alias("rarrow", Drawing.ShapeTypeValues.RightArrow);
        Alias("larrow", Drawing.ShapeTypeValues.LeftArrow);
        Alias("cross", Drawing.ShapeTypeValues.Plus);
        Alias("cylinder", Drawing.ShapeTypeValues.Can);
        return map;
    }

    /// <summary>
    /// Parse shape margin into 4 EMU insets (left, top, right, bottom).
    /// Accepts unit-qualified "14pt"/"0.5cm"/"0.2in"/bare-points for uniform
    /// inset, OR a 4-CSV "Lpt,Tpt,Rpt,Bpt" matching Get's readback format.
    /// CONSISTENCY(spacing-units): mirrors SpacingConverter usage so that
    /// margin's input vocabulary matches Get's "Npt"/"L,T,R,B" output.
    /// </summary>
    private static (int L, int T, int R, int B) ParseShapeMarginToEmu(string value)
    {
        var parts = (value ?? string.Empty).Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 4)
        {
            int Emu(string s) => (int)Math.Round(SpacingConverter.ParsePoints(s) * EmuConverter.EmuPerPoint);
            return (Emu(parts[0]), Emu(parts[1]), Emu(parts[2]), Emu(parts[3]));
        }
        if (parts.Length == 1)
        {
            var emu = (int)Math.Round(SpacingConverter.ParsePoints(parts[0]) * EmuConverter.EmuPerPoint);
            return (emu, emu, emu, emu);
        }
        throw new ArgumentException(
            $"Invalid 'margin' value '{value}'. Expected single length (e.g. '4pt', '0.5cm') or 4-CSV 'L,T,R,B'.");
    }

    private static Drawing.ShapeTypeValues ParseExcelShapePreset(string name)
    {
        var key = (name ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(key))
            return Drawing.ShapeTypeValues.Rectangle;
        if (_shapePresetMap.TryGetValue(key, out var val))
            return val;
        // R20-01: Unknown preset falls back to rectangle, but emit a stderr
        // warning so users notice (silent rect was found by audit). 'custom'
        // is the common case — it would require a custGeom path which
        // officecli doesn't expose, so suggest raw-set explicitly.
        if (key == "custom")
        {
            Console.Error.WriteLine(
                "Warning: preset='custom' requires a custGeom path which officecli does not expose; " +
                "falling back to preset='rect'. Use a 'rawset' / direct OOXML edit if you need a custom path.");
        }
        else
        {
            Console.Error.WriteLine(
                $"Warning: unknown shape preset '{name}'; falling back to preset='rect'. " +
                "Valid presets include rect, ellipse, roundRect, triangle, rightArrow, etc.");
        }
        return Drawing.ShapeTypeValues.Rectangle;
    }

    private static (int x, int y, int width, int height) ParseAnchorBounds(
        Dictionary<string, string> properties, string defX, string defY, string defW, string defH)
    {
        // CONSISTENCY(shape-h-w-alias): mirror PPTX shape Add — accept `w` as
        // alias for `width`, `h` as alias for `height`. Without this mapping,
        // ParseAnchorDimension never sees the user value and the negative-
        // number guard is silently bypassed (h=-100 left as default 3 cells).
        var widthRaw = properties.GetValueOrDefault("width")
            ?? properties.GetValueOrDefault("w")
            ?? defW;
        var heightRaw = properties.GetValueOrDefault("height")
            ?? properties.GetValueOrDefault("h")
            ?? defH;
        return (
            ParseAnchorOrigin(properties.GetValueOrDefault("x", defX) ?? defX, "x"),
            ParseAnchorOrigin(properties.GetValueOrDefault("y", defY) ?? defY, "y"),
            ParseAnchorDimension(widthRaw, "width"),
            ParseAnchorDimension(heightRaw, "height")
        );
    }

    /// <summary>
    /// Parse an anchor origin value (x/y) that is either a plain non-negative
    /// integer cell index ("0", "5") or a unit-qualified offset ("2cm", "1in",
    /// "72pt"). Unit-qualified values are converted to a cell index using the
    /// same approximate EMU/column and EMU/row factors as ParseAnchorDimension.
    /// CONSISTENCY(ole-width-units): symmetric with width/height units.
    /// </summary>
    private static int ParseAnchorOrigin(string value, string name)
    {
        if (int.TryParse(value, out var plainInt))
        {
            if (plainInt < 0)
                throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");
            return plainInt;
        }

        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected a non-negative cell index or a unit-qualified offset (e.g. '2cm', '1in') for {name}, got '{value}'.");
        }
        if (emu < 0)
            throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");

        const long emuPerColApprox = 609600;
        const long emuPerRowApprox = 190500;
        if (name == "y")
            return (int)(emu / emuPerRowApprox);
        return (int)(emu / emuPerColApprox);
    }

    /// <summary>
    /// Parse a width/height anchor value that is either a plain integer
    /// cell-count ("3", "5") or a unit-qualified size ("6cm", "2in", "72pt").
    /// Unit-qualified values are converted to an approximate cell count using
    /// Excel's default ~64px (~0.66cm) column width and ~15pt row height.
    /// CONSISTENCY(ole-width-units): Picture/Drawing elsewhere accept ParseEmu;
    /// anchor.x/y stay as cell coordinates, but width/height tolerate EMU units.
    /// </summary>
    private static int ParseAnchorDimension(string value, string name)
    {
        if (int.TryParse(value, out var plainInt))
        {
            // R30-1: negative cell-count is meaningless and silently
            // produced an invalid file. Reject up front. CONSISTENCY with
            // ParseAnchorDimensionEmu's negative-int guard.
            if (plainInt <= 0)
                throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");
            return plainInt;
        }

        // Not a plain integer — treat as EMU-convertible size string.
        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected an integer cell count or a unit-qualified size (e.g. '6cm', '2in') for {name}, got '{value}'.");
        }
        // R30-1: unit-qualified negative ("-2in") parses to a negative
        // EMU; reject so the shape branch matches picture behavior.
        if (emu <= 0)
            throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");

        // Rough conversion: 1 default Excel column ≈ 64px ≈ 0.677cm ≈ 609600 EMU.
        // 1 default Excel row    ≈ 15pt ≈ 0.529cm ≈ 190500 EMU.
        // For width/height passed as a unit, choose the larger of the two
        // converters so "6cm" yields a sensible ~9 columns result either axis.
        const long emuPerColApprox = 609600;
        const long emuPerRowApprox = 190500;
        if (name == "height")
            return Math.Max(1, (int)(emu / emuPerRowApprox));
        return Math.Max(1, (int)(emu / emuPerColApprox));
    }

    // CONSISTENCY(ole-width-units): OLE round-trip preserves sub-cell precision
    // by storing the full EMU extent in ObjectAnchor's From/To ColumnOffset and
    // RowOffset, instead of rounding to whole cells like ParseAnchorDimension.
    // Picture/shape branches keep the integer behavior for now.
    private const long EmuPerColApprox = 609600;
    private const long EmuPerRowApprox = 190500;

    /// <summary>
    /// Parse a width/height anchor value into EMU. Plain integers are treated
    /// as cell counts and multiplied by the default column/row EMU width.
    /// Unit-qualified values (e.g. "6cm", "2in") are parsed via EmuConverter.
    /// </summary>
    private static long ParseAnchorDimensionEmu(string value, string name)
    {
        if (long.TryParse(value, out var plainInt))
        {
            // R30-1: reject negative bare integers up front. Without this,
            // `width=-5` silently rounded to 0 (still invalid) and produced
            // an Excel-rejected file with cx=0/cy=0 anchors.
            if (plainInt <= 0)
                throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");
            // Bare integers are interpreted as cell counts (original grammar),
            // but values that exceed Excel's column max (16384) are clearly
            // EMU — for either axis. Using a single threshold (instead of
            // axis-specific MaxRows=1048576) keeps the heuristic symmetric
            // with ParseAnchorOriginCell so x/y/width/height all flip to
            // EMU at the same boundary.
            const int MaxCellIndex = 16384;
            // R39-2: cell-count form is rejected above the grid limit so
            // mistakes like `width=20000` raise a clear error instead of
            // being silently treated as raw EMU. Users passing EMU should
            // use a unit-qualified form (`914400emu`, `1in`) which is parsed
            // through EmuConverter further down. CONSISTENCY with
            // ParseAnchorOriginCell.
            if (plainInt > MaxCellIndex - 1)
                throw new ArgumentException(
                    $"Picture/shape {name} column/row index must be in [0, {MaxCellIndex - 1}] (got '{value}'). For EMU-scale sizes use a unit-qualified value like '1in' / '6cm' / '72pt'.");
            long perCell = (name == "height") ? EmuPerRowApprox : EmuPerColApprox;
            return plainInt * perCell;
        }

        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected an integer cell count or a unit-qualified size (e.g. '6cm', '2in') for {name}, got '{value}'.");
        }
        // R30-1: unit-qualified negatives (e.g. "-5cm") parse to a negative
        // EMU; reject so we don't write `<xdr:to><xdr:col>-2</xdr:col>...`
        // anchors that crash Excel on open.
        if (emu <= 0)
            throw new ArgumentException($"Picture/shape {name} must be positive (got '{value}').");
        return emu;
    }

    /// <summary>
    /// Parse an <c>anchor=</c> prop value as a cell-reference or cell-range
    /// (e.g. <c>"B2"</c> or <c>"B2:F7"</c>) into 0-based XDR column/row
    /// coordinates. Returns <c>false</c> for anchor-mode strings like
    /// <c>oneCell</c>/<c>twoCell</c>/<c>absolute</c>, which the caller should
    /// route to the anchorMode path instead. Throws <see cref="ArgumentException"/>
    /// for syntactically invalid range strings.
    ///
    /// When only a single cell is supplied, <c>toCol</c>/<c>toRow</c> are set
    /// to <c>-1</c> so callers can fall back to a size-derived extent (e.g.
    /// width/height × EMU-per-cell). The regex mirrors the OLE branch grammar.
    ///
    /// CONSISTENCY(xdr-coords): XDR ColumnId/RowId are 0-based; ColumnNameToIndex
    /// returns 1-based, so this helper subtracts 1 on the way out.
    /// </summary>
    internal static bool TryParseCellRangeAnchor(
        string? value, out int fromCol, out int fromRow, out int toCol, out int toRow)
    {
        fromCol = fromRow = 0;
        toCol = toRow = -1;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var m = System.Text.RegularExpressions.Regex.Match(
            value, @"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return false;
        fromCol = ColumnNameToIndex(m.Groups[1].Value) - 1;
        fromRow = int.Parse(m.Groups[2].Value) - 1;
        ValidateAnchorCell(fromCol, fromRow, value!.Split(':')[0]);
        if (m.Groups[3].Success)
        {
            toCol = ColumnNameToIndex(m.Groups[3].Value) - 1;
            toRow = int.Parse(m.Groups[4].Value) - 1;
            ValidateAnchorCell(toCol, toRow, value.Split(':')[1]);
            // Inverted ranges (D4:B2) written verbatim produce a twoCellAnchor
            // whose from exceeds to; real Excel refuses the file (0x800A03EC)
            // while schema validation stays green. Normalize per axis, same
            // as Excel treats a backwards drag-select.
            NormalizeAnchorRect(ref fromCol, ref fromRow, ref toCol, ref toRow);
        }
        return true;
    }

    /// <summary>Swap anchor corners per axis so from ≤ to.</summary>
    internal static void NormalizeAnchorRect(ref int fromCol, ref int fromRow, ref int toCol, ref int toRow)
    {
        if (toCol >= 0 && toCol < fromCol) (fromCol, toCol) = (toCol, fromCol);
        if (toRow >= 0 && toRow < fromRow) (fromRow, toRow) = (toRow, fromRow);
    }

    /// <summary>
    /// Bounds-check a parsed 0-based anchor cell against Excel's real grid
    /// (A1..XFD1048576). Row 0 (A0) parses to -1 and columns past XFD wrap
    /// into indices Excel refuses (0x800A03EC) while schema validation stays
    /// green — reject at parse time instead.
    /// </summary>
    internal static void ValidateAnchorCell(int col0, int row0, string original)
    {
        const int MaxCol0 = 16383;      // XFD
        const int MaxRow0 = 1048575;    // 1,048,576 rows, 0-based
        if (col0 < 0 || col0 > MaxCol0 || row0 < 0 || row0 > MaxRow0)
            throw new ArgumentException(
                $"Anchor cell '{original}' is outside Excel's grid (A1..XFD1048576).");
    }

    /// <summary>
    /// Clamp a TwoCellAnchor span computed from x/y/width/height (column/row
    /// units) to Excel's grid. The FROM marker must be a real cell
    /// (A1..XFD1048576) and still throws if outside it. The TO marker is the
    /// exclusive right/bottom edge, so it may legitimately sit one past the last
    /// cell (col 16384 / row 1048576) — a picture/shape placed near XFD simply
    /// ends at the grid edge, which is what real Excel does. A width/height that
    /// walks the TO marker further than that (e.g. x=1,width=16384 → toCol 16385)
    /// was silently persisted and made Excel refuse the file (0x800A03EC), so the
    /// TO marker is clamped to the ceiling here rather than written out of range.
    /// Returns the clamped (toCol, toRow).
    /// </summary>
    internal static (int toCol, int toRow) ClampAnchorSpan(int fromCol, int fromRow, int toCol, int toRow, string original)
    {
        ValidateAnchorCell(fromCol, fromRow, original);
        const int MaxToCol = 16384;     // exclusive right edge of XFD
        const int MaxToRow = 1048576;   // exclusive bottom edge of the last row
        return (Math.Clamp(toCol, fromCol, MaxToCol), Math.Clamp(toRow, fromRow, MaxToRow));
    }

    /// <summary>
    /// Return true if the given anchor= value is one of the recognized
    /// anchorMode tokens (oneCell/twoCell/absolute). Used by the picture
    /// branch to disambiguate mode-strings from cell-range strings.
    /// </summary>
    internal static bool IsAnchorModeToken(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        var v = value.Trim().ToLowerInvariant();
        return v is "onecell" or "twocell" or "absolute";
    }

    /// <summary>
    /// Parse x, y (cell indices) + width, height (EMU) for OLE anchors that
    /// need sub-cell precision. See ParseAnchorDimensionEmu for width/height
    /// semantics.
    /// </summary>
    private static (int x, int y, long widthEmu, long heightEmu) ParseAnchorBoundsEmu(
        Dictionary<string, string> properties, string defX, string defY, string defW, string defH)
    {
        return (
            ParseAnchorOriginCell(properties.GetValueOrDefault("x", defX) ?? defX, "x"),
            ParseAnchorOriginCell(properties.GetValueOrDefault("y", defY) ?? defY, "y"),
            ParseAnchorDimensionEmu(properties.GetValueOrDefault("width", defW) ?? defW, "width"),
            ParseAnchorDimensionEmu(properties.GetValueOrDefault("height", defH) ?? defH, "height")
        );
    }

    /// <summary>
    /// Parse anchor x/y origin into a cell index. Plain integers are normally
    /// cell counts, but values that exceed the sheet's column/row max can only
    /// be EMU offsets — fall back to dividing by the per-cell EMU constant so
    /// users passing inch-EMU values (e.g. x=914400) land on a sensible cell
    /// instead of overflowing the FromMarker. CONSISTENCY(ole-width-units):
    /// mirrors ParseAnchorDimensionEmu's "large bare int = EMU" heuristic for
    /// width/height.
    /// </summary>
    private static int ParseAnchorOriginCell(string value, string name)
    {
        if (long.TryParse(value, out var plainInt))
        {
            // R30-1: x/y origins are 0-based cell indices; negative values
            // would write an invalid <xdr:col>/-row anchor. Reject up front.
            if (plainInt < 0)
                throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");
            // Excel's column max (16384) is the tightest sheet-coordinate
            // bound — anything beyond that is unambiguously an EMU offset
            // (rows go to 1048576 but a row index that high is also clearly
            // EMU in practice). Use the same threshold for x and y so users
            // passing inch-EMU (914400) consistently land on a sensible cell
            // on either axis.
            const int MaxCellIndex = 16384;
            // R39-2: bare cell-count form must reject above-grid values
            // outright. Previously, x=20000 hit the "large bare int = EMU"
            // branch and divided by 609600, silently coercing the origin
            // back to col=0 (or row=0 for y). Cell-count input is small
            // by definition; if a user passes a number above the column
            // max, it's either a typo or an EMU value mistakenly fed
            // without a unit suffix. Either way, refuse rather than silently
            // remap. CONSISTENCY with R30-1 negative guard.
            if (plainInt > MaxCellIndex - 1)
                throw new ArgumentException(
                    $"Picture/shape {name} column/row index must be in [0, {MaxCellIndex - 1}] (got '{value}'). For EMU-scale offsets use a unit-qualified value like '1in' / '6cm' / '72pt'.");
            return (int)plainInt;
        }

        // Unit-qualified ("1in", "2cm") → EMU → cell count via the same per-cell constants.
        long emu;
        try
        {
            emu = OfficeCli.Core.EmuConverter.ParseEmu(value);
        }
        catch
        {
            throw new ArgumentException($"Expected an integer cell index or a unit-qualified offset (e.g. '1in', '2cm') for {name}, got '{value}'.");
        }
        if (emu < 0)
            throw new ArgumentException($"Picture/shape {name} must be non-negative (got '{value}').");
        long perCellOut = (name == "y") ? EmuPerRowApprox : EmuPerColApprox;
        return (int)(emu / perCellOut);
    }
}
